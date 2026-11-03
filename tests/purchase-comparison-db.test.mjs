import assert from "node:assert/strict";
import test from "node:test";
import { paymentDatabase, id, sqlFile, HOME } from "./helpers/payment-database.mjs";

test("comparison confirmation is atomic, idempotent, admin-only and shipping-safe", async (t) => {
  const { db, identity } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(`
    create table public.purchase_links(id uuid primary key,purchase_id uuid references public.purchases,url text);
    create table public.purchase_images(id uuid primary key,purchase_id uuid references public.purchases,storage_path text);
    alter table public.purchases add column qty integer not null default 1;
    alter table public.orders add column order_type text default 'shein',add column shein_type text default 'normal';
    grant select on public.orders to authenticated;
    insert into public.orders(id,order_name,arrived,placed_at_pickup,spent_amount,postal_fee,marketing_fee,
      home_profit_percent,rahaf_profit_percent,mira_profit_percent)
      values('${id(20)}','Existing Fast',false,false,5,2,1,20,30,50);
  `);
  for (const file of ["20261023020000_deduct_postal_fee_from_profit.sql", "20261025010000_add_unassigned_purchases.sql",
    "20261025020000_add_purchase_categories.sql", "20261025030000_add_purchase_comparison.sql"]) await db.exec(await sqlFile(file));
  for (const [n, price, qty, labels] of [[300,40,40,["fast"]], [301,60,60,["fast"]], [302,50,10,["makeup"]],
    [303,20,2,["fast","makeup"]], [304,9.76,1,[]], [305,100,101,["electronics"]]]) {
    await db.query(`insert into public.purchases(id,customer_name,price,paid_price,qty,pickup_point,ready_for_pickup,purchase_categories)
      values($1,'Customer',$2,$3,$4,$5,false,$6)`, [id(n),price,price+100,qty,HOME,labels]);
  }
  const row = async (n) => (await db.query("select * from public.purchases where id=$1", [id(n)])).rows[0];
  const item = async (n) => { const p = await row(n); return { id: p.id,price:p.price,qty:p.qty,categories:p.purchase_categories }; };
  const group = async (category, numbers, orderId = null) => ({ category, target: 100, name: `${category} comparison`, order_id: orderId,
    items: await Promise.all(numbers.map(item)) });
  const apply = (n, groups) => db.query("select public.apply_purchase_comparison($1,$2::jsonb) as result", [id(n), JSON.stringify(groups)]);
  const count = async (table) => Number((await db.query(`select count(*) as n from public.${table}`)).rows[0].n);
  const initialOrders = await count("orders");
  const cash = (await db.query("select * from public.home_cash_transactions order by id")).rows;
  const testPayload = [await group("fast", [300])];

  await t.test("viewers, pickup staff and anonymous cannot confirm or rewrite audit records", async () => {
    for (const actor of [2,3,4,5]) {
      await identity(actor);
      await assert.rejects(apply(600, testPayload));
      assert.equal(await count("purchase_comparison_commits"), 0);
    }
    await db.exec("reset role; set role anon");
    await assert.rejects(apply(600, testPayload));
    await identity(1);
    await assert.rejects(db.exec("delete from public.purchase_comparison_commits"));
  });
  await t.test("stale prices, quantities and labels reject confirmation", async () => {
    const payload = [await group("fast", [300])];
    for (const property of ["price","qty","categories"]) {
      const bad = structuredClone(payload);
      bad[0].items[0][property] = property === "categories" ? ["makeup"] : 999;
      await assert.rejects(apply(600,bad));
    }
    assert.equal((await row(300)).order_id,null);
  });
  await t.test("duplicates, mixed categories and capacity overflow cannot create partial orders", async () => {
    await assert.rejects(apply(600,[await group("fast",[300]), await group("makeup",[301])]));
    await assert.rejects(apply(600,[await group("fast",[300]), await group("fast",[300])]));
    await assert.rejects(apply(600,[await group("electronics",[305])]));
    assert.equal(await count("orders"),initialOrders);
    assert.equal(await count("purchase_comparison_commits"),0);
    assert.equal((await row(300)).order_id,null);
  });
  await t.test("destinations must be existing empty pending orders of the same category", async () => {
    await assert.rejects(apply(600,[await group("fast",[300],id(10))]));
    await assert.rejects(apply(600,[await group("makeup",[302],id(20))]));
    await assert.rejects(apply(600,[await group("fast",[300],id(999))]));
    await db.exec(`reset role; update public.orders set arrived=true where id='${id(20)}'`);
    await identity(1);
    await assert.rejects(apply(600,[await group("fast",[300],id(20))]));
    await db.exec(`reset role; update public.orders set arrived=false where id='${id(20)}'`);
    await identity(1);
    await assert.rejects(apply(600,[await group("fast",[300],id(20)), await group("fast",[301],id(20))]));
    assert.equal((await row(300)).order_id,null);
  });
  await t.test("stale assignment and payment history cannot be bypassed", async () => {
    const payload = [await group("fast",[300])];
    await db.exec(`reset role; update public.purchases set order_id='${id(20)}' where id='${id(300)}'`);
    await identity(1);
    await assert.rejects(apply(600,payload));
    await db.exec(`reset role; update public.purchases set order_id=null where id='${id(300)}';
      insert into public.purchase_payments(id,purchase_id,amount,pickup_point) values('${id(700)}','${id(300)}',1,'${HOME}');`);
    await identity(1);
    await assert.rejects(apply(600,payload));
    await db.exec(`reset role; delete from public.purchase_payments where id='${id(700)}'`);
    await identity(1);
  });
  let committed;
  let payload;
  await t.test("one confirmation combines existing/new orders, preserves labels, and uses original totals", async () => {
    payload = [await group("fast",[300,301],id(20)), await group("makeup",[302,303]), await group("electronics",[304])];
    committed = (await apply(601,payload)).rows[0].result;
    assert.equal(committed.orders.length,3);
    assert.equal(committed.purchase_count,5);
    assert.equal(committed.orders[0].total,100);
    assert.equal(committed.orders[0].qty,100);
    assert.equal(committed.orders[1].total,70);
    assert.equal(committed.orders[2].total,9.76);
    assert.equal(await count("orders"),initialOrders+2);
    assert.equal((await row(305)).order_id,null);
    assert.deepEqual((await row(303)).purchase_categories,["fast","makeup"]);
    const accrual=(await db.query("select * from public.order_profit_accruals where order_id=$1",[id(20)])).rows[0];
    assert.equal(Number(accrual.purchase_value),300);
    assert.equal(Number(accrual.distributable_profit),292);
    assert.equal(Number(accrual.mira_profit),146);
    for (const destination of committed.orders.slice(1)) {
      const order=(await db.query("select * from public.orders where id=$1",[destination.order_id])).rows[0];
      assert.equal(order.order_type,"shein");
      assert.equal(order.shein_type,destination.category);
      assert.equal(order.arrived,false);
      assert.equal(order.placed_at_pickup,false);
      assert.equal(order.mira_profit_percent,null);
    }
    assert.deepEqual((await db.query("select * from public.home_cash_transactions order by id")).rows,cash);
  });
  await t.test("identical retries return the original result, changed retries and stale fresh requests fail", async () => {
    assert.deepEqual((await apply(601,payload)).rows[0].result,committed);
    assert.equal(await count("orders"),initialOrders+2);
    const changed=structuredClone(payload); changed[0].target=200;
    await assert.rejects(apply(601,changed));
    await assert.rejects(apply(602,payload));
    assert.equal(await count("purchase_comparison_commits"),1);
    assert.equal(await count("orders"),initialOrders+2);
  });
});
