import assert from "node:assert/strict";
import test from "node:test";
import { paymentDatabase, id, sqlFile, HOME } from "./helpers/payment-database.mjs";
import { normalizePurchaseCategories } from "../src/lib/purchaseCategories.js";

test("Purchase categories accept any combination with no duplicates", () => {
  assert.deepEqual(normalizePurchaseCategories(["makeup", "fast", "fast"]), ["fast", "makeup"]);
  assert.deepEqual(normalizePurchaseCategories(["electronics", "fast", "makeup"]), ["fast", "makeup", "electronics"]);
  assert.deepEqual(normalizePurchaseCategories(["unknown"]), []);
  assert.deepEqual(normalizePurchaseCategories(null), []);
});

test("Unassigned purchases stay private, pending-only and outside cash/profits", async (t) => {
  const { db, identity } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(`
    create table public.purchase_links(id uuid primary key,purchase_id uuid references public.purchases on delete cascade,url text);
    create table public.purchase_images(id uuid primary key,purchase_id uuid references public.purchases on delete cascade,storage_path text);
    alter table public.purchase_links enable row level security;
    alter table public.purchase_images enable row level security;
    create policy links_read on public.purchase_links for select to authenticated using(true);
    create policy images_read on public.purchase_images for select to authenticated using(true);
    grant select on public.purchase_links,public.purchase_images to authenticated;
    alter table public.purchases alter column order_id set not null;
    update public.orders set arrived=false,placed_at_pickup=false where id='${id(10)}';
    update public.purchases set ready_for_pickup=false,ready_for_pickup_at=null where id in ('${id(100)}','${id(101)}');
    insert into public.orders(id,order_name,arrived,placed_at_pickup,home_profit_percent,rahaf_profit_percent,mira_profit_percent)
      values('${id(20)}','Pending destination',false,false,20,30,50);
  `);
  await db.exec(await sqlFile("20261023020000_deduct_postal_fee_from_profit.sql"));
  const migration = await sqlFile("20261025010000_add_unassigned_purchases.sql");
  await db.exec(migration);
  const categoriesMigration = await sqlFile("20261025020000_add_purchase_categories.sql");
  await db.exec(categoriesMigration);
  const move = (source, target, ids) => db.query("select public.transfer_order_purchases($1,$2,$3) as count", [source, target, ids]);
  const purchase = async (n) => (await db.query("select * from public.purchases where id=$1", [id(n)])).rows[0];
  const profit = async (n) => (await db.query("select * from public.order_profit_accruals where order_id=$1", [id(n)])).rows[0];
  const initial = Number((await profit(10)).purchase_value);
  const cash = (await db.query("select * from public.home_cash_transactions order by id")).rows;

  await t.test("admin creates and edits a normal unassigned purchase with no profit or cash", async () => {
    await identity(1);
    await db.query(`insert into public.purchases(id,order_id,customer_name,price,paid_price,pickup_point,ready_for_pickup)
      values($1,null,'Unassigned customer',60,50,$2,false)`, [id(300), HOME]);
    await db.query("update public.purchases set customer_name='Edited customer' where id=$1", [id(300)]);
    assert.equal((await purchase(300)).order_id, null);
    assert.equal(Number((await profit(10)).purchase_value), initial);
    assert.equal(Number((await profit(20)).purchase_value), 0);
    await db.exec("reset role");
    await db.exec(`insert into public.purchase_links values('${id(310)}','${id(300)}','https://example.com');
      insert into public.purchase_images values('${id(311)}','${id(300)}','unassigned/${id(300)}/image.jpg');`);
  });
  await t.test("viewers and pickup staff cannot see unassigned purchases or their metadata", async () => {
    for (const actor of [2,3,4,5,6]) {
      await identity(actor);
      assert.equal(await purchase(300), undefined);
      assert.equal((await db.query("select * from public.purchase_links")).rows.length, 0);
      assert.equal((await db.query("select * from public.purchase_images")).rows.length, 0);
      await assert.rejects(move(null, id(10), [id(300)]));
    }
    await identity(1);
    assert.equal((await db.query("select * from public.purchase_links")).rows.length, 1);
  });
  await t.test("unassigned purchases cannot be sent to pickup or assigned by direct update", async () => {
    await assert.rejects(db.query("update public.purchases set ready_for_pickup=true where id=$1", [id(300)]));
    await assert.rejects(db.query("update public.purchases set order_id=$2 where id=$1", [id(300),id(10)]));
    assert.equal((await purchase(300)).ready_for_pickup, false);
  });
  await t.test("optional category combinations persist without changing money; invalid labels are rejected", async () => {
    assert.deepEqual((await purchase(300)).purchase_categories, []);
    for (const combination of [["fast"], ["makeup", "electronics"], ["fast", "makeup", "electronics"], []]) {
      await db.query("update public.purchases set purchase_categories=$2 where id=$1", [id(300), combination]);
      assert.deepEqual((await purchase(300)).purchase_categories, combination);
    }
    for (const categories of [["unknown"], ["fast", null], null]) {
      await assert.rejects(db.query("update public.purchases set purchase_categories=$2 where id=$1", [id(300), categories]));
    }
    assert.equal(Number((await profit(10)).purchase_value), initial);
    assert.equal(Number((await profit(20)).purchase_value), 0);
    await db.query("update public.purchases set purchase_categories=$2 where id=$1", [id(300), ["fast", "makeup"]]);
  });
  await t.test("batch detach removes purchase values and recalculates source shares", async () => {
    assert.equal((await move(id(10), null, [id(100),id(101)])).rows[0].count, 2);
    assert.equal((await purchase(100)).order_id, null);
    assert.equal(Number((await profit(10)).purchase_value), initial - 142);
    await assert.rejects(move(id(10), null, [id(100)]));
  });
  await t.test("batch assignment preserves IDs, images, prices and recalculates destination shares", async () => {
    assert.equal((await move(null, id(20), [id(100),id(300)])).rows[0].count, 2);
    assert.equal((await purchase(300)).order_id, id(20));
    assert.equal((await purchase(300)).customer_name, "Edited customer");
    assert.deepEqual((await purchase(300)).purchase_categories, ["fast", "makeup"]);
    assert.equal(Number((await purchase(300)).paid_price), 50);
    assert.equal((await db.query("select storage_path from public.purchase_images")).rows[0].storage_path, `unassigned/${id(300)}/image.jpg`);
    assert.equal(Number((await profit(20)).purchase_value), 92);
    assert.equal(Number((await profit(20)).mira_profit), 46);
    await assert.rejects(move(null, id(20), [id(100)]));
  });
  await t.test("arrived sources and destinations, missing IDs and mixed-source batches fail atomically", async () => {
    await db.exec(`reset role; update public.orders set arrived=true where id='${id(20)}'`);
    await identity(1);
    await assert.rejects(move(null, id(20), [id(101)]));
    await assert.rejects(move(id(20), null, [id(100)]));
    await assert.rejects(move(null, id(10), [id(101),id(300)]));
    await assert.rejects(move(null, id(999), [id(101)]));
    await assert.rejects(move(null, null, [id(101)]));
    await assert.rejects(move(null, id(10), []));
    assert.equal((await purchase(101)).order_id, null);
  });
  await t.test("readiness and payment history prevent detach even after resetting order status", async () => {
    await assert.rejects(move(id(10), null, [id(103)]));
    await db.exec(`reset role; update public.purchases set ready_for_pickup=false where id='${id(102)}';
      insert into public.purchase_payments(id,purchase_id,amount,pickup_point)
      values('${id(500)}','${id(102)}',10,'${HOME}');`);
    await identity(1);
    await assert.rejects(move(id(10), null, [id(102)]));
  });
  await t.test("migration reruns preserve unassigned records, financial settings and cash", async () => {
    await db.exec("reset role");
    await db.exec(migration);
    await db.exec(categoriesMigration);
    assert.deepEqual((await purchase(300)).purchase_categories, ["fast", "makeup"]);
    assert.equal((await purchase(101)).order_id, null);
    assert.equal(Number((await profit(10)).spent), 20);
    assert.deepEqual((await db.query("select * from public.home_cash_transactions order by id")).rows, cash);
  });
});
