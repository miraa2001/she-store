import assert from "node:assert/strict";
import test from "node:test";
import { paymentDatabase, id, sqlFile } from "./helpers/payment-database.mjs";

test("Order purchase transfers preserve payments and recalculate both orders", async (t) => {
  const { db, identity } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(await sqlFile("20261023010000_transfer_purchases_between_orders.sql"));
  await db.exec(`
    update public.orders set placed_at_pickup=false where id='${id(10)}';
    update public.purchases set ready_for_pickup=false where id in ('${id(100)}','${id(101)}','${id(102)}');
    insert into public.orders(id,order_name,spent_amount,postal_fee,placed_at_pickup,
      home_profit_percent,rahaf_profit_percent,mira_profit_percent)
      values('${id(20)}','Destination',10,0,false,20,30,50);
  `);
  const move = (ids, target = id(20)) => db.query(
    "select public.transfer_order_purchases($1,$2,$3) as count", [id(10), target, ids]);
  const accrual = async (order) => (await db.query("select * from public.order_profit_accruals where order_id=$1", [id(order)])).rows[0];
  const initial = Number((await accrual(10)).purchase_value);
  await identity(1);

  await t.test("rejects empty, same-order and stale selections without moving anything", async () => {
    await assert.rejects(move([]));
    await assert.rejects(move([id(100)], id(10)));
    await assert.rejects(move([id(100), id(999)]));
    assert.equal(Number((await accrual(10)).purchase_value), initial);
  });
  await t.test("rejects dispatched purchases and orders", async () => {
    await assert.rejects(move([id(103)]));
    await db.exec("reset role");
    await db.exec(`update public.orders set placed_at_pickup=true where id='${id(20)}'`);
    await identity(1);
    await assert.rejects(move([id(100)]));
    await db.exec("reset role");
    await db.exec(`update public.orders set placed_at_pickup=false where id='${id(20)}'`);
    await identity(1);
  });
  await t.test("only admin can transfer", async () => {
    for (const user of [2, 4, 5]) {
      await identity(user);
      await assert.rejects(move([id(100)]));
    }
    await identity(1);
  });
  await t.test("moves a batch once, using agreed prices and destination percentages", async () => {
    const result = await move([id(100), id(101), id(100)]);
    assert.equal(result.rows[0].count, 2);
    const destination = await accrual(20);
    assert.equal(Number(destination.purchase_value), 142);
    assert.equal(Number(destination.distributable_profit), 132);
    assert.equal(Number(destination.mira_profit), 66);
    assert.equal(Number((await accrual(10)).purchase_value), initial - 142);
    assert.equal(Number((await accrual(10)).spent), 20);
    const moved = (await db.query("select order_id,pickup_point from public.purchases where id=$1", [id(100)])).rows[0];
    assert.equal(moved.order_id, id(20));
    assert.equal(moved.pickup_point, "من البيت");
    await assert.rejects(move([id(100)]));
  });
  await t.test("payment history prevents transfers even when not yet received", async () => {
    await db.exec("reset role");
    await db.exec(`insert into public.purchase_payments(id,purchase_id,amount,pickup_point)
      values('${id(500)}','${id(102)}',10,'توصيل')`);
    await identity(1);
    await assert.rejects(move([id(102)]));
  });
});
