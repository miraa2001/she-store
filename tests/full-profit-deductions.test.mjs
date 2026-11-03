import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { paymentDatabase, id, sqlFile } from "./helpers/payment-database.mjs";

test("client deducts the full amount from each party, including negative order shares", async () => {
  const source = await readFile(new URL("../src/lib/profits.js", import.meta.url), "utf8");
  const functionSource = source.match(/export function calculateProfitShares\([^]*?\n\}/)[0];
  const { calculateProfitShares: shares } = await import(`data:text/javascript,${encodeURIComponent(functionSource)}`);
  assert.deepEqual(shares(750, [0, 0, 100], [0, 0, 300]), [0, 0, 450]);
  assert.deepEqual(shares(250, [0, 0, 100], [0, 0, 300]), [0, 0, -50]);
  assert.deepEqual(shares(0, [35, 30, 35], [10, 20, 30]), [-10, -20, -30]);
  assert.deepEqual(shares(-10, [35, 30, 35]), [0, 0, 0]);
  assert.deepEqual(shares(9.76, [35, 30, 35], [4, 3, 4]), [-0.58, -0.07, -0.59]);
  assert.deepEqual(shares(250, [null, null, null], [0, 0, 300]), [null, null, null]);
});

test("full deductions recalculate accumulated balances and preserve payouts and cash", async t => {
  const { db } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(await sqlFile("20261025040000_add_order_misc_deduction.sql"));
  await db.exec(`select public.record_profit_payout('${id(600)}','home',1,false);
    update public.orders set home_profit_percent=100,rahaf_profit_percent=0,mira_profit_percent=0;
    insert into public.orders(id,order_name,home_profit_percent,rahaf_profit_percent,mira_profit_percent) values
      ('${id(20)}','Order 1',0,0,100),('${id(21)}','Order 2',0,0,100),('${id(22)}','Order 3',0,0,100);
    insert into public.purchases(id,order_id,customer_name,price,paid_price,pickup_point) values
      ('${id(300)}','${id(20)}','One',999,250,'من البيت'),
      ('${id(301)}','${id(21)}','Two',999,250,'من البيت'),
      ('${id(302)}','${id(22)}','Three',999,250,'من البيت');
    update public.orders set mira_profit_deduction=300 where id='${id(22)}';`);
  const balance = async () => (await db.query("select * from public.get_profit_balances() where party='mira'")).rows[0];
  const rows = async table => (await db.query(`select * from public.${table} order by id`)).rows;
  const cash = await rows("home_cash_transactions");
  const payouts = await rows("profit_payouts");
  assert.equal(Number((await balance()).balance), 500);
  const migration = await sqlFile("20261104000000_apply_full_party_profit_deductions.sql");
  await db.exec(migration);
  assert.equal(Number((await balance()).earned), 450);
  assert.equal(Number((await balance()).balance), 450);
  assert.equal(Number((await db.query('select mira_profit from public.order_profit_accruals where order_id=$1', [id(22)])).rows[0].mira_profit), -50);
  await db.exec(migration);
  assert.equal(Number((await balance()).balance), 450);
  assert.deepEqual(await rows("home_cash_transactions"), cash);
  assert.deepEqual(await rows("profit_payouts"), payouts);
  await db.exec(`select public.record_profit_payout('${id(601)}','mira',250,false);`);
  assert.equal(Number((await balance()).balance), 200);
  await db.exec(`update public.orders set mira_profit_deduction=1000,home_profit_deduction=10,rahaf_profit_deduction=20 where id='${id(22)}';`);
  assert.equal(Number((await balance()).balance), -500);
  const accrual = (await db.query('select * from public.order_profit_accruals where order_id=$1', [id(22)])).rows[0];
  assert.equal(Number(accrual.home_profit), -10);
  assert.equal(Number(accrual.rahaf_profit), -20);
  await assert.rejects(db.exec(`select public.record_profit_payout('${id(602)}','mira',1,false);`));
  assert.deepEqual(await rows("home_cash_transactions"), cash);
});
