import assert from "node:assert/strict";
import test from "node:test";
import { paymentDatabase, id, sqlFile } from "./helpers/payment-database.mjs";

test("order misc deductions refresh profit shares and preserve cash and payout history", async (t) => {
  const { db } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(await sqlFile("20261023020000_deduct_postal_fee_from_profit.sql"));
  await db.exec(`update public.orders set marketing_fee=2,home_profit_deduction=10 where id='${id(10)}';
    select public.record_profit_payout('${id(600)}','mira',1,false);`);
  const table = async (name) => (await db.query(`select * from public.${name} order by id`)).rows;
  const payouts = await table("profit_payouts");
  const cash = await table("home_cash_transactions");
  const migration = await sqlFile("20261025040000_add_order_misc_deduction.sql");
  await db.exec(migration);
  const accrual = async () => (await db.query("select * from public.order_profit_accruals where order_id=$1",[id(10)])).rows[0];
  assert.equal(Number((await accrual()).misc_deduction),0);
  assert.equal(Number((await accrual()).distributable_profit),305);
  await db.exec(`update public.orders set misc_deduction=50 where id='${id(10)}'`);
  const changed = await accrual();
  assert.equal(Number(changed.misc_deduction),50);
  assert.equal(Number(changed.distributable_profit),255);
  assert.equal(Number(changed.home_profit),66.5);
  assert.equal(Number(changed.rahaf_profit),76.5);
  assert.equal(Number(changed.mira_profit),102);
  const balance = (await db.query("select * from public.get_profit_balances() where party='mira'")).rows[0];
  assert.equal(Number(balance.balance),101);
  await db.exec(migration);
  assert.equal(Number((await accrual()).misc_deduction),50);
  for (const value of ["-1","'NaN'","'Infinity'","null"]) {
    await assert.rejects(db.exec(`update public.orders set misc_deduction=${value} where id='${id(10)}'`));
  }
  await db.exec(`update public.orders set misc_deduction=10000 where id='${id(10)}'`);
  assert.equal(Number((await accrual()).distributable_profit),0);
  assert.equal(Number((await accrual()).mira_profit),0);
  assert.deepEqual(await table("home_cash_transactions"),cash);
  assert.deepEqual(await table("profit_payouts"),payouts);
});
