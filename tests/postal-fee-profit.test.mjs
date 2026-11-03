import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { paymentDatabase, id, sqlFile } from "./helpers/payment-database.mjs";

test("Finance display subtracts postal fees with cent precision", async () => {
  const source = await readFile(new URL("../src/lib/finance.js", import.meta.url), "utf8");
  const functionSource = source.match(/export function calculateOrderFinanceProfit\([^]*?\n\}/)[0];
  const { calculateOrderFinanceProfit: profit } = await import(`data:text/javascript,${encodeURIComponent(functionSource)}`);
  assert.equal(profit(1000, 20, 200), 780);
  assert.equal(profit(9.76, 0.25, 1.01), 8.5);
  assert.equal(profit(10, 20, 0), -10);
  assert.equal(profit(1000, 20, 200, 50), 730);
  assert.equal(profit(9.76, 0.25, 1.01, 0.15), 8.35);
});

test("Postal correction recalculates shares without changing cash or payouts", async (t) => {
  const { db } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(`update public.orders set marketing_fee=10,mira_profit_deduction=2 where id='${id(10)}';
    select public.record_profit_payout('${id(600)}','mira',1,false);`);
  const rows = async (table) => (await db.query(`select * from public.${table} order by id`)).rows;
  const payouts = await rows("profit_payouts");
  const cash = await rows("home_cash_transactions");
  const migration = await sqlFile("20261023020000_deduct_postal_fee_from_profit.sql");
  await db.exec(migration);
  const accrual = async () => (await db.query("select * from public.order_profit_accruals where order_id=$1", [id(10)])).rows[0];
  let value = await accrual();
  assert.equal(Number(value.distributable_profit), Number(value.purchase_value) - 20 - 5 - 10);
  assert.equal(Number(value.mira_profit), Math.round(Number(value.distributable_profit) * 40) / 100 - 2);
  assert.deepEqual(await rows("profit_payouts"), payouts);
  assert.deepEqual(await rows("home_cash_transactions"), cash);
  await db.exec(migration);
  assert.equal(Number((await accrual()).distributable_profit), Number(value.distributable_profit));
  await db.exec(`update public.orders set postal_fee=15 where id='${id(10)}'`);
  assert.equal(Number((await accrual()).distributable_profit), Number(value.distributable_profit) - 10);
  await db.exec(`update public.orders set postal_fee=10000 where id='${id(10)}'`);
  value = await accrual();
  assert.equal(Number(value.distributable_profit), 0);
  assert.equal(Number(value.mira_profit), 0);
  assert.deepEqual(await rows("profit_payouts"), payouts);
});
