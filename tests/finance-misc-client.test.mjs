import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";

async function financeWith(stubs) {
  const key = `financeTest${Math.random().toString(36).slice(2)}`;
  globalThis[key] = stubs;
  const source = (await readFile(new URL("../src/lib/finance.js", import.meta.url), "utf8"))
    .replace(/^import .*;\r?\n/gm, "");
  return import(`data:text/javascript,${encodeURIComponent(`const {sb, selectOrdersWithOptionalProfitFields, FINANCE_START_AT, isFinanceOrderInPeriod} = globalThis[${JSON.stringify(key)}];\n${source}`)}`)
    .finally(() => { delete globalThis[key]; });
}

test("finance reads new misc deduction and falls back safely before migrations", async () => {
  const requests = [];
  const lib = await financeWith({
    FINANCE_START_AT: "2026-10-01", isFinanceOrderInPeriod: () => true,
    selectOrdersWithOptionalProfitFields: async (columns) => {
      requests.push(columns);
      if (columns.includes("misc_deduction")) return { error: { code: "42703", message: "misc_deduction does not exist" } };
      if (columns.includes("postal_fee")) return { error: { code: "42703", message: "postal_fee does not exist" } };
      return { data: [{ id: "order", spent_amount: 20 }] };
    }
  });
  const result = await lib.fetchFinanceOrders();
  assert.equal(result.miscDeductionAvailable, false);
  assert.equal(result.postalFeeAvailable, false);
  assert.equal(result.data.length, 1);
  assert.equal(requests.length, 3);
  const migrated = await financeWith({ FINANCE_START_AT: "2026-10-01", isFinanceOrderInPeriod: () => true,
    selectOrdersWithOptionalProfitFields: async () => ({ data: [{ id: "order", misc_deduction: 50 }] }) });
  const available = await migrated.fetchFinanceOrders();
  assert.equal(available.miscDeductionAvailable, true);
  assert.equal(available.data[0].misc_deduction, 50);
});

test("finance saves and validates misc deduction; legacy saves do not require the column", async () => {
  let payload, selected;
  const query = {
    update: (value) => { payload = value; return query; }, eq: () => query,
    select: (value) => { selected = value; return query; }, single: async () => ({ data: { id: "order", ...payload } })
  };
  const lib = await financeWith({ sb: { from: () => query } });
  const result = await lib.saveOrderFinance("order", { spent: "20", postalFee: "5", miscDeduction: "9.76" });
  assert.deepEqual(payload, { spent_amount: 20, postal_fee: 5, misc_deduction: 9.76 });
  assert.equal(result.misc_deduction, 9.76);
  assert(selected.includes("misc_deduction"));
  for (const miscDeduction of ["-1", "NaN", "Infinity", "0.001"]) {
    await assert.rejects(lib.saveOrderFinance("order", { spent: 20, postalFee: 5, miscDeduction }));
  }
  await lib.saveOrderFinance("order", { spent: 20, postalFee: 5 });
  assert.equal(Object.hasOwn(payload, "misc_deduction"), false);
  assert.equal(selected.includes("misc_deduction"), false);
});
