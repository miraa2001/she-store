import assert from "node:assert/strict";
import test from "node:test";
import { FINANCE_START_AT, isFinanceOrderInPeriod } from "../src/lib/financePeriod.js";

test("Finance begins on October 1, 2026 in Asia/Hebron", () => {
  assert.equal(Date.parse(FINANCE_START_AT), Date.parse("2026-09-30T21:00:00Z"));
  assert.equal(isFinanceOrderInPeriod({ created_at: "2026-09-30T20:59:59.999Z" }), false);
  assert.equal(isFinanceOrderInPeriod({ created_at: "2026-09-30T21:00:00Z" }), true);
  assert.equal(isFinanceOrderInPeriod({ created_at: "2026-10-01T00:00:00Z" }), true);
  assert.equal(isFinanceOrderInPeriod({ createdAt: "2027-01-01T00:00:00Z" }), true);
});

test("The cutoff uses creation time, not order date, and excludes missing dates", () => {
  assert.equal(isFinanceOrderInPeriod({ created_at: "2026-09-10T12:00:00Z", order_date: "2026-10-03" }), false);
  assert.equal(isFinanceOrderInPeriod({ created_at: "2026-10-03T12:00:00Z", order_date: "2026-09-10" }), true);
  for (const created_at of [null, undefined, "", "invalid"]) assert.equal(isFinanceOrderInPeriod({ created_at }), false);
});
