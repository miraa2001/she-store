import assert from "node:assert/strict";
import test from "node:test";
import { comparisonCategory, generateComparisons, selectedComparisonTotal, summarizeComparison, buildComparisonPayload, priceCents } from "../src/lib/purchaseComparison.js";

const purchase = (id, price, qty = 1, category = "fast") => ({ id, price, qty, paid_price: 9999,
  purchase_categories: [category], comparison_category: category, order_id: null });
const group = (key, target, category = "fast") => ({ key, target, category, name: `Order ${key}`, mode: "new" });

test("selection total uses only selected original prices and updates down to zero", () => {
  const rows = [purchase("a", 9.76, 8), purchase("b", 0.24, 5)];
  assert.equal(selectedComparisonTotal(rows), 10);
  assert.equal(selectedComparisonTotal(rows.filter((p) => p.id === "a")), 9.76);
  assert.equal(selectedComparisonTotal([]), 0);
});

test("comparison uses original price once per purchase, not paid price or price times qty", () => {
  const rows = [purchase("a", 9.76, 8), purchase("b", 0.24, 2)];
  const best = generateComparisons(rows, [group("x", 10)])[0];
  assert.equal(best.groups[0].cents, 1000);
  assert.equal(best.groups[0].qty, 10);
  assert.equal(best.deviation, 0);
});

test("separate targets produce distinct whole-purchase alternatives", () => {
  const rows = [purchase("a", 40), purchase("b", 60), purchase("c", 45), purchase("d", 55)];
  const groups = [group("x", 100), group("y", 100)];
  const plans = generateComparisons(rows, groups);
  assert.equal(plans[0].included, 4);
  assert.equal(plans[0].deviation, 0);
  assert.equal(plans.length, 3);
  assert.equal(new Set(plans.map((p) => rows.map((row) => p.assignments[row.id]).join())).size, plans.length);
  const unequal = generateComparisons([purchase("a", 75), purchase("b", 125)], [group("x", 75), group("y", 125)])[0];
  assert.equal(unequal.deviation, 0);
});

test("100-item cap uses qty, leaves oversized records out, and prioritizes purchase count", () => {
  const rows = [purchase("a", 100, 60), purchase("b", 50, 40), purchase("c", 1000, 1), purchase("d", 100, 101)];
  const best = generateComparisons(rows, [group("x", 100)])[0];
  assert.equal(best.included, 2);
  assert.equal(best.groups[0].qty, 100);
  assert.equal(best.groups[0].cents, 15000);
  assert(best.leftovers.some((p) => p.id === "d"));
  // Including two records wins over a single record that exactly meets the target.
  const countFirst = generateComparisons([purchase("a", 100, 100), purchase("b", 200, 50), purchase("c", 200, 50)], [group("x", 100)])[0];
  assert.equal(countFirst.included, 2);
  assert.equal(countFirst.groups[0].cents, 40000);
});

test("mixed and unlabelled purchases require an explicit valid comparison category", () => {
  const mixed = { ...purchase("a", 20), purchase_categories: ["fast", "makeup"], comparison_category: "" };
  assert.equal(comparisonCategory(mixed, "electronics"), "");
  assert.equal(comparisonCategory(mixed, "makeup"), "makeup");
  assert.throws(() => generateComparisons([mixed], [group("x", 20)]));
  mixed.comparison_category = "makeup";
  const bare = { ...purchase("b", 10), purchase_categories: [], comparison_category: "electronics" };
  const best = generateComparisons([mixed, bare], [group("x", 20, "makeup"), group("y", 10, "electronics")])[0];
  assert.equal(best.deviation, 0);
  assert.equal(best.assignments.a, "x");
  assert.equal(best.assignments.b, "y");
  assert.deepEqual(mixed.purchase_categories, ["fast", "makeup"]);
});

test("empty/category-incompatible groups cannot be fabricated", () => {
  assert.throws(() => generateComparisons([purchase("a", 20)], [group("x", 20), group("y", 10)]));
  assert.throws(() => generateComparisons([purchase("a", 20)], [group("x", 20, "makeup")]));
  assert.throws(() => priceCents("NaN"));
  assert.throws(() => generateComparisons([purchase("a", 20)], [group("x", 0)]));
});

test("manual edits and confirmation enforce category, nonempty groups, capacity and unique destinations", () => {
  const rows = [purchase("a", 40, 60), purchase("b", 50, 50), purchase("c", 10, 1, "makeup")];
  const groups = [group("x", 40), group("y", 50), group("z", 10, "makeup")];
  assert.equal(summarizeComparison(rows, groups, { a: "x", b: "x", c: "z" }).valid, false);
  assert.equal(summarizeComparison(rows, groups, { a: "x", b: "y", c: "x" }).valid, false);
  const assignment = { a: "x", b: "y", c: "z" };
  const payload = buildComparisonPayload(rows, groups, assignment);
  assert.equal(payload[0].items[0].price, 40);
  assert.equal(payload[0].items[0].qty, 60);
  assert.equal(payload[0].order_id, null);
  assert.throws(() => buildComparisonPayload(rows, groups.map((g) => ({ ...g, mode: "existing", orderId: "same" })), assignment));
});

test("small optimization problems match exhaustive count-first, deviation-second results", () => {
  for (let seed = 1; seed <= 6; seed++) {
    const rows = Array.from({ length: 5 }, (_, i) => purchase(String(i), (seed * 17 + i * 23) % 81 + 5, (seed * 11 + i * 13) % 60 + 1));
    const groups = [group("x", 80), group("y", 110)];
    let expected = { included: -1, deviation: Infinity };
    for (let code = 0; code < 3 ** rows.length; code++) {
      let value = code;
      const assignment = {};
      rows.forEach((row) => { const dest = value % 3; value = Math.floor(value / 3); if (dest) assignment[row.id] = groups[dest - 1].key; });
      const summary = summarizeComparison(rows, groups, assignment);
      if (summary.valid && (summary.included > expected.included || (summary.included === expected.included && summary.deviation < expected.deviation))) expected = summary;
    }
    const best = generateComparisons(rows, groups, { alternatives: 1, timeout: 10000 })[0];
    assert.equal(best.included, expected.included);
    assert.equal(best.deviation, expected.deviation);
  }
});
