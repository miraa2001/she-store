import assert from "node:assert/strict";
import test from "node:test";
import { filterPurchasesByCategory } from "../src/lib/purchaseCategories.js";
import { searchByName } from "../src/lib/search.js";

const rows = [
  { id: 1, customer_name: "A", price: 9.76, paid_price: 999, qty: 8, purchase_categories: ["fast"] },
  { id: 2, customer_name: "B", price: 20, purchase_categories: ["electronics"] },
  { id: 3, customer_name: "Mixed", price: 30, purchase_categories: ["fast", "makeup", "electronics"] },
  { id: 4, customer_name: "Untagged", price: 40, purchase_categories: [] }
];

test("category filters include mixed labels once, while all includes untagged purchases", () => {
  assert.equal(filterPurchasesByCategory(rows, "all"), rows);
  assert.deepEqual(filterPurchasesByCategory(rows, "fast").map((row) => row.id), [1, 3]);
  assert.deepEqual(filterPurchasesByCategory(rows, "electronics").map((row) => row.id), [2, 3]);
  assert.deepEqual(filterPurchasesByCategory(rows, "makeup").map((row) => row.id), [3]);
  assert.deepEqual(filterPurchasesByCategory([], "electronics"), []);
});

test("search intersects the category filter and totals use original price once per visible purchase", () => {
  const filtered = filterPurchasesByCategory(rows, "electronics");
  assert.equal(filtered.reduce((sum, row) => sum + row.price, 0), 50);
  const searched = searchByName(filtered, "mixed", (row) => row.customer_name);
  assert.deepEqual(searched.map((row) => row.id), [3]);
  assert.equal(searched.reduce((sum, row) => sum + row.price, 0), 30);
  assert.equal(searchByName(filtered, "absent", (row) => row.customer_name).length, 0);
  assert.equal(filterPurchasesByCategory(rows, "fast").reduce((sum, row) => sum + Math.round(row.price * 100), 0), 3976);
});
