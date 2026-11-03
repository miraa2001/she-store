import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import vm from "node:vm";

test("Reem shares Nablus manager details without gaining locations or admin actions", async () => {
  const source = await readFile(new URL("../src/pages/PickupPointPage.jsx", import.meta.url), "utf8");
  const start = source.indexOf('  const isRahaf = profile.role === "rahaf";');
  const end = source.indexOf("  const sidebarLinks", start);
  const policy = source.slice(start, end);
  for (const [role, location, access, details] of [
    ["rahaf", "nablus", true, true],
    ["rahaf", "maryamti", true, true],
    ["reem", "nablus", true, true],
    ["reem", "maryamti", false, false],
    ["nablus", "nablus", true, false],
    ["maryamti", "maryamti", true, false],
    ["rawand", "nablus", false, false],
  ]) {
    const result = vm.runInNewContext(`${policy}; ({canAccess, showManagerDetails, isRahaf})`, {
      profile: { role }, pickupLocation: { id: location, role: location }, showAllOrdersMode: false,
    });
    assert.equal(result.canAccess, access, `${role} at ${location}: access`);
    assert.equal(result.showManagerDetails, details, `${role} at ${location}: details`);
    assert.equal(result.isRahaf, role === "rahaf");
  }
  assert.match(source, /\{canAccess \? \(\s*<div className="pickuppoint-global-summary">/);
  assert.equal((source.match(/\{showManagerDetails \? <td>\{formatDateTime\(purchase\.picked_up_at\)\}/g) || []).length, 2);
  assert.match(source, /colSpan=\{isRahaf \? 7 : showManagerDetails \? 5 : 4\}/);
  for (const handler of ["collectPaymentRows", "collectAllOrders", "startEditPaid", "savePaidPrice", "openTransferDialog", "transferPickupLocation"]) {
    assert.match(source, new RegExp(`function ${handler}\\([^)]*\\) \\{\\s*if \\(!isRahaf`));
  }
});
