import solver from "javascript-lp-solver";
import { normalizePurchaseCategories } from "./purchaseCategories.js";

export const COMPARISON_ITEM_LIMIT = 100;
export const COMPARISON_ORDER_LIMIT = 20;

export function comparisonCategory(purchase, choice) {
  const labels = normalizePurchaseCategories(purchase.purchase_categories);
  if (labels.length === 1) return labels[0];
  return ["fast", "makeup", "electronics"].includes(choice)
    && (!labels.length || labels.includes(choice)) ? choice : "";
}

export function priceCents(value) {
  const amount = Number(value);
  if (value === null || value === "" || !Number.isFinite(amount) || amount < 0
    || !Number.isSafeInteger(Math.round(amount * 100))) throw new Error("السعر الأصلي غير صالح");
  return Math.round(amount * 100);
}

export function selectedComparisonTotal(purchases) {
  const cents = purchases.reduce((sum, purchase) => sum + priceCents(purchase.price), 0);
  if (!Number.isSafeInteger(cents)) throw new Error("مجموع الأسعار أكبر من الحد المدعوم");
  return cents / 100;
}

export function validateComparison(purchases, groups) {
  if (!purchases.length) throw new Error("اختاري مشتريات للمقارنة");
  if (purchases.length > 2000) throw new Error("اختاري 2000 مشترى كحد أقصى في المقارنة الواحدة");
  if (!groups.length || groups.length > COMPARISON_ORDER_LIMIT) throw new Error("عدد الطلبات يجب أن يكون بين 1 و20");
  if (new Set(purchases.map((p) => p.id)).size !== purchases.length) throw new Error("مشتريات مكررة");
  if (new Set(groups.map((g) => g.key)).size !== groups.length) throw new Error("طلبات مكررة");
  purchases.forEach((purchase) => {
    priceCents(purchase.price);
    if (!Number.isInteger(Number(purchase.qty)) || Number(purchase.qty) < 1) throw new Error("عدد القطع غير صالح");
    if (purchase.order_id || purchase.ready_for_pickup || purchase.picked_up || purchase.collected) throw new Error("المقارنة للمشتريات غير المعينة فقط");
    if (!comparisonCategory(purchase, purchase.comparison_category)) throw new Error("اختاري تصنيفاً لكل مشترى محدد");
  });
  if (!Number.isSafeInteger(purchases.reduce((sum, p) => sum + priceCents(p.price), 0))) throw new Error("مجموع الأسعار أكبر من الحد المدعوم");
  groups.forEach((group) => {
    if (!["fast", "makeup", "electronics"].includes(group.category) || group.target == null
      || String(group.target).trim() === "" || !Number.isFinite(Number(group.target)) || Number(group.target) <= 0
      || Number(group.target) > 1e9 || Math.abs(Number(group.target) * 100 - priceCents(group.target)) > 0.00001) {
      throw new Error("اختاري تصنيف الطلب والمبلغ المستهدف بمنزلتين عشريتين كحد أقصى");
    }
  });
}

export function summarizeComparison(purchases, groups, assignments) {
  const groupIds = new Set(groups.map((g) => g.key));
  const summaries = groups.map((group) => {
    const items = purchases.filter((p) => assignments[p.id] === group.key);
    const qty = items.reduce((sum, p) => sum + Number(p.qty), 0);
    const cents = items.reduce((sum, p) => sum + priceCents(p.price), 0);
    const categoryValid = items.every((p) => comparisonCategory(p, p.comparison_category) === group.category);
    return { ...group, items, qty, cents, difference: cents - priceCents(group.target || 0),
      valid: items.length > 0 && qty <= COMPARISON_ITEM_LIMIT && categoryValid };
  });
  const leftovers = purchases.filter((p) => !assignments[p.id]);
  return { groups: summaries, leftovers, included: purchases.length - leftovers.length,
    deviation: summaries.reduce((sum, g) => sum + Math.abs(g.difference), 0),
    valid: summaries.every((g) => g.valid) && purchases.every((p) => !assignments[p.id] || groupIds.has(assignments[p.id])) };
}

function buildModel(purchases, groups, timeout) {
  const model = { optimize: "included", opType: "max", constraints: {}, variables: {}, binaries: {}, options: { timeout, tolerance: 0 } };
  const edges = [];
  groups.forEach((group, j) => {
    model.constraints[`qty_${j}`] = { max: COMPARISON_ITEM_LIMIT };
    model.constraints[`nonempty_${j}`] = { min: 1 };
    model.constraints[`total_${j}`] = { equal: priceCents(group.target) };
    model.variables[`under_${j}`] = { [`total_${j}`]: 1, deviation: 1 };
    model.variables[`over_${j}`] = { [`total_${j}`]: -1, deviation: 1 };
  });
  purchases.forEach((purchase, i) => {
    model.constraints[`purchase_${i}`] = { max: 1 };
    if (Number(purchase.qty) > COMPARISON_ITEM_LIMIT) return;
    groups.forEach((group, j) => {
      if (comparisonCategory(purchase, purchase.comparison_category) !== group.category) return;
      const key = `x_${i}_${j}`;
      model.variables[key] = { included: 1, [`purchase_${i}`]: 1, [`qty_${j}`]: Number(purchase.qty),
        [`nonempty_${j}`]: 1, [`total_${j}`]: priceCents(purchase.price) };
      model.binaries[key] = 1;
      edges.push({ key, purchaseId: purchase.id, groupKey: group.key });
    });
  });
  return { model, edges };
}

export function generateComparisons(purchases, groups, { alternatives = 3, timeout = 1800 } = {}) {
  validateComparison(purchases, groups);
  const { model, edges } = buildModel(purchases, groups, timeout);
  const decode = (solution) => {
    if (!solution.feasible || edges.some(({ key }) => Math.abs((solution[key] || 0) - Math.round(solution[key] || 0)) > 0.00001)) return null;
    const assignments = {};
    for (const { key, purchaseId, groupKey } of edges) {
      if ((solution[key] || 0) > 0.5) {
        if (assignments[purchaseId]) return null;
        assignments[purchaseId] = groupKey;
      }
    }
    return summarizeComparison(purchases, groups, assignments).valid ? assignments : null;
  };
  const first = decode(solver.Solve(structuredClone(model)));
  if (!first) throw new Error("لم يتم العثور على تقسيم صالح. تحققي من تصنيفات الطلبات وعدد المشتريات وحد 100 قطعة.");
  const count = Object.keys(first).length;
  model.constraints.included = { equal: count };
  model.optimize = "deviation";
  model.opType = "min";
  const results = [];
  for (let n = 0; n < Math.max(1, Math.min(alternatives, 3)); n++) {
    const solution = decode(solver.Solve(structuredClone(model)));
    const assignments = solution || (n === 0 ? first : null);
    if (!assignments) break;
    results.push({ assignments, ...summarizeComparison(purchases, groups, assignments) });
    // Exclude the previous complete assignment, not just its price total.
    model.constraints[`alternative_${n}`] = { max: count - 1 };
    edges.forEach(({ key, purchaseId, groupKey }) => {
      if (assignments[purchaseId] === groupKey) model.variables[key][`alternative_${n}`] = 1;
    });
  }
  return results.sort((a, b) => b.included - a.included || a.deviation - b.deviation);
}

export function buildComparisonPayload(purchases, groups, assignments) {
  validateComparison(purchases, groups);
  const summary = summarizeComparison(purchases, groups, assignments);
  if (!summary.valid) throw new Error("كل طلب يحتاج مشتريات من نفس التصنيف وبحد أقصى 100 قطعة");
  const destinations = groups.filter((g) => g.mode === "existing").map((g) => g.orderId);
  if (destinations.some((id) => !id) || new Set(destinations).size !== destinations.length) throw new Error("اختاري طلباً فارغاً مختلفاً لكل وجهة");
  return summary.groups.map((group) => {
    if (group.mode !== "existing" && !String(group.name || "").trim()) throw new Error("أدخلي أسماء الطلبات الجديدة");
    return { category: group.category, target: Number(group.target), order_id: group.mode === "existing" ? group.orderId : null,
      name: String(group.name || "").trim(), items: group.items.map((p) => ({ id: p.id, price: p.price, qty: Number(p.qty), categories: p.purchase_categories || [] })) };
  });
}
