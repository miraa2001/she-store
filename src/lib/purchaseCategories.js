export const PURCHASE_CATEGORIES = [
  { value: "fast", label: "Fast" },
  { value: "makeup", label: "Makeup" },
  { value: "electronics", label: "Electronics" }
];

export function normalizePurchaseCategories(values) {
  return PURCHASE_CATEGORIES.filter(({ value }) => Array.isArray(values) && values.includes(value))
    .map(({ value }) => value);
}

export function filterPurchasesByCategory(purchases, category = "all") {
  if (category === "all") return purchases;
  return purchases.filter((purchase) => normalizePurchaseCategories(purchase.purchase_categories).includes(category));
}
