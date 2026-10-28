import { PURCHASE_CATEGORIES, normalizePurchaseCategories } from "../../lib/purchaseCategories";

export default function PurchaseCategoryLabels({ categories }) {
  const selected = normalizePurchaseCategories(categories);
  if (!selected.length) return null;
  return <span className="purchase-category-labels" aria-label="تصنيفات المشترى">
    {PURCHASE_CATEGORIES.filter(({ value }) => selected.includes(value)).map(({ value, label }) => (
      <span key={value} className={`purchase-category-label purchase-category-${value}`}>{label}</span>
    ))}
  </span>;
}
