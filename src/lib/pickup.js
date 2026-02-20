export const PICKUP_HOME = "من البيت";
export const PICKUP_DELIVERY = "توصيل";
export const PICKUP_POINT = "من نقطة الاستلام";

export const CUSTOMER_PICKUP_OPTIONS = [PICKUP_HOME, PICKUP_DELIVERY, PICKUP_POINT];

const PICKUP_ALIASES = [
  PICKUP_POINT,
  "من نقطه الاستلام",
  "La Aura",
  "la aura",
  "LAAURA",
  "لا اورا",
  "لا أورا",
  "لاورا"
];

export function normalizePickup(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[أإآ]/g, "ا");
}

export function isAuraPickup(value) {
  const normalized = normalizePickup(value || PICKUP_POINT);
  if (!normalized) return false;
  if (normalized.includes("بيت") || normalized.includes("توصيل")) return false;

  if (
    PICKUP_ALIASES.some((alias) => {
      const normalizedAlias = normalizePickup(alias);
      return normalizedAlias && (normalized === normalizedAlias || normalized.includes(normalizedAlias));
    })
  ) {
    return true;
  }

  if (normalized.includes("aura")) return true;
  if (normalized.includes("لاورا") || normalized.includes("لااورا")) return true;
  if (normalized.includes("نقطه") || normalized.includes("نقطة") || normalized.includes("استلام")) return true;
  return false;
}

