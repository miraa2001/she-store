// La Aura is intentionally hidden from the app UI while legacy database values stay supported here.
// To restore La Aura later, re-add a dedicated pickup config/route/navigation entry and unblock `laaura` in `src/lib/auth.js`.
export const PICKUP_HOME = "من البيت";
export const PICKUP_DELIVERY = "توصيل";
export const PICKUP_POINT = "من نقطة الاستلام";
export const PICKUP_POINT_NABLUS = "نقطة استلام نابلس";
export const PICKUP_POINT_LAAURA = `${PICKUP_POINT} - La Aura`;
export const PICKUP_POINT_MARYAMTI = `${PICKUP_POINT} - مريمتي`;
export const DEFAULT_PICKUP_OPTION = PICKUP_POINT;
export const NABLUS_CITY = "نابلس";

const LEGACY_LAAURA_ALIASES = [
  PICKUP_POINT_LAAURA,
  "La Aura",
  "la aura",
  "LAAURA",
  "لا اورا",
  "لا أورا",
  "لاورا"
];

const PICKUP_LOCATION_CONFIGS = {
  maryamti: {
    id: "maryamti",
    role: "maryamti",
    tabId: "pickup",
    routePath: "/pickuppoint",
    routeHash: "#/pickuppoint",
    navId: "pickuppoint",
    navLabel: "مريمتي",
    dashboardLabel: "مريمتي",
    pageTitle: "مريمتي",
    pageSubtitle: "طلبات الاستلام في مريمتي",
    pickupValue: PICKUP_POINT,
    pickupLabel: "مريمتي",
    whatsappLocationLine: "مريمتي - مجمع ابو طريف والزغل الطابق الثاني",
    whatsappHoursLine: "",
    aliases: [
      PICKUP_POINT,
      PICKUP_POINT_MARYAMTI,
      "Maryamti",
      "maryamti",
      "مريمتي",
      ...LEGACY_LAAURA_ALIASES
    ]
  },
  nablus: {
    id: "nablus",
    role: "nablus",
    tabId: "nablus",
    routePath: "/pickuppoint-nablus",
    routeHash: "#/pickuppoint-nablus",
    navId: "pickuppoint-nablus",
    navLabel: "الشخشير للأدوات المنزلية",
    dashboardLabel: "الشخشير للأدوات المنزلية",
    pageTitle: "الشخشير للأدوات المنزلية",
    pageSubtitle: "طلبات الاستلام في الشخشير للأدوات المنزلية",
    pickupValue: PICKUP_POINT_NABLUS,
    pickupLabel: "الشخشير للأدوات المنزلية",
    whatsappLocationLine: "الشخشير للأدوات المنزلية - وسط البلد - شارع غرناطة مقابل ال KFC بجانب عالم اللحوم - نابلس",
    whatsappHoursLine: "[ساعات الدوام]  : 9 صباحا - 6 مساءً",
    aliases: [PICKUP_POINT_NABLUS, "نابلس", "nablus", "الشخشير", "الشخشير للأدوات المنزلية"]
  }
};

export const PICKUP_POINT_LOCATIONS = Object.values(PICKUP_LOCATION_CONFIGS);
export const PICKUP_POINT_ROLE_IDS = PICKUP_POINT_LOCATIONS.map((location) => location.role);
export const CUSTOMER_PICKUP_OPTIONS = [PICKUP_HOME, PICKUP_POINT, PICKUP_POINT_NABLUS];

export function normalizePickup(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[أإآ]/g, "ا");
}

function matchesAliases(value, aliases = []) {
  const normalized = normalizePickup(value);
  if (!normalized) return false;

  return aliases.some((alias) => {
    const normalizedAlias = normalizePickup(alias);
    return normalizedAlias && (normalized === normalizedAlias || normalized.includes(normalizedAlias));
  });
}

function matchesLocationPickup(value, location) {
  const normalized = normalizePickup(value);
  if (!normalized || normalized.includes("بيت") || normalized.includes("توصيل")) return false;
  return matchesAliases(value, location?.aliases || []);
}

export function getPickupLocationById(locationId = "maryamti") {
  return PICKUP_LOCATION_CONFIGS[locationId] || PICKUP_LOCATION_CONFIGS.maryamti;
}

export function getPickupLocationByRole(role) {
  return PICKUP_LOCATION_CONFIGS[String(role || "").trim().toLowerCase()] || null;
}

export function getPickupLocationByTabId(tabId) {
  return PICKUP_POINT_LOCATIONS.find((location) => location.tabId === tabId) || null;
}

export function getPickupLocationByPickupPoint(value) {
  return PICKUP_POINT_LOCATIONS.find((location) => matchesLocationPickup(value, location)) || null;
}

export function isPickupPointRole(role) {
  return Boolean(getPickupLocationByRole(role));
}

export function getPickupRouteHashForRole(role) {
  return getPickupLocationByRole(role)?.routeHash || "#/orders";
}

export function getPickupRoutePathForRole(role) {
  return getPickupLocationByRole(role)?.routePath || "/orders";
}

export function isNablusCity(value) {
  return normalizePickup(value) === normalizePickup(NABLUS_CITY);
}

export function getDefaultPickupForCity(city, fallback = DEFAULT_PICKUP_OPTION) {
  if (isNablusCity(city)) return PICKUP_POINT_NABLUS;
  return fallback;
}

export function formatPickupDisplayLabel(value, fallback = "—") {
  const text = String(value || "").trim();
  if (!text) return fallback;
  const pickupLocation = getPickupLocationByPickupPoint(text);
  if (pickupLocation) return pickupLocation.pickupValue;
  if (normalizePickup(text) === normalizePickup(PICKUP_HOME)) return PICKUP_HOME;
  return text;
}

export function formatPickupFormValue(value, fallback = DEFAULT_PICKUP_OPTION) {
  const text = String(value || "").trim();
  if (!text) return fallback;
  const pickupLocation = getPickupLocationByPickupPoint(text);
  if (pickupLocation) return pickupLocation.pickupValue;
  if (normalizePickup(text) === normalizePickup(PICKUP_HOME)) return PICKUP_HOME;
  if (normalizePickup(text) === normalizePickup(PICKUP_DELIVERY)) return PICKUP_DELIVERY;
  return text;
}

export function getPreferredPickupForCustomer(customer, fallback = DEFAULT_PICKUP_OPTION) {
  const explicitPickup = formatPickupFormValue(customer?.usual_pickup_point, "");
  if (explicitPickup) return explicitPickup;
  return getDefaultPickupForCity(customer?.city, fallback);
}

export function getPickupOptionsWithCurrentValue(currentValue) {
  const options = [...CUSTOMER_PICKUP_OPTIONS];
  const displayValue = formatPickupFormValue(currentValue, "");
  if (!displayValue) return options;

  const exists = options.some((option) => normalizePickup(option) === normalizePickup(displayValue));
  if (!exists) options.push(displayValue);
  return options;
}

export function isLaauraPickup(value) {
  return matchesAliases(value, LEGACY_LAAURA_ALIASES);
}

export function isMaryamtiPickup(value) {
  return matchesLocationPickup(value, PICKUP_LOCATION_CONFIGS.maryamti);
}

export function isNablusPickup(value) {
  return matchesLocationPickup(value, PICKUP_LOCATION_CONFIGS.nablus);
}

export function isPickupPointPickup(value) {
  return Boolean(getPickupLocationByPickupPoint(value));
}

export function isPickupPointForLocation(value, locationId) {
  const location = getPickupLocationById(locationId);
  return matchesLocationPickup(value, location);
}

export const isAuraPickup = isLaauraPickup;
