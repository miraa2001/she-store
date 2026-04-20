export const PICKUP_HOME = "من البيت";
export const PICKUP_DELIVERY = "توصيل";
export const PICKUP_POINT = "من نقطة الاستلام";
export const PICKUP_POINT_LAAURA = `${PICKUP_POINT} - La Aura`;
export const PICKUP_POINT_MARYAMTI = `${PICKUP_POINT} - مريمتي`;

const PICKUP_LOCATION_CONFIGS = {
  laaura: {
    id: "laaura",
    role: "laaura",
    tabId: "laaura",
    routePath: "/pickuppoint",
    routeHash: "#/pickuppoint",
    navId: "pickuppoint-laaura",
    navLabel: "نقطة La Aura",
    dashboardLabel: "La Aura",
    pageTitle: "نقطة الاستلام - La Aura",
    pageSubtitle: "طلبات الاستلام من نقطة الاستلام",
    pickupValue: PICKUP_POINT_LAAURA,
    pickupLabel: PICKUP_POINT_LAAURA,
    whatsappLocationLine: "كافيه La Aura - سوق الذهب",
    whatsappHoursLine: "أوقات العمل: من ٨ صباحاً حتى ١٠ مساءً",
    aliases: [
      PICKUP_POINT,
      PICKUP_POINT_LAAURA,
      "La Aura",
      "la aura",
      "LAAURA",
      "لا اورا",
      "لا أورا",
      "لاورا"
    ]
  },
  maryamti: {
    id: "maryamti",
    role: "maryamti",
    tabId: "maryamti",
    routePath: "/pickuppoint-maryamti",
    routeHash: "#/pickuppoint-maryamti",
    navId: "pickuppoint-maryamti",
    navLabel: "نقطة مريمتي",
    dashboardLabel: "مريمتي",
    pageTitle: "نقطة الاستلام - مريمتي",
    pageSubtitle: "طلبات الاستلام من نقطة الاستلام",
    pickupValue: PICKUP_POINT_MARYAMTI,
    pickupLabel: PICKUP_POINT_MARYAMTI,
    whatsappLocationLine: "مريمتي - مجمع ابو طريف والزغل الطابق الثاني",
    whatsappHoursLine: "",
    aliases: [PICKUP_POINT_MARYAMTI, "Maryamti", "maryamti", "مريمتي"]
  }
};

export const PICKUP_POINT_LOCATIONS = Object.values(PICKUP_LOCATION_CONFIGS);
export const PICKUP_POINT_ROLE_IDS = PICKUP_POINT_LOCATIONS.map((location) => location.role);
export const CUSTOMER_PICKUP_OPTIONS = [PICKUP_HOME, PICKUP_DELIVERY, ...PICKUP_POINT_LOCATIONS.map((location) => location.pickupValue)];

export function normalizePickup(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[أإآ]/g, "ا");
}

function matchesLocationPickup(value, location) {
  const normalized = normalizePickup(value);
  if (!normalized || normalized.includes("بيت") || normalized.includes("توصيل")) return false;

  return location.aliases.some((alias) => {
    const normalizedAlias = normalizePickup(alias);
    return normalizedAlias && (normalized === normalizedAlias || normalized.includes(normalizedAlias));
  });
}

export function getPickupLocationById(locationId = "laaura") {
  return PICKUP_LOCATION_CONFIGS[locationId] || PICKUP_LOCATION_CONFIGS.laaura;
}

export function getPickupLocationByRole(role) {
  return PICKUP_LOCATION_CONFIGS[String(role || "").trim().toLowerCase()] || null;
}

export function getPickupLocationByTabId(tabId) {
  return PICKUP_POINT_LOCATIONS.find((location) => location.tabId === tabId) || null;
}

export function getPickupLocationByPickupPoint(value) {
  if (matchesLocationPickup(value, PICKUP_LOCATION_CONFIGS.maryamti)) {
    return PICKUP_LOCATION_CONFIGS.maryamti;
  }
  if (matchesLocationPickup(value, PICKUP_LOCATION_CONFIGS.laaura)) {
    return PICKUP_LOCATION_CONFIGS.laaura;
  }
  return null;
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

export function isLaauraPickup(value) {
  return matchesLocationPickup(value, PICKUP_LOCATION_CONFIGS.laaura);
}

export function isMaryamtiPickup(value) {
  return matchesLocationPickup(value, PICKUP_LOCATION_CONFIGS.maryamti);
}

export function isPickupPointPickup(value) {
  return Boolean(getPickupLocationByPickupPoint(value));
}

export function isPickupPointForLocation(value, locationId) {
  const location = getPickupLocationById(locationId);
  return matchesLocationPickup(value, location);
}

export const isAuraPickup = isLaauraPickup;
