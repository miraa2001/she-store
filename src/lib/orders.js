import { sb } from "./supabaseClient";

export const IMAGE_BUCKET = "purchase-images";

export const ORDER_STATUS = {
  PENDING: "pending",
  ARRIVED: "arrived",
  AT_PICKUP: "at_pickup",
  COLLECTED: "collected"
};

export const ORDER_STATUS_LABELS = {
  [ORDER_STATUS.PENDING]: "\u0642\u064a\u062f \u0627\u0644\u0627\u0646\u062a\u0638\u0627\u0631",
  [ORDER_STATUS.ARRIVED]: "\u062a\u0645 \u0648\u0635\u0648\u0644 \u0627\u0644\u0637\u0644\u0628",
  [ORDER_STATUS.AT_PICKUP]: "\u0641\u064a \u0646\u0642\u0637\u0629 \u0627\u0644\u0627\u0633\u062a\u0644\u0627\u0645",
  [ORDER_STATUS.COLLECTED]: "\u062a\u0645 \u0627\u0644\u062a\u062d\u0635\u064a\u0644"
};

export function parsePrice(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value)
    .replace(/[, ]/g, "")
    .replace(/[\u20AA]/g, "");

  const match = cleaned.match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

export function formatILS(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return "0";
  const fixed = number.toFixed(2);
  return fixed.endsWith(".00") ? String(Math.round(number)) : fixed;
}

export function normalizeOrderStatus(rawStatus) {
  const status = String(rawStatus || "").trim().toLowerCase();
  if (status === ORDER_STATUS.ARRIVED) return ORDER_STATUS.ARRIVED;
  if (status === ORDER_STATUS.AT_PICKUP) return ORDER_STATUS.AT_PICKUP;
  if (status === ORDER_STATUS.COLLECTED) return ORDER_STATUS.COLLECTED;
  return ORDER_STATUS.PENDING;
}

export function isPurchaseFullyCollected(purchase) {
  // "Collected" is authoritative, even when price/paid fields are zero.
  return !!purchase?.collected;
}

export function isOrderFullyCollected(purchases = []) {
  if (!Array.isArray(purchases) || !purchases.length) return false;
  return purchases.every((purchase) => isPurchaseFullyCollected(purchase));
}

export function deriveOrderStatus({ arrived, placedAtPickup, purchaseCount = 0, allCollected = false }) {
  if (purchaseCount > 0 && allCollected) return ORDER_STATUS.COLLECTED;
  if (placedAtPickup) return ORDER_STATUS.AT_PICKUP;
  if (arrived) return ORDER_STATUS.ARRIVED;
  return ORDER_STATUS.PENDING;
}

function monthKey(iso) {
  if (!iso) return "\u2014";
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "\u2014";

  return date.toLocaleDateString("ar", {
    month: "long",
    year: "numeric"
  });
}

function shortOrderNo(id) {
  const text = String(id || "").replace(/-/g, "");
  if (!text) return "----";
  return text.slice(-4);
}

export async function fetchOrdersWithSummary() {
  const { data: orders, error } = await sb
    .from("orders")
    .select("id, order_name, created_at, arrived, placed_at_pickup")
    .order("created_at", { ascending: false });

  if (error) throw error;

  const normalizedOrders = (orders || []).map((order) => ({
    id: order.id,
    name: String(order.order_name || "").trim() || "\u0637\u0644\u0628 \u0628\u062f\u0648\u0646 \u0627\u0633\u0645",
    createdAt: order.created_at,
    arrived: !!order.arrived,
    placedAtPickup: !!order.placed_at_pickup,
    orderNo: shortOrderNo(order.id)
  }));

  if (!normalizedOrders.length) {
    return [];
  }

  const ids = normalizedOrders.map((order) => order.id);

  const { data: purchases, error: purchasesError } = await sb
    .from("purchases")
    .select("order_id, price, paid_price, collected")
    .in("order_id", ids);

  if (purchasesError) throw purchasesError;

  const totals = new Map();
  const purchaseCounts = new Map();
  const fullyCollectedByOrder = new Map();

  (purchases || []).forEach((purchase) => {
    const id = purchase.order_id;
    if (!id) return;

    const next = (totals.get(id) || 0) + parsePrice(purchase.price);
    totals.set(id, next);
    purchaseCounts.set(id, (purchaseCounts.get(id) || 0) + 1);

    const currentAllCollected = fullyCollectedByOrder.has(id) ? fullyCollectedByOrder.get(id) : true;
    fullyCollectedByOrder.set(id, currentAllCollected && isPurchaseFullyCollected(purchase));
  });

  return normalizedOrders.map((order) => {
    const total = totals.get(order.id) || 0;
    const purchaseCount = purchaseCounts.get(order.id) || 0;
    const allCollected = purchaseCount > 0 && fullyCollectedByOrder.get(order.id) === true;

    return {
      ...order,
      amountRaw: total,
      amountLabel: `${formatILS(total)} \u20AA`,
      purchaseCount,
      allCollected,
      status: deriveOrderStatus({
        arrived: order.arrived,
        placedAtPickup: order.placedAtPickup,
        purchaseCount,
        allCollected
      })
    };
  });
}

export function currentMonthStart(now = new Date()) {
  return new Date(now.getFullYear(), now.getMonth(), 1);
}

export function isOlderThanCurrentMonth(isoDate, now = new Date()) {
  if (!isoDate) return false;
  const parsed = new Date(isoDate);
  if (Number.isNaN(parsed.getTime())) return false;
  return parsed < currentMonthStart(now);
}

export async function fetchArrivedOrders() {
  const { data, error } = await sb
    .from("orders")
    .select("id, order_name, created_at, arrived, placed_at_pickup, placed_at_pickup_at")
    .eq("arrived", true)
    .order("created_at", { ascending: false });

  if (error) throw error;

  return (data || []).map((order) => ({
    id: order.id,
    name: String(order.order_name || "").trim() || "\u0637\u0644\u0628 \u0628\u062f\u0648\u0646 \u0627\u0633\u0645",
    createdAt: order.created_at,
    arrived: !!order.arrived,
    placedAtPickup: !!order.placed_at_pickup,
    placedAtPickupAt: order.placed_at_pickup_at || null,
    orderNo: shortOrderNo(order.id)
  }));
}

export async function updateOrderPlacedAtPickup(orderId, enabled) {
  const payload = enabled
    ? { placed_at_pickup: true, placed_at_pickup_at: new Date().toISOString() }
    : { placed_at_pickup: false, placed_at_pickup_at: null };

  const { error } = await sb.from("orders").update(payload).eq("id", orderId);
  if (error) throw error;
  return payload;
}

export async function updateOrderWorkflowStatus(orderId, nextStatus) {
  const status = normalizeOrderStatus(nextStatus);
  const nowIso = new Date().toISOString();
  let payload;

  if (status === ORDER_STATUS.PENDING) {
    payload = { arrived: false, placed_at_pickup: false, placed_at_pickup_at: null };
  } else if (status === ORDER_STATUS.ARRIVED) {
    payload = { arrived: true, placed_at_pickup: false, placed_at_pickup_at: null };
  } else {
    payload = { arrived: true, placed_at_pickup: true, placed_at_pickup_at: nowIso };
  }

  const { error } = await sb.from("orders").update(payload).eq("id", orderId);
  if (error) throw error;

  return {
    status,
    payload
  };
}

export async function createOrder(orderName) {
  const name = String(orderName || "").trim();
  if (!name) throw new Error("اسم الطلب مطلوب.");

  const { data, error } = await sb
    .from("orders")
    .insert({ order_name: name })
    .select("id, order_name")
    .single();

  if (error) throw error;

  return {
    id: data?.id,
    name: String(data?.order_name || name).trim() || name
  };
}

export async function updateOrderName(orderId, orderName) {
  const name = String(orderName || "").trim();
  if (!name) throw new Error("اسم الطلب مطلوب.");

  const { error } = await sb
    .from("orders")
    .update({ order_name: name })
    .eq("id", orderId);

  if (error) throw error;
  return { id: orderId, name };
}

export async function deleteOrderById(orderId) {
  const { error } = await sb.from("orders").delete().eq("id", orderId);
  if (error) throw error;
}

export function groupOrdersByMonth(orders, query = "") {
  const needle = String(query || "").trim().toLowerCase();

  const filtered = !needle
    ? orders
    : orders.filter((order) => order.name.toLowerCase().includes(needle));

  const grouped = [];
  const map = new Map();

  filtered.forEach((order) => {
    const key = monthKey(order.createdAt);
    if (!map.has(key)) {
      const item = { month: key, orders: [] };
      map.set(key, item);
      grouped.push(item);
    }
    map.get(key).orders.push(order);
  });

  return grouped;
}
