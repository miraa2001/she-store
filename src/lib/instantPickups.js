import { parseFinanceAmount } from "./finance";
import { PICKUP_DELIVERY, PICKUP_HOME, PICKUP_POINT, PICKUP_POINT_NABLUS } from "./pickup";
import { sb } from "./supabaseClient";

export const INSTANT_PICKUP_LOCATIONS = [
  { value: PICKUP_HOME, label: "البيت" },
  { value: PICKUP_POINT, label: "مريمتي" },
  { value: PICKUP_POINT_NABLUS, label: "نابلس" },
  { value: PICKUP_DELIVERY, label: "توصيل" }
];
export const INSTANT_PICKUPS_CHANGED = "instant-pickups-changed";

function instantError(error) {
  if (["PGRST202", "42P01", "42883"].includes(error?.code)) {
    return new Error("الاستلام الفوري غير متاح حالياً.");
  }
  return error;
}

function notifyChanged() {
  window.dispatchEvent(new Event(INSTANT_PICKUPS_CHANGED));
}

export async function fetchInstantPickups(pickupPoint, collected = false) {
  const rows = [];
  for (let offset = 0; ; offset += 500) {
    const { data, error } = await sb.from("instant_pickups")
      .select("id, customer_name, price, pickup_point, picked_up, picked_up_at, collected, collected_at, created_at")
      .eq("pickup_point", pickupPoint).eq("collected", collected)
      .order("created_at", { ascending: false }).order("id", { ascending: false })
      .range(offset, offset + 499);
    if (error) throw instantError(error);
    rows.push(...(data || []));
    if (!data || data.length < 500) return rows;
  }
}

export async function createInstantPickup({ requestId, customerName, price, pickupPoint }) {
  const name = String(customerName || "").trim();
  if (!name || name.length > 200) throw new Error("أدخلي اسم الزبون.");
  if (String(price).trim() === "") throw new Error("أدخلي السعر.");
  const amount = parseFinanceAmount(price);
  if (!INSTANT_PICKUP_LOCATIONS.some((location) => location.value === pickupPoint)) throw new Error("اختاري نقطة الاستلام.");
  const { data, error } = await sb.rpc("create_instant_pickup", {
    p_request_id: requestId, p_customer_name: name, p_price: amount, p_pickup_point: pickupPoint
  });
  if (error) throw instantError(error);
  notifyChanged();
  return data;
}

export async function updateInstantPickup(id, action, pickupPoint = null) {
  const { data, error } = await sb.rpc("update_instant_pickup", { p_id: id, p_action: action, p_pickup_point: pickupPoint });
  if (error) throw instantError(error);
  notifyChanged();
  return data;
}

export async function collectInstantPickups(ids) {
  const { data, error } = await sb.rpc("collect_instant_pickups", { p_ids: ids });
  if (error) throw instantError(error);
  notifyChanged();
  return data;
}

export async function deleteInstantPickup(id) {
  if (!id) throw new Error("تعذر تحديد المستلم للحذف.");
  const { data, error } = await sb.rpc("delete_instant_pickup", { p_id: id });
  if (error) {
    if (["PGRST202", "42883"].includes(error.code)) {
      throw new Error("الحذف غير متاح: يلزم تشغيل تحديث حذف الاستلام الفوري في قاعدة البيانات.");
    }
    throw instantError(error);
  }
  notifyChanged();
  return data;
}
