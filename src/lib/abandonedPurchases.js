import { sb } from "./supabaseClient";
import { parseFinanceAmount } from "./finance";
import { INSTANT_PICKUP_LOCATIONS, INSTANT_PICKUPS_CHANGED } from "./instantPickups";

export const ABANDONED_PURCHASES_CHANGED = "abandoned-purchases-changed";
export function canArchivePurchase(purchase) {
  return !!purchase.order_id && !purchase.abandoned_at && !purchase.picked_up && !purchase.collected
    && Number(purchase.received_amount || 0) === 0 && Number(purchase.remitted_amount || 0) === 0
    && !(purchase.purchase_payments || []).length;
}

function archiveError(error) {
  if (["PGRST202", "42P01", "42703", "PGRST204"].includes(error?.code)) return new Error("يلزم تشغيل تحديث أرشيف المشتريات المتروكة في Supabase أولاً.");
  return error;
}
function changed() {
  window.dispatchEvent(new Event(ABANDONED_PURCHASES_CHANGED));
  window.dispatchEvent(new Event(INSTANT_PICKUPS_CHANGED));
}
async function archiveRpc(name, args) {
  const { data, error } = await sb.rpc(name, args);
  if (error) throw archiveError(error);
  if (!data) throw new Error("لم يتم تأكيد حفظ العملية.");
  changed();
  return data;
}
export async function archiveAbandonedPurchases(requestId, purchaseIds) {
  return archiveRpc("archive_abandoned_purchases", { p_request_id: requestId, p_purchase_ids: purchaseIds });
}
export async function restoreAbandonedPurchases(requestId, archiveIds) {
  return archiveRpc("restore_abandoned_purchases", { p_request_id: requestId, p_archive_ids: archiveIds });
}
export async function resellAbandonedPurchase(archiveId, { requestId, customerName, price, pickupPoint }) {
  const amount = parseFinanceAmount(price);
  if (!String(customerName || "").trim() || String(price).trim() === "") throw new Error("أدخلي اسم الزبون والسعر.");
  if (!INSTANT_PICKUP_LOCATIONS.some(location => location.value === pickupPoint)) throw new Error("اختاري نقطة الاستلام.");
  return archiveRpc("resell_abandoned_purchase", { p_request_id: requestId, p_archive_id: archiveId,
    p_customer_name: String(customerName).trim(), p_price: amount, p_pickup_point: pickupPoint });
}
export async function fetchAbandonedPurchases() {
  const rows = [];
  for (let from = 0; ; from += 500) {
    const { data, error } = await sb.from("abandoned_purchase_archives")
      .select("*,resale:instant_pickups!instant_pickup_id(id,customer_name,price,pickup_point,picked_up,collected)")
      .order("archived_at", { ascending: false }).order("id", { ascending: false }).range(from, from + 499);
    if (error) throw archiveError(error);
    rows.push(...(data || []));
    if (!data || data.length < 500) return rows;
  }
}
