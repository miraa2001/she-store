export const agreedAmount = (row) => Number(row.paid_price ?? row.price ?? 0);
export const paidAmount = (row) => Number(row.received_amount ?? 0);
export const remainingAmount = (row) => Math.max(0, Math.round((agreedAmount(row) - paidAmount(row)) * 100) / 100);
export const collectibleAmount = (row) => Math.max(0, Math.round((paidAmount(row) - Number(row.remitted_amount ?? 0)) * 100) / 100);
export const canRecordPayment = (role) => role === "rahaf" || role === "reem";
export const canConfirmStaffPickup = (role) => role === "maryamti" || role === "nablus";

export async function paymentRpc(client, name, args) {
  const { data, error } = await client.rpc(name, args);
  if (error) {
    if (["PGRST202", "42883", "42703"].includes(error.code)) {
      throw new Error(name === "confirm_pickup_staff_receipt"
        ? "يلزم تشغيل تحديث استلام نقاط مريمتي ونابلس في Supabase أولاً."
        : "يلزم تشغيل تحديث الدفعات الجزئية في Supabase أولاً.");
    }
    throw error;
  }
  if (data === null || data === undefined) throw new Error("لم يتم تأكيد حفظ العملية.");
  return data;
}

export async function recordPayment(client, { requestId, purchase, amount, instant = false, receive = false }) {
  if (!/^\d+(\.\d{1,2})?$/.test(String(amount).trim()) || Number(amount) <= 0) throw new Error("أدخلي مبلغاً أكبر من صفر بمنزلتين عشريتين.");
  const result = await paymentRpc(client, "record_purchase_payment", {
    p_request_id: requestId, p_id: purchase.id, p_amount: Number(amount), p_receive: receive, p_instant: instant
  });
  if (result.purchase?.id !== purchase.id || typeof result.changed_receipt !== "boolean") throw new Error("لم يتم تأكيد حفظ الدفعة.");
  return result;
}

export async function collectPayments(client, ids, instant = false) {
  const data = await paymentRpc(client, "collect_purchase_payments", { p_ids: ids, p_instant: instant });
  if (!Number.isFinite(Number(data)) || Number(data) < 0) throw new Error("لم يتم تأكيد التحصيل.");
  return Number(data);
}

export async function confirmStaffPickup(client, { requestId, purchase, instant = false }) {
  const result = await paymentRpc(client, "confirm_pickup_staff_receipt", {
    p_request_id: requestId, p_id: purchase.id, p_instant: instant
  });
  if (result.purchase?.id !== purchase.id || result.purchase.picked_up !== true
    || !result.purchase.picked_up_at || typeof result.changed_receipt !== "boolean"
    || remainingAmount(result.purchase) !== 0) {
    throw new Error("لم يتم تأكيد حفظ الاستلام والدفع الكامل.");
  }
  return result;
}
