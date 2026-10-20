export async function savePurchaseReceipt(client, purchaseId, pickedUp) {
  const { data, error } = await client.rpc("set_purchase_received", {
    p_purchase_id: purchaseId,
    p_picked_up: pickedUp
  });
  if (error) {
    if (["PGRST202", "42883"].includes(error.code)) {
      throw new Error("يلزم تشغيل تحديث صلاحيات الاستلام في Supabase أولاً.");
    }
    throw error;
  }
  if (!data || String(data.id) !== String(purchaseId) || data.picked_up !== pickedUp
    || (pickedUp && !data.picked_up_at) || (!pickedUp && data.picked_up_at !== null)
    || typeof data.changed !== "boolean") {
    throw new Error("لم يتم تأكيد حفظ حالة الاستلام. أعيدي المحاولة.");
  }
  return data;
}
