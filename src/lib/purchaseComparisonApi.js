import { sb } from "./supabaseClient";

export async function fetchEmptyComparisonOrders() {
  const rows = [];
  for (let from = 0; ; from += 500) {
    const { data, error } = await sb.from("orders")
      .select("id,order_name,order_type,shein_type,arrived,placed_at_pickup,purchases(count)")
      .eq("arrived", false).eq("placed_at_pickup", false).eq("order_type", "shein")
      .order("created_at", { ascending: false }).order("id", { ascending: true }).range(from, from + 499);
    if (error) throw error;
    rows.push(...(data || []).filter((order) => Number(order.purchases?.[0]?.count) === 0));
    if ((data || []).length < 500) break;
  }
  return rows;
}

export async function confirmPurchaseComparison(requestId, groups) {
  const { data, error } = await sb.rpc("apply_purchase_comparison", { p_request_id: requestId, p_groups: groups });
  if (error) {
    if (["PGRST202", "42883", "42P01"].includes(error.code)) {
      const failure = new Error("يلزم تشغيل تحديث قاعدة بيانات المقارنة أولاً");
      failure.code = error.code;
      throw failure;
    }
    throw error;
  }
  if (!data?.orders?.length || !data.purchase_count) throw new Error("تعذر التأكد من حفظ التقسيم؛ أعيدي المحاولة");
  return data;
}
