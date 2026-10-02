import { selectOrdersWithOptionalProfitFields } from "./orders";
import { sb } from "./supabaseClient";
import { FINANCE_START_AT, isFinanceOrderInPeriod } from "./financePeriod";

const PAGE_SIZE = 1000;
const ORDER_COLUMNS = "id, order_name, order_date, created_at, spent_amount";

function isMissingPostalFee(error) {
  return ["42703", "PGRST204"].includes(error?.code) &&
    String(error?.message || "").includes("postal_fee");
}

export function parseFinanceAmount(value) {
  const text = String(value ?? "").trim();
  const amount = text === "" ? 0 : Number(text);
  if (!Number.isFinite(amount) || amount < 0) {
    throw new Error("أدخلي مبلغاً صحيحاً لا يقل عن صفر.");
  }
  const cents = amount * 100;
  if (!Number.isSafeInteger(Math.round(cents)) || Math.abs(cents - Math.round(cents)) > 0.000001) {
    throw new Error("أدخلي المبلغ بمنزلتين عشريتين كحد أقصى.");
  }
  return Math.round(cents) / 100;
}

export function calculateOrderFinanceProfit(purchaseValue, postalFee, spent) {
  return (Math.round(purchaseValue * 100) + Math.round(postalFee * 100) - Math.round(spent * 100)) / 100;
}

export async function fetchFinanceOrders() {
  const orders = [];
  let postalFeeAvailable = true;

  for (let offset = 0; ; offset += PAGE_SIZE) {
    const range = { from: offset, to: offset + PAGE_SIZE - 1, createdFrom: FINANCE_START_AT };
    let result = await selectOrdersWithOptionalProfitFields(
      postalFeeAvailable ? `${ORDER_COLUMNS}, postal_fee` : ORDER_COLUMNS,
      range
    );
    if (isMissingPostalFee(result.error)) {
      postalFeeAvailable = false;
      result = await selectOrdersWithOptionalProfitFields(ORDER_COLUMNS, range);
    }
    if (result.error) throw result.error;
    orders.push(...(result.data || []));
    if ((result.data || []).length < PAGE_SIZE) break;
  }

  return { data: orders.filter(isFinanceOrderInPeriod), postalFeeAvailable };
}

export async function fetchFinancePurchases() {
  const purchases = [];
  for (let offset = 0; ; offset += PAGE_SIZE) {
    const { data, error } = await sb.from("purchases")
      .select("id, order_id, price, paid_price, pickup_point, collected, picked_up")
      .order("id", { ascending: true })
      .range(offset, offset + PAGE_SIZE - 1);
    if (error) throw error;
    purchases.push(...(data || []));
    if ((data || []).length < PAGE_SIZE) break;
  }
  return { data: purchases };
}

export async function saveOrderFinance(orderId, { spent, postalFee }) {
  const id = String(orderId || "").trim();
  if (!id) throw new Error("تعذر تحديد الطلبية للحفظ.");

  const payload = {
    spent_amount: parseFinanceAmount(spent),
    postal_fee: parseFinanceAmount(postalFee)
  };
  const { data, error } = await sb.from("orders")
    .update(payload)
    .eq("id", id)
    .select("id, spent_amount, postal_fee")
    .single();
  if (error) {
    if (isMissingPostalFee(error)) throw new Error("رسوم البريد غير متاحة حالياً. تعذر حفظ المالية.");
    throw error;
  }
  if (!data?.id) throw new Error("تعذر حفظ مالية الطلبية.");
  return data;
}
