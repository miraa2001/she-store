import { parseFinanceAmount } from "./finance";
import { parsePrice } from "./orders";
import { sb } from "./supabaseClient";

export const HOME_CASH_PAGE_SIZE = 25;
export const HOME_CASH_EXPENSE_CATEGORIES = [
  { value: "bags", label: "أكياس" },
  { value: "pins", label: "دبابيس" },
  { value: "name_stickers", label: "ستكرات أسماء الزبائن" },
  { value: "postal_delivery", label: "توصيل البريد" },
  { value: "atm_deposit", label: "وضع مصاري في الصراف" }
];
export const HOME_CASH_SOURCE_LABELS = { home: "البيت", maryamti: "مريمتي", nablus: "نابلس", delivery: "توصيل" };

function cashError(error) {
  if (["PGRST202", "42P01", "42883"].includes(error?.code)) {
    return new Error("تتبع مصاري البيت غير متاح حالياً.");
  }
  return error;
}

export async function fetchHomeCash(offset = 0) {
  const [summaryResult, activityResult] = await Promise.all([
    sb.rpc("get_home_cash_summary"),
    sb.rpc("get_home_cash_activity", { p_offset: offset, p_limit: HOME_CASH_PAGE_SIZE + 1 })
  ]);
  if (summaryResult.error) throw cashError(summaryResult.error);
  if (activityResult.error) throw cashError(activityResult.error);
  if (!summaryResult.data) throw new Error("تتبع مصاري البيت غير متاح حالياً.");
  const rows = activityResult.data || [];
  return {
    summary: {
      ...summaryResult.data,
      balance: parsePrice(summaryResult.data.balance),
      income: parsePrice(summaryResult.data.income),
      expenses: parsePrice(summaryResult.data.expenses)
    },
    activity: rows.slice(0, HOME_CASH_PAGE_SIZE),
    hasMore: rows.length > HOME_CASH_PAGE_SIZE
  };
}

export async function recordHomeCashExpense({ requestId, category, amount, quantity = 1, note }) {
  const total = parseFinanceAmount(amount);
  const count = Number(quantity);
  if (total <= 0) throw new Error("أدخلي مبلغاً أكبر من صفر.");
  if (!Number.isInteger(count) || count < 1 || count > 2147483647) throw new Error("أدخلي كمية صحيحة لا تقل عن واحد.");
  if (!HOME_CASH_EXPENSE_CATEGORIES.some((item) => item.value === category)) throw new Error("اختاري نوع المصروف.");
  const { data, error } = await sb.rpc("record_home_cash_expense", {
    p_request_id: requestId,
    p_category: category,
    p_amount: total,
    p_quantity: count,
    p_note: String(note || "").trim() || null
  });
  if (error) throw cashError(error);
  return data;
}

export async function cancelHomeCashExpense(expenseId) {
  const { data, error } = await sb.rpc("cancel_home_cash_expense", { p_expense_id: expenseId });
  if (error) throw cashError(error);
  return data;
}
