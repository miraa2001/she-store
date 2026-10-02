import { parseFinanceAmount } from "./finance";
import { sb } from "./supabaseClient";

export const PROFIT_PARTIES = [
  { value: "home", label: "البيت" }, { value: "rahaf", label: "رهف" }, { value: "mira", label: "ميرا" }
];
export const PROFIT_PAGE_SIZE = 25;

export function calculateProfitShares(profit, percentages, deductions = [0, 0, 0]) {
  if (percentages.some((value) => value === null || value === undefined)) return [null, null, null];
  const cents = Math.max(0, Math.round(profit * 100));
  const shares = percentages.map((percent, index) => {
    const numerator = BigInt(cents) * BigInt(Math.round(percent * 100));
    return { index, base: Number(numerator / 10000n), remainder: Number(numerator % 10000n) };
  });
  const remaining = cents - shares.reduce((sum, share) => sum + share.base, 0);
  shares.sort((a, b) => b.remainder - a.remainder || a.index - b.index);
  const result = [];
  shares.forEach((share, rank) => {
    result[share.index] = Math.max(0, share.base + (rank < remaining ? 1 : 0) - Math.round((deductions[share.index] ?? 0) * 100)) / 100;
  });
  return result;
}

function profitError(error) {
  if (["PGRST202", "42P01", "42883"].includes(error?.code)) return new Error("يلزم تشغيل تحديث توزيع الأرباح في قاعدة البيانات أولاً.");
  return error;
}

export async function fetchProfitData(offset = 0) {
  const [balances, payouts] = await Promise.all([
    sb.rpc("get_profit_balances"),
    sb.from("profit_payouts").select("id, party, amount, deduct_home_cash, actor_email, created_at")
      .order("created_at", { ascending: false }).order("id", { ascending: false })
      .range(offset, offset + PROFIT_PAGE_SIZE)
  ]);
  if (balances.error) throw profitError(balances.error);
  if (payouts.error) throw profitError(payouts.error);
  const orders = [];
  for (let from = 0; ; from += 500) {
    const { data, error } = await sb.from("order_profit_accruals").select("*")
      .order("updated_at", { ascending: false }).order("order_id", { ascending: false }).range(from, from + 499);
    if (error) throw profitError(error);
    orders.push(...(data || []));
    if (!data || data.length < 500) break;
  }
  return { balances: balances.data || [], orders, payouts: (payouts.data || []).slice(0, PROFIT_PAGE_SIZE), hasMore: (payouts.data || []).length > PROFIT_PAGE_SIZE };
}

export async function recordProfitPayout({ requestId, party, amount, deductHomeCash }) {
  const value = parseFinanceAmount(amount);
  if (value <= 0) throw new Error("أدخلي مبلغاً أكبر من صفر.");
  const { data, error } = await sb.rpc("record_profit_payout", {
    p_request_id: requestId, p_party: party, p_amount: value, p_deduct_home_cash: deductHomeCash
  });
  if (error) throw profitError(error);
  return data;
}
