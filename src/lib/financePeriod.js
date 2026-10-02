export const FINANCE_START_AT = "2026-10-01T00:00:00+03:00";

export function isFinanceOrderInPeriod(order) {
  const created = Date.parse(order.created_at ?? order.createdAt);
  return Number.isFinite(created) && created >= Date.parse(FINANCE_START_AT);
}
