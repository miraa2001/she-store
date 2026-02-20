import { formatILS } from "./orders";

const NTFY_TOPIC = "she-store-rahaf-2001-2014";

export function buildPickupStatusMessage({ picked, customerName, price, pickupLabel }) {
  const status = picked ? "✅" : "❌";
  return [
    `تم استلام الطلب ${status}`,
    `الزبون: ${customerName || ""}`,
    `نقطة الاستلام: ${pickupLabel || ""}`,
    `السعر: ${formatILS(price)}`
  ].join("\n");
}

export function buildCollectedMoneyMessage({ pickupLabel, amountText }) {
  return [
    "تم استلام تحصيل النقود ✅",
    `المكان: ${pickupLabel}`,
    `المبلغ: ${amountText}`
  ].join("\n");
}

export function notifyPickupStatus(message) {
  return fetch(`https://ntfy.sh/${NTFY_TOPIC}`, {
    method: "POST",
    body: message
  }).catch((error) => {
    console.error("NTFY ERROR:", error);
  });
}

