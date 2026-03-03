import { sb } from "./supabaseClient";
import { formatILS } from "./orders";
import { PICKUP_HOME, PICKUP_POINT } from "./pickup";

const EMOJI = {
  sparkleHeart: "\u{1F496}",
  package: "\u{1F4E6}",
  pin: "\u{1F4CD}",
  alarm: "\u23F0",
  kiss: "\u{1F48C}",
  heart: "\u2764\uFE0F",
  bell: "\u{1F514}",
  question: "\u2753"
};

function normalizePhone(value) {
  const converted = String(value || "")
    .replace(/[٠-٩]/g, (d) => String("٠١٢٣٤٥٦٧٨٩".indexOf(d)))
    .replace(/[۰-۹]/g, (d) => String("۰۱۲۳۴۵۶۷۸۹".indexOf(d)));

  return converted.replace(/[^\d]/g, "");
}

export function toWhatsappPhone(value) {
  let phone = normalizePhone(value);
  if (!phone) return "";
  if (phone.startsWith("00")) phone = phone.slice(2);
  if (phone.startsWith("970") || phone.startsWith("972")) return phone;

  if (phone.startsWith("0")) {
    const local = phone.slice(1);
    if (local.startsWith("59") || local.startsWith("56")) return `970${local}`;
    if (local.startsWith("5")) return `972${local}`;
  }

  return phone;
}

export function isValidWhatsappPhone(value) {
  const phone = toWhatsappPhone(value);
  return /^(970|972)\d{8,9}$/.test(phone);
}

export function buildWhatsappUrl(phone, message) {
  const safePhone = toWhatsappPhone(phone);
  const safeText = String(message || "").normalize("NFC");
  return `https://wa.me/${safePhone}?text=${encodeURIComponent(safeText)}`;
}

function extractFirstName(value) {
  const fullName = String(value || "").trim();
  if (!fullName) return "";
  return fullName.split(/\s+/)[0] || "";
}

export function buildArrivalNotifyMessage({ pickupPoint, price, customerName }) {
  const priceText = formatILS(price);
  const name = extractFirstName(customerName);
  const greeting = name ? `مرحباً ${name} ${EMOJI.sparkleHeart}` : `مرحباً حبيبتي ${EMOJI.sparkleHeart}`;

  if (pickupPoint === PICKUP_POINT) {
    return [
      greeting,
      `${EMOJI.bell} طلبك جاهز في نقطة الاستلام`,
      `${EMOJI.pin} كافيه La Aura - سوق الذهب`,
      `${EMOJI.alarm} أوقات العمل: من ٨ صباحاً حتى ١٠ مساءً`,
      `${EMOJI.package} حسابك: ${priceText} شيكل`
    ].join("\n");
  }

  if (pickupPoint === PICKUP_HOME) {
    return [
      greeting,
      `${EMOJI.bell} طلبك جاهز عندي في البيت`,
      `خبريني قبل بوقت متى رح تستلمي ${EMOJI.kiss}`,
      `${EMOJI.package} حسابك: ${priceText} شيكل`
    ].join("\n");
  }

  return [
    greeting,
    `${EMOJI.bell} طلبك جاهز للاستلام (${pickupPoint || "—"})`,
    `${EMOJI.package} حسابك: ${priceText} شيكل`
  ].join("\n");
}

export function buildPickupInquiryMessage() {
  return [
    `مرحباً حبيبتي، طلبك وصل ${EMOJI.package}`,
    "بتحبي تستلمي من:",
    `${EMOJI.pin} بيتي في الحي الجنوبي`,
    "أو",
    `${EMOJI.pin} نقطة الاستلام كافيه La Aura - سوق الذهب ${EMOJI.question}`,
    `خبريني لو سمحتي ${EMOJI.heart}`
  ].join("\n");
}

export async function resolvePurchaseWhatsappTarget(purchase) {
  if (!purchase?.customer_id) {
    throw new Error("لا يوجد رقم هاتف لهذا الزبون.");
  }

  const { data: customer, error } = await sb
    .from("customers")
    .select("phone, name")
    .eq("id", purchase.customer_id)
    .maybeSingle();

  if (error || !customer?.phone) {
    throw new Error("لا يوجد رقم هاتف لهذا الزبون.");
  }

  const phone = toWhatsappPhone(customer.phone);
  if (!isValidWhatsappPhone(phone)) {
    throw new Error("رقم الهاتف غير صالح.");
  }

  return {
    phone,
    customerName: String(customer.name || purchase.customer_name || "").trim()
  };
}
