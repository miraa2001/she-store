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

const EMPTY_EMOJI = {
  sparkleHeart: "",
  package: "",
  pin: "",
  alarm: "",
  kiss: "",
  heart: "",
  bell: "",
  question: ""
};

function shouldUseWhatsappEmoji() {
  if (typeof navigator === "undefined") return true;
  const ua = String(navigator.userAgent || "");
  const isIOS = /iPhone|iPad|iPod/i.test(ua);
  if (!isIOS) return true;

  const iosMatch = ua.match(/OS (\d+)_/i);
  const iosMajor = iosMatch ? Number(iosMatch[1]) : 0;
  return Number.isFinite(iosMajor) ? iosMajor < 18 : true;
}

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
  const emoji = shouldUseWhatsappEmoji() ? EMOJI : EMPTY_EMOJI;
  const greeting = name ? `مرحباً ${name} ${emoji.sparkleHeart}` : `مرحباً حبيبتي ${emoji.sparkleHeart}`;

  if (pickupPoint === PICKUP_POINT) {
    return [
      greeting,
      `${emoji.bell} طلبك جاهز في نقطة الاستلام`,
      `${emoji.pin} كافيه La Aura - سوق الذهب`,
      `${emoji.alarm} أوقات العمل: من ٨ صباحاً حتى ١٠ مساءً`,
      `${emoji.package} حسابك: ${priceText} شيكل`
    ].join("\n");
  }

  if (pickupPoint === PICKUP_HOME) {
    return [
      greeting,
      `${emoji.bell} طلبك جاهز عندي في البيت`,
      `خبريني قبل بوقت متى رح تستلمي ${emoji.kiss}`,
      `${emoji.package} حسابك: ${priceText} شيكل`
    ].join("\n");
  }

  return [
    greeting,
    `${emoji.bell} طلبك جاهز للاستلام (${pickupPoint || "—"})`,
    `${emoji.package} حسابك: ${priceText} شيكل`
  ].join("\n");
}

export function buildPickupInquiryMessage() {
  const emoji = shouldUseWhatsappEmoji() ? EMOJI : EMPTY_EMOJI;
  return [
    `مرحباً حبيبتي، طلبك وصل ${emoji.package}`,
    "بتحبي تستلمي من:",
    `${emoji.pin} بيتي في الحي الجنوبي`,
    "أو",
    `${emoji.pin} نقطة الاستلام كافيه La Aura - سوق الذهب ${emoji.question}`,
    `خبريني لو سمحتي ${emoji.heart}`
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
