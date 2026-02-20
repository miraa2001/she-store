import { sb } from "./supabaseClient";
import { CUSTOMER_PICKUP_OPTIONS } from "./pickup";

export const CUSTOMER_CITIES = [
  "طولكرم",
  "نابلس",
  "جنين",
  "رام الله",
  "سلفيت",
  "قلقيلية",
  "أريحا",
  "بيت لحم",
  "القدس",
  "الداخل"
];

export { CUSTOMER_PICKUP_OPTIONS };

export function normalizePhone(value) {
  const converted = String(value || "")
    .replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)));

  return converted.replace(/[^\d]/g, "");
}

export function isValidCustomerPhone(value) {
  const phone = normalizePhone(value);
  return /^(970|972)\d{7,9}$/.test(phone);
}

export function customerFriendlyError(error, fallback = "حصل خطأ غير متوقع.") {
  const message = String(error?.message || "");
  const code = String(error?.code || "");

  if (
    code === "23505" ||
    message.includes("duplicate key value") ||
    message.toLowerCase().includes("unique constraint")
  ) {
    return "هذا الزبون موجود مسبقًا. تأكدي من الاسم أو رقم الهاتف.";
  }

  return fallback || message || "حصل خطأ غير متوقع.";
}

export async function fetchCustomers() {
  const { data, error } = await sb
    .from("customers")
    .select("id,name,phone,city,usual_pickup_point")
    .order("created_at", { ascending: false });

  if (error) throw error;
  return data || [];
}

export async function createCustomer(payload) {
  const { data, error } = await sb
    .from("customers")
    .insert(payload)
    .select("id,name,phone,city,usual_pickup_point")
    .single();
  if (error) throw error;
  return data;
}

export async function updateCustomer(id, payload) {
  const { error } = await sb.from("customers").update(payload).eq("id", id);
  if (error) throw error;
}

export async function deleteCustomer(id) {
  const { error } = await sb.from("customers").delete().eq("id", id);
  if (error) throw error;
}
