import { sb } from "./supabaseClient";

const EMAIL_ROLE_FALLBACK = {
  "rahaf@she-store.com": "rahaf",
  "reem@she-store.com": "reem",
  "rawand@she-store.com": "rawand",
  "laaura@she-store.com": "laaura"
};

function normalizeRole(rawRole, email = "") {
  const role = String(rawRole || "").trim().toLowerCase();
  const normalizedEmail = String(email || "").trim().toLowerCase();

  if (!role) {
    if (normalizedEmail.includes("rahaf")) return "rahaf";
    if (normalizedEmail.includes("rawand")) return "rawand";
    if (normalizedEmail.includes("reem")) return "reem";
    if (normalizedEmail.includes("laaura") || normalizedEmail.includes("la.aura") || normalizedEmail.includes("aura")) {
      return "laaura";
    }
    return "viewer";
  }

  if (role === "rahaf" || role === "owner" || role === "admin") return "rahaf";

  if (role === "laaura" || role === "pickup" || role === "la aura" || role === "aura") {
    return "laaura";
  }

  if (role === "rawand") return "rawand";
  if (role === "reem") return "reem";

  if (role === "viewer" || role === "view" || role === "readonly" || role === "read_only") {
    if (normalizedEmail.includes("rawand")) return "rawand";
    return "reem";
  }

  return role;
}

export async function getCurrentUserProfile() {
  const {
    data: { session }
  } = await sb.auth.getSession();

  if (!session) {
    return { authenticated: false, role: "viewer", user: null, email: "" };
  }

  const {
    data: { user }
  } = await sb.auth.getUser();

  const email = String(user?.email || "").toLowerCase();
  const userId = String(user?.id || "");

  if (!userId) {
    return {
      authenticated: true,
      role: normalizeRole(EMAIL_ROLE_FALLBACK[email], email),
      user,
      email
    };
  }

  const { data, error } = await sb
    .from("user_roles")
    .select("role")
    .eq("user_id", userId)
    .maybeSingle();

  if (error) {
    console.error("role read error", error);
  }

  const role = normalizeRole(data?.role || EMAIL_ROLE_FALLBACK[email], email);

  return {
    authenticated: true,
    role,
    user,
    email
  };
}
