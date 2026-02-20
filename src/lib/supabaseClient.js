import { createClient } from "@supabase/supabase-js";

const DEFAULT_SUPABASE_URL = "https://maucjvxhrnkdeybltjco.supabase.co";
const DEFAULT_SUPABASE_ANON_KEY = "sb_publishable_0N65DXPMsk_3WVcLQ9wavQ_p9HgB_Bo";

function normalizeSupabaseUrl(raw) {
  const value = String(raw || "").trim();
  if (!value) return DEFAULT_SUPABASE_URL;
  if (/^https?:\/\//i.test(value)) return value;

  // Accept project-ref shorthand like: maucjvxhrnkdeybltjco
  if (/^[a-z0-9-]+$/i.test(value)) {
    return `https://${value}.supabase.co`;
  }

  return value;
}

function parseJwtPayload(token) {
  try {
    const [, payload] = String(token || "").split(".");
    if (!payload) return null;
    const normalized = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized + "=".repeat((4 - (normalized.length % 4 || 4)) % 4);
    const json = atob(padded);
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function isServiceRoleToken(token) {
  const payload = parseJwtPayload(token);
  return payload?.role === "service_role";
}

const envSupabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const envSupabaseKey = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "").trim();

const resolvedUrl = normalizeSupabaseUrl(envSupabaseUrl);
const resolvedAnonKey =
  envSupabaseKey && !isServiceRoleToken(envSupabaseKey)
    ? envSupabaseKey
    : DEFAULT_SUPABASE_ANON_KEY;

if (envSupabaseKey && isServiceRoleToken(envSupabaseKey)) {
  console.error(
    "[Supabase] VITE_SUPABASE_ANON_KEY contains a service_role token. Falling back to publishable key."
  );
}

let sb;

try {
  sb = createClient(resolvedUrl, resolvedAnonKey);
} catch (error) {
  console.error("[Supabase] Invalid config. Falling back to default project settings.", error);
  sb = createClient(DEFAULT_SUPABASE_URL, DEFAULT_SUPABASE_ANON_KEY);
}

export { sb };
