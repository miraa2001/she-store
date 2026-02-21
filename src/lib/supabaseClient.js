import { createClient } from "@supabase/supabase-js";

function normalizeSupabaseUrl(raw) {
  const value = String(raw || "").trim();
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;

  // Accept project-ref shorthand from CI secrets.
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

function assertPublicSupabaseConfig(url, key) {
  const errors = [];

  if (!url) {
    errors.push("VITE_SUPABASE_URL is missing.");
  } else {
    try {
      const parsed = new URL(url);
      if (!/^https?:$/.test(parsed.protocol)) {
        errors.push("VITE_SUPABASE_URL must be an HTTP/HTTPS URL.");
      }
    } catch {
      errors.push("VITE_SUPABASE_URL is invalid.");
    }
  }

  if (!key) {
    errors.push("VITE_SUPABASE_ANON_KEY is missing.");
  } else if (isServiceRoleToken(key)) {
    errors.push("VITE_SUPABASE_ANON_KEY must be a publishable/anon key, not service_role.");
  }

  if (errors.length) {
    throw new Error(`[Supabase config] ${errors.join(" ")}`);
  }
}

const SUPABASE_URL = normalizeSupabaseUrl(import.meta.env.VITE_SUPABASE_URL);
const SUPABASE_ANON_KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "").trim();

assertPublicSupabaseConfig(SUPABASE_URL, SUPABASE_ANON_KEY);

export const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
