import { useEffect, useRef } from "react";
import { useNavigate } from "react-router-dom";
import { sb } from "../lib/supabaseClient";

const SESSION_TIMEOUT_MS = 15 * 60 * 1000;
const ACTIVITY_STORAGE_KEY = "she_store:last_activity_at";
const REMEMBER_ME_STORAGE_KEY = "she_store:remember_me";
const ACTIVITY_EVENTS = [
  "pointerdown",
  "keydown",
  "touchstart",
  "wheel",
  "scroll"
];

function readLastActivity() {
  const raw = Number(window.localStorage.getItem(ACTIVITY_STORAGE_KEY) || 0);
  return Number.isFinite(raw) && raw > 0 ? raw : Date.now();
}

function writeLastActivity() {
  window.localStorage.setItem(ACTIVITY_STORAGE_KEY, String(Date.now()));
}

function isRememberMeEnabled() {
  return window.localStorage.getItem(REMEMBER_ME_STORAGE_KEY) === "1";
}

export default function SessionTimeoutGuard() {
  const navigate = useNavigate();
  const timerRef = useRef(null);
  const hasSessionRef = useRef(false);
  const signingOutRef = useRef(false);

  const clearTimer = () => {
    if (timerRef.current) {
      window.clearTimeout(timerRef.current);
      timerRef.current = null;
    }
  };

  const timeoutSignOut = async () => {
    if (signingOutRef.current || !hasSessionRef.current) return;
    signingOutRef.current = true;
    clearTimer();
    hasSessionRef.current = false;

    try {
      await sb.auth.signOut();
    } catch (error) {
      console.error(error);
    } finally {
      window.localStorage.removeItem(ACTIVITY_STORAGE_KEY);
      const currentHash = window.location.hash || "#/";
      const currentRoute = currentHash.startsWith("#") ? currentHash.slice(1) : currentHash;
      const nextTarget = currentRoute.startsWith("/login")
        ? "/login"
        : `/login?next=${encodeURIComponent(currentRoute || "/")}`;
      navigate(nextTarget, { replace: true });
      signingOutRef.current = false;
    }
  };

  const scheduleTimeout = () => {
    clearTimer();
    if (!hasSessionRef.current) return;
    if (isRememberMeEnabled()) return;

    const elapsed = Date.now() - readLastActivity();
    const remaining = SESSION_TIMEOUT_MS - elapsed;

    if (remaining <= 0) {
      timeoutSignOut();
      return;
    }

    timerRef.current = window.setTimeout(() => {
      timeoutSignOut();
    }, remaining);
  };

  useEffect(() => {
    let mounted = true;

    const onActivity = () => {
      if (!hasSessionRef.current) return;
      writeLastActivity();
      scheduleTimeout();
    };

    const onStorage = (event) => {
      if (event.key !== ACTIVITY_STORAGE_KEY && event.key !== REMEMBER_ME_STORAGE_KEY) return;
      if (!hasSessionRef.current) return;
      scheduleTimeout();
    };

    async function init() {
      try {
        const {
          data: { session }
        } = await sb.auth.getSession();
        if (!mounted) return;
        hasSessionRef.current = !!session;
        if (session) {
          if (!window.localStorage.getItem(ACTIVITY_STORAGE_KEY) && !isRememberMeEnabled()) {
            writeLastActivity();
          }
          scheduleTimeout();
        }
      } catch (error) {
        console.error(error);
      }
    }

    init();

    const { data: authSubscription } = sb.auth.onAuthStateChange((_event, session) => {
      hasSessionRef.current = !!session;

      if (session) {
        if (!isRememberMeEnabled()) writeLastActivity();
        scheduleTimeout();
        return;
      }

      clearTimer();
      window.localStorage.removeItem(ACTIVITY_STORAGE_KEY);
    });

    ACTIVITY_EVENTS.forEach((eventName) => {
      window.addEventListener(eventName, onActivity, { passive: true });
    });
    window.addEventListener("storage", onStorage);

    return () => {
      mounted = false;
      clearTimer();
      authSubscription?.subscription?.unsubscribe?.();
      ACTIVITY_EVENTS.forEach((eventName) => {
        window.removeEventListener(eventName, onActivity);
      });
      window.removeEventListener("storage", onStorage);
    };
  }, [navigate]);

  return null;
}
