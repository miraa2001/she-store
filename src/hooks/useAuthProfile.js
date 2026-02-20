import { useCallback, useEffect, useState } from "react";
import { getCurrentUserProfile } from "../lib/auth";

const DEFAULT_PROFILE = {
  loading: true,
  authenticated: false,
  role: "viewer",
  email: "",
  user: null
};

export function useAuthProfile() {
  const [profile, setProfile] = useState(DEFAULT_PROFILE);

  const refreshProfile = useCallback(async () => {
    try {
      const result = await getCurrentUserProfile();
      setProfile({
        loading: false,
        authenticated: result.authenticated,
        role: result.role || "viewer",
        email: result.email || "",
        user: result.user || null
      });
      return result;
    } catch (error) {
      console.error(error);
      setProfile({
        loading: false,
        authenticated: false,
        role: "viewer",
        email: "",
        user: null
      });
      return null;
    }
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const result = await getCurrentUserProfile();
        if (cancelled) return;
        setProfile({
          loading: false,
          authenticated: result.authenticated,
          role: result.role || "viewer",
          email: result.email || "",
          user: result.user || null
        });
      } catch (error) {
        console.error(error);
        if (cancelled) return;
        setProfile({
          loading: false,
          authenticated: false,
          role: "viewer",
          email: "",
          user: null
        });
      }
    }

    load();
    return () => {
      cancelled = true;
    };
  }, []);

  return { profile, setProfile, refreshProfile };
}
