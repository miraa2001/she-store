import { sb } from "./supabaseClient";

export async function signOutAndRedirect(targetHash = "#/login") {
  try {
    await sb.auth.signOut();
  } catch (error) {
    console.error(error);
  } finally {
    window.location.hash = targetHash;
  }
}

