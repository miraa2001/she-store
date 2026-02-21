import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL =
  import.meta.env.VITE_SUPABASE_URL || "https://maucjvxhrnkdeybltjco.supabase.co";

const SUPABASE_ANON_KEY =
  import.meta.env.VITE_SUPABASE_ANON_KEY || "sb_publishable_0N65DXPMsk_3WVcLQ9wavQ_p9HgB_Bo";

export const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
