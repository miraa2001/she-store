import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { SignJWT, importPKCS8 } from "https://deno.land/x/jose@v4.15.5/index.ts";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const GOOGLE_SHEET_ID =
  Deno.env.get("GOOGLE_SHEET_ID") ?? "1d2yQ_ovd7w0CamppPpnvoQ2fb_3U6Q-m19tv7As1jiQ";
const GOOGLE_CLIENT_EMAIL = Deno.env.get("GOOGLE_SERVICE_ACCOUNT_EMAIL") ?? "";
const GOOGLE_PRIVATE_KEY = (Deno.env.get("GOOGLE_PRIVATE_KEY") ?? "").replace(/\\n/g, "\n");
const BUCKET = "purchase-images";

const sb = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);

type PurchaseRecord = {
  id: string;
  customer_name: string | null;
  qty: number | null;
  price: number | null;
  pickup_point: string | null;
  note: string | null;
  picked_up: boolean | null;
  picked_up_at: string | null;
  collected: boolean | null;
  collected_at: string | null;
  purchase_links?: { url: string }[];
  purchase_images?: { storage_path: string }[];
};

type OrderRecord = {
  id: string;
  order_name: string | null;
  purchases?: PurchaseRecord[];
};

function sanitizeSheetTitle(title: string) {
  const cleaned = title.replace(/[\[\]\*\/\\\?\:]/g, " ").trim();
  return cleaned.slice(0, 100) || "Order";
}

function publicImageUrl(path: string) {
  return `${SUPABASE_URL}/storage/v1/object/public/${BUCKET}/${path}`;
}

async function getAccessToken() {
  if (!GOOGLE_CLIENT_EMAIL || !GOOGLE_PRIVATE_KEY) {
    throw new Error("Missing Google service account credentials.");
  }

  const key = await importPKCS8(GOOGLE_PRIVATE_KEY, "RS256");
  const now = Math.floor(Date.now() / 1000);

  const jwt = await new SignJWT({
    scope: "https://www.googleapis.com/auth/spreadsheets",
  })
    .setProtectedHeader({ alg: "RS256", typ: "JWT" })
    .setIssuer(GOOGLE_CLIENT_EMAIL)
    .setAudience("https://oauth2.googleapis.com/token")
    .setIssuedAt(now)
    .setExpirationTime(now + 60 * 60)
    .sign(key);

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion: jwt,
  });

  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Failed to get access token: ${err}`);
  }

  const data = await res.json();
  return data.access_token as string;
}

async function fetchOrder(orderId: string) {
  const { data, error } = await sb
    .from("orders")
    .select(
      `
      id,
      order_name,
      purchases (
        id,
        customer_name,
        qty,
        price,
        pickup_point,
        note,
        picked_up,
        picked_up_at,
        collected,
        collected_at,
        purchase_links ( url ),
        purchase_images ( storage_path )
      )
    `,
    )
    .eq("id", orderId)
    .maybeSingle();

  if (error) {
    throw new Error(`Supabase error: ${error.message}`);
  }

  return data as OrderRecord | null;
}

async function ensureSheet(accessToken: string, title: string) {
  const sheetsRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    },
  );

  if (!sheetsRes.ok) {
    const err = await sheetsRes.text();
    throw new Error(`Failed to read spreadsheet: ${err}`);
  }

  const sheetsData = await sheetsRes.json();
  const existing = (sheetsData.sheets || []).find(
    (sheet: { properties: { title: string } }) => sheet.properties.title === title,
  );

  if (existing) return;

  const createRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}:batchUpdate`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        requests: [{ addSheet: { properties: { title } } }],
      }),
    },
  );

  if (!createRes.ok) {
    const err = await createRes.text();
    throw new Error(`Failed to create sheet: ${err}`);
  }
}

function buildRows(order: OrderRecord) {
  const header = [
    "اسم الطلبية",
    "رقم المشترى",
    "اسم الزبون",
    "العدد",
    "السعر",
    "مكان الاستلام",
    "ملاحظة",
    "تم الاستلام",
    "تاريخ الاستلام",
    "تم التحصيل",
    "تاريخ التحصيل",
    "روابط",
    "صور",
  ];

  const rows = (order.purchases || []).map((p) => {
    const links = (p.purchase_links || []).map((l) => l.url).join("\n");
    const images = (p.purchase_images || [])
      .map((img) => publicImageUrl(img.storage_path))
      .join("\n");

    return [
      order.order_name || "",
      p.id,
      p.customer_name || "",
      p.qty ?? "",
      p.price ?? "",
      p.pickup_point || "",
      p.note || "",
      p.picked_up ? "نعم" : "لا",
      p.picked_up_at || "",
      p.collected ? "نعم" : "لا",
      p.collected_at || "",
      links,
      images,
    ];
  });

  return [header, ...rows];
}

async function updateSheet(accessToken: string, title: string, rows: string[][]) {
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}/values/${encodeURIComponent(
      `${title}!A1`,
    )}?valueInputOption=RAW`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ values: rows }),
    },
  );

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Failed to update sheet: ${err}`);
  }
}

serve(async (req) => {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  try {
    const payload = await req.json();
    const orderId = payload.order_id || payload.orderId || payload.record?.order_id;
    if (!orderId) {
      return new Response("Missing order_id", { status: 400 });
    }

    const order = await fetchOrder(orderId);
    if (!order) {
      return new Response("Order not found", { status: 404 });
    }

    const accessToken = await getAccessToken();
    const sheetTitle = sanitizeSheetTitle(order.order_name || order.id);

    await ensureSheet(accessToken, sheetTitle);
    await updateSheet(accessToken, sheetTitle, buildRows(order));

    return new Response(JSON.stringify({ ok: true }), {
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    console.error(err);
    return new Response(
      JSON.stringify({ ok: false, error: String(err?.message || err) }),
      { status: 500, headers: { "Content-Type": "application/json" } },
    );
  }
});
