const GEMINI_API_KEY = String(import.meta.env?.VITE_GEMINI_API_KEY || "").trim();
const GEMINI_MODEL_ID = "gemini-3-flash-preview";

const GEMINI_PROMPT = [
  "You are analyzing a shopping cart screenshot (like SHEIN).",
  "Identify Products: Find every unique item name.",
  "Price Selection: Pick the current final price. Ignore strikethrough (old) prices. If two prices are visible, pick the lower one.",
  "Clutter Filtering: Explicitly ignore status bar numbers (time, battery %, signal) and notification badge counts.",
  "Reasoning: Explain why you chose certain numbers and ignored others.",
  "Return JSON that matches the provided schema."
].join("\n");

let genaiPromise = null;

function getApiKey() {
  if (!GEMINI_API_KEY || GEMINI_API_KEY.includes("__GEMINI_API_KEY__")) return "";
  return GEMINI_API_KEY;
}

async function loadGeminiClient() {
  if (!genaiPromise) {
    genaiPromise = import("https://esm.sh/@google/genai@1.40.0");
  }
  const moduleRef = await genaiPromise;
  return moduleRef.default?.GoogleGenAI ? moduleRef.default : moduleRef;
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("تعذر قراءة الصورة."));
    reader.onload = () => {
      const result = reader.result;
      if (typeof result !== "string") {
        reject(new Error("بيانات الصورة غير صالحة."));
        return;
      }
      const base64 = result.split(",")[1];
      if (!base64) {
        reject(new Error("تعذر تحويل الصورة."));
        return;
      }
      resolve(base64);
    };
    reader.readAsDataURL(file);
  });
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("تعذر قراءة الصورة."));
    reader.onload = () => {
      const result = reader.result;
      if (typeof result !== "string") {
        reject(new Error("بيانات الصورة غير صالحة."));
        return;
      }
      const base64 = result.split(",")[1];
      if (!base64) {
        reject(new Error("تعذر تحويل الصورة."));
        return;
      }
      resolve(base64);
    };
    reader.readAsDataURL(blob);
  });
}

async function urlToBase64(url) {
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error("تعذر تحميل إحدى الصور.");
  }
  const blob = await res.blob();
  const data = await blobToBase64(blob);
  const mimeType =
    blob.type || (String(url).toLowerCase().endsWith(".png") ? "image/png" : "image/jpeg");
  return { data, mimeType };
}

function buildResponseSchema(Type) {
  return {
    type: Type.OBJECT,
    properties: {
      reasoning: { type: Type.STRING },
      items: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            name: { type: Type.STRING },
            price: { type: Type.NUMBER }
          },
          required: ["name", "price"]
        }
      },
      totalCalculated: { type: Type.NUMBER },
      ambiguities: {
        type: Type.ARRAY,
        items: { type: Type.STRING }
      }
    },
    required: ["reasoning", "items", "totalCalculated", "ambiguities"]
  };
}

function getResponseText(response) {
  if (!response) return "";
  if (typeof response.text === "function") return String(response.text() || "").trim();
  return String(response.text || "").trim();
}

export function hasGeminiKey() {
  return Boolean(getApiKey());
}

export function resolveTotalFromGemini(result) {
  const total = Number(result?.totalCalculated);
  if (Number.isFinite(total)) return total;
  const items = Array.isArray(result?.items) ? result.items : [];
  return items.reduce((sum, item) => sum + (Number(item?.price) || 0), 0);
}

export async function runGeminiCartAnalysis({ files = [], urls = [], onProgress }) {
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error("مفتاح Gemini غير مضبوط. أضيفي VITE_GEMINI_API_KEY داخل ملف .env.");
  }

  const { GoogleGenAI, Type } = await loadGeminiClient();
  if (!GoogleGenAI || !Type) {
    throw new Error("تعذر تحميل مكتبة Gemini.");
  }

  const client = new GoogleGenAI({ apiKey });
  const parts = [{ text: GEMINI_PROMPT }];
  const imageParts = [];

  for (let i = 0; i < files.length; i += 1) {
    onProgress?.(`تحويل صورة جديدة ${i + 1}/${files.length}...`);
    const data = await fileToBase64(files[i]);
    imageParts.push({
      inlineData: {
        mimeType: files[i].type || "image/png",
        data
      }
    });
  }

  for (let i = 0; i < urls.length; i += 1) {
    onProgress?.(`تحميل صورة محفوظة ${i + 1}/${urls.length}...`);
    const imgData = await urlToBase64(urls[i]);
    imageParts.push({
      inlineData: {
        mimeType: imgData.mimeType || "image/jpeg",
        data: imgData.data
      }
    });
  }

  if (!imageParts.length) {
    throw new Error("لا توجد صور للتحليل.");
  }

  onProgress?.("إرسال الصور إلى Gemini...");
  const response = await client.models.generateContent({
    model: GEMINI_MODEL_ID,
    contents: [
      {
        role: "user",
        parts: [...parts, ...imageParts]
      }
    ],
    config: {
      responseMimeType: "application/json",
      responseSchema: buildResponseSchema(Type)
    }
  });

  const rawText = getResponseText(response);
  if (!rawText) {
    throw new Error("لم يرجع النموذج أي نتيجة.");
  }

  try {
    return JSON.parse(rawText);
  } catch {
    throw new Error("تعذر قراءة نتيجة التحليل. جربي صورة أوضح.");
  }
}
