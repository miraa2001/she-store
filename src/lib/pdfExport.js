import { formatILS } from "./orders";

const PDF_FONT_REGULAR_URL =
  "https://cdn.jsdelivr.net/gh/googlefonts/noto-fonts/hinted/ttf/NotoNaskhArabic/NotoNaskhArabic-Regular.ttf";
const PDF_FONT_BOLD_URL =
  "https://cdn.jsdelivr.net/gh/googlefonts/noto-fonts/hinted/ttf/NotoNaskhArabic/NotoNaskhArabic-Bold.ttf";

let pdfLibPromise = null;
let fontkitPromise = null;
let fontBytesPromise = null;

function normalizeUrl(url) {
  const raw = String(url || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  return `https://${raw}`;
}

function getSafeOrderFileName(orderName) {
  const value = String(orderName || "Order")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, "-");
  return value || "Order";
}

async function loadPdfLib() {
  if (!pdfLibPromise) {
    pdfLibPromise = import("https://esm.sh/pdf-lib@1.17.1");
  }
  const moduleRef = await pdfLibPromise;
  return moduleRef.default?.PDFDocument ? moduleRef.default : moduleRef;
}

async function loadFontkit() {
  if (!fontkitPromise) {
    fontkitPromise = import("https://esm.sh/@pdf-lib/fontkit@1.1.1");
  }
  const moduleRef = await fontkitPromise;
  return moduleRef.default || moduleRef;
}

async function loadFontBytes() {
  if (!fontBytesPromise) {
    fontBytesPromise = Promise.all([fetch(PDF_FONT_REGULAR_URL), fetch(PDF_FONT_BOLD_URL)]).then(
      async ([regularRes, boldRes]) => {
        if (!regularRes.ok || !boldRes.ok) {
          throw new Error("تعذر تحميل الخط العربي.");
        }
        const [regularBytes, boldBytes] = await Promise.all([
          regularRes.arrayBuffer(),
          boldRes.arrayBuffer()
        ]);
        return { regularBytes, boldBytes };
      }
    );
  }

  return fontBytesPromise;
}

async function embedFonts(pdfDoc, PDFLib) {
  const fontkit = await loadFontkit();
  pdfDoc.registerFontkit(fontkit);
  const { regularBytes, boldBytes } = await loadFontBytes();
  const { StandardFonts } = PDFLib;

  return {
    regular: await pdfDoc.embedFont(regularBytes, { subset: true }),
    bold: await pdfDoc.embedFont(boldBytes, { subset: true }),
    latinRegular: await pdfDoc.embedFont(StandardFonts.Helvetica),
    latinBold: await pdfDoc.embedFont(StandardFonts.HelveticaBold)
  };
}

function addLinkAnnotation(pdfDoc, page, PDFLib, url, x, y, width, height) {
  if (!url) return;
  const { PDFName, PDFString } = PDFLib;
  const rect = [x, y, x + width, y + height];
  const annotation = pdfDoc.context.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Link"),
    Rect: rect,
    Border: [0, 0, 0],
    A: pdfDoc.context.obj({
      Type: PDFName.of("Action"),
      S: PDFName.of("URI"),
      URI: PDFString.of(url)
    })
  });
  page.node.addAnnot(annotation);
}

async function buildImageThumbs(pdfDoc, imageUrls) {
  const items = [];

  for (const url of imageUrls) {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error("failed");
      const bytes = await res.arrayBuffer();
      const contentType = (res.headers.get("content-type") || "").toLowerCase();
      const isPng = contentType.includes("png") || String(url).toLowerCase().endsWith(".png");
      const image = isPng ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes);
      const maxW = 160;
      const maxH = 120;
      const scale = Math.min(maxW / image.width, maxH / image.height, 1);
      const dims = image.scale(scale);
      items.push({ type: "image", image, width: dims.width, height: dims.height });
    } catch {
      items.push({ type: "error", width: 140, height: 18 });
    }
  }

  if (!items.length) {
    items.push({ type: "error", width: 140, height: 18 });
  }

  return items;
}

async function addPurchasePage(pdfDoc, purchase, order, fonts, PDFLib) {
  const { rgb, PageSizes } = PDFLib;
  const page = pdfDoc.addPage(PageSizes?.A4 || [595.28, 841.89]);
  const { width, height } = page.getSize();
  const { regular, bold, latinRegular } = fonts;

  const title = purchase.customer_name || order.name || "Purchase";
  const qtyValue = purchase.qty ?? "—";
  const unitPrice =
    purchase.price === null || purchase.price === undefined ? "—" : formatILS(purchase.price);
  const paidPrice =
    purchase.paid_price === null || purchase.paid_price === undefined
      ? "—"
      : formatILS(purchase.paid_price);

  const linkLines = Array.isArray(purchase.links) && purchase.links.length ? purchase.links : ["—"];
  const imageUrls = (purchase.images || []).map((img) => img?.url).filter(Boolean);
  const thumbItems = await buildImageThumbs(pdfDoc, imageUrls);

  const lines = [
    { type: "text", text: title, size: 22, font: bold, color: rgb(0, 0, 0), gap: 14, align: "right" },
    { type: "pair", label: "الكمية", value: qtyValue, size: 14, gap: 8 },
    { type: "pair", label: "سعر القطعة", value: unitPrice, size: 14, gap: 8 },
    { type: "pair", label: "السعر المدفوع", value: paidPrice, size: 14, gap: 14 },
    { type: "text", text: "رابط المنتج:", size: 12, font: bold, color: rgb(0, 0, 0), gap: 6, align: "right" }
  ];

  linkLines.forEach((url) => {
    lines.push({
      type: "text",
      text: String(url || "—"),
      size: 11,
      font: latinRegular || regular,
      color: url === "—" ? rgb(0, 0, 0) : rgb(0.1, 0.4, 0.9),
      link: url === "—" ? null : normalizeUrl(url),
      gap: 6,
      align: "left"
    });
  });

  lines.push({ type: "text", text: "الصور:", size: 12, font: bold, color: rgb(0, 0, 0), gap: 8, align: "right" });

  const thumbsGap = 12;
  const rowHeight = Math.max(...thumbItems.map((item) => item.height));
  const rowWidth =
    thumbItems.reduce((sum, item) => sum + item.width, 0) + thumbsGap * (thumbItems.length - 1);
  const contentHeight = lines.reduce((sum, line) => sum + line.size + line.gap, 0) + rowHeight;

  const blockWidth = Math.max(0, width - 120);
  const blockX = (width - blockWidth) / 2;
  let y = (height + contentHeight) / 2;

  const drawTextLine = (line) => {
    const safeText = String(line.text ?? "");
    const textWidth = line.font.widthOfTextAtSize(safeText, line.size);
    let x = blockX + (blockWidth - textWidth) / 2;
    if (line.align === "right") x = blockX + (blockWidth - textWidth);
    if (line.align === "left") x = blockX;
    page.drawText(safeText, {
      x,
      y,
      size: line.size,
      font: line.font,
      color: line.color,
      maxWidth: width - 80
    });
    if (line.link) {
      addLinkAnnotation(pdfDoc, page, PDFLib, line.link, x, y, textWidth, line.size + 2);
    }
  };

  const drawPairLine = (line) => {
    const labelText = String(line.label ?? "");
    const valueText = String(line.value ?? "—");
    const labelWidth = regular.widthOfTextAtSize(labelText, line.size);
    const valueWidth = (latinRegular || regular).widthOfTextAtSize(valueText, line.size);
    const labelX = blockX + blockWidth - labelWidth;
    const valueX = blockX;
    page.drawText(labelText, { x: labelX, y, size: line.size, font: regular, color: rgb(0, 0, 0) });
    page.drawText(valueText, {
      x: valueX,
      y,
      size: line.size,
      font: latinRegular || regular,
      color: rgb(0, 0, 0),
      maxWidth: Math.max(90, blockWidth - labelWidth - 24)
    });
  };

  for (const line of lines) {
    y -= line.size;
    if (line.type === "pair") {
      drawPairLine(line);
    } else {
      drawTextLine(line);
    }
    y -= line.gap;
  }

  const rowX = Math.max(40, (width - rowWidth) / 2);
  const rowY = y - rowHeight;
  let x = rowX;

  thumbItems.forEach((item) => {
    if (item.type === "image") {
      page.drawImage(item.image, { x, y: rowY, width: item.width, height: item.height });
    } else {
      page.drawText("Image unavailable", {
        x,
        y: rowY + Math.max(0, (rowHeight - item.height) / 2),
        size: 10,
        font: latinRegular || regular,
        color: rgb(0.4, 0.4, 0.4),
        maxWidth: item.width
      });
    }
    x += item.width + thumbsGap;
  });
}

function downloadPdf(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 0);
}

export async function exportOrderPdf({ order, purchases }) {
  const PDFLib = await loadPdfLib();
  const { PDFDocument, PageSizes } = PDFLib;
  const pdfDoc = await PDFDocument.create();
  const fonts = await embedFonts(pdfDoc, PDFLib);

  if (!Array.isArray(purchases) || purchases.length === 0) {
    const page = pdfDoc.addPage(PageSizes.A4);
    page.drawText("No purchases", {
      x: 50,
      y: page.getHeight() / 2,
      size: 16,
      font: fonts.bold
    });
  } else {
    for (const purchase of purchases) {
      await addPurchasePage(pdfDoc, purchase, order || {}, fonts, PDFLib);
    }
  }

  const bytes = await pdfDoc.save();
  const filename = `${getSafeOrderFileName(order?.name)}-Purchases.pdf`;
  downloadPdf(bytes, filename);
}
