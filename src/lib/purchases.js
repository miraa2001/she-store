import { sb } from "./supabaseClient";
import { IMAGE_BUCKET } from "./orders";
import {
  PICKUP_DELIVERY,
  PICKUP_HOME,
  PICKUP_POINT,
  PICKUP_POINT_NABLUS,
  formatPickupFormValue
} from "./pickup";

const IMAGE_UPLOAD_CONCURRENCY = 3;
const DEFAULT_BAG_SIZE = "كيس صغير";

function safeFileName(name) {
  return String(name || "image")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^a-zA-Z0-9._-]/g, "");
}

export function toPublicImageUrl(storagePath) {
  if (!storagePath) return "";
  const { data } = sb.storage.from(IMAGE_BUCKET).getPublicUrl(storagePath);
  return data?.publicUrl || "";
}

export function sanitizeLinks(list) {
  const out = [];

  (Array.isArray(list) ? list : []).forEach((raw) => {
    const value = String(raw || "").trim();
    if (!value) return;

    const withScheme = /^https?:\/\//i.test(value) ? value : `https://${value}`;

    try {
      const parsed = new URL(withScheme);
      if (!/^https?:$/i.test(parsed.protocol)) return;
      out.push(parsed.toString());
    } catch {
      // skip invalid url
    }
  });

  return Array.from(new Set(out));
}

export async function fetchPurchasesByOrder(orderId) {
  if (!orderId) return [];

  const { data, error } = await sb
    .from("purchases")
    .select(
      "id, order_id, customer_id, customer_name, qty, price, paid_price, bag_size, pickup_point, ready_for_pickup, ready_for_pickup_at, note, created_at, collected, picked_up, purchase_links(url), purchase_images(id,storage_path)"
    )
    .eq("order_id", orderId)
    .order("created_at", { ascending: false });

  if (error) throw error;

  return (data || []).map((purchase) => {
    const links = (purchase.purchase_links || []).map((item) => String(item?.url || "")).filter(Boolean);
    const images = (purchase.purchase_images || []).map((img) => ({
      id: img.id,
      storage_path: img.storage_path,
      url: toPublicImageUrl(img.storage_path)
    }));

    return {
      ...purchase,
      links,
      images
    };
  });
}

export async function fetchCustomersList() {
  const { data, error } = await sb
    .from("customers")
    .select("id, name, phone, city, usual_pickup_point")
    .order("name", { ascending: true });

  if (error) throw error;
  return data || [];
}

export async function uploadPurchaseImages({ files, orderId, purchaseId, onProgress }) {
  const validFiles = (Array.isArray(files) ? files : []).filter(
    (file) => file && file.type && file.type.startsWith("image/")
  );

  if (!validFiles.length) {
    return { uploadedPaths: [], errors: [] };
  }

  const uploadedPaths = [];
  const errors = [];
  let done = 0;

  const uploadOne = async (file) => {
    const random = Math.random().toString(36).slice(2, 8);
    const path = `${orderId}/${purchaseId}/${Date.now()}-${random}-${safeFileName(file.name)}`;
    const { error } = await sb.storage.from(IMAGE_BUCKET).upload(path, file);
    if (error) throw new Error(error.message || "Upload failed");
    return path;
  };

  for (let i = 0; i < validFiles.length; i += IMAGE_UPLOAD_CONCURRENCY) {
    const batch = validFiles.slice(i, i + IMAGE_UPLOAD_CONCURRENCY);
    const results = await Promise.allSettled(batch.map(uploadOne));

    results.forEach((result, batchIndex) => {
      done += 1;
      onProgress?.(done, validFiles.length);

      if (result.status === "fulfilled") {
        uploadedPaths.push(result.value);
      } else {
        errors.push({
          file: batch[batchIndex]?.name || `image-${i + batchIndex + 1}`,
          message: result.reason?.message || "Upload failed"
        });
      }
    });
  }

  return { uploadedPaths, errors };
}

export async function createPurchaseWithRelations({
  purchase,
  links,
  files,
  onUploadProgress
}) {
  const normalizedPurchase = { ...purchase, bag_size: DEFAULT_BAG_SIZE };
  const { data: created, error: createError } = await sb
    .from("purchases")
    .insert(normalizedPurchase)
    .select()
    .single();

  if (createError) throw createError;

  const cleanLinks = sanitizeLinks(links);

  if (cleanLinks.length) {
    const { error: linksError } = await sb
      .from("purchase_links")
      .insert(cleanLinks.map((url) => ({ purchase_id: created.id, url })));

    if (linksError) throw linksError;
  }

  const uploadResult = await uploadPurchaseImages({
    files,
    orderId: normalizedPurchase.order_id,
    purchaseId: created.id,
    onProgress: onUploadProgress
  });

  if (uploadResult.uploadedPaths.length) {
    const { error: metaError } = await sb
      .from("purchase_images")
      .insert(uploadResult.uploadedPaths.map((storage_path) => ({ purchase_id: created.id, storage_path })));

    if (metaError) {
      uploadResult.errors.push({ file: "purchase_images", message: "Image metadata save failed." });
    }
  }

  return { purchase: created, uploadErrors: uploadResult.errors };
}

export async function updatePurchaseWithRelations({
  purchaseId,
  purchasePatch,
  links,
  removeImages,
  newFiles,
  onUploadProgress
}) {
  const normalizedPatch = { ...purchasePatch, bag_size: DEFAULT_BAG_SIZE };
  const { error: updateError } = await sb.from("purchases").update(normalizedPatch).eq("id", purchaseId);
  if (updateError) throw updateError;

  const { error: deleteLinksError } = await sb.from("purchase_links").delete().eq("purchase_id", purchaseId);
  if (deleteLinksError) throw deleteLinksError;

  const cleanLinks = sanitizeLinks(links);
  if (cleanLinks.length) {
    const { error: insertLinksError } = await sb
      .from("purchase_links")
      .insert(cleanLinks.map((url) => ({ purchase_id: purchaseId, url })));

    if (insertLinksError) throw insertLinksError;
  }

  const toRemove = Array.isArray(removeImages) ? removeImages.filter((img) => img?.id) : [];
  if (toRemove.length) {
    const ids = toRemove.map((img) => img.id);
    const paths = toRemove.map((img) => img.storage_path).filter(Boolean);

    if (paths.length) {
      await sb.storage.from(IMAGE_BUCKET).remove(paths);
    }

    const { error: deleteImagesError } = await sb.from("purchase_images").delete().in("id", ids);
    if (deleteImagesError) throw deleteImagesError;
  }

  const uploadResult = await uploadPurchaseImages({
    files: newFiles,
    orderId: normalizedPatch.order_id,
    purchaseId,
    onProgress: onUploadProgress
  });

  if (uploadResult.uploadedPaths.length) {
    const { error: metaError } = await sb
      .from("purchase_images")
      .insert(uploadResult.uploadedPaths.map((storage_path) => ({ purchase_id: purchaseId, storage_path })));

    if (metaError) {
      uploadResult.errors.push({ file: "purchase_images", message: "Image metadata save failed." });
    }
  }

  return { uploadErrors: uploadResult.errors };
}

export async function deletePurchaseById(purchaseId) {
  const { error } = await sb.from("purchases").delete().eq("id", purchaseId);
  if (error) throw error;
}

export async function restoreDeletedPurchase(snapshot) {
  if (!snapshot?.purchase?.id) return;

  const { purchase, links, images } = snapshot;
  const normalizedPurchase = { ...purchase, bag_size: DEFAULT_BAG_SIZE };
  const { error: purchaseError } = await sb.from("purchases").insert(normalizedPurchase);
  if (purchaseError) throw purchaseError;

  const cleanLinks = sanitizeLinks(links || []);
  if (cleanLinks.length) {
    const { error: linksError } = await sb
      .from("purchase_links")
      .insert(cleanLinks.map((url) => ({ purchase_id: purchase.id, url })));

    if (linksError) throw linksError;
  }

  if (Array.isArray(images) && images.length) {
    const rows = images
      .filter((img) => img?.storage_path)
      .map((img) => ({
        id: img.id,
        purchase_id: purchase.id,
        storage_path: img.storage_path
      }));

    if (rows.length) {
      const { error: imagesError } = await sb.from("purchase_images").insert(rows);
      if (imagesError) throw imagesError;
    }
  }
}

export async function markPurchasePaidPrice(purchaseId, paidPrice) {
  const { error } = await sb.from("purchases").update({ paid_price: paidPrice }).eq("id", purchaseId);
  if (error) throw error;
}

export function normalizeOperationalPickupPoint(value) {
  const normalized = formatPickupFormValue(value, PICKUP_HOME);
  if (!normalized || normalized === PICKUP_DELIVERY) return PICKUP_HOME;
  return normalized;
}

export function getTransferPickupOptions(currentValue) {
  const options = [
    { value: PICKUP_HOME, label: "البيت" },
    { value: PICKUP_POINT, label: "مريمتي" },
    { value: PICKUP_POINT_NABLUS, label: "نابلس" }
  ];
  const normalizedCurrent = normalizeOperationalPickupPoint(currentValue);
  if (normalizedCurrent && !options.some((option) => option.value === normalizedCurrent)) {
    options.push({
      value: normalizedCurrent,
      label: formatPickupFormValue(normalizedCurrent, PICKUP_HOME)
    });
  }

  return options;
}

export async function movePurchaseToPickupLocation(purchaseId, pickupPoint) {
  const id = String(purchaseId || "").trim();
  const nextPickupPoint = normalizeOperationalPickupPoint(pickupPoint);

  if (!id) {
    throw new Error("تعذر تحديد المشترى للنقل.");
  }

  if (!nextPickupPoint) {
    throw new Error("اختاري مكان الاستلام الجديد.");
  }

  const { data, error } = await sb
    .from("purchases")
    .update({
      pickup_point: nextPickupPoint,
      ready_for_pickup: true,
      ready_for_pickup_at: new Date().toISOString()
    })
    .eq("id", id)
    .eq("collected", false)
    .select("id, pickup_point")
    .maybeSingle();

  if (error) throw error;
  if (!data?.id) {
    throw new Error("لا يمكن نقل مشترى تم تحصيله.");
  }

  return data;
}

export async function searchPurchasesByCustomerName(query, limit = 50) {
  const text = String(query || "").trim();
  if (!text) return [];

  const { data, error } = await sb
    .from("purchases")
    .select("id, order_id, customer_name, price, paid_price, qty, created_at")
    .ilike("customer_name", `%${text}%`)
    .order("created_at", { ascending: false })
    .limit(limit);

  if (error) throw error;
  return data || [];
}
