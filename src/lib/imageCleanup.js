const PAGE_SIZE = 500;

async function readAll(client, table, columns) {
  const rows = [];
  for (let offset = 0; ; offset += PAGE_SIZE) {
    const { data, error } = await client.from(table).select(columns)
      .order("id", { ascending: true }).range(offset, offset + PAGE_SIZE - 1);
    if (error) throw error;
    rows.push(...(data || []));
    if (!data || data.length < PAGE_SIZE) return rows;
  }
}

export function eligibleImages(purchases, images) {
  const incompleteOrders = new Set(purchases.filter((p) => !p.collected).map((p) => p.order_id));
  const eligibleIds = new Set(purchases.filter((p) => p.order_id && !incompleteOrders.has(p.order_id)).map((p) => p.id));
  const protectedPaths = new Set(images.filter((img) => !eligibleIds.has(img.purchase_id)).map((img) => img.storage_path));
  return images.filter((img) => eligibleIds.has(img.purchase_id) && img.storage_path && !protectedPaths.has(img.storage_path));
}

export async function cleanupCollectedOrderImages(client) {
  let deleted = 0;
  try {
    const purchases = await readAll(client, "purchases", "id,order_id,collected");
    const images = await readAll(client, "purchase_images", "id,purchase_id,storage_path");
    const candidates = eligibleImages(purchases, images);
    for (let offset = 0; offset < candidates.length; offset += 100) {
      // Recheck collection status and shared paths before each irreversible Storage request.
      const current = eligibleImages(
        await readAll(client, "purchases", "id,order_id,collected"),
        await readAll(client, "purchase_images", "id,purchase_id,storage_path")
      );
      const chunk = candidates.slice(offset, offset + 100).filter((img) => current.some((row) => row.id === img.id && row.storage_path === img.storage_path));
      if (!chunk.length) continue;
      const { error: storageError } = await client.storage.from("purchase-images").remove([...new Set(chunk.map((img) => img.storage_path))]);
      if (storageError) throw storageError;
      const { data, error } = await client.from("purchase_images").delete().in("id", chunk.map((img) => img.id)).select("id");
      if (error) throw new Error("تم حذف الملفات من التخزين، لكن تعذر حذف سجلات الصور. أعيدي المحاولة.");
      deleted += data?.length || 0;
      if (data?.length !== chunk.length) throw new Error("تعذر حذف بعض سجلات الصور. تحققي من صلاحيات الحذف وأعيدي المحاولة.");
    }
    return { deleted };
  } catch (error) {
    error.deleted = deleted;
    throw error;
  }
}
