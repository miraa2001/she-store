import assert from "node:assert/strict";
import test from "node:test";
import { cleanupCollectedOrderImages, eligibleImages } from "../src/lib/imageCleanup.js";

function fixture({ storageError = false, dbError = false, denied = false, count = 1, undoCollection = false } = {}) {
  const purchases = [{ id: "p", order_id: "o", collected: true }];
  let images = Array.from({ length: count }, (_, i) => ({ id: String(i).padStart(4, "0"), purchase_id: "p", storage_path: `image-${i}` }));
  const calls = [];
  let purchaseReads = 0;
  const client = {
    from(table) {
      let deleting = false, ids;
      const query = {
        select() { return query; }, order() { return query; },
        delete() { deleting = true; return query; },
        in(_, values) { ids = values; return query; },
        range(start, end) {
          if (table === "purchases" && ++purchaseReads > 1 && undoCollection) purchases[0].collected = false;
          return Promise.resolve({ data: (table === "purchases" ? purchases : images).slice(start, end + 1) });
        },
        then(resolve) {
          assert(deleting);
          calls.push("database");
          const data = denied ? [] : images.filter(img => ids.includes(img.id));
          if (!dbError && !denied) images = images.filter(img => !ids.includes(img.id));
          return Promise.resolve({ data, error: dbError ? Error("db failure") : null }).then(resolve);
        }
      };
      return query;
    },
    storage: { from(bucket) {
      assert.equal(bucket, "purchase-images");
      return { async remove(paths) { calls.push("storage"); assert(paths.length <= 100); return { error: storageError ? Error("storage failure") : null }; } };
    } }
  };
  return { client, calls, remaining: () => images.length };
}

test("only fully collected orders qualify; shared active files and standalone records stay", () => {
  const purchases = [{ id: "a", order_id: "one", collected: true }, { id: "b", order_id: "two", collected: true }, { id: "c", order_id: "two", collected: false }, { id: "d", order_id: null, collected: true }];
  const images = [{ id: "1", purchase_id: "a", storage_path: "safe" }, { id: "2", purchase_id: "b", storage_path: "partial" }, { id: "3", purchase_id: "a", storage_path: "shared" }, { id: "4", purchase_id: "c", storage_path: "shared" }, { id: "5", purchase_id: "d", storage_path: "instant" }];
  assert.deepEqual(eligibleImages(purchases, images).map(img => img.id), ["1"]);
});

test("pagination and batching remove Storage before metadata, including over 1000 images", async () => {
  const f = fixture({ count: 1001 });
  assert.deepEqual(await cleanupCollectedOrderImages(f.client), { deleted: 1001 });
  assert.equal(f.remaining(), 0);
  assert.deepEqual(f.calls, Array.from({ length: 11 }, () => ["storage", "database"]).flat());
});

test("storage errors preserve metadata", async () => {
  const f = fixture({ storageError: true });
  await assert.rejects(cleanupCollectedOrderImages(f.client), /storage failure/);
  assert.equal(f.remaining(), 1);
  assert.deepEqual(f.calls, ["storage"]);
});

test("database failure and silent RLS denial are not reported as success", async () => {
  for (const options of [{ dbError: true }, { denied: true }]) {
    const f = fixture(options);
    await assert.rejects(cleanupCollectedOrderImages(f.client));
    assert.equal(f.remaining(), 1);
  }
});

test("empty cleanup does not call Storage", async () => {
  const f = fixture({ count: 0 });
  assert.deepEqual(await cleanupCollectedOrderImages(f.client), { deleted: 0 });
  assert.deepEqual(f.calls, []);
});

test("collection undone before a batch protects its images", async () => {
  const f = fixture({ undoCollection: true });
  assert.deepEqual(await cleanupCollectedOrderImages(f.client), { deleted: 0 });
  assert.equal(f.remaining(), 1);
  assert.deepEqual(f.calls, []);
});
