# Database Maintenance

Archive and PDF export used no dedicated tables, so removing those features requires no schema deletion or migration.

Use the Orders menu's image cleanup button to remove images only for orders with purchases that are all collected. The app deletes files through Storage first, then their purchase_images rows. Purchase, order, customer, cash, and profit records are retained. This uses the existing admin SELECT/DELETE policies on purchase_images and the purchase-images bucket; do not remove them.

Run inspect_storage_and_tables.sql in the Supabase SQL Editor for read-only table sizes, indexes, and potentially unreferenced image files. Review unreferenced files before deleting them: uploads in progress or another consumer may still need them. Delete confirmed unused files through the Storage Dashboard/API, not DELETE FROM storage.objects, which only deletes metadata and leaves the file behind.

Keep financial ledgers, profit accruals, payouts, receipt tracking, roles, and foreign-key/check constraints. Do not drop old columns until all remaining app and database function references have been checked. Normal vacuuming handles dead rows; a smaller number of records does not guarantee an immediate reduction of the allocated database file size.

Reference: https://supabase.com/docs/guides/storage/management/delete-objects
