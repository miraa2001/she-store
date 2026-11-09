# Abandoned Purchases Archive

Run `supabase/migrations/20261109010000_add_abandoned_purchase_archive.sql` in the Supabase SQL Editor before deploying the updated frontend. It requires the existing partial-payment, location-readiness and instant-pickup migrations.

The migration does not backfill or update existing purchases, payments, profits or Home Cash. The two new tables have RLS enabled and are readable only by Rahaf. Mutations use admin-only RPCs.

## Workflow

- Rahaf can archive individual purchases or select a batch from an order. A purchase must be linked to an order, unreceived, uncollected and have no payment history, including cancelled payments.
- Archived purchases disappear from pickup queues. They remain linked to their original order and continue to contribute to its finances exactly as before. Original customer, prices, notes, images and links are preserved.
- Available stock can be restored to the original order. Pickup readiness follows the order's current status and location. Repeated archive/restore cycles retain their history.
- Reselling creates a separate instant pickup with a new customer, price and location. Its receipts follow existing cash rules and never enter order profits.
- Deleting an unreceived, unpaid resale returns its archived item to available stock. Received resales or resales with payment history cannot be deleted.

## Data Protection

Archived originals and their image/link metadata are protected from editing or deletion. The retained history also prevents deleting their parent order. Restore an available purchase before editing it.

Images are referenced, not duplicated. Original purchases remain uncollected, so existing collected-order image cleanup does not remove their images. Operation records preserve who archived, restored, resold or cancelled a resale. Retried requests use stable IDs to prevent duplicate operations.

No SQL in this file should be used to mark historical purchases received or collected, or to reset any financial balances.
