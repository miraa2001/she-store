# Unassigned Purchases

Run `supabase/migrations/20261025010000_add_unassigned_purchases.sql` in the Supabase SQL Editor before deploying the updated frontend. It uses the existing payment and profit schema and replaces the previous order-transfer RPC.

- An unassigned purchase is a normal `purchases` row with `order_id = null`, not an instant pickup or a placeholder order.
- Only admins (Rahaf in the app) can create, read, edit, or transfer these purchases. Their image/link metadata follows purchase visibility.
- Unassigned purchases cannot be marked ready, received, or collected. They contribute to neither order profits nor home cash.
- The transfer dialog accepts multiple purchases. Every participating order must still be pending, checked again in the database at save time.
- Ready/received/collected purchases and any purchase with payment history cannot be transferred.
- Transfers preserve purchase IDs, prices, customers, notes, links, and image paths. Source/destination profits recalculate using each order's existing settings; spending, marketing, percentages, deductions, and payouts do not move.
- New images for unassigned purchases use `unassigned/<purchase-id>/...`. Assigning an order does not move or re-upload those files.

The migration does not unassign or delete existing purchases. Rerunning it preserves existing assignments, profits, and cash transactions. No old column or table needs to be deleted.

For category labels, also run `supabase/migrations/20261025020000_add_purchase_categories.sql`. Each purchase can have any combination of Fast, Makeup, and Electronics. Existing purchases start without labels. Labels are retained when assigning/detaching purchases, but do not restrict destinations or change prices/profits. Edit an unassigned purchase to select or clear its labels.
