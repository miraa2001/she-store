# Explicit Home Collection

Run `supabase/migrations/20261109020000_defer_home_cash_until_collection.sql` in the Supabase SQL Editor to activate the updated workflow.

1. Confirm Home pickup: the customer payment is saved, but the purchase remains uncollected. Home Cash does not increase.
2. Click collect as Rahaf: only uncollected payments enter Home Cash. A fully paid purchase becomes collected; a partially paid purchase remains outstanding.
3. Later installments also wait for collection. Repeating collection does not add the same payment twice.

The same rules apply to Home instant pickups. Other locations keep their existing deferred-collection behavior. Zero-value Home pickups also require explicit collection, without creating cash entries.

There is no backfill, balance reset or data update in the migration. Existing collected purchases, previously recognized Home cash, profit calculations and payout history are preserved. If an older Home installment already entered cash, only its new uncollected installments enter cash later. Do not unmark historical collected purchases to apply the new workflow retroactively.

Permissions remain unchanged: Rahaf and Reem can record permitted customer payments; only Rahaf can collect. Cash logs retain the collection account and the original pickup account through the existing audit functions.
