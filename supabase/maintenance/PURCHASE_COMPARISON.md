# Purchase Comparison

Run `supabase/migrations/20261025030000_add_purchase_comparison.sql` in the Supabase SQL Editor before deploying the frontend. The unassigned-purchase, category-label, payment, and profit migrations must already be applied.

## Workflow

Rahaf opens the comparison tool from the unassigned-purchases tab. Select all or individual purchases. Choose one comparison category for purchases with multiple labels or no labels. Single-labelled purchases keep that category. These choices do not overwrite saved labels.

Choose 1-20 proposed orders, each with its own category and original-price target. Each destination can be a new Shein order or a different empty pending Shein order of the matching subtype. Fast maps to the existing `normal` subtype. New orders start pending with zero expenses and unconfigured profit percentages; existing order settings are preserved.

The bundled `javascript-lp-solver` runs in a cancellable Web Worker. Binary variables keep each purchase whole and in at most one order. Every group must be nonempty, category-compatible, and at most 100 items using the sum of `qty`. The objective first maximizes the number of purchase records included, then minimizes the sum of absolute differences from each target. Targets are not spending caps and may be exceeded. Totals use `price` once per purchase, not `paid_price` and not `price * qty`. Other purchases stay unassigned.

Up to three distinct assignments are proposed. Difficult searches are time-bounded; proposals are valid candidates, not a promise of global optimality. Each result shows totals, quantities, signed target differences, and leftovers. Manual moves enforce category and quantity limits; empty groups cannot be confirmed.

## Confirmation

Nothing is saved before confirmation. One transactional RPC locks and rechecks source purchases and destinations, including original prices, quantities, category labels, payment history, and pending/empty status. All assignments and new orders commit together or all roll back. Purchase IDs, images, notes, prices, and labels remain intact. Normal profit accruals refresh after assignment; this does not change the website's profit formula or home cash.

`purchase_comparison_commits` stores the request UUID, snapshot payload, result, actor, and timestamp. Identical retries return the original result, including after a lost response, rather than creating duplicate orders. Changed payloads cannot reuse a request UUID. Only admins can execute confirmation or read these records; client roles cannot edit/delete them.

The migration is rerunnable and does not change existing assignments or reset any financial data. Existing transfer controls remain available and unchanged.

Verification: `npm run test:comparison` exercises optimization, exhaustive small-case comparisons, permissions, snapshot validation, capacity, atomic rollback, mixed new/existing destinations, and idempotent retries.
