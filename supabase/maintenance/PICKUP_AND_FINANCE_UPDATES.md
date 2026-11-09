# Pickup and finance updates

Run these files in Supabase SQL Editor, in order, before deploying the frontend:

1. `supabase/migrations/20261106000000_location_readiness_and_pickup_actor.sql`
2. `supabase/migrations/20261106010000_home_recipients_and_cash_log_views.sql`

Home and Delivery purchases become ready when their order arrives. Other locations
still require the order to be placed at pickup. Returning an order to pending removes
readiness from unpaid, unreceived purchases. Purchases with payment history retain
their readiness and financial records. New purchases and location transfers also
derive readiness from the destination order and location.

Household profit remains one party with its existing percentage, deductions and
balance. New household payouts require Reem, Rawand or Mama as recipient. Previous
payouts retain a null recipient and display as legacy household payouts. No old
amounts are redistributed. The existing `get_profit_balances` function is not
replaced, preserving any exclusions already applied to the live database.

Cash receipts and their corrections appear in the purchase-receipts view. Expenses,
expense cancellations, profit payouts and the opening entry appear in the manual
view. Pagination is applied after filtering; running balances still include every
cash transaction, across both views.

New pickups record the account ID and email from the authenticated database session.
Cash receipts preserve the pickup account separately from the cash movement account.
Only verified non-legacy receipt-payment records backfill older purchase actors;
other old pickups and old cash entries show an unknown pickup account. Neither
migration deletes orders, resets cash, nor replays historical payments.
