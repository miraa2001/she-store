# Partial Payments

Run migrations in timestamp order, including `20261021000000_add_purchase_receipt_rpc.sql` and then `20261021010000_add_purchase_payments.sql`, before deploying the updated frontend. Back up production first. The partial-payment migration requires the existing Home Cash, instant pickup, and profit migrations.

The agreed price remains `coalesce(paid_price, price)` for ordinary purchases and `price` for instant pickups. Profit calculations are unchanged. `received_amount` tracks customer payments; `remitted_amount` tracks how much reached Home Cash. The remaining debt is the agreed price minus received_amount. Money available for collection is received_amount minus remitted_amount.

Only Rahaf and Reem can enter custom payments. Reem can do so at Home, Nablus, and Delivery, not Maryamti. Run `20261022000000_restore_pickup_staff_receipts.sql` to let Maryamti and Nablus confirm ordinary fully-paid pickups at their own locations. Their checkbox opens a fixed full-payment confirmation, without an editable amount. This records a payment held at the pickup point, not income in Home Cash until Rahaf collects it. Staff cannot cancel payments, undo paid handovers, enter partial amounts, or settle purchases already marked received with a partial balance. Other roles retain read-only access. Only Rahaf collects proceeds or cancels a payment. The RPCs enforce permissions independently of UI controls.

Clicking receive opens a dialog with the remaining balance prefilled. A positive first payment is required. Later installments use the Add Payment button. Home payments immediately add to cash; other locations add cash only when Rahaf collects the payments recorded so far. Collection does not clear the customer's remaining debt. A purchase is complete (`collected`) only after receipt, full payment, and full remittance. Image cleanup therefore excludes incomplete payments.

The migration baselines already-received purchases as fully paid. It does not replay historical money into Home Cash. Previously uncollected location proceeds can still be collected. Legacy cash excluded before tracking started stays excluded. The old checkbox-driven cash triggers are retired; do not reapply earlier migrations after this migration.

Payment entry and collection are idempotent, amount-checked and serialized using the same Home Cash lock as other cash actions. Cancellation keeps the original payment, records a reason and actor, and reverses only the amount actually recognized in Home Cash. A cancellation is a correction, not a general customer-refund feature.

Any payment history (including cancelled payments) blocks location transfers and deleting the purchase/order that owns it. This protects audit records. Received payments also prevent lowering the agreed price below the total paid or resetting parcel readiness. Existing orders with no payment history remain editable and transferable.

Run `npm run test:payments` for the database, permissions, cash, history, client validation, and notification tests.
