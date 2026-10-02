# She-Store Operations Dashboard

Production-facing operations dashboard for managing orders, pickup workflows, collections, and finance reporting.

## Overview
This project is a React + Vite single-page application used by internal staff roles to operate the full order lifecycle:
- Create and manage orders and purchases
- Track pickup status across home pickup and pickup point flows
- Monitor collected vs pending amounts
- Export order summaries and send customer-facing WhatsApp templates

The app is role-aware and route-protected, with different capabilities for each staff account.

## Tech Stack
- React 18
- Vite 5
- React Router (hash routing for GitHub Pages)
- Supabase (database + storage)
- Styled CSS modules/files per page + shared design tokens

## Core Product Areas
- `#/orders` : primary order workspace (orders list, purchases, customers)
- `#/pickup-dashboard` : pickup workbench with tabs for home, pickup point, and collections
- `#/homepickup` : home pickup processing view
- `#/pickuppoint` : pickup point processing view
- `#/finance` : order-level and monthly financial breakdowns
- The finance ledger stores spending and postal fees per order. Purchase value is the sum of paid purchase prices, falling back to the original price when no paid price is set, and profit is `(purchase value + postal fee) - spending`.
- `#/archive` : historical orders overview
- `#/collections` : collection tracking view
- `#/login` : authentication entry

## Access Model
Current role behavior is implemented in the app routing and UI controls:
- `rahaf`: full operational access
- `reem`, `rawand`: restricted/view-focused access
- `maryamti`: pickup point focused access

## Architecture Notes
- Data access is centralized in `src/lib/*` helpers (`orders`, `purchases`, `customers`, `pickup`, `session`, `whatsapp`).
- Shared UI elements and patterns are in `src/components/common/*` and `src/components/orders/*`.
- Pickup pages share styles through `src/pages/pickup-common.css` for consistent menus, lists, and states.
- Sidebar navigation and role-based item visibility are configured in `src/lib/navigation.js`.

## Notable Engineering Decisions
- Hash-based routes to support static hosting on GitHub Pages.
- Session inactivity timeout guard for safer shared-device usage.
- Unified design system and icon structure under `src/assets/icons/*`.
- Reusable orders menu pattern across pickup/finance/collections pages for consistent UX.

## Local Development
```bash
npm install
npm run dev
```

Create `.env` from `.env.example` and set:
- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`
- `VITE_GEMINI_API_KEY` (optional, only needed for extraction feature)

## Build
```bash
npm run build
npm run preview
```

## Finance Database Setup
Run `supabase/migrations/20261002000000_add_order_postal_fee.sql` in the Supabase SQL Editor before using the finance ledger. It adds `orders.postal_fee` with a default of zero and reuses the existing `spent_amount` column. Existing orders and manually configured profit shares are preserved.

### Home Cash
Run `supabase/migrations/20261002010000_add_home_cash_ledger.sql` in the Supabase SQL Editor to enable the Home Cash tab in Finance. Tracking begins at zero when the migration runs; previously received Home purchases and previously collected purchases are excluded. Rerunning the migration preserves the balance and history.

- Home purchases add cash when marked received. Collecting those same purchases later does not add cash again.
- Maryamti, Nablus, and Delivery add cash when marked collected, not when marked received.
- Receipt amounts use `paid_price`, falling back to `price` only when no paid price is set. Corrections and reversals create adjustment entries; parcel transfers do not move cash already received.
- Expenses record a category, total paid amount, and optional note. New entries use an internal quantity of one for compatibility with the existing database. Categories include bags, pins, name stickers, postal delivery, and ATM deposits. Depositing money at an ATM reduces cash held at home.
- Expenses cannot exceed available cash. Cancelling an expense restores its amount and adds a reversal without deleting the original entry.
- The activity log records amounts, running balances, dates, customer/order details, and the acting user. Deleting an order or purchase does not erase cash already recorded.

Automatic receipts are handled by a database trigger, including staff actions outside Finance. Cash history and expense RPCs are restricted to database admins; the Finance UI retains its existing Rahaf-only access.

Run `npm run test:home-cash` for isolated PostgreSQL regression tests covering the actual migration, receipt rules, corrections, expense retries, cancellations, pagination, and access policies. Tests do not connect to the live Supabase database.

## Profit Distribution Setup
Finance shows only orders created on or after October 1, 2026 in Asia/Hebron, in the finance ledger and per-order profit table. This is a display cutoff: historical cash, accrued profit balances, and payment records are not deleted or reset. The Finance tabs are Finance Ledger, Profit Distribution, and Home Cash; Home Cash shows only the cash balance summary, with expenses and activity history below.

After the postal-fee and Home Cash migrations, run `supabase/migrations/20261002040000_add_profit_distribution.sql` in Supabase SQL Editor.

On first installation, old manually entered `total_profit`, `mira_profit`, and `rahaf_profit` values are zeroed as requested. Spending, purchase prices, receipts, and cash history are preserved. New percentages are blank for every order, and marketing defaults to zero. Rerunning the migration does not repeat this reset.

Order Settings accepts Home, Rahaf, and Mira percentages totaling 100%, plus a marketing amount. Distributable profit is `max(0, sum(paid_price ?? price) + postal_fee - spent_amount - marketing_fee)`. Percentages must be configured before an order contributes to party balances. Earnings count immediately, without waiting for collections; cents are allocated deterministically so the three shares sum to the distributable profit.

The Finance Profit Distribution tab shows per-order shares, accumulated balances, total paid, and a paginated payout log. A partial payout subtracts only that amount: 400 earned minus 250 paid leaves 150. Each payout requires an explicit Yes/No choice for deduction from Home Cash; the Home party is independent of the cash account. Deductions and payouts are atomic, checked against both balances, and safe to retry after a lost response. Only Rahaf/database admins can record payouts.

Editing order finance or purchase prices automatically recalculates earnings without deleting payout history. If earnings fall below amounts already paid, the negative remaining balance represents overpayment and further payouts are blocked. Deleting an order preserves its last earnings snapshot and payout history rather than erasing paid obligations. Instant pickups never enter order profit.

Run `npm run test:profits` for isolated PostgreSQL checks of percentages, marketing, paid prices, immediate earnings, partial payouts, optional cash deductions, access controls, rounding, and migration reruns.

For per-party deductions, run `supabase/migrations/20261003000000_add_party_profit_deductions.sql` after the profit-distribution migration. Each order has three NIS deduction fields defaulting to zero. A party's earnings are `max(0, rounded percentage share - party deduction)`; deductions do not increase another party's earnings or change Home Cash. Payout history is preserved, and Finance shows the deduction beneath each affected share. The pickup dashboard now contains only the four pickup locations; individual and location collection controls remain available.

## Instant Pickup Setup
After installing Home Cash, run `supabase/migrations/20261002020000_add_instant_pickups.sql` in the Supabase SQL Editor.

Instant pickups are standalone records, with no order or customer-directory relationship. Rahaf adds a name, price, and one of the four pickup locations from the Orders action menu, beside New Order, or from a pickup section. New records are ready for pickup but neither received nor collected.

Rahaf and Reem see the `استلام فوري` section at locations they can access. Pickup-point staff see entries grouped by their creation date in the Asia/Hebron timezone, without the instant-pickup label. Staff receive purchases only at their own location; only Rahaf collects or transfers them. Collected entries cannot be transferred.

Instant receipts feed Home Cash using the same receipt and collection rules as normal purchases, including duplicate protection and audit entries. They never enter order purchase values, spending, or profit calculations. The migration does not change existing orders or purchases.

The Orders page also has an Instant Pickup tab for Rahaf and Reem. On desktop it appears beside Customers (where permitted); on mobile it is available in the floating actions. This view supports name search, location filtering, receipt and collection controls, and transfers while preserving existing location permissions. It uses the same instant-pickup migration, with no additional database setup.

Run `supabase/migrations/20261002030000_add_instant_pickup_delete.sql` to enable Rahaf's instant-pickup delete buttons in the Orders tab and pickup dashboards. Deletion requires confirmation and is blocked after receipt or collection. Undo receipt first for an uncollected entry so its cash adjustment is logged; existing cash history is retained after deletion.

Run `npm run test:instant-pickups` for isolated PostgreSQL checks of creation, cash timing, transfers, bulk collection, deletion, location permissions, and audit preservation.

## Deployment (GitHub Pages)
Workflow file: `.github/workflows/deploy-react-pages.yml`

Requirements:
1. Repository Pages source set to `GitHub Actions`
2. Repository secrets configured:
- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`
- `VITE_GEMINI_API_KEY`

On push to `main`, the workflow builds and deploys `dist/`.

## QA / Verification
Use `MIGRATION_CHECKLIST.md` for regression checks across roles and routes.

Minimum release checks:
- Login/logout + session timeout
- Orders CRUD and customer flows
- Pickup and collections updates
- Finance calculations
- PDF export and WhatsApp actions

## Repository Structure
```text
src/
  assets/
  components/
    common/
    orders/
    tabs/
  hooks/
  lib/
  pages/
public/
  legacy/
.github/workflows/
```

## Roadmap Candidates
- Additional test coverage for role-specific UI states
- Bundle size optimization via route-level code splitting
- Stronger typed data contracts around Supabase responses
