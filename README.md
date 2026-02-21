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
- `#/archive` : historical orders overview
- `#/collections` : collection tracking view
- `#/login` : authentication entry

## Access Model
Current role behavior is implemented in the app routing and UI controls:
- `rahaf`: full operational access
- `reem`, `rawand`: restricted/view-focused access
- `laaura`: pickup point focused access

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
