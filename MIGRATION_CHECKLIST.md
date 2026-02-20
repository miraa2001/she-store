# Regression Checklist (Legacy Parity)

Use this checklist after every migration step.

## Current Automated Status (2026-02-14)
- [x] React production build passes (`npm run build`).
- [x] Native routes are wired for all staff pages.
- [x] Global 15-minute inactivity timeout guard is enabled.
- [x] GitHub Pages deploy workflow is configured.

## Manual QA Required Before Go-Live

## Auth + Access
- [ ] Login works from `login.html` equivalent flow.
- [ ] Role gating works (`rahaf`, `reem`, `rawand`, `laaura`).
- [ ] Unauthorized sessions redirect correctly.

## Orders (index)
- [ ] Orders list loads and selects correctly.
- [ ] Add purchase modal opens/closes and submit flow works.
- [ ] Edit purchase flow works (including image previews).
- [ ] Delete purchase works.
- [ ] Arrived toggle updates state and reflects immediately.
- [ ] Filters and sorting produce same results.
- [ ] Search by customer works in edit/view/customers contexts.

## Actions + Integrations
- [ ] WhatsApp action buttons open with correct templated messages.
- [ ] PDF export works for current order.
- [ ] Gemini extraction flow still works.
- [ ] Undo behavior works where currently available.
- [ ] Lightbox works for purchase images.

## Sidebar + Navigation
- [ ] Global sidebar links remain role-correct.
- [ ] Tab switching (orders/view/customers) remains functional.
- [ ] Workspace list collapse/expand behaves on desktop + mobile.

## Other Pages
- [ ] pickup-dashboard
- [ ] pickuppoint
- [ ] finance
- [ ] archive
- [ ] collections
- [ ] homepickup

## Visual + Responsive
- [ ] RTL layout remains correct.
- [ ] No horizontal overflow on mobile.
- [ ] Keyboard focus visible for interactive controls.
