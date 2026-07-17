# Build Prompt: SupplyDesk (SupplySarthi Clone) — Multi-Site B2B Supply, Billing & Credit System

> Paste this entire prompt into your AI coding tool (Claude Code, Cursor, Windsurf, etc.). It is fully self-contained.

## Project Overview

Build a Google Apps Script (GAS) web application called **SupplyDesk** for a B2B wholesale/distribution business that supplies materials to corporate clients across multiple physical sites (e.g. one client, five factory/office locations). It replaces manual WhatsApp/phone ordering, Excel-based site pricing, and manual GST invoice calculation with a single system covering: site-wise ordering, delivery reconciliation, dynamic credit-limit enforcement, auto-GST invoicing, a 24-hour complaint window with auto credit notes, and a Tally Prime-compatible export.

## Reusable Working-Style Playbook (mandatory — apply exactly, every project)

### Backend (Code.gs)
1. `initializeSheets()` — idempotent setup: creates all sheets + headers + formatting if missing, seeds dummy data covering every status/edge case, returns a summary object. Safe to run repeatedly (skip sheets that already exist).
2. `resetDummyData()` — companion function that clears seeded rows for a clean demo state, without touching real rows (use an ID-threshold convention: seeded rows have low IDs, real rows added later get IDs > 1000, or similar — keep it consistent with the sample below).
3. Standard response wrapper on **every** public function: `success(data)` / `fail(message)` — return `{ success, data, error }` always, never a bare value or a thrown error to the client.
4. `LockService.getScriptLock()` wraps every write operation (create/update/delete), released in a `finally` block.
5. `CacheService` (script or user cache) on expensive reads, 30–60s TTL, with a coarse `invalidateEntityCache(entity)` called after every write.
6. A `Config` sheet (`key | value`) + `getConfig(key)` / `setConfig(key, value)` helpers — no hardcoded business rules (credit limits, SLA days, tax rates, etc. all live here).
7. Centralized `logError(functionName, error)` → writes to an `ErrorLog` sheet (id, ts, fn, message, stack). Never let a function throw silently.
8. Try/catch on every public handler via a `safeCall(fnName, fn)` wrapper — never leak raw stack traces to the frontend; always return `fail(e.message)`.
9. `PropertiesService` (Script Properties) for any API keys/secrets — never hardcoded in source.
10. Reusable data-access trio, used everywhere instead of ad-hoc range math: `getSheetAsObjects(sheetName)`, `appendRowFromObject(sheetName, obj)`, `updateRowById(sheetName, idColumn, id, updatedFields)`.
11. Batch writes — `setValues()` once for multi-row writes (e.g. CSV import); **never** call `appendRow()` inside a loop.
12. `getPagedData(entity, page, pageSize, filters, sortKey, sortDir)` — range-based reads for pagination, never pull the full sheet to compute a page.
13. `getCurrentUser()` / `requireRole(allowedRoles)` pattern for role gating on every write/sensitive-read function, backed by a `Users` sheet (`id | email | name | role | active`).
14. `logAudit(actor, action, entity, entityId, details)` → `AuditLog` sheet on every create/update/delete/stage-change, for traceability.
15. A single `bootstrap()` function the frontend calls once at startup, returning `{ me, config, lookups/enums, ... }` in one round trip instead of many small calls.
16. `onOpen()` installs a custom Sheets menu (e.g. "⚙️ Setup") with "Initialize System", "Reset Dummy Data", and any install/remove-trigger actions — this is the operator's setup UX, not just a dev convenience.

### Frontend (Index.html — single file, SPA-style, partials via `include()`)
17. `doGet()` serves one `HtmlService.createTemplateFromFile('Index')` template; role-based views are switched **client-side** via JS after `bootstrap()` returns `me.role` — do not build separate `doGet` routes per role unless a view must be reachable by an unauthenticated/public user (e.g. a public self-service form), in which case that one exception gets its own template file, everything else stays in `Index.html`.
18. Shared CSS/JS pulled into `Index.html` via `<?!= include('Name') ?>` partials — no duplicated style/script blocks.
19. Pagination controls with configurable page size (10/20/25/50/100).
20. Client-side page cache (a `Map` keyed by `entity+filters+page+size`) with auto-invalidation on writes and a soft TTL.
21. Background prefetch of the next page while the user views the current one.
22. Pick **one** loading-state pattern per project and apply it consistently: skeleton shimmer, animated fade/stagger with a top progress bar, or blur-to-sharp crossfade.
23. SweetAlert2 for all alerts/confirmations/toasts — no native `alert()`/`confirm()`.
24. Soft-delete + "Undo" toast pattern instead of hard deletes (a `deleted` boolean column, filtered out of reads, restorable).
25. A notification bell/center in the top bar for in-app alerts (e.g. SLA breaches, pending approvals, credit-limit warnings — whatever is relevant to the project).

### Deliverable format (non-negotiable)
Output **exactly two files**, matching this pattern precisely (plus optional named `.html` partial files only for shared CSS/JS via `include()`, e.g. `Styles.html`, `ClientUtils.html`):
- `Code.gs` — all backend logic, structured in this order: constants → `doGet`/`onOpen`/menu handlers → response wrappers → `initializeSheets`/seed functions/`resetDummyData` → data-access trio → config helpers → auth (`getCurrentUser`/`requireRole`) → audit/error logging → `bootstrap()` → pagination → entity CRUD functions (grouped by entity) → domain-specific business logic functions → triggers.
- `Index.html` — single SPA page, tab/view switching in client JS based on `me.role` from `bootstrap()`.

Include a top-of-file docblock comment in `Code.gs` exactly like this, adapted to the project:
```
/**
 * ============================================================================
 *  <PROJECT NAME> — Google Apps Script Web App (single-spreadsheet backend)
 * ============================================================================
 *  DEPLOYMENT
 *  ----------
 *  1. Open Google Sheets → Extensions → Apps Script.
 *  2. Paste this file as Code.gs and paste Index.html as an HTML file named
 *     exactly "Index".
 *  3. Reload the sheet; a custom menu appears. Click "Initialize System" once
 *     (safe to run again — idempotent).
 *  4. Deploy → New deployment → Web app → Execute as: Me, Access: <as needed>.
 *  SHEET STRUCTURE (created by initializeSheets)
 *  ---------------------------------------------
 *  <list every sheet with its exact column headers>
 * ============================================================================
 */
```

This structure must work if pasted as-is into **any** AI coding tool (Claude Code, Cursor, Windsurf, etc.) with zero extra context needed — the prompt is fully self-contained.

---

## Roles ("Dual-Console Gateway" — role field on Users sheet, gated via `requireRole`)

1. **Client Console** — a site supervisor's login, scoped to their own site only: places orders against the catalog at their site's negotiated rates, sees their own site's order/delivery/invoice history, raises complaints within the 24-hour window.
2. **Admin/Ops Console** — full access: catalog and site-wise pricing management, credit limit configuration and override, order consolidation and dispatch, invoice generation, complaint resolution and credit notes, Tally export, client/site directory management.

Both consoles live inside the single `Index.html` and switch client-side based on `me.role` from `bootstrap()` — the ordering flow for Client role and the ops flow for Admin role are just different view states, not separate deployments.

---

## Sheet Structure (for `initializeSheets()`)

- **Users** — `id | email | name | role | siteId | active | createdAt` (`role` = Client or Admin; `siteId` links a Client-console user to their site, blank for Admin)
- **Clients** — `id | name | gstin | homeState | creditLimit | createdAt`
- **Sites** — `id | clientId | siteName | address | state | contactPerson | contactPhone | createdAt`
- **Catalog** — `id | itemName | sku | uom | defaultRate | hsnCode | gstRatePct | createdAt`
- **SitePricing** — `id | siteId | itemId | rate` (site-specific override; falls back to `Catalog.defaultRate` if no row exists — this is the "Client Site-Wise Pricing Master")
- **Orders** — `id | siteId | status | createdAt | deleted` (`status`: `Placed | Consolidated | Dispatched | Delivered | Cancelled`)
- **OrderLines** — `id | orderId | itemId | qtyOrdered | qtyDelivered | damagedQty`
- **Deliveries** — `id | siteId | deliveryDate | orderIds (comma-separated or JSON) | status` — represents a consolidated batch of orders dispatched together (the "Consolidated Deliveries Optimizer")
- **Invoices** — `id | clientId | siteId | invoiceDate | subtotal | cgst | sgst | igst | total | tallyExported | createdAt`
- **InvoiceLines** — `id | invoiceId | orderLineId | itemId | qty | rate | amount`
- **Payments** — `id | clientId | amount | type (Payment/CreditNote/Debit) | reference | createdAt` (the running ledger used to compute outstanding balance for the Credit Limit Guard)
- **Complaints** — `id | orderLineId | raisedAt | deadline | status (Open/Resolved/Expired) | reason | resolutionNotes`
- **CreditNotes** — `id | complaintId | amount | issuedAt`
- **Config** — `key | value` (seller's home state for GST split logic, default credit-limit-override-requires-reason flag, complaint window hours = 24)
- **AuditLog**, **ErrorLog** — per playbook items 7 and 14.

Seed dummy data covering: 2 clients each with 2–3 sites in different states (to demonstrate CGST/SGST vs IGST split), site-specific pricing overrides on a subset of items, one order near its credit limit, one delivery with a damaged-quantity line, one open complaint inside its 24-hour window, one expired complaint, and one already-issued credit note — so every status is demoable immediately after `initializeSheets()`.

---

## Features to Build

### 1. Client Site-Wise Pricing Master
- When a Client-console user builds an order, line pricing looks up `SitePricing` for `(siteId, itemId)` first, falls back to `Catalog.defaultRate` if no override exists.
- Admin view: a pricing grid to set/edit per-site overrides against the catalog.

### 2. Dynamic Credit Limit Guard
- On order placement (or at dispatch — confirm which checkpoint is stricter for this business), compute `outstanding = SUM(Invoices.total for client) - SUM(Payments/CreditNotes for client)`.
- If `outstanding + this order's estimated value > Clients.creditLimit`, block the action inside a `LockService`-protected check and return a clear `fail()` message.
- Admin can override with a mandatory reason, which is written to `AuditLog`.

### 3. Consolidated Deliveries Optimizer
- Admin dispatch view groups all `Placed` orders by `siteId` + target delivery date into a single `Deliveries` row, so a site's multiple line requests become one physical delivery trip.
- On delivery confirmation, capture actual `qtyDelivered` and `damagedQty` per `OrderLines` row (these will differ from `qtyOrdered` — this is the reconciliation step the article calls out as the core month-end pain).

### 4. Live Demand Planner Engine
- A simple aggregation (not ML): sum/average ordered quantity per item over a trailing N-day window (configurable in `Config`), shown as a small chart/table on the Admin dashboard, to flag items trending up before stockouts happen.

### 5. Auto GST-Detect Invoicing Pipeline
- On invoice generation from delivered `OrderLines`: compare `Sites.state` (place of supply) against the seller's home state (`Config`).
- Same state → split tax into `CGST + SGST` using `Catalog.gstRatePct` (half each). Different state → apply the full rate as `IGST`.
- Write the resulting `Invoices`/`InvoiceLines` rows; this removes the manual state-by-state tax calculation the article describes as error-prone.

### 6. 24-Hour Strict Issue Desk & Auto Credit Notes
- A Client-console user can raise a `Complaints` row against a delivered `OrderLines` entry only while `now <= deliveredAt + 24 hours` (`Config`-driven window) — enforce this server-side, not just in the UI.
- On raising a complaint tied to a `damagedQty` line, auto-generate a proposed `CreditNotes` amount (`damagedQty × rate`) for Admin to confirm/adjust before it posts to `Payments` as a `CreditNote` entry and reduces the client's outstanding balance.
- After the 24-hour window, mark any un-actioned complaint attempt as blocked/expired rather than silently allowing it.

### 7. Native Tally Prime XML Exporter
- On an `Invoices` row, generate a Tally Prime-compatible Sales Voucher XML (standard `ENVELOPE > BODY > IMPORTDATA` structure with ledger name, GST breakdown, item allocations) and offer it as a downloadable file from the Admin invoice view.
- Mark `Invoices.tallyExported = true` after a successful export, to avoid duplicate imports.

---

## Open Decisions to Confirm Before/During Build
- Credit-limit check: enforce at order placement, at dispatch, or both?
- Tally export: exact XML schema/version target (Tally Prime XML format has changed across releases — confirm the version in use)?
- Demand Planner: trailing-window length default (e.g. 30 days)?
- Multi-currency or single-currency (assume INR-only unless told otherwise)?
