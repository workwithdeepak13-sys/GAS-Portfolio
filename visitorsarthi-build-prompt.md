# Build Prompt: GateDesk (VisitorSarthi Clone) — Gate Register Digitization System

> Paste this entire prompt into your AI coding tool (Claude Code, Cursor, Windsurf, etc.). It is fully self-contained.

## Project Overview

Build a Google Apps Script (GAS) web application called **GateDesk** that replaces a manual paper visitor register at a factory/office/warehouse gate with a digital visitor management system.

Core workflow: a visitor scans a QR code at the gate → fills a mobile self-registration form (with photo capture) → the host is notified and approves/rejects → guard/reception sees a live "who's inside" dashboard → admin gets full history, audit search, and CSV export.

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

## Roles (role field on Users sheet, gated via `requireRole`)

1. **Guard/Reception** — live queue, manual check-in/out, phone-number search.
2. **Host (employee)** — sees only visits addressed to them, approves/rejects, own visitor history.
3. **Admin** — full dashboard, all history, CSV export, host directory management, QR code display, config.

**Public exception**: the visitor self-registration form is unauthenticated (visitors use their own phone, no Google login). This is the one view that gets its own template — everything else (Guard/Host/Admin) lives inside `Index.html` and switches by role.

---

## Sheet Structure (for `initializeSheets()`)

- **Users** — `id | email | name | role | active | createdAt`
- **Visitors** — `id | phone | name | company | photoUrl | idProofLast4 | createdAt` (master directory, enables returning-visitor auto-fill by phone lookup)
- **VisitLogs** — `id | visitorId | hostId | purpose | checkInTime | checkOutTime | durationMins | status | approvalMethod | source | createdAt | deleted`
  - `status` enum: `Pending Approval | Approved | Rejected | Inside | Checked Out`
  - `source` enum: `QR Self-Registration | Guard Manual Entry`
- **Hosts** — `id | name | phone | department | active`
- **Config** — `key | value` (site name, QR target URL, working hours, notify-enabled flag)
- **AuditLog**, **ErrorLog** — per playbook items 7 and 14.

Seed dummy data covering: a visitor mid-visit ("Inside"), one checked out with a full duration, one pending approval, one rejected, and one returning visitor (same phone, two VisitLogs rows) — so the auto-fill and history search features are demoable immediately after `initializeSheets()`.

---

## Features to Build

### 1. QR Self-Registration Form (public, unauthenticated, own template)
- `doGet(e)` checks `e.parameter.page === 'checkin'` and serves a separate `CheckinForm.html` template instead of `Index` — this is the one intentional exception to the single-page rule, because visitors have no Google identity to key views off of.
- Phone field (10-digit) → on blur, server call checks `Visitors` sheet by phone; if found, auto-fill name/company/photo (editable).
- Live camera capture (`getUserMedia`) → base64 → save to a Drive folder → store URL. Optional ID proof: capture last-4-digits only as a typed field — **do not store a photographed ID image as a default**; flag this as an explicit decision point back to me if full ID image storage is wanted instead.
- Host dropdown (searchable, from `Hosts`) + purpose field.
- On submit: upsert `Visitors`, create `VisitLogs` row with `status = Pending Approval`, trigger host notification.

### 2. Host Notification & Approval
- **Reality check**: GAS cannot send free-form WhatsApp messages without a paid BSP (Gupshup/Interakt/Twilio/Meta Cloud API). Build a pluggable `Backend_Notify` module with two swappable implementations, chosen via `Config`:
  - **Option A (default, free)**: `MailApp`/`GmailApp` notification to the host with visitor details + a link to the in-app approve/reject view. Optionally include a `wa.me` deep link.
  - **Option B (real WhatsApp, scaffold only)**: `UrlFetchApp` call to a BSP REST API with an approve/reject template, plus a `doPost(e)` webhook to receive the host's reply and update `VisitLogs.status`.
- In-app approve/reject view reachable by the host (and by Admin acting on a host's behalf) inside `Index.html`.

### 3. Guard/Reception View (inside Index.html, role = Guard)
- Live counts (Today's Total / Currently Inside / Checked Out / Pending Approval) via `CacheService`, refreshed on action + 60s poll.
- Manual check-in for visitors without smartphones.
- One-tap check-out from the "Inside" list.
- Phone-number search box → paginated `VisitLogs` history (the "5-second audit" feature from the article).

### 4. Admin View (inside Index.html, role = Admin)
- Full paginated history with filters (date range, host, status), CSV export.
- Host directory CRUD.
- Duration analytics (avg. visit duration, busiest hours).
- QR code display/print page (render the check-in URL via a QR chart API).

### 5. Offline Resilience (adapted for GAS's real constraints)
- **Reality check**: a GAS web app cannot execute without reaching the server at all — there's no true offline execution like a native/PWA app would have with a service worker.
- Build the closest practical equivalent: on submit failure, queue the payload in `localStorage`; retry on the browser's `online` event and every 30s; show a persistent "Offline — N records pending sync" banner.
- Flag to me if a true installable PWA with a service worker is wanted instead — that's a bigger scope decision, don't default into it.

### 6. Automated Duration Tracking
- On check-out: `durationMins = checkOutTime - checkInTime`, written back to the row.
- Optional daily trigger (`installable time-based trigger`, per playbook item 16's menu pattern) to flag anyone still "Inside" past a configurable `Config` cutoff, for Admin review.

---

## Open Decisions to Confirm Before/During Build
- Notification: Option A (free) or Option B (paid WhatsApp BSP)?
- ID proof: last-4-digits only, or full photographed ID image too?
- True offline PWA vs. simple localStorage queue?
- Single-gate or multi-gate/multi-site support?
