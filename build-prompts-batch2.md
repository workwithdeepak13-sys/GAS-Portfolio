# Build Prompts — Batch 2 (5 New SMB Systems, Google Apps Script)

> **How to use this file:** Each project section below is a complete build prompt.
> To build a project, copy the **Reusable Working-Style Playbook** section (immediately below)
> **plus** the one project section you want, and paste both into your AI coding tool
> (Claude Code, Cursor, Windsurf, etc.). The playbook + one project section together are
> fully self-contained — zero extra context needed.

---

## Reusable Working-Style Playbook (mandatory — apply exactly, every project)

### Backend (Code.gs)
1. `initializeSheets()` — idempotent setup: creates all sheets + headers + formatting if missing, seeds dummy data covering every status/edge case, returns a summary object. Safe to run repeatedly (skip sheets that already exist).
2. `resetDummyData()` — companion function that clears seeded rows for a clean demo state, without touching real rows (use an ID-threshold convention: seeded rows have low IDs, real rows added later get IDs > 1000, or similar — keep it consistent).
3. Standard response wrapper on **every** public function: `success(data)` / `fail(message)` — return `{ success, data, error }` always, never a bare value or a thrown error to the client.
4. `LockService.getScriptLock()` wraps every write operation (create/update/delete), released in a `finally` block.
5. `CacheService` (script or user cache) on expensive reads, 30–60s TTL, with a coarse `invalidateEntityCache(entity)` called after every write.
6. A `Config` sheet (`key | value`) + `getConfig(key)` / `setConfig(key, value)` helpers — no hardcoded business rules (rates, SLA hours, penalty percentages, tax rates, etc. all live here).
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
25. A notification bell/center in the top bar for in-app alerts (e.g. SLA breaches, expiring documents, overdue payments — whatever is relevant to the project).

### Deliverable format (non-negotiable)
Output **exactly two files** (plus optional named `.html` partials only for shared CSS/JS via `include()`):
- `Code.gs` — all backend logic, structured: constants → `doGet`/`onOpen`/menu handlers → response wrappers → `initializeSheets`/seeds/`resetDummyData` → data-access trio → config helpers → auth → audit/error logging → `bootstrap()` → pagination → entity CRUD (grouped by entity) → domain business logic → triggers.
- `Index.html` — single SPA page, tab/view switching in client JS based on `me.role` from `bootstrap()`.

Include a top-of-file docblock in `Code.gs`:
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

---
---

# PROJECT 1 — ServiceSarthi: AMC & Field Service Management System

## Project Overview

Build a Google Apps Script (GAS) web application called **ServiceSarthi** for small service businesses that maintain installed equipment at customer premises — AC servicing companies, water purifier (RO) dealers, lift/elevator AMC providers, CCTV/security installers, medical equipment servicers, diesel generator maintainers.

These businesses run on **Annual Maintenance Contracts (AMC)**: a customer pays yearly for N scheduled preventive services plus breakdown support with an SLA. Today the owner tracks contracts in a diary/Excel, forgets renewal dates (losing recurring revenue), misses scheduled visits (losing customer trust), can't prove which technician visited when (disputes), and has no idea which spare parts were consumed against which contract (margin leakage).

ServiceSarthi covers: contract lifecycle (quote → active → expiring → renewed/lost), auto-generated preventive maintenance visit schedules, breakdown ticket intake with SLA countdown, technician day-plan and mobile-friendly job closure with photo proof + customer OTP sign-off, spare parts consumption against tickets, and renewal-pipeline revenue forecasting.

## Roles (role field on Users sheet, gated via `requireRole`)

1. **Admin/Owner** — full access: contracts, customers, pricing, technician management, renewal pipeline, reports, config.
2. **Coordinator/Dispatcher** — creates tickets, assigns technicians, reschedules visits, follows up renewals; cannot edit contract pricing or delete records.
3. **Technician** — mobile-first view: today's assigned jobs, navigate (map link), start/close jobs, record parts used, capture photos + customer OTP; sees only their own jobs.
4. **Public exception**: a customer breakdown-request form (unauthenticated, reached via QR sticker on the machine — `doGet(e.parameter.page === 'request')` serves a separate `RequestForm.html`). Customer enters the equipment code printed on the sticker + issue description + phone; this creates a ticket in `Pending Assignment`.

## Sheet Structure (for `initializeSheets()`)

- **Users** — `id | email | name | role | phone | active | createdAt`
- **Customers** — `id | name | phone | email | address | area | gstin | createdAt | deleted`
- **Equipment** — `id | customerId | equipmentType | brand | model | serialNo | installDate | location | qrCode | createdAt | deleted` (`qrCode` = the short code printed on the sticker, e.g. `EQ-1042`)
- **Contracts** — `id | customerId | contractType | startDate | endDate | value | visitsIncluded | visitsUsed | slaHours | status | renewedFromId | createdAt | deleted`
  - `contractType`: `Comprehensive (parts included) | Non-Comprehensive (labour only) | Per-Visit`
  - `status`: `Draft | Active | Expiring Soon | Expired | Renewed | Lost`
- **ContractEquipment** — `id | contractId | equipmentId` (one contract covers multiple machines)
- **Tickets** — `id | contractId | equipmentId | type | source | issueDescription | priority | status | assignedTechId | scheduledDate | slaDeadline | startedAt | closedAt | resolutionNotes | photoUrls | customerOtpVerified | createdAt | deleted`
  - `type`: `Preventive (scheduled) | Breakdown | Installation | Inspection`
  - `source`: `Auto-Schedule | Coordinator | Customer QR Form | Phone`
  - `status`: `Pending Assignment | Assigned | In Progress | Completed | Cancelled | SLA Breached`
- **PartsCatalog** — `id | partName | partCode | uom | costPrice | sellPrice | stockQty | reorderLevel | createdAt | deleted`
- **PartsUsed** — `id | ticketId | partId | qty | billable | rate | amount | createdAt` (`billable` = false when the contract is Comprehensive — this is the margin-leakage tracker)
- **RenewalPipeline** — `id | contractId | expiryDate | followUpStage | nextFollowUpDate | quotedValue | outcome | notes | createdAt` (`followUpStage`: `Not Contacted | Quoted | Negotiating | Won | Lost`)
- **Config** — `key | value` (expiring-soon window days = 45, default SLA hours by priority, OTP length, visits auto-schedule spacing, service-due reminder days)
- **AuditLog**, **ErrorLog** — per playbook.

Seed dummy data covering: one Comprehensive and one Non-Comprehensive active contract, one contract expiring inside the 45-day window (so the renewal pipeline is populated), one expired-and-renewed chain (`renewedFromId` linkage), a preventive ticket scheduled for today, one breakdown ticket inside SLA, one SLA-breached ticket, one completed ticket with parts consumed (both billable and non-billable lines), and one open customer-QR-sourced ticket pending assignment.

## Features to Build

### 1. Contract Lifecycle & Auto Visit Scheduler
- On contract activation, auto-generate `visitsIncluded` preventive `Tickets` rows spaced evenly across the contract period (e.g. 4 visits/year → one per quarter), each with `type = Preventive`, `status = Pending Assignment`, `scheduledDate` pre-filled. Spacing logic is `Config`-driven.
- A daily time-based trigger flips `Active` contracts to `Expiring Soon` when `endDate - today <= Config.expiringSoonDays`, and to `Expired` past `endDate` — and auto-inserts a `RenewalPipeline` row when a contract enters `Expiring Soon` (skip if one already exists).

### 2. Breakdown Ticket Intake + SLA Countdown
- Ticket creation (any source) computes `slaDeadline = createdAt + slaHours` from the contract (fallback: `Config` default by priority).
- Dashboard shows live SLA countdown chips (green > 50% remaining, amber < 50%, red breached). A time-based trigger marks `SLA Breached` server-side and fires a notification-bell alert to Admin/Coordinator — SLA state must never depend only on client-side clocks.
- The public QR form resolves `qrCode → Equipment → active Contract` server-side; if no active contract covers the machine, still create the ticket but flag it `No Active Contract` so the coordinator can quote a per-visit charge (this is a real upsell moment, don't reject the request).

### 3. Technician Day Plan (mobile-first)
- Technician role lands on a card list of today's jobs sorted by scheduled time: customer name, address (as a `https://maps.google.com/?q=` link), equipment, issue, SLA chip.
- Job flow enforced server-side as a state machine: `Assigned → In Progress (startedAt stamped) → Completed (requires: resolutionNotes, ≥1 photo, customer OTP)`. Reject out-of-order transitions with `fail()`.
- Photo capture via `getUserMedia` → base64 → Drive folder → URL stored in `photoUrls` (JSON array).
- Customer OTP sign-off: on "Complete", server generates a short OTP, sends it to the customer's phone via `MailApp` (email) by default — build a pluggable `Backend_Notify` with an SMS/WhatsApp BSP scaffold behind `Config`, same reality-check pattern as GateDesk: GAS cannot send free SMS/WhatsApp; do not fake it. Technician enters the OTP the customer received; server verifies before closing. This is the dispute-proof "customer was present and satisfied" record.

### 4. Parts Consumption & Comprehensive-Contract Margin Guard
- When closing a ticket, technician selects parts used. If the contract is `Comprehensive`, lines default `billable = false` but still decrement `PartsCatalog.stockQty` and record `costPrice` — so the Admin dashboard can show **true cost per contract** (contract value minus labour estimate minus non-billable parts cost). This is the number that tells the owner which AMC customers are unprofitable at renewal time.
- Stock decrement inside `LockService`; block closure if `stockQty` would go negative unless Admin overrides with a reason (audited).
- Reorder alert in the notification bell when `stockQty <= reorderLevel`.

### 5. Renewal Pipeline & Revenue Forecast
- Kanban-style view of `RenewalPipeline` by `followUpStage` with drag/click stage moves, next-follow-up date sorting, and a "renew" action that clones the old contract into a new `Draft` with `renewedFromId` set (pricing editable before activation).
- Forecast widget: sum of `quotedValue` weighted by stage (`Config`-driven weights, e.g. Quoted 40%, Negotiating 70%) per month of expiry — the owner's next-quarter recurring-revenue view.
- Renewal conversion report: renewed vs lost count and value, trailing 12 months.

### 6. Service History per Machine
- From any `Equipment` row: full chronological ticket history (visits, breakdowns, parts, photos, who serviced it). This is the answer to "this machine keeps failing — show me its record" and the strongest sales artifact at renewal time.

### 7. Reports & CSV Export
- Technician productivity (jobs/day, avg time-to-close, SLA compliance %), contract profitability ranking, parts consumption by month, pending-visit backlog. All paginated + CSV export.

## Open Decisions to Confirm Before/During Build
- OTP delivery default: email-only (free) or wire a paid SMS/WhatsApp BSP from day one?
- Preventive visit auto-spacing: even split, or fixed calendar months chosen at contract creation?
- Should technicians see customer phone numbers, or route all calls through the coordinator?
- Per-visit (non-contract) tickets: quote/collect payment inside the app, or just record the amount?

---
---

# PROJECT 2 — FleetSarthi: Fleet Trip, Fuel & Compliance Management System

## Project Overview

Build a Google Apps Script (GAS) web application called **FleetSarthi** for small transporters and businesses running their own vehicles — 5 to 50 trucks/tempos/vans/buses: local goods carriers, school-bus operators, construction-material suppliers, distribution fleets, staff-transport contractors.

Their reality: trip records in a paper "trip sheet" book, diesel bills in a shoebox, driver cash advances remembered in the owner's head, and — most dangerously — insurance/permit/fitness/PUC expiry dates tracked nowhere, discovered only at an RTO checkpoint fine or a rejected insurance claim. Fuel is 40–55% of operating cost and the single biggest leakage point (inflated bills, siphoning), yet no one computes per-vehicle mileage (km/l) consistently enough to catch it.

FleetSarthi covers: vehicle master with document expiry tracking + auto-alerts, trip sheets with revenue and expense capture, fuel log with odometer-based mileage computation and anomaly flags, driver ledger (advances vs settlements), maintenance/service records with cost history, and per-vehicle profit & loss.

## Roles (role field on Users sheet, gated via `requireRole`)

1. **Admin/Owner** — everything: vehicles, documents, driver ledger settlements, P&L, config.
2. **Supervisor/Manager** — creates trips, enters fuel/expenses, closes trips, records maintenance; cannot settle driver ledgers or edit document masters.
3. **Driver** — mobile-first, own data only: sees assigned trip, submits start/end odometer + trip expenses with receipt photos, views own ledger balance. (No public/unauthenticated view in this project — all views live in `Index.html`.)

## Sheet Structure (for `initializeSheets()`)

- **Users** — `id | email | name | role | phone | licenseNo | licenseExpiry | active | createdAt` (drivers are Users with role Driver; license expiry participates in the compliance engine)
- **Vehicles** — `id | regNo | type | make | model | year | ownerName | currentOdometer | status | createdAt | deleted` (`status`: `Active | In Workshop | Sold | Inactive`)
- **VehicleDocs** — `id | vehicleId | docType | docNumber | issueDate | expiryDate | reminderDays | fileUrl | createdAt`
  - `docType` enum: `Insurance | Permit | Fitness Certificate | PUC | Road Tax | National Permit | Goods Carriage Permit` (extendable via Config)
- **Trips** — `id | vehicleId | driverId | customerName | route | startDate | endDate | startOdometer | endOdometer | kmRun | freightAmount | advanceReceived | balanceReceived | status | createdAt | deleted`
  - `status`: `Planned | Running | Completed | Settled | Cancelled`
- **TripExpenses** — `id | tripId | category | amount | paidFrom | receiptUrl | notes | createdAt` (`category`: `Toll | Loading/Unloading | Food | Parking | Police/RTO | Repair En-route | Other`; `paidFrom`: `Driver Advance | Company Cash | Driver Own Pocket`)
- **FuelLogs** — `id | vehicleId | tripId | date | station | litres | ratePerLitre | amount | odometerAtFill | paymentMode | receiptUrl | computedKmpl | anomalyFlag | createdAt`
- **DriverLedger** — `id | driverId | date | type | amount | reference | notes | createdAt` (`type`: `Advance Given | Expense Approved | Salary | Settlement Received | Deduction` — running balance computed, never stored denormalized without a recompute function)
- **Maintenance** — `id | vehicleId | date | type | odometer | workshop | description | partsCost | labourCost | totalCost | nextDueKm | nextDueDate | createdAt | deleted` (`type`: `Scheduled Service | Breakdown Repair | Tyre | Battery | Body Work | Other`)
- **Config** — `key | value` (default reminder days per doc type, mileage anomaly threshold %, expected km/l per vehicle type, financial-year start month)
- **AuditLog**, **ErrorLog** — per playbook.

Seed dummy data covering: 4 vehicles (one with insurance expiring in 10 days, one with PUC already expired, one in workshop), a driver with an outstanding advance balance, one running trip, one completed-but-unsettled trip, one settled trip with full expense lines, fuel logs producing one normal and one anomalous km/l reading, and a maintenance row with `nextDueKm` close to the vehicle's current odometer.

## Features to Build

### 1. Compliance Engine (the killer feature)
- Every `VehicleDocs` row has `expiryDate` + `reminderDays`. A daily time-based trigger computes three buckets: `Expired`, `Expiring ≤ reminderDays`, `OK`.
- Dashboard compliance board: vehicles × doc types grid with red/amber/green cells — the owner sees the whole fleet's legal standing in one screen.
- Alerts: notification bell + daily digest email (`MailApp`) to Admin listing everything expired/expiring, including **driver license expiries** from the Users sheet.
- Hard guard: block assigning a `Planned` trip to a vehicle with any `Expired` critical doc (which doc types are "critical" is `Config`-driven) unless Admin overrides with a mandatory reason → `AuditLog`. This is the feature that prevents the ₹2-lakh uninsured-accident story.

### 2. Trip Sheet Lifecycle
- `Planned → Running` (start odometer required; must be ≥ vehicle's `currentOdometer`, else `fail()` with the discrepancy shown) → `Completed` (end odometer required, `kmRun` computed, vehicle's `currentOdometer` advanced inside `LockService`) → `Settled` (freight balance received + all expenses approved).
- Trip P&L computed on completion: `freightAmount − (fuel attributed to trip + approved TripExpenses + a Config-driven per-km maintenance reserve)` — shown on the trip card so every trip answers "did we make money on this run?"
- Driver mobile view: current trip card with big Start/End odometer inputs and expense-photo submission; expenses land as pending until Supervisor approves (approval writes to `DriverLedger` if `paidFrom = Driver Own Pocket`).

### 3. Fuel Log & Mileage Anomaly Detection
- On each fuel entry with `odometerAtFill`, compute `computedKmpl = (odometerAtFill − previous fill's odometer) / litres` (tank-to-tank method; first fill for a vehicle gets no km/l).
- Compare against the vehicle's trailing-average km/l (window `Config`-driven): deviation beyond `Config.anomalyThresholdPct` sets `anomalyFlag = true` and raises a bell notification — this quietly surfaces bill inflation and siphoning without accusing anyone.
- Fuel dashboard: per-vehicle km/l trend chart, monthly fuel spend, cost-per-km ranking across the fleet (the worst vehicle is often the one to sell).

### 4. Driver Ledger
- Running balance per driver: advances given minus approved own-pocket expenses minus settlements. Balance shown to the driver (own row only) and to Admin (all).
- Settlement flow: Admin records `Settlement Received` / `Deduction` with reference; ledger history is append-only (corrections are new reversing entries, never edits — audit integrity).
- Alert when any driver's outstanding advance exceeds a `Config` ceiling.

### 5. Maintenance & Service Due Tracking
- Each maintenance entry can set `nextDueKm` / `nextDueDate`. Daily trigger + dashboard flag when a vehicle's `currentOdometer` approaches `nextDueKm` (within a `Config` margin) or the date nears.
- Per-vehicle maintenance cost history with cost-per-km trend — rising cost-per-km is the objective "time to replace this vehicle" signal.

### 6. Per-Vehicle P&L and Fleet Reports
- Monthly per-vehicle P&L: freight revenue − fuel − trip expenses − maintenance − (optional fixed costs from Config: EMI, driver salary allocation, insurance amortization).
- Fleet utilization: days-with-trips vs idle days per vehicle per month.
- All reports paginated, filterable by date range/vehicle/driver, CSV export.

## Open Decisions to Confirm Before/During Build
- Fixed costs (EMI, salaries) in per-vehicle P&L: include via Config allocations, or keep P&L variable-cost-only?
- Fuel entry: driver-submitted (with supervisor approval) or supervisor-only entry?
- Multi-day trips with multiple fuel fills: attribute fuel to trips by date range, or by explicit `tripId` selection at entry?
- GPS/odometer photo proof on trip start/end: require the odometer photo, or trust typed values?

---
---

# PROJECT 3 — VidyaSarthi: Coaching Institute Fee, Batch & Attendance Management System

## Project Overview

Build a Google Apps Script (GAS) web application called **VidyaSarthi** for small coaching institutes, tuition centers, and training academies — 50 to 1000 students: JEE/NEET coaching, spoken-English institutes, computer-training centers, music/dance academies, competitive-exam academies.

Their operating pain: fee collection is installment-based but tracked in a paper register, so nobody knows the true "total pending dues" number; parents are chased inconsistently (awkward for staff, leaky for revenue — 5–15% of billed fees typically slip through as silently-unpaid installments); attendance is a paper register that parents never see until a student has already disengaged; and enquiry-to-admission conversion is untracked, so marketing spend is blind.

VidyaSarthi covers: enquiry → admission pipeline, batch and course masters, installment-plan fee management with receipts and dues aging, automated fee-reminder engine, daily attendance with absence alerts to parents, exam/test score recording with progress reports, and a parent/student self-service view.

## Roles (role field on Users sheet, gated via `requireRole`)

1. **Admin/Owner** — everything: courses, batches, fee plans, discounts/waivers, staff, reports, config.
2. **Front Desk/Accountant** — enquiries, admissions, fee collection + receipt generation, follow-ups; cannot create courses or grant discounts beyond a `Config` ceiling.
3. **Teacher** — own batches only: mark attendance, enter test scores, view student lists.
4. **Parent/Student** — self-service, own record only: fee status + receipts, attendance summary, test scores, notices. Login via email matching the `Students.parentEmail` — no public unauthenticated views in this project.

## Sheet Structure (for `initializeSheets()`)

- **Users** — `id | email | name | role | phone | active | createdAt`
- **Courses** — `id | name | durationMonths | totalFee | description | active | createdAt | deleted`
- **Batches** — `id | courseId | name | teacherId | startDate | endDate | schedule | capacity | active | createdAt | deleted` (`schedule` = human string like "Mon-Wed-Fri 5-7pm")
- **Enquiries** — `id | name | phone | parentName | courseInterest | source | stage | followUpDate | notes | convertedStudentId | createdAt | deleted`
  - `source`: `Walk-in | Referral | Google | Instagram | Pamphlet | Other` — `stage`: `New | Follow-up | Demo Scheduled | Demo Done | Admitted | Lost`
- **Students** — `id | enquiryId | name | phone | parentName | parentPhone | parentEmail | batchId | admissionDate | status | photoUrl | createdAt | deleted` (`status`: `Active | Completed | Dropped | Suspended`)
- **FeePlans** — `id | studentId | courseId | totalFee | discount | discountReason | netFee | installmentsJson | createdAt` (`installmentsJson` = JSON array of `{n, dueDate, amount}` — generated at admission, editable by Admin only)
- **FeeInstallments** — `id | feePlanId | studentId | installmentNo | dueDate | amount | status | createdAt` (`status`: `Upcoming | Due | Overdue | Paid | Waived` — materialized rows, one per installment, so dues queries never parse JSON)
- **Payments** — `id | installmentId | studentId | receiptNo | amount | mode | reference | collectedBy | date | createdAt` (`mode`: `Cash | UPI | Card | Bank Transfer | Cheque`; `receiptNo` auto-sequenced from Config, gap-free)
- **Attendance** — `id | batchId | studentId | date | status | markedBy | createdAt` (`status`: `Present | Absent | Late | Leave`)
- **Exams** — `id | batchId | name | date | maxMarks | createdAt | deleted`
- **ExamScores** — `id | examId | studentId | marks | remarks | createdAt`
- **Notices** — `id | title | body | audience | batchId | publishedAt | createdAt | deleted` (`audience`: `All | Batch | Parents`)
- **Config** — `key | value` (receipt number prefix + next sequence, overdue grace days, reminder schedule offsets e.g. `-3,0,7` days relative to due date, front-desk discount ceiling %, absence-alert threshold, late fee rule)
- **AuditLog**, **ErrorLog** — per playbook.

Seed dummy data covering: 2 courses, 3 batches, enquiries in every stage (one converted with linkage to its student), students with fee plans in every installment state (`Upcoming/Due/Overdue/Paid/Waived`), one student with a discount + reason, payments across all modes with sequential receipt numbers, two weeks of attendance including a student with 3 consecutive absences, one exam with scores entered, and one published notice.

## Features to Build

### 1. Enquiry → Admission Pipeline
- Kanban of `Enquiries` by stage; follow-up date sorting with an "overdue follow-ups" filter; conversion action opens the admission form pre-filled and links `convertedStudentId` back.
- Source-wise conversion report (enquiries → admissions → revenue by `source`) — this makes marketing spend accountable.

### 2. Installment Fee Engine (the heart of the system)
- At admission: pick course → `totalFee` copied → optional discount (front-desk capped by `Config`, beyond that requires Admin, always with `discountReason` → audited) → choose installment split (equal-N or custom rows) → server generates `FeeInstallments` rows.
- A daily trigger transitions `Upcoming → Due` (on due date) `→ Overdue` (past due + grace days). Status transitions are server-owned; the UI never computes them.
- Collection screen: search student → open installments listed → collect (full or partial — partial creates a `Payments` row and splits the installment: paid portion recorded, remainder stays open with adjusted amount) → auto-generated receipt (printable HTML with receipt number, institute name from Config) → optional email of the receipt to `parentEmail` via `MailApp`.
- **Dues dashboard**: total outstanding, aging buckets (0–15 / 16–30 / 31–60 / 60+ days), batch-wise dues, defaulter list ranked by amount — the single number the owner never had.

### 3. Automated Fee Reminder Engine
- Daily trigger reads `Config` reminder offsets (e.g. 3 days before due, on due date, 7 days after) and sends templated emails to `parentEmail` via `MailApp` with student name, installment amount, due date, and total pending.
- Reality check (same pattern as GateDesk): free WhatsApp/SMS is not possible from GAS — build a pluggable `Backend_Notify` with email as the working default and a BSP (WhatsApp/SMS) scaffold behind `Config`. Log every reminder to a `RemindersSent` region in AuditLog details (or a dedicated sheet if cleaner) so no parent is double-pinged the same day.
- Escalation: after N reminders (Config) with no payment, flag the student on the defaulter list for a personal call — automation hands off to a human at the right moment instead of nagging forever.

### 4. Attendance + Absence Alerts
- Teacher view: pick batch + date → roster with one-tap Present/Absent/Late/Leave → single batched `setValues()` write.
- Consecutive-absence detector: on marking, if a student hits the `Config` threshold (e.g. 3 consecutive absences), alert Admin's bell + email the parent — early-warning for dropouts, which are the #1 silent revenue killer (a dropped student is 12 months of lost fees).
- Monthly attendance % per student, visible to the parent role.

### 5. Exams & Progress Reports
- Teacher enters scores for an exam (batch roster grid, batch write). Percentile/rank computed within the batch.
- Parent view: score history per subject/exam with a simple trend chart.
- Printable progress report per student (attendance % + exam trend + fee status footer).

### 6. Parent/Student Self-Service View
- Read-only: fee plan with paid/pending installments and downloadable receipts, attendance calendar, exam scores, published notices. Strictly scoped server-side to the logged-in parent's student(s) — never filter client-side only.

### 7. Reports & CSV Export
- Daily collection register (by mode, by collector — this is the cash-reconciliation sheet), monthly revenue vs dues, batch occupancy vs capacity, enquiry conversion funnel. Paginated + CSV.

## Open Decisions to Confirm Before/During Build
- Late fee: auto-add to overdue installments (Config rule), or manual-only?
- Receipt numbering: single global sequence, or per-financial-year reset (e.g. `2026-27/0001`)?
- One student in multiple batches/courses simultaneously — support now or later?
- Parent login: email-based Google auth only, or also a phone+OTP flow (requires paid SMS)?

---
---

# PROJECT 4 — RentSarthi: Equipment Rental & Hire Management System

## Project Overview

Build a Google Apps Script (GAS) web application called **RentSarthi** for equipment rental businesses — construction equipment (scaffolding, shuttering plates, mixers, vibrators), event equipment (tents, chairs, sound, lighting), tools & machinery hire, furniture rental, medical equipment rental (wheelchairs, oxygen concentrators, hospital beds).

Their pain is unique and badly served by generic inventory software: **the same physical item cycles out and back endlessly**, and money leaks at every joint — items go out without a signed record, come back late with nobody computing the extra days, return damaged or short with no photo proof of original condition, security deposits are refunded from memory (or disputed for weeks), and the owner cannot answer "where are my 500 shuttering plates right now?" across a dozen concurrent customer sites.

RentSarthi covers: rental item master with quantity pools, rental agreements (challans) with itemized dispatch, per-day/week/month rate engine with automatic duration billing, security deposit ledger, return processing with damage/shortage assessment against photo proof, overdue-return alerts, and item-utilization economics.

## Roles (role field on Users sheet, gated via `requireRole`)

1. **Admin/Owner** — everything: item catalog + rates, deposit refunds, damage waivers, reports, config.
2. **Store Operator** — creates rentals, dispatches, processes returns, records payments; cannot refund deposits or waive damage charges.
3. **Customer (optional read-only)** — sees own active rentals, running charges, deposit status. (Decide at build time whether to enable; all views stay in `Index.html`.)

## Sheet Structure (for `initializeSheets()`)

- **Users** — `id | email | name | role | phone | active | createdAt`
- **Customers** — `id | name | phone | address | idProofType | idProofLast4 | creditWorthy | notes | createdAt | deleted`
- **Items** — `id | itemName | category | uom | totalQty | availableQty | ratePerDay | ratePerWeek | ratePerMonth | replacementCost | active | createdAt | deleted`
  - Quantity-pool model: fungible items (500 identical plates) tracked as counts. `availableQty` is **derived-but-materialized**: updated inside `LockService` on every dispatch/return, plus a `recomputeAvailability()` repair function that rebuilds it from movement history (menu action) — never trust a running counter without a rebuild path.
- **Rentals** — `id | rentalNo | customerId | siteAddress | startDate | expectedReturnDate | status | depositAmount | notes | createdAt | deleted`
  - `status`: `Draft | Dispatched | Partially Returned | Closed | Overdue`
- **RentalLines** — `id | rentalId | itemId | qtyOut | qtyReturned | qtyDamaged | qtyLost | rateBasis | rate | dispatchPhotoUrls | createdAt` (`rateBasis`: `Day | Week | Month`)
- **Returns** — `id | rentalId | returnDate | processedBy | notes | createdAt` (a return event; lines below)
- **ReturnLines** — `id | returnId | rentalLineId | qtyReturnedOk | qtyDamaged | qtyLost | damageChargePerUnit | damageNotes | returnPhotoUrls | createdAt`
- **Charges** — `id | rentalId | type | description | amount | createdAt` (`type`: `Rent | Damage | Lost Item | Late Fee | Transport | Discount(-) | Other` — the rental's bill is the sum of its Charges rows, always reconstructable)
- **Payments** — `id | rentalId | customerId | type | amount | mode | reference | date | createdAt` (`type`: `Deposit In | Rent Payment | Deposit Refund | Deposit Forfeit`)
- **Config** — `key | value` (late-fee rule e.g. 1.5× daily rate after grace days, grace days, deposit % default of item replacement value, overdue alert offsets, rentalNo sequence)
- **AuditLog**, **ErrorLog** — per playbook.

Seed dummy data covering: fungible (plates, chairs) and low-count (mixer machine) items, one rental fully out and running, one partially returned with a damaged-quantity return line + photos, one overdue rental past its expected return, one closed rental with a deposit partially forfeited against damage and remainder refunded, and one draft rental — every status demoable immediately.

## Features to Build

### 1. Rental Agreement (Challan) & Dispatch
- Build a rental: customer → lines (item, qty, rate basis; rate auto-filled from Items, editable with audit) → deposit computed (`Config` % of replacement value, editable) → `Draft`.
- Dispatch action: validates `availableQty` per line inside `LockService` (block or allow partial per the operator's choice), decrements availability, stamps `Dispatched`, captures dispatch condition photos per line (`getUserMedia` → Drive → URLs), and renders a **printable dispatch challan** (HTML print view: rental number, customer, site, itemized lines, deposit, terms footer from Config). The signed challan is the legal artifact these businesses currently improvise on letterheads.

### 2. Rate Engine & Running-Charge Meter
- Duration billing per line: elapsed = dispatch date → (return date | today), charged on the line's `rateBasis` with part-period rounding rules from `Config` (e.g. any started day counts; weeks round up after N days — make the rules explicit and configurable, never implicit).
- Every open rental shows a live **running charges** figure (computed server-side, cached 60s) — customer-facing transparency and the operator's answer to "what do I owe so far?"
- On each return event, post crystallized `Rent` rows into `Charges` for the returned quantities' elapsed duration; remaining quantities keep accruing. This is the tricky part — partial returns split the billing timeline, and the Charges-row design keeps every rupee reconstructable.

### 3. Return Processing with Damage & Shortage Assessment
- Return flow: pick rental → per line enter `qtyReturnedOk / qtyDamaged / qtyLost` (validated: cannot exceed outstanding) → damage charge per unit (default from `Config` % of replacement cost, editable by Admin only beyond the default) → capture return-condition photos → server posts `Damage`/`Lost Item` Charges rows, restores `availableQty` by OK quantity only, and flips rental status (`Partially Returned` or, if nothing outstanding, ready-to-close).
- Side-by-side dispatch-photo vs return-photo view during assessment — this kills the "it was already broken" dispute.

### 4. Deposit Ledger & Settlement
- Deposit lifecycle as `Payments` rows: `Deposit In` at dispatch → at closure, settlement screen shows: total Charges − Rent Payments received = balance; deposit applied against balance → `Deposit Forfeit` (portion consumed) + `Deposit Refund` (remainder) rows, both Admin-only actions with audit.
- A rental cannot be `Closed` until: all quantities accounted (returned + damaged + lost = out) AND charges settled AND deposit dispositioned — enforce server-side as a closure checklist returned by the API, shown as a UI checklist.

### 5. Overdue Tracking & "Where Is My Stock" Board
- Daily trigger flags rentals past `expectedReturnDate` as `Overdue`, applies the `Config` late-fee rule as accruing `Late Fee` charges, alerts the bell, and emails the customer (pluggable `Backend_Notify`, email default — same BSP reality-check as other projects).
- **Stock deployment board**: for every item — total / available / out, and out-quantities broken down by customer + site with days-out. This one screen answers "where are my 500 plates?" and is the feature owners will pay for.

### 6. Utilization & Economics Reports
- Per item: utilization % (item-days rented ÷ item-days owned), revenue earned vs `replacementCost` (payback progress), damage/loss rate. Ranks which inventory to buy more of and which is dead capital.
- Customer ranking: revenue, damage rate, average days-late — feeds the `creditWorthy` flag.
- Monthly revenue by category, outstanding balances aging. Paginated + CSV throughout.

## Open Decisions to Confirm Before/During Build
- Serialized tracking (per-unit IDs with individual condition history) for high-value items alongside the quantity-pool model — now or phase 2?
- GST invoice generation on closure (rental services are taxable) — include an invoice module, or keep Charges + challan only?
- Transport/delivery charges: flat per rental, or per-trip lines?
- Customer read-only role: enable at launch?

---
---

# PROJECT 5 — JobSarthi: Manufacturing Job-Work Challan & Material Reconciliation System

## Project Overview

Build a Google Apps Script (GAS) web application called **JobSarthi** for small manufacturers and fabricators who send or receive materials for **job work** — outsourced processing stages: powder coating, electroplating, CNC machining, heat treatment, stitching/embroidery (garments), casting/forging, assembly. Nearly every small factory in an industrial cluster lives inside this send-process-return web daily.

The pain is structural: under GST (Section 143, Rule 45), material sent for job work moves on a **delivery challan without tax**, but it must return within **1 year (inputs) / 3 years (capital goods)** or the principal owes GST on it as a deemed supply — and quarterly **ITC-04** reporting requires exact challan-wise reconciliation. In practice, challans are paper books; nobody tracks what's still lying at which job worker; process scrap/wastage is never reconciled against agreed norms (a classic silent-theft channel); and at audit time the accountant reconstructs a year of movements from memory. This is a compliance time bomb combined with a material-leakage hole, and off-the-shelf software for it starts at enterprise prices.

JobSarthi covers both directions: **Principal mode** (we send material out for processing) and **Job Worker mode** (we receive others' material and process it) — challan issue/receipt, work-order rate management, receipt reconciliation with wastage-norm checking, ageing against statutory deadlines, job-work invoicing for labour charges, and ITC-04-ready reporting.

## Roles (role field on Users sheet, gated via `requireRole`)

1. **Admin/Owner** — everything: parties, items, process rates, wastage norms, deadline overrides, reports, config.
2. **Store/Dispatch** — creates outward challans, records inward receipts, enters reconciliation quantities; cannot edit rates/norms or close disputed reconciliations.
3. **Accounts** — views everything, generates job-work invoices, runs ITC-04 reports and CSV exports; cannot create/modify challans.

## Sheet Structure (for `initializeSheets()`)

- **Users** — `id | email | name | role | active | createdAt`
- **Parties** — `id | name | gstin | address | state | type | contactPerson | phone | createdAt | deleted` (`type`: `Job Worker | Principal | Both` — the same party master serves both directions)
- **Items** — `id | itemName | itemCode | uom | hsnCode | ratePerUom | itemType | createdAt | deleted` (`itemType`: `Input | Semi-Finished | Capital Goods` — drives the 1-year vs 3-year statutory clock)
- **Processes** — `id | processName | description | createdAt | deleted` (powder coating, machining, plating…)
- **WorkOrders** — `id | direction | partyId | processId | rateType | rate | wastageNormPct | validFrom | validTo | active | createdAt | deleted`
  - `direction`: `Outward (we are principal)` | `Inward (we are job worker)` — `rateType`: `Per Piece | Per Kg | Lumpsum`
  - The commercial agreement: what we pay (outward) or charge (inward) for a process, and the **agreed wastage norm %** — the reconciliation yardstick.
- **Challans** — `id | challanNo | direction | partyId | workOrderId | challanDate | vehicleNo | ewayBillNo | statutoryDeadline | status | createdAt | deleted`
  - `challanNo` auto-sequenced per direction per financial year from Config (e.g. `JW-OUT/26-27/0041`) — gap-free, since these numbers go on statutory documents.
  - `status`: `Open | Partially Received | Fully Received | Overdue Statutory | Closed Short (deemed supply)`
  - `statutoryDeadline` = challanDate + 1yr or 3yr by the line items' `itemType` (server-computed).
- **ChallanLines** — `id | challanId | itemId | description | qtySent | qtyReceivedOk | qtyScrapDeclared | qtyShort | createdAt`
- **Receipts** — `id | receiptNo | challanId | receiptDate | jobWorkerChallanRef | notes | createdAt` (one challan may be returned across multiple receipts — partial returns are the norm, not the exception)
- **ReceiptLines** — `id | receiptId | challanLineId | qtyOk | qtyScrap | qtyRejected | reconNotes | createdAt`
- **JobWorkInvoices** — `id | invoiceNo | direction | partyId | invoiceDate | challanIds | processedQty | rate | amount | gstRatePct | cgst | sgst | igst | total | createdAt | deleted` (labour/processing charges — GST applies to the job-work **service**, split CGST/SGST vs IGST by party state vs Config home state, same logic pattern as SupplyDesk)
- **Config** — `key | value` (home state, challan sequence counters, input deadline months = 12, capital goods months = 36, wastage tolerance buffer %, FY start month, overdue-warning offsets e.g. 90/30/7 days before statutory deadline)
- **AuditLog**, **ErrorLog** — per playbook.

Seed dummy data covering: two outward work orders (per-piece and per-kg with different wastage norms) and one inward, an open outward challan, one partially received (two receipt events), one fully received and cleanly reconciled, one reconciled **over the wastage norm** (flagged), one challan aged past its 90-day warning offset, one crossed the statutory deadline (`Closed Short (deemed supply)`), and one job-work invoice each direction (one intra-state CGST/SGST, one inter-state IGST).

## Features to Build

### 1. Outward Challan Issue (Principal mode)
- Create challan: party + work order → lines (item, qty, description) → server assigns gap-free `challanNo`, computes `statutoryDeadline` from the strictest line's `itemType`, renders a **printable GST-compliant delivery challan** (HTML print view with the Rule 55 mandatory fields: consecutive serial number, date, both parties' names/addresses/GSTINs, HSN, description, quantity, taxable value, place of supply, signature block, and the "Material sent for job work — not a supply — no tax payable" declaration).
- E-way bill number field (generated outside the system on the govt portal; recorded here for the audit trail).

### 2. Receipt & Reconciliation Engine (the core)
- Record receipts against a challan, line by line: `qtyOk / qtyScrap / qtyRejected`. Multiple receipts per challan; server maintains per-line running totals inside `LockService` and updates challan status (`Partially Received` → `Fully Received` when `received + scrap = sent` per every line).
- **Wastage-norm check on every receipt**: cumulative scrap % per line vs the work order's `wastageNormPct` + `Config` tolerance buffer. Over the norm → receipt saves but the line is flagged `Norm Exceeded`, bell alert to Admin, and the reconciliation cannot be *closed* (only Admin can close a flagged line, with a mandatory reason → audit). This turns the silent scrap-leakage channel into an exception report.
- Rejected quantities: track as pending-rework at the job worker (a re-challan is not needed — they still hold it against the original challan) until a later receipt clears them.

### 3. Statutory Deadline Ageing & Deemed-Supply Guard
- Daily trigger: for every `Open`/`Partially Received` challan, compute days-to-deadline; raise bell + email alerts at the `Config` offsets (90/30/7 days).
- Past deadline with quantities still out: flip status to `Overdue Statutory`, compute the **deemed-supply exposure** (outstanding qty × item rate × GST rate) and show it as a red rupee figure on the dashboard — the number that makes the owner act. Admin closure as `Closed Short (deemed supply)` records the exposure permanently for the accountant.
- Ageing dashboard: all material lying outside, grouped by party, with days-out and deadline countdown — "what of mine is sitting in whose factory" in one screen.

### 4. Inward Mode (we are the job worker)
- Mirror flow: record inward challans received from principals (their challan number as reference), track material held, record our dispatch-back events, and — critically — our **labour billing**: processed quantity × work-order rate → `JobWorkInvoices` with CGST/SGST vs IGST auto-split (party state vs `Config` home state).
- Held-material register: everything we currently hold belonging to others, by principal — our side of the same audit question.

### 5. Job-Work Invoicing & Party Ledger
- Generate processing-charge invoices from reconciled quantities (invoice references the challans it covers — `challanIds`). Printable invoice HTML.
- Simple party ledger: invoices raised vs payments noted (a `Payments`-style entry on the invoice, or keep it invoice-status-only — confirm at build), outstanding by party.

### 6. ITC-04-Ready Reporting
- Quarterly report generator: for a selected quarter, produce the two ITC-04 tables as paginated views + CSV export — **Table 4** (goods sent to job worker: challan no, date, party GSTIN, item, UQC, qty, taxable value) and **Table 5A** (goods received back / supplied from job worker's premises: original challan ref, receipt date, qty). Losses/wastage reported as required.
- Do **not** generate the government JSON upload format by default — produce clean CSVs the accountant maps into the GST portal offline tool; flag JSON generation as a phase-2 decision (the schema changes across portal versions).

### 7. Dashboard
- Direction-tabbed: material outside (value + qty + top parties + nearest deadlines) / material held for others / norm-exceeded flags / deemed-supply exposure total / this-quarter ITC-04 readiness status (any unreconciled challans in the quarter?).

## Open Decisions to Confirm Before/During Build
- One challan restricted to one `itemType` class (clean single deadline per challan) vs mixed lines with the strictest deadline — which matches your users' behaviour?
- Scrap disposition: does the job worker return scrap, retain it (value adjusted in labour billing), or is it sold at the job worker's end? All three exist in practice; affects reconciliation math.
- Party ledger: full payments module, or invoice status only (`Unpaid/Paid`)?
- Multi-stage chains (A sends to B, B sub-sends to C — legal under GST with conditions): support now or phase 2?
