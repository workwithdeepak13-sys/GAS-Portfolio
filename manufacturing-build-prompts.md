# Manufacturing Systems — Enterprise-Grade Build Prompts (Batch 3)

> 5 complete, enterprise-grade Google Apps Script applications solving real, chronic manufacturing problems.
> Each prompt is paste-ready and self-contained. Build each as a standalone GAS Web App.
>
> **Naming convention for this batch:** plain, professional "System" names. No brand suffixes.

---

## Shared Engineering Playbook (applies to ALL 5 prompts)

Every application in this batch MUST follow these rules:

1. **Two-file deliverable**: `code.gs` (server) + `index.html` (single-page client with embedded CSS/JS). No external hosting, no libraries that require billing.
2. **Google Sheets as the database.** Each entity gets its own sheet (tab). First row = headers. Server code must reference columns by header name (build a header-index map), never by hardcoded column letters, so users can reorder columns without breaking the app.
3. **Auto-bootstrap**: on first run, `setup()` creates every sheet with headers, a `Config` sheet, and realistic seeded demo data (minimum 25–40 rows per major entity) that demonstrates every state of every workflow — including edge cases (overdue items, rejected items, partially completed items).
4. **Config sheet drives all business rules** — thresholds, tolerance percentages, shift timings, escalation hours, email recipients. Zero hardcoded business values in code.
5. **Role-based access** via a `Users` sheet (email, name, role, department, active). Roles differ per app but always include at least: Admin, Manager, Operator/Executive, Viewer. The UI must hide/disable actions the current role cannot perform, AND the server must re-validate the role on every write (never trust the client).
6. **Every write is audited**: an `AuditLog` sheet capturing timestamp, user email, action, entity, record ID, old value → new value (JSON), and source screen.
7. **Concurrency safety**: use `LockService` around all writes. Generate IDs via a counter in Config with prefixes (e.g., `NCR-2026-0001`).
8. **Server API pattern**: one `doGet()` serving the SPA; all data operations via `google.script.run` calling granular server functions that return `{ok, data, error}` envelopes. Client must handle every error with a toast, never a silent failure.
9. **Dashboard-first UI**: the landing screen is a role-aware dashboard with KPI cards, trend indicators (vs. previous period), an exception list ("what needs my attention today"), and drill-down navigation. Managers should get answers in 10 seconds without opening a single sheet.
10. **Mobile-usable**: shop-floor screens (data entry, approvals) must work on a phone — large touch targets, minimal typing, dropdowns and steppers over free text.
11. **Automated engines via time-driven triggers**: escalation checks, digest emails, and status recalculation run on `ScriptApp` triggers created by `setup()`. All trigger runs are logged to a `JobRuns` sheet with duration and outcome.
12. **Email notifications** via `MailApp` with clean HTML templates: action-request emails (with deep links back into the app), escalation emails, and a daily 7 AM management digest. Respect a per-user notification preference column.
13. **Printable documents**: any formal document (inspection report, work order, maintenance job card) must render as a clean print-ready HTML view (A4, company header from Config, signature blocks).
14. **Search + filter + export** on every register screen: text search, status/date/department filters, and CSV export of the filtered result.
15. **Reality-check section**: end the build by documenting honest GAS limitations (6-min execution cap, ~30 simultaneous users practical ceiling, no true real-time push — use 60-second polling for boards) and the mitigations used.
16. **Zero-training design**: every screen has a one-line purpose statement at top; every form field has helper text; every status badge has a hover tooltip explaining what it means and what happens next.

---

## PROMPT 1 — Production Planning & Shop-Floor Execution System

**Domain:** Discrete/batch manufacturing — plan vs. actual, downtime, OEE.

Build a complete **Production Planning & Shop-Floor Execution System** as a Google Apps Script Web App, following the Shared Engineering Playbook above.

### The business problem
Production planning lives in a planner's personal Excel. The shop floor runs on verbal instructions and a whiteboard. Nobody knows, at any given moment: what was planned today, what has actually been produced, which machine is down and for how long, and whether today's dispatch commitment is safe. The daily production meeting is 45 minutes of "as per my data…" arguments between planning, production, and dispatch — because there are three versions of the truth.

### Core entities (sheets)
`Machines` (code, name, section, rated capacity/hr, status), `Products` (SKU, name, std cycle time per machine, std scrap %), `ProductionPlans` (plan ID, date, shift, machine, SKU, planned qty, priority, status), `ProductionEntries` (entry ID, plan ID, hour slot, actual qty OK, rejected qty, rework qty, operator, remarks), `DowntimeEvents` (machine, start, end, category [breakdown/changeover/no-material/no-operator/power/quality-hold], sub-reason, minutes, logged by), `Shifts` (name, start, end), `Users`, `Config`, `AuditLog`, `JobRuns`.

### Must-have functionality
1. **Weekly plan builder** (Planner role): grid of machines × days; assign SKU + qty per shift; system validates against rated capacity and flags overloaded slots in red before saving. Copy-last-week and drag-to-move conveniences.
2. **Hourly production logging** (Operator, mobile-first): operator selects their machine → sees the active plan → logs OK/rejected/rework quantities per hour with a large-button stepper UI. Late entries (logged >2 hours after the slot) are accepted but flagged.
3. **Downtime capture with reason tree**: stopping a machine requires category + sub-reason (two-level dropdown from Config). A machine cannot log production while an open downtime event exists on it. Downtime >X minutes (Config) auto-emails the section head; >Y minutes escalates to plant head.
4. **Live plant board** (auto-refreshing, wall-display friendly): every machine as a tile — green (running, on plan), amber (running, behind plan >10%), red (down, with minutes counter), grey (no plan). Shows plan vs. actual % per machine, per section, plant-wide.
5. **OEE engine** (nightly trigger): computes Availability × Performance × Quality per machine per shift from logged data; stores results in an `OEEResults` sheet; dashboard shows OEE trend per machine over 30 days with the worst 5 machines highlighted.
6. **Plan-vs-actual variance report**: per day/shift/machine/SKU — planned, produced, shortfall, downtime minutes by category, rejection %. One-click "morning meeting pack" — a single printable page replacing the 45-minute argument.
7. **Dispatch risk flag**: if any SKU's cumulative actual falls behind cumulative plan by more than the Config tolerance while a dispatch date approaches (from plan priority), the SKU appears in a "Commitment at Risk" panel on the dashboard.
8. **Changeover tracking**: changeover downtime is separated in analytics; report on average changeover time per machine per SKU-pair to expose sequencing waste.

### Roles
Admin, Planner, Section Head, Operator, Plant Head (viewer with full analytics).

### Seeded demo must show
Machines in all four board states, a shift that beat plan, a shift that missed plan due to a 90-minute breakdown, an overloaded plan slot, and 30 days of history so OEE trends render.

---

## PROMPT 2 — Quality Inspection, NCR & CAPA Management System

**Domain:** In-plant quality assurance — incoming, in-process, and final inspection with non-conformance and corrective-action workflows. (Explicitly NOT customer complaints — this is internal quality only.)

Build a complete **Quality Inspection, NCR & CAPA Management System** as a GAS Web App, following the Shared Engineering Playbook.

### The business problem
Inspection results live in paper registers. When a lot fails, the "rejection note" is a WhatsApp photo. Root-cause analysis happens verbally and is never verified, so the same defect returns every quarter. During customer audits (ISO 9001 / IATF-style), the quality head spends three panic-days reconstructing records. There is no data answering: which supplier, which machine, which defect type is actually costing us the most?

### Core entities (sheets)
`InspectionPlans` (item/SKU, stage [incoming/in-process/final], characteristics to check, spec min/max/target, sampling qty rule, method), `Inspections` (ID, date, stage, item, lot no, lot qty, sample qty, inspector, result [pass/fail/conditional], linked NCR), `InspectionReadings` (inspection ID, characteristic, observed values, within-spec flag), `NCRs` (ID, date, source inspection, item, defect code, qty affected, disposition [rework/reject/use-as-is/return-to-supplier], disposition approver, status, cost impact), `CAPAs` (ID, NCR ref, problem statement, root cause [5-Why fields: why1…why5], corrective action, responsible, target date, verification date, verified by, effectiveness check result, status), `DefectCodes` (code, description, category), `Suppliers`, `Users`, `Config`, `AuditLog`, `JobRuns`.

### Must-have functionality
1. **Inspection execution screen** (mobile-first): inspector picks stage + item → system loads the inspection plan → shows each characteristic with spec limits → inspector enters observed readings → out-of-spec readings highlight instantly → result auto-computed by sampling rule from Config (e.g., any critical failure = lot fail).
2. **Auto-NCR creation**: a failed inspection forces NCR creation in the same flow — defect code, affected qty, photos-optional note. The lot status becomes "Quality Hold" and the system emails the responsible department head.
3. **Disposition workflow with authority matrix**: disposition options require approval per Config matrix (e.g., "use-as-is" needs Quality Head + Plant Head; "rework" needs Section Head). Approvals happen in-app; pending dispositions >48h escalate.
4. **CAPA with enforced 5-Why**: a CAPA cannot move to "Action Defined" until all five Why fields are filled (or explicitly marked N/A with reason). Corrective actions get a responsible person and target date; overdue actions escalate; closed CAPAs require an effectiveness verification after N days (trigger creates the verification task automatically).
5. **Repeat-defect detection**: nightly trigger flags any defect code + item combination occurring ≥X times in Y days (Config) and auto-opens a "Repeat Defect Alert" requiring a mandatory CAPA — this is the mechanism that breaks the quarterly-repeat cycle.
6. **Supplier quality scorecard**: incoming inspection data rolls into per-supplier metrics — lot acceptance rate, PPM defective, top defect codes, trend. Exportable one-page supplier report for vendor meetings.
7. **Cost-of-poor-quality dashboard**: rejection cost + rework cost (rates from Config) aggregated by defect code, machine, supplier, and month — answers "what is quality actually costing us and where."
8. **Audit binder generator**: select a date range → system produces a printable index of all inspections, NCRs, and CAPAs with statuses and closure evidence — the 3-panic-day audit prep becomes a 3-minute export.

### Roles
Admin, Quality Head, Inspector, Section Head, Plant Head.

### Seeded demo must show
Passing and failing inspections at all three stages, an NCR in every disposition path, an overdue CAPA, a repeat-defect alert already fired, and two suppliers with visibly different scorecards.

---

## PROMPT 3 — Machine Maintenance & Spare Parts Management System

**Domain:** Preventive + breakdown maintenance with spares inventory and reliability analytics.

Build a complete **Machine Maintenance & Spare Parts Management System** as a GAS Web App, following the Shared Engineering Playbook.

### The business problem
Maintenance is 100% reactive. Preventive schedules exist in a forgotten laminated sheet. Breakdowns are reported by shouting across the shop floor; nobody records how long the machine was down or why. Critical spares are discovered missing at 2 AM during a breakdown. The maintenance head cannot answer: which machine is our most unreliable, what is breakdown actually costing, and are we doing PMs on time?

### Core entities (sheets)
`Machines` (code, name, section, criticality [A/B/C], make, model, install date), `PMTemplates` (machine/machine-type, task checklist items, frequency [daily/weekly/monthly/quarterly], est. duration, required spares/consumables), `PMSchedules` (generated instances: machine, template, due date, status, completed date, done by, checklist results), `BreakdownTickets` (ID, machine, reported by, reported at, problem description, severity, assigned to, response at, resolved at, root cause category, action taken, downtime minutes, spares consumed, status), `Spares` (code, name, category, machine compatibility, current stock, min level, location bin, unit cost, lead time days), `SpareTransactions` (issue/receipt/return, spare, qty, against ticket/PM, by whom), `Users`, `Config`, `AuditLog`, `JobRuns`.

### Must-have functionality
1. **PM auto-scheduler**: nightly trigger generates upcoming PM instances from templates. Dashboard shows PM compliance % (done-on-time / due) per machine and per technician — the single number the maintenance head owns.
2. **PM execution with checklist** (mobile): technician opens today's PM → ticks checklist items → any item marked "abnormal" spawns a follow-up ticket automatically → records spares consumed (auto-deducts stock).
3. **Breakdown ticket lifecycle**: anyone can raise a ticket in 20 seconds (machine + problem + severity). Auto-assignment by section from Config. Timestamps captured at report → response → resolution give true MTTR. Severity-1 tickets on A-criticality machines immediately email + create an escalation timer (Config minutes) to the maintenance head.
4. **Downtime cost meter**: each machine has an hourly loss rate in Config; every ticket displays its running cost while open and final cost when closed. Dashboard shows monthly breakdown cost by machine and by root-cause category.
5. **Reliability analytics** (nightly): MTBF and MTTR per machine, 90-day trend, worst-5 machines panel, repeat-failure detection (same machine + same root cause ≥X times in Y days → flagged "Chronic Issue" requiring engineering review sign-off).
6. **Spares min-level engine**: stock deductions against tickets/PMs; below-min items appear in a reorder panel with suggested qty (min level + lead-time buffer); weekly reorder email to purchase. Issue history answers "where did all the bearings go."
7. **Spares criticality cross-check**: flag any A-criticality machine whose compatible critical spares have zero stock — the "2 AM discovery" prevented at 2 PM.
8. **Machine health card**: printable per-machine one-pager — PM compliance, breakdown history, MTBF/MTTR, total maintenance cost YTD, open issues. Used for repair-vs-replace decisions.

### Roles
Admin, Maintenance Head, Technician, Production User (can raise tickets + view own machines), Plant Head.

### Seeded demo must show
Overdue PMs, on-time PMs, an open severity-1 breakdown with a running cost meter, a chronic-issue flag, spares below min level, and one A-criticality machine with a zero-stock critical spare.

---

## PROMPT 4 — Tooling, Die & Fixture Lifecycle Management System

**Domain:** Tool-room management for press shops, molding units, machining, and forging — an almost entirely unserved software niche.

Build a complete **Tooling, Die & Fixture Lifecycle Management System** as a GAS Web App, following the Shared Engineering Playbook.

### The business problem
Dies, molds, jigs, and fixtures are the most expensive assets after machines — and the least managed. Nobody tracks how many shots/strokes a die has run against its rated life. Dies fail mid-production run, killing a delivery commitment. Tools go for outside refurbishment and disappear for weeks with no follow-up. When a customer asks "is my die still in good condition?" (customer-owned tooling is common), the answer is a guess. Tool location itself is tribal knowledge — finding a die takes an hour.

### Core entities (sheets)
`Tools` (tool ID, description, type [die/mold/jig/fixture/gauge], owner [company/customer name], linked products, rated life shots, life consumed shots, last PM shots, PM interval shots, current location [rack bin/machine/vendor], condition status, acquisition cost, date), `ToolMovements` (tool, from, to, movement type [issue-to-machine/return-to-rack/send-to-vendor/receive-from-vendor], by, date, remarks), `UsageLogs` (tool, date, machine, shots/strokes run, source [production entry], logged by), `RefurbJobs` (ID, tool, vendor, sent date, expected return, actual return, cost, work description, status, condition after), `ToolPMs` (tool, due-at shots, done date, checklist result, done by), `Gauges` (as tool type with calibration due date, certificate no, agency), `Users`, `Config`, `AuditLog`, `JobRuns`.

### Must-have functionality
1. **Tool master with life meter**: every tool shows a visual life bar — consumed vs. rated shots — with bands: green (<70%), amber (70–90%), red (>90%). Red tools appear on the dashboard "Plan replacement/refurb NOW" panel.
2. **Shot-count accumulation**: production quantities logged against a tool (simple entry: tool + date + shots, or bulk paste) accumulate its life counter and its since-last-PM counter. PM due when since-last-PM ≥ interval → auto-task created.
3. **Issue/return workflow with location truth**: a tool must be formally issued to a machine and returned to a rack bin. The system always answers "where is tool X" in one search. A tool cannot be issued to two places; return requires condition status entry (OK / needs attention / damaged — damaged forces a follow-up job).
4. **Refurbishment tracking**: send-to-vendor creates a RefurbJob with expected return date; overdue jobs escalate by email; receiving back requires cost + condition entry; refurb history and cumulative refurb cost per tool visible on its card (feeds repair-vs-retire decisions: flag tools whose cumulative refurb cost > X% of replacement cost, from Config).
5. **Customer-owned tooling register**: filter/report of all customer tools with condition, life consumed, location, and last-used date — the professional one-page answer when the customer audits their tooling. Idle customer tools (unused > N days) flagged, since customers bill for lost tooling.
6. **Gauge calibration module**: gauges carry calibration due dates; expiring within Config window → amber list + email; expired gauges show "DO NOT USE" status and are excluded from issue.
7. **Die-failure risk panel**: cross-references red-life tools against upcoming production plans (simple linked sheet or manual flag) — "these 4 dies are >90% life AND scheduled to run next week" — the mid-run failure prevented before it kills a dispatch.
8. **Tool room dashboard**: tools by status, life-band distribution, tools at vendor (with overdue count), calibration alerts, this-month refurb spend, top-10 most-used tools.

### Roles
Admin, Tool Room In-charge, Production Supervisor (issue/return + shot logging), Quality (gauge calibration), Plant Head.

### Seeded demo must show
Tools in every life band, one tool overdue at a vendor, a damaged return with follow-up job, an expired gauge, a customer-owned tool idle 200 days, and one red-life die flagged against next week's plan.

---

## PROMPT 5 — Order Profitability & Delivery Commitment Truth System

**Domain:** The MD's chronic, unsolved pain. Order-level margin truth + commitment accountability across departments.

Build a complete **Order Profitability & Delivery Commitment Truth System** as a GAS Web App, following the Shared Engineering Playbook.

### The business problem — the one the MD is tired of
Every manufacturing MD lives this loop: **"We are running full capacity, sales are growing — so why is the bank balance not growing? And why does every delivery date slip with everyone blaming everyone else?"**

He has asked for the answer for years and never gotten it, because:
- Quotes are priced on a 3-year-old cost sheet. Actual material prices, actual rejection rates, actual overtime, and actual freight are never traced back to the order. **He genuinely does not know which orders make money and which quietly lose it.** Accounting gives him a company-level P&L 45 days later — useless for decisions.
- Delivery commitments are made verbally by sales, adjusted verbally by planning, missed silently by production, and discovered by the customer. In the review meeting, sales blames planning, planning blames stores, stores blames purchase, purchase blames finance. **No system records who committed what, when, and what actually happened** — so accountability is impossible and the same failure repeats.

Existing ERP modules technically "have" this, but they require costing discipline and data entry rigor no mid-size plant achieves — so the modules stay empty and the MD stays blind. This system is designed to work with the data a real plant CAN capture.

### Core entities (sheets)
`Orders` (order ID, customer, SKU/description, qty, quoted unit price, order value, quoted cost basis [material/labor/overhead per unit as quoted], committed dispatch date, commitment made by, status), `CommitmentEvents` (order, event type [initial-commitment/revision/internal-slip-alert/customer-informed], old date → new date, reason category, responsible department, entered by, timestamp), `CostCaptures` (order, cost head [material/subcontract/rejection-loss/rework-labor/overtime/freight/penalty/other], amount, reference doc, date, entered by), `MilestoneGates` (order, gate [material-ready/production-start/production-complete/QC-clear/dispatch], planned date, actual date, owner department, delay reason if late), `Customers`, `Departments`, `Users`, `Config`, `AuditLog`, `JobRuns`.

### Must-have functionality
1. **Order margin ledger — quoted vs. actual, live**: every order carries its quoted cost stack. As actual costs are captured against it (simple, one-minute entries by the relevant department: stores posts material issue value, HR posts overtime hours × rate, QC posts rejection loss, dispatch posts freight), the system shows a live margin bar: quoted margin vs. current actual margin. The moment actual margin falls below the Config alert threshold, the order turns amber; below zero, red + MD email. **The MD sees loss-making orders while they can still be acted on, not 45 days later.**
2. **Margin leak analytics**: aggregate variance by cost head across all orders — "we lost ₹18L this quarter and 61% of it is rejection-loss on two SKUs, 24% is freight on rush dispatches." Ranked list of most and least profitable customers and SKUs by ACTUAL margin, not quoted margin. This single screen reprices the next quotation round.
3. **Commitment ledger — the accountability engine**: every dispatch date and every revision is a recorded event with a named person, a reason category, and a responsible department. Nobody can change a date silently. The order's commitment history reads like a timeline: who promised, who slipped it, why, and whether the customer was informed.
4. **Milestone gate tracking with early-warning**: each order has 5 gates with planned dates back-calculated from the committed dispatch date (offsets in Config). A gate going late triggers an internal slip alert to the NEXT gate's owner and to the MD digest — the slip is visible weeks before the dispatch date, when recovery is still possible, instead of the day after.
5. **Blame-proof departmental scorecard**: monthly per-department metrics from gate data — on-time gate %, average delay days, top delay reasons. The review meeting stops being an argument because the data is signed, timestamped, and reason-coded by the departments themselves.
6. **MD's 7 AM digest** (the centerpiece — one email, phone-readable): (a) orders that turned red/amber on margin yesterday, (b) commitments at risk in the next 7 days with the currently-late gate and owner, (c) any silent slip (gate late >2 days with no reason entered — the worst signal), (d) this month's actual-margin % vs. quoted, (e) top 3 margin leaks. Five sections, thirty seconds, complete truth.
7. **Customer-facing commitment report**: for any order, a clean printable status letter — current confirmed date, stage, and history of communicated changes — so the customer hears about a slip from the company, never the other way around.
8. **Won/lost margin review**: on order closure, final actual margin locks and the variance-vs-quote is stored; quarterly report shows quote-accuracy trend per estimator — closing the loop back into better pricing.

### Roles
Admin, MD (full visibility, digest recipient), Sales, Planning, Stores, Production, QC, Dispatch — each department can only post costs/gates it owns (department-to-cost-head mapping in Config).

### Seeded demo must show
A healthy-margin order, an order that turned red from rejection losses, an order with 3 commitment revisions across 2 departments, a silent slip currently flagged, a customer whose actual margin is negative despite healthy quoted margins, and a full month of gate data so the departmental scorecard renders.

---

## Suggested Build Order

1. **Prompt 5 (Order Profitability & Commitment Truth)** — the highest-impact, MD-level flagship; pure workflow + analytics, no hardware dependency.
2. **Prompt 1 (Production Planning & Shop-Floor Execution)** — the operational backbone; teaches OEE and plan-vs-actual discipline.
3. **Prompt 2 (Quality/NCR/CAPA)** — deep workflow engineering (approval matrices, enforced 5-Why, effectiveness loops).
4. **Prompt 3 (Maintenance & Spares)** — reliability math (MTBF/MTTR) plus inventory logic.
5. **Prompt 4 (Tooling Lifecycle)** — the niche differentiator almost no portfolio has.
