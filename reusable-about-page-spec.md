# REUSABLE SECTION: In-App "About & System Logic" Page

> **How to use:** Paste this file along with any build prompt (and the UI skill).
> It adds a mandatory page to every app. The agent must treat every requirement
> here as part of the app's acceptance criteria — the app is NOT complete
> without this page fully populated with REAL content from the actual build.

---

## Purpose

Every app must ship with a built-in **"About this System"** page, reachable from
the last item in the sidebar navigation. This page is the living documentation of
the app: what it is, why it exists, and — most importantly — the EXACT logic it
runs on. It proves the depth of the system to any stakeholder (client, auditor,
MD, or portfolio reviewer) without them ever opening the code or the spreadsheet.

**Hard rule: zero placeholder content.** Every formula, threshold, sheet name,
column name, and trigger schedule shown on this page must match what the app
actually implements. If the app changes, this page's content must be generated
from the same constants/config the code uses (single source of truth — read
values from the Config sheet / CONSTANTS object, never duplicate them as static
text).

## Page Structure (7 sections, in this order)

### 1. What This System Is
- 3–4 sentence plain-language summary: the business problem, who uses it, and
  the single most important outcome it delivers.
- A "Built for" line listing the roles (e.g., Plant Head, QA Manager, Stores).
- Version number, build date, and environment (read from Config).

### 2. The Problem It Solves
- 4–6 bullet points of the concrete pains that existed before this system
  (missed deadlines, untracked costs, Excel chaos, blame cycles, etc.).
- One line each on the measurable cost of that pain.

### 3. How Data Flows (Architecture)
- A rendered diagram (inline SVG or styled HTML boxes with arrows, NOT an image)
  showing: User roles → Web App (HtmlService) → Server functions (.gs) →
  Google Sheets (each sheet named) → Automations (triggers) → Outputs
  (emails/PDFs/digests).
- Below the diagram, a table of every sheet: **Sheet name | What it stores |
  Key columns | Written by | Read by**.

### 4. Exact Business Logic (the heart of the page)
For EVERY calculated value, score, status, or automated decision in the app,
render a logic card containing:
- **Name** of the rule/metric (e.g., "OEE", "Dues Aging Bucket", "PM Due Date").
- **Exact formula** in readable notation, e.g.
  `OEE = Availability × Performance × Quality`, then each term expanded:
  `Availability = (Planned time − Downtime) / Planned time`.
- **Inputs**: which sheet + column each variable comes from.
- **Thresholds**: the actual configured values and what happens at each band
  (e.g., "≥ 85% green · 70–85% amber · < 70% red — values read live from
  Config!B12:B14").
- **Trigger/timing**: when it is computed (on edit, on submit, time-driven
  trigger at 07:00, etc.).
- **Edge cases**: what happens with missing data, zero denominators, backdated
  entries.
- Group cards by module. This section should typically contain 10–25 cards.

### 5. Automations & Triggers
A table of every trigger: **Trigger | Schedule | What it does | Who gets
notified | What it writes**. Include email/digest examples rendered as styled
previews.

### 6. Roles & Permissions
A matrix table: rows = every user role, columns = every module/action, cells =
Full / View / None. Must match the actual permission checks in code.

### 7. Assumptions & Limits (honesty section)
- Data volume limits (rows before performance degrades), concurrency notes,
  Apps Script quota limits that apply (email/day, trigger runtime).
- Business assumptions baked in (single currency, single plant, shift timings).
- What this system deliberately does NOT do, and the recommended upgrade path.

## Presentation Requirements

- Follows the UI skill fully: same tokens, typography, and density. The page is
  a single scrollable view with a sticky in-page section nav (left, 11px
  uppercase links) that highlights the current section.
- Formulas render in mono font inside bordered `--surface-2` blocks.
- Logic cards: title 14px/600, body 13px, threshold bands shown as the app's
  actual status pills.
- A "Print / Export as PDF" button (top-right) that produces a clean printable
  version via CSS `@media print` (nav hidden, cards not split across pages).
- Include a "Last verified against build" timestamp updated whenever the
  config-driven content is regenerated.

## Acceptance Checks

1. Pick any 3 numbers shown on any dashboard — this page must let a reader
   reproduce all 3 by hand from raw sheet data using the formulas shown.
2. Every sheet name/column referenced here exists in the actual spreadsheet.
3. Thresholds displayed change immediately when Config values are edited.
4. The page prints to at most ~8 clean A4 pages.
5. No sentence on the page could be copy-pasted into a different app unchanged
   (i.e., zero generic filler).
