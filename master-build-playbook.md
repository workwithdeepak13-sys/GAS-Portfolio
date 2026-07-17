# MASTER PLAYBOOK: Getting the Best Out of Every Build

> This is YOUR operating manual — not pasted into prompts. It tells you how to
> combine the files in this repo, how to run a build session, how to judge the
> output, and how to turn each finished app into portfolio and client value.

---

## 1. The File System (what goes where)

| File | Role | Paste into agent? |
|---|---|---|
| `*-build-prompts*.md` | The WHAT — one project's full spec | Yes — one project at a time |
| `skill-ui-design.md` | The LOOK — anti-slop UI law | Yes — with every prompt |
| `reusable-about-page-spec.md` | The PROOF — in-app living documentation | Yes — with every prompt |
| `*-why-these-problems-matter.md` | The WHY — sales/portfolio narrative | No — for you and clients |
| `master-build-playbook.md` | The HOW — this file | No — for you only |

**Paste order in a fresh chat (order matters):**
1. `skill-ui-design.md` (sets the visual law first)
2. `reusable-about-page-spec.md` (sets the documentation law)
3. ONE project prompt (never two — quality collapses when agents split attention)
4. One closing line: *"Follow the two skill files above as hard constraints.
   Build the complete system. Do not simplify, do not stub, do not skip the
   About page."*

## 2. The Build Session Protocol

**Phase 1 — Skeleton (first response).** Ask the agent to output ONLY: the sheet
schemas, the Config sheet contents, the module list, and the trigger list.
Review it yourself before any code. This 5-minute review kills 80% of rework.
Check: Are all thresholds in Config (not hardcoded)? Does every module map to a
sheet? Is anything from the prompt silently missing?

**Phase 2 — Build in module order.** Tell it to build in this exact order:
1. Config + sheet setup + seed data function
2. Auth/roles + app shell (sidebar, top bar, routing)
3. Core transaction module (the one that creates the most rows)
4. The calculation/engine module
5. Dashboards + digests + triggers
6. About & System Logic page (LAST — so it documents reality, not intention)

**Phase 3 — The Hostile Review.** Before accepting, run these prompts verbatim:
- *"List every feature from the original prompt and mark each one: fully built /
  partially built / missing. Do not be generous."*
- *"Open the app as a brand-new user with an empty spreadsheet. What breaks?"*
- *"Show me 3 numbers on the dashboard and trace each to its raw rows and
  formula. If you cannot, fix the calculation."*
- *"Run the Self-Review Gate from the UI skill and report each answer honestly."*

**Phase 4 — Seed & screenshot.** Run the seed function, open every page, and
take screenshots at 1366×768. If any page looks empty or sloppy, it goes back.

## 3. Red Flags That Mean "Reject and Redo"

- Any hardcoded threshold that the prompt said belongs in Config
- `TODO`, `// implement later`, or functions that return fake static data
- A dashboard number you cannot reproduce by hand from the sheets
- The About page containing a sentence that could describe any other app
- Tables that re-fetch the entire sheet on every action (quota killer)
- Business logic living in the client-side JS instead of `.gs` server code
- Purple. Anywhere.

## 4. Testing Matrix (run once per app, 30 minutes)

| Test | How |
|---|---|
| Role isolation | Log in as lowest role — verify restricted modules are invisible, not just disabled |
| Concurrency | Two browser tabs, same record, conflicting edits — verify no silent overwrite |
| Bad input | Negative qty, future dates, blank required fields, 5000-char text |
| Empty state | Fresh spreadsheet, no seed — every page must render, not crash |
| Scale smoke | Seed 2,000 transaction rows — dashboard must load < 5s |
| Trigger dry-run | Manually run each time-driven function — verify emails/writes once, not double |
| Print | Export the About page to PDF — clean pages, no cut cards |

## 5. Turning Builds into Portfolio Value

For every finished app, produce a **one-page case study** (reuse the matching
`why-these-problems-matter` section):
1. Problem (2 sentences, with the money at stake)
2. Solution (3 bullets: the engine, the enforcement, the visibility)
3. One dashboard screenshot + one About-page logic card screenshot
4. "Built with: Google Apps Script, Sheets, HtmlService — zero licence cost"
5. The customization hook: "Every threshold, role, and rule is config-driven —
   adapted to a new client in days, not months."

Order your portfolio by business sophistication, ending with the MD-level
profitability/commitment system — it is the strongest closer.

## 6. Client Conversation Cheat Sheet

- Lead with the pain, not the tech: open with section 2 of the why-file, ask
  "which of these is costing you the most right now?"
- Demo the About page early — it disarms "is this just a spreadsheet?" instantly.
- Price on outcome bands, not hours: recovered margin, saved penalties,
  audit-readiness. The why-files contain the numbers to anchor on.
- Always scope Phase 1 to ONE module live in 2 weeks. Land, prove, expand.
- Objection "we'll outgrow Sheets": agree — say the system is designed with an
  explicit upgrade path (section 7 of the About page) and the workflows/logic
  transfer to any future stack; the thinking is the asset, not the spreadsheet.

## 7. Maintenance & Versioning Discipline

- Keep one master copy of each app; client copies are made from it. Never edit
  a client copy's logic directly — change master, then port.
- Version format `v<major>.<minor>` stored in Config and shown on the About
  page. Bump minor for logic/threshold changes, major for new modules.
- Keep a `CHANGELOG` sheet in each spreadsheet: date | version | change | who.
- Quarterly: re-run the Testing Matrix on the master copy against current
  Apps Script quotas (they change).

## 8. When to Say No to Apps Script

Recommend a different stack (and win trust) when the client needs:
- > ~50 concurrent writers or sub-second real-time updates
- > ~200k rows in any hot table
- Offline-first factory-floor terminals
- Hard regulatory data residency / audit trails beyond Sheets' revision history
Everything below those lines: Apps Script wins on cost, speed, and ownership.
