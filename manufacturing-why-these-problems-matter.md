# Manufacturing Systems (Batch 3) — Why These Problems Matter

> Companion to `manufacturing-build-prompts.md`. Explains who has each problem, why it hurts,
> why existing software fails this segment, and what business-systems depth each project demonstrates.

---

## Why manufacturing, and why these five

Small and mid-size manufacturers (10–500 employees) are the single largest underserved software segment:

- **They cannot afford real MES/QMS/CMMS/ERP modules.** SAP/Oracle module implementations start in the tens of lakhs; even mid-market tools (Tulip, MaintainX Enterprise, ETQ, Plex) are priced per-seat per-month in dollars, for plants where an operator's monthly salary is less than one seat license.
- **Even when they buy ERP, the operational modules stay empty.** ERP assumes data discipline (routings, BOMs, standard costs, real-time confirmations) that a mid-size plant never achieves. The finance module runs; production, quality, maintenance, and costing modules rot. The plant runs on Excel + WhatsApp + shouting.
- **Google Workspace is already there.** Sheets + Apps Script means zero licensing, zero hosting, full customizability, and data the owner physically controls — exactly matching this repo's philosophy.

The five problems were chosen because each one (a) burns real money monthly, (b) has no affordable packaged answer, and (c) exercises a DIFFERENT class of manufacturing system, so the batch together demonstrates end-to-end plant-management knowledge.

---

## 1. Production Planning & Shop-Floor Execution System

**Who has this problem:** Every discrete/batch manufacturer with more than ~5 machines — auto components, plastics, packaging, textiles, food processing, fabrication.

**Why it hurts:**
- Without plan-vs-actual visibility, capacity is silently wasted. Industry studies consistently put unmeasured OEE at 40–60% while managers believe they run at 85%. On a plant with ₹50L/month machine capacity value, every 10 OEE points recovered is ₹5L/month.
- Downtime that is not categorized is downtime that never gets fixed — "the machine is always down" is not actionable; "changeover consumed 22% of machine 7's month" is.
- The daily production meeting consumes 45 minutes × 8 managers × 26 days — roughly 150 manager-hours/month arguing over whose Excel is right.

**Why existing software fails them:** True MES needs machine connectivity (sensors, PLCs, integrators) — a ₹20L+ project. This system deliberately uses hourly human logging, which is 90% of the analytical value at 2% of the cost, and is honest about that trade-off.

**Systems knowledge demonstrated:** OEE mathematics (Availability × Performance × Quality), finite-capacity plan validation, downtime reason-tree taxonomy, shift-based data modeling, wall-board UX, exception-driven management (dispatch-risk panel).

---

## 2. Quality Inspection, NCR & CAPA Management System

**Who has this problem:** Any manufacturer supplying to OEMs, exporters, or anyone with ISO 9001 / IATF 16949 / customer-audit obligations — plus every plant that keeps rediscovering the same defect.

**Why it hurts:**
- Cost of poor quality in unmanaged plants typically runs 10–20% of revenue (rejection + rework + returns + expediting). Almost none of it is visible because nobody aggregates it by defect, machine, or supplier.
- A repeat defect is pure system failure: the plant already paid to discover the problem once and then paid again because the corrective action was never verified. The enforced 5-Why + mandatory effectiveness check is the mechanism that breaks this loop — this is textbook CAPA discipline that even expensive QMS software fails to enforce culturally.
- Failed customer audits lose contracts. Three panic-days of record reconstruction before every audit is the norm; the audit-binder generator makes it a 3-minute export.

**Why existing software fails them:** Enterprise QMS (ETQ, MasterControl, Intelex) is priced for regulated industries. Mid-market plants get nothing between "paper register" and "₹15L/year."

**Systems knowledge demonstrated:** Inspection planning with characteristic-level specs, sampling logic, disposition authority matrices, CAPA lifecycle with effectiveness verification, supplier quality metrics (PPM, lot acceptance), cost-of-poor-quality accounting, audit-trail design for compliance contexts.

**Deliberate scope boundary:** This is internal quality only (incoming/in-process/final). Customer complaints are explicitly out of scope because that system already exists in the portfolio — the two systems would complement, not overlap.

---

## 3. Machine Maintenance & Spare Parts Management System

**Who has this problem:** Every plant. Reactive-only maintenance is the default state of small manufacturing.

**Why it hurts:**
- Reactive maintenance costs 3–5× planned maintenance for the same failure (emergency labor, expedited parts, collateral damage, and the downtime itself).
- A single unplanned breakdown on a bottleneck machine can kill a dispatch commitment — connecting this system to Batch-3 Prompt 5's commitment engine is the enterprise story.
- The "critical spare discovered missing at 2 AM" event turns a 2-hour repair into a 3-day stoppage. The criticality cross-check (A-machine × zero-stock critical spare) prevents exactly this, proactively.
- Without MTBF/MTTR, repair-vs-replace decisions are gut feel on the plant's most expensive assets.

**Why existing software fails them:** CMMS tools (MaintainX, Fiix, UpKeep) charge per-seat monthly; a 10-technician team costs more per year than this entire system costs to build. Most small plants therefore track nothing.

**Systems knowledge demonstrated:** PM scheduling engines, reliability mathematics (MTBF/MTTR), severity/criticality matrices, escalation timers, spares min-max inventory with lead-time buffers, downtime costing, chronic-failure detection — the full CMMS domain.

---

## 4. Tooling, Die & Fixture Lifecycle Management System

**Who has this problem:** Press shops, injection/blow molding, die casting, forging, machining — any plant whose production physically depends on dies, molds, jigs, fixtures, and gauges.

**Why it hurts:**
- A die can cost ₹2–50L — often the second-most-expensive asset class after machines — yet tool life tracking is almost universally absent. Running a die past rated life risks catastrophic failure mid-run: scrapped material, a dead delivery commitment, and a multi-week die repair.
- Customer-owned tooling is a contractual liability: customers audit their dies, and "we're not sure where it is or what condition it's in" damages the relationship and can trigger penalty clauses.
- Tools at outside refurbishment vendors are money in limbo — no follow-up, no cost accumulation, no repair-vs-retire logic.
- Expired gauges silently invalidate every inspection done with them — a compliance time bomb that connects directly to System 2.

**Why existing software fails them:** This is a genuine software desert. Tool management modules exist only inside high-end MES/PLM suites. There is effectively NO affordable standalone product — which makes this the strongest portfolio differentiator of the batch: it signals knowledge that only comes from actually understanding plant operations.

**Systems knowledge demonstrated:** Asset lifecycle modeling with consumable life (shot counters), location/custody tracking, vendor job tracking, calibration compliance, customer-asset fiduciary reporting, replace-vs-refurbish economics.

---

## 5. Order Profitability & Delivery Commitment Truth System — the MD's unsolved pain

**Who has this problem:** Practically every mid-size manufacturing MD/owner. This is the problem they have personally tried to solve for years — through meetings, Excel formats that die in a month, and ERP modules that stay empty — and never solved.

**Why it is THE pain:**
- **"Busy but not profitable" is the most common and most dangerous state of a mid-size plant.** Sales grow, machines run full, and cash doesn't grow — because 20–30% of orders quietly lose money (stale cost sheets, untraced rejection losses, overtime, rush freight) and nobody can identify which ones. The company-level P&L arrives 45 days late and averages the losers into the winners, hiding them completely.
- **Delivery slips with zero accountability are an organizational disease.** When commitments are verbal, every review meeting is a blame circle, the same failure mode repeats monthly, and the MD becomes the plant's only integration layer — personally chasing every order. That is exactly the treadmill he is tired of.
- The two halves are the same problem: **no order-level truth.** Margin truth and commitment truth both die at departmental boundaries. This system creates a single, timestamped, signed record that crosses those boundaries.

**Why it has stayed unsolved:**
- ERP costing modules demand standard-cost discipline, routings, and real-time confirmations that mid-size plants cannot sustain — so the MD was sold the answer and never received it. This system is designed around what a real plant CAN capture: one-minute cost postings by the department that already knows the number, and five milestone gates instead of fifty routing steps. **Its realism is the innovation.**
- No packaged product does cross-departmental commitment accountability, because it's an organizational-design problem wrapped in software. Building it demonstrates understanding of management systems, not just code.

**Systems knowledge demonstrated:** Order-level contribution costing, quoted-vs-actual variance analysis by cost head, customer/SKU profitability ranking, event-sourced commitment ledgers, back-scheduled milestone gates with early-warning logic, accountability system design, and executive information design (the 7 AM digest — five sections, thirty seconds, complete truth).

---

## How the batch fits together (the enterprise story)

These five are not isolated apps — they form a coherent plant operating system, and the portfolio should say so:

| Layer | System | Feeds |
|---|---|---|
| Execution | 1. Production Planning & Shop-Floor | actual output + downtime → OEE, dispatch risk |
| Assurance | 2. Quality/NCR/CAPA | rejection losses → cost heads in System 5; gauge status ← System 4 |
| Asset reliability | 3. Maintenance & Spares | breakdown downtime → System 1's downtime categories |
| Asset lifecycle | 4. Tooling & Die | die-failure risk → System 1's plan; gauges → System 2 |
| Management truth | 5. Order Profitability & Commitment | consumes signals from all four; answers the MD |

Batch 1 (repo) covered horizontal SMB systems (CRM, HRMS, IMS, expenses, projects). Batch 2 covered service-economy verticals. **Batch 3 goes vertical into the factory itself** — planning, quality, maintenance, tooling, and order economics — the five pillars of manufacturing operations management. Together they demonstrate that the portfolio's author understands not just how to code business apps, but how a manufacturing business actually works and fails.

---

## Positioning line for the repo README

> "Enterprise-grade manufacturing operations systems — production execution, quality (NCR/CAPA), maintenance reliability, tooling lifecycle, and order-level profitability truth — built entirely on Google Apps Script, priced at zero licensing, designed around the data a real mid-size plant can actually capture."
