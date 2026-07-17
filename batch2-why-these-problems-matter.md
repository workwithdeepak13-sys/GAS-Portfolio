# Batch 2 — Why These 5 Problems Matter (and Why This Portfolio Solves Them)

> Companion to `build-prompts-batch2.md`. This document explains the business reasoning
> behind each project: who has the problem, why it hurts, why existing software fails them,
> and what business-systems knowledge each project demonstrates.

---

## The Portfolio Thesis (unchanged from Batch 1, extended here)

Small and medium businesses run on paper registers, WhatsApp messages, Excel files, and the
owner's memory. They cannot afford ERP licenses (₹2–15 lakh + AMC), cannot staff an IT team,
and generic SaaS either doesn't fit their workflow or prices per-user in dollars. The result
is a predictable set of failure modes that repeat across every industry:

1. **Revenue leakage** — dues nobody totals, renewals nobody chases, extra rental days nobody bills.
2. **Compliance time bombs** — expiry dates and statutory deadlines tracked nowhere until the fine/audit.
3. **Unprovable disputes** — no photo, no signature, no timestamp; the SMB always loses the argument.
4. **Invisible unit economics** — no per-vehicle, per-contract, per-item, per-batch profitability, so bad decisions repeat.
5. **Silent internal leakage** — fuel, scrap, spare parts, cash: leakage channels that only exception reports can close.

Google Apps Script + Google Sheets is deliberately the right stack for this segment, not a
compromise: zero hosting cost, zero license cost, data lives in a spreadsheet the owner
already trusts and can open directly, Google login is free auth, MailApp is free notifications,
and the entire system is customizable per business — which is the actual demand pattern of
this market (every SMB insists "our process is different", and here it genuinely can be).

Batch 1 covered: CRM, Expenses, HRMS, Inventory ERP, ISP Billing, Project Management,
B2B Supply/Credit/GST (SupplyDesk), Visitor Management (GateDesk).

Batch 2 deliberately moves into **operational verticals** where the domain logic is deeper:
recurring-contract operations, asset compliance, installment receivables, cyclical asset
custody, and GST statutory material flows. Together the two batches demonstrate coverage of
every core business system class: sales, people, money-in, money-out, stock, projects,
service delivery, assets, compliance, and receivables.

---

## Project 1 — ServiceSarthi (AMC & Field Service Management)

**Who:** AC servicing companies, RO/water-purifier dealers, lift AMC providers, CCTV installers,
DG-set maintainers, medical equipment servicers. Typically 2–20 technicians, hundreds of
contracts. One of the most common SMB types in every Indian city.

**The problem:**
- AMC is a *promise of future visits*. Paper/Excel tracking means promised visits get missed,
  and a missed visit is a broken promise the customer remembers at renewal.
- Renewal dates live in the owner's head. Every forgotten renewal is 100% recurring revenue
  lost to a competitor who happened to call first. Renewal rate is the single number that
  decides whether these businesses grow or stagnate.
- "Your technician never came" disputes are unwinnable without proof of visit.
- Comprehensive contracts (parts included) leak margin invisibly: nobody sums the parts
  consumed against a contract, so unprofitable customers get renewed at the same price.

**Why existing software fails them:** Field-service SaaS (ServiceTitan-class, or Indian
equivalents) is priced per technician per month and built around dispatch-heavy US trades.
The Indian AMC operator needs contract-visit-renewal logic first, dispatch second.

**Why solve it:** Every feature maps to money: the visit auto-scheduler protects renewals,
the renewal pipeline recovers forgotten revenue, OTP + photo closure kills disputes, and the
parts-vs-contract cost report tells the owner which contracts to reprice. This is a system
where the ROI story writes itself.

**Business-systems knowledge demonstrated:** recurring-revenue contract lifecycle management,
SLA engineering (server-owned deadlines, breach states), field-workforce state machines,
proof-of-service design (photo + OTP), contract-level cost accounting, and pipeline-weighted
revenue forecasting.

---

## Project 2 — FleetSarthi (Fleet Trip, Fuel & Compliance)

**Who:** Small transporters, school-bus operators, construction-material suppliers,
distribution fleets — 5 to 50 vehicles. Owner-operated, supervisor-managed.

**The problem:**
- **Compliance is existential, not administrative.** An expired insurance policy on the day of
  an accident means the claim is rejected and the owner absorbs the full loss. Expired permits/
  fitness/PUC mean checkpoint fines and vehicle detention. These dates are tracked nowhere.
- **Fuel is 40–55% of operating cost** and the biggest leakage channel (inflated bills,
  siphoning). The only honest detector is consistent tank-to-tank km/l computation — which
  no one does on paper.
- Driver cash advances are remembered, not ledgered → settlement disputes and quiet losses.
- No per-vehicle P&L means the owner cannot identify the loss-making vehicle in the fleet —
  and there is almost always one.

**Why existing software fails them:** Fleet telematics platforms sell GPS-hardware
subscriptions per vehicle; transport ERPs target 200+ vehicle fleets. The 5–50 vehicle
operator needs trip sheets, a fuel log, a document alarm clock, and a driver ledger — not
hardware.

**Why solve it:** The compliance engine alone (red/amber/green document board + trip-assignment
block on expired docs) prevents losses that dwarf the cost of any software. The mileage
anomaly detector addresses leakage without confrontation — the data raises the flag, not the owner.

**Business-systems knowledge demonstrated:** asset compliance management with statutory
document lifecycles, exception-based fraud detection (statistical deviation from trailing
baselines), append-only ledger design (driver accounts), derived-vs-materialized counters
(odometer advancement under locks), and per-asset P&L construction with allocated fixed costs.

---

## Project 3 — VidyaSarthi (Coaching Institute Fee, Batch & Attendance)

**Who:** Coaching institutes, tuition centers, training academies, music/dance/skill schools —
50 to 1000 students. One of the largest and most underserved SMB categories in India.

**The problem:**
- Fees are **installment receivables**, not one-time sales. Paper registers cannot answer
  "what is my total pending dues right now?" — the most important number in the business.
  5–15% of billed fees typically slip through as silently-unpaid installments because chasing
  is manual, inconsistent, and socially awkward for staff.
- A dropped student is ~12 months of lost fees, and disengagement shows up first as
  consecutive absences — which paper registers record but never alert on.
- Marketing spend (pamphlets, Instagram, referrals) is blind because enquiry→admission
  conversion is untracked by source.
- Cash-mode fee collection with hand-written receipts is a reconciliation and pilferage risk.

**Why existing software fails them:** School-ERP products are priced and structured for formal
schools (academic years, government reporting). Coaching-specific SaaS exists but charges
per student per month — brutal at 500 students — and resists the per-institute customization
(fee structures, batch patterns) this segment demands.

**Why solve it:** The installment engine + dues aging + automated reminders directly recover
the leaked 5–15%. The absence-alert is a retention system disguised as attendance. Gap-free
receipt numbering + daily collection register by collector closes the cash hole.

**Business-systems knowledge demonstrated:** receivables management (installment plans,
aging buckets, partial payments, dunning/escalation design), gap-free document sequencing,
funnel analytics (source-attributed conversion), early-warning churn detection, multi-role
data scoping (parent sees only own child — enforced server-side), and cash-reconciliation
controls.

---

## Project 4 — RentSarthi (Equipment Rental & Hire)

**Who:** Construction equipment hire (scaffolding, shuttering, mixers), event rentals
(tents, chairs, sound), tool/machinery hire, furniture and medical equipment rental.
Massive, fragmented, almost entirely paper-run.

**The problem:**
- Rental is **cyclical custody**, which generic inventory software cannot model: the same
  item goes out and comes back forever, and every joint leaks — unbilled extra days,
  unassessed damage, unaccounted shortages, disputed deposits.
- Partial returns are the norm (300 of 500 plates come back today, the rest next month),
  which splits the billing timeline — the hardest math in the domain and the place where
  paper systems simply give up and under-bill.
- Deposit disputes ("it was already broken", "you never returned 40 plates") are unwinnable
  without dispatch-condition photo proof.
- The owner cannot answer "where is my stock?" across a dozen concurrent customer sites —
  which means over-buying inventory that is actually just lost.

**Why existing software fails them:** Rental-management SaaS exists but targets Western
equipment-rental chains with barcoded serialized fleets. The Indian hire business runs on
fungible quantity pools (500 identical plates) and handwritten challans.

**Why solve it:** The running-charge meter and partial-return billing engine directly convert
currently-unbilled days into revenue. The dispatch-vs-return photo comparison and the
deposit settlement ledger turn every dispute into a lookup. The stock deployment board is
the "wow" screen that sells the system in one demo.

**Business-systems knowledge demonstrated:** custody-cycle asset modeling, quantity-pool
availability with lock-protected counters + rebuild-from-history repair functions,
time-based revenue recognition with partial-fulfillment splits, deposit/settlement ledger
design, evidence-chain workflows (photo pairing), and asset utilization economics
(payback tracking, dead-capital identification).

---

## Project 5 — JobSarthi (Job-Work Challan & Material Reconciliation)

**Who:** Small manufacturers and fabricators in every industrial cluster — anyone who sends
material out for powder coating, plating, machining, heat treatment, stitching, casting — and
the job workers on the receiving side. Nearly every small factory is on one or both sides of
this flow daily.

**The problem:**
- This is the most **compliance-critical** flow in the batch. Under GST (Section 143, Rule 45),
  material moves to a job worker on a tax-free delivery challan — but must return within
  1 year (inputs) / 3 years (capital goods), or the principal owes GST on it as a deemed
  supply, with interest. Quarterly ITC-04 filing demands challan-wise reconciliation.
  Paper challan books make both requirements a year-end archaeology project.
- **Process scrap is a silent theft channel.** Every process has an agreed wastage norm
  (e.g. 3% in machining); nobody reconciles actual scrap against it, so the delta walks out
  of the job worker's gate as sellable metal — invisibly, forever.
- Nobody can answer "what of mine is lying in whose factory right now, and for how long?" —
  working capital sitting untracked in other people's premises.

**Why existing software fails them:** Job-work modules exist only inside full ERPs
(SAP B1, larger Tally-ecosystem add-ons) priced far beyond this segment. There is essentially
no affordable standalone tool — which is exactly why the paper challan book survives.

**Why solve it:** The deadline-ageing engine with a rupee-denominated deemed-supply exposure
figure converts an abstract legal risk into a number the owner acts on. The wastage-norm
exception flag closes the scrap channel. The ITC-04 CSV generator turns a quarterly
accounting nightmare into a button. This is the project that most strongly signals deep,
India-specific statutory-process knowledge.

**Business-systems knowledge demonstrated:** GST statutory material flows (Sec 143 / Rule 45 /
Rule 55 challan fields / ITC-04 Tables 4 & 5A), dual-perspective system design (the same
system serving principal and job worker modes), norm-vs-actual exception reporting,
multi-receipt partial reconciliation under locks, statutory gap-free document numbering,
and deemed-supply exposure quantification.

---

## Coverage Map After Batch 2

| Business system class | Batch 1 | Batch 2 |
|---|---|---|
| Sales pipeline / CRM | ORBIT CRM | Renewal pipeline (ServiceSarthi), Enquiry funnel (VidyaSarthi) |
| Money out (payables/claims) | Expense Tracker | Driver ledger (FleetSarthi) |
| People / HR | HRMS | Technician workforce mgmt (ServiceSarthi) |
| Stock (linear flow) | Inventory ERP | — |
| Stock (custody cycle) | — | RentSarthi |
| Stock (statutory movement) | — | JobSarthi |
| Recurring billing | ISP Billing | AMC contracts (ServiceSarthi) |
| Installment receivables | — | VidyaSarthi |
| Projects / delivery | FlowPulse PMS | — |
| B2B credit + GST invoicing | SupplyDesk | Job-work invoicing (JobSarthi) |
| Front-desk / access | GateDesk | — |
| Asset compliance | — | FleetSarthi |
| Field service / SLA | — | ServiceSarthi |
| Fraud/leakage exception reporting | — | Fuel anomalies, wastage norms, parts-vs-contract cost |

**Common architecture across all 13 systems** (the portfolio's signature): idempotent
`initializeSheets()` with seeded demo states, `success/fail` response contracts, LockService
on writes + CacheService on reads, Config-driven business rules (never hardcoded), role-gated
single-SPA frontends, audit + error logs on everything, soft deletes with undo, server-owned
state machines and deadlines, and honest "reality-check" handling of platform limits
(no fake WhatsApp/SMS — pluggable notify modules with free email defaults).

---

## Suggested Build Order

1. **ServiceSarthi** — broadest applicability, easiest to demo, strongest ROI story.
2. **VidyaSarthi** — huge market, and the installment/receivables engine is highly reusable.
3. **RentSarthi** — the hardest billing math in the batch; a differentiator project.
4. **FleetSarthi** — the compliance engine pattern generalizes to many asset businesses.
5. **JobSarthi** — the deepest statutory-knowledge signal; best built last with the most care.
