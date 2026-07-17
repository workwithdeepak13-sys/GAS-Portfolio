# SKILL: Enterprise UI Design for Google Apps Script Web Apps

> **How to use:** Paste this entire file along with any build prompt. It overrides
> every default styling instinct of the AI agent. If any instruction in the build
> prompt conflicts with this skill, THIS SKILL WINS on all UI/visual matters.

---

## 0. The Prime Directive

You are not building a "demo app UI". You are building software that a plant head,
CFO, or MD will look at for 4+ hours a day. Every screen must look like it was
designed by a senior product designer at a company like Linear, Stripe, or Vercel —
dense, calm, precise, and fast. If a screen would look at home in a hackathon
submission, it has failed.

**The AI-slop tells you must never produce:**
- Purple/violet/indigo gradients anywhere
- Giant hero sections with centered text inside an internal tool
- Emoji used as icons (📊 ❌ ✅ 🔧)
- Rounded-2xl cards floating on pastel backgrounds with huge drop shadows
- Random accent colors per card ("rainbow dashboard")
- Oversized padding that fits 4 rows of data on a full screen
- Placeholder text like "Lorem ipsum" or "Coming soon"
- Default browser fonts or Comic-Sans-adjacent playful fonts
- Buttons that say "Click Here" or "Submit"

---

## 1. Color System (exactly this, no improvisation)

Define ALL colors as CSS variables in one `:root` block. Use ONLY these tokens
in the entire app. Never hardcode a hex value inside a component style.

```css
:root {
  /* Neutrals — the app is 90% neutral */
  --bg:            #f7f7f5;   /* app background, warm off-white */
  --surface:       #ffffff;   /* cards, tables, panels */
  --surface-2:     #f1f1ef;   /* table header rows, hover states, input bg */
  --border:        #e4e4e0;   /* all borders, 1px only */
  --text:          #1a1a18;   /* primary text */
  --text-2:        #6b6b66;   /* secondary text, labels, captions */
  --text-3:        #9c9c96;   /* placeholders, disabled, timestamps */

  /* Single brand color — pick ONE per app, use sparingly */
  --brand:         #0f5132;   /* deep green (or #1e3a5f steel blue, or #7c2d12 rust) */
  --brand-soft:    #e7f0eb;   /* brand at ~8% for selected states */

  /* Status — semantic only, never decorative */
  --ok:            #15803d;
  --ok-soft:       #dcf2e3;
  --warn:          #b45309;
  --warn-soft:     #fdf0dc;
  --danger:        #b91c1c;
  --danger-soft:   #fde8e8;
  --info:          #1d4ed8;
  --info-soft:     #e3ebfd;
}
```

Rules:
- The brand color appears ONLY in: primary buttons, active nav item, focus rings,
  selected tabs, key metric emphasis. Never as a page background.
- Status colors appear ONLY on status pills, alert rows, and trend indicators.
- A soft color is ALWAYS paired with its dark text color (e.g. `--ok-soft` bg
  with `--ok` text). Never white text on soft backgrounds.
- Dark sidebar variant is allowed: sidebar `#191918`, text `#a3a39e`,
  active item white — everything else stays light.

## 2. Typography

Exactly two fonts, loaded from Google Fonts in the HTML `<head>`:

```html
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
```

- **Inter** — everything.
- **JetBrains Mono** — numbers in tables, IDs, codes, timestamps, quantities,
  currency amounts. Tabular numbers make dashboards look instantly professional:
  `font-variant-numeric: tabular-nums;`

Scale (px, use these only): 11 (micro labels, uppercase, letter-spacing 0.05em),
12 (table meta, captions), 13 (table body, form labels), 14 (default body, inputs),
16 (section titles), 20 (page title), 28 (KPI numbers only).

- Page titles: 20px / 600. NEVER larger inside an internal tool.
- KPI numbers: 28px / 600 / mono. Their labels: 11px uppercase `--text-2`.
- Line-height 1.5 for body, 1.2 for headings and KPI numbers.

## 3. Layout & Density

This is enterprise software: **density is a feature.**

- App shell: fixed left sidebar (220px, collapsible to 56px icon rail) +
  top bar (48px: page title left, global search center, user chip right) +
  content area with `max-width: 1400px` and 24px padding.
- Content spacing scale: 4 / 8 / 12 / 16 / 24 / 32. Nothing else.
- Cards: `--surface` bg, 1px `--border`, 8px radius, NO shadow (or at most
  `0 1px 2px rgb(0 0 0 / 0.04)`). Padding 16px. Never rounded-2xl.
- Tables must show 15+ rows per screen: row height 36–40px, 13px text,
  header row 11px uppercase on `--surface-2`, right-align all numeric columns.
- KPI strip: single row of 4–6 stat cards at the top of a dashboard, each with
  label / mono number / small trend arrow + delta vs last period. No icons needed.
- Forms: single column, max-width 560px, labels ABOVE inputs (13px / 500),
  inputs 36px tall, 1px border, 6px radius, focus ring = 2px `--brand` outline.
- Empty states: one 13px sentence + one primary action button. No illustrations,
  no giant icons, no "It's quiet here!" copy.

## 4. Components

**Buttons.** Primary: `--brand` bg, white text, 32px tall, 6px radius, 13px/500 —
maximum ONE per view. Secondary: white bg, 1px border. Destructive: white bg,
`--danger` text/border; solid red only inside confirm dialogs. Verb labels always:
"Create work order", "Approve NCR" — never "Submit".

**Status pills.** 20px tall, 10px radius, 11px/600 text, soft bg + dark text of the
same status family. Include a 6px dot before the text. These are the ONLY colorful
element in tables.

**Tables.** Sticky header. Hover row = `--surface-2`. First column is the entity
ID/name and is a link-styled cell. Actions in a rightmost `⋯` kebab menu, not a
row of icon buttons. Pagination bottom-right: "1–50 of 1,240".

**Navigation.** Sidebar items: 32px tall, 13px, `--text-2` default; active item
gets `--brand-soft` bg + `--brand` text + 2px left brand bar. Group headers 11px
uppercase `--text-3`. Section order mirrors the workflow, not the alphabet.

**Modals & drawers.** Edits/details open in a RIGHT-SIDE DRAWER (480px) instead
of centered modals whenever possible — it feels more like enterprise software.
Centered modal only for confirmations (400px max).

**Toasts.** Bottom-right, `--surface` with border and 4px status bar on the left
edge, auto-dismiss 4s, verb-first message: "Work order WO-0142 created."

## 5. Icons

Use inline SVG icons from Lucide (https://lucide.dev) — copy the SVG paths
directly into the HTML. 16px in tables/buttons, 18px in the sidebar.
`stroke-width: 1.75`. Color inherits from text color. Never emoji, never
Material icon font, never mixed icon sets.

## 6. Data Visualization

- Bar and line charts only, drawn with Chart.js (CDN) or inline SVG.
- One series = `--brand`. Comparisons = `--brand` + `#9c9c96`. Status splits use
  the status tokens. NEVER a 6-color rainbow palette.
- No 3D, no pie charts (use horizontal bar rankings), no gradient fills
  (flat or 8% opacity fill under a line only), gridlines `--border` dashed,
  axis labels 11px `--text-2`.
- Every chart gets a 13px title and a "vs previous period" context line.
  A number without comparison is decoration.

## 7. Motion

- Transitions: 120ms ease-out on hover/focus states only.
- Drawer slide-in: 180ms ease-out. Nothing else animates.
- No fade-in-on-scroll, no bouncing, no skeleton shimmer longer than 800ms.

## 8. Apps Script-specific execution

- Single HTML file served via `HtmlService` with `<style>` and `<script>` inline.
- All CSS handwritten with the tokens above — no Tailwind CDN, no Bootstrap.
- Show a 2px indeterminate progress bar under the top bar during every
  `google.script.run` call; disable the triggering button with a subtle spinner
  replacing its label text.
- Optimistic UI where safe; on failure, toast the error and revert.
- Render tables client-side from JSON returned by the server; never rebuild the
  whole DOM on refresh — patch rows.
- The web app must be fully usable at 1366×768 without horizontal scroll.

## 9. Self-Review Gate (run before declaring done)

Reject your own UI if ANY answer is "no":
1. Could this screenshot pass as a Linear/Stripe internal tool?
2. Is there exactly one brand color, used in fewer than 10 places per screen?
3. Are all numbers in mono font with tabular spacing, right-aligned in tables?
4. Does a 1366×768 dashboard show at least one full table AND the KPI strip?
5. Are there zero emojis, zero gradients, zero purple, zero rounded-2xl cards?
6. Does every chart answer "compared to what?"
7. Would a 50-year-old plant manager understand every label without a tooltip?
