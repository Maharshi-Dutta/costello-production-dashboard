# Phase pipeline + floating selection wheel (spec, 2026-09-07)

Dashboard-only. NOTHING is written to the workbook by either feature; no Graph calls,
no new sheets, no localStorage beyond one UI preference. Repo: `d:\Costello Windows\Excel
Dashboard\web` (static site, plain ES2017, no modules; read `app.js`, `parser.js`,
`checkpoints.js`, `index.html`; offline test pattern in `test_checkpoints.js`).

## Feature 1 - the phase pipeline is the primary progress
In the job drawer, the current "Progress" section (five date steps: Sold, Stamp, Ivana,
Ready to print, Sent to floor) is replaced as the primary view by a **phase pipeline**:

`In office → Sent to floor → Cutting → In fabrication → In glazing → Quality check → Fitted / delivered`

Derived purely from data already parsed (a pure function `jobPhase(j)` → index 0..6,
in `checkpoints.js` or a new `phases.js` loaded before app.js; export on `window`):
- 0 In office: no `dates.floor` and no checkpoint activity.
- 1 Sent to floor: `dates.floor` set (or any checkpoint in process/done) and nothing below applies.
- 2 Cutting: any product F/S/T cell carries the sheet's "Cut" colour (see parser note) and no fabrication yet.
- 3 In fabrication: any product F/S/T item status "process" (yellow) or windows/doors in process.
- 4 In glazing: every product F/S/T item that exists is "done" (gold), job not ready to deliver, and not all glass/windows/doors done.
- 5 Quality check: every checkpoint item that exists (products, glass, windows, doors) is "done" and the job is not marked ready to deliver.
- 6 Fitted / delivered: job marked ready to deliver (`j.done`) or no longer on the Production sheet (`cat === "past"`).
The highest applicable phase wins. Phases 4-6 are shown as steps regardless; nothing sets
them by hand and nothing writes anywhere.

Parser note: `parser.js` today reads yellow (`FFFF00`) and gold (`FFE699`) fills. Add
"cut" = the sheet's green Cut colour: detect it from the legend cell whose text is "Cut="
on the header rows (its fill), falling back to `92D050`/`00B050`/`C6EFCE`. Record it in
`j.cp.prod[name][sub] = "cut"` (rank below process) - purely additional, `verify.js` must
still agree exactly (it compares fields the Python reader produces; `cp` is not one).

UI: a horizontal stepper with the seven steps, filled up to the current phase, the current
step highlighted with a word (never colour alone), small labels; on phone it wraps to two
rows. The date steps move into a **collapsed "Dates" section** directly under it
(chevron, remembered in `state.collapsed` like group collapse, key `"dates"`), default
collapsed. Checkpoints section unchanged. The job list row gets the phase word in the
existing status badge spot (replace the stage badge text with the phase name for jobs in
production; keep "Ready to deliver"/section badges as they are).

## Feature 2 - floating selection wheel
When one or more jobs are ticked (`state.picked`), a **floating round button** appears at
the bottom-right (fixed, 56 px, shows the count). Tapping it fans out options in a
quarter-circle "wheel" animation (CSS transitions, 200-300 ms, staggered): **Alert to…**
(admin only, opens `renderAlertMenu` for the ticked jobs), **Export** (opens the Export
window with Scope preset to "Ticked jobs"), **Move to…** (the existing move menu), and
**Clear** (clears the selection). Tapping the button again or outside closes it. Escape
closes it. The wheel never covers the Download/Save buttons of an open window (hide it
while any window/drawer is open at phone width). Theme-aware, 40 px+ targets, no library.

## Tests (`test_phases.js`, vm-loaded real code)
`jobPhase` for each phase with fixtures (including cut colour, mixed states, past jobs);
parser picks up the Cut colour from a legend cell; the wheel renders only when something
is ticked and its options wire to the existing functions (assert on calls with the ticked
ids); Export opens with scope "ticked". Keep every existing suite green.

## Deliverables
Summary per file, pasted output of node --check (all JS) and every suite. No commits.
