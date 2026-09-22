# Glazing board polish: alignment, gold from the sheet, no bar

**Status: approved to build, 2026-09-22** (owner's answers in chat that day:
Q1 yes, Q2 done, Q3 keep, Q4 yes).

**Builds on** `2026-09-21-glazing-station.md` sections C and D. Office side
only. **The glazing tablet (`glazing.html`, `glazing.js`) is not touched.**

## Context

The owner opened Show ▸ Glazing station on the live dashboard (2026-09-22) and
asked for three things: the columns do not line up; a job whose whole
`Production` row is already painted gold (4998, 5091 were the examples) shows
no colour; the progress bar is noise — "I only need the number of doors and
windows and colour as ready (completed, gold) and in progress (yellow)".

## Hard rules

- Display only. **No list write, no workbook write, no seed.** The
  `Glazing station` list is not written by anything in this brief; the office
  steppers (`glzOfficeEdit`) keep writing exactly what they write today.
- The gold-row fact comes from the parser's existing `j.done` (parser.js
  `rowDone`: ≥ 10 gold cells in the spine or ≥ 40 in the row). Nothing new is
  read off the sheet.
- Section E (the floor's voice in the phase bar) is **unchanged**: a gold row
  says nothing new to `floorPhaseOf`; a job's phase already reads the sheet.
- Welding board keeps its bar and its green. Glass board untouched.
- No real names. Verification line, rule-2 grep, `test_pages.js` all clean.

## Behaviour

### 1. Alignment (`index.html`)
`glzRowHtml` (app.js ~3832) emits nine cells into welding's `.wohead` grid,
which has ten tracks and starts with an 18px toggle column glazing does not
have, so every glazing cell sits one track left. Fix: the glazing board's host
gets its own class (`stboard weldboard glzboard`, app.js ~7937) and
`.glzboard .wohead` gets its own `grid-template-columns` with one track per
cell actually emitted, in this order: job 78px · customer minmax(90px,1.4fr) ·
section minmax(80px,1fr) · count 64px · quantity words (Windows · Doors)
minmax(120px,1fr) · steppers auto · notes 34px · last touch minmax(90px,1fr).
Phone layout (`max-width:900px`) gets its own line too: job · customer · count
· steppers, rest hidden — same pattern as b5103ba.

### 2. Gold from the sheet, and from a full count (app.js + glazing-core.js)
A new pure helper in `glazing-core.js`, `glzOfficeColour(c, rowDone)`:
- `rowDone` true (the job's `Production` row is gold) → `"gold"`;
- else `c.glazed >= c.total && c.total > 0` → `"gold"` (Q1: the floor has
  finished, the office moves the row later);
- else `c.glazed > 0` → `"yellow"`; else `""`.
`glzColour` (the tablet's rule, green) is left alone — the tablet keeps its
own sentence. The office board passes `!!(byId(c.job) || {}).done`. A job not
on the sheet any more (`byId` null) uses its count alone.

Row class: `worow c-gold` gets a CSS line beside `.worow.c-green` using the
existing `--done` / `--done-bg`-style tokens of index.html (whatever the glass
board's gold row uses — reuse that token, do not invent a colour). Finished
rows still sort to the bottom (`glzCardsShown`), and a gold-row job counts as
finished for that sort.

On a gold-row job the steppers are **hidden** (Q2: "done"): the `wobtns` cell
is emitted empty and shows the word "done" in the muted ink; the count still
shows `glazed / total`. A full-count job that is not row-gold keeps its
steppers (the office may still take a unit back).

### 3. No bar (app.js)
Remove `weldBarHtml(...)` from `glzRowHtml`. Keep `glazed / total` (Q3) and
the quantity words (`GLZC.glzQtyWords`, e.g. "4 windows · 2 doors"). Nothing
else on the row changes.

### 4. Drawer line (app.js `glzDrawerLine`, Q4)
Same colour rule as the board: the `weldline` gets `c-gold` / `c-yellow`
class from `glzOfficeColour(c, !!j.done)`, and a gold-row job reads
"Glazing · done" beside the count. A job with `total` 0 stays hidden as now.

## Tests to deliver
- `test_glazing.js`: `glzOfficeColour` — rowDone with 0/N → gold; N/N no
  rowDone → gold; 1/N → yellow; 0/N → ""; total 0 → "" even with rowDone
  false; rowDone with total 0 → gold (the sheet outranks a count of nothing).
- The browser rig (`glazing_check.js`, scratchpad): a fixture job with `done`
  set shows a gold row with no steppers and the word "done"; a row has no
  `.wobar`; header cells of two rows share the same x (alignment check the way
  `welding_check.js` measured it).
- Suites unchanged elsewhere; `test_pages.js` 4/4.

## Amendments after review

- **The steppers track is a fixed 170px, not `auto`** (review, 2026-09-22).
  Section 1 said "steppers auto". Each `.wohead` is its own grid, so an `auto`
  track sizes to that row's own content: a row with the four buttons made it
  ~165px wide and a gold row showing only the word "done" made it ~32px, and
  the `fr` tracks either side swallowed the difference — the section cell was
  measured in the rig at x=430 on one row and x=474 on the next, which is the
  misalignment this brief exists to remove. 170px is the four buttons plus
  their three 5px gaps (164px measured in Edge at the board's font size). The
  phone rule uses the same fixed track for the same reason.

## What to report
Diff summary, suite counts, screenshots of the board (light and dark) with one
gold-row job, one full-count job, one in-progress job, one untouched job, and
the drawer line on the gold-row job.
