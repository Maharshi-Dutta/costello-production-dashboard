# Fabrication: glass status and door glazing — brief (2026-10-01)

Status: shipped 2026-10-02, build 20261002-0817.

Read first: `docs/specs/2026-09-25-fabrication-station.md` (and Amendments),
`docs/fabrication-station.md`, `docs/STATIONS.md` (Fabrication, Glass and
Glazing sections), `docs/specs/2026-09-11-doors-and-window-types.md` (the
DOORS DONE cells as checkpoint items `door:<slot>`),
`docs/specs/2026-09-11-status-list-is-truth.md`, `docs/HISTORY.md` B31, B32.

## Why

The glazing tablet counts windows only (2026-09-23), so nobody records door
glazing. And the fabricators cannot see whether a job's glass is ready: "if a
job has windows and doors not fabricated but glass is done, they know they
should start that job".

## A. Glass status on every fabrication job (read only)

- New feeder column on `Fabrication station`: **`Glass`** (single line text),
  written by the office feeder on every row of the job, one of:
  - `done` — cutting and hotmelting complete for all the job's glass, **and
    Tuff complete when the job has Tuff** (owner), or the office's record says
    the job's glass is done (`OfficeDone` meaning);
  - `part:<cut>/<total>:<hotmelt>/<total>` (plus `:tuff <n>/<m>` when the job
    has Tuff) — started, not complete;
  - `none` — the job has glass and nothing is recorded;
  - blank — the job has no glass.
  Derive it in a pure core function from what `app.js` already holds (the
  glass station rows / `ST` glass board record and the office checkpoint
  record). No new workbook read. `Glass` joins `FB_FEEDER_FIELDS` and the
  feeder hash, so a glass change re-feeds.
- Tablet and office board: a chip on the card head — green "Glass ✓ done",
  amber "Glass: cut 12/20 · hotmelt 5/20", grey "Glass: not started"; nothing
  when blank. When Glass is `done` and the job's fabrication done count is 0
  (In production, not finished on sheet): a line "Glass is ready and nothing
  is fabricated yet — start this job".
- Tablet: a toggle **Glass ready first** (remembered per tablet,
  `cw_fabglassfirst`): sorts glass-done-and-unfinished jobs to the top (after
  urgent, which stays first). Works with every view of the My work filter.
- The fabrication tablet never reads the glass list itself.

## B. Door glazing — a fourth part on door groups

- New part key **`glazing`**, label "Door glazing", drawn as a fourth line
  under **PVC DOOR** and **PVC SMART** only, when its total > 0.
- Total = number of the job's DOORS DONE cells (`j.doors`) by code:
  `SS` → PVC SMART; `CD`, `SFCD`, `BF`, `ACSD`, `ACSS` → not counted;
  **every other non-empty code** (PVC, DD, SD, …) → PVC DOOR (owner:
  "anything under PVC Doors and PVC Smart … SD, DD, PVC and more but not
  CD"). Only fed on a row that exists (the job has that group); if the job has
  glazeable doors but no such group row, feed nothing and say so in the
  report-back with a count from the local workbook copy.
- New columns: `GlazeTotal` (Number, feeder), `GlazeDone` (Number, floor),
  `GlazeBy`, `GlazeAt` (text, floor). `DoneBy/DoneAt` updated as for any tap.
- **Role only, no assignment** (owner): a person may tap the line when
  `Stages` grants them the `glazing` part of that group (`PVC DOOR:glazing`,
  `ALL:glazing`). **A bare group name still means frames+sashes+transoms
  only** — glazing must be named. The assignment gate (`fbCanTap`'s Assigned
  row requirement), Take, the office Assign picker, splits and the
  All/None-only-when-whole-line-is-yours rule do **not** apply to the glazing
  part: eligible = may tap, All/None allowed. Fails closed like the rest when
  the person is not eligible.
- "My work" shows the glazing line to a glazing-role person; "Assigned to me"
  never shows it (nothing is assigned).
- The glazing line does **not** count toward the group's/card's fabrication
  done/total, colour, or the lavender/purple sheet paint, and a job is
  "finished" for the tabs exactly as before (frames/sashes/transoms). It is
  its own count.
- Office board: the same line with − + All None (office edit: counter,
  GlazeBy/At, DoneBy/At, one `Dashboard Log` line; never `Station log`).
  `Station log`: tablet taps log `Stage = glazing`, `GlassType` = group.
- Finished-on-sheet green jobs: the line is drawn full and locked like the
  others.

## C. Gold on the sheet when door glazing is complete (rule 1, owner 2026-10-01)

Owner: "when door glazing is done make door done/golden in excel". Yes to:

- When a job's door glazing is **full** on a group (GlazeDone >= GlazeTotal >
  0), the **office dashboard** sets the checkpoint record of each of that
  job's counted door cells for that group (`door:<slot>` items whose code maps
  to that group by the table in B) to **done**, through the existing
  checkpoint record path (the `Dashboard progress` list is the truth; the
  existing writer then paints the cell gold and logs it as it does for an
  office tick). Source `fabrication`, Who = the glazer's name from `GlazeBy`
  (list data, not repo).
- **Partial counts paint nothing.** **Raise only**: a record already `done` is
  left; if the count later drops, nothing is cleared — only the office clears.
  A record the office has set to anything is never lowered.
- CD / SFCD / BF / AC* cells are never touched.
- Same safety as HISTORY B32: decide from the record read **now**, inside the
  serialised path the checkpoint writes already use; skip while `cpPending`
  for that item; skip jobs that are gold rows, off-sheet, or finished on
  sheet; skip when the record list is unreadable or the import is pending.
- The tablet never writes the workbook or `Dashboard progress`.
- Update rule 1 in `web/CLAUDE.md`, dated: the door cells (third sanctioned
  fill) gain a second writer, fabrication's door glazing, gold only, raise
  only.

## D. Welding and glazing office boards: redraw after a capped feed

Review of e716f16 confirmed `feedWelding` (~app.js:2918) and `feedGlazing`
(~3626) have the stale-board bug `feedFabrication` had: a continuation feed
replaces the cache without a redraw and the next identical read is skipped by
`sameRows`. Add the same one-line `redrawWelding()` / `redrawGlazing()` after
a feed that wrote rows. Test each.

## Hard rules

- Tablet files: rule-2 grep stays 0. Feeder never writes floor columns
  (`GlazeDone/By/At`), never deletes. Office never writes `Station log` /
  `Station people`.
- Quantities from `Production` alone (`j.doors`, `j.prodsMain`).
- If `GlazeTotal`/`Glass` columns are missing on the list, the feed must not
  be refused wholesale: detect and say which column is missing (quiet
  explained state), keep everything else working. (The manager adds the
  columns by script before the push.)
- No real names in the repo. Edit/Write only. No commit, push, live services.
- 10-inch portrait, no sideways scroll at 700–1100 px, both themes, header
  two rows at 700 px with the new toggle.

## Tests

`test_fabrication.js`: glass status derivation (done incl. Tuff rule, part,
none, blank); door code → group table incl. unknown codes and CD excluded;
`Stages` parsing with `glazing` (bare group excludes it; `ALL:glazing`);
gate: role only, no assignment needed, All/None allowed; glazing not in
fabrication totals/colour/tabs/painter; gold plan: full → the mapped door
items done, partial → nothing, already-done → nothing, count dropping →
nothing cleared, CD never, pending/finished-on-sheet skipped; Glass ready
first sort; welding/glazing redraw (D). Browser rig: chip states, the start
line, the toggle, the glazing line for a glazing-role person and a
non-glazing person, 700 px both themes.

## Report back

Changed functions one line each, suite tails, rig results, screenshot paths,
the count of jobs with glazeable doors and no group row, anything
interpreted.
