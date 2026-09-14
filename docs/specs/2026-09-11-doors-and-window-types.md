# Doors by type, windows grouped by type, status colours on the home list (brief, 2026-09-11)

**Status: approved to build, 2026-09-14** (the owner said "continue" with the
five answers; assumptions (a) and (b) below stand unless refused). The
status-list rebuild (`2026-09-11-status-list-is-truth.md`) shipped 2026-09-11
(build 20260911-1642), so door cells join it from the start, so door cells
join the "list is the truth, the Excel colour is a copy" model rather than
the old hold-and-settle one. Owner consulted 2026-09-11; answers recorded
below. Hard rules as in `CLAUDE.md`; **one new owner-sanctioned fill** is
recorded here (see §3).

## What the owner decided (2026-09-11)

1. The doors of a job are the five DOORS DONE cells (columns BT–BX): one door
   per cell, the cell text is the door's type code (`SD`, `CD`, `DD`, `SS`,
   `SFCD`, shown as they are), the fill is its status: blank not started,
   yellow in fabrication, gold done. Doors have two stages only.
2. The window types are the product groups that carry F/S/T (Casement,
   Polaris 85mm Casement, 4000 Casement, 7000 Casement, Polaris 85 Tilt &
   Turn, 4000 T&T, 7000 T&T, PVC French Wds, Arch & Angles, TH W, Alu clad
   windows, Aluclad T&T, …): exactly the checkpoints that exist today,
   regrouped under one **Windows** heading in the job card.
3. Text colours stay row flags (red urgent, green booked, blue on hold) and
   are not per-cell statuses.
4. Home list: the Windows and Doors numbers carry the status colour, with a
   hover breakdown of types and statuses.
5. Two-way with Excel for doors: set in the dashboard → the door's own cell
   is painted; painted in Excel → shown on the dashboard.

Assumptions put to the owner (correct the brief if either is refused):
(a) the Windows (M) and Doors (N) quantity cells stop being separate ticks
and are painted as an aggregate of their types (gold when all done, yellow
when any started, blank otherwise); (b) the fills written are the dashboard's
existing yellow `FFFF00` and gold `FFE699`.

## 1. Parser (`parser.js`)

- `mapSheet`: find the DOORS DONE group (header contains "doors done") and
  its five sub-columns `1`..`5`. Per job read each cell's text (the code,
  trimmed, upper-case) and fill → `j.doors = [{slot: 1..5, code, status}]`
  where status is `""` / `process` / `done` by the same fill rule used for
  glass and product cells. Empty cells produce nothing. Keep column
  positions in `PRODMAP` so the writer can address the cell (`doors[slot]`).
- Nothing else in the parser changes; product groups already carry F/S/T.

## 2. Checkpoint model (`checkpoints.js`, the status list)

- Each door is a checkpoint item `door:<slot>` with `total = 1`, label = its
  code, statuses blank / process / done. It lives in the status list like
  every other item (Title `JOB|door:<slot>`, Item, Done, Total, Status, Who,
  When, Source) with the Excel colour as the copy, exactly as the rebuild
  defines for glass and products.
- The Windows (M) and Doors (N) items become **derived**: no tick of their
  own; their status is the aggregate of the job's window types (every F/S/T
  of every window type done → done; any process/done → process) and of its
  doors respectively. The writer paints M/N from that aggregate (assumption a).

## 3. The writer (`app.js`, `graph.js`)

- Setting a door's status writes the list first (the truth), then paints
  **that one cell** (BT–BX) yellow, gold or clear. **New sanctioned fill,
  owner's decision 2026-09-11:** the five DOORS DONE cells of a job row.
  Nothing else on the row is written; the code text is never touched.
  **Owner's words (2026-09-11, in the consultation for this brief):** "when
  updated from dashboad it should update the cell with correct color as
  well. and vice versa" and "Door doesnt have 3 process so yellow fabrication
  adn doen Golden should be ok". This is the dated owner decision that
  `CLAUDE.md` rule 1 requires; rule 1 gets its own dated edit in the commit
  that ships this brief, and not before.
- M and N are painted from the aggregates, as they are painted today.
- One `Dashboard Log` line per change, as for every checkpoint.

## 4. Job card (drawer)

- **Windows** line: the count, the aggregate colour, and an arrow that opens
  the list of window types present on the job, each with its Frames / Sashes
  / Transoms rows exactly as today (steppers, All done). Closed by default;
  the open/closed state is remembered per browser.
- **Doors** line: the count, the aggregate colour, and an arrow that opens
  one row per door: code · status word · two buttons, **In fabrication**
  and **Done**, plus **Clear**; the current one is highlighted. Buttons
  ≥ 40 px. Same PENDING/hold behaviour as any checkpoint under the rebuild.

## 5. Home list

- The `WND/DRS` cell colours each number by its aggregate: gold done, yellow
  started, plain otherwise; the number itself is unchanged. The word is in
  the hover (title): "Doors: CD done · CD in fabrication · SS not started";
  "Windows: Polaris 85 frames in progress · 4000 casement not started". A
  job with no doors shows the number plain and no hover.
- No new column (owner's rule: per-job detail goes in the card).

## 6. Tests

`test_checkpoints.js`: door items from a fixture with the five cells (codes,
blank/yellow/gold, a stray non-code cell ignored), aggregates for M and N,
the writer paints only the addressed door cell and never the code text, the
list line per change. `test_export.js`: exports unchanged. A new
`test_doors.js` (offline, stubbed fetch): the drawer lists, the two-button
row, the home cell colours and hover text, Excel→dashboard on refresh.
Every suite green; rehearse the door paint on a workbook copy (scratch
`rehearse.js`) before the first live write.

## 7. Docs

`docs/REFERENCE.md` (§5 checkpoints), `docs/USER-GUIDE.md` (job card,
home list), `docs/ARCHITECTURE.md` (the new fill), `CLAUDE.md` rule 1
(third sanctioned fill), `docs/specs/README.md`.

## Report back

Files changed; full pasted suite output; one paragraph per numbered owner
decision saying where it is done and which test covers it; anything unsure.

## 8. Implementation notes against the shipped code (2026-09-14, manager)

Read these files first; the doors must join what is there, not a parallel path.

- **Item keys** (`checkpoints.js` ~line 150): `"win"`, `"drs"`, `"glass:<type>"`,
  `"prod:<name>:<f|s|t>"`, built by `cpItems(j)`; totals by `cpTotal`, labels
  by `cpLabel`, columns by `cpColumn(item, map)`. Add `"door:<slot>"` (slot
  1..5) with `total = 1`, label = the code, group `"drs"`, groupLabel "Doors";
  `cpColumn` returns the DOORS DONE column for that slot from the map.
- **The record**: every status change goes through `cpSaveRow(o)` (`app.js`
  ~748) into the `Dashboard progress` list (Title `JOB|ITEM`, Job, Item,
  Done, Total, Status, Who, When, Source); the Excel fill is painted from the
  record and registered with `cpPainted(job, item, status)`; nothing reads a
  cell colour as status. **A hand-painted cell is adopted** by the safeguard
  (`cpAdoptCandidates` / `cpAdoptConfirm` / `cpAdoptRun`, `cpAdoptDrawerJob`
  for the open drawer): extend the candidate set to the door cells so a door
  coloured in Excel is adopted into the list the same way (that is the
  owner's "vice versa"). Read the safeguard's comment block before touching it.
- **Parser** (`parser.js` ~226): `skip = ['doors done', …]` currently drops the
  DOORS DONE group; keep it out of the product groups but map its five columns
  separately (`m.doors[slot] = col`) and read code + fill per job into
  `j.doors`. `fillOf(cell)` is the existing fill reader; use its status
  mapping so yellow/gold mean the same as everywhere.
- **Derived M and N**: `"win"` and `"drs"` stay as keys so nothing else breaks,
  but they are no longer tickable; their status is computed (windows: over
  every `prod:*` item; doors: over `door:*`) and their cells are painted from
  that aggregate through the same paint path, logged as a copy. If a job has
  products but no F/S/T quantities, or no door codes, the aggregate is blank.
- **Drawer**: `cpSectionHtml(j, ed)` / `cpGroupHtml(...)` (`app.js` ~5196).
  The product groups become children of one collapsible **Windows** group;
  the Doors group lists `door:*` rows with three buttons (In fabrication /
  Done / Clear). `cpPatchSection` must keep patching in place.
- **Home list**: `rowHtml`'s WND/DRS cell: colour class from the aggregates,
  `title` with the breakdown; no new column.
- **Rule 1** in `web/CLAUDE.md`: add the third sanctioned fill with the date
  and the owner's words from §3, in this commit. `docs/HISTORY.md`: one line
  in section C for the build, one in section A for the owner's decision.
- **Rehearsal before the first live write**: `rehearse.js` against a OneDrive
  copy — paint one door cell yellow, then gold, then clear; confirm with the
  Excel API that only that cell changed and its text is intact.

---

## Amendments after review

_(append here; do not rewrite the sections above)_

### A. As built (2026-09-14)

**A1. Any non-empty text is a door.** The brief named five codes. The live
workbook, read on the day of the build, carries ten across 160 jobs — CD 149,
DD 44, SS 30, SD 14, PVC 13, DOOR 4, ACSS 2, "1 DOOR" 1, ACSD 1, SFCD 1 — and
28 yellow cells, 18 gold and 214 with no fill. So `doorCode()` accepts **any**
non-empty text (trimmed, single-spaced, upper-cased, capped at 24 characters)
and there is deliberately no list of known codes to fall behind what the
office types. Whitespace alone is not a door. A job may carry the same code in
two cells — the sheet often does — and they are told apart by their slot, not
by their code.

**A2. The doors are their own group, not part of `drs`.** Item key
`door:<slot>`, group `"door"`, groupLabel "Doors". Putting them in the `drs`
group would have handed `setGroupDone("drs", …)` a set of items to write over
a derived line. There is no "all doors done" button, by the same reasoning as
the brief's: a door is one tap already.

**A3. A door's write carries a word, not a count.** A door is one thing with
three states, so `cpStatusFor(0, 1)`/`cpStatusFor(1, 1)` cannot say "in
fabrication". `cpWriteItem` gained an optional `o.status` (and `o.fromText` /
`o.toText` for the log line); `setDoorStatus` is the only caller that uses it,
and everything counted still derives the word from the count exactly as it
did. The burst/queue machinery carries `status` and `was` through
`localStorage`, so a door tap survives the tab closing like any other.

**A4. M and N are painted by two paths, on purpose.** `cpAggregatePaint(job)`
paints them at the click, so Excel does not sit a minute behind the drawer;
`cpRepaintPlan` covers derived cells too, so a paint that was refused is put
right on the next load. Both compare the aggregate with `PAINTED` and write
nothing when they agree — and neither will write white into a cell that says
nothing and that this browser has never painted, which is what stops a tick on
the windows writing into the Doors cell of a job that has no doors.

**A5. M and N are never adopted from the sheet.** `cpAdoptCandidates` skips
derived items: they are a copy of a copy, so there is no record to adopt a
hand-paint into, and `cpRepaintRun` simply paints over it with what the items
below say. The five DOORS DONE cells are ordinary managed cells and *are*
adopted, which is the owner's "vice versa".

**A6. `cpLabel("door:3")` is "Door 3", not "Door 3 (CD)".** `diffJobs` names
the same change from the parsed file with the same function, and `dropMine`
folds the two into one line in Changes by comparing those names — so the name
has to be the same whether or not the job object is at hand. The code is in
the drawer, in the export's Checkpoints sheet, and in the home list's hover.

**A7. The home list's windows breakdown is per window TYPE, not per F/S/T.**
The brief's example ("Polaris 85 frames in progress") reads per sub-column; a
job with six product groups would then be eighteen phrases in one tooltip. It
is one phrase per type instead — "Polaris 85mm Casement in progress · 4000
Casement not started". The doors breakdown is per door, exactly as the brief
has it.

**A8. What this costs, and it should be said to the owner.** A job that has a
windows quantity in M but **no window type carrying F/S/T quantities** can no
longer have its windows ticked at all: the aggregate has nothing under it, so
it reads blank and stays blank. That follows from assumption (a) and from §8
("the aggregate is blank"), and it is the one behaviour in this change that
takes something away. Any `win` or `drs` rows the one-time import already
wrote to `Dashboard progress` are now ignored rather than deleted; pruning
them is the same later chore the status-list spec §7 already records.

**A9. Suites.** `test_doors.js` is new (21 checks, `fetch` throws so nothing
can be nearly-offline). `test_checkpoints.js` gained sections 1b and 17 and
had its machinery tests moved off `"win"` — which is no longer tickable — onto
a window type's frames, an item with the same total in a different column;
nothing about the debounce, the queue, the write order, the safeguard or the
repaint changed with it. `test_export.js` and `test_phases.js` had the
fixtures that used to colour M on its own given a window type or a door
instead, which is what colouring M now means.

**A10. Not done, and deliberately.** The PDF job card still shows one line per
checkpoint GROUP, so the individual door codes appear in the Excel export's
Checkpoints sheet and in the drawer but not on a printed card. Nobody asked
for them there, and the card is already the tallest thing in that export.

## Amendments after review (2026-09-14, manager; owner's answers 2026-09-14)

Owner's words, 2026-09-14: "already golden means it finished even if all the
cell or component has no ticked. u should not change anything in the excel
sheet. if yellow mean in fabrication and golden means done; if a window has
some yellow and some component white it means it in process." · quantity
mismatch: "that mean the quantity is wrong and should give a warning" ·
"Person name should be according to login in dashboard or login in excel" ·
"if something is ticked it means fabrication or done".

1. **Assumption (a) is withdrawn.** The Windows (`win`) and Doors (`drs`)
   lines **stay tickable** with their own record rows, exactly as before this
   brief. Their shown status = the **higher** of the line's own record and
   the aggregate of its children (window types' F/S/T; the doors): any child
   in process → at least `process`; every child done → `done`. The cell paint
   follows that shown status **upwards only**: the aggregate may paint yellow
   or gold, it never paints white. White on M/N is written only by the
   office's own Clear on that line (as today). An existing gold M/N cell is
   therefore never touched by this feature. `jobPhase` works off the shown
   status. Tests: an M gold with un-ticked types stays gold and reads done;
   ticking one door paints N yellow; ticking every door paints N gold;
   clearing the last door leaves N as it was.
2. **Doors import.** The one-time import runs once more **for door items
   only**, without clearing the 2026-09-11 marker and without reopening the
   colour fallback: a per-kind marker (`cw_cpimported_doors`), planned from
   the parsed colours, `Source = "import"`, `Who = "the sheet"`, one log line
   per job as the first import did. Theme-coloured fills (accent tint) must
   resolve exactly as `fillOf` resolves them elsewhere (R5033 slots 1–2 are
   gold via theme colour, slot 3 yellow). From then on: dashboard ticks carry
   the dashboard login; a hand-coloured door cell is adopted by the safeguard
   under the file's last saver, as for every other cell.
3. **No phantom Changes line**: `diffJobs` skips `win`/`drs` (a copy of a copy
   is never news); the aggregate paint is described inside the child tick's
   own log line ("… · Doors cell painted gold").
4. Finding 4 is closed by amendment 1 (the line is tickable again).
5. Parser fallback for the door columns stops at the next labelled group.
6. **Quantity warning**: when the Doors quantity (N) differs from the number
   of coded door cells (including quantity 0 with codes present), the drawer's
   Doors line shows "quantity says 3 · 2 doors listed" in the warning style
   and the home-list hover carries the same words. No paint decision changes.
7. `cpInProgress` counts children only, never a derived line and its children.
8. A door is its slot; overwriting a code keeps the status (owner: yes).
   Documented in the guide.
9. A cell whose whole text is a number or a date is not a door.
10. Quantity 0 with codes: the warning of amendment 6; N stays unpainted.
    Documented.
11. Gold rows: cells the parser forced to done because the row is gold are
    not safeguard candidates.

### B. The fix pass as built (2026-09-14)

All eleven points above are in the working tree. Where each one is, and what
proves it:

**1. Assumption (a) withdrawn.** `cpOwnState` is the line's own record row
(what every ordinary item answers); `cpDerivedOf` is the aggregate;
`itemState` returns the **higher** of the two on `CP_RANK`. `setItemProgress`
and the drawer's stepper are back on `win` and `drs` unchanged, so the office
ticks them exactly as it did before this brief. The paint is upwards only in
two places: `cpAggregatePlan` drops a cell whose want is `""`, and
`cpRepaintPlan` skips a derived line that has no record row and an empty want
— which is the reviewer's reproduction, and it is now impossible by
construction rather than by a guard, because `want` for a derived line can
never be lower than its own record. Tests: `test_checkpoints.js` §7 (a gold N
reads done with two of four doors untouched, and drops to its doors' word when
its own row is removed), §17 (ticking one door paints N yellow, every door
paints it gold, and the Doors line ticked on its own stays done with a door
un-ticked under it), §19 (a gold M with nothing ticked and no record is never
whitened; the office's own Clear still is; clearing the last door writes
nothing to N). `test_doors.js` §2 asserts both lines still carry their stepper
and their All done, and that the fold heads carry no control at all.

**2. Doors import.** `cw_cpimported_doors` + `cpDoorImportRun()` +
`cpImportPlan(..., "door")`, run from `load()` **before** `cpAdoptRun`.
`cpImportOwns(key)` is what keeps the safeguard off an unrecorded door cell
until that pass has drained. It runs only when the general import has already
settled, and `cpImportCheck` settles both markers together, so the two can
never both create a row for one cell. Theme fills need no new code:
`fillOf` already resolves accent-4-at-tint-.6 to `FFE699`, and
`test_doors.js` proves it on an R5033-shaped row (slots 1–2 gold via the theme,
slot 3 a literal yellow) and that `cpImportPlan(..., "door")` finds them.
`test_checkpoints.js` §18 is the pass itself: three coloured cells imported as
`Source = "import"` / `Who = "the sheet"`, the blank one skipped, nothing
painted, one summary line, **no** "adopted from Excel" line, `cw_cpimported`
untouched, the colour fallback never reopened, and nothing on the second load.

**3. No phantom Changes line.** `diffJobs` returns early on `cpDerived(x.key)`;
the child's own `cpWriteItem` call carries the words in its `toText`, built by
`cpAggregateWords(cpAggregatePlan(job, item))` — computed before the write,
because the record already moved at the click. `setGroupDone` does the same on
its `to`. Tests: `test_doors.js` (a door tick plus both cells moving in the
file is **one** line, "Door 1"); `test_checkpoints.js` §3 and §17 assert the
words in the line itself.

**4.** Closed by 1.

**5. Parser fallback.** `doorCols` is capped at five as it is collected, and
`group` changes at the next labelled column, so it cannot reach past the group.
`test_doors.js`: three columns before the next labelled group map three slots,
five where there is room, and the next group's column is never taken.

**6 and 10. Quantity warning.** `cpDoorWarning(j)` in `checkpoints.js`, drawn
by `cpLineHtml`'s `warn` argument on the Doors line — or on the doors fold head
when the quantity is 0 and there is no line — and appended to `wndDrsCell`'s
`title`. `.cpwarn` is the urgent ink. It decides nothing: quantity 0 leaves
`cpTotal(j,"drs")` at 0, so there is no Doors item and `cpAggregatePlan` plans
no paint for N. Tests: `test_doors.js` §4.

**7. `cpInProgress`** skips a derived line that has children. Test:
`test_doors.js` §5 (two real pieces of work in progress, not four).

**8. A door is its slot.** Already true — the record key is `door:<slot>` and
never the code. Asserted in `test_doors.js` (re-typing a code keeps the status
and changes only the label) and written into the user guide.

**9. Numbers and dates.** `DOOR_NUM_RE` / `DOOR_DATE_RE` in `doorCode`.
`test_doors.js` walks both lists, including `"1 DOOR"` and `"2 CD"`, which
survive.

**11. Gold rows.** `cpAdoptCandidates` returns early on `j.done`, as
`cpRepaintPlan` already did. Test: `test_checkpoints.js` §17 (not one cell of a
gold row is adopted, and not one request goes out for it).

**Docs.** `CLAUDE.md` rule 1 now says M and N keep their tick and are painted
upwards only; `docs/HISTORY.md` A18 is rewritten and A19 added for the warning;
`REFERENCE.md` §5, `ARCHITECTURE.md` (the fill rule, the module map,
`cw_cpimported_doors`), `USER-GUIDE.md` (the one-way rule, the warning, a door
is its cell, quantity 0) and `SUPPORT.md` (three new sections: the warning, a
gold cell with nothing ticked under it, and the doors import line in Changes).
