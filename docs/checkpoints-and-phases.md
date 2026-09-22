# Checkpoints and the phase pipeline

`checkpoints.js`, drawer in `app.js`.

## 5. Checkpoints (`checkpoints.js`, drawer in `app.js`)

**What.** Per-job counts for each glass type and each product's F/S/T, plus
one checkpoint per **door**, ticked from the drawer and shown as bars.

**Doors and the two derived lines — since 2026-09-14.** Spec:
[`docs/specs/2026-09-11-doors-and-window-types.md`](specs/2026-09-11-doors-and-window-types.md).

- A job's **doors** are the five DOORS DONE cells of its row (BT–BX as the
  sheet stands). Each cell that has text in it is one door: the text is its
  type code (`CD`, `DD`, `SS`, `SD`, `PVC`, `DOOR`, `ACSS`, `SFCD`, `1 DOOR` —
  **any** non-empty text, there is no list of known codes), the fill is its
  status, and a door has **two stages and a blank**: white not started, yellow
  in fabrication, gold done. Item key `door:<slot>`, total 1, label the code;
  `j.doors` on the job model, `PRODMAP.doors` for the columns. The code text
  is never written by anything. Ticking one writes the `Dashboard progress`
  row first and paints that one cell from it, like every other checkpoint; a
  door painted in Excel by hand is adopted by the safeguard, which is the
  owner's "vice versa".
- **Windows (M) and Doors (N) keep their own tick and also answer for what is
  under them.** `cpOwnState` is the line's own record row, exactly as before;
  `cpDerivedOf` is the aggregate of the job's `prod:*` items and of its
  `door:*` items; `itemState` returns the **higher** of the two. A gold M
  therefore reads done however little is ticked underneath, and a line nobody
  has touched still goes yellow the moment one window type or one door moves.
  **The paint follows that upwards only**: `cpAggregatePaint` (at the click)
  and `cpRepaintPlan` (once a load) may take M or N to yellow or gold and
  never back to white — only the office's own Clear on the line, which leaves
  a record row saying so, whitens it. A doors quantity that disagrees with the
  number of coded cells is reported by `cpDoorWarning` in the drawer's Doors
  line and in the home list's hover; it decides no paint and blocks nothing.
- **The doors' own one-time import.** The general import (`cw_cpimported`) had
  drained in every browser three days before the doors existed, so the door
  cells get one pass of their own behind `cw_cpimported_doors`
  (`cpDoorImportRun`, `cpImportPlan(..., "door")`): `Source = "import"`,
  `Who = "the sheet"`. It does not clear the general marker and does not
  reopen the colour fallback. Until it drains, `cpImportOwns` keeps the
  safeguard off the door cells, so a door that has carried a colour since
  before this shipped is not recorded as somebody's hand-paint.
- **The job card** puts every window type inside one collapsible **Windows**
  fold and the doors inside a **Doors** fold, each headed by its derived line;
  open/closed is remembered per browser in `cw_cpopen`. **The home list's
  WND / DRS numbers** carry the aggregate colour (gold done, yellow started,
  plain otherwise) with the breakdown in the `title` — no new column (A6).

**How, since 2026-09-11 — read [[status-list-is-truth]] (§19), which replaced the mechanism.** Status
lives in the SharePoint list **`Dashboard progress`**, one row per `JOB|ITEM`,
and `itemState`/`cpStatus` answer from it and from nothing else. A click
writes that row first, then paints the Excel cell (white 0 / yellow part /
gold done), then writes the `Dashboard Log` line. There is no `PENDING` hold
for a checkpoint and nothing to reconcile.

**How it used to work, until 2026-09-11.** Excel got only the colour and the
colour WAS the status; exact counts lived in the `Dashboard Progress` *sheet*;
a refresh let the Excel colour win over a stored count; and a `PENDING` hold
kept the office's own click alive for up to three minutes while the ~36 s-old
download caught up. That sheet is no longer written by anything and is left
exactly as it stands.

Pure logic (`cpItems`, `itemState`, `cpRow`, `cpDerivedOf`, write ordering
record→fill→log, bursts and a replayable queue) is in `checkpoints.js`; the
drawer UI is `cpSectionHtml`/`cpPatchSection`/`cpFoldHtml`/`cpDoorHtml`, and a
door is written by `setDoorStatus` (a word, not a count — `cpWriteItem` takes
`o.status` for it). **Clearing a job's glass also clears the floor's
counters** since 2026-09-10 — see [[glass-office-clear]] (§18). Tests: `test_checkpoints.js` and
`test_doors.js`.

## 6. Phase pipeline and hand-set phases

**What.** The primary progress line for a job: In office → Sent to floor →
Cutting → In fabrication → In glazing → Quality check → Fitted / delivered.

**How.** `jobPhase(j)` in `checkpoints.js` derives the phase from the sheet
(dates, cut colour, yellow, gold, ready, off the sheet; highest wins).
Clicking a step sets it by hand: stored in the SharePoint list `Dashboard
phases` (Title=job, unique; Phase; PhaseName; SetBy; SetAt), never in the
workbook; `effectivePhase = max(sheet-derived, hand-set)`. `readPhases()` in
`app.js` reads the list quietly and shows a plain "permission needed" state
if the tenant has not granted the list scope; the admin can grant it from a
click. Tests: `test_phases.js`, `test_phases_list.js`.

Since 2026-09-21 `effectivePhase` gained a third voice, the floor's own —
see [[glazing-station-phase-bar]].

## See also

- [[sections-and-row-moves]] — previous: sections, categories and row moves
- [[ui-extras]] — next: the selection wheel and versions
- [[status-list-is-truth]] — where checkpoint status actually lives since 2026-09-11
- [[glazing-station-phase-bar]] — the phase pipeline's third voice, the floor
