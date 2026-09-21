# The glazing station, and the phase bar hearing the floor

**Status: approved to build, 2026-09-21** (owner's answers in chat that day),
**except section F**, which is waiting on the owner's choice and must not be
built yet.

**Builds on** `2026-09-21-glass-split-no-glazing.md` and
`2026-09-21-day-sheets-and-station-reports.md` (both in the working tree).

## Context

On 2026-09-21 the owner took glazing out of the glass station: "glazing is the
last step for the whole job, not just glass, so it will be its own dashboard
with its section." The same day, after the demo: "make the glazing dashboard as
well, and I should be able to see what is glazed in the glaze station window in
the main dashboard. I should be able to edit glazing from that window, and when
glazing is in process or done the phase bar in the master dashboard should move
the progress into the In glazing section. And when something is in welding,
in process or done, it should be In fabrication. If something is marked in
glazing it means it has gone through fabrication; if glazing is not marked but
welding is, it is in fabrication."

This is the **third station**. `docs/STATIONS.md` § "Adding a station" is the
checklist and welding (`welding-core.js`, `welding.js`, `welding.html`, the
`weld*` block in `app.js`, `test_welding.js`) is the worked example. Follow the
checklist; do not copy a page — what two pages share goes in `station-core.js`
or `station-ui.js`.

## Decisions taken (owner, 2026-09-21)

1. The glazer counts **one number per job**: units glazed, out of the job's
   windows plus doors.
2. The board carries **every job in the In production section**, the same scope
   as welding.
3. The office edits glazing from the Glazing station window; every edit is
   logged.
4. The phase bar hears the floor (section E).
5. What happens when glazing reaches its total — the owner's words: "Ready to
   deliver, and it moves up to the ready to deliver section" — is **section F,
   not decided in detail yet**.

## Hard rules

- The tablet never touches the workbook. The grep gate
  (`setFill|clearFill|setValues|appendLog|saveProgress|moveJobRow|batchWrite|/workbook`)
  covers the new station files (`glazing-core.js`, `glazing.js`,
  `glazing.html`) as well as the existing ones.
- **Nothing in sections A–E writes the workbook**, from either side, except
  the `Dashboard Log` line the existing `noteChange` path writes for an office
  edit.
- **Read the `Production` sheet only** (owner's standing rule, 2026-09-18:
  `Production (2)` and the other sheets are not real data). Quantities come
  from what the parser reads off `Production` for the job — use the same
  Production-only source welding uses (`j.prodsMain` and the job's own
  Windows / Doors quantities as parsed from `Production`); never the
  cross-sheet merge.
- Rule 3 applies to everything the feeder puts on the list: the customer name
  and any comment go through `ST.stripContact`.
- No real person's name, email or domain. The person is "the glazer".
- No list is created by code; a missing list is a quiet explained state that
  asks again later (the `Station comments` / welding pattern).
- Edit/Write tools only for repo files. Do not commit. No live SharePoint.
- Layout: 10-inch tablet, portrait, one full-width column (`minmax(0,1fr)`),
  never scroll sideways at 700–1100 px, steppers in one fixed column, tap
  targets ≥ 40 px. (The glass page had the `1fr` min-content trap until
  2026-09-21 — do not reintroduce it.)

## A. Data model — `Glazing station`, in the `Floor stations` site

One row per job. `site: "floor"` in the definition (the site exists since
2026-09-17; welding lives there).

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, unique | feeder | the job number, upper-case |
| Job, Customer, Section, Seq, Active | text | feeder | job facts, as welding feeds them (customer stripped per rule 3) |
| Windows, Doors | number | feeder | the job's quantities on the `Production` sheet |
| Total | number | feeder | `Windows + Doors` |
| Comment | multiple lines | feeder | the sheet's COMMENT, stripped per rule 3, as welding does |
| FedAt, FedBy | text | feeder | when / whose dashboard |
| Glazed | number | tablet; office board | units glazed, 0…Total |
| GlazedBy, GlazedAt, DoneBy, DoneAt | text | tablet; office board | who / when, as on the other stations |

- No seed: an untouched row starts at nought. (There is no office record of
  glazing to seed from. Say so in STATIONS.md.)
- A job with `Total = 0` is not fed.
- The feeder never writes the floor's columns, never deletes, deactivates
  (`Active = No`) a job that has left In production — all exactly as welding.
- `Station people` rows: `Station = Glazing`, `Stages = glaze`. `Station log`
  lines: `Station = Glazing`, `Stage = glaze`. `Station comments`:
  `Station = Glazing`. All three lists are the shared ones already in the site.

## B. The definition and the core — `glazing-core.js`

Pure, loads in Node and the browser, exports `GLZC` with `GLZC.GLAZE`, the
station definition (every field in the STATIONS.md table, plus the report
fields the 2026-09-21 report brief added: `stageLabel`, `reportStages`,
`reportLogStages`, `reportJobs`, and no `daySheets`). The slice (In production
jobs with `Total > 0`), the board/card shape, the tap clamp, the write body
(`{Glazed, GlazedBy, GlazedAt, DoneBy, DoneAt}` and nothing else), `finished`
= `Glazed >= Total > 0`.

Reuse from `station-core.js`: `feedPlan`, `sliceHash`, `floorOnly`, `logFields`,
`stationComments`, `stationPeople`, the delta merge, `boardDiff`,
`stripContact`, `atCmp`, `stClock`. If welding-core has a helper glazing needs
too, move it to `station-core.js` and have both call it (welding's 61 checks
unchanged).

## C. The tablet — `glazing.html` + `glazing.js`

Same shell as welding's page through `station-ui.js` (theme, sign-in gate,
"Who are you?", PIN pad). One card per job: job number, customer, "N units"
(with "W windows · D doors" under it), the comment, one stepper (− n + All),
"N left", the note composer (`ST.stationComments`, station `Glazing`). Header:
"GLAZING", the person, "N left" across the board, Switch person, theme, Sign
out. Finished jobs sink to a Finished group. Offline queue, 10 s delta poll,
build check, idle lock — as welding.

## D. The office — the Glazing station window

- `STATIONS` gains `["glazing", "Glazing station"]`; Show ▸ Glazing station.
- The board: one card per job, the counter with **− + All None** for the
  office (welding's `weldOfficeEdit` pattern, **with the writing flag set
  before the first await**), clamped, one PATCH of the five floor fields, one
  `Dashboard Log` line ("Glazing", from → to), never a `Station log` line.
  Card head opens the job drawer. Unread-notes badge and the notes under the
  card as on the other boards. **Floor log** and **Report** chips. Gold card,
  sorted last, when finished.
- The job drawer gains a read-only **Glazing** line (n of Total, who, when)
  under the Welding line, and its timeline lines.
- The job row shows nothing new (owner's rule: no new columns).
- The feeder hook and the poll follow welding's (`weldFeed…`, `weldPoll`,
  `weldReadIfNeeded`): its own try/catch so a glazing failure can never stall
  the glass colour writer or the welding feed.
- `web/CLAUDE.md` gains the dated exception: the office may write the five
  floor fields of a `Glazing station` row from the board; logged in `Dashboard
  Log`.

## E. The phase bar hears the floor

`checkpoints.js` owns phases (`effectivePhase`: the higher of the hand-set phase
and what the sheet shows). Add a third voice, **the floor**, as a pure function
there:

- any welding recorded on the job (any `Welding station` row of that job with
  frames or sashes done > 0) → at least **In fabrication**;
- any glazing recorded (`Glazed > 0`) → at least **In glazing**;
- glazing outranks welding ("if something is marked in glazing it has gone
  through fabrication");
- the result is the **highest** of hand-set, sheet and floor — the floor can
  only move a job forward, never back past what the sheet or a person says;
- a counter tapped back to nought withdraws that voice (the phase falls back to
  the next highest) — computed live on every render, **stored nowhere**: no
  list write, no workbook write, no `Dashboard phases` row.

`app.js` passes the floor's records in wherever `effectivePhase` is called (the
drawer's phase strip, the home list's status word / colours, the phase
pipeline, the export's status column, alerts digests if they read the phase —
grep every caller and decide each on purpose; an export must say the same
phase the screen says). The drawer's strip says where the phase came from when
it is the floor's ("from the glazing station" / "from the welding station"),
the way it already distinguishes hand-set from sheet.

Welding and glazing lists may not have been read yet when the list first
draws: the phase must not flicker backwards and forwards — draw from what is
known, and repaint through the existing quiet path when the floor lists
arrive (extend the `chipsNow()` signature so a change in a job's floor phase
repaints rows; do not add a timer).

## F. NOT TO BE BUILT YET — glazing complete → Ready to deliver

The owner's words: when glazing reaches its total, the job is "Ready to
deliver, and it moves up to the ready to deliver section." Moving a job to the
Ready to deliver section is a **mark-ready fill plus a whole-row move on the
`Production` sheet** — the heaviest write the dashboard has (insert, copy,
verify, delete). The manager has put two ways of doing it to the owner
(automatic, or one click in the office); until the owner chooses, build
nothing here. **Leave the seam:** `GLZC` exposes `finished`, and the office
board already knows which jobs are complete; nothing else is needed yet.

## Tests to deliver — `test_glazing.js`, in the style of `test_welding.js`

- slice: In production only, `Total = Windows + Doors` from the Production-only
  source, `Total = 0` not fed, customer and comment stripped (a phone number in
  three shapes and an eircode never reach a feed body).
- feeder: adds, patches only feeder fields, never a floor field, never a
  delete, deactivates a job that left the section, never lowers anything on a
  touched row; its failure does not stop the glass or welding feed.
- tablet: tap clamp 0…Total, the five-field PATCH, one `Station log` line with
  `Stage = glaze`, offline queue replay, a person without `glaze` cannot tap,
  the workbook gate over the three new files, no request to `/workbook` in the
  run.
- office: steppers incl. None, one PATCH, one `Dashboard Log` line, no
  `Station log` POST, double-click → one PATCH, failed write restores the
  number, card head opens the drawer, drawer Glazing line.
- phases: welding progress alone → In fabrication; glazing progress → In
  glazing even with no welding; sheet or hand-set phase further on wins; tapped
  back to nought withdraws; nothing is written (assert zero list/workbook
  writes across the phase tests); the export's status equals the screen's.
- report: `stationReport(GLZC.GLAZE, "glaze", …)` produces Summary, Days, Jobs,
  Activity, Notes with no glazing-specific branch in `export.js`.
- all existing suites green with unchanged counts.

Add `test_glazing.js` and the three new source files to the verification
command in `web/CLAUDE.md`. `build.py` must stamp `glazing.html`'s script tags
— check how it finds pages and add the page if it is a list.

## Docs to update

`docs/STATIONS.md` (the Glazing station data model with the owner's list
recipe; people rows; what the office can do; "Adding a station" gains anything
learned), `docs/REFERENCE.md` (new section), `docs/SUPPORT.md` (the glazer's
page; what the phase bar now means and that the floor only ever moves it
forward), `docs/ARCHITECTURE.md` (any new localStorage key; the third feeder),
`web/CLAUDE.md` (rule exception, pages list, checks), `docs/specs/README.md`.

## Report back

Pasted suite output before and after; files added; functions added/changed
with file:line; the exact feed body and tap body; every caller of
`effectivePhase` and what you decided for each; the list recipe (columns and
types) exactly as the owner's script must create it; anything not done as
written, and why; anything that looks wrong but is outside the brief (list, do
not fix).
