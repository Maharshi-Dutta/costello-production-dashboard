# Glass: one page per stage, no glazing, a new colour rule, an editable office board

**Status: approved to build, 2026-09-21** (owner's answers in chat that day;
the seven decisions are quoted under "Decisions taken").

## Context

The glass station shipped 2026-09-08 as one tablet page with three stages
(Cutting, Hotmelting, Glazing) plus the Tuff counter, and since 2026-09-10 the
office dashboard paints DG/TG/TUFF/NOT TUFF on the `Production` sheet from the
floor's counters: yellow when cutting and hotmelting are complete, gold when
glazing is complete.

The owner's correction, 2026-09-21: **glazing is the last step for the whole
job, not a glass step.** It must stop deciding whether glass is done, and it
leaves this station altogether (a glazing station of its own is a later brief,
not this one). Cutting and hotmelting each get a tablet of their own. The
owner likes the office's Glass station window and wants to work from it:
edit counters there and open the job from there.

## Decisions taken (owner, 2026-09-21)

1. The new colour rule applies to **every job at once**: a job yellow today
   (cut + hotmelted, not glazed) turns gold; a job with exactly one of the two
   stages complete, blank today, turns yellow. "thats ok change that."
2. "Done" for a stage means its counter has reached its total. Partial counts
   paint nothing.
3. NOT TUFF follows the DG/TG rule.
4. **No glazing page is built.** "i want to remove glazing option and rule
   from existing dashboard and code."
5. Glass is complete — the tablet's finished state, and what the office lock
   means — when cutting **and** hotmelting are complete.
6. The office may edit the glass counters from the board; every edit is logged
   in `Dashboard Log`. This is a new dated exception to rule 2.
7. Three physical tablets. Two are used by this brief (Cutting, Hotmelting).
8. **Do not change any data already recorded.** No counter, stamp, log line,
   note or people row is rewritten, reset or deleted by this change. The
   `Glazed`, `GlazedBy`, `GlazedAt` columns stay on the list with whatever is
   in them; nothing reads or writes them any more.

Tuff is unchanged: a fourth counter the cutter taps, its total carried from the
sheet, outside the glass total and outside the finished rule. (**Superseded in
part** by "Owner, after the demo" under the amendments below: it is still
outside the glass total, the lock and the header count, but a job with tuff
still to do is no longer *shown* as finished.)

## Hard rules (unchanged, re-stated because this brief touches all of them)

- The tablet pages never touch the workbook. Grep gate as always:
  `setFill|clearFill|setValues|appendLog|saveProgress|moveJobRow|batchWrite|/workbook`
  must not appear in `station-core.js`, `station-ui.js`, `station.js`,
  `glass.html`.
- The `Production` sheet is changed only by the sanctioned fills. This brief
  changes **when** the four glass columns (AY DG, AZ TG, BA TUFF, BB NOT TUFF)
  are painted, not which cells may be painted. BC–BF (ARCH, ASTRAGAL, FANCY,
  EXTRA) are never written.
- `Dashboard progress` (the status list) stays the truth; an Excel fill is a
  copy written from it and never read back as status.
- No real person's name, email or company domain anywhere. The person who cuts
  is "the cutter"; placeholders `Person A`, `example.test`.
- Code edits go through the Edit/Write tools, never a shell heredoc or a
  python write (HISTORY.md B16). Do not commit. Do not touch live SharePoint.
- The site pin is **not** changed in this brief: `ST.GLASS.site` stays `"own"`.
  Moving the glass lists to `Floor stations` is a separate, later step
  ("Not built here").

## A. Glazing leaves the station

- `STAGES` in `station-core.js` becomes Cutting and Hotmelting. Everything
  derived from it follows: `ALL_STAGES`, `ALL_STAGE_KEYS`, `STAGE_FIELD`,
  `STAGE_ROW`, `STAGE_BY`, `STAGE_AT`, `STAGE_TOTAL_ROW`, `CLEAR_WORD`,
  `counterFields`, the board record's `bars`, `cardSig`.
- `FLOOR_FIELDS`, `STATION_FIELDS`, `SEED_FIELDS`, `OFFICE_CLEAR_FIELDS` lose
  `Glazed`, `GlazedBy`, `GlazedAt`. Consequences, all intended:
  - the tablet can no longer write those columns (whitelist);
  - the feeder no longer seeds `Glazed`;
  - an office clear no longer writes `Glazed` — it leaves whatever is there;
  - reads no longer `$select` them.
- A queued tap for `glazed` sitting in a tablet's `cw_stationq` from before the
  new build must be **dropped quietly** on replay, not sent and not an error
  (the whitelist filter on the way out of the queue should already do this —
  prove it with a test).
- A person whose `Stages` holds `glazed` keeps the token in SharePoint (rule 8);
  the code ignores unknown tokens. A person left with no known stage shows on
  no page; STATIONS.md tells the owner to edit that row.
- **History stays readable.** `Station log` lines with `Stage = glazed` still
  show in the Floor log window and the drawer timeline, labelled "Glazing". Keep
  a display-only label for it (a `LEGACY_STAGE_LABELS` map or similar); it must
  not make `glazed` a tappable or filterable-as-current stage. The Floor log's
  stage filter may list it under the stages that exist in the rows on screen.
- The job **phase** strip in the drawer ("In glazing") is a different thing
  (`checkpoints.js` phases) and is not touched.

## B. The colour rule

In `station-core.js` (the block that starts "the colour the Production sheet
should be showing", `stageComplete`, `COLOUR_TYPES` and the plan function the
office's glass colour writer calls):

| column | blank | yellow | gold |
|---|---|---|---|
| DG, TG, NOT TUFF | neither stage complete | exactly one of cutting / hotmelting complete | both complete |
| TUFF | tuff count not complete | — | tuff count complete (and `tuffTotal > 0`) |

- Reversible exactly as today: a counter tapped back down walks gold → yellow →
  blank. "Complete" still needs a total to reach: a row with no glasses has
  finished nothing.
- `finished` on the board record = cutting and hotmelting complete (was: all
  three). `OfficeDone` semantics unchanged: the record says this job's glass
  is done; it locks the job on the tablets; only the office unlocks.
- **The seed** (`seedFields` / the function around the comment "Glazed stays at
  NOUGHT", used for a row the floor has never tapped): office record gold →
  `Cut = Hotmelt = total`; office record yellow → `Cut = total, Hotmelt = 0`
  (cutting comes first; one complete stage is what yellow now means); anything
  else → noughts. Untouched rows only, exactly as today.
- Everything else about the two-way arrangement (most recent action wins, ties
  to the office, the office's click discarding the writer's un-landed paint,
  `Dashboard Log` line per paint "Floor glass colours") is unchanged.
- The rehearsal (manager's job, not the implementer's) will run the new plan
  against a copy of the real workbook and the real list and report how many
  cells change on day one. The plan function must therefore stay **pure** and
  callable from Node, as it is now.

## C. One tablet page per stage

One page, `glass.html`, serving both tablets — no second HTML file:

- The stage comes from `?stage=cut` or `?stage=hotmelt` in the URL; failing
  that, from `localStorage` key `cw_stationstage`; failing that, the page shows
  a two-button chooser ("Cutting" / "Hotmelting") before the person picker and
  remembers the answer in `cw_stationstage`. A small "Cutting ▾" control in the
  header changes it (confirm first; it signs the person out).
- The header reads "GLASS · CUTTING" or "GLASS · HOTMELTING".
- **People picker:** only active people of station `Glass` whose `Stages` hold
  this page's stage. On the Cutting page, a person who holds `tuff` also shows.
- **Steppers on a card:** this page's stage only; on the Cutting page also Tuff
  when the job has tuff and the signed-in person holds `tuff`. A person's other
  stages are not shown on the wrong page even if their row holds them.
- **Header count and card "N left":** this page's stage only (Tuff stays out of
  it, as today).
- **A card is done on this page** (gold, sinks to the bottom) when this page's
  stage is complete — the cutter's finished jobs leave the cutter's way even if
  hotmelting has not started. The office lock (`OfficeDone`) still locks the
  card on both pages.
- Notes (`Station comments`) keep `Station = "Glass"` on both pages: no data
  model change, both tablets share the job's glass thread.
- `Station log` lines are unchanged (`Stage` already says which).
- Layout rule (owner, 2026-09-17): 10-inch tablet, portrait, one column, never
  scroll sideways, steppers in one fixed column.

## D. The office's Glass station board becomes a working surface

`stationBoardHtml()` and its wiring in `app.js`. Model it on the welding
board's office steppers (`weldOfficeEdit`, `wireWeldBoard`, `weldOfficeLineHtml`)
— reuse the pattern and any helper that can be shared; do not copy a second
implementation of clamping or of the log line.

- Each stage line on a card (Cutting, Hotmelting, and Tuff when the job has
  tuff) gets `−`, `+` and "All" controls. Clamp 0…total. One PATCH per click to
  the `Glass station` row: the counter, its `By` (the signed-in office account's
  display name) and `At`, plus `DoneBy`/`DoneAt`.
- Optimistic on screen, reconciled by the existing 10 s delta poll; a failed
  write puts the old number back and says so in a toast.
- **Logged in `Dashboard Log`** through the existing change path
  (`noteChange`), one line per click: job, "Glass cutting" / "Glass hotmelting"
  / "Glass tuff", from → to. **Not** written to `Station log` (that is the
  floor's record of the floor's work — same decision as welding, 2026-09-16).
- The glass colour writer then sees the new counters on its next run and
  paints by rule B. The office edit itself paints nothing directly.
- A job locked by `OfficeDone` shows its steppers disabled with the existing
  "glass complete — clear it in the job card" wording.
- **Clicking the card head (job number / customer) opens the job drawer**, the
  same drawer as from the job list (`state.sel = id; openDrawer()`); the
  unread-notes badge keeps its own click. A board job that is no longer in the
  workbook parse (left production) has no drawer: the head is not clickable
  and says nothing.
- The drawer's "Glass station" section shows two stage lines (and Tuff), no
  glazing line.
- `web/CLAUDE.md` rule 2 gains the dated exception: "2026-09-21, owner: the
  office may edit `Cut`, `Hotmelt`, `Tuff` and their By/At stamps on a `Glass
  station` row from the Glass station board; logged in `Dashboard Log`, never
  `Station log`."

## Not built here

- A glazing station. The owner will brief it separately.
- The move of the glass lists to `Floor stations`. After this build is live:
  a scratch script copies `Glass station` rows and the glass rows of `Station
  people` / `Station log` / `Station comments` into the `Floor stations` site
  exactly (rehearsed on a test site first), the station account is confirmed as
  a member there, and `ST.GLASS.site` becomes `"floor"` in a one-word ship.

## Tests to deliver

Update `test_station.js` and `test_glasscolour.js` to the new rule rather than
deleting coverage: every glazing-gate assertion becomes its new-rule
equivalent. New or changed checks must include:

- colour plan: neither / cut only / hotmelt only / both, for DG, TG, NOT TUFF;
  TUFF gold on its own count and blank otherwise; BC–BF never planned; walking
  back down; a row with total 0 plans nothing.
- a row carrying old `Glazed` values (e.g. `Glazed = total`, `Cut = 0`) plans
  **blank** — glazing data has no influence at all.
- `finished` = cut and hotmelt complete; `Glazed` irrelevant.
- seed: gold → cut + hotmelt; yellow → cut only; never `Glazed`.
- no request from the tablet or the feeder or an office clear ever carries
  `Glazed`, `GlazedBy` or `GlazedAt` (assert over every request body in the
  run).
- a queued `glazed` tap from an old build is dropped on replay without a
  request and without an error state.
- page stage: `?stage=hotmelt` shows only hotmelting steppers and only people
  holding `hotmelt`; the Cutting page shows Tuff for a tuff-holder; the
  chooser writes `cw_stationstage`.
- a legacy log line with `Stage = glazed` renders with the label "Glazing".
- office board edit: clamps, one PATCH with exactly the expected fields, a
  `Dashboard Log` line, **no** `Station log` POST, disabled under `OfficeDone`,
  old value restored on a failed write.
- the workbook gate: no `/workbook` request from the station files.

All existing suites stay green (`test_welding.js` included — welding must be
untouched by this brief).

## Docs to update

`docs/REFERENCE.md` (glass sections), `docs/STATIONS.md` (data model, what the
office sees, "Adding a person" — stages are now `cut`, `hotmelt`, `tuff`),
`docs/SUPPORT.md` (tell the office and the floor: **yellow now means one of
cutting / hotmelting is finished; gold means both**; glazing no longer appears
on the glass tablets), `docs/ARCHITECTURE.md` (`cw_stationstage`),
`web/CLAUDE.md` (rule 2 exception; the colour rule wording), `docs/specs/README.md`.
HISTORY.md is the manager's at ship.

## Amendments after review (2026-09-21)

Independent review of the first build (commit `48e921e`) plus the manager's
offline rehearsal on a saved copy of the real `Glass station` list (138 active
rows, 116 of them never tapped by the floor). All to be fixed before ship.

- **R1 — the feeder must never lower a seeded counter (decision 8).** The old
  build seeded an untouched row whose office record was yellow to
  `Cut = Hotmelt = total`; the new `officeSeed` answers `{cut: total,
  hotmelt: 0}` for the same record and `feedPlan`'s untouched branch writes any
  difference, in both directions — so the first load after ship would PATCH
  `Hotmelt` from the total back to 0 (7 rows on the saved copy, e.g. 43 → 0)
  and hand the hotmelting tablet finished work as work to do. Fix: on an
  untouched row a seed field is only ever **raised**, unless the whole seed is
  noughts (a real office un-tick, which still clears). Test the downward
  direction explicitly.
- **R2 — Tuff is outside the office lock.** The lock now lands the moment
  hotmelting finishes, not after glazing, so it used to freeze the Tuff stepper
  and `dropBlocked()` deleted queued tuff taps while the cutter was still
  working. `OfficeDone` locks the glass stages only: Tuff stays tappable on the
  Cutting page and on the office board, and queued tuff taps are never dropped
  by the lock.
- **R3 — two fast clicks on an office stepper.** `glassWriting[k]` is set after
  `await CW.hasListConsent()`, so two clicks can both pass the guard and both
  PATCH from the same base. Set the flag (and redraw) before the first await
  and clear it on every early return. `weldOfficeEdit` has the same ordering:
  fix it the same way, with `test_welding.js` counts unchanged.
- **R4 — a typo in `?stage=`** must fall back to the device's remembered stage,
  not to the chooser: only a valid URL value wins.
- **R5 — a Tuff counter nobody has tapped is no opinion.** Under the old rule
  glazing painted TUFF gold even with `Tuff = 0`; under the new one such a cell
  would be planned blank and an existing gold wiped. When `TuffAt` is empty the
  plan names **nothing** for TUFF (neither paint nor clear). Tapping Tuff down
  to nought on purpose still clears it.
- **Not code, told to the owner:** the two existing gates in `glassColourPlan`
  still apply to "every job at once" — a row the floor has never tapped, and a
  job whose office record is newer than the floor's stamp, keep their colour
  until somebody acts; and one office stepper click makes a row "touched" for
  good, handing that job's four glass cells to the counters (brief §D, by
  design).
- **Owner, after the demo (2026-09-21): a job that still owes TUFF is not shown
  as finished.** The owner's words: *"why are moving the jobs to complete when
  tuff is left? if job has tuff and is not done dont move it, if done then
  move."* A job with `TuffTotal > 0` whose `Tuff` count is not complete is not
  drawn gold and does not sink into the Finished group — on the office's Glass
  station board, on the job row's `Glass 8/16` chip, and on the **cutting**
  tablet, whose bench the tuff is. The **hotmelting** tablet is unaffected: tuff
  is not that bench's work and never appears on it. A job with no tuff finishes
  on the glass stages as before. One pure helper, `ST.tuffOwed(g)`
  (`tuffTotal > 0 && !stageComplete(g, "tuff")`), used in `buildJobs` and in
  `station.js`'s `boardNow`.

  **This reverses one clause of the 2026-09-10 decision** ("Tuff … outside the
  glass total and outside the `finished` rule") and **only** that clause. It is
  a display rule: `tuffOwed` decides nothing about the colour the `Production`
  sheet is painted (`glassColours` reads the counters direct), nothing about the
  office's lock (`officeComplete` asks the office's own DG and TG checkpoints
  and has never known about tuff — see R2), nothing about the feeder's `Active`,
  seed or hash, and nothing about the header's "N left", which is this page's
  stage alone. `test_station.js` carries an assertion for each of those four
  saying the same row with and without tuff owed answers identically.

## Report back

Pasted output of every suite; functions added/changed with line numbers;
every place `glazed` still appears in code and why it is allowed to; the exact
fields an office board click PATCHes; anything in this brief you could not do
as written, and why.
