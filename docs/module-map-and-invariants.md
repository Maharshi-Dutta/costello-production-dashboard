# Module map, key invariants, and where state lives

## Module map

| module | responsibility |
|---|---|
| `graph.js` | MSAL sign-in, token acquisition (with the quiet/popup split above), all Graph HTTP calls (`call()`), workbook session handling and retry-on-`InvalidSession`, `findFile()`/download/version history, the dashboard-owned-sheet upserts, `moveJobRow()` and its supporting row-capture/write/border functions, the SharePoint list layer (`listId`/`listItems`/`listUpsert`/`listDelete`/`listAdd`/`listPatch`/`listDelta`), batching (`batchGet`/`batchWrite`/`batchRun`) |
| `parser.js` | `parseWorkbook()`: turns a downloaded workbook into the job model; **two sets of product counts per job — `j.prods`, the maximum across every sheet (what every older reader uses), and `j.prodsMain`, the `Production` sheet's alone, which is what a station feeds from (2026-09-17, HISTORY B20)**; **the five DOORS DONE cells** (`mapSheet().doors`, `doorCode()` — any non-empty text is a door — into `j.doors` as `{slot, code, status}`); section/divider detection (`blocksFromValues`); fill-colour reading and the Cut/process/done ranking; **font-colour reading and the row colour code** (`fontOf`/`flagOf` → `j.flag`/`j.flagHex`, by hue range, hyperlinks and the two hyperlink theme colours excluded); **`parseJohnSheet()`** — "Production (2)" read on its own terms for the John print sheet, and never merged into the job model; `mapSheet()` header detection; `templateForJob()` (captures formatting for a move) |
| `checkpoints.js` | pure logic only, no DOM: the doors of a job (`cpDoors`/`cpDoorAt`, one item per DOORS DONE cell since 2026-09-14); the Windows/Doors aggregate that sits alongside those two lines' own record rows (`cpOwnState` / `cpDerivedOf`, the higher of the two shown, the paint upwards only) and the doors quantity warning (`cpDoorWarning`); the record of checkpoint status (`cpRow`/`cpRowsFrom`/`itemState`, the `Dashboard progress` list since 2026-09-11 — **not** the Excel colour, which is `cpFileStatus`), colour choice, the per-(job,item) tap debounce and its `localStorage`-backed queue, the two write sequences (`cpWriteItem`/`cpWriteGroup`, record→fill→log), the one-time import and the doors' own (`cpImportPlan`, with its per-kind filter) and the phase pipeline (`jobPhase`, `effectivePhase`, `PHASES`, and since 2026-09-21 the floor's third voice — `floorPhaseOf`/`setFloorHook`/`floorPhase`/`phaseSource`, computed on every render and **stored nowhere**) |
| `export.js` | pure logic only, no DOM, no network: filtering (`exportFilter`), row shaping, filename/summary text, building an ExcelJS workbook and a pdfmake document definition; §7b holds the fixed **John print sheet** (`exportJohnRows`/`buildJohnWorkbook`/`buildJohnDoc`), built from Production (2)'s own rows and the one export the owner sanctioned to carry a phone number; the eircode exclusion, and the phone exclusion from every other template, live here as hard rules with standing tests |
| `app.js` | everything with a DOM: rendering the job list/drawer/windows, wiring every tap to a write, `PENDING`/`PENDV`/`PENDA` optimistic holds, sign-in/session boot, the phase-list read/write UI, the Export/Alerts/Versions windows, the radial selection menu, build-freshness polling |
| `station-core.js` | pure glass-station logic (shipped 2026-09-08) — `glassTotal`, `officeSeed`, `glassSlice`, `feedPlan`, `sliceHash`, `jobBoard`, `boardFilter`, `applyTap`, who may record what (`stationPeople`/`canStage`/`pinOk`/`personExpired`), the write and log shapes (`floorOnly`/`tapFields`/`logFields`), reading the log back (`logRows`/`logFilter`/`logCounts`/`logLast`), and the two functions that keep a ten-second poll cheap (`mergeDelta`, `boardDiff`) — no DOM, no Graph, loads in Node and the browser like `checkpoints.js` |
| `station.js` | the glass station page UI (shipped 2026-09-08), using `CW` from `graph.js` and the pure functions from `station-core.js` |
| `station-ui.js` | the floor tablet's **shell**, station-independent (2026-09-16): the device's light/dark choice, "how long ago" in words, the sign-in gate, and the "Who are you?" picker with its PIN pad. Everything is prefixed `stu`/`STU` so a page can load it beside any station's script. Used by `welding.js`; `glass.html` loads it and `station.js` has not moved onto it yet (its test harness runs `station.js` in a context that does not load this file) |
| `welding-core.js` | pure welding-station logic (shipped 2026-09-16): the `WELD` definition, the deny-list of product groups, the rule-3 comment strip (`weldStripDigits`), the slice, the records and cards, the three-level colour rule, the tap clamp, the write bodies (`weldTapFields`/`weldOfficeFields`/`weldFloorOnly`), the queued-tap re-base and the card signature — no DOM, no Graph, no workbook, self-contained |
| `welding.js` | the welding station page UI (2026-09-16), using `CW`, `ST`, `STU` and `WELDC` |
| `glazing-core.js` | pure glazing-station logic (2026-09-21): the `GLAZE` definition (**no `seedFields`, no `daySheets`**), the slice off the `Production`-only quantities, the records and cards, the three-level colour rule, the tap clamp, the write bodies (`glzTapFields`/`glzOfficeFields`/`glzFloorOnly`, the last of which is `ST.floorOnly` with this definition), the queued-tap re-base, the card signature and the report adapter — no DOM, no Graph, no workbook. Borrows from `station-core.js` rather than repeating it |
| `glazing.js` | the glazing station page UI (2026-09-21), using `CW`, `ST`, `STU` and `GLZC` |
| `build.py` | stamps a timestamp build id onto every script tag's `?v=` and into the footer of all four pages, writes `version.json` for the freshness poll |

## Key invariants

- **Rows are re-found by job number before every write.** `rowForJob(sheet,
  jobId)` reads column C fresh and never trusts a remembered row number,
  because rows move (section moves, and anyone editing the sheet directly).
  `moveJobRow()` re-locates both the source and target positions at each
  step and verifies the row count after writing before it will delete the
  original.
- **Borders follow the shared-edge rule.** Excel keeps one border definition
  per shared edge between two rows; writing an edge on a row makes Excel
  strip that edge from its neighbour. So a moved row gets its bottom edge and
  verticals from its own template (never its top), and the row that stays
  where it was has its bottom edge explicitly re-asserted afterwards
  (`restoreBottomEdge`) — the line "belongs" to whichever row didn't move.
- **The text-date apostrophe guard.** A text cell whose content looks like a
  number, a formula, `TRUE`/`FALSE`, or a date fragment (`guardText()` in
  `graph.js`) is written with a leading apostrophe, Excel's own way of
  forcing text interpretation, so a job note like `07/04` doesn't silently
  become a date on the next write.
- **Only fills are ever written to `Production`, for exactly three sanctioned
  reasons** (repo rule 1): a checkpoint colour, a glass colour from the floor,
  and — since 2026-09-14 — **one of the five DOORS DONE cells** of a job's
  row, plus the Windows (M) and Doors (N) cells, which carry the higher of
  their own record row and what sits under them and are only ever painted
  upwards — nothing under the line whitens either. Every one of those is a **copy** of a `Dashboard
  progress` row (or, for M and N, of that row and the aggregate of the rows
  below it), written
  after the record and never read back as status. **The door's code text is
  never written**: the fill of that cell is the only thing this app may touch
  there.
- **Contact details leave the app in exactly one place.** The eircode
  (`j.eir`) is never read by any export path, in any format. The phone number
  is carried by the **John print sheet alone**, sanctioned by the owner on
  2026-09-09: it comes from Production (2)'s own Phone column via
  `parseJohnSheet`, with `j.ph` read only by `exportJohnRows` and only for a
  job Production (2) does not have. The Default template's builders read
  neither, and every John print's `Dashboard Log` line says in words that the
  file has phone numbers in it.
- **Production and Production (2) are two sheets, not one.** Production is the
  job model, its colours and its status; Production (2) is the paper John
  works from. Neither is read for the other: the John print sheet view and the
  John print come from `parseJohnSheet` alone, and the job model comes from
  `parseWorkbook` alone.
- **Writes are serialised per sheet (and per list).** `serialised(key, fn)`
  chains every write against a given dashboard sheet (or `"list:" +
  displayName` for a SharePoint list) into one promise queue, because every
  upsert is "read the used range, work out the target row, write it" — two
  of those racing on the same sheet would silently clobber one write with
  the other. An earlier failure in the chain never blocks the next write.

## Where state lives

| kind of state | where |
|---|---|
| the job model itself | never stored — always re-derived from the downloaded workbook on each `load()` |
| checkpoint status — what is ticked off, per job per item | the SharePoint list `Dashboard progress` (2026-09-11). Read in full at boot, then `listDelta` every 10 s; written at the click, before the Excel colour. Nothing about it lives in this browser |
| the colour this dashboard last painted into each managed cell | `cw_painted`, so the safeguard can tell a hand-paint in Excel from a stale download, and so a fill the workbook refused is painted again on the next load |
| whether the one-time import of the sheet's colours has drained in this browser | `cw_cpimported`; until it has, an item with no row reads as the sheet's colour so nothing on any screen changes during the switch-on window |
| ... and the same question for the door cells alone | `cw_cpimported_doors` — the doors shipped three days after the general import had drained, so they get one pass of their own. It does not reopen the colour fallback; it only keeps the safeguard off an unrecorded door cell until that pass is done |
| a person's own unconfirmed edit (section move, mark-ready, product status, alert, hand-set phase) | in memory + `localStorage`, for up to 180 s, until the downloaded file (or list) agrees — `cw_pending`, `cw_pendv`, `cw_penda`, and the `phase` field inside `cw_pending`. **Neither a checkpoint tick nor a glass colour is one of them** since 2026-09-11: both go straight onto the record, and `cw_pending` can no longer gain a `cp` or `gc` key at all |
| exact checkpoint counts | `Dashboard Progress` sheet (Excel keeps only the cell colour) |
| the audit trail | `Dashboard Log` sheet — every write anywhere in the app appends one line here |
| custom groupings / saved views | `Dashboard Views` sheet |
| email subscriptions | `Dashboard Alerts` sheet |
| admin address and other read-only config | `Dashboard Config` sheet (read by the dashboard, never written) |
| hand-set phase per job | `Dashboard phases` SharePoint list (not the workbook) |
| the extra note printed on a John print | `Dashboard print notes` SharePoint list (not the workbook) — shipped 2026-09-09 |
| the rows John's paper is printed from | never stored — re-read from the `Production (2)` sheet of every download into `JOHNROWS` |
| glass station job facts and floor counters | `Glass station` SharePoint list, in the separate `Floor stations` site — shipped 2026-09-08 |
| who may record which stage on the floor | `Station people` SharePoint list, same site — shipped 2026-09-08 |
| who moved which counter, and when | `Station log` SharePoint list, same site; written by the tablet, read by the master, never deleted from — shipped 2026-09-08 |
| a note the floor left against a job | `Station comments` SharePoint list, same site; written by a floor tablet, read by the master in that job's drawer, append-only and never deleted from — 2026-09-15 |
| a note typed on the tablet but not sent yet | nowhere but the page's own memory: the draft lives in the channel's state and is re-drawn from it, and a refused send keeps it in the box until Send is tapped again |
| a queued/not-yet-sent checkpoint tap | `cw_cpqueue`, replayed on the next page load |
| a queued/not-yet-sent station tap | `cw_stationq` (counters) and `cw_stationlogq` (the log lines they owe) — tablet only, shipped 2026-09-08 |
| a queued tap the office's lock arrived under, so it was dropped rather than sent | `cw_stationblocked` — tablet only, drawn on the card in red until the office unlocks the job (2026-09-10) |
| who is at the station tablet, and when they last tapped | `cw_person` — tablet only; the stages always come back from the list, never from storage |
| **which of the two glass tablets this device is** | `cw_stationstage` — tablet only, `cut` or `hotmelt` (2026-09-21). A `?stage=` in the URL wins over it and is written to it; a device holding neither is asked once, on screen, before the person picker |
| when the master last fed the `Glass station` list, and the hash of what it sent | `cw_stationfeed` — read by `feedStation()` to decide whether a run can be skipped |
| the cutter's end-of-day sheets, and the weekly target they are measured against | `Station day sheets` and `Station targets` SharePoint lists, same site as the station's other lists (resolved through the definition's `site`) — one row per person per station-stage per day, and one row per station-stage. Append-only from the tablet; the office may correct the counts and the note, and is the only writer of the target — 2026-09-21 |
| a day sheet typed on the tablet but not saved | `cw_daysheetdraft` — tablet only, `{day, stage, counts, note}`. Dropped at the change of day and on save: yesterday's typing is not today's sheet (2026-09-21) |
| a saved day sheet this tablet still owes | `cw_stationdayq` — tablet only, one entry per Title, rebuilt through `ST.dayFields` on load so an edited storage can put no column on the wire that a save could not. Drained by `flushDay()` inside the page's existing `flushQueue()` (2026-09-21) |
| welding station job facts and floor counters | `Welding station` SharePoint list, in the `Floor stations` site — one row per job **and product group** — 2026-09-16 |
| a queued/not-yet-sent welding tap | `cw_weldq` (counters) and `cw_weldlogq` (the log lines they owe) — the welding tablet only, deliberately separate from the glass tablet's two so one page can never read the other's queue (2026-09-16) |
| who is at the welding tablet, and when they last tapped | `cw_wperson` — the welding tablet only |
| which tab the welding tablet's board is on, On floor or Finished | `cw_weldtab` — `"floor"` or `"finished"`, the device's own choice, On floor by default (replaced `cw_weldsent` and the "Sent to floor" chip, 2026-09-24) |
| when the master last fed the `Welding station` list, and the hash of what it sent | `cw_weldfeed` — read by `feedWelding()` to decide whether a run can be skipped |
| glazing station job facts and the floor's counter | `Glazing station` SharePoint list, in the `Floor stations` site — one row per **job** — 2026-09-21 |
| a queued/not-yet-sent glazing tap | `cw_glzq` (counters) and `cw_glzlogq` (the log lines they owe) — the glazing tablet only, deliberately separate from the other two tablets' queues so one page can never read another's (2026-09-21) |
| who is at the glazing tablet, and when they last tapped | `cw_glzperson` — the glazing tablet only |
| when the master last fed the `Glazing station` list, and the hash of what it sent | `cw_glzfeed` — read by `feedGlazing()` to decide whether a run can be skipped |
| where a job has got to, as the FLOOR sees it | **nowhere** — `effectivePhase`'s third voice is computed on every render out of `WELD_ITEMS` and `GLZ_ITEMS`, which are already in memory. No list row, no workbook cell, no `Dashboard phases` write (2026-09-21) |
| workbook and list ids, once resolved | `cw_fileref`, `cw_listids`, `cw_stationsite` (the legacy resolver, which nothing in the app asks for any more), and one key per pinned channel: `cw_stationsite_own` (glass, the workbook's own site) and `cw_stationsite_floor` (welding **and glazing**, which share the one `Floor stations` channel) — 2026-09-16 |
| export presets (filters/fields/format only, never job data) | `cw_exportpresets` |
| whether the job card's Windows and Doors folds were left open | `cw_cpopen` — one flag per fold, per browser; nothing about a job is in it |
| which floor notes the Changes panel has already announced | `cw_notesseen` — the ids of announced `Station comments` rows, capped at 500, seeded silently on a browser's first look (2026-09-15) |
| which jobs' floor notes a person on this screen has opened | `cw_notesread` — `{ JOB: { at, ids } }`: the `At` of the newest dated note seen and the ids of undated ones; drives the unread icon on the job row and the station board card; per browser, never written to SharePoint (2026-09-16) |
| UI-only preferences | `cw_theme`, `cw_hidden`, `cw_collapsed`, `cw_changes`, `cw_stationtheme` (the tablet's own light/dark, dark by default) |

## See also

- [[station-feeder]] — previous: the station feeder and the station page
- [[overview-and-process]] — the documentation set and test suite this module map underlies
- [[glazing-station-phase-bar]] — the newest addition to `effectivePhase`
