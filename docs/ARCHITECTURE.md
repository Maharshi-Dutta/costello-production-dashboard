# Architecture

## Data flow, in text

```
  SharePoint site (ProductionProgress)
        |
        |  1. sign in (MSAL, Entra ID)                     <- graph.js
        |  2. findFile(): locate the workbook by name
        |  3. downloadWorkbook(): GET the whole .xlsx       <- graph.js
        v
  ExcelJS workbook in memory (full fidelity: values, fills, formats)
        |
        |  parseWorkbook(wb)                                <- parser.js
        v
  job model: one plain object per job (id, cust, dates, prods,
  glass, cp, blk, seq, cat, ...)
        |
        |  applyPending() merges in this browser's own
        |  optimistic writes (see "the 36 s lag" below)      <- app.js
        v
  ALL (in-memory array) --> rendered as the job list, the drawer,
                             the phase pipeline, the export window
        |
        |  a tap / a form submit
        v
  a write, always through the Excel API in a workbook session:
    - setFill() on one cell (checkpoints)                    <- graph.js
    - moveJobRow() (the only structural write: insert, copy,
      verify, delete)                                        <- graph.js
    - upserts against Dashboard Log / Views / Progress /
      Alerts (own sheets, created on first use)               <- graph.js
    - upserts against the Dashboard phases / Glass station
      SharePoint lists (never the workbook at all)            <- graph.js
```

The workbook is fetched whole because that is the only way to read cell fill
colours cheaply for every job at once; every write goes back through the
targeted Excel API (`PATCH .../format/fill`, `PATCH .../range(...)`,
`/insert`, `/delete`) so the file itself is never rewritten wholesale, except
`restoreVersion` — the one deliberate whole-file operation, used only to roll
back to an earlier SharePoint version.

### The 36-second file lag and `PENDING` holds

SharePoint takes roughly 35 seconds to fold a change made through the Excel
API into the file you'd get from `downloadWorkbook()`; the Excel API itself
reflects the same change in about a second. If the dashboard only trusted
what it downloads, a person's own edit would appear to work and then vanish
the next time the page refreshed and re-downloaded a still-stale file.

The fix is `PENDING` (`app.js`): every optimistic write is held in memory
and in `localStorage` (`cw_pending`) for up to 180 seconds (`PENDING_MS`),
each field tracked with its own timestamp. `applyPending(list, fresh)`
overlays these holds onto whatever was just parsed; when a *freshly*
downloaded file agrees with a held value, that hold is dropped early so
everyone else's later edits aren't shadowed forever. The same pattern
repeats for section moves (`PENDV` / `cw_pendv`), alerts (`PENDA` /
`cw_penda`), and hand-set phases (`p.phase` inside `PENDING`).

### Dashboard-owned sheets

Created on first write, addressed only through a hard-coded constant so no
call site can ever be pointed at `Production` by accident:

| sheet | constant | written by |
|---|---|---|
| `Dashboard Log` | `LOG_SHEET` | `appendLog()` — every change, one line each |
| `Dashboard Views` | `VIEWS_SHEET` | `saveAssignment()` / `clearAssignment()` — custom groupings |
| `Dashboard Progress` | `PROGRESS_SHEET` | `saveProgress()` / `saveProgressMany()` — checkpoint counts |
| `Dashboard Alerts` | `ALERTS_SHEET` | `addAlert()` / `removeAlert()` — email subscriptions |
| `Dashboard Config` | `CONFIG_SHEET` | **read-only** — never created or written by code; holds the admin address, allowed export/company text, etc. |

### SharePoint lists and the scope split

Two kinds of Graph scope exist in `graph.js`:

- `SCOPES` (`Files.ReadWrite.All`, `User.Read`) — sign-in and every workbook
  call. Every dashboard user has these from the moment they sign in.
- `LIST_SCOPES` (`Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read`)
  — needed only for SharePoint list requests. Asked for quietly
  (`token(scopes, quiet=true)`) so a background list read never pops a
  consent dialog; asked for with a popup only from a genuine click
  (`listConsent()`), so a tenant that hasn't granted `Sites.ReadWrite.All`
  yet still signs in and uses everything workbook-related — the list
  features simply report that they need permission (`PHASE_NEED_CONSENT`).
  `headers()` picks the right scope set and drops the workbook session
  header for any path containing `/lists`, since a stale workbook session id
  on a list request earns an unrelated `InvalidSession` error.

Lists in use: `Dashboard phases` (hand-set phase per job) and the floor's
three, shipped 2026-09-08: `Glass station` (job/glass facts plus the floor's
counters), `Station people` (who may record which stage) and `Station log`
(one line per counter write) — see `docs/STATIONS.md`.
`Dashboard phases` lives in the same SharePoint site as the workbook. The
floor's three are meant to live in a *separate* site, `Floor stations`, so the
station account never needs any permission on the workbook's own site — but
creating that site needs an administrator the owner has not got yet, so for now
they live in the workbook's site too and the station account is a member of it
(which does mean it can reach the workbook; see `docs/STATIONS.md`).
`stationSite()` resolves whichever of the two is there, preferring `Floor
stations`, caches the answer in `cw_stationsite` with which site it is, and
re-checks every ten minutes while it is on the fallback so the pages move over
by themselves when the real site appears. The fallback is resolved by a plain
site lookup, never through `findFile()`: nothing on that path touches `/drive/`
or `/workbook`.

Both screens keep those lists current with `listDelta()` rather than
re-reading them: the first call enumerates and hands back a `deltaLink`,
every call after it passes that token and gets only what moved. A 4xx from
the delta endpoint — including the 410 Gone that carries a
`resyncChanges...` code when a token is too old — is re-thrown as one
recognisable error (`isDeltaRestart`), and the caller answers it by reading
the list once and starting a fresh enumeration.

### The station feeder and the station page

Shipped 2026-09-08. The data flow, in one line: the **feeder** (inside the
master dashboard) writes job/glass-type facts into `Glass station`; the
**tablet** writes back only its own counters plus one `Station log` line per
counter write; the **office** reads `Glass station` and `Station log` by
delta (and `Station people` once, for the log window's filters) and writes
none of the three.

- **Feeder → `Glass station`.** After every successful `load()` of the
  workbook, `feedStation()` (`app.js`) turns the freshly parsed jobs into a
  slice (`ST.glassSlice`, one row per job-and-glass-type, jobs "in
  production" only) and a plan of adds/patches (`ST.feedPlan`), then sends
  it. It skips the run entirely when the signed-in account has not granted
  the SharePoint list permission, or when the slice's hash
  (`ST.sliceHash`) matches the last run and that run was under ten minutes
  ago. A job that leaves production is patched `Active = No`, never
  deleted; the feeder never sends `Cut`, `Hotmelt`, `Glazed`, or any of the
  By/At columns.
- **Tablet → counters + `Station log`.** The station page (`glass.html` +
  `station.js`) reads `Glass station` and `Station people`, asks who is at
  the tablet (a name and, optionally, a PIN — a deterrent, not a secret;
  see `docs/STATIONS.md`), lets them tap only the stages they hold, and
  queues a PATCH of that row's changed counter plus its By/At and the
  overall DoneBy/DoneAt (`cw_stationq`, whitelisted to
  `ST.FLOOR_FIELDS`). Once a counter PATCH lands, one `Station log` line is
  queued and POSTed (`cw_stationlogq`) — never before, and never if the
  counter write never happened.
- **Office reads by delta.** The master dashboard never writes `Glass
  station`'s floor columns, and never writes `Station people` or `Station
  log` at all. It reads all three (the people list once, for the log
  window's filters) and keeps the board, the drawer's timeline and the log
  window current with `listDelta()` — every 10 seconds while the station
  board, the log window, or a drawer for a job with glass is open
  (`stationTick()`/`stationWatching()`), once a minute otherwise.

See `docs/STATIONS.md` for the full data model, admin setup and
troubleshooting, and `docs/specs/2026-09-08-glass-station.md` plus its v2
for the binding spec (the code wins over either where they disagree).

## Module map

| module | responsibility |
|---|---|
| `graph.js` | MSAL sign-in, token acquisition (with the quiet/popup split above), all Graph HTTP calls (`call()`), workbook session handling and retry-on-`InvalidSession`, `findFile()`/download/version history, the dashboard-owned-sheet upserts, `moveJobRow()` and its supporting row-capture/write/border functions, the SharePoint list layer (`listId`/`listItems`/`listUpsert`/`listDelete`/`listAdd`/`listPatch`/`listDelta`), batching (`batchGet`/`batchWrite`/`batchRun`) |
| `parser.js` | `parseWorkbook()`: turns a downloaded workbook into the job model; section/divider detection (`blocksFromValues`); fill-colour reading and the Cut/process/done ranking; `mapSheet()` header detection; `templateForJob()` (captures formatting for a move) |
| `checkpoints.js` | pure logic only, no DOM: the Excel-colour-vs-stored-count merge rule (`itemState`), colour choice, the per-(job,item) tap debounce and its `localStorage`-backed queue, the two write sequences (`cpWriteItem`/`cpWriteGroup`), and the phase pipeline (`jobPhase`, `effectivePhase`, `PHASES`) |
| `export.js` | pure logic only, no DOM, no network: filtering (`exportFilter`), row shaping, filename/summary text, building an ExcelJS workbook and a pdfmake document definition; the phone/eircode exclusion lives here as a hard rule with a standing test |
| `app.js` | everything with a DOM: rendering the job list/drawer/windows, wiring every tap to a write, `PENDING`/`PENDV`/`PENDA` optimistic holds, sign-in/session boot, the phase-list read/write UI, the Export/Alerts/Versions windows, the radial selection menu, build-freshness polling |
| `station-core.js` | pure glass-station logic (shipped 2026-09-08) — `glassSlice`, `feedPlan`, `sliceHash`, `jobBoard`, `applyTap`, who may record what (`stationPeople`/`canStage`/`pinOk`/`personExpired`), the write and log shapes (`floorOnly`/`tapFields`/`logFields`), reading the log back (`logRows`/`logFilter`/`logCounts`/`logLast`), and the two functions that keep a ten-second poll cheap (`mergeDelta`, `boardDiff`) — no DOM, no Graph, loads in Node and the browser like `checkpoints.js` |
| `station.js` | the glass station page UI (shipped 2026-09-08), using `CW` from `graph.js` and the pure functions from `station-core.js` |
| `build.py` | stamps a timestamp build id onto every script tag's `?v=` and into the footer, writes `version.json` for the freshness poll |

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
| a person's own unconfirmed edit (checkpoint count, section move, alert, phase) | in memory + `localStorage`, for up to 180 s, until the downloaded file (or list) agrees — `cw_pending`, `cw_pendv`, `cw_penda`, and the `phase` field inside `cw_pending` |
| exact checkpoint counts | `Dashboard Progress` sheet (Excel keeps only the cell colour) |
| the audit trail | `Dashboard Log` sheet — every write anywhere in the app appends one line here |
| custom groupings / saved views | `Dashboard Views` sheet |
| email subscriptions | `Dashboard Alerts` sheet |
| admin address and other read-only config | `Dashboard Config` sheet (read by the dashboard, never written) |
| hand-set phase per job | `Dashboard phases` SharePoint list (not the workbook) |
| glass station job facts and floor counters | `Glass station` SharePoint list, in the separate `Floor stations` site — shipped 2026-09-08 |
| who may record which stage on the floor | `Station people` SharePoint list, same site — shipped 2026-09-08 |
| who moved which counter, and when | `Station log` SharePoint list, same site; written by the tablet, read by the master, never deleted from — shipped 2026-09-08 |
| a queued/not-yet-sent checkpoint tap | `cw_cpqueue`, replayed on the next page load |
| a queued/not-yet-sent station tap | `cw_stationq` (counters) and `cw_stationlogq` (the log lines they owe) — tablet only, shipped 2026-09-08 |
| who is at the station tablet, and when they last tapped | `cw_person` — tablet only; the stages always come back from the list, never from storage |
| when the master last fed the `Glass station` list, and the hash of what it sent | `cw_stationfeed` — read by `feedStation()` to decide whether a run can be skipped |
| workbook and list ids, once resolved | `cw_fileref`, `cw_listids`, `cw_stationsite` |
| export presets (filters/fields/format only, never job data) | `cw_exportpresets` |
| UI-only preferences | `cw_theme`, `cw_hidden`, `cw_collapsed`, `cw_changes`, `cw_stationtheme` (the tablet's own light/dark, dark by default) |
