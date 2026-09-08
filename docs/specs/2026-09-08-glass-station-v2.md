# Glass station v2: people, stages, log, real time (brief, 2026-09-08)

Builds on `2026-09-08-glass-station.md` (read it first, including its
amendments). Everything there still holds unless changed here. The hard rules
are unchanged: the workbook is never written; the tablet can only change its
counters; the feeder never touches the floor's columns and never deletes; no
phone/eircode/county/product data on the floor; no real person's name, address
or domain in the repo (tests use example.test and made-up names); existing
suites green with `set -o pipefail`; no commit, no push, no live systems.

## What the owner decided (2026-09-08)

1. One shared Glass account; three people use it. Each person picks their
   **name** on the tablet; the name decides which stages they may edit. The
   other stages stay visible, greyed, read-only.
2. The stages are **Glass cut, Hotmelt, Glazing** (Toughening is gone).
   Hotmelt applies to **every** glass type.
3. A person may hold **one or more stages** (e.g. one person: hotmelt and cut).
4. **PIN per person, on.** Picking a name asks for the PIN.
5. One tablet passed around now, maybe three later: the chosen name **locks
   itself after 10 minutes without a tap** and there is a Switch person button.
6. The owner likes the stacked layout of the phone-width dark screenshot
   (label above a full-width bar): use that layout for the bars at every width,
   and make the tablet default to the dark theme with a light/dark toggle in the
   header (saved on the device).
7. From the master dashboard, clicking a job shows the whole progress of that
   order and **who changed what, when**. There is a **log per section and per
   person**.
8. Updates must reach the office in **real time** (a floor tap visible on the
   colleague's screen within ~10 s) and the tablet must be fast.

## Data model (all lists in the `Floor stations` site; none created by code)

### `Glass station` (changed)

Columns: Title (`JOB|TYPE`, unique), Job, Customer, GlassType, Total, Seq,
Active, FedAt, FedBy (feeder) — and, floor-owned: **Cut, Hotmelt, Glazed**
(numbers), **CutBy, CutAt, HotmeltBy, HotmeltAt, GlazedBy, GlazedAt** (text),
DoneBy, DoneAt (last touch). `Toughened` no longer exists. The feeder's field
set is unchanged; the floor's whitelist becomes exactly the twelve floor
columns above, and a PATCH for a stage carries only that stage's counter,
its By/At, and DoneBy/DoneAt.

`STAGES = [["cut","Glass cut"],["hotmelt","Hotmelt"],["glazed","Glazing"]]`,
`STAGE_FIELD = {cut:"Cut", hotmelt:"Hotmelt", glazed:"Glazed"}`, plus
`STAGE_BY`/`STAGE_AT` maps. `TOUGH_TYPES` and the `tough` flag are removed;
every stage applies to every type.

### `Station people` (new)

| column | type | meaning |
|---|---|---|
| Title | text | the person's display name as shown on the tablet |
| Station | text | `Glass` |
| Stages | text | comma-separated stage keys, e.g. `hotmelt,cut` |
| PIN | text | 4–6 digits; empty = no PIN asked for this person |
| Active | text | `Yes`/`No`; only `Yes` rows are offered |

Maintained by the owner in SharePoint. Read by the tablet at start and every
10 minutes, and by the master for the log filters. The PIN is compared on the
tablet; document in STATIONS.md that this is a deterrent on a shared device,
not a secret (the station account can read the list).

### `Station log` (new)

| column | type | meaning |
|---|---|---|
| Title | text | job number |
| Station | text | `Glass` |
| GlassType | text | e.g. `TG` |
| Stage | text | `cut` / `hotmelt` / `glazed` |
| From | number | counter before |
| To | number | counter after |
| Who | text | the person's name |
| At | text | ISO timestamp of the tap that produced this value |

One line per **sent write**: the tablet's queue already merges quick
successive taps on one row/stage, so `From` is the value last known to be in
the list and `To` is the value sent. The log is written by the tablet only,
after the counter PATCH succeeds (POST via `listAdd`); a failed log write is
retried with the queue and never blocks the counter. Nothing ever deletes
from it. The master reads it, never writes it.

## Tablet (`glass.html`, `station.js`, `station-core.js`)

- **Who are you?** After sign-in (and whenever locked), a full-screen picker:
  one big button per active person from `Station people` (min 64 px tall). If
  the person has a PIN, a numeric pad appears (buttons ≥ 56 px); wrong PIN =
  shake + "Try again", no lockout. The choice is kept in memory and in
  localStorage `cw_person` with the time of the last tap.
- **Lock**: 10 minutes without a tap (`PERSON_LOCK_MS`, one constant) returns
  to the picker; so does the **Switch person** button in the header (replaces
  the old name prompt). The header shows "Person A · Glass cut" (name and the
  stages they hold).
- **Stage gating**: steppers for stages the person does not hold are rendered
  disabled (`disabled` + `aria-disabled`, greyed, values still shown). Taps on
  them do nothing, and `tap()` itself refuses a stage the person does not hold
  (belt and braces; tested). The `All`/`None` button follows the same rule.
- **Writes**: a tap queues `{<Stage>: n, <Stage>By: name, <Stage>At: iso,
  DoneBy, DoneAt}` for the row; after the PATCH succeeds, a `Station log` line
  is queued and sent. Queue entries carry the person's name captured at tap
  time (as today for DoneBy).
- **Real time**: poll every 10 s (`REFRESH_MS`) using a **delta query**
  (`listDelta`, below); merge changed rows into `ITEMS` (removed → drop).
  On a delta failure fall back to a full read once and start a new delta.
- **Fast drawing**: keep cards keyed by job; on each refresh compute which
  cards changed (a pure `boardDiff(prev, next)` in station-core.js, tested)
  and update only those nodes; never `innerHTML` the whole board after the
  first paint; an open card stays open; a card that becomes finished moves
  to the Finished group (collapsed by default, with a count, tap to expand).
- **Layout**: the stacked bar style at all widths (label row with the count
  on the right, full-width track under it); two-column card grid from 700 px.
  Dark by default, toggle in the header (`cw_stationtheme`).
- Everything else from v1 stays: no delete, no job-fact edits, no link to the
  master, checkBuild reload only with an empty queue, error states.

## Master (`app.js`, `index.html`)

- **Drawer, Glass station section**: for each stage a bar plus "who · when"
  from the By/At columns ("Cut 8 of 8 · Person A · Tue 14:02"); under it a
  **timeline** of the job's `Station log` lines, newest first, max 12, with a
  "Full log" link that opens the log window filtered to that job. Read-only.
- **Log window** (`#logbtn` next to the board, and from the drawer): every
  `Station log` line, newest first, filters: person, stage, job, day; a count
  per person and per stage for the filtered range at the top. No export, no
  delete. Paged at 200 rows on screen with "show more".
- **Real time**: while the station board, the log window, or a drawer for a
  job with glass is open, poll `Glass station` and `Station log` every 10 s
  by delta and redraw only what changed; otherwise every 60 s. Never toast
  on a poll; a failure keeps the last data and shows the existing
  "cannot reach SharePoint" line.
- `Show ▸ Glass station` board: bars use the same three stages; each card
  gets a one-line "last: Person A cut TG 5→8, 2 min ago" from the log.

## `graph.js`

- `listDelta(name, opts)` → `{items:[{id, fields, removed:boolean}], next}`
  where `opts = {siteId, fields, token}`; first call without `token` does the
  initial `.../items/delta?expand=fields(select=…)` and follows `@odata.nextLink`
  pages, returning the final `@odata.deltaLink` as `next`; later calls pass
  `token` (the deltaLink, run through `graphPath`). A 4xx from the delta
  endpoint throws a recognisable error so callers fall back to `listItems`.
  Uses `LIST_SCOPES` quietly like every list call. Tests cover pages, removed
  items, and the fallback.
- Nothing else changes in graph.js.

## Setup script (manager's job, not yours)

The manager creates the three lists with a scratch script. You do not create
lists or columns. If a list or column is missing, the pages say which one.

## Tests (extend `test_station.js`; target ≥ 100 checks total)

People and gating (hold one stage, two stages, none; PIN right/wrong/empty;
lock after 10 min; switch person clears), log line composition (From/To/Who/At,
merged taps produce one line, log queued only after the counter PATCH
succeeded, retried on failure, never blocks the counter), delta merge
(add/update/removed, page following, fallback), `boardDiff` (changed cards
only; open state kept), master timeline formatting and the log window filters,
the stage rename (no `Toughened` anywhere in the new code), and the full
"nothing touched the workbook, nothing personal sent" sweep. Keep every
existing suite green.

## Report back

Files changed; full pasted output of all suites; one paragraph per numbered
owner decision above saying where it is implemented and which test covers it;
anything you were unsure of.

## Amendments after review (2026-09-08, manager)

From the independent review and the manager's browser pass. All in one fix pass.

1. **Delta result vs a feed reset (blocker).** In the master's `stationDelta`, capture whether a token was in use *before* the await (`const had = opts.token || null`), keep a generation counter bumped by `readStation`/`feedStation`, and discard a delta result whose generation is stale. `feedStation` also waits for / skips while `stationPolling`.
2. **Drawer timeline matches the job exactly**, never by substring; the log window's free-text job box keeps substring search (`logFilter` gains an exact mode).
3. **Log window re-rendering**: a poll or a keystroke redraws only the rows and the counts, never the filter bar; focus and caret survive. The filter-clear button is labelled "Clear filters".
4. **Drawer redraw from a poll** preserves the comment box text (`#cbox`) like the other inputs, or is skipped while focus is inside the drawer.
5. **Log `From` is captured at tap time** in `queueTap` when the entry is created and kept across merged taps; never re-derived at flush time, so a lost response cannot swallow the log line.
6. **People read retry**: while `!PEOPLE_READ` the fast tick retries `readPeople()`, and the "Try again" button shows whenever the board cannot be shown.
7. **`mergeDelta` defensive**: a non-removed change with an empty `fields` bag keeps the existing row's fields; `listDelta` re-appends `expand=fields(select=…)` to a token path that lacks it.
8. A list that refuses delta sets a `DELTA_OFF` flag (retried every 5 minutes) instead of one failing delta per poll.
9. Expanding a card and a tap at the clamp both count as activity for the 10-minute lock (`touch()`).
10. The office reads `Station people` without the `PIN` column.
11. `loadPerson` clamps a stored `at` to now; STATIONS.md states that the tablet's rules are a deterrent, not a secret.
12. No human-looking full names in tests: use "Customer One" style placeholders.
13. The initial `Station log` read is limited to the last 90 days (by `At`), with a console warning if the page cap is hit; `logRows` is memoised between renders.
14. Clicking a job number in the log window only opens the drawer if the job is on the sheet; otherwise the window stays.
15. `stationTick()` is re-armed from the Show dropdown, `openDrawer`, and `openStationLog`, so the first fast poll comes within 10 s of looking at the floor.

## Interim site fallback (2026-09-08) and its review amendments

The owner cannot create the `Floor stations` site yet, so `stationSite()`
falls back to the workbook's own site (lists created there by the manager's
script) and re-checks for the real site every 10 minutes, switching over on
its own. Review amendments, all in one fix pass:

1. **Every forget bumps the move counter.** `forgetStationSite()` (and any
   path that nulls the cached site) increments `stationSiteMoves()`
   unconditionally, so a later switch-over always drops delta tokens on both
   pages. Test: fallback with a token in hand → forget → site appears → the
   counter moved and the next list request names the new site.
2. **Queued taps are site-stamped.** Each `cw_stationq` entry records the site
   id it was captured in; after a move, entries from another site are
   re-matched by `JOB|TYPE` Title against the freshly read items (the Title is
   unique) and dropped if not found. Never PATCH an item id from another site.
   Test: queue an entry, move site, assert the PATCH targets the re-matched id
   and the log line names the right job.
3. **The tablet never writes without a resolved site.** `flushQueue`/`flushLog`
   re-arm and return when `SITEID` is null; `listSiteId(opts)` throws rather
   than reaching `findFile()` when `opts` carries an explicit null `siteId`.
   Test: a tap during the "site forgotten" window makes no `/drive/` request.
4. **Forget keeps the cadence.** Only a site-shaped failure (404/403 on a site
   lookup) forgets the site; a failed list call keeps `stationSiteMissAt` and
   `stationSoftMiss`, so the 10-minute/1-minute re-check holds survive, and a
   single 403 on the re-check does not flip the pages to the fallback.
5. **A move clears the feeder hash and the office's people cache**, so the new
   site is fed at the next load and the log filter shows the new roster.
6. **No site name in any message**: "The 'Glass station' list is not in the
   floor's site yet" style wording on both pages; test asserts neither page's
   error strings contain "Floor stations", "FloorStations" or the workbook
   site's name.
7. `needsListScope`: `/lists` is tested before `/drive/`.
