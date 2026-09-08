# Glass station: the first floor dashboard (brief, 2026-09-08)

## Product context

The Costello production dashboard (`index.html` + `app.js`, GitHub Pages, MSAL +
Microsoft Graph) is the **master dashboard**. Two office users (the admin and one
colleague) use it. It reads the shared SharePoint workbook (Production sheet) by
downloading it, and writes to it only in sanctioned ways (fills, whole-row section
moves, dashboard-owned sheets). Hand-set phases already live in a SharePoint list
called `Dashboard phases` in the same site as the workbook (see `graph.js`, block
"SharePoint lists", and `docs/specs/2026-09-07-phases-list.md`).

This feature adds the first **floor station**: a separate page for the glass area,
used on a shared tablet signed in with a shared station account. The floor sees
only: job number, customer name, the glass units of that job by type, and three
progress bars per job (**Glass cutting**, **Toughening**, **Glazing**). The floor
records progress by tapping. The office sees the floor's progress inside the
master dashboard, from the "Sheet" dropdown (renamed as described below) and in
the job drawer.

The station account has **no access to the workbook at all**. It is a member only
of a small separate SharePoint site, `Floor stations`, which holds one list per
station. The master dashboard is the **feeder**: every time it loads the workbook
it pushes the glass slice into the `Glass station` list. Nothing else feeds it.

## Hard rules (provable; the reviewer greps for them)

1. **The workbook is never written by this feature.** No fills, no values, no
   rows, no new sheets, not even a Dashboard Log line. Grep the diff: no
   `setFill`, `clearFill`, `setValues`, `appendLog`, `saveProgress`,
   `moveJobRow`, `batchWrite` or any `/workbook` path may be introduced by this
   feature. The station page must not load `parser.js`' workbook download path
   or call `downloadWorkbook` at all.
2. **The station page has no option to delete or change anything** except the
   three counters (Cut, Toughened, Glazed) of a glass row. No delete, no edit of
   job facts, no comments, no links to the master page, no export.
3. **The feeder never touches the floor's counters.** It writes only job facts
   (Job, Customer, GlassType, Total, Seq, Active, FedAt, FedBy). It never sends
   `Cut`, `Toughened`, `Glazed`, `DoneBy`, `DoneAt`. It never deletes a list
   item: a job that leaves production is marked `Active = No`.
4. **No phone numbers, eircodes, counties, prices, comments, product or product
   counts** reach the station list or page. Only job number, customer name,
   glass type and count.
5. **No real email address, person's name or domain in the repo** beyond what
   already exists (`SITE_PATH` hostname is already there). Tests use
   `example.test`. Station account naming appears only in docs as placeholders.
6. **Existing checks stay green**: `verify.js`, `test_move.js`,
   `test_checkpoints.js`, `test_alerts.js`, `test_export.js`, `test_phases.js`,
   `test_phases_list.js`, `automation/test_digest.js`. Run with
   `set -o pipefail`.
7. **Do not commit, do not push, do not touch live SharePoint or the live
   workbook.** All verification is offline (fake `fetch`) or in a browser with
   stubbed data.

## Data model

SharePoint site: `Floor stations`, path `/sites/FloorStations` on the same
hostname as `SITE_PATH` in `graph.js`. Resolve its id with
`GET /sites/{hostname}:/sites/FloorStations` (hostname = the part of `SITE_PATH`
before the colon). Cache it in localStorage (`cw_stationsite`). The site may not
exist yet: the pages must say so plainly, never throw uncaught.

List: `Glass station` (display name; template genericList). **One item per job
and glass type.** Columns (all created by the manager by hand or via Graph; the
code never creates lists or columns):

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, unique | feeder | `JOB|TYPE`, upper-case, e.g. `R5303|TG` |
| Job | text | feeder | job number, upper-case |
| Customer | text | feeder | customer name from the Production sheet (`j.cust`), max 70 chars |
| GlassType | text | feeder | the parser's glass key, upper-case: DG, TG, TUFF, NOT TUFF, ARCH, ASTRAGAL, FANCY, EXTRA |
| Total | number | feeder | units of that type for the job (`j.glass[type]`) |
| Seq | number | feeder | the job's position in the master list (sheet order), so the floor sees the office's order |
| Active | text | feeder | `Yes` while the job is in production and has that glass type; `No` afterwards |
| FedAt | text | feeder | ISO timestamp of the last feed that changed this item |
| FedBy | text | feeder | who was signed in to the master dashboard |
| Cut | number | station | units cut |
| Toughened | number | station | units toughened |
| Glazed | number | station | units glazed |
| DoneBy | text | station | "Name (Glass station)" from the shift name tap, or the account name |
| DoneAt | text | station | ISO timestamp of the last counter change |

Toughening does not apply to every type. `TOUGH_TYPES = ["TUFF", "TG", "DG"]`
(one constant in `station-core.js`, documented, easy to change after the demo).
For other types the Toughened counter is hidden and excluded from the job's
Toughening bar total.

"In production" for `Active`: the job is in `live()` (i.e. `j.cat !== "past"`)
and its section name (`BLOCKNAMES[j.blk]`) is the "In production" section (match
case-insensitively on the prefix `in production`). Jobs in Ready to fit, Collect
& supply only, or any other section are `Active = No`. Reuse the predicate the
alerts digest uses if it is exported; otherwise write one small pure function
and test it.

## Files to create / change

New files:

- `station-core.js` — pure logic, no DOM, no Graph. Loads in the browser and in
  Node like `checkpoints.js` (assign to `window`/`globalThis`). Exports:
  - `glassSlice(jobs, blockNames)` → array of `{title, job, customer, type,
    total, seq, active}` for every job that has any glass type with count > 0
    (active or not; a job with no glass at all produces nothing).
  - `feedPlan(slice, items)` → `{adds:[fields], patches:[{id, fields}],
    unchanged:n}`. `items` are current list items (`{id, fields}`). A slice row
    with no item → add. An item whose feeder fields differ → patch with only
    the changed feeder fields plus `FedAt`/`FedBy`. An item not in the slice
    and `Active !== "No"` → patch `{Active:"No"}`. Never includes Cut,
    Toughened, Glazed, DoneBy, DoneAt. Deterministic order (Seq, then Title).
  - `sliceHash(slice)` → string; the feeder skips a run when it equals the last
    fed hash and the last feed is under 10 minutes old.
  - `jobBoard(items)` → grouping for display: `[{job, customer, seq, rows:[{id,
    type, total, cut, toughened, glazed, tough:boolean}], bars:{cut:{done,total},
    tough:{done,total}, glazed:{done,total}}, finished:boolean}]` for
    `Active === "Yes"` only, sorted by Seq then job. Counters clamp to
    `[0, total]` for display; `finished` = every bar full.
  - `applyTap(row, stage, delta|"all"|"none")` → the new counter value with
    clamps (0..total). No ordering rule between stages is enforced (the floor may
    glaze before recording toughening); only the clamp.
  - Constants: `STATION_LIST = "Glass station"`, `STATION_SITE = "FloorStations"`,
    `TOUGH_TYPES`, `STAGES = [["cut","Glass cutting"],["tough","Toughening"],["glazed","Glazing"]]`.
- `station.js` — the station page UI (tablet). Uses `CW` from `graph.js` and
  `station-core.js`. See "Station page" below.
- `glass.html` — the station page shell + styles (same fonts and tokens as
  `index.html`, both themes, phone width). Loads `vendor/msal-browser.min.js`,
  `graph.js`, `station-core.js`, `station.js`. **Does not load `exceljs`,
  `parser.js`, `checkpoints.js`, `export.js` or `app.js`.**
- `test_station.js` — offline checks (target ≥ 30): slice/plan/hash/board/tap
  logic incl. failure paths (job leaves production → Active No patch; total
  shrinks below Cut → board clamps, plan never writes Cut; duplicate items for
  one Title → board uses the oldest; item missing fields → no crash), the graph
  list generalisation (site + fields selection, paging), the station write queue
  (optimistic value shown, PATCH sent once, retry on 503, replay after reload),
  and the feeder guard (no run without silent list token; no run for a
  non-master; throttle). Use the same `vm.runInThisContext` + fake `fetch`
  pattern as `test_phases_list.js`.

Changed files:

- `graph.js`: generalise the list layer without changing the behaviour of the
  phases list. Add an optional `opts` argument: `listId(name, opts)`,
  `listItems(name, opts)`, `listItemsFor(name, title, opts)`,
  `listUpsert(name, title, fields, opts)` where `opts = {siteId, fields:[...]}`.
  Default (no opts) = the workbook's site and the phases field selection, exactly
  as today. Add `listAdd(name, fields, opts)`, `listPatch(name, itemId, fields,
  opts)` (plain POST/PATCH, no dedupe read, for the feeder and the station) and
  `stationSite()` (resolves and caches the `Floor stations` site id; returns
  null if the site does not exist, cached negative for 60 s only). Add
  `signIn(scopes)` so the station page can sign in asking for `LIST_SCOPES`
  directly (the master keeps `SCOPES`). Add `hasListConsent()` = a quiet token
  attempt for `LIST_SCOPES` that answers true/false without a popup. Export the
  new functions on `window.CW`, plus `_setStationSite` for tests.
- `app.js`:
  - **Sheet dropdown**: the options become `All jobs (n)` and `Glass station`
    (future stations are added by extending one array `STATIONS` next to
    `SHEETNAMES`). The other workbook sheets are removed from this dropdown.
    `SHEETNAMES` stays for the export filter and the row chips (unchanged).
    Selecting `Glass station` renders the **station board** in place of the job
    list: read-only, every active job with its three bars and per-type counts,
    `FedAt` freshness ("fed 3 min ago"), and a clear state when the site or list
    is missing or the permission is not granted (reuse the wording pattern of
    `PHASE_LIST_MISSING` / `PHASE_NEED_CONSENT`). Reading the board needs the
    list permission: use the quiet path and show the consent hint for the admin
    exactly as phases do.
  - **Drawer**: a "Glass station" section under the existing "Glass units"
    section with the three bars for that job (read-only), or "not fed yet".
  - **Feeder**: after a successful `load()` (after `readPhases()`), call
    `feedStation()`: skip unless `hasListConsent()`; skip if `sliceHash` matches
    the last run and it ran under 10 minutes ago; resolve `stationSite()` (skip
    silently if null); read the list once (`listItems` with the station fields);
    `feedPlan`; send adds and patches with concurrency 3 and at most 60 writes
    per run (the rest goes next run); remember the hash + time in memory and
    localStorage (`cw_stationfeed`). Failures are logged to the console and set
    a small status word in the footer ("station feed: 2 min ago" / "station
    feed failed"); they never toast and never block the dashboard. Only one
    feed at a time (busy flag). Not admin-only: any master user feeds.
  - **Wrong account**: if `downloadWorkbook()`/`findFile()` fails with 403 or
    404 for the signed-in account, the sign-in overlay shows "This account has
    no access to the production workbook. Station accounts use the Glass
    station page." with a link to `glass.html`. No retry loop.
- `index.html`: add `<script src="station-core.js">` before `app.js`; styles
  for the station board (bars reuse the checkpoint bar styles where sensible).
- `build.py`: stamp `station-core.js` and `station.js` too (both html files).
- `README.md` etc.: see "Docs" — written by a separate docs agent; leave a
  short section in README pointing to `docs/STATIONS.md` if you touch README.

## Station page (`glass.html`, `station.js`)

- Sign-in overlay like the master, button "Sign in", requests `LIST_SCOPES`.
  After sign-in: resolve the station site and list; if missing, show the plain
  message and a "Try again" button; if permission is missing, show "Ask the
  office to grant the SharePoint permission" (no admin consent flow on the
  tablet).
- Header: "Glass station", the shift name ("Working: Person A" — tap to change;
  stored in localStorage `cw_shiftname`; if empty, shows "Tap to add your
  name"), a refresh indicator ("updated 12 s ago"), Sign out. Nothing else.
- Board: cards sorted by Seq. Card = job number (large, condensed font like the
  master's job number), customer name, chips per glass type with `cut/total`
  style counts, and three bars (Glass cutting, Toughening, Glazing) with
  `done/total`. Finished jobs (all bars full) drop to a "Finished" group at the
  bottom, greyed; they leave the board when the feeder marks them inactive.
- Tapping a card expands it: one row per glass type with three steppers
  (− count + and an "All" button) for Cut, Toughened (only if `tough`), Glazed.
  Buttons ≥ 48 px. Every tap updates the screen at once, queues a PATCH of that
  row's changed counter + DoneBy + DoneAt, and retries on failure (queue in
  localStorage `cw_stationq`, replayed on reload; the same row's later tap
  supersedes the earlier one in the queue). A failed write shows a small red
  "not saved yet — retrying" on the row, never a modal.
- Refresh: re-read the list every 30 s and after each successful write; the
  local queued value wins over the list until the list agrees.
- Theme follows the OS (same tokens as `index.html`). Layout works at 768 px
  (tablet portrait) and 390 px.
- No links to `index.html`; no export; no delete; no editing of job facts.

## Tests and the report

Deliver `test_station.js` and keep every existing suite green. Report back with:
the list of files changed, the pasted output of every suite, a paragraph per hard
rule saying how the diff respects it, screenshots are NOT required (the manager
takes them), and anything you were unsure about.

## What to reuse

`serialised()`, `call()`, `graphPath()`, `listItems` paging, the dedupe rule of
`listUpsert` (not needed for the feeder: Title is unique and the feeder is the
only writer of job facts, so a refused POST means "exists, patch it"), `toast`,
`friendly`, `mk`, `esc`, `$`, the checkpoint bar CSS (`.cpbar` or whatever
`cpGroupHtml` uses in `index.html`), the phase list state words.

## Amendments after review (2026-09-08, manager)

Found by the browser pass and the independent review; all part of the single fix pass.

1. **Feed only jobs that are in production.** `glassSlice` returns rows for active jobs only. `feedPlan` still marks any existing item whose Title is not in the slice as `Active = No` (never deletes). Result: the list holds the floor's jobs plus a tail of finished ones, not every job in the workbook.
2. **`needsListScope()` must decide on the shape of the path, never on a colon**: any `/workbook` or `/drive/` path uses `SCOPES` (and keeps the interactive re-auth popup); `/lists` paths and a by-path site lookup other than `SITE_PATH` use `LIST_SCOPES`. Add a test asserting the scope chosen for a `range(address='A1:K600')` path and for the two site lookups.
3. **Tablet queue drains, never snapshots**: `flushQueue` re-reads the queue keys until empty; a tap arriving during an in-flight write is sent; the retry timer is armed whenever the queue is non-empty at the end (including the `!READY` exit); a successful `readList()` calls `flushQueue()`. Add a "tap during an in-flight write" test that drives `tap()` without touching `flushing` by hand. `DoneBy`/`DoneAt` are captured at tap time in `queueTap`, not at flush time. The queue read from localStorage is whitelisted to the three counter fields, on load and before each PATCH.
4. **Wrong-account gate**: `start()` calls `openSession()` before `load()`, and `findFile()` throws there, so the gate never shows and the page sits at "starting…". Route a no-access error from anywhere in `start()` to `showNoAccessGate()`. `NOACCESS_RE` matches only the status field of `call()`'s message (`-> 403` / `-> 404`) or `accessDenied`/`itemNotFound`, never a row number. `NOACCESS = false` at the top of `start()`.
5. **Drawer station section**: read the station list once when a drawer opens and `STATION_OK === null` (the hash skip must not leave it on "checking…"); look the job up over all items, active or not, so a finished job still shows the floor's record (with "finished on the floor" wording) rather than "Not fed to the floor yet".
6. **Errors on the tablet and the office board**: only a genuine 404/`itemNotFound` on the site or list clears the board; any other failure keeps the last board on screen with a small "cannot reach SharePoint — retrying" line and the underlying status in the console. MSAL `interaction_required`/`login_required` shows "Tap Sign out, then Sign in again" with the sign-in button, not the consent wording.
7. **`stationSite()` cache**: a failed list call against a cached site id clears the cache so the next call re-resolves; "Try again" on the tablet also clears it.
8. **Feeder**: `stationBusy = true` immediately after the guard; a refused POST re-reads that Title and patches it; a run cut short by the cap or a failure schedules a follow-up run 30 s later (once); `STATION_FEED_ERR` is cleared on any run that skips or succeeds; `FedBy` stores `account.name`, falling back to the local part of the address.
9. **Tablet build refresh**: `glass.html` polls `version.json` every 2 minutes like the master's `checkBuild()` and reloads itself when the build changes and the queue is empty.
10. Footer separator when `#stationfeed` is empty; the `stationSend` comment matches the code.
