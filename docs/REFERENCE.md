# Reference: what has been built, and how

One page per feature, in the order they were built. Each entry says what the
feature does, where the code is, how it works underneath, what it writes and
where, and which tests cover it. Read this before touching anything; read the
spec in `docs/specs/` for the full brief of a feature. Names of people and
addresses are placeholders throughout ("the admin", "the colleague").

Last updated 2026-09-10 (glass colours into the Production sheet, the tuff counter, the office's lock, and an office clear reaching the floor's counters).

---

## 0. The shape of the whole thing

- **One static site, two pages**, served by GitHub Pages from `main`:
  `index.html` + `app.js` (the master dashboard, office) and `glass.html` +
  `station.js` (the Glass station, floor tablet). No bundler, no framework.
- **One workbook** on SharePoint (site `ProductionProgress`) is the source of
  truth for jobs. Its `Production` sheet is read by downloading the file
  (`graph.js: downloadWorkbook`) and parsed by `parser.js`.
- **Three kinds of state**, in order of how sacred they are:
  1. The `Production` sheet: only two sanctioned writes (§3, §4, §5).
  2. Dashboard-owned sheets in the same workbook: `Dashboard Log`, `Dashboard
     Views`, `Dashboard Progress`, `Dashboard Alerts` (written), `Dashboard
     Config` (read only).
  3. SharePoint lists: `Dashboard phases` (workbook's site); `Glass station`,
     `Station people`, `Station log` (separate site `Floor stations`).
- **Sign-in**: MSAL browser, Entra public client, delegated Graph. Scopes are
  split so sign-in never depends on a permission the tenant may not have
  granted (§1).
- **Tests** are offline Node scripts that load the real files with a fake
  `fetch` (§13).

---

## 1. Sign-in and the Graph layer (`graph.js`)

**What.** Everything that talks to Microsoft: sign-in, workbook download,
surgical workbook writes, dashboard sheets, SharePoint lists, the floor
stations site.

**How.**
- `initAuth`/`signIn(scopes)`/`signOut`; `token(scopes, quiet)` gets a token
  silently and only opens a popup when `quiet` is false.
- `SCOPES = Files.ReadWrite.All + User.Read` for sign-in and every workbook
  call. `LIST_SCOPES` adds `Sites.ReadWrite.All` and is used **only** for
  `/lists` paths and a by-path site lookup other than the workbook's site
  (`needsListScope(path)`, decided on the *shape* of the path, never on
  characters inside a range address). Background list reads are quiet;
  `listConsent()` is the one place a consent popup may open, from a click.
- `call(method, path, body)` wraps `fetch` with retries for 429/5xx and the
  Excel `InvalidSession` 400 (re-opens the workbook session and retries once).
  The `workbook-session-id` header is sent only on `/workbook` paths.
- `findFile()` resolves site → drive → the workbook by name and caches it in
  `cw_fileref`. `downloadWorkbook()` fetches `/content` into ExcelJS. The
  downloaded file lags the Excel API by about 36 s, which is why the UI holds
  its own writes locally (§2).
- Workbook writes: `setFill/clearFill/setValues` (single ranges), `batchWrite`
  (`$batch` with a concurrency cap; flooding it earns `OperationQueueFull`),
  `moveJobRow` (§3), `appendLog`, `saveAssignment`, `saveProgress`,
  `addAlert/removeAlert`, each behind `serialised(sheet, fn)` so one tab's
  writes to a sheet run in order.
- Lists: `listId/listItems/listItemsFor/listUpsert/listDelete` (phases, with
  dedupe on Title), `listAdd/listPatch` (plain writes), `listDelta`
  (§11, delta queries with `@odata.deltaLink`, 410 resync detection),
  `stationSite()` (resolves and caches the `Floor stations` site id;
  `forgetStationSite()` clears it).

**Lessons.** Adding a scope to the sign-in list breaks sign-in for everyone
until admin consent exists. A colon test for "is this a list path" matched
every A1 range and killed every workbook write. Both are covered by tests.

---

## 2. Parsing and the job model (`parser.js`, `app.js: load`)

**What.** Turns the `Production` sheet into jobs the UI can use.

**How.**
- `mapSheet` finds columns by header text (identity columns, dates, quantities
  M/N, glass unit columns AY..BF, product F/S/T groups). `blocksFromValues`
  finds the five sections by their divider rows (a job row is never a divider).
  `parseWorkbook` yields `jobs[]` with `id`, `cust`, `blk` (section index),
  `cat`, `glass{type:count}`, `prods[]`, `cp` (checkpoint colours read from the
  cells: white / yellow "process" / gold "done" / the legend's green "cut").
- `load()` in `app.js` downloads, parses, reads the dashboard sheets and the
  phases list, then lays the local holds over the result: `PENDING` (row
  moves and fills, 180 s per key), `PENDV` (view assignments), `PENDA`
  (alerts). A hold is released the moment the file agrees with it.
- `poll()` every 12 s compares `lastModifiedDateTime` and reloads on change.
- `verify.js` compares the JS parser with a Python reference extract of a
  real download (local files, not in the repo).

---

## 3. Sections, categories and row moves

**What.** Jobs sit in the sheet's five sections (In production, Ready to fit,
Collect & supply only, …). Moving a job between sections in the dashboard
moves its **row** in Excel so the floor's Excel and the office's dashboard
agree. Custom categories live only in `Dashboard Views`.

**How.** `moveJobsInSheet` → `graph.js: moveJobRow`: locate the job live
(A1:K600), capture the row (values + number formats in one read; fills/fonts
per run from the download), insert a row after the target section's last job,
write it (text starting with a digit/=/+/- gets an apostrophe so Excel keeps
it text), verify exactly two copies exist, delete the original. Borders:
Excel keeps one definition per shared edge, so the landing row writes bottom +
verticals but never its top, and the row above a landing/vacated slot gets its
bottom edge re-asserted (`restoreBottomEdge`). Row heights come from the
downloaded file (the API rounds to pixels). Verified with desktop Excel via
COM in a scratch script. Tests: `test_move.js`.

---

## 4. Mark ready

**What.** Marks a job ready to deliver: gold fill on its status cell and a
row move to "Ready to fit" (C/S-prefixed jobs go to "Collect & supply only").
Undo clears the fill and moves the row to the bottom of "In production".
Dates are never touched. `markReady` in `app.js`.

---

## 5. Checkpoints (`checkpoints.js`, drawer in `app.js`)

**What.** Per-job counts for windows, doors, each glass type and each
product's F/S/T, ticked from the drawer, shown as bars.

**How.** Excel gets **only the colour** in the item's own cell (white 0 /
yellow part / gold done). Exact counts live in `Dashboard Progress`
(Job|Item|Done|Total|Who|When) plus a `Dashboard Log` line. On refresh the
Excel colour wins over a stored count. Pure logic (`cpItems`, `itemState`,
write ordering progress→fill→log, bursts and a replayable queue) is in
`checkpoints.js`; the drawer UI is `cpSectionHtml`/`cpPatchSection`.
**Clearing a job's glass also clears the floor's counters** since 2026-09-10 —
see §18. Tests: `test_checkpoints.js`.

---

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

---

## 7. Selection wheel (radial menu)

**What.** Ticking jobs shows a floating button; it opens a full-circle radial
menu (Alert to, Export, Move to, Untick all). `renderFab`/`fabToggle` in
`app.js`; motion in `--fab-*` CSS variables in `index.html`, timeline in
`docs/motion-notes.md`. Nothing here writes anything.

---

## 8. Versions and rollback

**What.** The Versions window lists the workbook's SharePoint version history
and can restore one (`renderVersions`, `graph.js: listVersions /
downloadVersion / restoreVersion`). This is the rollback for any dashboard
write to the workbook.

---

## 9. Job alerts by email

**What.** The admin subscribes addresses to jobs; every third day at 08:00 an
email per address lists its jobs still in production, job number and
comments only.

**How.** Subscriptions live in `Dashboard Alerts` (Job|Email|Added by|When);
only the admin (client-side check against `Dashboard Config` key `admin`)
can add or remove; recipients must be on the admin's domain. The mailer is an
Office Script (`automation/alerts-digest.ts`, generated from
`automation/digest-core.js` by `build-script.js`) run by a Power Automate
flow the owner creates by hand from `automation/SETUP.md`, sending from the
admin's mailbox. No address or domain is in the repo; they come from
`Dashboard Config` at run time. Tests: `test_alerts.js`,
`automation/test_digest.js`.

---

## 10. Export to Excel / PDF (`export.js`)

**What.** Download the current jobs (or the ticked ones) as an Excel workbook
or a PDF, with filters and a field picker.

**How.** Pure builders on `window.XP`: filters → rows → ExcelJS workbook
(Jobs, Comments, Checkpoints, Export info sheets) or a pdfmake document (job
cards, or a compact 14-column table). pdfmake is lazy-loaded from `vendor/`.
Hard rules: **no eircodes ever, and no phone numbers in this, the Default
template** (the John print sheet of §16 is the single sanctioned exception, and
the two templates share no row builder); one `Dashboard Log` line per export
(who, format, job count, filters); no network. Since 2026-09-09 the field
picker also offers **Flag** — the Production sheet's row colour code (§16) —
which comes out as a last column carrying the word and the sheet's own ink.
Logo slot `assets/logo.png` (not yet supplied). Tests: `test_export.js`.

---

## 11. Glass station (floor dashboards)

**What.** A separate page for the glass area on a shared tablet, showing only
job number, customer name and **one total number of glasses** (DG + TG), with
three stages (**Cutting, Hotmelting, Glazing**) recorded per person. There are
no glass types on the floor at all. The office sees it all from the master
dashboard. The floor cannot see the master.

**Where.** `station-core.js` (pure logic, loads in browser and Node),
`station.js` (tablet UI), `glass.html` (shell + styles), plus the station
blocks in `app.js` (search `STATIONS`, `feedStation`, `stationPoll`,
`renderStationLog`, `stationSectionHtml`) and `listDelta`/`stationSite` in
`graph.js`. Specs: `docs/specs/2026-09-08-glass-station.md` (v1) and
`…-v2.md` (both with "Amendments after review" sections) and `…-v3.md`, which
made it one row per job, one number per job and cards you tap directly. Admin
guide: `docs/STATIONS.md`.

**How, in five parts.**
1. **Separation by permission, not by screen.** The station account is a
   member only of a private SharePoint site `Floor stations`; it has no
   access to the workbook. If it opens the master page, `start()` shows a
   plain "wrong page" notice with a link to the station page and stops
   polling.
2. **The office dashboard is the feeder.** After every `load()`,
   `feedStation()` computes the glass slice of jobs in the "In production"
   section (`glassSlice`, one row per job, `Total = DG + TG`), diffs it
   against the `Glass station` list (`feedPlan`), and sends adds/patches
   (3 lanes, 60 writes per run, a 30 s follow-up if cut short). It writes job
   facts (Job, Customer, GlassType = the literal `GLASS`, Total, Seq, Active,
   FedAt, FedBy) and — the one exception to "never the floor's columns", added
   in v3 — the three counters, **only on a row it is creating or a row whose
   `DoneAt` is empty**, seeded from the office's own glass checkpoints
   (`glassCounts` in app.js → `ST.officeSeed`). That guard is re-checked with a
   one-item `listItem()` read immediately before every write that carries a
   counter (`stationFeedPatch`), because the plan is made from one read and
   sent as up to sixty writes; the same read drops the seed when the row's
   `FedAt` is newer than this dashboard's `lastStamp`. It never writes a By, an
   At or the last-touch pair, never touches a counter after the floor's first tap,
   and never deletes (a job that leaves production is `Active = No`). Skipped
   quietly without list consent; throttled by a hash of the slice, which
   includes the seed.
3. **The tablet.** Sign in once with the station account (asks for the list
   scopes at the door). "Who are you?" picker from `Station people`
   (Title=name, Station, Stages comma list, PIN, Active); PIN pad when the
   row has one; a person may hold one or more stages; steppers for other
   stages are greyed and `tap()` refuses them too. One card per job, with the
   job number, the customer and the job's own glass count — `glassWords(total)`
   = DG + TG, plus `· N tuff` where there is tuff — which is **the same for
   everybody**, and the steppers (−, +, All / None) on the card itself. Since
   2026-09-10 each stepper row the person **holds** carries its own
   `ST.stageLeft(g, stage)` under the stage's name ("Cutting 20 left"): one
   number per stage, **never summed** (an earlier build summed them and a real
   job read 157 — 49 glasses × 3 stages + 10 tuff — which the owner rejected on
   sight), tuff counted against `TuffTotal` and never against DG + TG. The
   number sits in the label cell of the row whose buttons move it, so the card
   is no taller with four numbers than with one; a row is either yours (a
   number) or not (the words "not yours"). Two people holding the same stage
   read the **same** number, because `stageLeft` never sees a person — nothing
   enforces that and nothing can break it. Nothing to expand, no bars, nothing
   to scroll inside — plus a search box in the header (`boardFilter`, job
   number or customer) and, beside it, `#gtotal`: `ST.boardLefts`, the same
   per-stage numbers summed down the board in the same words and the same order
   ("Cutting 8 left · Tuff 11 left"), never narrowed by the search and hidden
   whenever the search box is — and absent entirely for somebody who holds no
   stages. All of it is derived on the device at every draw and stored nowhere.
   Ten quiet minutes lock the
   tablet back to the picker; Switch person does the same. Every tap shows at
   once, is queued in `cw_stationq`, PATCHes only that stage's counter plus
   `<Stage>By/<Stage>At/DoneBy/DoneAt`, then POSTs one `Station log` line
   (Title=job, Station, GlassType, Stage, From, To, Who, At) with `From`
   captured at tap time. The queue drains until empty, retries, and survives
   a reload; a replayed queue is whitelisted to the counter fields.
4. **Real time.** Both pages poll by delta (`listDelta`, 10 s on the tablet;
   10 s on the office while the board, a glass job's drawer or the log window
   is open, 60 s otherwise). `mergeDelta` applies last-occurrence-wins and
   `@removed`; a 410 resync or a refused delta falls back to one full read
   (and a 5-minute "delta off" flag). The tablet repaints only the cards that
   changed (`boardDiff`), re-bases a queued tap whose row has risen under it
   (`rebaseQueue`), and folds finished jobs — all three counters at the
   total, with a total > 0 — into a collapsed "Finished · n" group, gold.
5. **The office view.** Show ▸ Glass station replaces the job list with a
   read-only board: one card per job with "12 glasses" — the office's number
   is still the size of the job, where the tablet's is one person's remaining;
   they answer different questions and are not meant to agree — the three
   counters with who · when under each, "last: …" from the log, and the same gold rule
   (finished cards go gold and sink to the bottom). The drawer shows the same
   three lines, a 12-line timeline and Full log; the Floor log window filters
   by person, stage, job and day with counts, read-only, 90-day horizon, and
   no glass-type column. The office never writes `Station people` or
   `Station log`.

**Where the lists live (interim, 2026-09-08).** `stationSite()` prefers the
separate `Floor stations` site; while it does not exist (the owner has no
Global Administrator to create it) the same three lists live in the
workbook's own site and both pages fall back to it, re-checking for the real
site every ten minutes and switching over on their own (delta tokens and
list-id caches are dropped on the move). In that interim the station account
is a member of the workbook's site and can therefore open the workbook; the
owner accepted that. See `docs/STATIONS.md`, "Interim".

**Honest limits (told to the owner).** PIN, gating and the lock are a
deterrent on a shared device, not security; the real boundary is site
membership. Graph has no push; 10 s delta polling is the real-time mechanism.
Jobs reach the floor only while an office dashboard is open.

**On the ordinary job list (2026-09-09).** Every job row carries a `Glass`
column of its own (the ninth, 96px, between `Components F·S·T` and `On
sheets`): one badge reading the floor's progress across all three stages —
`Glass 12/24` for eight glasses cut 8, hotmelted 4, glazed 0 — gold on the same
"finished" rule the board and the tablet use, with the breakdown and "the Excel
file is not involved" in its `title`. A job the floor has never been fed
renders that cell empty. The rows repaint when the floor taps (`redrawStation`
now redraws the plain list, not only a floor-station board) and the poll runs
at the fast rate whenever the rows on screen carry the floor's work
(`ROWS_GLASS`). A repaint is held off only while somebody is typing in the list
or dragging a row, and is then *owed* (`ROWS_STALE` / `stationCatchUp`) rather
than dropped; it fires only when the chips would actually say something
different (`ROWS_CHIPS`), and without the row entry animation. `stationRecords()`
caches `ST.jobRecords(items)` on `STATION_ITEMS` array identity — O(rows +
items) instead of O(rows × items), measured 11–16 ms → 0.71 ms per render.
**Still no workbook write and no list write:** the one read added is the
existing shared `stationReadIfNeeded()`, once per session. Spec:
`docs/specs/2026-09-09-glass-chip-on-job-row.md`.

**Since 2026-09-10.** The floor's counters now reach the `Production` sheet's
four glass columns as fills, painted by the office dashboard; there is a
fourth counter (tuff) with its own total; and the office can make a job
read-only on the tablet. See §17 — and note that `STAGE_KEYS` still means the
three glass stages everywhere in this section, with the four in
`ALL_STAGE_KEYS`.

**Tests.** `test_station.js` (235 checks) proves, over the whole run, that no
request touched the workbook, no DELETE was sent, every floor PATCH is a
subset of the floor columns, every feeder write is inside `ST.FEEDER_WRITES`
(the job facts plus the three counters, never a By, an At or the last touch),
and no body carried a phone, eircode, county, price or comment.

---

## 12. Documentation set

| file | purpose |
|---|---|
| `CLAUDE.md` | rules, process, checks, deploy — read first |
| `README.md` | what it is, features, file map, how to run |
| `docs/REFERENCE.md` | this file: what was built and how |
| `docs/ARCHITECTURE.md` | data flow, module map, invariants, where state lives |
| `docs/SUPPORT.md` | troubleshooting |
| `docs/STATIONS.md` | floor stations: data model and admin setup |
| `docs/specs/README.md` | index of every feature brief with status |
| `docs/motion-notes.md` | the radial menu's timing |
| `automation/SETUP.md` | the alerts flow, step by step |

---

## 13. Tests, build, deploy

| suite | covers | checks |
|---|---|---|
| `verify.js` | parser vs a Python reference extract (local files) | agree/disagree |
| `test_move.js` | row-move protocol | 7 |
| `test_checkpoints.js` | checkpoint logic and writes | 36 |
| `test_alerts.js` | subscriptions, admin gate, pending holds | 30 |
| `test_export.js` | builders, no phone/eircode, logging | 53 |
| `test_john.js` | the John print sheet, template and print notes | 31 |
| `test_phases.js` | phase derivation and the wheel | 26 |
| `test_phases_list.js` | the phases list, scope split, dedupe | 35 |
| `test_station.js` | floor stations end to end (offline), incl. tuff, the lock, the seed, the office clear's vocabulary and a queued tap meeting its zeros | 243 |
| `test_glasscolour.js` | glass colours into `Production`: the rule, last-writer-wins (per job, outside the settling window), idempotence, the cap, the backoff, fills only, the office clear end to end, the observed un-tick repaint, and the office-absolute guard | 65 |
| `automation/test_digest.js` | the alerts digest | 26 |

Pattern: `vm.runInThisContext` loads the real source; `global.fetch` is a
fake that serves Graph-shaped JSON and records every request; assertions
include "this was never requested". Always run with `set -o pipefail`.

Build: `python build.py` stamps `?v=<build>` on every script tag of both
pages and writes `version.json` (both pages poll it and offer/perform a
reload). Deploy = push to `main`. Commit only after the owner has approved a
demo; commit messages end with the `Co-Authored-By: Claude …` trailer.

---

## 14. Scratch tooling (outside the repo, in the session scratchpad)

- `graph.py`, `token.json`, `fileref.json`: a device-code token and the
  workbook reference for one-off Graph scripts.
- `rehearse.js` (in the repo) + `rehearse_*.js`, `live_*.js`: run the real
  `graph.js` against a OneDrive copy of the workbook, then one net-zero check
  on the live file. `fidelity.py`, `cp_fidelity.py`, `com_check.py`: compare
  a rehearsed workbook against the original, including drawn borders via
  desktop Excel.
- `make_phase_list.py`, `make_station_lists.py`: create/complete the
  SharePoint lists with the exact columns (idempotent; `check` mode).
- `stub_station.js` + `cors_server.py`: a fake sign-in library and fake Graph
  injected with Playwright so both pages run offline against a fixture for
  screenshots and behaviour checks (delta, real-time, gating).

---

## 15. Lessons that cost time

1. Never add a scope to the sign-in list; ask for it only on the requests
   that need it, quietly in the background, with a popup only from a click.
2. Decide "which scope" on the shape of a path, never on characters in it.
3. Excel keeps one border definition per shared edge; write the landing
   row's bottom and verticals, never its top, and re-assert the row above.
4. The downloaded file lags the API by ~36 s; hold your own writes locally
   until the file agrees.
5. `values` PATCH auto-parses "07/04" and "9434": guard text with an
   apostrophe.
6. `$batch` floods earn `OperationQueueFull`; cap concurrency.
7. A `| tail -1` masked a failing suite once; always `set -o pipefail`.
8. A queue that snapshots its keys loses the tap that lands mid-write;
   drain until empty and re-arm the retry whenever anything is owed.
9. Capture "before" values at the moment of the action, never at flush time.
10. A delta result must be discarded if the list was replaced while it was in
    flight (generation counter).
11. Every feature gets an independent review by a different agent before the
    demo; both blockers in the station work were found that way, not by tests.

---

## 16. Row colour code, section select, the John print sheet and print notes

Shipped 2026-09-09. Spec: `docs/specs/2026-09-09-john-template.md`, including
its "Amendments after review and the owner's second round". Office page only —
the tablet is untouched by all of it. The owner named the view, the template
and the file **"John print sheet"**, and that name is his recorded exception to
the no-real-names rule.

**The row colour code (`parser.js`, `app.js`, `export.js`).** The Production
sheet says four things with the colour of a job row's *text*, and nowhere
else: red = urgent, green = booked (an exact date is agreed), pink/magenta =
trade order, blue = on hold. `fontOf()` reads the job-number cell's font
colour (theme colours resolve through the same palette and tint the fills use)
and falls back to the customer cell **only when the job number's own ink is
unset or plain black** (`blackInk`). A cell carrying a hyperlink, and theme 10
and 11 — Excel's own two hyperlink colours — are never read at all, because a
link is blue by default and would otherwise arrive as "on hold". `flagOf()`
maps the hex by **hue range, not exact value** — the office picks these by hand
and one row's red is `FF0000` while the next is `FF3300`. Saturation and
lightness gate the hue first, so grey, black and white say nothing, and
orange/yellow/violet fall through rather than being guessed at. The result is
`j.flag` (the word) and `j.flagHex` (the colour that said so). A red row is
urgent as well as the old rule (the word in a comment), so `j.urg` is now
"either". Shown as a chip in the row's last cell, beside the other badges (the
cell that already handles overflow), and in the drawer head — **always the
word**, with the colour only as its ink, clamped by `inkFor()` so it stays
readable in the dark theme. An exported file always keeps the sheet's own hex.
Production (2) has its own colours and is read for them separately; neither
sheet's colours are merged into the other.

**Section / category select (`app.js`).** In the grouped view every section
header carries a "Select all n" tick box (indeterminate when part of the
section is picked, and present whether the section is open or collapsed); in
the flat list a "Select all shown (n)" box appears at the left of the toolbar,
but only when something is actually narrowing the list. The John print sheet
view has the same per-section box. All of them go through `pickMany(ids, on)`:
one pass over the ids, one `renderAll()`, whatever the size of the section.

**Production (2), read on its own terms (`parser.js: parseJohnSheet`).**
Production (2) is not a lagging copy of Production: it is a separate sheet the
office keeps for the paper John works from, with its own rows, order, sections
and colouring. `parseJohnSheet(wb)` reads it into
`{id, section, ready, cust, phone, area, wnd, drs, notes, fillHex, inkHex, seq}`
— the fixed columns C/G/I/J/K/M/N, plus the "Notes from Brendan's office"
column found by its header (column BY on the real sheet, past a run of hidden
ones), the row's own solid fill and the row's own ink, and the sections from
its own divider rows. **Nothing from it is merged into the job model, and
nothing from Production is used to draw or print it.** `load()` keeps the
result in `JOHNROWS`; a workbook without that sheet gives `[]`, never null.

**The John print sheet view (`app.js`).** The Show dropdown (`BOARDS`) gains
"John print sheet" after Glass station. It renders in the job list's place:
the eight printed columns, a divider row per section, each row wearing its own
fill and its own text colour with the colour's word beside it, a tick box per
row and a "Select all n" per section, all narrowed by the header search box.
Ticks go into `state.picked` like anywhere else, so the wheel works unchanged.
Nothing in the view polls, feeds or writes anything — it is drawn from the
workbook download the page already made.

**The John print sheet template (`export.js` §7b).** `exportJohnRows(ids,
notes, ctx)` takes job *numbers* and builds the rows from `ctx.john`
(Production (2)) in that sheet's own order, with a divider row per section. A
chosen job that sheet does not have is still printed — from the Production
model, with **no colour at all** and a grey italic "not on John's sheet" line
in its Notes. `buildJohnWorkbook` makes a one-sheet ExcelJS workbook named
"John print sheet" (two-row header with QUANTITY over Wnd/Drs, the sheet's
widths, thin borders, landscape, fit to one page wide, print titles `1:2`) and
`buildJohnDoc` the A4-landscape pdfmake equivalent with page numbers and a
"Printed … by …" footer. File name `John print sheet <date>.xlsx` / `.pdf`.
Choosing **John print sheet** in the Export window hides the field picker and
the card/table choice and leaves the scope (ticked jobs / everything the view
is showing), the format, and a **Continue to notes** button. Opening the
window from the John view preselects that template — on the way in only, so
the choice stays clickable.

**This is the one export in the app that carries a phone number.** The owner
sanctioned it on 2026-09-09 for this template only; the number printed is
Production (2)'s own. `j.ph` (added to the job model) is read only by
`exportJohnRows`, and only for a job Production (2) does not have; the eircode
is carried by nothing anywhere. Every John print's `Dashboard Log` line says so
in words: `Excel · n jobs · John print sheet · with phone numbers`.

**Print notes (`app.js`).** One extra note per job, typed in the notes window
(`#nhost`, "Print notes — John print sheet") before a print: job number,
customer, colour chip, Production (2)'s own note in grey with the Production
comment labelled underneath when the two differ, and a box (2000 characters)
pre-filled from the list. **Print** clears last time's failure marks, asks for
the list permission **once** (`CW.listConsent()` from the click, before any
write — three writes starting together used to open three consent dialogs),
saves the notes that changed one after another through `CW.listUpsert` into
`Dashboard print notes` (Title = job number, `Note`, `By`, `At`, in the
workbook's own site), then builds the file from **exactly the jobs the window
listed** (`NSTATE.ids`, not a fresh lookup — a poll can land in between) and
downloads it. A note that will not save marks its own row "not saved" and does
not stop the print. Clicking the scrim with unsaved typing asks first; Cancel
and Escape do not. The list is read when the notes window opens, not on every
page load; the drawer shows the note read-only under Comments. **No workbook
sheet is written by any of this** beyond the one export log line, Production
(2) included, and the list is never created by code.

Tests: `test_john.js` (the flag parser and `parseJohnSheet` on a two-sheet
ExcelJS fixture with hidden columns, a divider, fills, four inks and a
hyperlink; section select in all three views; the view's rows, dividers,
search and tick boxes; Export preselecting the template from the view and
leaving it clickable; the file taking Production (2)'s phone, area, date and
note; a job missing from that sheet; one consent dialog for three changed
notes; printing `NSTATE.ids` after the job list changed underneath; the
cleared failure mark; and a sweep of every request the run makes) and the John
sections of `test_export.js`.

---

## 17. Glass colours: the floor's work reaches the Production sheet

Built 2026-09-10. Spec: `docs/specs/2026-09-10-glass-colours-two-way.md`,
including its long "Amendments after review" section, which records the one
gap in the brief and the twelve decisions taken while building it. **This is
the first feature in which something done on the floor changes the master
sheet**, so read the boundaries before touching any of it.

**What.** When the floor finishes cutting and hotmelting a job, its DG, TG and
NOT TUFF cells go **yellow** in the workbook; when glazing is done all four
glass cells (TUFF included) go **gold**. It walks back down as well as up —
gold to yellow to blank — mirroring whatever the floor now says. TUFF's yellow
comes from its own new counter rather than from cutting and hotmelting.
**ARCH, ASTRAGAL, FANCY and EXTRA are never written by this feature**, in
either direction: the office ticks those by hand, as it always has.

**Who writes it.** The **office dashboard** (`app.js`), never the tablet. The
tablet has no access to the workbook at all and that is the actual security
boundary; this is the feeder in reverse — the office watches the floor's list
and paints the cells.

**What reaches Production.** A single-cell **fill**, and nothing else, ever:
no value, no row, no formula, no number format, no other sheet. Not even a
`Dashboard Progress` row or a `Dashboard Log` line — the first would be an
invented per-type count (the floor counts one combined DG + TG number) and the
second would become the office's own stamp on the next pass and poison the
last-writer-wins comparison below.

**Where the code is.**

- `station-core.js` — the pure half: `glassColours(record)` returns
  `gold` / `yellow` / `""` per column; `floorStamp(record)` is the newest ISO
  stamp on the floor's row; `COLOUR_TYPES` is the four columns; `stageComplete`
  is "this counter has reached its own total, and there is a total".
- `app.js` — `glassColourPlan(j)` (what to write, or null), `glassColourWrite`
  (the write), `glassColourRun()` (one pass over every job), `stampMs`,
  `glassOfficeStamp`, and the `gc` branch of `pend` / `applyPending`.

**How, in five parts.**

1. **When it runs, and how much it may do at once.** After every `load()`,
   once the feed has brought the list up to date, and after every station poll
   that actually moved something (10 s while somebody is looking at the floor,
   60 s otherwise). It is a walk of the jobs in memory and **not one request**
   when there is nothing to do, which is what lets it be called six times a
   minute.

   It is also **capped**, for the same reason the feeder is and with the same
   numbers: `GLASS_MAX = 60` cells per pass, three writes in flight
   (`stationSend`, the feeder's own lane runner), and one follow-up armed 30 s
   later when the cap cuts a pass short (`glassColourAgain` — exactly one
   timer, never one per job). Rehearsed against the owner's real file: 141 fed
   rows, and switching the feature on after the floor has worked a week with no
   office dashboard open would otherwise plan **362 cell fills in one burst**.
   Day one is about six cells, so this bites on a catch-up, which is precisely
   when nobody is watching. A job is never split across two passes: half a
   job's colours would be a lie on screen for thirty seconds and the other half
   would only be re-planned anyway.
2. **Idempotence, which is the whole reason it is safe to call that often.**
   The desired colour is compared with what the sheet is already showing —
   `j.cp.glass[type]`, i.e. the *held* colour while a write is in the air and
   the *downloaded* colour once the file has caught up — and a cell that
   already says it is not written. Ten turns of the poll after a write send
   zero requests (`test_glasscolour.js` asserts exactly that). A cell carrying
   a colour this feature does not own (the sheet's own Cut green) is left
   alone in both directions.
3. **Last writer wins.** The floor's stamp is `DoneAt` (and the per-stage
   `At`s) on the `Glass station` row — the newest **by parsed time**, skipping
   anything unreadable, because every one of those columns is hand-editable in
   SharePoint and a text maximum let `CutAt: "zzz"` win, answer a time nothing
   could read, and switch the feature off for that job in silence. The office's
   is the newest of the job's glass `Dashboard Progress` rows (`Who`/`When`,
   read through `cpStored`) and the `Dashboard Log` lines for that job's glass
   items (already in memory as `CHANGES`) — so the comparison costs **no extra
   read**.

   Three formats meet in `stampMs`, and it returns **the latest instant a
   stamp can be describing**, not the instant it literally names. That
   distinction is the whole of it: `nowStamp()` writes no seconds, so an office
   tick at 15:00:40 is recorded as `15:00`, and read literally a floor tap at
   15:00:20 looked like the later action and painted the office's own tick back
   out. A stamp that names only a minute therefore covers that minute
   (`+59_999 ms`); one that carries seconds is taken exactly. It also makes the
   answer stable across the round trip — `noteChange` puts a full-second entry
   in `CHANGES` and `load()` later replaces it with the Log sheet's
   minute-precision copy, and both readings now answer the same side, where
   before the dashboard quietly changed its mind half a minute later.

   **A tie goes to the office** (a tie is therefore no write at all). **No
   floor stamp means never written** — an untapped row's counters are the
   office's own seed echoed back. **No office stamp means the floor wins**,
   which is the honest limit: a colour painted by hand in Excel leaves no dated
   record and will be walked back; ticking it in the drawer gives it a stamp
   and it wins.
4. **The hold.** Every write is held in `PENDING` under a new `gc` key — the
   colour's own word per glass type, with its own timestamp — and let go the
   moment the downloaded file agrees, exactly as a checkpoint tick is. Without
   it the next refresh would read the 36-second-old file and the cell would
   flicker. A **reversal** is held the same way, which is what stops gold
   becoming yellow and then flickering back to gold. Both sides can hold one
   cell at once (the office ticked in the drawer, a floor tap arrived two
   seconds later): the newer hold is what is drawn, by comparing
   `t["gc:<type>"]` with `t["cp:glass:<type>"]` — the same rule as the
   writer's. A refused write drops its own hold rather than showing a colour
   the sheet does not have.
5. **Ordering.** Each job's write runs on that job's existing checkpoint chain
   (`CP.cpChain`), so an office tick and a floor colour for the same job can
   never interleave, and the fills themselves go through
   `CW.serialised("Production", ...)` (now exported from `graph.js`) with the
   row re-found by job number immediately before writing. `serialised()` is a
   promise chain inside **one tab** and can say nothing about a second
   dashboard on another desk — what makes that safe is the idempotence test
   and last-writer-wins, not the chain. Two passes cannot overlap either
   (`glassRunning`); the second arms the follow-up rather than giving up, so a
   long catch-up still finishes.
6. **When the workbook refuses.** A failed write drops its own holds — what the
   sheet still says *is* the old colour — and then **backs off**, because
   without one it would go out again on the next poll, and `stationPoll`
   answers "moved" every ten seconds while the floor is tapping. That is a
   write storm against the live workbook, and the trigger is ordinary: somebody
   opens the file exclusively in desktop Excel mid-shift. Each job counts its
   own failures and waits 1 minute, then 5, then 15
   (`GLASS_BACKOFF_MS`); after `GLASS_FAIL_MAX` (5) it is given up on until the
   page is reloaded. A success clears the record. All of it is surfaced in the
   footer the way the station feed's own failures already are — *"glass colours
   not saved"*, with the count, the reason and what happens next in the
   tooltip — because a workbook refusing a write is not something to leave in a
   console nobody has open.

**The TUFF counter (`Glass station`, `station-core.js`, `station.js`).** A
fourth thing the same person counts on the tablet, with its own quantity off
the sheet's TUFF column. New columns: `Tuff`, `TuffBy`, `TuffAt` (the floor's)
and `TuffTotal` (a job fact, fed like `Total`). It is deliberately **not** one
of the three glass stages: `STAGES` / `STAGE_KEYS` still mean the three and
`ALL_STAGES` / `ALL_STAGE_KEYS` are the four, so the job's `Total` is still
DG + TG, `finished` (the gold card, the Finished group, the job row's
`Glass 8/24` chip) is still the three, and the feeder never seeds `Tuff`. It
gets a row and a number of its own (`stageLeft` against `TuffTotal`): eight
glasses to cut and eleven tuff show as **two** numbers, "Cutting 8 left" and
"Tuff 11 left", never the nineteen that adding them would have said. The stepper is drawn only on jobs that
have tuff on them. **NOT TUFF gets no counter at all** — its colour is derived
from the glass stages.

**The office's own yellow (`ST.officeSeed`, owner 2026-09-10).** A yellow glass
cell now *means* "cut and hotmelted", so the seed reads it that way: when
**every** one of a job's DG and TG items is at least yellow, the floor's row
starts with `Cut` and `Hotmelt` at the total and **`Glazed` at nought**, and
the cell round-trips to yellow instead of being flattened to blank. It had to
be fixed — a yellow item has no stored count, so the row used to seed at
nought and the first tap on that job painted the office's own mark out. It was
found on the single yellow glass cell in the owner's whole workbook. Glazed
staying at nought is the point: yellow says glazing is *not* done, and seeding
it would make the cell read gold. A gold item is untouched and still
round-trips to gold. **Every, not some**: the floor's row holds one combined
DG + TG number and there is nowhere to put a per-type split, so a job with DG
yellow and TG blank keeps the office's own count instead — seeding it to the
total would tell the floor that glass which still needs cutting is cut, and a
wrong instruction on the workshop screen is worse than a lost colour on a
report. That job can still lose its yellow once the floor taps; it is a known
limit, recorded in the spec's Amendment B5. None of this widens rule 3's
seeding exception — it changes how the three counters are *derived*, not which
columns may be seeded, and `Tuff` is still never seeded.

**The lock (`OfficeDone`).** When the office has ticked every DG and TG item
of a job off, the feeder writes `OfficeDone = "Yes"` and the job goes
read-only on the tablet: every stepper greyed, `tap()` refusing, and a line on
the card saying why. No new control in the office — it is derived from the
glass checkpoints the drawer has always had, so un-ticking one unlocks it, and
only the office can. The drawer's Glass station section says so, so the office
knows why the floor cannot move it. A tap already queued on the tablet when
the lock arrives is **dropped rather than sent** (the office acted later) but
never quietly: `dropBlocked()` moves it to a `BLOCKED` note in `localStorage`,
warns to the console and draws it on the card in red — "Cutting 5 was not
saved — the office marked this job finished first" — until the office unlocks
the job. No log line is written, because the counter never landed. The check
runs after every read of the list **and** again at the top of `flushQueue`,
because the lock can arrive in the poll that ran while the queue was waiting.

**Honest limits, told to the owner.** Colours only move while an office
dashboard is open (the floor working on a Saturday means the sheet catches up
on Monday). A mis-tap on the floor now reaches the master sheet — reversible
by tapping back, or by the office ticking over it, but no longer contained to
the tablet. Yellow is a new convention: one cell in the live sheet used it
before this. And a colour set by hand in Excel rather than in the dashboard
will always lose the last-writer contest, because it leaves no dated record.

**Tests.** `test_glasscolour.js` (41 checks): the rule per column in both
directions and all the way back to blank; gold requiring glazing with no
exception; `floorStamp` skipping a stamp that will not parse rather than
letting it win; the stamp formats **and the minute a seconds-less stamp
covers**, including the boundary and the round trip that used to change its
mind; last-writer-wins both ways, the same-minute case, and each side with no
stamp at all; the Log as the office's stamp (right job, right item);
idempotence over ten polls; the hold, its release and a reversal without a
flicker; an untapped row and a job marked ready both left alone; the office's
yellow surviving the seed round trip (and the mixed job that cannot); a
refused write, the growing backoff, giving up, the footer's words, and one
success clearing it; the 60-cell cap, exactly one follow-up, the remainder
going out on later passes and no job split; two passes unable to overlap; the
drawer's own ticks unchanged; and over the whole run — every write to
Production a single-cell fill whose body is a colour and nothing else, every
fill in column AY/AZ/BA/BB and never in BC-BF, and no sheet touched but
Production and the dashboard's own two. `test_station.js` (235) adds the tuff
counter, the lock, the dropped queued tap, the seed's new yellow rule, and
proves over its own run that the tablet still touches no workbook.

---

## 18. An office clear reaches the floor's counters

**Spec.** `docs/specs/2026-09-10-office-clears-the-floor.md`. Changes
`CLAUDE.md` rule 3, by the owner, on 2026-09-10.

**The bug.** The owner spent an afternoon on what looked like three faults and
was one gap. Un-ticking a job's glass in the office cleared the **workbook**
correctly every time — the cell really did go white — but nothing ever cleared
the **floor's** counters. They stayed at 49/49/49, the tablet stayed gold for
ever, and the only way back was somebody tapping `−` forty-nine times. Two
earlier fixes in this area shipped and neither addressed it; the master
dashboard showing gold for up to a minute afterwards is SharePoint's ~36 s
download lag and is not this.

**What happens now.** When the office **clears a job's glass checkpoints** —
the group **Clear**, or a per-item clear that leaves no other DG/TG ticked —
the office writes that job's `Glass station` row's `Cut`, `Hotmelt`, `Glazed`
and `Tuff` to **zero**, plus `DoneBy`/`DoneAt`. Six fields, and nothing else.
`DoneBy`/`DoneAt` are written on purpose: last-writer-wins reads
`ST.floorStamp` off them, so zeros under a stale stamp would read as old news
and the tablet's "last touch" line would name the wrong person. The per-stage
`By`/`At` pairs are never written — they say who did that stage's work, and
nobody did. No `Station log` line: it is an office action and goes in
`Dashboard Log`, as `Floor glass counters`.

**The trigger is a transition, not a state.** `officeGlassEmpty()` asks
`ST.officeSeed` — the very function that seeds a row from the office's record —
whether that record now says nothing of the job's DG and TG is done;
`officeClearsGlass()` fires only when it *newly* says so. A tick made when the
record already said nothing is not a clear, so clearing a hand-ticked ARCH on a
job whose DG is blank cannot wipe the floor's counters. There is deliberately
**no reconciliation pass**: a state-driven version of this would zero any row
the office happens not to have ticked, including a job the floor is working on
right now.

**A row the floor never tapped is left alone.** Its counters are the office's
own seed echoed back, and the **feeder** puts them right on its next run
(`sliceHash` carries the seed, so that is the next load). Writing `DoneAt`
there would mark the row touched for ever and switch its seeding off — a worse
bug than the one being fixed. This is a deliberate narrowing of the brief,
which wrote unconditionally.

**The confirmation.** When the row has a floor stamp and any counter above
nought, the office is asked first, naming what will go: *"The floor has
recorded 49 cut, 49 hotmelted, 12 glazed on this job. Clearing the glass here
will set all of those back to zero. Clear it anyway?"* (`ST.clearWarning`, a
pure function). Answering **no writes nothing at all** — not the floor's
counters and not the workbook half, because the question is asked before
`pend()` and before the burst.

**A queued tap that meets the zeros is decided on the stamps.** A queue entry
holds an **absolute** number — a `+1` made against 40 carries 41 — so letting a
drop through unchanged would not "apply the tap on top of the zeros", it would
put the whole count back and leave the row at 41 cut, 0 hotmelted, 0 glazed.
`rebaseQueue` therefore compares the row's `DoneAt` with the tap's own `at`,
which is exactly what writing `DoneAt` on a clear is for: **a row stamped after
the tap was made is the later word and the tap is dropped**; a tap made after
the clear applies on top of the zeros. A tie, or a stamp that will not parse,
leaves the tap alone. Nothing is `BLOCKED` either way — unlike the lock, a
clear **unlocks** the job, so a dropped tap can simply be tapped again, which
is why stranding it the way `dropBlocked()` does would be the wrong shape here.
The real case this closes is undramatic: the workshop wifi drops, the floor
taps, the office clears the job *because something went wrong on it*, the wifi
returns — and without the stamp check the tablet would go back to 49 cut, which
to the office is indistinguishable from "the un-tick didn't stick".

**Why it cannot be reached any other way.** Four independent locks: two call
sites, both on a clear branch; `FLOORCLEAR_OK`, the office's answer to the
question, set in one place and spent once — and keyed to the **write** that
must land (`job|item` or `job|group`), never to the job, because a job has
several checkpoint bursts at once and any of them landing would otherwise spend
the answer and clear the floor before the glass write had landed;
`clearFloorGlass()` re-deriving the reason from the office's own record and the
floor's row at write time rather than trusting the caller; and `ST.officeClearFields(who, at)` — a name and a
time in, six fields out, every counter a literal nought — so the path has no
argument through which a counter or a per-stage stamp could enter.
`test_station.js` asserts the counts of each of those in the source.

**A refused write is owed, not lost.** The workbook half has landed by then, so
forgetting it would leave the tablet gold — exactly the bug. `FLOORCLEAR_OWED`
counts the refusals, one timer retries after 30 s, and after three attempts the
office is told in a toast that names the numbers still on the tablet. The
retry re-derives, so a job the office has re-ticked meanwhile is dropped.

**The office is absolute: the writer stands down while an office change is
settling.** The owner restated the rule on 2026-09-10 — *"the hierarchy is
Excel, then master dashboard, then glass. Any change from the dashboard is
absolute … an un-tick from the dashboard is absolute — no thinking, no
arguing."* So `glassColourPlan` opens with `if (officeSettling(j, log)) return
null;` and makes **no decision at all** about a job while the office's own
change on it is still settling — no stamp comparison, no plan, no write.
`officeSettling` is true while a `glass:` checkpoint hold exists on the job
(`pend()` dates it at the click; `applyPending` releases it when the download
agrees) **or** while the per-job office stamp is younger than `PENDING_MS`.
Both, because the office must be safe in every place its action is read from,
and the window is **per job**.

The reason is not that the comparison is wrong — it was corrected twice — but
that inside that window it is fed copies that have not caught up (the download
~36 s behind, the floor's list a poll behind, the clear's own `DoneAt`
indistinguishable from a tap), and this feature exists to carry the **floor's**
work up to the sheet, not to second-guess the office. Outside the window
nothing changes: last-writer-wins still decides and a genuine floor tap still
paints. **The cost, plainly: a floor tap made inside the window is deferred by
up to three minutes** — never lost, because the writer re-plans from the
current list on every pass, and a pass that deferred anything now arms the same
30 s follow-up a capped pass uses. **The hole it does not close** is a second
dashboard, or a reload, inside the download lag: it has no record of the
office's action at all, so it cannot stand down for it. Closing that needs an
`OfficeAt` column on the `Glass station` row — a decision for the owner.

**The office's stamp is per JOB, and the clear's own `DoneAt` is the office's.**
Fixed 2026-09-10 after the owner saw it in production: `glassOfficeStamp` was
asked **per column**, so for every column the office had not named *in that
action* it found no hold, no Progress row and no Log line and answered a
literal `0` — and any floor stamp at all beat it. Pressing Clear on TG alone
therefore painted out the TUFF and NOT TUFF the office had ticked by hand. And
the floor stamp it lost to was one the office had itself just written:
`clearFloorGlass`'s own `DoneAt`, a fraction of a second after the click, while
`glassLogStamps` deliberately refused to count the `Floor glass counters` line
that records the same act. One half of the office's own clear was counted for
the floor and the other half for nobody.

`glassOfficeJobStamp(j, log)` now answers **the newest thing the office has
said about this job's glass, on any column**: every glass item's hold (stamped
at the click by `pend()`), every `Glass …`/`Glass: …` Log line for the job
whichever column it named, the `Floor glass counters` line, every glass item's
Progress `When`, and `OFFICE_FLOOR_AT[job]` — this dashboard's own memory of
the `DoneAt` it wrote, which is the only record of a clear until the Log line
has been written and read back. `glassOfficeStamp` takes the max of that and
its per-column sources. It is computed once per job per pass (`byId` is a
linear scan) and `OFFICE_FLOOR_AT[job]` is dropped once the floor has moved
that row since, so the map stays bounded. The rule is unchanged: a floor tap
after the office's action still carries the later stamp and still wins.

**Known gap.** `FLOORCLEAR_OK` lives in memory only, so **any clear interrupted
before its workbook write lands leaves the floor's counters standing, with no
automatic recovery** — three ways: a per-item burst that never fired is
replayed by `cpReplay` and clears the workbook only; one already in flight is
marked `sent` before the await and is never replayed at all; and the **group**
path has no queue and no replay of any kind, which is the owner's usual way of
clearing a job. Persisting the answer was rejected — it would put a
hand-editable token in `localStorage` on the one path that writes a floor
column. The recovery, in `docs/SUPPORT.md`: tick the glass done again, then
clear it again. Not forty-nine taps.
