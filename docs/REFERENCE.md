# Reference: what has been built, and how

One page per feature, in the order they were built. Each entry says what the
feature does, where the code is, how it works underneath, what it writes and
where, and which tests cover it. Read this before touching anything; read the
spec in `docs/specs/` for the full brief of a feature. Names of people and
addresses are placeholders throughout ("the admin", "the colleague").

Last updated 2026-09-09 (row colour code, section select, the John print sheet).

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
Tests: `test_checkpoints.js`.

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
   job number, the customer, "12 glasses" and the three steppers (−, +, All /
   None) on the card itself — nothing to expand, no bars, nothing to scroll
   inside — plus a search box in the header (`boardFilter`, job number or
   customer). Ten quiet minutes lock the
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
   read-only board: one card per job with "12 glasses", the three counters
   with who · when under each, "last: …" from the log, and the same gold rule
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

**Tests.** `test_station.js` (195 checks) proves, over the whole run, that no
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
| `test_export.js` | builders, no phone/eircode, logging | 45 |
| `test_phases.js` | phase derivation and the wheel | 26 |
| `test_phases_list.js` | the phases list, scope split, dedupe | 35 |
| `test_station.js` | floor stations end to end (offline) | 158 |
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
