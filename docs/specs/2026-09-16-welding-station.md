# Welding station: the second floor station, and the pattern for the next ones

**Status: approved to build, 2026-09-16** (the owner answered the three open
decisions; see "Decisions taken" at the end).

## Context

The glass station (`docs/specs/2026-09-08-glass-station-v3.md` and after)
gave the glass area a tablet page fed from the master dashboard. The owner
now wants the same for **PVC welding**: a tablet page for the welders, a
board on the master dashboard for the office, progress kept in SharePoint,
and **nothing written into the workbook**. A third station (PA Lam) follows
later and must be a small job once this one exists.

The owner supplied the `Cut & Weld PVC` sheet as a **header template**: its
column layout is the Production sheet's, and the header colours say which
columns the floor may see (green) and which it may not (red). The dashboard
reads the `Production` sheet as always; the template sheet is read by
nobody. Colours as read from the owner's PDF, 2026-09-16:

| green (floor sees) | red or white (never on the floor) |
|---|---|
| COMMENT, JOB NO., CUSTOMER name, QUANTITY (WND, DRS), the **sent to floor** date | OFFICE NO., PHONE NO., AREA, EIRCODE, WINDOWS COLOUR, SOLD, STAMP, the third date column, Ready to print |
| 14 product groups: CASEMENT WINDOWS, POLARIS 85MM CASEMENT, 4000 CASEMENT, 7000 CASEMENT, Polaris 85 TILT & TURN, 4000 TILT & TURN, 7000 TILT & TURN, PVC FRENCH WDS, Arch & Angles, TH W, SIDELIGHTS, PVC DOOR, Super door, PVC SMART (Smart slide) | ALU CLAD WINDOWS, ALUCLAD TILT & TURN, every GLASS UNITS column, Bifold, COMPOSITE, DOORS EXTRA, DOORS DONE |

Within a green group only **F (frames) and S (sashes)** reach the floor.
T (transoms) is green on the template in most groups and is **still never
shown**: the owner's words, "there is no transomes even if it is green".

The set of groups is **not hard-coded**: it is every F/S/T product group the
parser already finds on the Production sheet (`mapSheet` → `j.prods`) minus a
short deny-list (`ALU CLAD WINDOWS`, `ALUCLAD TILT & TURN`, `BIFOLD`,
`COMPOSITE`), so a new green group on the sheet is fed without a code change
and a new red one needs one line.

## Owner consulted, 2026-09-16 (answers, in substance)

1. **COMMENT on the floor:** yes, with runs of six or more digits removed
   (option B). Recorded as a rule-3 decision: the comment column can carry
   phone numbers, and this strip is the guard.
2. **sent to floor:** shown on the card; a blank stays blank; the tablet can
   filter to "sent to floor" jobs.
3. **F and S cells** hold plain numbers: frames to weld, sashes to weld.
   Blank means none.
4. **Steppers** (−, +, All) per Frames line and per Sashes line, like glass.
5. **Seed from the office:** an F or S cell **gold** means that component is
   fully welded; yellow or anything else means not done. The floor sees
   "In production" jobs only; the master sees every job, finished ones
   included.
6. **Office powers:** the same as glass, and editable from the office's own
   welding view: set and clear progress per job, even on a finished job
   ("they might make it wrong"). People and what they did must be logged and
   visible on the office's welding view.
7. **Finished:** a job whose every frame and sash is welded goes to the
   Finished group on the tablet and comes back when the office reduces it.
8. **Same tablet, same station account**, a different station page.
9. **The `Floor stations` site is created by script**; the owner now has
   admin. Glass stays where it is for now.
10. Station comments already shipped (`30c9f86`); this builds on top.
11. **Code shape is my call**, with the instruction to make later changes
    one-place changes and to document how the pieces fit.
12. **Colours:** no colour = nothing welded; **yellow** = started, not
    finished; **green** = every frame and sash welded. Same words at three
    levels: a Frames/Sashes line, a product-group line, the job card.

## Hard rules (unchanged, and how they bite here)

- Rule 1: **no workbook write of any kind** for this feature. The welding
  station never paints a cell, and the office's welding edits never do
  either. The gold F/S cells are read as a seed, once, on an untouched row.
- Rule 2: floor stations never touch the workbook; the office never writes
  `Station people` or `Station log`; the feeder never writes the floor's
  counters on a touched row and never deletes. **One extension, mirroring the
  glass clear:** the office's welding edits write the floor's counters and
  the last-touch pair (`DoneBy`/`DoneAt`) from the office's welding board,
  logged in `Dashboard Log`, never in `Station log`.
- Rule 3: no phone numbers or eircodes off the workbook. COMMENT is fed with
  every run of 6+ digits replaced by `…` (`stripDigits`) and anything the
  export's eircode check would catch removed too, capped at 140 characters;
  no other job fact leaves the green set above.
- Rule 4: no real names in the repo. The station people are "the welder",
  "Person A" in every test, doc and fixture.

## Data model

All floor lists for this station live in the **`Floor stations`** site
(created by script, see "Site and lists"). The glass lists stay in the
workbook's site until a separate, planned move.

### `Welding station` — one row per job **and product group**

A job with casement windows and a PVC door has two rows. Job facts are
repeated on every row of the job so a row is complete on its own (the
tablet groups rows into cards by `Job`).

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, unique | feeder | `JOB\|GROUP`, e.g. `R5303\|CASEMENT WINDOWS` (group in the sheet's own upper-cased header words, whitespace collapsed) |
| Job | text | feeder | the job number |
| Group | text | feeder | the product group label as on the sheet, for display |
| GroupSeq | number | feeder | the group's column order on the sheet, so cards list groups in the sheet's order |
| Customer | text | feeder | customer name, max 70 chars |
| Comment | text | feeder | COMMENT after the rule-3 strip, max 140 chars, may be blank |
| SentToFloor | text | feeder | the "sent to floor" cell as text (typed by hand on the sheet; shown as typed), may be blank |
| Wnd / Drs | number | feeder | the QUANTITY cells |
| Frames / Sashes | number | feeder | the F and S cells of this group: to weld |
| Seq | number | feeder | the job's position in the master list |
| Section | text | feeder | the job's section name on the sheet ("In production", "Finished", …) — the tablet shows `In production` only; the master shows all |
| Active | text | feeder | `Yes` while the job is on the sheet with this group having F or S > 0; `No` afterwards. Never deleted |
| FedAt / FedBy | text | feeder | last feed that changed this row, and who was signed in |
| FramesDone / SashesDone | number | station; feeder as a **seed on an untouched row only**; the office from its welding board | welded so far |
| FramesBy / FramesAt, SashesBy / SashesAt | text | station; the office | who last moved that counter and when (ISO) |
| DoneBy / DoneAt | text | station; the office | last touch of any counter on this row |

Seeding (feeder, only when creating a row or when `DoneAt` is empty, checked
twice like glass): `FramesDone = Frames` when the office's record for
`prod:<group>:f` says `done` (gold), else `0`; same for sashes with `:s`.
Yellow seeds nothing (owner, answer 5). The office's record is
`cpStatus(j, item)` from `Dashboard progress` — the same source the Excel
gold is a copy of — so a cell painted gold by hand in Excel reaches the seed
through the existing safeguard (`docs/REFERENCE.md` §19), not by reading the
colour here.

### `Station people`, `Station log`, `Station comments` — in the new site

Created fresh in `Floor stations` with the same columns as the existing
ones (`docs/STATIONS.md`). Welding people rows: `Station = Welding`,
`Stages = weld`. The glass site's three lists stay as they are; the glass
rows are copied across on the day glass moves (not this spec).

`Station log` lines for welding: `Station = Welding`, `GlassType = <GROUP>`
(the column name is kept, the value is the product group), `Stage = frames`
or `sashes`, From/To/Who/At as today.

### Which site does a page use — the flip hazard

`graph.js` `stationSite()` looks for a site named `FloorStations` first and
falls back to the workbook's site when it is missing. **Creating the site
would therefore move the glass page to a site without a `Glass station`
list on its next re-check.** Decision needed (see "Open decisions" 1).

## Behaviour

### The feeder (master dashboard, after every load)

`feedWelding()` beside `feedStation()`: build the welding slice from
`ALL` (every job on the sheet, every allowed group with F or S > 0; the
section name on each row), diff against the list (`ST.feedPlan` generalised
to take the station's field lists), send adds and patches in the same three
lanes and 60-writes-per-run budget, seed per the rule above, never delete,
throttle on a hash. A missing list is reported in the footer the way
`Glass station` is, and never blocks the glass feed or the checkpoint
writes: the two feeders run one after the other, each in its own try.

### The tablet page `welding.html`

Same shell as `glass.html`: sign in with the station account, "Who are
you?" from `Station people` filtered to `Station = Welding`, PIN, ten-minute
lock, Switch person, search, Floor log, the comments composer and thread
(from `ST.stationComments({ station: "Welding" })`), 10 s delta polling,
incremental drawing.

**Card, one per job** (rows grouped by `Job`), in `Seq` order, `In
production` rows only, `Active = Yes`:

```
R5303  Person A                              sent to floor 12.09
Wnd 6 · Drs 1     "…rang, will collect Friday"
┌ CASEMENT WINDOWS ─────────────────────── yellow ┐
│ Frames  3 / 6   [−] [+] [All]   ▮▮▮▮▮▯▯▯▯▯      │
│ Sashes  0 / 8   [−] [+] [All]   ▯▯▯▯▯▯▯▯▯▯      │
├ PVC DOOR ───────────────────────────────  green ┤
│ Frames  1 / 1   [−] [+] [All]   ▮▮▮▮▮▮▮▮▮▮      │
│ Sashes  1 / 1   [−] [+] [All]   ▮▮▮▮▮▮▮▮▮▮      │
└─────────────────────────────────────────────────┘
```

- A line with 0 to weld is not drawn. A group with nothing to weld is not
  fed.
- Colour: line = green when done = total, yellow when 0 < done < total,
  none at 0. Group = green when every drawn line is green, yellow when any
  line is started, none otherwise. Card = the same over its groups.
- Groups on the board: **Active** and **Finished** (every line green). A
  finished card is still tappable (−) and comes back when reduced, by the
  floor or by the office.
- Header: "N left" for the signed-in person = frames left + sashes left over
  their visible cards (`weld` is the only stage, so anyone signed in holds
  it). A filter chip **Sent to floor** hides cards with a blank `SentToFloor`;
  remembered in localStorage; default per open decision 3.
- Every tap is one absolute write (`FramesDone` or `SashesDone`, its By/At,
  `DoneBy`/`DoneAt`) plus a `Station log` line, queued and rebased exactly as
  glass (`applyTap`, `rebaseQueue`). No lock column: there is no `OfficeDone`
  for welding, because the office edits the same counters instead.

### The office's welding board (master dashboard, Show ▸ Welding station)

A `BOARDS` entry `["welding", "Welding station"]`. In the job list's place:

- **One row per job** on the sheet, every section, in sheet order, with:
  job, customer, section, an overall bar (`welded / to weld` over the job's
  frames and sashes), `Frames x/y`, `Sashes x/y`, last touch (who, when),
  the unread-notes icon from station comments. Coloured by the same
  none/yellow/green rule.
- **Expand a row** (click): one line per group with its own bar, and the
  **office's steppers** (−, +, All, None) on Frames and Sashes. A change
  writes the list at once (the office's name in By/At, `DoneBy`/`DoneAt`),
  one `Dashboard Log` line per change ("Welding: R5303 CASEMENT WINDOWS
  frames 3 → 6"), no `Station log` line (rule 2). The tablet sees it on its
  next poll; a queued floor tap older than the office's edit is dropped and
  shown, as with the glass clear.
- **Floor log** panel below the board: `Station log` and `Station comments`
  filtered to Welding, newest first, with the counts bar — the existing
  `renderStationLog` pointed at the welding lists.
- Search and section filter as the home list has.
- The job **drawer** gets one read-only line under Windows: "Welding 9 / 16"
  (link "open the welding board"). No new column on the job row (owner's
  rule, 2026-09-09).

The office's edits are the second sanctioned office write of a floor
counter, after the glass clear. They are gated by list consent like every
list write and skipped quietly without it.

## Code shape (answer 11: my call)

- `station-core.js` keeps the station-independent pieces and gains nothing
  glass-specific: people, PIN, lock, log rows, delta helpers, comments,
  `feedPlan`/`sliceHash`/`floorOnly` generalised by a **station definition**
  argument (defaults = the glass definition, so every existing call and all
  246 station checks pass unchanged).
- `welding-core.js`: the welding definition (`WELD`: list name, fields,
  stages, seed, slice, tap, colour and done rules, `stripDigits`), pure,
  loads in Node and the browser.
- `welding.js` + `welding.html`: the tablet page, built on the same
  primitives `station.js` uses; anything the two pages share verbatim moves
  into `station-core.js` or a small `station-ui.js` rather than being copied.
- `app.js`: `feedWelding`, the board renderer, the office steppers, the
  drawer line. `graph.js`: no change beyond whatever the site decision needs.
- A new section in `docs/STATIONS.md`, **"Adding a station"**: the
  definition object, the three files, the lists, the people rows, the feeder
  hook, the board entry — written so PA Lam is that list and nothing more.
- `docs/ARCHITECTURE.md` gets the welding data flow; `docs/REFERENCE.md` a
  §21; `docs/HISTORY.md` the decisions above; `CLAUDE.md` rule 2 gains the
  office-edits extension with the date.

## Site and lists (script, owner watching)

`make_floor_site.py` in the scratchpad (never the repo): signs the owner in
interactively, creates the private site `Floor stations` (the admin as
owner, the station account and the colleague as members), then the four
lists with the exact columns above, `Title` unique on `Welding station`
only, and the welding `Station people` rows the owner dictates. Idempotent,
with a `check` mode that lists what exists and what is missing. Rehearsed
first against a throwaway site name, then that site deleted, then the real
run. If site creation through the API is refused for this tenant, the
fallback is the two-minute manual site creation in `docs/STATIONS.md` §2 and
the script does the lists.

## Tests to deliver

`test_welding.js` (Node, no network): the slice from a fixture with three
jobs (one two-group, one with T only → not fed, one in the Finished
section), the rule-3 strip (six digits stripped, five kept, an eircode-shaped
token removed), the seed rule (gold → total, yellow → 0, the untouched-row
guard), `feedPlan` with the welding definition (add, patch, no-op, never
delete), tap and rebase, the three-level colour rule, done/finished
grouping, "N left", the sent-to-floor filter, the office edit fields (By/At,
DoneBy/At, nothing else), and the Title key. The glass suites must pass
unchanged. Grep gates: no `setFill|clearFill|setValues|appendLog|
saveProgress|moveJobRow|batchWrite|/workbook` in `welding*.js`; no real
names anywhere.

## Browser check

`stub_station.js` extended with the four lists in the new site: tablet card
tap, colour change, finished group, office stepper on the board reaching
the tablet on the next poll, missing-list quiet state, comments on the
welding card. Screenshots to the owner.

## Not built (say so)

- Nothing paints the Production sheet from welding progress (owner, answer
  5b). The office ticks F/S gold by hand or in the drawer as today.
- No reply channel from the office to the welding tablet (as for comments).
- No move of the glass lists (owner, answer 9).
- Cutting is not a stage here (owner: "no cutting, it should just have
  welding").

## Open decisions (need the owner's answer before code)

1. **The site flip.** Creating `Floor stations` moves the glass page to it
   automatically (see the hazard above). Options: (a) pin glass to the
   workbook's site in `graph.js` with a one-line constant until the glass
   move is done — recommended, smallest change, reversible; (b) do the glass
   move in the same script run (copy the three lists' rows, verified,
   rehearsed) — the planned end state, but a bigger day; (c) name the new
   site something other than `Floor stations` — avoids the flip, leaves two
   sites for ever.
2. **Office edits and `Station log`.** Rule 2 says the office never writes
   `Station log`. As specified, office edits go to `Dashboard Log` only, so
   the tablet's timeline shows the office's change as a last-touch line
   ("the office, 14:02") but not as a log entry. Fine, or should rule 2 get
   a dated exception so the floor's log carries office edits too?
3. **Sent-to-floor filter default** on the tablet: off (all production jobs)
   or on (only jobs with the date)?

## Decisions taken (owner, 2026-09-16)

1. **Site flip: option (a).** `graph.js` pins the glass page to the workbook's
   site with one constant (`STATION_SITE_PINNED_OWN = true` on the glass
   definition) until the planned glass move; the welding page and the
   welding feeder resolve `Floor stations` directly. Reversible by flipping
   the constant on the day glass moves.
2. **Office edits stay out of `Station log`** (rule 2 unchanged). The owner's
   answer was about visibility: the floor's log must be visible on the
   office's welding board, which the board's Floor log panel does. Office
   edits are recorded in `Dashboard Log` and show on the tablet as the
   row's last touch ("the office").
3. **Sent-to-floor filter: same as glass**, i.e. every "In production" job
   is shown by default and the chip narrows it when tapped.

### Owner, later on 2026-09-16: "make sure you do not break the glass dashboard"

The glass tablet is live and in daily use. Nothing in this build may change
what it does, corrupt its rows, or move its lists. Option (a) stands because
it is the option that touches glass least (one constant; no data moves).
Option (b), moving glass into the new site, is the fallback **only** if the
pin cannot be shown to work, and would be its own rehearsed day. Two
sequencing rules follow, and they are hard:

1. **The pinned build ships before the site exists.** The live build number
   must show the pin before `make_floor_site.py` creates `Floor stations`;
   otherwise live glass tablets would flip to the new site on their next
   ten-minute re-check. The site script refuses to create the real site
   unless the operator confirms the live build number on the command line.
2. **The rehearsal site is named `FloorStationsTest`**, never
   `FloorStations`, so a rehearsal can never match the glass lookup.

Glass regression proof before ship: every glass suite green and unchanged,
the glass browser check (`prev/browser_check.js` pattern) passing on the
new code, and a read of the diff for every line that `station.js` or the
glass feeder executes.
