# Status lives in a list; the Excel colour is a copy

**Date:** 2026-09-11
**Status:** brief for the owner — **not approved, not built**
**Replaces the mechanism behind:** `2026-09-10-glass-colours-two-way.md`,
`2026-09-10-office-clears-the-floor.md` and every hold/reconcile amendment on
them. The features stay; the plumbing changes.
**Consulted:** the owner asked for it on 2026-09-11 ("yeah that should do
it"), after two days in which five commits each made one mechanism more
careful and none made a stale copy true.

---

## 1. The problem in one paragraph

The dashboard never reads the truth; it reads copies — a download of the
workbook ~36 s behind, and a poll of the floor's list ~10 s behind. A tick
never makes those copies disagree, so it looks instant. An un-tick makes every
copy disagree with the office for about a minute, and four mechanisms exist
only to keep believing the office over the copies: the `PENDING` hold, the
45 s reconcile, the colour writer's last-writer-wins, and the feeder's
`OfficeDone` derivation. Each failed this week in turn, and on 2026-09-11 the
writer painted a job gold at 08:15 with nobody clicking, from a copy of the
floor's row that said 49/49/49 while the tablet said 0. **No mechanism can
make a copy true. So status stops being read from a copy.**

## 2. The rule

**A SharePoint list is the single record of checkpoint status. Both screens
read it directly. The Excel colour is written *from* it for people to look at
in Excel, and is never read back to decide anything.**

- The list answers in about a second, and the dashboard already polls lists
  every 10 s. There is no lag to hold against, so **there is no hold**.
- The office and the floor each write the list. The list is one record, so
  two dashboards cannot disagree, and a refresh changes nothing.
- The Excel cell colour becomes an *output*: when the list changes, the
  dashboard paints the cell. If the paint is late, Excel is late — the
  dashboard and the tablet are not.

The sheet stays the master for everything the office types into it (dates,
comments, quantities, sections, row colours). Nothing about how the office
uses Excel changes. **One rule the office must accept:** checkpoint colours
are set in the dashboard, not by painting cells in Excel. A hand-painted cell
is not status any more.

## 3. Data model

A new list in the workbook's site: **`Dashboard progress`** (replacing the
`Dashboard Progress` *sheet*, which stops being written and is left in place).

| column | type | meaning |
|---|---|---|
| Title | text | `JOB|ITEM`, unique — e.g. `R5033|glass:tg` |
| Job | text | job number |
| Item | text | checkpoint key, exactly as today (`win`, `drs`, `glass:tg`, `prod:<name>:<f|s|t>`) |
| Done | number | count done |
| Total | number | count total (so the list is readable on its own) |
| Status | text | `done` / `process` / `` — the word, so nobody has to derive it |
| Who | text | short name |
| When | text | ISO, second precision |
| Source | text | `office` or `floor` — who last set it |

The dashboard reads this list the way it reads the station lists: a full read
at boot, then delta every 10 s (and at once after its own write). It caches
the rows in memory keyed by `JOB|ITEM`.

## 4. What changes, feature by feature

**The drawer's checkpoint rows** read `Status`/`Done` from the list. A click
writes the list first (one `listUpsert`, ~1 s), updates the in-memory cache
at once so the row moves instantly, then paints the Excel cell as a copy.
There is no `PENDING` hold for checkpoints and nothing to reconcile: the next
delta simply confirms what the cache already says.

**The job row's `Glass` column, the office board, the drawer's Glass station
section** — unchanged; they already read the floor's list.

**The floor's work reaching the sheet.** The office watches the floor's list
as today, derives the colour word for DG/TG/TUFF/NOT TUFF as today
(`glassColours`), and **writes it to `Dashboard progress` with
`Source = floor`**, then paints Excel from that. It never reads the Excel
colour. Last-writer-wins survives in one place only: the list row's `When`,
which both sides write at second precision, so a floor tap and an office click
on the same item resolve on real timestamps with no lag in the loop.

**An office un-tick** writes the list (`Status = ""`, `Source = office`),
paints Excel white, clears the floor's counters (as granted on 2026-09-10) and
releases the lock — and that is the end of it. Nothing can argue, because
nothing reads a copy. **This is the whole point of the change.**

**The lock.** `OfficeDone` is derived from the list, never from what the
screen happens to be showing.

**The parser** stops reading checkpoint colours as status. It still reads
everything else. One-time import: for a job with no list row, adopt what the
Excel colour says *once*, write it to the list with `Source = import`, and
never look at the colour again.

## 4a. Safeguard: a colour changed directly in Excel is adopted (owner, 2026-09-11)

The owner asked: *"if anyone changes the colour in the Excel sheet directly,
within ~2 minutes recognise the change was done outside the dashboard, update
the list with the changes, and log who made them and what changed."*

The owner then said the two minutes was "a measure, not an actual value" and
asked for it as fast as possible. **The window is replaced by an exact check;
there is no waiting period at all.**

**Why no window is needed.** A window only existed to be sure a differing
colour was not a stale copy of the dashboard's own paint. But the Excel API —
the channel the dashboard writes through — reads the *live* file with no lag.
So the download is used to *find* candidates and the API to *confirm* them:

**The rule.** The dashboard records, per managed cell, the colour it last
painted (`PAINTED[job][item]`; persisted). On every download, for every
managed checkpoint cell:

1. **file colour ≠ last-painted colour** → candidate. (Rare: a hand-paint or a
   stale copy.)
2. Read **that cell's fill through the Excel API** (`range/format/fill`,
   batched, one read per candidate). Instant, and never stale.
3. **API colour == download colour** → a real external change → adopt it now:
   `listUpsert` the row with `Status` from the colour, `Done` from its meaning
   (done = total, blank = 0, process = leave `Done` as it was),
   `Source = excel`, `Who` = the file's last editor, `When` = now; one
   `Dashboard Log` line: *"Adopted from Excel: R5033 TG → gold (file last
   edited by X, 14:02)"*. Nothing is painted — Excel already has the colour.
   `PAINTED` is updated so it is not adopted twice.
   **API colour == last-painted colour** → the download was stale → ignore,
   silently.
4. **The open drawer's job is checked live:** on every 10 s station poll, the
   drawer's job's managed cells are read through the API (a handful of reads
   in one batch), so a hand-paint on the job somebody is looking at shows
   within ~10 s without waiting for a download at all. Every other job waits
   for the download to notice it — checking every cell of every job through
   the API each poll would be thousands of reads.

**What that makes the real speed:** a hand-paint anywhere shows on the
dashboard and the tablet **~40–50 s** after it is made (download lag ~36 s +
up to one 12 s poll + the confirming read); on the job that is open in the
drawer, **~10 s**. The download is how the dashboard *notices*; nothing makes
it notice sooner, but nothing waits after it has.

**Honest limits, to be stated in the UI wording and the log:**
- Excel gives one *last modified by* for the whole file, not per cell. The log
  says *"file last edited by"*, never *"painted by"*.
- A cell hand-painted and then painted again by the dashboard inside the same
  download interval shows the dashboard's colour; the hand-paint is not
  adopted, because by the time it is noticed the API says the dashboard's
  colour. The office's action wins, which is the owner's own rule.
- The confirming API read costs one request per candidate; a workbook where
  many cells are hand-painted at once (a bulk colouring) produces one batch of
  up to 20 reads per 20 candidates — bounded, and only on the download that
  first shows them.

**Tests:** a hand-paint seen in the download and confirmed by the API is
adopted on that same load, with the right list row and log line; a stale
download (API says the dashboard's own colour) is ignored and nothing is
written; the dashboard's own paint is never adopted as external; the open
drawer's job picks up a hand-paint on the next station poll without a
download; a bulk of 25 candidates is confirmed in two batches and adopted
once each.

## 5. What is deleted

Everything that exists only to survive the lag, for checkpoints: the cp/gc
`PENDING` holds and their persistence, `scheduleReconcile`/`bootReconcile`
for checkpoints, hold expiry and give-up and its red toast, `officeSettling`
(the settle window), `glassOfficeStamp`/`glassOfficeJobStamp`/`OFFICE_FLOOR_AT`,
`glassCellNow` as an input to any decision, and the writer's `have`-vs-`want`
against the parsed file. Roughly a thousand lines and their tests. The
`PENDING` machinery stays for the things that still ride the download: row
moves, mark-ready, phases.

## 6. Hard rules that still hold

1. The tablet never touches the workbook. (It gains nothing here; it already
   reads and writes lists only.)
2. Only fills are ever written to `Production`; only DG/TG/TUFF/NOT TUFF from
   the floor; arch/astragal/fancy/extra by office click only.
3. The office never writes `Station people` or `Station log`; the two granted
   exceptions on the floor's columns stand exactly as written.
4. No real person's name, email or domain anywhere.
5. No list is created by code. The owner creates `Dashboard progress` (or runs
   a dry-run script); missing list → the drawer says so plainly and checkpoints
   are read-only until it exists.

## 7. Honest limits

- **Excel can be a few seconds behind the dashboard**, the other way round
  from today. Someone looking at Excel sees the colour ~1–3 s after the click.
  Nobody has ever complained about that direction.
- **Hand-painting a checkpoint cell in Excel no longer counts.** The office
  must be told once.
- **The list grows** (one row per job per item, ~30 per job). Rows for jobs
  that have left the sheet should be pruned by a later chore; not this spec.
- **The one-time import trusts today's Excel colours once.** Any cell that is
  wrong today becomes wrong in the list until somebody clicks it.

## 8. Steps and order

1. Owner creates the `Dashboard progress` list (five minutes in the browser,
   or the dry-run script). Same permissions the dashboard already has.
2. Build: list read/write + drawer on the list + import — **checkpoints only,
   glass writer untouched** — ship, live with it two days.
3. Build: glass writer writes the list, not the file; delete the hold/settle
   machinery — ship.
4. Later, when a new department comes: its stages write the same list.

## 9. Tests to deliver (step 2)

- A click updates the row at once and writes exactly one list upsert then one
  fill; a refresh or a second dashboard reads the same thing; no `PENDING`
  entry is ever created for a checkpoint.
- An un-tick then a refresh then a stale download: the row stays white,
  because nothing reads the download for status. (The harness in the
  scratchpad — `untick_repro_reload.js` / `stub_office_reload.js` — must pass
  this with the download lagged 36 s and the page reloaded, R1–R8 all white.)
- Missing list: read-only checkpoints and a plain message; nothing written.
- The import: a job with no rows adopts the Excel colour once, `Source =
  import`; a job with rows never consults the colour.
- The standing proofs: no workbook write but fills; tablet walled off;
  nothing personal.

## Amendments after review

_(append here; do not rewrite the sections above)_

---

### A. Step 2 as built (2026-09-11)

Built in the working tree, not committed. Every suite green, and the browser
harness `scratchpad/untick_repro_reload.js` + `stub_office_reload.js` run
R1-R8. What the code does that the sections above do not say, and where it
departs from them.

**A1. `j.cp` keeps the parsed colours under the same name.** Renaming it would
have touched `parser.js`, `export.js`, the colour writer and every fixture for
no gain. Instead there is exactly one accessor, `cpFileStatus(j, item)`, and
`cpStatus`/`itemState` no longer touch `j.cp` at all. Three readers of the
parsed colour survive, each named in the code: the phase pipeline's Cut green
(nothing writes that colour, so the list can never carry it), the one-time
import, and the safeguard. `glassCellNow` also still reads it, deliberately -
it asks what the CELL is showing, which is what the colour writer's
idempotence test needs.

**A2. `glassCounts` keeps one reading of the file: the Cut green.** A green
"cut" cell seeds the floor's `Cut` counter through `ST.officeSeed`, and the
record can never carry that word. Without the fallback that feature would have
gone silently.

**A3. The `Dashboard Progress` SHEET is still parsed, and never written.**
Section 4 says to read it once at import "and then never again". It is in fact
parsed out of the workbook that has already been downloaded on every load - no
extra request - because two readers still need it: the import's exact count
behind a yellow cell, and `glassOfficeStamp`'s Progress-`When` source, which
step 3 retires with the rest of the colour writer.

**A4. The office's stamp moved with the hold.** Section 5 deletes the `cp`
holds; `glassOfficeStamp`, `glassOfficeJobStamp` and `officeSettling` all read
one. They now take the office's instant from the record row's `When`, counting
**only rows whose `Source` is `office`** - an `import` row is dated at
switch-on and would otherwise read as the office having just touched every job
on the sheet. Same instant, better precision, and it now survives a reload and
reaches the desk next door. The `pend()` gc-void of 2026-09-11 morning became
an explicit `glassVoidPaint()` at the click.

**A5. Which verb writes the list.** `listUpsert` reads the whole list before
writing. This list will be the biggest in the tenant, so: a row whose id is
known -> one `listPatch`; a new row from a **click** -> `listUpsert` (dedupes
on Title, once per item ever); a new row from the **import** -> `listAdd`, a
plain POST. Two dashboards importing the same cell in the same minute would
leave two rows; `cpRowsFrom` resolves duplicates on `When` and the first office
click removes the loser.

**A6. A blank cell is never imported.** No row already means nothing done, so
importing every blank would be a row per checkpoint of every job on the sheet
and the sixty-a-load cap would never catch up. A blank cell coloured later
belongs to the safeguard (see B5).

**A7. A refused fill does not undo the record.** `cpWriteItem` writes the
record first for exactly this reason and marks the error `cpRecordStands`; the
office's decision stands and Excel is behind, which is the direction section 7
chose. A refused RECORD write is a click that did not take, and the row goes
back.

**A8. The colour writer gained one line.** It records what it paints in
`PAINTED`. Observed in the harness: without it the writer's own paint read as
somebody else's and was adopted with `Source = "excel"` and the file's last
editor as `Who`.

**A9. Group writes are sequential.** SharePoint offers no batch of list writes
through this API and `graph.js` has no helper for one.

**A10. Told to the owner, plainly.** The first load after switch-on imports
roughly 2,000-3,000 rows at sixty a load with a thirty-second follow-up: about
fifteen to twenty-five minutes of quiet background writing, once, ever. And
**step 2 must not ship on its own**: until step 3, the colour writer paints the
floor's work into the `Production` sheet without writing the record, so the
floor's glass would reach the sheet but not the drawer's own glass checkpoint
rows. Steps 2 and 3 ship together.

---

### B. Independent review, and the fixes (2026-09-11)

The reviewer re-ran the harness (948 samples, none reverting) and could not
break the one guarantee - that no parsed colour becomes status while the list
is present. Verdict: fit as the base for step 3 with these fixed first. All
nine are fixed in the working tree; the flip table is in the build report.

**B1 (M1). The switch-on window.** Nothing asked whether the import had
finished, so for fifteen to twenty-five minutes every consumer that used to
read the colour read an empty record: `jobPhase` collapsed every job to In
office, `glassCounts` -> `officeComplete` was false everywhere so the feeder
wrote `OfficeDone: "No"` across the floor's list and then `"Yes"` again as rows
landed (two writes per job; jobs unlocking and re-locking on the tablet),
`officeSeed` zeroed untouched rows, and a job marked ready printed "all
checkpoints are complete" over lines reading 0. **Fix:** a per-browser marker
(`cw_cpimported`) set the first time `cpImportPlan(ALL, cpStored, 1)` comes
back empty. While it is unset - and whenever the list cannot be read -
`itemState` answers **an item that has no row** from the sheet's own colour
(`cpFileState`), so every screen and every feeder write reads exactly what it
read before the switch-over. An item that HAS a row always answers from the
row, so an un-tick made inside the window is still absolute. Once set, the
fallback is unreachable.

*Deviation, deliberate:* the reviewer also asked for the feeder's
`OfficeDone`/seed derivation to be **suppressed** until the import completes.
With the fallback in place the feeder's answers are byte-identical to the old
build's, so suppressing would itself change the plan - jobs the office really
has finished would sit unlocked on the tablet for twenty minutes. The stated
goal ("the same feeder plan as before switch-on") is met by the fallback and is
asserted directly in `test_station.js`, including that a half-imported record
plans the same thing and sends no write.

**B2 (M2). The import could stall.** The thirty-second follow-up fired into
`if (cpImporting) return 0` and armed nothing, so a pass that outlasted its own
timer stopped the import until the next `load()`. **Fix:** the guard arms the
next timer before returning, and the in-body decision moved into a `finally`
with `owed` defaulting to true, so a throw re-arms as surely as a refusal.

**B3 (M3). No repair path for the Excel copy.** A refused fill correctly left
`PAINTED` alone - but the safeguard compares the FILE with `PAINTED`, and after
a refused fill those two agree, so the cell was never a candidate and the
record said done while Excel said white for ever. That is the failure class
this whole change exists to end. **Fix:** `cpRepaintRun`, once per load, caps
at twenty cells, compares the RECORD with `PAINTED` (both in memory, no read)
and paints the difference through the ordinary fill path on the job's own
`cpChain`. `PAINTED` only moves when a fill lands, so one success ends it and a
refusal is tried again on the next load, never in a loop inside one. The four
glass columns are excluded while step 3 is outstanding: the colour writer owns
those cells, and repainting the record over them would be two writers fighting
for one cell. **Step 3 must remove that exclusion.**

**B4 (M4). An open drawer was expensive.** The live check skipped the candidate
rule, so every managed item of the drawer's job was read through the Excel API
on every ten-second poll - about a hundred workbook operations a minute for as
long as the drawer stayed open, adopting nothing. **Fix:** the same candidate
rule on both paths (the download DISCOVERS, the API CONFIRMS) plus at most one
check every thirty seconds per drawer job. Asserted: sixty seconds with a
drawer open and nothing changed sends no request at all.

*The cost of that fix, for the owner:* section 4a promised a hand-paint on the
job somebody is looking at would show in about ten seconds without waiting for
a download. It now shows on the same ~40-50 s as everywhere else, because
nothing looks at the cell until the download suggests it has moved. Buying the
ten seconds back would cost roughly thirty workbook operations a minute per
open drawer. **The owner's call, not a session's.**

**B5 (M5). A blank cell painted later.** `cpAdoptCandidates` returned early
when there was no record, so a blank cell hand-painted after switch-on was
picked up by the IMPORT - `Source = "import"`, `Who = "the sheet"`, no API
confirmation and no per-cell log line. Section 4a assigns that case to the
safeguard. **Fix:** once the import has drained, a cell with no row and a
colour is a candidate like any other, confirmed through the API and recorded
`Source = "excel"` with its own "adopted from Excel" line.

**B6 (M6). A background write could open a consent window.** `cpSaveRow` falls
through to `listUpsert` when no id is known, and `listUpsert` calls
`listConsent()`, which is allowed to throw a popup at somebody - fine from a
click, never from a load or a ten-second timer. **Fix:** background callers
pass `quiet`, and a background write that would have to create a row writes
nothing and leaves the cell to the next click.

**B7 (M7). A group write that failed half way** threw a bare error and left the
job part-ticked with nothing said. **Fix:** `cpWriteGroup` records which rows
landed and the toast names both halves.

**B8 (M8). One generation per feed.** `stationResetFeed` bumped a shared
counter, so a full read of `Dashboard progress` threw away an in-flight delta
of the floor's board. Each feed now has its own.

**B9 (M9).** The harness was re-run on the fixed build with correct build
metadata.

**B10. Still outstanding, for step 3 to close.** The `PAINTED` exclusion of the
four glass columns in `cpRepaintRun`; `glassOfficeStamp`'s use of the frozen
`Dashboard Progress` sheet; the `gc` holds and everything section 5 lists;
and the import's summary log line, which carries a blank `Job` and is therefore
written to `Dashboard Log` but not read back into the Changes window.

---

### C. Step 3 as built (2026-09-11)

Built on top of step 2 in the same uncommitted tree. **Steps 2 and 3 ship
together**: step 2 alone leaves the floor's work reaching the `Production`
sheet without reaching the record, and the reviewer showed the import would
then swallow the writer's own paint as `Source = import`, `Who = "the sheet"`,
and drive the tablet's lock from it.

**C1. The writer writes the record, and Excel is painted from it.**
`glassColourRun` still watches the floor's list and still derives the colour
with `ST.glassColours`, but `glassColourWrite` now goes: one `cpSaveRow` per
DG/TG/TUFF/NOT TUFF item (`Status` from the colour, `Done` = the total for
gold / the row's existing count for yellow / 0 for blank, `Who` from the row's
`DoneBy`, `When` = `ST.floorStamp(g)` — **the floor's own action time, not
now** — `Source = "floor"`), then the fill, then `cpPainted`, then the
`Floor glass colours` log line. One checkpoint write path, and the floor's
work goes down it.

**C2. Last-writer-wins is two fields of one row.** `glassColourPlan` compares
`ST.floorStamp(g)` with the existing row's `When` **when its `Source` is
`office` or `excel`** — a person deciding something. A tie goes to the office.
An `import` row (the switch-over reading today's colours) and a `floor` row
(this feature's own earlier work) never block. Idempotence came with it: if the
record already says what the floor says, whoever wrote it, there is nothing to
do — and if the SHEET is behind the record, that is `cpRepaintRun`'s job.

**C3. Deleted.** Measured on the pre-step-3 file:

| gone | ~lines |
|---|---|
| `glassOfficeStamp`, `glassOfficeJobStamp`, `officeSettling`, `GLASS_SETTLING`, `glassLogStamps`, `OFFICE_FLOOR_AT` | 186 |
| the hold machinery: `HOLD_GIVEUP`, `holdForceAt`, `reconcileForHolds`, `holdStuck`, `holdGaveUp` and their essay | 83 |
| `applyPending`'s `gc` settle, expiry, give-up and overlay | 30 |
| `pend`'s `gc` branch and its comment | 22 |
| `glassVoidPaint` and its two call sites | 19 |
| `glassCellNow`, `pendEmpty`'s `gc` clause, the repaint's glass exclusion, the `OFFICE_FLOOR_AT` write in `clearFloorGlass` | 12 |
| `glassColourPlan` / `glassColourWrite` rewritten | 162 out, 150 in |

About **510 lines of mechanism**, of which roughly 90 are replaced by named
tombstone comments — every one of those mechanisms was written for a real and
expensive bug, and the next person needs to know they were answered rather
than forgotten. `test_glasscolour.js` lost 851 lines and gained 608.

**C4. What is left of `PENDING`, and why.** The three-minute optimistic hold
survives for the things that genuinely ride the downloaded workbook and have
nowhere else to live: a whole-row section move (`blk`), mark-ready (`done`), a
product status (`prods`), and a hand-set phase (`phase`, which rides the
`Dashboard phases` list). `bootReconcile` stays for exactly those.
`scheduleReconcile` stays — it asks for the file again after a write so the
parse catches up. The boot migration that strips leftover `cp` **and** `gc`
holds out of `cw_pending` stays for as long as a browser might still carry one.

**C5. One reading of the parsed colour survives in the writer, on purpose.** A
cell carrying a colour this feature does not own — the sheet's own Cut green —
is not painted over, in either direction (the colours spec, 2026-09-10). That
is a fact about the CELL, not about status, and it is the same exception
`cpFileStatus` already documents. `cpRepaintRun` keeps off such a cell too. An
office **click** still paints one, exactly as it always did.

**C6. The `Dashboard Progress` SHEET is no longer parsed at all** once the
one-time import has drained (`cpImportSettled()`), which is every load but the
first few, ever. Removing `glassOfficeStamp` left the switch-on window as its
only reader. It is still never written.

**C7. The import's summary log line now carries a job** (the first in the
plan), because `load()` skips a `Dashboard Log` row with no job when it reads
the sheet back — so the blank one was written and then never seen again in the
Changes window. B10 closed.

**C8. Deviations from the step-3 brief, and why.**

- The brief said to keep a backoff "via the same mechanism M3's repaint uses".
  M3's repaint has no backoff — it simply retries once per load. The writer's
  existing per-job backoff (1 min, 5, 15, then given up on, with the footer
  saying so) is kept instead: it is better protection against a workbook or a
  list that keeps refusing, it is what stops a write storm while the floor is
  tapping every ten seconds, and it has four passing tests. The failure it now
  brakes is the **record** write; a refused **fill** leaves the record standing
  and belongs to `cpRepaintRun`.
- The brief said to prefer deleting `glassCellNow`. Done — but the Cut-green
  guard above needed one reading of the parsed colour, which now goes through
  `cpFileStatus` like the other two.

**C9. What this closes, in the owner's own words.** The sequence in
`test_glasscolour.js` section 13 is the whole complaint, in order: the floor
finishes a job (record → Excel gold → the tablet locks); the office presses
Clear (record → Excel white → the floor's counters to nought → the tablet
unlocks); five stale downloads and a stale list read change nothing; and a
genuine floor tap a moment later puts the gold back. Not one of those four
steps holds anything, anywhere. Section 14 adds the two races that used to lose
it — a poll landing mid-clear, and one landing inside the stepper's 800 ms
debounce — and both are now decided by a row that was written at the click.

---

### D. Second review of steps 2 + 3, and the last fix pass (2026-09-11)

Verdict: **fit as the release candidate** — the guarantee holds without
exception, last-writer-wins is asserted in every case, the end-to-end test is
real, the deletion is clean, and there is no write-storm path. Seven findings,
all handled below. The tree is held uncommitted until the owner says go.

**D1 (F1, moderate). The repaint trusted the download.** `cpRepaintPlan` had a
second shortcut: if the ~36 s-old download already showed what the record said,
it concluded the sheet was right, wrote that into `PAINTED` and skipped. That
is B3's own failure class through another door, and the sequence is ordinary:
the office un-ticks (record `""`, the fill lands, `PAINTED ""`); the floor
finishes and the writer's **fill is refused** (record `"done"`, `PAINTED` still
`""`); the download in hand is still the pre-un-tick gold, so it *agrees* with
the record; the shortcut wrote `PAINTED = "done"` and the cell was never a
candidate again — Excel white, record done, for ever. **The shortcut is gone.**
The file is not consulted at all: a fill that turns out to have been
unnecessary is one idempotent write, and it only happens when the record and
what this dashboard painted already disagree.

*The cost, recorded honestly:* a browser whose `localStorage` has been cleared
has no `PAINTED` at all, so every managed cell with a record row is a candidate
once. Bounded at 20 a load and self-terminating (each landed fill sets
`PAINTED`), but on a full sheet that is a few hundred idempotent fills spread
over a couple of hours of background loads. Correctness over churn.

**D2 (F4, minor). A list read that has gone backwards.** `STATION_ITEMS` can
regress — a full read dispatched before a delta can land after it — and for a
moment this dashboard holds an older copy of the floor's row than the one it
has already recorded. `glassColourPlan` now refuses to write a floor row over a
`floor` row whose `When` is *newer* than the stamp in hand. The same instant is
not a regression and still writes.

**D3 (F5, minor). A floor colour whose fill failed left no trace.**
`noteChange` sat after the fill, so a refused fill threw first and the colour —
which had been **recorded** — never reached Changes; `cpRepaintRun` painted the
cell on a later load and said nothing either. **The log line moved to
immediately after the record write and before the fill**, rather than adding a
second line to the repaint. The reason: the record *is* the event and the
colour in the cell is its copy, so one line at the moment the record moved is
both the truthful account and the only one — two lines for one change would put
the repaint's plumbing into the office's history.

**D4 (F3). The rule-2 `[PENDING]` draft described the step-2 world** — "two
readers, the import and the glass colour writer's office stamp". `glassOfficeStamp`
is gone. The draft now says **one** reader, the one-time import, and not parsed
at all once it has drained.

**D5 (F7). The rule-3 `[PENDING]` draft still called `OfficeDone` "the office's
read-only lock".** It now means *"the record says this job's glass is done"* —
still a feeder column the tablet can never write, still greying the steppers,
but since the floor's work goes onto the record too, **the floor finishing a
job is what locks it**, and only an office row stamped later unlocks it. One
clause added to the draft.

**D6 (F6, recorded, not changed). `glassWaiting` brakes a whole job on one
cell's refusal** — if one of a job's four glass fills is refused, all four wait
out the 1/5/15-minute backoff. Deliberately conservative: the failures this
brake exists for (the workbook open exclusively in desktop Excel, SharePoint
refusing writes) are whole-file conditions rather than per-cell ones, and a
per-cell brake would be four times the state for a case that does not occur on
its own. Written down rather than changed; if a per-cell refusal is ever seen
in the wild, this is the note to come back to.

**D7 (F2). The harness was re-run on the final tree** after all of the above,
with honest build metadata (`steps 2 + 3 + review fixes F1-F7`). R1–R9, and the
table is in the build report.

**Flips proved for this pass** (each fix undone one at a time, suite run, file
restored): F1 → *"the cell is repainted: a stale download agreeing with the
record proves nothing"*; F4 → *"a record of the floor's work NEWER than the
list copy in hand is never written over"*; F5 → *"the floor's colour is in the
office's history even though the fill was refused"*.
