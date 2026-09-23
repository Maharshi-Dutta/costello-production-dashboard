# The running record: what was built, what broke, what we decided

Every feature, every bug that cost real time, and every decision the owner
made, in one file, so that a problem which comes back can be recognised
instead of re-solved.

**If something is going wrong, read section B first.** It is written from the
symptom — what a person actually sees — not from the code. Several of these
problems looked identical to each other and had completely different causes,
which is exactly why this file exists.

Keep it current at every ship: a new feature gets a line in section C, a new
owner decision gets an entry in section A, and anything that took more than
one attempt to fix gets an entry in section B.

Related files: [`REFERENCE.md`](REFERENCE.md) is how the system works today; the briefs in
[`specs/`](specs/README.md) are how each piece was designed; this file is the history and the
reasoning. [`CLAUDE.md`](../CLAUDE.md) in the repo root holds the hard rules that came out of
the decisions below.

---

## A. Decisions register

These are the owner's calls. They outrank anything in the code, and a session
that wants to change one has to ask. The date is when the owner said it.

| # | Date | Decision | Why |
|---|---|---|---|
| A1 | 2026-09-03 | The `Production` sheet is the source of truth. The dashboard reads a downloaded copy of the workbook and never becomes a second database. | The office already lives in Excel. A dashboard that disagrees with the sheet is worse than no dashboard. |
| A2 | 2026-09-04 | The `Production` sheet is written in exactly two ways: sanctioned single-cell fills, and whole-row moves between sections. Nothing else, ever. | One bad write to a live shared workbook costs a day of everyone's work. |
| A3 | 2026-09-07 | Dashboard-only state lives in dashboard-owned sheets, never in `Production`. | Keeps A2 true while still letting the dashboard remember things. |
| A4 | 2026-09-08 | Floor stations never touch the workbook at all. The office dashboard performs every write on their behalf. | The tablet is a shared device on the factory floor with no individual sign-in. It must not be able to damage the workbook. |
| A5 | 2026-09-08 | Interim: the floor's SharePoint lists live in the workbook's own site, not a separate `Floor stations` site. | The owner has no Global Administrator on the tenant and cannot create the site yet. Accepted knowingly; the station account can also open the workbook under this arrangement. |
| A6 | 2026-09-09 | No new columns on the job row. Per-job detail goes in the job card. | "I can click the job card and go in it to see the progress." The row is already wide enough to read at a glance. |
| A7 | 2026-09-09 | Run the test suites and check the feature actually works before every commit. | A commit that breaks the live page costs the office a morning. |
| A8 | 2026-09-09 | The John print template may carry phone numbers. Eircodes never, in any export. | The paper sheet is used to phone customers. Eircodes serve no purpose on it. |
| A9 | 2026-09-10 | Yellow on a glass cell means cut and hotmelted. Gold means glazed. Glazing gates gold for every glass column. | The floor's three stages needed two visible states in Excel, and the workbook had exactly one yellow glass cell before this, so yellow was free to redefine. |
| A10 | 2026-09-10 | The hierarchy is: Excel sheet, then the master dashboard, then the glass tablet. A change made from the dashboard is absolute. An un-tick from the dashboard happens with no thinking and no argument. | Five commits of increasingly careful "who spoke last" logic had failed. The owner cut it off: the office wins, full stop. |
| A11 | 2026-09-10 | An office un-tick of a job's glass also clears the floor's counters to nought, with a confirmation prompt if the floor had really tapped it. | An un-tick that leaves the floor's numbers intact just paints itself gold again on the next pass. Recorded as the second exception to repo rule 3. |
| A12 | 2026-09-10 | Only the office can unlock a job on the tablet. | The lock exists so the floor stops working on glass that is finished. The floor cannot be the one to decide that. |
| A13 | 2026-09-10 | The `Production` sheet is the main document. Do not edit its text. Changing it by colour is fine. Other sheets created in SharePoint by our scripts are fine to change. | Text is what the office reads and prints. Colour is what the dashboard was given permission to speak in. |
| A14 | 2026-09-11 | Checkpoint status lives in the `Dashboard progress` SharePoint list. That list is the record. The Excel colour is a copy written from it, never read back to decide status. | "The office is absolute; an un-tick must be as boring as a tick." No amount of hold logic can make a 36-second-old copy of a file tell the truth. See B14. |
| A15 | 2026-09-11 | A cell painted by hand in Excel is adopted into the list as fast as possible, with no waiting window, and who and what is logged. | People will keep colouring cells in Excel. Their work must not be silently overwritten by the dashboard. |
| A16 | open | Should an open job card poll the workbook every ~10 s for hand-painted cells, instead of the default ~40–50 s? Costs about 30 extra workbook operations per minute per open card. Default kept for now. | Owner's call. Nothing is broken either way; it only changes how fast a hand-paint is noticed. |
| A17 | 2026-09-11 | **The five DOORS DONE cells are the third sanctioned fill.** A job's doors are those five cells: the text is the door's type code, the fill is its status, and a door has two stages — yellow in fabrication, gold done. "When updated from dashboad it should update the cell with correct color as well. And vice versa." | The doors were the one part of a job the dashboard could not tick, so the office was colouring them in Excel by hand — which, since A14, does not count as status. Repo rule 1 gained its third reason in the commit that shipped it. |
| A18 | 2026-09-14 | **Windows (M) and Doors (N) keep their own tick AND answer for what is under them.** What is shown is the higher of the line's own record and the aggregate of its window types or its doors. The paint only ever goes **upwards**: a tick underneath may take M or N to yellow or gold, and **nothing under the line ever repaints it white** — a gold M or N is never touched by this feature, and only the office's own Clear on that line whitens it. | The doors brief's assumption (a) made the two lines purely derived; the owner withdrew it on sight of the build: *"already golden means it finished even if all the cell or component has no ticked. u should not change anything in the excel sheet. if yellow mean in fabrication and golden means done; if a window has some yellow and some component white it means it in process."* The review had also reproduced the derived-only version planting white into a gold M cell. |
| A19 | 2026-09-14 | **A doors quantity that disagrees with the coded cells is a warning, not a decision.** The drawer's Doors line and the home list's hover say "quantity says 3 · 2 doors listed"; nothing is blocked and no colour changes. | Owner: *"that mean the quantity is wrong and should give a warning."* The sheet is the master for how much work there is; the dashboard says when it does not add up and leaves it to a person. |
| A20 | 2026-09-16 | **The saved view named after a colleague in `app.js` keeps its name.** The rename to "Sheet order" was declined. | Standing item since the station-comments review flagged the rule-4 violation; the owner's call, not the session's, and the answer is no for now. |
| A21 | 2026-09-16 | **The unread-notes icon's seen state is per screen, not shared.** Two office computers each show 💬 on a job until each one, separately, opens that job. | Amendment E's own design: the state is a plain `localStorage` key, and nothing writes it to a shared list. Sharing it across screens is possible later but was not asked for. |
| A22 | 2026-09-21 | **Glazing is the last step of the whole job, not a glass step.** It gets its own station rather than folding into glass. | Owner: "glazing is the last step for the whole job, not just glass, so it will be its own dashboard with its section." |
| A23 | 2026-09-21 | The new glass colour rule (yellow = one of cutting/hotmelting done, gold = both) applies to every job at once, but only a row the floor or the office actually touches repaints. | The two existing gates — an untapped row, an office record newer than the floor's stamp — still hold a colour until somebody acts; the rule says what yellow and gold mean, not "repaint every row overnight." |
| A24 | 2026-09-21 | **Recorded data is not touched to land this ship.** `Glazed`/`GlazedBy`/`GlazedAt` stay on `Glass station` exactly as they are, read and written by nothing; old `glazed` log lines still show, labelled "Glazing." | "Do not change any data already recorded." |
| A25 | 2026-09-21 | **Tuff moves back inside "shown finished"**, reversing one clause of A9: still outside the glass total and the office lock, but a job still owing tuff is not drawn gold or sunk into Finished. | Owner, after the demo: "why are moving the jobs to complete when tuff is left? if job has tuff and is not done dont move it, if done then move." |
| A26 | 2026-09-21 | The weekly target is **one number — total sheets per week** — not one per glass type, not per day. | Simplest number the office can set and the cutter can read against. |
| A27 | 2026-09-21 | **The cutter cannot correct a saved day sheet. The office can**, from the dashboard. | The end-of-day sheet is the cutter's own record; only the office is trusted to amend it after the fact. |
| A28 | 2026-09-21 | **Glazing counts one number per job — windows plus doors —** read off the `Production` sheet alone, never the cross-sheet merge. | Standing rule since 2026-09-18: `Production (2)` and the other sheets are not real data. |
| A29 | 2026-09-21 | **Glazing complete → Ready to deliver will be one click in the office, not automatic.** Not built yet; the seam (`GLZC.finished`, the office board already knowing which jobs are complete) is left for it. | Owner's "yes" to the manager's recommendation over the automatic option: the row move is the heaviest write the dashboard makes (insert, copy, verify, delete) and the owner wants to see it before it fires. |
| A30 | 2026-09-21 | **No unlock of a locked job from any office board**, glazing included. | Owner: "no." Unlocking stays the job card's Clear, same as glass. |
| A31 | 2026-09-21 | Rehearsal finding R4941 ("cannot remove TG") raised and dropped, not fixed. | Owner's call after seeing it live. |
| A32 | 2026-09-22 | **The Glazing station board shows gold when the sheet already says the job is done** (a whole gold `Production` row, or a full glazed count), hides the steppers for the word "done" on that row, and drops the progress bar. | Owner, live on the board: a job whose row was already gold (4998, 5091) showed no colour, and the bar was noise — "I only need the number of doors and windows and colour as ready (completed, gold) and in progress (yellow)." |
| A34 | 2026-09-23 | **A remake counter on the welding tablet, beside All on every Frames and Sashes line.** How many times a frame or sash of that group had to be welded again. A record, not progress: it moves no done count, no colour, no "left", no report unit; no ceiling. Written by the tablet only; the office reads it (board line, drawer, report Jobs sheet) and never writes it. | Owner, 2026-09-23: "like if she needed to remake the frame or the sashes she will add a count like I had to remake that frame 3 times before it was completed". Office view read-only by the owner's answer. Built directly by the session at the owner's word, one Sonnet reviewer pass. |
| A33 | 2026-09-23 | **The glazing unit is a window, not windows plus doors.** A job's glazed total is its window count alone; a door-only job leaves the glazing station entirely (no card, no board row); card and board words say "windows" only. | Owner: "if a job has 11 windows and 2 doors the glazing should be 11 and the Total count should also be just all jobs (windows each job)." Doors are not glazed at this station. |

**Rules that came out of these:** repo rule 1 (two sanctioned fill reasons; a
third is the owner's decision, not a session's), rule 2 (dashboard-owned
sheets only; `Dashboard Progress` frozen since A14), rule 3 (lists hold what
the workbook cannot hold safely, with two named exceptions), rule 4 (no real
names in the public repo).

---

## B. Problems and fixes

Written from the symptom. Each entry says what was actually wrong, because in
almost every case the obvious explanation was not it.

### B1. "I made a change and it disappeared when I refreshed"
**Cause:** the dashboard re-read the workbook before SharePoint had finished
saving, so the fresh read was older than the change.
**Fix:** hold your own write locally until the downloaded file agrees with it
(`69cf61d`, `b590000`).
**Still relevant:** this is the origin of the whole `PENDING` hold idea, which
went on to cause B10 through B14. Since A14 the checkpoint path no longer
needs it.

### B2. "The write worked but the screen never changed"
**Cause:** the write path and the render path did not share a model; the log
also recorded dates wrongly.
**Fix:** `aec8869`. Render from one model, updated at the moment of the action.

### B3. Sign-in problems, three separate ones
- The overlay would not close after a successful login (`27778d5`).
- The redirect URI needed to be the origin alone, nothing appended (`79b5de8`).
- Sign-in failed entirely for anyone who had not granted the SharePoint list
  permission (`63afa1f`), and then needed to become a quiet explained state
  rather than an error (`d285bf7`).

**Lesson, still true:** never add a scope to the sign-in request. Ask for it
only on the requests that need it, in the background, and only show a popup
from a real click.
**If it happens on localhost:** `http://localhost:8000` must be registered as
a SPA redirect URI in Entra, or sign-in fails with `user_cancelled`.

### B4. The workbook session times out mid-use
**Cause:** Excel API sessions expire.
**Fix:** detect it and re-establish automatically (`65e521c`).

### B5. Row moves left broken borders
**Cause:** Excel keeps one border definition per shared edge between two rows.
Writing the landing row's top border silently rewrites the bottom border of
the row above it.
**Fix:** write the landing row's bottom and vertical borders only, never its
top, and re-assert the row above (`4c60c31`).

### B6. Job numbers and dates got mangled on write
**Cause:** a `values` PATCH auto-parses text that looks like a date or a
number, so "07/04" became a date and "9434" became a number.
**Fix:** guard text with a leading apostrophe.

### B7. Bulk writes failed under load
**Cause:** flooding `$batch` earns `OperationQueueFull` from Graph.
**Fix:** cap concurrency.
**Related, fixed later:** a `$batch` reply that is missing one of its
responses was being treated as success (`c8c7184`). It is a failure now.

### B8. A test suite was failing and nobody noticed
**Cause:** the verification command ended in `| tail -1`, which returns the
exit status of `tail`, not of the test.
**Fix:** always `set -o pipefail`. It is in the standard verification command
in `CLAUDE.md`.

### B9. A floor tap made during a write was lost
**Cause:** the write queue snapshotted its keys at the start of a flush, so
anything arriving mid-write was dropped.
**Fix:** drain until the queue is genuinely empty, and re-arm the retry
whenever anything is still owed.
**Related:** a delta result must be discarded if the list was replaced while
that delta was in flight. Guarded with a generation counter.

### B10. "The glass count says 157 and it should be 49"
**Cause:** the tablet header was summing every stage for every person into one
number, which is not a quantity that means anything.
**Fix:** two numbers per stage instead — "Cutting 20 left", "Hotmelting 35
left", with tuff counted separately (`a223716`). The per-person summed
functions were deleted outright.
**Lesson:** when the owner says a number is wrong, ask what the number should
mean before changing how it is computed.

### B11. The glass progress chip was clipped in the job row
**Cause:** it was put in the badge cell, which is too narrow for "Glass 48/4".
**Fix:** its own column (`adcad80`) — and then decision A6, no more columns
ever.

### B12. The feeder stopped sending jobs to the tablet
**Cause:** the feed is refused wholesale if a column it writes is missing from
the `Glass station` list. `TuffTotal` and `OfficeDone` did not exist yet.
**Fix:** the owner added five columns (`TuffTotal`, `OfficeDone`, `Tuff`,
`TuffBy`, `TuffAt`).
**Watch for this:** any new feeder column means the owner has to add it to the
list *first*, or every job stops reaching the floor. Ask; never assume it is
done.

### B13. A real person's name was committed to a public repo
**Cause:** three specs used a floor worker's real first name.
**Fix:** replaced with "the cutter". One occurrence survives in the message of
commit `ab0c85c`; the owner chose to leave it rather than rewrite history.
**Rule:** repo rule 4. No real name, email address or company domain anywhere
in the public repo, including commit messages. Placeholders are "the admin",
"the colleague", "Person A", `example.test`.

### B14. The big one: "I un-tick the glass and it goes back to gold"

This took **six commits over two days** and four wrong diagnoses. Every fix
was a real defect. Only the last one was the defect the owner was seeing.

| Attempt | Theory | Commit | Was it the cause? |
|---|---|---|---|
| 1 | The office was stamped when the write landed, not when the button was clicked | `a223716` | No |
| 2 | Stamps were per column, and the office's own clear out-ranked the office | `1f2f27d` | No |
| 3 | The colour writer needed to stand down while the office's change settled | `a6b31c6` | No |
| 4 | The clear did not release the tablet lock | `9aa6b96` | No, but a genuine bug |
| 5 | A page reload killed the reconcile timer, so a hold expired into a stale gold parse | `eea3ec6` | Partly, and a genuine bug |
| 6 | `applyPending` only *masked* the colour writer's un-landed paint, never *dropped* it | `c8c7184` | **Yes** |

**The actual mechanism (attempt 6).** At boot the colour writer legitimately
painted a job gold, because the floor's counters said finished and the
office's un-tick from the previous day had been made by hand in Excel, so
there was no dated office record to beat it. The feeder then locked the
tablet. The office pressed Clear 33 seconds later. The download caught up and
released the office's hold — and the older, still-un-landed gold paint
reasserted itself for 11 minutes 40 seconds, long enough to re-lock the
tablet. The fix is four lines: an office click on a cell now discards the
writer's un-landed paint for that cell instead of merely out-ranking it.

**Why it was so hard.** The dashboard reads copies. The workbook download runs
about 36 seconds behind the Excel API, and the list runs a poll behind. Inside
that window, a stale copy of the office's own change is indistinguishable from
the floor genuinely disagreeing. No amount of care about *who spoke last* can
fix that, because the question itself is unanswerable from a stale copy.

**What finally fixed it properly:** decision A14. Status moved out of the
workbook into the `Dashboard progress` list, and the Excel colour became a
copy that is written and never read back. About 510 lines of hold machinery
were deleted (`a7aa370`, `2f28da0`). The class of bug is gone, not patched.

**The lesson, and it is the most expensive one in this file:** four fixes were
shipped by reading the code and reasoning about it. None was right. The cause
was found in twenty minutes by a stubbed browser replaying the owner's exact
clicks with the download deliberately lagged 36 seconds *and a page reload
after the un-tick*, which the owner did every time because the tablet looked
locked. **For any report of the form "I did X and it came back" or "it took a
minute", reproduce it before theorising.** Ask the owner the exact control
they pressed and whether they refreshed. The harness lives in the session
scratchpad, not the repo: `untick_repro_reload.js` with `stub_office_reload.js`
for the reverting case, `morning_repro.js` with `stub_morning.js` for the
boot-time case.

### B18. The welding board blew the stack the first time it was opened in a browser

**Symptom:** `Maximum call stack size exceeded` the moment Show ▸ Welding
station was picked. Nothing on screen, no board, no error the user could read.

**Not what it looked like.** It was not the renderer, not the list, not the
data. `weldLogReadIfNeeded(then)` answered its callback **synchronously** when
the list had already been read — and the callback it is given, from inside
`weldBoardHtml()`, is "draw the board again". So `renderRows()` called
`renderRows()` until the stack ran out.

**The fix, and the lesson:** already read means `return`, not "call the
callback anyway". The glass pair (`stationReadIfNeeded` /
`stationLogReadIfNeeded`) has always returned bare and says so in a comment;
the welding three were written from the same shape but with the guard's `then`
left in. All three now return bare.

Worth recording for one reason: **no offline suite could have found it.** The
Node harness never calls `renderRows()` for the board, and 40 offline checks
were green while the page was unopenable. It was the first thing the real
browser check found, and it is why the manager loop has that step.

### B19. The welding fix pass: what an independent review found, and why each one mattered

Twelve findings on the first build of the welding station, all closed in one
pass on 2026-09-16. Four are worth remembering because none of them was a
mistake in the feature — each was a correct-looking piece of code that was
wrong about something outside itself.

**A pin that was not a pin (ship-blocking).** `stationSite("own")` returned the
cached site and otherwise fell through to the legacy resolver — which *prefers*
`Floor stations`. It reads like a pin and behaves like one right up until the
cache empties, and three ordinary events empty it: a 404 on any glass list call
(which calls `forgetStationSite`), localStorage cleared or a new device, and a
browser profile that has never opened the dashboard. After any of them the next
resolve moved the glass feeder, the glass colour writer and `clearFloorGlass`
onto a site with no `Glass station` list — permanently, from one 404, with
"Try again" re-resolving the same wrong site. It is a lookup by path per
channel now, with no route from one channel to the other. **The lesson: a cache
with a fallback is not a pin. If the fallback can be reached, it will be.**

**A regex that only knew one way to write a phone number (ship-blocking, rule
3).** `\d{6,}` catches `0871234567` and nothing else, and a phone number in a
workshop comment is almost never typed that way. `086 123 4567`,
`+353 86 123 4567`, `087-123-4567`, `086.123.4567`, `(086) 123 4567` and
`08712 34567` all reached the floor's list unchanged. Fixed with a separator-
aware pass in front of the contiguous one, and the same strip was extended to
the `Customer` column, which is free text off the same sheet. **The lesson: a
redaction rule is only as good as the list of shapes it was tested against, and
that list has to come from how people actually type, not from how the field is
defined.**

**One station's poll inside another's.** `weldPoll()` was called from inside
`stationPoll()`, after its `if (STATION_OK !== true) return false`. A glass list
that could not be read — no permission yet, list not made, one 404 — silently
froze the welding board on a dashboard where the welding lists were perfectly
healthy. Each station's poll is its own `try` on the tick now, side by side.
**The lesson: "in its own try" is not enough if the try is inside somebody
else's early return.**

**A ten-second-old board writing an absolute number.** The office's stepper
derived its value from `WELD_ITEMS`, which is up to one poll old. A welder's
"All" at 14:00:01 on a board last polled at 13:59:56 showing 3 of 49, plus the
office pressing + at 14:00:05, wrote 4 — forty-six taps destroyed, under the
office's later stamp so nothing could argue them back. The row is read
immediately before the PATCH now and the number derived from what it says, with
a quiet "updated from the floor first" when it had moved. **The lesson: any
absolute write derived from a polled copy needs a read immediately before it,
however short the window looks.**

Also closed in the same pass, and each one small: a repaint owed on the welding
board was never taken (`stationCatchUp` called `redrawStation`, which returns
early there without clearing the flag); the three `weld*ReadIfNeeded` helpers
retried a failed read at network speed because a redraw asked again immediately
(now on a clock, like the notes channel's); `weldJobCard` ignored `Active`, so
the drawer's "Welding 9 / 16" counted groups that had left the sheet; and
`weldSlice` could in principle emit one `Title` twice, which the unique rule on
the list would then refuse on every run for ever (now first-one-wins with a
warning).

**One inherited bug was deliberately NOT fixed.** In `welding.js`'s
`flushQueue`, `e` and `QUEUE[k]` were the same object, so the post-write check
`now.value === e.value` compared a thing with itself and always agreed: a tap
or a re-base that landed while the PATCH was in the air was deleted as though
it had been sent. `welding.js` now captures what it sent before the await and
decides against that. **`station.js` has the identical shape and was left
alone** — it is inherited rather than introduced, and changing it means
changing the frozen glass suite. It is written down here so the next person to
touch `station.js`'s queue fixes it there too.

### B20. "This job has four window types it has not got" — a station fed from the wrong sheet

**Symptom (owner, 2026-09-17).** Jobs on the welding tablet showing product
groups that are not on them. R5053 had `4000 CASEMENT`, `4000 TILT & TURN`,
`SIDELIGHTS` and `PVC SMART` on its card, none of which that job has, and its
`POLARIS 85MM CASEMENT` said 19 sashes where the sheet says 16. The
`Welding station` list had gone from 485 rows to 786 overnight.

**Not a welding bug at all.** Nothing in `welding-core.js` or the feeder had
changed. `parser.js` merges each product group's F/S/T **across every sheet**
by `Math.max` — one line, there since the beginning, and right for what it was
written for: the copies are meant to say the same thing and differ only by
being out of date, so the furthest-along number is the safe one.

**The cause.** The owner inserted a column into `Production` that morning.
`Production (2)` — John's print sheet — is **formulas that moved with it**,
while its own **headers did not move**. So on that sheet every number ended up
one column to the right of the header describing it: the parser read
`Production (2)`'s POLARIS sashes out of the transoms column, found a
`4000 CASEMENT` number where there is only a neighbour's, and `Math.max` took
whichever was larger. The merge did exactly what it was told. What broke was
the assumption underneath it — that two sheets' columns mean the same thing.

**The fix.** `parser.js` now also keeps `j.prodsMain`: the same counts taken
from the sheet named `Production` and merged with nothing. `weldSlice` reads
that, and a job with no `prodsMain` is not on `Production` and is not fed at
all. `j.prods` is deliberately unchanged — the drawer, the checkpoints, the
exports and the John print sheet all still read the merged one, and whether
they should is a separate decision for the owner rather than something to
change while fixing a live bug.

Same pass, same cause in miniature: `SUPER DOOR` is welded in sashes only, but
the sheet carries a frames number in that group's F column, so the floor was
being shown frames to weld that are somebody else's count.
`WELD_GROUP_PARTS` now says which components a named group actually has, and
the rest are fed as nought.

**What the correction costs.** Nothing is deleted, ever. On the real list the
next feed makes **0 adds, 629 patches, 0 deletes**: 322 rows go `Active = No`
(267 the merge invented, 55 jobs that have genuinely left the sheet) and the
rest are counts patched back. At 60 writes a run that is about eleven runs, so
the list settles over a few loads rather than in one.

**The lesson, and it is the general one.** *A station feeds from the main sheet
only.* The owner said so in the first brief for the welding station and it read
like a preference; it is a correctness rule. A copy of a sheet is only a copy
while somebody keeps it aligned, and nothing in a workbook tells you when that
has stopped being true. Any reader that must not inherit another sheet's
mistakes has to name the sheet it trusts — and `Math.max` across sources is a
silent way of trusting all of them.

**How it was found, and how long it should have taken.** The owner saw it on
the tablet. Reproducing it took one script against a fresh download
(`scratchpad/repro_r5053.js`): print `j.prods` for the job, then print what
each sheet says on its own. The difference was obvious in one screen. That
script is worth keeping the shape of for any "the numbers are wrong for one
job" report — parse, then ask each sheet separately, before reading any code.

### B15. Things that are not bugs and will be reported as bugs
- **An un-tick can show as gold on the master dashboard for up to a minute.**
  SharePoint's downloadable copy lags about 36 seconds. Not fixable, corrects
  itself.
- **Clearing one glass type of a two-type job does not clear the floor.** The
  floor has one combined DG plus TG number, so there is nothing to clear
  partially.
- **A tab closed part-way through a clear can leave the floor uncleared.**
  Tick done and clear again.
- **A tablet 404 on the `Floor stations` site lookup is expected.** That site
  does not exist; the fallback to the workbook's site is by design (A5).

### B16. A subagent wiped `app.js` to 0 bytes mid-task
**Cause:** a python heredoc run through the shell, mid-edit on `app.js`,
collapsed the literal `\uD83D` (the 💬 icon's surrogate half) to `\uD83D`
being written as an actual control sequence; `open(p, "w")` had already
truncated the file before the encode error landed, so the file was left at
zero bytes. Nothing had been committed, so `git` had nothing to restore it
from.
**Fix:** recovered `git show HEAD:app.js` as a base and replayed the previous
day's implementer's 28 Edit calls from its own transcript by hand, checked by
spot-checking line numbers recorded earlier that day, by the `355/8` numstat
matching the diff stat recorded before the wipe, and by all suites passing
afterwards.
**Lesson:** code edits by a subagent go through the Edit/Write tools, never a
shell heredoc — a heredoc can silently mistranslate escapes and a shell
redirect truncates before it errors. Uncommitted feature work of any size
should be committed to a branch early, precisely so a wipe like this one has
something to restore from.

### B17. Two review bugs in the unread-notes icon (amendment E), both closed before ship
**Cause 1 — stamps compared as text.** The stored "last seen" stamp and every
note's own `At` were compared as strings. The tablet writes milliseconds
(`toISOString()`); `"…10:00:00.100Z" < "…10:00:00Z"` as text, because `.`
sorts below `Z`. A note landing in the same second as the stored stamp could
therefore read as already seen, permanently.
**Fix 1:** one comparator, `ST.atCmp(a, b)` in `station-core.js` — parses both
sides as dates and compares the moments, falling back to the old text compare
only when a side will not parse (a hand-typed "yesterday"). Used everywhere a
stamp decides "seen" or note order.
**Cause 2 — an undated note born after the job was first opened read as seen
from birth.** The seen state was a single stamp per job; a note with no `At`
(typed by hand into SharePoint) has nothing to compare against a stamp, so it
was silently treated as older than whatever the stamp already held.
**Fix 2:** the stored value became `{ at, ids }` — a stamp for dated notes,
and the note's own item id remembered for every undated one, because a stamp
can say "and everything before this moment" but only an id can say "and that
one, too".
**Found by:** the second independent review, before ship, not by a live
report.

### B21. Review R1 — the new seed would have undone finished hotmelting on first load

**Would have been the symptom:** the first load after ship patches `Hotmelt`
back from a finished total to 0 on rows the floor had already done (7 live
rows on the saved copy, e.g. 43 → 0), handing the hotmelting tablet finished
work as work to do.
**Cause:** the new seed answers `{cut: total, hotmelt: 0}` for an untouched
row whose office record is yellow, but `feedPlan`'s untouched branch wrote
any difference between seed and list in **both** directions — so a
higher-than-seed value already on the row was lowered to match.
**Fix:** an untouched row's seed field is only ever raised, never lowered,
unless the whole seed is noughts (a real office un-tick).

### B22. Review R2 — the earlier office lock froze Tuff and dropped its queued taps

**Cause:** the office lock now lands at hotmelt-complete instead of after
glazing. It used to freeze the Tuff stepper too, and `dropBlocked()` deleted
any tuff taps queued while the cutter kept working.
**Fix:** `OfficeDone` locks the glass stages only — Tuff stays tappable on
the Cutting page and the office board under a locked job, and its queued
taps are never dropped by the lock.

### B23. D1 — the welding station report came back empty

**Cause:** `stationReport` filtered the log on `Stage === "weld"`, but the
welding tablet logs `Stage = frames | sashes`; the test fixture had been
hand-built with the wrong stage and agreed with the bug.
**Fix:** the station definition now says which log stages feed a report
stage (`reportLogStages`); the welding fixture is built through
`weldLogEntry` itself, never typed by hand.

### B24. D2/D4 — a day-sheet draft with no owner and no bedtime

**Cause:** the draft was keyed without the person, so `switchPerson()` could
load one person's half-typed sheet under another; and a draft that crossed
midnight had nothing telling it which day it belonged to or when to be
dropped.
**Fix:** the draft carries `who` and closes on `switchPerson()`/the idle
lock; it is filed under the day it was started and dropped at the end of the
**following** day.

### B25. D13 — the glass tablet scrolled sideways at 800px

**Cause:** a bare `1fr` grid track is `minmax(auto, 1fr)`, not
`minmax(0, 1fr)`; one `nowrap` span was enough to widen the track past the
viewport, and a leftover 700px two-column media query sat inside the
tablet's own 700–1100px band and fought it. A test had been asserting the
sideways scroll as correct.
**Fix:** explicit `minmax(0,1fr)` tracks, the stray media query removed.

### B26. G2 — a job kept reading "In glazing" long after it left the floor

**Cause:** a floor row outlives its job's section by design — nothing clears
it when the job moves on by hand — and the phase bar's floor voice had no
gate on that, so a job moved anywhere else kept reading whatever the floor
last said, forever.
**Fix:** the floor's voice is gated on the job still being in a live
section, the same test the tablets already use to decide what they show;
once a job leaves that scope the floor says nothing and `effectivePhase`
falls back to the sheet or a hand-set phase, as before this station existed.

### B27. Process — an implementer agent died mid-build on a usage limit

**Cause:** the agent hit its usage limit mid-task and stopped.
**Fix:** the working tree was snapshotted as a WIP commit before anything
else touched it, and the same agent resumed with its own context intact;
nothing was lost.
**Lesson:** commit WIP to the branch after every agent hand-back — done all
day on this ship, which is why nothing worse is recorded here.

### B28. Measured, not fixed — the idle job list repaints itself about three times a minute

**Symptom:** the owner's "flickering" — the master list visibly redraws with
nothing on it changed.
**Measured:** a test rig sending a real no-change delta shows the ~230-row
list rebuilding roughly three times a minute on its own.
**Cause:** not yet traced. Left open, section E.
**Update:** the owner reports it fixed on their side, 2026-09-22. Not
re-measured in the rig, so left open in section E until it is.

### B29. The glazing tablet's Sign in button did nothing

**Symptom:** since the glazing station shipped 2026-09-21, tapping Sign in on
`glazing.html` had no effect at all — no Microsoft popup, nothing.
**Cause:** `glazing.js` declared `const G = GLZC` at the top level; `graph.js`
already declares a global `const G` (the Graph base URL). Classic `<script>`
files share one global lexical scope, so `glazing.js` failed at parse time
with `Identifier 'G' has already been declared` — the whole file never ran,
so the Sign in button never got its click handler.
**Fix:** the alias renamed to `GZ` (`da89168`, build 20260922-1244).
**Why every gate missed it:** `node --check` is per file and cannot see a
name collide across files; the browser rig stubs `graph.js` with a fake `CW`
that has no `const G`, so the collision never had a chance to fire in a test.
Found by a headless probe of the live page — the welding page on the same
probe opened the Microsoft popup, glazing did not.
**New gate:** `test_pages.js` concatenates every page's own `<script src>`
files and parses them once as a single `vm.Script` (parse only, nothing
runs), so a duplicate top-level name across files now fails a suite instead
of only a live tablet. Failed on `glazing.html` before the rename, passes
4/4 after.
**Lesson:** stubbing `graph.js` away hides anything `graph.js` contributes to
the global scope.

### B30. The steppers track re-broke the alignment the same brief was meant to fix

**Symptom, caught in review before ship:** a row with the four stepper
buttons and a gold row showing only the word "done" landed at different x
positions on the Glazing station board — the exact misalignment the brief
existed to remove.
**Cause:** the steppers track was specified as `auto`. Each `.wohead` is its
own CSS grid, so an `auto` track sizes to that row's own content — about
165px for four buttons, about 32px for the word "done" — and the `fr` tracks
on either side absorbed the difference. The rig measured the section cell at
x=430 on one row and x=474 on the next.
**Fix:** a fixed 170px track (four buttons plus their three 5px gaps),
measured in Edge at the board's font size — the same lesson as the welding
board's fixed 34px notes column (`b5103ba`, 2026-09-18, build log C/2026-09-21).
**Lesson:** in a grid drawn once per row, `auto` is not safe for a track
whose content differs row to row; only a fixed track keeps two rows agreeing.

---

## C. Build log

In order. Each line is one shipped commit.

### 2026-09-03 — the dashboard exists
| Commit | What it added |
|---|---|
| `16a9a56` | Live dashboard over the SharePoint workbook |
| `eecb3fd` | MSAL and ExcelJS served from the repo instead of a CDN |
| `79b5de8` | Origin alone as the redirect URI (B3) |
| `27778d5` | Sign-in overlay closes after login (B3) |
| `2f817bf` | Show what actually changed, not just that something did |
| `d1180e6` | Shared audit log in a `Dashboard Log` sheet |
| `65e521c` | Automatic recovery from workbook session timeout (B4) |
| `aec8869` | Writes now reach the screen; log dates fixed (B2) |
| `b590000` | Your own edits show immediately (B1) |
| `c68b7c5` | Build id stamped and timings shown, so slowness can be measured |
| `69cf61d` | Your own change survives a refresh (B1) |
| `b5f5ff4` | Views, custom categories, grouping, drag and drop |
| `6cacdc9` | Sort by job number, ignoring the letter prefix |
| `b27c883` | "Not sent to floor" renamed "Ready to fit" |
| `f8bd722` | Per-job comments restored after the Graph rebuild dropped them |

### 2026-09-04 — history and checkpoints
| Commit | What it added |
|---|---|
| `0a6381e` | Versions window: browse SharePoint history, see diffs, roll back |
| `4c60c31` | Section moves move the Excel row; mark-ready moves it too (B5) |
| `c93d2b0` | Checkpoints: tick windows, doors, glass and F/S/T off per job |

### 2026-09-07 — alerts, export, phases
| Commit | What it added |
|---|---|
| `56bbab5` | Job alerts: subscribe an address, mailed every third day at 08:00 |
| `f6fff87` | Export chosen jobs as an Excel workbook or a PDF |
| `62d9d53` | Phase pipeline as the primary progress, plus a selection wheel |
| `d40927d` | Selection wheel becomes a full-circle radial menu |
| `96c47dd` | Wheel motion reworked; hand-set phases kept in a SharePoint list |
| `63afa1f` | Sign-in no longer depends on the list permission (B3) |
| `d285bf7` | An ungranted permission is a quiet explained state (B3) |

### 2026-09-08 — the floor gets a tablet
| Commit | What it added |
|---|---|
| `616daee` | Glass station: a floor page for the glass area, fed from the office |
| `6340439` | One number of glasses per job; cards tapped on the spot |
| `1115945` | User Guide and Support and Maintenance, Markdown and PDF |

### 2026-09-09 — the print sheet, and the floor reaches the office
| Commit | What it added |
|---|---|
| `bf3470b` | John print sheet: a Production (2) view, a print template with notes, section select, row colour flags |
| `adcad80` | The floor's glass progress on the master dashboard's job rows (B11, then A6) |
| `3d26677` | Glass total beside the tablet's search box |
| `06dff7c` | The tablet counts down what the signed-in person has left |

### 2026-09-10 — glass colours, and the un-tick saga
| Commit | What it added |
|---|---|
| `ab0c85c` | **Glass colours: the floor's work reaches the `Production` sheet.** DG, TG, TUFF and NOT TUFF painted from the floor's counters (A9) |
| `a223716` | Per-stage counts on the tablet (B10); office stamped at the click (B14 attempt 1) |
| `c9bbee5` | An office un-tick clears the floor's counters (A11) |
| `1f2f27d` | The office's own clear no longer out-ranks the office (B14 attempt 2) |
| `a6b31c6` | The writer stands down while an office change settles (B14 attempt 3) |
| `9aa6b96` | An office clear unlocks the job in the same write (B14 attempt 4) |
| `eea3ec6` | A held change is protected until the file catches up, reload or not (B14 attempt 5) |

### 2026-09-11 — the rebuild
| Commit | What it added |
|---|---|
| `c8c7184` | **An office click voids the writer's un-landed paint of that cell** — the actual cause of B14 |
| `a7aa370` | Checkpoint status lives in a list; the Excel colour is a copy (built, reviewed, held) |
| `2f28da0` | The rebuild shipped: rules in force, build `20260911-1642` (A14, A15) |

### 2026-09-14 — doors, and window types under one heading
| Commit | What it added |
|---|---|
| _(this build)_ | **Doors by type.** The five DOORS DONE cells (BT–BX) become one checkpoint each — the cell's text is the door's code, its fill is its status, two stages only — ticked from the job card with In fabrication / Done / Clear, joined to the `Dashboard progress` record like every other checkpoint, and adopted by the safeguard when somebody paints one in Excel (A17). **Windows (M) and Doors (N) keep their own tick and also answer for what is under them** — the higher of the two is shown, and the paint only ever goes upwards, so a gold M or N is never repainted white by anything under it (A18). A doors quantity that disagrees with the coded cells is said in plain words in the card and in the hover (A19). The door cells get a one-time import of their own, behind `cw_cpimported_doors`, because the general one had drained three days before they existed. The window types move under one collapsible **Windows** heading in the card, the doors under **Doors**, both remembered per browser. On the home list the WND / DRS numbers carry the aggregate colour with the breakdown in the hover — no new column (A6). **Repo rule 1 gains its third sanctioned fill in this commit.** Spec: [`specs/2026-09-11-doors-and-window-types.md`](specs/2026-09-11-doors-and-window-types.md). Tests: `test_doors.js`, plus doors and derived-aggregate sections in `test_checkpoints.js` |

### 2026-09-16 — station comments, and the floor's notes on the job row
commit and build: see git log

| Commit | What it added |
|---|---|
| `the commit stamped build 20260916-1259` | **Station comments**, built 2026-09-15: a note channel from any floor station to the office, `ST.stationComments(cfg)` built once for every station, the `Station comments` list, the tablet's own-station thread, and the drawer's cross-station read-only **Floor notes** section — a Changes line per new note, nothing written to the workbook (REFERENCE.md §20). **Amendment E**, added 2026-09-16 after the owner's demo: the Components F·S·T cell blanked (E1); an unread-notes 💬 badge on the job row and on the Glass station board card, per-screen seen state in `localStorage` (`cw_notesread`), opening a job scrolling the drawer to Floor notes (E2, B17); every station's notes drawn on the board cards themselves (E3) (REFERENCE.md §20a). Real-browser checks run this time: tablet 13/13, office 12/12, amendment 22/22, headed-Edge demo approved by the owner. Build: `20260916-1259` |

---

### 2026-09-16/17 — the welding station, and a shape for the next one

| Commit | What it added |
|---|---|
| 2026-09-18, build 20260918-1446 | **Welding tablet header: two capsules.** The single "N left" added frames and sashes together (2310 on the live list: 1213 frames + 1097 sashes over 155 in-production jobs); the owner asked for one capsule each. `weldLeftByPart` in `welding-core.js`; scope unchanged (active rows, In production, the cards shown). |
| 2026-09-18, build 20260918-0836 | **Welding reads the `Production` sheet only** (`bfbabe7`). The slice had used the parser's cross-sheet max; `Production (2)`'s headers had not moved with its formula-linked data after a column insert, so R5053 was fed four product groups it has not got. `parser.js` now keeps `prodsMain` (Production alone); welding reads that; Super door feeds sashes only. The next feed marks 322 invented/left rows `Active = No`. Owner's standing rule recorded: `Production (2)` is not real data (B20). |
| 2026-09-17, build 20260917-1231 | **Welding board colours follow the theme.** The row and group tints were hard-coded light green/yellow; in dark mode that was a pale box with light text, seen live by the owner. Six CSS lines in `index.html` now use `--green/--green-bg` and `--fab/--fab-bg`. |
| 2026-09-17, build 20260917-1212 | **The welding station**, the second floor page: `welding.html` + `welding.js` + `welding-core.js` for the PVC welders, a `Welding station` list in the `Floor stations` site (one row per job **and product group**), an office board at Show ▸ Welding station with the office's own steppers, and a read-only Welding line in the job drawer. F and S only, never T; every product group on the sheet minus a four-name deny-list; the COMMENT carried with a rule-3 strip over it. **No workbook write anywhere in the feature.** `station-core.js` was generalised by a **station definition** (default: glass, so every older call means what it meant) and the station-independent tablet shell moved into a new `station-ui.js`. `CLAUDE.md` rule 2 gains a dated third exception: the office's welding edits write the floor's counters, logged in `Dashboard Log` and never in `Station log`. Spec: `docs/specs/2026-09-16-welding-station.md`; what was built: `docs/REFERENCE.md` §21; how to add the third station: `docs/STATIONS.md`, "Adding a station" |

**Owner decisions taken that day** (full wording in the spec's "Decisions
taken"): the COMMENT reaches the floor with runs of six or more digits removed
(recorded as a rule-3 decision); a gold F/S cell seeds that component as
welded and a yellow one seeds nothing; the office may correct welding progress
on a finished job, "they might make it wrong"; office edits stay out of
`Station log` and the floor's log is shown on the office's board instead; the
sent-to-floor filter defaults off, the same as glass; and the **site flip** is
handled by pinning the glass page to the site its lists are already in rather
than moving them — `ST.GLASS.site`, one word, reversible on the day the glass
lists are copied across.

---

### 2026-09-21 — glazing leaves glass, the cutter's day sheet, station reports, the glazing station

Build `20260921-1858`. One push to `main`: ten branch commits landed as they
are, plus one ship commit, and `b5103ba` (the welding board's fixed 34px
notes track, built 2026-09-18) riding along unpushed until now. Three
briefs: [`specs/2026-09-21-glass-split-no-glazing.md`](specs/2026-09-21-glass-split-no-glazing.md)
(glazing out of the glass station, one tablet page per stage, the new
yellow-means-one/gold-means-both colour rule, an editable office board),
[`specs/2026-09-21-day-sheets-and-station-reports.md`](specs/2026-09-21-day-sheets-and-station-reports.md)
(the cutter's end-of-day sheet, a weekly target the office sets, one Excel
report template for every station), and
[`specs/2026-09-21-glazing-station.md`](specs/2026-09-21-glazing-station.md)
(the third floor station, sections A–E; section F — glazing complete →
Ready to deliver — briefed but not built). Three independent reviews, 5 + 13
+ 3 findings, all fixed before ship (B21–B26). Suites at ship: move 7,
checkpoints 70, alerts 30, export 54, phases 26, phases-list 35, station
270, glasscolour 72, john 31, doors 29, comments 58, welding 61, daysheets
54, glazing 53, digest 26/0. Real-browser harness: glass/office 61/61,
glazing 26/26. Rehearsed on a saved copy of the real `Glass station` list
(138 active rows, 116 never tapped by the floor): zero gold cells wiped,
zero counters lowered.

### 2026-09-22 — the glazing tablet could not sign in; the glazing board polished

| Commit | What it added |
|---|---|
| `da89168`, build 20260922-1244 | **Hotfix: the glazing tablet's Sign in did nothing** (B29) — `glazing.js`'s top-level `const G` collided with `graph.js`'s global `const G` in the shared classic-script scope; renamed to `GZ`. New gate `test_pages.js` (parses every page's concatenated scripts once, so a duplicate top-level name across files fails a suite; `node --check` alone cannot see it) — passes 4/4. |
| `9fa1528` → `afbe16b` (fix pass) → `697c920`, build 20260922-1309 | **Glazing board polish**, branch `glazing-polish` (`docs/specs/2026-09-22-glazing-board-polish.md`, A32). Office's Glazing station board only; the tablet untouched. The board gets its own CSS grid tracks instead of inheriting welding's (nine cells were sitting one track left). A job whose `Production` row is already gold, or whose glazed count is full, shows gold (`GLZC.glzOfficeColour`); its steppers are replaced with the word "done"; the progress bar is removed, the count and quantity words kept. Independent review found five things, all fixed in one pass: the steppers track had to be a fixed 170px, not `auto` (B30); the same fix applied to the phone layout; `byId` was an O(n) scan called twice per comparison inside the board's sort, now decorated once per card; "done" in `--ink-4` read about 3:1 contrast, moved to `--done`; a clamp-branch test added. Suites: pages 4/4, move 7, checkpoints 70, alerts 30, export 54, phases 26, phases-list 35, station 270, glasscolour 72, john 31, doors 29, comments 58, welding 61, glazing 54, digest 26/0. Browser rig 11/12 (the 12th needed a fixture row the saved list did not have). |

### 2026-09-23 — welding remake counter

| Commit | What it added |
|---|---|
| `6cbeac5` → `1eb7c61` → `d5e2ed7` (fix pass) → `6def102`, build 20260923-1115 | **Welding remake counter**, branch `welding-remake-counter` (`docs/specs/2026-09-23-welding-remake-counter.md`, A34). Two new Number columns on `Welding station`, `FramesRemade` / `SashesRemade`, **added by script after the push** (scratchpad `add_remake_columns.py`: `check`, `rehearse` on a throwaway list, `create`; reusing `make_floor_site.py`'s sign-in; the build went live ~40 min before the columns, so a remake tap in that window sat owed and landed after). The check-in had said "by hand"; the owner asked why, when the script existed — no good reason, script first next time. `welding-core.js`: counter keys `frames-remake` / `sashes-remake` beside `frames` / `sashes` (`WELD_COUNTER_KEYS`), every field table extended, `weldApplyTap` without a ceiling, `weldRebase` without a clamp, `weldRecord` carries `remade` per part and on each line, `weldCardSig` sees it, the report's Jobs sheet gains `Frames remade` / `Sashes remade`. `welding.js`: the pill `− REMADE N +` after All, through the same queue, log queue and re-base as a done tap; `welding.html`: 48 × 56 px buttons, wraps under below 520 px. `app.js`: read-only `↻ N` on the board line (its own 40 px track in `index.html`), `↻ N remade` on the drawer's Welding line, `weldOfficeEdit` refuses a remake key. Independent review (Sonnet, on the commits) found two real bugs, fixed in one pass: (1) the remake keys had been added to `reportLogStages`, and the station report's Summary/Days sum every line of those stages as "Units recorded" — five remake taps would have read as five frames welded; back to `frames`/`sashes`; (2) the remake PATCH stamped `DoneBy`/`DoneAt`, and `DoneAt` empty is the feeder's only gate for seeding a row's done counts from the office record — a remake on a fresh row would have closed it for ever; the PATCH is now the count alone, who/when on the `Station log` line. Suites: welding 63 (was 61), pages 4/4, rest unchanged; browser rig `remake_check.js` 23/23 (one PATCH with exactly one key, one log line `Stage = frames-remake`, done count and colour unmoved, 700–1100 px no overflow and one column, office row `↻ 3` with no button). Rule-2 grep 0, names 0, no workbook write. |

### 2026-09-23 — glazing counts windows only

| Commit | What it added |
|---|---|
| `762e4e1` → `9f60b8d` (fix pass) → `e3d671b`, build 20260923-0942 | **Glazing counts windows only**, branch `glazing-windows-only` (`docs/specs/2026-09-23-glazing-windows-only.md`, A33). Owner's decision: the glazing unit is a window, not windows plus doors — "if a job has 11 windows and 2 doors the glazing should be 11." `glazing-core.js`: `glzSlice`'s `total = wnd` (was `wnd + drs`); `glzQtyWords`/`glzUnitWords` say windows only ("N windows" on the tablet card head); the tablet's `cfacts` line, which had printed the same number twice, now carries the comment only; the station report column renamed `Doors (not glazed)`. List columns unchanged, `Doors` still fed as a fact. Independent review found no blockers, six minor issues and three nits, all fixed in one pass: an untested migration case, stale "windows plus doors" comments across `app.js`, `glazing.js`, `STATIONS.md`, `SUPPORT.md`, `2026-09-21-glazing-station.md` and the repo `CLAUDE.md`, the duplicated tablet number, and the report header. Live effect on the office's next load, through the existing feeder diff: `Total` PATCHed on every fed row with doors (~103 of 230 rows in the 2026-09-16 workbook copy), `Active = No` on door-only rows (~76); `Glazed`/`Done*` never written (a new migration test asserts the exact patch keys). One accepted behaviour recorded in the spec's Amendments: a door-only job In production whose glazer had recorded units goes inactive, so its "In glazing" phase voice withdraws and falls back to the sheet's or welding's. Suites: pages 4/4, glazing 56, station 270, welding 61, export 54, checkpoints 70, phases 26, daysheets 54, rest unchanged. Rule-2 grep clean, no workbook write. |

---

## D. Symptom index

| What you see | Read |
|---|---|
| A change vanishes on refresh | B1 |
| A write worked but the screen did not move | B2 |
| Cannot sign in, or sign-in says `user_cancelled` | B3 |
| Everything stops working after a while | B4 |
| Borders look wrong after a job moved section | B5 |
| A job number or date came out mangled in Excel | B6 |
| Bulk operations fail under load | B7 |
| Tests "pass" but something is broken | B8 |
| A tap on the tablet was lost | B9 |
| A count on the tablet looks far too big | B10 |
| Jobs stopped appearing on the tablet | B12 |
| A colour reverts, or takes a minute to settle | B14, then B15 |
| A job locked itself overnight | B14, attempt 6 |
| The tablet cannot find the `Floor stations` site | B15 |
| A job shows product groups or components it has not got | B20 |
| A count on a station board does not match the `Production` sheet | B20 |
| A station list grows by hundreds of rows overnight | B20 |
| A board is empty or frozen while another station's is fine | B19 |
| An office edit overwrote what the floor had just done | B19 |
| The glazing tablet's Sign in does nothing | B29 |
| Columns on a station board do not line up | B30 |

---

## E. What is still open

- **A16**, the only open decision: whether an open job card should check the
  workbook every 10 seconds for hand-painted cells.
- The Power Automate alerts flow, from `automation/SETUP.md`.
- The PDF logo file, `assets/logo.png`.
- The radial menu motion looked unchanged on the owner's device and was never
  resolved. Next step is to ask for the footer build number and their Windows
  "Animation effects" setting.
- Deferred by the owner: a request-and-approve gate for deletions, and a
  mobile layout for the master page.
- Parked features: DG and TG progress bars inside the job card, and a PDF
  export with a per-job template and custom notes.
- **A latent copy of the text-stamp bug (B17) remains in `logRows`'s sort for
  `Station log`** (`station-core.js` ~878) — left out of the 2026-09-16
  commit on purpose; only `commentRows`/`notesUnreadMap` were fixed via
  `ST.atCmp`.
- Shared seen state for the unread-notes icon across office screens, if the
  owner wants it later (A21 records the per-screen choice as deliberate for
  now).
- The owner still has to create the `Station comments` SharePoint list
  (recipe in `docs/STATIONS.md`) and do a live check on the tablet against
  the real list; nothing else is blocked on it, both pages already show a
  quiet explained state without it.
- The permanent `Floor stations` site, once the owner has Global Administrator
  (A5).
- Each glass tablet must be told its stage once — the `?stage=` bookmark or
  the on-page chooser.
- **Section F** (glazing complete → one-click Move to Ready to deliver) is
  briefed (A29) but not built.
- **The flicker (B28):** owner reports it fixed on their side, 2026-09-22;
  not re-measured in the rig, so left open here until it is.
- ~~`test_daysheets.js` fails at line 514 on `main` (2026-09-22)~~ — closed the
  same day, test fixture, not the app: the block seeds one saved sheet for
  Person A on this week's Monday, then saves a draft dated yesterday for
  Person A; every Tuesday those are the same day and "one sheet per person per
  day" rightly refuses. The D4 block now starts with nothing saved.
- The glass lists' move to `Floor stations` (`ST.GLASS.site` `"own"` →
  `"floor"`) still to do.
- **G4** (welding's and glazing's polls each delta their own copy of the
  shared `Station log`/`Station comments` tokens) and **G5** (a hand-set
  phase click on a floor-set step can't then be withdrawn by the floor) left
  by decision, not required for this ship.
- `#weldopen` is a dead button outside Edit mode.
- The "Floor log" chip opens the glass log on every board, not that board's
  own station.
