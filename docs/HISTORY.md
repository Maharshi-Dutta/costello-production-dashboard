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

Related files: `REFERENCE.md` is how the system works today; the briefs in
`specs/` are how each piece was designed; this file is the history and the
reasoning. `CLAUDE.md` in the repo root holds the hard rules that came out of
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
| _(this build)_ | **Doors by type.** The five DOORS DONE cells (BT–BX) become one checkpoint each — the cell's text is the door's code, its fill is its status, two stages only — ticked from the job card with In fabrication / Done / Clear, joined to the `Dashboard progress` record like every other checkpoint, and adopted by the safeguard when somebody paints one in Excel (A17). **Windows (M) and Doors (N) keep their own tick and also answer for what is under them** — the higher of the two is shown, and the paint only ever goes upwards, so a gold M or N is never repainted white by anything under it (A18). A doors quantity that disagrees with the coded cells is said in plain words in the card and in the hover (A19). The door cells get a one-time import of their own, behind `cw_cpimported_doors`, because the general one had drained three days before they existed. The window types move under one collapsible **Windows** heading in the card, the doors under **Doors**, both remembered per browser. On the home list the WND / DRS numbers carry the aggregate colour with the breakdown in the hover — no new column (A6). **Repo rule 1 gains its third sanctioned fill in this commit.** Spec: `specs/2026-09-11-doors-and-window-types.md`. Tests: `test_doors.js`, plus doors and derived-aggregate sections in `test_checkpoints.js` |

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
- The permanent `Floor stations` site, once the owner has Global Administrator
  (A5).
