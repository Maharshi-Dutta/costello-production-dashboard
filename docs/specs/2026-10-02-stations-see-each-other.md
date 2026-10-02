# Stations see each other directly, faster ticks, no idle lock

Status: **draft for the owner, not approved, nothing built.** 2026-10-02.

## 1. Why this brief exists

Three sessions built the same thing three different ways in one week: one
station's tablet showing another station's progress.

| Feature | Branch | Route | Delay after a tap | State |
|---|---|---|---|---|
| Fabrication tablet sees glass status | `main` (build 20261002-0817) | the office copies a word into the `Glass` column of `Fabrication station` | about 1.5 to 2.5 minutes, and only while an office dashboard is open | live |
| Glass tablet sees fabrication progress | `glass-sees-fabrication` (2 commits) | the tablet reads `Fabrication station` itself | 10 seconds | built, reviewed once, not demoed, not pushed |
| Glazing tablet sees glass and fabrication status | `glazing-ready-chips` (4 commits) | the office copies two words into `Glass` and `Fabrication` columns of `Glazing station` | about 1 minute, office open | built both sides, reviewed, not pushed |

The owner's requirement (2026-10-02): every dashboard sees the others with no
noticeable lag, no missing data and no disagreeing numbers, and people are not
sent back to the name screen.

The office-copy route cannot meet that. A copied word is late by design, stops
when the office closes its browser, and is a second version of a fact that can
disagree with the first. The direct-read route already works on the glass
branch. This brief makes it the one route for all three.

## 2. What the owner has done (as understood, please correct)

- Two floor accounts. **The glass account** runs the 3 glass tablets. **The
  welding account** runs the other 10 (welding and every fabrication person).
- Both accounts are members of the `Floor stations` site.
- 2026-10-02: the welding account was made a **member of the workbook's site**
  (ProductionProgress), so its tablets can read `Glass station` there.
- The columns `Glass` (on `Fabrication station`) and `Glass` / `Fabrication`
  (on `Glazing station`) exist on the live lists.

Not known yet: which account the glazing tablet signs in with.

## 3. The design rule

**One owner per fact. Everyone who needs the fact reads the owner's list.
Nobody copies it.**

| Fact | Owner (the only place it is written) |
|---|---|
| glass cut / hotmelt / Tuff counts | `Glass station` (workbook's site) |
| frames / sashes / transoms / door glazing counts | `Fabrication station` (`Floor stations`) |
| welding counts | `Welding station` (`Floor stations`) |
| glazed windows / astragal | `Glazing station` (`Floor stations`) |
| job facts (customer, totals, section) | the `Production` sheet, fed by the office as today |

The glass lists stay where they are. No list is moved, no column is added, no
row is copied. The workbook is not touched by any tablet (rule 2 unchanged).

## 4. What each station will see

Status is **per job**, because the floor's counters are per job. The glass
count is one combined DG + TG number, not one number per window or per door.
The screens can say "this job's glass is ready"; they cannot say which single
window's glass is ready. That is a limit of the data, not of this design.

### 4.1 Fabrication tablet sees glass

Purpose: know which doors can be glazed and which windows are worth
fabricating next because their glass is ready.

- The existing glass chip and "Glass ready first" switch stay exactly as they
  look today. Only the source changes: the tablet reads `Glass station`
  directly instead of the office-fed `Glass` column.
- The chip gains the breakdown when glass is part-done:
  `Glass: cut 14/20 · hotmelt 9/20 · tuff 0/2`.
- "Done" keeps today's rule: cut and hotmelt full (or the office's
  glass-complete mark), and Tuff full when the job has Tuff.
- A Door glazing line shows the glass chip beside it, so the glazer of doors
  sees at once whether that job's glass is ready.

### 4.2 Glass tablets see fabrication, and hotmelt sees cutting

Purpose: know which job to do next.

- Fabrication line, "FABRICATION DONE, GLASS WAITING" badge, "N waiting on
  glass" capsule and its filter: taken as built on `glass-sees-fabrication`.
- **Hotmelt page shows cutting.** Since the 2026-09-21 split a card draws only
  the page's own stage, so the hotmelt tablet today shows no cutting count at
  all (and no Tuff). Restored as a read-only, greyed line on the hotmelt card:
  `Cut 14/20` and `Tuff 0/2` when the job has Tuff. Never tappable there.
- **"Cut, ready to hotmelt" first.** A switch on the hotmelt page, like
  fabrication's "Glass ready first": jobs whose cutting is ahead of their
  hotmelting come first. Off by default, remembered per tablet. Costs nothing:
  the cutting count is on the row the tablet already reads.

### 4.3 Glazing tablet sees glass and fabrication

Purpose: know which jobs can be glazed now.

- Chips, "READY TO GLAZE" badge and "Ready first" switch: taken as built on
  `glazing-ready-chips`.
- Source changes from the two office-fed columns to direct reads of
  `Glass station` and `Fabrication station`.
- Glass chip: `Glass ✓ done`, or `cut a/t · hotmelt b/t · tuff n/m`, or
  `not started`.
- Fabrication chip: see question Q3. The branch counts every group, doors
  included. The owner's words were "what windows have been fabricated", and
  glazing counts windows only since 2026-09-23. Proposed: **window groups
  only**; door groups are left out of the glazing chip.
- Ready = glass done, and fabrication done or nothing to fabricate.
- Unknown is never ready: until both lists have been read once in the session
  the chip says "checking".

### 4.4 Office boards

The office already reads all four lists. Its board chips are computed from
those reads with the same functions the tablets use. The three office-fed
columns stop being written. They stay on the lists, unused, as `Glazed*` did
after the glass split.

## 5. How it is built

1. **One shared reader.** The glass branch's `readFab` is generalised into one
   helper: given a site (`"own"` or `"floor"`), a list and a field set, it
   keeps that list in memory by delta read, one read in flight at a time, its
   own token, its own retry clock. Glass, glazing and fabrication use it; no
   second copy.
2. **One definition of each status.** "Glass done" (`fbGlassStatus`) and
   "fabrication done" (`glassFabOf`, `glzFabStatus`) are each defined once and
   used by every tablet and the office.
3. **Fail quiet, never guess.** A list that cannot be read (no permission, not
   found, throttled) shows "not available" or keeps the last good read. It is
   never read as "not started" or "ready".
4. **Order of work.** New branch from `main`. Merge `glass-sees-fabrication`
   and `glazing-ready-chips` into it, then change the sources. Both sessions
   hold their branches from the moment this brief is approved.

Files: `station-core.js`, `station.js`, `glazing-core.js`, `glazing.js`,
`fabrication-core.js`, `fabrication.js`, `welding.js` (tick only), `graph.js`
(throttling only), `app.js` (stop three column feeds, board chips, office
tick), the four test suites, docs.

## 6. Delays found in the audit, and what changes

Measured from the code on `main`, not from a live run.

| # | Delay today | Cause | Change |
|---|---|---|---|
| 1 | cross-station chips 1 to 2.5 min | office copy, 60 s debounce | direct read (sections 4 and 5) |
| 2 | tablet sees another tablet in up to 10 s | 10 s tick | **5 s tick; 2 s for 20 s after a tap or a change seen** |
| 3 | office board and sheet painting 60 s or more when no floor window is open | the office drops to a 60 s poll unless it is "watching" | always the fast poll |
| 4 | each office tick is slow | five list polls run one after another | run them together |
| 5 | fabrication tablet reads the whole assignments list every tick | full read | delta read |
| 6 | ticks can pile up when a reply is slow | no in-flight guard on the tablet tick | skip a tick while one is running |
| 7 | throttled calls retry on a fixed clock | `Retry-After` is ignored | wait what Microsoft asks, one shared back-off |
| 8 | a new job reaches tablets in 30 s to 1.5 min; a first feed of 460 rows takes about 5 min | feeders run in the office browser, one after another, 60 writes per pass | **not changed here**, see section 9 |
| 9 | an office dashboard left in a background tab slows to about one poll a minute | the browser throttles hidden tabs | not fixable in the page; keep the dashboard in its own window. Section 9 removes the dependence |

No tablet pauses or sleeps. Every tablet polls all the time (owner,
2026-10-02).

### Request budget

Microsoft limits each account to about 3,000 requests in 5 minutes (600 a
minute). To be confirmed against Microsoft's current page before the build.

| Tablet | Lists read per tick | At 5 s |
|---|---|---|
| glass | own board, fabrication | 24 / min |
| welding | own board | 12 / min |
| glazing | own board, glass, fabrication | 36 / min |
| fabrication | own board, assignments, glass | 36 / min |

Glass account, 3 tablets: about 72 a minute. Welding account, 10 tablets, the
worst case of all ten on fabrication: about 360 a minute. The 2 s bursts are
short and the limit is counted over 5 minutes, so both accounts stay inside
it. The demo rig will count real requests per account and the number goes in
the amendments.

## 7. No more return to the name screen

Today a tablet goes back to the person picker after 10 minutes without a tap,
search or tab change (scrolling does not count), and after a reload more than
10 minutes after the last touch. The person then picks a name and types a PIN.

Change: **the idle lock is removed on all four pages.** The person stays
selected until somebody presses Switch, across reloads and new builds. The
picker still appears when the person is removed from `Station people`, and
when a glass tablet's stage is changed.

What the owner gives up, plainly:

- Taps are logged under whoever was picked last. If person A walks away and
  person B taps without pressing Switch, the log and the day sheet say A.
- A PIN now protects only the moment of switching, not the tablet.
- On fabrication, eligibility follows the picked person: B gets A's lines
  until B switches.

To soften that, the header shows the person's name large, with Switch beside
it, on every page.

### The Microsoft sign-in

Separate from the name screen. The Microsoft session behind a browser app is
renewed silently, but the renewal token for this kind of app lasts 24 hours.
When silent renewal fails the tablet shows "sign-in has expired" with a
"Sign in again" button. Proposed: when silent renewal fails, the page tries
one automatic full-page renewal that needs no typing as long as the browser
still holds the Microsoft session ("Stay signed in" ticked). Only if that
fails does the button appear.

**Cannot be guaranteed:** how long Microsoft keeps a session is the tenant's
policy, not the page's. This has to be tried on a real tablet over two days.

## 8. Hard rules (unchanged)

- No tablet touches the workbook. The gate grep covers every station file.
- A tablet writes only its own station's floor columns. Cross-station reads
  are read-only; the shared reader has no write path.
- No phone numbers or eircodes; the glass customer name keeps its strip.
- No real name, address or domain in the repository.

## 9. Not in this brief (the next two steps)

1. **One `Jobs` list.** Job facts fed once instead of once per station: four
   feeders become one, no per-station caps, a sheet change reaches every
   station together.
2. **The feeder and the sheet painter out of the office browser**, on a
   scheduled flow. Then nothing on the floor depends on an office page being
   open. Needs the owner's licence checked.

Hand edits made directly in Excel still take about 40 to 60 seconds to show:
SharePoint's downloadable copy lags. Clicks made in the dashboard do not.

## 10. Tests to deliver

- Shared reader: first read, delta merge, one in flight, 403 / 404 / 410 /
  429 each leave the right state, "not available" is never "ready".
- Each view: the status words from list rows alone, unknown never ready, the
  hotmelt cutting line never tappable, cut-first order.
- Idle lock: no lock after any idle time, person survives a reload, Switch
  and removal from `Station people` still return to the picker.
- Tick: no overlapping ticks, burst starts and ends, `Retry-After` obeyed.
- Feeder: the three columns are in no PATCH (exact-keys test).
- Browser rig: four tablets and the office on shared fake lists; a tap on one
  shows on the others inside the tick; requests counted per account.

## 11. Process

Brief approved → implementers on the new branch → independent review of the
commits → one fix pass → suites, gate greps, rig → live clickable demo for the
owner → push only on the owner's go.

## 12. Questions for the owner

- **Q1.** Which account does the glazing tablet use? It needs read access to
  `Glass station` in the workbook's site, as the welding account now has.
- **Q2.** How are the 13 tablets split (glass 3; welding, glazing,
  fabrication how many each)?
- **Q3.** Glazing's fabrication chip: window groups only (proposed), or doors
  too?
- **Q4.** The welding account is a full member of the workbook's site, so
  anyone at those tablets can open and edit the workbook in a browser. It only
  needs to read there. Change it to Visitor (read only), or leave it?
- **Q5.** Removing the idle lock means taps log under the last picked name
  (section 7). Accepted?
- **Q6.** Do the tablets show "sign-in has expired" most mornings? If yes the
  automatic renewal in section 7 is worth building now; if no it can wait.
- **Q7.** Hotmelt "cut first": a switch that is off by default (proposed), or
  always on?
