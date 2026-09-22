# The glazing station: the phase bar hears the floor

Continues [[glazing-station]] (§24).

### The phase bar hears the floor (section E)

`effectivePhase` in `checkpoints.js` gained a **third voice**. It was the higher
of the sheet's own evidence (`jobPhase`) and the hand-set phase (`phaseHook`,
the `Dashboard phases` list); it is now the highest of those two and the floor's
(`floorHook`, registered by `app.js` as `floorPhaseRec`):

- any welding recorded on the job → at least **In fabrication** (3);
- any glazing recorded → at least **In glazing** (4);
- glazing outranks welding — *"if something is marked in glazing it means it has
  gone through fabrication"*;
- the floor can only ever move a job **forward**, never past the sheet or a
  person;
- a counter tapped back to nought withdraws that voice and the phase falls back
  to the next highest;
- **and the floor only speaks for a job still in a live section** (review
  finding G2). A floor row is never deleted — it goes `Active = No` when its job
  leaves the sheet, and its counter stays exactly where the floor put it, on
  purpose. Section F was deliberately not built, so *nothing* clears a finished
  glazing row when the office moves the job on by hand. Without the gate, a job
  sitting in Ready to fit would read "In glazing" for ever — and it is the
  reading that outranks the sheet. `floorPhaseRec` therefore begins with
  `ST.inProduction(j, BLOCKNAMES)`, which is `ST.sectionInProduction` of the
  job's own section — the same test the tablets use to decide what they show —
  plus the `cat === "past"` guard, so a job gone from the sheet altogether is
  covered by the same line. It is asked of the **job** rather than of the list
  row because the sheet is a feed ahead of the row: the office moves the job,
  this load's parse sees it, and the floor falls silent on *that* load rather
  than on the next feed. Falling silent is not "not fed" and not "nothing done"
  — it is the floor having no opinion, so `effectivePhase` falls back to the
  sheet and any hand-set phase, exactly as before this station existed.

`floorPhaseOf(rec)` is pure and takes `{welded, glazed}`. `phaseSource(j)`
answers `"sheet" | "hand" | "welding" | "glazing"` — naming the sheet and the
person first on a tie, because the floor only ever raises — and is what lets the
drawer's strip say *"Moved here by the floor — from the glazing station"* the
way it already distinguishes hand-set from sheet.

**Nothing is stored for it.** No `Dashboard phases` row, no list write, no
workbook write: it is computed on every render out of lists already in memory,
through the two cached per-version maps (`weldRecordsNow().byJob`,
`glzRecordsNow().byJob`), so it costs two object lookups per job row rather than
two walks of a list. A whole section of `test_glazing.js` asserts **zero
requests of any kind** across every phase check.

**Why it cannot flicker.** A station that has read nothing says nothing, so the
phase starts at whatever the sheet says and is *raised* when the lists arrive.
The repaint goes down the existing quiet path rather than a new timer:
`chipsNow()` now carries each drawn row's floor phase, and **both**
`redrawWelding` and `redrawGlazing` call `floorPhaseRepaint()` when they are off
their own board (the welding half was review finding G1 — it returned bare, so a
welding-only job's badge waited for the glass list to move or for a workbook
reload). `stationAfterFeed()` asks both lists for their once-per-session read,
so a dashboard nobody has opened a floor board on still shows the right word —
and it now catches its own throws (G3), because it sits in one promise chain
ahead of `glassColourRun()` as `.then(stationAfterFeed, () => {})` and the link
after it carries an `onRejected`: a throw there would not merely lose a repaint,
it would be swallowed by the next link's handler and skip that load's colour
write entirely.

### Every caller of `effectivePhase`, and what each got

| caller | decision |
|---|---|
| `checkpoints.js` `phaseName(j)` | unchanged — it is `PHASES[effectivePhase(j)]` and inherits the floor |
| `app.js` `statusWord(j)` (the list badge and the drawer head) | unchanged — inherits the floor, which is the owner's whole request |
| `app.js` `phasePipeHtml(j)` (the drawer's strip) | inherits the floor, **plus** the new "Moved here by the floor" line from `phaseSource(j)` |
| `app.js` `setPhaseByHand(j, n)` (the `before` of the log line) | unchanged — "from" should read what the screen was showing, which now includes the floor |
| `jobPhase(j)` in `setPhaseByHand` / `phasePipeHtml` | deliberately **not** changed: it is "what the sheet alone says", the floor for a hand-set phase. The floor must not be able to disable a step in the picker |
| the export | **there is no phase in any export.** `EXPORT_FIELD_KEYS` has no phase, status or stage column (its `Section` is the sheet's section), so nothing an export carries can disagree with the screen. Asserted in the suite rather than assumed |
| the alerts digest (`automation/`) | reads no phase at all — checked by grep |

### Tests

`test_glazing.js`, 53 checks, in the offline pattern: the slice and the
Production-only quantities, rule 3 in three phone shapes and an eircode,
`feedPlan` with the glazing definition (adds, patches, deactivation, never a
floor column, never a delete), the board and the three colours, the tap clamp,
the five-field body and what the filter refuses, an offline queue replay across
a simulated reload, the re-base, the log line, the people gate, the office's
edit on the real `app.js` (one PATCH, one `Dashboard Log` line, no `Station
log`, double-click → one write, a stale board never overwriting the floor, a
refused write leaving the number alone), the feeder on the real `app.js`, a
missing list as a quiet state that stops no other station, the whole phase
section with zero requests, the station report with no glazing branch in
`export.js`, and the workbook gate over the three new files. The last section is
the three review findings (G1–G3), each written so it fails on the build that
went to review — they are kept apart from the rest because every one of them is
about a **join** between two separately-correct parts, which is the only kind of
fault a suite of pure functions can miss.

### The review, and what each finding changed

| # | finding | fix |
|---|---|---|
| **G1** | `redrawWelding()` returned bare off its own board, so section E's quiet repaint only ever reached the glazing list. A welding-only job's badge waited for the glass list to move or a workbook reload | one branch, `floorPhaseRepaint()`, the pair of the one `redrawGlazing()` already had |
| **G2** | a job that has left In production kept reading "In glazing" for ever — nothing clears a finished floor row, because section F was deliberately not built | `floorPhaseRec` gated on `ST.inProduction(j, BLOCKNAMES)`, for the welding record and the glazing record alike |
| **G3** | `stationAfterFeed()` had no try of its own and sits ahead of `glassColourRun()` in a chain whose next link swallows rejections | body moved to `stationAfterFeedBody()`, wrapped, `console.warn` on catch |
| G4 | `weldPoll`/`glzPoll` both delta the shared `Station log` and `Station comments` in the same site — a few redundant requests a minute per open screen | **not fixed, by decision.** Not a correctness problem; the fix if it ever shows up as load is one shared delta channel for those two lists |
| G5 | clicking the step the floor put a job on writes a hand-set phase the floor can then never withdraw | **not fixed, by decision.** Only reachable by a deliberate click in Edit mode; belongs to a future brief on the hand-set picker |
| G6 | a job with quantities on another sheet but nothing on `Production` would not be fed | **no fix needed**: checked against the workbook copy — 0 of 232 active jobs affected; the gap exists only among jobs already off the sheet, which are not fed |

### Not built here

- **Section F of the spec — glazing complete → Ready to deliver.** Moving a job
  to the Ready to deliver section is a mark-ready fill plus a whole-row move on
  the `Production` sheet, the heaviest write the dashboard has, and the owner
  has not yet chosen between doing it automatically and doing it on one click in
  the office. The seam is left and nothing more: `GLZC.glzColour`/`finished` say
  which jobs are complete and the office board already sorts them last.
- A day sheet or a weekly target for glazing (one `daySheets` line, when asked).
- Any change to glass or welding behaviour. Their suites kept every count.

## See also

- [[glazing-station]] — previous: what it is, the five things worth knowing
- [[checkpoints-and-phases]] — the phase pipeline this section extends
