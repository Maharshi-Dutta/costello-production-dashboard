# Glass two-stage: the office board, review findings, tests

Continues [[glass-two-stage]] (§22).

**D. The office's Glass station board is a working surface.** Modelled on the
welding board, and using the tablet's own two functions rather than a second
implementation: each stage line (Cutting, Hotmelting, and Tuff where the job
has any) carries `−`, `+` and `All`, clamped 0…total. One click is **one
PATCH** of five fields — `ST.floorOnly(ST.tapFields(stage, value, who, at))`,
so that stage's counter, that stage's `By`/`At`, and `DoneBy`/`DoneAt` — with
the row **re-read immediately before** the number is derived (the welding
board's reason: the board's copy is up to ten seconds old and the floor is
tapping the same counter). It leaves **one `Dashboard Log` line** through
`noteChange` — "Glass cutting" / "Glass hotmelting" / "Glass tuff", from → to —
and **no `Station log` line**: `app.js` calls `ST.logFields` nowhere at all. A
refused write puts the old number back and toasts. A job locked by `OfficeDone`
has its **glass** steppers disabled and those clicks refused, with the drawer's
own wording: *"glass complete — clear it in the job card"* — the **Tuff line
stays editable** (R2). Clicking the card head opens the
job drawer; a board job no longer in the workbook parse has no drawer and its
head is not clickable. `glassWriting[k]` goes up **before the first await** so
two fast clicks are one PATCH (R3). **Nothing on this path goes near the
workbook** — the sheet's colours follow on `glassColourRun`'s next pass, from
the new counters, by rule B.

**`web/CLAUDE.md` rule 3 gains a dated fourth exception** for that write.

### The five review findings, and what each changed

Independent review of the first build plus a rehearsal on a saved copy of the
real `Glass station` list (138 active rows, 116 never tapped by the floor).

- **R1 — the feeder must never lower a seeded counter.** `officeSeed` answers
  `{cut: total, hotmelt: 0}` for a record that used to seed both, and
  `feedPlan`'s untouched branch wrote every difference in both directions — so
  the first load after the ship would have PATCHed `Hotmelt` from the total
  back to nought on seven rows and handed the hotmelting tablet finished work
  as work to do. A seed field is now only ever **raised**, unless the whole
  seed is noughts (a real office un-tick, which still clears the row).
- **R2 — Tuff is outside the office lock.** `OfficeDone` is
  `ST.officeComplete` of the job's DG and TG: it says the office has ticked the
  **glass** off and has never said anything about tuff. That cost nothing while
  the lock landed after glazing, by which time the tuff was long counted; it
  now lands the moment hotmelting finishes, so a locked job is routinely one
  the cutter is still counting tuff on. The tuff stepper stays live on the
  tablet and on the office board, `tap()` and `glassOfficeEdit` both let it
  through, and `dropBlocked()` never takes a queued tuff tap away. The card's
  line reads "the office has marked this job's **glass** finished".
- **R3 — two fast clicks on an office stepper.** `glassWriting[k]` was set
  *after* `await CW.hasListConsent()`, so two clicks in one tick both passed
  the guard, both read the same row and both PATCHed from the same base. The
  flag (and its redraw) now go up before the first await, and every early
  return clears it. `weldOfficeEdit` had the same ordering and got the same fix.
- **R4 — a typo in `?stage=`** falls back to the device's remembered stage.
  Only a valid URL value wins; taking the typo as "told nothing" put a working
  tablet in front of the chooser and signed its person out.
- **R5 — a Tuff counter nobody has tapped is no opinion.** `Tuff` is the one
  counter the feeder may never seed, so on an untapped row it reads nought.
  Glazing used to paint TUFF gold and that nought said nothing; under the new
  rule it would read as "the floor says no tuff is done" and wipe a gold TUFF
  cell the office ticked by hand. `glassColours` answers `null` for it while
  `TuffAt` is empty and the plan names the column not at all. Tapping tuff down
  to nought on purpose still clears it.

**Told to the owner, not code:** the two existing gates in `glassColourPlan`
still apply to "every job at once" — a row the floor has never tapped, and a
job whose office record is newer than the floor's stamp, keep their colour
until somebody acts; and one office stepper click makes a row "touched" for
good, handing that job's glass cells to the counters (§D, by design).

**Tests.** `test_station.js` 246 → 266 and `test_glasscolour.js` 68 → 72;
`test_welding.js` unchanged at 61. Every glazing-gate assertion was converted
to its new-rule equivalent rather than deleted. New: the colour table re-written
to the count-of-complete-stages rule; a row carrying old `Glazed` values
planning blank; `finished` without glazing;
the seed both ways; a run-wide assertion that **no request body of either suite
ever carries `Glazed`, `GlazedBy` or `GlazedAt`**; the queued `glazed` tap
dropped on replay; the hotmelting page's people and steppers; the chooser and
what it remembers; a legacy log line reading "Glazing"; and the office board's
clamp, its five-field PATCH, its `Dashboard Log` line, its silence in
`Station log`, the `OfficeDone` lock and a failed write. And one per finding:
the seed refusing to go down and still clearing (R1); the tuff stepper and a
queued tuff tap surviving the lock on both screens (R2); two office clicks in
one tick being one PATCH and one log line (R3); `?stage=hotmlet` landing on the
hotmelting page with nobody signed out (R4); and a row with `Glazed = total`,
`Tuff = 0`, no `TuffAt` and a gold TUFF record planning no TUFF write (R5).

### Not built here

- A glazing station. The owner will brief it separately.
- The move of the glass lists to `Floor stations`: `ST.GLASS.site` is still
  `"own"` and that one word is still the whole of the move ([[welding-station-office-and-site]], §21).

## See also

- [[glass-two-stage]] — previous: glazing leaves the station, the new colour rule
- [[day-sheets-and-reports]] — next: the end-of-day sheet built on the cutting tablet
- [[glazing-station]] — the glazing station briefed separately, as promised here
