# Glass colours: the hold, ordering, TUFF and the lock

Continues [[glass-colours]] (§17).

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
of the glass stages: `STAGES` / `STAGE_KEYS` mean the glass stages and
`ALL_STAGES` / `ALL_STAGE_KEYS` are all of them, so the job's `Total` is still
DG + TG and the feeder never seeds `Tuff`.

> **CHANGED 2026-09-21 by [[glass-two-stage]] (§22, owner, after the demo).** `finished` — the gold
> card, the Finished group, the job row's `Glass 8/16` chip — **does** ask
> about tuff now: a job that has tuff and has not counted it is not shown as
> finished. Only that clause moved; tuff is still outside the glass total, the
> lock and the header count. `ST.tuffOwed(g)` is the whole rule.

It gets a row and a number of its own (`stageLeft` against `TuffTotal`): eight
glasses to cut and eleven tuff show as **two** numbers, "Cutting 8 left" and
"Tuff 11 left", never the nineteen that adding them would have said. The stepper is drawn only on jobs that
have tuff on them. **NOT TUFF gets no counter at all** — its colour is derived
from the glass stages.

**The office's own yellow (`ST.officeSeed`, owner 2026-09-10).** A yellow glass
cell *means* whatever yellow means, so the seed reads it that way. **Since
2026-09-21 ([[glass-two-stage]], §22) yellow means one of the two glass stages is complete**, and
cutting comes first: when **every** one of a job's DG and TG items is at least
yellow, the floor's row starts with `Cut` at the total and **`Hotmelt` at
nought**, and
the cell round-trips to yellow instead of being flattened to blank. It had to
be fixed — a yellow item has no stored count, so the row used to seed at
nought and the first tap on that job painted the office's own mark out. It was
found on the single yellow glass cell in the owner's whole workbook. Hotmelt
staying at nought is the point: seeding both would make the cell read gold. A
gold item is untouched and still
round-trips to gold. **Every, not some**: the floor's row holds one combined
DG + TG number and there is nowhere to put a per-type split, so a job with DG
yellow and TG blank keeps the office's own count instead — seeding it to the
total would tell the floor that glass which still needs cutting is cut, and a
wrong instruction on the workshop screen is worse than a lost colour on a
report. That job can still lose its yellow once the floor taps; it is a known
limit, recorded in the spec's Amendment B5. None of this widens rule 3's
seeding exception — it changes how the counters are *derived*, not which
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

## See also

- [[glass-colours]] — previous: the two-way colour rule
- [[glass-office-clear]] — next: an office clear reaches the floor's counters
- [[glass-two-stage]] — the 2026-09-21 rule change (yellow/gold, tuff, one tablet per stage)
