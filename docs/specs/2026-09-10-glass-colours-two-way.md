# Glass colours: the floor's work reaches the Production sheet

**Date:** 2026-09-10
**Status:** consulted, not yet approved to build
**Consulted:** yes — the owner described the colour logic on 2026-09-09/10 and
answered five rounds of questions. Every decision below is theirs and is dated.

---

## 1. Why this is different from everything before it

Every feature so far has kept the floor's work off the `Production` sheet. This
one deliberately puts it there: when the floor finishes cutting and hotmelting,
the job's glass cells go **yellow** in the workbook; when glazing is done they
go **gold**. That is a real change of policy and the owner made it explicitly
("excelfile too", 2026-09-10).

Two boundaries do **not** move:

- **The floor tablet still never touches the workbook.** It cannot — the
  station account's lack of workbook access is the actual security boundary.
  The **office dashboard** does every write, watching the floor's list and
  painting the cells. This is the same shape as the feeder, in reverse.
- **Only fills are ever written to `Production`.** No value, no text, no row,
  no formula. The owner restated this on 2026-09-10: *"the production sheet is
  main doc, don't edit that as in text; by colour is ok"*. Cell fills are
  already a sanctioned write (`CLAUDE.md` rule 1); this feature adds a new
  *reason* to make one, not a new *kind* of write.

## 2. The colour rule

| column | yellow when | gold when |
|---|---|---|
| DG (AY), TG (AZ) | cutting **and** hotmelting complete | glazing complete |
| TUFF (BA) | David's tuff count complete | glazing complete |
| NOT TUFF (BB) | cutting **and** hotmelting complete | glazing complete |
| arch, astragal, fancy, extra | **never written — office ticks these by hand, as today** (owner, 2026-09-10) |

Glazing is the gate for everything: nothing is gold until the job is glazed.

**Fully reversible** (owner, 2026-09-10 — *"do not make it one way because it
would be hard to reverse it when it done by mistake"*). Colours walk back down
as well as up: gold → yellow → blank, mirroring whatever the floor now says. A
job tapped back to zero leaves its glass cells with no fill at all.

Colours read and written use the constants already in `parser.js`: gold is
`FFE699` (also accepted: `FFC000`), yellow is `FFFF00`.

**Measured against the live workbook on 2026-09-10** (corrected after the
rehearsal — the figures first written here quoted a denominator of 704 filled
glass cells across 192 jobs, which matches nothing in the file): across the
eight glass columns of the **193** glass jobs in production there are **130
filled** glass cells — **129 gold** and exactly **one yellow** (one job's DG
column; the rehearsal knows which, and a live job number does not belong in a
public repository). The gold and yellow counts were right; the denominator was
not.

So the yellow convention is new — the office today only marks glass gold when
it is finished. Nobody's existing habit is broken by adopting yellow, but
nobody will recognise it either until they are told. And that single yellow
cell turned out to matter: see Amendment B5.

## 3. Two-way, and who wins

Both sides may change the same cell, so the rule is **the most recent action
wins**, decided on timestamps that both sides already carry:

- The floor's side is stamped in the `Glass station` list (`DoneAt`, and the
  per-stage `At` columns).
- The office's side is stamped in `Dashboard Progress` (`Who`/`When`) and in
  `Dashboard Log`, which every checkpoint tick already writes.

So: the office ticks DG gold by hand and it stays gold until the floor next
moves that job; the floor taps and the colour follows, until the office ticks
it again. This satisfies both of the owner's statements — *"the master
dashboard should not be overridden"* and *"any update from the glass area
should show in the job id"* — because neither side is silently overwritten by
an **older** action.

**The office keeps every control it has today.** The drawer's per-type ticks
(`−`, `+`, `All glass done`) are unchanged for every glass column. Owner,
2026-09-10: *"I like the option that the office dashboard already has and don't
change the options."*

### 3.1 The lock

When the office marks a job's glass **complete**, that job becomes
**read-only on the tablet** — the steppers grey out and `tap()` refuses them,
the same way a stage somebody does not hold is already refused. **Only the
office can unlock it** (owner, 2026-09-10), by un-marking it.

The tablet learns this from a new list column, not by reading the workbook.

## 4. The TUFF counter

TUFF becomes a fourth thing David counts on the tablet, with its own total from
the sheet's TUFF column (11 in the owner's example). *"Tuff was a different
department but same person operates it. David will update the status. It will
be the same as the cutting component David already has"* (owner, 2026-09-10).

- Its total is fed from the workbook's TUFF column, the same way the glass
  total is fed from DG + TG.
- It is **not** part of the DG+TG total and does not change the job's glass
  number. The floor's existing "N left" arithmetic counts it as its own stage
  the person may hold.
- NOT TUFF gets **no** counter — it is derived from the glass stages
  (owner: *"NOT TUFF doesn't need any component"*).

### 4.1 New columns on `Glass station`

`Tuff` (number), `TuffBy` (text), `TuffAt` (text), and `OfficeDone` (text
Yes/No, for the lock in §3.1).

The owner has approved a script for this — *"other sheets that you created in
SharePoint using scripts are ok to change"* (2026-09-10). **The script is
written for the owner to run; it is never run against the live site from a
session without the owner watching, and it is rehearsed on a copy first.** No
list is created by code (`CLAUDE.md` rule 3); adding columns to a list this
project already owns is what is approved here, and nothing else.

## 5. Hard rules

1. **The tablet never touches the workbook.** No new request from
   `station.js`/`station-core.js` to anything but the three lists.
2. **Only fills are written to `Production`.** Never a value, never a row,
   never any other sheet. Assert it in tests over the whole run.
3. Writes go through the existing `serialised()` machinery so two open office
   dashboards cannot interleave on one sheet.
4. Every write is held locally (`PENDING`) until the downloaded file agrees,
   as every other sanctioned write already is — the download lags the API by
   about 36 seconds.
5. No phone numbers or eircodes anywhere near this. No real person's name,
   email or domain in code, tests, docs or commit messages.
6. **Nothing runs against the live workbook until it has been rehearsed on a
   OneDrive copy and the owner has seen the result** (`CLAUDE.md` rule 7). This
   is the first feature where a floor tap changes the master sheet, so the
   rehearsal is not a formality.

## 6. Honest limits, told to the owner

- **Colours only move while an office dashboard is open.** Accepted by the
  owner on 2026-09-10. The floor working on a Saturday with nobody in the
  office means the sheet catches up on Monday.
- **A mis-tap on the floor now reaches the master sheet.** Reversible by
  tapping back, or by the office ticking over it — but it is no longer
  contained to the tablet.
- **Yellow will be unfamiliar.** One cell in the live sheet uses it today.

## 7. Tests to deliver

1. The colour rule, per column, in both directions including all the way back
   to blank.
2. Gold requires glazing, for every column, with no exception.
3. TUFF: its own total, its own counter, its own By/At; it does not change the
   job's glass total; NOT TUFF never gets a counter.
4. Last-writer-wins, both ways, on stamps — including the case where the two
   stamps are equal.
5. The lock: office-complete greys the tablet's steppers and `tap()` refuses;
   only the office clears it; a queued tap made before the lock is not
   silently dropped without saying so.
6. Arch, astragal, fancy and extra are **never** written by this feature.
7. Over the whole run: no value written to `Production`, no other sheet
   touched, nothing written by the tablet except the floor's own columns.
8. The office drawer's existing controls behave exactly as they do today.

## 8. Out of scope

- Glass types on the floor. The tablet still counts one combined DG+TG number
  plus tuff.
- Any new column on the master job row (owner, 2026-09-09).
- The four hand-ticked glass columns.

## Amendments after review

_(append here; do not rewrite the sections above)_

### A. Built 2026-09-10 — one gap in the brief, and eleven decisions taken

**Status: built, tests green, not demoed, not committed.**

#### A1. A gap: `TuffTotal`. §4 is not buildable as written

§4 says the tuff counter's total "is fed from the workbook's TUFF column, the
same way the glass total is fed from DG + TG" — but §4.1's list of new columns
is `Tuff`, `TuffBy`, `TuffAt`, `OfficeDone`, with **no column to hold that
total**. The glass total has `Total`; the tuff count has nowhere to count
towards, so it cannot be clamped, cannot be shown as "4 of 11", and cannot be
"complete", which is what §2 needs for TUFF to go yellow.

`Total` cannot be reused: it is DG + TG, which §4 says explicitly the tuff
number is not part of. So a fifth column was added — **`TuffTotal` (number)**,
fed as a job fact exactly like `Total` — and the owner should be told, because
§4.1 is what the column script was approved against. **The script the owner
runs must create five columns, not four.**

#### A2. Blank cells are painted **white**, not cleared

§2 says a job tapped back to zero "leaves its glass cells with no fill at
all". The code writes `#FFFFFF` instead. The `Production` cells carry an
explicit white fill, and `format/fill/clear` takes the fill away and leaves a
hole that looks nothing like the rows around it — the reason `checkpoints.js`
has written white and never cleared since the day it shipped. The parser reads
white and no-fill as the same nothing, so on screen and on paper the result is
what §2 asked for. A cell that already has *neither* colour is left alone
rather than painted white for the sake of it.

#### A3. An equal stamp goes to the **office**

§7.4 asks for the equal case to be tested but §3 does not say who wins it.
Chosen: the office. Three reasons — the owner's own *"the master dashboard
should not be overridden"*; the office's stamp is written to the minute
(`nowStamp()`), so it is already biased *early* and loses ties it should win;
and it is the quiet answer, because a tie resolves to no write at all and
therefore can never loop.

#### A4. No floor stamp ⇒ **never** written

A `Glass station` row whose `DoneAt` is empty has never been tapped, and its
counters are the office's own seed (`ST.officeSeed`) echoed back. Painting the
sheet from them would be the dashboard arguing with itself, so a row with no
floor stamp is skipped entirely. This is also what stops the seed becoming a
colour.

#### A5. No office stamp ⇒ the **floor** wins, and that is an honest limit

An office tick made before this feature existed, and any colour somebody
painted by hand **in Excel**, leaves no dated record in `Dashboard Progress`
or `Dashboard Log`. Such a cell therefore cannot be shown to be the later
action, and a dated floor tap will walk it back. The office keeps every
control: ticking it again in the drawer gives it a stamp and it wins.

**Told plainly: somebody who colours the glass columns in Excel rather than in
the dashboard will always lose this contest.** There is no fix inside the
design the owner approved — `lastModified` gives one stamp for the whole file,
not one per cell.

#### A6. A queued tap that meets the lock is **dropped, and said**

§7.5 asks that it not be "silently dropped". The office marking a job complete
is the later action, so by the same rule the office keeps it; and on a locked
row nobody on the floor could correct a number afterwards. So the entry is
taken out of the queue, moved to a `BLOCKED` note in `localStorage`, written
to the console, and drawn on the card in red — *"Cutting 5 was not saved — the
office marked this job finished first"* — until the office unlocks the job.
No log line is written, because the counter never landed.

#### A7. `OfficeDone` is **derived**, not a new button

§3.1 says "when the office marks a job's glass complete", and §3 says the
office's controls do not change. So the lock is derived from the office's own
glass checkpoints: `OfficeDone = "Yes"` when every **DG and TG** item of the
job reads done. ARCH, ASTRAGAL, FANCY and EXTRA are not asked — they are hand
ticked and describe work the floor never sees. Un-ticking one glass checkpoint
unlocks the job, which is the "only the office can unlock it" §3.1 asked for.

#### A8. The tuff counter is **not seeded**

The feeder's one exception to "never the floor's columns" was given for three
counters, in the v3 spec and nowhere else (`CLAUDE.md` rule 3). Nobody has
widened it, so `Tuff` starts at nought on every row and only the floor moves
it — even on a job whose TUFF the office has already ticked. If the owner
wants it seeded, that is a new decision.

#### A9. Tuff is a fourth stage, but not one of the three

`ST.STAGES` / `STAGE_KEYS` still mean the three glass stages; `ALL_STAGES` /
`ALL_STAGE_KEYS` are the four. The difference matters in three places, all
deliberate:

- **the job's total** is still DG + TG. Tuff counts against `tuffTotal`.
- **"finished"** (gold on both boards, the Finished group, the job row's chip)
  is still the three. A job whose tuff nobody has counted is still a job the
  glass area has finished.
- **the job row's `Glass 8/24` chip** is still glasses × three stages. Adding
  tuff would make the denominator mean nothing.

`jobLeftFor` **does** count it, per §4: a person who cuts and counts tuff on an
8-glass, 11-tuff job has 19 left, not 8 and not 11.

The tuff stepper is drawn only on jobs that have tuff on them — a "Tuff 0 of 0"
row on every card is a fourth line of nothing on a workshop screen.

#### A10. A colour this feature does not own is left alone

If a glass cell is carrying something that is neither gold, nor yellow, nor
nothing — the sheet's own Cut green, or a colour painted for a reason of
somebody's own — it is not written, in either direction. **Consequence worth
knowing:** the floor's work will not show in such a cell until the office
clears it. The glass columns do not carry the Cut green today (it is a product
cell convention), so this should never be seen; it is there so the feature can
never destroy information it does not understand.

#### A11. `serialised()` is per tab, not per dashboard

§5.3 says writes go through `serialised()` "so two open office dashboards
cannot interleave on one sheet". `serialised()` is a promise chain inside one
page's JavaScript: it orders **this tab's** writes and can say nothing about
another tab's. It is used as asked (`CW.serialised("Production", …)`, now
exported), and each job's writes also run on that job's existing checkpoint
chain (`CP.cpChain`), so an office tick and a floor colour for one job can
never interleave. But two dashboards open on two desks are still two writers,
and the thing that makes that safe is the idempotence test and last-writer-
wins, not `serialised()`.

#### A12. No `Dashboard Progress` row and no `Dashboard Log` line

The colour writer writes **fills and nothing else**. Two reasons: the floor
counts one combined DG + TG number, so a per-type count like "DG 6 of 8" would
be invented; and a log line written by this feature would become the office's
own stamp on the next pass and poison the very comparison in §3. The floor's
own record of who did what and when is the `Station log`, which is unchanged.
A floor-driven colour change still appears in the office's Changes window on
the next load, as a cell difference like any other.

### B. Fix pass, 2026-09-10 — four review findings, one rehearsal finding, two owner decisions

**Status: fixed, tests green, still not demoed and not committed.**

#### B1. Minute-precision stamps now cover their whole minute (review, MAJOR)

`nowStamp()` and `appendLog` write `"YYYY-MM-DD HH:MM"` with **no seconds**, so
an office tick made at 15:00:40 is recorded as 15:00 — up to 59 seconds early.
Read literally, a floor tap at 15:00:20 looked like the later action and the
dashboard painted the office's own tick back out. The reviewer's run:

```
office really acted at 15:00:40 -> stamp recorded as 15:00:00
floor tapped at 15:00:20        -> 15:00:20
verdict: FLOOR WINS   <- wrong, the office acted 20 s later
```

Worse, it was a *delayed* betrayal. `noteChange` puts a full-second entry in
`CHANGES`, so the tab that made the tick won at first; `load()` then dropped
that local entry once the Log sheet carried the same line at minute precision,
and the dashboard silently changed its mind and repainted.

`stampMs` now returns **the latest instant a stamp can be describing**: a
stamp that names only a minute is taken as the end of that minute
(`+59_999 ms`). It is applied in the parser itself, so every reader gets it —
Progress `When`, Log `at`, and any future one. Consequences: anything inside
that minute loses to it (consistent with A3, ties to the office); the next
minute still beats it; and one office tick now reads the same before and after
the round trip, so the change of mind is gone. A stamp that carries seconds is
still taken exactly.

#### B2. A refused write now backs off, and says so (review, MAJOR)

`glassColourWrite` dropped its holds and cleared `GLASSC_BUSY`, so the next
pass re-planned the same cells at once — and `stationPoll` answers "moved"
every ten seconds while the floor is tapping. A failing `$batch` therefore went
out **every ten seconds, indefinitely**: a write storm against the live
workbook, the worst outcome this feature has available. Real trigger: somebody
opens the workbook exclusively in desktop Excel, or a token loses its write
scope, mid-shift.

Everything comparable in the file already had a brake (`stationDelta` marks a
refusing list off for five minutes; `cpQueue` bounds its replays; the feeder
has `stationFeedAgain`). This one now does too: each job counts its own
failures and waits **1 minute, then 5, then 15**, and after **five** failures
is given up on until the page is reloaded. A success clears the record. It is
surfaced the way `STATION_FEED_ERR` already is — the footer reads *"glass
colours not saved"* with the count, the reason and what happens next in the
tooltip — so a person can see it rather than only a console.

#### B3. One pass is capped, with one follow-up (rehearsal, MAJOR)

Rehearsed against the owner's real file: the feeder produces **141 rows**, and
switching this feature on after the floor has been recording with no office
dashboard open planned **up to 140 batched requests carrying 362 cell
PATCHes** in a single `load()`. Nothing capped it; `serialised()` only ordered
it, and it cannot order across tabs (A11). Day one is only about 6 cells, so
this bites on a *catch-up* — which is exactly when nobody is watching.

Fixed with the feeder's own idiom and the feeder's own numbers rather than a
second mechanism: `GLASS_MAX = 60` cells per pass, dispatched through
`stationSend` at three at a time, and one follow-up armed 30 s later when the
cap cuts a pass short (`glassColourAgain`, guarded to **one** timer — the bug
the feeder already had and fixed). A job is never split across two passes:
half a job's colours would be a lie on screen for thirty seconds and the other
half would only be re-planned anyway.

#### B4. `floorStamp` takes the newest by time, not by text (review, MINOR)

It took a lexicographic maximum. Every column it reads is hand-editable in
SharePoint, so `CutAt: "zzz"` won the comparison, `stampMs` answered 0, and
that job was **never coloured, silently, with nothing anywhere saying so**. It
failed closed, which was the right direction, but one stray character switched
the feature off for a job with no way to diagnose it. It now takes the newest
by parsed time and passes over anything it cannot read. A row on which nothing
parses still has no stamp, and still fails closed — deliberately now.

#### B5. An office yellow is read as "cut and hotmelted" (owner, 2026-09-10)

The rehearsal found that the **one** yellow glass cell in the owner's whole
workbook — one job's DG — was erased to blank by the realistic day-one run.
`officeSeed` turned a *gold* item into a count but a *yellow* one into nothing
(`cpStored` holds no count for a yellow item), so the row seeded at nought, the
floor's reading was "blank", and the first tap on that job painted the mark
out. A colour destroyed by a feature that had learned nothing from doing it.

**The owner's decision:** treat an office yellow as *"cutting and hotmelting
are complete"*, because under this feature's own rules that is exactly what
yellow now means. So `officeSeed` gained a branch: **every** DG/TG item at
least yellow seeds `Cut` and `Hotmelt` to the total and **`Glazed` to nought**.
Glazed staying at nought is the point of it, not an omission — yellow says
glazing is not done, and seeding it would make the cell read gold and assert
finished work that has not happened. The row round-trips to yellow and the
mark survives, by idempotence rather than by luck. A gold item is untouched:
it still seeds all three and still round-trips to gold.

**Every, not some, and why.** The floor's row carries one combined DG + TG
number and there is nowhere to put a per-type split (§8). A job with DG yellow
and TG blank therefore falls through to the count, as before. Seeding it to the
total would tell the **floor** that glass which still needs cutting is cut, and
a wrong instruction on the workshop screen is worse than a lost colour on a
report. **The residual limit, stated plainly: a job with one glass type part
done and another untouched can still lose that yellow once the floor taps.**
Whether the live yellow cell is on such a job is a question for the rehearsal,
not for the code: if that job's TG is blank with a quantity, this fix does not
save it; if it has no TG, or its TG is also marked, it does. **Worth checking
before the demo.**

This stays inside rule 3's seeding exception. It changes how `Cut`, `Hotmelt`
and `Glazed` are *derived*, not which columns may be seeded. **`Tuff` is still
not seeded**, and a row the floor has tapped (`DoneAt` non-empty) is still
never re-seeded, re-checked with the one-item read immediately before each
write.

#### B6. TUFF gold on glazing alone — confirmed (owner, 2026-09-10)

Amendment A raised it as a surprise worth checking: §2's table gives every
column gold on glazing, so TUFF goes gold when the job is glazed **even with
its own tuff count short** (3 of 11, say). **The owner has decided it stays as
built.** Recorded here so nobody "fixes" it later.

#### B7. Also corrected

§2's live-workbook figures. The rehearsal could not reproduce "704 filled glass
cells across 192 jobs"; the real figures are **193 glass jobs in production,
130 filled / 129 gold / 1 yellow** across the eight glass columns. Gold and
yellow were effectively right; the denominator matched nothing. §2 now carries
the measured numbers, because the owner may end up quoting them.
