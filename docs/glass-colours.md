# Glass colours: the floor's work reaches the Production sheet

## 17. Glass colours: the floor's work reaches the Production sheet

Built 2026-09-10. Spec: [`docs/specs/2026-09-10-glass-colours-two-way.md`](specs/2026-09-10-glass-colours-two-way.md),
including its long "Amendments after review" section, which records the one
gap in the brief and the twelve decisions taken while building it. **This is
the first feature in which something done on the floor changes the master
sheet**, so read the boundaries before touching any of it.

**What.** The floor's counters decide the colour of DG, TG, TUFF and NOT TUFF
on the job's own row. It walks back down as well as up — gold to yellow to
blank — mirroring whatever the floor now says. **ARCH, ASTRAGAL, FANCY and
EXTRA are never written by this feature**, in either direction: the office
ticks those by hand, as it always has.

> **THE RULE CHANGED 2026-09-21 ([[glass-two-stage]], §22).** Until then: yellow once cutting AND
> hotmelting were complete, gold once **glazing** was. Now it is a count of how
> many of the two glass stages are complete — blank for neither, **yellow for
> exactly one**, **gold for both** — and TUFF is gold on its own count, blank
> once somebody has tapped it, and **not named at all** while nobody has
> (`TuffAt` empty). Everything below about glazing gating gold is history.

**Who writes it.** The **office dashboard** (`app.js`), never the tablet. The
tablet has no access to the workbook at all and that is the actual security
boundary; this is the feeder in reverse — the office watches the floor's list
and paints the cells.

> **CHANGED 2026-09-11 by [[status-list-is-truth]] (§19), step 3. Everything below about holds, stamps
> and settling is history.** The writer now puts the floor's colour on the
> **`Dashboard progress` record** first (`Source = "floor"`, `When` = the
> floor's own stamp, `Who` from `DoneBy`) and paints Excel from it, through the
> same `cpSaveRow` → fill → log path an office click uses. Last-writer-wins is
> two fields of one row: the floor does not write when an `office` or `excel`
> row's `When` is at or after `ST.floorStamp(g)`. Gone with it:
> `glassOfficeStamp`, `glassOfficeJobStamp`, `officeSettling`, `glassLogStamps`,
> `OFFICE_FLOOR_AT`, `glassCellNow`, the `gc` `PENDING` holds and the whole
> "a hold is never let go into a stale copy" mechanism — about 510 lines. What
> survives unchanged: the colour rule itself (`ST.glassColours`), the four
> columns, ARCH/ASTRAGAL/FANCY/EXTRA never written, the 60-cell cap and its
> follow-up, the per-job backoff, and the rule that a colour this feature does
> not own is never painted over.

**What reaches Production.** A single-cell **fill**, and nothing else, ever:
no value, no row, no formula, no number format, no other sheet. Not even a
`Dashboard Progress` row or a `Dashboard Log` line — the first would be an
invented per-type count (the floor counts one combined DG + TG number) and the
second would become the office's own stamp on the next pass and poison the
last-writer-wins comparison below.

**Where the code is.**

- `station-core.js` — the pure half: `glassColours(record)` returns
  `gold` / `yellow` / `""` per column; `floorStamp(record)` is the newest ISO
  stamp on the floor's row; `COLOUR_TYPES` is the four columns; `stageComplete`
  is "this counter has reached its own total, and there is a total".
- `app.js` — `glassColourPlan(j)` (what to write, or null), `glassColourWrite`
  (the write), `glassColourRun()` (one pass over every job), `stampMs`,
  `glassOfficeStamp`, and the `gc` branch of `pend` / `applyPending`.

**How, in five parts.**

1. **When it runs, and how much it may do at once.** After every `load()`,
   once the feed has brought the list up to date, and after every station poll
   that actually moved something (10 s while somebody is looking at the floor,
   60 s otherwise). It is a walk of the jobs in memory and **not one request**
   when there is nothing to do, which is what lets it be called six times a
   minute.

   It is also **capped**, for the same reason the feeder is and with the same
   numbers: `GLASS_MAX = 60` cells per pass, three writes in flight
   (`stationSend`, the feeder's own lane runner), and one follow-up armed 30 s
   later when the cap cuts a pass short (`glassColourAgain` — exactly one
   timer, never one per job). Rehearsed against the owner's real file: 141 fed
   rows, and switching the feature on after the floor has worked a week with no
   office dashboard open would otherwise plan **362 cell fills in one burst**.
   Day one is about six cells, so this bites on a catch-up, which is precisely
   when nobody is watching. A job is never split across two passes: half a
   job's colours would be a lie on screen for thirty seconds and the other half
   would only be re-planned anyway.
2. **Idempotence, which is the whole reason it is safe to call that often.**
   The desired colour is compared with what the sheet is already showing —
   `j.cp.glass[type]`, i.e. the *held* colour while a write is in the air and
   the *downloaded* colour once the file has caught up — and a cell that
   already says it is not written. Ten turns of the poll after a write send
   zero requests (`test_glasscolour.js` asserts exactly that). A cell carrying
   a colour this feature does not own (the sheet's own Cut green) is left
   alone in both directions.
3. **Last writer wins.** The floor's stamp is `DoneAt` (and the per-stage
   `At`s) on the `Glass station` row — the newest **by parsed time**, skipping
   anything unreadable, because every one of those columns is hand-editable in
   SharePoint and a text maximum let `CutAt: "zzz"` win, answer a time nothing
   could read, and switch the feature off for that job in silence. The office's
   is the newest of the job's glass `Dashboard Progress` rows (`Who`/`When`,
   read through `cpStored`) and the `Dashboard Log` lines for that job's glass
   items (already in memory as `CHANGES`) — so the comparison costs **no extra
   read**.

   Three formats meet in `stampMs`, and it returns **the latest instant a
   stamp can be describing**, not the instant it literally names. That
   distinction is the whole of it: `nowStamp()` writes no seconds, so an office
   tick at 15:00:40 is recorded as `15:00`, and read literally a floor tap at
   15:00:20 looked like the later action and painted the office's own tick back
   out. A stamp that names only a minute therefore covers that minute
   (`+59_999 ms`); one that carries seconds is taken exactly. It also makes the
   answer stable across the round trip — `noteChange` puts a full-second entry
   in `CHANGES` and `load()` later replaces it with the Log sheet's
   minute-precision copy, and both readings now answer the same side, where
   before the dashboard quietly changed its mind half a minute later.

   **A tie goes to the office** (a tie is therefore no write at all). **No
   floor stamp means never written** — an untapped row's counters are the
   office's own seed echoed back. **No office stamp means the floor wins**,
   which is the honest limit: a colour painted by hand in Excel leaves no dated
   record and will be walked back; ticking it in the drawer gives it a stamp
   and it wins.

## See also

- [[john-print-sheet]] — previous: row colour code and the John print sheet
- [[glass-colours-tuff-and-lock]] — next: the hold, ordering, TUFF and the lock
- [[glass-two-stage]] — the 2026-09-21 rule change this section's banner points to
- [[status-list-is-truth]] — the record this writer now writes through
