# Glass station v3: one number per job, cards you tap directly (brief, 2026-09-08)

Builds on v1 and v2 (`2026-09-08-glass-station.md`, `…-v2.md`, both with their
amendments). Everything there still holds unless changed here. Hard rules
unchanged except where §"Seeding" says so explicitly. No commit, no push, no
network, no live systems; all nine suites green with `set -o pipefail`.

## What the owner decided (2026-09-08, after first use on the tablet)

1. **No glass types on the floor.** Per job the tablet shows: job number,
   customer name, **one total number of glasses**, and the three stages
   Cutting, Hotmelting, Glazing. No DG/TG/TUFF/NOT TUFF anywhere on the floor.
2. **Total glasses = DG + TG**, nothing else. TUFF / NOT TUFF describe those
   same units and are not added; ARCH, ASTRAGAL, FANCY, EXTRA are ignored.
3. **No progress bars, no expanding, no scrolling to update.** Every card
   shows its three counters with their buttons on the card itself. One tap on
   *All* finishes the person's stage for that job; + and − for part of a job.
4. **The current state carries over.** A job whose glass the office has
   already ticked done or part done starts on the floor with those numbers
   (R4941 was complete in the office but showed nothing on the floor).
5. **A finished job goes gold**: when all three counters equal the total (and
   the total is > 0) the whole card turns gold, on the tablet and on the
   office board, and drops to a Finished group at the bottom.
6. Real time stays as v2: delta polling every 10 s on both sides.

## Data model

### `Glass station` list (one row per job)

| column | written by | meaning |
|---|---|---|
| Title | feeder | the job number, upper-case, unique |
| Job | feeder | same as Title |
| Customer | feeder | customer name, max 70 chars |
| GlassType | feeder | the literal `GLASS` (column kept; nothing reads it) |
| Total | feeder | DG + TG for the job |
| Seq, Active, FedAt, FedBy | feeder | as before |
| Cut, Hotmelt, Glazed | floor, **or the feeder while the floor has not touched the row** (see Seeding) | counters |
| CutBy/At, HotmeltBy/At, GlazedBy/At, DoneBy, DoneAt | floor | as before; the feeder never writes these |

The list is empty today, so there is no migration. `glassSlice` yields one row
per job with glass (`total = dg + tg`, jobs with `total = 0` produce nothing).
`Station log` lines write `GlassType = "GLASS"`.

### Seeding from the office's own record

The office's per-glass state is the glass checkpoints (`checkpoints.js`:
`cpStatus`/`cpStored` per glass item; statuses `done` (gold), `process`
(yellow, with a done/total count in Dashboard Progress), `cut` (green), or
nothing). Add a pure `officeSeed(job, counts)` in `station-core.js` that
looks only at the DG and TG items and answers `{cut, hotmelt, glazed}`:

- every DG/TG item `done` → all three = total;
- otherwise if any item is `process` or `done` → all three = the office's
  done count summed over DG/TG (a `done` item counts its full quantity);
- otherwise if any item is `cut` → cut = total, the others 0;
- otherwise all 0.

The feeder writes those three counters **only** on a row it is creating, or
on an existing row whose `DoneAt` is empty (the floor has never tapped it).
Once `DoneAt` is set the feeder never writes a counter again. This is the one
change to the feeder hard rule; `feedPlan` takes the seed per row and the
tests prove both halves (seeded while untouched; never after the first tap).

## Tablet (`glass.html`, `station.js`, `station-core.js`)

- Header as v2: name · stages, Switch person, Light/Dark, Sign out, plus a
  **search box** (job number or customer) that filters the cards as you type.
- **Cards**, one per job, in Seq order, stacked (one column under 700 px,
  two above). A card is:

  ```
  R4941   Edward D…                       12 glasses
  Cutting      [ − ]  12  [ + ]  [ All ]
  Hotmelting   [ − ]   0  [ + ]  [ All ]
  Glazing      [ − ]   0  [ + ]  [ All ]
  ```

  The person's own stage rows are bright and tappable; the other rows show
  their numbers but are greyed and disabled (`disabled` + `aria-disabled`,
  "not yours"). Buttons ≥ 56 px. `All` becomes `None` when the stage is at the
  total. No chips, no bars, no expand/collapse, nothing to scroll inside.
- **Gold card** when `cut === hotmelt === glazed === total` and `total > 0`:
  class `done`, gold background (`--done-bg`) and gold edge, still readable;
  such cards move under a "Finished · n" heading at the bottom (collapsed by
  default, tap to show).
- Everything else from v2 stays: picker + PIN, lock, gating in `tap()`, queue
  with site stamp, log line per sent write, delta poll, incremental repaint
  (`boardDiff` keyed by job), never a write without a resolved site, checkBuild.

## Office (`app.js`, `index.html`)

- Board cards: job, customer, "12 glasses", the three counters with who·when
  under each ("Cutting 12 of 12 · Person A · Tue 14:02"), the "last: …" line
  from the log; gold card when finished, same rule. No glass types.
- Drawer section: same three lines plus the timeline and Full log.
- Log window: the glass-type column is removed (the line reads job, stage,
  from → to, who, when).
- Seeding runs inside the existing `feedStation()`; it needs `cpStored` /
  `cpStatus` from `checkpoints.js` (already loaded before `app.js`), passed
  into `ST.glassSlice`/`officeSeed` as plain data so `station-core.js` stays
  free of app globals.

## Tests (extend `test_station.js`)

One row per job with total = DG + TG and nothing from the other types; the
four seed rules; the feeder seeds a new row, re-seeds an untouched row when
the office count changes, and never touches a row with `DoneAt` set; the
tablet card markup has no expand control, no bars, no type names, and the
three steppers are on the card; gating unchanged; gold class on completion on
both boards; search filter; the whole-run sweep (no workbook path, no DELETE,
floor PATCH keys ⊆ floor fields, feeder writes ⊆ feeder fields ∪ the three
counters only when allowed). Keep every existing suite green.

## Docs

Update `docs/STATIONS.md` (data model and the seeding rule), `docs/REFERENCE.md`
§11, `docs/ARCHITECTURE.md` (one sentence), `docs/specs/README.md` (v3 listed,
v2 superseded). Placeholders only.

## Report back

Files changed; the full pasted suite output; one paragraph per numbered owner
decision saying where it is done and which test covers it; anything unsure.

## Amendments after review (2026-09-08, manager)

1. **Seed guard at write time (blocker).** A feeder write that carries any of
   the three counters GETs that one item's `DoneAt` immediately before the
   PATCH and drops the counters if it is no longer empty (the guard the
   refused-POST path already has). Test: plan a seed, set `DoneAt` on the row
   between the read and the write, assert the PATCH carries no counter.
2. **Old rows.** `jobRecord` prefers the item whose Title equals the job
   number; `jobBoard` de-duplicates by job (Title = job wins, else oldest
   id), so leftover `JOB|TYPE` rows can neither feed the drawer nor count as
   a finished card. The manager clears the old rows by script before deploy;
   `docs/STATIONS.md` says the list is one row per job and that older
   `JOB|TYPE` rows are ignored.
3. **Re-base a waiting tap.** When a delta raises a counter for a row/stage
   with a queued entry whose `from` is below the new list value, the entry is
   re-based (`value = listValue + (value - from)`, clamped to the total) and
   its `from` reset, so a seed that arrived while the tablet was offline is
   not overwritten. Test with a queued `+1` and a seed of 12 arriving.
4. **Switch person clears the search** (`QUERY` and the input value).
5. A seed write is skipped when the row's `FedAt` is newer than this
   dashboard's workbook stamp (`lastStamp`), so an older dashboard cannot park
   a row on older numbers.
6. `#search` font-size 16px (no iOS zoom).
7. `docs/REFERENCE.md` §11: drop "or glass type" from the sweep sentence.
8. `total === 0`: the All/None button is disabled and reads "All".
