# The welding station: the office board, generalisation and the site pin

Continues [[welding-station]] (§21).

### The office's board

One row per job, **every section**, in sheet order: an overall bar,
`Frames x/y`, `Sashes x/y`, the last touch, and the unread-notes count from
`Station comments`. Click a row to open it: one block per product group, each
with the office's own **−, +, All, None** on Frames and Sashes.

An office edit writes **exactly five fields** — that part's counter, that
part's `By`/`At`, and `DoneBy`/`DoneAt` — through
`WELDC.weldOfficeFields` → `WELDC.weldFloorOnly`, which is the same filter the
tablet's own queue runs on, so no job fact can reach the body. It leaves **one
`Dashboard Log` line** per change ("Welding: R5303 CASEMENT WINDOWS frames,
3 → 6") and **no line of `Station log`**: that list is the floor's and stays
the floor's (rule 2, unchanged; the owner's answer was about visibility, and
the board's own **Floor log** panel is where the floor's log is visible).

The tablet sees the office's change as the row's last touch on its next poll.
A queued floor tap made *before* the office's edit is dropped and said so in
red on the card — `WELDC.weldRebase` decides that, and it is the same
last-writer-wins rule the glass clear uses: a counter that has **risen** under
a waiting tap re-bases the movement onto the new number, one that has **fallen**
under a row stamped **after** the tap is the later word and the tap goes.

The job drawer gets one read-only line under the window types — "Welding 9 /
16", with a button that switches the list to the board. **No new column on the
job row** (owner's rule, 2026-09-09).

### What was generalised, and what was deliberately not

`station-core.js` gained a **station definition** argument and nothing else:
`feedPlan`, `sliceHash` and `floorOnly` take one (default `ST.GLASS`, so every
older call site means exactly what it meant), `stationPeople` takes the
station's stage list, `boardDiff` takes the station's card signature. Nothing
glass-specific was added to it and nothing welding-specific went into it — the
welding rules are all in `welding-core.js`, and the station-independent tablet
shell (theme, gate, picker, PIN pad) is the new `station-ui.js`.

### The site pin

`CW.stationSite(which)` honours the definition's `site`, and each word is a
separate lookup by path that can only ever answer one site:

- **`"floor"`** — the `Floor stations` site, with no fallback. Missing is
  `null` and the page says so quietly.
- **`"own"`** — the workbook's own site, where the glass lists still are. The
  `Floor stations` site is **never** looked up on this channel: not on an empty
  cache, not after a forget, not on **Try again**.

The two keep separate cached ids (`cw_stationsite_own`,
`cw_stationsite_floor`), separate miss clocks and separate move counters, so a
glass 404 can never move the welding board and vice versa.

**The first build of this pin was wrong and the review caught it.** It returned
the cached site and otherwise fell through to the legacy resolver, which
*prefers* `Floor stations`. A 404 on any glass list call (which calls
`forgetStationSite`), a cleared localStorage, a new tablet or a new office PC
each empty that cache — and the very next resolve would have moved the glass
feeder, the glass colour writer and `clearFloorGlass` onto a site with no
`Glass station` list in it, permanently. There is no gap left to describe: the
only thing that can ever move the glass page is `ST.GLASS.site` being changed
from `"own"` to `"floor"` on the day its three lists are copied across.

### Not built, and said so

- Nothing paints the `Production` sheet from welding progress (owner, answer
  5b). The office ticks F/S gold by hand or in the drawer, as today.
- No reply channel from the office to the welding tablet.
- No move of the glass lists (owner, answer 9).
- Cutting is not a stage here: "no cutting, it should just have welding".

## See also

- [[welding-station]] — previous: what it is, rule 3, the seed, the tablet
- [[glass-two-stage]] — next: the glass split that followed this station's pattern
- [[station-feeder]] — the data-flow diagram this section's site pin is part of
