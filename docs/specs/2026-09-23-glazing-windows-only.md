# Glazing counts windows only

**Status: approved to build, 2026-09-23** (owner's answers in chat that day:
door-only jobs leave glazing, yes; card words say windows only, yes; no
separate rehearsal — the feeder mechanism in daily use since 2026-09-21 is the
rehearsal).

**Builds on** `2026-09-21-glazing-station.md` (section B, the feeder) and
`2026-09-22-glazing-board-polish.md`.

## Context

Owner, 2026-09-23: "the number of glazed units should be only windows … so if
a job has 11 windows and 2 doors the glazing should be 11 and the Total count
should also be just all jobs (windows each job)". Doors are not glazed at this
station.

## Hard rules

- Unchanged: no workbook write anywhere in glazing; the feeder writes only its
  own columns (`Windows`, `Doors`, `Total`, job facts), never `Glazed*`,
  `Done*`, never deletes. Rule-2 grep clean.
- The `Glazing station` list keeps its columns. `Doors` is still fed as a fact;
  it is simply not counted.
- No real names.

## Behaviour (all in `glazing-core.js`; nothing else changes)

1. `glzSlice`: `total = wnd` (was `wnd + drs`). A job with no windows on the
   `Production` sheet is not fed (`total > 0` gate, unchanged); if it was fed
   before, the existing feeder plan sets it `Active = No` and it leaves the
   tablet and the office board.
2. `glzQtyWords`: windows only — "11 windows" / "1 window"; doors are not
   mentioned on the card, the board row or the drawer.
3. Everything that reads `Total` (tablet card and header "N left", office
   board, drawer line, station report, the phase voice) follows without code.

## What happens live, once, on the office's next load

The existing feeder diff PATCHes `Total` on every fed row that has doors, and
`Active = No` on every door-only row. A row the glazer had already tapped
above its new total shows clamped (`11 / 11`, gold); the list keeps the old
`Glazed` until the next tap. Nothing in Excel, nothing in `Station log`.

## Tests to deliver

- `test_glazing.js`: the "windows plus doors" assertions become windows only
  (R8001: `total` 6, `Total: 6` in the feeder fields and the fixture rows; the
  quantity words "6 windows"); one new case: a door-only job (windows 0,
  doors N) is not in the slice; `glzDoorsOf` still reads `drsMain` (the fact
  is still fed).
- Suites otherwise unchanged; `test_pages.js` 4/4.

## Amendments after review

- **A door-only job the glazer had already recorded units against loses its
  voice in the phase bar.** Its row goes `Active = No` on the office's next
  load, so `floorPhaseOf` stops hearing it and that job's phase falls back to
  the sheet's own reading, or to welding's if welding has recorded anything.
  Nothing is cleared, reset or deleted — the counter stays exactly where the
  floor left it on a row nobody reads any more. Accepted: it is what "door-only
  jobs leave glazing entirely" means, and a job that is not glazed here should
  not be saying "In glazing" anywhere.

## What to report

`git diff --stat`, suite lines, and the count of live rows the feeder will
PATCH (from the local workbook copy: rows with doors > 0 and windows > 0, and
door-only rows).
