# Welding: a remake counter beside All (2026-09-23)

**Status:** shipped 2026-09-23, build 20260923-1115 (`6cbeac5` build,
`1eb7c61` spec, `d5e2ed7` fix pass, `6def102` build stamp), HISTORY A34.

## Context

The welder sometimes has to weld a frame or a sash again before it is right.
Nothing recorded that. The owner asked (2026-09-23) for a remake counter on
the welding tablet, beside the All button of every component line, that the
welder moves up and down: "I had to remake that frame 3 times before it was
completed."

Owner's answers before coding: build it directly (no implementer subagent);
the office gets a read-only view; the welder's name never enters the repo.

## Hard rules (unchanged)

- Nothing in this feature writes the workbook. Rule 2 stands: the office never
  writes `Station log`; the feeder never writes the floor's columns.
- The office's welding board writes eight floor columns and **never a remake
  count**. The remake counts are the tablet's alone.
- Rule 3 (no phone numbers, no eircodes) is untouched: no new free text.

## Data model

Two new Number columns on `Welding station` (`Floor stations` site), added
by script (`add_remake_columns.py`, rehearsed on a throwaway list first)
after the push on 2026-09-23:

| column | written by | meaning |
|---|---|---|
| `FramesRemade` | the tablet only | how many times a frame of this group was welded again |
| `SashesRemade` | the tablet only | the same for sashes |

They are floor columns (`WELD_FLOOR_FIELDS`, `WELD_COUNTER_FIELDS`): the
feeder never writes or seeds them, `weldFloorOnly` lets them through only as
numbers. A remake tap's PATCH is the count alone (see Amendments), and it
logs one `Station log` line with `Stage = frames-remake` / `sashes-remake`
(`From`/`To` the old and new count, `Who`/`At` the welder and the moment).

**Sequencing:** the columns must exist before the build is live, else the
tablet's PATCH is refused and the tap is owed until they do.

## Behaviour

- A remake count is a record, not progress: it moves no done count, no
  colour, no "left", no header capsule.
- No ceiling (a frame can be remade more often than the job has frames);
  never below nought. `All` means nothing to it; `None` sets it to nought.
- Queued, sent, re-based and logged through the same queue as a done tap,
  under its own counter key (`frames-remake`, `sashes-remake`), so an
  offline tap survives a reload and a rise under a queued tap re-bases
  without a clamp.
- Locked for anyone not holding the weld stage, as the steppers are.
- Office: read-only `↻ N` on each expanded line of the welding board; the
  drawer's Welding line adds `↻ N remade` when the job has any; the station
  report's Jobs sheet gains `Frames remade` / `Sashes remade` columns.
  `weldOfficeEdit` refuses a remake key.

## UI

Tablet: a pill after All, set off by a wider gap: `− REMADE N +`. Buttons
48 × 56 px (narrower than the done buttons on purpose, tapped a few times a
week). `−` disabled at nought. Amber border once N > 0. Below 520 px the pill
wraps under the done buttons rather than pushing the page sideways.

## Tests delivered

- `test_welding.js` 7b: record/line carry the count; `weldApplyTap` on a
  remake key (no ceiling, floor at nought, All ignored); `weldTapFields`
  keys; `weldFloorOnly`; `weldRebase` unclamped; feeder never writes or
  seeds it; card signature changes; report columns; log stage.
- `test_welding.js` 9: the office cannot move a remake count (no request);
  the open board row shows `↻ 3` and offers no button for it.
- Feeder test: the never-written set now includes the two new columns.
- Browser rig (scratchpad `remake_check.js`, headless Edge, stubbed Graph):
  23/23 — pill on every line, one PATCH carrying the count alone, one log
  line, done count and colour unmoved, All beside it still works, 700 to
  1100 px no overflow and one column, 400 px no overflow, office row shows
  `↻ 3` with no remake button, steppers aligned.

## Amendments after review

Independent review (2026-09-23), two findings, both fixed in one pass:

1. **A remake line was summed as work in the station report.** The first
   build put the remake keys into `reportLogStages`, meaning to show them on
   the Activity sheet; but the report's Summary and Days sum every line of
   those stages as "Units recorded", so five remake taps would have read as
   five more frames welded. Fix: `reportLogStages` is `frames`/`sashes`
   again. Remake lines stay in `Station log` and on the office's Floor log
   panel ("Frames remade · 2 → 3"); the report carries the counts on its
   Jobs sheet only.
2. **A remake tap stamped `DoneAt`, closing the feeder's seed for ever.**
   `DoneAt` empty is the only thing that lets the feeder seed a row's done
   counts from the office's own record. A welder tapping the remake pill on
   a fresh row before any done tap would have made that row "touched" with
   nothing welded. Fix: the remake PATCH is the count alone — no By/At, no
   `DoneBy`/`DoneAt`. Who and when are on the log line.

## What to report

Suites, rule-2 grep on the diff, name grep, the rig count, and four
screenshots (tablet light/dark, after taps, office board).
