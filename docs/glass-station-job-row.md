# The glass chip on the ordinary job list

Part of [[glass-station-overview]] (§11).

**On the ordinary job list (2026-09-09).** Every job row carries a `Glass`
column of its own (the ninth, 96px, between `Components F·S·T` and `On
sheets`): one badge reading the floor's progress across both glass stages —
`Glass 12/16` for eight glasses cut 8, hotmelted 4 — gold on the same
"finished" rule the board and the tablet use, with the breakdown and "the Excel
file is not involved" in its `title`. A job the floor has never been fed
renders that cell empty. The rows repaint when the floor taps (`redrawStation`
now redraws the plain list, not only a floor-station board) and the poll runs
at the fast rate whenever the rows on screen carry the floor's work
(`ROWS_GLASS`). A repaint is held off only while somebody is typing in the list
or dragging a row, and is then *owed* (`ROWS_STALE` / `stationCatchUp`) rather
than dropped; it fires only when the chips would actually say something
different (`ROWS_CHIPS`), and without the row entry animation. `stationRecords()`
caches `ST.jobRecords(items)` on `STATION_ITEMS` array identity — O(rows +
items) instead of O(rows × items), measured 11–16 ms → 0.71 ms per render.
**Still no workbook write and no list write:** the one read added is the
existing shared `stationReadIfNeeded()`, once per session. Spec:
[`docs/specs/2026-09-09-glass-chip-on-job-row.md`](specs/2026-09-09-glass-chip-on-job-row.md).

**Since 2026-09-10.** The floor's counters now reach the `Production` sheet's
four glass columns as fills, painted by the office dashboard; there is a
fourth counter (tuff) with its own total; and the office can make a job
read-only on the tablet. See [[glass-colours]] (§17) — and note that `STAGE_KEYS` still means the
three glass stages everywhere in this section, with the four in
`ALL_STAGE_KEYS`.

**Tests.** `test_station.js` (235 checks) proves, over the whole run, that no
request touched the workbook, no DELETE was sent, every floor PATCH is a
subset of the floor columns, every feeder write is inside `ST.FEEDER_WRITES`
(the job facts plus the three counters, never a By, an At or the last touch),
and no body carried a phone, eircode, county, price or comment.

## See also

- [[glass-station-overview]] — previous: glass station (floor dashboards)
- [[john-print-sheet]] — next: row colour code and the John print sheet
- [[glass-colours]] — the floor's counters reaching the Production sheet
