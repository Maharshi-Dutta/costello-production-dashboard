# Day sheets for welding and hotmelting; one Floor log for every station

## 26. Day sheets for welding and hotmelting, and one Floor log

Built 2026-09-28 (not shipped yet). Spec:
[`docs/specs/2026-09-28-day-sheets-welding-hotmelt-floor-log.md`](specs/2026-09-28-day-sheets-welding-hotmelt-floor-log.md).
Builds on [[day-sheets-and-reports]] (§23) and [[day-sheets-office]].

### Two more stages with a sheet, by definition

- `ST.GLASS.daySheets.hotmelt` — counts `DG`, `TG` (units), `counted: true`,
  `verb: "hotmelted"`, `target: false`.
- `WELDC.WELD.daySheets.weld` — one count `Squares` (squares), `target: false`.

`target: false` means: no `Station targets` read on the tablet or in the
office, no target line, no `WeekTarget` on the row (`ST.dayFields` drops it
even if handed one), no "Set target" control, and no Target / Difference
columns in the report. `ST.dayHasTarget(def, stage)` is the one question.

`counted: true` adds the `Counted` column to the read (`ST.dayFieldsFor`) and to
the row (`ST.dayFields(..., { counted })`; not written when unknown). It is
never in `ST.dayOfficeFields`' output, because that builder only takes the
definition's counts.

### The sheet moved into `station-ui.js`

`STU.stuDaySheet(cfg)` is the whole tablet sheet — draft (per day, per person),
save (one POST per person per stage per day, confirm first), owed-sheet queue
(`refused` count, read-back on a refusal), drawing and wiring — as one
controller per page. `station.js` and `welding.js` each make one (`DAY`) and
call `DAY.sheet() / read() / open() / html() / wire() / save() / flush() /
persist() / owing() / reset()`. Its storage keys are the page's: glass keeps
`cw_stationdayq` / `cw_daysheetdraft` (a sheet owed across the upgrade is still
owed after it), welding uses `cw_welddayq` / `cw_welddaydraft`.

**Hotmelting's own count.** `ST.dayCounted(logRows, who, day, stage)` is the
sum of `To − From` over the person's `Station log` lines at that stage on that
LOCAL day, never below nought. The hotmelt tablet reads `Station log` through
the delta feed when the sheet opens (the glass tablet did not read the log
before), adds its own lines still owed (`LOGQ`), and shows "You hotmelted N
units today". DG and TG are typed, not prefilled. When DG + TG differs from a
non-zero N, the one save question starts "You counted N taps today and typed M.
Save anyway?"; no taps at all is not a disagreement.

**Welding.** End of day on the header's first row beside the capsules; under
900 px the "updated" words are hidden so the row fits (measured 700–1100 px in
headless Edge, both rows exactly their width). The sheet resolves
`Station day sheets` through `WELD.site` — the second list, in `Floor stations`.

### The office's Day sheets window

One state per station-stage (`DAYST[id]`, id `"<Station>|<stage>"`,
`dayS(id)`), each read only when the window is on it (or, for Cutting, when the
glass board's week line asks — `dayBoardId()`). A selector at the top lists
`dayStages()` off `stationDefs()`; the window opens on the board's own stage or
on Cutting. Hotmelting shows `Counted` read-only beside DG / TG and a small
"≠ counted". Corrections and targets write through that stage's definition and
site; the `Dashboard Log` job column stays `"(<Station> <stage>)"`.

The station report's Days sheet now carries welding's squares and hotmelting's
DG / TG, plus a "Counted by the tablet" column where the stage counts
(`stationReportData` passes the stage's own rows, when the window has read them).

### The Floor log, every station

`openStationLog(job, station)`: a station filter (All / Glass / Welding /
Glazing / Fabrication); from a board it opens on that board's station, from the
job list on All, from the glass drawer on Glass. Sources are the arrays the
boards already hold (`STATION_LOG`, `WELD_LOG`, `GLZ_LOG`, `FABR_LOG`, each
through `ST.logRows(items, name)`), read if needed through the boards' own
`*ReadIfNeeded`. The office's edits come out of `CHANGES` through
`ST.officeLogRows` — the "what" words are matched in one list,
`ST.OFFICE_FLOOR_EDITS`:

| what | station · stage |
|---|---|
| `Glass cutting` / `Glass hotmelting` / `Glass tuff` | Glass · cut / hotmelt / tuff |
| `Floor glass counters` (an office clear) | Glass |
| `Welding: <job> <group> <part>` | Welding · part |
| `Glazing: <job>` / `Glazing astragal: <job>` | Glazing · glaze / astragal |
| `Fabrication: <job> <group> <part>` | Fabrication · part |
| `Day sheet corrected`, job `(<Station> <stage>)` | that station; unknown → All only |

`ST.floorLogMerge` sorts by the moment each stamp names (`ST.logWhenMs` reads
ISO, `yyyy-mm-dd hh:mm` and `dd/mm/yyyy hh:mm`, the last two local). Office
lines are marked **office**, count as lines but never as units in the
per-person counts. A station whose list is missing is one quiet line.

### Tests

`test_daysheets.js` (definitions, columns, Title keys, POST bodies, no
`WeekTarget`, `dayOfficeFields` never carries `Counted`, `dayCounted` at
23:30 / 00:30, the hotmelt tablet end to end, the office selector),
`test_welding.js` §12b (the welding tablet end to end, Floor stations),
`test_floorlog.js` (new: matching list, stamps, merge, filters),
`test_station.js` (the window with office lines and a missing station).

## See also

- [[day-sheets-and-reports]] — the cutting sheet this grew from
- [[day-sheets-office]] — the office's window and writes
- [[station-reports]] — the report template
