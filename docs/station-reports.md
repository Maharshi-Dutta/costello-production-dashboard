# The station report

Part of [[day-sheets-and-reports]] (§23).

### The station report

A third export template beside Default and the John print sheet, reached from
the **Report** chip on a station board (with that station preselected) or from
the Export window's Template row. Inputs: a station-stage and a period (this
week, last week, this month, custom). Excel only; PDF is not built.

`stationReport(def, stage, data, period)` in `export.js` is pure and **has no
station-specific branch in it**. Everything a station says about itself comes
off its definition:

| what it asks | glass | welding |
|---|---|---|
| `def.name`, `def.stageLabel(stage)` | Glass, Cutting / Hotmelting | Welding |
| `ST.daySheetOf(def, stage)` — the day-sheet counts | the four, for `cut` | none |
| `def.reportJobs(data, stage)` — **the adapter** | `ST.glassReportJobs`, one row per job | `WELDC.weldReportJobs`, one row per job **and product group** |
| `def.reportLogStages(stage)` — which `Station log` stages feed this report stage | `k => [k]` | its part keys, `frames` and `sashes` |
| `def.reportGroupLabel` — has this station's log a product group in it? | absent | `"Group"` |

**`reportLogStages` is why the welding report was empty at first.** The welding
tablet logs `Stage = frames | sashes` (`weldLogEntry`), never "weld", so
matching the report's stage name against the log found nothing: no Activity
sheet and an empty Days. The fixture that missed it hand-wrote `Stage: "weld"`;
it is now built through `ST.logFields(WELDC.weldLogEntry(...))`, the way the
tablet builds it.

The other four sheets are built from the three lists every station shares.

| sheet | what is on it | left out when |
|---|---|---|
| Summary | station, stage, period, who and when; then a line per ISO week: units recorded, jobs touched, jobs complete at this stage, and — for a stage with a day sheet — the counts, the total, the target and the difference | never (an empty period gets a Summary saying so) |
| Days | a row per day: units recorded and by whom; for a day-sheet stage the four counts, the total, the note and whether the office corrected it | no log line and no sheet in the period |
| Jobs | the station's own adapter's answer | the board is empty |
| Activity | the `Station log` lines in the period, oldest first, with a note row saying the office's own edits are in `Dashboard Log`. **Group** only where the definition names one, and **Part** only where the report stage is made of more than one log stage — glass has neither and gets six columns, welding has both and gets eight | no lines |
| Notes | `Station comments` in the period: when, who, job, text | no notes |

**One definition of "day" for the whole report.** A day sheet is filed under
the tablet's local date; log lines and notes are bucketed by
`ST.dayKey(new Date(at))` — the local date of the stamp — and not by slicing
the ISO string, which is the UTC date. In Irish summer time that put everything
recorded between midnight and one in the morning on the day before, in a
different row of Days from the sheet it was cut for and sometimes in a
different ISO week of the Summary. `data.dayOf` overrides it for a caller
reporting in a fixed zone, and is how the test pins the behaviour down without
being able to change the machine's own zone.

**Rule 3.** The free-text columns — a day sheet's note, a floor note and **the
customer name on the Jobs sheet** — go through `ST.stripContact` on their way
into the file. The customer joined them after review: it is free text off the
sheet like any other, welding's feeder has stripped it since that station
shipped, and "Customer Two 086 123 4567" typed into a name is exactly how a
phone number reaches a file that has no column for one. That is the strip
`welding-core.js` has used for the COMMENT column since 2026-09-16, **moved to
`station-core.js` on 2026-09-21** so it has one home: `weldStripDigits` is now a
call to it with the comment's own cap, welding's behaviour and its 61 checks
are unchanged, and it fails closed (no strip to hand, no free text at all).
Nothing else in the report is free text and no sheet has a column for a phone
number, an eircode or an address. File name: `Station report - <Station>
<Stage> - <from> to <to>.xlsx`, and every report writes the usual
`Dashboard Log` line through `noteChange` (`exportLogFrom(..., "station", …)`
names the template, the station, the stage and the period).

### Tests

`test_daysheets.js` (52 checks), plus `test_export.js` 53 → 54 — the standing
"no phone number and no eircode in any export" scan now builds a station report
from fixtures that type a phone number in three written shapes and an
address code into a day note and a floor note, and runs over the rows and over
the real workbook. The suite covers the row shape and the refusals, one
save per person per day, the unique-value refusal read as already-saved, the
offline save, the week helper across a year boundary and at 23:30 local, "this
week N of T" including the draft, both office writes, the week subtotal's stored
target, the filters, a missing list on both screens, all five report sheets for
glass cutting, hotmelting without the day-sheet columns, welding's per-group
Jobs sheet, an empty period, rule 3 in four written shapes, and the workbook
gate over the tablet's files and over every request in the run.

**The review's thirteen findings each brought a test that fails without its
fix** — the first build's suites were green and caught none of them, which is
the thing worth remembering about this feature. A fixture that agrees with the
code instead of with the tablet (`Stage: "weld"`, hand-written) proved nothing;
so did a suite that never asked what happens at midnight, to a second person at
the same tablet, to a duplicate row, to an empty target box, or to a
correction while the poll was running. The layout one (D13) has no offline test
at all and was found by a browser harness measuring `scrollWidth` against
`innerWidth` at 800 px; what `test_station.js` can hold is the CSS rules
themselves, and it now does.

### Not built here

- Day sheets for hotmelting or welding (one line of `daySheets` each, when the
  owner asks and says which counts they want).
- PDF station reports. Charts. Emailing a report. Per-day or per-glass-type
  targets.

## See also

- [[day-sheets-office]] — previous: the office's window, the two writes
- [[glazing-station]] — next: the glazing station
- [[export]] — the export builders this template sits beside
