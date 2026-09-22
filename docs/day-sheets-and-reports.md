# The end-of-day sheet, a weekly target, and a report for every station

## 23. The end-of-day sheet, a weekly target, and a report for every station

Built 2026-09-21. Spec:
[`docs/specs/2026-09-21-day-sheets-and-station-reports.md`](specs/2026-09-21-day-sheets-and-station-reports.md).
Builds on [[glass-two-stage]] (§22).

**What it is.** The cutter fills in a paper work sheet every day — name, day,
four numbers, and a line when something got in the way ("air for cutting not
working 9:20–11:21"). That sheet is now on the Cutting tablet, the entries are
in SharePoint, the office reads them by date, weekday and week against a
**weekly target it sets**, and any station's whole record can be exported.

### Two new lists, and where they live

`Station day sheets` (one row per person per station-stage per day) and
`Station targets` (one row per station-stage). Both are resolved through the
station definition's `site`, exactly like the station's other lists — so when
the glass lists move to `Floor stations` these move with them and no code
changes. Columns and the owner's creation recipe: `docs/STATIONS.md`.

**No list is created by code.** Until they exist both screens say which list is
missing, plainly, and write nothing; each asks again on its own clock (one
minute after a dropped connection, five after "there is no such list"), so the
morning the owner makes them every open screen picks them up without a reload.

### Switched on by the definition, not by the page

`ST.GLASS.daySheets` is `{ cut: { counts: [["Clear", "Clear glass sheets cut"],
…], unit: "sheets" } }`. `ST.daySheetOf(def, stage)` answers with that or with
nothing, and **nothing** — the header button, the reads, the office window, the
report's day-sheet columns — happens for a stage that answers nothing. There is
no `if (stage === "cut")` anywhere in `station.js`. Hotmelting or welding get a
sheet by gaining one line of definition.

### The tablet (Cutting only, today)

A header button **End of day** opens the sheet in the board's place: one
full-width column, portrait, 16px boxes (`inputmode="numeric"`), nothing under
44 px and no sideways scroll. The person's name, the weekday and the date are
at the top and not editable; under the four boxes and the note, "Today: N
sheets" and "This week: N of T" with a bar — N being this person's saved sheets
this ISO week **plus what is typed now**, T the target, or "no target set" in
those words rather than "of 0".

- **A draft belongs to the day it was started**, carries **whose** it is, and
  survives a poll, a reload and midnight (`cw_daysheetdraft`). The form's
  header shows that day and says so when it is not today; Save files it under
  that day, that day's ISO week and the target that was in force then. It is
  dropped at the end of the **following** day. Keying it on "today" was the
  first build's bug in both directions: typing at 23:55 and saving at 00:01
  filed the evening under tomorrow, and the same tick threw an unsaved draft
  away. `who` is why the next person to pick their name is never handed
  somebody else's numbers on a form that would save them under their own —
  and **switching person, or the ten-minute idle lock, closes the sheet.**
- **Save is dead until something has been written.** An untouched form is not
  four noughts. A day with nothing cut but a line about why ("machine down all
  day") is a real entry and saves perfectly well.
- **Save asks once** ("It cannot be changed from the tablet afterwards"),
  queues one row (`cw_stationdayq`) and sends it through the page's existing
  queue — so offline it says "waiting to send" and goes when the wifi is back.
  Re-opening later that day shows the saved sheet read-only, with
  "saved 17:02 — ask the office to correct a mistake"; the next day it is a
  blank form again.
- **Only a list that is genuinely missing refuses a save.** A list that could
  not be *read* is the workshop wifi, and the sheet is queued exactly like any
  other offline write; refusing it was the one way this feature could lose
  somebody's day. `DAY_MISSING` is the flag that tells the two apart.
- **A count that is not a whole number, or over 9999, is refused at the form**,
  not rounded: `ST.dayFields` answers null for a minus, a decimal, a word or a
  stuck finger, so there is nothing to queue and nothing to send.
- **A unique-value refusal means ALREADY SAVED, never an error.** A second
  tablet on the same key, or this tablet replaying a row whose answer was lost,
  makes `flushDay` read the list back; if the row is there the owed item is let
  go and the saved sheet is what the person sees.
- **A row SharePoint keeps refusing stops pointing at the wifi.** After three
  4xx refusals of the same row the words become "could not be saved — tell the
  office", and the row **stays queued** (the tablet cannot show it, so dropping
  it is the only way the day is really lost). The count survives a reload with
  the row. An owed day sheet no longer holds `checkBuild` back from reloading
  onto a new build — the queue is in `localStorage` and is rebuilt on load, and
  a new build may be exactly the fix a refusing list needs (`owingWrites`,
  which is the counter and log queues only, is what a reload waits for).
- The person's last seven saved days are listed under the form, read-only.

The tablet POSTs and reads. It never PATCHes or DELETEs either list, and it can
never write `Station targets` at all.

## See also

- [[glass-two-stage-office-board]] — previous: the office board, review findings, tests
- [[day-sheets-office]] — next: the office's window, the two writes
- [[station-reports]] — the export template built on the same two lists
