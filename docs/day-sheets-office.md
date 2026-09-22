# Day sheets: the office's window and the two writes

Continues [[day-sheets-and-reports]] (§23).

### The office

A **Day sheets** chip beside Floor log, and a line under the Glass station
board's head ("Cutting this week: N of T sheets"). The window is the Floor log
window's shell: filters that a poll cannot take out from under somebody's
fingers, "Show more" paging, rows grouped by ISO week newest first, each week
with a subtotal carrying **its own stored `WeekTarget`** and the difference — so
changing the target today cannot rewrite what last week was measured against.
A week with no stored target falls back to the live one, which is what the
current week needs before anybody has saved a sheet in it.

The table carries **one sticky heading row**, built from the station
definition: Day · Who · each count's **short** label (the third element of its
`counts` entry — `ST.dayCountShort` falls back to the long question when a
definition gives none) · Total · Note · Saved. `app.js` holds no list of its
own, so a station with three counts or five gets its own headings. Without it
the counts were bare numbers with nothing to say which column was which.

**Every time on both screens is the local clock** (`ST.stClock`, one
implementation in `station-core.js`, which the office's `stWhen` is built on
and the tablet's "saved 14:02" uses). Slicing the ISO stamp gives the UTC
clock, and a sheet saved at 12:28 read 11:28 all summer.

**A duplicate row is read once.** `ST.dayRows` de-duplicates by `Title` with
the oldest item id winning, the same rule `buildJobs`, `weldRecords` and
`targetOf` use. Enforce-unique-values on `Title` is still the owner's step in
SharePoint, but the code no longer depends on it having been done — without
this, one duplicate doubled that person's day in the window, in the subtotal,
in the board's line, in the tablet's own week and in the report, silently.

**A correction is not wiped by the poll.** The window re-reads the list every
20 seconds; the *paint* waits while a row is in edit mode and is owed until
the edit is saved or cancelled (`DAY_PAINT_OWED`), so what somebody is typing
into a row cannot be thrown away under them. The read still happens, and the
moment the edit closes everything that arrived meanwhile is on screen. A
correction that fails to save leaves the row open with the typing in it.

**Two writes, and they are the whole of the new rule-2 exception:**

| write | body | log |
|---|---|---|
| the weekly target | `CW.listUpsert` of `Station targets` with `ST.targetFields(n, who, at)` — a number, a name and a time. **A whole number of one or more**: an empty box or a nought is refused with a toast and nothing is written, because a target of 0 is not "no target" (every week then reads as beating it, and the tablet says "35 of 0"). Removing a target is not built | one `Dashboard Log` line, "Cutting weekly target", from → to |
| a correction | `CW.listPatch` of one `Station day sheets` row with `ST.dayOfficeFields(counts, e)` — the counts, the note, `EditedBy`, `EditedAt`, and nothing else | one `Dashboard Log` line, "Day sheet corrected", the day and person, old → new total |

Neither writes `Station log`. Neither goes near the workbook. **No day sheet is
ever deleted, by either side.** The one delete in the feature is `listUpsert`'s
own duplicate settle on `Station targets` — if two browsers set the target at
once and the unique rule is off, the oldest row is kept and the surplus one
deleted — which is the settle `Dashboard phases` has always used and can only
ever remove a duplicate of the row it is writing (`web/CLAUDE.md` rule 3 says
so).

## See also

- [[day-sheets-and-reports]] — previous: the tablet, the two new lists
- [[station-reports]] — next: the station report export template
