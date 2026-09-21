# The cutter's end-of-day sheet, a weekly target, and a report export for every station

**Status: approved to build, 2026-09-21** (owner's answers in chat that day).
**Builds on** `2026-09-21-glass-split-no-glazing.md`, which must be in the
working tree first: this brief assumes `glass.html?stage=cut` is the Cutting
page and that glazing is gone.

## Context

The cutter fills in a paper "Glass cutting work sheet" every day: name, day and
date, four numbers — clear glass sheets cut, K-glass sheets cut, satin sheets
cut, other obscure sheets cut — and a line of free text when something got in
the way ("machine not working", "air for cutting not working 9:20–11:21",
"helped on another bench through the day"). Saturdays are sometimes worked.
The owner wants that sheet on the Cutting tablet, the entries logged in our
system, visible from the office's Glass station window by date, by weekday and
by week, measured against a **weekly target the owner sets from the office
board**, and exportable. And the export is to be built once as a template, so
hotmelting and welding (and stations that do not exist yet) get the same
"full report of this station" from the office dashboard.

## Decisions taken (owner, 2026-09-21)

1. The target is **total sheets per week** — one number, not one per glass
   type, not per day.
2. The cutter **cannot correct** a saved day. **The office can**, from the
   dashboard.
3. The day sheet is for **Cutting only** for now; built so another station
   page can switch it on later without a new list.
4. **Export: yes.** What was cut in the day and in the week, what statuses he
   updated, a full report. Keep it a template; the same export for every
   station, present and future — "export a full report from my dashboard for
   each station's jobs."

## Hard rules

- The tablet never touches the workbook (the usual grep gate on the station
  files). Nothing in this brief writes the workbook at all, from either side,
  except the one `Dashboard Log` line per export and per office edit that the
  existing `noteChange` path already writes.
- **Rule 3: no phone numbers or eircodes in any export, ever; every export is
  logged.** The note a person types is free text, so it goes through the same
  strip welding uses for the COMMENT column (`weldStripDigits` and the eircode
  pattern in `welding-core.js`) **on its way into a report file** — move that
  strip to `station-core.js` as a shared pure helper if it is not reachable
  from `app.js`/`export.js` already, leaving welding's behaviour and tests
  exactly as they are. The on-screen view shows the note as typed.
  `test_export.js` has a standing test that no export carries a phone number
  or an eircode; extend it to the station report.
- No real person's name, email or domain anywhere. The person is "the cutter".
- No list is created by code. The owner runs a scratch script (manager's job)
  to make the two lists; until they exist both screens say so quietly and
  write nothing (the `Station comments` missing-list pattern).
- Edit/Write tools only for repo files. Do not commit. No live SharePoint.

## Data model — two new lists, in whichever site the station's lists live

Resolved through the station definition's `site`, exactly like the station's
other lists (`CW.stationSite(def.site)`), so when the glass lists move to
`Floor stations` these move with them and no code changes.

### `Station day sheets` — one row per person per station-stage per day

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, **unique** | tablet | `<Station>\|<Stage>\|<YYYY-MM-DD>\|<Who>` — the key; a second save for the same key is refused by SharePoint as well as by the page |
| Station | text | tablet | `Glass` |
| Stage | text | tablet | `cut` |
| Day | text | tablet | `YYYY-MM-DD`, the tablet's local date at save |
| Who | text | tablet | the signed-in person, as resolved for taps |
| Clear, KGlass, Satin, Obscure | number | tablet; office on an edit | sheets cut, whole numbers ≥ 0, blank on paper = 0 |
| Note | multiple lines of text | tablet; office on an edit | the free-text line, 500 characters at most |
| WeekTarget | number | tablet | the weekly target in force when the row was saved (blank if none set) — so an old week stays true after the target changes |
| SavedAt | text | tablet | ISO stamp |
| EditedBy, EditedAt | text | office | set only when the office corrects the row |

Append-only from the tablet: one POST per person per day, no PATCH, no DELETE.
The office may PATCH `Clear`, `KGlass`, `Satin`, `Obscure`, `Note`,
`EditedBy`, `EditedAt` and nothing else, and never deletes.

### `Station targets` — one row per station-stage

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, unique | office | `<Station>\|<Stage>`, e.g. `Glass\|cut` |
| WeeklyTarget | number | office | total sheets per week |
| SetBy, SetAt | text | office | who set it, when |

Upserted by the office (`CW.listUpsert`), read-only on the tablet.

**Weeks** are ISO weeks, Monday to Sunday, labelled `2026-W39` with the
Monday's date beside it on screen. One shared pure helper for "which week is
this day in" (`station-core.js`), used by the tablet, the office window and the
export, with a test across a year boundary.

**Rule 2 exception to record in `web/CLAUDE.md`, dated 2026-09-21:** the office
writes `Station targets`, and may correct the count and note columns of a
`Station day sheets` row; both are logged in `Dashboard Log`.

## Tablet — the Cutting page only

Switched on by the station-stage's definition, not by an `if (stage === "cut")`
in the page: add `daySheet: { counts: [["Clear","Clear glass sheets cut"],
["KGlass","K-glass sheets cut"],["Satin","Satin sheets cut"],
["Obscure","Other obscure sheets cut"]] }` (or the equivalent) to what describes
the Cutting stage, and have the page draw the form from that. A stage without
it gets no button. That is the whole of "built so another page can switch it
on".

- A header button **"End of day"**. It opens a full-width sheet, one column,
  portrait, big number inputs (`inputmode="numeric"`), no sideways scroll.
- Top of the sheet, not editable: the person's name, the weekday and date.
- Four number fields, a note box, a running **"Today: N sheets"**.
- Under it, the week: **"This week: N of T"** where N is the sum of this
  person's saved sheets this ISO week **plus** what is typed now, and T is the
  weekly target; "no target set" when there is none. A plain progress bar is
  enough.
- **Save** asks once ("Save today's sheet? It cannot be changed from the tablet
  afterwards."), POSTs one row, then shows the saved sheet read-only with
  "saved 17:02 — ask the office to correct a mistake". Re-opening the button
  later that day shows the same read-only view; the next day it is a blank
  form again.
- A draft typed but not saved survives a background poll and a reload
  (`localStorage` key `cw_daysheetdraft`, cleared on save and at the change of
  day). Offline at save: the row is queued and sent when the connection is
  back, using the page's existing queue pattern, and the sheet shows "waiting
  to send" rather than "saved".
- A second tablet saving the same key, or a replayed queue item that already
  landed, gets SharePoint's unique-value refusal: treat that as **already
  saved**, read the row back and show it; never as a red error.
- Below the form: this person's last seven saved days, read-only, one line
  each (day, four counts, total).

## Office — the Glass station window

Beside the existing **Floor log** chip on the Glass station board: a **Day
sheets** chip opening a window built like the Floor log window (same shell,
filters that do not steal focus on a poll, "Show more" paging).

- **Target control** at the top: "Weekly cutting target: [ 250 ] sheets · set
  by … on …" with a Save button. Upserts `Station targets`, logs "Cutting
  weekly target" from → to in `Dashboard Log`.
- **Filters:** from / to dates (default: this week and last), weekday, week,
  person.
- **Table:** one row per sheet — day, weekday, who, Clear, K-glass, Satin,
  Obscure, total, note, saved at, and "edited by the office …" when it was.
  Grouped by week, newest week first, each week with a **subtotal row: total
  of T target, and the difference**. The target used for a week is the
  `WeekTarget` on that week's most recently saved sheet; for the current week
  with no sheet yet, the live target.
- **Edit** on a row opens the four counts and the note; Save PATCHes the
  allowed columns plus `EditedBy`/`EditedAt`, logs one `Dashboard Log` line
  ("Day sheet corrected", the day and person, old → new totals). No delete.
- Read on the floor's own poll clock like the other floor lists; a missing
  list is a quiet explained state, and asks again later on its own.
- The cutting card area of the board shows one line under the board's head:
  "This week: N of T sheets" for the Cutting stage.

## Export — the station report, one template for every station

A new export **template**, "Station report", beside Default and the John print
sheet in `export.js`, reached two ways: a **Report** chip on each station board
(Glass, Welding) with that station preselected, and the Export window's
template choice.

Inputs: station (`Glass · Cutting`, `Glass · Hotmelting`, `Welding`, built from
the station definitions so a new station appears by being defined), and a
period (this week, last week, this month, custom from/to). Format: **Excel**
(`buildWorkbook`'s library and styling conventions). PDF is not built now.

The report is assembled by one pure function, `stationReport(def, stage, data,
period)`, returning sheets as plain arrays, where `data` is what the office
page already holds for that station (board rows, `Station log` rows,
`Station comments` rows, day sheets, target). Sheets, each omitted when the
station has nothing for it:

1. **Summary** — station, stage, period, generated when and by whom; per week:
   units recorded (from the log), jobs touched, jobs completed at this stage;
   for a stage with day sheets: sheets cut by type, total, target, difference.
2. **Days** — one row per day in the period: units recorded that day per
   person (from `Station log`), and for Cutting the day sheet's four counts,
   total, note (stripped per rule 3) and whether the office corrected it.
3. **Jobs** — every job on the station's board: job, customer, section/status
   words the board already shows, total, done at this stage, left, who last
   moved it and when, complete or not. For welding: one row per job and product
   group, Frames and Sashes as the board shows them.
4. **Activity** — the `Station log` lines in the period for this station and
   stage, oldest first: when, who, job, stage, from → to. (For welding the
   office's own edits live in `Dashboard Log`, not here; say so in a note row.)
5. **Notes** — `Station comments` for this station in the period: when, who,
   job, text (stripped per rule 3).

Customer names appear, as they do in the Default export. **No phone number, no
eircode, no address** — none of these sheets has a column for one, and the
free-text columns are stripped. Every export writes the usual `Dashboard Log`
line through `noteChange("(export)", "Export", …)`, with the from-text naming
the template, station, stage and period (extend `exportLogFrom`).
File name: `Station report - <Station> <Stage> - <from> to <to>.xlsx`.

Welding's definition (`WELDC.WELD`) and the glass definition must both produce
a report with **no station-specific branch inside `stationReport`**: anything
a station needs to say about itself (its stage list, its counter columns, how
a board row reads as "done / total / left") comes from its definition or from
a small adapter function on the definition. If welding needs such an adapter,
add it to `welding-core.js` and nothing else there changes.

## Not built here

- Day sheets for hotmelting or welding (one line of definition each, when the
  owner asks and says which counts they want).
- PDF station reports. Charts. Emailing a report.
- Per-day or per-glass-type targets.

## Tests to deliver

A new offline suite `test_daysheets.js` in the style of `test_comments.js`
(fake `fetch`, stub DOM), plus additions to `test_export.js`:

- row shape for a save (all columns, Title key, whole numbers, blanks → 0,
  note capped); negative, decimal and non-numeric input refused at the form.
- one save per person per day: the second attempt makes no POST; a unique-value
  refusal from SharePoint is read as already-saved.
- the tablet never PATCHes or DELETEs either list, and never writes
  `Station targets`; a stage without `daySheet` draws no button and makes no
  request for these lists.
- week helper: Monday start, Sunday end, year boundary (2026-12-31 and
  2027-01-01 share a week; 2027-01-04 starts the next), local-date not UTC-date
  at 23:30.
- "This week N of T": sums this person's week, includes the typed draft, uses
  the live target; no target → the words, not "of 0".
- office: target upsert body and its log line; an edit PATCHes exactly the
  allowed columns; week subtotal uses the week's stored `WeekTarget`; filters
  by weekday and week; missing list → quiet state on both pages, board intact.
- `stationReport`: glass cutting with day sheets produces all five sheets;
  glass hotmelting omits day-sheet columns; welding produces Jobs per product
  group; an empty period produces a Summary that says so and no crash.
- rule 3: a note and a comment carrying a phone number in three written shapes
  and an eircode come out of the report stripped; the standing "no phone, no
  eircode in any export" test covers the new template; the export is logged.
- the workbook gate over the station files; all existing suites green with
  unchanged counts (`test_welding.js` included).

Add `test_daysheets.js` to the verification list in `web/CLAUDE.md`.

## Docs to update

`docs/STATIONS.md` (the two lists with an owner's creation recipe; the office
window; the rule-2 exception), `docs/REFERENCE.md` (a new section),
`docs/SUPPORT.md` (the cutter cannot change a saved sheet — the office can;
what "waiting to send" means), `docs/ARCHITECTURE.md` (`cw_daysheetdraft`, the
queue key if a new one is added), `web/CLAUDE.md` (rule 2 exception, the new
suite), `docs/specs/README.md`.

## Report back

Pasted suite output before and after; functions added with file:line; the
exact request bodies for a save, a target upsert and an office correction;
the localStorage keys added; how `stationReport` stays free of station-specific
branches (name the adapter, if any); anything not done as written, and why.
