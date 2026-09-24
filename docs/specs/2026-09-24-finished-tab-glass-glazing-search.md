# Finished tab on the glass and glazing tablets; search across both tabs

Status: shipped 2026-09-24, build 20260924-1248.

Follows `2026-09-24-welding-finished-tab.md` (shipped, build 20260924-1057).

## Owner's decisions (2026-09-24)

1. The glazing and glass tablets get the same **On floor · N / Finished · N**
   tabs as the welding tablet, remembered on the device; the collapsed
   Finished group at the bottom goes. Finished holds this tablet's finished In
   production jobs AND every job still on the sheet in a later section, with
   the section name on the card head.
2. Glass cards in Finished behave exactly as they do today: tappable unless
   the office has locked the job (`OfficeDone`); when the office clears a job
   that is In production it goes back to On floor. No new rule.
3. **Search, all three tablets (welding too):** typing searches both tabs. If
   the tab you are on has no match and the other tab does, the tablet switches
   to that tab and shows the job (typing 3708 on On floor lands on Finished
   showing R3708). If both tabs match, it stays and shows a tappable line
   "N more in Finished" / "N more on floor" that switches. Clearing the search
   returns to the tab you were on before you started typing.
4. No remake on glass or glazing. No "Sent to floor" button anywhere.
5. Glass and glazing ship together with the search change in one push.

## Hard rules (unchanged)

- Tablets never touch the workbook.
- Portrait 10-inch, one column, never scroll sideways (700-1100 px).
- No real names. Rule 3 strip on anything free-text fed to a list.
- `Active` on `Glass station` keeps its meaning exactly: "In production". The
  office Glass station board, the job drawer, the OfficeDone lock and the glass
  colour writer keep reading it as today.

## Part A - glazing tablet + search (files: glazing.js, glazing-core.js,
glazing.html, welding.js, welding-core.js if needed, test_glazing.js,
test_welding.js)

`Glazing station` already carries every job on the sheet (Active = on the
sheet) with `Section`. Mirror welding: `glzTabs(items)` -> {floor, finished},
one pass over `glzOfficeBoard`; floor = any row In production && card not
finished; else finished. Header numbers unchanged. Section badge. Tab key
`cw_glztab`. Then the search behaviour (decision 3) on welding and glazing.

## Part B - glass tablets (files: station-core.js, station.js, glass.html,
app.js feeder call only, test_station.js, test_glasscolour.js if touched)

Data model: two new Single-line-text columns on `Glass station`, added by
script BEFORE the push (the feeder refuses to feed when a column it writes is
missing):
- `Section` - the job's section on the sheet (display name, as welding).
- `OnSheet` - "Yes" when the job is still on the sheet (cat != past), else "No".

Feeder (`glassSlice` / `feederFields` / `feedPlan`):
- also feeds jobs with glass (`glassTotal > 0`) in any live section, not only
  In production; `Active` = In production as today; `Section`, `OnSheet` on
  every fed row.
- a row whose job has left the sheet: `Active = No`, `OnSheet = No` (today it
  only sets Active = No). Never deletes.
- seed for a newly fed later-section row: `officeSeed` from the office record,
  exactly as for a new In production row; the "only raises a seed on an
  untouched row" rule stays.
- must not trip the colour writer: an untouched row (DoneAt empty) never
  repaints; check the gates in `glassColourPlan` still hold for rows that are
  Active = No / OnSheet = Yes.
- office Glass station board (`jobBoard` = Active) unchanged.

Tablet (`station.js`): On floor = Active (In production) && not finished at
this page's stage (existing `g.finished` for the page); Finished = finished
here, or OnSheet = Yes && not In production. Rows with OnSheet = No never show.
Tab key `cw_glasstab`. Search behaviour as decision 3. Lock (`officeDone`)
behaviour untouched.

Report: count of rows the first feed after the push would POST and PATCH,
computed with `glassSlice`/`feedPlan` against the local workbook copy
(`C:\Costello Windows\Excel Dashboard\2026 Production work in progress Jan 2026
oakley.xlsx`) and an empty / current-shaped list, so the owner is told before
the push.

## Tests to deliver

- glazing: `glzTabs` split incl. mixed-section job; header unchanged.
- welding/glazing/glass: pure search helper (which tab to show for a query,
  the "N more" count) tested in node.
- station: feeder emits later-section rows with Section/OnSheet, Active = In
  production; leaving the sheet -> Active No + OnSheet No; FEEDER_FIELDS
  includes the two new columns; tab split per stage; colour writer does not
  paint an untouched later-section row.
- Browser rigs for each tablet: tabs, section badge, search auto-switch,
  no sideways scroll 700/900/1100.
