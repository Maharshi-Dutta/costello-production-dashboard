# Section select, the John print template, print notes and row colours (brief, 2026-09-09)

Office (master) dashboard only; the tablet is untouched. Hard rules as in
`CLAUDE.md`, with **one owner-sanctioned exception** recorded here: the John
template carries the customer's phone number (never the eircode). No other
export may. No commit, no push, no network, no live systems; all nine suites
green with `set -o pipefail`; no real person's name, address or domain in the
repo (placeholders; tests use `Customer One`, `Person A`, `example.test`).

## What the owner decided (2026-09-09)

1. **Select a whole section or category** and act on it with the wheel
   (Export, Move to, Alert to).
2. **A "John" export template** that reproduces the sheet "Production (2)"
   exactly as it prints for John: columns Job no · Ready to print · Customer ·
   Phone no · Area · Wnd · Drs · Notes from Brendan's office; sheet order;
   the section divider rows kept; every row's **fill** and **text colour**
   reproduced. Phone numbers are allowed in this template only.
3. **A print-notes window before a John export**: one row per selected job
   (job, customer, colour chip, the existing note, a box for more), scrollable,
   pre-filled from what already exists, then Print. Notes are kept by the
   dashboard, never written to the workbook. John only for now.
4. **The row colour code is read and shown.** Text colour of the job row in
   the Production sheet: red = Urgent, green = Booked (exact date), pink =
   Trade order, blue = On hold. Row fill as today (yellow = in fabrication,
   gold = ready to deliver). The Production sheet is the source; Production (2)
   is a copy and is not read for colours.

## 1. Section / category selection (`app.js`)

- Grouped view (View = the colleague's sheet-order view, or any custom view):
  each group header gets a tick box "Select all n". Ticking it selects every
  job in that group that is currently shown; unticking clears them. Partial
  state shows as indeterminate.
- Flat list: whenever a filter narrows the list (a summary tile, a category,
  a search, a Show/Sort that hides rows), a "Select all shown (n)" tick box
  appears at the left of the toolbar next to the count; it selects every job
  in `filtered()`.
- The wheel and the existing "n selected" chip work unchanged on the result.
  Selecting 200 jobs must not stall: reuse `state.picked` as today and render
  once.

## 2. Row colour flag (`parser.js`, `app.js`, `export.js`)

- `parser.js`: when parsing the Production sheet, read the **font colour** of
  the job row (take the job-number cell; fall back to the customer cell if the
  job-number font is black) and map: `FF0000`-ish red → `urgent`, `00B050` /
  green → `booked`, `FF3399` / `FF00FF`-ish pink/magenta → `trade`, `00B0F0` /
  `0070C0` blue → `hold`. Use hue tolerance, not exact matches (document the
  ranges in a small table in the code). Store as `j.flag` (`""` when black).
  Also store the raw hex as `j.flagHex` for the exports. Theme colours
  (`theme` index without an rgb) map through the workbook's theme palette the
  way fills already do; if unknown, `""`.
- `URG_RE` (the word "urgent" in the comment) still sets urgency; `j.flag ===
  "urgent"` counts as urgent too (`j.urgent` true either way).
- Job list: a small chip after the job number with the word (Urgent / Booked
  / Trade / On hold) in the flag's colour, never colour alone. Drawer header:
  the same chip. Default export: a "Flag" column when the field is chosen.

## 3. The John template (`export.js`, Export window in `app.js`)

- Export window gains a **Template** choice at the top: Default (as today) /
  **John**. Choosing John hides the field picker and the card/table choice
  (the layout is fixed) and shows: the scope (selected jobs or filtered), the
  format (Excel / PDF), and a "Continue to notes" button instead of Download.
- `exportJohnRows(jobs, notes, ctx)` (pure, in `export.js`): jobs in sheet
  order (`j.blk`, then row order as parsed), grouped by section with a divider
  row carrying the section name (as the sheet does: "Supply and collection"
  etc.), columns exactly: Job no, Ready to print (date, `dd-MMM`), Customer,
  Phone no, Area, Wnd, Drs, Notes. Notes = the existing sheet notes (Comment
  column + "Notes from Brendan's office", joined with " · ") followed by the
  dashboard print note if any.
- Colours per row: fill from the job's status (in fabrication → `FFFF00`,
  ready to deliver → `FFE699`, else none) and text colour from `j.flagHex`
  (black otherwise). Divider rows: bold, light grey fill.
- **Excel** (ExcelJS): one sheet "John", the eight columns with widths close to
  the sheet's (11, 10, 20, 14, 12, 6, 5, 60), a two-row header like the
  sheet's (QUANTITY over Wnd/Drs), row fills and font colours, borders thin,
  landscape, fit to one page wide, repeat header rows, print title. File name
  `John print <date>.xlsx`.
- **PDF** (pdfmake): A4 landscape, the same table with the header band, row
  fills and coloured text, page numbers, and a footer "Printed <date> by <who>".
  File name `John print <date>.pdf`.
- **Phone numbers** appear only in this template. `test_export.js`'s standing
  "no phone/eircode" checks stay as they are for the Default template; add
  tests that the John template carries the phone and never the eircode, and
  that a Default export still carries neither.
- Log line per John export: `exportLogFrom` gains the template name and the
  words "with phone numbers".

## 4. Print notes (`app.js`, `graph.js` list plumbing)

- SharePoint list **`Dashboard print notes`** in the workbook's own site (the
  manager creates it; code never creates lists): Title = job number
  (unique), `Note` (text, multi-line), `By`, `At`. Read with the existing list
  functions (`listItems` with `opts.fields`), written with `listUpsert`
  (dedupe on Title). Never written to any workbook sheet.
- The **notes window** (`#nhost`, same window style as Alerts/Export): title
  "Print notes for John", the selected jobs in sheet order, each row: job
  number, customer, colour chip, the existing sheet note(s) in grey (read
  only), a text box pre-filled with the dashboard note, "cleared" if the job
  has none. Scrolls. Buttons: **Print** (saves every changed note to the list,
  then builds the file with the notes just typed, then downloads) and Cancel.
  Saving uses `listUpsert` per changed job, three at a time; a failure shows
  the row's "not saved" and still lets the export go ahead with the typed
  text. Keyboard: Tab between boxes, Escape cancels.
- The drawer gets a "Print notes" line under Comments showing the dashboard
  note and who/when; read only there.
- The notes list is read once on the first John export (and again when the
  window opens), not on every load.

## 5. Tests

Extend `test_export.js` (John rows, order, dividers, colours, phone yes /
eircode no, Default still clean, log words) and add `test_john.js` (offline,
same pattern as `test_phases_list.js`): the flag parser on a small ExcelJS
workbook with coloured fonts incl. a theme colour and a near-red; section
select in the grouped and flat views (`state.picked` outcome and the count);
the notes window pre-fill (sheet note + list note), a changed note is upserted
and an unchanged one is not, a failed save still exports, the log line; the
sweep that nothing in the change writes any workbook sheet (no
`setFill/clearFill/setValues/appendLog` beyond the one export log line, no
`/workbook` PATCH) and that `Dashboard print notes` is written only through
`listUpsert`. Keep every existing suite green.

## 6. Docs

`docs/REFERENCE.md` §10 (export) and a new short §16 (John template, notes,
flags), `docs/USER-GUIDE.md` §3 (section select, template choice, notes
window, colour chips), `docs/ARCHITECTURE.md` (the notes list, the flag),
`docs/specs/README.md`. Placeholders only.

## Report back

Files changed; full pasted suite output; one paragraph per numbered owner
decision saying where it is done and which test covers it; anything unsure.

## Amendments after review and the owner's second round (2026-09-09, manager)

The owner confirmed the colour flags on the master view (they come from the
Production sheet and are correct) and accepted the name: the view, the template
and the window are called **"John print sheet"**; file names `John print sheet
<date>.xlsx/.pdf`. This is the owner's named exception to the no-real-names
rule, recorded here.

### A. The John print sheet view (new)

1. Production (2) is a **separate sheet** for John, maintained on its own. The
   parser gains `parseJohnSheet(wb)` → rows in that sheet's own order:
   `{id, section, ready, cust, phone, area, wnd, drs, notes, fillHex, inkHex,
   seq}` read **only from "Production (2)"**: job no (col C), Ready to print
   (G), Customer (I), Phone (J), Area (K), Wnd (M), Drs (N), "Notes from
   Brendan's office" (the column whose header contains "brendan"), the row's
   solid fill and the job-number cell's font colour (customer cell fallback),
   the section from that sheet's divider rows. Nothing from this sheet is
   merged into the Production job model, and nothing from Production is used
   to draw this view.
2. `STATIONS`/Show dropdown gains **"John print sheet"** after Glass station.
   It renders in the job list's place: the eight columns, section divider
   rows, each row with its own fill and text colour (the word chip too, so
   colour is never the only signal), a tick box per row, "Select all n" per
   section, and the search box narrowing rows. Ticks go into `state.picked`
   like any other view, so the wheel works unchanged.
3. **Export from either view.** From the John print sheet view the Export
   window opens with the John template preselected. From the master (the
   colleague's view or flat list) the John template is chosen as today. In
   both cases the John template prints **from Production (2)'s rows** for the
   selected job numbers (colours, notes, ready date, area, phone from that
   sheet); a selected job that is not on Production (2) is printed from the
   Production model with no colour and a grey "not on John's sheet" note.
4. The notes window's grey "existing note" is Production (2)'s Brendan note
   (plus the Production comment if different, labelled). Print notes stay in
   `Dashboard print notes`; nothing is written to Production (2).

### B. Review fixes

5. **Blocker:** `johnPrint` calls `CW.listConsent()` once from the click
   before `savePrintNotes` when `PRINT_NOTE_OK !== true`; a failure falls
   through and the print still goes. Fix the comment about "three lanes"
   (`listUpsert` is serialised per list).
6. The flag chip on the master list moves beside the existing badges in the
   last cell (which already handles overflow); the drawer keeps the full word.
7. `johnDownload` prints exactly `NSTATE.ids` (the jobs the window showed),
   never a fresh `xpJohnJobs()`.
8. `johnPrint` clears `NSTATE.failed` for the jobs being printed before the
   save pass.
9. `fontOf` skips a cell with a hyperlink and never maps theme 10/11
   (hyperlink colours) to a flag; the customer-cell fallback applies only when
   the job-number font is black or unset.
10. `readPrintNotes` sets `PRINT_NOTE_OK = false` on the early-return branch;
    the drawer copy says "read when the print-notes window opens".
11. Collapsed-section "Select all" stays (owner's choice); the label keeps
    the count.
12. `exportJohnRows` guards a non-numeric `blk`; a `maxlength` of 2000 on the
    note box; clicking the scrim with typed notes asks before discarding;
    `SUPPORT.md` says the Note column must be plain text, not enhanced rich
    text; on-screen chip ink is clamped for contrast in dark theme while the
    file keeps the sheet's own hex.

### C. Tests

`test_john.js`: `parseJohnSheet` on an ExcelJS fixture with hidden columns,
a divider row, fills, red/pink/blue/green fonts and a hyperlink; the view's
rows and dividers; select-all per section in the view; export from the view
preselects John; the John rows come from Production (2) data (a job whose
Production (2) phone/area/notes differ from Production prints the (2) values);
a job missing from (2); the consent-once test (a fake MSAL that counts popups:
exactly one for three changed notes); print-from-`NSTATE.ids` when `ALL`
changed; the cleared failure mark; hyperlink not a flag. Keep every suite
green; update USER-GUIDE (the new view), REFERENCE §16, ARCHITECTURE,
specs index.
