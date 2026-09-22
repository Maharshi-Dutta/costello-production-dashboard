# Row colour code and the John print sheet

## 16. Row colour code, section select, the John print sheet and print notes

Shipped 2026-09-09. Spec: [`docs/specs/2026-09-09-john-template.md`](specs/2026-09-09-john-template.md), including
its "Amendments after review and the owner's second round". Office page only —
the tablet is untouched by all of it. The owner named the view, the template
and the file **"John print sheet"**, and that name is his recorded exception to
the no-real-names rule.

**The row colour code (`parser.js`, `app.js`, `export.js`).** The Production
sheet says four things with the colour of a job row's *text*, and nowhere
else: red = urgent, green = booked (an exact date is agreed), pink/magenta =
trade order, blue = on hold. `fontOf()` reads the job-number cell's font
colour (theme colours resolve through the same palette and tint the fills use)
and falls back to the customer cell **only when the job number's own ink is
unset or plain black** (`blackInk`). A cell carrying a hyperlink, and theme 10
and 11 — Excel's own two hyperlink colours — are never read at all, because a
link is blue by default and would otherwise arrive as "on hold". `flagOf()`
maps the hex by **hue range, not exact value** — the office picks these by hand
and one row's red is `FF0000` while the next is `FF3300`. Saturation and
lightness gate the hue first, so grey, black and white say nothing, and
orange/yellow/violet fall through rather than being guessed at. The result is
`j.flag` (the word) and `j.flagHex` (the colour that said so). A red row is
urgent as well as the old rule (the word in a comment), so `j.urg` is now
"either". Shown as a chip in the row's last cell, beside the other badges (the
cell that already handles overflow), and in the drawer head — **always the
word**, with the colour only as its ink, clamped by `inkFor()` so it stays
readable in the dark theme. An exported file always keeps the sheet's own hex.
Production (2) has its own colours and is read for them separately; neither
sheet's colours are merged into the other.

**Section / category select (`app.js`).** In the grouped view every section
header carries a "Select all n" tick box (indeterminate when part of the
section is picked, and present whether the section is open or collapsed); in
the flat list a "Select all shown (n)" box appears at the left of the toolbar,
but only when something is actually narrowing the list. The John print sheet
view has the same per-section box. All of them go through `pickMany(ids, on)`:
one pass over the ids, one `renderAll()`, whatever the size of the section.

**Production (2), read on its own terms (`parser.js: parseJohnSheet`).**
Production (2) is not a lagging copy of Production: it is a separate sheet the
office keeps for the paper John works from, with its own rows, order, sections
and colouring. `parseJohnSheet(wb)` reads it into
`{id, section, ready, cust, phone, area, wnd, drs, notes, fillHex, inkHex, seq}`
— the fixed columns C/G/I/J/K/M/N, plus the "Notes from the colleague's office"
column found by its header (column BY on the real sheet, past a run of hidden
ones), the row's own solid fill and the row's own ink, and the sections from
its own divider rows. **Nothing from it is merged into the job model, and
nothing from Production is used to draw or print it.** `load()` keeps the
result in `JOHNROWS`; a workbook without that sheet gives `[]`, never null.

**The John print sheet view (`app.js`).** The Show dropdown (`BOARDS`) gains
"John print sheet" after Glass station. It renders in the job list's place:
the eight printed columns, a divider row per section, each row wearing its own
fill and its own text colour with the colour's word beside it, a tick box per
row and a "Select all n" per section, all narrowed by the header search box.
Ticks go into `state.picked` like anywhere else, so the wheel works unchanged.
Nothing in the view polls, feeds or writes anything — it is drawn from the
workbook download the page already made.

**The John print sheet template (`export.js` §7b).** `exportJohnRows(ids,
notes, ctx)` takes job *numbers* and builds the rows from `ctx.john`
(Production (2)) in that sheet's own order, with a divider row per section. A
chosen job that sheet does not have is still printed — from the Production
model, with **no colour at all** and a grey italic "not on John's sheet" line
in its Notes. `buildJohnWorkbook` makes a one-sheet ExcelJS workbook named
"John print sheet" (two-row header with QUANTITY over Wnd/Drs, the sheet's
widths, thin borders, landscape, fit to one page wide, print titles `1:2`) and
`buildJohnDoc` the A4-landscape pdfmake equivalent with page numbers and a
"Printed … by …" footer. File name `John print sheet <date>.xlsx` / `.pdf`.
Choosing **John print sheet** in the Export window hides the field picker and
the card/table choice and leaves the scope (ticked jobs / everything the view
is showing), the format, and a **Continue to notes** button. Opening the
window from the John view preselects that template — on the way in only, so
the choice stays clickable.

**This is the one export in the app that carries a phone number.** The owner
sanctioned it on 2026-09-09 for this template only; the number printed is
Production (2)'s own. `j.ph` (added to the job model) is read only by
`exportJohnRows`, and only for a job Production (2) does not have; the eircode
is carried by nothing anywhere. Every John print's `Dashboard Log` line says so
in words: `Excel · n jobs · John print sheet · with phone numbers`.

**Print notes (`app.js`).** One extra note per job, typed in the notes window
(`#nhost`, "Print notes — John print sheet") before a print: job number,
customer, colour chip, Production (2)'s own note in grey with the Production
comment labelled underneath when the two differ, and a box (2000 characters)
pre-filled from the list. **Print** clears last time's failure marks, asks for
the list permission **once** (`CW.listConsent()` from the click, before any
write — three writes starting together used to open three consent dialogs),
saves the notes that changed one after another through `CW.listUpsert` into
`Dashboard print notes` (Title = job number, `Note`, `By`, `At`, in the
workbook's own site), then builds the file from **exactly the jobs the window
listed** (`NSTATE.ids`, not a fresh lookup — a poll can land in between) and
downloads it. A note that will not save marks its own row "not saved" and does
not stop the print. Clicking the scrim with unsaved typing asks first; Cancel
and Escape do not. The list is read when the notes window opens, not on every
page load; the drawer shows the note read-only under Comments. **No workbook
sheet is written by any of this** beyond the one export log line, Production
(2) included, and the list is never created by code.

Tests: `test_john.js` (the flag parser and `parseJohnSheet` on a two-sheet
ExcelJS fixture with hidden columns, a divider, fills, four inks and a
hyperlink; section select in all three views; the view's rows, dividers,
search and tick boxes; Export preselecting the template from the view and
leaving it clickable; the file taking Production (2)'s phone, area, date and
note; a job missing from that sheet; one consent dialog for three changed
notes; printing `NSTATE.ids` after the job list changed underneath; the
cleared failure mark; and a sweep of every request the run makes) and the John
sections of `test_export.js`.

## See also

- [[glass-station-job-row]] — previous: the glass chip on the ordinary job list
- [[glass-colours]] — next: glass colours reaching the Production sheet
- [[export]] — the export builders this template sits beside
