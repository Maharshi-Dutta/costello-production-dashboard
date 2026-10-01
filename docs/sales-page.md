# 27. The Sales page

Brief: [`specs/2026-09-30-sales-page.md`](specs/2026-09-30-sales-page.md). Built 2026-09-30.

**What.** A page for the two sales people, `sales.html`, with the office
dashboard's look: the same header, tiles, list and drawer, but customer
columns (Job no, Customer, Address, Eircode, Phone, Office no, Wnd/Drs,
Status, Delivery). On top of the office's read-only drawer it adds: the six
customer cells, Urgent / Booked (the row's text colour), a delivery date, a
move once the office has marked the job ready, delete with restore, and
requests to the office. Floor boards are view only. Both people sign in with
one shared account and pick their name; every write is logged under it.

**How - one app, one switch.** `sales.html` sets `window.CW_PAGE = "sales"`
before `app.js` loads. `isSales()` in `app.js` is the one test; every
difference goes through it, as a hook into `sales.js` (loaded by
`sales.html` only) or a branch:

| hook in `app.js` | in `sales.js` |
|---|---|
| `whoAmI()` | `salesWho()` - the picked name |
| `start()` | `salesStart()` - name picker, Requests / Deleted jobs buttons |
| `load()` after `applyPending` | `salesOverlay()` - our writes held 180 s, deleted rows hidden |
| `renderTiles` / `renderChips` | `salesTiles` / `salesChips` |
| `renderRows` (the list) | `salesRowHtml` |
| `renderDrawer` | `salesDrawerHtml` / `salesWireDrawer` (the office sections follow, `ed = false`) |
| `readSalesLists()` | `salesRepaint()` |

The rules and every write are in `sales-core.js` (`SALESC`), loaded by both
pages, tested by `test_sales.js`. The workbook primitives are in `graph.js`:
`salesLocate`, `salesSetCell`, `salesSetFont`, `salesDeleteRow`,
`salesInsertRow`. Every one of them, and every `SALESC` write, throws unless
the page is the Sales page, before anything reaches Graph.

**What stands down on the Sales page.** `cpWritable()` is false (no tick, no
door, no gold, no import, no adoption, no repaint, no colour writers);
`markReady` throws; the four feeders, the board edit functions, phases,
alerts (`isAdmin()` false), day sheets, print notes, saved views and
rollback all return at once; `setFill`/`clearFill` throw. The office's load
chain (import, safeguard, feeders, colour writers) is not run at all. The
boards' write controls are removed before wiring (`viewOnlyBoard`). No
multi-select, no drag and drop, no grouped views.

**After the review (fix pass, same day).** The locate reads `A1:K600` (the same
600 rows `moveJobRow` sees, not `rowForJob`'s 400) and refuses unless the
job is on the sheet exactly once; the row located for the backup is the row
written or deleted, or nothing happens. The backup's Section is the live
section. All six customer cells are written as text through `formulas` with
the apostrophe marker (writeRow's route), after the column's header is read
and still matches; phone and eircode go into the Dashboard Log as `•••` plus
the last three characters. Delete and restore are logged the moment the row
goes or comes back; a failed border step is reported, not a failed write.
The Sales move runs in `SALESC.serial` with every write, and `SALESC.busy()`
holds every write control on the page while anything is in flight. A pink
(Trade order) or blue (On hold) row asks before it is replaced. Nothing is
written or exported before a name is picked. The office's reply reads the
request first and refuses if it has been answered.

**Amendments after the demo (A-C, same day).**
- *A:* the View select offers the flat list and the sheet-order grouped view
  (sections in sheet order, collapse, Select all per section, the
  Categories show/hide filter); rows carry a tick box; "N selected - move
  to...", "Export selected" (the export window on the ticked jobs) and the
  selection wheel. Sales columns in both views. No saved views or
  categories are written.
- *B:* a drag onto a section, the "move to" menu and the wheel's Move go
  through `SALESC.moveMany`: same gate and targets as the drawer's Move,
  same queue; a job not `j.done` is refused and named in the toast
  ("the office has not marked this job ready - send a request"), and
  nothing moves for it. The office page's drag and move are unchanged.
- *C (both pages):* `renderDrawer` leaves the DOM alone when the HTML is
  identical, and patches `.dhead`/`.dbody` inside the open drawer for the
  same job (no slide or fade replayed, scroll kept); `paintRows` skips
  identical HTML for the list, boards and tiles; a background re-read
  (`load` after the first, `stationAfterFeed`) draws quietly;
  `stationFull` and the welding/glazing/fabrication full reads report a
  change only when the rows differ (`sameRows`). A click on a control, a
  change, or a drag in the list/drawer forgets the last HTML so the next
  redraw always goes through.

**The writes** (spec "Writes, exactly"):
- *Customer cell* - locate by column C, check the one cell, PATCH the cell
  (phone and eircode with Excel's apostrophe text marker, the same one
  `writeRow` uses, so a leading 0 stays), `noteChange`, then a `Kind = edit`
  backup item.
- *Text colour* - locate and check, capture the row, add a `Kind = colour`
  backup with `Row` and read it back, then PATCH `A{r}:CL{r}/format/font`
  `{color}` after checking column C again, `noteChange("Text colour")`.
- *Delete* - locate and check, capture, `Row` JSON (empty entries dropped,
  refused over 60,000 characters - a real 90-column row is about 15,000),
  add `Kind = delete`, read back identical or stop, re-locate and re-check,
  delete the row, restore the bottom edge above it as `moveJobRow` does,
  `noteChange("Deleted")`.
- *Restore* - from the Deleted jobs window: refused if the job is on the
  sheet or the section has no divider/jobs; insert at the section's
  `last + 1`, `writeRow`, verify exactly one copy there or take the inserted
  row out again; mark the backup Restored, add `Kind = restore`,
  `noteChange("Restored")`.
- *Move* - `moveJobsInSheet`, refused unless every job is `j.done` and the
  target is not In production or Can sell as second hand.
- All `SALESC` writes run one at a time (`SALESC.serial`).

**Lists** (workbook's site): `Sales people`, `Sales jobs`,
`Sales job backups`, `Sales requests` - columns in the brief. A missing list
gives a plain line and no write. `Sales requests` and `Sales jobs` are read
every 30 s on both pages while visible (`readSalesLists` in `app.js`).

**The office gains** a bell, "Sales requests N" (N unanswered; hidden while
the list is missing), a window with unanswered first, a reply box (writes
`Reply`, `ReplyBy`, `ReplyAt`) and Open job, and "Delivery date: Wed 07 Oct
&middot; set by ..." at the top of the drawer. Nothing else.

**Reports, complaints and photos (2026-10-01).** Brief:
[`specs/2026-10-01-sales-reports-complaints-photos.md`](specs/2026-10-01-sales-reports-complaints-photos.md).
One form (`salesSendForm` in `sales.js`), opened from the drawer (job filled
in), from "Send to office" in the top bar and from the Requests window: type
Status / Move / Report / Customer complaint, job no. (optional for a report or
complaint, then a Customer box), text, up to 8 photos. Each photo is shrunk in
the browser (`createImageBitmap` + canvas, longest side 1600, JPEG 0.82,
`SALESC.fitSize`); a non-image or a file the browser cannot decode (HEIC) is
refused. `SALESC.sendMessage`: add the item (`Photos` 0, `Status` Open for a
complaint), upload the photos one by one to the `Sales photos` library
(`CW.salesPhotoPut`, `PUT /sites/{site}/drives/{drive}/root:/<folder>/<n>.jpg:/content`,
`conflictBehavior=replace`, numbered past the highest file already in the
folder, folder = the item's Title with `|`/`:` made `-`), then PATCH `Photos`
to the folder listing's count; a failed photo leaves the message and the
toast says "N of M photos failed - send them again from the message", and
"Add photos" on the person's own message (`SALESC.addPhotos`) sends more. One
`Dashboard Log` line per message: kind, job, photo count, never the text.
Complaints: `SALESC.setComplaint` (both pages) reads the item first and refuses
if it already says Resolved (or Open), naming who and when; writes `Status`,
`ResolvedBy`, `ResolvedAt`; logged with `noteChange("Complaint")`. The office
bell counts unanswered messages plus open complaints (`SALESC.bellCount`); open
complaints come first, red. Thumbnails (`CW.salesPhotoList`, Graph
`thumbnails`, click opens `webUrl`) are read only while a requests window is
open, once per photo count. The five new columns (`Customer`, `Status`,
`ResolvedBy`, `ResolvedAt`, `Photos`) are looked for with `CW.listColumns`;
until all five are there they are neither read nor written, no photo is sent,
and the form and window say so. At 400 px the form and the Requests window are
one column with 40 px tap targets. Photos and message text never reach an
export.

**Exports.** On the Sales page the Default Excel template carries Phone no
and Eircode beside Customer (`xpSales()` in `export.js`); the office's is
unchanged. The PDF layouts are unchanged on both pages. Still one
`Dashboard Log` line per export.

**Tests.** `test_sales.js` (move gate, urgent/booked, Row JSON size, delete
order, restore, customer cells, nothing runs off the Sales page, requests);
`test_export.js` (Sales columns); `test_pages.js` covers `sales.html`.

## See also

- [[sections-and-row-moves]] - `moveJobRow`, whose steps the delete and
  restore follow
- [[export]] - the export and its phone/eircode rule
