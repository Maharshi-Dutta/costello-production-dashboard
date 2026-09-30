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
