# The Sales page

Status: briefed 2026-09-30. The owner approved the design mock-up ("looks good")
and the defaults listed under "Decisions taken".

## Context

The office dashboard (`index.html` + `app.js`) shows the `Production` sheet of
the shared workbook as a live job list with a job drawer, station boards and
exports. The sales team (two people) want **their own page with the same look
and layout as the office dashboard**. The office list columns are replaced by
customer columns. They get a small set of edits the office page does not
have: customer details, the row's text colour, a delivery date, delete with
restore, and a request/reply channel to the office.

Both sales people sign in with **one shared Microsoft account**, which already
has edit access to the workbook's site (the owner confirmed this 2026-09-30).
The shared account's address and domain must never appear in the repo. The
page tells the two people apart by a name picker (below). The names live
only in a SharePoint list.

## Decisions taken (owner, 2026-09-30)

1. **A separate page, `sales.html`**, sharing `app.js` and every other script
   with the office page. It is not a floor station.
2. **Rule 1 gains three sanctioned writes to `Production`, from the Sales page
   only:**
   a. **Row text colour.** Set the font colour of the whole row, `A{r}:CL{r}`:
      - red `#FF0000` = urgent
      - green `#00B050` = booked for delivery
      - black (automatic) = neither
      Fills are never touched by this write.
   b. **Customer cells.** The six identity cells of the job's own row:
      Customer, Phone no, Area, Eircode, Office no and Windows colour. Columns
      are located by header label through the parser's `m.ident` map, never by
      fixed position.
   c. **Delete a whole row**, only after a full copy has been saved to
      `Sales job backups` and read back. A **restore** puts the row back at the
      bottom of the section it came from.
3. **Rule 3 gains an exception:** exports made on the Sales page include Phone
   and Eircode. Office exports keep the rule unchanged, and every export is
   still logged.
4. **Sales never paint a fill.** No checkpoint tick, no door cell, no mark
   ready, no gold. Marking a job ready (the whole row gold plus the move to
   Ready to fit) stays the office's action.
5. **Moving a job:**
   - Sales can move a job to another section **only once the office has
     marked it ready**, i.e. `j.done` is true.
   - Until then the Move control is disabled, and a **Send a request**
     button is shown instead.
   - The allowed targets are every section except In production and Can sell
     as second hand.
6. **Urgent and Booked** can be switched off again (text goes back to black).
   The two are exclusive, because a row has one text colour.
7. **The delivery date** is stored in a SharePoint list, never on the sheet.
   It is shown on both pages.
8. **Floor boards** (Glass, Welding, Glazing, Fabrication) are **view only** on
   the Sales page.
9. **The office sees what Sales did**: every Sales write is logged with
   `noteChange` into `Dashboard Log`, with Who = the picked name. The office's
   drawer shows the delivery date.
10. **Deleted jobs window**, on the Sales page only. It lists every delete,
    restore and customer-detail change. Restore is available to anyone on the
    Sales page (the sales team, or the owner opening the page).
11. **Requests.** Sales send a question about a job, such as status, timing
    or "please move". The office gets a bell with a count of unanswered
    requests, reads them in a window and replies. Sales see the reply, with a
    count of unread replies.
12. **Delete asks for one confirm** (an in-page dialog, not `window.confirm`).

## Hard rules for the implementer

- **Nothing is written to `Production` except by writes 2a, 2b and 2c,** and
  only from the Sales page. The Sales page also uses the existing move
  (`moveJobsInSheet`), and only when gated as in decision 5. A mode check must
  guard each new write, so the office page can never reach them.
- The office page gains **only** the bell, the requests window with its
  reply, and a read-only delivery date in the drawer. It gains no delete, no
  customer edit and no text-colour write.
- **Checkpoint writes are off in sales mode.** `cpWritable()` returns false.
  The Edit toggle and mark ready are hidden. The board steppers
  (`glassOfficeEdit`, `weldOfficeEdit`, `glzOfficeEdit`, `fabrOfficeEdit`) are
  neither rendered nor wired. Also hidden or disabled: Versions/rollback,
  Categories editing, Alerts editing, saved-view saving, multi-select move and
  drag-and-drop onto a section.
- Tablet pages and their js files are not touched.
- No real names or addresses anywhere: code, tests, docs or commit. Use
  "Person A" and "Person B" in fixtures, and `example.test` for addresses.
- Edit files with Edit/Write only. Never use a shell heredoc or python to
  write a file.
- Do not commit, and do not touch live SharePoint or the workbook.

## Mode

`sales.html` sets `window.CW_PAGE = "sales"` in an inline script before
`app.js` loads. A helper `isSales()` in `app.js` reads it. The office page
leaves it unset. Every branch in `app.js` goes through `isSales()`. Put the
sales-only code in a new `sales.js`, loaded only by `sales.html`, and its pure
logic in `sales-core.js` (tested by node). Keep changes to `app.js` to hooks
and branches. **Watch the shared global scope** (HISTORY B29): classic scripts
share one scope, so a second top-level `const G` breaks the page. Add
`sales.html` to `test_pages.js`.

## Who is using it

After sign-in on the Sales page:
- If `localStorage.cw_salesperson` is empty, or not an active name in
  `Sales people`, show a full-screen picker with one button per active name.
  There is no PIN.
- The name button in the top bar re-opens the picker.
- `whoAmI()` returns the picked name on the Sales page, so every log line
  carries it.
- If the list is missing, the picker says so and the page stays read-only.

## Lists (workbook's own site, the default `listSiteId`; the session creates them by script)

| list | columns | written by |
|---|---|---|
| `Sales people` | Title (the name), Active (Yes/No) | the owner, by hand |
| `Sales jobs` | Title = job id (unique), DeliveryDate (text `YYYY-MM-DD`), SetBy, SetAt | Sales page |
| `Sales job backups` | Title (`JOB\|ISO time`), Job, Kind (`delete` / `restore` / `edit` / `colour`), Section, Field, From, To, Row (multi-line text: the captured row as JSON), Who, At, Restored (Yes/No), RestoredBy, RestoredAt | Sales page; nothing ever deletes an item |
| `Sales requests` | Title (`JOB\|ISO time`), Job, Kind (`status` / `move` / `other`), Text, From, At, Reply, ReplyBy, ReplyAt, ReplySeen (Yes/No) | Sales page (request, ReplySeen) and office page (Reply*) |

A missing list gives a quiet explained state and no write, the same as the
day-sheet feature. Use the existing helpers: `listItems`, `listAdd`,
`listPatch`, `listUpsert` and `listItem`.

## Sales page UI

Same header, tiles, list and drawer shell as `index.html`. Copy the markup;
CSS may be shared by a link or duplicated, whichever keeps both pages simple.

- **Top bar:** brand "Sales"; search `#q` (job no, customer, area, eircode,
  phone digits, office no); Requests (count of unread replies); Deleted jobs;
  Export; Refresh; theme; the name button.
- **Tiles:** All jobs, Urgent, Ready to fit (`j.done`), Booked
  (`j.flag==='booked'`), In production.
- **List columns:** Job no, Customer, Address (area), Eircode, Phone (full
  `j.ph`), Office no, Wnd/Drs, Status pills (Urgent, Ready to fit, Booked,
  otherwise the section word; plus "Request open" when a request is
  unanswered), Delivery date.
  - The row text is red when `j.flag==='urgent'` and green when
    `j.flag==='booked'`. The rest of the row styling copies the office row.
- **Drawer, in order:**
  1. **Customer form** with the six fields and "Save customer details".
  2. **Status:**
     - an Urgent toggle and a Booked toggle
     - a delivery date `<input type="date">`
     - a note: "only the text colour changes"
  3. **Move:**
     - a section select and a Move job button when `j.done`
     - otherwise the button is disabled, with the note "the office has not
       marked this job ready" and a Send a request button
     - Delete job…
  4. **Requests for this job:** a thread and a send form with a kind select
     and text.
  5. **The office drawer's own sections, read only.** These are
     `cpSectionHtml` with `ed=false`, the dates, the counts, glass, the
     station lines and floor notes. Reuse `renderDrawer` and branch it; do
     not copy it.
- **Deleted jobs window:** a table with Job, What happened, By, When, Changes
  (the edit/colour items for that job), and a Restore button on
  not-yet-restored delete items.
- **Requests window:** every request with its reply, newest first. Opening
  it marks the shown replies `ReplySeen = Yes`.

## Office page additions

- **Bell button** "Sales requests N": N is the number of requests with an
  empty Reply. It is hidden while the list is missing.
- **Window:** unanswered requests first, a reply textarea with a Reply button
  (writes Reply, ReplyBy = `whoAmI()`, ReplyAt), and Open job.
- **Drawer:** "Delivery date: Tue 07 Oct · set by <name>" when present.
- **Polling:** read `Sales requests` and `Sales jobs` every 30 s while the
  page is visible, on both pages. `listItems` is fine; the lists are small.

## Writes, exactly

All writes are serialised: one at a time per page, reusing the pattern the
page already uses.

**Customer cell.**
1. `CW.rowForJob("Production", id)`.
2. Re-read that row's column C and refuse if it is not this job.
3. Write the one cell.
4. Write Phone and Eircode **as text**, so a leading `0` is never lost: set
   the cell's `numberFormat` to `@` in the same PATCH, or use the proven
   equivalent the graph layer already uses.
5. Log with `noteChange(id, "<Field>", from, to)`.
6. Add a backup item with Kind `edit`, Field, From and To.

**Text colour.**
1. Locate and check the row as above.
2. Add a backup item with Kind `colour`, Row = `CW.captureRow` of the row,
   From and To.
3. Only after that item has been read back, PATCH
   `range(address='A{r}:CL{r}')/format/font` with `{color}`. Black means
   `#000000`.
4. Log with `noteChange(id, "Text colour", from, to)`.

**Delete.**
1. Locate and check the row.
2. `cap = CW.captureRow(r, tmpl)`.
3. `listAdd` a backup item: Kind `delete`, Section = `BLOCKNAMES[j.blk]`,
   Row = JSON of `{cap, sectionName}`.
4. `listItem` must read it back with an identical Row. If the read-back fails,
   stop and delete nothing.
5. Re-locate the row, since the sheet may have moved meanwhile, and re-check
   column C.
6. POST `rows(n:n)/delete {shift:"Up"}`, then `restoreBottomEdge` on its
   neighbours, exactly as `moveJobRow` does.
7. Log with `noteChange(id, "Deleted", section, "")`.

Factor the insert/delete/verify steps out of `moveJobRow` into small graph.js
helpers **only if** that keeps `moveJobRow`'s behaviour byte-for-byte.
`test_move.js` must stay green unchanged.

**Restore.**
1. Refuse if the job id is already on `Production`.
2. Locate the section by name. If it no longer has a divider or any job,
   refuse with a clear message, as `moveJobRow` does.
3. Insert a blank row at the section's `last + 1` and `writeRow(cap)`.
4. Verify the id now appears exactly once. On failure, delete the inserted
   row.
5. PATCH the backup item: Restored = Yes, RestoredBy, RestoredAt.
6. Add a Kind `restore` item.
7. Log with `noteChange(id, "Restored", "", section)`.

**Row JSON size.** A multi-line text column holds about 63,000 characters.
Drop null and empty entries from `cap` before stringifying. If the result is
still over 60,000 characters, **refuse the delete** with a message. A test
must show a real-width row (90 columns with formulas, fills and fonts) fits
the limit.

**Export (Sales page).** Add Phone (`j.ph`) and Eircode (`j.eir`) columns to
the default template, gated by `isSales()`. The office path must not change,
and `test_export.js`'s standing rule-3 test stays as it is. Add a test that
the office path still has neither column and the Sales path has both. Still
log one `Dashboard Log` line per export.

## Tests to deliver

A new `test_sales.js`, run in node against `sales-core.js`, plus a fake Graph
for the write order:
- the move gate and the allowed targets
- the urgent/booked exclusivity and colour values
- the Row JSON size limit (a 90-column real-shape row fits, and an oversize
  one is refused)
- **delete order:** no row delete when `listAdd` fails, or when the read-back
  differs, or when column C no longer matches
- **restore:** refused when the id is already on the sheet, and refused when
  the section is gone; the inserted row is removed when verification fails
- customer write: refused on a row mismatch; Phone keeps its leading zero
  (text format sent)
- `isSales()` false: none of the new write functions can run (each one
  throws or returns without calling fetch)
- the request unread count and the ReplySeen marking

Also: `test_pages.js` covers `sales.html`; `test_export.js` gets the Sales
column test; the root verification command gets `sales-core.js`, `sales.js`
and `test_sales.js`.

## Docs to update in the same change

- `web/CLAUDE.md`: rule 1 gets the three Sales writes (dated 2026-09-30,
  owner's decision); rule 3 gets the Sales export exception.
- A new topic note `web/docs/sales-page.md` linked from `REFERENCE.md`.
- `STATIONS.md` does not change.

## Report back

- a diff summary with file:line
- the verification output (tail lines)
- the output of `node test_sales.js`
- the greps proving:
  1. the new write functions each check `isSales()`
  2. no `setFill` or `clearFill` is reachable from `sales.js`
  3. no real names
- anything you could not do, or any place you deviated from this brief
