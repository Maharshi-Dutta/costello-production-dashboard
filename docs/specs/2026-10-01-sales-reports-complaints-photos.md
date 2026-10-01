# Sales: reports and customer complaints with photos

Status: briefed 2026-10-01, owner approved. Builds on the Sales page
(`docs/specs/2026-09-30-sales-page.md`, `docs/sales-page.md`); read both first.

## Owner's answers (2026-10-01)

- Sales send reports and customer complaints **through the dashboard only**.
  There is no email and no Power Automate alert.
- A report is **free text plus photos**, with no fixed fields.
- Photos can be uploaded from a PC (file picker). On a phone, the same picker
  offers the camera. The send form and the Requests window must work at
  phone width.
- **Both the office and Sales** can mark a complaint Resolved.

## What to build

### Data

**`Sales requests` list** (workbook's site), new columns:
- `Customer`: text, for a complaint about a job no longer on the sheet.
- `Status`: `Open` / `Resolved`. Only used for Kind `complaint`.
- `ResolvedBy`, `ResolvedAt`
- `Photos`: number, how many photos were uploaded.

The session adds these columns by script before the push. The code must
tolerate them missing: show a quiet note and do not send photos or status.

**Kind gains `report` and `complaint`.** It is a plain text column, so no
schema change is needed for that.

**A new document library `Sales photos`** in the workbook's site, created by
the session's script.
- Each message gets one folder named after its list item Title, with `|`
  and `:` made safe.
- Upload through Graph drive: `PUT /sites/{site}/drives/{drive}/root:/<folder>/<n>.jpg:/content`.
  The page's existing `Files.ReadWrite.All` scope covers it, so **no new
  scope**.
- Find the library's drive by name, the way `listId` finds lists.

### Sending (Sales page)

- The request form in the drawer and a new **"Send to office"** button in
  the top bar open one form:
  - **type**: Status / Move / Report / Customer complaint
  - **job no.** Prefilled from the drawer. Optional for Report and Complaint;
    then a Customer field shows instead.
  - **text**, required.
  - **"Add photos"**: `<input type=file accept="image/*" multiple>`. At most
    8 photos. Thumbnails before sending, each with a remove button.
- **Shrink each photo in the browser** to a longest side of 1600 px, as JPEG
  quality 0.82, using canvas. A file that isn't an image, or fails to decode
  (HEIC on some browsers), is refused with a plain message.
- **Send order:**
  1. Add the list item, with `Photos` = 0 and `Status` = Open for a
     complaint.
  2. Upload the photos one by one.
  3. PATCH `Photos` to the number that actually landed.
  4. If an upload fails, the message stays sent and the toast says
     "N of M photos failed — send them again from the message".
  5. Add a "Add photos" action on an existing message, so a failed photo can
     be sent again.
- Run the send through `SALESC.serial` and require a picked name, the same as
  every other Sales write. Log one `Dashboard Log` line per message: kind,
  job and photo count, never the text.

### Reading and resolving (both pages)

- **Office:**
  - The `Sales requests` bell counts unanswered messages **plus open
    complaints**.
  - In the window, open complaints come first, marked red "Complaint".
  - Each message shows its text, photo thumbnails (Graph `thumbnails`;
    click opens the full image in a new tab via its `webUrl`), job link or
    customer, reply, and **Mark resolved** / **Reopen** for complaints.
- **Sales:** the same thread, with the same Mark resolved / Reopen, and
  "Add photos" on their own messages.
- **Resolving:** PATCH `Status`, `ResolvedBy` and `ResolvedAt`.
  - Re-read the item first. If someone already resolved it, show who and
    when, and do nothing.
  - Logged with `noteChange`.
- **Thumbnails** load only when the window is open. Don't poll the library.

### Phone width

At 400 px wide, the send form and the Requests window must work: one column,
no sideways scroll, tap targets ≥ 40 px. The rest of the Sales page may stay
desktop-shaped.

## Hard rules

- Nothing new is written to the workbook.
- Photos and message text never go into any export.
- No real names anywhere in the repo.
- Office-page changes are limited to the requests window and the bell.
- Tablets are untouched.
- Edit with Edit/Write only. Don't push. Don't touch live SharePoint, since
  the tests use fakes.
- Keep `test_pages.js` and the global-scope rule in mind (HISTORY B29).

## Tests

`test_sales.js`:
- the resize maths (longest side becomes 1600, aspect kept, small images not
  upscaled)
- refusal of more than 8 photos and of non-images
- **send order:** the item is added first; a failed upload leaves the
  message, with `Photos` = the number that landed
- resolve is refused when already resolved
- the bell count includes open complaints
- no name means no send

Browser rig (scratchpad, extend `sales_check.js`):
- send a complaint with 2 photos
- the office sees it red with 2 thumbnails
- the office resolves it, and Sales sees it resolved
- Sales reopens it
- a 400 px screenshot of the form

## Report back

Diff summary, the verification command output, `test_sales` output, rig
results, screenshots (office window, Sales form at desktop and at 400 px), and
any deviations.
