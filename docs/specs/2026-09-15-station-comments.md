# Station comments: a note channel from any floor station to the office

**Status: approved to build, 2026-09-15** ("let's start developing the
feature," after the consult below).

## Context

The glass tablet can update a job's own progress, but a floor worker has no
way to leave a word for the office against a job — a shortage, a mistake on
the sheet, anything that needs a person's attention rather than a tick. The
owner wants this on the glass tablet now, and says plainly that more station
tablets are coming (cutting, fabrication, glazing, windows, doors) and every
one of them will want the same channel. This spec builds the channel once,
generically, so a future station page needs no new list and no new drawer
code — only its own station name.

## Owner consulted, 2026-09-15

Asked in chat: what I understood, three points needing a decision, and what
I'd default to if the owner did not correct me. The owner's reply was to
proceed ("let's start developing the feature"), which stands as agreement to
the three defaults below unless the owner says otherwise on review:

1. **Notification.** No new column on the job row (owner's rule from
   2026-09-09, `docs/HISTORY.md` A6). Default: a new comment is logged into
   the dashboard's existing Changes panel, the same surface that already
   reports other events — "New floor note on `<job>`, from `<station>`,
   `<who>`." No new UI element.
2. **Edit or delete.** Default: append-only, like `Station log`. Neither side
   can edit or delete a comment once sent. A mistake gets a follow-up note.
3. **A real name in the owner's own message must not enter the repo**
   (`CLAUDE.md` rule 4, public repo). Every mention of who receives these
   notes, in code, docs, specs, or commit messages, reads "the office." This
   is not a placeholder for a station worker's name (that pattern is "the
   cutter", "the colleague") — it names the reader, not a person on the
   floor, so it gets its own placeholder.

Direction is one-way for this build: floor writes, office reads. The owner's
own framing ("to contact the office or to notify something for the job")
describes exactly that; a reply channel back to the tablet is a plausible
fast-follow and is called out under "Not built" below, not attempted here.

## Data model

One new SharePoint list, in the same site the other floor lists live in
under the current interim arrangement (`docs/STATIONS.md` §"Interim: lists in
the workbook's site" — the owner has no Global Administrator yet, so this
follows `Glass station`/`Station people`/`Station log` rather than waiting
for a `Floor stations` site that does not exist).

### `Station comments`

Named to match `Station people` and `Station log` — both already describe
"any floor station," not just glass, so the name does not need to change
when a second station page ships.

| column | type | written by | meaning |
|---|---|---|---|
| Title | text | tablet | a unique row id, e.g. `<JOB>\|<unix-ms>` — this list is append-only, one row per comment, never upserted by Title the way the counters are |
| Job | text | tablet | the job number, upper-case |
| Station | text | tablet | which floor station wrote it: `Glass` today; `Cutting`, `Fabrication`, `Glazing`, `Windows`, `Doors` as those ship — free text, not an enum, so a new station is a new value, not a schema change |
| Who | text | tablet | the signed-in person's display name, exactly as already resolved for stage taps (`Station people`.`Title`) |
| Text | multiple lines of text | tablet | the note itself |
| At | text | tablet | ISO timestamp, stamped at send, matching the `At`-suffixed fields elsewhere in this codebase |

No `enforce unique values` on Title (unlike `Glass station`): this list is a
log, every row is a separate record, same as `Station log`.

**No phone numbers, eircodes, or other exportable personal data belong in
`Text`.** This list is never part of any export path (`CLAUDE.md` rule 3
governs exports specifically; this list is internal-only, read only inside
the drawer), but the note is free text a person types, so nothing stops
someone writing a customer's number into it by habit. `SUPPORT.md` gets a
line telling the office and the floor this channel is for job problems, not
customer contact details.

## Tablet (`glass.html` / `station.js`, generalised in `station-core.js`)

- A note icon on each job card opens a small composer: a text box and a send
  button.
- Under the composer, that job's own station's earlier comments show
  chat-style, oldest first, each line "`who` · `relative time`: `text`" — so
  a second shift does not retype the first shift's note. Only this station's
  own comments show here; a glass tablet does not see a cutting note and vice
  versa (the office drawer is where everything about a job comes together).
- Sending appends one row via the same `listAdd` primitive `station.js`
  already uses for `Station log` (`station.js:440`) — no new write path, no
  feeder involvement, consistent with rule 2 (the tablet writes its own
  lists directly; it is the workbook it must never touch).
- The composer, the thread render, and the list-read/write calls live in
  `station-core.js` behind one function taking a station name, so
  `glass.html` passes `"Glass"` and a future `cutting.html` passes `"Cutting"`
  and gets the same feature for free.

## Master dashboard (`app.js`, drawer)

- A "Floor notes" section in the job drawer: every station's comments for
  that job, oldest first, each line tagged with station, person, and time.
  Read-only — no reply box in this build.
- Read the same way `Dashboard print notes` and `Dashboard phases` are read
  today: on the drawer's own poll cycle, no new polling infrastructure.
- A new comment is also written to the Changes panel the moment the
  dashboard's poll notices it, per the notification default above.

## What is not built here

- **A reply from the office back to the tablet.** The list and the drawer
  are shaped so this is a small addition later (a `From` value of `"Office"`
  on a row, shown on the tablet's own thread), but no UI for it ships now —
  the owner asked for floor-to-office, not a conversation.
- **Read/unread state per comment.** The Changes-panel line is the only
  signal that something arrived; there is no per-comment "seen by the
  office" flag. If notes get missed in practice, this is the first place to
  extend.
- **A station page for cutting, fabrication, glazing, windows, or doors.**
  Only `station-core.js` is made ready for them; no new tablet page ships in
  this spec.

## Tests to deliver

- `test_comments.js`, offline, no network: composer produces the right row
  shape for a given job/station/person/text; the tablet's own-station thread
  filters correctly when the list holds rows from more than one station; the
  drawer's cross-station thread does not filter by station; Title uniqueness
  is never assumed (no upsert path exists for this list); a missing
  `Station comments` list is a quiet, explained state on both pages (pattern
  already used for `Floor stations` not existing, `docs/STATIONS.md`), never
  a crash.
- Extend `test_station.js` only if the composer touches shared render paths;
  otherwise keep this list's tests self-contained in `test_comments.js`.
- Grep gate as usual: no `setFill|clearFill|setValues|batchWrite|/workbook`
  in station files touched by this feature; no real name anywhere in the
  diff, including the one named in the owner's own consult message.

## What to report at demo

- Screenshots: the composer and thread on the tablet; the Floor notes section
  in the drawer; a Changes-panel line for a new comment.
- Confirmation that a job with comments from more than one station (staged in
  the fixture, since only Glass exists live) renders all of them in the
  drawer, tagged correctly.
- Confirmation that `station-core.js`'s new function takes a station name as
  its only station-specific input — the reviewer should be able to point at
  the one line a future `cutting.html` would change.

## For the owner: the list to create

Same process as every floor list so far — plain list, no template, in the
same site the other three floor lists are in:

1. Create a list named exactly `Station comments`.
2. Add the columns `Job`, `Station`, `Who`, `Text` (Multiple lines of text),
   `At` — all Single line of text except `Text`. `Title` already exists.
3. Do **not** turn on enforce-unique-values on Title — this list is a log,
   like `Station log`, not an upsert target like `Glass station`.
4. No rows to pre-fill; it starts empty.

This does not touch any existing list or the workbook, so there is nothing to
rehearse before creating it. Building can start now, against a test fixture,
while this is created — the live rehearsal step before ship needs the real
list to exist, the coding does not.

## Amendment E — the owner's additions, 2026-09-16 (approved: "yes")

Asked for after the real-browser demo of the build above, before the push.
Three changes, all on the master dashboard (`index.html` / `app.js`). The
tablet page, the feeder and the workbook are not touched; the office still
never writes `Station comments`.

### E1. The Components F·S·T column is emptied

The owner: "we dont need this bar for now, empty this space as no real info
comes from this, i might add something later." The job row's Components cell
(the three-colour `mini` bar and the count, `app.js` row markup) and the header
text `Components F·S·T` (`index.html` line ~684) are both blanked. The column
keeps its width so the layout does not move and the space is reserved. The
drawer's stat tile ("2 components"), the exports and the parser are unchanged.

### E2. A new-note icon on the job row

- A small speech-bubble icon with a count, shown on a job row only while that
  job has floor notes **not yet seen on this screen**. It lives in the badge
  cell (the cell that already holds the flag chip, "Urgent", "In fab", "N
  comments"), never in a column of its own — the owner's 2026-09-09 rule.
- **Hover** shows the unread notes in a tooltip, one line each:
  `Glass · Person A · Wed 08:58 · two DG units short, need more stock`. Use the
  native `title` attribute (the codebase's existing convention, e.g. the flag
  chip); no new tooltip component.
- **Click** opens that job's drawer and scrolls it to the Floor notes section
  (give the section an id or `data-` anchor and `scrollIntoView`). The click
  must not toggle the row's pick checkbox and must not start a drag.
- **Seen** means: this job's drawer was opened on this screen at or after the
  note's `At`. Opening the drawer by any route (icon, row click, jump link)
  marks every note of that job seen. The icon disappears on the next row
  paint (the existing `quietRows` / chips diff path — extend `chipsNow()` or
  its equivalent so a change in the unread set repaints the rows).
- **Per screen, not shared.** Store the seen state in localStorage, keyed
  `cw_notesread`, as `{ "<JOB>": "<ISO At of the newest note seen>" }` — one
  small entry per job rather than one per note, so it never needs the 500-id
  cap that `cw_notesseen` (the Changes-panel set) has. Keep it separate from
  `cw_notesseen`: that set answers "has the Changes panel announced this
  note", this one answers "has a person on this screen opened the job". The
  two must not be merged. Documented limit for SUPPORT.md: two office
  computers each show the icon until each opens the job.

  > **Amendment E, fix pass 2026-09-16.** Two bugs found in review, both
  > fixed before ship, and both change the paragraph above.
  >
  > 1. **The stored value is now an object, not a string:**
  >    `{ "<JOB>": { "at": "<ISO At of the newest DATED note seen>",
  >    "ids": ["<item id of each UNDATED note seen>"] } }`. A note with no
  >    `At` cannot be placed on a timeline, and the stamp standing in for one
  >    made every undated note typed *after* the job was first opened read as
  >    seen from birth — invisible to the office for ever. An id can say "and
  >    that one"; a stamp cannot. Dated notes still cost nothing per note, so
  >    the "no 500-id cap needed" reasoning holds: the only ids ever stored
  >    are those of rows somebody typed into SharePoint by hand. The key was
  >    never shipped, so there is nothing to migrate, but `notesReadRestore()`
  >    reads a bare string as `{ at: <string>, ids: [] }` so a browser that ran
  >    the build in progress does not throw.
  > 2. **Stamps are compared as moments, never as text.** The tablet writes
  >    milliseconds (`toISOString()`); a hand-typed row usually has none; and
  >    `"…10:00:00.100Z" < "…10:00:00Z"` as *text*, because `.` sorts below
  >    `Z`. An undated-looking stamp as the stored value therefore hid every
  >    note that followed it within the same second, permanently. One
  >    comparator, `ST.atCmp(a, b)` in `station-core.js`, now decides both
  >    "is this note newer than what I have seen" and the order the notes are
  >    drawn in on both screens; a stamp neither side can parse as a date
  >    falls back to the text comparison it always had.
- **First load on a fresh browser:** no seeding. Every note whose `At` is
  newer than the stored stamp for its job (or with no stamp at all) is
  unread. The list is new and empty, so there is no flood; and a screen that
  has never opened a job genuinely has not seen its notes.
- A note with no `At` (typed into SharePoint by hand) counts as unread until
  the drawer is opened, then as seen (store the current time).
- The icon is also drawn on the station board card head (E3) with the same
  rule and the same click.

### E3. Notes on the Glass station board

`stationBoardHtml()` (`app.js` ~5250): under each card's "last:" line, that
job's floor notes, oldest first, the same line shape the drawer uses
(`floorNotesHtml` — factor the per-note line out so both call it). Every
station's notes, not only Glass — the board is the office's view. Cards with
no notes show nothing extra. Read-only; no reply box. The three channel states
(checking / missing / unreachable) show nothing on the cards — the drawer is
where the explained state lives; do not add a banner to the board.

### Tests to add (`test_comments.js`)

- unread set: job with notes and no stamp → unread N; stamp older than one
  note → unread 1; stamp equal to newest → unread 0; note with no `At` →
  unread until opened.
- opening the drawer for a job stores the newest `At` under `cw_notesread`
  and the row's icon html is empty afterwards; a note arriving later makes it
  non-empty again.
- tooltip text carries station, who, time and text for each unread note and
  is HTML-escaped.
- E1: the row html contains no `mini` bar and no F/S/T count; the header cell
  text is empty; the drawer tile still says "N components".
- E3: the board html for a job with notes from two stations carries both
  lines, oldest first; a job with no notes carries no note markup.
- The existing gates keep passing: no request from `app.js` ever POSTs,
  PATCHes or DELETEs `Station comments`; no `/workbook` request from the
  station files.

### Report back

Test output for all 12 suites and the new checks; the exact lines changed in
`index.html`; confirmation that `glass.html`, `station.js`, `station-core.js`
(other than a shared helper, if any) are unchanged; which localStorage key
was added and what it holds.
