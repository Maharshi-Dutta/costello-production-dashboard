# Station comments: a word from the floor to the office

## 20. Station comments: a word from the floor to the office

Built 2026-09-15. Spec: `docs/specs/2026-09-15-station-comments.md`.

**What.** A floor worker can leave a note against a job — a shortage, a mistake
on the sheet, anything that needs a person's attention rather than a tick — and
the office reads it in that job's drawer. One-way in this build: the floor
writes, the office reads. **Append-only on both sides**: nothing edits or
deletes a note, and no affordance for either exists in the code.

**The list.** `Station comments`, beside the other floor lists (the interim
arrangement of [[glass-station-overview]] (§11) applies unchanged: the workbook's own site until a
`Floor stations` site exists). One row per note — `Title` (`JOB|unix-ms`, a row
id and **not** a key: no unique rule, nothing looks a row up by it), `Job`,
`Station`, `Who`, `Text`, `At`. Named like `Station people` and `Station log`
because it is every station's, not the glass station's. Never created by code;
a missing list is a plain, explained state on both pages and nothing is written
anywhere.

**Built once, for every station.** The whole channel — the composer, the thread,
the row builder and the two list calls — is `ST.stationComments(cfg)` in
`station-core.js`, and **`cfg.station` is its only station-specific input**.
`station.js` passes `ST.STATION_NAME`; a future `cutting.html` passes
`"Cutting"` and has the feature, with no new list, no new column and nothing to
add to the office's drawer. The Graph calls are injected (`cfg.listItems`,
`cfg.listAdd`), so `station-core.js` still knows nothing about Graph and the
tests need no network.

**The tablet** (`station.js`, styles in `glass.html`). A note button at the foot
of each job card, carrying the count of **this station's** notes on that job.
Open, it is that station's own thread — oldest first, "who · how long ago ·
what they said" — a box and Send. A glass tablet never shows a cutting note:
the office drawer is where a job's notes come together. Sending is one `listAdd`
to `Station comments`, the same primitive and the same directness as a
`Station log` line — no feeder, no upsert, no id looked up, and **no workbook of
any kind** (the grep gate over `station.js`/`station-core.js`/`glass.html` is
asserted as a test, not only run by hand). A refused send keeps the typing in
the box and says so in red; it can never take the board away. A note typed while
the last one is still in the air is kept rather than wiped, and the box says
which it is — *"sent — what is in the box is a new note"* — because writing left
behind on purpose otherwise looks exactly like a note that failed, and the
obvious answer to that (tap Send again) would post the first note twice. The draft is
deliberately **not** part of the card's repaint signature — a card rebuilt on
every keystroke is a caret lost on every keystroke — and `dressCard` puts the
caret back when something else redraws the card mid-sentence. The list is read
once at start-up and then only while a composer is actually open, at most every
`COMMENT_POLL_MS` (20 s): a tablet nobody is writing on sends no request for it.

**The office** (`app.js`). `Station comments` is a fourth entry in
`STATION_FEEDS`, polled by delta on the floor's own clock (10 s while somebody
is looking at the floor or has a glass job's drawer open, 60 s otherwise) — the
spec's "the drawer's own poll cycle, no new polling infrastructure". Read at
every load (`stationAfterFeed`) rather than only when a drawer opens, because
the Changes line is the whole notification. `floorNotesHtml(j)` draws a **Floor
notes** section in the drawer: every station's notes on that job, oldest first,
tagged with station, person and time, read-only, for **every** job and not only
a glass one. **No new column on the job row** (owner's rule, 2026-09-09) and no
new window. A note nobody has seen puts one line in **Changes** — "New floor
note on `<job>`, from `<station>`, `<who>`", `src: "floor"`, which renders as
its own **Floor** badge and its own entry in the source filter. The first read
of the list is a baseline and announces nothing, so a page opening on two
hundred old notes posts no lines. **Nothing about a note is written into the
workbook** — not a cell and not a `Dashboard Log` line: the note is already
recorded in the floor's own list, and a log line per open dashboard would record
the same note once per screen.

**Not built** (spec's own list): a reply from the office back to the tablet;
read/unread state per note; any station page but glass.

**Honest limits.** A note on a job with no glass on it reaches the office on the
slow poll (up to a minute), because the fast rate is decided by whether anybody
is looking at the floor's work.

## See also

- [[status-list-is-truth-safeguard]] — previous: the switch-on window, import and the safeguard
- [[station-comments-details]] — next: the ninety-day window, review bugs, board reorder
- [[station-comments-amendment]] — amendment E: the job-row icon and board notes
