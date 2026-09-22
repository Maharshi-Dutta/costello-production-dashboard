# Floor notes on the job row and the board (amendment E)

### 20a. Floor notes on the job row and the board (amendment E, 2026-09-16)

Asked for after the real-browser demo of [[station-comments]] (§20), before the push.
Spec: `docs/specs/2026-09-15-station-comments.md`, Amendment E. Three changes,
all in `index.html` / `app.js`; the tablet, the feeder and the workbook are
untouched, and the office still never writes `Station comments`.

**E1 — the Components F·S·T cell is blanked.** The owner: "we dont need this
bar for now... i might add something later." The row's Components cell and
its header text are both emptied in `rowHtml`; the column keeps its width so
nothing else on the row moves. The drawer's own "N components" tile, the
exports and the parser are unchanged.

**E2 — an unread-notes icon on the job row.** `notesIconHtml(job)` draws a
💬-and-count badge in the row's badge cell (never a column of its own — the
2026-09-09 rule) for whatever that job's floor notes nobody has opened on
*this screen* yet, with a native `title` tooltip listing each one (station,
who, when, text). Empty string, not a placeholder, the moment nothing is
unread. `wireNotesIcons(scope)` binds its click and swallows the drag-start so
the badge never starts dragging the row underneath it; the click calls
`openJobNotes(id)`, which selects the job, opens the drawer and
`scrollIntoView`s the `#floornotes` section. "Seen" is decided by
`notesUnreadMap()`, one entry per job in memory (`NOTES_UNREAD`, rebuilt by
`notesUnreadFresh()` once per paint rather than once per row), against what
is stored in localStorage under `cw_notesread` and loaded by
`notesReadRestore()`. The stored shape is `{ "<JOB>": { at: "<ISO of the
newest DATED note seen>", ids: ["<item id>", ...] } }` — an id for every
*undated* note (typed by hand into SharePoint) because a stamp cannot say
"and that one, too", only "and everything before this moment". Opening a job
calls `markNotesRead(job)`, which folds every note's `at` or `id` into that
entry, saves it, and returns whether anything changed. Every stamp comparison
anywhere in this feature — newest-seen, note ordering — goes through
`ST.atCmp(a, b)` (`station-core.js`): parses both sides as dates and compares
the moments, falling back to a plain text compare only when one side will not
parse. Comparing the raw ISO strings as text was the first review bug (see
`docs/HISTORY.md`): a stamp with milliseconds text-sorts *before* the same
instant without them, because `.` is below `Z`.

**Unread notes repaint rows without touching the floor's poll rate.**
`chipsNow()` builds one comparison string per paint from two independent
halves: a `glass` half (each row's station counters) that also sets
`CHIPS_GLASS` — the only thing that decides whether the floor's lists get
polled at the fast 10 s rate — and a `notes` half (each row's unread count,
from the same `notesUnreadFresh()` call the icon uses) that feeds only the
diff `ROWS_CHIPS` compares against to decide whether a repaint is owed. A note
arriving, or a job being opened on this screen, changes the `notes` half and
so repaints the rows; it never changes `CHIPS_GLASS` and so never speeds up
the poll a note nobody asked about.

**E3 — every station's notes on the Glass station board card.**
`floorNoteLineHtml(r)` is the one line-drawing helper — station badge, who,
`stWhen(r.at)`, the text — used by both `floorNotesHtml(j)` (the drawer's
existing Floor notes section, now with a `#floornotes` anchor id for E2's
scroll) and the new `floorNotesCardHtml(job)`, which prints the same lines
under a board card's "last:" line for every station, not only Glass, oldest
first. A card with no notes prints nothing; the channel's checking/missing/
unreachable states are never shown on the board — that explained state stays
in the drawer only, so four hundred cards do not each grow a banner.

## See also

- [[station-comments-details]] — previous: the reading window, review bugs, board reorder
- [[welding-station]] — next: the welding station, the second floor page
