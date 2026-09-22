# Station comments: the reading window, review bugs, and typing on the board

Continues [[station-comments]] (§20).

**The office holds a ninety-day window of it**, the same one `Station log` gets
and through the same function (`stationNotesRecent` → `stationLogRecent`),
applied in `readStationNotes` and in `feedRows` for both the full read and the
delta's first enumeration. A row with no stamp at all is kept, exactly as the
log keeps one. Nothing deletes from the list itself — this is a reading window,
not data loss — and `listItems` still follows `@odata.nextLink` only to
`LIST_PAGE_CAP` (50 pages of 999) before warning and handing back what it has,
so a list that ever grew past ~50,000 rows would be read in part. That is the
remaining future work, and it is years away at a few notes a day.

**Two bugs the second review found in the first fix pass, both now closed:**

- **The seen-ids cap used to re-announce old notes for ever.** `saveNotesSeen`
  trimmed to the newest 500 ids while `noteFloorNotes` walks *every* row the
  feed holds, so the moment the list passed 500 rows the oldest ids fell out of
  the memory, were announced again as new, and evicted a different block, which
  was announced again next pass — measured at exactly (rows − 500) false lines
  **per pass**, and `noteFloorNotes` runs on every successful delta (six times a
  minute for five minutes whenever a list has refused one). The 400-entry
  Changes panel would fill with ghosts and push every real change out. Fixed by
  making the cap a cap on **memory of rows that are gone**: an id still in the
  feed is never evicted, so re-announcing one is impossible, and the window
  above is what keeps that bounded.
- **A list that vanished after a good read reported healthy for ever.**
  `stationFull` returned `false` both for "nothing to do" and for "there is no
  such list", so a list renamed, deleted, or left behind in the old site when
  the others moved across (exactly the migration `docs/STATIONS.md` describes —
  and `Station comments` postdates those notes) left `STATION_NOTES_OK === true`
  with no clock armed: a fresh enumeration of a list that is not there every ten
  seconds, ~8,600 requests a day per open screen, a drawer rendering rows nobody
  could refresh, and no Changes line ever again. Fixed by making both
  `stationFull` and `stationDelta` answer **`null`** for "there is no such
  list" — which every existing caller reads as falsy exactly as before — and by
  the poll's notes branch degrading on it the same way `readStationNotes` does.

**A note with a blank `Station`** shows on every station's tablet, because an
empty string passes every station's filter — the same rule `logRows` has always
had. It is only reachable by hand-typing a row into SharePoint; hiding it
everywhere would lose a note somebody wrote, so it is deliberate. Recorded in
`docs/STATIONS.md`.

**A board reorder interrupts typing. Not solved — mitigated.** A card is moved
(`appendChild`) when another job finishes or the office marks one done, and
moving a node blurs what is inside it. `paintBoard` puts the caret back in the
note box afterwards, the same way `dressCard` does after a redraw, and the text
was never at risk (it lives in the draft, and the node is moved rather than
rebuilt) — but **a programmatic `focus()` does not re-open a tablet's on-screen
keyboard**, only a finger does, so the person has to **tap the box again to keep
typing**. Closing it properly means not re-appending a card whose composer is
open, which is a change to the board-diff logic and was not made. Told plainly
in `docs/SUPPORT.md`.

**Tests.** `test_comments.js` (44 checks): the row the composer builds; the
tablet's own-station thread and the drawer's cross-station one; oldest first on
both, and two notes in one second in the order they were written; a hand-made
row with only a Title; the card's button and its count; sending as one POST from
the right station and the right person; **two notes sharing a Title and both
surviving** — no upsert, no dedupe, nothing assuming uniqueness; a refused send
keeping the typing; **a note typed while the last one is still in the air never
being wiped by it landing**; typing never redrawing the card; every boundary of
the "how long ago" a thread line shows; a second station getting
the feature from one word; the drawer's section, read-only, with no control of
any kind; the Changes line in the owner's own words, **a note that arrived while
the dashboard was shut being announced on the next load and only once**, the
first look by a browser announcing nothing, and the seen-ids cap trimming its
map and its list together; **announcing a note calling neither `noteChange` nor
`CW.appendLog` and sending no write at all** (the real functions, counted, not a
string match); a missing list on the tablet (an explained line, no box, nothing
written) and in the drawer, with the floor's board untouched by it; **a list
that does not exist yet being asked about again and found the moment it is
made**; **a notes-list failure inside `stationPoll` reaching neither the board's
error line nor `glassColourRun`, and then backing off**; **a list longer than
the seen-ids cap announcing nothing twice however many passes run** (driven
through the real `noteFloorNotes`, since a synthetic `notesSeenMark` test cannot
see that failure mode); the ninety-day window; **a list that vanishes after a
good read degrading, arming its clock and then asking for nothing**; and
over the whole run — no workbook, no `/drive`, no DELETE and no PATCH of a note.

## See also

- [[station-comments]] — previous: station comments, a word from the floor to the office
- [[station-comments-amendment]] — next: amendment E, the job-row icon and board notes
