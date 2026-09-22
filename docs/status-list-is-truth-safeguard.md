# Status list: the switch-on window, the import, and the hand-paint safeguard

Continues [[status-list-is-truth]] (§19).

**The switch-on window.** The import takes fifteen to twenty-five minutes on
the real sheet, and until it has drained an item with no row would read as
"nothing done" — which would collapse every job's phase, make the feeder write
`OfficeDone` twice per job as rows land, zero the floor's seeds and print "all
checkpoints are complete" over a gold row's zeroes. So while it is outstanding
(a per-browser marker, `cw_cpimported`, set the first time the import plan
comes back empty), an item **with no row** answers from the sheet's own colour
and every screen reads exactly what it read before the switch-over. An item
that has a row always answers from the row, so an un-tick made inside the
window is still absolute. Once the marker is set the fallback is unreachable.
An unreadable list falls back the same way.

**The one-time import.** A job+item with no row and a colour in the sheet
adopts that colour once: `Source = "import"`, `Done` = the total for gold, the
`Dashboard Progress` **sheet's** own count for yellow (the last thing that
sheet is for) and nought when it has none, `Who = "the sheet"`. Lazy, capped at
60 rows a load over three lanes (the feeder's numbers), with a 30 s follow-up
when the cap cuts it short, and one summary `Dashboard Log` line per load
("Imported 40 checkpoints from the sheet's colours"). **A blank cell is
deliberately never imported**: no row already means nothing done, and importing
every blank would be a row per checkpoint of every job on the sheet.

**The safeguard: a colour changed in Excel by hand is adopted (spec §4a).**
There is no waiting period. `PAINTED[job][item]` (in `localStorage`, key
`cw_painted`) remembers the colour this dashboard last painted into every
managed cell — which is **every checkpoint item the drawer shows**, not only
the four glass columns. On every download, a managed cell whose parsed colour
differs from `PAINTED` is a *candidate*; each candidate is confirmed with one
Excel API read of that cell's fill (`range/format/fill`, batched 20 per
`$batch` through `CW.batchGet`, plus one values read per job to prove the row
has not moved), and the API reads the live file with no lag. API agrees with
the download → a real outside change, adopted at once (`Source = "excel"`,
`Who` = the file's last editor, one log line naming the item and saying
"adopted from Excel"). API says what we painted → the download was stale,
ignored silently. Capped at 60 candidates a download. Once the import has
drained, a cell with **no** row and a colour is a candidate too — that is a
hand-paint, not something the import missed. **The open drawer's job is checked
on the 10 s pass**, at most every 30 s, and only for cells the download already
disagrees about: the download DISCOVERS, the API CONFIRMS. A drawer left open
on a job nobody is touching sends no request at all. The colour writer records
its own paints in `PAINTED` too, or its work would read as somebody else's.

**Putting the Excel copy right.** It covers the four glass columns too since
step 3 — they are ordinary managed cells now. A fill the workbook refuses leaves the record
saying one thing and the sheet another, and the safeguard cannot see it — it
compares the file with `PAINTED`, and after a refused fill those agree. So each
load also compares the **record** with `PAINTED` (both in memory, no read) and
repaints the difference through the ordinary fill path, on the job's own
`cpChain`, capped at 20 cells and never looping inside one load. `PAINTED` only
moves when a fill lands, so one success ends it. It never concludes from the
**download** that the sheet is already right: the download is ~36 s old and
this runs precisely when something did not land, so a stale file agreeing with
the record proves nothing (review finding F1). A cell carrying a colour this
feature does not own — the sheet's own Cut green — is left alone, in either
direction.

**Honest limits.** A hand-paint on the job in the open drawer no longer shows
in ten seconds; it shows on the same ~40–50 s as everywhere else, which is what
the review traded for a drawer that costs nothing to leave open. Excel is now a second or two behind the dashboard rather
than the other way round. Hand-painting a checkpoint cell no longer *sets*
status, it is *adopted* as one. The list grows (about 30 rows per job that has
work started on it) and pruning is a later chore. The import trusts today's
colours once. A first click on an item that has no row yet costs one full read
of the list (`listUpsert`, for the dedupe); every click after it is one PATCH.
Step 3 closed the last of it: the glass colour writer writes the record, so the
floor's work reaches the drawer's glass checkpoint rows, `OfficeDone` and the
sheet from one place.

**Tests.** `test_checkpoints.js` (53): the switch-on window and its end; the record decides and the colour does
not; eight stale gold downloads leaving every checkpoint white with no
`PENDING` entry; a refresh; the write order record→fill→log for an item and for
a group; a refused record write and a refused fill, told apart; a missing list;
the import, its cap and that nothing is imported twice; the safeguard adopting
a confirmed hand-paint, ignoring a stale download, never adopting our own
paint, the drawer's live check, and 25 candidates in whole batches.
`test_glasscolour.js` (76) rewrites its hold section as the same properties
without a hold. the repaint after a refused fill; an open drawer costing nothing; a background
write never asking for consent; and a group write that fails half way naming
both halves. `test_station.js` (246) proves `OfficeDone` follows the record and
does not flap during the switch-on window.
Reproduced in a stubbed browser with `scratchpad/untick_repro_reload.js` +
`stub_office_reload.js`, cases R1–R8.

## See also

- [[status-list-is-truth]] — previous: status lives in a list, the Excel colour is a copy
- [[station-comments]] — next: station comments, a word from the floor to the office
