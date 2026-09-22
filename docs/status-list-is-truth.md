# Status lives in a list; the Excel colour is a copy

## 19. Status lives in a list; the Excel colour is a copy

Built 2026-09-11. Spec: [`docs/specs/2026-09-11-status-list-is-truth.md`](specs/2026-09-11-status-list-is-truth.md). This
is **step 2 of three** in that spec: checkpoints read and write the list, and
the glass colour writer is not changed (step 3 does that). Read the spec
before touching any of it.

**The problem it answers.** The dashboard never read the truth, it read
copies: a download of the workbook about 36 s behind, and a poll of the
floor's list about 10 s behind. A tick never makes those copies disagree, so
it looked instant; an un-tick makes every copy disagree with the office for
about a minute. Four mechanisms existed only to keep believing the office over
the copies — the `PENDING` hold, the 45 s reconcile, the colour writer's
last-writer-wins and the feeder's `OfficeDone` derivation — and each of them
failed in turn in the week of 2026-09-08. **No mechanism can make a copy
true.** So status stopped being read from a copy.

**The rule.** The SharePoint list `Dashboard progress` — in the workbook's own
site, one row per `JOB|ITEM`, Title the dedupe key, columns
`Title, Job, Item, Done, Total, Status, Who, When, Source` — is the single
record of checkpoint status. The Excel colour is written **from** it so people
can see it in Excel, and is never read back to decide anything. **One rule the
office must accept: checkpoint colours are set in the dashboard, not by
painting cells in Excel** — and a hand-painted cell is adopted rather than
obeyed (the safeguard, see [[status-list-is-truth-safeguard]]).

**What reads what now.**

| asked | answered by |
|---|---|
| `itemState(j, item)` / `cpStatus(j, item)` — the status and the count | `cpRow(job, item)`, the list. Return shape unchanged, so the drawer, `cpSummaryHtml`, `cpInProgress`, `glassCounts`, `exportCheckpoints` and the export's checkpoint filter all follow it with no change of their own |
| `cpFileStatus(j, item)` — the **parsed colour** | `j.cp`, which the parser still fills exactly as before. Three readers only: the phase pipeline's Cut green, the one-time import, and the safeguard |
| `jobPhase(j)` | the list for done/process, `cpFileStatus` for the sheet's own Cut green — nothing writes that colour, so the list can never carry it |
| `glassCounts(j)` → `ST.officeSeed` / `ST.officeComplete` → the feeder's `OfficeDone` | the list, with the same Cut-green exception |
| `glassCellNow(j, type)` — the colour writer's idempotence test | still `j.cp.glass`, deliberately: it asks what the **cell** is showing, not what is done |

**The write, in order: record, fill, log.** `cpSaveRow` puts the row in this
dashboard's copy of the list at the click (so the row moves before the request
goes out), then writes it — `listPatch` on a row whose id is known, otherwise
`listUpsert`, which dedupes on Title. Then `cpWriteItem`/`cpWriteGroup` paints
the cell and writes the `Dashboard Log` line, exactly as they did. A refused
**record** write is a click that did not take and the row goes back; a refused
**fill** leaves the record standing and says so — the office decided, and Excel
is behind, which is the direction this change chose.

**No holds.** `PENDING` no longer carries a `cp` entry: the map updated at the
click is the optimistic state and the next 10 s delta confirms it. `PENDING`
still holds row moves, mark-ready, phases and the colour writer's own `gc`
paints (step 3 retires the last of those). What the `cp` hold also did — void
the writer's un-landed paint of a cell the office has just acted on, the fix of
the morning of 2026-09-11 — is now an explicit `glassVoidPaint()` at the click.

**The floor's work goes down the same path (step 3).** `glassColourRun`
watches the floor's list, derives the colour with `ST.glassColours`, and writes
one record row per DG/TG/TUFF/NOT TUFF item with `Source = "floor"`, `When` =
the floor's own stamp and `Who` from the row's `DoneBy` — then paints Excel
from it and writes the `Floor glass colours` log line **between the record and the
fill**, so a colour that was recorded reaches Changes even when the workbook
refuses the paint. Last-writer-wins is the row's `When` against
`ST.floorStamp(g)`, and only an `office` or `excel` row can block; ties go to
the office. One floor row can block a floor write: its own, when it is newer
than the copy of the list in hand, which is a list read that has gone
backwards. If the record already says what the floor
says, nothing happens at all, which is what lets the writer be called six times
a minute.

**Kept current.** A full read at boot (only when the list has not answered
yet), then `listDelta` every 10 s on its own clock (`cpTick`, `CP_POLL_MS`) and
again 1.5 s after this dashboard's own write. It is a third entry in
`STATION_FEEDS` and uses the floor's own delta machinery — the token, the 410
resync, the five-minute "delta off" — but on its own timer, because the floor's
poll drops to a minute when nobody is looking at the floor and a colleague's
tick has to arrive in ten seconds whatever is on screen.

**Missing list.** No list is created by code. `listItems` answering null puts
`CP_LIST_OK = false`, the drawer's Checkpoints section says which list is
missing and every control on it is dead, and nothing is written anywhere —
not the list, not the workbook, not the log.

## See also

- [[glass-office-clear-history]] — previous: the bugs found on the way here
- [[status-list-is-truth-safeguard]] — next: the switch-on window, import and the hand-paint safeguard
- [[checkpoints-and-phases]] — the checkpoint drawer this list stands behind
