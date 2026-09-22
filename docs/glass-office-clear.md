# An office clear reaches the floor's counters

## 18. An office clear reaches the floor's counters

> **CHANGED 2026-09-11 by [[status-list-is-truth]] (§19), step 3.** The clear itself is unchanged — seven
> fields, the confirmation, the four locks, the owed retry. What is gone is
> everything this section says about the *contest* around it: `OFFICE_FLOOR_AT`
> (the clear zeroes the counters, so the colours the writer derives from them
> already say what the office's own rows say), `officeSettling` and the
> stand-down window, `glassOfficeJobStamp`, and `glassVoidPaint`. An office
> click now simply writes a row stamped at the click, and a floor colour older
> than that row never reaches the sheet. The "known gap" about a second
> dashboard inside the download lag is closed by the same thing: the record is
> shared, so there is no second dashboard that cannot see the office's action.

**Spec.** [`docs/specs/2026-09-10-office-clears-the-floor.md`](specs/2026-09-10-office-clears-the-floor.md). Changes
`CLAUDE.md` rule 3, by the owner, on 2026-09-10.

**The bug.** The owner spent an afternoon on what looked like three faults and
was one gap. Un-ticking a job's glass in the office cleared the **workbook**
correctly every time — the cell really did go white — but nothing ever cleared
the **floor's** counters. They stayed at 49/49/49, the tablet stayed gold for
ever, and the only way back was somebody tapping `−` forty-nine times. Two
earlier fixes in this area shipped and neither addressed it; the master
dashboard showing gold for up to a minute afterwards is SharePoint's ~36 s
download lag and is not this.

**What happens now.** When the office **clears a job's glass checkpoints** —
the group **Clear**, or a per-item clear that leaves no other DG/TG ticked —
the office writes that job's `Glass station` row's `Cut`, `Hotmelt`
and `Tuff` to **zero**, plus `DoneBy`/`DoneAt` and `OfficeDone = "No"`. Six
fields, and nothing else. (`Glazed` was a fourth nought until 2026-09-21; the
clear leaves whatever is in it, like everything else about glazing — [[glass-two-stage]], §22.) The unlock rides along so the **card frees itself on
the tablet's next ten-second poll** (owner, 2026-09-10: it used to sit greyed,
"the office has marked this job finished", for over a minute, because only the
feeder released the lock and the feeder derives it from what the master
currently shows). `OfficeDone` is a feeder column — the office's own, which the
tablet can never write — so nothing about rule 3 moves.
`DoneBy`/`DoneAt` are written on purpose: last-writer-wins reads
`ST.floorStamp` off them, so zeros under a stale stamp would read as old news
and the tablet's "last touch" line would name the wrong person. The per-stage
`By`/`At` pairs are never written — they say who did that stage's work, and
nobody did. No `Station log` line: it is an office action and goes in
`Dashboard Log`, as `Floor glass counters`.

**The trigger is a transition, not a state.** `officeGlassEmpty()` asks
`ST.officeSeed` — the very function that seeds a row from the office's record —
whether that record now says nothing of the job's DG and TG is done;
`officeClearsGlass()` fires only when it *newly* says so. A tick made when the
record already said nothing is not a clear, so clearing a hand-ticked ARCH on a
job whose DG is blank cannot wipe the floor's counters. There is deliberately
**no reconciliation pass**: a state-driven version of this would zero any row
the office happens not to have ticked, including a job the floor is working on
right now.

**A row the floor never tapped is left alone.** Its counters are the office's
own seed echoed back, and the **feeder** puts them right on its next run
(`sliceHash` carries the seed, so that is the next load). Writing `DoneAt`
there would mark the row touched for ever and switch its seeding off — a worse
bug than the one being fixed. This is a deliberate narrowing of the brief,
which wrote unconditionally.

**The confirmation.** When the row has a floor stamp and any counter above
nought, the office is asked first, naming what will go: *"The floor has
recorded 49 cut, 49 hotmelted, 12 tuff on this job. Clearing the glass here
will set all of those back to zero. Clear it anyway?"* (`ST.clearWarning`, a
pure function). Answering **no writes nothing at all** — not the floor's
counters and not the workbook half, because the question is asked before
`pend()` and before the burst.

**A queued tap that meets the zeros is decided on the stamps.** A queue entry
holds an **absolute** number — a `+1` made against 40 carries 41 — so letting a
drop through unchanged would not "apply the tap on top of the zeros", it would
put the whole count back and leave the row at 41 cut, 0 hotmelted, 0 tuff.
`rebaseQueue` therefore compares the row's `DoneAt` with the tap's own `at`,
which is exactly what writing `DoneAt` on a clear is for: **a row stamped after
the tap was made is the later word and the tap is dropped**; a tap made after
the clear applies on top of the zeros. A tie, or a stamp that will not parse,
leaves the tap alone. Nothing is `BLOCKED` either way — unlike the lock, a
clear **unlocks** the job, so a dropped tap can simply be tapped again, which
is why stranding it the way `dropBlocked()` does would be the wrong shape here.
The real case this closes is undramatic: the workshop wifi drops, the floor
taps, the office clears the job *because something went wrong on it*, the wifi
returns — and without the stamp check the tablet would go back to 49 cut, which
to the office is indistinguishable from "the un-tick didn't stick".

**Why it cannot be reached any other way.** Four independent locks: two call
sites, both on a clear branch; `FLOORCLEAR_OK`, the office's answer to the
question, set in one place and spent once — and keyed to the **write** that
must land (`job|item` or `job|group`), never to the job, because a job has
several checkpoint bursts at once and any of them landing would otherwise spend
the answer and clear the floor before the glass write had landed;
`clearFloorGlass()` re-deriving the reason from the office's own record and the
floor's row at write time rather than trusting the caller; and `ST.officeClearFields(who, at)` — a name and a
time in, seven fields out, every counter a literal nought — so the path has no
argument through which a counter or a per-stage stamp could enter.
`test_station.js` asserts the counts of each of those in the source.

**A refused write is owed, not lost.** The workbook half has landed by then, so
forgetting it would leave the tablet gold — exactly the bug. `FLOORCLEAR_OWED`
counts the refusals, one timer retries after 30 s, and after three attempts the
office is told in a toast that names the numbers still on the tablet. The
retry re-derives, so a job the office has re-ticked meanwhile is dropped.

## See also

- [[glass-colours-tuff-and-lock]] — previous: the hold, ordering, TUFF and the lock
- [[glass-office-clear-history]] — next: the bugs found on the way here
- [[status-list-is-truth]] — the record that closed the "second dashboard" gap
