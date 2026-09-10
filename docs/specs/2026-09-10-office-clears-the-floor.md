# An office un-tick clears the floor's counters

**Date:** 2026-09-10
**Status:** built 2026-09-10, all suites green — **not demoed, not committed**
**Changes a standing rule.** See §2. Only the owner can grant this, and they
did, on 2026-09-10.

---

## 1. What the owner was actually hitting

They spent an afternoon on what looked like three different bugs and was really
one gap. Un-ticking a job's glass in the office cleared the workbook correctly
— confirmed by opening the file directly, the cell stays white — but:

- the **master dashboard** showed gold again for up to a minute (SharePoint
  takes ~36 s to put a change into the downloadable copy the dashboard reads:
  latency, self-correcting, not a fault);
- the **tablet** showed gold **for ever**, because an office un-tick never
  cleared the floor's counters. They stayed at 49/49/49.

So "remove the all done" undid the office's half and left the floor's half
standing, with no way back except somebody tapping `−` forty-nine times.

## 2. The rule this changes

`CLAUDE.md` rule 3 says: *"The office (`app.js`) never writes `Station people`
or `Station log`, and never writes the floor's own columns of `Glass station`
(`Cut`, `Hotmelt`, `Glazed`, `Tuff` and their By/At pairs, `DoneBy`/`DoneAt`)"*
— with one exception, the seeding of a row the floor has not yet tapped.

**A second exception is granted, by the owner, on 2026-09-10:** when the office
clears a job's glass checkpoints, the office may write that job's floor
counters to zero.

Everything else in rule 3 stands. The office still never writes `Station
people`, never writes `Station log`, and never writes the floor's columns for
any other reason.

## 3. What happens on an office clear

When the office clears a job's glass — the per-item **Clear**, or **All glass
done** toggled off — the office writes to that job's `Glass station` row:

| field | value |
|---|---|
| `Cut`, `Hotmelt`, `Glazed`, `Tuff` | `0` |
| `DoneBy` | the office user's short name |
| `DoneAt` | now, ISO |

**Why `DoneBy`/`DoneAt` and not nothing:** the last-writer-wins rule reads
`floorStamp` off those fields. If the counters go to zero with a stale stamp,
the floor's row looks like old news and the tablet's "last touch" line lies
about who moved it. Writing the pair keeps both honest. The per-stage `By`/`At`
pairs are **not** written — those say who did that stage's work, and nobody
did.

**No `Station log` line.** Rule 3's ban on the office writing that list is not
lifted. The clearing is recorded in `Dashboard Log` like every other office
action, which is where an office action belongs.

## 4. The safety this needs

**A confirmation when the floor has actually done work.** If the row's `DoneAt`
is set — the floor has tapped this job — clearing it destroys real recorded
work: potentially a whole afternoon, from one click, on a job somebody is
standing at. So the office is asked first, in plain words, naming what will go:

> The floor has recorded 49 cut, 49 hotmelted, 12 glazed on this job. Clearing
> the glass here will set all of those back to zero. Clear it anyway?

If `DoneAt` is empty the counters are only the office's own seed echoed back,
so there is nothing to lose and nothing to ask.

**A queued tap on the tablet must not be silently stranded.** The tablet may
have an owed write for that job when the zeros land. `rebaseQueue` currently
re-bases a tap whose row has *risen* underneath it; a drop to zero is the
opposite. Decide explicitly and say which: either the queued tap applies on top
of the zeros (the floor tapped after the office clicked, so it wins), or it is
dropped and shown to the floor the way an office lock already drops one. **The
first is more consistent with last-writer-wins** and is the recommendation.

## 5. Hard rules that still hold

1. The tablet still never touches the workbook.
2. Only fills are ever written to `Production`.
3. The office still never writes `Station people` or `Station log`.
4. The office still never writes a per-stage `By`/`At`.
5. No real person's name, email or company domain in code, tests, docs or
   commit messages — including quotations of the owner.

## 6. Tests

1. An office clear on a job the floor has tapped writes exactly
   `Cut/Hotmelt/Glazed/Tuff = 0` plus `DoneBy`/`DoneAt`, and nothing else —
   asserted as an exact key set, the way the feeder's vocabulary already is.
2. It writes no `Station log` line and no per-stage `By`/`At`.
3. The confirmation appears only when `DoneAt` is set, and declining it writes
   nothing at all.
4. After the clear, the colour writer wants blank for every glass column and
   the tablet's card reads zero — the two agree, with no second round of
   writes.
5. A queued tap that meets the zeros behaves as §4 decides, asserted.
6. Marking done still behaves exactly as it does today; this is the un-tick
   path only.
7. The standing proofs: no workbook access from the tablet, no DELETE, nothing
   personal reaching the floor.

## 7. Out of scope

- The ~36 s download latency on the master dashboard. That is SharePoint's, not
  ours, and no delay on our side shortens it. If the dashboard is still showing
  gold **five minutes** after an un-tick, that is a separate bug and is being
  investigated on its own.
- Any change to marking done, to the lock, or to the colour rule.

## Amendments after review

_(append here; do not rewrite the sections above)_

### A. Built 2026-09-10 — one narrowing of §3, the §4 decision, and what was found

**Status: built, all suites green, not demoed, not committed.**

#### A1. §3 is too wide by one case: an untouched row must NOT be written

§3 says the clear writes the six fields whenever the office clears a job's
glass. §4 then says that if `DoneAt` is empty "the counters are only the
office's own seed echoed back, so there is nothing to lose and nothing to ask".
Both are true, but §3's write on such a row is not merely unnecessary — **it is
harmful**, and it would have introduced a second bug of the same family as the
first:

- On a row the floor has never tapped, the **feeder already clears the
  counters**. `ST.officeSeed` of an office record that says nothing is
  `{0,0,0}`, `feedPlan` writes the seed diff on any row whose `DoneAt` is
  empty, and `sliceHash` carries the seed — so the zeros go out on the **next
  load**, not in ten minutes. Nothing new is needed for that case.
- Writing `DoneAt` there would mark the row **touched for ever**. The seeding
  guard is `DoneAt` being empty and nothing else, so from that moment the
  feeder could never seed that row again: the office ticking the job's glass
  done next week would lock the row (`OfficeDone = "Yes"`) while its counters
  stayed at nought, and the tablet would show a locked, unfinishable job. That
  is the same class of fault as the one this spec fixes, pointing the other
  way.

**So the write is made only when the row has a floor stamp**, which is exactly
the case where §4 requires the confirmation. The two conditions collapse into
one: the office is asked whenever anything is written, and nothing is written
when nothing is asked. `ST.floorWorkToClear()` also requires at least one
counter above nought, because four noughts over four noughts tells nobody
anything.

#### A2. The trigger is a TRANSITION, not a state — and it is asked of DG and TG

§3 says "the per-item **Clear**, or **All glass done** toggled off". Taken
literally, any per-item Clear on a glass item would zero the floor. That cannot
be right: the floor's row carries **one combined DG + TG number** (the colours
spec §8), so clearing DG while TG is still gold would tell the floor that
finished glass needs doing again — the mirror image of the reason Amendment B5
of the colours spec refuses to seed a half-yellow job to the total. A wrong
instruction on a workshop screen is worse than a stale one.

Built instead: the clear fires when the office's own record of the job's **DG
and TG** goes **from saying something to saying nothing**, asked through
`ST.officeSeed` — the very function that seeds a row from that record, so the
two readings cannot drift apart. Consequences worth naming:

- the group **Clear** on Glass always fires it (every glass item goes blank);
- a per-item Clear fires it when no other DG/TG item is still ticked — which
  on a job whose only glass is DG is every time;
- `−` down to nought, or typing 0 into the box, is the same act through a
  different control and fires it too;
- a tick made when the record **already** said nothing is not a clear. This is
  what stops clearing a hand-ticked ARCH, on a job whose DG is blank, from
  wiping counters nobody asked about.

There is deliberately **no reconciliation pass**. A state-driven version of
this feature — "any row whose office record is empty and whose counters are
not" — would zero every job the office simply has not ticked yet, including one
the floor is working on at that moment. The write is event-driven or it is
nothing.

#### A3. §4's queued tap: it applies ON TOP of the zeros

Decided as §4 recommends, and for a reason §4 does not give. `rebaseQueue`
re-bases a tap whose row has **risen**; a drop already falls through untouched,
so the tap is sent as the floor made it and the floor's number lands after the
office's zeros. Under last-writer-wins that is right: the floor tapped after
the office clicked.

The alternative — dropping it the way `dropBlocked()` drops a tap that meets
the office's **lock** — was rejected, and the difference between the two cases
is the point. A lock is the office saying *"nobody may move this"*, so a tap
that meets one can never be corrected by the person holding the tablet and has
to be shown to them in red. A clear is the opposite: it **unlocks** the job
(`OfficeDone` goes back to `No`). Stranding the tap would lose a number the
person is looking at, on a card nothing is stopping them from using, and would
put a red apology on it for no reason.

Pinned by the test *"a tap owed when the office's zeros land is applied on top
of them"* in `test_station.js`, which drives the real `readList()` and
`flushQueue()` with the real `ST.officeClearFields` body landing in the list
underneath, and asserts that the tap is neither dragged to nought nor thrown
away, that `BLOCKED` stays empty, that the card shows the tapped number
meanwhile, and that the stages nobody re-tapped stay at the office's nought.

#### A4. `DoneBy` carries `feedWho()`, not `whoAmI()`

§3 says "the office user's short name". `whoAmI()` can be a whole email
address; `feedWho()` is the function written for exactly this — a display name,
or the local part of the address, never the address itself — and it is what
`FedBy` already uses. Nothing else about an office account reaches the floor.

#### A5. The clear is logged as `Floor glass counters`, and the name is load-bearing

§3 says the clearing is recorded in `Dashboard Log` like every other office
action. It is, as `Floor glass counters`, from `49/49/8/11` to `0/0/0/0`.
The name matters: `glassLogStamps()` reads any log entry beginning `Glass ` or
`Glass:` as the office's stamp on a glass column, and this entry must never
become one — it is a record of what happened to the **list**, not a tick on a
cell, and treating it as a stamp would feed the last-writer contest with the
feature's own writes. Asserted.

#### A6. A refused list write is owed and retried

Not in the brief, but the failure it covers is the bug itself: by the time the
list write is attempted the **workbook half has landed**, so a refusal that was
simply forgotten would leave the tablet gold with the sheet clear — precisely
the state §1 describes. `FLOORCLEAR_OWED` counts the refusals, one timer comes
back after 30 s, and after three attempts the office is told in a toast naming
the numbers still on the tablet. Every retry re-derives, so a job the office
has re-ticked in the meantime is dropped rather than cleared.

#### A7. Known gap, stated plainly — and it is wider than first written

The office's answer to the confirmation (`FLOORCLEAR_OK`) lives in memory only,
so **any clear interrupted before its workbook write lands leaves the floor's
counters standing, with no automatic recovery.** The first draft of this
amendment described only one of the three ways that happens; all three:

- **A per-item clear whose burst never fired** (the tab died inside the 800 ms
  debounce, before `cpFireAll` on `pagehide` could run): `cpReplay` replays it
  on the next visit because it is not marked `sent`, and that replay clears the
  **workbook only** — the answer is gone with the page.
- **A per-item clear already in flight**: `cpFire` sets `sent = 1` *before* it
  awaits anything, deliberately, because nobody can tell whether an interrupted
  write landed and replaying it would log it twice. So it is **never replayed**:
  the workbook may or may not have been cleared, and the floor certainly was
  not.
- **The group path** — `setGroupDone`, "All glass done" toggled off, which is
  the owner's own usual way of clearing a job — **has no queue and no replay at
  all.** If the tab closes mid-write, nothing is retried by anything.

Persisting the answer was rejected: it would put a hand-editable token in
`localStorage` on the one path in this app that writes a floor column, and the
whole containment argument rests on that path being unreachable by anything but
a confirmed click. The recovery is the same in all three cases and is now in
`docs/SUPPORT.md`: **tick the job's glass done again, then clear it again** —
the second clear sees the counters still standing and puts them to nought. Not
forty-nine taps, and nothing to edit in SharePoint by hand. The re-derivation
guard means nothing wrong is written on any of these paths, only that nothing
may be written at all.

#### A8. What proves it cannot be reached any other way

Four independent locks, each asserted: two call sites, both on a clear branch;
`FLOORCLEAR_OK`, set in one place and spent once; `clearFloorGlass()`
re-deriving the reason from the office's own record and the floor's row at
write time rather than trusting its caller; and `ST.officeClearFields(who, at)`
— a name and a time in, six fields out, every counter a literal nought — so the
path has no argument through which a counter or a per-stage stamp could enter.
`test_station.js` counts each of those in the source of `app.js` and asserts
that `station.js` mentions none of them.

#### A9. A bug this brief's own tests caught

The first build filtered the job's glass counts with
`ST.TOTAL_TYPES.indexOf(type.toLowerCase())`. `TOTAL_TYPES` is upper case
(`["DG", "TG"]`), so the filter matched nothing, `officeGlassEmpty` answered
`false` for every job, and the whole feature was silently dead — no
confirmation, no write, no error. It now goes through `ST.jobKey`, the same
normaliser `officeSeed` filters with. Worth recording because the failure mode
was invisible: the feature simply never fired.

### B. Fix pass, 2026-09-10 — two review findings, one pinned assumption, one doc

**Status: fixed, all suites green, still not demoed and not committed.**

#### B1. A3 was wrong: a queued tap does not "apply on top of" the zeros (MAJOR)

Amendment A3 said the drop case already fell through `rebaseQueue` and that
this *was* last-writer-wins. It was neither.

`rebaseQueue` compares **numbers only** — `if (!isFinite(now) || !isFinite(was)
|| now <= was) return;` — and never looks at a stamp. A queue entry's `value` is
**absolute**: a `+1` tapped against a row at 40 carries `value: 41`. So after a
clear the entry flushed and wrote `Cut: 41`, leaving the row at **41 cut, 0
hotmelted, 0 glazed, 0 tuff** — a state neither side asked for. The first
version of the test pinned that shape and called it correct.

Worse, A3's justification ("the floor tapped after the office clicked") was
*asserted and never checked*. The failing case needs no unusual timing at all:
the workshop wifi drops; the floor taps; the tap sits in `cw_stationq`, which
survives reloads and drains only when `READY && SITEID`; the office clears the
job *because something went wrong on it*; the wifi comes back; the tablet goes
back to 49 cut. To the office that is indistinguishable from "the un-tick
didn't stick" — a fifth report of the same complaint, caused by the fix for it.

**Fixed with this feature's own machinery.** `DoneAt` is written on a clear
precisely so `floorStamp` stays honest (§3, A5) — so the comparison it exists
for is now actually made. When a counter has gone **down** under a queued tap,
`rebaseQueue` compares the row's `DoneAt` with the tap's own `at`:

- row stamped **after** the tap → the office (or another tablet) moved it last
  → **the tap is dropped**, with a console warning naming it;
- tap made **after** the row's stamp → the floor moved it last → the tap
  applies on top of the zeros, which is what §4 recommended;
- a tie, or a stamp that will not parse, leaves the tap alone: the floor's own
  statement is what the tablet is for, and it fails towards keeping it.

Dropping is the right shape here for the reason A3 gave for the opposite
conclusion: a clear **unlocks** the job, so the floor can simply tap again. That
is also why nothing is put in `BLOCKED` — a red apology belongs on a card
nobody may touch, not on one that is working normally. The rise-rebase is
untouched, and so is the tuff clamp.

Both orderings are pinned in `test_station.js`, driving the real `readList()`
and `flushQueue()` with the real `ST.officeClearFields` body landing underneath:
tap before the clear → dropped, nothing at all goes out, nothing BLOCKED, the
card shows the row's real number at once; tap after the clear → one PATCH
carrying the floor's number, the stages nobody re-tapped left at the office's
nought, and the log line reading `0 -> 1` because `from` was captured against
what the office could actually see. A third case pins the unreadable stamp.

#### B2. `FLOORCLEAR_OK` was granted per clear but spendable by any write on
that job (MAJOR)

`floorClearAfterWrite` was keyed on the **job**, and it is called from
`cpFlushItem` for whichever burst lands and unconditionally at the end of
`setGroupDone` for any group. Burst keys are `job|item`, so a glass clear and a
Windows tick on one job are separate bursts: un-tick the glass and tick Windows
on the same job inside the 800 ms debounce, the Windows write lands first,
spends the answer, and the floor's counters are cleared **before the glass
workbook write has landed**. If that write then failed, the floor would be at
nought with the sheet still gold — the mirror image of the bug this spec exists
to remove. The same keying let an answer that was granted and never spent (the
office confirms, then immediately re-ticks, cancelling the burst) be spent later
by an unrelated write on that job.

**Fixed** by keying the answer to the **write that must land** — the burst key,
`job|item` for the per-item path and `job|group` for the group one — so only
that write can spend it. A leftover answer is now inert: nothing but the very
write it was given for can consume it, and if that write ever does land,
`clearFloorGlass` re-derives and refuses unless the office's record still says
nothing is done. Pinned in `test_glasscolour.js` by starting the unrelated
burst **first** and asserting on request order that the list write comes after
the glass fill, and separately by planting a leftover answer and showing an
unrelated write cannot spend it.

#### B3. Verified, not changed: a re-tick inside the same minute goes to the office

The reviewer asked that minute-precision office stamps be treated as covering
their whole minute before this ships. That already landed in `ab0c85c`:
`stampMs` adds `STAMP_MINUTE_MS` when the seconds group is absent. So a clear
at 15:00:10 (a full ISO `DoneAt`, read exactly) followed by a re-tick recorded
as `"2026-09-10 15:00"` (read as 15:00:59.999) resolves to the office, which is
right — otherwise the colour writer would paint the floor's blank straight back
over the re-tick and the office would watch its own work undone.

Nothing defended that sequence, and it is now an ordinary workflow rather than
a race: clearing is a deliberate click and re-ticking is the obvious next move.
`test_glasscolour.js` now pins it, together with the other side of the same
boundary — a floor stamp in the **next** minute takes those cells back, so the
test proves a rule and not merely "the office always wins".

#### B4. A7 was too narrow, and the recovery now lives in `docs/SUPPORT.md`

A7 described one of three ways a clear can be interrupted. It has been
corrected in place: a burst that never fired is replayed and clears the
workbook only; a burst already in flight is marked `sent` **before** the await
and is therefore never replayed at all; and the **group** path — "All glass
done" toggled off, the owner's usual way of clearing a job — has no queue and
no replay of any kind. In every case the way back is the same and it is now in
the troubleshooting doc rather than only in a spec amendment: **tick the job's
glass done again, then clear it again.**

#### B5. One wording fix

The failure toast said `49/49/12/3` while the question said "49 cut, 49
hotmelted, 12 glazed". `ST.clearWords()` now produces those words once and both
the question and the apology are built from it, so the office cannot be shown
two descriptions of the same row. The `Dashboard Log` line reads from the same
words to `nothing`.
