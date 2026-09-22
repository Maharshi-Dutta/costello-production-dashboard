# An office clear reaches the floor's counters — the bugs found on the way

Continues [[glass-office-clear]] (§18). This is the historical record of the
holds/stamps/settling mechanism that existed before [[status-list-is-truth]]
(§19, step 3) replaced it — kept for the reasoning it recorded, not because
any of it still runs.

**An office action voids the writer's un-landed paint of that cell** (the
morning of 2026-09-11). `applyPending` only ever *masked* a `gc` hold with a
newer `cp` hold on the same glass type; it never dropped it. So the writer
painted a job gold at boot and held it; the office pressed Clear 33 s later
and its `cp` hold outranked the gold, so the drawer went white; six seconds
after that the download agreed with the **Clear**, the office's own hold was
let go — and the writer's older gold, still underneath, came straight back
over a sheet the office had just made white. With holds now kept while the
file disagrees (below) that lie stood for **11 min 40 s**, crossing the
feeder's ten-minute window, whose `officeComplete` read the held gold as "the
office says done" and re-locked the tablet at 0/0/0/0. Fixed in `pend()`: an
office `cp` hold on `glass:<type>` discards any `gc` hold for that type —
value, stamp and give-up count. Only what was painted *before* the office
acted; a genuine floor tap after it still paints.

Three things found with it: a `$batch` reply **missing** a response counted as
success (`find` returns `undefined` both for a missing entry and for no
failure) — now `findIndex`, and a missing answer throws; the writer wrote **no
Dashboard Log line**, so a job going gold by itself was invisible in Changes —
it now writes one per paint, named `Floor glass colours` because
`glassLogStamps` must never read it as an office stamp (the same constraint as
`Floor glass counters`, and asserted); and `holdGaveUp` said "your change …
may not have saved" for a colour the office never made — now worded by source.

**A hold never expires into a stale copy** (2026-09-10, and this was the
un-tick gold — not the writer, not a stamp). The owner un-ticks, then
**refreshes** (they refresh after an un-tick, not after a tick: the asymmetry
is in that, not in the code). `cw_pending` survives the reload, so the screen is
right — but the 45 s reconcile timer died with the old page, and `poll()` only
downloads when `lastModified` *moves*, which this dashboard's own write was the
last thing to do. Measured: one `/content` download after the reload and none
for the next 215 s. The page then sits on the stale gold parse with the hold as
its only cover, and the hold's expiry was a pure clock test — so the next
`applyPending` from anywhere (a floor tap on **another** job, through
`glassColourRun`) dropped it and **unmasked the gold**.

Now: `bootReconcile()` arms a re-read for a page that starts up holding
anything (5 s if the hold is already old); a fresh parse that still disagrees
keeps one armed; and an expired cp/gc hold whose parse still disagrees is
**kept**, with a read demanded at once (bounded to one a minute), until the
file agrees. After `HOLD_GIVEUP` (12) fresh parses still disagreeing it is let
go and the office is told once, in red, naming the job and the item. What
counts as "disagrees" is the **colour** — what a release would unmask — so a
held count whose colour the file already shows still expires quietly.
`officeSettling` follows the hold, so the writer keeps standing down.
**Known gap:** SharePoint can serve an older copy than the one before it, and
after a hold is released nothing protects the item from that; the fix needs a
per-cell modified stamp, which Graph does not give for a workbook cell.

**The office is absolute: the writer stands down while an office change is
settling.** The owner restated the rule on 2026-09-10 — *"the hierarchy is
Excel, then master dashboard, then glass. Any change from the dashboard is
absolute … an un-tick from the dashboard is absolute — no thinking, no
arguing."* So `glassColourPlan` opens with `if (officeSettling(j, log)) return
null;` and makes **no decision at all** about a job while the office's own
change on it is still settling — no stamp comparison, no plan, no write.
`officeSettling` is true while a `glass:` checkpoint hold exists on the job
(`pend()` dates it at the click; `applyPending` releases it when the download
agrees) **or** while the per-job office stamp is younger than `PENDING_MS`.
Both, because the office must be safe in every place its action is read from,
and the window is **per job**.

The reason is not that the comparison is wrong — it was corrected twice — but
that inside that window it is fed copies that have not caught up (the download
~36 s behind, the floor's list a poll behind, the clear's own `DoneAt`
indistinguishable from a tap), and this feature exists to carry the **floor's**
work up to the sheet, not to second-guess the office. Outside the window
nothing changes: last-writer-wins still decides and a genuine floor tap still
paints. **The cost, plainly: a floor tap made inside the window is deferred by
up to three minutes** — never lost, because the writer re-plans from the
current list on every pass, and a pass that deferred anything now arms the same
30 s follow-up a capped pass uses. **The hole it does not close** is a second
dashboard, or a reload, inside the download lag: it has no record of the
office's action at all, so it cannot stand down for it. Closing that needs an
`OfficeAt` column on the `Glass station` row — a decision for the owner.

**The office's stamp is per JOB, and the clear's own `DoneAt` is the office's.**
Fixed 2026-09-10 after the owner saw it in production: `glassOfficeStamp` was
asked **per column**, so for every column the office had not named *in that
action* it found no hold, no Progress row and no Log line and answered a
literal `0` — and any floor stamp at all beat it. Pressing Clear on TG alone
therefore painted out the TUFF and NOT TUFF the office had ticked by hand. And
the floor stamp it lost to was one the office had itself just written:
`clearFloorGlass`'s own `DoneAt`, a fraction of a second after the click, while
`glassLogStamps` deliberately refused to count the `Floor glass counters` line
that records the same act. One half of the office's own clear was counted for
the floor and the other half for nobody.

`glassOfficeJobStamp(j, log)` now answers **the newest thing the office has
said about this job's glass, on any column**: every glass item's hold (stamped
at the click by `pend()`), every `Glass …`/`Glass: …` Log line for the job
whichever column it named, the `Floor glass counters` line, every glass item's
Progress `When`, and `OFFICE_FLOOR_AT[job]` — this dashboard's own memory of
the `DoneAt` it wrote, which is the only record of a clear until the Log line
has been written and read back. `glassOfficeStamp` takes the max of that and
its per-column sources. It is computed once per job per pass (`byId` is a
linear scan) and `OFFICE_FLOOR_AT[job]` is dropped once the floor has moved
that row since, so the map stays bounded. The rule is unchanged: a floor tap
after the office's action still carries the later stamp and still wins.

**Known gap.** `FLOORCLEAR_OK` lives in memory only, so **any clear interrupted
before its workbook write lands leaves the floor's counters standing, with no
automatic recovery** — three ways: a per-item burst that never fired is
replayed by `cpReplay` and clears the workbook only; one already in flight is
marked `sent` before the await and is never replayed at all; and the **group**
path has no queue and no replay of any kind, which is the owner's usual way of
clearing a job. Persisting the answer was rejected — it would put a
hand-editable token in `localStorage` on the one path that writes a floor
column. The recovery, in `docs/SUPPORT.md`: tick the glass done again, then
clear it again. Not forty-nine taps.

## See also

- [[glass-office-clear]] — previous: an office clear reaches the floor's counters
- [[status-list-is-truth]] — next: status lives in a list, the Excel colour is a copy
