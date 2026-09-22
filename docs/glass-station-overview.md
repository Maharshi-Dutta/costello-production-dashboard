# Glass station (floor dashboards)

## 11. Glass station (floor dashboards)

**What.** A separate page for the glass area on a shared tablet, showing only
job number, customer name and **one total number of glasses** (DG + TG), with
the stages recorded per person. There are
no glass types on the floor at all. The office sees it all from the master
dashboard. The floor cannot see the master.

> **CHANGED 2026-09-21 by [[glass-two-stage]] (§22).** The stages were three (Cutting, Hotmelting,
> **Glazing**) until then. Glazing is the last step of the whole job, not a
> glass step, so it left this station: two stages now, one **tablet per
> stage**, and the `Glazed`/`GlazedBy`/`GlazedAt` columns left untouched on the
> list and read by nothing. Read [[glass-two-stage]] before anything below about three stages,
> about glazing gating gold, or about one tablet showing them all.

**Where.** `station-core.js` (pure logic, loads in browser and Node),
`station.js` (tablet UI), `glass.html` (shell + styles), plus the station
blocks in `app.js` (search `STATIONS`, `feedStation`, `stationPoll`,
`renderStationLog`, `stationSectionHtml`) and `listDelta`/`stationSite` in
`graph.js`. Specs: [`docs/specs/2026-09-08-glass-station.md`](specs/2026-09-08-glass-station.md) (v1) and
`…-v2.md` (both with "Amendments after review" sections) and `…-v3.md`, which
made it one row per job, one number per job and cards you tap directly. Admin
guide: `docs/STATIONS.md`.

**How, in five parts.**
1. **Separation by permission, not by screen.** The station account is a
   member only of a private SharePoint site `Floor stations`; it has no
   access to the workbook. If it opens the master page, `start()` shows a
   plain "wrong page" notice with a link to the station page and stops
   polling.
2. **The office dashboard is the feeder.** After every `load()`,
   `feedStation()` computes the glass slice of jobs in the "In production"
   section (`glassSlice`, one row per job, `Total = DG + TG`), diffs it
   against the `Glass station` list (`feedPlan`), and sends adds/patches
   (3 lanes, 60 writes per run, a 30 s follow-up if cut short). It writes job
   facts (Job, Customer, GlassType = the literal `GLASS`, Total, Seq, Active,
   FedAt, FedBy) and — the one exception to "never the floor's columns", added
   in v3 — the three counters, **only on a row it is creating or a row whose
   `DoneAt` is empty**, seeded from the office's own glass checkpoints
   (`glassCounts` in app.js → `ST.officeSeed`). That guard is re-checked with a
   one-item `listItem()` read immediately before every write that carries a
   counter (`stationFeedPatch`), because the plan is made from one read and
   sent as up to sixty writes; the same read drops the seed when the row's
   `FedAt` is newer than this dashboard's `lastStamp`. It never writes a By, an
   At or the last-touch pair, never touches a counter after the floor's first tap,
   and never deletes (a job that leaves production is `Active = No`). Skipped
   quietly without list consent; throttled by a hash of the slice, which
   includes the seed.
3. **The tablet.** Sign in once with the station account (asks for the list
   scopes at the door). "Who are you?" picker from `Station people`
   (Title=name, Station, Stages comma list, PIN, Active); PIN pad when the
   row has one; a person may hold one or more stages; steppers for other
   stages are greyed and `tap()` refuses them too. One card per job, with the
   job number, the customer and the job's own glass count — `glassWords(total)`
   = DG + TG, plus `· N tuff` where there is tuff — which is **the same for
   everybody**, and the steppers (−, +, All / None) on the card itself. Since
   2026-09-10 each stepper row the person **holds** carries its own
   `ST.stageLeft(g, stage)` under the stage's name ("Cutting 20 left"): one
   number per stage, **never summed** (an earlier build summed them and a real
   job read 157 — 49 glasses × 3 stages + 10 tuff — which the owner rejected on
   sight), tuff counted against `TuffTotal` and never against DG + TG. The
   number sits in the label cell of the row whose buttons move it, so the card
   is no taller with four numbers than with one; a row is either yours (a
   number) or not (the words "not yours"). Two people holding the same stage
   read the **same** number, because `stageLeft` never sees a person — nothing
   enforces that and nothing can break it. Nothing to expand, no bars, nothing
   to scroll inside — plus a search box in the header (`boardFilter`, job
   number or customer) and, beside it, `#gtotal`: `ST.boardLefts`, the same
   per-stage numbers summed down the board in the same words and the same order
   ("Cutting 8 left · Tuff 11 left"), never narrowed by the search and hidden
   whenever the search box is — and absent entirely for somebody who holds no
   stages. All of it is derived on the device at every draw and stored nowhere.
   Ten quiet minutes lock the
   tablet back to the picker; Switch person does the same. Every tap shows at
   once, is queued in `cw_stationq`, PATCHes only that stage's counter plus
   `<Stage>By/<Stage>At/DoneBy/DoneAt`, then POSTs one `Station log` line
   (Title=job, Station, GlassType, Stage, From, To, Who, At) with `From`
   captured at tap time. The queue drains until empty, retries, and survives
   a reload; a replayed queue is whitelisted to the counter fields.
4. **Real time.** Both pages poll by delta (`listDelta`, 10 s on the tablet;
   10 s on the office while the board, a glass job's drawer or the log window
   is open, 60 s otherwise). `mergeDelta` applies last-occurrence-wins and
   `@removed`; a 410 resync or a refused delta falls back to one full read
   (and a 5-minute "delta off" flag). The tablet repaints only the cards that
   changed (`boardDiff`), re-bases a queued tap whose row has risen under it
   (`rebaseQueue`), and folds finished jobs — all three counters at the
   total, with a total > 0 — into a collapsed "Finished · n" group, gold.
5. **The office view.** Show ▸ Glass station replaces the job list with a
   read-only board: one card per job with "12 glasses" — the office's number
   is still the size of the job, where the tablet's is one person's remaining;
   they answer different questions and are not meant to agree — the three
   counters with who · when under each, "last: …" from the log, and the same gold rule
   (finished cards go gold and sink to the bottom). The drawer shows the same
   three lines, a 12-line timeline and Full log; the Floor log window filters
   by person, stage, job and day with counts, read-only, 90-day horizon, and
   no glass-type column. The office never writes `Station people` or
   `Station log`.

**Where the lists live (interim, 2026-09-08).** `stationSite()` prefers the
separate `Floor stations` site; while it does not exist (the owner has no
Global Administrator to create it) the same three lists live in the
workbook's own site and both pages fall back to it, re-checking for the real
site every ten minutes and switching over on their own (delta tokens and
list-id caches are dropped on the move). In that interim the station account
is a member of the workbook's site and can therefore open the workbook; the
owner accepted that. See `docs/STATIONS.md`, "Interim".

**Honest limits (told to the owner).** PIN, gating and the lock are a
deterrent on a shared device, not security; the real boundary is site
membership. Graph has no push; 10 s delta polling is the real-time mechanism.
Jobs reach the floor only while an office dashboard is open.

## See also

- [[export]] — previous: export to Excel / PDF
- [[glass-station-job-row]] — next: the glass chip on the ordinary job list
- [[glass-colours]] — the floor's counters reaching the Production sheet
- [[glass-two-stage]] — the 2026-09-21 redesign that changed this section
- [[station-feeder]] — the feeder/tablet/office data-flow diagram
