# The station feeder and the station page

### The station feeder and the station page

Shipped 2026-09-08. The data flow, in one line: the **feeder** (inside the
master dashboard) writes job/glass-type facts into `Glass station`; the
**tablet** writes back only its own counters plus one `Station log` line per
counter write; the **office** reads `Glass station` and `Station log` by
delta (and `Station people` once, for the log window's filters) and writes
none of the three.

- **Feeder → `Glass station`.** After every successful `load()` of the
  workbook, `feedStation()` (`app.js`) turns the freshly parsed jobs into a
  slice (`ST.glassSlice`, **one row per job**, `Total = DG + TG`, jobs "in
  production" only) and a plan of adds/patches (`ST.feedPlan`), then sends
  it. It skips the run entirely when the signed-in account has not granted
  the SharePoint list permission, or when the slice's hash
  (`ST.sliceHash`) matches the last run and that run was under ten minutes
  ago. A job that leaves production is patched `Active = No`, never
  deleted; the feeder never sends a By, an At or the last-touch pair, and it
  sends `Cut`/`Hotmelt` **only** on a row it is creating or a row
  whose `DoneAt` is still empty — the v3 seeding rule, which starts a row at
  what the office has already ticked off (`ST.officeSeed`) and stops dead at
  the floor's first tap. Every write carrying a counter re-reads that one row
  (`CW.listItem`) immediately before sending, because the plan is made from a
  single read and a tap can land before the write it was planned against.
- **Tablet → counters + `Station log`.** The station page (`glass.html` +
  `station.js`) reads `Glass station` and `Station people`, asks who is at
  the tablet (a name and, optionally, a PIN — a deterrent, not a secret;
  see `docs/STATIONS.md`), lets them tap only the stages they hold, and
  queues a PATCH of that row's changed counter plus its By/At and the
  overall DoneBy/DoneAt (`cw_stationq`, whitelisted to
  `ST.FLOOR_FIELDS`). Once a counter PATCH lands, one `Station log` line is
  queued and POSTed (`cw_stationlogq`) — never before, and never if the
  counter write never happened.
- **Office → the floor's counters, on a clear only.** Since 2026-09-10 the one
  other time the office writes a floor column: when it **clears a job's glass
  checkpoints**, `clearFloorGlass()` puts that row's `Cut`/`Hotmelt`/
  `Tuff` back to nought and stamps `DoneBy`/`DoneAt`, so the office, the
  workbook and the floor all say the same thing. Only on a row the floor has
  really tapped, only after the office has answered a question naming what will
  be destroyed, and only through `ST.officeClearFields(who, at)`, which cannot
  express any value but nought. See [[glass-office-clear]] (§18 of the old REFERENCE.md).
- **Office → the floor's counters, from the board.** Since 2026-09-21 the other
  time: the Glass station board's own `−`/`+`/`All` on a stage line
  (`glassOfficeEdit`). One PATCH of five fields, built only by
  `ST.floorOnly(ST.tapFields(stage, value, who, at))` — the tablet's own pair —
  with the row re-read immediately before the number is derived. One
  `Dashboard Log` line per click through `noteChange`, and **no `Station log`
  line**. Nothing on the path touches the workbook. See [[glass-two-stage-office-board]] (§22).
- **Tablet → `Station comments`.** Since 2026-09-15 a floor worker can also
  leave a **note** against a job, from a composer on that job's card. One
  `listAdd` per note, straight to the list, append-only — nothing edits or
  deletes one, on either side. The whole channel (composer, thread, row builder
  and both list calls) is `ST.stationComments(cfg)` in `station-core.js`, whose
  only station-specific input is `cfg.station`, so a second station page passes
  its own name and shares the list. The tablet shows a station **its own**
  notes; the office drawer shows every station's. See [[station-comments]] (§20).
- **Office reads by delta.** Apart from that one clear, the master dashboard
  never writes `Glass station`'s floor columns, and never writes `Station
  people`, `Station log` or `Station comments` at all. It reads all four (the people list once, for the log
  window's filters) and keeps the board, the drawer's timeline and the log
  window current with `listDelta()` — every 10 seconds while the station
  board, the log window, or a drawer for a job with glass is open
  (`stationTick()`/`stationWatching()`), once a minute otherwise.

See `docs/STATIONS.md` for the full data model, admin setup and
troubleshooting, and [`docs/specs/2026-09-08-glass-station.md`](specs/2026-09-08-glass-station.md) plus its v2 and
v3 for the binding spec (the code wins over any of them where they disagree).

### The welding station (2026-09-16), and what a second station costs

The same flow, in a different site, with one extra arrow. Spec:
[`docs/specs/2026-09-16-welding-station.md`](specs/2026-09-16-welding-station.md).

```
Production sheet ──parse──▶ ALL ──weldSlice──▶ feedWelding() ──▶ `Welding station`
  (F and S of every allowed                      (app.js)          (Floor stations site)
   product group; T never)                                            │      ▲
Dashboard progress ──cpStatus──▶ the seed, on an untouched row only ───┘      │
                                                                             │
welding.html + welding.js ──FramesDone/SashesDone + By/At + DoneBy/At─────────┤
  (the tablet)             ──one `Station log` line per landed counter────────┤
                           ──one `Station comments` row per note──────────────┘
                                                                             │
office board (Show ▸ Welding station) ──the same five fields, + one ──────────┘
                                         `Dashboard Log` line, never `Station log`
```

- **It reads the `Production` sheet and no other.** The slice is built from
  `j.prodsMain` — the parser's product counts taken from the sheet named
  `Production` alone — and **not** from `j.prods`, which is the maximum across
  every sheet. A job with no `prodsMain` is not on `Production` and is not fed.
  `j.prods` is unchanged and is still what the drawer, the checkpoints, the
  exports and the John print sheet read. Why the difference matters, with the
  live bug that proved it: `docs/HISTORY.md` B20.
- **No workbook write anywhere in it.** The welding station never paints a cell
  and neither do the office's welding edits. The gold F/S cells are read once,
  as a seed, through the **record** (`cpStatus`, the `Dashboard progress`
  list) — never as a colour, never written back. The one workbook write on the
  whole path is the `Dashboard Log` line an office edit leaves, and `Dashboard
  Log` is a dashboard-owned sheet.
- **A different site.** The welding lists are in `Floor stations`; the glass
  lists are still in the workbook's own site. `CW.stationSite(which)` honours
  the definition's `site` (`"floor"` / `"own"`), so the two can never share a
  site id, a delta token or a "the lists moved" counter — see `docs/STATIONS.md`,
  "Adding a station", step 7, which also records the one gap that pin leaves.
- **What was generalised, and what was not.** `ST.feedPlan`, `ST.sliceHash` and
  `ST.floorOnly` now take a **station definition** whose default is the glass
  one, and `ST.stationPeople` and `ST.boardDiff` take the station's stage list
  and card signature. Nothing glass-specific was added to `station-core.js` and
  nothing welding-specific went into it: the welding rules are all in
  `welding-core.js`, and the shared tablet shell (theme, gate, picker, PIN pad)
  is `station-ui.js`.
- **The office writes a floor counter here, on purpose.** The second sanctioned
  case after the glass clear, dated in `CLAUDE.md` rule 2 on 2026-09-16: the
  owner asked to be able to correct welding progress from the office, "even on
  a finished job — they might make it wrong". It goes through
  `WELDC.weldOfficeFields` → `WELDC.weldFloorOnly`, which is the same filter the
  tablet's own queue runs on, so a job fact cannot get into the body. It is
  recorded in `Dashboard Log` and **never** in `Station log`.

## See also

- [[sheets-and-lists-scope]] — previous: dashboard-owned sheets and the scope split
- [[module-map-and-invariants]] — next: module map, key invariants, where state lives
- [[glass-station-overview]] — the glass station this feeder was built for
- [[welding-station]] — the second station this page describes
