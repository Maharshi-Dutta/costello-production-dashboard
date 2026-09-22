# The glazing station

## 24. The glazing station, and the phase bar hearing the floor

Built 2026-09-21. Spec: [`docs/specs/2026-09-21-glazing-station.md`](specs/2026-09-21-glazing-station.md). The **third**
floor station, and the first one built from the "Adding a station" checklist in
[`docs/STATIONS.md`](STATIONS.md) rather than from a page.

### What it is

The owner, the same day glazing left the glass station: *"glazing is the last
step for the whole job, not just glass, so it will be its own dashboard with its
section."* And, after the demo: *"make the glazing dashboard as well, and I
should be able to see what is glazed in the glaze station window in the main
dashboard. I should be able to edit glazing from that window."*

```
Production sheet ──parse──▶ ALL ──glzSlice──▶ feedGlazing() ──▶ `Glazing station`
                                (wndMain + drsMain)                    │
                                                                       ▼
index.html + app.js  ◀──Glazed / GlazedBy / GlazedAt / DoneBy / DoneAt──┤
  Show ▸ Glazing station, the drawer's Glazing line, the phase bar      │
glazing.html + glazing.js ──────────────────────────────────────────────┘
```

**One number per job**: units glazed, out of the job's windows plus doors
(owner's decision 1). No stages, no product groups, no glass types — which is
why `glazing.js` is materially shorter than `welding.js` rather than a copy of
it: a card is a job, one stepper and a note box.

### The five things worth knowing

- **No workbook write anywhere in it**, from either side. The only workbook
  write in the whole feature is the `Dashboard Log` line an office edit leaves,
  which is a dashboard-owned sheet (rule 2). `test_glazing.js` asserts on the
  request log that the entire run contains no `/workbook`, no `/drive` and no
  `DELETE`.
- **`Production` only.** `j.wnd`/`j.drs` take the *first* sheet that has a
  number, so `parser.js` gained `j.wndMain`/`j.drsMain` — the Production-only
  pair, beside `j.prodsMain` which welding reads — and the slice reads those. A
  fixture whose two disagree (`wnd: 11, wndMain: 3`) is in the suite, because
  that is exactly the shape [`HISTORY.md`](HISTORY.md) B20 took.
- **No seed.** Glazing was never an office checkpoint, so there is nothing to
  seed from: `GLAZE.seedFields` is `[]` and `seedOf` answers `{}`. That makes
  the feeder the simplest of the three — with no seed field in the plan, no
  patch can ever name a floor column, so there is no "re-read the row before
  seeding" guard to write.
- **Five fields, one builder, both sides.** A tap from the tablet and a click on
  the office's board go through the same `glzTapFields(value, who, at)` and the
  same `glzFloorOnly` (which is `ST.floorOnly` with this station's definition).
  There is no second body-builder that could carry a job fact in. The office
  writes **no** `Station log` line — there is no call to `ST.logFields` on that
  path at all — and leaves one `Dashboard Log` line ("Glazing: R5303, 3 → 6").
- **No day sheet.** `GLAZE` has no `daySheets` entry, so the tablet draws no
  button, the office shows no chip and neither day-sheet list is read.

### What was shared, and what was not

Borrowed from `station-core.js` rather than repeated: `feedPlan`, `sliceHash`,
`floorOnly`, `logFields`, `stationComments`, `stationPeople`, `mergeDelta`,
`boardDiff`, `stripContact`, `atCmp`, `canStage`, `pinOk`, `personExpired`,
`REFRESH_MS`, `PERSON_LOCK_MS`. From `station-ui.js`: the theme, the sign-in
gate, the "Who are you?" picker and its PIN pad.

One helper moved *into* `station-core.js` for this: **`ST.sectionInProduction`**,
the "is this section name In production" regex. `ST.inProduction` (the job-side
question) now calls it, so the feeder's slice test and a tablet's row filter can
never disagree about what the phrase means. Welding's own one-line
`weldInProduction` was deliberately **left alone** — moving it would be churn on
a frozen suite for no behaviour change.

Deliberately **not** extracted: the tablet's queue / flush / delta-poll loop,
which `station.js`, `welding.js` and now `glazing.js` each have a version of.
Pulling it out would touch the glass tablet, whose suite is frozen at 270
checks, and it is the kind of change that wants its own brief rather than a
ride on a feature's.

## See also

- [[station-reports]] — previous: the station report
- [[glazing-station-phase-bar]] — next: the phase bar hears the floor, tests, review
- [[welding-station]] — the earlier station this one borrows its pattern from
- [[glass-two-stage]] — where glazing left the glass station
