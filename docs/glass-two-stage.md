# Glass: one page per stage, no glazing, a new colour rule

## 22. Glass: one page per stage, no glazing, a new colour rule, an editable office board

Built 2026-09-21. Spec: [`docs/specs/2026-09-21-glass-split-no-glazing.md`](specs/2026-09-21-glass-split-no-glazing.md).
It changes [[glass-station-overview]] (§11), [[glass-colours]] (§17) and [[glass-office-clear]] (§18), and the notes at the top of each of those point
here.

**The owner's correction.** *"Glazing is the last step for the whole job, not a
glass step."* It stops deciding whether glass is done and leaves this station
altogether; a glazing station of its own is a later brief. Cutting and
hotmelting each get a tablet.

**A. Glazing leaves the station.** `STAGES` is `cut` and `hotmelt`;
`STAGE_FIELD`/`STAGE_ROW`/`STAGE_BY`/`STAGE_AT`/`STAGE_TOTAL_ROW`/`CLEAR_WORD`
lose their `glazed` entries, and `FLOOR_FIELDS`, `STATION_FIELDS`,
`SEED_FIELDS` and `OFFICE_CLEAR_FIELDS` lose `Glazed`, `GlazedBy`, `GlazedAt`.
Consequences, all intended: the tablet cannot write those columns (the
whitelist has no slot for them), the feeder does not seed `Glazed`, an office
clear does not zero it, and a read does not `$select` it.

**Rule 8 of the brief: no existing data is rewritten, reset or deleted.** The
three columns stay on the list with whatever is in them. A row carrying
`Glazed = total` and `Cut = 0` plans **blank**: glazing has no influence in
either direction. A `glazed` token left in somebody's `Stages` column is
ignored rather than edited out (a person left holding *no* known stage shows on
no page — `STATIONS.md` tells the owner to fix that row by hand). A `glazed`
tap sitting in an old tablet's `cw_stationq` is dropped at load, quietly, with
no request and no error state.

**History stays readable.** `Station log` lines with `Stage = "glazed"` are
still shown, labelled **"Glazing"** from `ST.LEGACY_STAGE_LABELS` — a
display-only map read by `stageLabel` and nothing else. The Floor log window's
stage filter offers the word when the rows on screen carry it. A label is not a
stage: it cannot be tapped, seeded, held or written.

**B. The colour rule** (`ST.glassColours`, still pure and still callable from
Node): how many of the two glass stages are complete decides DG, TG and NOT
TUFF — none blank, **one yellow**, **both gold**. Still reversible, still
needing a total to reach. `finished` on a board record is cutting and
hotmelting complete. The seed: office gold → `Cut = Hotmelt = total`; office
yellow → `Cut = total, Hotmelt = 0`; the sheet's own Cut green → `Cut = total`
as before; anything else the office's own count, as before — **and on an
untouched row a seed field is only ever raised** (R1 below). Everything about
the two-way arrangement — last writer wins, ties to the office, the
`Dashboard Log` line per paint — is unchanged.

**A job that still owes TUFF is not shown as finished** (owner, after the demo
on 2026-09-21: *"why are moving the jobs to complete when tuff is left? if job
has tuff and is not done dont move it, if done then move."*). `ST.tuffOwed(g)`
— `tuffTotal > 0 && !stageComplete(g, "tuff")` — is the whole rule, and it is
**display only**: it decides the gold card, the Finished group and the job
row's chip on the office's board and on the **cutting** tablet (the bench whose
work tuff is), and nothing else. The **hotmelting** tablet is unaffected — tuff
never appears on it. This **reverses one clause** of the 2026-09-10 decision
that put tuff "outside the finished rule": tuff is still outside the job's
glass total, still outside the office's lock (R2), still outside the header's
"N left", and still never seeded. A job with no tuff on it is unaffected.

**TUFF has THREE answers, not two** (R5 below): `"gold"` once its own count is
complete, `""` once somebody has tapped it and it is not, and **`null` — no
opinion at all — while `TuffAt` is empty**. `ST.tuffSpoken(g)` is that question
asked on its own. The `null` is carried to `glassColourPlan`, which leaves the
column out of the plan entirely: neither a paint nor a clear.

**C. One tablet page per stage.** One `glass.html` serves both. The stage comes
from `?stage=cut` / `?stage=hotmelt`, else from the device's own
`cw_stationstage`, else from a two-button chooser shown before the person
picker. **Only a VALID `?stage=` wins** (R4 below): a typo on a kiosk bookmark
falls back to the stage the device already knows it is, and only a tablet that
has been told nothing at all is asked. The header reads `GLASS · CUTTING` /
`GLASS · HOTMELTING` with a `Cutting ▾` control beside it that switches the
tablet (confirm first; it signs the person out, because the name was chosen for
the stage that is leaving). The people picker offers only the Glass people who
hold this page's stage — plus, on the cutting page, anyone who holds `tuff`. A
card draws this page's stage and, on the cutting page, Tuff where the job has
any and the signed-in person holds it; the header count and the card's "N left"
are this page's stage alone. **A card is done on THIS page when this page's
stage is complete — and, on the cutting page, when the job's tuff is counted
too** (`boardNow()` re-reads `finished` for the page), so the
cutter's finished jobs leave the cutter's way before hotmelting has started —
while the record's own `finished`, which the office reads, needs both glass
stages and the tuff.
The office lock (`OfficeDone`) locks the card's **glass** stages on both pages
— **Tuff stays live under it** (R2 below). Notes keep
`Station = "Glass"` on both, so the two tablets share one thread per job.

## See also

- [[glass-colours-tuff-and-lock]] — previous: the earlier three-stage colour rule this replaces
- [[glass-two-stage-office-board]] — next: the office board, review findings, tests
- [[welding-station]] — the pattern this station's site and definition follow
- [[day-sheets-and-reports]] — the day sheet this split's cutting tablet gained
