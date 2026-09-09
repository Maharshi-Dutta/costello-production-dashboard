# Specs index

One line per spec, oldest first. "Shipped" means the feature is live in
`main` (checked against the git log); "in progress" means a spec exists but
the code isn't merged yet.

| date | spec | status |
|---|---|---|
| 2026-09-04 | [`2026-09-04-checkpoints.md`](2026-09-04-checkpoints.md) — per-job progress ticking (windows, doors, glass, product F/S/T) | shipped |
| 2026-09-04 | [`2026-09-04-alerts.md`](2026-09-04-alerts.md) — job alerts by email, subscriptions in the dashboard, mailing by a Power Automate flow | shipped |
| 2026-09-07 | [`2026-09-07-export.md`](2026-09-07-export.md) — export jobs to Excel or PDF, filtered and field-picked, no phone/eircode ever | shipped |
| 2026-09-07 | [`2026-09-07-phases-fab.md`](2026-09-07-phases-fab.md) — the phase pipeline as primary progress, plus the first floating selection wheel (quarter-circle fan) | shipped — the wheel's look was replaced by the next spec |
| 2026-09-07 | [`2026-09-07-radial-menu.md`](2026-09-07-radial-menu.md) — the selection wheel redone as a full-circle radial menu (supersedes the fan in the previous spec) | shipped |
| 2026-09-07 | [`2026-09-07-phases-list.md`](2026-09-07-phases-list.md) — hand-set phases stored in a SharePoint list (`Dashboard phases`), never in the workbook | shipped |
| 2026-09-08 | [`2026-09-08-glass-station.md`](2026-09-08-glass-station.md) — the glass floor station: a separate tablet page fed by the master dashboard, its own SharePoint list, no workbook access | shipped 2026-09-08 — superseded by the v2 spec below |
| 2026-09-08 | [`2026-09-08-glass-station-v2.md`](2026-09-08-glass-station-v2.md) — glass station, second iteration: people and PINs on the tablet, the stages renamed to Glass cut / Hotmelt / Glazing, a `Station log` of who changed what, and ten-second delta polling on both screens | shipped 2026-09-08 — superseded by the v3 spec below |
| 2026-09-08 | [`2026-09-08-glass-station-v3.md`](2026-09-08-glass-station-v3.md) — glass station, third iteration: one row and one number per job (DG + TG, no glass types on the floor), the counters seeded from the office's own checkpoints, cards you tap directly with no bars or expanding, a search box, and a gold card when a job is finished | shipped 2026-09-08 |
| 2026-09-09 | [`2026-09-09-john-template.md`](2026-09-09-john-template.md) — selecting a whole section, the Production sheet's row colour code (urgent / booked / trade / on hold), the **John print sheet** view and export (the `Production (2)` sheet on its own terms, with phone numbers), and the print-notes window backed by the `Dashboard print notes` list. Includes the owner's second round and the review fixes in its Amendments section | shipped 2026-09-09 (bf3470b) |
| 2026-09-09 | [`2026-09-09-glass-chip-on-job-row.md`](2026-09-09-glass-chip-on-job-row.md) — the floor's glass progress on the master dashboard's own job rows: a `Glass` column showing all three stages as one number, gold when finished, repainting when the floor taps and polling fast while glass jobs are listed. No workbook write and no list write. Two review rounds in its Amendments section | shipped 2026-09-09 |
