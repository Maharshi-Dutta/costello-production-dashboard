# Costello production dashboard

A live view of Costello Windows' production workbook on SharePoint — a
static site with no build step and no framework: plain ES2017 browser
JavaScript, vendored MSAL and ExcelJS, no bundler, no modules.

Two pages, two audiences:

- **`index.html`** — the master dashboard. Used by the office (a couple of
  people who need the full picture: every job, every section, checkpoints,
  phases, alerts, exports).
- **`glass.html`** — the glass floor station, shipped 2026-09-08. Used on a
  shared tablet by whoever is on the glass line, signed in with a shared
  station account that cannot see the workbook at all — only job number,
  customer, one number of glasses and progress on Cutting / Hotmelting /
  Glazing. See `docs/specs/2026-09-08-glass-station.md` (and its v2 and v3)
  and `docs/STATIONS.md` for the full data model and setup.

The dashboard reads the workbook by downloading it whole (so cell fill
colours come through) and writes back surgically, cell by cell, through the
Microsoft Graph Excel API — it never rewrites the file. Four SharePoint lists
hold state that must never touch the workbook: hand-set phases
(`Dashboard phases`, in the workbook's site) and the glass station's three
(`Glass station`, `Station people`, `Station log`, in the separate private
site `Floor stations`). The master dashboard feeds `Glass station`'s job
facts and reads all three floor lists, but never writes a floor counter, a
person or a log line — only the tablet does that.

## Features

- **Sections and categories** — jobs grouped by the Production sheet's own
  divider rows (Ready to deliver, Collect & supply, Won't take, Second hand,
  In production…), read live so a moved divider updates the dashboard with
  no code change.
- **Mark ready** — move a job's row into a different section; the only
  structural write this app makes, re-finding the row by job number and
  copying every format, border and fill with it.
- **Checkpoints** — tick windows, doors, glass units and each product's
  frame/sash/transom counts off individually; the sheet keeps only the
  colour (white/yellow/gold), the exact count lives in `Dashboard Progress`.
- **Phase pipeline** — a seven-step "In office → … → Fitted / delivered"
  reading of where a job has got to, worked out purely from what's already
  on the sheet, plus an optional hand-set override stored in the
  `Dashboard phases` SharePoint list (the sheet's own evidence always wins
  when it's further along).
- **Radial selection menu** — tick one or more jobs and a floating button
  fans out Alert to…, Export, Move to…, and Untick all on a full circle.
- **Versions and rollback** — browse SharePoint's own version history of the
  workbook, see a diff, restore an older version (itself becomes a new,
  rollable version).
- **Alerts by email** — subscribe an address to a job; every third day a
  Power Automate flow mails each subscriber their jobs and comments (see
  `automation/SETUP.md`). The dashboard only manages subscriptions — it
  never sends mail itself.
- **Export to Excel or PDF** — filtered, field-picked, direct download; no
  phone numbers or eircodes are ever included, and every export is logged.
- **Glass station** — a separate tablet page for the glass floor (shipped
  2026-09-08), fed by the master dashboard on every load; a person picks
  their name (and PIN) and taps to record Cutting / Hotmelting / Glazing,
  only for the stages they hold. The master shows the floor's progress and
  its log — who changed what, when — read-only, in the "Sheet" dropdown, the
  job drawer, and a log window.

## File map

| file | purpose |
|---|---|
| `index.html` | master dashboard: page shell, styles, script tags |
| `glass.html` | glass station page shell and styles (shipped 2026-09-08) |
| `graph.js` | Entra sign-in (MSAL) and every Microsoft Graph read/write call |
| `parser.js` | workbook → job model; runs in the browser and in Node |
| `checkpoints.js` | checkpoint merge rule, colours, debounced writes, the phase pipeline |
| `export.js` | pure export logic: filtering, rows, Excel workbook, PDF document definition |
| `station-core.js` | pure glass-station logic (slice/plan/board/tap/people/log) — shipped 2026-09-08 |
| `station.js` | glass station page UI (shipped 2026-09-08) |
| `app.js` | master dashboard UI: rendering, drawer, windows, wiring |
| `build.py` | stamps a build id (`?v=…`) onto every script tag on `index.html` and `glass.html`, and writes `version.json` |
| `verify.js` | dev-only check: compares the JS parser against a Python reference extract |
| `automation/` | the alerts mailer: Office Script source, its generator, and setup instructions |
| `vendor/` | vendored MSAL, ExcelJS and pdfmake — no CDN dependency |
| `test_*.js` (incl. `test_station.js`), `automation/test_digest.js` | offline test suites, one per feature |
| `docs/` | architecture, support, stations and the specs index (this README's companions) |

## Running it locally

```bash
python -m http.server
```

Then open `http://localhost:8000/index.html` (or `glass.html`) and sign in
with a Costello account. Sign-in uses MSAL against Entra ID; a session sees
exactly the jobs that account can already open in the workbook in Excel.
There is no local/offline mode for the live app — the offline tests
(below) are how logic gets exercised without a network.

## Testing

Every feature has an offline Node test file that loads the real source with
`vm.runInThisContext` and drives it against a fake `fetch()` — no network,
no live workbook, ever. Run each one directly, always with `set -o pipefail`:

```bash
set -o pipefail
node --check parser.js graph.js checkpoints.js export.js app.js station-core.js station.js
node test_move.js && node test_checkpoints.js && node test_alerts.js \
  && node test_export.js && node test_phases.js && node test_phases_list.js \
  && node automation/test_digest.js && node test_station.js
```

See `CLAUDE.md` for the full command list and what `verify.js` needs.

## Deploying

`python build.py`, then commit and push to `main` — GitHub Pages serves the
branch directly, so a push **is** the deploy. Check the build number in the
page footer to confirm the live site picked up the change (it can lag a few
minutes behind a push).

## More documentation

- `CLAUDE.md` — rules, workflow and commands for anyone (human or Claude)
  working in this repo.
- `docs/REFERENCE.md` — what has been built and how, feature by feature.
- `docs/ARCHITECTURE.md` — data flow, module responsibilities, invariants.
- `docs/SUPPORT.md` — troubleshooting sign-in, permissions and stale builds.
- `docs/STATIONS.md` — what a floor station is and how to set one up.
- `docs/specs/README.md` — every feature spec, with date and status.
