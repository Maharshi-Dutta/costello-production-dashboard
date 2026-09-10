# Costello production dashboard — read this first

This is a static web app: no build step, no framework, no bundler, plain
ES2017 browser JavaScript. It is the production dashboard for Costello
Windows, reading and writing one shared SharePoint workbook (the `Production`
sheet) through the Microsoft Graph Excel API, plus four SharePoint lists for
state that must never touch the workbook. The GitHub repository is **public**.

Two pages ship from this repo:

- `index.html` + `app.js` — the master dashboard, used by the office.
- `glass.html` + `station.js` + `station-core.js` — the glass floor station,
  used on a shared tablet, shipped 2026-09-08 — see
  `docs/specs/2026-09-08-glass-station.md` and its v2 and v3, and
  `docs/STATIONS.md` for the built data model and setup. Check
  `docs/specs/README.md` for the status of every spec.

Read `docs/REFERENCE.md` (what has been built and how, feature by feature),
`docs/ARCHITECTURE.md`, `docs/SUPPORT.md`, `docs/STATIONS.md` and
`docs/specs/README.md` before making any change. They are kept accurate to
the code — if something here or there looks wrong, the code wins and the
docs need fixing.

## Non-negotiable rules from the owner

These override any convenience, any "it would be quicker to", any assumption
about what a feature "should" also do. They apply to every session, not just
the one that first wrote them down.

1. **The master `Production` sheet is only ever changed in two ways:**
   sanctioned fills (single-cell colour writes) and whole-row section moves
   (`moveJobRow` in `graph.js`, which re-finds the row by job number, copies
   it, and deletes the original — the only structural write in the app). No
   feature may write a value, a new row, a new sheet, or a colour anywhere
   else on `Production`, ever.
   There are exactly **two** sanctioned reasons to make a fill, and the owner
   gave each of them in a dated spec:
   - **checkpoint colours** — the office ticking work off in the drawer
     (`docs/specs/2026-09-04-checkpoints.md`);
   - **glass colours from the floor** — the office dashboard painting DG, TG,
     TUFF and NOT TUFF from what the floor has recorded on the tablet
     (`docs/specs/2026-09-10-glass-colours-two-way.md`, owner 2026-09-10:
     *"the production sheet is main doc, don't edit that as in text; by colour
     is ok"*). ARCH, ASTRAGAL, FANCY and EXTRA stay hand-ticked and are never
     written by that feature. Adding a **third** reason is a new decision for
     the owner, not a judgement call for a session.
2. **No other sheet on the workbook is touched** except the dashboard's own
   sheets, created and owned by the dashboard: `Dashboard Log`,
   `Dashboard Views`, `Dashboard Progress`, `Dashboard Alerts`. `Dashboard
   Config` is **read-only** for the dashboard — it is never created or
   written by code, only read.
3. **SharePoint lists** hold state that must never go near the workbook at
   all: `Dashboard phases` (hand-set phase per job, in the workbook's own
   site) and the floor's three — `Glass station`, `Station people` and
   `Station log` — in the separate private site `Floor stations`, which the
   station account can reach and the workbook's site cannot be reached from
   (shipped 2026-09-08). No list is created by code; if one is missing the
   feature says which one, plainly, and does nothing. **The office
   (`app.js`) never writes `Station people` or `Station log`, and never
   writes the floor's own columns of `Glass station`** (`Cut`, `Hotmelt`,
   `Glazed`, `Tuff` and their By/At pairs, `DoneBy`/`DoneAt`) — it only feeds
   that list's job-fact columns (which since 2026-09-10 include `TuffTotal`,
   the sheet's own TUFF quantity, and `OfficeDone`, the office's read-only
   lock on a job) and reads all three. The tablet, for its part, can never
   write `TuffTotal` or `OfficeDone`: `ST.floorOnly` drops both.
   There are **exactly two** exceptions, each granted by the owner in a dated
   spec, each for a named case; a **third** is a new decision for the owner,
   not a judgement call for a session.
   - **Seeding** (owner, 2026-09-08, the v3 spec's "Seeding" section and
     nowhere else): the feeder writes `Cut`/`Hotmelt`/`Glazed` on a row it is
     *creating*, or on a row whose `DoneAt` is empty, seeding them from the
     office's own glass checkpoints so a job already ticked off in the office
     does not arrive on the floor reading nothing done. It never writes a By,
     an At or the last-touch pair, and once the floor's first tap has set
     `DoneAt` it never writes a counter on that row again — a guard checked
     again with a one-item read immediately before each such write, not only
     when the plan was made. `Tuff` is **not** seeded.
   - **An office clear** (owner, 2026-09-10,
     `docs/specs/2026-09-10-office-clears-the-floor.md`): when the office
     **clears a job's glass checkpoints** — the moment its own record of that
     job's DG and TG goes from saying something to saying nothing — it writes
     that job's row's `Cut`, `Hotmelt`, `Glazed` and `Tuff` to **zero**, plus
     `DoneBy`/`DoneAt`, and **nothing else**. Only on a row the floor has
     really tapped (`DoneAt` set) and only with the office's answer to a
     confirmation naming what will be destroyed; a row the floor never tapped
     is left to the feeder. That is the whole exception. It is **not** a
     general "the office writes the floor" path: no per-stage `By`/`At`, no
     value but zero, no reconciliation pass, and never `Station people` or
     `Station log`. The body can only be built by `ST.officeClearFields(who,
     at)`, which takes a name and a time and so cannot carry a counter in.
   Only the tablet (`station.js`) writes a By, an At, a last touch or a log
   line.
4. **No phone number and no eircode leave the app in any export, ever.**
   `j.ph3` and `j.eir` (and the sheet's PHONE NO. / EIRCODE columns) are
   never read into an export path, in any format, under any filter or
   preset. `export.js` carries a standing test for this — do not weaken it.
5. **Every export is logged**: one `Dashboard Log` line per export, via
   `noteChange`, and nothing else about the exported data is stored anywhere.
6. **No real address, person's name or company domain in the public repo** —
   not in code, comments, tests, fixtures, docs, or commit messages — beyond
   what is already there: the SharePoint hostname constant `SITE_PATH` in
   `graph.js`, and the site/list/sheet *names* themselves (`Floor stations`,
   `Glass station`, `Dashboard phases`, `Dashboard Log`, and so on — names
   are fine, people and addresses are not). Use placeholders:
   `admin@yourcompany.example`, "the admin", "the colleague", "the station
   account".
7. **Consult the owner before coding any feature.** State "what I understood
   and what I will do" and get a yes before writing code — every spec in
   `docs/specs/` records that this happened and when.
8. **The manager loop** (see the `agent-manager` skill) is how features get
   built: brief written to `docs/specs/` → implementer agent → independent
   review agent → verify the result yourself → demo to the owner → owner
   approval → push. One build, one review, one fix, then ship — do not loop
   review/fix indefinitely; if a second real problem turns up, that is a new,
   small cycle, not an excuse to keep re-reviewing the same diff.

## Before any code change: the check-in rule (owner, 2026-09-08)

Before coding or changing anything, always write to the owner, in plain words:
1. **What the owner has done so far** on their side (accounts, sites, lists,
   permissions, approvals), as you understand it, so mistakes are caught early.
2. **What you are planning to do**, step by step, and what it will touch.
Then wait for a yes. This applies to fixes and small changes as well as
features. Never assume a step on the owner's side is done; ask.

## Running the checks

Every JS file must parse cleanly and every offline test suite must stay
green before anything is called done. Always run with `set -o pipefail` so a
failing test doesn't get masked by a later command in the same pipeline.

```bash
set -o pipefail
node --check parser.js && node --check graph.js && node --check checkpoints.js \
  && node --check export.js && node --check app.js \
  && node --check station-core.js && node --check station.js

node test_move.js
node test_checkpoints.js
node test_alerts.js
node test_export.js
node test_phases.js
node test_phases_list.js
node automation/test_digest.js
node test_station.js
node test_glasscolour.js

node verify.js   # dev-only cross-check, see below
```

`verify.js` compares the JS parser's output against a reference extract
produced by a Python reader of a real downloaded workbook. It needs two
files that are **not** in the repo and never should be (`live.xlsx` and a
`jobs.json` reference) — it is a local sanity check, not part of CI, and it
is fine for it to be unrunnable on a machine that doesn't have those files.

## Building and deploying

```bash
python build.py   # stamps a ?v=<timestamp> onto every script tag and writes version.json
```

`build.py` rewrites the cache-busting query string on every `<script src="…">`
tag matching `parser`, `graph`, `checkpoints`, `station-core`, `station`,
`export` or `app` — on **both** `index.html` and `glass.html` — and updates
the `<span id="build">` footer text on each. `glass.html` carries the same
build stamp as `index.html`, which matters more there than anywhere else: a
tablet left signed in for weeks is exactly where a stale cached script does
the most damage, and `glass.html`'s own `checkBuild()` polls `version.json`
and reloads itself once the build moves on (only when its write queue is
empty). Adding a script tag to a page that `build.py` doesn't already stamp
means extending its `SCRIPTS` regex or the `stamp()` calls at the bottom of
the file; it only touches the files and tags it's told to.

Deploy:

1. Run the checks above — every one green.
2. `python build.py`.
3. Commit, with the `Co-Authored-By` trailer, only when the owner has
   approved (see the manager loop above). Never commit or push on your own
   initiative mid-session.
4. Push to `main` — this **is** the deploy: GitHub Pages serves straight
   from the branch. There is no separate deploy step.
5. After a push, check the footer build number in the live page updates
   (GitHub Pages can lag by up to ~10 minutes; `checkBuild()` in `app.js`
   polls `version.json` every two minutes and offers a reload banner once
   the served build moves past what's already open).

## The offline test pattern

Every test file (`test_*.js`, `automation/test_digest.js`) runs in plain
Node with no network and no real Graph/SharePoint/Excel:

- The real source files (`parser.js`, `graph.js`, `checkpoints.js`,
  `export.js`, `station-core.js`, `station.js`, sometimes `app.js`) are
  loaded with
  `vm.runInThisContext(fs.readFileSync(...), { filename: ... })` — the
  actual shipped code runs, not a reimplementation of it.
- `global.window`, `global.localStorage` and (where needed) a stub DOM
  element factory stand in for the browser. `global.ExcelJS` is either a
  tiny fake or the real `exceljs` package (already a dependency, used to
  build in-memory workbooks for fixtures).
- `global.fetch` is a hand-written fake that inspects the method/URL/body of
  each call against an in-memory fixture (a fake workbook's cell ranges, or a
  fake SharePoint list) and returns canned Graph-shaped JSON. Several suites
  assert that fetch is *never* called down certain paths (e.g. the export
  path, or anything under `/workbook` from the phases-list feature) — that's
  how "no network" and "workbook never touched" get proven, not asserted by
  comment.
- Nothing here is rehearsed against a live tenant. See the next section for
  where that happens instead.

## Where scratch and rehearsal scripts live

Anything written to poke at a live token, a live list, or a live workbook
during development is **not** part of this repository. It lives outside the
repo entirely, in the session's scratchpad directory. `rehearse.js` in this
repo is the harness that such a scratch script drives (it takes a token and
a workbook reference on the command line and runs the real `graph.js`
against them) — but the token, the reference file and the decision of which
workbook to point it at are never checked in.

**Live rehearsals run against a OneDrive copy of the production workbook
first**, never the live SharePoint file, until a change has been demoed and
approved. Only after owner approval does a change get pushed and touch the
real workbook through real dashboard use.

## Further reading

- `docs/ARCHITECTURE.md` — data flow, module map, invariants, where state lives.
- `docs/SUPPORT.md` — troubleshooting for sign-in, permissions, stale builds, alerts.
- `docs/STATIONS.md` — what a floor station is, the glass station data model, admin setup.
- `docs/specs/README.md` — index of every feature spec, with status.
