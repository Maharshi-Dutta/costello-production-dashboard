# Overview and process

The shape of the whole system, the documentation set itself, how tests/build/
deploy work, the scratch tooling that lives outside the repo, and the lessons
that cost time. Read this first; [[REFERENCE]] is the index that links here
and to every other topic note.

## 0. The shape of the whole thing

- **One static site, two pages**, served by GitHub Pages from `main`:
  `index.html` + `app.js` (the master dashboard, office) and `glass.html` +
  `station.js` (the Glass station, floor tablet). No bundler, no framework.
- **One workbook** on SharePoint (site `ProductionProgress`) is the source of
  truth for jobs. Its `Production` sheet is read by downloading the file
  (`graph.js: downloadWorkbook`) and parsed by `parser.js`.
- **Three kinds of state**, in order of how sacred they are:
  1. The `Production` sheet: only two sanctioned writes (§3, §4, §5 —
     [[sections-and-row-moves]], [[checkpoints-and-phases]]).
  2. Dashboard-owned sheets in the same workbook: `Dashboard Log`, `Dashboard
     Views`, `Dashboard Progress`, `Dashboard Alerts` (written), `Dashboard
     Config` (read only).
  3. SharePoint lists: `Dashboard phases` (workbook's site); `Glass station`,
     `Station people`, `Station log` (separate site `Floor stations`).
- **Sign-in**: MSAL browser, Entra public client, delegated Graph. Scopes are
  split so sign-in never depends on a permission the tenant may not have
  granted (§1 — [[auth-and-graph]]).
- **Tests** are offline Node scripts that load the real files with a fake
  `fetch` (§13, below).

## 12. Documentation set

| file | purpose |
|---|---|
| `CLAUDE.md` | rules, process, checks, deploy — read first |
| `README.md` | what it is, features, file map, how to run |
| `docs/REFERENCE.md` | index page into the linked notes on what was built and how (2026-09-22 reorg) |
| `docs/ARCHITECTURE.md` | index page into the linked notes on data flow, module map, invariants, where state lives (2026-09-22 reorg) |
| `docs/SUPPORT.md` | troubleshooting |
| `docs/STATIONS.md` | floor stations: data model and admin setup |
| `docs/specs/README.md` | index of every feature brief with status |
| `docs/motion-notes.md` | the radial menu's timing |
| `automation/SETUP.md` | the alerts flow, step by step |

## 13. Tests, build, deploy

| suite | covers | checks |
|---|---|---|
| `verify.js` | parser vs a Python reference extract (local files) | agree/disagree |
| `test_move.js` | row-move protocol | 7 |
| `test_checkpoints.js` | checkpoint logic and writes, incl. a hold that will not expire into a stale file | 37 |
| `test_alerts.js` | subscriptions, admin gate, pending holds | 30 |
| `test_export.js` | builders, no phone/eircode, logging | 53 |
| `test_john.js` | the John print sheet, template and print notes | 31 |
| `test_phases.js` | phase derivation and the wheel | 26 |
| `test_phases_list.js` | the phases list, scope split, dedupe | 35 |
| `test_station.js` | floor stations end to end (offline), incl. tuff, the lock and its release by a clear, the seed, the office clear's vocabulary and a queued tap meeting its zeros | 245 |
| `test_glasscolour.js` | glass colours into `Production`: the rule, last-writer-wins (per job, outside the settling window), idempotence, the cap, the backoff, fills only, the office clear end to end, the observed un-tick repaint, the office-absolute guard, a hold surviving a reload, and the office's Clear voiding the writer's paint | 77 |
| `automation/test_digest.js` | the alerts digest | 26 |

Pattern: `vm.runInThisContext` loads the real source; `global.fetch` is a
fake that serves Graph-shaped JSON and records every request; assertions
include "this was never requested". Always run with `set -o pipefail`.

Build: `python build.py` stamps `?v=<build>` on every script tag of both
pages and writes `version.json` (both pages poll it and offer/perform a
reload). Deploy = push to `main`. Commit only after the owner has approved a
demo; commit messages end with the `Co-Authored-By: Claude …` trailer.

## 14. Scratch tooling (outside the repo, in the session scratchpad)

- `graph.py`, `token.json`, `fileref.json`: a device-code token and the
  workbook reference for one-off Graph scripts.
- `rehearse.js` (in the repo) + `rehearse_*.js`, `live_*.js`: run the real
  `graph.js` against a OneDrive copy of the workbook, then one net-zero check
  on the live file. `fidelity.py`, `cp_fidelity.py`, `com_check.py`: compare
  a rehearsed workbook against the original, including drawn borders via
  desktop Excel.
- `make_phase_list.py`, `make_station_lists.py`: create/complete the
  SharePoint lists with the exact columns (idempotent; `check` mode).
- `stub_station.js` + `cors_server.py`: a fake sign-in library and fake Graph
  injected with Playwright so both pages run offline against a fixture for
  screenshots and behaviour checks (delta, real-time, gating).

## 15. Lessons that cost time

1. Never add a scope to the sign-in list; ask for it only on the requests
   that need it, quietly in the background, with a popup only from a click.
2. Decide "which scope" on the shape of a path, never on characters in it.
3. Excel keeps one border definition per shared edge; write the landing
   row's bottom and verticals, never its top, and re-assert the row above.
4. The downloaded file lags the API by ~36 s; hold your own writes locally
   until the file agrees.
5. `values` PATCH auto-parses "07/04" and "9434": guard text with an
   apostrophe.
6. `$batch` floods earn `OperationQueueFull`; cap concurrency.
7. A `| tail -1` masked a failing suite once; always `set -o pipefail`.
8. A queue that snapshots its keys loses the tap that lands mid-write;
   drain until empty and re-arm the retry whenever anything is owed.
9. Capture "before" values at the moment of the action, never at flush time.
10. A delta result must be discarded if the list was replaced while it was in
    flight (generation counter).
11. Every feature gets an independent review by a different agent before the
    demo; both blockers in the station work were found that way, not by tests.

## See also

- [[auth-and-graph]] — next: sign-in and the Graph layer
- [[REFERENCE]] — the index this note is linked from
- [[module-map-and-invariants]] — the module map these files describe
