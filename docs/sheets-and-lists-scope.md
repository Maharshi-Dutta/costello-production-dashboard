# Dashboard-owned sheets, and the SharePoint list scope split

### Dashboard-owned sheets

Created on first write, addressed only through a hard-coded constant so no
call site can ever be pointed at `Production` by accident:

| sheet | constant | written by |
|---|---|---|
| `Dashboard Log` | `LOG_SHEET` | `appendLog()` — every change, one line each |
| `Dashboard Views` | `VIEWS_SHEET` | `saveAssignment()` / `clearAssignment()` — custom groupings |
| `Dashboard Progress` | `PROGRESS_SHEET` | `saveProgress()` / `saveProgressMany()` — checkpoint counts |
| `Dashboard Alerts` | `ALERTS_SHEET` | `addAlert()` / `removeAlert()` — email subscriptions |
| `Dashboard Config` | `CONFIG_SHEET` | **read-only** — never created or written by code; holds the admin address, allowed export/company text, etc. |

### SharePoint lists and the scope split

Two kinds of Graph scope exist in `graph.js`:

- `SCOPES` (`Files.ReadWrite.All`, `User.Read`) — sign-in and every workbook
  call. Every dashboard user has these from the moment they sign in.
- `LIST_SCOPES` (`Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read`)
  — needed only for SharePoint list requests. Asked for quietly
  (`token(scopes, quiet=true)`) so a background list read never pops a
  consent dialog; asked for with a popup only from a genuine click
  (`listConsent()`), so a tenant that hasn't granted `Sites.ReadWrite.All`
  yet still signs in and uses everything workbook-related — the list
  features simply report that they need permission (`PHASE_NEED_CONSENT`).
  `headers()` picks the right scope set and drops the workbook session
  header for any path containing `/lists`, since a stale workbook session id
  on a list request earns an unrelated `InvalidSession` error.

Lists in use: `Dashboard phases` (hand-set phase per job), `Dashboard print
notes` (2026-09-09 — one extra note per job for the John print template:
Title = job number, `Note`, `By`, `At`; written only through `listUpsert`,
read only when the print-notes window opens) and the floor's three, shipped
2026-09-08: `Glass station` (job/glass facts plus the floor's counters),
`Station people` (who may record which stage) and `Station log`
(one line per counter write) — see `docs/STATIONS.md`.
`Dashboard phases` and `Dashboard print notes` live in the same SharePoint
site as the workbook. The
floor's three are meant to live in a *separate* site, `Floor stations`, so the
station account never needs any permission on the workbook's own site — but
creating that site needs an administrator the owner has not got yet, so for now
they live in the workbook's site too and the station account is a member of it
(which does mean it can reach the workbook; see `docs/STATIONS.md`).
`stationSite()` resolves whichever of the two is there, preferring `Floor
stations`, caches the answer in `cw_stationsite` with which site it is, and
re-checks every ten minutes while it is on the fallback so the pages move over
by themselves when the real site appears. The fallback is resolved by a plain
site lookup, never through `findFile()`: nothing on that path touches `/drive/`
or `/workbook`.

Both screens keep those lists current with `listDelta()` rather than
re-reading them: the first call enumerates and hands back a `deltaLink`,
every call after it passes that token and gets only what moved. A 4xx from
the delta endpoint — including the 410 Gone that carries a
`resyncChanges...` code when a token is too old — is re-thrown as one
recognisable error (`isDeltaRestart`), and the caller answers it by reading
the list once and starting a fresh enumeration.

## See also

- [[data-flow-and-lag]] — previous: data flow and the 36-second lag
- [[station-feeder]] — next: the station feeder and the station page
- [[glass-station-overview]] — the interim arrangement this section describes
