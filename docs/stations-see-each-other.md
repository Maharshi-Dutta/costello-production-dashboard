# Stations see each other

Built 2026-10-02 on branch `cross-station-direct-read`. Brief, with its
Amendments after review: `specs/2026-10-02-stations-see-each-other.md`.
Decisions: HISTORY A44. Problems met: HISTORY B37, B38.

## One owner per fact

Every fact is written in one place. A screen that needs it reads that place.
Nobody copies it into a second list, so two screens cannot disagree.

| Fact | Owner (the only place it is written) |
|---|---|
| glass cut / hotmelt / Tuff counts | `Glass station` (workbook's site) |
| frames / sashes / transoms / door glazing counts | `Fabrication station` (`Floor stations`) |
| welding counts | `Welding station` (`Floor stations`) |
| glazed windows / astragal | `Glazing station` (`Floor stations`) |
| job facts (customer, totals, section, and whether a job has glass or windows) | the `Production` sheet, fed by the office |

The office no longer feeds any status word. The `Glass` column of
`Fabrication station` is unread and unwritten (kept on the list). The `Glass`
and `Fabrication` columns of `Glazing station` hold yes/no job facts (see
"Glazing" below), not status.

## The shared reader: `STU.stuListReader`

In `station-ui.js`. One helper, used by glass, glazing and fabrication.

- **Given:** `{ site: "own" | "floor", list, fields, tag, retryMs? }`.
- **Gives:** `read()` (answers true when something changed), `ready`, `items`,
  `missing`, `asOf`.
  - `ready`: false until the first read has answered. Say "checking".
  - `items`: the rows, or `null` when the list cannot be read. Never read
    `null` as "nothing done" or "ready".
  - `missing`: `items` is null because the list is positively not there.
- **Read only.** No write path. It never touches the page's own site id, token,
  `PROBLEM` or `SOFT`. Two readers on one page share nothing.
- **One read in flight.** A second `read()` while one is running answers false.
- **How it reads:** `CW.listDelta` on its own token. A list that refuses delta
  is read the plain way for five minutes.
- **Failures:**

| Answer | What the reader does |
|---|---|
| 403 | rows dropped (`items` null), retry after the retry clock (a minute) |
| 404 | rows dropped, `missing`, that channel's cached site forgotten, retry later |
| 410 | token dropped; a fresh enumeration on the next turn; not a failure |
| 429 | kept as is: last rows and token stay, ask again next turn (`graph.js` no longer calls it "refuses delta") |
| offline, 5xx | last good read kept, retry after the clock |

`graph.js` also waits out `Retry-After` (up to 60 s) with one back-off shared
by all list calls, and gives up on a list call after 30 s.

## The ticker: `STU.stuTicker`

One timer beats every second; `ST.tickDue` says whether a turn is due.

- **5 s** normally (`ST.REFRESH_MS`).
- **2 s for 20 s** after the tablet's own tap (`TICK.burst()`), and only then.
  A change seen in a poll does not start a burst (HISTORY B38).
- **In-flight guard.** A turn never starts while the last is running. One that
  has not come back after 5 minutes is given up on.
- No pause and no visibility check: a tablet polls all the time.

The office station poll is 5 s, always, and its five polls run together.

## What each page reads per tick

| Page | Lists read | Per minute at 5 s |
|---|---|---|
| glass (cut, hotmelt) | own board, `Fabrication station` | 24 |
| welding | own board | 12 |
| glazing | own board, `Glass station`, `Fabrication station` | 36 |
| fabrication | own board, assignments (by delta), `Glass station` | 36 |

Microsoft limits an account to about 600 requests a minute (3,000 in 5
minutes). Two floor accounts:

- **Glass account** (cutting, hotmelting, glazing): 24 + 24 + 36 = 84 a minute.
  Measured idle in the rig: 84.
- **Welding account** (welding and fabrication, ten tablets, worst case):
  measured idle 336 a minute. The burst only adds requests for the tablet that
  was tapped.

## What each station sees of the others

All per job, never per window: the floor's glass count is one DG + TG number.

- **Fabrication sees glass.** The glass chip and "Glass ready first"
  (`cw_fabglassfirst`) as before, from `Glass station` directly
  (`FABC.fbGlassOf`). "Checking" before the first answer, "not available" if
  the list cannot be read. A glass row off the sheet (not `OnSheet`, not
  `Active`) reads as no glass. A Door glazing line carries the job's chip.
  Take is held until an assignments read that started after the request (60 s
  ceiling).
- **Glass sees fabrication.** A line on each card, a red badge "fabrication
  done, glass waiting" and the "N waiting on glass" filter (`ST.fabOfJob`,
  `ST.glassWaiting`). The hotmelt tablet is badged for every waiting job, even
  when none of its glass is cut yet.
- **Hotmelt sees cutting.** A greyed read-only `Cut a/t` line, and `Tuff n/m`
  when the job has Tuff. Never tappable. A "Cut first" switch
  (`cw_glasscutfirst`, off by default, per tablet) puts jobs whose cutting is
  ahead of their hotmelting first (`ST.cutFirst`).
- **Glazing sees glass and fabrication.** Two chips, a READY TO GLAZE badge and
  a Ready first switch. Fabrication counts WINDOW groups only (door groups are
  left out). `GLZC.glzStatusNow` works the words out from the two lists.
  - **READY** (`glzReady`): nothing still owed by glass or fabrication, and at
    least one of them really done. Two blanks is not ready.
  - **Needs facts.** The office feeder writes `Glass` = yes/no and
    `Fabrication` = yes/no on each `Glazing station` row (`GLZC.glzNeeds`). A
    job with no row on a list is "checking", never ready, unless the fact says
    no. These are facts off the sheet, not a status copy.
  - The first office load after the ship PATCHes every Glazing row once.
- **The office boards** work the same words out from the lists the office
  already holds, with the same functions. The glass-colour and glazing chips
  redraw when those lists move.

## No idle lock

The ten-minute lock is gone on all four pages. The picked name stays across
reloads and new builds. The picker still returns when:

- somebody presses Switch;
- the person is removed from `Station people`, or holds nothing the page draws;
- a glass tablet's stage is changed.

Cost, accepted by the owner: a tap is logged under the last picked name. A day
sheet is the same. On fabrication, eligibility follows the picked person.

## Sign-in renewal

`CW.stationRenew(idle)` is called by each station page. When silent renewal
fails with "interaction required", the page tries once: a hidden `ssoSilent`,
then one full-page `acquireTokenRedirect` with `prompt=none`. It cannot loop:
the redirect time is kept in `sessionStorage` and there is no second one inside
an hour. It never leaves while a tap is queued. If it fails, the "Sign in
again" button shows as before; an attempt that did not get that far is retried
after a minute.

**Owner step:** register the four tablet page URLs (glass, welding, glazing,
fabrication) as SPA redirect URIs in Entra. Until then the renewal does
nothing. How long Microsoft keeps a session is tenant policy; it must be tried
on a real tablet over two days.

## Honest limits

- Status is per job, not per window or per door.
- Hand edits in Excel still lag about 40 to 60 s (SharePoint's downloadable
  copy). Dashboard clicks do not.
- A new job still reaches the tablets only through an office page that is
  open (the feeders run there). The next briefs: one `Jobs` list; the feeder
  and painter out of the office browser.
- A tablet's account must be a member of both sites to read both lists.

## See also

[[glass-station-overview]], [[glazing-station]], [[fabrication-station]],
[[welding-station]], [[fabrication-glass-and-door-glazing]], [[auth-and-graph]],
[[REFERENCE]], [[ARCHITECTURE]]
