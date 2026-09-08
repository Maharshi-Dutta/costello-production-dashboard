# Support and troubleshooting

Plain-language fixes for the problems that come up most. If something here
doesn't match what the code actually does, trust the code and fix this file
— see `CLAUDE.md`.

## Sign-in fails

The dashboard uses MSAL (Entra ID) against a fixed client and tenant. If
sign-in won't complete:

- Make sure the account is a real Costello account with access to the
  `ProductionProgress` SharePoint site. The dashboard shows exactly the jobs
  that account could already open in Excel — it grants nothing extra.
- Check the browser isn't blocking the sign-in popup (`loginPopup`). Some
  browsers/extensions block popups by default on a first visit.
- A stuck redirect usually clears with a hard refresh (see "stale build"
  below) — MSAL's redirect handling runs once at page load
  (`handleRedirectPromise`).

## "Permission needed" / phases (or the glass station) cannot be set

This means the signed-in account has the workbook scopes
(`Files.ReadWrite.All`) but not the SharePoint list scope
(`Sites.ReadWrite.All`) yet. The dashboard deliberately never pops a consent
dialog for this in the background — only a genuine click on a phase step (or
a station action) will ask for it, and only once per tenant, ever, needs to
be **granted org-wide**:

1. Sign in to the Microsoft Entra admin centre as a tenant administrator.
2. **App registrations** → find the dashboard's app registration → **API
   permissions**.
3. Confirm `Sites.ReadWrite.All` (delegated) is listed. If not, add it.
4. **Grant admin consent for `<the organisation>`** — this is the button
   that makes the permission apply to every user without each of them seeing
   an individual consent prompt.

Until this is granted, phases and the glass station both show a plain message
asking for the permission rather than failing silently or popping a consent
dialog nobody asked for. On the master dashboard this reads "Seeing the
floor's progress needs a SharePoint permission that has not been granted yet"
(the board, or the drawer's "Glass station" section); on the tablet it is
shorter — "Ask the office to grant the SharePoint permission." — because
there is no admin consent flow on a shared device and nothing more useful for
the floor to do than wait for the office to grant it. Either way, the fix is
the same four steps above, done once by a tenant administrator.

## "List not found"

The dashboard never creates a SharePoint list — `Dashboard phases`, `Glass
station`, `Station people` and `Station log` are all created by hand by the
admin, once (see `docs/STATIONS.md` → "Setting up the glass station"). If the
dashboard or the tablet says a list is missing:

- Check the list exists in the right site, with the exact display name the
  spec calls for (`Dashboard phases` in the workbook's own site; the other
  three — `Glass station`, `Station people`, `Station log` — in the separate
  `Floor stations` site, not the workbook's).
- Check the columns match the spec exactly (see
  `docs/specs/2026-09-07-phases-list.md` and `docs/STATIONS.md`) — the code
  reads columns by name and will not repair a mis-typed one.
- List ids are cached in `localStorage` (`cw_listids`) once found; a list
  that didn't exist yet is never cached as missing, so creating it and
  reloading the page (or tapping "Try again" on the tablet) is enough — no
  cache to clear.
- The three floor lists live in a site the code resolves independently of
  the workbook's (`cw_stationsite`, resolved by `CW.stationSite()`); if the
  *site* itself is missing rather than one list in it, both the tablet and
  the master dashboard say so separately — see the next section.

## Workbook locked / `InvalidSession`

- **Locked (423)**: usually means someone has the workbook open in the
  desktop Excel app without AutoSave on. Ask them to close it or turn
  AutoSave on; SharePoint's own co-authoring needs AutoSave to let the API
  in alongside a desktop session.
- **`InvalidSession`**: the workbook session the dashboard was using went
  idle and Graph expired it. The dashboard already retries this
  automatically — it drops the stale session, opens a fresh one, and
  re-sends the request once. If it keeps happening, it usually means the
  workbook is under heavy load; wait a few seconds and try again.

## `429` / `OperationQueueFull`

Excel's own API queues requests per session and refuses a flood. This shows
up on bulk operations (moving several jobs, a big batch of checkpoint
writes). The dashboard already retries with backoff on `429`, `500`, `503`
and `504`, and on the `OperationQueueFull` message specifically. If it still
fails after retries, it usually means several people are writing to the
workbook at once — spacing bulk actions out a little avoids it.

## Stale build

GitHub Pages caches `index.html` for up to about ten minutes, so a browser
can be sitting on an old build right after a deploy.

- The **footer build number** (`build YYYYMMDD-HHMM`) is the source of
  truth. `checkBuild()` polls `version.json` every two minutes and offers a
  "newer version available — Reload" banner once the server is ahead of what
  is loaded.
- If it doesn't show up on its own, a **hard refresh**
  (Ctrl+Shift+R / Cmd+Shift+R) forces the browser past its own cache.

## Alerts not sending

The dashboard only manages *subscriptions* (`Dashboard Alerts` sheet) — the
mail itself is sent by a Power Automate flow running an Office Script, set
up once by the admin. See `automation/SETUP.md` for the full click-by-click,
including its own troubleshooting table (no Automate tab, red squiggles
after pasting the script, "file not found", `[]` returned every time, mail
full of `<p>` tags, someone not getting mail, mail arriving at the wrong
time). Nothing about the mail flow lives in this app's own code — it cannot
be debugged from the dashboard.

## Station tablet says a site or list is missing

The glass station page resolves its own SharePoint site (`Floor stations`)
and its three lists (`Glass station`, `Station people`, `Station log`)
independently of the workbook's site, and says so plainly if any of them is
missing rather than throwing — one of "The Floor stations site is not there
yet…", "The 'Glass station' list is not in the Floor stations site yet…" or
the same for "Station people". This is expected on a tablet before the admin
has finished the setup in `docs/STATIONS.md` — walk through that setup, then
tap the tablet's "Try again" button (this also clears the cached site id, in
case it was the stale part).

## Station tablet: "Tap Sign out, then Sign in again"

This is the tablet's wording for an expired sign-in (MSAL's
`interaction_required`/`login_required`), not for a missing permission — it
means the station account's session itself needs refreshing, which happens
occasionally over weeks of a tablet being left signed in. Tap **Sign out** in
the header, then use the "Sign in again" button with the station account.
Nothing recorded is lost: any taps or log lines not yet sent are kept on the
tablet and go out as soon as it is signed in again.

## A person is not offered on the tablet's picker

The picker only offers rows from `Station people` where `Station` matches
`Glass` (not case-sensitive) and `Active` is exactly `Yes`. If someone is
missing:

- Check their row exists, `Station` reads `Glass`, and `Active` reads `Yes`
  — a blank or misspelled `Active` is treated as inactive, not as a typo to
  be forgiven.
- The tablet re-reads `Station people` every ten minutes on its own, and
  also on start-up, so a fix in SharePoint reaches it without anyone
  reloading anything — allow up to ten minutes, or reload the page.
- If they are offered but their steppers are greyed for a stage they should
  hold, check `Stages` — it is read forgivingly (comma, semicolon or slash
  separated, case-insensitive) but only `cut`, `hotmelt` and `glazed` are
  recognised; anything else is silently dropped.

## PIN not accepted on the tablet

The PIN is compared as typed digits against the `PIN` column for that
person, on the tablet itself — there is no server-side check and no
lockout, so a wrong PIN just shakes the dots and says "Try again". If it
keeps failing, check the `PIN` column in `Station people` for stray spaces
or the wrong digits. Remember the PIN is a deterrent on a shared device, not
a secret (see `docs/STATIONS.md`) — anyone signed in as the station account
can read the column it's compared against.

## A job is missing from the floor's board

Two different reasons, both normal:

- **The job is not "in production" yet** (or has left it). Only jobs whose
  section is "In production…" on the sheet are fed to the floor at all — a
  job in Ready to fit, Collect & supply, or any other section simply has
  nothing to show there.
- **It hasn't been fed yet.** The master dashboard is the only thing that
  ever pushes a job onto the floor's list, and only when someone has it open
  in a browser (`feedStation()` runs after every successful load). If
  nobody has opened the master dashboard since the job moved into
  production, the floor will not see it yet — opening `index.html` once is
  enough; the footer shows "station feed: just now" once it has run. The
  feed also skips a run when nothing has changed and the last one was under
  ten minutes ago, so a very fresh change can take a few minutes to appear.

## The office board or drawer says "cannot reach SharePoint — retrying"

This is different from a missing site or list: it means the site and lists
were read successfully before, but a later poll failed for some other
reason (the network, a rate limit, a bad gateway). The dashboard keeps
showing the last board or drawer it read rather than clearing the screen,
with this line above it, and keeps retrying on its own ten-second or
one-minute clock — there is nothing to click. If it does not clear up on its
own within a few minutes, check the tenant's own SharePoint/Graph status.

## The station account opened the master page

The station account has no access to the workbook at all. If it's used to
open `index.html` (the master dashboard) instead of `glass.html`,
`findFile()`/`downloadWorkbook()` fails with a 403 or 404, and the sign-in
overlay shows "This account has no access to the production workbook.
Station accounts use the Glass station page." with a link to `glass.html`,
rather than retrying in a loop. If that message doesn't appear, it's worth
checking the account genuinely has no workbook access — if it does, someone
gave the station account more than it should have.

## The log window shows nothing for a job

The `Station log` list is never deleted from, so on a long-lived
installation it can hold thousands of lines going back a long way. Neither
the drawer's timeline nor the log window looks back further than the last
90 days (`LOG_DAYS` in `station-core.js`) — a job whose only glass work
happened longer ago than that will show "Nothing recorded on the floor for
this job yet." even though the lines still exist in SharePoint. This is a
display window, not data loss: nothing is ever removed from the list itself.

## Rolling back

Two different kinds of rollback:

- **A bad edit to the workbook** (a wrong colour, a wrong section): open the
  **Versions** window in the dashboard, browse SharePoint's own version
  history, see the diff, and restore the version from before the mistake.
  Restoring is itself a new version, so it can be undone the same way.
- **A bad dashboard release**: `git revert` the offending commit, run
  `python build.py`, and push to `main` — the revert becomes the new deploy,
  same as any other push. There is no separate rollback mechanism for the
  code; the site is always whatever `main` currently is.
