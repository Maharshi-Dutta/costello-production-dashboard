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

The dashboard never creates a SharePoint list — `Dashboard phases`,
`Dashboard print notes`, `Glass station`, `Station people` and `Station log`
are all created by hand by the admin, once (see `docs/STATIONS.md` → "Setting
up the glass station"). If the dashboard or the tablet says a list is missing:

- Check the list exists in the right site, with the exact display name the
  spec calls for (`Dashboard phases` and `Dashboard print notes` in the
  workbook's own site; the other three — `Glass station`, `Station people`,
  `Station log` — in the separate `Floor stations` site, not the workbook's).
- `Dashboard print notes` (2026-09-09) is a plain list with Title (the job
  number, "enforce unique values" on), `Note` (multiple lines of text, with
  **"Use enhanced rich text" turned OFF** — the dashboard reads and writes
  plain text, and rich text comes back as HTML in the printed note), `By` and
  `At` (single line of text). Without it the John print sheet still prints —
  the notes typed in the window go into the file, they just are not
  remembered for next time.
- Check the columns match the spec exactly (see
  [`docs/specs/2026-09-07-phases-list.md`](specs/2026-09-07-phases-list.md) and `docs/STATIONS.md`) — the code
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
  separated, case-insensitive) but only `cut`, `hotmelt` and `tuff` are
  recognised; anything else is silently dropped.
- **Since 2026-09-21 each glass tablet shows one stage**, so somebody who only
  hotmelts is not offered the *cutting* tablet at all, and vice versa. Check
  the header: it reads `GLASS · CUTTING` or `GLASS · HOTMELTING`. If it is the
  wrong one, tap the `Cutting ▾` button beside it. Somebody whose `Stages`
  column still holds only the old `glazed` now holds nothing and appears on
  neither tablet — give them `cut` or `hotmelt`.
- If the card says *"the office has marked this job's glass finished"* and the
  Cutting stepper is greyed, that is the lock, not a fault: un-tick one of that
  job's glass checkpoints in the office to free it. **The Tuff stepper beside
  it is not greyed** and never is — the lock is about the glass.
- If the tablet shows the **"Which tablet is this?"** chooser on a device that
  has been working for weeks, its remembered answer has been cleared (site data
  wiped, a new browser profile). Answer it once and it will not ask again. A
  typo in a `?stage=` bookmark does *not* cause this: the page falls back to
  what the device already knew.

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

## Floor notes: what the channel is for, and when one does not appear

**What it is for.** The note box on a tablet's job card is for something about
**the job** that needs a person rather than a tick: short by two units, the
measurement on the sheet looks wrong, the glass arrived broken. It is **not**
for customer contact details — no phone numbers and no eircodes in a note,
ever. Nothing stops somebody typing one (it is a free-text box), and nothing
should: the list is internal, it is read in the drawer and it is never part of
any export, but the rule is the rule.

**Nobody can edit or delete a note**, on either screen — not the floor, not the
office. A note that was wrong gets a follow-up note saying so. That is
deliberate: this is a record of what was said, like `Station log`.

**A note the office cannot see.** In order:

1. **"The 'Station comments' list is not in the floor's site yet."** It has not
   been created. `docs/STATIONS.md`, `Station comments`, has the five-minute
   recipe. Until it exists the tablet shows the same line inside the composer
   and writes nothing anywhere.
2. **It is on another job.** The drawer shows one job's notes and nothing else.
3. **Give it a minute.** A note reaches the office on the floor's own poll —
   ten seconds while somebody is looking at the floor or has a glass job's
   drawer open, up to a minute otherwise. The Changes panel gets a line the
   moment it arrives.
4. **The tablet says "not sent — tap Send again".** Then it never left the
   tablet. The typing is still in the box; tapping Send again sends it. This is
   usually the workshop wifi.

**A glass tablet does not show a cutting tablet's notes** and vice versa — each
station sees its own. The job's drawer in the office is the one place they all
come together.

**The speech-bubble icon on a job row** (and on a Glass station board card)
means that job has notes **this screen has not opened yet**. Hovering it shows
them; clicking it opens the job at its Floor notes. Opening the job by any
route — the icon, the row, a jump link — marks that job's notes seen and the
icon goes. **It is per computer, not shared**: each office screen keeps its own
memory (in the browser, under `cw_notesread`), so two computers each show the
icon until each of them opens the job, and clearing the browser's site data
puts every note back to unread on that screen. Nothing about it is written to
SharePoint or to the workbook.

**The keyboard closes while you are typing a note, if another job finishes.**
This is a known limitation and it is **not fixed** — expect it. When another job
on the board finishes, or the office marks one done, the cards are re-ordered,
and moving a card closes the tablet's on-screen keyboard. **Nothing is lost**:
every word already typed is still in the box. **You will need to tap the box
again to carry on typing.** The dashboard puts the cursor back where it was, but
no web page can make a tablet re-open its keyboard on its own — only a finger
can — so the tap is unavoidable until the board is redrawn differently.

## What yellow and gold mean on the glass cells (changed 2026-09-21)

**Tell the office and the floor.** Glazing is the last step of the whole job,
not a glass step, so on 2026-09-21 it left the glass station. What the four
glass cells (DG, TG, TUFF, NOT TUFF) mean on the `Production` sheet changed
with it:

- **blank** — neither cutting nor hotmelting is finished on that job;
- **yellow** — **exactly one of the two is finished**;
- **gold** — **both are finished**.

Before that day, yellow meant "cut *and* hotmelted" and nothing went gold until
glazing was done, so **every job that was yellow turns gold, and some jobs that
were blank turn yellow**, the first time an office dashboard looks at them.
That is the change, applied to every job at once, and the owner asked for it.

**TUFF is not touched unless somebody on the floor has tapped it.** It has its
own count and is gold on its own, with no yellow of its own — but the dashboard
says nothing at all about a job's TUFF cell while the tuff counter has never
been tapped. That is deliberate: tuff is the one counter that never starts from
what the office has ticked, so a nought on it means "nobody has been there
yet", not "there is none done". A TUFF cell ticked gold by hand in the office
stays gold. Once the cutter taps tuff, the cell follows the counter from then
on — including back to blank if they tap it down to nought.

**A job with tuff still to count is not shown as finished** (owner, after the
demo on 2026-09-21). It does not go gold and does not drop into the Finished
group — on the office's Glass station board, on the job row's `Glass 8/16`
chip, and on the **cutting** tablet, which is the bench the tuff belongs to.
The **hotmelting** tablet is not affected: tuff never appears on it, so a job
that is hotmelted through is finished there whatever its tuff says. A job with
no tuff on it is finished on the glass stages, exactly as before. This is only
about how a card *looks*: the tuff count is still not part of the job's glass
total, the office's lock still has nothing to do with it, and the "N left"
number beside the search box is still this tablet's own stage alone.

**Glazing no longer appears on the glass tablets at all.** There is no glazing
stepper, no glazing number and no glazing person. Nothing that was recorded has
been deleted: the old counts are still on the list and the old log lines are
still shown in the Floor log window, labelled "Glazing". They simply decide
nothing any more. A glazing station of its own is a separate piece of work.

**There are now two glass tablets, one for cutting and one for hotmelting** —
see `docs/STATIONS.md` for how each is told which it is, and the entry above
about a person not being offered on a picker.

## The end-of-day sheet: who can change one, and what "waiting to send" means

New 2026-09-21. The **End of day** button is on the Cutting tablet only. It
asks for the day's four counts and a line about anything that got in the way.

**The cutter cannot change a sheet once it is saved. The office can.** That is
the owner's decision, and the tablet says so where the sheet was: *"saved 17:02
— ask the office to correct a mistake"*. Every time on both screens is the
clock on the wall here, not UTC. In the office, Show ▸ Glass station,
the **Day sheets** chip, **Edit** on the row: the four counts and the note, and
nothing else. The correction is stamped with who made it, shown on the row as
"edited by the office", and written to `Dashboard Log`. Nothing is ever deleted.

**"waiting to send"** under a saved sheet means the tablet has the sheet and
SharePoint has not got it yet — the wifi dropped, or the site was slow. It is
kept on the device and goes on its own within a few seconds of the connection
coming back; nothing is lost and nobody needs to retype it. It is safe to
reload the tablet, or to let it reload itself onto a new build: the sheet is in
the device's own storage and is sent again afterwards.

**"could not be saved — tell the office"** is a different thing and means what
it says: SharePoint has refused the same sheet three times, so it is not the
wifi and waiting will not fix it (a column renamed, the station account's
permission changed, the list locked). **The sheet is still on the tablet and is
not lost** — it will go the moment the list takes it. Tell the office; they can
read the numbers off the screen in the meantime.

**"saved" appearing for a sheet somebody else already saved** is the right
answer, not an error: one sheet per person per day is enforced by SharePoint as
well as by the page, and a second attempt (two tablets, or a retry that already
landed) is read back as the sheet that is there.

**A sheet started before midnight saves under the day it was started.** If
somebody types the sheet at 23:55 and taps Save at 00:05, it is filed under the
day they worked, not the new one — the form says which day it will save under,
in as many words, when that day is not today. An unsaved sheet is still there
the next morning for the same reason; the day after that it is cleared away.

**Save stays greyed until something has been written.** An untouched form is
not a day of four noughts. A day with nothing cut is still a real entry as long
as there is a line saying why ("machine down all day") — write that and Save
comes alive.

**"Whole numbers only, please"** — a count cannot be a minus, a decimal, a word
or anything over 9999. Nothing is rounded or guessed; Save simply waits for a
whole number.

**"no target set"** means nobody has set a weekly target for cutting yet. The
office sets it at the top of the Day sheets window; the tablet picks it up
within ten minutes, and the target saved with each sheet is the one that week
is measured against for ever after. **A target has to be one or more** — an
empty box or a nought is refused and nothing is written, because a target of
nought would read as every week beating it. Taking a target away again is not
something the dashboard can do; set it to the number you want instead.

**A correction being typed in the Day sheets window is safe from the poll.**
The window re-reads the floor's list every twenty seconds, but a row that is
open for editing is left exactly as it is until Save or Cancel — and anything
that arrived while it was open appears the moment it closes.

**"The 'Station day sheets' list is not in the floor's site yet"** on either
screen means exactly that: the two lists are made by hand in SharePoint (see
`docs/STATIONS.md`). Nothing is written anywhere until they exist, the board
and the job list are unaffected, and both screens pick the lists up on their
own within five minutes of them being made — no reload needed.

## The office cleared a job's glass but the tablet still shows the old counts

Clearing a job's glass in the office also puts the floor's counters back to
nought (`docs/REFERENCE.md` §18). It is one write to the `Glass station` list,
made **after** the workbook half has landed, and the office is asked to confirm
it first. So if the tablet is still showing 49/49/12 afterwards, one of these
happened:

1. **The list write was refused.** The dashboard retries it twice, 30 seconds
   apart, and then says so in a red message naming the job and the numbers
   still on the tablet. If you saw that message, the sheet was cleared and the
   floor was not — go to step 4.
2. **The tablet has not polled yet.** It reads the list every 10 seconds while
   somebody is using it. Give it a moment.
3. **The clear was interrupted** — the tab was closed, the browser crashed, or
   the machine slept in the second or two while the write was going out. There
   is **no automatic recovery for this**, and it is worth knowing why: a
   per-item tick that had not yet been sent is replayed on the next visit, but
   that replay clears the **workbook only** — it does not know the office was
   asked about the floor. A tick that was already in flight is not replayed at
   all, deliberately, because nobody can tell whether it landed. And **"All
   glass done" toggled off — the usual way to clear a job — has no replay of
   any kind**: if the tab closes mid-write, nothing is retried.
4. **The way back, in either case: tick the job's glass done again, then clear
   it again.** The second clear sees the floor's counters still standing, asks
   you to confirm, and puts them to nought. Nobody has to tap `−` forty-nine
   times on the tablet, and nothing needs to be edited in SharePoint by hand.

If the floor tapped the job **after** you cleared it, that is not this: the
floor's tap is the later action and it stands, by design. Clear it again if it
was tapped in error — a clear leaves the job unlocked, so they can also just
tap it back themselves.

## The job card says "quantity says 3 · 2 doors listed"

Not a fault, and nothing is blocked by it. The DRS quantity on the job's row
and the number of DOORS DONE cells with a code typed in them do not agree. One
of the two is wrong — either a code has not been typed yet, or the quantity
is out — and the dashboard will not guess which, so it says so and leaves it
to a person. **No colour in the sheet changes because of this message**, and
the doors that *are* typed go on working normally.

The same words appear when you hover the `9 / 3` numbers in the job list.

Two related things that are also not faults:

- **Quantity 0 with codes typed.** The DRS cell is left completely unpainted:
  as far as the sheet is concerned the job has no doors, so there is nothing
  for the dashboard to colour. The warning still shows, and each typed door
  still ticks off on its own.
- **A number or a date in a DOORS DONE cell is ignored.** Those are a quantity
  or a fitting date in the wrong column, not a door type, so no door appears
  for that cell and the dashboard never paints it.

## A WND or DRS cell went gold and none of its parts are ticked

That is correct, and it is the owner's own rule: a gold cell means the work is
finished, whether or not each window type or each door was ticked off one by
one. The line keeps its own tick **and** follows what is under it, and
whichever says more is what you see.

It only ever goes one way. Ticking a window type or a door can take that cell
to yellow or gold; **nothing under the line ever takes it back**, and nothing
under it ever repaints a gold cell white. Un-ticking the last door leaves the
DRS cell exactly as it was — to clear it, press **Clear** on the Doors line
itself.

## "Imported n doors from the sheet's colours" in Changes

Once, on the first load after the doors shipped, per browser. The door cells
already carrying a colour are read into the record so they are not mistaken
for somebody hand-painting them that morning; the line says `the sheet` as who
did it, because that is honestly all Excel can tell us. It never happens twice,
and it paints nothing — the colours are already there.

## "The sheet still does not show your change to … — it may not have saved"

A change made here is held on screen until the downloaded copy of the workbook
catches up with it — normally about 36 seconds. This red message means the
dashboard has re-read the file about a dozen times over roughly ten minutes and
it **still** shows the old value, so it has stopped holding the new one and is
telling you rather than quietly putting the old colour back.

It usually means the write did not land: the workbook was locked in desktop
Excel, the sign-in lost its edit rights, or the browser lost the network at the
wrong moment. **Open the sheet and look at the cell named in the message.** If
it is not what you asked for, make the change again. If it *is* what you asked
for, the message was the download lagging unusually badly and nothing is wrong.

Until that message appears, a held change is never dropped into a copy that
disagrees with it — including across a page refresh, which used to be the one
way to see an old colour come back (`docs/REFERENCE.md` §18).

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
