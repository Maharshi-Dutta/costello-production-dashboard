# Floor stations

Everything in this document describes the **glass station**, the first floor
station, **shipped 2026-09-08** — built to
`docs/specs/2026-09-08-glass-station.md` and its second and third iterations
`docs/specs/2026-09-08-glass-station-v2.md` and
`docs/specs/2026-09-08-glass-station-v3.md` (each supersedes the one before;
see `docs/specs/README.md`). This document is kept accurate to the code
(`station-core.js`, `station.js`, `glass.html`, and the station parts of
`app.js`/`graph.js`) — if something here looks wrong, the code wins and this
file needs fixing.

## What a station is

A station is a separate, single-purpose page for one area of the factory
floor, used on a shared tablet signed in with a shared account that has **no
access to the production workbook at all**. It shows only what that area
needs — for the glass station: job number, customer name, **one total number
of glasses**, and progress on cutting, hotmelting and glazing — and lets the
floor record progress by tapping. There are no glass types on the floor at
all: DG, TG, TUFF and the rest are the office's business. It never shows job facts, comments,
prices, phone numbers, eircodes, counties, or anything from any other area.

The master dashboard is the **only** thing that ever writes job facts into a
station's data. It does this as a **feeder**: every time it loads the
workbook, it pushes an up-to-date slice into the station's SharePoint list.
A station page never talks to the workbook, and the feeder never touches a
station's own progress counters.

## The Glass station data model

Three SharePoint lists, all in a separate site (`Floor stations`, not the
workbook's own site): `Glass station` (the work), `Station people` (who may
record what) and `Station log` (what was recorded, by whom, when). None of
them is ever created by code — if one is missing, the pages say which.

### `Glass station`

**One item per job** (this changed in v3; it used to be one per job and glass
type, with a Title of `JOB|TYPE`). Both screens read the list one row per job:
where a job still has old `JOB|TYPE` rows beside its new one, **the row whose
Title is the job number wins and the older rows are ignored** — they cannot
put a job on the board twice, cannot feed the drawer, and cannot make a job
read as finished because some old ARCH row happened to be full. The old rows
are cleared out of the live list by hand (a one-off script) before this
iteration is deployed; the code tolerates them either way. Columns:

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, unique | feeder | the job number, upper-case, e.g. `R5303` |
| Job | text | feeder | the same job number |
| Customer | text | feeder | customer name, max 70 chars |
| GlassType | text | feeder | the literal `GLASS` on every row — the column is kept, nothing reads it |
| Total | number | feeder | **glasses on the job = DG + TG** |
| TuffTotal | number | feeder | tuff units on the job, off the sheet's own TUFF column. **Never added to `Total`** (added 2026-09-10) |
| Seq | number | feeder | the job's position in the master list, so the floor sees the office's order |
| Active | text | feeder | `Yes` while the job is in production and has glass, `No` afterwards |
| OfficeDone | text | feeder; **also set to `No` by an office clear** | `Yes` when the office has ticked this job's DG and TG off. The job is then **read-only on the tablet** (added 2026-09-10) |
| FedAt | text | feeder | ISO timestamp of the last feed that changed this item |
| FedBy | text | feeder | who was signed in to the master dashboard at the time |
| Cut | number | station; the feeder while the floor has not touched the row; **zeroed by an office clear** | glasses cut |
| Hotmelt | number | station, same | glasses hotmelted |
| Glazed | number | station, same | glasses glazed |
| Tuff | number | station; **never seeded**, but **zeroed by an office clear** | tuff units counted, against `TuffTotal` (added 2026-09-10) |
| CutBy / CutAt | text | station only | who last moved the Cut counter, and when (ISO) |
| HotmeltBy / HotmeltAt | text | station only | the same, for Hotmelting |
| GlazedBy / GlazedAt | text | station only | the same, for Glazing |
| TuffBy / TuffAt | text | station only | the same, for Tuff |
| DoneBy | text | station; **also an office clear** | the person's name from the last touch of any counter |
| DoneAt | text | station; **also an office clear** | ISO timestamp of that last touch |

The three **glass** stages are **Cutting, Hotmelting, Glazing**. (Toughening
was in the first iteration and is gone; there is no `Toughened` column any
more.) **Tuff** is a fourth counter beside them, added 2026-09-10: a different
department, the same person, its own quantity. It is deliberately *not* one of
the three, and the difference shows in three places — it is not part of the
job's `Total`, it does not decide whether a job is finished (the gold card,
the Finished group, the job row's `Glass 8/24` chip are all still the three
glass stages), and it is never seeded by the feeder. It *is* counted in the
tablet's "N left": eight glasses to cut plus eleven tuff to count is nineteen
things left for whoever holds both. In `station-core.js` the three are
`STAGES` / `STAGE_KEYS` and the four are `ALL_STAGES` / `ALL_STAGE_KEYS`.

**The lock (`OfficeDone`).** When the office has ticked every DG and TG item
of a job off, the feeder writes `OfficeDone = "Yes"` and that job goes
read-only on the tablet: every stepper on the card is greyed, `tap()` refuses
them, and the card says *"the office has marked this job finished"*. There is
no new control in the office for this — it is derived from the glass
checkpoints the drawer has always had, so **un-ticking one of them unlocks the
job**, and only the office can. ARCH, ASTRAGAL, FANCY and EXTRA are not asked:
they are hand-ticked and describe work the floor never sees.
A tap already sitting in the tablet's queue when the lock arrives is **dropped
rather than sent** — the office acted later — but never quietly: it is written
to the console and drawn on the card in red, naming the stage and the number
that was lost, until the office unlocks the job.

**Why DG + TG.** TUFF and NOT TUFF describe those same units — adding them
would send the floor to cut sheets that do not exist — and ARCH, ASTRAGAL,
FANCY and EXTRA are not glass the floor cuts, hotmelts and glazes. A job whose
only glass is one of those is never fed at all. One constant,
`TOTAL_TYPES` in `station-core.js`, says which kinds count.

**Seeding, the first of the two exceptions to "the office never writes the
floor's columns".** The office ticks a job's glass off in the workbook's own colours
long before a tablet appears on the floor, and a job already finished in the
office must not arrive on the tablet reading nothing done. So the feeder
writes `Cut`, `Hotmelt` and `Glazed` — and only those three — on **a row it is
creating, or a row whose `DoneAt` is empty** (the floor has never tapped it).
The first tap sets `DoneAt`, and from that moment the feeder never writes a
counter on that row again, whatever the office does afterwards. That guard is
checked **twice**: once when the plan is made from the list, and again with a
one-item read immediately before each write that carries a counter — the plan
is made from a single read and then sent as up to sixty writes, and a tap
landing in between must win. The same read skips the seed when the row's
`FedAt` is newer than this dashboard's own copy of the workbook, so a
dashboard holding a stale file cannot park a row on older numbers. The By/At
pairs and the last-touch pair are never the feeder's, seeded or not: they say
who did the work, and the office did not.

The seed is read from the glass checkpoints (`checkpoints.js`) by these rules,
in order:

1. every DG/TG item **gold** → all three counters at the total;
2. otherwise every DG/TG item at least **yellow** → `Cut` and `Hotmelt` at the
   total and **`Glazed` at nought** (owner, 2026-09-10). Yellow on a glass cell
   means "cut and hotmelted" since the glass-colours feature, so the row
   round-trips back to yellow instead of being flattened to blank — which is
   what used to happen, because a yellow item has no stored count. Glazed stays
   at nought deliberately: yellow says glazing is *not* done, and seeding it
   would make the cell read gold;
3. otherwise **any** item gold or yellow → all three at the office's own done
   count (a gold item counting its whole quantity). This is where a job with
   one type yellow and another blank lands: the floor's row holds one combined
   DG + TG number, so rule 2 would tell the floor that glass which still needs
   cutting is cut;
4. otherwise any item green ("cut") → `Cut` at the total, the other two nought;
5. otherwise nought.

It is `ST.officeSeed()`, a pure function, and `test_station.js` covers every
rule. `Tuff` is not seeded by any of them.

**An office clear, the second exception (owner, 2026-09-10 —
`docs/specs/2026-09-10-office-clears-the-floor.md`).** Un-ticking a job's glass
in the office cleared the workbook correctly but left the floor's counters
standing, so the tablet stayed gold for ever and the only way back was somebody
tapping `−` forty-nine times. So a clear now reaches the list: the office
writes `Cut`, `Hotmelt`, `Glazed` and `Tuff` to **zero**, plus `DoneBy`/`DoneAt`
and `OfficeDone = "No"` — seven fields, and nothing else, ever.

- **When.** Only when the office's own record of the job's **DG and TG** goes
  from saying something to saying nothing — the group **Clear**, or a per-item
  clear that leaves no other glass ticked. Clearing DG while TG is still gold
  is *not* a clear of the job's glass: the floor's row holds one combined
  number and telling them nought would be a wrong instruction on a workshop
  screen.
- **Only on a row the floor has really tapped** (`DoneAt` set). On an untouched
  row the counters are the office's own seed echoed back and the **feeder**
  puts them right on its next run — and writing `DoneAt` there would mark the
  row touched for ever and switch its seeding off.
- **The office is asked first**, in plain words naming what will go (*"The
  floor has recorded 49 cut, 49 hotmelted, 12 glazed on this job…"*). Answering
  no writes **nothing at all**, not even the workbook half.
- **`OfficeDone = "No"` rides along**, so the clear unlocks the card itself
  (owner, 2026-09-10: after an un-tick the card went correctly black with its
  counters at nought and then sat greyed, saying the office had marked the job
  finished, for over a minute). The lock used to be released only by the
  feeder's next run, and the feeder derives it from what the master currently
  shows — so while the master had not caught up it wrote the lock straight back
  on. Now the tablet frees the card on its next ten-second poll. This widens
  nothing: `OfficeDone` is a **feeder** column, the office's own, which the
  office already writes and the tablet never can (`ST.floorOnly` drops it).
- **`DoneBy`/`DoneAt` are written on purpose.** Last-writer-wins reads
  `ST.floorStamp` off those fields; counters dropped to nought under a stale
  stamp would read as old news and the tablet's "last touch" line would name
  the wrong person. The per-stage `By`/`At` pairs are **not** written — they
  say who did that stage's work, and nobody did.
- **No `Station log` line.** The clear is an office action and is recorded in
  `Dashboard Log`, as `Floor glass counters`, like every other one — and it is
  read back as one: that line, and the `DoneAt` the clear writes, are both the
  office's, so the colour writer counts them as an office action on the job.
  Counting the `DoneAt` for the floor and the log line for nobody is what made
  the office's own clear out-rank the office (`docs/REFERENCE.md` §18).
- **A tap already queued on the tablet is decided on the stamps.** A queued tap
  holds an *absolute* number, so letting it through after a clear would put the
  whole count back, not lay one tap on top of it. `rebaseQueue` compares the
  row's `DoneAt` with the tap's own time: a row moved *after* the tap was made
  wins and the tap is dropped; a tap made after the clear applies on top of the
  zeros. Nothing is marked BLOCKED either way — unlike the lock, a clear
  *unlocks* the job (`OfficeDone` goes to `No`), so a dropped tap can simply be
  tapped again.

It is `ST.officeClearFields(who, at)` — a name and a time in, seven fields out,
every counter a literal nought and the lock a literal `No`, so the path cannot
express anything else — with
`ST.floorWorkToClear()` and `ST.clearWarning()` beside it. In `app.js` the
whole of it is `clearFloorGlass()`, reached only from the two clear paths and
only after the office has confirmed.

Only job number, customer name and a number of glasses reach this list: no
glass type, no phone number, no eircode, no county, no price, no comment, no
product name or count.

### `Station people`

Maintained by the owner in SharePoint. Read by the tablet at start and every
ten minutes, and by the master for the log window's filters.

| column | type | meaning |
|---|---|---|
| Title | text | the person's display name, as shown on the tablet |
| Station | text | `Glass` |
| Stages | text | comma-separated stage keys, e.g. `hotmelt,cut` (one or more) |
| PIN | text | 4–6 digits; **empty means no PIN is asked for this person** |
| Active | text | `Yes`/`No`; only `Yes` rows are offered on the tablet |

**Everything the tablet enforces is a deterrent, not a secret.** The PIN is
compared on the tablet, against a column the station account itself can read,
so anybody who can sign in as the station account can read every PIN in the
list. The same goes for the stage gating and the ten-minute lock: they are
drawn and enforced by a page running on the device, so somebody with the
device and the will can get past all three. They exist to stop one person
tapping in another's name, or moving a stage that is not theirs, on a tablet
that is passed around a workshop — none of it is authentication, and nothing
of value should be protected by it. The real boundary is the station account
itself: it can reach the `Floor stations` site and nothing else, and no amount
of tampering with the tablet changes that.

### Interim: lists in the workbook's site

Creating a SharePoint site needs a Global Administrator, which the owner has
not got on this tenant yet. So, for now, **the three lists live in the
workbook's own site** (the site named by `SITE_PATH` in `graph.js`) and the
station account is a member of that site.

What that costs, stated plainly: **in this arrangement the station account can
also open the production workbook.** The separate site is what would have made
that impossible, and until it exists the isolation is a matter of the station
page not offering the workbook rather than the account being unable to reach
it. The owner accepted that for now. Everything else still holds — the tablet
still writes only its own counters and its own log lines, and the feeder still
never touches the floor's columns.

**No code change is needed either way.** `stationSite()` in `graph.js` looks
for the `Floor stations` site first and falls back to the workbook's own site
when it is not there (resolved by a plain site lookup — never through
`findFile()`, so nothing in that path goes near the file's drive). While it is
running on the fallback it looks for the real site again every ten minutes, and
switches over on its own the moment it appears, forgetting the list ids it
found in the old site so the lists are found again in the new one. Neither
screen says which site is in use, because to everyone using them it makes no
difference.

**The later move, when an administrator is available**, is four steps and no
deploy:

1. Create the private team site `Floor stations` (step 2 below).
2. Add the owner as an owner and the station account as a member (step 2).
3. Run the list-creation script against the new site (step 3 below).
4. Copy the rows across from the three lists in the workbook's site, and
   remove the station account from that site.

Both pages pick the new site up within ten minutes, or immediately on their
next reload. Nothing in the repository changes.

A tap that a tablet had not managed to send when the lists moved is not lost
and is not written to the wrong row: each queued tap is stamped with the site
it was made in, and after a move the row is found again by its Title — the job
number, which is unique and is the one thing about a row that does not change
between the two lists. A tap whose row did not come across at all is dropped
with a line in the console rather than written onto whatever row happens to
carry that item id in the new list.

### `Station log`

One line per counter write that the tablet actually sent. Written by the
tablet only; the master reads it and never writes it, and nothing anywhere
deletes from it.

| column | type | meaning |
|---|---|---|
| Title | text | job number |
| Station | text | `Glass` |
| GlassType | text | the literal `GLASS` — the column is kept, nothing reads it |
| Stage | text | `cut` / `hotmelt` / `glazed` |
| From | number | the counter before |
| To | number | the counter after |
| Who | text | the person's name |
| At | text | ISO timestamp of the tap that produced this value |

The tablet merges a run of quick taps on one row and stage into a single
write, so one line reads `5 → 8` rather than three lines counting up to it.

## What the office sees

From the master dashboard's "Sheet" dropdown (next to the search box),
picking **Glass station** replaces the job list with a read-only board: one
card per active job, sorted in the office's own order, with the job number,
customer, "12 glasses", the three counters (Cutting, Hotmelting, Glazing) each
reading "12 of 12" with who last moved it and when under it, "fed 3 min ago"
(or "not fed yet"), and — once anything has been recorded — a one-line "last:
Person A Cutting 5→8, 2 min ago" pulled from the log. A job whose three
counters have all reached the total goes **gold** and drops to the bottom of
the board, the same rule and the same colour as the tablet. Nothing on this
board can be clicked to change anything.

Opening a job's **drawer** (from the ordinary job list) shows a "Glass
station" section under "Glass units" with the same three lines, and under them
a timeline of that job's own log lines, newest first, capped at twelve, with
a "Full log" link. A job that has left production but was fed at some point
still shows its record here, worded "Finished on the floor" rather than "Not
fed to the floor yet".

The **Floor log window** (the button next to the board, or "Full log" from a
drawer) lists every `Station log` line, newest first, with four filters
(person, stage, job — free-text, matches as a substring — and day) and a
count of lines and units per person and per stage for whatever is currently
filtered. It pages 200 rows at a time with "Show more". Clicking a job number
opens that job's drawer, but only if the job is still on the sheet — a job
that has since left production stays as plain text so the window does not
close on nothing. It is read-only in the strongest sense: there is no code
path in the master dashboard that writes or deletes a line, and log lines are
never part of an export.

**Keeping up in real time.** A tap on the tablet is meant to reach the office
within about ten seconds. The tablet polls the `Glass station` list every
10 seconds; the master dashboard does the same for both `Glass station` and
`Station log`, but only while somebody is actually looking at the floor's
data — the station board is showing, the log window is open, or a job with
glass is open in the drawer — and drops back to once a minute otherwise. Both
sides use SharePoint's delta feed (`listDelta`) rather than re-reading the
whole list each time, merging in only what changed; a stale delta token (a
410) is answered by reading the list once and starting a fresh delta
enumeration, and a list that refuses delta outright is polled the plain way
for five minutes before being tried again. A failed poll never pops a toast or
clears the screen — the last board or log stays exactly as it was, with a
small "cannot reach SharePoint — retrying" line above it.

## Setting up the glass station (admin, one-off)

Do these in order. Every name below is a placeholder — substitute your own.

### 1. Create the shared station account

1. Microsoft 365 admin centre → **Users** → **Active users** → **Add a
   user**.
2. Give it a name that identifies the station, e.g. "Glass station" — not a
   real person's name.
3. Assign it a licence that includes SharePoint (e.g. Microsoft 365 Business
   Basic).
4. Give it **no admin roles** — it should be exactly as unprivileged as any
   other member account.
5. Set the password to never expire, or to whatever your organisation's
   shared-device policy requires.
6. Sign in with it once on the tablet, so it's a known, working sign-in
   before anyone relies on it.

### 2. Create the private team site

1. SharePoint start page → **Create site** → **Team site**.
2. Set visibility to **Private**.
3. Name it `Floor stations`.
4. Add the admin as an **owner**.
5. Add the colleague and the station account as **members**.

### 3. Create the three lists

1. In the `Floor stations` site, create three plain, generic lists (no
   template), named exactly `Glass station`, `Station people` and
   `Station log`.
2. Add every column from the three tables above, with the exact name and type
   shown (Title already exists on every list; the rest are new columns).
3. On the `Glass station` list's **Title** column, turn on **enforce unique
   values** — this is what lets the feeder upsert safely without ever
   creating a duplicate row for the same job/type. `Station log` must NOT
   have it: every line there is a separate record of the same job.
4. Fill in `Station people`: one row per person, `Station` = `Glass`,
   `Stages` a comma-separated list of `cut`, `hotmelt`, `glazed` and — since
   2026-09-10 — `tuff`, `PIN` 4–6 digits or blank, `Active` = `Yes`. Read the
   PIN warning above before relying on it for anything.

### 3a. Adding the 2026-09-10 columns to a list that already exists

The glass-colours feature needs five columns that the lists made in
September 2026 do not have. They go on `Glass station` and nowhere else:

| column | type |
|---|---|
| TuffTotal | Number |
| OfficeDone | Single line of text |
| Tuff | Number |
| TuffBy | Single line of text |
| TuffAt | Single line of text |

Add them by hand in the list's settings, or run the column script the session
wrote for this (it lives outside the repository, in the scratch area, like
`make_station_lists.py`). Either way: **rehearse on a copy of the list first,
and run it with the owner watching** (`CLAUDE.md` rule 7). Nothing needs
back-filling — a blank `OfficeDone` reads as "not locked", a blank `Tuff` as
nought, and the feeder fills `TuffTotal` in on its next run.

Until the columns exist, both pages still work: the reads ask for fields that
are not there and SharePoint simply does not return them, so no job gets a
tuff stepper and no job is ever locked. The feeder's writes of the two office
columns are what will fail, and they fail the way every refused feeder write
already does — counted, reported in the footer, retried on the next run.

### 4. Grant the app's SharePoint permission (once, tenant-wide)

The dashboard's app registration needs `Sites.ReadWrite.All` to read and
write any SharePoint list, including this one. If it hasn't already been
granted for the phase-list feature, grant it now — see
`docs/SUPPORT.md` → "Permission needed" for the exact steps (Entra admin
centre → App registrations → the app → API permissions → grant admin
consent). This is a single tenant-wide grant; it is not done per station.

### 5. Open the station page on the tablet

Open `glass.html` (not `index.html`) on the tablet, signed in with the
station account. It should never be able to open the master dashboard
successfully — see `docs/SUPPORT.md` if it can.

### 6. Optional: kiosk mode

Locking the tablet to just this page is optional but sensible for a shared
factory-floor device. Pick whichever fits your hardware:

- **Windows 11**: Settings → Accounts → Other users → Set up a kiosk →
  Assigned access, running Edge in kiosk mode pointed at the `glass.html`
  URL.
- **iPad**: Settings → Accessibility → Guided Access, started once
  `glass.html` is open in Safari.
- **Android tablet**: Settings → Security → Screen pinning, pinned to the
  browser tab showing `glass.html`.
- **Managed fleet**: an Intune kiosk profile pointed at the same URL.

## Adding a person

Add a row to `Station people` in SharePoint: `Title` = the name shown on the
tablet, `Station` = `Glass` (case doesn't matter, but it must match), `Stages`
= a comma-, semicolon- or slash-separated list of the stage keys they hold
(`cut`, `hotmelt`, `glazed` — any that don't match one of these three are
silently ignored, so a typo is a missing stage, not an error), `PIN` = 4–6
digits or left blank for no PIN, `Active` = `Yes`. The tablet re-reads this
list every ten minutes, so a new row (or a stage added to an existing one, or
`Active` flipped to `No` to retire someone) reaches the tablet on its own —
nobody needs to sign out or reload it. A person who is offered on the picker
but should not be able to move a stage is fixed the same way: edit `Stages`
in SharePoint, not on the tablet — there is nothing to edit there.

## Adding a future station

The intended path (not yet needed until a second station is built):

1. Extend the `STATIONS` array in `app.js` (next to `SHEETNAMES`) with the
   new station's name — this is what adds it to the "Sheet" dropdown on the
   master dashboard.
2. Create a new SharePoint list for it, following the same shape as
   `Glass station`: feeder-owned fact columns, station-owned progress
   columns, unique `Title`.
3. Copy `glass.html` (and `station.js`, adapted) into a new page for the new
   area, pointed at the new list.

This keeps every station's data isolated in its own list and its own page —
a station account for one area never needs, and never gets, access to
another area's data.

## What the feeder does, and when

The feeder runs inside the master dashboard, not the station page. After
every successful `load()` of the workbook (i.e. after the phases are read),
the dashboard:

1. Skips the run entirely if the signed-in account hasn't granted the
   SharePoint list permission yet (checked quietly, no popup).
2. Skips the run if the data hasn't changed since the last successful feed
   (a hash of the slice) and the last feed was under ten minutes ago.
3. Resolves the `Floor stations` site and the `Glass station` list; skips
   silently if either doesn't exist yet.
4. Reads the list once, works out which items need adding, which need their
   fact columns patched, which need their counters seeded (see "Seeding"
   above — a new row, or one the floor has never tapped), and which jobs have
   left production (and so need `Active` set to `No`) — never deleting an
   item. The hash in step 2 includes the seed, so a job the office has just
   ticked off reaches the floor on the next load rather than in ten minutes.
5. Sends the writes (a modest number per run, so one page load never floods
   the API), and remembers when it last ran.

Feeding is not restricted to the admin — any signed-in master user's browser
feeds the station list on load. Feed failures are logged to the console and
shown as a small status word in the footer; they never interrupt the
dashboard and never block on a retry loop.

## What the floor can and cannot do

**Can:**
- See every active job on the glass line, in the office's own order, with one
  total number of glasses each (no glass types).
- Type a job number or a customer name into the header's search box to narrow
  the cards down (Switch person clears it again for the next person).
- Pick their own name from `Station people` (and enter their PIN, if the
  owner has given them one). The chosen name locks itself after ten minutes
  without a tap, and there is a **Switch person** button in the header.
- Tap to record progress on **cutting, hotmelting and glazing**, per job, on
  the card itself: −, +, and one **All** (which becomes **None** at the
  total). There is nothing to expand and nothing to scroll inside a card.
  They may move **only the stages that person holds**; the stages they do not
  hold are shown greyed and disabled, values still readable, and a tap on one
  does nothing.
- Watch a job go **gold** and drop into the collapsed "Finished · n" group at
  the bottom when all three counters reach the total.
- Switch the tablet between the dark and light themes; it starts dark.

**Cannot:**
- See or change any job fact (customer, dates, comments, other checkpoints).
- Move a stage they do not hold.
- Delete anything, including a line of the station log.
- Export anything.
- See or reach the master dashboard, or any other station's data.
- Write anything to the production workbook — the station page never loads
  the workbook download path at all.

## Keeping the tablet signed in

The station account's sign-in token is refreshed silently while the tablet is
in use. After a long period (weeks, or a password change, or a policy reset)
Microsoft may demand a fresh interactive sign-in. The station page then shows
"The sign-in has expired. Tap Sign out, then Sign in again." with a "Sign in
again" button: tap **Sign out** in the header, then use that button to sign in
again with the station account. No data is lost: taps and log lines that
could not be sent are kept on the tablet (`cw_stationq`, `cw_stationlogq`) and
sent as soon as it is signed in again.

The person's name (picked from `Station people`, above) is stored on the
tablet only, in `cw_person`, alongside the time of their last tap — never the
stages they hold, which are always re-read from the list so an edit to the
tablet's storage can never hand somebody a stage that is not theirs. The name
is written next to every counter change as `CutBy`/`HotmeltBy`/`GlazedBy` and
`DoneBy`, and into the log's `Who` column, captured at the moment of the tap,
not when the write eventually goes out.
