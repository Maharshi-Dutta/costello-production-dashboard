# Floor stations

Everything in this document describes the **glass station**, the first floor
station, **shipped 2026-09-08** — built to
[`docs/specs/2026-09-08-glass-station.md`](specs/2026-09-08-glass-station.md) and its second and third iterations
[`docs/specs/2026-09-08-glass-station-v2.md`](specs/2026-09-08-glass-station-v2.md) and
[`docs/specs/2026-09-08-glass-station-v3.md`](specs/2026-09-08-glass-station-v3.md) (each supersedes the one before;
see `docs/specs/README.md`). This document is kept accurate to the code
(`station-core.js`, `station.js`, `glass.html`, and the station parts of
`app.js`/`graph.js`) — if something here looks wrong, the code wins and this
file needs fixing.

## What a station is

A station is a separate, single-purpose page for one area of the factory
floor, used on a shared tablet signed in with a shared account that has **no
access to the production workbook at all**. It shows only what that area
needs — for the glass station: job number, customer name, **one total number
of glasses**, and progress on this tablet's own stage — and lets the
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
| OfficeDone | text | feeder; **also set to `No` by an office clear** | `Yes` when the office has ticked this job's DG and TG off. The job's **glass stages** are then read-only on the tablet and on the office's board; **Tuff is outside the lock** since 2026-09-21 (added 2026-09-10) |
| FedAt | text | feeder | ISO timestamp of the last feed that changed this item |
| FedBy | text | feeder | who was signed in to the master dashboard at the time |
| Cut | number | station; the feeder while the floor has not touched the row; **zeroed by an office clear** | glasses cut |
| Hotmelt | number | station, same | glasses hotmelted |
| Glazed | number | **nobody, since 2026-09-21** | glasses glazed. Left on the list with whatever is in it; read and written by nothing |
| Tuff | number | station; **never seeded**, but **zeroed by an office clear** | tuff units counted, against `TuffTotal` (added 2026-09-10) |
| CutBy / CutAt | text | station; **also the office's board steppers** (2026-09-21) | who last moved the Cut counter, and when (ISO) |
| HotmeltBy / HotmeltAt | text | the same | the same, for Hotmelting |
| GlazedBy / GlazedAt | text | **nobody, since 2026-09-21** | left exactly as they are |
| TuffBy / TuffAt | text | station; **also the office's board steppers** | the same, for Tuff |
| DoneBy | text | station; **also an office clear and an office board edit** | the person's name from the last touch of any counter |
| DoneAt | text | the same | ISO timestamp of that last touch |

The **glass** stages are **Cutting** and **Hotmelting**, one tablet each.
(Toughening was in the first iteration and is gone; there is no `Toughened`
column any more. **Glazing** was a third stage until 2026-09-21 and is now a
step of the whole job rather than a glass step — the owner's correction. Its
three columns stay on the list untouched and nothing reads or writes them; old
`Station log` lines saying `glazed` are still shown, labelled "Glazing".)
**Tuff** is a third counter beside them, added 2026-09-10: a different
department, the same person, its own quantity. It is deliberately *not* one of
the glass stages, and the difference shows in three places — it is not part of
the job's `Total`, it is never seeded by the feeder, and the office's lock
(`OfficeDone`) is never about it. In `station-core.js` the two glass stages are
`STAGES` / `STAGE_KEYS` and all three are `ALL_STAGES` / `ALL_STAGE_KEYS`.

**It does, since 2026-09-21, decide whether a job is shown as FINISHED.** The
owner, watching the demo: *"why are moving the jobs to complete when tuff is
left? if job has tuff and is not done dont move it, if done then move."* So a
job that has tuff on it (`TuffTotal > 0`) and has not counted it is not drawn
gold and does not sink into the Finished group — on the office's board, on the
job row's `Glass 8/16` chip, and on the **cutting** tablet, whose bench the tuff
is. The **hotmelting** tablet is unaffected: tuff is not that bench's work and
never appears on it. A job with no tuff finishes on the glass stages exactly as
before. One helper, `ST.tuffOwed(g)`, is the whole rule.

This **reverses one clause** of the 2026-09-10 decision and only that clause.
Tuff is still outside the job's glass total, still outside the lock, and still
outside the header's "N left" — that number is this tablet's own stage. What
changed is what a card is allowed to *look* like.

**The lock (`OfficeDone`).** When the office has ticked every DG and TG item
of a job off, the feeder writes `OfficeDone = "Yes"` and that job's **glass
stages** go read-only on the tablet: the Cutting and Hotmelting steppers are
greyed, `tap()` refuses them, and the card says *"the office has marked this
job's glass finished"*. There is
no new control in the office for this — it is derived from the glass
checkpoints the drawer has always had, so **un-ticking one of them unlocks the
job**, and only the office can. ARCH, ASTRAGAL, FANCY and EXTRA are not asked:
they are hand-ticked and describe work the floor never sees.

**Tuff is outside the lock** (2026-09-21). `OfficeDone` is the office's word
about the job's DG and TG; it has never said anything about tuff, which is a
different department counting a different quantity. That cost nothing while the
lock landed after glazing, by which time the tuff was long counted — the lock
now lands the moment hotmelting finishes, so a locked job is routinely one the
cutter is still counting tuff on. The Tuff stepper therefore stays live under
the lock, on the tablet and on the office's board alike.

A **glass** tap already sitting in the tablet's queue when the lock arrives is
**dropped rather than sent** — the office acted later — but never quietly: it is
written to the console and drawn on the card in red, naming the stage and the
number that was lost, until the office unlocks the job. A queued **tuff** tap
is never dropped by the lock, for the same reason its stepper stays live.

**Why DG + TG.** TUFF and NOT TUFF describe those same units — adding them
would send the floor to cut sheets that do not exist — and ARCH, ASTRAGAL,
FANCY and EXTRA are not glass the floor cuts and hotmelts. A job whose
only glass is one of those is never fed at all. One constant,
`TOTAL_TYPES` in `station-core.js`, says which kinds count.

**Seeding, the first of the two exceptions to "the office never writes the
floor's columns".** The office ticks a job's glass off in the workbook's own colours
long before a tablet appears on the floor, and a job already finished in the
office must not arrive on the tablet reading nothing done. So the feeder
writes `Cut` and `Hotmelt` — and only those two — on **a row it is
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

1. every DG/TG item **gold** → both counters at the total;
2. otherwise every DG/TG item at least **yellow** → `Cut` at the
   total and **`Hotmelt` at nought** (owner, 2026-09-10; re-read 2026-09-21).
   Yellow on a glass cell means "one of the two glass stages is complete", and
   cutting comes first, so the row
   round-trips back to yellow instead of being flattened to blank — which is
   what used to happen, because a yellow item has no stored count. Hotmelt stays
   at nought deliberately: seeding both would make the cell read gold;
3. otherwise **any** item gold or yellow → both at the office's own done
   count (a gold item counting its whole quantity). This is where a job with
   one type yellow and another blank lands: the floor's row holds one combined
   DG + TG number, so rule 2 would tell the floor that glass which still needs
   cutting is cut;
4. otherwise any item green ("cut") → `Cut` at the total, the other two nought;
5. otherwise nought.

It is `ST.officeSeed()`, a pure function, and `test_station.js` covers every
rule. `Tuff` is not seeded by any of them.

**An office clear, the second exception (owner, 2026-09-10 —
[`docs/specs/2026-09-10-office-clears-the-floor.md`](specs/2026-09-10-office-clears-the-floor.md)).** Un-ticking a job's glass
in the office cleared the workbook correctly but left the floor's counters
standing, so the tablet stayed gold for ever and the only way back was somebody
tapping `−` forty-nine times. So a clear now reaches the list: the office
writes `Cut`, `Hotmelt` and `Tuff` to **zero**, plus `DoneBy`/`DoneAt`
and `OfficeDone = "No"` — six fields, and nothing else, ever. (`Glazed` was a
fourth nought until 2026-09-21; a clear leaves it exactly as it is.)

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
  floor has recorded 49 cut, 49 hotmelted, 12 tuff on this job…"*). Answering
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
| Stage | text | `cut` / `hotmelt` — and `glazed` on lines written before 2026-09-21, still shown and labelled "Glazing" |
| From | number | the counter before |
| To | number | the counter after |
| Who | text | the person's name |
| At | text | ISO timestamp of the tap that produced this value |

The tablet merges a run of quick taps on one row and stage into a single
write, so one line reads `5 → 8` rather than three lines counting up to it.

### `Station comments`

Added 2026-09-15 (spec: `docs/specs/2026-09-15-station-comments.md`). One row
per **note** a floor worker leaves against a job — a shortage, a mistake on the
sheet, anything that needs a person's attention rather than a tick. Written by
a floor tablet only; the office reads it in the job's drawer and never writes
it, and nothing anywhere deletes from it.

Like `Station people` and `Station log`, this list is **every** station's, not
the glass station's: a second station page passes its own name and shares it.

| column | type | meaning |
|---|---|---|
| Title | text | a row id, `<JOB>\|<unix-ms>` — **not** a key: no unique rule, nothing looks a row up by it, two rows may share one |
| Job | text | the job number, upper-case |
| Station | text | which station wrote it: `Glass` today, free text so a new station is a new value rather than a schema change |
| Who | text | the person signed in at the tablet's picker |
| Text | multiple lines of text | the note |
| At | text | ISO timestamp, stamped at send |

**Append-only, on both sides.** There is no edit and no delete anywhere in the
code, and the tablet sends one POST per note. A mistake gets a follow-up note.

**A row with the `Station` column left blank shows on every station's tablet.**
That is deliberate, and it matches the rule `Station log` has always used: a row
can only get there by somebody typing it into SharePoint by hand, and hiding it
from every screen would lose a note rather than tidy one. The office drawer tags
such a row "floor" rather than naming a station. Filling the column in fixes it.

**This channel is for job problems, not customer contact details.** The note is
free text somebody types, so nothing stops a phone number being written into
it by habit — and nothing should be. The list is never part of any export.

**Creating it** (the owner, once): a plain list named exactly
`Station comments`, in the same site the other three floor lists are in, with
the columns `Job`, `Station`, `Who`, `At` as Single line of text and `Text` as
Multiple lines of text (`Title` already exists). **Do not** turn on
enforce-unique-values on Title — this list is a log, like `Station log`, not an
upsert target like `Glass station`. It starts empty. Until it exists, both
screens say so plainly and nothing is written anywhere.

### `Station day sheets`

Added 2026-09-21 (spec:
[`docs/specs/2026-09-21-day-sheets-and-station-reports.md`](specs/2026-09-21-day-sheets-and-station-reports.md)).
One row per **person per station-stage per day**: the cutter's paper
"Glass cutting work sheet", on the tablet. Written by a floor tablet, one POST
at the end of the day; the office reads it, may correct it, and never deletes.

Like the three lists above it is **every** station's — which stage has a sheet,
and what is counted on it, is the station **definition**'s business
(`ST.GLASS.daySheets`), not this list's. Today only `Glass` · `cut` has one.

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, **unique** | tablet | `<Station>\|<Stage>\|<YYYY-MM-DD>\|<Who>` — the key. A second save for the same key is refused by SharePoint as well as by the page |
| Station | text | tablet | `Glass` |
| Stage | text | tablet | `cut` |
| Day | text | tablet | `YYYY-MM-DD`, the tablet's **local** date at save |
| Who | text | tablet | the person signed in at the picker |
| Clear, KGlass, Satin, Obscure | number | tablet; office on a correction | sheets cut, whole numbers ≥ 0; blank on paper is 0 |
| Note | multiple lines of text | tablet; office on a correction | the free-text line, 500 characters at most |
| WeekTarget | number | tablet | the weekly target in force when the row was saved, so an old week stays true after the target changes. Not written at all when no target is set |
| SavedAt | text | tablet | ISO timestamp |
| EditedBy, EditedAt | text | office | set only when the office corrects the row |

**Append-only from the tablet**: one POST, no PATCH, no DELETE, ever. **The
office may PATCH `Clear`, `KGlass`, `Satin`, `Obscure`, `Note`, `EditedBy` and
`EditedAt` and nothing else** — the proof is the shape of
`ST.dayOfficeFields(counts, e)`, which takes counts, a note, a name and a time,
so there is no argument that could carry a Day, a Who or a Title in. Every
correction leaves one `Dashboard Log` line and **no `Station log` line**.

**Creating it** (the owner, once): a plain list named exactly `Station day
sheets`, in the same site the other floor lists are in, with `Station`,
`Stage`, `Day`, `Who`, `SavedAt`, `EditedBy`, `EditedAt` as **Single line of
text**, `Note` as **Multiple lines of text** (plain text), and `Clear`,
`KGlass`, `Satin`, `Obscure`, `WeekTarget` as **Number** (0 decimal places).
Then turn **enforce unique values on `Title`** — that is what makes a second
save of the same person-day impossible from two tablets at once. It starts
empty. Until it exists both screens say so plainly and nothing is written.

### `Station targets`

Added 2026-09-21, beside the one above. One row per station-stage: the weekly
target the office sets. **The office is its only writer**; every tablet reads it
and can never write it.

| column | type | written by | meaning |
|---|---|---|---|
| Title | text, unique | office | `<Station>\|<Stage>`, e.g. `Glass\|cut` |
| WeeklyTarget | number | office | total sheets per week — one number, not one per glass type and not one per day (owner's decision 1). **One or more**: the dashboard refuses an empty box or a nought, because a target of nought is not "no target" |
| SetBy, SetAt | text | office | who set it, when |

**Creating it**: a plain list named exactly `Station targets`, same site, with
`WeeklyTarget` as **Number** (0 decimals) and `SetBy`, `SetAt` as **Single line
of text**; enforce unique values on `Title`. It starts empty, and "no target
set" is a state both screens say in those words rather than showing "of 0".

## What the office sees

From the master dashboard's "Sheet" dropdown (next to the search box),
picking **Glass station** replaces the job list with a board: one
card per active job, sorted in the office's own order, with the job number,
customer, "12 glasses", the counters (Cutting, Hotmelting, and Tuff where the
job has any) each
reading "12 of 12" with who last moved it and when under it, "fed 3 min ago"
(or "not fed yet"), and — once anything has been recorded — a one-line "last:
Person A Cutting 5→8, 2 min ago" pulled from the log. A job whose two glass
counters have both reached the total goes **gold** and drops to the bottom of
the board, the same colour as the tablet.

**Since 2026-09-21 the office can work from it** (owner's decision 6): every
stage line carries `−`, `+` and `All`, clamped 0…total. One click writes five
fields of that `Glass station` row — the counter, its `By`/`At`, and
`DoneBy`/`DoneAt` — and leaves **one `Dashboard Log` line** ("Glass cutting" /
"Glass hotmelting" / "Glass tuff", from → to). It writes **no `Station log`
line**: that list is the floor's record of the floor's own work, the same
decision as the welding board's. Nothing on this path touches the workbook; the
sheet's glass colours follow on the colour writer's next pass, from the new
counters. A job the office has ticked off shows its **glass** steppers disabled
— *"glass complete — clear it in the job card"* — while its **Tuff** line stays
editable, because the lock is the office's word about the glass and says
nothing about tuff. Clicking a card's head opens that job's
drawer, unless the job has left the sheet, in which case there is no drawer to
open and the head is not clickable.

**The day sheets** (2026-09-21) have a chip of their own beside **Floor log**,
and the Glass station board carries a line under its head: "Cutting this week:
N of 250 sheets". The window shows one row per saved sheet — day, weekday,
person, the four counts, the total, the note, when it was saved, and "edited by
the office" where it was — grouped by ISO week (Monday to Sunday, `2026-W39`)
newest week first, each week with a subtotal row carrying its **own** target and
the difference. The target at the top of the window is the office's to set — a
whole number of one or more, or it is refused with nothing written; it writes
`Station targets` and logs "Cutting weekly target", from → to. **Edit** on a
row opens the counts and the note, and Save logs "Day sheet corrected"; a row
being edited is left alone by the twenty-second poll until it is saved or
cancelled. No day sheet is ever deleted, by either side.

**Report** (2026-09-21) is the chip beside them on any station's board: it opens
the Export window on the **Station report** template with that station already
picked. One template for every station, present and future — pick a
station-stage and a period (this week, last week, this month, custom dates) and
the file comes out with a Summary (a line per ISO week), Days, Jobs, Activity
and Notes, each sheet left out when there is nothing for it. Excel only for now.
Free text is stripped of anything that could be a phone number or an eircode on
its way into the file, and every report writes the usual `Dashboard Log` line.

Opening a job's **drawer** (from the ordinary job list) shows a "Glass
station" section under "Glass units" with the same three lines, and under them
a timeline of that job's own log lines, newest first, capped at twelve, with
a "Full log" link. A job that has left production but was fed at some point
still shows its record here, worded "Finished on the floor" rather than "Not
fed to the floor yet".

Under that, since 2026-09-15, a **Floor notes** section: every station's notes
about this job, oldest first, each tagged with the station, the person and the
time. Read-only — there is no reply box in this build. It is shown for **every**
job, not only a glass one, because the channel belongs to every station. A new
note also appears in the **Changes** panel, as "New floor note on `<job>`, from
`<station>`, `<who>`", within the same ten seconds (or a minute when nobody is
looking at the floor) as the rest of the floor's work. There is no new column
on the job row — per-job detail goes in the card (owner's rule, 2026-09-09).
Since 2026-09-16, a new note also shows a 💬 badge on the job row and on that
job's Glass station board card, until the job is opened on that screen.

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
   `Stages` a comma-separated list of `cut`, `hotmelt` and — since
   2026-09-10 — `tuff`, `PIN` 4–6 digits or blank, `Active` = `Yes`. Read the
   PIN warning above before relying on it for anything. (`glazed` stopped being
   a stage on 2026-09-21. A row that still holds the word is fine — it is
   ignored — but somebody whose `Stages` column holds **only** `glazed` now
   holds nothing and appears on neither tablet: give them `cut` or `hotmelt`,
   or set `Active` = `No`.)

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

### 5. Open the station page on the tablet, and tell it which one it is

Open `glass.html` (not `index.html`) on the tablet, signed in with the
station account. It should never be able to open the master dashboard
successfully — see `docs/SUPPORT.md` if it can.

**Since 2026-09-21 there are two glass tablets and one page.** Tell each one
which stage it is, in whichever of these three ways suits:

- open **`glass.html?stage=cut`** on the cutting tablet and
  **`glass.html?stage=hotmelt`** on the hotmelting one — the URL always wins,
  so this is the one to bookmark or pin in a kiosk profile;
- or open plain `glass.html` and answer the **"Which tablet is this?"** chooser
  that appears before the person picker. The answer is remembered on that
  device (`cw_stationstage`) and it is never asked again;
- or change it later from the **`Cutting ▾`** button in the header. It asks
  first and signs the person out, because the name on screen was picked for the
  stage that is leaving.

The header then reads `GLASS · CUTTING` or `GLASS · HOTMELTING`, and the
tablet shows only that stage: only the people who hold it, only its steppers,
only its numbers. The cutting tablet also shows **Tuff** on the jobs that have
any, for whoever holds `tuff`.

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
(`cut`, `hotmelt`, `tuff` — any that don't match one of these three are
silently ignored, so a typo is a missing stage, not an error, and so is the
`glazed` left over from before 2026-09-21), `PIN` = 4–6
digits or left blank for no PIN, `Active` = `Yes`. The tablet re-reads this
list every ten minutes, so a new row (or a stage added to an existing one, or
`Active` flipped to `No` to retire someone) reaches the tablet on its own —
nobody needs to sign out or reload it. A person who is offered on the picker
but should not be able to move a stage is fixed the same way: edit `Stages`
in SharePoint, not on the tablet — there is nothing to edit there.

## Adding a station

Written down on 2026-09-16, the day the **second** station (welding) was built,
and **followed for the third (glazing) on 2026-09-21** — which is the evidence
that it works: the glazing station is three new files, one `STATIONS` line, one
feeder, one poll and one renderer, and it changed nothing in the other two
stations' behaviour (their suites kept every count). Everything below was
actually done twice; where a file is named, that file already contains a worked
example to copy.

**What the third one added to this list** (read these before starting a fourth):

- `seedFields` may be **empty**. Glazing has no office record to seed from, so
  `GLAZE.seedFields` is `[]` and `seedOf` answers `{}`. That makes the feeder
  *simpler*, not harder: with no seed field in the plan, no patch can ever name
  a floor column, so there is no `weldFeedPatch`-shaped "re-read the row before
  seeding" guard to write at all.
- `daySheets` may be **absent**, and then the station has no day sheet anywhere
  — no button on the tablet, no chip in the office, no read of either list.
- Borrow before writing. Glazing's core calls `ST.floorOnly`, `ST.stripContact`,
  `ST.sectionInProduction`, `ST.atCmp`, `ST.feedPlan`, `ST.sliceHash`,
  `ST.logFields`, `ST.boardDiff`, `ST.mergeDelta`, `ST.stationPeople` and
  `ST.stationComments` rather than repeating any of them. Only the definition,
  the slice, the card shape, the clamp, the write body and the re-base are its
  own.
- **A quantity a station reads must come off `Production` alone.** `j.prodsMain`
  was the first Production-only field; `j.wndMain`/`j.drsMain` joined it for
  glazing. A station that reads `j.wnd`, `j.drs` or `j.prods` is reading a
  cross-sheet merge and is one inserted column away from [`HISTORY.md`](HISTORY.md) B20.
- If the station has an opinion about **where a job has got to**, it belongs in
  `effectivePhase`'s floor voice (`checkpoints.js`, `setFloorHook`), it is
  **computed, never stored**, and it must be **gated on the job still being in a
  live section**. A floor row outlives its job's time in production by design —
  it is never deleted, only deactivated — so an ungated voice asserts a stale
  phase for ever, and it is the voice that outranks the sheet (review finding
  G2, 2026-09-21).
- Whatever repaints the plain job list when your list moves must be wired on
  **every** station's redraw, not only the new one. The glazing build wired
  `redrawGlazing()` and left `redrawWelding()` returning bare, so welding-only
  jobs kept a stale badge (review finding G1).

Each station keeps its data in its own list and its own page — a station
account for one area never needs, and never gets, another area's data.

### 1. Write the station's definition

A **station definition** is a plain object. It is the only thing
`station-core.js` is told about a station, and it lives beside that station's
own rules, never inside `station-core.js`:

| field | what it is |
|---|---|
| `key` | the word in the URL and in `STATIONS` (`"glass"`, `"welding"`) |
| `name` | the word in the `Station` column of `Station people`, `Station log` and `Station comments` |
| `list` | the station's own SharePoint list |
| `site` | `"own"` = the workbook's own site, resolved by path; `"floor"` = the `Floor stations` site, resolved by path. Each is the only site that channel can ever answer. See step 7 below |
| `stages` | the stage keys this station's people can hold (glass has four, welding has one: `weld`) |
| `fields` | every column, for the `$select` of a read |
| `feederFields` | the job-fact columns the feeder owns |
| `floorFields` | the columns the tablet may write, **and nothing else** |
| `counterFields` | which of those are numbers rather than stamps |
| `seedFields` | the counters the feeder may seed on an untouched row |
| `feederOf(row)` | one slice row's job facts, in list shape |
| `seedOf(row)` | one slice row's seed, in list shape |
| `hashOf(row)` | everything about a row worth re-feeding for |
| `stageLabel(stage)` | the word a stage is called by, for headings and reports |
| `daySheets` | *optional* (2026-09-21). `{ <stage>: { counts: [[column, question, short], …], unit } }` — the end-of-day sheet this stage asks for. The **question** is what the tablet asks beside the box; the **short** label is what the office's table puts over the column (it falls back to the question when a definition gives none). **A stage with no entry gets no button, reads neither day-sheet list and sends no request for either.** That one line of definition is the whole of switching it on |
| `reportStages` | which stages a **station report** may be run for (glass: `cut`, `hotmelt` — not `tuff`, which is a counter rather than a station somebody reports on) |
| `reportJobs(data, stage)` | the report's **Jobs** sheet, as `{ columns, rows, jobs }` — the adapter. Glass answers one row per job; welding answers one row per job **and product group**. `jobs` is the same rows read as `{job, done, total}`, which is all the Summary needs. It is the only reason `stationReport` needs no branch per station |
| `reportLogStages(stage)` | which **`Station log` stages** feed one report stage. Glass answers `[stage]`; welding answers its part keys (`frames`, `sashes`), because its tablet logs one line per part and never the word "weld". Getting this wrong is invisible: the report simply comes out with no Activity sheet and an empty Days, which is how the welding report shipped at first |
| `reportGroupLabel` | *optional*: the heading for the product-group column of the report's Activity sheet, for a station whose log lines carry one (welding: `"Group"`). A station that names none loses the column rather than printing its own name down it |

`ST.GLASS` (in `station-core.js`), `WELDC.WELD` (in `welding-core.js`) and
`GLZC.GLAZE` (in `glazing-core.js`) are the three that exist. `ST.feedPlan`,
`ST.sliceHash` and `ST.floorOnly` take one; omitting it means the glass
definition, which is what they were about before there was a second station.

### 2. Three files

| file | what goes in it |
|---|---|
| `<station>-core.js` | the definition, the slice, the board/card shapes, the colour rule, the tap clamp, the write bodies, the re-base. **Pure**: no DOM, no Graph, no workbook. Loads in Node and the browser |
| `<station>.js` | the tablet page: the queue, the poll, the drawing |
| `<station>.html` | the tablet page's shell and styles |

What the pages share lives in `station-core.js` (lists, people, PIN, the log
line, the note channel, the delta merge, the board diff) and in
`station-ui.js` (the theme, the sign-in gate, the "Who are you?" picker and
its PIN pad). Copy nothing between station pages that could live in one of
those two instead.

### 3. The lists

- one list of the station's own, `Title` unique, with the feeder's columns and
  the floor's columns (the two sets must not overlap);
- `Station people`, `Station log` and `Station comments` are **shared** by
  every station in the site; they are never created per station. The station's
  `name` is what separates its rows.

No list is created by code, ever. A missing one says which list, plainly, on
whichever screen is asking, and writes nothing.

### 4. The people rows

One row per person in `Station people`: `Title` = the name, `Station` = the
definition's `name`, `Stages` = the stage keys they hold, `PIN` = digits or
blank, `Active` = `Yes`.

### 5. The feeder hook, in `app.js`

A `feed<Station>()` beside `feedStation()`, called from `load()`'s chain **in
its own link** so a failure of one feeder can never stop the other, nor the
checkpoint writes, nor the glass colour writer:

```js
.then(() => feedStation(), () => {})
.then(() => feedWelding(), () => feedWelding())
.then(() => feedGlazing(), () => feedGlazing())
```

The poll is the same shape: its own `try` inside `stationPoll()`, its own
tokens, its own three states, and nothing inside it may touch the other
station's.

### 6. The board entry, in `app.js`

One line in `STATIONS` (next to `SHEETNAMES`) and a renderer for its key in
`renderRows()`'s `if (state.board)` branch. That is what puts the station in
the Show dropdown and gives it the job list's place.

### 7. The site pin

`CW.stationSite(which)` honours the definition's `site`:

- **`"floor"`** resolves the `Floor stations` site **by path, and nothing
  else**. There is no fallback of any kind: a missing site is `null` and the
  page shows the quiet explained state.
- **`"own"`** resolves **the workbook's own site, by path, and nothing else**.
  The `Floor stations` site is never looked up on this channel — not on an
  empty cache, not after a forget, not when somebody taps **Try again**. That
  is the glass station, whose three lists are still in the workbook's own site
  (the interim arrangement of 2026-09-08).

Each channel keeps its own cached id (`cw_stationsite_own`,
`cw_stationsite_floor`), its own miss clock and its own "the lists moved"
counter, and neither can ever answer the other's site. `forgetStationSite(look,
which)` and `stationSiteMoves(which)` take the same word.

**Why it is a lookup and not a cache.** The first version of the pin returned
whatever site was cached and otherwise fell through to the old
resolver — which *prefers* `Floor stations`. That is not a pin. Three ordinary
things empty the cache: a 404 on any glass list call (which calls
`forgetStationSite`), localStorage cleared or a new device, and a browser
profile that has never opened the dashboard. After any of them the next resolve
landed on `Floor stations` and cached it, and the glass feeder, the glass
colour writer and `clearFloorGlass` were pointed at a site with no glass list
in it — permanently, from one 404. There is now no route from the `"own"`
channel to the other site at all.

**The only way the glass page ever moves** is somebody changing
`ST.GLASS.site` from `"own"` to `"floor"` in `station-core.js`, on the day the
three glass lists are copied into `Floor stations`. One word, and it is the
last step of that move.

### 8. Tests and docs

A `test_<station>.js` in the offline pattern, added to the verification command
in both `CLAUDE.md` files; a section in `REFERENCE.md`; the data flow in
`ARCHITECTURE.md`; a row in `specs/README.md`; and the station's own column
table in this file.

## The Welding station data model

Shipped 2026-09-16 (`docs/specs/2026-09-16-welding-station.md`). The welders'
tablet is `welding.html`; the office sees it at **Show ▸ Welding station**.

**Nothing in this feature writes the workbook.** The welding station never
paints a cell, and neither do the office's welding edits. The gold F/S cells
on the `Production` sheet are read **once**, as a seed, through the office's
own `Dashboard progress` record — never as a colour and never written back.

### `Welding station` — one row per job **and product group**

In the **`Floor stations`** site. `Title` is unique. A job with casement
windows and a PVC door has two rows; the job's facts are repeated on every row
of the job, so a row is complete on its own and the tablet builds one card per
job out of them.

| column | type | written by | meaning |
|---|---|---|---|
| `Title` | Single line, **unique** | feeder | `JOB\|GROUP`, e.g. `R5303\|CASEMENT WINDOWS` — the group upper-cased with whitespace collapsed |
| `Job` | Single line | feeder | the job number |
| `Group` | Single line | feeder | the product group as on the sheet, for display |
| `GroupSeq` | Number | feeder | the group's column order on the sheet, so a card lists its groups the sheet's way |
| `Customer` | Single line | feeder | customer name, max 70 characters |
| `Comment` | Single line | feeder | the sheet's COMMENT after the rule-3 strip, max 140 characters, may be blank |
| `SentToFloor` | Single line | feeder | the "sent to floor" cell as typed; a blank stays blank |
| `Wnd` / `Drs` | Number | feeder | the QUANTITY cells |
| `Frames` / `Sashes` | Number | feeder | the group's F and S cells: how many to weld |
| `Seq` | Number | feeder | the job's position in the master list |
| `Section` | Single line | feeder | the job's section on the sheet. The tablet shows `In production` only; the office board shows every section |
| `Active` | Single line | feeder | `Yes` while the job is on the sheet with this group having F or S > 0; `No` afterwards. **Never deleted** |
| `FedAt` / `FedBy` | Single line | feeder | the last feed that changed this row, and whose dashboard did it |
| `FramesDone` / `SashesDone` | Number | the tablet; the feeder as a **seed on an untouched row only**; the office from its welding board | welded so far |
| `FramesRemade` / `SashesRemade` | Number | **the tablet only** (2026-09-23); never the feeder, never the office | how many times a frame / sash of this group was welded again. A record, not progress: it moves no count and no colour, and has no ceiling. Tapping it stamps that part's By/At and DoneBy/At like any tap, and logs a `Station log` line with `Stage = frames-remake` / `sashes-remake` |
| `FramesBy` / `FramesAt`, `SashesBy` / `SashesAt` | Single line | the tablet; the office | who last moved that counter, and when (ISO) |
| `DoneBy` / `DoneAt` | Single line | the tablet; the office | the last touch of any counter on this row |

Do **not** switch on "enforce unique values" for anything but `Title`.

**Which product groups reach the floor.** Every F/S/T product group the parser
finds on the `Production` sheet, minus a deny-list of four: `ALU CLAD WINDOWS`,
`ALUCLAD TILT & TURN`, `BIFOLD`, `COMPOSITE`. So a new green group on the sheet
is fed with no code change, and a new red one is one line in
`welding-core.js`. Within a group, **F and S only**: T is never shown, however
the template colours it — the owner's words, "there is no transomes even if it
is green".

**THE SLICE READS `Production` AND NOTHING ELSE.** The parser keeps two sets of
product counts per job: `j.prods`, the **maximum across every sheet**, which is
what the drawer, the checkpoints, the exports and the John sheet have always
read; and `j.prodsMain`, the `Production` sheet's own numbers, never merged.
**The welding slice reads `prodsMain`.** A job with no `prodsMain` is not on
`Production`, so it is not the welding floor's business and is not fed.

This is not a nicety. On 2026-09-17 a column was inserted into `Production`;
`Production (2)` is formulas that moved with it and headers that did not, so on
that sheet every number ended up one column to the right of the header
describing it. The cross-sheet maximum took the neighbours: one job gained four
product groups it has not got and a sashes count from the column next door, and
the feeder wrote 300 wrong rows into `Welding station`. See `docs/HISTORY.md`
B20. If a station ever needs the counts, it takes them from the main sheet.

**Components a group has not got.** `WELD_GROUP_PARTS` in `welding-core.js`
overrides the default `["frames","sashes"]` for a named group. Today it holds
one entry — `SUPER DOOR` is welded in **sashes only** — so a Super door's
frames number is fed as `0` rather than carried, and no Frames line is drawn
for it on the tablet or on the office board. Adding another is one line.

**Seeding.** Only when the feeder is creating a row, or when `DoneAt` is empty
(checked twice, as the glass list's seed is): `FramesDone = Frames` when the
office's own record for `prod:<group>:f` says `done` (gold), else `0`; the same
for sashes with `:s`. **Yellow seeds nothing.** The moment the floor's first tap
sets `DoneAt`, the feeder never writes a counter on that row again.

**Rule 3 and the COMMENT.** The comment column is free text and does carry
phone numbers. It is stripped on the way into the list — every run of six or
more digits and anything eircode-shaped becomes `…` — so no phone number and no
eircode is ever stored where the floor could read one.

**The colours.** Three levels, the same three words at each: no colour =
nothing welded, **yellow** = started, **green** = every one welded. A
Frames/Sashes line, then the product group, then the job card.

### Welding rows in the shared lists

- `Station people`: `Station` = `Welding`, `Stages` = `weld` (one stage —
  "no cutting, it should just have welding").
- `Station log`: `Station` = `Welding`, `GlassType` = the product **group**
  (the column name is kept; nothing creates columns), `Stage` = `frames` or
  `sashes`, `From`/`To`/`Who`/`At` as before. **Written by the tablet only.**
- `Station comments`: `Station` = `Welding`, exactly as for glass.

### What the office can do, and what it cannot

The office's welding board shows one row per job, every section, in sheet
order, with an overall bar, `Frames x/y`, `Sashes x/y`, the last touch and the
unread-notes count. Opening a row gives **−, +, All, None** on each group's
Frames and Sashes.

An office edit writes exactly five fields — that part's counter, that part's
`By`/`At`, and `DoneBy`/`DoneAt` — and leaves **one `Dashboard Log` line** per
change ("Welding: R5303 CASEMENT WINDOWS frames, 3 → 6"). It writes **no line
of `Station log`**: that list is the floor's and stays the floor's. The tablet
sees the change as the row's last touch on its next ten-second poll, and a
queued floor tap made *before* the office's edit is dropped and said so on the
card.

## The Glazing station data model

Shipped 2026-09-21 ([`docs/specs/2026-09-21-glazing-station.md`](specs/2026-09-21-glazing-station.md)). The **third**
station, and the first one built from "Adding a station" above rather than from
a page. The glazer's tablet is `glazing.html`; the office sees it at
**Show ▸ Glazing station**.

Glazing used to be the glass station's third stage. The owner took it out on
2026-09-21: *"glazing is the last step for the whole job, not just glass, so it
will be its own dashboard with its section."* The glass list's `Glazed`,
`GlazedBy` and `GlazedAt` columns are **not** reused and are not read by this
station — they stay on `Glass station` exactly as they are, written by nothing.

**Nothing in this feature writes the workbook**, from either side. The only
workbook write anywhere in it is the `Dashboard Log` line an office edit leaves,
which is a dashboard-owned sheet (rule 2).

### `Glazing station` — one row per **job**

In the **`Floor stations`** site. `Title` is unique.

| column | type | written by | meaning |
|---|---|---|---|
| `Title` | Single line, **unique** | feeder | the job number, upper-case, e.g. `R5303` |
| `Job` | Single line | feeder | the same job number |
| `Customer` | Single line | feeder | customer name, max 70 characters, **stripped** per rule 3 |
| `Section` | Single line | feeder | the job's section on the sheet. The tablet shows `In production` only; the office board shows every section |
| `Seq` | Number | feeder | the job's position in the master list, so the floor sees the office's order |
| `Active` | Single line | feeder | `Yes` while the job is on the sheet with something to glaze; `No` afterwards. **Never deleted** |
| `Windows` / `Doors` | Number | feeder | the job's quantities **on the `Production` sheet** |
| `Total` | Number | feeder | `Windows` — the units to glaze. It was `Windows + Doors` until 2026-09-23, when the owner took the doors out of the count; `Doors` is still fed as a fact and is counted nowhere |
| `Comment` | Multiple lines of text | feeder | the sheet's COMMENT after the rule-3 strip, max 140 characters, may be blank |
| `FedAt` / `FedBy` | Single line | feeder | the last feed that changed this row, and whose dashboard did it |
| `Glazed` | Number | the tablet; the office from its glazing board | units glazed so far, 0…`Total` |
| `GlazedBy` / `GlazedAt` | Single line | the same | who last moved the counter, and when (ISO) |
| `DoneBy` / `DoneAt` | Single line | the same | the last touch of any counter on this row |

Do **not** switch on "enforce unique values" for anything but `Title`.

**THE QUANTITIES COME OFF `Production` AND NO OTHER SHEET.** `j.wnd`/`j.drs`
take the *first* sheet that has a number, which is `Production` whenever
`Production` has one and somebody else's sheet when it has not — the same door
[`HISTORY.md`](HISTORY.md) B20 came through. The parser therefore carries
`j.wndMain`/`j.drsMain`, the Production-only pair, beside `j.prodsMain` which
welding reads, and **the glazing slice reads those**. A job with no `Production`
quantity at all reads nought, so it is not this station's business and is not
fed. A job whose `Total` would be `0` is never fed.

**There is no seed**, and that is a fact about the office rather than an
omission: glazing was never an office checkpoint, so there is no record anywhere
of a job already glazed. `GLAZE.seedFields` is empty and `seedOf` answers `{}`,
so the feeder's whole vocabulary holds not one floor column — which is why this
station needs no `weldFeedPatch`-shaped guard. An untouched row starts at
nought.

**The colours.** Three levels, the same three words welding uses: no colour =
nothing glazed, **yellow** = started, **green** = every unit glazed.

### Glazing rows in the shared lists

- `Station people`: `Station` = `Glazing`, `Stages` = `glaze` (one stage).
- `Station log`: `Station` = `Glazing`, `GlassType` = the literal `GLAZING` (the
  column name is kept; nothing creates columns), `Stage` = `glaze`,
  `From`/`To`/`Who`/`At` as before. **Written by the tablet only.**
- `Station comments`: `Station` = `Glazing`, exactly as for the other two.
- `Station day sheets` / `Station targets`: **not used.** `GLAZE` has no
  `daySheets` entry, so the tablet draws no button, reads neither list and sends
  no request for either.

### Creating the list (the owner, once)

A plain, generic list named exactly `Glazing station`, in the `Floor stations`
site, with these columns (`Title` already exists):

| column | type |
|---|---|
| Job, Customer, Section, Active, FedAt, FedBy, GlazedBy, GlazedAt, DoneBy, DoneAt | Single line of text |
| Comment | Multiple lines of text (plain text) |
| Seq, Windows, Doors, Total, Glazed | Number (0 decimal places) |

Then turn **enforce unique values on `Title`** — that is what lets the feeder
upsert safely without ever creating a duplicate row for a job. It starts empty.
Until it exists both screens say so plainly and nothing is written anywhere; no
list is created by code, ever.

Then add the glazer to `Station people`: `Title` = the name shown on the tablet,
`Station` = `Glazing`, `Stages` = `glaze`, `PIN` = 4–6 digits or blank,
`Active` = `Yes`. Read the PIN warning above before relying on it for anything.

### What the office can do, and what it cannot

The office's glazing board shows one row per job, every section, in sheet order,
with the counter, a bar, the windows/doors breakdown, the last touch and the
unread-notes count, and **−, +, All, None** on the row itself. Finished jobs go
gold and sort last. The card's head opens that job's drawer, unless the job has
left the sheet, in which case there is no drawer to open and the head is not
clickable. **Floor log** sits under the board, read-only and filtered to
Glazing; **Report** is the usual chip beside the board.

An office edit writes exactly five fields — `Glazed`, `GlazedBy`, `GlazedAt`,
`DoneBy`, `DoneAt` — and leaves **one `Dashboard Log` line** per change
("Glazing: R5303, 3 → 6"). It writes **no line of `Station log`**. The row is
read once immediately before the PATCH, so a tap the board had not yet seen is
never overwritten; a double-click is one write, because the busy flag goes up
before the first await.

The **job drawer** gains a read-only **Glazing** line under the Welding line:
`n / Total`, who last moved it and when, and a button to the board. There is no
new column on the job row (owner's rule, 2026-09-09).

### The phase bar hears the floor (2026-09-21)

The owner: *"when glazing is in process or done the phase bar in the master
dashboard should move the progress into the In glazing section. And when
something is in welding, in process or done, it should be In fabrication. If
something is marked in glazing it means it has gone through fabrication; if
glazing is not marked but welding is, it is in fabrication."*

So the floor is a **third voice** in `effectivePhase` (`checkpoints.js`), beside
the sheet's own evidence and the hand-set phase, and it is read the same way:

- any welding recorded on the job (frames or sashes done > 0) → at least
  **In fabrication**;
- any glazing recorded (`Glazed` > 0) → at least **In glazing**;
- glazing outranks welding;
- the result is the **highest** of the three, so the floor can only ever move a
  job forward and never past what the sheet or a person says;
- a counter tapped back to nought withdraws that voice and the phase falls back
  to the next highest;
- **and the floor only ever speaks for a job still in a live section.** A floor
  row is never deleted — it goes `Active = No` when its job leaves the sheet,
  and its counter stays exactly where the floor put it. Nothing clears a
  finished glazing row when the office moves the job on by hand, so without this
  a job in Ready to fit would read "In glazing" for ever. `floorPhaseRec` is
  gated on `ST.inProduction(j, BLOCKNAMES)` — `ST.sectionInProduction` of the
  job's own section, the same test the tablets use, plus the `past` guard — and
  it gates the welding record and the glazing record alike. Once a job leaves
  that scope **the floor says nothing**; it does not clear, reset or delete
  anything, and the phase falls back to the sheet and any hand-set phase exactly
  as it did before this station existed.

**Nothing is stored for it.** It is computed on every render out of lists
already in memory: no `Dashboard phases` row, no list write, no workbook write.
`test_glazing.js` asserts zero requests of any kind across the whole phase
section. The drawer's phase strip says so in words when the phase is the
floor's — "Moved here by the floor — from the glazing station" — the way it
already distinguishes hand-set from sheet.

The two floor lists may not have been read when the job list first draws, which
is deliberately harmless: a station that has read nothing says nothing, so the
phase starts at whatever the sheet says and is *raised* when the lists arrive —
it never flickers backwards. The repaint goes down the existing quiet path
(`chipsNow()` carries each drawn row's floor phase, `floorPhaseRepaint()` takes
it), and no timer was added for it. **Both** `redrawWelding()` and
`redrawGlazing()` take that branch when they are off their own board, so a
welding-only job's badge does not have to wait for the glass list to move.

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
- Tap to record progress on **this tablet's own stage** (cutting or
  hotmelting — and tuff as well on the cutting tablet, where the job has any
  and the person holds it), per job, on
  the card itself: −, +, and one **All** (which becomes **None** at the
  total). There is nothing to expand and nothing to scroll inside a card. The
  other tablet's stage is not drawn here at all, even for somebody who holds
  both; a stage they do not hold is greyed and disabled.
- Watch a job go **gold** and drop into the collapsed "Finished · n" group at
  the bottom when **this tablet's stage** reaches the total — the cutter's
  finished jobs leave the cutter's way whether or not hotmelting has started.
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
is written next to every counter change as `CutBy`/`HotmeltBy`/`TuffBy` and
`DoneBy`, and into the log's `Who` column, captured at the moment of the tap,
not when the write eventually goes out.
