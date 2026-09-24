# The welding station: the second floor page, and the pattern for the next

## 21. The welding station: the second floor page, and the pattern for the next

**Shipped 2026-09-16.** Spec:
[`docs/specs/2026-09-16-welding-station.md`](specs/2026-09-16-welding-station.md).
Data model, admin setup and the checklist for the third station:
[`docs/STATIONS.md`](STATIONS.md).

The PVC welders get their own tablet page (`welding.html`), the office gets a
board at **Show ▸ Welding station**, progress lives in one SharePoint list in
the `Floor stations` site, and **nothing in the feature writes the workbook**.

### What it is

- **`Welding station`**, one row per job **and product group**, so a job with
  casement windows and a PVC door is two rows carrying the same job facts.
  `Title` = `JOB|GROUP`, unique.
- The groups fed are **every** F/S/T product group the parser finds **on the
  `Production` sheet and on no other** (`j.prodsMain`), minus a deny-list of
  four (`ALU CLAD WINDOWS`, `ALUCLAD TILT & TURN`, `BIFOLD`, `COMPOSITE`). A
  new green group on the sheet needs no code change; a new red one is one line
  in `welding-core.js`.
- **The slice reads `Production` only, and `prodsMain` is what makes that
  true** (2026-09-17, after a live bug — `docs/HISTORY.md` B20). `j.prods` is
  the parser's cross-sheet **maximum** of each group's F/S/T, which is right
  for the things it was written for and wrong the moment a sheet's headers and
  its data come apart. A column inserted into `Production` left
  `Production (2)` — formulas that moved with it, headers that did not — with
  every number one column right of its own header, and the max gave one job
  four product groups it has not got and a sashes count from the column next
  door. `j.prodsMain` is the `Production` sheet's own numbers, never merged;
  a job with no `prodsMain` is not on `Production` and is not fed at all.
- **A group is fed only the components it actually has.** `WELD_GROUP_PARTS`
  in `welding-core.js` overrides the default `["frames","sashes"]` per group:
  `SUPER DOOR` is sashes only, so the frames number on the sheet is fed as
  **0** rather than carried, and neither board draws a Frames line for it
  (both already skip a line with nothing to weld). The Title and the columns
  are unchanged — nothing is created or removed, the number is simply nought.
- Within a group, **F (frames) and S (sashes) only**. T is green on the owner's
  header template in most groups and is still never shown — "there is no
  transomes even if it is green" (owner, 2026-09-16).
- **Three colours, three levels, the same three words:** no colour = nothing
  welded, **yellow** = started, **green** = every one welded. On a
  Frames/Sashes line, on the product group, and on the job card.

### Rule 3 and the COMMENT (`weldStripDigits`)

The owner asked for the sheet's COMMENT on the floor, with phone numbers
removed (option B, 2026-09-16), and that answer is recorded as a rule-3
decision. The strip runs **on the way into the list**, so no phone number and
no eircode is ever stored where the floor could read one:

- every run of **six or more digits** becomes `…`;
- anything **eircode-shaped** (a letter, two digits or a digit and `W`, then
  four alphanumerics, with or without a space) becomes `…`;
- whitespace is collapsed and the result is capped at 140 characters.

A job number is deliberately not caught: `R5303` is a letter and four digits,
two characters short of the eircode shape. `export.js` had no eircode regex to
borrow — it protects the eircode by never reading `j.eir` at all, which is the
stronger rule where it applies and no help at all for free text — so the shape
is written down once, in `welding-core.js`, and tested in `test_welding.js`
against both a phone number and an eircode.

A dash- or dot-joined run of sizes is over-stripped the same right-side-of-wrong
way as an ISO date — `cill 150-2100` → `cill …`, `1200-900-1500` → `…`, `pane
1.2.3.4.5.6` → `…` — because it has the same digit-and-separator shape as a
phone number; that is the same owner's decision, the same way round (rule 3
wins over a readable size). An `x`-separated size (`2 x 1200 x 900`) has no
such shape and survives.

### The seed

Only when the feeder is **creating** a row, or when `DoneAt` is still empty
(checked twice, exactly as the glass list's seed is): `FramesDone = Frames`
when the office's own record for `prod:<group>:f` says `done`, else `0`; the
same for sashes with `:s`. **Yellow seeds nothing** (owner, answer 5) — yellow
says the work is not finished, and seeding it would tell the floor that frames
still to weld are welded.

The record is `cpStatus(j, item)` — the `Dashboard progress` list, [[status-list-is-truth]] (§19) — and
**not** the Excel colour, so a cell painted gold in Excel by hand reaches the
seed through the existing safeguard rather than by anything here reading a
fill. Nothing in this feature can paint one either: rule 1 still has exactly
three sanctioned fills and welding is not a fourth.

### The tablet

`welding.html` + `welding.js`, the same shell as the glass page and the same
shared station account: the picker, the PIN, the ten-minute lock, Switch
person, the search box, the note channel (`ST.stationComments({ station:
"Welding" })` — one word changed), ten-second delta polling, queued absolute
writes rebased against the list, incremental redrawing.

One card per job, its groups inside it in the sheet's own order, Frames and
Sashes lines with `−` / `+` / `All` (which becomes `None` at the total) and a
bar. A line with nothing to weld is not drawn; a group with nothing to weld is
never fed. Groups on the board are **Active** and **Finished**; a finished card
is still tappable and comes back the moment it is reduced, by the floor or by
the office. The header carries **two capsules, "Frames N left" and "Sashes N
left"** (each part's own left over the whole board, never narrowed by the
search box, added by `weldLeftByPart`; a group fed only one part, e.g. Super
door's sashes, contributes 0 to the other without a special case) and two
**tabs, On floor and Finished**, each with its count (2026-09-24, replacing the
Sent to floor chip and the collapsed Finished group). On floor is In production
and not finished; Finished is every other Active card of the office's board -
finished In production jobs and every job in any other section, whose card head
names its section (`weldTabs`). The header capsules stay the In production
board's. The tab is remembered on the device (`cw_weldtab`).

There is **no lock column**. Glass has `OfficeDone`, which greys the tablet;
welding does not need one, because the office edits the same counters instead.

## See also

- [[station-comments-amendment]] — previous: amendment E, the job-row icon and board notes
- [[welding-station-office-and-site]] — next: the office board, generalisation and the site pin
- [[status-list-is-truth]] — the record welding's seed reads
- [[glazing-station]] — the third station, built from the checklist this one started
