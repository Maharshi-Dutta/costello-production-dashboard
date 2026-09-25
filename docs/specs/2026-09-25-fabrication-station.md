# Fabrication station — brief (2026-09-25)

Status: **approved by the owner 2026-09-25, Part A in build.** Part B follows
after Part A is reviewed.

Read first: `CLAUDE.md` (repo root of this folder and `web/CLAUDE.md`),
`docs/STATIONS.md` → "Adding a station" (follow it step by step), the Welding
station data model in the same file, and `docs/HISTORY.md` B20, B29, B31.

## 1. Context

Fabrication is the floor step **after welding**, for windows and doors (no
glass). Seven people share a few tablets with PIN sign-in, like the glass
tablets. Each person can only do certain product types. The supervisor (in
code and docs: **"the supervisor"**, never a name) assigns the work, splits it
between people, approves self-picked work and flags urgent work, all from a new
office tab. The owner and the supervisor both use that office tab.

Everyone can **see** every job and every product group. A person can only
**change** the lines of a product group they are eligible for (Part A), and in
Part B only the lines assigned to them.

## 2. Hard rules (non-negotiable)

1. **The tablet never touches the workbook.** No `setFill`, `clearFill`,
   `setValues`, `appendLog`, `saveProgress`, `moveJobRow`, `batchWrite` or
   `/workbook` in `fabrication.js`, `fabrication-core.js`, `fabrication.html`.
2. **The office never writes `Station people` or `Station log`.** The
   feeder never writes the floor's columns and never deletes a row.
3. The COMMENT goes through `ST.stripContact` (no phone numbers, no
   eircodes) before it reaches a list.
4. No real person's name anywhere in the repo. "the supervisor", "Person A".
5. **A quantity comes off `Production` alone** (`j.prodsMain`, `j.doors`),
   never `j.prods`, `j.wnd`, `j.drs`. `Production (2)` and other sheets are not
   real data.
6. **New sanctioned fill (owner's decision, 2026-09-25) — the fourth.** The
   office dashboard (never the tablet) paints a product group's own F / S / T
   cell on `Production` in two new colours:
   - `FAB_PROCESS_HEX = "#D9D2E9"` (light lavender) — fabrication started on
     that line (0 < done < total);
   - `FAB_DONE_HEX = "#B4A7D6"` (purple) — that line fully fabricated.
   Neither colour occurs anywhere on the `Production` sheet today (checked on
   the 2026-09-16 workbook copy: fills in use are white, gold FFE699/theme 7,
   yellow FFFF00, red FF0000, orange FFC000, greens 92D050/00B050/theme 9, grey
   theme 2/6, one light blue theme 8). They are deliberately **not yellow and
   not gold**, so nothing that reads the sheet (welding's gold seed, the
   checkpoint record, the office's hand-paint adopter) mistakes fabrication
   for the office's own status.
   **Never lower the office.** The painter only paints a cell when the
   office's checkpoint record for that item (`prod:<group>:<f|s|t>`, read with
   the existing `cpStatus`) is blank. If the office's record is yellow, gold or
   cut, fabrication leaves the cell alone. The office may lower or clear its
   own record at any time; when it does, fabrication may paint again on the
   next pass. Fabrication clears (paints white) **only a cell currently showing
   one of its own two colours** and only when its own done count for that line
   is back to 0. Update rule 1 in `web/CLAUDE.md` in the same commit, dated,
   with the owner's words: "never lowers office but office can lower theirs".
7. The fab colour is a **copy**. Nothing reads it back as status. The list
   `Fabrication station` is the truth.
8. Classic `<script>` files share one global scope: no top-level name in
   `fabrication*.js` may repeat one in `graph.js`, `station-core.js`,
   `station-ui.js` or `checkpoints.js` (`test_pages.js` must stay green; B29).
9. Subagents edit code with Edit/Write only, never a shell heredoc. Do not
   commit, do not push, do not touch live lists or the live workbook.

## 3. Product groups

Windows: CASEMENT WINDOWS, POLARIS 85MM CASEMENT, 4000 CASEMENT, 7000
CASEMENT, POLARIS 85 TILT & TURN, 4000 TILT & TURN, 7000 TILT & TURN, PVC
FRENCH WDS, ARCH ANGLES, TH W (the owner's "Arch & AnglesTH W"), ALU CLAD WINDOWS, ALUCLAD TILT & TURN.
Doors: SIDELIGHTS, SUPER DOOR, BIFOLD, PVC SMART, COMPOSITE (the owner's "Composite PVC door"), PVC DOOR. Sheet headers confirmed by the owner 2026-09-25.

Implement as an **allow-list** (`FAB_GROUPS`) compared through the same key
normaliser welding uses (`wKey`-style: upper case, whitespace collapsed).
**Before writing it, list the actual F/S/T group headers the parser returns
for the local workbook copy** (`../2026 Production work in progress Jan 2026
oakley.xlsx`) and match the owner's names to them; report every header that
matches none of the owner's names and every owner name that matches no header.
Do not guess a match — report it.

**All three parts** — Frames, Sashes, Transoms — where the sheet's count is
> 0. (Welding has no T; fabrication has.)

**Doors: the counts are already right on the sheet.** The owner: "if a job has
2 CD and 1 DD it will have PVC F 3, COMPOSITE S 2 and PVC S 2". So the feeder
takes F/S/T straight from `prodsMain`; the DOORS DONE codes (`j.doors`, e.g.
CD, DD, SS, SFCD, BF, PVC) are **labels only**, fed as a text column on each
door-group row of the job (e.g. `2 CD, 1 DD`) and shown on the card. Nothing
is computed from them. For reference only (the owner's rules, not code):
CD = 1 frame from PVC DOOR + 1 composite sash; PVC = 1 F + 1 S from PVC DOOR;
DD = 1 F + 2 S from PVC DOOR; SS = 1 F + 1 S (PVC SMART); SFCD = Super door
frame + 1 composite sash; BF = 1 frame + up to 6 sashes from BIFOLD.

## 4. Data model (all in `Floor stations`, definition `site: "floor"`)

### `Fabrication station` — one row per job **and product group** (Part A)

| column | type | written by | meaning |
|---|---|---|---|
| `Title` | text, unique | feeder | `JOB\|GROUPKEY` |
| `Job`, `Group`, `GroupSeq`, `Customer`, `Comment`, `Seq` | as welding | feeder | as welding |
| `Doors` | text | feeder | door codes of the job as labels, door groups only, else blank |
| `Frames` / `Sashes` / `Transoms` | Number | feeder | the group's F / S / T on `Production` |
| `Section` | text | feeder | the job's section |
| `Active` | text | feeder | `Yes` while In production with this group > 0 |
| `OnSheet` | text | feeder | `Yes` while the job is on the sheet at all (Finished tab reads it, as glass) |
| `FedAt` / `FedBy` | text | feeder | as welding |
| `FramesDone` / `SashesDone` / `TransomsDone` | Number | tablet; office board | fabricated so far |
| `FramesBy/At`, `SashesBy/At`, `TransomsBy/At` | text | tablet; office | last mover of that counter |
| `DoneBy` / `DoneAt` | text | tablet; office | last touch of any counter |
| `Urgent` | text | **office only** (Part B) | blank, or a comma list of `group`, `frames`, `sashes`, `transoms` |

No seed (`seedFields: []`): fabrication has no office record to seed from.

### `Fabrication assignments` (Part B) — one row per piece of work given to a person

| column | type | meaning |
|---|---|---|
| `Title` | text, unique | `JOB\|GROUPKEY\|PART\|<random 6>` |
| `Job`, `Group`, `Part` | text | `Part` = `frames` / `sashes` / `transoms` |
| `Person` | text | `Station people` Title |
| `Qty` | Number | how many of that part are this person's |
| `Status` | text | `Requested` (tablet "Take"), `Assigned` (office), `Refused`, `Removed` |
| `RequestedBy/At`, `DecidedBy/At` | text | who asked, who decided |

A split is several `Assigned` rows on one part; their `Qty` sum may not exceed
the part's total (the office refuses the save and says so). Removing an
assignment sets `Status = Removed`; rows are never deleted.

### Shared lists

- `Station people`: `Station = Fabrication`, `Stages` = comma list of product
  group names the person may do (matched through the key normaliser). Seven
  placeholder rows, Person A–G, created by the manager's script, not by code.
- `Station log`: `Station = Fabrication`, `GlassType` = group, `Stage` =
  `frames` / `sashes` / `transoms`. Tablet only.
- `Station comments`: `Station = Fabrication`, as the other stations.

No list is created by code. A missing list shows the quiet explained state and
writes nothing.

## 5. Behaviour

### Part A — tablet, eligibility, office board, colours

- `fabrication.html` / `fabrication.js` / `fabrication-core.js` (`FABC.FAB`
  definition). Copy the welding page's shape; share through `station-core.js`
  and `station-ui.js`; do not copy anything that can live there.
- Tablet: one card per job, every group, each group's F / S / T lines with
  − / + / All / None like welding. **Lines of a group the signed-in person is
  not eligible for are shown read-only** (greyed steppers, a small "not your
  line" hint on tap), never hidden. Door codes shown on door-group heads.
- Two tabs **On floor / Finished** + cross-tab search exactly as welding
  (`weldTabs` shape, tab key `cw_fabtab`); Finished = every other on-sheet card.
- Colours on the tablet and board: none / lavender (started) / purple (done)
  at line, group and card — theme tokens, both themes. Not welding's yellow and
  green.
- Office: `STATIONS` gets `["fabrication", "Fabrication station"]`; a
  `feedFabrication()` in its own link of `load()`'s chain and its own `try` in
  `stationPoll()`; a board like welding's (one row per job, every section, open
  a row for − + All None on each line; an office edit writes that counter, its
  By/At, DoneBy/DoneAt, and one `Dashboard Log` line; never `Station log`);
  a read-only Fabrication line in the job drawer; `redrawFabrication()` wired
  wherever the other stations' redraws are (STATIONS.md "Adding a station",
  G1).
- **Colour painter** in `app.js`, modelled on `glassColourPlan` / the glass
  colour writer: plan per job → cells to paint per rule 6 → one `$batch`,
  one `Dashboard Log` line per paint ("Fabrication colours"). Only rows the
  floor or office has touched (`DoneAt` set) produce paint. **Verify and
  report** that (a) the checkpoint writer never whitens a cell whose record is
  blank on its own, and (b) the hand-paint adopter does not treat a lavender /
  purple cell as a hand change. If either is false, stop and report — do not
  work around it silently.

### Part B — assignments, approval, urgent, notifications

- Office board, opened job: per part, the assigned pieces (person, qty),
  **Assign** (picker offers only people eligible for that group; qty defaults
  to what is unassigned), **Remove**, and at the job, group and part level an
  **Urgent** toggle. A **Requests** strip at the top of the board: Approve /
  Refuse. A **Who is doing what** view (Show ▸ Fabrication station, second
  toggle): one block per person, their open assignments with done/qty, urgent
  first.
- Tablet: a person can tap a line only when eligible **and** holding an
  `Assigned` row on that part. On an unassigned part an eligible person sees
  **Take** (creates a `Requested` row); the line stays locked until approved
  (owner: no tapping before approval). Their card shows "yours: 12".
- Urgent: an urgent job/group/part sorts to the top of every worker's list,
  with a red icon beside exactly what is urgent. Only the office sets or
  clears it.
- Notifications on the tablet (page open): a new or changed assignment for the
  signed-in person, or an urgent flag on something they hold or are viewing,
  shows a banner + a short sound + a badge on the card until tapped. Seen state
  per person in localStorage (`cw_fabseen`). The delay is the delta poll
  (~10 s). Say in the docs plainly: no notification while the page is closed.
- Every office write (assign, remove, approve, refuse, urgent) logs one
  `Dashboard Log` line.

## 6. Tests to deliver

- `test_fabrication.js` in the offline pattern: allow-list and key matching;
  slice from `prodsMain` only (a `prods`-only group is not fed); door labels;
  F/S/T with T; tabs; eligibility gate; clamp; the painter plan (blank record →
  paints; yellow/gold/cut record → never; own colour + done 0 → white; foreign
  colour + done 0 → nothing); Part B: split sum refused over total, locked
  until approved, urgent sort, notification seen-state.
- Add it and the three new files to the verification command in both
  `CLAUDE.md` files; `test_pages.js` must include `fabrication.html`.
- A browser rig in the scratchpad (reuse `stub_all_shared.js` pattern from
  scratchpad `48e0c3b2…`): tablet at 700 and 1100 px, no sideways scroll,
  both themes.

## 7. Report back

Summary of files changed, the group-header matching table (section 3), the two
"verify and report" answers (section 5), pasted tail of every suite, and
anything in this brief you could not do as written.

## 8. Lists to create (manager, by script, before the push)

`Fabrication station`, `Fabrication assignments` in `Floor stations`, and the
seven `Station people` rows — by a script with modes check / rehearse /
create, rehearsed on a throwaway list first, reusing `make_floor_site.py`'s
sign-in (scratchpad `82777b7e…`).
