# The fabrication station: the fourth floor page, a fourth sanctioned fill

## 25. Fabrication: eligibility per group and part, assignments, the fourth fill

Part A built 2026-09-25, Part B the same day, demo amendments 2026-09-28
(build 20260928-0959). Spec:
[`docs/specs/2026-09-25-fabrication-station.md`](specs/2026-09-25-fabrication-station.md)
(read its Amendments section — it changes eligibility and adds the
finished-on-sheet green). Data model, columns, "who may move what":
[`docs/STATIONS.md`](STATIONS.md) → "The Fabrication station data model".

### What it is

The floor step **after welding**, windows and doors, no glass. Built from the
`fabrication.html` / `fabrication.js` / `fabrication-core.js` (`FABC.FAB`)
shape the welding station started (`docs/STATIONS.md` "Adding a station").
Two lists in `Floor stations`, created by script 2026-09-25
(rehearsed on throwaway lists first, Title made unique through the API):
`Fabrication station` (one row per job **and product group**, 27 columns) and
`Fabrication assignments` (one row per piece of work given to a person, 10
columns). Twelve `Station people` rows (`Station = Fabrication`, no PIN)
added 2026-09-28.

Unlike welding, fabrication counts **all three parts** — Frames, Sashes,
**and Transoms** — wherever the sheet's own count is greater than nought, off
`Production` alone (`j.prodsMain`, `j.doors` for the door code labels; rule 5
of the brief — `Production (2)` is still not real data). ARCH ANGLES, TH W
and COMPOSITE are fed as the owner names them on the sheet (2026-09-25,
`8dd3fa4`); aluminium doors and aluclad sliders are left out — the sheet
carries no F/S column for either.

### Eligibility, per group and per part (2026-09-28 amendment)

Everybody sees every job and every line; a person can only **move** the
lines they are eligible for. `Station people`'s `Stages` column is a comma
list of entries, each a bare group name (every part of it) or
`GROUP:parts` (`frames`/`sashes`/`transoms` joined by `+`), with `ALL` for
"every fed group" so the 255-character column does not run out
(`ALL:sashes`). `fbParseStages` reads it; `fbEligible(person, group, part)`
answers everywhere the gate is asked: the tablet (`fbCanTap`, greyed lines,
"not your line" vs "not your part"), Take, the office's Assign picker, and
Approve (refused, with a message, for a part the person does not do). This
replaced a simpler "eligible for the group, full stop" rule from Part A after
the owner's demo: *"some people can do frames but not sash or transom and
other combinations."*

### The fourth sanctioned fill

The office (never the tablet) paints a product group's own F/S/T cell on
`Production`: lavender `#D9D2E9` (0 < done < total), purple `#B4A7D6` (done =
total). **Never lowers the office's own checkpoint record; the office may
lower or clear its own record at any time**, and fabrication may paint again
on the next pass (`web/CLAUDE.md` rule 1, updated in the same commit with the
owner's words). The painter is a copy — nothing reads it back as status, the
`Fabrication station` list is the truth. Neither colour existed anywhere on
the sheet before this (checked against the 2026-09-16 workbook copy).

### Part B: assignments, approval, urgent, notifications

`Fabrication assignments`: a tablet's Take creates a `Requested` row; the
office's Assign/Approve/Refuse/Remove PATCHes `Status`/`Qty`/`Decided*`, never
deletes. A split's `Qty` sum may not exceed the part's sheet total. A person
may tap a line only when eligible **and** holding an `Assigned` row on that
part; with no assignments list at all, Part A's eligibility-only gate applies.
Urgent (job/group/part) is office-only, sorts to the top of every worker's
list. Notifications — banner, short sound, card badge — fire only **while the
tablet's page is open**; nothing arrives to a closed page. Seen state is
per person, in `cw_fabseen`.

### Two review lessons carried into the build (see `docs/HISTORY.md` B32)

- The colour plan is made from a checkpoint record up to ten seconds old and a
  workbook download about 36 seconds old; the write now re-reads the live
  `/format/fill` of every planned cell (plus its two header cells) inside the
  serialised write, right before painting, and decides again from what it
  finds.
- The tablet's assignment gate has three states, not two — *read* (know the
  index), *missing* (no list, Part A applies), and *not known yet* — and the
  third one locks every line rather than defaulting to open.

### Finished on the sheet — green, display only (2026-09-28 amendment)

A job in **Ready to fit**, **Ready, customer won't take**, or **Collect &
supply only** (exactly those three) is shown finished everywhere fabrication
appears — every line full, a new green "finished on sheet" colour (`--sht*`
tokens, both themes; purple still means "done in fabrication"). Nothing is
written, nothing is tappable, a queued tap for such a job is dropped and
said so, and the colour painter skips it outright.

### localStorage keys (tablet, per device)

- `cw_fabperson` — the signed-in person (survives a reload, like the other
  tablets).
- `cw_fabq` — the queued taps not yet flushed to `FramesDone`/`SashesDone`/
  `TransomsDone`.
- `cw_fablogq` — the queued `Station log` lines.
- `cw_fabtab` — which tab is showing, `"floor"` or `"finished"` (On floor /
  Finished, the same shape as `cw_weldtab`/`cw_glztab`/`cw_glasstab`).
- `cw_fabseen` — per-person seen state for the Part B notification badges.
- `cw_fabview`, `cw_fabglassfirst` — added 2026-10-02, see [[fabrication-glass-and-door-glazing]].
- Since 2026-10-02 the tablet also reads `Glass station`, polls every 5 s and has
  no idle lock (`cw_fabperson` no longer expires): [[stations-see-each-other]].

## See also

- [[fabrication-glass-and-door-glazing]] — the My work filter, glass chip, door
  glazing and its gold on the door cells (2026-10-02)
- [[welding-station]] — the pattern this station's shape and eligibility gate
  were built from
- [[glazing-station]] — the third station, and the "Adding a station"
  checklist this one also followed
- [[status-list-is-truth]] — the checkpoint record the colour painter reads
  (and never lowers)
- [[STATIONS]] — the full column-by-column data model and "who may move what"
