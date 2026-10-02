# Fabrication: My work filter, glass status, door glazing

## 25b. Filter, glass chip, door glazing, gold on the door cells

Built 2026-10-01, shipped 2026-10-02 (build 20261002-0817). Briefs:
[`docs/specs/2026-10-01-fabrication-my-work-filter.md`](specs/2026-10-01-fabrication-my-work-filter.md)
and
[`docs/specs/2026-10-01-fabrication-glass-status-door-glazing.md`](specs/2026-10-01-fabrication-glass-status-door-glazing.md).
Decisions: HISTORY A43. Base note: [[fabrication-station]].

### The view filter (tablet only, display only)

A three-way switch beside On floor / Finished: **Everything** (today's view),
**My work** (only groups and parts the signed-in person is eligible for by
`fbEligible`, so `GROUP:parts` and `ALL:parts` count), **Assigned to me**
(only parts with an `Assigned` row, Qty > 0; a `Requested` row does not count;
"Assignments not loaded yet" when the list is unread). Default Everything,
remembered per person. Cards are placed by the lines the person actually
sees: a group with no shown part is not drawn, a job with nothing left is not
drawn. Counts, search (and its cross-tab line), header totals and urgent-first
all follow the filter. The filter never grants a tap; `fbCanTap` is unchanged.
Each tab has its own empty state.

### Glass status chip, Glass ready first

The office feeder writes a `Glass` word on every row of a job, from the glass
station record and the checkpoint record already held (no new workbook read):
`done`, `part:<cut>/<total>:<hotmelt>/<total>` (plus `:tuff n/m` with Tuff),
`none`, or blank for no glass. **Done requires Tuff complete when the job has
Tuff.** The chip (green done, amber partial, grey not started) is one
**combined DG+TG count, not per type**. When glass is done and nothing is
fabricated yet, the card says so. The tablet toggle **Glass ready first**
sorts glass-done unfinished jobs to the top, after urgent. The tablet never
reads the glass list. The word follows the glass list at most once a minute.

### Door glazing: a role-only fourth part

Part key `glazing`, "Door glazing", drawn under **PVC DOOR** and **PVC SMART**
only. Total = the job's DOORS DONE cells by code: `SS` goes to PVC SMART;
`CD`, `SFCD`, `BF`, `ACSD`, `ACSS` never count; any other non-empty code goes
to PVC DOOR. Columns `GlazeTotal`, `GlazeDone`, `GlazeBy`, `GlazeAt`. **Role
only, no assignment:** a person taps it only when `Stages` names it
(`PVC DOOR:glazing`, `ALL:glazing`); **a bare group name never grants it**.
Take, Assign, splits and the whole-line rule do not apply. It does not count
toward the group's fabrication done/total, colour, the lavender/purple paint
or the Finished tab. The office board has the same line with − + All None
(logged in `Dashboard Log`, never `Station log`). A job with a glazeable door
code but no PVC DOOR / PVC SMART group row gets no line.

### Gold on the DOORS DONE cells (second writer of the third fill)

When a group's door glazing is **full** (`GlazeDone >= GlazeTotal > 0`) the
office sets the checkpoint record of that job's counted `door:<slot>` items to
done **through the record**; the existing writer paints the cell gold. Full
count only, **raise only**, never lowered by fabrication, CD/SFCD/BF/AC*
cells never touched. Same safety as HISTORY B32: decided from the record read
now, inside the serialised chain; skipped while pending, for gold or
off-sheet or finished jobs, and while the record is unreadable. Review
lessons in HISTORY B36: gold-vs-tap time uses the list's server `Modified`,
never two devices' clocks, and a stale second office screen refreshes inside
the chain and PATCHes, never adds a duplicate record row. Source
`fabrication`, Who from `GlazeBy`.

### Also fixed in the same ship

The office fabrication, welding and glazing boards redraw after a capped
feed's continuation (HISTORY B35).

### localStorage keys (tablet, per device)

- `cw_fabview` — `{ "<person>": "all" | "mine" | "assigned" }`.
- `cw_fabglassfirst` — the Glass ready first toggle.

### Not built / deferred

The sheet's yellow/gold raising fabrication counts (owner, 2026-10-01: "leave
for now"). Glass tablets seeing the fabrication stage is a separate approved
design, not built.

## See also

- [[fabrication-station]] — the base station
- [[status-list-is-truth]] — the checkpoint record the gold goes through
- [[STATIONS]] — columns and the `Stages` syntax
