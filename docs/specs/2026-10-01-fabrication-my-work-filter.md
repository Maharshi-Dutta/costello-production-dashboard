# Fabrication tablet: "My work" filter — brief (2026-10-01)

Status: shipped 2026-10-02, build 20261002-0817.

Read first: `docs/specs/2026-09-25-fabrication-station.md` (incl. its
Amendments), `docs/fabrication-station.md`, `docs/STATIONS.md` (Fabrication
section, `Stages` syntax).

## What the owner asked

"An option or filter that allows the floor dashboard to see only the job he
can do and component he can do." Sign-in stays by name (the existing people
picker). Plus a third option: only what is assigned to me.

## Behaviour (tablet only, `fabrication.html`)

A three-way switch beside On floor / Finished:

**Everything | My work | Assigned to me**

- **Everything** — exactly today's view (every group and part shown; ones the
  person may not do are locked).
- **My work** — only jobs with at least one group **and part** the signed-in
  person is eligible for (`fbEligible(person, group, part)`, i.e. the
  `Stages` syntax incl. `GROUP:parts` and `ALL:parts`). Inside each card only
  those groups and parts are drawn; a group with no eligible part is not
  drawn at all; a job with nothing left is not drawn.
- **Assigned to me** — only parts where the signed-in person holds an
  `Assigned` row in `Fabrication assignments` (Qty > 0); same hiding rule. A
  pending `Requested` row does not count. If the assignments list is missing
  or not yet read, say so in the empty state ("Assignments not loaded yet")
  rather than showing nothing silently.
- Default after sign-in: **Everything**. The choice is remembered **per
  person** in localStorage key `cw_fabview` (`{ "<person>": "all" | "mine" |
  "assigned" }`), wrapped in try/catch.
- The On floor / Finished counts, the search (incl. its cross-tab "N more
  in …" line), the header totals ("N left") and the urgent-first sort all
  follow the filter. Urgent items hidden by the filter are hidden (they are not
  the person's work).
- Finished-on-sheet green jobs follow the same rule (shown under My work only
  if they have an eligible part).
- An empty result shows a plain line: "Nothing for you here — switch to
  Everything to see all jobs."
- Notifications are not filtered (an assignment is always the person's).

## Hard rules

- **Display only.** No list write, no new column, no workbook anything; the
  tap gate (`fbCanTap`) is unchanged — the filter never grants a tap.
- Tablet files keep the rule-2 grep at 0.
- 10-inch portrait, no sideways scroll at 700–1100 px, both themes; the header
  stays two rows at 700 px (`#upd` hiding etc. as glass did if needed).
- No real names in the repo.
- Edit/Write only, no commit/push, no live services.

## Tests

`test_fabrication.js`: a pure core function (e.g. `fbViewFilter(cards,
person, mode, assignIdx)`) — mine hides ineligible groups/parts/jobs with
`GROUP:parts` and `ALL:sashes`; assigned shows only Assigned (not Requested,
not Removed); everything is identity; tabs/search counts follow; per-person
memory. Browser rig: switch present, filters, no sideways scroll at 700 px in
both themes.
