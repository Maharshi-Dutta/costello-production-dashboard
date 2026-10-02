# Glazing: glass and fabrication status, ready to glaze (2026-10-01)

Status: built on branch `glazing-ready-chips`, awaiting the owner's go.

## Owner's words

- "Glazing should be able to see what jobs' glasses and fabrication are done so
  he knows which one to glaze next."
- Tuff must be complete for glass to count as done.
- Glass and fabrication only for now (no welding chip).
- Show it on both the tablet and the office Glazing board.
- Nothing moves by itself. A **Ready first** switch re-orders the list, and it
  is off by default.
- Follow the design the fabrication session built
  (`2026-10-01-fabrication-glass-status-door-glazing.md`, section A).

## What was built

- **List:** two text columns on `Glazing station` in `Floor stations`, `Glass`
  and `Fabrication`. They were added by script on 2026-10-01: rehearsed on a
  throwaway list, then created and verified. No row was touched. Both are
  feeder facts: `GLZ_FEEDER_FIELDS` carries them, and `floorOnly` drops them
  from any tablet write.
- **Glass word:** `""` (no glass), `none`, `part:<cut>/<t>:<hotmelt>/<t>[:tuff n/m]`
  or `done`. It is worked out by `FABC.fbGlassStatus`, the same function the
  fabrication tablet's chip uses, so the two tablets cannot disagree.
- **Fabrication word:** `""` (nothing to fabricate), `none`, `part:<done>/<total>`
  or `done`, from the job's fabrication card (`GLZC.glzFabStatus`). A job shown
  finished on the fabrication board is `done`.
- **Office feed (`app.js` `glzStatusOf`):** reads only lists the dashboard
  already holds. There is no new read and no workbook write. A list that
  cannot be read right now gives `null`, and `GLZC.glzKeepStatus` then keeps
  the word already on the row, so an unread list never feeds a ready job back
  to blank. The feed comes back a minute later (`glzFeedSoon`, debounced 60 s).
  The same debounce re-feeds when the glass or fabrication list moves.
- **Ready:** neither glass nor fabrication is still owed, and at least one of
  them is done (`glzReady`). Two blanks mean a row not fed yet, not a ready job.
- **Tablet:** read-only chips under the card head (green done, amber part,
  grey not started) and a READY TO GLAZE badge. A switch above the list reads
  "N ready to glaze · Ready first: on/off". It is off by default, remembered
  on the device (`cw_glzreadyfirst`), and applies to the On floor tab only.
  The tablet reads neither the glass list nor the fabrication list.
- **Office Glazing board:** the same chips on each row.

## Limits told to the owner

- The words update only while an office dashboard is open somewhere.
- They follow a glass or fabrication tap by up to about a minute.

## Checks

- `test_glazing.js`: 66, including three new blocks.
- Every other suite is unchanged.
- Browser rig `ready_check.js` (session scratchpad): 16/16 at 700, 800 and
  1100 px.
