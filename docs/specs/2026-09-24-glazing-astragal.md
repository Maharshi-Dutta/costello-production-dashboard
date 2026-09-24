# Glazing: astragal as its own count (2026-09-24)

Status: built on branch `glazing-astragal`, awaiting the owner's go.

## Owner's words

- "make astragal as a separate component that shows the number of astragal
  units there are and still to be done in a job, no relation to windows or
  doors"
- "if a job has no windows but has doors and astragal, show it on the glazing
  tablet and dashboard"
- A job is done only when windows AND astragal are both full (yes).
- The count line changes when − / + / All are pressed, the same as the windows line.
- The tablet header shows both: windows left and astragal left. Read from the
  `Production` sheet.
- Keep it simple: the count and the done count only. No `AstragalBy` or `AstragalAt`.

## What was built

- **List:** two Number columns on `Glazing station` in `Floor stations`:
  `AstragalTotal` (fed by the office feeder) and `Astragal` (the done count,
  written by the tablet and by the office board). The owner gave permission to
  add them from this session. They were added by script on 2026-09-24:
  rehearsed on a throwaway list, then created and verified. No row was touched.
- **Parser:** `j.astrMain`, the ASTRAGAL quantity from `Production` alone
  (owner's rule of 2026-09-18). `j.glass.astragal` is the cross-sheet maximum,
  so it is not read here.
- **Feeder:** a job is fed when it has windows OR astragal. A doors-only job
  is still not fed. `AstragalTotal` is one more feeder column. On the first
  office load after the ship, every existing row gets `AstragalTotal` once.
  Rows that went `Active = No` on 2026-09-23 and have astragal come back to `Yes`.
- **Tablet:** each card shows a `Windows` line and/or an `Astragal` line, each
  with its own count, bar, − + and All/None. The header reads
  "N windows left · M astragal left". An astragal tap PATCHes `Astragal`,
  `DoneBy` and `DoneAt`, and logs one `Station log` line with
  `Stage = astragal` and `GlassType = ASTRAGAL`. The station report's glazing
  days count `Stage = glaze` only, so astragal units are never added to
  window units there. The queue keeps the windows' entry under the bare item
  id, as before, and the astragal entry under `<id>|a`.
- **Office board:** one head line per counter, on the same grid, so every
  column stays aligned. The second line says "Astragal". The drawer line reads
  "3 / 11 windows · 0 / 4 astragal". The station report gains `Astragal` and
  `Astragal done` columns.
- **Done / colour:** green on the tablet, gold on the office board, only when
  every line is full. Yellow as soon as either line has started. A gold
  `Production` row still wins, as before. The phase bar hears "In glazing"
  when windows or astragal have been recorded.
- **Excel:** untouched. The ASTRAGAL cell stays hand-ticked.

## Checks

- `test_glazing.js`: 60, including four new astragal checks.
- The rest of the verification line is unchanged.
- Browser rig `astr_check.js` (session scratchpad): 12/12.
- The older rigs still pass: glazing_check 26/26, glz_polish_check 11/12
  (same fixture gap as before).
