# The office's top search bar drives every station board

Status: briefed 2026-09-30, owner approved the design ("looks good").

## Context

The office dashboard (`index.html` + `app.js`) has one search box in the top
bar, `#q`, bound to `state.q` (`app.js` ~10347) and used by the home list
(`renderRows`) and the John print sheet. Three of the station boards opened
from Show ▸ carry their own second search box inside the board:

| board | own box | variable | filter |
|---|---|---|---|
| Welding | `#wq` (`weldBoardHtml`, ~3307) | `WELD_Q` | `WELDC.weldFilter(all, q)` (~3230) |
| Glazing | `#zq` (`glzBoardHtml`, ~3978) | `GLZ_Q` | `GLZC.glzFilter(all, q)` (~3895) |
| Fabrication | `#fq` (`fabrBoardHtml`, ~4702) | `FABR_Q` | `FABC.fbFilter(cards, q)` (~4465) |
| Glass | none | — | — |

The owner wants one search: the top bar filters the job list **and** whichever
board is open.

## Behaviour

1. Remove the `#wq`, `#zq`, `#fq` inputs and their `oninput` wiring. Each board
   filters by `state.q` instead of its own variable. Delete `WELD_Q`, `GLZ_Q`,
   `FABR_Q` (grep every reader; nothing may still read them).
2. The Glass station board (`stationBoardHtml`, ~8468) also filters by
   `state.q`: job number or customer, case-insensitive substring, same rule as
   the other boards use. Reuse an existing filter shape; do not invent a new
   matching rule.
3. Typing in `#q` while a board is open re-renders that board (it already calls
   `renderRows()`, which picks the board renderer). Focus stays in `#q`
   (it lives in the header, not in the re-rendered host, so this should hold
   for free — verify it).
4. Switching board, or back to the list, keeps the text in `#q`; the new view
   is filtered by it. No per-board memory of a query.
5. The empty-result state on a board says "No job matches “<q>”" (reuse
   whatever each board already shows for an empty filter).
6. The section selectors on the boards (`WELD_SECT`, `GLZ_SECT`, `FABR_SECT`)
   stay as they are.
7. Tablets (`glass.html`, `welding.html`, `glazing.html`, `fabrication.html`)
   are **not** touched. They have no top bar and keep their own boxes.

## Hard rules

- No workbook write, no list write, no new column. UI only, `app.js` and, if
  needed, CSS in `index.html`.
- Do not touch `station.js`, `welding.js`, `glazing.js`, `fabrication.js`,
  `station-ui.js`.
- No real names anywhere (code, comments, tests, commit).

## Tests

- A small pure check if any board filter logic changes shape (the core
  `*Filter` functions should not need to change). If none changes, say so.
- Full verification command from the root `CLAUDE.md` must stay green.

## Report

Diff summary, the grep showing `WELD_Q|GLZ_Q|FABR_Q|#wq|#zq|#fq` gone from
`app.js`/`index.html`, the verification output (failures only).
