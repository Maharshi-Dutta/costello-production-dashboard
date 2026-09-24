# Welding tablet: Finished tab, no "Sent to floor" chip, bigger remake count

Status: approved by the owner 2026-09-24, building.

## Context

The welding tablet (`welding.html` / `welding.js` / `welding-core.js`) shows only
jobs whose `Section` is "In production" (`weldBoard`). A job moved to a later
section of the `Production` sheet (Ready to fit, Collect & supply only, Ready,
customer won't take, Can sell as second hand) drops off the tablet, although the
office's Welding station board (`weldOfficeBoard`) still shows it. Example: jobs
3708 and 5272 are visible on the office board and not on the tablet. The owner
wants the welder to reach those jobs, mainly to record a remake on a job that
came back with a fault.

## Owner's decisions (2026-09-24)

1. The "Sent to floor" chip goes. Any job on the `Production` sheet has been
   sent to the floor, date or not.
2. Two tabs at the top of the tablet board, where the chip was: **On floor** and
   **Finished**, each with its count.
   - On floor: Active rows, Section "In production", card not finished.
   - Finished: every other Active card - finished In production jobs AND every
     job in any other section still on the sheet (the office board's set).
     Jobs that left the sheet are Active = No and stay off, as today.
   - The collapsed "Finished" group at the bottom of the list goes; the tab
     replaces it.
3. Finished cards stay fully tappable (counts and remake), as finished cards are
   today. A remake does not move a job back to On floor; it stays finished and
   shows its remake count.
4. The remake counter stays on every Frames and Sashes line, works on any job
   (finished or not, any section), and is restyled: bigger and easier to read
   from a distance. Same data, same queue, same log, same PATCH as shipped
   2026-09-23 (count only).
5. Welding tablet only. Glass and glazing tablets are not touched.

## Hard rules (unchanged)

- The tablet never touches the workbook (no `setFill|clearFill|setValues|
  appendLog|saveProgress|moveJobRow|batchWrite|/workbook`).
- Portrait 10-inch tablet, one column, never scroll sideways (700-1100 px).
- No real names in code, tests or docs.

## Behaviour details

- The selected tab is remembered on the device (localStorage, try/catch), key
  replacing `cw_weldsent`. Default On floor.
- Search searches within the selected tab. Header "Frames N left / Sashes N
  left" stays the In production board number exactly as today (unaffected by
  the tab and the extra sections).
- Card order within a tab: the existing order (seq, then job).
- A card for a job not In production shows its section name in small text on
  the card head so the welder knows why it is there (e.g. "Ready to fit").
- Empty-tab messages: On floor empty / Finished empty, plain words.
- Office side unchanged: remake already shows as `↻ N` on the office board and
  drawer, logged in `Station log`.

## Tests to deliver

- `test_welding.js`: new core function(s) splitting cards into the two tabs
  (finished In production -> Finished; unfinished In production -> On floor;
  any other section, finished or not -> Finished; Active = No -> neither);
  `weldSentFilter` removed or unused and its tests adjusted; header left count
  still the In production board only.
- Browser rig: tabs switch, 3708-style job in "Ready to fit" appears under
  Finished, remake + tap works on it, no sideways scroll at 700/900/1100 px.
