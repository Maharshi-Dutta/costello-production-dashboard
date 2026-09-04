# Checkpoints: per-job progress ticking (spec, approved by the user 2026-09-04)

## Product context
Costello Windows runs its factory from one SharePoint workbook (`Production` sheet).
This dashboard (`d:\Costello Windows\Excel Dashboard\web`, a static site: `index.html`,
`app.js`, `parser.js`, `graph.js`, vendored MSAL + ExcelJS, no bundler, plain ES2017,
no modules) reads the workbook by downloading it and writes surgically through the
Microsoft Graph Excel API. Users: 3–4 office/factory people; Abin ticks work off from
his phone.

The user (Maharshi) approved this feature with these answers:
1. Windows, doors and glass cells: **gold when finished, yellow while in process**.
2. **Also** do counts for each product's F / S / T (Frame, Sashes, Transom).
3. When everything is ticked do **not** auto mark-ready (QA may still be pending); leave the existing button.
4. Lowering a count / un-ticking is allowed (and logged).
5. Exact counts live in the dashboard's own sheet; **Excel gets only the colour**.

## Hard rules (non-negotiable)
- The ONLY cells ever written on the `Production` sheet by this feature are fills
  (colours) of the job's own row in: column M (WND), column N (DRS), the glass type
  columns (AY..BF, from `mapSheet().glass`), and product F/S/T columns (from
  `mapSheet().prod[product][sub]`). Never values, never other rows, never other sheets
  except the dashboard's own `Dashboard Progress` and `Dashboard Log`.
- Row numbers are looked up by job number immediately before every write
  (`CW.rowForJob("Production", id)`); never cached.
- Colours: white `#FFFFFF` for 0 done (the sheet's cells already carry an explicit
  white fill - do NOT use `fill/clear`), yellow `#FFFF00` for partial, gold `#FFE699`
  for complete. Constants `GOLD_HEX`, `YELLOW_HEX` exist in `app.js`; add `WHITE_HEX`.
- Excel's colour wins over the stored count on every refresh (see merge rule).
- Never call Graph / SharePoint / the live workbook from tests. Tests run offline with
  the fake-fetch pattern in `test_move.js`. Do NOT run `rehearse.js`. Do NOT commit.
- Keep `node --check` green for app.js / parser.js / graph.js, keep `node test_move.js`
  and `node verify.js` passing. Do not modify `moveJobRow`, `writeRow`, `captureRow`,
  `restoreBottomEdge` or the border logic in `graph.js`.
- No new dependencies. Follow the existing code style: small plain functions, comments
  that say *why*, theme-aware CSS using the variables already in `index.html`
  (`--surface`, `--ink`, `--line`, `--done`, `--fab`, `--single`, `--accent`, ...).

## Data model
An **item** is one countable thing on a job:
| item key | label | total | Excel cell (column) |
|---|---|---|---|
| `win` | Windows | `j.wnd` | `PRODMAP.qty.wnd` (M) |
| `drs` | Doors | `j.drs` | `PRODMAP.qty.drs` (N) |
| `glass:<TYPE>` | Glass TG / Glass TUFF ... | `j.glass[TYPE]` | `PRODMAP.glass[TYPE]` |
| `prod:<name>:<f|s|t>` | `<Product> frames/sashes/transoms` | `p.f / p.s / p.t` | `PRODMAP.prod[name][sub]` |

Only items with total > 0 exist for a job. `PRODMAP = mapSheet(productionSheet)` is
already kept in `app.js` after each download; use `CW.A1(col)` for column letters.

**Stored progress** lives in a new sheet `Dashboard Progress` (create on first write,
same pattern as `ensureViewsSheet` / `saveAssignment` in `graph.js`):
`Job | Item | Done | Total | Who | When` (all text, numberFormat `"@"`), one row per
(Job, Item), upserted. Add to `graph.js`: `PROGRESS_SHEET` constant,
`ensureProgressSheet()`, `saveProgress(job, item, done, total, who)`, and
`saveProgressMany(job, [{item, done, total}], who)` (one usedRange read, then all
row writes in one `batchWrite` - `batchWrite`/`batchRun` already exist). Export via
`window.CW`. Hard rule as with the log: the sheet name is a constant, never a parameter.

**Reading**: in `app.js` `load()`, after the Views sheet, read `Dashboard Progress`
from the downloaded workbook into `PROGRESS = { [job]: { [item]: {done, total, who, when} } }`.

**Excel status** (colour) per item comes from the parser. `parser.js` already reads
fills for product F/S/T (`j.status[pname].process/done`). Extend it so each job gets
`cp: { win: st, drs: st, glass: { TYPE: st }, prod: { name: { f: st, s: st, t: st } } }`
with `st` in `""` (no colour / white), `"process"` (yellow), `"done"` (gold). A fully gold
row (`rowDone`) counts as `"done"` for every item. Keep `prods[].st` working as today.
Export nothing new from parser beyond the job fields.

**Merge rule** (`app.js`, pure function `itemState(j, item)` → `{done, total, status}`):
- total T from the job; if T <= 0 the item does not exist.
- status S from Excel colour (`j.cp`), stored D from `PROGRESS[j.id][item].done`.
- S === "done" → done = T, status "done".
- S === "process" → status "process"; done = D if 0 < D < T else `null` ("in progress,
  count unknown" - render bar at half, label "in progress").
- otherwise → done = 0, status "".
- A held (pending) value overrides everything for 180 s (below).

**Instant updates**: extend `PENDING` (`pend(id, {cp: {item: done}})`, persisted in
localStorage like today) and `applyPending` so a held count sets `j.cp` status
(0 → "", < T → "process", >= T → "done") and a `j.cpDone[item]` count that `itemState`
prefers over `PROGRESS`. Expiry and persistence exactly like `done`/`prods`.

## Writes (all in the background; the screen updates first)
`setItemProgress(j, item, newDone)` in `app.js`:
1. clamp 0..T; if unchanged return; `pend()` it; `ALL = applyPending(ALL)`; re-render
   rows + drawer immediately.
2. **Debounce per (job, item) 800 ms** so a run of + taps becomes one write: after the
   pause, do in order: `CW.setFill("Production", cell, colour)` (row via `rowForJob`),
   `CW.saveProgress(...)`, `noteChange(j.id, label, oldDone + " of " + T, newDone + " of " + T)`.
   Colour: 0 → white, partial → yellow, complete → gold.
3. On failure: toast `friendly(e)`, put the held value back to the previous count.
4. 45 s later `load("checking…", true)` (existing pattern).

`setGroupDone(j, group)` for the "All ... done" buttons (`win`, `drs`, `glass`, or one
product): sets every item in the group to its total: one `batchWrite` of fill PATCHes
(`format/fill` per cell), one `saveProgressMany`, one log line per group
("Glass: all done (TG 25, TUFF 11, ...)"). Reverse action "Clear" sets them to 0 (white).

## UI (drawer, `renderDrawer` in `app.js`; styles in `index.html`)
A **Checkpoints** section directly under the Progress steps, visible always; the
controls are enabled only in Edit mode (consistent with every other write in the app -
the Edit button is at the top of the drawer). Groups in this order:
- **Windows** - one line: label, `6 of 10`, progress bar, `−` `+` buttons, a number input,
  and an **All done** button (becomes **Clear** when complete).
- **Doors** - same.
- **Glass** - one line per type the job has (TG, TUFF, NOT TUFF, ...) as above, plus one
  **All glass done** button for the group.
- **Products** - the existing components table becomes count-based: for each product,
  F / S / T each show `done of total` with a mini bar and `−`/`+`; a per-product
  **All done** button. Remove the old In-fabrication / Process-done `<select>`
  (keep `setProductStatus` only if still used elsewhere; otherwise delete it).
Bars: 0 = empty, partial = yellow (`--fab`), complete = gold (`--done`); "in progress,
count unknown" = half bar in yellow with the words "in progress". Tap targets ≥ 40 px
(phones). A per-job summary line at the top of the section: "Windows 6/10 · Doors 0/2 ·
Glass 0/52 · Frames/Sashes/Transoms 0/35". A tiny "who/when" under a line when the
stored row has it ("Abin, 10:42").
The list row (`rowHtml`) gets a small "cp" badge when any item is partially done
(e.g. "3 in progress") - keep it subtle.

Do NOT auto mark-ready. The existing mark-ready button stays as it is.

## Logging text
- Single item: what = label (e.g. `Windows`, `Glass TG`, `7000 CASEMENT frames`),
  from = `"2 of 10"`, to = `"6 of 10"`.
- Group: what = `Windows: all done` / `Glass: all done` / `<Product>: all done` /
  `... cleared`, from = "", to = the list of items.

## Tests to deliver (`test_checkpoints.js`, run with `node test_checkpoints.js`)
Use the vm shim pattern from `test_move.js` (load parser.js + graph.js; for app.js
logic either load app.js with a minimal DOM stub or factor the pure parts -
`itemState`, colour choice, pend/applyPending for `cp`, debounce coalescing - so they
can be exercised without a DOM). Cover:
1. merge rule: gold cell + stored 3/10 → 10/10 done; yellow + stored 6/10 → 6; yellow +
   no stored → null/"in progress"; white + stored 6 → 0; whole row gold → all done.
2. pending override wins for 180 s and expires.
3. debounce: 5 rapid `+` taps → exactly one fill write, one progress write, one log line
   "1 of 10 → 6 of 10"... (i.e. old = value before the burst).
4. `saveProgress` upsert: existing (Job, Item) row is updated in place; new pair appends;
   `saveProgressMany` uses one batch.
5. colour choice: 0 → white, 1..T-1 → yellow, T → gold; clamping.
6. parser: `cp` statuses from fills for M, N, glass, F/S/T; rowDone → all done
   (build a small ExcelJS workbook in memory for this, `exceljs` is installed for Node).

## Deliverables from the implementer
- Code changes in `app.js`, `parser.js`, `graph.js`, `index.html` (+ `test_checkpoints.js`).
- `node --check` on the three JS files; `node test_move.js`, `node verify.js`,
  `node test_checkpoints.js` all passing (paste the output).
- A short written summary: what was added where, the exact Graph calls a single tap
  makes, and anything you were unsure about. No commits.
