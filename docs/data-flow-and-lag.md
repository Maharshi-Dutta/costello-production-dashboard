# Data flow, and the 36-second file lag

## Data flow, in text

```
  SharePoint site (ProductionProgress)
        |
        |  1. sign in (MSAL, Entra ID)                     <- graph.js
        |  2. findFile(): locate the workbook by name
        |  3. downloadWorkbook(): GET the whole .xlsx       <- graph.js
        v
  ExcelJS workbook in memory (full fidelity: values, fills, formats)
        |
        |  parseWorkbook(wb)                                <- parser.js
        v
  job model: one plain object per job (id, cust, dates, prods,
  glass, cp, blk, seq, cat, ...)
        |
        |  applyPending() merges in this browser's own
        |  optimistic writes (see "the 36 s lag" below)      <- app.js
        v
  ALL (in-memory array) --> rendered as the job list, the drawer,
                             the phase pipeline, the export window
        |
        |  a tap / a form submit
        v
  a write, always through the Excel API in a workbook session:
    - setFill() on one cell (checkpoints)                    <- graph.js
    - moveJobRow() (the only structural write: insert, copy,
      verify, delete)                                        <- graph.js
    - upserts against Dashboard Log / Views / Progress /
      Alerts (own sheets, created on first use)               <- graph.js
    - upserts against the Dashboard phases / Glass station
      SharePoint lists (never the workbook at all)            <- graph.js
```

The workbook is fetched whole because that is the only way to read cell fill
colours cheaply for every job at once; every write goes back through the
targeted Excel API (`PATCH .../format/fill`, `PATCH .../range(...)`,
`/insert`, `/delete`) so the file itself is never rewritten wholesale, except
`restoreVersion` — the one deliberate whole-file operation, used only to roll
back to an earlier SharePoint version.

### The 36-second file lag and `PENDING` holds

SharePoint takes roughly 35 seconds to fold a change made through the Excel
API into the file you'd get from `downloadWorkbook()`; the Excel API itself
reflects the same change in about a second. If the dashboard only trusted
what it downloads, a person's own edit would appear to work and then vanish
the next time the page refreshed and re-downloaded a still-stale file.

The fix is `PENDING` (`app.js`): every optimistic write is held in memory
and in `localStorage` (`cw_pending`) for up to 180 seconds (`PENDING_MS`) — for a section move, mark-ready, a product status and a hand-set phase only, since 2026-09-11 —
each field tracked with its own timestamp. `applyPending(list, fresh)`
overlays these holds onto whatever was just parsed; when a *freshly*
downloaded file agrees with a held value, that hold is dropped early so
everyone else's later edits aren't shadowed forever. The same pattern
repeats for section moves (`PENDV` / `cw_pendv`), alerts (`PENDA` /
`cw_penda`), and hand-set phases (`p.phase` inside `PENDING`).

Since 2026-09-10 a glass colour written from the floor's counters is held the
same way, under `p.gc` inside `PENDING`: the colour's own word (`gold` /
`yellow` / `""`) per glass type, with its own timestamp. Two things need it —
the cell must not flicker back to the 36-second-old file, and a *reversal*
(gold walking back to yellow) is the case where a flicker would be most
visible. Where a job's own drawer tick (`p.cp["glass:dg"]`) and a floor colour
(`p.gc.dg`) are both held, the newer timestamp is drawn, which is the same
last-writer-wins rule the writer itself applies. See [[glass-colours]] (§17 of the old REFERENCE.md).

## See also

- [[overview-and-process]] — the shape of the whole system this diagram belongs to
- [[sheets-and-lists-scope]] — next: dashboard-owned sheets and the scope split
- [[glass-colours]] — the `gc` hold this page describes in detail
