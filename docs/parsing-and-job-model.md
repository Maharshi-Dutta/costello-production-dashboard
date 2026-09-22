# Parsing and the job model

`parser.js`, `app.js: load`.

## 2. Parsing and the job model (`parser.js`, `app.js: load`)

**What.** Turns the `Production` sheet into jobs the UI can use.

**How.**
- `mapSheet` finds columns by header text (identity columns, dates, quantities
  M/N, glass unit columns AY..BF, product F/S/T groups). `blocksFromValues`
  finds the five sections by their divider rows (a job row is never a divider).
  `parseWorkbook` yields `jobs[]` with `id`, `cust`, `blk` (section index),
  `cat`, `glass{type:count}`, `prods[]`, `cp` (checkpoint colours read from the
  cells: white / yellow "process" / gold "done" / the legend's green "cut").
- `load()` in `app.js` downloads, parses, reads the dashboard sheets and the
  phases list, then lays the local holds over the result: `PENDING` (row
  moves and fills, 180 s per key), `PENDV` (view assignments), `PENDA`
  (alerts). A hold is released the moment the file agrees with it.
- `poll()` every 12 s compares `lastModifiedDateTime` and reloads on change.
- `verify.js` compares the JS parser with a Python reference extract of a
  real download (local files, not in the repo).

## See also

- [[auth-and-graph]] — previous: sign-in and the Graph layer
- [[sections-and-row-moves]] — next: sections, categories and row moves
- [[data-flow-and-lag]] — how this fits into the overall data flow
