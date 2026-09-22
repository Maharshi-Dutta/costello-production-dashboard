# Sign-in and the Graph layer

`graph.js` — everything that talks to Microsoft.

## 1. Sign-in and the Graph layer (`graph.js`)

**What.** Everything that talks to Microsoft: sign-in, workbook download,
surgical workbook writes, dashboard sheets, SharePoint lists, the floor
stations site.

**How.**
- `initAuth`/`signIn(scopes)`/`signOut`; `token(scopes, quiet)` gets a token
  silently and only opens a popup when `quiet` is false.
- `SCOPES = Files.ReadWrite.All + User.Read` for sign-in and every workbook
  call. `LIST_SCOPES` adds `Sites.ReadWrite.All` and is used **only** for
  `/lists` paths and a by-path site lookup other than the workbook's site
  (`needsListScope(path)`, decided on the *shape* of the path, never on
  characters inside a range address). Background list reads are quiet;
  `listConsent()` is the one place a consent popup may open, from a click.
- `call(method, path, body)` wraps `fetch` with retries for 429/5xx and the
  Excel `InvalidSession` 400 (re-opens the workbook session and retries once).
  The `workbook-session-id` header is sent only on `/workbook` paths.
- `findFile()` resolves site → drive → the workbook by name and caches it in
  `cw_fileref`. `downloadWorkbook()` fetches `/content` into ExcelJS. The
  downloaded file lags the Excel API by about 36 s, which is why the UI holds
  its own writes locally (§2 — [[parsing-and-job-model]]).
- Workbook writes: `setFill/clearFill/setValues` (single ranges), `batchWrite`
  (`$batch` with a concurrency cap; flooding it earns `OperationQueueFull`),
  `moveJobRow` (§3 — [[sections-and-row-moves]]), `appendLog`, `saveAssignment`, `saveProgress`,
  `addAlert/removeAlert`, each behind `serialised(sheet, fn)` so one tab's
  writes to a sheet run in order.
- Lists: `listId/listItems/listItemsFor/listUpsert/listDelete` (phases, with
  dedupe on Title), `listAdd/listPatch` (plain writes), `listDelta`
  (§11 — [[glass-station-overview]], delta queries with `@odata.deltaLink`, 410 resync detection),
  `stationSite()` (resolves and caches the `Floor stations` site id;
  `forgetStationSite()` clears it).

**Lessons.** Adding a scope to the sign-in list breaks sign-in for everyone
until admin consent exists. A colon test for "is this a list path" matched
every A1 range and killed every workbook write. Both are covered by tests.

## See also

- [[overview-and-process]] — previous: the shape of the whole thing
- [[parsing-and-job-model]] — next: parsing and the job model
- [[station-feeder]] — the list layer this section describes, in use
