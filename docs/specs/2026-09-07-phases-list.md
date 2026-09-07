# Hand-set phases stored in a SharePoint list (spec, approved "option A" 2026-09-07)

## Context
Dashboard: `d:\Costello Windows\Excel Dashboard\web` (static, plain ES2017, no modules).
The drawer's phase pipeline (`jobPhase(j)` in checkpoints.js: 0 In office … 6 Fitted /
delivered, derived from the sheet) becomes clickable. The chosen phase must be shared by
everyone WITHOUT touching the production workbook in any way (no tab, colour, column).
It is stored in a **SharePoint list** in the ProductionProgress site.

## The list
Name `Dashboard phases` (display name; internal list name may differ - look it up by
display name once and cache the id in localStorage `cw_listids`). Columns: `Title` = job
number (upper-case), `Phase` (number 0-6), `PhaseName` (text), `SetBy` (text), `SetAt`
(text, ISO). One item per job; the dashboard upserts (find by Title, PATCH fields, else
POST). Read all items on every `load()` (`GET …/lists/{id}/items?expand=fields&$top=999`,
follow `@odata.nextLink`). The list is created by the manager, not by the code; if the
list is missing the feature shows a plain message and does nothing.

Graph: site id is already known (`findFile()` → `fileRef.siteId`). Endpoints:
`GET /sites/{siteId}/lists?$select=id,displayName`, `GET /sites/{siteId}/lists/{listId}/items?expand=fields(select=Title,Phase,PhaseName,SetBy,SetAt)&$top=999`,
`POST /sites/{siteId}/lists/{listId}/items` body `{fields:{Title,Phase,PhaseName,SetBy,SetAt}}`,
`PATCH /sites/{siteId}/lists/{listId}/items/{itemId}/fields` body `{Phase,…}`,
`DELETE …/items/{itemId}` for clearing. All through the existing `call()`; add
`Sites.ReadWrite.All` to `SCOPES` in graph.js (users consent once at sign-in).
Functions in graph.js (exported on `window.CW`): `listId(displayName)`,
`listItems(displayName)`, `listUpsert(displayName, title, fields)`, `listDelete(displayName, title)`.
Serialise writes per list with the existing `serialised()` queue.

## Behaviour
- `PHASES_SET = { [job]: {phase, name, who, at} }` read in `load()` (after the workbook,
  failures tolerated: keep the last known map, toast once).
- Effective phase `effectivePhase(j)` = max(jobPhase(j), hand-set phase). The sheet's own
  evidence wins when it passes the hand-set one; clearing the hand-set phase returns the
  job to `jobPhase(j)`. `phaseName`, the row badge and the pipeline use the effective phase.
- Drawer pipeline: in Edit mode each step is a 40 px+ button. Click a step → set that
  phase (optimistic hold like PENDING for 180 s, `pend(id,{phase:n})`), write the list item,
  `noteChange(id, "Phase", oldName, newName)`. Click the current hand-set step again →
  clear (delete the item, log "Phase: X → sheet"). A hand-set phase shows a small line
  "set by <who>, <time>" and, if it differs from the sheet's own reading, "sheet says: <name>".
  Steps below the sheet's own phase cannot be chosen (disabled, tooltip "the sheet already
  shows this job past this step").
- Not in Edit mode: read-only as today. Non-admins may set phases (it is not destructive).
- On write failure: toast `friendly(e)`, drop the hold.
- No workbook write of any kind. No new localStorage beyond the hold and `cw_listids`.

## Tests (`test_phases_list.js`, vm-loaded real code, fake fetch)
list id lookup + caching; items read with paging; upsert = PATCH when present, POST when
absent; delete on clear; effectivePhase rules (sheet wins when higher; clear returns to
sheet); hold survives a stale read and releases when the list agrees; log lines; disabled
steps; no Production/worksheet request ever made (assert on the request log); scope added.
Keep every existing suite green.

## Deliverables
Summary per file, the exact Graph calls for a set and a clear, pasted output of node
--check on every JS file and all suites. No commits.
