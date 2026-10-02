# Glass tablets see the fabrication stage — brief (2026-10-01)

Status: design approved by the owner 2026-10-01, not built.

Read first: `docs/fabrication-station.md`, `docs/glass-two-stage.md`,
`docs/station-feeder.md`, `docs/STATIONS.md` ("Adding a station", step 7 — the
site pin), `docs/specs/2026-09-24-finished-tab-glass-glazing-search.md`.

## What the owner asked

The cutting and hotmelting tablets (`glass.html?stage=cut` / `?stage=hotmelt`)
should show, on each job, what stage fabrication is at. A job whose
fabrication is done while its glass is not must stand out, so the glass people
know to do that job first.

Owner's decisions (2026-10-01):

1. The tablet reads the `Fabrication station` list **directly**, read only.
   Not copied onto `Glass station` by the office.
2. A waiting job gets a **badge and a filter**. It is **not** sorted to the
   top; card order stays as it is.
3. **One line per job**, summed over the job's product groups. Not one line
   per group.

## Hard rules

- **Read only.** No list write of any kind for this feature, no new column,
  no new list, nothing in the workbook. The rule-2 grep on the station files
  stays at 0.
- The glass lists stay pinned to the workbook's own site (`ST.GLASS.site =
  "own"`). This feature adds **one** separate, read-only lookup of the
  `Floor stations` site for the `Fabrication station` list. It must not share
  a site id, a delta token or a "lists moved" counter with the glass lists,
  and it must never change where the glass lists resolve.
- **Fail quiet.** If the site or the list cannot be read (no membership, no
  consent, 403, 404, offline), the fabrication line reads "Fabrication: not
  available" in the muted colour, and everything else on the glass page works
  exactly as today. No retry storm: re-check on the same slow cadence the
  other "list missing" states use.
- Reads `Production`-derived data only, through the list. Nothing from
  `Production (2)` or other sheets.
- 10-inch portrait, no sideways scroll at 700–1100 px, both themes, header
  stays two rows at 700 px.
- No real names, addresses or company domains in the repo.
- Edit/Write tools only for code. No commit, no push, no live services.

## Data

`Fabrication station`: one row per job and product group, with a total and a
done count for each of Frames, Sashes and Transoms, and the row's `Section`.
Reuse `fabrication-core.js` for reading rows into records and for the
finished-on-sheet rule (`fbRecord`, `fbRollUp` or whichever existing function
already answers it) rather than writing a second reader. `glass.html` then
loads `fabrication-core.js`; `test_pages.js` must still pass (no top-level
name collision in the page's script chain).

One new pure function in `station-core.js`, e.g.
`ST.fabOfJob(fabRows)` → for one job:

- `state`: `"none"` (no fabrication rows for the job), `"notstarted"` (every
  done count 0), `"progress"`, or `"done"` (every line full, or the job is in
  one of the three finished sections — the same rule the fabrication screens
  use for their green).
- `parts`: `{frames:{done,total}, sashes:{…}, transoms:{…}}`, summed over the
  job's rows, only parts whose total is above 0.

And `ST.glassWaiting(job, fab, stages)` → true when `fab.state === "done"`
and the signed-in person's own stage or stages are not complete for that job.
"Not complete" is the same test the card's "N left" uses for those stages; a
job with Tuff owed (`ST.tuffOwed`) counts as not complete for a person who
holds the Tuff counter.

Polling: the same delta cadence the tablet already uses for its other lists.

## UI (tablet only)

On each glass card, one line under the counters:

- `none` — no line.
- `notstarted` — "Fabrication: not started", muted.
- `progress` — "Fabrication: in progress · Frames 4/4 · Sashes 1/4 ·
  Transoms 0/2", lavender (the fabrication tokens).
- `done` — "Fabrication: done · Frames 4/4 · …", purple; or the green
  finished-on-sheet tokens when the done comes from the section.
- Waiting — a red badge on the card head, "FABRICATION DONE — GLASS
  WAITING", plus the `done` line.

Header: one capsule, "N waiting on glass", shown only when N > 0. Tapping it
toggles a filter that shows only waiting cards (in the current tab); tapping
again clears it. The filter is not remembered across a reload, and is cleared
where the search is cleared (`pickPerson` / `switchPerson` / idle lock). It
works together with search and the On floor / Finished tabs.

The fabrication line must be part of the card signature so a fabrication tap
redraws only the cards it changes (the incremental draw stays incremental).

The office page is untouched.

## Tests to deliver

`test_station.js` (or a small new suite added to the verification line):

- `fabOfJob`: none / notstarted / progress / done; sums across two groups;
  parts with total 0 left out; finished section gives done.
- `glassWaiting`: true only when fab done and the person's stage is not
  complete; false for a cutter whose cutting is complete while hotmelting is
  not; Tuff owed case.
- The list unreadable: the page function returns the "not available" state
  and no glass behaviour changes.
- No write: assert no list write call is made on the fabrication list.

Browser rig (scratchpad, stubbed Graph): line in each state, badge, capsule
count, filter on and off, filter with search, no sideways scroll at 700 px in
both themes, fabrication list returning 403.

## What to report

Files changed, the functions reused from `fabrication-core.js`, how the
second site lookup is kept apart from the glass pin, suite counts, anything
in this brief that the code made impossible or wrong.
