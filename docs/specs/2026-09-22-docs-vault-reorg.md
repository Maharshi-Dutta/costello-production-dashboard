# Docs reorg: REFERENCE.md and ARCHITECTURE.md into a linked vault

Owner-requested, 2026-09-22. Pure documentation restructuring — no app code
changes, no workbook/list access, nothing to rehearse. Delegated as a
Sonnet chore per `CLAUDE.md` rule 6.

## Context

`web/docs/REFERENCE.md` (2203 lines) and `web/docs/ARCHITECTURE.md` (364
lines) have grown into single long files. The owner wants them broken into
smaller, topic-scoped Markdown notes (each under 150 lines) inside
`web/docs/`, cross-linked with Obsidian-style `[[wikilink]]` syntax, so a
session only has to load the note it actually needs instead of the whole
file.

`web/docs/HISTORY.md` is explicitly **out of scope** — it stays exactly as
it is, one file, untouched. It is "the running record" and is read
symptom-first (see its own section B); splitting it is a separate decision
the owner has not made.

## Hard rules

- **No information loss.** Every sentence currently in REFERENCE.md or
  ARCHITECTURE.md must land in exactly one new file. Do not summarize,
  compress, or drop detail — this is a file split, not a rewrite. Copy text
  verbatim into its new home; only the surrounding headers/links are new.
- **Do not touch HISTORY.md, STATIONS.md, SUPPORT.md, or anything under
  `docs/specs/`.** Read-only for this task.
- **Do not touch any code file** (`.js`, `.html`).
- **No real names, emails, or company domains** in any new file (rule 4 —
  same standard as the rest of the public repo). If either source file
  already violates this, flag it in your report instead of copying the
  violation forward.
- **14 other files in the repo reference `REFERENCE.md` or `ARCHITECTURE.md`
  by name** (found by grep: `CLAUDE.md`, `README.md`, `docs/STATIONS.md`,
  `docs/SUPPORT.md`, and several files under `docs/specs/`). To avoid
  editing all of them, **keep `REFERENCE.md` and `ARCHITECTURE.md` as their
  own files, but shrink each to a short index page** (well under 150 lines):
  a one-paragraph orientation plus a linked table of contents into the new
  topic files, in the same reading order as today. Anything that already
  links to `docs/REFERENCE.md#some-heading` should still land somewhere
  sensible — note any such anchor links you find so we can check them.
- **Every new file gets a one-line YAML-free Markdown title (`# Title`) and,
  at the bottom, a short "See also" list of `[[wikilink]]` links** to the
  files it's most connected to (the previous/next topic in reading order,
  plus anything it references by name, e.g. a welding note linking to
  `[[glass-two-stage]]` where the two stations share a pattern).

## Suggested split (adjust if the content doesn't cut cleanly here — you're
reading the real file, this is a starting map, not a spec to force-fit)

`REFERENCE.md` sections (numbered `##` headings, current line numbers):
- 0, 12, 13, 14, 15 (shape of the whole thing; doc set; tests/build/deploy;
  scratch tooling; lessons that cost time) → `overview-and-process.md`
- 1 (sign-in and Graph layer) → `auth-and-graph.md`
- 2 (parsing and the job model) → `parsing-and-job-model.md`
- 3, 4 (sections/categories/row moves; mark ready) → `sections-and-row-moves.md`
- 5, 6 (checkpoints; phase pipeline) → `checkpoints-and-phases.md`
- 7, 8 (selection wheel; versions/rollback) → `ui-extras.md`
- 9 (job alerts by email) → `alerts.md`
- 10 (export to Excel/PDF) → `export.md`
- 16 (row colour code, John print sheet, print notes) → `john-print-sheet.md`
- 11 (glass station intro) → `glass-station-overview.md`
- 17 (glass colours reach Production) → `glass-colours.md`
- 18 (office clear reaches the floor) → `glass-office-clear.md`
- 19 (status list is truth) → `status-list-is-truth.md`
- 20, 20a (station comments + floor notes amendment) → `station-comments.md`
- 21 (welding station) → `welding-station.md`
- 22 (glass two-stage redesign, no glazing gate) → `glass-two-stage.md`
- 23 (day sheets, weekly target, station reports) → `day-sheets-and-reports.md`
- 24 (glazing station, phase bar hears the floor) → `glazing-station.md`

Some of these (17, 19, 21, 22, 23, 24 especially) are themselves 150–250+
lines — split further along their own `###` subsection boundaries (e.g.
`glazing-station.md` + `glazing-station-phase-bar.md`) rather than
compressing content. Use your judgement; the 150-line cap is the hard
constraint, the file list above is not.

`ARCHITECTURE.md` sections:
- Data flow in text, the 36-second file lag / PENDING holds →
  `data-flow-and-lag.md`
- Dashboard-owned sheets; SharePoint lists and the scope split →
  `sheets-and-lists-scope.md`
- The station feeder and the station page; the welding station (what a
  second station costs) → `station-feeder.md`
- Module map; key invariants; where state lives → `module-map-and-invariants.md`

## What to also do

1. Delete `web/graphify-out/` (directory) and `web/.graphifyignore` — stale
   leftovers from the graphify tool, removed 2026-09-15, already marked
   "safe to delete" in `CLAUDE.md`.
2. Update the root `CLAUDE.md`'s "Where things are" table (and the "Read it,
   then `web/CLAUDE.md` ... and `web/docs/REFERENCE.md`" line near the top)
   so it describes the new layout: `REFERENCE.md` and `ARCHITECTURE.md` are
   now index pages into linked notes in `web/docs/`; `HISTORY.md` is
   unchanged and still the first read for a bug report. Do not remove the
   graphify history paragraph (section "Seeing the docs as a graph") — it's
   a historical record, not a pointer to a live tool; you may add one
   sentence noting the leftover files were deleted on this date.
3. Check `web/CLAUDE.md` for the same kind of direct reference and update if
   needed.

## What to report back

- The full list of new files created, each with its line count.
- Confirmation every new file is under 150 lines.
- A short map of what moved where (old heading → new file), so the reviewer
  can spot-check without re-deriving it.
- Any anchor links (`REFERENCE.md#...`) found elsewhere in the repo that now
  point somewhere different, and what you did about them.
- Any real name/email/domain you found in the source material and skipped.
- Confirmation `HISTORY.md`, `STATIONS.md`, `SUPPORT.md`, and `docs/specs/*`
  were not modified, and no `.js`/`.html` file was touched.
- `git status --short` and `git diff --stat` output from `web/`.

Do not commit. Do not push. Report back for review.
