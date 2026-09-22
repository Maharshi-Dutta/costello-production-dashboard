# Sections, categories, row moves and mark ready

## 3. Sections, categories and row moves

**What.** Jobs sit in the sheet's five sections (In production, Ready to fit,
Collect & supply only, …). Moving a job between sections in the dashboard
moves its **row** in Excel so the floor's Excel and the office's dashboard
agree. Custom categories live only in `Dashboard Views`.

**How.** `moveJobsInSheet` → `graph.js: moveJobRow`: locate the job live
(A1:K600), capture the row (values + number formats in one read; fills/fonts
per run from the download), insert a row after the target section's last job,
write it (text starting with a digit/=/+/- gets an apostrophe so Excel keeps
it text), verify exactly two copies exist, delete the original. Borders:
Excel keeps one definition per shared edge, so the landing row writes bottom +
verticals but never its top, and the row above a landing/vacated slot gets its
bottom edge re-asserted (`restoreBottomEdge`). Row heights come from the
downloaded file (the API rounds to pixels). Verified with desktop Excel via
COM in a scratch script. Tests: `test_move.js`.

## 4. Mark ready

**What.** Marks a job ready to deliver: gold fill on its status cell and a
row move to "Ready to fit" (C/S-prefixed jobs go to "Collect & supply only").
Undo clears the fill and moves the row to the bottom of "In production".
Dates are never touched. `markReady` in `app.js`.

## See also

- [[parsing-and-job-model]] — previous: parsing and the job model
- [[checkpoints-and-phases]] — next: checkpoints and the phase pipeline
- [[module-map-and-invariants]] — the row-move and border invariants
