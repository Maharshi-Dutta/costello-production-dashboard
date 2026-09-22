# Export to Excel / PDF

`export.js`.

## 10. Export to Excel / PDF (`export.js`)

**What.** Download the current jobs (or the ticked ones) as an Excel workbook
or a PDF, with filters and a field picker.

**How.** Pure builders on `window.XP`: filters → rows → ExcelJS workbook
(Jobs, Comments, Checkpoints, Export info sheets) or a pdfmake document (job
cards, or a compact 14-column table). pdfmake is lazy-loaded from `vendor/`.
Hard rules: **no eircodes ever, and no phone numbers in this, the Default
template** (the John print sheet of [[john-print-sheet]] (§16) is the single sanctioned exception, and
the two templates share no row builder); one `Dashboard Log` line per export
(who, format, job count, filters); no network. Since 2026-09-09 the field
picker also offers **Flag** — the Production sheet's row colour code
([[john-print-sheet]], §16) —
which comes out as a last column carrying the word and the sheet's own ink.
Logo slot `assets/logo.png` (not yet supplied). Tests: `test_export.js`.

## See also

- [[alerts]] — previous: job alerts by email
- [[glass-station-overview]] — next: glass station (floor dashboards)
- [[john-print-sheet]] — the one export allowed to carry a phone number
- [[station-reports]] — the third export template, a station report
