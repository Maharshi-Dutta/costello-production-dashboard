# Reference: what has been built, and how

This used to be one 2200-line file. As of 2026-09-22 it is an index into
linked topic notes under `web/docs/`, cross-referenced with Obsidian-style
`[[wikilink]]` syntax, so a session only has to open the note it actually
needs. Every sentence that used to live here moved verbatim into one of the
notes below — nothing was summarised or dropped. Read `docs/HISTORY.md`
first for a bug report (see its own section B); read this page first for "how
does feature X work". Names of people and addresses are placeholders
throughout ("the admin", "the colleague").

Last updated 2026-09-22 (split into linked notes; content unchanged).

## Reading order

Each entry below keeps the section number(s) it had in the old single file,
so an old `§N` reference in another doc still resolves to the right note.

0, 12–15. **Shape of the whole thing; docs; tests/build/deploy; scratch
tooling; lessons** — [[overview-and-process]]

1. **Sign-in and the Graph layer** — [[auth-and-graph]]

2. **Parsing and the job model** — [[parsing-and-job-model]]

3, 4. **Sections, categories, row moves; mark ready** — [[sections-and-row-moves]]

5, 6. **Checkpoints; the phase pipeline** — [[checkpoints-and-phases]]

7, 8. **Selection wheel; versions and rollback** — [[ui-extras]]

9. **Job alerts by email** — [[alerts]]

10. **Export to Excel / PDF** — [[export]]

11. **Glass station (floor dashboards)** — [[glass-station-overview]], and the
    ordinary job list's glass chip — [[glass-station-job-row]]

16. **Row colour code and the John print sheet** — [[john-print-sheet]]

17. **Glass colours reach Production** — [[glass-colours]], continued in
    [[glass-colours-tuff-and-lock]] (the hold, ordering, TUFF, the lock)

18. **An office clear reaches the floor's counters** — [[glass-office-clear]],
    continued in [[glass-office-clear-history]] (the bugs found on the way)

19. **Status lives in a list; the Excel colour is a copy** —
    [[status-list-is-truth]], continued in
    [[status-list-is-truth-safeguard]] (switch-on window, import, safeguard)

20, 20a. **Station comments** — [[station-comments]], continued in
    [[station-comments-details]] (reading window, review bugs) and
    [[station-comments-amendment]] (amendment E: job-row icon, board notes)

21. **The welding station** — [[welding-station]], continued in
    [[welding-station-office-and-site]] (office board, generalisation, site pin)

22. **Glass: one page per stage, no glazing** — [[glass-two-stage]],
    continued in [[glass-two-stage-office-board]] (office board, review, tests)

23. **Day sheets, weekly target, station reports** — [[day-sheets-and-reports]],
    continued in [[day-sheets-office]] (the office's window and writes) and
    [[station-reports]] (the export template)

24. **The glazing station** — [[glazing-station]], continued in
    [[glazing-station-phase-bar]] (the phase bar hears the floor, tests, review)

## See also

- [[ARCHITECTURE]] — the companion index: data flow, module map, invariants
- [[overview-and-process]] — start here for the shape of the whole system
