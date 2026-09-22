# Selection wheel and versions

## 7. Selection wheel (radial menu)

**What.** Ticking jobs shows a floating button; it opens a full-circle radial
menu (Alert to, Export, Move to, Untick all). `renderFab`/`fabToggle` in
`app.js`; motion in `--fab-*` CSS variables in `index.html`, timeline in
`docs/motion-notes.md`. Nothing here writes anything.

## 8. Versions and rollback

**What.** The Versions window lists the workbook's SharePoint version history
and can restore one (`renderVersions`, `graph.js: listVersions /
downloadVersion / restoreVersion`). This is the rollback for any dashboard
write to the workbook.

## See also

- [[checkpoints-and-phases]] — previous: checkpoints and the phase pipeline
- [[alerts]] — next: job alerts by email
