# Glass progress on the master dashboard's job row

**Date:** 2026-09-09
**Status:** in progress
**Consulted:** yes — the owner reported on 2026-09-09 that tapping + / − / All
on the tablet "is not updating the status in the job card in the master
dashboard", and approved this design (dashboard only, no workbook writes) the
same day. The owner separately agreed that mirroring the floor's progress into
the `Production` sheet's glass checkpoints is a **different** feature, to be
specced on its own later — it is out of scope here.

---

## 1. Context

The glass station (see `2026-09-08-glass-station-v3.md`) is live. The floor
taps counters on `glass.html`; those taps are PATCHed into the `Glass station`
SharePoint list, and one line per tap is POSTed to `Station log`. The office
dashboard already reads both lists and already shows the floor's numbers in
three places:

- **Show ▸ Glass station** — the read-only board (`stationBoardHtml`).
- **A glass job's drawer** — the "Glass station" section
  (`stationSectionHtml`).
- **The Floor log window** — `paintStationLog`.

What it does **not** do is show anything about the floor on the ordinary job
list. That is the whole of this spec.

### Why the owner saw nothing (the investigation, so you don't repeat it)

Three separate reasons, all of which must be closed:

1. **`rowHtml(j, i, max)` in `app.js`** builds the job row entirely from the
   workbook model: job number, customer, county, status word, windows/doors,
   the F/S/T mini bar, and a run of badges (sheet chips, the colour-code flag
   chip, Urgent, "In fab", "n in progress", the comments count). It never
   reads `STATION_ITEMS`. There is no element on the card for a floor tap to
   change.
2. **`redrawStation()` in `app.js`** repaints the rows only when the station
   board is open — `if (state.board) renderRows();`. On the ordinary job list
   a poll that brought in a floor change would not repaint anything.
3. **`stationWatching()` in `app.js`** returns true only for the board, the
   log window, or a drawer open on a glass job. On the plain job list the poll
   therefore runs at `STATION_SLOW_MS` (60 s), not `STATION_FAST_MS` (10 s).

The data itself is fine — that is why the Floor log window was up to date.

---

## 2. Hard rules (non-negotiable — the review will grep for these)

1. **No workbook write of any kind is added by this feature.** No
   `setFill`, `clearFill`, `setValues`, `moveJobRow`, `batchWrite`, no request
   to any `/workbook` URL on any path this spec introduces. This feature is
   read-and-render only.
2. **No write to any SharePoint list.** The office already never writes
   `Station people` or `Station log`, and never writes the floor's own columns
   of `Glass station`. Nothing here changes that. The chip is built from data
   already in memory (`STATION_ITEMS`); it must not add a read of its own, a
   fetch, or a new list call.
3. **No phone number, no eircode** anywhere near this code.
4. **No real person's name, email address or company domain** in code,
   comments, tests or this spec. Placeholders only ("the admin", "Person A",
   `example.test`).
5. The existing test suites stay green — in particular `test_station.js`'s
   standing proofs that no request touched the workbook, that no DELETE was
   sent, and that every floor PATCH is a subset of the floor's own columns.

---

## 3. Data you already have

Use what exists; add no new state and no new read.

- `STATION_ITEMS` — the `Glass station` list as last read, kept current by
  `stationPoll()` / `feedStation()`. May be `null` before the first read.
- `ST.jobRecord(STATION_ITEMS, id)` — one job's record whatever its `Active`
  is, or `null` when the floor has never been fed that job. Already wrapped in
  `app.js` as **`stationForJob(id)`** — use that.
- The record's shape (from `station-core.js`): `{ id, job, customer, seq,
  fedAt, active, total, cut, hotmelt, glazed, by:{}, at:{}, bars:{},
  finished }`. `STAGE_KEYS` is `["cut","hotmelt","glazed"]`; `STAGE_ROW` maps a
  stage key to the field on the record. `finished` is
  `total > 0 && every stage counter >= total`.
- `ST.glassTotal(j)` — the floor's own reckoning of whether a job has glass
  (DG + TG). A job with no glass the floor works on is not on their board.
- `ST.glassWords(n)` — `"12 glasses"` / `"1 glass"`, worded the same on both
  screens.

---

## 4. What to build

### 4.1 The chip on the job row

In `rowHtml`, in the same badge cell as "In fab" and "n in progress", add one
badge for a job the floor has a record of. Rules:

- **Nothing at all** when `STATION_ITEMS` is null (not read yet), when
  `stationForJob(j.id)` is null (never fed to the floor), or when the record's
  `total` is 0. A job the floor has never seen must look exactly as it does
  today — no empty chip, no "0/0", no placeholder.
- **Otherwise** the chip reads `Glass <done>/<total>`, where `<done>` is the
  **sum of the three stage counters** and `<total>` is `3 × record.total` —
  i.e. progress across all three stages, not one of them. Put the plain-words
  explanation in the `title` attribute so hovering says what it means, e.g.
  `"The floor has recorded 8 of 24 stage steps on 8 glasses: cutting 4,
  hotmelting 3, glazing 1. Recorded on the Glass station page — the Excel file
  is not involved."` Build that sentence from the record; do not hard-code
  numbers.
- **Gold when `record.finished`** — the same "finished" rule and the same gold
  the board and the tablet already use, so the three screens agree. Reuse the
  existing gold custom property the board uses rather than a new hex literal.
- **A job that has left the floor** (`active` false) still shows its chip: what
  the floor recorded is still true. Do not hide it.
- The chip must survive a narrow screen: the badge cell already has
  `overflow:hidden`, so keep the text short — the long form goes in `title`.
- Escape everything through `esc()`, as every other badge does.

### 4.2 Repaint the job list when the floor moves

In `redrawStation()`, repaint the ordinary job list as well as the board.
Today it is `if (state.board) renderRows();` — it must call `renderRows()`
whether or not the board is open.

Two things to be careful of, both already documented in that function:

- Do **not** repaint the filter bar. The existing comment says why: a poll
  landing while somebody is halfway through typing a job number must not take
  the box out from under them. `renderRows()` is the right call; do not widen
  it to a full re-render.
- Repainting the rows must not disturb the tick boxes. Selection lives in
  `state.picked` and `rowHtml` reads it, so a normal `renderRows()` is safe —
  but add a test that proves a picked job is still picked after a poll
  repaints the list.

### 4.3 Refresh at the fast rate while glass jobs are listed

In `stationWatching()`, add a fourth reason to be watching: the ordinary job
list is on screen and the rows currently rendered include at least one job the
floor has a record of.

- Base it on what is actually being shown (the filtered/visible rows), not on
  `ALL` — a dashboard filtered down to jobs with no glass has no reason to
  poll six times a minute.
- Keep it cheap: it runs on every tick. If working out the visible rows is
  expensive, cache the answer and recompute it in `renderRows()`.
- When the board, the log window or a glass job's drawer is open, the existing
  reasons still apply and win.

---

## 5. Tests to deliver

Add to **`test_station.js`** (it already loads the real `app.js`,
`station-core.js` and a stub DOM — follow the patterns already there). Every
one of these must be a real assertion on the shipped code, not a
reimplementation:

1. **No chip when the floor has never seen the job** — `stationForJob` returns
   null → `rowHtml` output contains no "Glass" badge, and the row is otherwise
   byte-identical to the row built with `STATION_ITEMS = null`.
2. **No chip when `total` is 0.**
3. **The chip's numbers** — a record with total 8 and counters 4/3/1 renders
   `Glass 8/24`; the `title` names all three stage numbers.
4. **Gold when finished** — counters at 8/8/8 of total 8 → the chip carries
   the finished styling; at 8/8/7 it does not.
5. **A job that has left the floor** (`active` false) still renders its chip.
6. **`redrawStation()` repaints the job list when the board is closed** — with
   `state.board` null, a poll that moves a counter results in a re-rendered
   row carrying the new number.
7. **A picked job stays picked** across that repaint.
8. **`stationWatching()`** is true when the visible rows include a job the
   floor has a record of, false when they do not, and still true for the
   board / log window / glass drawer.
9. **The standing safety proofs still hold** over the whole new run: no
   request to `/workbook`, no DELETE, no write to `Station people` or
   `Station log`, no floor column written by the office beyond the seed the v3
   spec already allows. Do not weaken these; if a new stub is needed, extend
   the fixture rather than relaxing an assertion.

Run and paste the output of the full suite listed in `CLAUDE.md`
("Running the checks"), not just `test_station.js`.

---

## 6. What to report back

- The diff, function by function, with the reasoning for anything you chose
  that this spec did not pin down.
- Pasted output of every check in `CLAUDE.md` § "Running the checks".
- Anything in this spec you found to be wrong, impossible, or in conflict with
  the code — say so plainly rather than working around it silently.
- Confirmation, in your own words, that no workbook write and no list write
  was added, and how you satisfied yourself of that.

## 7. Out of scope — do not build

- Any write to the `Production` sheet, including colouring the glass
  checkpoints from the floor's counters. That is a separate spec, deferred by
  the owner on 2026-09-09.
- Changing the tablet (`glass.html`, `station.js`), the board, the drawer
  section or the log window. They already work.
- Changing the feeder, the seeding rule, or anything in `station-core.js`.
  This feature is a renderer and a poll rate, nothing more.
- Any change to `index.html` beyond what `build.py` stamps. If you believe a
  style rule is needed for the chip, add it to the existing badge styles and
  say so in your report.

## Amendments after review

### 1. §3 was wrong, and it caused the review's critical finding (manager, 2026-09-09)

§3 said "it must not add a read of its own". Followed literally — as it was —
that left the feature with no way to get its data at all:

- `feedStation()` returns **before reading the list** when the glass slice hash
  is unchanged and the last feed was under ten minutes ago, and that state is
  persisted in `localStorage` so it survives a reload. On that path
  `STATION_ITEMS` stays `null` and `STATION_OK` stays `null`.
- `stationPoll()` bails on `if (STATION_OK !== true) return false;`, so the
  poll can never bootstrap it.
- The only two callers of `stationReadIfNeeded()` require the station board to
  be picked or a glass job's drawer to be open. The ordinary job list called
  neither.

The result: sign in, chips appear; press F5 six minutes later and there are no
chips for the rest of that session — the owner's original complaint, reproduced
by the feature built to fix it. It would not have shown up in a demo, because a
demo starts from a fresh feed.

**The rule is amended to:** do not add a new `fetch` or a new list call of your
own. Calling the *existing* `stationReadIfNeeded()` on the plain-list path is
allowed and expected — it is a shared read with its own in-flight guard, not a
new one.

The implementer followed the spec as written; §6 asked it to flag a spec that
was wrong and it did not. Both halves of that are worth remembering.

### 2. Fixes taken in the one review pass (2026-09-09)

Critical and major, all with tests that fail before and pass after:

1. `STATION_ITEMS` is read on the ordinary-list path, so the chip survives the
   feeder's ten-minute skip and a page reload.
2. The rows repaint after `feedStation()` resolves whether or not a board or
   drawer is open — otherwise the first load of every session shows no chips
   for up to 60 s and polls at the slow rate.
3. The `redrawStation()` hold-off is narrowed to genuine text entry, and a
   dirty flag forces one repaint on the next tick. Previously a focused row
   checkbox — every row has one, and ticking one does not re-render — held the
   repaint off indefinitely with no catch-up, leaving the office reading a
   stale number with nothing to say so.
4. Poll-driven repaints no longer re-run the row entry animation. Replacing all
   of `#rows` on every `moved` poll made ~750 rows fade in and out as often as
   every ten seconds, destroying text selection and hover state.
5. A job→record cache keyed by `STATION_ITEMS` array identity replaces the
   per-row `ST.jobRecord` walk (measured at 11–16 ms per render at ~100 floor
   rows, and the `Glass station` list is never pruned — rows leaving production
   are only flipped to `Active = No` — so it grows without bound).
6. The chip renders *after* the comments badge, so it cannot clip the comment
   count out of an `overflow:hidden` row on a narrow screen.
7. Section 14b clears `stationPollT`, as the pre-existing sections do.

### 3. Known deviations, accepted rather than fixed (2026-09-09)

- `ROWS_GLASS` is computed from `filtered()`, not from the truly visible rows:
  per-group search boxes and collapsed sections are not subtracted. It can only
  ever err toward polling *too fast*, never too slow.
- While a hold-off is active the poll still runs at 10 s and discards the
  result. Wasted requests, not a correctness bug.

### 4. The chip has its own column (owner, 2026-09-09)

**This supersedes §4.1's "keep the text short" and Amendment 2 item 6.** Both of
those were attempts to fit the chip inside the row's last cell (`On sheets`,
`overflow:hidden`), and both failed. Measured in a real browser at 1440px:
every single glass row clipped — the cell offered 224px and wanted 236–292px —
and `Glass 48/48` rendered as `Glass 48/4`, a plausible but wrong number, which
is worse than showing nothing.

The row grid therefore gained a ninth column, **96px**, between
`Components F·S·T` and `On sheets`, with a `Glass` heading. 96px was measured,
not guessed: the widest chip any real job can produce (`Glass 120/360`,
`Glass 999/999`) renders at 89.13px in the fallback font. The cell deliberately
has **no** `overflow:hidden` — with slack in hand, an unexpected font should
spill a pixel into the gap rather than clip a digit and print a wrong number
again. A job the floor has never been fed renders that cell as an empty
`<span></span>`: same cell count as any other row, no nought, no placeholder.

The narrow-screen rule at `index.html` needed no change — the new cell is child
8 of 9, so the existing `.thead>:nth-child(n+5),.row>:nth-child(n+5)` catches
it. `.thead` and `.row` are each emitted in exactly one place, and neither the
glass station board (`.stcard`, its own card grid) nor John's print sheet
(`.jhead`/`.jrow`, its own nine-track grid) uses this grid, so nothing else
misaligned. A test now reads `index.html` from disk and asserts head cells, row
cells and grid tracks all number nine, so the three drifting apart fails a
check.

### 5. Known, pre-existing, not caused by this feature

- The `On sheets` badge cell still overflows on a maximal row (two sheet chips
  + Urgent + In fab + "n in progress" + a comment count wants ~388px and gets
  241px), so the comment count can clip. This change makes it *better* — the
  cell sheds a ~90px chip and loses only 28px of track — but it does not fix
  it, and no test asserts it.
- Below 900px, row cells 7 and 9 carry inline `style="display:flex…"`, which
  beats the media query's `display:none` on specificity, so they are not hidden
  and spill under the County/Customer headings. Reproduced identically on the
  pre-change markup. The master page has no mobile layout; the owner has
  deferred one.

### 6. Still deferred (owner, 2026-09-09)

Mirroring the floor's counters into the `Production` sheet's glass checkpoint
colours. The owner asked for it, was told it would be the first automatic write
to the master sheet and can race the office ticking the same job, and agreed to
spec it separately. **The `Production` sheet is still never written by this
feature.**
