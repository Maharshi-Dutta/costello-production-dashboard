# Day sheets for welding and hotmelting; one Floor log for every station

Owner request 2026-09-28, answers the same morning. Status: **approved to build
2026-09-28** (check-in done in the session; owner answered every question).

## What the owner asked, in their words

- *"on the top of [the] welding dashboard she can write in a separate box how
  many squares she welded on a daily basis, her welding machine tells her how
  many she did daily."* One number: squares.
- Glass cutting: the existing End of day sheet (Clear, K-glass, Satin, Obscure,
  shipped 2026-09-21) already covers it. **Do not touch it.**
- *"for hotmelting, how many units done per day, d/g & t/glaze ... maybe more
  automated like a counter"*. Owner chose **the mix**: the tablet shows the
  person's own hotmelt count for today, counted from their taps; the person
  types the DG / TG split.
- Weekly targets for welding and hotmelt: **not now** ("maybe in the future").
- *"the floor log should cover fabrication and welding dashboard logs"*. Owner
  chose **B**: the office's Floor log window covers every station, and also
  shows the office's own edits to floor counters (read from `Dashboard Log`),
  marked "office".
- Welding's day sheet list lives in `Floor stations` (owner: yes). The hotmelt
  count is **per person** (owner: yes).

## Hard rules (unchanged, restated because they bite here)

1. The workbook is not written by anything in this feature. No `/workbook` call
   is added anywhere. The Floor log reads `Dashboard Log` only from the
   workbook download the office page **already** makes (`app.js` ~5211, the
   `CHANGES` parse); no new read of the file.
2. Tablets never touch the workbook. The office never writes `Station log`,
   `Station people`, `Station comments`. The Floor log window is **read-only**.
3. `Station day sheets` stays append-only from the tablet: one POST per person
   per station-stage per day, no PATCH, no DELETE. The office may PATCH only
   the count columns, `Note`, `EditedBy`, `EditedAt` (`ST.dayOfficeFields`
   shape). No code path deletes a day sheet.
4. No real names/addresses in repo, tests, fixtures, docs. Placeholders only.
5. Edit files with the Edit/Write tools only. **Never** a shell heredoc or
   python rewrite of a source file (HISTORY B16: that wiped `app.js`).
6. Do not commit, do not push, do not touch any live list or tenant.

## Data model

### `Station day sheets` — three new Number columns (0 decimals)

| column | station · stage | meaning |
|---|---|---|
| `Squares` | Welding · `weld` | squares welded that day, from the machine's own display |
| `DG` | Glass · `hotmelt` | DG units hotmelted that day |
| `TG` | Glass · `hotmelt` | TG units hotmelted that day |

Hotmelt also stores what the tablet counted, so the office can see it beside
what was typed:

| `Counted` | Glass · `hotmelt` | the tablet's own count of this person's hotmelt taps today at the moment of saving (net of `−` taps, floor 0). Written by the tablet only; the office never corrects it |

So four new columns in all: `Squares`, `DG`, `TG`, `Counted`. The manager adds
them by script (rehearsed first) to the existing list in the workbook's site,
and creates a **second** `Station day sheets` list in `Floor stations` for
welding (same columns as the first, `Title` unique). The implementer does
nothing on the tenant; missing lists/columns must already be a quiet,
explained state (the existing `DAY_MISSING_*` path).

Title key stays `<Station>|<Stage>|<YYYY-MM-DD>|<Who>`. Welding's is
`Welding|weld|…`, hotmelt's `Glass|hotmelt|…`.

`WeekTarget` is simply not written for these two stages (no target is set).

### Definitions (the way the 2026-09-21 code was designed to grow)

- `ST.GLASS.daySheets.hotmelt = { counts: [["DG", "DG units hotmelted", "DG"],
  ["TG", "TG units hotmelted", "TG"]], unit: "units", counted: <fn or flag>,
  target: false }`
- `WELDC.WELD.daySheets.weld = { counts: [["Squares", "Squares welded today",
  "Squares"]], unit: "squares", target: false }`
- `cut` gains nothing; `target` absent means today's behaviour (target shown).
- `target: false` means: the tablet never reads `Station targets` for that
  stage, shows no target line, writes no `WeekTarget`; the office's Day sheets
  window shows no "Set target" control for it. A missing `Station targets`
  list in `Floor stations` must produce **no** warning on the welding tablet.
- `Counted` must be allowed through the tablet's POST field filter for the
  hotmelt stage only, and must **not** be in `dayOfficeFields`' output.

## Behaviour

### Welding tablet (`welding.html` / `welding.js`)

- An **End of day** button in the header's top row, beside the Frames / Sashes
  capsules (same look as the Cutting page's). Tablet layout rule: 10-inch
  portrait, 700–1100 px wide, **never scrolls sideways**, header stays two rows.
- Opens the same sheet the Cutting page opens: date (today), who (the person
  signed in), one number box "Squares welded today", the note line, Save.
  Once saved: shows the saved number, "ask the office to correct a mistake",
  no second save that day (list-unique + page check, exactly as cutting).
- Draft kept per device like cutting's (`cw_…` draft key must be per page so
  the welding and glass drafts never collide on one tablet).
- **Reuse, do not copy.** The day-sheet UI lives in `station.js` today
  (`openDaySheet`, `DAYDRAFT`, `draftNow`, `saveDraft`, `clearDraft`, the save
  path). Move the station-agnostic part into `station-ui.js` (shared by the
  tablet pages) and call it from both `station.js` and `welding.js`. The core
  helpers stay in `station-core.js` (`daySheetOf`, `dayFieldsFor`, `dayFields`,
  `dayTitle`, …). Watch the shared global scope: `test_pages.js` must stay
  4/4+ (B29 — a duplicated top-level `const` breaks a page in the browser).
- Lists resolve through welding's own site (`WELD.site`, `Floor stations`).

### Hotmelt tablet (`glass.html?stage=hotmelt`)

- Same End of day button as the Cutting page, now also on the hotmelt page.
- The sheet opens with a line at the top: **"You hotmelted N units today"**,
  N = sum over this person's `Station log` lines with `Stage = hotmelt` and a
  local-day `At` of today of `(To − From)`, floored at 0. Pure function in
  `station-core.js` (e.g. `ST.dayCounted(logRows, who, day, stage)`), tested.
  The tablet must have today's log lines for this; if it does not already
  read them, read them the way it reads the log for anything else (delta
  token), no new list.
- Two boxes: DG units, TG units. **Prefill neither** (the person types). When
  DG + TG ≠ N on Save, show one plain confirmation ("You counted N taps today
  and typed M. Save anyway?"); Save anyway saves, Back returns. N = 0 with no
  taps is not an error.
- `Counted = N` is written with the row.
- The Cutting page's sheet is unchanged: same counts, target, words.

### Office: Day sheets window (`app.js`, `openDaySheets` / `renderDaySheets`)

- Today it shows one station-stage (`dayStage()`). It gains a **station-stage
  selector** at the top: Cutting / Hotmelting / Welding, from the definitions
  (`ST.daySheetStages` over every station definition the office knows:
  `ST.GLASS`, `WELDC.WELD`; glazing and fabrication have none). Default: the
  current board's stage if it has one, else Cutting.
- Columns come off the chosen stage's `counts` (`dayCountShort`). Hotmelt also
  shows `Counted` read-only beside DG / TG, and flags a row where DG + TG ≠
  Counted (a small "≠ counted" marker, no colour alarm).
- Correction works per stage exactly as cutting's (`ST.dayOfficeFields` with
  that stage's counts), one `Dashboard Log` line.
- Target controls only for a stage without `target: false`.
- Welding's rows come from `Station day sheets` in `Floor stations`; glass
  stages' from the workbook's site. Each read only when the window is open on
  that stage (or already cached); a missing list is the quiet explained state.
- The "Day sheets" chip shows on every board now that more than one stage has
  a sheet (keep the definition-driven test, `dayStage()` generalised).

### Station reports

- The welding report and the glass report's hotmelt stage carry their day
  sheet columns on the Days sheet, the same way cutting's do today
  (`reportStages` / `reportLogStages` unchanged; the report reads the sheet
  counts through the definition). `Counted` included for hotmelt. Rule 3 strip
  on `Note` as today. One `Dashboard Log` line per export as today.

### Office: Floor log window (`openStationLog`, ~`app.js` 7722, 8172)

Today it lists glass lines only (`ST.logRows(STATION_LOG)` defaults to
`Glass`), whatever board is open. After:

- **Station filter**: All / Glass / Welding / Glazing / Fabrication, beside
  the existing person / stage / job / day filters. Opening from a board starts
  on that board's station; from the home list, All.
- **Sources**: `STATION_LOG` (glass), `WELD_LOG`, `GLZ_LOG`, `FABR_LOG`,
  each through `ST.logRows(items, <that station's name>)`, merged newest
  first. Each station's log is read on open if not already loaded (reuse the
  existing `weldReadIfNeeded` / `glzReadIfNeeded` / `fabrReadIfNeeded` /
  `stationReadIfNeeded`); one station's list missing or unreadable shows one
  quiet line for that station and the rest still show.
- **Office edits** (owner's B): lines from the office's own `Dashboard Log`
  rows that are edits of a floor counter — the `noteChange` "what" values the
  boards already write: "Glass cutting", "Glass hotmelting", "Glass tuff",
  welding's, "Glazing: <job>", fabrication's (grep `noteChange(` in the
  welding/glazing/fabrication/glass board code for the exact words; match
  them by a list kept in one place). Shown in the same list with a small
  **office** tag and the person, from → to, time. They come from the
  already-parsed `CHANGES` / sheet rows; nothing new is read. Day sheet
  corrections ("Day sheet corrected") also count as office floor edits.
  `Dashboard Log` time strings are `dd/mm/yyyy hh:mm` local; parse them to
  sort alongside the ISO `At` stamps. A Dashboard Log line whose station is
  unknown goes under All only.
- Stage filter options follow the chosen station (glass: cut / hotmelt /
  tuff; welding: frames / sashes / remake; glazing: glaze / astragal;
  fabrication: frames / sashes / transoms). Person filter lists floor people
  and office people who appear.
- Each line shows its station (a short tag) when the filter is All.
- The job drawer's glass timeline (`stationTimelineHtml`) is **unchanged**.
- The per-board Floor log panels (welding/glazing/fabrication boards) stay.
- Fix the latent sort bug while here only if it is on this path: `logRows`
  sorts ISO text; the merged list must sort by parsed time.

## Tests to deliver (offline, existing pattern)

- `test_daysheets.js`: hotmelt and weld definitions; `dayFieldsFor` per stage;
  Title keys; POST body for weld (`Squares`) and hotmelt (`DG`, `TG`,
  `Counted`) with no `WeekTarget`; `dayOfficeFields` for hotmelt never
  carries `Counted`/`Day`/`Who`/`Title`; `ST.dayCounted` (today only, this
  person only, hotmelt only, net of minus taps, floor 0, local day boundary
  at 23:30/00:30); `target: false` reads no targets list; cutting unchanged
  (every existing assertion still passes).
- `test_welding.js`: the End of day button exists; save goes to the
  `Floor stations` site's `Station day sheets`; no workbook call; second save
  same day refused.
- A floor-log test (new `test_floorlog.js` or inside an existing suite, add it
  to the verification command): merge of four stations + office lines, order,
  station filter, stage filter per station, one station missing still shows
  the rest, office-line matching list, `dd/mm/yyyy hh:mm` parse.
- `test_pages.js` still passes for all five pages.
- Every suite in the root `CLAUDE.md` verification command green.

## Report back

Files changed, what moved from `station.js` to `station-ui.js`, new functions,
exact test output of the full verification command (failures only is fine plus
the tail lines), anything in this brief you could not do or did differently and
why, and the exact `noteChange` "what" strings you matched for office lines.

## Not built

Weekly targets for welding and hotmelt; any automatic DG/TG split (the list
does not know a job's DG/TG — `GlassType` is one literal per row); a machine
link; changes to the cutting sheet.
