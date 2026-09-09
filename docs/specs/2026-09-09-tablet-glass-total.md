# A running total of glass on the tablet's header

**Date:** 2026-09-09
**Status:** shipped 2026-09-09 (`3d26677`) — **superseded the same day** by
`2026-09-09-tablet-my-work-left.md`, which replaced this floor-wide total with a
per-person countdown. `boardGlassTotal` no longer exists; do not build from §4
of this file.
**Consulted:** yes — the owner asked for "a total number of glass including all
the jobs in the glass section dashboard, beside the search bar" on 2026-09-09,
and settled the two ambiguities the same day (see §2).

---

## 1. Context

The glass station tablet (`glass.html` + `station.js` + `station-core.js`) shows
one card per job with the job number, the customer, "12 glasses" and the three
steppers. Its header carries the brand, who is signed in, Switch person, a
"Find a job" search box, an "updated …" line, the theme button and Sign out.

There is nowhere on that screen that says how much glass is on the floor
altogether. This spec adds it, beside the search box.

## 2. The two decisions the owner made

1. **The total does NOT follow the search box.** Typing in "Find a job" narrows
   the cards; the total stays exactly as it was. The number always means "this
   is how much glass is on the floor", never "how much matches what I typed",
   so it cannot be misread as a smaller workload.
2. **Finished jobs do NOT count.** The number is glass **still to do**. A job
   whose three counters are all at its total drops out of the number entirely.
   A job that is partly done still contributes its **whole** total — the number
   counts glass on the floor, not remaining stage steps. (If the owner later
   wants remaining *steps* instead, that is a different number and a different
   spec.)

## 3. Hard rules

1. **The tablet never touches the workbook.** Nothing in this feature reads or
   writes the Excel file, and no new request of any kind goes out — the number
   is arithmetic over the board the tablet already holds in memory.
2. **No write to any SharePoint list.** This is display only.
3. No phone number, eircode, customer detail or anything personal in the
   header — it is a count and a word.
4. No real person's name, email address or company domain in code, comments,
   tests or this spec.
5. `test_station.js`'s standing proofs stay green, in particular that the
   station page cannot reach the workbook and that no DELETE is ever sent.

## 4. What to build

### 4.1 The number

`station-core.js` gains one pure function, exported on `ST` beside
`glassWords`:

```
boardGlassTotal(board)   // sum of g.total for every g where !g.finished
```

- `board` is what `ST.jobBoard(items)` returns — the same array the tablet
  already draws from, **before** `ST.boardFilter` is applied.
- A job with `total` 0 contributes nothing (it already cannot be `finished`).
- Returns a plain number, never `NaN`, never negative.

Put it next to `glassWords` in `station-core.js` and export it on `ST`. It is
pure and read-only, like `glassWords` and `boardFilter`.

### 4.2 Where it goes

`glass.html`: a `<span id="gtotal">` immediately **after** `<input id="search">`
in `header.top`, before `<span id="upd">`.

Style it to sit with the search box rather than shout over it: the same
`var(--brand-2)` background, `var(--brand-line)` border and 24px radius as
`#whois` and `#search`, `min-height:40px`, centred text. It is read, not
tapped, so it does not need a 44px target — but it must not shrink the ones
either side of it, and the header must still wrap cleanly on a narrow tablet
(`flex-wrap:wrap` is already on `header.top`).

Both themes: use the existing custom properties only. No new hex literals.

### 4.3 What it says

`ST.glassWords(n)` — the tablet's own wording, already used on every card:
`"1 glass"`, `"12 glasses"`, `"248 glasses"`.

### 4.4 When it shows and when it does not

It follows the same rule as the search box, which is already computed in
`render()` as `boarding`:

```
const boarding = !PROBLEM && PEOPLE_READ && !!PERSON && READY;
```

- Hidden whenever the search box is hidden — on the picker, on any error
  message, before a person has been chosen. A total floating over "Ask the
  office to grant the SharePoint permission" only looks broken.
- Shown whenever the board is shown, **including when a search matches
  nothing**. The board being empty because of a search does not change how much
  glass is on the floor, and that is exactly the case where a total that
  vanished would be most confusing.

Set its text in the same place the header's other pieces are set (`render()`,
next to `#upd` and `#switchbtn`), from `boardNow()` — **not** from the filtered
board.

## 5. Tests to deliver

In `test_station.js`, following the patterns already there:

1. `ST.boardGlassTotal` sums the totals of unfinished jobs and **excludes**
   finished ones; a board of all-finished jobs gives 0; an empty board gives 0.
2. A job with `total` 0 contributes nothing.
3. The number ignores the search: with a `QUERY` that matches one job, the
   header still shows the whole board's total.
4. The header shows it while the board shows, and hides it on the picker and on
   an error, exactly as the search box does.
5. It survives a search that matches nothing (board empty, total still shown
   and still the full figure).
6. A tap that finishes a job drops that job's glass out of the total on the
   next render.
7. No request goes out to compute or draw it — assert over the whole run as the
   existing sweeps do.

## 6. What to report

- The diff, and anything the spec did not pin down that you had to decide.
- Verbatim output of every check in `CLAUDE.md` § "Running the checks".
- Confirmation that no workbook read/write and no list write was added.
- Anything in this spec that is wrong or impossible — say so rather than
  working around it.

## 7. Out of scope

- The office's own Glass station board and the master dashboard. This is the
  tablet only.
- Any change to the cards, the steppers, the finished-group folding, the
  search behaviour or the poll.
- Any new column anywhere. The owner has asked for no more columns
  (2026-09-09).

## Amendments after review

### 1. Reviewed clean; one comment corrected (2026-09-09)

An independent review found **no critical and no major** findings. The one
thing fixed in place: the `#gtotal` CSS comment had its causality backwards. It
claimed `flex:none` stops the pill "eating into" its neighbours. It does not —
`flex:none` stops the *pill* shrinking, which makes the header more likely to
wrap, not less. What actually protects the 44px tap targets either side is
`flex-wrap:wrap` on `header.top`: items that will not fit wrap rather than
shrink. The declaration was right and stays; only the explanation was wrong,
and a wrong comment in this codebase is a defect in its own right.

### 2. Recorded, not fixed

- **The header wraps sooner on a wide tablet.** The header's content is roughly
  1126px before this change, so it already wraps at 1024px landscape; the
  ~115px pill makes wrapping likely at 1280px too. `header.top` is
  `position:sticky`, so a wrapped second row costs about 52px of card space
  permanently on a wall-mounted screen. Estimated, not measured — check it on
  the real tablet rather than a desktop window.
- **Two small test gaps.** The malformed-board assertion covers `-4`, `""` and
  `null` but not `"abc"`, `NaN` or `Infinity` (all handled by `stNum`'s
  `isFinite` gate, verified by the reviewer but not defended by the suite). The
  no-hex-literal regex would not catch a named colour or `rgb()`.
- **`boardGlassTotal` throws on a non-array `board`.** Unreachable: the only
  caller passes `boardNow()`, which always returns an array, and `ITEMS` starts
  as `[]`.

### 3. The one thing the owner should confirm at the demo

With the Finished group expanded, its gold cards are on screen but excluded
from the header total, so adding up every visible card gives a larger number
than the header shows. That is exactly what §2.2 asks for, and the group is
separately headed "Finished · n" — but it is the one place somebody on the
floor could hand-compute a different figure. The same applies to §2.1: one card
reading "2 glasses" under a header reading "248 glasses" is the chosen
behaviour, and worth seeing once before it is settled.

### 4. Proofs the review established, worth not re-deriving

- The `boardNow()` → `live` substitution in `render()` is provably equivalent:
  the early-return condition is exactly `!boarding` by De Morgan, `boardNow()`
  is pure and is called the same number of times in the same states, and no
  statement between the two points can run script.
- The header cannot drift from the cards, structurally: `paintBoard` has
  exactly one caller, `render()`, so both are computed from the same
  `boardNow()` result in the same call.
- `finished` is one shared field set once in `buildJobs`, read by the gold
  class, the group fold and the total alike — so a job cannot be gold on the
  board and counted in the header, or the reverse.
