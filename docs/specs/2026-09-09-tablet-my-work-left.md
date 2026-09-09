# The tablet counts down the work *you* have left

**Date:** 2026-09-09
**Status:** in progress
**Supersedes:** the header total shipped in `2026-09-09-tablet-glass-total.md`
(commit `3d26677`, build `20260909-1509`). That total was the floor's remaining
glass, whole jobs at a time. This replaces it with a per-person countdown.
**Consulted:** yes — the owner described the behaviour on 2026-09-09, was asked
to confirm the reading back, and confirmed it, including the multi-stage rule.

---

## 1. What the owner asked for, in their own terms

> "each job id has 3 stages cutting which is done by David, glazing and
> hotmelting — it should show 14 for each of them. But for David, when he
> clicks 1 out of 14 done it changes the total job for him which is 14 to 13.
> And it would be separate for each person."

And, for the header: job 5035 has 14, job 5041 has 13, so the total is 27; one
tap takes it to 26.

## 2. The model

**Every stage has its own count.** A 14-glass job is 14 to cut, 14 to hotmelt
and 14 to glaze — three independent counts, which is already exactly how the
`Glass station` list stores them (`Cut`, `Hotmelt`, `Glazed`, each clamped to
`Total`). Nothing about the data model changes; only what is drawn from it.

**A person's remaining on one job** is the number of glasses that still need
something from *them*:

```
mine(job, stages) = total - min( counter[s] for s in stages )
```

- One stage: `total - counter[stage]`. David cuts one of 14 → he sees 13.
- Two stages (say cut and hotmelt): a glass is through *his* part only when
  both are done on it, so `min` is right. Cut 3, hotmelt 1 → one glass is fully
  through his part → **13**. The owner chose one number for such a person, with
  the individual stages still tracked underneath (which the steppers already
  show).
- No stages at all: `min` over nothing is meaningless, so it falls back to the
  job's plain total — a person who cannot tap anything has all of it "left".
  They cannot change it either; the steppers are already greyed for them.

**The counts are independent between people.** David cutting one does not
change what the glazing person sees. That falls out of the rule above — no code
needs to enforce it.

**The header total** is the sum of `mine(job, stages)` over every job on the
board, for the person signed in now.

## 3. Consequences the owner has been told about, and accepted

1. **The number is personal, not the floor's.** It changes on *Switch person*.
   Two people at the same tablet would see different totals. This is inherent
   to what was asked for and the owner said yes to it explicitly.
2. **The card's headline number changes meaning** — from "how big this job is"
   to "how much of it is left for you". That is a change to what the floor
   looks at, not an addition to it.
3. **"Finished jobs are excluded", from the superseded spec, is now
   automatic.** A job where the person's stages are all complete contributes 0
   to their total by arithmetic; no special case is needed. Delete the special
   case rather than keeping it alongside — two rules doing one job is how they
   drift apart.
4. **The office and the tablet now say different things about the same job,
   deliberately.** The master dashboard's `Glass` column counts all three
   stages across everyone (`Glass 12/24`); the tablet counts one person's
   remaining. Neither is wrong; they answer different questions. Do not try to
   make them agree.

## 4. Hard rules

1. **The tablet never touches the workbook.** No read, no write, no request of
   any kind added — this is arithmetic over the board already in memory.
2. **No write to any SharePoint list.** Display only. In particular this must
   not write a per-person total anywhere: it is derived on the device, every
   time it is drawn.
3. Nothing personal in the header or on the card beyond what is already there.
4. No real person's name, email or company domain in code, comments, tests or
   this spec. **The owner's screenshot of the live tablet contains a real
   customer name — it must not be reproduced anywhere in the repo.** Use
   "Person A"/"Person B" and "Customer One".
5. `test_station.js`'s standing proofs stay green, especially that the station
   page cannot reach the workbook and that no DELETE is ever sent.

## 5. What to build

### 5.1 `station-core.js` — the arithmetic

Replace `boardGlassTotal(board)` (added in the superseded spec) with the
per-person pair. Both pure, both exported on `ST`, both beside `glassWords`:

```
jobLeftFor(g, stages)      // total - min(counter over stages); total when stages is empty
boardLeftFor(board, stages) // sum of jobLeftFor over the board
```

- `stages` is an array of stage keys (`cut` / `hotmelt` / `glazed`), i.e. what
  `PERSON.stages` already holds. Ignore any key that is not a real stage —
  `station-core.js` already treats an unknown stage as absent rather than an
  error, and this must not be the one place that throws.
- Clamp: never negative, never `NaN`, whatever a malformed row says. Reuse
  `stNum` and the existing clamping, as `boardGlassTotal` did.
- `boardLeftFor(null, …)` is 0; an empty board is 0.
- **Delete `boardGlassTotal` and its tests.** Do not leave it behind as dead
  code "in case" — it is superseded, and a second total is exactly the thing
  that drifts.

### 5.2 The card

The headline number on each card becomes `jobLeftFor(g, PERSON.stages)`.

**Wording.** It currently reads "14 glasses" (`ST.glassWords`). A number that
now means *remaining* must say so, or it is a lie by omission on a workshop
screen. Render it as **"14 left"** ("1 left" in the singular), and keep
`glassWords` itself untouched — the office still uses it and it still means the
job's size there. Add the new wording as its own small helper next to it.

Flag this in your report: the owner asked for the number to count down but did
not ask for the words to change, and I am choosing the wording. It is the one
thing here they may want different.

### 5.3 The header

`#gtotal` shows `boardLeftFor(board, PERSON.stages)` in the same wording.

Everything else about the header stays exactly as shipped: it does **not**
follow the search box, it is hidden whenever the search box is hidden, it is
shown when a search matches nothing, and it is cleared rather than left stale.
Those rules and their tests survive this change — only the number inside it
changes. Do not re-litigate them.

### 5.4 Redraw

The card's number now depends on `PERSON`, so a *Switch person* must redraw
every card, not only the header. Check `boardDiff`/`dressCard`: if the card
signature does not include the person's stages, a switch between two people
holding different stages could leave stale numbers on screen. Fix it if so, and
say what you found either way.

## 6. Tests to deliver

1. `jobLeftFor`: one stage, counting down one tap at a time; two stages using
   `min` (cut 3 / hotmelt 1 of 14 → 13); no stages → the job's total; a
   completed job → 0; unknown stage keys ignored; malformed counters never give
   `NaN` or a negative.
2. Independence: David cutting one does not change what the glazing person's
   number shows for the same job.
3. `boardLeftFor` over the owner's own example — 14 and 13 → 27, one cut tap →
   26 — asserted from the rule, not read off the implementation.
4. A tap changes the card and the header **before the write leaves the tablet**
   (the queue overlay in `boardNow()`), as the superseded spec's test did.
5. *Switch person* to someone with different stages redraws every card with
   that person's numbers.
6. A person with no stages sees the job totals and can still tap nothing.
7. The header rules that survive: not narrowed by the search, hidden with the
   search box, shown when a search matches nothing, cleared when hidden.
8. No request goes out to compute or draw any of it.

## 7. Out of scope

- The office dashboard and its `Glass` column. See §3.4 — they are meant to
  differ.
- Any new column anywhere (owner, 2026-09-09).
- Any change to the counters, the queue, the log, the picker, the PIN or the
  poll.

## Amendments after review

### 1. `min` becomes a sum (owner, 2026-09-09) — supersedes §2 and §5.1

The review demonstrated, by running the shipped function, what `min` looks like
on a wall for a person holding two stages:

```
cut 0, hotmelt 1 -> "14 left"
cut 1, hotmelt 1 -> "13 left"
cut 2, hotmelt 1 -> "13 left"
...
cut 6, hotmelt 1 -> "13 left"
```

Six taps, each doing real work and writing a real log line, with the headline
frozen and the stepper directly beneath it counting up correctly. That is
precisely the symptom that started this whole line of work — "it is not
updating the status" — and it would have been read as a fault on the floor.

The owner confirmed that people in the live `Station people` list **do** hold
two stages, and chose to count each stage separately. The rule is now:

```
jobLeftFor(g, stages) = sum over stages of (total - counter[stage])
```

- One stage: `total - counter`, unchanged. The owner's original example still
  holds exactly: 14 + 13 = 27, one tap → 26.
- Two stages on a 14-glass job with nothing done: **28**. Every tap of either
  stage drops it by one, so the number always moves when the person works.
- Three stages: 42.
- No stages, or only unknown stage words: the job's plain total, unchanged.
- The result is an integer in `0 … total × (stages held)`; never negative,
  never `NaN`.

**Consequence, accepted:** for a multi-stage person the number is no longer a
count of *glasses* — it is a count of what they have left to do. This is why
the wording stays "28 left" and does not go back to "glasses": "left" is true
for both cases, "glasses" would be false for one.

### 2. The gold rule stays job-level (owner, 2026-09-09)

The review pointed out that a single-stage person can finish their part of
every job and end the day with a board of "0 left" cards, all still in full
colour, none folded into the Finished group — because gold and the fold are
driven by the *job* being complete for everyone, not by the person's own part.

The owner was asked and chose to leave it. Gold means the whole job is done.
Do not make it per-person without asking again.

### 3. Fixed a bug that was already live, unrelated to this feature

`readPeople()` re-resolves the signed-in person from the list every ten
minutes, so editing somebody's `Stages` column in SharePoint swaps
`PERSON.stages` with no picker in between. Neither `boardDiff` nor `qState`
could see that, so the greyed steppers went stale — and in the direction that
matters: a stage *removed* left a live-looking, tappable button that `tap()`
then silently refused. `PSIG`/`pState()`, added here for the headline number,
closes both directions. The tablet ships more correct than it was.

### 4. Documentation debt cleared at the same time

The superseded `2026-09-09-tablet-glass-total.md` still read "in progress" with
no pointer here, so a cold session following `CLAUDE.md` would have found a
live-looking brief telling it to build a function that no longer exists. Marked
superseded. Both tablet specs were also missing from `docs/specs/README.md`,
which `CLAUDE.md` calls the index of every spec; both added.

Still outstanding at ship: `docs/USER-GUIDE.md` (line ~19 and §4.2) still
describes the card as "one total number of glasses", which is no longer true.
It has a generated PDF in `docs/pdf/`, so the Markdown and the PDF must be
redone together or they desync.
