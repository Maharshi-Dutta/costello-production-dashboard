# The tablet counts glasses per stage, and tuff on its own

**Date:** 2026-09-10
**Status:** consulted, approved to build
**Supersedes:** `2026-09-09-tablet-my-work-left.md` (shipped `06dff7c`), whose
per-person single number multiplied by the stages held. The owner saw the
result — a job with 49 glasses and 10 tuff reading **157** — and rejected it.

---

## 1. What the owner said

> "the total for each job id is TG + DG. If either is not there the total
> number of glass is whichever is present for each job id (5033 has TG 49, no
> DG, that's the total number of glass). And for grand total that shows the
> total number of glasses altogether (every job id added) — so for example 5033
> has 49 glass + 5035 has 14 glass = 63."

> "for every person the job unit will be same. Like for [the cutter] he can cut and
> tuff, so for him the job glass counter goes down when the glass is cut (you
> can have separate counters for tuff) and the tuff counter goes down when tuff
> is done. Same for the others as it is a job unit — so if [the cutter] finishes 20
> cuts out of 40, 20 should be left for others."

And, asked what a person holding two glass stages should see: **"two numbers —
20 left to cut, 35 left to hotmelt."**

## 2. The rule

**A job's glass count is `DG + TG`** — whichever of the two is present. This is
unchanged (`ST.glassTotal`), and 5033 with TG 49 and no DG is 49.

**The card shows one number per stage the signed-in person holds**, each being
that job's glass count less that stage's own counter:

```
left(job, stage) = glassTotal(job) - counter[stage]
```

- The cutter holds cutting: 40-glass job with 20 cut → **"Cutting 20 left"**.
- Someone holding cutting and hotmelting, 20 cut and 5 hotmelted → **two
  numbers: "Cutting 20 left" and "Hotmelting 35 left"**.
- **Nothing is summed across stages.** The 157 was `49×3 + 10`; both the
  multiplication and the tuff were wrong.

**TUFF is its own counter, never folded into the glass number.** It counts
against `TuffTotal`, not against `DG + TG`:

```
left(job, "tuff") = tuffTotal(job) - counter["tuff"]
```

**The numbers are the job's, not the person's.** Two people who both cut see
the same "20 left", because 20 is genuinely what is left to cut on that job —
the owner's "20 should be left for others". This falls out of the rule above;
no code enforces it.

**The header sums the same numbers across the board**, one per stage held, so
the header and the cards always agree and both count down as work is done. A
person holding one glass stage — the common case — sees a single figure that
behaves exactly like the owner's example: 49 + 14 = 63, falling as they work.

## 3. What this changes

- `jobLeftFor(g, stages)` — a single summed number — is **replaced**, not kept
  beside a new function. Delete it and its tests; a second answer to the same
  question is how the two drift apart.
- `leftWords` stays, but the card and header now render one entry per stage.
- Nothing about the office dashboard changes. Nothing about the glass colours
  feature (`ab0c85c`) changes: `glassColours`, `stageComplete`, `COLOUR_TYPES`
  and `floorStamp` are untouched, and the colour rule still reads the raw
  counters, not these derived numbers.
- The gold card and the Finished group still mean the whole job is done by
  everyone (owner, 2026-09-09). Unchanged.

## 4. Hard rules

1. The tablet still never touches the workbook: no new request from
   `station.js`/`station-core.js` to anything but the three lists.
2. No write of any kind is added — these numbers are derived on the device at
   every draw and stored nowhere.
3. No real person's name, email or company domain anywhere.
4. `test_station.js`'s standing proofs stay green, especially that the station
   page cannot reach the workbook and that no DELETE is ever sent.

## 5. Tests

1. `left(job, stage)` per stage: cutting, hotmelting, glazing, tuff; tuff
   against `TuffTotal` and never against `DG + TG`.
2. A person with two glass stages gets **two** numbers, the owner's 20/35 case
   asserted exactly.
3. A person with one stage gets one; a person with none gets the job's glass
   count and can tap nothing.
4. Two people holding the same stage see the **same** number on the same job.
5. The header is the per-stage sums across the board, and the owner's
   49 + 14 = 63 example is asserted from the rule.
6. A tap moves both the card and the header before the write leaves the tablet.
7. `ST.jobLeftFor` is gone (`assert.strictEqual(ST.jobLeftFor, undefined)`).
8. The colour feature is unaffected: the same fixtures produce the same
   `glassColours` output as before this change.
9. No request goes out to compute or draw any of it.

## Amendments after review

_(append here; do not rewrite the sections above)_
