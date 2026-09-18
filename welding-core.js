/* Welding station - the pure logic of the second floor station.

   The welders work on their own page (welding.html) signed in with the same
   shared station account the glass tablet uses. It sees four SharePoint lists
   in the `Floor stations` site - "Welding station" (the work), "Station
   people", "Station log" and "Station comments" - and NO workbook of any kind.
   Nothing this feature does writes a cell, a colour or a value into the
   Production sheet, in either direction (spec 2026-09-16, "Hard rules").

   What the floor is told about a job is deliberately short, and the list of it
   is the owner's own reading of the `Cut & Weld PVC` header template: the job
   number, the customer's name, the COMMENT (with the rule-3 strip below over
   it), the "sent to floor" date, the WND/DRS quantities, and - per product
   group - the number of FRAMES and SASHES to weld. Never a phone number, never
   an eircode, never an area, never a price, never a colour, never a transom:
   "there is no transomes even if it is green" (owner, 2026-09-16).

   Nothing in this file touches a DOM, Graph, or a workbook. It is the slice the
   feeder should send, the rows the two boards draw, the clamp on a tap, the
   three-level colour rule, and the fields one write may carry. test_welding.js
   exercises every one of them directly.

   Loaded in the browser by index.html (before app.js) and by welding.html, and
   in Node by the tests, the same way station-core.js is. It is self-contained
   on purpose - its own small helpers rather than station-core's - so it can be
   loaded, read and tested on its own.

   ADDING THE NEXT STATION: docs/STATIONS.md, "Adding a station". This file is
   the shape to copy; station-core.js is the shape NOT to copy into, because
   everything in there is station-independent already.                       */

/* ---- names ---------------------------------------------------------------- */
const WELD_LIST = "Welding station";
/* the word in the `Station` column of `Station people`, `Station log` and
   `Station comments`. It is the ONE station-specific input each of those three
   shared lists takes - see ST.stationComments and ST.logRows. */
const WELD_NAME = "Welding";
/* Which site this station's lists are in. "floor" is the `Floor stations`
   site and nothing else; the glass station is pinned to "own" (the workbook's
   own site) until its lists are moved across - owner's decision 1 of
   2026-09-16, and the reason CW.stationSite() takes an argument at all. */
const WELD_SITE = "floor";
/* There is one stage here, not three: "no cutting, it should just have
   welding" (owner). Anybody signed in holds it, which is why the header's
   "N left" is the whole board's rather than a per-stage list. */
const WELD_STAGE = "weld";

/* ---- what is counted -------------------------------------------------------
   Two numbers per product group, both plain counts off the sheet: frames to
   weld and sashes to weld. `sub` is the sheet's own sub-header letter, and it
   is what makes the office's record reachable: the checkpoint item for a
   group's frames is `prod:<group>:f`. T (transoms) is deliberately absent. */
const WELD_PARTS = [["frames", "Frames", "f"], ["sashes", "Sashes", "s"]];
const WELD_PART_KEYS = WELD_PARTS.map(p => p[0]);
const WELD_PART_LABEL = { frames: "Frames", sashes: "Sashes" };
const WELD_PART_SUB = { frames: "f", sashes: "s" };
/* the list columns each part owns: the quantity (feeder), the counter (floor
   and the office's own board), and the counter's By/At pair */
const WELD_TOTAL_FIELD = { frames: "Frames", sashes: "Sashes" };
const WELD_DONE_FIELD = { frames: "FramesDone", sashes: "SashesDone" };
const WELD_BY_FIELD = { frames: "FramesBy", sashes: "SashesBy" };
const WELD_AT_FIELD = { frames: "FramesAt", sashes: "SashesAt" };
const weldPartLabel = k => WELD_PART_LABEL[String(k).trim().toLowerCase()] || String(k);

/* ---- the columns -----------------------------------------------------------
   The feeder writes Title, the job facts and FedAt/FedBy; the floor writes the
   eight columns of WELD_FLOOR_FIELDS and nothing else. The office's welding
   board writes the same eight - the second sanctioned office write of a floor
   counter after the glass clear (CLAUDE.md rule 2, dated 2026-09-16) - and
   never a line of `Station log`.                                            */
const WELD_FEEDER_FIELDS = ["Job", "Group", "GroupSeq", "Customer", "Comment", "SentToFloor",
                            "Wnd", "Drs", "Frames", "Sashes", "Seq", "Section", "Active"];
const WELD_FLOOR_FIELDS = ["FramesDone", "SashesDone", "FramesBy", "FramesAt",
                           "SashesBy", "SashesAt", "DoneBy", "DoneAt"];
/* the two floor columns that are numbers; everything else on that list is text.
   floorOnly() uses this to refuse a counter that is not a number and a stamp
   that is not a string, whatever somebody has edited into localStorage. */
const WELD_COUNTER_FIELDS = ["FramesDone", "SashesDone"];
/* The ONE exception to "the feeder never writes the floor's columns", and it is
   the same one the glass list has: a row the feeder is CREATING, or a row the
   floor has never tapped (DoneAt empty), starts from what the office's own
   record already says. A gold F cell means those frames are welded; yellow
   seeds nothing (owner, answer 5). Once the floor's first tap has set DoneAt
   the feeder never writes a counter on that row again. */
const WELD_SEED_FIELDS = ["FramesDone", "SashesDone"];
const WELD_FIELDS = ["Title"].concat(WELD_FEEDER_FIELDS, ["FedAt", "FedBy"], WELD_FLOOR_FIELDS);
const WELD_FEEDER_WRITES = ["Title"].concat(WELD_FEEDER_FIELDS, ["FedAt", "FedBy"], WELD_SEED_FIELDS);

/* ---- which product groups reach the floor ----------------------------------
   NOT a hard-coded list of the fourteen green ones. It is every F/S/T product
   group the parser already finds on the Production sheet (mapSheet -> j.prods)
   minus this short deny-list, so a new green group on the sheet is fed without
   a code change and a new red one is one line here. The names are matched after
   weldKey() - upper-cased, every run of non-alphanumerics collapsed to one
   space - because the parser's own header normalisation drops the "&" out of
   "ALUCLAD TILT & TURN" and nothing should depend on that detail twice. */
const WELD_DENY = ["ALU CLAD WINDOWS", "ALUCLAD TILT & TURN", "BIFOLD", "COMPOSITE"];

/* ---- which COMPONENTS a group has, where it is not both ---------------------
   Most groups have frames and sashes. A few do not, and the sheet still has a
   number in the other column - so "it is on the sheet" is not the same as "the
   floor welds it". A group named here is fed ONLY the parts listed:

     SUPER DOOR   sashes only. The frames number on the sheet is somebody
                  else's count and is fed as 0, not carried across.

   Fed as 0 rather than left out, deliberately: the column stays on the list
   (nothing creates or deletes columns), and a part at 0 draws no line on the
   tablet or the office board because both already skip a line with nothing to
   weld. So the row is complete, the Title is unchanged, and the welder is not
   shown a component their group has not got. Matched after weldKey(). */
const WELD_GROUP_PARTS = { "SUPER DOOR": ["sashes"] };
/** The parts this group is welded in, in the board's own order. */
function weldPartsFor(group) {
  const only = WELD_GROUP_PARTS[wKey(group)];
  if (!only) return WELD_PART_KEYS.slice();
  return WELD_PART_KEYS.filter(k => only.indexOf(k) >= 0);
}

const WELD_CUSTOMER_MAX = 70;
const WELD_COMMENT_MAX = 140;

/* ---- rule 3: nothing that could be a phone number or an eircode ------------
   The COMMENT column is free text typed by the office and it does carry phone
   numbers. The owner's answer of 2026-09-16 was "yes, show it, with runs of six
   or more digits removed" (option B), recorded as a rule-3 decision. This is
   that guard, and it is applied on the way INTO the list - so no phone number
   is ever stored where the floor, or anything reading the floor's list, could
   see one.

   THREE patterns, and the first of them is the one that matters. export.js has
   no eircode regex to borrow: it protects the eircode by never reading `j.eir`
   at all, which is the stronger rule where it applies and no help at all here,
   because a comment is a comment. So the shapes are written down once, here,
   and every one of them is tested in test_welding.js:

     · a PHONE NUMBER AS PEOPLE TYPE ONE - six or more digits with spaces,
       brackets, dots or dashes between them, and an optional leading "+":

           086 123 4567          +353 86 123 4567      087-123-4567
           086.123.4567          (086) 123 4567        08712 34567

       The first build of this had only the contiguous rule below, and every
       one of those six reached the list unchanged. A phone number in a
       workshop comment is almost never typed as ten digits in a row.
     · any run of six or more digits          0871234567  ->  …
       (kept as its own pass: it is the plain case and it reads as the rule)
     · anything shaped like an eircode        D02 X285    ->  …
       (a letter, two digits or a digit and W, then four alphanumerics)

   WHAT SURVIVES, and is tested so it stays surviving: a job number (`R5303` -
   a letter and four digits, two short of the eircode shape and four short of
   the phone one), the sheet's own short dates (`12.09`, `12/09/2026` - `/` is
   not a separator here, so neither half is long enough), and small quantities
   (`Wnd 6`, `4000 CASEMENT`, `7000`).

   WHAT IS OVER-STRIPPED, said rather than discovered: a full ISO date typed
   into a comment (`2026-09-16`) is eight digits with dashes between them and
   goes. That is the right way round to be wrong - the owner's rule is that no
   phone number leaves the app, ever, and a lost date in a floor comment costs
   nobody anything. The same is true of a dash- or dot-joined run of sizes -
   `cill 150-2100` -> `cill …`, `1200-900-1500` -> `…`, `pane 1.2.3.4.5.6` ->
   `…` - because it has the same digit-and-separator shape as a phone number;
   that is the same owner's decision, the same way round (rule 3 wins over a
   readable size). An `x`-separated size (`2 x 1200 x 900`) has no such shape
   and survives. */
const WELD_PHONE_RE = /(?:\+?\d[\s().‐-―-]{0,2}){6,}\d?/g;
/* a fourth shape, added 2026-09-16: digit groups joined by `/`, `,`, `:` or
   `_` - a shape none of the three patterns above catches, because none of
   those separators is in WELD_PHONE_RE's class and a run like `086/123/4567`
   is short of WELD_DIGITS_RE's six-CONTIGUOUS-digits bar. It only fires when
   the whole match holds nine or more digits, so a date (`12/09/2026`, eight
   digits) or a small size (`spacer 4/20/4`, four digits) is not caught. */
const WELD_SEP_RE = /\d+(?:[\/,:_]\d+)+/g;
const WELD_DIGITS_RE = /\d{6,}/g;
const WELD_EIR_RE = /\b[A-Za-z]\d(?:\d|[Ww])\s?[A-Za-z0-9]{4}\b/g;

/* ---- small helpers ---------------------------------------------------------
   Its own, deliberately: this file is loaded beside station-core.js and could
   borrow stTxt/stNum/stClamp from it, but then it could not be read or tested
   without it, and a second station's rules should not depend on the first
   station's file happening to be on the page. */
const wTxt = v => String(v == null ? "" : v);
const wKey = v => wTxt(v).trim().toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
function wNum(v, dflt) {
  if (v == null || wTxt(v).trim() === "") return dflt;
  const n = Number(v);
  return isFinite(n) ? n : dflt;
}
const wInt = (v, dflt) => Math.round(wNum(v, dflt == null ? 0 : dflt));
const wClamp = (n, total) => Math.max(0, Math.min(total, Math.round(wNum(n, 0))));
/** SharePoint item ids are numbers in a string; the oldest is the smallest. */
function wItemAge(a, b) {
  const na = Number(a && a.id), nb = Number(b && b.id);
  if (isFinite(na) && isFinite(nb)) return na - nb;
  const sa = wTxt(a && a.id), sb = wTxt(b && b.id);
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/** The group name as the list carries it: the sheet's own words, upper-cased,
    whitespace collapsed. It is half of the Title key, so it has to be stable
    whatever the sheet's header spacing does. */
const weldGroupKey = wKey;
/** May this product group go to the floor? Everything except the deny-list. */
function weldAllowed(group) {
  const k = wKey(group);
  if (!k) return false;
  return !WELD_DENY.some(d => wKey(d) === k);
}
/** The Title of one list row: `JOB|GROUP`, unique on the list. */
const weldTitle = (job, group) => wKey(job) + "|" + wKey(group);

/** Rule 3's strip: a phone number however it was typed, then any long run of
    digits, then anything eircode-shaped - each replaced by an ellipsis - then
    whitespace tidied and the result capped. Blank in, blank out.

    `max` defaults to the comment's cap. It is passed explicitly for the
    CUSTOMER column, which is also free text off the sheet and can also have a
    number in it ("Customer One 086 123 4567", a site contact typed into the
    name): the same guard, the same one function, a different cap. */
function weldStripDigits(text, max) {
  const cap = Math.max(1, wInt(max, WELD_COMMENT_MAX));
  const t = wTxt(text)
    .replace(WELD_PHONE_RE, "…")
    .replace(WELD_SEP_RE, m => ((m.match(/\d/g) || []).length >= 9 ? "…" : m))
    .replace(WELD_DIGITS_RE, "…")
    .replace(WELD_EIR_RE, "…");
  return t.replace(/\s+/g, " ").trim().slice(0, cap);
}

/* ---- what the feeder should be sending --------------------------------------
   ONE row per job AND product group. A job with casement windows and a PVC door
   is two rows; the job's facts are repeated on both, so a row is complete on its
   own and the tablet can build a card out of the rows of one job without
   looking anything up.

   EVERY job on the sheet, every section - the master sees them all and the
   owner asked for the office to be able to correct a finished job ("they might
   make it wrong"). The tablet is what narrows to "In production", off the
   Section column. A job that has left the sheet (cat "past") is not in the
   slice at all; feedPlan marks whatever row it already made Active = No and
   never deletes it.

   `statusOf` is how the office's own record reaches the seed: a function
   (job, itemKey) -> the record's word for that checkpoint item. It is
   cpStatus() in app.js and nothing else; passing nothing seeds every row at
   nought, which is what a caller with no record to hand should send.       */
function weldSectionOf(j, blockNames) {
  const names = blockNames || [];
  return wTxt(names[j && j.blk] || "").trim();
}
/** The job's COMMENT, as the parser left it on j.notes. */
function weldCommentOf(j) {
  const hit = ((j && j.notes) || []).find(n => n && n.k === "comment");
  return weldStripDigits(hit ? hit.t : "");
}
/** The "sent to floor" cell, shown as it was typed. A blank stays blank. */
function weldSentOf(j) {
  return wTxt(((j && j.dates) || {}).floor || "").trim().slice(0, 40);
}

function weldSlice(jobs, blockNames, statusOf) {
  const names = blockNames || (jobs && jobs.blockNames) || [];
  const look = (j, item) => {
    if (typeof statusOf !== "function") return "";
    try { return wTxt(statusOf(j, item)).trim().toLowerCase(); } catch (e) { return ""; }
  };
  const out = [];
  /* GUARD: one Title, one row in the slice. `Title` is unique on the list, so a
     slice carrying it twice would make the feeder POST a row SharePoint then
     refuses, every run, for ever. It should not be reachable - the parser folds
     a job's product groups into one entry each - but a sheet with the same
     group header in two column blocks, or a job number appearing twice, would
     do it, and the feeder is the wrong place to find that out. The first one
     wins and the second is said out loud rather than swallowed. */
  const emitted = {};
  (jobs || []).forEach(j => {
    if (!j || !j.id) return;
    if (j.cat === "past") return;                 // not on the sheet any more
    const job = wKey(j.id);
    if (!job) return;
    /* the CUSTOMER column gets the same rule-3 strip the comment does. It is
       free text off the sheet like any other, and a site contact's number
       typed into a customer name would otherwise walk straight onto the
       floor's list. */
    const customer = weldStripDigits(j.cust, WELD_CUSTOMER_MAX);
    const comment = weldCommentOf(j);
    const sent = weldSentOf(j);
    const section = weldSectionOf(j, names);
    const wnd = Math.max(0, wInt(j.wnd, 0)), drs = Math.max(0, wInt(j.drs, 0));
    const seq = wNum(j.seq, 99999);
    /* `prodsMain`, NOT `prods`: the `Production` sheet's own numbers, never
       merged with any other sheet's (parser.js, the block by j.prodsMain).
       The owner's rule for this station from the first brief is that it reads
       `Production` only, and on 2026-09-17 the difference stopped being
       theoretical - a column inserted into `Production` left `Production (2)`
       with its headers one column away from its data, and the cross-sheet max
       gave R5053 four product groups it has not got. A job with no
       `prodsMain` at all is not on `Production`, so it is not this station's
       business and is not fed. */
    (j.prodsMain || []).forEach((p, i) => {
      if (!p || !p.n) return;
      const group = weldGroupKey(p.n);
      if (!weldAllowed(group)) return;            // a red group: never fed
      /* F and S only. T is green on the template in most groups and is still
         never shown - there are no transoms (owner, 2026-09-16). And only the
         parts this group actually has: a Super door's frames number belongs to
         somebody else and is fed as 0 rather than carried. */
      const parts = weldPartsFor(group);
      const frames = parts.indexOf("frames") >= 0 ? Math.max(0, wInt(p.f, 0)) : 0;
      const sashes = parts.indexOf("sashes") >= 0 ? Math.max(0, wInt(p.s, 0)) : 0;
      if (!(frames > 0 || sashes > 0)) return;    // nothing to weld: not fed
      const title = weldTitle(job, group);
      if (emitted[title]) {
        if (typeof console !== "undefined" && console.warn)
          console.warn("[welding] " + title + " appears twice on the sheet: the second one is " +
                       "not being fed. Title is unique on the Welding station list.");
        return;
      }
      emitted[title] = 1;
      const row = { title: title, job: job, group: group,
                    /* the group's column order on the sheet, so a card lists
                       its groups the way the sheet does */
                    groupSeq: i,
                    customer: customer, comment: comment, sentToFloor: sent,
                    wnd: wnd, drs: drs, frames: frames, sashes: sashes,
                    seq: seq, section: section, active: true };
      /* gold on the office's record means that component is welded; yellow and
         everything else mean nothing is (owner, answer 5) */
      row.seed = {
        framesDone: look(j, "prod:" + p.n + ":f") === "done" ? frames : 0,
        sashesDone: look(j, "prod:" + p.n + ":s") === "done" ? sashes : 0
      };
      out.push(row);
    });
  });
  out.sort(weldRowOrder);
  return out;
}

/** Rows sort by the office's own order, then the group's column order, then the
    Title - so one slice always makes one plan. */
function weldRowOrder(a, b) {
  const d = wNum(a.seq, 99999) - wNum(b.seq, 99999);
  if (d) return d;
  const g = wNum(a.groupSeq, 999) - wNum(b.groupSeq, 999);
  if (g) return g;
  return a.title < b.title ? -1 : a.title > b.title ? 1 : 0;
}

/** The job facts of one slice row, in list shape. Thirteen columns, and not one
    of them a phone number, an eircode, an area, a price or a transom. */
function weldFeederFields(row) {
  return { Job: row.job, Group: row.group, GroupSeq: wInt(row.groupSeq, 0),
           Customer: wTxt(row.customer), Comment: wTxt(row.comment),
           SentToFloor: wTxt(row.sentToFloor),
           Wnd: Math.max(0, wInt(row.wnd, 0)), Drs: Math.max(0, wInt(row.drs, 0)),
           Frames: Math.max(0, wInt(row.frames, 0)), Sashes: Math.max(0, wInt(row.sashes, 0)),
           Seq: wNum(row.seq, 99999), Section: wTxt(row.section),
           Active: row.active ? "Yes" : "No" };
}
/** The two counters of a slice row's seed, in list shape, clamped to their own
    quantity. */
function weldSeedFields(row) {
  const s = (row && row.seed) || {};
  return { FramesDone: wClamp(s.framesDone, Math.max(0, wInt(row && row.frames, 0))),
           SashesDone: wClamp(s.sashesDone, Math.max(0, wInt(row && row.sashes, 0))) };
}
/** Everything about one slice row that could make the feeder want to write it.
    ST.sliceHash turns this into the fingerprint a quiet afternoon is skipped on,
    so a fact missing from here is a change that would not reach the floor for
    ten minutes. */
function weldHashRow(r) {
  const seed = (r && r.seed) || {};
  return [r.title, r.job, r.group, r.groupSeq, r.customer, r.comment, r.sentToFloor,
          r.wnd, r.drs, r.frames, r.sashes, r.seq, r.section, !!r.active,
          seed.framesDone || 0, seed.sashesDone || 0];
}

/* ---- what the screens draw --------------------------------------------------
   One RECORD per list row (a job and a group), and one CARD per job built out
   of the records that share its job number. Counters are clamped for display:
   if the office shortens a group from 8 sashes to 4 after the floor has welded
   6, the line reads 4 of 4 rather than 150%. The clamp is display only -
   nothing here writes a corrected number back anywhere.                     */
const weldActive = f => wTxt(f && f.Active).trim().toLowerCase() === "yes";
/** The tablet's own filter: "In production" and nothing else, matched on the
    prefix and ignoring case, exactly as ST.inProduction reads a section name. */
const weldInProduction = f => /^\s*in production/i.test(wTxt(f && f.Section));
/** The colour of anything that has a done and a total, at all three levels.
    "" = nothing welded · "yellow" = started · "green" = every one welded.
    A line with nothing to weld has finished nothing, so it is not coloured. */
function weldColour(done, total) {
  const t = Math.max(0, wInt(total, 0));
  const d = wClamp(done, t);
  if (!(t > 0)) return "";
  if (d >= t) return "green";
  return d > 0 ? "yellow" : "";
}
/** The colour of a group or a card from the colours of what is under it. */
function weldRollUp(colours) {
  const c = (colours || []).filter(x => x !== null && x !== undefined);
  if (!c.length) return "";
  if (c.every(x => x === "green")) return "green";
  return c.some(x => x) ? "yellow" : "";
}

/** One list row as the boards read it. */
function weldRecord(it) {
  const f = (it && it.fields) || {};
  const g = { id: wTxt(it && it.id), title: wTxt(f.Title),
              job: wKey(f.Job) || wKey(wTxt(f.Title).split("|")[0]),
              group: wKey(f.Group) || wKey(wTxt(f.Title).split("|")[1]),
              groupSeq: wNum(f.GroupSeq, 999),
              customer: wTxt(f.Customer), comment: wTxt(f.Comment),
              sentToFloor: wTxt(f.SentToFloor),
              wnd: Math.max(0, wInt(f.Wnd, 0)), drs: Math.max(0, wInt(f.Drs, 0)),
              seq: wNum(f.Seq, 99999), section: wTxt(f.Section),
              active: weldActive(f), fedAt: wTxt(f.FedAt),
              /* the first tap writes this and nothing else ever does but an
                 office edit, so it is the honest answer to "has anybody moved
                 this row" - and the half of last-writer-wins the queue reads */
              doneAt: wTxt(f.DoneAt), doneBy: wTxt(f.DoneBy),
              by: {}, at: {}, lines: [] };
  let done = 0, total = 0;
  WELD_PART_KEYS.forEach(k => {
    const t = Math.max(0, wInt(f[WELD_TOTAL_FIELD[k]], 0));
    const d = wClamp(f[WELD_DONE_FIELD[k]], t);
    g[k] = d;
    g[k + "Total"] = t;
    g.by[k] = wTxt(f[WELD_BY_FIELD[k]]);
    g.at[k] = wTxt(f[WELD_AT_FIELD[k]]);
    /* a line with nothing to weld is not drawn at all - "Sashes 0 / 0" on a
       door is a row of nothing on a screen read at arm's length */
    if (t > 0) {
      g.lines.push({ part: k, label: weldPartLabel(k), done: d, total: t,
                     by: g.by[k], at: g.at[k], colour: weldColour(d, t) });
      done += d; total += t;
    }
  });
  g.done = done; g.total = total;
  g.colour = weldRollUp(g.lines.map(l => l.colour));
  g.finished = total > 0 && done >= total;
  g.left = Math.max(0, total - done);
  return g;
}

/** Every record of the list that `keep` wants, de-duplicated by Title with the
    OLDEST id winning - the same rule the phases list and the glass list use.
    Title is unique on this list and the feeder is its only writer, so a
    duplicate should not happen; if one does, the feeder's own row is the one
    both boards read and the stray is left alone rather than deleted. */
function weldRecords(items, keep) {
  const best = {};
  (items || []).forEach(it => {
    if (!it) return;
    const f = it.fields || {};
    if (keep && !keep(f)) return;
    const t = wKey(f.Title) || (wKey(f.Job) + "|" + wKey(f.Group));
    if (!t || t === "|") return;
    const prev = best[t];
    if (prev && wItemAge(prev, it) <= 0) return;
    best[t] = it;
  });
  return Object.keys(best).map(t => weldRecord(best[t]));
}

/** One card per JOB, its groups inside it in the sheet's own order. */
function weldCards(items, keep) {
  const recs = weldRecords(items, keep);
  const byJob = {};
  recs.forEach(r => {
    if (!r.job) return;
    const c = byJob[r.job] || (byJob[r.job] = {
      job: r.job, customer: r.customer, comment: r.comment, sentToFloor: r.sentToFloor,
      wnd: r.wnd, drs: r.drs, seq: r.seq, section: r.section, fedAt: r.fedAt,
      doneAt: "", doneBy: "", groups: [], done: 0, total: 0 });
    /* the job's facts are repeated on every one of its rows; the first row in
       group order is taken as the card's, and a blank never overwrites a value
       (a row written by an older feeder can be short of a column) */
    ["customer", "comment", "sentToFloor", "section", "fedAt"].forEach(k => {
      if (!c[k] && r[k]) c[k] = r[k];
    });
    if (!(c.wnd > 0) && r.wnd > 0) c.wnd = r.wnd;
    if (!(c.drs > 0) && r.drs > 0) c.drs = r.drs;
    if (r.seq < c.seq) c.seq = r.seq;
    /* the card's last touch is the newest of its rows', by the moment the
       stamp names rather than as text */
    if (r.doneAt && (!c.doneAt || weldAtCmp(r.doneAt, c.doneAt) > 0)) {
      c.doneAt = r.doneAt; c.doneBy = r.doneBy;
    }
    c.groups.push(r);
    c.done += r.done; c.total += r.total;
  });
  const out = Object.keys(byJob).map(j => {
    const c = byJob[j];
    c.groups.sort((a, b) => (wNum(a.groupSeq, 999) - wNum(b.groupSeq, 999)) ||
                            (a.group < b.group ? -1 : a.group > b.group ? 1 : 0));
    c.colour = weldRollUp(c.groups.map(g => g.colour));
    c.finished = c.total > 0 && c.done >= c.total;
    c.left = Math.max(0, c.total - c.done);
    return c;
  });
  out.sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
  return out;
}

/** Two `At` stamps compared by the MOMENT they name rather than as text - the
    same care ST.atCmp takes, and for the same reason: "…10:00:00.100Z" sorts
    BEFORE "…10:00:00Z" as a string and a second after it as a time. */
function weldAtCmp(a, b) {
  const ta = Date.parse(wTxt(a)), tb = Date.parse(wTxt(b));
  if (isFinite(ta) && isFinite(tb)) return ta < tb ? -1 : ta > tb ? 1 : 0;
  const sa = wTxt(a), sb = wTxt(b);
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/** The tablet's board: the jobs on the floor now. */
function weldBoard(items) {
  return weldCards(items, f => weldActive(f) && weldInProduction(f));
}
/** The office's board: every section, finished jobs included (owner, answer 5). */
function weldOfficeBoard(items) { return weldCards(items, weldActive); }
/** One job's card, or null - for the office's drawer line.

    ACTIVE = YES, the same as both boards. It used to take every row of the job
    whatever its Active said, which made the drawer's "Welding 9 / 16" count
    groups that are not on the sheet any more: the office would read a total
    that no board anywhere agrees with and that nobody can work off. A row goes
    Active = No when its group leaves the sheet and is never deleted, so the
    inactive ones pile up behind a job for ever. */
function weldJobCard(items, job) {
  const want = wKey(job);
  if (!want) return null;
  const mine = weldCards(items, f => weldActive(f) &&
    (wKey(f.Job) || wKey(wTxt(f.Title).split("|")[0])) === want);
  return mine[0] || null;
}

/** The cards somebody typing in the search box is looking for: a job number, a
    customer name or a group, matched anywhere. An empty box is every card. */
function weldFilter(cards, q) {
  const want = wTxt(q).trim().toLowerCase();
  if (!want) return (cards || []).slice();
  return (cards || []).filter(c =>
    (c.job + " " + c.customer + " " + c.groups.map(g => g.group).join(" ")).toLowerCase()
      .indexOf(want) >= 0);
}
/** The "Sent to floor" chip: on, it hides a card whose SentToFloor is blank.
    Off by default, so every "In production" job is shown - the same as glass
    (owner's decision 3, 2026-09-16). */
function weldSentFilter(cards, on) {
  if (!on) return (cards || []).slice();
  return (cards || []).filter(c => !!wTxt(c.sentToFloor).trim());
}
/** "N left" for the header: frames left plus sashes left over the cards given.
    It is the BOARD's number, never the searched one - somebody looking a job up
    must not make the day's work read smaller than it is. */
function weldLeft(cards) {
  return (cards || []).reduce((n, c) => n + Math.max(0, wInt(c.left, 0)), 0);
}
/** "N left" per part (frames, sashes) over the cards given - same board number
    as weldLeft, split by WELD_PART_KEYS instead of summed. A group fed only
    some parts (Super door: sashes only, per WELD_GROUP_PARTS) already carries
    0 in the part it is not fed, so it contributes 0 to that part here without
    any extra check. */
function weldLeftByPart(cards) {
  const out = {};
  WELD_PART_KEYS.forEach(k => { out[k] = 0; });
  (cards || []).forEach(c => (c.groups || []).forEach(g => {
    WELD_PART_KEYS.forEach(k => {
      out[k] += Math.max(0, wInt(g[k + "Total"], 0) - wInt(g[k], 0));
    });
  }));
  return out;
}
const weldLeftWords = n => Math.max(0, wInt(n, 0)) + " left";
/** "6 windows · 1 door" - the quantities under the job number. */
function weldQtyWords(c) {
  const bits = [];
  if (c && c.wnd > 0) bits.push("Wnd " + c.wnd);
  if (c && c.drs > 0) bits.push("Drs " + c.drs);
  return bits.join(" · ");
}

/* ---- the writes one tap makes ----------------------------------------------
   Two, in order: the counter, and then - only if that succeeded - the log line.
   The counter PATCH carries that part's number, that part's By and At, and the
   last-touch pair. Nothing else can get into it: weldFloorOnly() is the same
   filter the tablet's queue runs on the way in and on the way out, so a job
   fact typed into localStorage by somebody holding the tablet still cannot
   reach the list.

   THE OFFICE WRITES THE SAME EIGHT COLUMNS through the same two functions
   (spec 2026-09-16, "The office's welding board"). That is deliberate: there is
   no second body-builder that could carry a job fact, a section or an Active in
   on the office's path, and the office's name goes into the same By/At the
   floor's does. What the office does NOT do is write a `Station log` line -
   rule 2 is unchanged - and that is enforced by there being no call to
   ST.logFields on the office's path at all.                                 */
function weldFloorOnly(fields) {
  const out = {};
  WELD_FLOOR_FIELDS.forEach(k => {
    if (!fields || !(k in fields)) return;
    const v = fields[k];
    if (WELD_COUNTER_FIELDS.indexOf(k) >= 0) {
      if (typeof v === "number" && isFinite(v)) out[k] = v;    // a counter is a number
      return;
    }
    if (typeof v === "string" && v) out[k] = v;                // a By or an At is text
  });
  return out;
}
/** The new value of one counter after a tap. delta is a number, "all" or
    "none". The only rule is the clamp: a finished card is still tappable, and
    reducing it is what brings it back off the Finished group. */
function weldApplyTap(row, part, delta) {
  const k = wTxt(part).trim().toLowerCase();
  if (WELD_PART_KEYS.indexOf(k) < 0) return null;
  const total = Math.max(0, wInt(row && row[k + "Total"], 0));
  const now = wClamp(row && row[k], total);
  if (delta === "all") return total;
  if (delta === "none") return 0;
  const d = Number(delta);
  if (!isFinite(d)) return now;
  return wClamp(now + d, total);
}
/** The counter PATCH for one tap, from the floor or from the office. */
function weldTapFields(part, value, who, at) {
  const k = wTxt(part).trim().toLowerCase();
  if (WELD_PART_KEYS.indexOf(k) < 0) return null;
  const when = wTxt(at) || new Date().toISOString();
  const name = wTxt(who);
  const out = {};
  out[WELD_DONE_FIELD[k]] = Math.max(0, wInt(value, 0));
  out[WELD_BY_FIELD[k]] = name;
  out[WELD_AT_FIELD[k]] = when;
  out.DoneBy = name;
  out.DoneAt = when;
  return out;
}
/** The office's edit, said in its own words so the call site reads as what it
    is. Identical body to a tap's, on purpose - see the block comment above. */
const weldOfficeFields = weldTapFields;

/** One `Station log` line's worth of an entry, in the shape ST.logFields takes.
    `GlassType` keeps its column name and carries the product GROUP: the column
    is on the shared list already and nothing creates columns (rule 3). Written
    by the tablet and by nothing else - never by the office. */
function weldLogEntry(e) {
  return { job: wKey(e && e.job), station: WELD_NAME, type: wKey(e && e.group),
           stage: wTxt(e && e.part).trim().toLowerCase(),
           from: wInt(e && e.from, 0), to: wInt(e && e.to, 0),
           who: wTxt(e && e.who), at: wTxt(e && e.at) };
}
/** The `Dashboard Log` words for one office edit: "R5303 CASEMENT WINDOWS
    frames". The from/to go in noteChange's own columns. */
function weldLogWords(job, group, part) {
  return "Welding: " + wKey(job) + " " + wKey(group) + " " + wTxt(part).trim().toLowerCase();
}

/* ---- a queued tap meets somebody else's write --------------------------------
   A tap is owed as a NUMBER, and a number only means something against what the
   list said when it was made. Three answers, and they are the glass tablet's,
   word for word:

     · the counter has RISEN under the tap (the office edited it, or the feeder
       seeded the row): re-base the same movement onto the new number, so +1 on
       twelve becomes thirteen rather than one;
     · the counter has FALLEN and the row was stamped AFTER the tap was made:
       somebody said something later, so the tap is dropped and said so. This is
       the "a queued floor tap older than an office edit is dropped" of the spec;
     · anything else - a fall with no later stamp, a tie, a stamp that will not
       parse - leaves the tap alone. The floor's own statement is what the
       tablet is for.

   Pure, and separate from the tablet, so all three can be tested without a
   queue, a clock or a network.                                              */
function weldRebase(e, fields) {
  const f = fields || {};
  const k = wTxt(e && e.part).trim().toLowerCase();
  if (WELD_PART_KEYS.indexOf(k) < 0) return { action: "keep" };
  const now = Number(f[WELD_DONE_FIELD[k]]);
  const was = Number(e && e.from);
  if (!isFinite(now) || !isFinite(was)) return { action: "keep" };
  if (now < was) {
    const rowAt = Date.parse(wTxt(f.DoneAt));
    const tapAt = Date.parse(wTxt(e && e.at));
    if (!isFinite(rowAt) || !isFinite(tapAt) || rowAt <= tapAt) return { action: "keep" };
    return { action: "drop" };
  }
  if (now <= was) return { action: "keep" };
  const total = Math.max(0, wInt(f[WELD_TOTAL_FIELD[k]], 0));
  return { action: "rebase",
           value: wClamp(Math.round(now + (Number(e.value) - was)), total),
           from: Math.round(now) };
}

/** What one card is currently drawing, for the tablet's repaint diff. Anything
    not in here cannot make a card redraw. */
function weldCardSig(c) {
  return JSON.stringify([c.job, c.customer, c.comment, c.sentToFloor, c.wnd, c.drs,
    c.seq, c.finished, c.colour,
    c.groups.map(g => [g.id, g.group, g.colour,
                       WELD_PART_KEYS.map(k => [g[k], g[k + "Total"], g.by[k], g.at[k]])])]);
}

/* ---- the station definition -------------------------------------------------
   The one object station-core.js is handed so that feedPlan, sliceHash and
   floorOnly can do for this station exactly what they do for glass, with no
   glass in them and no welding in them. Adding the third station is writing one
   more of these - see docs/STATIONS.md, "Adding a station".                 */
const WELD = {
  key: "welding",
  name: WELD_NAME,
  list: WELD_LIST,
  site: WELD_SITE,
  stages: [WELD_STAGE],
  fields: WELD_FIELDS,
  feederFields: WELD_FEEDER_FIELDS,
  floorFields: WELD_FLOOR_FIELDS,
  counterFields: WELD_COUNTER_FIELDS,
  seedFields: WELD_SEED_FIELDS,
  feederWrites: WELD_FEEDER_WRITES,
  feederOf: weldFeederFields,
  seedOf: weldSeedFields,
  hashOf: weldHashRow
};

const WELDC = {
  WELD, WELD_LIST, WELD_NAME, WELD_SITE, WELD_STAGE,
  WELD_PARTS, WELD_PART_KEYS, WELD_PART_LABEL, WELD_PART_SUB,
  WELD_TOTAL_FIELD, WELD_DONE_FIELD, WELD_BY_FIELD, WELD_AT_FIELD,
  WELD_FIELDS, WELD_FEEDER_FIELDS, WELD_FLOOR_FIELDS, WELD_COUNTER_FIELDS,
  WELD_SEED_FIELDS, WELD_FEEDER_WRITES, WELD_DENY, WELD_GROUP_PARTS, weldPartsFor,
  WELD_CUSTOMER_MAX, WELD_COMMENT_MAX, WELD_PHONE_RE, WELD_DIGITS_RE, WELD_EIR_RE,
  weldPartLabel, weldGroupKey, weldAllowed, weldTitle, weldStripDigits,
  weldSlice, weldRowOrder, weldFeederFields, weldSeedFields, weldHashRow,
  weldSectionOf, weldCommentOf, weldSentOf,
  weldActive, weldInProduction, weldColour, weldRollUp,
  weldRecord, weldRecords, weldCards, weldAtCmp,
  weldBoard, weldOfficeBoard, weldJobCard, weldFilter, weldSentFilter,
  weldLeft, weldLeftByPart, weldLeftWords, weldQtyWords,
  weldFloorOnly, weldApplyTap, weldTapFields, weldOfficeFields,
  weldLogEntry, weldLogWords, weldRebase, weldCardSig,
  weldKey: wKey
};
if (typeof window !== "undefined") window.WELDC = WELDC;
else if (typeof globalThis !== "undefined") globalThis.WELDC = WELDC;
if (typeof module !== "undefined" && module.exports) module.exports = WELDC;
