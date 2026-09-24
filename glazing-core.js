/* Glazing station - the pure logic of the third floor station.

   Glazing used to be the glass station's third stage. On 2026-09-21 the owner
   took it out of there - "glazing is the last step for the whole job, not just
   glass, so it will be its own dashboard with its section" - and this is that
   dashboard's half of it (docs/specs/2026-09-21-glazing-station.md).

   ONE NUMBER PER JOB: units glazed, out of the job's WINDOWS. Not per glass
   type, not per product group, not per stage - the glazer counts finished units
   and nothing else (owner's decision 1). Doors were counted with the windows
   until 2026-09-23, when the owner said they are not glazed at this station -
   "if a job has 11 windows and 2 doors the glazing should be 11" - so the
   quantity is the windows alone and a job of nothing but doors is not fed
   (docs/specs/2026-09-23-glazing-windows-only.md). The `Doors` column stays on
   the list and is still fed as a fact; it is simply not counted.

   It reads four SharePoint lists in the `Floor stations` site - "Glazing
   station" (the work), "Station people", "Station log" and "Station comments" -
   and NO workbook of any kind. Nothing in this feature writes a cell, a colour
   or a value into the Production sheet, from either side.

   Nothing in this file touches a DOM, Graph or a workbook: it is the slice the
   feeder should send, the rows the two boards draw, the clamp on a tap and the
   five fields one write may carry. test_glazing.js exercises every one.

   It is the THIRD definition, so it borrows rather than repeats: station-core.js
   already holds feedPlan, sliceHash, floorOnly, logFields, stationComments,
   stationPeople, mergeDelta, boardDiff, stripContact, atCmp and the section
   test, and every one of them is called from here rather than copied. What is
   left below is what is genuinely this station's: which jobs, what is counted,
   and the shape of a write.

   Loaded in the browser by index.html (before app.js) and by glazing.html, and
   in Node by the tests, the same way station-core.js and welding-core.js are. */

/* ---- names ---------------------------------------------------------------- */
const GLZ_LIST = "Glazing station";
/* the word in the `Station` column of the three shared lists */
const GLZ_NAME = "Glazing";
/* the `Floor stations` site and nothing else - the site has existed since
   2026-09-17 and welding's lists are already in it (docs/STATIONS.md, step 7) */
const GLZ_SITE = "floor";
/* one stage, as welding has one: there is nothing here to divide */
const GLZ_STAGE = "glaze";

/* ---- the columns -----------------------------------------------------------
   The feeder writes Title, the job facts and FedAt/FedBy; the floor writes the
   five of GLZ_FLOOR_FIELDS and nothing else. The office's own board writes the
   same five - the dated exception in CLAUDE.md rule 3 - and never a line of
   `Station log`.

   THERE IS NO SEED, and that is a fact about the office rather than an omission:
   glazing was never an office checkpoint, so there is no record anywhere of a
   job already glazed. An untouched row starts at nought. `seedFields` is empty
   and `seedOf` answers {}, which is what feedPlan needs to plan nothing.     */
const GLZ_FEEDER_FIELDS = ["Job", "Customer", "Section", "Seq", "Active",
                           "Windows", "Doors", "Total", "Comment", "AstragalTotal"];
/* `Astragal` (2026-09-24) is the second counter: the job's ASTRAGAL units done,
   out of `AstragalTotal`. The owner asked for the count and nothing else - no
   By/At of its own; a tap on it moves the last-touch pair like any tap. */
const GLZ_FLOOR_FIELDS = ["Glazed", "GlazedBy", "GlazedAt", "DoneBy", "DoneAt", "Astragal"];
const GLZ_COUNTER_FIELDS = ["Glazed", "Astragal"];
const GLZ_SEED_FIELDS = [];
const GLZ_FIELDS = ["Title"].concat(GLZ_FEEDER_FIELDS, ["FedAt", "FedBy"], GLZ_FLOOR_FIELDS);
const GLZ_FEEDER_WRITES = ["Title"].concat(GLZ_FEEDER_FIELDS, ["FedAt", "FedBy"]);

const GLZ_CUSTOMER_MAX = 70;
const GLZ_COMMENT_MAX = 140;

/* ---- small helpers ---------------------------------------------------------
   Its own three, exactly as welding-core keeps its own: this file is loaded
   beside station-core.js and could borrow stTxt/stNum/stClamp, but a station's
   rules should not stop being readable because another station's file is not on
   the page. Everything with real behaviour in it is borrowed; these are not. */
const gTxt = v => String(v == null ? "" : v);
const gKey = v => gTxt(v).trim().toUpperCase();
function gNum(v, dflt) {
  if (v == null || gTxt(v).trim() === "") return dflt;
  const n = Number(v);
  return isFinite(n) ? n : dflt;
}
const gInt = (v, dflt) => Math.round(gNum(v, dflt == null ? 0 : dflt));
const gClamp = (n, total) => Math.max(0, Math.min(total, Math.round(gNum(n, 0))));

/** station-core.js, reached at CALL time rather than at load time, so this file
    still loads and parses on its own. It is before this one on every page and
    in every test that loads this one at all. */
function stCore() {
  return (typeof ST !== "undefined" && ST) ||
         (typeof window !== "undefined" && window.ST) || null;
}
/** Rule 3's strip, with this station's own caps. The customer name and the
    sheet's COMMENT are both free text off the sheet and both carry phone
    numbers often enough that welding has stripped them since the day it
    shipped. ST.stripContact is the one implementation.

    FAIL CLOSED, and that is rule 3's own direction: with no strip to hand no
    free text goes on the list at all, rather than an unguarded phone number. */
function glzStrip(text, max) {
  const S = stCore();
  if (!S || typeof S.stripContact !== "function") return "";
  return S.stripContact(text, gInt(max, GLZ_COMMENT_MAX));
}
/** Is this list row's Section "In production"? ST.sectionInProduction is the
    one regex, shared with the feeder's own slice test. */
function glzSectionLive(section) {
  const S = stCore();
  return S ? !!S.sectionInProduction(section) : false;
}

/* ---- what the feeder should be sending --------------------------------------
   ONE ROW PER JOB, every section the master shows, exactly as welding feeds:
   the tablet is what narrows to "In production", and the office asked to be
   able to correct a finished job. A job that has left the sheet (cat "past") is
   not in the slice at all; feedPlan marks whatever row it made Active = No and
   never deletes it.

   THE QUANTITIES COME OFF `Production` AND NO OTHER SHEET (owner's standing
   rule of 2026-09-18, and HISTORY B20 is what it cost to learn). `j.wndMain`
   and `j.drsMain` are the parser's Production-only pair, beside `j.prodsMain`
   which welding reads; `j.wnd`/`j.drs` take the first sheet that has a number
   and are deliberately not read here.                                        */
function glzWindowsOf(j) { return Math.max(0, gInt(j && j.wndMain, 0)); }
function glzDoorsOf(j) { return Math.max(0, gInt(j && j.drsMain, 0)); }
/** ASTRAGAL off `Production` (2026-09-24): its own component, no relation to
    the windows or the doors - the number of astragal units the job needs. */
function glzAstragalOf(j) { return Math.max(0, gInt(j && j.astrMain, 0)); }
/** The job's COMMENT, as the parser left it on j.notes, stripped. */
function glzCommentOf(j) {
  const hit = ((j && j.notes) || []).find(n => n && n.k === "comment");
  return glzStrip(hit ? hit.t : "");
}
function glzSectionOf(j, blockNames) {
  return gTxt((blockNames || [])[j && j.blk] || "").trim();
}

function glzSlice(jobs, blockNames) {
  const names = blockNames || (jobs && jobs.blockNames) || [];
  const out = [], emitted = {};
  (jobs || []).forEach(j => {
    if (!j || !j.id) return;
    if (j.cat === "past") return;                  // not on the sheet any more
    const job = gKey(j.id);
    if (!job) return;
    const wnd = glzWindowsOf(j), drs = glzDoorsOf(j);
    /* the WINDOWS alone (owner, 2026-09-23): doors are not glazed here. The
       doors go on the row as a fact and are not added to the count, so a job of
       nothing but doors falls out on the next line and the feeder's own plan
       marks whatever row it made before that day Active = No. */
    const total = wnd;
    const astr = glzAstragalOf(j);
    /* windows OR astragal puts a job on the board: a job of doors and astragal
       and no windows is fed with the astragal line alone (owner, 2026-09-24) */
    if (!(total > 0) && !(astr > 0)) return;       // nothing to glaze: not fed
    /* one Title, one row. Title is unique on the list, so a slice carrying a
       job twice would make the feeder POST a row SharePoint then refuses, every
       run, for ever. The first wins and the second is said out loud. */
    if (emitted[job]) {
      if (typeof console !== "undefined" && console.warn)
        console.warn("[glazing] " + job + " appears twice on the sheet: the second one is not " +
                     "being fed. Title is unique on the Glazing station list.");
      return;
    }
    emitted[job] = 1;
    out.push({ title: job, job: job,
               customer: glzStrip(j.cust, GLZ_CUSTOMER_MAX),
               comment: glzCommentOf(j),
               wnd: wnd, drs: drs, total: total, astr: astr,
               seq: gNum(j.seq, 99999), section: glzSectionOf(j, names), active: true });
  });
  out.sort(glzRowOrder);
  return out;
}
/** The office's own order, then the Title - so one slice always makes one plan. */
function glzRowOrder(a, b) {
  const d = gNum(a.seq, 99999) - gNum(b.seq, 99999);
  if (d) return d;
  return a.title < b.title ? -1 : a.title > b.title ? 1 : 0;
}
/** The job facts of one slice row, in list shape. Nine columns, and not one of
    them a phone number, an eircode, an area, a price or a glass type. */
function glzFeederFields(row) {
  return { Job: row.job, Customer: gTxt(row.customer), Section: gTxt(row.section),
           Seq: gNum(row.seq, 99999), Active: row.active ? "Yes" : "No",
           Windows: Math.max(0, gInt(row.wnd, 0)), Doors: Math.max(0, gInt(row.drs, 0)),
           Total: Math.max(0, gInt(row.total, 0)), Comment: gTxt(row.comment),
           AstragalTotal: Math.max(0, gInt(row.astr, 0)) };
}
/** No seed, ever: there is no office record of glazing to seed from. */
function glzSeedFields() { return {}; }
/** Everything about a slice row that could make the feeder want to write it. */
function glzHashRow(r) {
  return [r.title, r.job, r.customer, r.comment, r.wnd, r.drs, r.total,
          r.seq, r.section, !!r.active, r.astr];
}

/* ---- what the screens draw -------------------------------------------------
   One CARD per job, because there is one list row per job. The counter is
   clamped for display: if the office shortens a job from 8 units to 4 after the
   glazer has done 6, the card reads 4 of 4 rather than 150%. Display only -
   nothing here writes a corrected number back anywhere.                      */
const glzActive = f => gTxt(f && f.Active).trim().toLowerCase() === "yes";
const glzInProduction = f => glzSectionLive(f && f.Section);
/** "" nothing glazed · "yellow" started · "green" every unit glazed. A job with
    nothing to glaze has finished nothing, so it is not coloured. */
function glzColour(done, total) {
  const t = Math.max(0, gInt(total, 0));
  const d = gClamp(done, t);
  if (!(t > 0)) return "";
  return d >= t ? "green" : d > 0 ? "yellow" : "";
}

/** The OFFICE board's colour, which is a different sentence from the tablet's
    (owner, 2026-09-22: jobs whose whole `Production` row was already gold showed
    no colour at all). `rowDone` is the parser's own `j.done` - nothing new is
    read off the sheet - and it speaks first: a gold row is finished work
    whatever this list's counter says, including a job with nothing to glaze. A
    full count is gold too, because the floor has finished and the office moves
    the row to its next section later. Display only; this writes nothing and
    `glzColour` (green, the tablet's) is deliberately left alone. */
function glzOfficeColour(c, rowDone) {
  if (rowDone) return "gold";
  const col = glzCardColour(c);
  return col === "green" ? "gold" : col;
}
/** One card's colour over BOTH counts (2026-09-24): green only when the windows
    and the astragal are each full, yellow once either has started. */
function glzCardColour(c) {
  const t = Math.max(0, gInt(c && c.total, 0)), at = Math.max(0, gInt(c && c.astrTotal, 0));
  if (!(t > 0) && !(at > 0)) return "";
  const d = gClamp(c && c.glazed, t), ad = gClamp(c && c.astr, at);
  if (d >= t && ad >= at) return "green";
  return (d > 0 || ad > 0) ? "yellow" : "";
}

/** One list row as the boards read it. */
function glzRecord(it) {
  const f = (it && it.fields) || {};
  const total = Math.max(0, gInt(f.Total, 0));
  const done = gClamp(f.Glazed, total);
  const astrTotal = Math.max(0, gInt(f.AstragalTotal, 0));
  const astr = gClamp(f.Astragal, astrTotal);
  const rec = { id: gTxt(it && it.id), title: gTxt(f.Title),
           job: gKey(f.Job) || gKey(f.Title),
           customer: gTxt(f.Customer), comment: gTxt(f.Comment),
           wnd: Math.max(0, gInt(f.Windows, 0)), drs: Math.max(0, gInt(f.Doors, 0)),
           seq: gNum(f.Seq, 99999), section: gTxt(f.Section),
           active: glzActive(f), fedAt: gTxt(f.FedAt),
           total: total, glazed: done, left: Math.max(0, total - done),
           by: gTxt(f.GlazedBy), at: gTxt(f.GlazedAt),
           /* the first tap writes these and nothing else does but an office
              edit, so they are the honest answer to "has this row moved" */
           doneBy: gTxt(f.DoneBy), doneAt: gTxt(f.DoneAt),
           astrTotal: astrTotal, astr: astr, astrLeft: Math.max(0, astrTotal - astr) };
  rec.colour = glzCardColour(rec);
  /* finished = every line full: windows AND astragal (owner, 2026-09-24) */
  rec.finished = rec.colour === "green";
  return rec;
}

/** Every card `keep` wants, de-duplicated by Title with the OLDEST id winning -
    the same rule every other list here uses. Title is unique and the feeder is
    its only writer, so a duplicate should not happen; if one does, the feeder's
    own row is what both boards read and the stray is left alone, never
    deleted. */
function glzCards(items, keep) {
  const age = (a, b) => {
    const na = Number(a && a.id), nb = Number(b && b.id);
    if (isFinite(na) && isFinite(nb)) return na - nb;
    const sa = gTxt(a && a.id), sb = gTxt(b && b.id);
    return sa < sb ? -1 : sa > sb ? 1 : 0;
  };
  const best = {};
  (items || []).forEach(it => {
    if (!it) return;
    const f = it.fields || {};
    if (keep && !keep(f)) return;
    const t = gKey(f.Title) || gKey(f.Job);
    if (!t) return;
    const prev = best[t];
    if (prev && age(prev, it) <= 0) return;
    best[t] = it;
  });
  const out = Object.keys(best).map(t => glzRecord(best[t]));
  out.sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
  return out;
}
/** The tablet's board: the jobs on the floor now. */
function glzBoard(items) {
  return glzCards(items, f => glzActive(f) && glzInProduction(f));
}
/** The office's board: every section, finished jobs included. */
function glzOfficeBoard(items) { return glzCards(items, glzActive); }
/** One job's card, or null - for the office's drawer line. Active rows only,
    the same as both boards, so the drawer can never read a total no board
    anywhere agrees with. */
function glzJobCard(items, job) {
  const want = gKey(job);
  if (!want) return null;
  const mine = glzCards(items, f => glzActive(f) && (gKey(f.Job) || gKey(f.Title)) === want);
  return mine[0] || null;
}

/** The cards somebody typing in the search box is looking for: a job number or
    a customer name, matched anywhere. An empty box is every card. */
function glzFilter(cards, q) {
  const want = gTxt(q).trim().toLowerCase();
  if (!want) return (cards || []).slice();
  return (cards || []).filter(c => (c.job + " " + c.customer).toLowerCase().indexOf(want) >= 0);
}
/** The tablet's two tabs (owner, 2026-09-24; replaced the collapsed Finished
    group), as welding's weldTabs. ONE pass over the office's board: On floor =
    In production and not finished; Finished = every other Active card - a
    finished In production job AND every job in any other section still on the
    sheet. Active = No is on neither. Each keeps glzCards' order. */
function glzTabs(items) {
  const out = { floor: [], finished: [] };
  glzOfficeBoard(items).forEach(c => {
    out[glzSectionLive(c.section) && !c.finished ? "floor" : "finished"].push(c);
  });
  return out;
}
/** Which tab a search shows (owner, 2026-09-24, decision 3). The tab you are on
    keeps the screen while it has a match; with none there and a match in the
    other tab, the other tab is shown. `other` is how many matches the tab NOT
    shown holds, for the tappable "N more" line. An empty box changes nothing. */
function glzSearchTab(tabs, q, current) {
  const cur = current === "finished" ? "finished" : "floor";
  const alt = cur === "floor" ? "finished" : "floor";
  if (!gTxt(q).trim()) return { tab: cur, other: 0 };
  const here = glzFilter(tabs && tabs[cur], q).length;
  const there = glzFilter(tabs && tabs[alt], q).length;
  if (!here && there) return { tab: alt, other: 0 };
  return { tab: cur, other: there };
}
/** "N left" over the cards given. The BOARD's number, never the searched one -
    somebody looking a job up must not make the day's work read smaller. */
function glzLeft(cards, key) {
  const k = key || "left";
  return (cards || []).reduce((n, c) => n + Math.max(0, gInt(c[k], 0)), 0);
}
/** The tablet header: both counts, windows first (owner, 2026-09-24). */
function glzHeadWords(cards) {
  return glzLeft(cards) + " windows left · " + glzLeft(cards, "astrLeft") + " astragal left";
}
const glzLeftWords = n => Math.max(0, gInt(n, 0)) + " left";
/** "6 windows" at the head of the tablet's card. It said "6 units" until
    2026-09-23; now that the count is the windows and nothing else, the word for
    what is being counted is the honest one to print. */
function glzUnitWords(n) {
  const t = Math.max(0, gInt(n, 0));
  return t + (t === 1 ? " window" : " windows");
}
/** The windows and nothing else. The doors are on the row as a fact but they
    are not glazed here (owner, 2026-09-23), so no screen says a door out loud -
    a number the glazer cannot act on would only make the card read wrong. */
function glzQtyWords(c) {
  const w = Math.max(0, gInt(c && c.wnd, 0));
  const a = Math.max(0, gInt(c && c.astrTotal, 0));
  return [w > 0 ? w + (w === 1 ? " window" : " windows") : "",
          a > 0 ? a + " astragal" : ""].filter(Boolean).join(" · ");
}
/** Which counter a part names: "astragal" is `Astragal`, anything else the
    windows' `Glazed` - with the card keys and the total column that go with it. */
const GLZ_PARTS = {
  windows: { field: "Glazed", done: "glazed", total: "total", totalField: "Total" },
  astragal: { field: "Astragal", done: "astr", total: "astrTotal", totalField: "AstragalTotal" }
};
const glzPart = part => GLZ_PARTS[part === "astragal" ? "astragal" : "windows"];

/* ---- the writes one tap makes ----------------------------------------------
   Two, in order: the counter, and then - only if that succeeded - the log line.
   The PATCH carries the counter, its By and At, and the last-touch pair.
   Nothing else can get into it: ST.floorOnly(fields, GLAZE) is the same filter
   the tablet's queue runs on the way in and on the way out, so a job fact typed
   into localStorage by somebody holding the tablet still cannot reach the list.

   THE OFFICE WRITES THE SAME FIVE COLUMNS through the same builder (spec
   2026-09-21 section D). That is deliberate: there is no second body-builder
   that could carry a job fact, a Section or an Active in on the office's path.
   What the office does NOT do is write a `Station log` line, and that is
   enforced by there being no call to ST.logFields on the office's path.      */
/** The new value after a tap. delta is a number, "all" or "none". The only rule
    is the clamp: a finished card is still tappable, and reducing it is what
    brings it back off the Finished group. */
function glzApplyTap(row, delta, part) {
  const P = glzPart(part);
  const total = Math.max(0, gInt(row && row[P.total], 0));
  const now = gClamp(row && row[P.done], total);
  if (delta === "all") return total;
  if (delta === "none") return 0;
  const d = Number(delta);
  if (!isFinite(d)) return now;
  return gClamp(now + d, total);
}
/** The counter PATCH for one tap, from the floor or from the office. */
function glzTapFields(value, who, at, part) {
  const when = gTxt(at) || new Date().toISOString();
  const name = gTxt(who);
  const n = Math.max(0, gInt(value, 0));
  /* astragal: its count and the last-touch pair, no By/At of its own */
  if (part === "astragal") return { Astragal: n, DoneBy: name, DoneAt: when };
  return { Glazed: n, GlazedBy: name, GlazedAt: when, DoneBy: name, DoneAt: when };
}
/** The office's edit, said in its own words so the call site reads as what it
    is. Identical body to a tap's, on purpose - see the block comment above. */
const glzOfficeFields = glzTapFields;
/** The five fields and nothing else, through station-core's one filter. */
function glzFloorOnly(fields) {
  const S = stCore();
  return S ? S.floorOnly(fields, GLAZE) : {};
}

/** One `Station log` line's worth of an entry, in the shape ST.logFields takes.
    `GlassType` keeps its column name and carries this station's own word: the
    column is on the shared list already and nothing creates columns. Written by
    the tablet and by nothing else - never by the office. */
function glzLogEntry(e) {
  /* an astragal line says so in Stage, so the report's glazing days (which
     count Stage = glaze) never add astragal units to window units */
  const astr = !!(e && e.part === "astragal");
  return { job: gKey(e && e.job), station: GLZ_NAME, type: astr ? "ASTRAGAL" : GLZ_NAME.toUpperCase(),
           stage: astr ? "astragal" : GLZ_STAGE,
           from: gInt(e && e.from, 0), to: gInt(e && e.to, 0),
           who: gTxt(e && e.who), at: gTxt(e && e.at) };
}
/** The `Dashboard Log` words for one office edit. The from/to go in
    noteChange's own columns. */
const glzLogWords = (job, part) => (part === "astragal" ? "Glazing astragal: " : "Glazing: ") + gKey(job);

/* ---- a queued tap meets somebody else's write --------------------------------
   A tap is owed as a NUMBER, and a number only means something against what the
   list said when it was made. Three answers, and they are the other two
   tablets', word for word:

     · the counter has RISEN under the tap (the office edited it): re-base the
       same movement onto the new number, so +1 on twelve becomes thirteen;
     · the counter has FALLEN and the row was stamped AFTER the tap was made:
       somebody said something later, so the tap is dropped and said so;
     · anything else - a fall with no later stamp, a tie, a stamp that will not
       parse - leaves the tap alone. The floor's own statement is what the
       tablet is for.                                                          */
function glzRebase(e, fields) {
  const f = fields || {};
  const P = glzPart(e && e.part);
  const now = Number(f[P.field]);
  const was = Number(e && e.from);
  if (!isFinite(now) || !isFinite(was)) return { action: "keep" };
  if (now < was) {
    const rowAt = Date.parse(gTxt(f.DoneAt));
    const tapAt = Date.parse(gTxt(e && e.at));
    if (!isFinite(rowAt) || !isFinite(tapAt) || rowAt <= tapAt) return { action: "keep" };
    /* one DoneAt now serves two counters (2026-09-24): a later stamp that is
       this same person's own tap on the OTHER line is not somebody else
       saying something later, so it cannot drop this tap (review finding 1) */
    if (gTxt(f.DoneBy) && gTxt(f.DoneBy) === gTxt(e && e.who)) return { action: "keep" };
    return { action: "drop" };
  }
  if (now <= was) return { action: "keep" };
  const total = Math.max(0, gInt(f[P.totalField], 0));
  return { action: "rebase",
           value: gClamp(Math.round(now + (Number(e.value) - was)), total),
           from: Math.round(now) };
}

/** What one card is currently drawing, for the tablet's repaint diff. Anything
    not in here cannot make a card redraw. */
function glzCardSig(c) {
  return JSON.stringify([c.job, c.customer, c.comment, c.section, c.wnd, c.drs, c.total,
                         c.seq, c.glazed, c.by, c.at, c.finished, c.colour, c.astr, c.astrTotal]);
}

/** THE ADAPTER the station report asks this station for. The report itself
    (export.js, stationReport) has no station in it: the Jobs sheet is whatever
    the station's own definition says its board rows are, and this is glazing's
    answer - one row per job. `jobs` is the same rows read as per-job progress,
    which is all the Summary needs.

    `stage` is unused here - this station has one - and is taken anyway so the
    three adapters have the same shape. The CUSTOMER goes through the strip
    again on the way out (rule 3, as the glass adapter does): the feeder already
    stripped it, and an export is the one place where being sure twice costs
    nothing. */
function glzReportJobs(data, stage) {
  const S = stCore();
  const cap = (S && S.REPORT_CUSTOMER_MAX) || 60;
  const cards = (data && data.board) || [];
  const rows = [], jobs = [];
  cards.forEach(c => {
    /* complete = every line full, the same test as the Complete column - so
       an astragal-only job can complete and 11/11 windows with astragal owed
       does not (review finding 3) */
    jobs.push({ job: c.job, done: c.finished ? 1 : 0, total: 1 });
    rows.push([c.job, glzStrip(c.customer, cap), c.section, c.wnd, c.drs,
               c.total, c.glazed, c.left, c.astrTotal, c.astr, c.doneBy, c.doneAt,
               c.finished ? "Yes" : "No"]);
  });
  /* `Doors (not glazed)` says in the header what the numbers would otherwise
     make somebody work out: the doors are a fact off the sheet and are not in
     Total, which is the windows (owner, 2026-09-23). */
  return { columns: ["Job", "Customer", "Section", "Windows", "Doors (not glazed)", "Total",
                     "Glazed", "Left", "Astragal", "Astragal done", "Last moved by", "When", "Complete"],
           rows: rows, jobs: jobs };
}

/* ---- the station definition -------------------------------------------------
   The one object station-core.js is handed so that feedPlan, sliceHash and
   floorOnly do for this station exactly what they do for the other two, with no
   glazing in them. docs/STATIONS.md, "Adding a station", is the field list.

   NO `daySheets`: the glazer has no paper end-of-day sheet (spec 2026-09-21).
   One line of definition switches one on, and nothing else.                  */
const GLAZE = {
  key: "glazing",
  name: GLZ_NAME,
  list: GLZ_LIST,
  site: GLZ_SITE,
  stages: [GLZ_STAGE],
  stageLabel: () => "Glazing",
  reportStages: [GLZ_STAGE],
  /* the tablet logs `Stage = glaze`, so the report stage and the log stage are
     the same word here - unlike welding, whose tablet logs one line per part */
  reportLogStages: () => [GLZ_STAGE],
  reportJobs: glzReportJobs,
  fields: GLZ_FIELDS,
  feederFields: GLZ_FEEDER_FIELDS,
  floorFields: GLZ_FLOOR_FIELDS,
  counterFields: GLZ_COUNTER_FIELDS,
  seedFields: GLZ_SEED_FIELDS,
  feederWrites: GLZ_FEEDER_WRITES,
  feederOf: glzFeederFields,
  seedOf: glzSeedFields,
  hashOf: glzHashRow,
  rowOrder: glzRowOrder
};

const GLZC = {
  GLAZE, GLZ_LIST, GLZ_NAME, GLZ_SITE, GLZ_STAGE,
  GLZ_FIELDS, GLZ_FEEDER_FIELDS, GLZ_FLOOR_FIELDS, GLZ_COUNTER_FIELDS,
  GLZ_SEED_FIELDS, GLZ_FEEDER_WRITES, GLZ_CUSTOMER_MAX, GLZ_COMMENT_MAX,
  glzStrip, glzSlice, glzRowOrder, glzFeederFields, glzSeedFields, glzHashRow,
  glzWindowsOf, glzDoorsOf, glzAstragalOf, glzCommentOf, glzSectionOf,
  glzActive, glzInProduction, glzColour, glzOfficeColour, glzCardColour, glzPart, glzHeadWords,
  glzRecord, glzCards, glzBoard, glzOfficeBoard, glzJobCard, glzFilter, glzTabs, glzSearchTab,
  glzLeft, glzLeftWords, glzUnitWords, glzQtyWords,
  glzApplyTap, glzTapFields, glzOfficeFields, glzFloorOnly,
  glzLogEntry, glzLogWords, glzRebase, glzCardSig, glzReportJobs,
  glzKey: gKey
};
if (typeof window !== "undefined") window.GLZC = GLZC;
else if (typeof globalThis !== "undefined") globalThis.GLZC = GLZC;
if (typeof module !== "undefined" && module.exports) module.exports = GLZC;
