/* Checkpoints - ticking a job's work off item by item.
   Everything here is pure logic: what a checkpoint item's status IS, the
   colour choice, the burst coalescer and the two write sequences. It is kept
   out of app.js because none of it needs a DOM, so test_checkpoints.js can
   load and exercise it directly.
   Loaded after graph.js (it calls CW) and before app.js.

   CHANGED 2026-09-11 (docs/specs/2026-09-11-status-list-is-truth.md, step 2).
   Checkpoint status used to be READ OUT OF THE EXCEL COLOUR - a copy of the
   workbook about 36 s behind the API - with a stored count from the
   `Dashboard Progress` sheet filling in the number. Every mechanism in this
   app that argued with a stale copy existed because of that, and none of them
   could make a copy true.

   Status now lives in one SharePoint list, `Dashboard progress`, one row per
   JOB|ITEM, and BOTH screens read it directly. The Excel colour is written
   FROM it so people can see it in Excel, and is never read back to decide
   anything. So:

     - itemState / cpStatus answer from the list (cpRow), never from j.cp;
     - j.cp still carries the PARSED colours, under the same name, because
       other things still want them: the phase pipeline's Cut green, the glass
       colour writer's idempotence test, the one-time import and the safeguard
       that adopts a cell somebody painted by hand in Excel. They are reached
       through cpFileStatus(), which is the ONLY way the parsed colour may be
       read, and never through cpStatus();
     - there is no hold and nothing to reconcile: the list answers in about a
       second, the map is updated at the click, and the next delta confirms
       it.                                                                   */

/* The sheet carries only the colour. White is written, never cleared: the
   Production cells already have an explicit white fill, and fill/clear would
   take the fill away instead of setting it back to white. */
const WHITE_HEX = "#FFFFFF", YELLOW_HEX = "#FFFF00", GOLD_HEX = "#FFE699";
const CP_DEBOUNCE_MS = 800;               // a run of + taps settles into one write
const CP_PROD_SHEET = "Production";
const CP_SUB = { f: "frames", s: "sashes", t: "transoms" };

/* ---- the record: the "Dashboard progress" SharePoint list -----------------
   One row per JOB|ITEM. Title is `JOB|ITEM` and is the dedupe key, exactly as
   `Dashboard phases` uses Title. Read in full at boot and by delta every ten
   seconds; written by a click, which updates this map at once so the row moves
   before the request has even gone out.

   The list is made by hand in SharePoint. When it is not there the dashboard
   says so and checkpoints are read-only - nothing is written anywhere. */
const CP_LIST = "Dashboard progress";
const CP_LIST_FIELDS = ["Title", "Job", "Item", "Done", "Total", "Status", "Who", "When", "Source"];
const cpTitle = (job, item) =>
  String(job == null ? "" : job).trim().toUpperCase() + "|" + String(item == null ? "" : item).trim();

let CPROWS = {};              // "JOB|ITEM" -> { id, job, item, done, total, status, who, when, source }
function cpRowsSet(map) { CPROWS = map || {}; }
function cpRowsAll() { return CPROWS; }
/** What the list says about one item of one job, or null. THE answer. */
function cpRow(job, item) { return CPROWS[cpTitle(job, item)] || null; }
/** Put one row in (or take it out, with null). This is what a click does
    before it writes: the map IS the optimistic state, and the next delta
    merely confirms it. There is no expiry and nothing to unmask. */
function cpRowPut(job, item, row) {
  const k = cpTitle(job, item);
  if (row == null) delete CPROWS[k]; else CPROWS[k] = row;
  return row || null;
}
const CP_WORDS = { done: "done", process: "process", "": "" };
/** The list's items, as the map. A Status the list does not know reads as
    blank; Done and Total as numbers or nought. Two rows for one Title (a list
    without the unique rule, or two browsers at the same second) resolve on
    When, newest wins - the same rule readPhases uses on SetAt. */
function cpRowsFrom(items) {
  const out = {};
  (items || []).forEach(it => {
    const f = (it && it.fields) || {};
    const title = String(f.Title == null ? "" : f.Title).trim();
    let job = String(f.Job == null ? "" : f.Job).trim().toUpperCase();
    let item = String(f.Item == null ? "" : f.Item).trim();
    if (!job || !item) {
      const i = title.indexOf("|");
      if (i > 0) { job = job || title.slice(0, i).trim().toUpperCase(); item = item || title.slice(i + 1).trim(); }
    }
    if (!job || !item) return;                     // a half-filled row is not a status
    const row = { id: it && it.id != null ? String(it.id) : null, job: job, item: item,
                  done: Number(f.Done) || 0, total: Number(f.Total) || 0,
                  status: CP_WORDS[String(f.Status == null ? "" : f.Status).trim().toLowerCase()] || "",
                  who: String(f.Who == null ? "" : f.Who),
                  when: String(f.When == null ? "" : f.When),
                  source: String(f.Source == null ? "" : f.Source).trim().toLowerCase() };
    const k = cpTitle(job, item);
    const had = out[k];
    if (had && String(had.when || "") > String(row.when || "")) return;
    out[k] = row;
  });
  return out;
}
/** The list item body for one row. listUpsert adds Title itself; a PATCH of a
    known id carries it here, so the two paths write exactly the same fields. */
function cpRowFields(job, item, done, total, status, who, when, source) {
  return { Title: cpTitle(job, item), Job: String(job).trim().toUpperCase(), Item: String(item).trim(),
           Done: Math.max(0, Math.round(Number(done) || 0)),
           Total: Math.max(0, Math.round(Number(total) || 0)),
           Status: CP_WORDS[String(status || "")] || "",
           Who: String(who == null ? "" : who), When: String(when == null ? "" : when),
           Source: String(source == null ? "" : source) };
}

/* Exact counts from the Dashboard Progress SHEET: job -> item -> {done,total,who,when}.
   NOT status, and no longer written by anything: since 2026-09-11 the sheet is
   left exactly as it stands. It is still parsed out of the workbook this
   dashboard has already downloaded, for two readers and no others - the
   one-time import, which recovers the exact count behind a yellow cell, and
   the glass colour writer's office stamp, which step 3 moves to the list with
   the rest of that feature. */
let PROGRESS = {};
function cpSetProgress(map) { PROGRESS = map || {}; }
function cpStored(job, item) { return (PROGRESS[job] || {})[item] || null; }

/* ---- the switch-on window (review finding M1) -----------------------------
   The one-time import takes fifteen to twenty-five minutes on the owner's real
   sheet, because it is lazy and capped at sixty rows a load. Until it has
   drained, most items have no row - and an item with no row reads as "nothing
   done", which was quietly catastrophic: every job collapsed to In office in
   the phase pipeline, `glassCounts` said nothing was ticked so the feeder
   wrote `OfficeDone: "No"` across the floor's list and then "Yes" again as
   rows landed (jobs unlocking and re-locking on the tablet), the floor's seeds
   went to nought, and a job marked ready to deliver printed "all checkpoints
   are complete" over lines reading 0.

   So while the import is outstanding, an item WITH NO ROW answers from the
   sheet's own colour - exactly what it answered before the switch-over, so
   nothing on any screen changes and no feeder write flaps. An item that HAS a
   row always answers from the row, so an un-tick made during the window is
   still absolute. The flag is set by app.js (cpImportCheck) and is false the
   moment the import has drained, after which this is never consulted again. */
let CP_IMPORT_PENDING = false;
function cpSetImportPending(v) { CP_IMPORT_PENDING = !!v; }
function cpImportPending() { return CP_IMPORT_PENDING; }
/** What the sheet's own colour says about an item, in itemState's own shape.
    Only ever reached while the import is outstanding and the item has no row. */
function cpFileState(j, item, total) {
  const f = cpFileStatus(j, item);
  if (f === "done") return { done: total, total: total, status: "done" };
  if (f === "process") {
    const st = cpStored(String((j && j.id) || "").toUpperCase(), item);
    const n = st ? Math.round(Number(st.done) || 0) : 0;
    return { done: (n > 0 && n < total) ? n : null, total: total, status: "process" };
  }
  return { done: 0, total: total, status: "" };
}

/* ---- item keys: "win", "drs", "glass:<type>", "prod:<name>:<f|s|t>" ----
   Product names come from the sheet's own headers, normalised to lower case
   words, so they can never contain a colon and the split is unambiguous. */
function cpTotal(j, item) {
  const p = String(item).split(":");
  if (item === "win") return j.wnd || 0;
  if (item === "drs") return j.drs || 0;
  if (p[0] === "glass") return (j.glass || {})[p[1]] || 0;
  if (p[0] === "prod") { const x = (j.prods || []).find(q => q.n === p[1]); return x ? (x[p[2]] || 0) : 0; }
  return 0;
}
/** What the PARSED Excel colour says about an item: "done" / "process" /
    "cut" / "". This is a copy of the workbook about 36 s behind the API and it
    is NOT status any more. Exactly three things may read it, and each of them
    reads it as a colour rather than as a fact:

      - the phase pipeline, for the sheet's own Cut green, which nothing in the
        dashboard ever writes and the list therefore never carries;
      - the one-time import, which adopts today's colours once;
      - the safeguard, which spots a cell somebody painted in Excel by hand.

    Nothing in the checkpoint path may call it. cpStatus is the answer. */
function cpFileStatus(j, item) {
  const cp = (j && j.cp) || {}, p = String(item).split(":");
  if (item === "win") return cp.win || "";
  if (item === "drs") return cp.drs || "";
  if (p[0] === "glass") return (cp.glass || {})[p[1]] || "";
  if (p[0] === "prod") return ((cp.prod || {})[p[1]] || {})[p[2]] || "";
  return "";
}
/** The status of one item of one job: the list's word, and nothing else. */
function cpStatus(j, item) {
  const st = itemState(j, item);
  return st ? st.status : "";
}
/** The column on the Production sheet for an item, from mapSheet(). 0 = not on the sheet. */
function cpColumn(item, map) {
  if (!map) return 0;
  const p = String(item).split(":");
  if (item === "win") return (map.qty || {}).wnd || 0;
  if (item === "drs") return (map.qty || {}).drs || 0;
  if (p[0] === "glass") return (map.glass || {})[p[1]] || 0;
  if (p[0] === "prod") return ((map.prod || {})[p[1]] || {})[p[2]] || 0;
  return 0;
}
/** The log's name for an item: "Windows", "Glass TG", "7000 CASEMENT frames". */
function cpLabel(item) {
  const p = String(item).split(":");
  if (item === "win") return "Windows";
  if (item === "drs") return "Doors";
  if (p[0] === "glass") return "Glass " + p[1].toUpperCase();
  if (p[0] === "prod") return p[1].toUpperCase() + " " + (CP_SUB[p[2]] || p[2]);
  return String(item);
}

/** Every countable thing on a job, in drawer order. Only totals > 0 exist. */
function cpItems(j) {
  const out = [];
  const add = (key, label, total, group, groupLabel) => { if (total > 0) out.push({ key, label, total, group, groupLabel }); };
  add("win", "Windows", j.wnd || 0, "win", "Windows");
  add("drs", "Doors", j.drs || 0, "drs", "Doors");
  Object.keys(j.glass || {}).forEach(k => add("glass:" + k, k.toUpperCase(), j.glass[k], "glass", "Glass"));
  (j.prods || []).forEach(p => ["f", "s", "t"].forEach(s =>
    add("prod:" + p.n + ":" + s, CP_SUB[s], p[s], "prod:" + p.n, p.n)));
  return out;
}

/* ---- the phase pipeline ----
   Where a job has got to, in one word, worked out from what the sheet already
   says: the "sent to floor" date and the checkpoint colours. Nothing sets a
   phase by hand, nothing is stored and nothing is written - it is a reading of
   the sheet, not a new field on it. The highest phase that applies wins. */
const PHASES = ["In office", "Sent to floor", "Cutting", "In fabrication",
                "In glazing", "Quality check", "Fitted / delivered"];

function jobPhase(j) {
  if (!j) return 0;
  /* gold, or gone from the Production sheet altogether: it has left the floor */
  if (j.done || j.cat === "past") return 6;
  const items = cpItems(j);                       // only things with a total > 0
  const st = x => cpStatus(j, x.key);
  const prod = items.filter(x => x.group.indexOf("prod:") === 0);
  const rest = items.filter(x => x.group.indexOf("prod:") !== 0);
  const allDone = a => a.length > 0 && a.every(x => st(x) === "done");
  if (allDone(items)) return 5;                   // everything ticked, not yet marked ready
  if (allDone(prod) && !allDone(rest)) return 4;  // frames done, the glass/windows/doors are not
  if (prod.some(x => st(x) === "process") ||
      cpStatus(j, "win") === "process" || cpStatus(j, "drs") === "process") return 3;
  /* the sheet's own Cut green: a colour nothing in the dashboard writes, so
     the list never carries it and this one reading stays on the file */
  if (prod.some(x => cpFileStatus(j, x.key) === "cut")) return 2;
  if ((j.dates && j.dates.floor) || items.some(x => st(x) === "process" || st(x) === "done")) return 1;
  return 0;
}
/* A phase can also be set by hand. That does not live here and it is never in
   the workbook - it comes from a SharePoint list that app.js reads - so this
   file only keeps the hook. When one is registered, every reading of the phase
   goes through effectivePhase(): the sheet's own evidence and the hand-set
   value, whichever is further on. The sheet therefore wins the moment it
   catches up, and clearing the hand-set phase drops straight back to it. */
let phaseHook = null;
function setPhaseHook(fn) { phaseHook = typeof fn === "function" ? fn : null; }
function effectivePhase(j) {
  const sheet = jobPhase(j);
  const hand = phaseHook ? phaseHook(j) : null;
  return hand == null ? sheet : Math.max(sheet, hand);
}
const phaseName = j => PHASES[effectivePhase(j)];

/* ---- colours ---- */
const cpStatusFor = (done, total) => done <= 0 ? "" : (done >= total ? "done" : "process");
const cpColour = (done, total) => { const s = cpStatusFor(done, total); return s === "done" ? GOLD_HEX : s === "process" ? YELLOW_HEX : WHITE_HEX; };
/* null, not 0: a blank or unreadable box means "no number given", and writing
   white into the sheet because someone cleared the field would be a lie. */
const cpClamp = (v, total) => {
  const s = String(v == null ? "" : v).trim();
  if (!s) return null;
  const n = Number(s);
  return isFinite(n) ? Math.max(0, Math.min(total, Math.round(n))) : null;
};

/** What one item of one job stands at. The list's row and nothing else: no
    Excel colour, no held count, no merge. No row at all is "nothing done",
    which is what a blank cell has always meant - so the list only ever carries
    rows for work that has actually started.

    The shape is exactly what it always was, so every reader of it - the
    drawer, cpSummaryHtml, glassCounts, the export, jobPhase - is unchanged:
    { done, total, status }, with done null for "in progress, count unknown".
    Returns null when the job has no such item. */
function itemState(j, item) {
  const total = cpTotal(j, item);
  if (!(total > 0)) return null;
  const row = cpRow(j && j.id, item);
  /* the switch-on window, and nothing else: no row yet AND the import has not
     drained, so the sheet's own colour answers rather than a hole */
  if (!row && CP_IMPORT_PENDING) return cpFileState(j, item, total);
  const s = row ? row.status : "";
  if (s === "done") return { done: total, total: total, status: "done" };
  if (s === "process") {
    const d = row.done == null ? null : Number(row.done);
    return { done: (d > 0 && d < total) ? d : null, total: total, status: "process" };
  }
  return { done: 0, total: total, status: "" };
}

/* cpWithHeld() lived here until 2026-09-11. It rewrote a job's parsed colours
   from the counts a `cp` PENDING hold was holding, so the screen could show a
   tick before the downloaded file caught up. There is no cp hold any more -
   the list answers in a second and the map is updated at the click - so there
   is nothing to apply and nothing to unmask. */

/* ---- what the sheet's colours stand for, both ways ------------------------
   The import and the safeguard both have to turn a colour into a status and
   back. A colour this feature does not own - the sheet's Cut green, or
   anything else somebody has used - reads as null: it is never adopted, never
   painted over and never counted as a disagreement. */
const CP_WORD_HEX = { done: GOLD_HEX, process: YELLOW_HEX, "": WHITE_HEX };
const CP_GOLD_HEX = ["FFE699", "FFC000"], CP_YELLOW_HEX = ["FFFF00"];
function cpWordForHex(hex) {
  let h = String(hex == null ? "" : hex).trim().toUpperCase().replace(/^#/, "");
  if (!h) return "";                                  // no fill at all reads as white
  if (h.length === 8) h = h.slice(-6);
  if (h === "FFFFFF") return "";
  if (CP_GOLD_HEX.indexOf(h) >= 0) return "done";
  if (CP_YELLOW_HEX.indexOf(h) >= 0) return "process";
  return null;                                        // not a colour checkpoints own
}

/** The one-time import (spec section 4): every item that has no row in the
    list yet and whose Excel cell says something. Pure, so the cap and the
    "never twice" property can be tested without a workbook.

    A BLANK cell is deliberately not imported. A blank item reads as "nothing
    done" with no row at all, so importing every one of them would write a row
    per checkpoint of every job on the sheet - tens of thousands of rows
    carrying no information - and the cap would never catch up. A blank cell
    somebody colours later is caught by the safeguard, not by this. */
function cpImportPlan(jobs, storedFor, cap) {
  const out = [];
  for (let i = 0; i < (jobs || []).length; i++) {
    const j = jobs[i];
    if (!j || !j.id) continue;
    const items = cpItems(j);
    for (let k = 0; k < items.length; k++) {
      const it = items[k];
      if (cpRow(j.id, it.key)) continue;                       // already imported, or clicked
      const st = cpFileStatus(j, it.key);
      if (st !== "done" && st !== "process") continue;         // blank, or a colour we do not own
      const stored = storedFor ? storedFor(String(j.id).toUpperCase(), it.key) : null;
      const n = stored ? Math.round(Number(stored.done) || 0) : 0;
      out.push({ job: j.id, item: it.key, total: it.total, status: st,
                 done: st === "done" ? it.total : (n > 0 && n < it.total ? n : 0) });
      if (cap && out.length >= cap) return out;
    }
  }
  return out;
}

/* ---- one write per burst of taps ----
   Tapping + five times must not send five writes. Each tap restarts the timer
   for that (job, item); when it finally fires, the flush gets the count from
   before the burst started and the count it ended on. The column and the total
   are worked out again at flush time, never taken from the tap - a column
   inserted in Excel meanwhile would otherwise send the fill one cell over.
   A burst waiting to be sent is also kept in localStorage, because closing the
   tab used to lose the write while the screen went on showing the new count
   for three minutes.                                                        */
const CPBURST = {};            // waiting, timer running
const CPSEND = {};             // handed to the flush, not settled yet
const CP_QUEUE_KEY = "cw_cpqueue";
const CP_QUEUE_MS = 3600000;   // older than an hour: the sheet has moved on, do not replay

function cpQueue() {
  const out = [];
  [CPBURST, CPSEND].forEach(m => Object.keys(m).forEach(k => {
    const b = m[k];
    out.push({ key: k, job: b.job, item: b.item, col: b.col, from: b.from, to: b.to,
               total: b.total, at: b.at, who: b.who, sent: b.sent ? 1 : 0 });
  }));
  return out;
}
function cpSaveQueue() { try { localStorage.setItem(CP_QUEUE_KEY, JSON.stringify(cpQueue())); } catch (e) {} }
function cpLoadQueue() {
  let q = [];
  try { q = JSON.parse(localStorage.getItem(CP_QUEUE_KEY) || "[]"); } catch (e) { q = []; }
  const now = Date.now();
  return (q || []).filter(x => x && x.job && x.item && (now - (x.at || 0)) < CP_QUEUE_MS);
}
function cpClearQueue() { try { localStorage.removeItem(CP_QUEUE_KEY); } catch (e) {} }

/* One chain per JOB, shared by single items and whole groups: a group write
   that supersedes an item write must land after it, and two taps on "All done"
   must not interleave. Ordering per job is total. */
const cpChains = {};
function cpChain(key, fn) {
  const prev = cpChains[key] || Promise.resolve();
  const next = prev.then(fn, fn);
  cpChains[key] = next.catch(() => {});
  return next;
}

function cpBurst(key, o, flush, ms) {
  /* `prev` is the list row as it stood before the FIRST tap of this burst, so
     a refused write puts back exactly what was there rather than a count
     worked out afterwards. It is not persisted with the queue: a burst
     replayed on a later visit has no optimistic row to put back, because the
     list already says what it says. */
  const b = CPBURST[key] || (CPBURST[key] = { key: key, from: o.from, prev: o.prev });
  b.job = o.job; b.item = o.item; b.col = o.col; b.total = o.total; b.to = o.to; b.who = o.who;
  b.at = Date.now(); b.flush = flush; b.sent = 0;
  if (b.t) clearTimeout(b.t);
  b.t = setTimeout(() => cpFire(key), ms == null ? CP_DEBOUNCE_MS : ms);
  cpSaveQueue();
  return b;
}
/** Send a waiting burst now (its timer, a group write, or the page closing). */
function cpFire(key) {
  const b = CPBURST[key];
  if (!b) return null;
  if (b.t) clearTimeout(b.t);
  b.t = null;
  delete CPBURST[key];
  /* marked before anything is awaited: if the tab dies mid-write we cannot know
     whether it landed, and replaying it would log it twice and re-assert an old
     count over a newer tick. The 45 s re-read shows what really happened. */
  b.sent = 1;
  CPSEND[key] = b;
  cpSaveQueue();
  cpChain(b.job, () => Promise.resolve(b.flush(b)).then(() => {}, () => {}));
  return b;
}
function cpFireAll() { Object.keys(CPBURST).forEach(cpFire); }
/** The flush is over (either way): the write is no longer owed. */
function cpSettled(key) { delete CPSEND[key]; cpSaveQueue(); }
function cpPending(key) { return CPBURST[key] || CPSEND[key] || null; }
/** Drop a waiting burst - a group write covers its items, so its own write must not follow. */
function cpCancelBurst(key) {
  const b = CPBURST[key];
  if (b && b.t) clearTimeout(b.t);
  delete CPBURST[key];
  cpSaveQueue();
}
/** Taps this browser owed when it was last closed: never sent, still this
    person's, and young enough that the sheet has not moved on without them. */
function cpReplay(flush, who) {
  const q = cpLoadQueue().filter(x => !x.sent && (!x.who || !who || x.who === who));
  cpClearQueue();
  q.forEach(x => {
    const key = x.key || (x.job + "|" + x.item);
    if (CPBURST[key] || CPSEND[key]) return;        // this session already has something newer
    const b = { key: key, job: x.job, item: x.item, col: x.col, from: x.from, to: x.to,
                total: x.total, at: x.at, who: x.who, sent: 1, flush: flush };
    CPSEND[key] = b;
    cpChain(b.job, () => Promise.resolve(flush(b)).then(() => {}, () => {}));
  });
  cpSaveQueue();
  return q.length;
}

/* ---- the writes ----
   THE LIST FIRST, then the fill, then the log line. That order is the whole
   change of 2026-09-11: the list is the record, so nothing has happened until
   it says so; the Excel colour is a copy written for people to look at; and
   the log line records what was done.

   Only fills reach `Production`, only on the job's own row, only in that
   item's own column, and the row is looked up immediately before writing
   because rows move.

   `o.save` is the list write, handed in by app.js (it needs the item ids and
   the consent state, which live there). A refused list write throws before
   anything else happens, so the click simply did not take and the caller puts
   the row back. A refused FILL does NOT undo the list: the office's decision
   stands and Excel is behind, which is the direction this change chose. The
   caller is told, and o.paint is not called, so the safeguard goes on knowing
   what is really in the cell.                                              */
async function cpWriteItem(o) {
  await o.save({ job: o.job, item: o.item, done: o.done, total: o.total,
                 status: cpStatusFor(o.done, o.total), who: o.who });
  try {
    const row = await CW.rowForJob(CP_PROD_SHEET, o.job);
    await CW.setFill(CP_PROD_SHEET, CW.A1(o.col) + row, cpColour(o.done, o.total));
    if (o.paint) o.paint(o.job, o.item, cpStatusFor(o.done, o.total));
    if (o.log) o.log(o.job, cpLabel(o.item), o.from + " of " + o.total, o.done + " of " + o.total);
    return row;
  } catch (e) {
    /* the record landed and the colour did not. The caller must NOT put the
       record back: the office decided, the record says so, and Excel is
       behind - which is the direction this change chose. */
    e.cpRecordStands = true;
    throw e;
  }
}

/** A whole group at once: one list write per item (SharePoint has no batch of
    list writes and graph.js has no helper for one - see cpSaveRow in app.js),
    one batch of fills, one log line. */
async function cpWriteGroup(o) {
  /* which rows actually landed, so a failure half way down can say so rather
     than leaving the job half ticked with a bare error (review finding M7) */
  const saved = [];
  for (let i = 0; i < o.items.length; i++) {
    const x = o.items[i];
    try {
      await o.save({ job: o.job, item: x.item, done: x.done, total: x.total,
                     status: cpStatusFor(x.done, x.total), who: o.who });
      saved.push(x.item);
    } catch (e) {
      e.cpSaved = saved.slice();
      e.cpFailedAt = x.item;
      throw e;
    }
  }
  try {
    const row = await CW.rowForJob(CP_PROD_SHEET, o.job);
    const f = await CW.findFile();
    const S = f.base + "/worksheets('" + CP_PROD_SHEET + "')";
    await CW.batchWrite(o.items.map(x => ({
      method: "PATCH",
      url: S + "/range(address='" + CW.A1(x.col) + row + "')/format/fill",
      body: { color: cpColour(x.done, x.total) }
    })));
    if (o.paint) o.items.forEach(x => o.paint(o.job, x.item, cpStatusFor(x.done, x.total)));
    if (o.log) o.log(o.job, o.what, "", o.to);
    return row;
  } catch (e) {
    e.cpRecordStands = true;                  // as above: the record is the record
    e.cpSaved = saved.slice();
    throw e;
  }
}

if (typeof window !== "undefined") window.CP = {
  itemState, cpItems, cpTotal, cpStatus, cpFileStatus, cpColumn, cpLabel, cpColour, cpStatusFor, cpClamp,
  jobPhase, phaseName, PHASES, effectivePhase, setPhaseHook,
  CP_LIST, CP_LIST_FIELDS, cpTitle, cpRow, cpRowPut, cpRowsSet, cpRowsAll, cpRowsFrom, cpRowFields,
  cpImportPlan, cpWordForHex, CP_WORD_HEX, cpSetImportPending, cpImportPending, cpFileState,
  cpSetProgress, cpStored, cpWriteItem, cpWriteGroup,
  cpBurst, cpFire, cpFireAll, cpSettled, cpPending, cpCancelBurst, cpReplay, cpQueue, cpClearQueue, cpChain,
  WHITE_HEX, YELLOW_HEX, GOLD_HEX, CP_DEBOUNCE_MS
};
/* the pipeline is read by the drawer and by every job row, so it is a plain
   global as well, the way the parser's own helpers are */
if (typeof window !== "undefined") { window.jobPhase = jobPhase; window.phaseName = phaseName; window.PHASES = PHASES;
                                     window.effectivePhase = effectivePhase; window.setPhaseHook = setPhaseHook; }
