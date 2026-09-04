/* Checkpoints - ticking a job's work off item by item.
   Everything here is pure logic: the merge rule between Excel's colour and the
   dashboard's own counts, the colour choice, the burst coalescer and the two
   write sequences. It is kept out of app.js because none of it needs a DOM, so
   test_checkpoints.js can load and exercise it directly.
   Loaded after graph.js (it calls CW) and before app.js. */

/* The sheet carries only the colour. White is written, never cleared: the
   Production cells already have an explicit white fill, and fill/clear would
   take the fill away instead of setting it back to white. */
const WHITE_HEX = "#FFFFFF", YELLOW_HEX = "#FFFF00", GOLD_HEX = "#FFE699";
const CP_DEBOUNCE_MS = 800;               // a run of + taps settles into one write
const CP_PROD_SHEET = "Production";
const CP_SUB = { f: "frames", s: "sashes", t: "transoms" };

/* Exact counts from the Dashboard Progress sheet: job -> item -> {done,total,who,when}.
   Excel's colour still wins over these (see itemState) - the sheet is the truth. */
let PROGRESS = {};
function cpSetProgress(map) { PROGRESS = map || {}; }
function cpStored(job, item) { return (PROGRESS[job] || {})[item] || null; }

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
function cpStatus(j, item) {
  const cp = j.cp || {}, p = String(item).split(":");
  if (item === "win") return cp.win || "";
  if (item === "drs") return cp.drs || "";
  if (p[0] === "glass") return (cp.glass || {})[p[1]] || "";
  if (p[0] === "prod") return ((cp.prod || {})[p[1]] || {})[p[2]] || "";
  return "";
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

/** The merge rule. Excel's colour decides; the stored count only fills in the
    number when the colour says "part way there". null done = "in progress,
    count unknown". Returns null when the job has no such item. */
function itemState(j, item) {
  const total = cpTotal(j, item);
  if (!(total > 0)) return null;
  const held = (j.cpDone && Object.prototype.hasOwnProperty.call(j.cpDone, item)) ? Number(j.cpDone[item]) : null;
  const row = cpStored(j.id, item);
  const d = held != null ? held : (row ? Number(row.done) : null);
  const s = cpStatus(j, item);
  if (s === "done") return { done: total, total: total, status: "done" };
  if (s === "process") return { done: (d > 0 && d < total) ? d : null, total: total, status: "process" };
  return { done: 0, total: total, status: "" };
}

/** A copy of the job's Excel statuses with the held (pending) counts applied,
    so the screen shows the new colour before the file catches up. */
function cpWithHeld(j, held) {
  const src = j.cp || {};
  const cp = { win: src.win || "", drs: src.drs || "", glass: Object.assign({}, src.glass), prod: {} };
  Object.keys(src.prod || {}).forEach(n => { cp.prod[n] = Object.assign({}, src.prod[n]); });
  Object.keys(held || {}).forEach(item => {
    const total = cpTotal(j, item);
    if (!(total > 0)) return;
    const st = cpStatusFor(Number(held[item]), total), p = String(item).split(":");
    if (item === "win") cp.win = st;
    else if (item === "drs") cp.drs = st;
    else if (p[0] === "glass") cp.glass[p[1]] = st;
    else if (p[0] === "prod") (cp.prod[p[1]] = cp.prod[p[1]] || {})[p[2]] = st;
  });
  return cp;
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
  const b = CPBURST[key] || (CPBURST[key] = { key: key, from: o.from });
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
   Only fills, only on the job's own row, only in that item's own column. The
   row is looked up immediately before writing because rows move.
   Our own count goes first: it is ours to lose. If the fill then fails, the
   count is put back, so the sheet and the stored number never disagree in a
   way nobody was told about. */
async function cpWriteItem(o) {
  await CW.saveProgress(o.job, o.item, o.done, o.total, o.who);
  try {
    const row = await CW.rowForJob(CP_PROD_SHEET, o.job);
    await CW.setFill(CP_PROD_SHEET, CW.A1(o.col) + row, cpColour(o.done, o.total));
    if (o.log) o.log(o.job, cpLabel(o.item), o.from + " of " + o.total, o.done + " of " + o.total);
    return row;
  } catch (e) {
    /* 0 when the count was unknown: with the cell still yellow that reads back
       as "in progress", which is exactly where it was */
    try { await CW.saveProgress(o.job, o.item, o.from == null ? 0 : o.from, o.total, o.who); } catch (e2) {}
    throw e;
  }
}

/** A whole group at once: one progress write, one batch of fills, one log line. */
async function cpWriteGroup(o) {
  await CW.saveProgressMany(o.job, o.items, o.who);
  try {
    const row = await CW.rowForJob(CP_PROD_SHEET, o.job);
    const f = await CW.findFile();
    const S = f.base + "/worksheets('" + CP_PROD_SHEET + "')";
    await CW.batchWrite(o.items.map(x => ({
      method: "PATCH",
      url: S + "/range(address='" + CW.A1(x.col) + row + "')/format/fill",
      body: { color: cpColour(x.done, x.total) }
    })));
    if (o.log) o.log(o.job, o.what, "", o.to);
    return row;
  } catch (e) {
    if (o.before && o.before.length) {
      try { await CW.saveProgressMany(o.job, o.before.map(x => ({ item: x.item, done: x.done == null ? 0 : x.done, total: x.total })), o.who); } catch (e2) {}
    }
    throw e;
  }
}

if (typeof window !== "undefined") window.CP = {
  itemState, cpItems, cpTotal, cpStatus, cpColumn, cpLabel, cpColour, cpStatusFor, cpClamp,
  cpWithHeld, cpSetProgress, cpStored, cpWriteItem, cpWriteGroup,
  cpBurst, cpFire, cpFireAll, cpSettled, cpPending, cpCancelBurst, cpReplay, cpQueue, cpClearQueue, cpChain,
  WHITE_HEX, YELLOW_HEX, GOLD_HEX, CP_DEBOUNCE_MS
};
