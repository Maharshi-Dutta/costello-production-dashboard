/* Costello production dashboard - UI over the live SharePoint workbook. */

const $ = s => document.querySelector(s);
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const cap = s => String(s || "").replace(/\b\w/g, m => m.toUpperCase());
const dshort = d => d ? d.slice(8, 10) + "/" + d.slice(5, 7) : "—";

const CATS = {
  deliver: { l: "Ready to deliver", c: "--done" }, collect: { l: "Collect/Supply", c: "--single" },
  wonttake: { l: "Won't take", c: "--notsent" }, secondhand: { l: "Second hand", c: "--ink-3" }
};
const STAGE = { floor: { l: "On floor", c: "--single" }, ready: { l: "Waiting", c: "--fab" }, office: { l: "In office", c: "--notsent" } };
const STEPS = [["sold", "Sold"], ["stamp", "Stamp"], ["ivana", "Ivana"], ["ready", "Ready to print"], ["floor", "Sent to floor"]];
const SORTS = [["id", "Job no (A-Z)"], ["num", "Job number (ignore letter)"], ["cat", "Category"],
               ["urgent", "Urgent first"], ["size", "Biggest first"], ["wait", "Longest wait"], ["county", "County"]];
/* R5244, C4794, S5136 ... sort on the digits alone, so the list runs in job
   order rather than being grouped by whichever letter the job happens to carry. */
const jobNum = j => { const m = /(\d+)/.exec(j.id); return m ? parseInt(m[1], 10) : 0; };
const CATORDER = ["deliver", "collect", "wonttake", "secondhand", "floor", "ready", "office"];
const SHEETNAMES = ["Production", "Production (2)", "PA Lam", "Glass", "Wds Prep", "Glazing", "Cut & Weld", "PVC Doors", "Smart Slides", "THWS", "Call Log"];
/* The floor stations, as the dropdown above the job list offers them. One
   entry per station page: [key, what the dropdown calls it]. Adding the next
   station is adding a line here and a renderer for its key - the dropdown, the
   empty states and the reset in renderChips all follow from this array.
   SHEETNAMES stays exactly as it was, for the export filter and the row chips. */
const STATIONS = [["glass", "Glass station"]];
/* Everything the Show dropdown can put in the job list's place. The floor
   stations, and then the John print sheet - which is not a station at all: it
   is the office's own second sheet, "Production (2)", shown on its own terms
   (see JOHNROWS below). Anything keyed here that is not in STATIONS must never
   reach the station polling, feeding or list code. */
const BOARDS = STATIONS.concat([["john", "John print sheet"]]);
/* "Production (2)" as parseJohnSheet last read it: the rows John's paper is
   printed from, in that sheet's own order. Never merged into ALL. */
let JOHNROWS = [];

/* ---- the row colour code ----
   The Production sheet says four things with the colour of a job row's text.
   parser.js reads it into j.flag (the word) and j.flagHex (the colour that said
   so). Everywhere it is shown, the WORD is shown: the colour only ever goes
   with it, because some of the people reading these screens and printouts
   cannot tell red from green. */
const FLAGWORD = { urgent: "Urgent", booked: "Booked", trade: "Trade order", hold: "On hold" };
/* what to draw it in when the sheet's own hex is missing - close to the
   colours the office uses, and readable on both themes */
const FLAGINK = { urgent: "var(--urgent)", booked: "var(--green)", trade: "#B5179E", hold: "var(--single)" };
/* the sheet's own hex, pushed into a readable range for whichever theme is on
   (inkFor, below) - the exported FILE always keeps the sheet's hex exactly */
const flagInk = j => inkFor((j && j.flagHex) || "", false) || (FLAGINK[j && j.flag] || "var(--ink)");
/** Tick or untick a whole list of jobs at once - a section, or everything the
    list is showing. Two hundred jobs is two hundred assignments and ONE draw,
    which is the whole reason this is a function and not a loop that calls
    renderAll each time round. Hands back how many are ticked afterwards. */
function pickMany(ids, on) {
  (ids || []).forEach(id => { if (on) state.picked[id] = 1; else delete state.picked[id]; });
  renderAll();
  return Object.keys(state.picked).length;
}

/** The little chip after a job number, in the list and at the head of the
    drawer. Nothing at all when the row's text is plain black. */
function flagChip(j) {
  if (!j || !j.flag || !FLAGWORD[j.flag]) return "";
  return '<span class="badge flagchip" style="color:' + flagInk(j) + '" title="The colour of this row on the Production sheet">' +
    esc(FLAGWORD[j.flag]) + "</span>";
}

let ALL = [], PRODMAP = null, lastStamp = null, busy = false;
/* SharePoint takes ~35s to write our change into the downloadable file, while
   the Excel API reflects it in ~1s. So we apply our own writes locally at once
   and hold them until the file catches up - otherwise the next refresh reads a
   stale file and appears to undo what you just did. */
let PENDING = {};
const PENDING_MS = 180000;
/* kept in localStorage, not just memory: refreshing the page used to discard
   these, and the freshly downloaded file is still ~36s behind, so your own
   change would disappear the moment you pressed F5. */
try { PENDING = JSON.parse(localStorage.getItem("cw_pending") || "{}"); } catch (e) { PENDING = {}; }
const savePending = () => { try { localStorage.setItem("cw_pending", JSON.stringify(PENDING)); } catch (e) {} };
/* The two vocabularies for one cell, and the translation between them. The
   parser reads a fill and says "done" / "process" / "cut" / nothing; the glass
   colour feature (§ glass colours, below) decides "gold" / "yellow" / "" and
   writes the hex. A colour this feature does not own - the sheet's own Cut
   green - has no word here at all, which is exactly how the writer knows to
   leave that cell alone. */
const GLASS_COLOUR_WORD = { done: "gold", process: "yellow", "": "" };
const GLASS_WORD_CP = { gold: "done", yellow: "process", "": "" };
/** What the Production sheet is showing for one of the four glass columns,
    in the colour feature's own words - undefined when it is a colour this
    feature has no business touching. */
const glassCellNow = (j, type) => GLASS_COLOUR_WORD[((j.cp || {}).glass || {})[type] || ""];

/* Each held thing carries its own timestamp: ticking a checkpoint must not
   extend the hold on an unrelated change made two minutes earlier. */
function pend(id, patch) {
  const p = PENDING[id] || (PENDING[id] = { at: 0, prods: {} });
  const now = Date.now(), was = p.at || now;
  const t = p.t = p.t || {};
  /* entries written by an older build have one timestamp for the whole job */
  if ("done" in p && !t.done) t.done = was;
  if (p.blk != null && !t.blk) t.blk = was;
  Object.keys(p.prods || {}).forEach(k => { if (!t["prod:" + k]) t["prod:" + k] = was; });
  Object.keys(p.cp || {}).forEach(k => { if (!t["cp:" + k]) t["cp:" + k] = was; });
  Object.keys(p.gc || {}).forEach(k => { if (!t["gc:" + k]) t["gc:" + k] = was; });
  p.at = now;
  if ("done" in patch) { p.done = patch.done; t.done = now; }
  if (patch.prod) { p.prods[patch.prod.name] = patch.prod.status; t["prod:" + patch.prod.name] = now; }
  if (patch.cp) {
    p.cp = p.cp || {};
    /* null means "stop holding this one" - used when a write failed and the
       count it replaced was itself unknown, so there is nothing to put back */
    for (const k in patch.cp) {
      if (patch.cp[k] == null) { delete p.cp[k]; delete t["cp:" + k]; }
      else {
        p.cp[k] = patch.cp[k]; t["cp:" + k] = now;
        /* THE OFFICE IS ABSOLUTE, AND THIS IS WHERE THE MORNING OF 2026-09-11
           WENT WRONG. A `gc` hold is the colour writer saying "I have painted
           this cell and the download has not caught up". applyPending only
           ever MASKED one with a newer cp hold on the same glass type - it
           never dropped it. So the office pressed Clear, its cp hold outranked
           the writer's gold and the drawer went white; six seconds later the
           download agreed with the CLEAR, the office's own hold was let go,
           and the writer's older gold - still sitting underneath - came
           straight back over a sheet the office had just made white. It then
           stood for eleven and a half minutes (a hold is now kept while the
           file disagrees, twelve reads), long enough to cross the feeder's ten
           minute window and re-lock the tablet at 0/0/0/0.

           The moment the office acts on a glass cell, the writer's un-landed
           paint of that cell is VOID. Not masked - discarded. */
        const ty = k.indexOf("glass:") === 0 ? k.slice(6) : "";
        if (ty && p.gc && Object.prototype.hasOwnProperty.call(p.gc, ty)) {
          delete p.gc[ty]; delete t["gc:" + ty];
          if (p.n) delete p.n["gc:" + ty];
        }
      }
    }
  }
  /* A glass colour the floor's work has just put into the sheet, held the same
     way and for the same reason as a checkpoint tick: the downloaded file is
     about 36 s behind the API, and without the hold the next refresh would
     read the old colour back and the cell would flicker gold, blank, gold. The
     value is the colour's own word ("gold" / "yellow" / ""); null drops the
     hold, which is what a failed write does. */
  if (patch.gc) {
    p.gc = p.gc || {};
    for (const k in patch.gc) {
      if (patch.gc[k] == null) { delete p.gc[k]; delete t["gc:" + k]; }
      else { p.gc[k] = patch.gc[k]; t["gc:" + k] = now; }
    }
  }
  if ("blk" in patch) { if (patch.blk == null) { delete p.blk; delete t.blk; } else { p.blk = patch.blk; t.blk = now; } }
  /* A hand-set phase, held while the SharePoint list catches up. 0-6 is a
     phase; -1 means "held as cleared" - the item has just been deleted and a
     read a second later may still be showing it; null drops the hold, the way
     it does for blk, which is what a failed write does. */
  if ("phase" in patch) { if (patch.phase == null) { delete p.phase; delete t.phase; } else { p.phase = patch.phase; t.phase = now; } }
  savePending();
}
const blkCat = b => b === 0 ? "secondhand" : b === 1 ? "wonttake" : b === 2 ? "collect" : "active";
const pendEmpty = p => !("done" in p) && p.blk == null && p.phase == null &&
  !Object.keys(p.prods || {}).length && !Object.keys(p.cp || {}).length &&
  !Object.keys(p.gc || {}).length;

/* ---- a hold is never let go into a copy that still disagrees with it -------
   Owner's bug, observed 2026-09-10 with the mechanism. Un-tick a job's glass,
   then REFRESH the page - which the owner does after an un-tick, because the
   tablet looked locked, and not after a tick. `cw_pending` survives the reload
   byte for byte, so the screen is correctly white. But:

     · the 45 s reconcile timer died with the old page - it lives in memory -
       and poll() only downloads when `lastModified` MOVES, which this
       dashboard's own write was the last thing to do and the boot load has
       already recorded. So nothing ever re-reads the file, and the page sits
       on the stale gold parse with the hold as its only cover;
     · and the hold's expiry was a pure clock test. The next time anything at
       all called applyPending - a floor tap on a DIFFERENT job, through
       glassColourRun - the three-minute-old hold was dropped and the gold
       underneath was UNMASKED. Minutes after an un-tick, and it stayed until
       something else happened to move lastModified.

   So expiry no longer means "let go". It means "ask again":

     · a hold kept because the parse still disagrees keeps a re-read armed;
     · an EXPIRED hold whose parse still disagrees is kept, and a read is
       demanded now rather than waited for;
     · it is let go the moment the file agrees - and only then;
     · after HOLD_GIVEUP fresh parses that still disagree it is let go anyway
       and the office is told, in red, naming the job and the item: at that
       point the write really may not have saved, and quietly showing a colour
       we know is older than our own change is the one thing not to do.

   What counts as "disagrees" is what a release would UNMASK, which is the
   colour: the sheet carries only that. A held count whose colour the file
   already shows expires quietly, exactly as before - otherwise an item whose
   exact count the Dashboard Progress sheet cannot answer for would be held for
   ever and warned about for nothing. */
const HOLD_GIVEUP = 12;             // fresh parses still disagreeing: ~9 min at 45 s apart
let holdForceAt = 0;                // when a read was last demanded for a hold
/** Keep a re-read coming while a hold is waiting for the file. `soon` demands
    one now - bounded to one a minute, because applyPending is called from
    every render and a download per call would be a storm. */
function reconcileForHolds(soon) {
  if (soon) {
    if (Date.now() - holdForceAt < 45000) { if (!reconcileT) scheduleReconcile(); return; }
    holdForceAt = Date.now();
    scheduleReconcile(0);
    return;
  }
  if (!reconcileT) scheduleReconcile();
}
/** Would letting this checkpoint hold go change what the row SHOWS? The sheet
    carries the colour and nothing else, so the colour is what a release can
    unmask - and what this refuses to unmask while it is still wrong. */
function cpHoldUnmasks(j, item, held) {
  const total = cpTotal(j, item);
  if (!(total > 0)) return false;
  const st = itemState(j, item);
  return !!st && cpStatusFor(held, total) !== st.status;
}
/** An expired hold the file still disagrees with. Answers whether to keep it.
    Only a FRESH parse counts towards giving up: a parse is what could have
    caught up, and applyPending is called far more often than the file is read. */
function holdStuck(p, key, fresh) {
  const n = p.n = p.n || {};
  if (fresh) n[key] = (n[key] || 0) + 1;
  if ((n[key] || 0) >= HOLD_GIVEUP) { delete n[key]; return false; }
  reconcileForHolds(!fresh);
  return true;
}
/** Said once, when a change is given up on. Never one per read. */
function holdGaveUp(job, what, floor) {
  /* worded by WHOSE change it was. "your change may not have saved" is wrong
     for a colour this dashboard painted from the floor's counters: the office
     never made that change and has nothing to check for having made it. */
  toast(floor
    ? job + ": the colour painted from the floor's work for " + what +
      " is not showing in the sheet yet. Check the sheet."
    : job + ": the sheet still does not show your change to " + what +
      " — it may not have saved. Check the sheet.", true);
}
/** A page that starts up holding something asks to be re-read, because after a
    reload nothing else ever will: the reconcile timer is gone with the old
    page and poll() only downloads when lastModified moves - which this
    dashboard's own write was the last thing to do. Sooner when the hold is
    already old, because that write went out before the reload. */
function bootReconcile() {
  let oldest = 0;
  Object.keys(PENDING).forEach(id => {
    const p = PENDING[id];
    if (pendEmpty(p)) return;
    const t = p.t || {};
    const ages = Object.keys(t).map(k => t[k]).concat([p.at || 0]).filter(Boolean);
    const a = ages.length ? Math.min.apply(null, ages) : p.at || 0;
    if (!oldest || a < oldest) oldest = a;
  });
  if (!oldest) return;
  scheduleReconcile(Date.now() - oldest > 30000 ? 5000 : 45000);
}

/** `fresh` = this is a newly parsed workbook, so a held count can be compared
    with what the file now says and let go once the two agree. */
function applyPending(list, fresh) {
  const now = Date.now();
  let dropped = false;
  /* the jobs this call can actually compare a hold against */
  const seen = {};
  (list || []).forEach(x => { const j = (x && x.raw) || x; if (j && j.id) seen[j.id] = 1; });
  Object.keys(PENDING).forEach(id => {
    const p = PENDING[id], t = p.t = p.t || {};
    const old = k => now - (t[k] || p.at || 0) > PENDING_MS;
    if ("done" in p && old("done")) { delete p.done; delete t.done; dropped = true; }
    if (p.blk != null && old("blk")) { delete p.blk; delete t.blk; dropped = true; }
    if (p.phase != null && old("phase")) { delete p.phase; delete t.phase; dropped = true; }
    Object.keys(p.prods || {}).forEach(k => { if (old("prod:" + k)) { delete p.prods[k]; delete t["prod:" + k]; dropped = true; } });
    /* cp and gc holds are settled against the parsed job below, where there is
       something to compare them WITH. A hold on a job this call cannot see -
       gone from the sheet, or a list that does not carry it - has nothing to
       be compared against and still goes on the clock alone. */
    if (!seen[id]) {
      Object.keys(p.cp || {}).forEach(k => { if (old("cp:" + k)) { delete p.cp[k]; delete t["cp:" + k]; dropped = true; } });
      Object.keys(p.gc || {}).forEach(k => { if (old("gc:" + k)) { delete p.gc[k]; delete t["gc:" + k]; dropped = true; } });
    }
    if (pendEmpty(p)) { delete PENDING[id]; dropped = true; }
  });
  const out = list.map(x => {
    /* always re-apply to the parsed job, never to an already patched copy:
       letting a hold go has to give back exactly what the file says */
    const j = x.raw || x;
    const p = PENDING[j.id];
    if (!p) return j;
    /* the file has caught up with this tick: stop holding it, so what Excel
       says takes over again straight away */
    /* the list now says what we clicked (or that it is gone again): let go, so
       what everyone else can see takes over */
    if (fresh && p.phase != null) {
      const set = PHASES_SET[String(j.id).toUpperCase()];
      if (p.phase < 0 ? !set : (set && set.phase === p.phase)) { delete p.phase; delete p.t.phase; dropped = true; }
    }
    const stale = k => now - ((p.t || {})[k] || p.at || 0) > PENDING_MS;
    const letCpGo = k => { delete p.cp[k]; delete p.t["cp:" + k];
                           if (p.n) delete p.n["cp:" + k]; dropped = true; };
    if (p.cp) Object.keys(p.cp).forEach(k => {
      const st = itemState(j, k);
      /* the file has caught up with this tick, to the count: let go */
      if (fresh && st && st.done != null && st.done === p.cp[k]) { letCpGo(k); return; }
      if (!stale("cp:" + k)) return;                    // still inside its three minutes
      /* expired. If letting go would put back a colour we know is older than
         our own change, it is NOT let go - the file is asked for again. */
      if (!cpHoldUnmasks(j, k, p.cp[k])) { letCpGo(k); return; }
      if (holdStuck(p, "cp:" + k, fresh)) return;
      const label = typeof cpLabel === "function" ? cpLabel(k) : k;
      letCpGo(k);
      holdGaveUp(j.id, label);
    });
    /* the same rule for a glass colour: the file now shows what we painted, so
       stop holding it and let the sheet speak for itself again */
    const letGcGo = k => { delete p.gc[k]; delete p.t["gc:" + k];
                           if (p.n) delete p.n["gc:" + k]; dropped = true; };
    if (p.gc) Object.keys(p.gc).forEach(k => {
      /* a gc hold IS a colour, so agreement and unmasking are the same test */
      if (glassCellNow(j, k) === p.gc[k]) { letGcGo(k); return; }
      if (!stale("gc:" + k)) return;
      if (holdStuck(p, "gc:" + k, fresh)) return;
      letGcGo(k);
      holdGaveUp(j.id, "Glass " + String(k).toUpperCase(), true);
    });
    if (pendEmpty(p)) { delete PENDING[j.id]; dropped = true; return j; }
    const c = Object.assign({}, j);
    c.raw = j;
    if ("done" in p) c.done = p.done;
    if (p.cp && Object.keys(p.cp).length) { c.cp = cpWithHeld(j, p.cp); c.cpDone = Object.assign({}, p.cp); }
    /* Both sides can be holding the same cell at once: the office ticked DG in
       the drawer at 14:00:00 and a floor tap reached this dashboard two seconds
       later, and both writes are still in the air. The newer of the two holds
       is the one shown - which is the same last-writer-wins rule the writer
       itself applies, applied to the two things this browser has not yet seen
       land. Never edited in place: `c.cp` can still be the parsed job's own
       object, and `x.raw` has to keep saying exactly what the file said. */
    if (p.gc && Object.keys(p.gc).length) {
      const glass = Object.assign({}, (c.cp || {}).glass);
      Object.keys(p.gc).forEach(k => {
        if ((p.t["cp:glass:" + k] || 0) > (p.t["gc:" + k] || 0)) return;
        glass[k] = GLASS_WORD_CP[p.gc[k]];
        /* the count the office's own tick was holding is no longer what this
           cell is about, so the drawer must not go on showing it */
        if (c.cpDone) delete c.cpDone["glass:" + k];
      });
      c.cp = Object.assign({}, c.cp, { glass: glass });
    }
    if (p.blk != null) { c.blk = p.blk; c.cat = blkCat(p.blk); }   // moved in Excel; file still catching up
    c.prods = j.prods.map(y => Object.prototype.hasOwnProperty.call(p.prods, y.n)
      ? Object.assign({}, y, { st: p.prods[y.n] ? [p.prods[y.n]] : [] }) : y);
    c.stage = c.done ? "deliver" : (c.dates.floor ? "floor" : (c.dates.ready ? "ready" : "office"));
    return c;
  });
  if (dropped) savePending();
  /* anything still held after a fresh parse is something the file has not
     caught up with, so keep a re-read coming until it has. This is also the
     answer to a download served OLDER than the one before it, which SharePoint
     does: the hold stands and the next read is asked for. */
  if (fresh && Object.keys(PENDING).some(id => {
    const p = PENDING[id];
    return Object.keys(p.cp || {}).length || Object.keys(p.gc || {}).length;
  })) reconcileForHolds(false);
  if (list.blockNames) out.blockNames = list.blockNames;   // the grouped view reads them off the list
  return out;
}

/* Placements in dashboard-only categories get the same treatment: kept here
   until the downloaded file shows them, so a refresh cannot undo a move. */
let PENDV = {};
try { PENDV = JSON.parse(localStorage.getItem("cw_pendv") || "{}"); } catch (e) { PENDV = {}; }
const savePendV = () => { try { localStorage.setItem("cw_pendv", JSON.stringify(PENDV)); } catch (e) {} };
function pendView(view, id, group, order) { PENDV[view + "|" + id] = { group: String(group), order: order, at: Date.now() }; savePendV(); }
function applyPendV() {
  const now = Date.now(); let changed = false;
  Object.keys(PENDV).forEach(k => {
    const p = PENDV[k], i = k.indexOf("|"), view = k.slice(0, i), id = k.slice(i + 1);
    const cur = VIEWS[view] && VIEWS[view][id];
    if (now - p.at > PENDING_MS || (cur && String(cur.group) === p.group)) { delete PENDV[k]; changed = true; return; }
    (VIEWS[view] = VIEWS[view] || {})[id] = { group: p.group, order: p.order };
  });
  if (changed) savePendV();
}

let CHANGES = [];                     // what has changed while this page has been open
try { CHANGES = JSON.parse(localStorage.getItem("cw_changes") || "[]"); } catch (e) { CHANGES = []; }
/* Alert lines carry an email address, and a subscription lives in the workbook
   only - never in this browser beyond the 180 s hold below. They are left out
   of the cached copy and come back from the Dashboard Log sheet on the next
   read, so the Changes window still shows them. */
const saveChanges = () => { try {
  localStorage.setItem("cw_changes", JSON.stringify(CHANGES.filter(c => c.what !== "Alert").slice(0, 400)));
} catch (e) {} };
let state = { q: "", cat: null, sheet: null, sort: "id", desc: false, sel: null, edit: false,
              scope: "", view: "flat", picked: {}, collapsed: {}, hidden: {},
              board: null };     // null = the job list · "glass" = the Glass station board
let VIEWS = {};            // view name -> { job -> {group, order} }  (from the workbook)
let BLOCKNAMES = [];
try { state.hidden = JSON.parse(localStorage.getItem("cw_hidden") || "{}"); } catch (e) {}
try { state.collapsed = JSON.parse(localStorage.getItem("cw_collapsed") || "{}"); } catch (e) {}
/* the drawer's five date steps start folded away: the phase pipeline above them
   is the progress people actually read. Same store, same rule as a group in the
   list - a truthy value means collapsed - so it only needs seeding once. */
if (!("dates" in state.collapsed)) state.collapsed.dates = 1;
const saveUi = () => { try {
  localStorage.setItem("cw_hidden", JSON.stringify(state.hidden));
  localStorage.setItem("cw_collapsed", JSON.stringify(state.collapsed));
} catch (e) {} };
const viewNames = () => Object.keys(VIEWS).sort();
/** which group a job sits in, for a given view: an explicit placement wins,
    otherwise its natural block from the Production sheet. */
function groupOf(j, view) {
  if (view === "Abin") return j.blk;                 // the sheet is the truth: a move here moves the row in Excel
  const v = VIEWS[view];
  if (v && v[j.id] && v[j.id].group !== "") return Number(v[j.id].group);
  return j.blk;
}
function inView(j, view) {
  if (view === "flat") return true;
  const v = VIEWS[view] || {};
  if (view === "Abin") return j.blk >= 0;          // the whole sheet, grouped
  return !!v[j.id];                                 // custom views: only what was put there
}

const live = () => ALL.filter(j => j.cat !== "past");
const byId = id => ALL.find(j => j.id === id);
const comp = j => j.prods.reduce((a, p) => ({ f: a.f + p.f, s: a.s + p.s, t: a.t + p.t }), { f: 0, s: 0, t: 0 });
const tot = c => c.f + c.s + c.t;
function catOf(j) {
  if (j.cat === "collect") return "collect";
  if (j.cat === "wonttake") return "wonttake";
  if (j.cat === "secondhand") return "secondhand";
  if (j.done) return "deliver";
  return j.stage;
}
const label = j => CATS[catOf(j)] || STAGE[j.stage];

function toast(msg, isErr) {
  const t = document.createElement("div");
  t.className = "toast" + (isErr ? " err" : "");
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), isErr ? 7000 : 3200);
}
/** Turn a Graph error into something a person can act on. */
function friendly(e) {
  const m = (e && e.message) || String(e);
  if (/InvalidSession/i.test(m)) return "The Excel connection timed out and was renewed - please try that again.";
  if (/no longer on the/i.test(m)) return m;
  if (/40[13]/.test(m)) return "Excel refused that change - you may not have edit rights on the workbook.";
  if (/423|locked/i.test(m)) return "The workbook is locked, usually because someone has it open in desktop Excel without AutoSave.";
  if (/50[0234]|timeout/i.test(m)) return "Excel was busy and did not answer. Try again in a moment.";
  return "Write failed: " + m.slice(0, 160);
}
function setStatus(text, kind) {
  $("#status").textContent = text;
  $("#livedot").className = "dot" + (kind ? " " + kind : "");
}

/* ---------- what changed ---------- */
const CATNAME = { deliver:"Ready to deliver", collect:"Collect & supply", wonttake:"Won't take",
                  secondhand:"Second hand", floor:"On floor", ready:"Waiting for floor", office:"In office" };

const CPWORD = { "": "not started", process: "in fabrication", done: "done" };
const GROUPDONE = /: (all done|cleared)$/;

/** Drop the differences that describe what this dashboard itself just did, so a
    change is not listed twice. A group write logs one line ("Glass: all done")
    but shows up in the sheet as one difference per item ("Glass TG"), so the
    group's own label covers those too. */
function dropMine(list, changes, now) {
  const t = now || Date.now();
  const recent = (changes || []).filter(c => c.src === "dashboard" && t - new Date(c.at).getTime() < 600000);
  const mine = new Set(recent.map(c => c.job + "|" + c.what));
  const groups = recent.filter(c => GROUPDONE.test(c.what))
    .map(c => ({ job: c.job, up: c.what.replace(GROUPDONE, "").toUpperCase() }));
  return list.filter(c => {
    if (mine.has(c.job + "|" + c.what)) return false;
    const w = String(c.what).toUpperCase();
    return !groups.some(g => g.job === c.job && (w === g.up || w.indexOf(g.up + " ") === 0));
  });
}

/** Compare two parses of the workbook and describe every difference in plain terms. */
function diffJobs(prev, next, who, at) {
  const out = [], pm = {}, nm = {};
  prev.forEach(j => pm[j.id] = j);
  next.forEach(j => nm[j.id] = j);
  const add = (job, what, from, to) => out.push({ at, who, job, what, from: String(from), to: String(to), src: "sheet" });

  next.forEach(n => {
    const p = pm[n.id];
    if (!p) { if (n.cat !== "past") add(n.id, "Added to the sheet", "", CATNAME[catOf(n)] || n.cat); return; }
    if (!!p.done !== !!n.done) add(n.id, "Ready to deliver", p.done ? "yes" : "no", n.done ? "yes" : "no");
    if (p.cat !== n.cat) add(n.id, "Category", CATNAME[p.cat] || p.cat, CATNAME[n.cat] || n.cat);
    if (p.blk !== n.blk && p.blk >= 0 && n.blk >= 0)
      add(n.id, "Section", (prev.blockNames || BLOCKNAMES)[p.blk] || ("section " + p.blk),
          (next.blockNames || BLOCKNAMES)[n.blk] || ("section " + n.blk));
    ["sold","stamp","ivana","ready","floor"].forEach(k => {
      if ((p.dates[k] || "") !== (n.dates[k] || ""))
        add(n.id, k === "floor" ? "Sent to floor" : k.charAt(0).toUpperCase() + k.slice(1), p.dates[k] || "blank", n.dates[k] || "cleared");
    });
    if (p.wnd !== n.wnd || p.drs !== n.drs) add(n.id, "Quantity wnd/drs", p.wnd + "/" + p.drs, n.wnd + "/" + n.drs);
    if ((p.cust || "") !== (n.cust || "")) add(n.id, "Customer", p.cust || "blank", n.cust || "blank");
    const ps = {}, ns = {};
    p.prods.forEach(x => ps[x.n] = x); n.prods.forEach(x => ns[x.n] = x);
    Object.keys(ns).forEach(k => {
      const b = ps[k], c = ns[k];
      if (!b) { add(n.id, "Product added - " + cap(k), "", c.f + "/" + c.s + "/" + c.t); return; }
      if (b.f !== c.f || b.s !== c.s || b.t !== c.t)
        add(n.id, cap(k) + " F/S/T", b.f + "/" + b.s + "/" + b.t, c.f + "/" + c.s + "/" + c.t);
    });
    Object.keys(ps).forEach(k => { if (!ns[k]) add(n.id, "Product removed - " + cap(k), ps[k].f + "/" + ps[k].s + "/" + ps[k].t, ""); });
    /* checkpoint colours, named the same way the dashboard's own log names them,
       so a tick made here is not also listed as a change spotted in Excel */
    cpItems(n).forEach(x => {
      const was = cpStatus(p, x.key), is = cpStatus(n, x.key);
      if (was !== is) add(n.id, cpLabel(x.key), CPWORD[was], CPWORD[is]);
    });
    if (p.notes.length !== n.notes.length) {
      const old = p.notes.map(x => x.t), fresh = n.notes.filter(x => old.indexOf(x.t) < 0);
      fresh.forEach(x => add(n.id, "Note added", "", x.t.slice(0, 90)));
    }
  });
  prev.forEach(p => { if (!nm[p.id] && p.cat !== "past") add(p.id, "Removed from the sheet", CATNAME[catOf(p)] || p.cat, ""); });
  return out;
}

/** Comments for one job, oldest first. They are stored as log rows, so they are
    shared with everyone and survive browsers and devices. */
function commentsFor(id) {
  return CHANGES.filter(c => c.job === id && c.what === "Comment")
    .slice().reverse();
}
async function addComment(id, text) {
  noteChange(id, "Comment", "", text);      // local at once, log row behind it
  toast("Comment added");
}

function whoAmI() {
  const a = CW.account;
  return (a && (a.username || a.name)) || "unknown";
}
function noteChange(job, what, from, to) {
  CHANGES.unshift({ at: new Date().toISOString(), who: whoAmI(), job, what,
                    from: String(from), to: String(to), src: "dashboard" });
  saveChanges(); updateChangeBtn();
  CW.appendLog(whoAmI(), job, what, from, to);   // permanent, shared, log sheet only
}
function updateChangeBtn() {
  const b = $("#changebtn"); if (!b) return;
  b.textContent = CHANGES.length ? "Changes (" + CHANGES.length + ")" : "Changes";
}

/* ---------- hand-set phases -------------------------------------------------
   A phase can be set by hand when the sheet has not caught up - the frames are
   cut but nobody has coloured the cell yet. That decision is shared through a
   SharePoint list in the same site, "Dashboard phases", and NOTHING about it
   goes into the workbook: no tab, no column, no colour, no cell. The list is
   made by hand in SharePoint; if it is not there this whole feature says so
   plainly and writes nothing at all.

   The sheet always wins when it is further on: effectivePhase() (checkpoints.js)
   takes the higher of the two, so clearing a hand-set phase drops the job
   straight back to what the sheet says.                                      */
const PHASE_LIST = "Dashboard phases";
const PHASE_LIST_MISSING = "The \u201cDashboard phases\u201d list is not in SharePoint yet, so phases cannot be set here. " +
  "Ask the manager to add it - nothing in the Excel file is involved.";
const PHASE_BELOW_SHEET = "the sheet already shows this job past this step";
const PHASE_NEED_CONSENT = "Setting phases needs a SharePoint permission that has not been granted yet. Nothing in the Excel file is involved.";
const PHASE_CHECKING = "checking…";      // the first read of the list has not answered yet
let PHASES_SET = {};          // { JOB: { phase, name, who, at } } - everyone's hand-set phases
let PHASE_LIST_OK = null;     // null: not looked yet · false: no such list · true: read it
let PHASE_LIST_CONSENT = false; // the last read failed because the list permission has not been granted yet
let phaseWarned = false;      // one toast per page for a list that will not read
const PHASEBUSY = {};         // job -> true while its write is in flight

/** Read the whole list. Failures are tolerated: the last known map stays, and
    one toast is shown per page, not one per refresh. */
async function readPhases() {
  if (typeof CW === "undefined" || !CW || typeof CW.listItems !== "function") return PHASES_SET;
  try {
    const items = await CW.listItems(PHASE_LIST);
    if (items == null) { PHASE_LIST_OK = false; PHASES_SET = {}; return PHASES_SET; }
    PHASE_LIST_OK = true;
    const next = {};
    items.forEach(it => {
      const f = it.fields || {};
      const job = String(f.Title == null ? "" : f.Title).trim().toUpperCase();
      /* a blank cell is not a zero: Number("") is 0, and a row with a job but
         no phase on it is a half-filled row, not "In office" */
      const raw = f.Phase;
      const n = raw == null || String(raw).trim() === "" ? NaN : Number(raw);
      if (!job || !isFinite(n)) return;                    // a half-filled row is not a phase
      const at = String(f.SetAt == null ? "" : f.SetAt);
      /* Two browsers can add the same job at the same moment, and a list
         without the unique-Title rule will hold both. The newest SetAt is what
         somebody most recently decided, so that is the one that counts; the
         next write to this job clears the older ones away for good. */
      const seen = next[job];
      if (seen && seen.at >= at) return;                   // ISO stamps: newest sorts last
      next[job] = { phase: Math.max(0, Math.min(PHASES.length - 1, Math.round(n))),
                    name: String(f.PhaseName == null ? "" : f.PhaseName),
                    who: String(f.SetBy == null ? "" : f.SetBy),
                    at: at };
    });
    PHASES_SET = next;
  } catch (e) {
    PHASE_LIST_CONSENT = /permission needed/.test((e && e.message) || "");   // a click may ask for it
    if (PHASE_LIST_CONSENT) { PHASE_LIST_OK = false; }                         // known state, not an error: the drawer explains it
    else if (!phaseWarned) { phaseWarned = true; toast("Could not read the hand-set phases: " + friendly(e), true); }
  }
  return PHASES_SET;
}

/** The hold this browser is keeping on a phase it has just written.
    undefined = no hold · null = held as cleared · 0-6 = held at that phase. */
function pendPhase(id) {
  const p = PENDING[id];
  if (!p || p.phase == null) return undefined;
  return p.phase < 0 ? null : p.phase;
}
/** The hand-set phase for a job, or null. Our own click wins while it is held,
    so a stale read of the list cannot undo what was just done here. */
function handPhase(j) {
  if (!j) return null;
  const held = pendPhase(j.id);
  if (held !== undefined) return held;
  const s = PHASES_SET[String(j.id).toUpperCase()];
  return s ? s.phase : null;
}
/** Who set it and when, for the note under the pipeline. Not shown while our
    own hold is the only thing saying so. */
function handPhaseInfo(j) {
  if (!j) return null;
  const s = PHASES_SET[String(j.id).toUpperCase()];
  return s && handPhase(j) === s.phase ? s : null;
}
if (typeof setPhaseHook === "function") setPhaseHook(handPhase);

/** Set the phase of one job by hand, or - clicking the step it is already
    on - clear it and go back to the sheet. One list item per job; one log
    line per change; no workbook write of any kind. */
async function setPhaseByHand(j, n) {
  if (!j || PHASEBUSY[j.id]) return false;
  if (PHASE_LIST_OK === null || PHASE_LIST_CONSENT) {      // first read pending, or it failed for want of permission:
    try { if (CW.listConsent) await CW.listConsent(); } catch (e) { toast(friendly(e), true); return false; }
    await readPhases();                                     // a click is the one place the consent popup may open
  }
  if (PHASE_LIST_OK !== true) { toast(PHASE_LIST_CONSENT ? PHASE_NEED_CONSENT : PHASE_LIST_MISSING, true); return false; }
  n = Number(n);
  if (!isFinite(n) || n < 0 || n >= PHASES.length) return false;
  const sheet = jobPhase(j);
  /* the sheet's own evidence already puts it past this step: nothing to set */
  if (n < sheet) { toast(PHASE_BELOW_SHEET); return false; }
  const before = PHASES[effectivePhase(j)];
  const clearing = handPhase(j) === n;
  PHASEBUSY[j.id] = true;
  if (state.sel === j.id) renderDrawer();
  try {
    if (clearing) {
      await CW.listDelete(PHASE_LIST, j.id);
      delete PHASES_SET[String(j.id).toUpperCase()];
      pend(j.id, { phase: -1 });                       // held as cleared for 180 s
      ALL = applyPending(ALL);
      noteChange(j.id, "Phase", before, "sheet");
      toast(j.id + " phase cleared \u2014 back to what the sheet shows (" + PHASES[jobPhase(j)] + ")");
    } else {
      const who = whoAmI(), at = new Date().toISOString();
      await CW.listUpsert(PHASE_LIST, j.id, { Phase: n, PhaseName: PHASES[n], SetBy: who, SetAt: at });
      PHASES_SET[String(j.id).toUpperCase()] = { phase: n, name: PHASES[n], who: who, at: at };
      pend(j.id, { phase: n });
      ALL = applyPending(ALL);
      noteChange(j.id, "Phase", before, PHASES[n]);
      toast(j.id + " phase set to " + PHASES[n]);
    }
  } catch (e) {
    pend(j.id, { phase: null });                       // the write failed: drop the hold
    ALL = applyPending(ALL);
    toast(friendly(e), true);
    delete PHASEBUSY[j.id];
    renderAll(); if (state.sel === j.id) renderDrawer();
    return false;
  }
  delete PHASEBUSY[j.id];
  renderAll(); if (state.sel === j.id) renderDrawer();
  return true;
}

/* ---------- the glass station ------------------------------------------------
   The floor works on glass.html, signed in with a shared station account that
   has no access to the workbook at all. Everything it sees comes from one
   SharePoint list in a separate site, and the master dashboard is the only
   thing that puts job facts in there.

   HARD RULE, and the reason all of this lives in one block: the feeder writes
   Title, Job, Customer, GlassType, Total, Seq, Active, FedAt and FedBy, and -
   only on a row it is creating or a row the floor has never tapped (DoneAt
   empty) - the three counters, seeded from the office's own checkpoints so a
   job already ticked off in here does not arrive on the floor reading nothing
   done. It never sends a By, an At or the last-touch pair, it never touches a
   counter once the floor has tapped the row, and it never deletes an item. A
   job that leaves production is marked Active = No and keeps everything the
   floor recorded. This dashboard also
   reads two more of the floor's lists, Station log and Station people, and
   writes to neither. Not one line of this touches the workbook.             */
const STATION_SITE_MISSING = "The floor’s SharePoint site is not there yet, so their progress " +
  "cannot be shown here. Ask the manager to add it — nothing in the Excel file is involved.";
const STATION_LIST_MISSING = "The “Glass station” list is not in the floor’s site yet. " +
  "Ask the manager to add it — nothing in the Excel file is involved.";
const STATION_NEED_CONSENT = "Seeing the floor's progress needs a SharePoint permission that has not been granted yet. " +
  "Nothing in the Excel file is involved.";
const STATION_SIGN_IN_AGAIN = "The sign-in to SharePoint has expired. Sign out and sign in again to see the floor's progress.";
const STATION_LOG_MISSING = "The “Station log” list is not in the floor’s site yet, so who changed what " +
  "cannot be shown. Ask the manager to add it — nothing in the Excel file is involved.";
const STATION_UNREACHABLE = "cannot reach SharePoint — retrying";
const STATION_CHECKING = "checking…";

let STATION_ITEMS = null;      // the Glass station list, as last read - null until the first read answers
let STATION_LOG = null;        // the Station log list, likewise
let STATION_PEOPLE = null;     // the Station people list, read once for the log window's filters
let STATION_OK = null;         // null: not looked · false: no site, no list, or no permission · true: read it
let STATION_WHY = "";          // which of those, in words, for the board and the drawer
let STATION_ERR = "";          // a passing failure - the last board stays, with this line above it
let STATION_LOG_OK = null;     // the log list has its own three states: it can be missing on its own
let STATION_LOG_WHY = "";
let stationWarned = false;     // one console line per page for a list that will not read
let stationReading = null;     // the read in flight, so a drawer and the dropdown share one
let stationLogReading = null;

/* The two lists both screens poll by delta. One shape for each, so the poll
   below is one function rather than two nearly identical ones - and so a test
   can put a token back to null and watch the next call enumerate again. */
const STATION_FEEDS = {
  items: { list: () => ST.STATION_LIST, fields: () => ST.STATION_FIELDS,
           get: () => STATION_ITEMS, set: v => { STATION_ITEMS = v; }, token: null, off: 0 },
  log: { list: () => ST.LOG_LIST, fields: () => ST.LOG_FIELDS,
         get: () => STATION_LOG, set: v => { STATION_LOG = v; }, token: null, off: 0 }
};
/* A delta request is in the air for as long as SharePoint takes to answer it,
   and in that time a feed or a full read can replace the whole list and null
   the token underneath it. The answer that then comes back is two changed
   rows and no token in hand - which reads exactly like a fresh enumeration,
   and would collapse the board to those two rows until the next feed ten
   minutes later. So every wholesale replacement bumps a generation, and a
   delta whose generation has moved throws its answer away. */
let STATION_GEN = 0;
function stationResetFeed(key) {
  STATION_FEEDS[key].token = null;
  STATION_GEN++;
}
/* A list that will not serve a delta at all must not be asked for one six
   times a minute: it is marked off for five minutes and polled the plain way
   until then. A 410 "your token is too old" is NOT that - delta is working
   there, the token is simply stale - so it only resets the token. */
const DELTA_OFF_MS = 300000;
const deltaOff = f => !!(f.off && Date.now() - f.off < DELTA_OFF_MS);

/** Read the whole station list, quietly. Never pops a consent window, never
    toasts, never throws: a floor board nobody is looking at must not be able
    to interrupt the dashboard.

    A failure only clears the board when SharePoint actually says the site or
    the list is not there. Anything else - offline, a bad gateway, a refused
    token - leaves the last board on screen with a line saying it could not be
    reached, because a blank screen saying "ask the manager to make the list"
    is a lie when the list is fine and the wifi is not. */
async function readStation() {
  if (typeof CW === "undefined" || !CW || typeof CW.listItems !== "function") return null;
  if (typeof ST === "undefined") return null;
  try {
    if (CW.hasListConsent && !(await CW.hasListConsent())) {
      STATION_OK = false; STATION_WHY = STATION_NEED_CONSENT; STATION_ERR = ""; return null;
    }
    const siteId = await CW.stationSite();
    if (!siteId) { STATION_OK = false; STATION_WHY = STATION_SITE_MISSING; STATION_ERR = ""; return null; }
    const items = await CW.listItems(ST.STATION_LIST, { siteId: siteId, fields: ST.STATION_FIELDS });
    if (items == null) { STATION_OK = false; STATION_WHY = STATION_LIST_MISSING; STATION_ERR = ""; return null; }
    STATION_OK = true; STATION_WHY = ""; STATION_ERR = ""; STATION_ITEMS = items;
    stationResetFeed("items");              // a full read: the next poll starts a fresh delta
    return items;
  } catch (e) {
    stationTrouble(e);
    return null;
  }
}

/** The floor's log - who moved which counter, when. Read-only here, for ever:
    the office never writes a line and nothing anywhere deletes one. It is a
    separate list from the board, so it can be the one that is missing: that is
    a state of its own rather than a reason to hide the bars. */
async function readStationLog() {
  if (typeof CW === "undefined" || !CW || typeof CW.listItems !== "function") return null;
  if (typeof ST === "undefined") return null;
  try {
    if (CW.hasListConsent && !(await CW.hasListConsent())) {
      STATION_LOG_OK = false; STATION_LOG_WHY = STATION_NEED_CONSENT; return null;
    }
    const siteId = await CW.stationSite();
    if (!siteId) { STATION_LOG_OK = false; STATION_LOG_WHY = STATION_SITE_MISSING; return null; }
    const items = await CW.listItems(ST.LOG_LIST, { siteId: siteId, fields: ST.LOG_FIELDS });
    if (items == null) { STATION_LOG_OK = false; STATION_LOG_WHY = STATION_LOG_MISSING; return null; }
    /* nothing ever deletes from this list, so after a year it is thousands of
       lines of last spring. Neither screen looks past ninety days, so neither
       carries the rest around: a line with no stamp at all is kept, because
       dropping it would hide work rather than old work. */
    STATION_LOG = stationLogRecent(items);
    STATION_LOG_OK = true; STATION_LOG_WHY = "";
    stationResetFeed("log");
    return STATION_LOG;
  } catch (e) {
    stationTrouble(e);
    return null;
  }
}
/** The last ninety days of the log, by At. */
function stationLogRecent(items, now) {
  const since = ST.logSince(ST.LOG_DAYS, now);
  return (items || []).filter(it => {
    const at = String(((it && it.fields) || {}).At || "");
    return !at || at >= since;
  });
}

/** The people the floor picks from, read once - the log window's person filter
    offers everybody who may record something, not only everybody who has. A
    list that will not read answers an empty array rather than nothing, so the
    window falls back to the names in the log instead of asking again forever.
    Read WITHOUT the PIN column: the office has no business holding anybody's. */
let stationPeopleReading = null;
async function readStationPeople() {
  try {
    const siteId = await CW.stationSite();
    STATION_PEOPLE = (siteId &&
      await CW.listItems(ST.PEOPLE_LIST, { siteId: siteId, fields: ST.PEOPLE_FIELDS_OFFICE })) || [];
  } catch (e) {
    console.warn("[station] could not read the people list:", (e && e.message) || e);
    STATION_PEOPLE = [];
  }
  return STATION_PEOPLE;
}
function stationPeopleIfNeeded(then) {
  if (STATION_PEOPLE !== null) return;
  if (!stationPeopleReading)
    stationPeopleReading = readStationPeople().then(r => { stationPeopleReading = null; return r; },
                                                    () => { stationPeopleReading = null; });
  stationPeopleReading.then(() => { if (then) then(); });
}

/** What a thrown station error means, in one place, for the read and the feed. */
function stationTrouble(e) {
  const m = (e && e.message) || String(e || "");
  if (!stationWarned) { stationWarned = true; console.warn("[station] could not read the list:", m); }
  /* Only "there is no such thing here" is a reason to doubt the cached site.
     A refusal, a bad gateway or a dropped connection says nothing about where
     the lists are; dropping the site over one would restart the whole re-check
     cadence, and on a 403 could flip a dashboard that is happily reading the
     real site over to the fallback. */
  if (CW.isMissing && CW.isMissing(e) && CW.forgetStationSite) CW.forgetStationSite();
  if (/interaction_required|login_required/.test(m)) {
    STATION_OK = false; STATION_WHY = STATION_SIGN_IN_AGAIN; STATION_ERR = "";
  } else if (/permission needed/.test(m)) {
    STATION_OK = false; STATION_WHY = STATION_NEED_CONSENT; STATION_ERR = "";
  } else if (CW.isMissing && CW.isMissing(e)) {
    STATION_OK = false; STATION_WHY = STATION_LIST_MISSING; STATION_ERR = "";
  } else {
    STATION_ERR = STATION_UNREACHABLE;          // the last board stays exactly where it was
  }
}

/** Read the list once, if nobody has yet - for a drawer opening, or the board
    being picked, when the feeder's ten-minute skip means nothing has read it.
    One read is shared by every caller, so opening three drawers asks once. */
function stationReadIfNeeded(then) {
  /* already read: the caller has just drawn the current state, so there is
     nothing to redraw. Calling `then` here would have renderDrawer() call
     itself without end (it is the caller). */
  if (STATION_OK !== null) return;
  if (!stationReading) stationReading = readStation().then(r => { stationReading = null; return r; },
                                                           () => { stationReading = null; });
  stationReading.then(() => { if (then) then(); });
}
/** The same, for the log. Kept separate because the two lists fail separately:
    the board can be perfectly readable while the log list has not been made. */
function stationLogReadIfNeeded(then) {
  if (STATION_LOG_OK !== null) return;
  if (!stationLogReading) stationLogReading = readStationLog().then(r => { stationLogReading = null; return r; },
                                                                   () => { stationLogReading = null; });
  stationLogReading.then(() => { if (then) then(); });
}

/* ---- real time --------------------------------------------------------------
   A tap on the floor is meant to be on the colleague's screen inside ten
   seconds, and re-reading two whole lists six times a minute to find out that
   nothing moved would be six hundred rows a minute of unchanged text. So the
   poll asks Graph what has changed since last time (listDelta) and merges it.

   Ten seconds while somebody is actually looking at the floor - the station
   board, the log window, or a drawer for a job that has glass - and a minute
   otherwise, because a poll nobody is reading is only there to keep the drawer
   honest when it is next opened. A failure never toasts: it leaves the last
   data exactly where it is and says so in the line the board already has.   */
const STATION_FAST_MS = 10000, STATION_SLOW_MS = 60000;
let stationPollT = null, stationPolling = false;
let STATION_SITE_GEN = 0;               // which site the tokens in hand belong to

/* What the job list on screen is showing of the floor's work, worked out when
   the rows are drawn and remembered until they are drawn again. Three things
   are asked of it afterwards, all of them often enough to matter:

     · is anything on screen worth polling six times a minute for (ROWS_GLASS);
     · has anything the rows are showing actually changed (ROWS_CHIPS - a poll
       that moved a job nobody is looking at, or wrote a log line, must not
       rebuild seven hundred rows to draw exactly what is already there);
     · is a repaint owed, because the list was in use when one came round.

   ROWS_DRAWN is the rows the list last drew, so the second question can be
   asked without filtering and sorting the whole sheet again. */
let ROWS_GLASS = false, ROWS_CHIPS = "", ROWS_DRAWN = [], ROWS_STALE = false, ROWS_QUIET = false;

/** What the chips of the rows on screen say, as one string to compare against.
    A record with no glasses on it is in here too, as nothing: it draws no chip
    today, but a total arriving is a change the list has to show. */
function chipsNow() {
  let s = "";
  for (let i = 0; i < ROWS_DRAWN.length; i++) {
    const j = ROWS_DRAWN[i], g = stationForJob(j.id);
    if (g) s += j.id + ":" + g.cut + "," + g.hotmelt + "," + g.glazed + "/" + g.total + "|";
  }
  return s;
}
/** A repaint the list was too busy for is owed, not lost: the poll's own clock
    takes it as soon as whoever was typing or dragging has finished, so the rows
    can never sit on a number the floor has moved on from with nothing to say
    so. Cheap when nothing is owed, which is nearly always. */
function stationCatchUp() { if (ROWS_STALE) redrawStation(); }

/** A repaint nobody asked for: no entry animation, so the rows change their
    numbers where they stand instead of the whole list flashing. */
function quietRows() {
  ROWS_QUIET = true;
  try { renderRows(); } finally { ROWS_QUIET = false; }
}

/** Is anyone actually looking at the floor's data right now? */
function stationWatching() {
  /* the John print sheet is a board in the same slot, but it is the office's
     own second sheet: nothing on it can change because the floor tapped
     something, so it is no reason to poll the floor six times a minute */
  if (state.board && state.board !== "john") return true;
  if ($("#lhost")) return true;
  const j = state.sel ? byId(state.sel) : null;
  if (j && typeof ST !== "undefined" && ST.glassTotal(j) > 0) return true;
  /* the ordinary job list counts too, now that its rows carry the floor's
     chip: a tap on the tablet changes what is on this screen. Only when the
     rows being shown are jobs the floor has a record of - a list filtered
     down to jobs with no glass has nothing to wait for. */
  return ROWS_GLASS;
}
/** The plain read, for a list that will not serve a delta and for a token
    that has gone stale. true = the list was replaced. */
async function stationFull(key, siteId, gen) {
  const f = STATION_FEEDS[key];
  const all = await CW.listItems(f.list(), { siteId: siteId, fields: f.fields() });
  if (gen !== STATION_GEN) return false;           // somebody replaced it while we read
  if (all == null) return false;
  f.set(key === "log" ? stationLogRecent(all) : all);
  return true;
}
/** One list, brought up to date the cheap way. true = something moved. */
async function stationDelta(key, siteId) {
  const f = STATION_FEEDS[key];
  if (deltaOff(f)) return await stationFull(key, siteId, STATION_GEN);
  const opts = { siteId: siteId, fields: f.fields() };
  if (f.token) opts.token = f.token;
  /* BOTH captured before the await. f.token can be nulled by a feed or a full
     read while this request is in the air; reading it afterwards would make a
     delta of two changed rows look like a fresh enumeration of the whole
     list, and the board would collapse to two rows. */
  const had = opts.token || null;
  let gen = STATION_GEN;
  let d;
  try {
    d = await CW.listDelta(f.list(), opts);
  } catch (e) {
    if (!CW.isDeltaRestart || !CW.isDeltaRestart(e)) throw e;
    if (gen !== STATION_GEN) return false;         // that answer is about a list we no longer hold
    /* a 410 means the token is too old and delta is fine; anything else means
       this list will not serve one, so stop asking for five minutes */
    if (!(CW.isDeltaResync && CW.isDeltaResync(e))) {
      f.off = Date.now();
      console.warn("[station] the " + f.list() + " list refused a delta; polling it the plain way for five minutes");
    }
    stationResetFeed(key);
    gen = STATION_GEN;
    return await stationFull(key, siteId, gen);
  }
  if (gen !== STATION_GEN) return false;           // a feed or a full read landed while this was in the air
  if (d == null) return false;                     // the list is not there
  f.off = 0;                                       // it served one: it is not a refusing list
  f.token = d.next || null;
  if (!had) {                                      // the first pass enumerates the lot
    const rows = d.items.filter(x => !x.removed).map(x => ({ id: x.id, fields: x.fields }));
    f.set(key === "log" ? stationLogRecent(rows) : rows);
    return true;
  }
  if (!d.items.length) return false;
  f.set(ST.mergeDelta(f.get() || [], d.items));
  return true;
}
async function stationPoll() {
  if (stationPolling || stationBusy) return false;
  if (typeof ST === "undefined" || typeof CW === "undefined" || !CW || !CW.listDelta) return false;
  if (STATION_OK !== true) return false;           // nothing read yet: there is nothing to keep current
  stationPolling = true;
  try {
    if (!CW.hasListConsent || !(await CW.hasListConsent())) return false;
    const siteId = await CW.stationSite();
    if (!siteId) return false;
    /* the floor's lists can move from one site to another (see stationSite()).
       A delta token only means anything in the site it was issued in, so a
       move throws both of them away and starts again. */
    const gen = CW.stationSiteMoves ? CW.stationSiteMoves() : 0;
    if (gen !== STATION_SITE_GEN) { stationSiteMoved(gen); }
    let moved = await stationDelta("items", siteId);
    if (STATION_LOG_OK === true && (await stationDelta("log", siteId))) moved = true;
    STATION_ERR = "";
    if (moved) {
      redrawStation();
      /* a tap on the floor is what this feature exists to carry into the sheet,
         and this is the ten-second clock that hears about it. Not awaited: the
         poll must be back for its next turn whatever the workbook is doing, and
         the write reports its own failures. */
      glassColourRun().catch(e => console.warn("[glass] " + ((e && e.message) || e)));
    }
    return moved;
  } catch (e) {
    stationTrouble(e);
    redrawStation();
    return false;
  } finally {
    stationPolling = false;
  }
}
/** The floor's lists have moved to another site. Everything this dashboard is
    holding about them was true of the old one: the delta tokens, the "already
    fed, nothing changed" hash, and the roster the log window filters by. All
    of it goes, so the next load feeds the new site and the next window shows
    the people who are in it. */
function stationSiteMoved(gen) {
  STATION_SITE_GEN = gen;
  stationResetFeed("items"); stationResetFeed("log");
  STATION_FEEDS.items.off = 0; STATION_FEEDS.log.off = 0;
  STATION_FEED = { hash: "", at: 0 }; saveStationFeed();
  STATION_PEOPLE = null; stationPeopleReading = null;
  console.log("[station] the floor's lists have moved: feeding and re-reading the new site");
}

/** Is somebody in the middle of something the list must not be rebuilt under?
    Two things, and they are the same rule the filter bar has always had: a poll
    may not take something out from under somebody's hands. One is text being
    typed inside the list itself - the grouped view keeps a search box in every
    section heading. The other is a row being dragged, which a rebuild drops.

    A tick box is NOT one of them, and this is the whole reason the question is
    a function rather than "is anything in the list focused". Every row has a
    checkbox; ticking one leaves it focused and does not redraw anything, so
    treating focus as busy would stop the chips ever moving again for as long
    as the tick stood - and a stale number with nothing to say it is stale is
    exactly what this feature exists to stop. A repaint keeps what is ticked. */
function rowsInUse() {
  const rows = $("#rows");
  if (!rows) return false;
  if (rows.querySelector && rows.querySelector(".row.dragging")) return true;
  const a = document.activeElement;
  if (!a || !rows.contains || !rows.contains(a)) return false;
  if (a.isContentEditable) return true;
  const tag = String(a.tagName || "").toUpperCase();
  if (tag === "TEXTAREA") return true;
  if (tag !== "INPUT") return false;
  /* an input is text unless it says otherwise, which is what the DOM does too */
  return ["checkbox", "radio", "button", "submit", "reset", "range", "file", "image", "color"]
    .indexOf(String(a.type || "text").toLowerCase()) < 0;
}

/** Redraw only the parts of the dashboard that show the floor's data. */
function redrawStation() {
  /* the rows whatever is on screen, not only the board: an ordinary job row
     carries the floor's chip too, and a counter moved on the tablet has to
     reach it. Held off while the list is in use - and then owed, so the next
     tick takes the repaint that was skipped rather than losing it. */
  const onBoard = typeof STATIONS !== "undefined" && STATIONS.some(b => b[0] === state.board);
  if (rowsInUse()) ROWS_STALE = true;
  else if (onBoard) { ROWS_STALE = false; quietRows(); }
  /* on the plain list, only when what the rows are SAYING has changed. A poll
     that moved a job nobody is looking at, or that only wrote a log line, must
     not rebuild seven hundred rows to draw exactly what is already there - and
     a board that is not a floor station (the office's own print sheet) shows
     nothing of the floor's at all. */
  else if (!state.board && chipsNow() !== ROWS_CHIPS) { ROWS_STALE = false; quietRows(); }
  else ROWS_STALE = false;
  if (state.sel && $("#dhost")) renderDrawer();
  /* the rows and the counts, never the filter bar: a poll landing while
     somebody is half way through typing a job number must not take the box
     out from under them */
  if ($("#lhost")) paintStationLog();
}
/** The poll's own clock. Re-armed after every pass, so changing what is on
    screen changes the rate at the next tick rather than needing two timers. */
let STATION_TICK_MS = 0;                 // the rate currently armed, for the tests
function stationTick() {
  if (stationPollT) clearTimeout(stationPollT);
  STATION_TICK_MS = stationWatching() ? STATION_FAST_MS : STATION_SLOW_MS;
  stationPollT = setTimeout(async () => {
    stationPollT = null;
    /* both of them inside the one guard. stationCatchUp() reaches renderRows(),
       renderDrawer() and paintStationLog(), and it runs in exactly the unusual
       states - a search box in hand, a drag in flight. A throw out here would
       leave stationPollT null with stationTick() never reached: no timer would
       ever be armed again, and the whole dashboard would quietly stop updating
       until somebody reloaded the page. */
    try { await stationPoll(); stationCatchUp(); } catch (e) { /* neither throws; belt and braces */ }
    stationTick();
  }, STATION_TICK_MS);
}

/* ---- the feeder ----
   After every successful load(), the glass slice of the sheet is pushed into
   the list. It is deliberately dull: skip unless the permission is already
   granted, skip when nothing has changed and the last run is recent, read the
   list once, work out the plan (station-core.js), send it a few at a time, and
   remember what was sent. Failures go to the console and to one small word in
   the footer; they never toast and never block anything. */
const STATION_FEED_KEY = "cw_stationfeed";
const STATION_FEED_MS = 600000;        // an unchanged slice is re-fed at most every 10 minutes
const STATION_FEED_MAX = 60;           // writes per run; the rest goes next run
const STATION_FEED_LANES = 3;          // how many writes are in flight at once
let STATION_FEED = { hash: "", at: 0 };
try { STATION_FEED = JSON.parse(localStorage.getItem(STATION_FEED_KEY) || "null") || STATION_FEED; } catch (e) {}
let stationBusy = false;
let STATION_FEED_ERR = "";

function saveStationFeed() { try { localStorage.setItem(STATION_FEED_KEY, JSON.stringify(STATION_FEED)); } catch (e) {} }
/* The separator lives inside the wrapper, so an empty station word does not
   leave a stray middle dot sitting in the footer. */
function setStationFoot() {
  const el = $("#stationfeed"); if (!el) return;
  const wrap = $("#stationfeedwrap");
  const show = t => { el.textContent = t; if (wrap) wrap.hidden = !t; };
  /* first, because it is the more serious of the two: a feed that will not
     write means the floor's board goes stale, while this means the WORKBOOK is
     refusing a write and somebody has to know rather than read it in a console */
  if (GLASS_COLOUR_ERR) { show("glass colours not saved"); el.title = GLASS_COLOUR_ERR; return; }
  if (STATION_FEED_ERR) { show("station feed failed"); el.title = STATION_FEED_ERR; return; }
  if (!STATION_FEED.at) { show(""); el.title = ""; return; }
  show("station feed: " + agoWords(STATION_FEED.at));
  el.title = "The Glass station list was last brought up to date then. The Excel file is not involved.";
}
/** "just now" / "3 min ago" / "2 h ago", from a millisecond stamp or an ISO one. */
function agoWords(at) {
  const t = typeof at === "number" ? at : Date.parse(at);
  if (!t || !isFinite(t)) return "never";
  const s = Math.max(0, Math.round((Date.now() - t) / 1000));
  if (s < 45) return "just now";
  if (s < 5400) return Math.round(s / 60) + " min ago";
  if (s < 172800) return Math.round(s / 3600) + " h ago";
  return Math.round(s / 86400) + " days ago";
}

/** Run the writes a few lanes at a time, in order. Every failure is counted
    and the first message is kept: the run finishes what it can, and whatever
    it lost is planned again on the next run rather than retried in a tight
    loop here. */
async function stationSend(jobsToDo) {
  let sent = 0, failed = 0, err = "";
  let next = 0;
  const lane = async () => {
    for (;;) {
      const i = next++;
      if (i >= jobsToDo.length) return;
      try { await jobsToDo[i](); sent++; }
      catch (e) { failed++; if (!err) err = (e && e.message) || String(e); }
    }
  };
  const lanes = [];
  for (let i = 0; i < Math.min(STATION_FEED_LANES, jobsToDo.length); i++) lanes.push(lane());
  await Promise.all(lanes);
  return { sent: sent, failed: failed, err: err };
}

/** The name that goes in FedBy: whose dashboard fed the list. A display name
    when the account has one, the local part of the address when it does not -
    never the whole address, which the floor has no business reading. */
function feedWho() {
  const a = CW.account;
  const name = a && a.name ? String(a.name).trim() : "";
  if (name) return name;
  return String((a && a.username) || "").split("@")[0] || "dashboard";
}

/** Add one row, and treat a refusal the way listUpsert does: the Title is
    unique on the list, so a refused POST usually means the row is already
    there - look for it and patch it rather than losing the update. */
async function stationAdd(fields, opts) {
  try { return await CW.listAdd(ST.STATION_LIST, fields, opts); }
  catch (e) {
    const mine = await CW.listItemsFor(ST.STATION_LIST, fields.Title, opts);
    if (!mine.length) throw e;                 // a real failure: nothing was created
    /* it exists after all - another dashboard made it between our read and
       this write - so it is not a new row and the seeding rule applies to it
       like any other: if the floor has tapped it, its counters are not ours. */
    const have = mine[0].fields || {};
    const touched = String(have.DoneAt == null ? "" : have.DoneAt).trim() !== "";
    const body = {};
    Object.keys(fields).forEach(k => {
      if (touched && ST.SEED_FIELDS.indexOf(k) >= 0) return;
      body[k] = fields[k];
    });
    return CW.listPatch(ST.STATION_LIST, mine[0].id, body, opts);
  }
}

/** One patch of the feed's plan.

    A plan is made from ONE read of the list and then sent as up to sixty
    writes over three lanes, so a tap from the floor can land between the read
    and the write that was planned from it. Every patch that carries a seeded
    counter therefore re-checks that one row immediately before sending: if
    `DoneAt` has appeared in the meantime the floor has taken the row over and
    the counters are dropped, exactly as the refused-POST path in stationAdd
    already does. The job facts still go - what the job IS is still ours.

    The same GET answers the other question worth asking: whether this
    dashboard's copy of the workbook is older than the feed that last wrote
    this row. If it is, our seed is a reading of a stale file and would park
    the row on older numbers, so it is dropped too. */
async function stationFeedPatch(p, opts) {
  const fields = p.fields || {};
  if (!ST.SEED_FIELDS.some(k => k in fields))
    return CW.listPatch(ST.STATION_LIST, p.id, fields, opts);
  let now = null;
  try { now = await CW.listItem(ST.STATION_LIST, p.id, opts); }
  catch (e) { now = null; }                 // could not look: assume the worst and keep off the counters
  const have = (now && now.fields) || {};
  const touched = !now || String(have.DoneAt == null ? "" : have.DoneAt).trim() !== "";
  const fedAt = Date.parse(String(have.FedAt == null ? "" : have.FedAt));
  const mine = Date.parse(String(lastStamp || ""));
  const older = isFinite(fedAt) && isFinite(mine) && fedAt > mine;
  if (!touched && !older) return CW.listPatch(ST.STATION_LIST, p.id, fields, opts);
  const body = {};
  Object.keys(fields).forEach(k => { if (ST.SEED_FIELDS.indexOf(k) < 0) body[k] = fields[k]; });
  /* FedAt/FedBy alone are not worth a write: they only ride along with a
     change, and the change is exactly what was just dropped */
  if (!Object.keys(body).some(k => k !== "FedAt" && k !== "FedBy")) return null;
  return CW.listPatch(ST.STATION_LIST, p.id, body, opts);
}

/* A run that could not finish - the 60-write cap, or a write that was refused
   - comes back in half a minute rather than waiting for the next load. Once:
   the timer is only ever armed when none is already waiting. */
const STATION_AGAIN_MS = 30000;
let stationAgainT = null;
function stationFeedAgain() {
  if (stationAgainT) return;
  stationAgainT = setTimeout(() => { stationAgainT = null; feedStation().catch(() => {}); }, STATION_AGAIN_MS);
}

/** What the office has already ticked off on a job's glass, as plain data for
    ST.officeSeed - one entry per glass item with the sheet's own word for it
    and the office's count where there is one. The reading of the colour and
    the stored count is checkpoints.js' job (`cpStatus`, `itemState`, which is
    the merge of `cpStatus` with `cpStored`); station-core.js is handed the
    answer rather than any of the dashboard's globals.

    Nothing here writes anything: this is the workbook's own record, read. */
function glassCounts(j) {
  if (typeof cpStatus !== "function" || !j) return [];
  const glass = j.glass || {};
  const out = [];
  Object.keys(glass).forEach(k => {
    const total = Math.round(Number(glass[k]) || 0);
    if (!(total > 0)) return;
    const key = "glass:" + k;
    const s = typeof itemState === "function" ? itemState(j, key) : null;
    out.push({ type: k, total: total, status: cpStatus(j, key),
               done: s && s.done != null ? s.done : 0 });
  });
  return out;
}

async function feedStation() {
  if (stationBusy) return null;                          // only one feed at a time
  /* a poll is merging a delta into STATION_ITEMS right now. Feeding would
     replace the list and null the token under it, so it waits for the next
     load or the follow-up timer rather than pulling the rug. */
  if (stationPolling) { stationFeedAgain(); return null; }
  if (typeof ST === "undefined" || typeof CW === "undefined" || !CW || !CW.listAdd) return null;
  /* the flag goes up before the first await: two loads finishing together
     would otherwise both get past the guard and feed the list twice */
  stationBusy = true;
  try {
    if (!CW.hasListConsent || !(await CW.hasListConsent())) { STATION_FEED_ERR = ""; return null; }
    /* the slice carries the office's own record with it, so a job the office
       has already ticked off arrives on the floor showing it (see
       ST.officeSeed) rather than reading nothing done */
    const slice = ST.glassSlice(ALL, BLOCKNAMES, glassCounts);
    const hash = ST.sliceHash(slice);
    if (hash === STATION_FEED.hash && Date.now() - (STATION_FEED.at || 0) < STATION_FEED_MS) {
      STATION_FEED_ERR = "";                             // nothing to do is not a failure
      return null;
    }
    const siteId = await CW.stationSite();
    if (!siteId) { STATION_OK = false; STATION_WHY = STATION_SITE_MISSING; STATION_ERR = ""; return null; }
    /* the feed resolves the site as well, so it is one of the two places a
       move can first be noticed - and it must not feed the new site holding
       the old one's tokens */
    const fgen = CW.stationSiteMoves ? CW.stationSiteMoves() : 0;
    if (fgen !== STATION_SITE_GEN) stationSiteMoved(fgen);
    const opts = { siteId: siteId, fields: ST.STATION_FIELDS };
    const items = await CW.listItems(ST.STATION_LIST, opts);
    if (items == null) { STATION_OK = false; STATION_WHY = STATION_LIST_MISSING; STATION_ERR = ""; return null; }
    STATION_OK = true; STATION_WHY = ""; STATION_ERR = ""; STATION_ITEMS = items;
    stationResetFeed("items");             // a full read: the next poll starts a fresh delta
    const plan = ST.feedPlan(slice, items, { at: new Date().toISOString(), by: feedWho() });
    const all = plan.adds.map(f => () => stationAdd(f, opts))
      .concat(plan.patches.map(p => () => stationFeedPatch(p, opts)));
    const work = all.slice(0, STATION_FEED_MAX);
    const r = await stationSend(work);
    STATION_FEED_ERR = r.failed ? r.failed + " write" + (r.failed > 1 ? "s" : "") + " refused: " + r.err : "";
    /* the hash is only remembered when the whole plan went out: a run that was
       cut short by the 60-write cap, or that lost a write, must run again */
    const whole = !r.failed && work.length === all.length;
    STATION_FEED = whole ? { hash: hash, at: Date.now() } : { hash: "", at: Date.now() };
    saveStationFeed();
    if (!whole) stationFeedAgain();
    console.log("[station] fed " + r.sent + " of " + all.length +
                " (" + plan.adds.length + " new, " + plan.patches.length + " changed, " +
                plan.unchanged + " already right)");
    /* one more read, only when something actually went out, so the board and
       the drawer show what the list now says rather than what it said before */
    if (r.sent) {
      const after = await CW.listItems(ST.STATION_LIST, opts);
      if (after) { STATION_ITEMS = after; stationResetFeed("items"); }
    }
    return r;
  } catch (e) {
    STATION_FEED_ERR = (e && e.message) || String(e);
    console.warn("[station] feed failed:", STATION_FEED_ERR);
    stationTrouble(e);
    stationFeedAgain();
    return null;
  } finally {
    stationBusy = false;
    setStationFoot();
  }
}

/* Every job the floor has a row for, in one map, built once per version of the
   list. The job list asks about every row it draws - hundreds of questions,
   several times a minute once the floor is working - and answering each one by
   walking the whole list is the same list walked hundreds of times. The array
   itself is the cache key, exactly as the log's parsed rows are: every path
   that changes STATION_ITEMS replaces the array (a read assigns a new one,
   mergeDelta returns a new one), so an identity that has not moved is a list
   that has not moved. */
let STRECS = null, STRECS_OF = false;
function stationRecords() {
  if (STRECS_OF === STATION_ITEMS) return STRECS;
  STRECS_OF = STATION_ITEMS;
  STRECS = (STATION_ITEMS && typeof ST !== "undefined" && ST.jobRecords)
    ? ST.jobRecords(STATION_ITEMS) : {};
  return STRECS;
}

/** What the dashboard does once the feeder has had its go, whatever screen is
    open. Two things, and the second one is easy to think unnecessary:

    The rows are drawn again, because the plain job list carries the floor's
    chips now - not only the board and the drawer.

    And the list is READ if nothing has read it. The feeder normally leaves it
    read as a side effect of feeding, but it skips a slice it already fed less
    than ten minutes ago, and a skip reads nothing at all: sign in, look at the
    chips, press F5 six minutes later, and without this the rows would carry no
    chips for the rest of that session - which is the very complaint the chip
    was built to answer. The read is the shared one the board and the drawer
    ask for: once per session, whoever asks, and it answers at once thereafter. */
function stationAfterFeed() {
  if (state.board || state.sel) { renderAll(); if (state.sel) renderDrawer(); }
  else renderRows();
  stationReadIfNeeded(() => { if (!state.board) renderRows(); });
}

/** What the floor has on one job, whether it is still on their board or not.
    null = the job has never been fed. */
function stationForJob(id) {
  if (!STATION_ITEMS || typeof ST === "undefined") return null;
  if (!ST.jobRecords || !ST.jobKey) return ST.jobRecord(STATION_ITEMS, id);
  return stationRecords()[ST.jobKey(id)] || null;
}

/* ---------- load ---------- */
/* SharePoint needs about 35 s to put our change into the downloadable file, so
   every write asks for a re-read afterwards. One timer for all of them: a run
   of edits should re-read the file once, not once per edit. */
let reconcileT = null;
function scheduleReconcile(ms) {
  if (reconcileT) clearTimeout(reconcileT);
  reconcileT = setTimeout(() => { reconcileT = null; load("checking…", true); }, ms == null ? 45000 : ms);
}

async function load(reason, force) {
  if (busy && !force) return;
  busy = true;
  setStatus(reason || "reading sheet…", "busy");
  const t0 = performance.now();
  try {
    const wb = await CW.downloadWorkbook();
    LASTWB = wb;                                   // row formatting templates for moves
    const tDown = performance.now() - t0;
    const prev = ALL;
    const parsed = parseWorkbook(wb);
    BLOCKNAMES = parsed.blockNames || [];
    /* "Production (2)" on its own terms: John's own sheet, read straight out
       of the same download and never merged into the job model above. It is
       what the John print sheet view shows and what a John print prints. */
    JOHNROWS = parseJohnSheet(wb);
    /* the checkpoint counts, read before the held ticks are applied: a hold is
       let go the moment the file agrees with it, and that comparison needs the
       counts from this download, not the ones from the last */
    const pgw = wb.getWorksheet(CW.PROGRESS_SHEET);
    if (pgw) {
      const counts = {};
      for (let r = 2; r <= (pgw.rowCount || 0); r++) {
        const row = pgw.getRow(r);
        const cell = i => {
          const v = row.getCell(i).value;
          if (v == null) return "";
          if (v instanceof Date) {           // Excel may have taken the stamp for a date
            const q = n => (n < 10 ? "0" : "") + n;
            return v.getFullYear() + "-" + q(v.getMonth() + 1) + "-" + q(v.getDate()) +
                   " " + q(v.getHours()) + ":" + q(v.getMinutes());
          }
          return v.text != null ? v.text : String(v);
        };
        const job = cell(1).trim().toUpperCase(), item = cell(2).trim();
        if (!job || !item) continue;
        (counts[job] = counts[job] || {})[item] =
          { done: Number(cell(3)) || 0, total: Number(cell(4)) || 0, who: cell(5).trim(), when: cell(6).trim() };
      }
      cpSetProgress(counts);
    } else cpSetProgress({});          // no sheet, no counts - never the last download's

    /* the hand-set phases, from the SharePoint list - never from the workbook.
       Read before the holds are settled below, so a hold can be let go the
       moment the list agrees with it. */
    await readPhases();

    ALL = applyPending(parsed, true);   // our own recent writes win over a stale file
    ALL.blockNames = BLOCKNAMES;
    const ps = wb.getWorksheet("Production");
    PRODMAP = ps ? mapSheet(ps) : null;
    const m = await CW.lastModified();
    lastStamp = m.at;
    const t = new Date(m.at);
    const tAll = Math.round(performance.now() - t0);
    console.log("[dashboard] download " + Math.round(tDown) + "ms, total " + tAll + "ms, " + ALL.length + " jobs");
    const held = Object.keys(PENDING).length;
    setStatus("live · updated " + t.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) +
              (held ? " · " + held + " just changed" : " · " + (tAll / 1000).toFixed(1) + "s"));
    $("#srcinfo").textContent = "Production sheet · last edited by " + (m.by || "unknown");
    BLOCKNAMES = ALL.blockNames || parseWorkbook.__names || [];
    /* saved views: decisions only, never job data - so they cannot go stale */
    const vw = wb.getWorksheet("Dashboard Views");
    if (vw) {
      const next = {};
      for (let r = 2; r <= (vw.rowCount || 0); r++) {
        const row = vw.getRow(r);
        const t = i => { const v = row.getCell(i).value; return v == null ? "" : (v.text != null ? v.text : String(v)); };
        const view = t(1).trim(), job = t(2).trim().toUpperCase();
        if (!view || !job) continue;
        (next[view] = next[view] || {})[job] = { group: t(3).trim(), order: Number(t(4)) || 0 };
      }
      VIEWS = next;
    }
    applyPendV();
    /* who is the administrator, and who is subscribed to what. Both come out
       of the workbook on every read; nothing about them is kept in this
       browser except the 180 s hold on a change just made here. */
    readAlertSheets(wb);
    if ($("#ahost")) renderAlertsWindow();      // an open Alerts window must not go stale
    /* the shared history lives in the workbook, so everyone sees the same list */
    const logWs = wb.getWorksheet("Dashboard Log");
    if (logWs) {
      const fromSheet = [];
      for (let r = 2; r <= (logWs.rowCount || 0); r++) {
        const row = logWs.getRow(r);
        const cell = i => {
          const v = row.getCell(i).value;
          if (v == null) return "";
          if (v instanceof Date) {
            const p = n => (n < 10 ? "0" : "") + n;
            return p(v.getDate()) + "/" + p(v.getMonth() + 1) + "/" + v.getFullYear() +
                   " " + p(v.getHours()) + ":" + p(v.getMinutes());
          }
          return v.text != null ? v.text : String(v);
        };
        if (!cell(3)) continue;
        fromSheet.push({ at: cell(1), who: cell(2), job: cell(3), what: cell(4),
                         from: cell(5), to: cell(6), src: "dashboard", shared: true });
      }
      fromSheet.reverse();
      const inSheet = new Set(fromSheet.map(c => c.job + "|" + c.what + "|" + c.to));
      const local = CHANGES.filter(c => !c.shared && !inSheet.has(c.job + "|" + c.what + "|" + c.to));
      CHANGES = fromSheet.concat(local);
      saveChanges();
    }
    if (prev.length) {
      let d = diffJobs(prev, ALL, m.by || "someone in Excel", m.at);
      d = dropMine(d, CHANGES);   // a dashboard edit also shows up as a cell difference

      if (d.length) {
        CHANGES = d.concat(CHANGES); saveChanges();
        toast(d.length + " change" + (d.length > 1 ? "s" : "") + " — click Changes to see them");
      }
    }
    updateChangeBtn();
    renderAll();
    if (state.sel) renderDrawer();
    /* the floor's list, brought up to date from the sheet we have just read.
       Deliberately last, deliberately not awaited for the render above, and
       deliberately unable to fail loudly: the dashboard is finished by here. */
    /* and then, from the list the feed has just brought up to date, whatever
       the floor has finished since this dashboard last looked, painted into
       the four glass columns. Last, after the feed, because the feed is what
       puts STATION_ITEMS in front of it. */
    feedStation().then(stationAfterFeed, () => {})
      .then(() => glassColourRun(), () => {})
      .catch(e => console.warn("[glass] " + ((e && e.message) || e)));
  } catch (e) {
    /* A station account signing in here has no access to the workbook at all.
       That is not a fault to retry: it is the wrong page for them, and saying
       so once is more use than a read that fails every twelve seconds. */
    if (NOACCESS_RE.test((e && e.message) || "")) { showNoAccessGate(); busy = false; return; }
    setStatus("read failed", "err");
    toast("Could not read the workbook: " + e.message, true);
  }
  busy = false;
}

/* Graph answers 403 when the account may not open the file and 404 when it
   cannot even see the library it is in. Both mean the same thing here.
   Matched against the STATUS call() puts after its arrow, never against a bare
   number: row 403 of the Production sheet is in half the messages this app
   writes, and it is not a permission problem. */
const NOACCESS_RE = /->\s*40[34]\b|accessDenied|itemNotFound/;
let NOACCESS = false;
/** Put the sign-in overlay back with the one thing this person can act on: the
    station page. No retry loop - the poll and the reconcile both stand down. */
function showNoAccessGate() {
  NOACCESS = true;
  if (reconcileT) { clearTimeout(reconcileT); reconcileT = null; }
  const gate = $("#gate"); if (!gate) return;
  gate.hidden = false; gate.style.display = "flex";
  const top = $("#topbar"), main = $("#main");
  if (top) { top.hidden = true; top.style.display = "none"; }
  if (main) { main.hidden = true; main.style.display = "none"; }
  const box = $("#gateerr");
  if (box) {
    box.style.display = "block";
    box.innerHTML = "This account has no access to the production workbook. " +
      "Station accounts use the Glass station page.<br><br>" +
      '<a href="glass.html" style="color:#8ec5ff;font-weight:600">Open the Glass station page</a>';
  }
  const btn = $("#signinbtn");
  if (btn) btn.textContent = "Sign in with a different account";
}

/** Poll cheaply: only re-download when SharePoint says the file actually changed. */
async function poll() {
  if (busy || NOACCESS || !CW.account) return;
  try {
    const m = await CW.lastModified();
    if (m.at !== lastStamp) {
      toast("Sheet changed" + (m.by ? " — " + m.by : "") + ", reloading");
      await load("change detected…");
    }
  } catch (e) { /* transient; next tick will retry */ }
}

/* ---------- writes ----------
   GOLD_HEX / YELLOW_HEX / WHITE_HEX live in checkpoints.js, which loads first. */

/* Moving a job between the sheet's own sections moves its row in Excel. The
   list changes at once and holds the new place until the file catches up. */
let LASTWB = null, MOVING = {};
const sectionIdx = name => BLOCKNAMES.indexOf(name);
async function moveJobsInSheet(ids, idx) {
  const name = BLOCKNAMES[idx] || ("section " + idx);
  const before = {};
  ids.forEach(id => { const j = byId(id); if (j) before[id] = j.blk; });
  ids.forEach(id => pend(id, { blk: idx }));
  ALL = applyPending(ALL); state.picked = {}; renderAll();
  let ok = 0;
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    if (!(id in before)) continue;
    if (before[id] === idx) { ok++; continue; }
    MOVING[id] = 1; if (state.sel === id) renderDrawer();
    setStatus("moving " + id + " to " + name + " in Excel (" + (i + 1) + " of " + ids.length + ")…", "busy");
    try {
      const tmplFor = x => LASTWB ? templateForJob(LASTWB.getWorksheet("Production"), x) : null;
      const r = await CW.moveJobRow(id, idx, tmplFor, t => setStatus(id + ": " + t + "…", "busy"));
      if (r.moved) noteChange(id, "Section", r.from, r.to + " (row " + r.row + ")");
      ok++;
    } catch (e) {
      pend(id, { blk: before[id] });                 // put it back where it was shown before
      ALL = applyPending(ALL); renderAll();
      toast(id + ": " + friendly(e), true);
    }
    delete MOVING[id]; if (state.sel === id) renderDrawer();
  }
  setStatus("live");
  if (ok) toast(ok + " of " + ids.length + " moved to " + name + " in Excel");
  scheduleReconcile();
  return ok;
}

async function markReady(job, on) {
  const row = await CW.rowForJob("Production", job.id);   // re-found every time: rows move
  const addr = "A" + row + ":CL" + row;
  /* white, not "no fill": the sheet's cells carry an explicit white fill and
     clearing them leaves a hole that looks nothing like the rows around it */
  await CW.setFill("Production", addr, on ? GOLD_HEX : WHITE_HEX);
  return row;
}

/* ---------- checkpoints ----------
   Ticking work off. The screen changes at once and holds the new count; the
   write goes out after the taps stop, so a run of + is one write, not five. */
function cpRefresh(id) {
  ALL = applyPending(ALL);
  renderRows();
  if (state.sel === id) cpPatchSection(byId(id));   // patched in place: a tap must not rebuild the section under the finger
}

/** Set one item of one job to a count. Clamped, debounced, logged once. */
function setItemProgress(j, item, newDone) {
  if (!j || j.done) return;                        // a gold row is finished: no holes in it
  const s = itemState(j, item);
  if (!s) return;
  const want = cpClamp(newDone, s.total);
  if (want == null) return;                        // blank or not a number: nothing was asked for
  const cur = s.done == null ? 0 : s.done;
  if (want === cur && s.done != null) return;
  const col = cpColumn(item, PRODMAP);
  if (!col) { toast(cpLabel(item) + " is not a column on the Production sheet.", true); return; }
  /* is this the tick that leaves the job's glass reading nothing done? If it
     is, and the floor has really tapped this job, the office is asked before a
     single thing is written - and a "no" writes nothing at all, not even the
     workbook half. (An office clear reaches the floor, below.) */
  if (!confirmFloorClear(j, glassCountsAfter(j, item, want), j.id + "|" + item)) { cpRefresh(j.id); return; }
  pend(j.id, { cp: { [item]: want } });
  cpRefresh(j.id);
  /* from = the count before this burst of taps started, not before this tap.
     The column and total recorded here are only what the tap saw: the flush
     works both out again from the sheet it is about to write. */
  cpBurst(j.id + "|" + item,
    { job: j.id, item: item, col: col, who: whoAmI(),
      from: s.done == null ? null : cur, to: want, total: s.total },
    cpFlushItem);
}

/** One settled burst (or one replayed from a previous visit). */
async function cpFlushItem(b) {
  const j = byId(b.job);
  const total = j ? cpTotal(j, b.item) : b.total;
  /* worked out here, never at tap time: a column inserted in Excel in between
     would otherwise send this fill into somebody else's column */
  const col = cpColumn(b.item, PRODMAP);
  try {
    if (!col || !(total > 0)) {
      pend(b.job, { cp: { [b.item]: null } });          // nowhere to write it: stop showing it as held
      cpRefresh(b.job);
      toast(cpLabel(b.item) + " is not a column on the Production sheet - that tick was not saved.", true);
      return;
    }
    await cpWriteItem({ job: b.job, item: b.item, col: col, done: b.to, total: total,
                        from: b.from, who: b.who || whoAmI(), log: noteChange });
    /* the workbook half landed: if this was a confirmed clear of the job's
       glass, the floor's counters go back to nought too - and only THIS burst's
       own clear, never another burst's on the same job */
    floorClearAfterWrite(b.job, b.key);
    scheduleReconcile();                                  // reconcile once the file catches up
  } catch (e) {
    toast(friendly(e), true);
    /* put the screen back to what the sheet still says. When the count it
       replaced was unknown there is no number to go back to, so let the hold go
       and let Excel's colour speak for itself. */
    pend(b.job, { cp: { [b.item]: b.from == null ? null : b.from } });
    cpRefresh(b.job);
  } finally {
    cpSettled(b.key);
  }
}

const CPBUSY = {};        // job|group -> a group write is in the air

/** "All done" / "Clear" for a whole group: windows, doors, glass, or one product. */
async function setGroupDone(j, group, on) {
  if (!j || j.done) return;
  const items = cpItems(j).filter(x => x.group === group);
  if (!items.length) return;
  const missing = items.filter(x => !cpColumn(x.key, PRODMAP));
  if (missing.length) { toast(cpLabel(missing[0].key) + " is not a column on the Production sheet.", true); return; }
  /* "All glass done" toggled off leaves every glass item of the job reading
     nothing done. Ask before that destroys work the floor has recorded, and
     write nothing at all if the answer is no. */
  if (!on && group === "glass" &&
      !confirmFloorClear(j, glassCounts(j).map(c =>
        ({ type: c.type, total: c.total, status: "", done: 0 })), j.id + "|" + group)) return;
  const before = {}, held = {};
  items.forEach(x => {
    const s = itemState(j, x.key);
    before[x.key] = s && s.done != null ? s.done : null;  // null = it was "in progress, count unknown"
    held[x.key] = on ? x.total : 0;
    cpCancelBurst(j.id + "|" + x.key);                    // this write covers the item; drop its own
  });
  pend(j.id, { cp: held });
  CPBUSY[j.id + "|" + group] = 1;                         // no second tap while this one is in the air
  cpRefresh(j.id);
  const what = (items[0].groupLabel === "Glass" ? "Glass" : cap(items[0].groupLabel)) + (on ? ": all done" : ": cleared");
  const to = items.map(x => x.label.toUpperCase() + " " + (on ? x.total : 0)).join(", ");
  try {
    /* on the job's own chain, like the item writes: two taps on All done must
       land in the order they were made, or Excel and the counts disagree */
    await cpChain(j.id, () => {
      /* columns worked out here, not when the button was pressed: by now the
         sheet may have been reorganised under us */
      const cols = items.map(x => ({ item: x.key, col: cpColumn(x.key, PRODMAP), done: held[x.key], total: x.total }));
      const gone = cols.find(c => !c.col);
      if (gone) throw new Error(cpLabel(gone.item) + " is no longer a column on the Production sheet.");
      return cpWriteGroup({
        job: j.id, who: whoAmI(), what: what, to: to, log: noteChange, items: cols,
        before: items.map(x => ({ item: x.key, done: before[x.key], total: x.total }))
      });
    });
    /* the workbook half landed: a confirmed clear of the glass now reaches the
       floor's own counters as well, so both sides say the same thing */
    floorClearAfterWrite(j.id, j.id + "|" + group);
    toast(j.id + " · " + what);
    scheduleReconcile();
  } catch (e) {
    toast(friendly(e), true);
    const back = {};
    items.forEach(x => { back[x.key] = before[x.key]; });  // null drops the hold entirely
    pend(j.id, { cp: back });
  } finally {
    delete CPBUSY[j.id + "|" + group];
    cpRefresh(j.id);
  }
}

/* ---------- an office clear reaches the floor -------------------------------
   Shipped 2026-09-10 (docs/specs/2026-09-10-office-clears-the-floor.md), and
   it is the ONLY place in this file that writes one of the floor's own
   counters. Read the boundary before changing anything here.

   The bug it answers. Un-ticking a job's glass in the office cleared the
   workbook correctly every time - the cell really did go white - but nothing
   ever cleared the floor's counters, so the tablet stayed gold for ever and
   the only way back was somebody tapping minus forty-nine times. The office
   and the floor were left saying different things about the same job, which is
   the one thing this whole feature set exists to prevent.

   THE RULE, exactly (CLAUDE.md rule 3, second exception, owner 2026-09-10):
   when the office CLEARS a job's glass checkpoints it may write that job's
   `Glass station` row's Cut, Hotmelt, Glazed and Tuff to nought, plus
   DoneBy/DoneAt. That is all. The office still never writes `Station people`,
   never writes `Station log`, never writes a per-stage By/At, and never writes
   a floor column for any other reason.

   Four locks keep it that narrow, and they are independent of one another:

     1. Two call sites, both on a clear branch: setItemProgress and
        setGroupDone, and nowhere else in the file.
     2. FLOORCLEAR_OK - the office was ASKED and said yes, keyed to the WRITE
        that must land (job|item, or job|group) and not to the job, so no other
        burst on the same job can spend it. Set only where the clear is
        confirmed, consumed once, and nothing else writes to it. No token, no
        write, ever.
     3. clearFloorGlass re-derives the reason from the office's OWN record at
        write time (officeGlassEmpty, through ST.officeSeed - the very function
        that seeds a row) and from the floor's row (ST.floorWorkToClear). If
        the office's glass no longer reads "nothing done", nothing is written.
     4. The body can only be built by ST.officeClearFields - a name and a
        time in, no parameter that could carry a counter or a per-stage stamp. Every
        counter in it is a literal nought.

   A row the floor has NEVER tapped is deliberately left alone: its counters
   are the office's own seed echoed back and the feeder puts them right on its
   next run (sliceHash carries the seed, so the run is the next load, not ten
   minutes later). Writing DoneAt on such a row would mark it touched for ever
   and switch its seeding off - a worse bug than the one being fixed. */
const FLOORCLEAR_AGAIN_MS = 30000;
const FLOORCLEAR_TRIES = 3;
const FLOORCLEAR_OK = {};        // burst key -> the office was asked for this clear and said yes
/* What the Dashboard Log calls a clear of the floor's counters. Said once
   because glassLogStamps has to recognise it: it is an office action on the
   job and counts as the office's stamp on every one of its glass columns. */
const FLOORCLEAR_LOG = "Floor glass counters";
/* The DoneAt this dashboard last wrote on a job's row when clearing it. The
   clear stamps the floor's row so floorStamp stays honest (spec section 3),
   which means the office's own write would otherwise read as a FLOOR action
   newer than anything the office has - and the writer would paint the job's
   other glass columns out from the counters this very write zeroed. So it is
   kept, and glassOfficeJobStamp counts it as the office's. In memory only: the
   Dashboard Log line is what carries it across a reload and to a second
   dashboard. */
const OFFICE_FLOOR_AT = {};
const FLOORCLEAR_OWED = {};      // job -> how many times the write has been refused
let floorClearT = null;

/** Does the office's own record now say that NOTHING of this job's glass is
    done? Asked of DG and TG - the same glass the floor is given - through
    ST.officeSeed, which is the function that seeds a row FROM that record, so
    the two readings can never drift apart. */
function officeGlassEmpty(j, counts) {
  if (typeof ST === "undefined" || !j) return false;
  /* ST.TOTAL_TYPES is upper case and the job's own glass keys are lower, so the
     comparison goes through ST.jobKey - the same normaliser officeSeed itself
     filters with. Getting this wrong is silent: the filter simply matches
     nothing and the feature never fires. */
  const mine = (counts || glassCounts(j)).filter(c => ST.TOTAL_TYPES.indexOf(ST.jobKey(c.type)) >= 0);
  if (!mine.length) return false;                 // no DG and no TG: nothing the floor was ever given
  const seed = ST.officeSeed({ total: ST.glassTotal(j) }, mine);
  return !seed.cut && !seed.hotmelt && !seed.glazed;
}
/** This job's glass counts as they will read once one item is set to `want` -
    the office's record as it is ABOUT to be, so the clear can be recognised
    before anything is written. null when the item is not a glass item. */
function glassCountsAfter(j, item, want) {
  const p = String(item).split(":");
  if (p[0] !== "glass") return null;
  return glassCounts(j).map(c => String(c.type) !== p[1] ? c
    : { type: c.type, total: c.total, status: cpStatusFor(want, c.total), done: want });
}
/** Is this write the moment the office's glass record goes from saying
    something to saying nothing? That transition IS the clear. A tick made when
    the record already said nothing is not one, and must not reach the floor -
    otherwise clearing a hand-ticked ARCH on a job whose DG is blank would wipe
    the floor's counters, which nobody asked for. */
const officeClearsGlass = (j, after) => officeGlassEmpty(j, after) && !officeGlassEmpty(j);

/** Ask before a clear destroys work the floor has really recorded, and
    remember the answer. Returns false only when the office said no - and then
    NOTHING is written, not even the workbook half.

    Nothing is asked when the floor's row has no stamp on it: its counters are
    only the office's own seed echoed back, so there is nothing to lose. */
function confirmFloorClear(j, after, key) {
  if (typeof ST === "undefined" || !j) return true;
  if (!officeClearsGlass(j, after)) return true;              // not a clear of this job's glass
  const g = stationForJob(j.id);
  const say = g ? ST.clearWarning(g) : "";
  if (say && typeof confirm === "function" && !confirm(say)) return false;
  /* Said yes, or there was nothing to ask about: this job's floor counters may
     be put back to nought once THIS write has landed. The key is the burst key
     - job|item, or job|group - and not the job, because a job has several
     bursts at once: un-tick the glass and tick Windows on the same job inside
     the 800 ms debounce and the Windows write lands first. Keyed on the job,
     that write would have spent the answer and cleared the floor BEFORE the
     glass write landed - and if the glass write then failed, the floor would be
     at nought with the sheet still gold, the mirror image of the bug this
     feature exists to remove. */
  FLOORCLEAR_OK[key] = 1;
  return true;
}

/** Come back to a refused clear. ONE timer for all of them, the way the
    feeder's and the colour writer's follow-ups are. */
function floorClearAgain() {
  if (floorClearT) return;
  floorClearT = setTimeout(() => {
    floorClearT = null;
    Object.keys(FLOORCLEAR_OWED).forEach(id => { clearFloorGlass(id).catch(() => {}); });
  }, FLOORCLEAR_AGAIN_MS);
}

/** Put one job's floor counters back to nought. The one write. */
async function clearFloorGlass(job) {
  const id = String(job);
  if (typeof ST === "undefined" || typeof CW === "undefined" || !CW || !CW.listPatch) return false;
  const j = byId(id);
  /* re-derived here and not taken on trust from the caller: if the office's
     own record no longer says "nothing done" - somebody re-ticked it while the
     workbook write was in the air - then this is not a clear any more and
     nothing at all is written */
  if (!j || !officeGlassEmpty(j)) { delete FLOORCLEAR_OWED[id]; return false; }
  const g = stationForJob(id);
  if (!g || !ST.floorWorkToClear(g)) { delete FLOORCLEAR_OWED[id]; return false; }
  /* the same words the question used (ST.clearWords), so the apology for a
     write that could not be made and the question that authorised it cannot
     describe the same row differently */
  const was = ST.clearWords(g);
  const body = ST.officeClearFields(feedWho(), new Date().toISOString());
  try {
    if (STATION_OK !== true) throw new Error(STATION_WHY || "the floor's list is not readable");
    const siteId = await CW.stationSite();
    if (!siteId) throw new Error(STATION_SITE_MISSING);
    await CW.listPatch(ST.STATION_LIST, g.id, body, { siteId: siteId, fields: ST.STATION_FIELDS });
  } catch (e) {
    const n = (FLOORCLEAR_OWED[id] || 0) + 1;
    const why = (e && e.message) || String(e);
    console.warn("[station] could not clear " + id + "'s floor counters (" + n + "): " + why);
    if (n >= FLOORCLEAR_TRIES) {
      delete FLOORCLEAR_OWED[id];
      toast(id + ": the glass was cleared in the sheet, but the floor's counters could not be " +
            "reset - the tablet still shows " + was + ". " + friendly(e), true);
    } else { FLOORCLEAR_OWED[id] = n; floorClearAgain(); }
    return false;
  }
  delete FLOORCLEAR_OWED[id];
  /* the office has just moved this row, so remember WHEN before anything reads
     it back: DoneAt is the office's own stamp here, not the floor's, and the
     colour writer has to know that or it will paint this job's other glass
     columns out from the counters this very write zeroed */
  OFFICE_FLOOR_AT[id] = body.DoneAt;
  /* keep this dashboard's copy of the list in step at once, so the board, the
     rows and the drawer read nought without waiting for the ten-second poll -
     and so the colour writer plans from what the list now says. A NEW array,
     because stationRecords() caches on its identity. */
  STATION_ITEMS = (STATION_ITEMS || []).map(it => String(it.id) === String(g.id)
    ? { id: it.id, fields: Object.assign({}, it.fields || {}, body) } : it);
  redrawStation();
  /* an office action, in the office's own log, exactly as spec §3 says - and
     NOT in the Station log, which stays the floor's alone. glassLogStamps
     reads this line as the office's stamp on this job's glass, and that is
     what carries the clear across a reload and to a second dashboard: without
     it the DoneAt written above reads as a floor action and the office's own
     clear out-ranks the office. */
  noteChange(id, FLOORCLEAR_LOG, was, "nothing");
  return true;
}

/** The workbook half of a clear has landed. If the office asked for this job's
    glass to be cleared, and meant it, the floor's counters follow. */
function floorClearAfterWrite(job, key) {
  if (!FLOORCLEAR_OK[key]) return;
  delete FLOORCLEAR_OK[key];
  clearFloorGlass(job).catch(e => console.warn("[station] " + ((e && e.message) || e)));
}

/* Unsent taps must survive the tab closing, and be sent when it comes back. */
function cpWatchExit() {
  const go = () => cpFireAll();
  window.addEventListener("pagehide", go);
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") go(); });
}
function cpReplayQueue() {
  const n = cpReplay(cpFlushItem, whoAmI());
  if (n) toast(n === 1 ? "Sending a checkpoint from earlier" : "Sending " + n + " checkpoints from earlier");
  return n;
}

/* ---------- glass colours: the floor's work reaches the Production sheet ----
   Shipped 2026-09-10 (docs/specs/2026-09-10-glass-colours-two-way.md). This is
   the first feature in which something done on the floor changes the master
   sheet, so read the boundaries before changing anything here.

   The tablet still never touches the workbook - it cannot, and that is the
   real security boundary. THIS dashboard does every write: it watches the
   floor's list, works out what the four glass columns should be showing, and
   paints those cells. The feeder in reverse.

   Only a FILL is ever written, and only into DG, TG, TUFF and NOT TUFF of the
   job's own row on Production. Never a value, never a row, never a formula,
   never another sheet, and never ARCH, ASTRAGAL, FANCY or EXTRA, which the
   office ticks by hand (owner, 2026-09-10: "the production sheet is main doc,
   don't edit that as in text; by colour is ok").

   Four things keep it from being a nuisance:

     · A cell that is already the right colour is not written. The office polls
       the floor's list every ten seconds; without that test this would be a
       write storm against the workbook rather than a feature.
     · A job the floor has never tapped is never written at all. Its counters
       are the office's own seed echoed back (ST.officeSeed), and painting the
       sheet from them would be this dashboard arguing with itself.
     · Whoever acted last wins - the floor's stamp on the list row against the
       office's own stamp in Dashboard Progress and Dashboard Log. Neither side
       is overwritten by an older action, which is both halves of what the
       owner asked for.
     · Every write is held in PENDING until the download agrees, exactly as a
       checkpoint tick is, because the downloaded file lags the API by ~36 s.

   White, not "no fill", for a cell that should be showing nothing: the
   Production cells carry an explicit white fill, and clearing one leaves a
   hole that looks nothing like the rows around it. The parser reads white and
   no-fill as the same nothing, so a cell that already has neither colour is
   left alone rather than painted white for the sake of it.                  */
const GLASS_PROD_SHEET = "Production";
/* What the Dashboard Log calls a paint made from the floor's counters. Said
   once, and deliberately not beginning "Glass": glassLogStamps would read that
   as the office having spoken about the job's glass, and this is the floor's
   work, not the office's. */
const GLASSPAINT_LOG = "Floor glass colours";
const GLASS_HEX = { gold: GOLD_HEX, yellow: YELLOW_HEX, "": WHITE_HEX };
const GLASSC_BUSY = {};        // job -> a colour write for it is in the air
/* Cells per run, and three at a time - the same numbers and the same shape as
   the feeder's STATION_FEED_MAX / STATION_FEED_LANES, deliberately, because it
   is the same problem: switching this feature on after the floor has worked for
   a week with no office dashboard open plans a colour for every job at once.
   Rehearsed against the owner's real file: 141 fed rows, up to 362 cell fills
   in one burst if nothing caps it. A run cut short by the cap comes back in
   half a minute, like the feeder's own follow-up, and the rest goes then. */
const GLASS_MAX = 60;
const GLASS_AGAIN_MS = 30000;
let glassAgainT = null;
let glassRunning = false;
/** Come back for the rest. ONE timer, never one per job - the feeder had
    exactly that bug and fixed it the same way. */
function glassColourAgain() {
  if (glassAgainT) return;
  glassAgainT = setTimeout(() => {
    glassAgainT = null;
    glassColourRun().catch(e => console.warn("[glass] " + ((e && e.message) || e)));
  }, GLASS_AGAIN_MS);
}

/* A write the workbook keeps refusing must not be sent again every ten seconds.
   Everything comparable in this file already has a brake - stationDelta marks a
   refusing list off for five minutes, the feeder has stationFeedAgain, the
   checkpoint queue bounds its own replays - and this writes to the WORKBOOK, so
   it is the one that can least afford not to. The trigger is real and
   undramatic: somebody opens the file exclusively in desktop Excel, or a token
   loses its write scope, in the middle of a shift while the floor keeps tapping.

   So each job counts its own failures and waits longer each time, and after
   GLASS_FAIL_MAX it is given up on until the page is reloaded. A success clears
   the record. Either way the footer says so, because a workbook refusing a
   write is not something to leave in a console nobody has open. */
const GLASS_BACKOFF_MS = [60000, 300000, 900000];        // 1 min, then 5, then 15
const GLASS_FAIL_MAX = 5;
const GLASSC_FAIL = {};        // job -> { n, at, why }
let GLASS_COLOUR_ERR = "";
/** Is this job serving a backoff, or given up on? */
function glassWaiting(id) {
  const f = GLASSC_FAIL[id];
  if (!f) return false;
  if (f.n >= GLASS_FAIL_MAX) return true;
  return Date.now() - f.at < GLASS_BACKOFF_MS[Math.min(f.n - 1, GLASS_BACKOFF_MS.length - 1)];
}
/** What the footer says about it. One line, with the detail in the tooltip -
    the same arrangement the station feed's own word already uses. */
function setGlassFoot() {
  const bad = Object.keys(GLASSC_FAIL);
  if (!bad.length) { GLASS_COLOUR_ERR = ""; setStationFoot(); return; }
  const stopped = bad.filter(id => GLASSC_FAIL[id].n >= GLASS_FAIL_MAX).length;
  const why = GLASSC_FAIL[bad[0]].why || "";
  GLASS_COLOUR_ERR =
    bad.length + " job" + (bad.length > 1 ? "s" : "") + " could not have " +
    (bad.length > 1 ? "their" : "its") + " glass colours written to the Production sheet" +
    (stopped ? " — " + stopped + " of them given up on until this page is reloaded"
             : " — trying again shortly") +
    (why ? ". Last reason: " + why : "") +
    " Nothing else about the sheet is affected.";
  setStationFoot();
}

/** A stamp from any of the three places this app keeps one, as **the latest
    instant it can be describing**, in milliseconds - or 0 when there is
    nothing readable there.

    Three formats meet here and getting this wrong would silently decide every
    contest the wrong way round. The floor writes a full ISO stamp in UTC
    (`new Date().toISOString()`). Dashboard Progress and Dashboard Log write
    `nowStamp()` - "2026-09-10 14:03", the office's own local time with no zone
    on it - and the Log is read back out of the workbook as "10/09/2026 14:03".
    The last two are LOCAL times and are built as local dates here; Date.parse
    is left to handle the ISO one, where the Z means what it says. Doing this
    by Date.parse alone would work in one browser and be an hour out in
    another, all summer.

    THE SECONDS, which is the whole reason this returns "the latest instant"
    rather than "the instant". `nowStamp()` has no seconds in it, so an office
    tick made at 15:00:40 is written down as 15:00 - up to 59 seconds EARLY. A
    floor tap at 15:00:20 would then read as the later action and paint the
    office's own tick back out, which is the one thing the owner forbade. So a
    stamp that only names a minute is taken to cover that whole minute and
    anything inside it loses to it. It is the same reading as ties going to the
    office (spec Amendment A3), applied to the precision the sheets actually
    have.

    It also makes the answer STABLE across a round trip, which the minute-early
    reading did not. `noteChange` puts a full-second entry in CHANGES, so the
    tab that made a tick used to win - until `load()` dropped that local entry
    in favour of the Log sheet's minute-precision copy of the same line, and
    the dashboard quietly changed its mind and repainted. Now both readings of
    the same tick answer the same side. */
const STAMP_MINUTE_MS = 59999;         // ... so 15:00 means "any time before 15:01"
function stampMs(v) {
  const s = String(v == null ? "" : v).trim();
  if (!s) return 0;
  let m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)).getTime() +
                (m[6] === undefined ? STAMP_MINUTE_MS : 0);
  m = /^(\d{2})\/(\d{2})\/(\d{4})[ T](\d{2}):(\d{2})$/.exec(s);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]).getTime() + STAMP_MINUTE_MS;
  const t = Date.parse(s);
  return isFinite(t) ? t : 0;
}

/** When the office last said something about one glass type of one job, in
    milliseconds - 0 when it never has.

    Two records, both already in memory after load(), so this costs no request:
    the Dashboard Progress row the drawer writes for every tick (`Who`/`When`,
    read back through cpStored), and the Dashboard Log lines, which arrive as
    CHANGES. A whole-group write logs "Glass: all done" or "Glass: cleared" and
    counts for every type; a single tick logs the item's own name.

    0 is the honest answer for an office tick made before this feature existed,
    and for a cell somebody painted by hand in Excel: neither leaves a dated
    record anywhere, so neither can be shown to be the later action. See the
    honest limits in the spec. */
function glassOfficeStamp(job, type, log, jobAt) {
  const id = String(job).toUpperCase();
  const row = typeof cpStored === "function" ? cpStored(id, "glass:" + type) : null;
  let best = row ? stampMs(row.when) : 0;
  const mine = (log || glassLogStamps())[id];
  if (mine) {
    /* the item's own line, and any whole-group line, which is about all four */
    const t = Math.max(mine[type] || 0, mine["*"] || 0);
    if (t > best) best = t;
  }
  /* AND the office's stamp on the JOB, whatever column it was made on. This is
     the half that was missing, and it is what the owner saw on 2026-09-10.

     Asked per column, this function answered a literal 0 for every column the
     office had not named IN THAT ACTION - no hold, no Progress row, no Log
     line - so any floor stamp at all beat it. The office pressed Clear on TG
     alone and the writer promptly painted out the TUFF and NOT TUFF it had
     ticked by hand, because for those two columns the office had never
     "spoken". Worse, the floor stamp it lost to was one the office had itself
     just written: clearFloorGlass's own DoneAt, a fraction of a second after
     the click.

     The office's control is per JOB. A clear is a job-level act - it is what
     the drawer's Clear does - and the floor's row is one combined number with
     no per-type split in it at all (the colours spec, section 8). So an office
     action at time T on a job must not lose to a floor stamp on a column that
     action happened not to mention. That is all this adds, and it changes
     nothing about the contest itself: a floor tap made after T still carries
     the later stamp and still wins. */
  const at = jobAt == null ? glassOfficeJobStamp(byId(id) || byId(job), log) : jobAt;
  if (at > best) best = at;
  /* AND the tick this office has made but not landed yet, which is the third
     record and the one this used to miss. Both of the records above are
     written by the WRITE - Dashboard Progress at the start of it, Dashboard
     Log at the end - so between the click and the end of that write the office
     HAS acted and nothing here said so. The floor then won by default, its
     colour landed after the office's own (both go on the job's cpChain), and
     because the cell was then the colour the floor wanted it was never written
     again: an un-tick undone silently and for good. The owner's bug of
     2026-09-10, and it only showed on un-ticking because marking done moves
     the cell TOWARDS what the floor says, where there is nothing to plan.

     The hold IS the office's action, stamped at the click by pend(), so this
     does not change last-writer-wins: a floor tap made after the click still
     carries the later stamp and still wins. Only the glass items' own holds
     count - never a `gc` hold, which is this feature's own write. */
  const held = PENDING[id] || PENDING[job];
  if (held && held.cp && Object.prototype.hasOwnProperty.call(held.cp, "glass:" + type)) {
    const t = ((held.t || {})["cp:glass:" + type]) || held.at || 0;
    if (t > best) best = t;
  }
  return best;
}

/** The newest thing the OFFICE has said about this job's glass, on any column,
    in milliseconds - 0 when it never has. Four records, all of them already in
    memory, so this costs no request:

      - every glass item's hold, stamped at the click by pend(): the office has
        acted the moment the button is pressed, long before either sheet knows
        about it;
      - every "Glass ..." and "Glass: ..." line of the Dashboard Log for this
        job, whichever column it named;
      - the "Floor glass counters" line, which records the office clearing the
        floor's counters. It is an OFFICE action on this job and is counted as
        one. The earlier build deliberately excluded it, which was exactly
        backwards: the clear's other half - the DoneAt it writes on the floor's
        row - was being counted FOR THE FLOOR, so the office's own act
        out-ranked the office;
      - every glass item's Dashboard Progress row.

    Takes the job object rather than its id: byId is a linear scan of every job
    on the sheet, and this is asked once per job per pass. */
function glassOfficeJobStamp(j, log) {
  if (!j || !j.id) return 0;
  const id = String(j.id).toUpperCase();
  let best = 0;
  const held = PENDING[id] || PENDING[j.id];
  if (held && held.cp) {
    Object.keys(held.cp).forEach(k => {
      if (k.indexOf("glass:") !== 0) return;          // never a `gc` hold: that is our own write
      const t = ((held.t || {})["cp:" + k]) || held.at || 0;
      if (t > best) best = t;
    });
  }
  const mine = (log || glassLogStamps())[id];
  if (mine && mine.job > best) best = mine.job;
  if (typeof cpStored === "function") {
    Object.keys(j.glass || {}).forEach(k => {
      const row = cpStored(id, "glass:" + k);
      const t = row ? stampMs(row.when) : 0;
      if (t > best) best = t;
    });
  }
  /* and the clear this dashboard made itself, which is the only record of it
     until the Log line has been written and read back */
  const own = stampMs(OFFICE_FLOOR_AT[id] || OFFICE_FLOOR_AT[j.id] || "");
  if (own > best) best = own;
  return best;
}

/** Is an office change to THIS job's glass still settling? While it is, the
    colour writer stands down for this job (see glassColourPlan).

    The window is PENDING_MS - the same three minutes every other optimistic
    hold in this app uses, and for the same reason: it is how long the
    downloaded copy, the dashboard's own sheets and the floor's list can take
    to agree on something this dashboard has already done.

    It is one job's window, not the sheet's. An un-tick on one job says nothing
    about another and must never hold the floor's work off the whole sheet. */
/* How many jobs the guard stood down for in this pass. A deferred job is not a
   dropped one: the writer re-plans from whatever the list says every time it
   runs, so the colour lands on a later pass. But it only runs when something
   moves - a floor tap, or a load - and on a quiet evening nothing may move
   again for hours, which would leave the sheet stale for no good reason. So a
   pass that deferred anything comes back, through the follow-up the cap
   already uses: one timer, no requests, and it stops as soon as the windows
   have closed and there is nothing left to do. */
let GLASS_SETTLING = 0;
function officeSettling(j, log) {
  if (!j || !j.id) return false;
  const held = PENDING[j.id] || PENDING[String(j.id).toUpperCase()];
  /* a hold on any glass item of this job: the office has clicked and the file
     has not caught up. Never a `gc` hold, which is this feature's own write. */
  if (held && held.cp && Object.keys(held.cp).some(k => k.indexOf("glass:") === 0)) return true;
  const at = glassOfficeJobStamp(j, log);
  return at > 0 && (Date.now() - at) < PENDING_MS;
}

/** The Dashboard Log, indexed by job and glass type, in one pass.
    CHANGES can be four hundred lines and the writer asks about four columns of
    every job on the sheet, so asking it line by line would be the log walked
    a couple of thousand times per run. Built once per run and handed down;
    a caller with nothing to hand gets a fresh one, so the function above is
    still usable on its own. */
function glassLogStamps() {
  const out = {};
  for (let i = 0; i < CHANGES.length; i++) {
    const c = CHANGES[i];
    if (!c) continue;
    const what = String(c.what || "").toUpperCase();
    /* "GLASS DG" is one column; "GLASS: ALL DONE" and "GLASS: CLEARED" are the
       whole group and count for every one of them; and "FLOOR GLASS COUNTERS"
       is the office clearing the floor's own counters, which is an office
       action on this job's glass like any other.

       That last one used to be excluded here, deliberately and wrongly. The
       reasoning was that the feature must not feed the contest with its own
       writes - but the clear's OTHER half, the DoneAt it stamps on the floor's
       row, was being read as a FLOOR action by floorStamp. Counting one side
       and not the other is what let the office's own clear out-rank the
       office. Both halves are the office's, and both are counted now. */
    const group = what.indexOf("GLASS:") === 0;
    const clear = what === FLOORCLEAR_LOG.toUpperCase();
    if (!group && !clear && what.indexOf("GLASS ") !== 0) continue;
    const id = String(c.job || "").toUpperCase();
    if (!id) continue;
    const t = stampMs(c.at);
    if (!t) continue;
    const m = out[id] || (out[id] = {});
    /* every one of them counts for the JOB; only a column's own line counts
       for that column */
    if (t > (m.job || 0)) m.job = t;
    if (clear) continue;
    const key = group ? "*" : what.slice(6).toLowerCase();
    if (t > (m[key] || 0)) m[key] = t;
  }
  return out;
}

/** What this job's four glass cells should be repainted to, or null when there
    is nothing to do - which is the answer nearly every time this is asked.

    Read it as a series of reasons to write nothing, because that is what it
    mostly is: the job is finished, the floor has never touched it, the column
    is not on this sheet, the job has none of that glass, the cell already
    says it, the cell is a colour this feature does not own, or the office
    spoke more recently than the floor did. */
function glassColourPlan(j, log) {
  if (!j || j.done) return null;             // a gold row is finished work: leave it whole
  if (!PRODMAP || !PRODMAP.glass || typeof ST === "undefined") return null;
  /* THE OFFICE IS ABSOLUTE, AND THIS IS WHERE THAT IS ENFORCED.

     The owner's rule, 2026-09-10: "the hierarchy is Excel, then master
     dashboard, then glass. Any change from the dashboard is absolute. If the
     glass updates the ticks it comes golden instantly, correct, and should not
     change. An un-tick from the dashboard is absolute - no thinking, no
     arguing."

     So while an office change on this job's glass is still settling, this
     function makes NO decision about the job. It does not compare stamps, it
     does not plan, it does not write. Not because the comparison is wrong -
     it was corrected twice - but because inside that window it is being fed
     copies that have not caught up, and it cannot tell them apart from the
     real thing: the downloaded workbook is about 36 seconds behind, the
     floor's list is a poll behind, and the clear's own DoneAt looks exactly
     like a tap. The owner watched the gold come back after both stamp fixes.
     This one stops arguing instead of trying to win the argument.

     Two readings of "still settling", and either is enough, because the office
     must be safe in every one of the places its action is read from:

       - a hold on any of this job's glass items. pend() stamps it at the
         CLICK and applyPending lets it go only when the downloaded file agrees
         - which is precisely the window this needs;
       - the office's own stamp on the job (glassOfficeJobStamp: the holds, the
         Dashboard Progress rows, the Dashboard Log lines including the clear's
         own, and OFFICE_FLOOR_AT) being younger than PENDING_MS. That covers
         the moment after a hold is released, and a second dashboard or a
         reload, where the hold never existed.

     What this feature is FOR is carrying the FLOOR's work up to the sheet. It
     is not for second-guessing the office, and outside this window it does not
     change at all: last-writer-wins still decides, and a genuine floor tap
     still paints. */
  if (officeSettling(j, log)) { GLASS_SETTLING++; return null; }
  const g = stationForJob(j.id);
  if (!g) return null;                       // never fed to the floor
  const floorAt = stampMs(ST.floorStamp(g));
  if (!floorAt) return null;                 // the floor has never tapped it: nothing of theirs to show
  const want = ST.glassColours(g);
  /* worked out once for the job, not once per column: byId is a linear scan
     and this is asked of every job on the sheet, six times a minute */
  const jobAt = glassOfficeJobStamp(j, log);
  /* and the memory of our own clear is let go once the floor has moved that
     row since - it can decide nothing after that, and the map would otherwise
     grow for as long as the tab is open */
  const own = OFFICE_FLOOR_AT[j.id];
  if (own && stampMs(own) < floorAt) delete OFFICE_FLOOR_AT[j.id];
  const out = [];
  ST.COLOUR_TYPES.forEach(type => {
    const col = PRODMAP.glass[type];
    if (!col) return;                        // not a column on this sheet at all
    if (!((j.glass || {})[type] > 0)) return;   // the job has none of this glass
    const have = glassCellNow(j, type);
    /* undefined means the cell is carrying something else - the sheet's own Cut
       green, or a colour somebody painted for a reason of their own. This
       feature owns gold, yellow and nothing; it does not paint over anything
       it cannot recognise, in either direction. */
    if (have === undefined) return;
    if (have === want[type]) return;         // already right: never rewrite a cell
    /* last writer wins, and a tie goes to the office. Two reasons: the owner's
       "the master dashboard should not be overridden", and the office's own
       stamp is written to the minute, so it is already biased early - giving it
       the ties it would otherwise lose is the reading that is wrong less often.
       It is also the quiet answer: a tie resolves to no write at all. */
    if (glassOfficeStamp(j.id, type, log, jobAt) >= floorAt) return;
    out.push({ type: type, col: col, from: have, to: want[type] });
  });
  return out.length ? out : null;
}

/** Paint one job's cells. On the job's own checkpoint chain, so an office tick
    and a floor colour for the same job can never interleave, and inside
    serialised() for the sheet, so two of these cannot either. */
async function glassColourWrite(job, plan) {
  try {
    await CW.serialised(GLASS_PROD_SHEET, async () => {
      /* re-found immediately before writing, like every other fill in this app:
         rows move, and a remembered row number eventually paints somebody
         else's job */
      const row = await CW.rowForJob(GLASS_PROD_SHEET, job);
      const f = await CW.findFile();
      const S = f.base + "/worksheets('" + GLASS_PROD_SHEET + "')";
      await CW.batchWrite(plan.map(p => ({
        method: "PATCH",
        url: S + "/range(address='" + CW.A1(p.col) + row + "')/format/fill",
        body: { color: GLASS_HEX[p.to] }
      })));
    });
    delete GLASSC_FAIL[job];                 // it works again: forget the backoff
    setGlassFoot();
    /* one line per paint, in the office's own history. Until 2026-09-11 this
       feature wrote nothing anywhere (the colours spec's Amendment A12), so a
       job going gold by itself left no trace at all - which is exactly why the
       morning of 2026-09-11 was inexplicable until the whole thing was
       reproduced in a browser. A12's other reason stands and is not touched:
       no Dashboard Progress row, because the floor counts one combined DG + TG
       number and a per-type count would be invented.

       The WORDING is load-bearing. glassLogStamps reads any entry beginning
       "Glass " or "Glass:" as the OFFICE having spoken about that job's glass;
       this line is the FLOOR's work and must never be read as one, or the
       writer's own paint would out-rank the floor it came from. "Floor glass
       colours" begins with neither, exactly as "Floor glass counters" does -
       do not rename it to start with Glass. */
    const said = k => plan.map(p => p.type.toUpperCase() + " " + (p[k] || "blank")).join(", ");
    noteChange(job, GLASSPAINT_LOG, said("from"), said("to"));
    scheduleReconcile();                     // read the file back once it has caught up
  } catch (e) {
    /* let the holds go rather than putting an old colour back: what the sheet
       still says IS the old colour, and holding a guess over it for three
       minutes would be this dashboard telling the office something untrue */
    const back = {};
    plan.forEach(p => { back[p.type] = null; });
    pend(job, { gc: back });
    ALL = applyPending(ALL);
    if (!rowsInUse()) quietRows();
    if (state.sel === job && $("#dhost")) renderDrawer();
    const why = (e && e.message) || String(e);
    const f = GLASSC_FAIL[job] || { n: 0 };
    f.n++; f.at = Date.now(); f.why = why;
    GLASSC_FAIL[job] = f;
    setGlassFoot();
    console.warn("[glass] could not colour " + job + " (" + f.n + " time" + (f.n > 1 ? "s" : "") +
                 "): " + why);
    /* thrown on so the run that dispatched this can count it: the footer needs
       to know how the whole pass went, not only that one job is unhappy */
    throw e;
  } finally {
    delete GLASSC_BUSY[job];
  }
}

/** Look at every job once and paint whatever the floor has moved on.

    Run after every load and after every station poll that actually moved
    something. Cheap when there is nothing to do: it is a walk of the jobs in
    memory and not one request, which is what lets it be called six times a
    minute.

    Bounded when there is a great deal to do: at most GLASS_MAX cells go out in
    one pass, three writes at a time, and a pass the cap cut short arms one
    follow-up for the rest. A job is never split across two passes - half a
    job's colours would be a lie on screen for thirty seconds and the other
    half would only be re-planned anyway.

    Returns how many JOBS were sent, so a caller (and a test) can see the cap
    working. */
async function glassColourRun() {
  if (typeof ST === "undefined" || typeof CW === "undefined" || !CW || !CW.serialised) return 0;
  if (STATION_OK !== true || !PRODMAP || !ALL.length) return 0;
  /* a pass is already in the air. Arming the follow-up rather than simply
     returning is what guarantees the rest of a big catch-up still goes out. */
  if (glassRunning) { glassColourAgain(); return 0; }
  const all = [];
  GLASS_SETTLING = 0;
  const log = glassLogStamps();              // one pass over the log, not one per column
  for (let i = 0; i < ALL.length; i++) {
    const j = ALL[i];
    if (GLASSC_BUSY[j.id]) continue;         // one write per job in the air at a time
    if (glassWaiting(j.id)) continue;        // this one is serving a backoff, or given up on
    const plan = glassColourPlan(j, log);
    if (plan) all.push({ id: j.id, plan: plan });
  }
  if (!all.length) {
    /* nothing to paint now - but if that is because the office is still
       settling on some job, come back for it rather than waiting for the next
       thing to happen to move */
    if (GLASS_SETTLING) glassColourAgain();
    return 0;
  }
  /* the cap counts CELLS, because a cell fill is what reaches the workbook -
     one job is one batched request carrying up to four of them */
  const todo = [];
  let cells = 0;
  for (let i = 0; i < all.length; i++) {
    if (todo.length && cells + all[i].plan.length > GLASS_MAX) break;
    cells += all[i].plan.length;
    todo.push(all[i]);
  }
  const whole = todo.length === all.length;
  /* held before a single request leaves, and all of them before anything is
     redrawn: the screen must show what it is about to write, not flicker
     through the states it is writing */
  todo.forEach(x => {
    GLASSC_BUSY[x.id] = 1;
    const held = {};
    x.plan.forEach(p => { held[p.type] = p.to; });
    pend(x.id, { gc: held });
  });
  ALL = applyPending(ALL);
  if (!rowsInUse()) quietRows();
  if (state.sel && $("#dhost")) renderDrawer();
  console.log("[glass] painting " + todo.length + " job" + (todo.length > 1 ? "s" : "") +
              " (" + cells + " cell" + (cells > 1 ? "s" : "") + ") of " + all.length +
              " from the floor's counters");
  glassRunning = true;
  try {
    /* stationSend is the feeder's own three-lane runner: it finishes what it
       can, counts what it lost and keeps the first message. Each job still goes
       through its own checkpoint chain inside that, so an office tick and a
       floor colour for one job can never interleave. */
    const r = await stationSend(todo.map(x => () => cpChain(x.id, () => glassColourWrite(x.id, x.plan))));
    if (!whole || r.failed || GLASS_SETTLING) glassColourAgain();
  } finally {
    glassRunning = false;
  }
  return todo.length;
}

/* ---------- job alerts ----------------------------------------------------
   A person subscribes an email address to a job; a flow inside the tenant
   emails each address its jobs and their comments. The dashboard only manages
   the subscriptions, and they live in the Dashboard Alerts sheet - never here.
   HARD RULE: no address and no domain is written down in this file. The
   administrator's address comes from the Dashboard Config sheet at run time,
   and the only domain an address may belong to is that address's own.      */
const CONFIG_SHEET = "Dashboard Config";
/* said in one place, used in both: the drawer and the Alerts window */
const ALERT_NOTE = "Emails go out every 3rd day at 08:00 to each address listed here, one email per address.";
const ALERT_ADMIN_ONLY = "Only the administrator can change alerts.";
let CONFIG = {};                 // Dashboard Config, read-only for the dashboard
let ALERTS = {};                 // job -> [{email, who, when}], from Dashboard Alerts
let ALCONF = {};                 // job|email -> the inline "Remove?" is showing

/* The optimistic hold, the same idea as PENDV: SharePoint needs ~36 s to put
   our change into the downloadable file, so a chip appears (or goes) at once
   and is held until the file agrees, or for 180 s, whichever comes first. */
let PENDA = {};
try { PENDA = JSON.parse(localStorage.getItem("cw_penda") || "{}"); } catch (e) { PENDA = {}; }
/* An address may sit in this browser for the length of the hold and not a
   moment longer. Expiring it only during a successful load would leave one
   here for as long as a tab sat open on a workbook that never reloaded, or
   was simply closed and re-opened, so the rule is applied wherever the store
   is touched: on the way in, and before anything is written back out. */
function purgePendA() {
  const now = Date.now(); let changed = false;
  Object.keys(PENDA).forEach(k => {
    if (now - ((PENDA[k] || {}).at || 0) > PENDING_MS) { delete PENDA[k]; changed = true; }
  });
  return changed;
}
const savePendA = () => { purgePendA(); try { localStorage.setItem("cw_penda", JSON.stringify(PENDA)); } catch (e) {} };
if (purgePendA()) savePendA();       // holds left behind by an earlier visit, before anyone can see them
const alKey = (job, email) => job + "|" + email;
function pendAlert(job, email, act, who, when) {
  PENDA[alKey(job, email)] = { act: act, who: who || "", when: when || nowStamp(), at: Date.now() };
  savePendA();
}
function dropPendAlert(job, email) { delete PENDA[alKey(job, email)]; savePendA(); }
function alPut(job, email, who, when) {
  const list = ALERTS[job] = ALERTS[job] || [];
  if (!list.some(x => x.email === email)) list.push({ email: email, who: who || "", when: when || nowStamp() });
}
function alDrop(job, email) { if (ALERTS[job]) ALERTS[job] = ALERTS[job].filter(x => x.email !== email); }
/** Re-apply the held adds and removes over what a freshly read sheet says, and
    let a hold go the moment the file agrees with it. Called from load(). */
function applyPendA() {
  let changed = purgePendA();                     // the age rule lives in one place
  Object.keys(PENDA).forEach(k => {
    const p = PENDA[k], i = k.indexOf("|"), job = k.slice(0, i), email = k.slice(i + 1);
    const has = (ALERTS[job] || []).some(x => x.email === email);
    if (p.act === "add" ? has : !has) { delete PENDA[k]; changed = true; return; }
    if (p.act === "add") alPut(job, email, p.who, p.when); else alDrop(job, email);
  });
  if (changed) savePendA();
}

/** One cell of a dashboard sheet in the downloaded workbook, as text. Excel
    sometimes turns our text stamp back into a real date; write it out again
    in the same shape rather than letting a locale guess at it. */
function sheetText(row, i) {
  const v = row.getCell(i).value;
  if (v == null) return "";
  if (v instanceof Date) {
    const p = n => (n < 10 ? "0" : "") + n;
    return v.getFullYear() + "-" + p(v.getMonth() + 1) + "-" + p(v.getDate()) +
           " " + p(v.getHours()) + ":" + p(v.getMinutes());
  }
  return v.text != null ? v.text : String(v);
}

/** Read both alert sheets out of a freshly downloaded workbook. Neither is
    written here: Config is read-only for the dashboard, and a missing sheet or
    a missing key simply means nobody is the administrator. */
function readAlertSheets(wb) {
  const cfg = {};
  const cws = wb.getWorksheet(CONFIG_SHEET);
  if (cws) for (let r = 2; r <= (cws.rowCount || 0); r++) {
    const row = cws.getRow(r), k = sheetText(row, 1).trim().toLowerCase();
    if (k) cfg[k] = sheetText(row, 2).trim();
  }
  CONFIG = cfg;
  const next = {};
  const aws = wb.getWorksheet(CW.ALERTS_SHEET);
  if (aws) for (let r = 2; r <= (aws.rowCount || 0); r++) {
    const row = aws.getRow(r);
    const job = sheetText(row, 1).trim().toUpperCase(), email = sheetText(row, 2).trim().toLowerCase();
    if (!job || !email) continue;                      // a blanked line is a removed subscription
    const list = next[job] = next[job] || [];
    if (!list.some(x => x.email === email)) list.push({ email: email, who: sheetText(row, 3).trim(), when: sheetText(row, 4).trim() });
  }
  ALERTS = next;
  applyPendA();
}

const adminAddress = () => String(CONFIG.admin || "").trim().toLowerCase();
const adminDomain = () => adminAddress().split("@")[1] || "";
/* Client-side gating only. This decides what the dashboard offers; it is not a
   permission. Anyone who can edit the workbook can edit the sheet in Excel -
   SharePoint's own permissions are the real protection. */
function isAdmin() {
  const a = adminAddress();
  return !!a && String(whoAmI()).trim().toLowerCase() === a;
}

/** trim, lower-case, must look like an address, and must be in the same domain
    as the administrator's own address. Both come from the workbook. */
const ALERT_EMAIL_RE = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/;
function alertEmailCheck(raw) {
  const e = String(raw == null ? "" : raw).trim().toLowerCase();
  if (!e) return { ok: false, msg: "Type an email address first." };
  if (!ALERT_EMAIL_RE.test(e)) return { ok: false, msg: "That is not an email address." };
  const dom = adminDomain();
  if (!dom) return { ok: false, msg: "No administrator is set in the " + CONFIG_SHEET + " sheet, so nothing can be added." };
  if (e.split("@")[1] !== dom) return { ok: false, msg: "Only addresses in the administrator's own organisation can be added." };
  return { ok: true, email: e };
}

const alertsFor = job => (ALERTS[job] || []).slice().sort((a, b) => a.email.localeCompare(b.email));
/** Every address already used anywhere, for the one-tap quick-add chips. */
function alertAddresses() {
  const seen = {};
  Object.keys(ALERTS).forEach(j => (ALERTS[j] || []).forEach(a => { seen[a.email] = 1; }));
  return Object.keys(seen).sort();
}
function alertsByAddress() {
  const m = {};
  Object.keys(ALERTS).forEach(job => (ALERTS[job] || []).forEach(a =>
    (m[a.email] = m[a.email] || []).push({ job: job, who: a.who, when: a.when })));
  Object.keys(m).forEach(e => m[e].sort((x, y) => x.job.localeCompare(y.job)));
  return m;
}
/** A finished job is never emailed: it has left the "In production" section,
    or its row has gone gold. Nobody is told - it just stops. While the section
    names are unknown (nothing loaded yet) nothing is called finished. */
function alertFinished(j) {
  if (!j) return true;                       // no longer on the Production sheet at all
  if (j.done) return true;
  const idx = sectionIdx("In production");
  return idx < 0 ? false : j.blk !== idx;
}

/** Subscribe one address to one job: hold it, write it, log it. Resolves
    {ok, email} or {ok:false, msg} - the caller shows msg inline. */
async function addJobAlert(job, raw) {
  if (!isAdmin()) return { ok: false, msg: ALERT_ADMIN_ONLY };
  const v = alertEmailCheck(raw);
  if (!v.ok) return v;
  const id = String(job).trim().toUpperCase(), email = v.email, who = whoAmI();
  if ((ALERTS[id] || []).some(x => x.email === email)) return { ok: true, already: true, email: email };
  pendAlert(id, email, "add", who);
  alPut(id, email, who);
  refreshAlerts();
  try {
    await CW.addAlert(id, email, who);
    noteChange(id, "Alert", "", email);
    scheduleReconcile();
    return { ok: true, email: email };
  } catch (e) {
    dropPendAlert(id, email); alDrop(id, email); refreshAlerts();
    toast(friendly(e), true);
    return { ok: false, msg: friendly(e) };
  }
}

/** Unsubscribe one address from one job. */
async function removeJobAlert(job, email) {
  if (!isAdmin()) return { ok: false, msg: ALERT_ADMIN_ONLY };
  const id = String(job).trim().toUpperCase(), e = String(email).trim().toLowerCase();
  const had = (ALERTS[id] || []).find(x => x.email === e);
  if (!had) return { ok: true, already: true, email: e };
  pendAlert(id, e, "del", whoAmI());
  alDrop(id, e);
  delete ALCONF[alKey(id, e)];
  refreshAlerts();
  try {
    await CW.removeAlert(id, e);
    noteChange(id, "Alert", e, "");
    scheduleReconcile();
    return { ok: true, email: e };
  } catch (err) {
    dropPendAlert(id, e); alPut(id, e, had.who, had.when); refreshAlerts();
    toast(friendly(err), true);
    return { ok: false, msg: friendly(err) };
  }
}

/** The selection bar's "Alert to…": one address, every ticked job, one log
    line each. The writes are queued per sheet inside graph.js. */
async function alertMany(jobs, raw) {
  if (!isAdmin()) { toast(ALERT_ADMIN_ONLY, true); return { ok: false, msg: ALERT_ADMIN_ONLY, added: 0 }; }
  const v = alertEmailCheck(raw);
  if (!v.ok) return { ok: false, msg: v.msg, added: 0 };
  setStatus("subscribing " + jobs.length + " job" + (jobs.length > 1 ? "s" : "") + "…", "busy");
  let added = 0, stopped = null;
  for (let i = 0; i < jobs.length; i++) {
    const r = await addJobAlert(jobs[i], v.email);
    /* stop at the first refusal rather than working through the rest: whatever
       Excel objected to will object again, and the person would get a toast
       per job. addJobAlert has already said what went wrong once. */
    if (!r.ok) { stopped = r.msg; break; }
    added++;
  }
  if (stopped) {
    /* the status line was ours to set, so it is ours to correct - but this is
       not "live", it is an edit that did not land */
    setStatus("alerts: " + added + " of " + jobs.length + " saved", "err");
    return { ok: false, msg: stopped, added: added };
  }
  setStatus("live");
  toast(added + " of " + jobs.length + " job" + (jobs.length > 1 ? "s" : "") + " will be emailed");
  state.picked = {}; renderAll();
  return { ok: true, added: added, email: v.email };
}

/** Redraw whatever alert UI happens to be open. */
function refreshAlerts() {
  if ($("#ahost")) renderAlertsWindow();
  if ($("#dhost") && state.sel && byId(state.sel)) renderDrawer();
}

/* ---- the drawer's Alerts section ---- */
const alWhen = w => String(w || "").slice(0, 10);
function alChipHtml(job, a, admin) {
  const asking = !!ALCONF[alKey(job, a.email)];
  return '<span class="alchip"><span class="alwho"><b>' + esc(a.email) + '</b>' +
    (a.who || a.when ? '<span class="alsub">added by ' + esc(shortWho(a.who) || "someone") +
      (a.when ? " · " + esc(alWhen(a.when)) : "") + '</span>' : "") + '</span>' +
    (admin ? (asking
      ? '<span class="alconf">Remove?<button class="albtn yes" data-al-yes="' + esc(a.email) + '">Yes</button>' +
        '<button class="albtn" data-al-no="' + esc(a.email) + '">No</button></span>'
      : '<button class="alx" data-al-x="' + esc(a.email) + '" title="Remove this address">&times;</button>') : "") +
    '</span>';
}
function alertsSectionHtml(j) {
  const list = alertsFor(j.id), admin = isAdmin();
  const known = alertAddresses().filter(e => !list.some(x => x.email === e));
  return '<div class="sect"><span class="kick">Alerts (' + list.length + ')</span>' +
    (alertFinished(j) ? '<div class="alfin">Finished — no more emails for this job.</div>' : "") +
    (list.length ? '<div class="alchips">' + list.map(a => alChipHtml(j.id, a, admin)).join("") + '</div>'
                 : '<div style="font-size:13px;color:var(--ink-4)">Nobody is subscribed to this job.</div>') +
    (admin
      ? '<div class="alrow"><input class="alin" id="alnew" type="email" inputmode="email" autocomplete="off" ' +
          'spellcheck="false" placeholder="Email address"><button class="btn" id="aladd">Save</button></div>' +
        '<div class="alerr" id="alerr" hidden></div>' +
        (known.length ? '<div class="alquick"><span class="kick">Already used elsewhere</span>' +
          known.map(e => '<button class="chip alq" data-al-quick="' + esc(e) + '">' + esc(e) + '</button>').join("") + '</div>' : "")
      : '<div style="font-size:12px;color:var(--ink-4)">' + ALERT_ADMIN_ONLY + '</div>') +
    '<div class="alnote">' + ALERT_NOTE + '</div></div>';
}
function wireAlerts(host, id) {
  const err = m => { const b = $("#alerr"); if (b) { b.textContent = m; b.hidden = false; } };
  const save = async raw => {
    const r = await addJobAlert(id, raw);
    if (!r.ok) { err(r.msg); return; }
    /* the address is a chip now, so empty the box - and let it go, or the next
       re-render would put the text back out of the still-focused field */
    const box = $("#alnew");
    if (box) { box.value = ""; try { box.blur(); } catch (e) {} }
  };
  const add = $("#aladd");
  if (add) add.onclick = () => { const box = $("#alnew"); save(box ? box.value : ""); };
  const box = $("#alnew");
  if (box) box.onkeydown = e => { if (e.key === "Enter") { e.preventDefault(); save(box.value); } };
  host.querySelectorAll("[data-al-quick]").forEach(b => b.onclick = () => save(b.dataset.alQuick));
  host.querySelectorAll("[data-al-x]").forEach(b => b.onclick = () => {
    ALCONF[alKey(id, b.dataset.alX)] = 1; refreshAlerts();
  });
  host.querySelectorAll("[data-al-no]").forEach(b => b.onclick = () => {
    delete ALCONF[alKey(id, b.dataset.alNo)]; refreshAlerts();
  });
  host.querySelectorAll("[data-al-yes]").forEach(b => b.onclick = () => removeJobAlert(id, b.dataset.alYes));
}

/* ---- the Alerts window: every subscription, grouped by address ---- */
function renderAlertsWindow() {
  let host = $("#ahost");
  if (!host) { host = document.createElement("div"); host.id = "ahost"; document.body.appendChild(host); }
  const by = alertsByAddress(), addrs = Object.keys(by).sort(), admin = isAdmin();
  const n = addrs.reduce((a, e) => a + by[e].length, 0);
  host.innerHTML = '<div class="scrim" id="ascrim"></div><div class="logwin">' +
    '<div class="dhead"><div><div class="cond" style="font-size:25px;font-weight:700">Alerts</div>' +
      '<div style="font-size:12.5px;color:#a8a49a;margin-top:2px">' + n + ' subscription' + (n === 1 ? "" : "s") +
      ' across ' + addrs.length + ' address' + (addrs.length === 1 ? "" : "es") + ' · ' + ALERT_NOTE + '</div></div>' +
      '<button class="ghost" id="aclose">Close</button></div>' +
    '<div class="logbody">' + (addrs.length ? addrs.map(e => {
      const rows = by[e];
      return '<div class="agroup"><div class="ahead"><span class="gname">' + esc(e) + '</span>' +
        '<span class="gcount">' + rows.length + ' job' + (rows.length === 1 ? "" : "s") + '</span></div>' +
        rows.map(r => {
          const j = byId(r.job), fin = alertFinished(j), pair = esc(r.job) + "|" + esc(e);
          /* the same inline "Remove?" as the drawer chip, sharing ALCONF, so a
             stray tap in either place never removes anything on its own */
          const asking = !!ALCONF[alKey(r.job, e)];
          return '<div class="arow"><button class="stn jump" data-j="' + esc(r.job) + '" style="border:0;cursor:pointer">' +
            esc(r.job) + '</button>' +
            '<span class="ell">' + esc(j ? (j.cust || "—") : "—") + '</span>' +
            '<span class="ell asect" style="color:var(--ink-3)">' + esc(j ? (BLOCKNAMES[j.blk] || "—") : "not on the sheet") + '</span>' +
            '<span>' + (fin ? '<span class="badge" style="background:var(--surface-2);color:var(--ink-3)">Finished</span>'
                            : '<span class="badge" style="background:var(--green-bg);color:var(--green)">Emailing</span>') + '</span>' +
            '<span class="aact">' + (admin ? (asking
              ? '<span class="alconf">Remove?<button class="albtn yes" data-a-yes="' + pair + '">Yes</button>' +
                '<button class="albtn" data-a-no="' + pair + '">No</button></span>'
              : '<button class="albtn" data-a-ask="' + pair + '">Remove</button>') : "") + '</span>' +
            '</div>';
        }).join("") + '</div>';
    }).join("") : '<div class="empty">No job has an email alert yet.' +
        (admin ? '<br><span style="font-size:12px">Open a job and add an address in its Alerts section.</span>' : "") + '</div>') +
    '</div><div class="foot"><span>' + (admin ? "You can add and remove alerts" : ALERT_ADMIN_ONLY) +
    '</span><span>Click a job number to open it</span></div></div>';
  $("#ascrim").onclick = closeWin(host);
  $("#aclose").onclick = closeWin(host);
  host.querySelectorAll(".jump").forEach(b => b.onclick = () => {
    host.remove(); state.sel = b.dataset.j; state.edit = false; renderRows(); openDrawer();
  });
  const pairOf = v => { const i = String(v).indexOf("|"); return [String(v).slice(0, i), String(v).slice(i + 1)]; };
  host.querySelectorAll("[data-a-ask]").forEach(b => b.onclick = () => {
    const p = pairOf(b.dataset.aAsk); ALCONF[alKey(p[0], p[1])] = 1; renderAlertsWindow();
  });
  host.querySelectorAll("[data-a-no]").forEach(b => b.onclick = () => {
    const p = pairOf(b.dataset.aNo); delete ALCONF[alKey(p[0], p[1])]; renderAlertsWindow();
  });
  host.querySelectorAll("[data-a-yes]").forEach(b => b.onclick = () => {
    const p = pairOf(b.dataset.aYes); removeJobAlert(p[0], p[1]);
  });
  renderFab();                 // a window is open: the wheel steps aside
}

/** "Alert to…" for the ticked jobs: the addresses already in use, plus a box. */
function renderAlertMenu(anchor) {
  const old = $("#alertmenu"); if (old) { old.remove(); return; }
  const jobs = Object.keys(state.picked);
  const m = document.createElement("div"); m.id = "alertmenu"; m.className = "menu";
  const known = alertAddresses();
  m.innerHTML = '<div class="kick" style="padding:4px 10px 6px">Email alerts for ' + jobs.length +
      ' job' + (jobs.length > 1 ? "s" : "") + ' to</div>' +
    (known.length ? known.map(e => '<button class="mrow" data-al="' + esc(e) + '">' + esc(e) + '</button>').join("")
                  : '<div style="padding:2px 10px 6px;font-size:12px;color:var(--ink-4)">No address has been used yet.</div>') +
    '<div class="alrow" style="padding:8px 10px 4px;border-top:1px solid var(--line);margin-top:6px">' +
      '<input class="alin" id="almin" type="email" inputmode="email" autocomplete="off" spellcheck="false" placeholder="Email address">' +
      '<button class="btn" id="almadd">Add</button></div>' +
    '<div class="alerr" id="almerr" hidden style="margin:6px 10px"></div>' +
    '<div class="alnote" style="padding:6px 10px">' + ALERT_NOTE + '</div>';
  document.body.appendChild(m);
  menuAt(m, anchor);
  const go = async raw => {
    const res = await alertMany(jobs, raw);
    if (res.ok) m.remove();
    else { const b = $("#almerr"); if (b) { b.textContent = res.msg; b.hidden = false; } }
  };
  m.querySelectorAll("[data-al]").forEach(b => b.onclick = () => go(b.dataset.al));
  const addb = $("#almadd");
  if (addb) addb.onclick = () => { const box = $("#almin"); go(box ? box.value : ""); };
  setTimeout(() => document.addEventListener("click", function off(e) {
    if (!m.contains(e.target) && e.target !== anchor) { m.remove(); document.removeEventListener("click", off); }
  }), 0);
}

/* ---------- Export ----------
   The window that drives export.js. Everything it can do is decided here and
   carried out there: this half only collects choices, shows how many jobs they
   come to, and hands the finished file to the browser.

   Two things are worth saying out loud, because they are the promises made to
   the office when this was agreed:
   - Nothing in this path goes near the network. The jobs are the ones already
     in memory, the workbook and the PDF are built in the tab, and the file is
     handed over through an object URL. The single exception is the one
     Dashboard Log line written after the download, through the same noteChange
     every other change uses.
   - Phone numbers and eircodes are not options that happen to be switched off;
     they are not read at any point. See the note at the top of export.js.       */

let XSTATE = null;                 // the choices in the window, while it is open
let XBUSY = false;                 // a file is being built - the button waits
let XLOGO = null;                  // assets/logo.png as a data URL, when the file exists

function xpNewState() {
  return { format: "xlsx", layout: "cards", fields: exportAllFields(), f: exportDefaults(), preset: "",
           /* "default" is everything the window has always done; "john" is the
              fixed print layout, which has no field picker and no card/table
              choice because its whole point is to come out the same every time */
           template: "default",
           /* which groups are expanded: kept here rather than left to <details>,
              because a chip click redraws the window and would otherwise fold
              the group the person was working in */
           open: { filters: true, fields: true, sort: false, presets: false } };
}
/** The company name for the PDF header: the Config sheet if it says, otherwise
    the page's own title. Never a name written into the code. */
function xpCompany() {
  const t = String(CONFIG.company || document.title || "").split(/[·|—–]/)[0].trim();
  return t || "Production";
}
/** The logo is optional and always will be. It is read once at startup with an
    Image, not a fetch, and a missing file simply leaves XLOGO null - the PDF
    header then shows the company name on its own. */
function xpLoadLogo() {
  try {
    const img = new Image();
    img.onerror = () => { XLOGO = null; };
    img.onload = () => {
      try {
        const c = document.createElement("canvas");
        c.width = img.naturalWidth || img.width; c.height = img.naturalHeight || img.height;
        if (!c.width || !c.height) return;
        c.getContext("2d").drawImage(img, 0, 0);
        XLOGO = c.toDataURL("image/png");
      } catch (e) { XLOGO = null; }
    };
    img.src = "assets/logo.png";
  } catch (e) { XLOGO = null; }
}

/** The person's name for an exported file. The Dashboard Log keeps the full
    address, because that is an internal record; a workbook or a PDF often goes
    outside the office, so it carries a name and nothing else. */
function xpWhoName() {
  const a = CW.account;
  const n = (a && a.name) ? String(a.name).trim() : "";
  return n || String(whoAmI()).split("@")[0];
}

/* pdfmake is 1.4 MB and most days nobody exports a PDF, so it is not in the
   page: the two scripts are injected the first time one is asked for. The
   offline tests already have a pdfMake, so the loader hands that straight
   back. */
let XPDFLOAD = null;
function xpScript(src) {
  return new Promise((ok, no) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = false;                       // vfs_fonts.js must land after pdfmake
    s.onload = () => ok(src);
    s.onerror = () => no(new Error("could not load " + src));
    (document.head || document.body).appendChild(s);
  });
}
function xpLoadPdf() {
  if (typeof pdfMake !== "undefined" && pdfMake) return Promise.resolve(xpPdfReady());
  if (!XPDFLOAD) {
    XPDFLOAD = xpScript("vendor/pdfmake.min.js")
      .then(() => xpScript("vendor/vfs_fonts.js"))
      .then(() => xpPdfReady())
      .catch(e => { XPDFLOAD = null; throw e; });
  }
  return XPDFLOAD;
}

/** Everything export.js needs to know about this dashboard, right now. */
function xpCtxNow() {
  return { all: live(), view: filtered(), picked: state.picked, sections: BLOCKNAMES,
           /* Production (2)'s own rows, for the John print sheet and nothing
              else - no Default export reads them */
           john: JOHNROWS,
           comments: id => commentsFor(id), alerts: id => alertsFor(id) };
}
const xpMatched = () => exportJobs(xpCtxNow(), XSTATE.f);

const xpCounties = () => { const s = {}; live().forEach(j => { if (j.area) s[j.area] = 1; }); return Object.keys(s).sort(); };
const xpProductNames = () => { const s = {}; live().forEach(j => (j.prods || []).forEach(p => { s[p.n] = 1; })); return Object.keys(s).sort(); };
const xpGlassNames = () => { const s = {}; live().forEach(j => Object.keys(j.glass || {}).forEach(k => { if (j.glass[k] > 0) s[k] = 1; })); return Object.keys(s).sort(); };

/* ---- the little bits of markup the window is made of ---- */
const xpChip = (on, path, val, text) =>
  '<button class="xchip" aria-pressed="' + (on ? "true" : "false") + '" data-xset="' + path +
  '" data-xval="' + esc(String(val)) + '">' + esc(text) + "</button>";
const xpBox = (on, list, val, text) =>
  '<label class="xbox"><input type="checkbox" data-xtog="' + list + '" value="' + esc(String(val)) + '"' +
  (on ? " checked" : "") + "><span>" + esc(text) + "</span></label>";
/** One row of exclusive choices. `opts` is [value, label]; "" means "any". */
function xpRow(label, path, cur, opts) {
  return '<div class="xrow"><span class="xlab">' + esc(label) + "</span><span class=\"xopts\">" +
    opts.map(o => xpChip(String(cur == null ? "" : cur) === String(o[0]), path, o[0], o[1])).join("") +
    "</span></div>";
}
const XP_TRI = [["", "Any"], ["true", "Yes"], ["false", "No"]];
const XP_CPOPTS = [["", "Any"], ["none", "Not started"], ["process", "In fabrication"], ["done", "Done"]];

function xpDateRow(key, label) {
  const r = XSTATE.f.dates[key];
  return '<div class="xrow"><span class="xlab">' + esc(label) + "</span><span class=\"xopts\">" +
    '<input class="txt xdate" type="date" data-xdate="' + key + '.from" value="' + esc(r.from || "") + '" aria-label="' + esc(label) + ' from">' +
    '<span class="xto">to</span>' +
    '<input class="txt xdate" type="date" data-xdate="' + key + '.to" value="' + esc(r.to || "") + '" aria-label="' + esc(label) + ' to">' +
    '<button class="xchip" data-xdp="' + key + '|week">This week</button>' +
    '<button class="xchip" data-xdp="' + key + '|7">Last 7 days</button>' +
    '<button class="xchip" data-xdp="' + key + '|30">Last 30 days</button>' +
    '<button class="xchip" data-xdp="' + key + '|clear">Any</button>' +
    "</span></div>";
}

/** A collapsible group. Whether it is open is remembered in XSTATE, so a
    redraw leaves the window looking exactly as it did. */
const xpGroup = (title, inner, key) =>
  '<details class="xgrp" id="xg-' + key + '" data-xopen="' + key + '"' + (XSTATE.open[key] ? " open" : "") +
  "><summary>" + esc(title) + '</summary><div class="xgbody">' + inner + "</div></details>";

function renderExportWindow() {
  if (!XSTATE) XSTATE = xpNewState();
  /* opened from the John print sheet view, the window OPENS on that template -
     there is nothing else it could sensibly mean from in there. Only on the way
     in, though: this function re-runs on every click inside the window, and
     forcing it every time would make the Template choice unclickable. */
  if (state.board === "john" && !$("#xhost")) XSTATE.template = "john";
  let host = $("#xhost");
  if (!host) { host = document.createElement("div"); host.id = "xhost"; document.body.appendChild(host); }
  const keep = $("#xbody") ? $("#xbody").scrollTop : 0;
  const f = XSTATE.f, F = XSTATE.fields;
  const nPicked = Object.keys(state.picked).length;
  const presets = presetsLoad(), pnames = Object.keys(presets).sort();

  const sectionPick = BLOCKNAMES.length
    ? BLOCKNAMES.map((n, i) => xpBox(f.sections.indexOf(i) >= 0, "sections", i, n + " (" + live().filter(j => j.blk === i).length + ")")).join("")
    : '<span class="xnone">The sheet has no sections yet.</span>';

  const counties = xpCounties(), prods = xpProductNames(), glass = xpGlassNames();
  const john = XSTATE.template === "john";

  host.innerHTML = '<div class="scrim" id="xscrim"></div><div class="logwin xwin">' +
    '<div class="dhead"><div><div class="cond" style="font-size:25px;font-weight:700">Export</div>' +
      '<div style="font-size:12.5px;color:#a8a49a;margin-top:2px">' + (john
        ? "“Production (2)” as it prints for John, with the phone numbers. Eircodes are never included."
        : "Download the jobs you choose. No phone numbers and no eircodes are ever included.") +
      '</div></div>' +
      '<button class="ghost" id="xclose">Close</button></div>' +
    '<div class="logbody xbody" id="xbody">' +

      '<div class="xsec">' +
        xpRow("Template", "template", XSTATE.template, [["default", "Default"], ["john", "John print sheet"]]) +
        (john ? '<div class="xnote2">Fixed layout, printed from the <b>Production (2)</b> sheet: job no, ' +
                'ready to print, customer, phone no, area, windows, doors and the notes from Brendan’s ' +
                'office — in that sheet’s own order, with its section dividers, its row colours and its ' +
                'text colours. The field picker and the PDF layout do not apply. A chosen job that is not ' +
                'on that sheet is printed in grey, marked “not on John’s sheet”.</div>' : "") +
        xpRow("Format", "format", XSTATE.format, [["xlsx", "Excel workbook"], ["pdf", "PDF"]]) +
        (XSTATE.format === "pdf" && !john
          ? xpRow("PDF layout", "layout", XSTATE.layout, [["cards", "Job cards"], ["table", "Table"]])
          : "") +
        (john
          ? xpRow("Scope", "f.scope", f.scope === "ticked" ? "ticked" : "view",
                  [["view", state.board === "john" ? "All of John’s sheet" : "What I see now"],
                   ["ticked", "Ticked jobs (" + nPicked + ")"]])
          : xpRow("Scope", "f.scope", f.scope, [["view", "What I see now"], ["ticked", "Ticked jobs (" + nPicked + ")"], ["sections", "Sections…"]])) +
        (!john && f.scope === "sections" ? '<div class="xrow"><span class="xlab"></span><span class="xpick">' + sectionPick + "</span></div>" : "") +
      "</div>" +

      (john ? "" :
      xpGroup("Filters", '<div class="xsec">' +
        xpRow("Ready to deliver", "f.ready", f.ready, XP_TRI) +
        xpRow("Urgent", "f.urgent", f.urgent, [["", "Any"], ["true", "Urgent only"]]) +
        xpRow("Windows", "f.cp.win", f.cp.win, XP_CPOPTS) +
        xpRow("Doors", "f.cp.drs", f.cp.drs, XP_CPOPTS) +
        xpRow("Glass", "f.cp.glass", f.cp.glass, XP_CPOPTS) +
        xpRow("Products", "f.cp.prod", f.cp.prod, XP_CPOPTS) +
        xpDateRow("sold", "Sold") +
        xpDateRow("ready", "Ready to print") +
        xpDateRow("floor", "Sent to floor") +
        '<div class="xrow"><span class="xlab">County</span><span class="xpick">' +
          (counties.length ? counties.map(c => xpBox(f.county.indexOf(c) >= 0, "county", c, c)).join("") : '<span class="xnone">none</span>') +
        "</span></div>" +
        '<div class="xrow"><span class="xlab">Products</span><span class="xpick">' +
          (prods.length ? prods.map(p => xpBox(f.products.indexOf(p) >= 0, "products", p, p.toUpperCase())).join("") : '<span class="xnone">none</span>') +
        "</span></div>" +
        '<div class="xrow"><span class="xlab">Glass types</span><span class="xpick">' +
          (glass.length ? glass.map(g => xpBox(f.glassTypes.indexOf(g) >= 0, "glassTypes", g, g.toUpperCase())).join("") : '<span class="xnone">none</span>') +
        "</span></div>" +
        '<div class="xrow"><span class="xlab">On sheet</span><span class="xpick">' +
          SHEETNAMES.map(s => xpBox(f.sheets.indexOf(s) >= 0, "sheets", s, s)).join("") +
        "</span></div>" +
        xpRow("Has comments", "f.hasComments", f.hasComments, XP_TRI) +
        xpRow("Has email alerts", "f.hasAlerts", f.hasAlerts, XP_TRI) +
        '<div class="xrow"><span class="xlab">Search</span><span class="xopts">' +
          '<input class="txt xq" id="xq" type="search" placeholder="Job no, customer, county, office no or a comment" value="' + esc(f.q) + '">' +
        "</span></div>" +
      "</div>", "filters") +

      xpGroup("Fields", '<div class="xsec">' +
        '<div class="xrow"><span class="xlab">Include</span><span class="xopts">' +
          '<button class="xchip" id="xfall">All</button><button class="xchip" id="xfnone">None</button>' +
        "</span></div>" +
        '<div class="xrow"><span class="xlab"></span><span class="xpick">' +
          EXPORT_FIELDS.map(p => xpBox(!!F[p[0]], "fields", p[0], p[1])).join("") +
        "</span></div>" +
        '<div class="xnote2">Phone numbers and eircodes are never exported and cannot be turned on.</div>' +
      "</div>", "fields") +

      xpGroup("Sort & group", '<div class="xsec">' +
        '<div class="xrow"><span class="xlab">Sort by</span><span class="xopts">' +
          '<select class="txt" id="xsort">' + XP_SORTS.map(p => '<option value="' + p[0] + '"' +
            (f.sort === p[0] ? " selected" : "") + ">" + esc(p[1]) + "</option>").join("") + "</select>" +
          xpChip(!f.desc, "f.desc", "false", "▲ normal") + xpChip(!!f.desc, "f.desc", "true", "▼ reversed") +
        "</span></div>" +
        xpRow("Group by section", "f.groupBySection", !!f.groupBySection, [["false", "No"], ["true", "Yes"]]) +
      "</div>", "sort") +

      xpGroup("Presets", '<div class="xsec">' +
        '<div class="xrow"><span class="xlab">Saved</span><span class="xopts">' +
          '<select class="txt" id="xpsel">' + (pnames.length
            ? pnames.map(n => '<option value="' + esc(n) + '"' + (XSTATE.preset === n ? " selected" : "") + ">" + esc(n) + "</option>").join("")
            : '<option value="">no presets yet</option>') + "</select>" +
          '<button class="xchip" id="xpload"' + (pnames.length ? "" : " disabled") + ">Load</button>" +
          '<button class="xchip" id="xpdel"' + (pnames.length ? "" : " disabled") + ">Delete</button>" +
        "</span></div>" +
        '<div class="xrow"><span class="xlab">Save as</span><span class="xopts">' +
          '<input class="txt xq" id="xpname" placeholder="Name this set of choices" maxlength="40" value="">' +
          '<button class="xchip" id="xpsave">Save</button>' +
        "</span></div>" +
        '<div class="xnote2">A preset remembers the filters, the fields and the format. It never holds job data.</div>' +
      "</div>", "presets")) +

    "</div>" +
    '<div class="foot xfoot"><span id="xcount">…</span>' +
      (john ? '<button class="btn" id="xnotes">Continue to notes</button>'
            : '<button class="btn" id="xdl">Download</button>') + "</div>" +
    "</div>";

  if ($("#xbody")) $("#xbody").scrollTop = keep;
  xpWire(host);
  xpUpdateCount();
  renderFab();                 // the wheel is under a window now
}

/** Walk "f.cp.win" and put the value at the end of it. "" means "any". */
function xpSet(path, raw) {
  const v = raw === "" ? null : raw === "true" ? true : raw === "false" ? false : raw;
  const p = String(path).split(".");
  let o = XSTATE;
  for (let i = 0; i < p.length - 1; i++) o = o[p[i]];
  o[p[p.length - 1]] = v;
}
/** Tick or untick one value in one of the list filters. */
function xpToggleList(list, value, on) {
  if (list === "fields") { if (on) XSTATE.fields[value] = true; else delete XSTATE.fields[value]; return; }
  const arr = list === "sections" ? XSTATE.f.sections : XSTATE.f[list];
  const v = list === "sections" ? Number(value) : value;
  const i = arr.indexOf(v);
  if (on && i < 0) arr.push(v);
  if (!on && i >= 0) arr.splice(i, 1);
  if (list === "sections") XSTATE.f.sectionNames = XSTATE.f.sections.map(x => BLOCKNAMES[x] || ("section " + x));
}
/** This week (from Monday), the last 7 or 30 days including today, or clear. */
function xpDatePreset(key, kind) {
  const r = XSTATE.f.dates[key];
  if (kind === "clear") { r.from = null; r.to = null; return; }
  const now = new Date(), d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (kind === "week") d.setDate(d.getDate() - ((now.getDay() + 6) % 7));
  else d.setDate(d.getDate() - (kind === "7" ? 6 : 29));
  r.from = xpIsoDate(d); r.to = xpIsoDate(now);
}

/** The live count, and whether Download can do anything with it. Ticking every
    field off used to make ExcelJS throw on a table with no columns and leave
    the PDF as a cover page with nothing behind it, so the reason is shown
    beside the count rather than found out the hard way. */
function xpUpdateCount() {
  /* the John print sheet has no fields to tick and no layout to get wrong: the
     only question it can fail on is whether there is a job in the scope */
  if (XSTATE.template === "john") {
    let jn = 0;
    try { jn = xpJohnIds().length; } catch (e) { jn = 0; }
    const jc = $("#xcount");
    if (jc) jc.textContent = jn + " job" + (jn === 1 ? "" : "s") + (jn ? "" : " — nothing to print");
    const jb = $("#xnotes");
    if (jb) jb.disabled = XBUSY || !jn;
    return jn;
  }
  let n = 0, jobs = [];
  try { jobs = xpMatched(); n = jobs.length; } catch (e) { jobs = []; n = 0; }
  let can = { ok: false, why: "no jobs match" };
  try {
    can = exportBuildable(exportRows(jobs, XSTATE.fields, xpCtxNow()), XSTATE.fields, XSTATE.format, XSTATE.layout);
  } catch (e) { can = { ok: false, why: (e && e.message) || String(e) }; }
  const c = $("#xcount");
  if (c) c.textContent = n + " job" + (n === 1 ? "" : "s") + " match" + (n && !can.ok ? " — " + can.why : "");
  const b = $("#xdl");
  if (b) b.disabled = XBUSY || !can.ok;
  return n;
}

function xpWire(host) {
  const close = closeWin(host);
  if ($("#xscrim")) $("#xscrim").onclick = close;
  if ($("#xclose")) $("#xclose").onclick = close;

  host.querySelectorAll("[data-xset]").forEach(b => b.onclick = () => {
    xpSet(b.dataset.xset, b.dataset.xval); renderExportWindow();
  });
  host.querySelectorAll("[data-xtog]").forEach(cb => cb.onchange = () => {
    xpToggleList(cb.dataset.xtog, cb.value, cb.checked); xpUpdateCount();
  });
  host.querySelectorAll("[data-xdate]").forEach(inp => inp.onchange = () => {
    const p = String(inp.dataset.xdate).split(".");
    XSTATE.f.dates[p[0]][p[1]] = inp.value || null;
    xpUpdateCount();
  });
  host.querySelectorAll("[data-xopen]").forEach(d => d.ontoggle = () => { XSTATE.open[d.dataset.xopen] = !!d.open; });
  host.querySelectorAll("[data-xdp]").forEach(b => b.onclick = () => {
    const p = String(b.dataset.xdp).split("|");
    xpDatePreset(p[0], p[1]); renderExportWindow();
  });

  const q = $("#xq");
  /* the box keeps the focus while you type: only the count is redrawn */
  if (q) q.oninput = () => { XSTATE.f.q = q.value; xpUpdateCount(); };
  const sort = $("#xsort");
  if (sort) sort.onchange = () => { XSTATE.f.sort = sort.value; xpUpdateCount(); };
  if ($("#xfall")) $("#xfall").onclick = () => { XSTATE.fields = exportAllFields(); renderExportWindow(); };
  if ($("#xfnone")) $("#xfnone").onclick = () => { XSTATE.fields = {}; renderExportWindow(); };

  if ($("#xpsave")) $("#xpsave").onclick = () => {
    const box = $("#xpname"), name = String(box ? box.value : "").trim();
    if (!name) { toast("Give the preset a name first", true); return; }
    presetSave(name, { format: XSTATE.format, layout: XSTATE.layout, fields: XSTATE.fields, f: XSTATE.f });
    XSTATE.preset = name;
    toast('Preset "' + name + '" saved');
    renderExportWindow();
  };
  if ($("#xpload")) $("#xpload").onclick = () => {
    const sel = $("#xpsel"), name = sel ? sel.value : "";
    const p = presetsLoad()[name];
    if (!p) return;
    XSTATE.format = p.format; XSTATE.layout = p.layout;
    XSTATE.fields = p.fields; XSTATE.f = p.f; XSTATE.preset = name;
    XSTATE.f.sectionNames = XSTATE.f.sections.map(x => BLOCKNAMES[x] || ("section " + x));
    renderExportWindow();
  };
  if ($("#xpdel")) $("#xpdel").onclick = () => {
    const sel = $("#xpsel"), name = sel ? sel.value : "";
    if (!name) return;
    presetDelete(name);
    if (XSTATE.preset === name) XSTATE.preset = "";
    toast('Preset "' + name + '" deleted');
    renderExportWindow();
  };
  if ($("#xdl")) $("#xdl").onclick = () => xpDownload();
  /* the John print sheet never downloads straight from here: the notes window is
     the last step, and it is what builds the file */
  if ($("#xnotes")) $("#xnotes").onclick = () => openNotesWindow();
}

/** Build the file, hand it to the browser, then write the one log line.
    The log line is last on purpose: nothing is claimed to have been exported
    until the bytes have actually been handed over. */
async function xpDownload() {
  if (XBUSY) return null;
  if (XSTATE.template === "john") return null;     // that file is built by the notes window
  const ctx = xpCtxNow(), f = XSTATE.f;
  const jobs = exportJobs(ctx, f);
  const rows = exportRows(jobs, XSTATE.fields, ctx);
  const can = exportBuildable(rows, XSTATE.fields, XSTATE.format, XSTATE.layout);
  if (!can.ok) return null;
  XBUSY = true; xpUpdateCount();
  const btn = $("#xdl");
  if (btn) btn.textContent = "Building…";
  let name = null;
  try {
    const when = new Date();
    name = exportFilename(XSTATE.format, f, when, BLOCKNAMES);
    const buildEl = $("#build");
    const opts = {
      fields: XSTATE.fields, filters: f, layout: XSTATE.layout,
      who: xpWhoName(), when: when, company: xpCompany(), sections: BLOCKNAMES, logo: XLOGO,
      build: buildEl ? String(buildEl.textContent || "").replace("build ", "").trim() : ""
    };
    if (XSTATE.format === "pdf") {
      if (btn && typeof pdfMake === "undefined") btn.textContent = "Loading PDF engine…";
      const pm = await xpLoadPdf();
      if (btn) btn.textContent = "Building…";
      pm.createPdf(buildDocDefinition(rows, opts)).download(name);
    } else {
      const wb = buildWorkbook(rows, opts);
      const buf = await wb.xlsx.writeBuffer();
      downloadBlob(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), name);
    }
    noteChange("(export)", "Export", exportLogFrom(XSTATE.format, jobs.length), filtersSummary(f, BLOCKNAMES));
    toast("Downloading " + name);
  } catch (e) {
    name = null;
    toast("Export failed: " + ((e && e.message) || String(e)).slice(0, 160), true);
  }
  XBUSY = false;
  if ($("#xdl")) $("#xdl").textContent = "Download";
  xpUpdateCount();
  return name;
}

/* ---------- the John print sheet and its notes -------------------------------
   A second, fixed export layout (export.js §7b) plus the one window that
   drives it. Two things about it are worth having in front of you before
   changing a line of this block:

   1. It is the ONLY export that carries a phone number. The owner sanctioned
      that on 2026-09-09 for this template alone; the eircode is not carried
      here or anywhere else. Every John print writes one Dashboard Log line
      that says, in words, that the file has phone numbers in it.
   2. The notes people type here are the dashboard's, not the sheet's. They go
      to the SharePoint list "Dashboard print notes" - Title = job number,
      Note, By, At - through listUpsert and nothing else. Not one cell of any
      workbook sheet is written by this block, including the Production sheet.
      A note that will not save does not stop the print: the file is built from
      what was typed, and the row says it was not saved.

   The list is made by hand in SharePoint, like every other list here. If it is
   not there the window says so plainly and the print still works.

   The rows the file is built from come from "Production (2)" (export.js §7b),
   never from this block: nothing here touches that sheet either.           */
const PRINT_NOTE_LIST = "Dashboard print notes";
const PRINT_NOTE_FIELDS = ["Title", "Note", "By", "At"];
const PRINT_NOTE_MISSING = "The “Dashboard print notes” list is not in SharePoint yet, so notes cannot be " +
  "saved. Ask the manager to add it — nothing in the Excel file is involved. You can still print with what you type here.";
const PRINT_NOTE_CONSENT_MSG = "Saving print notes needs a SharePoint permission that has not been granted yet. " +
  "Nothing in the Excel file is involved. You can still print with what you type here.";
let PRINTNOTES = {};            // { JOB: { note, who, at } } - everyone's print notes
let PRINT_NOTE_OK = null;       // null: not looked yet · false: no list, or no permission · true: read it
let PRINT_NOTE_CONSENT = false; // the last read failed for want of the list permission
let printNotesReading = null;   // the read in flight, so two callers share one

/** Read the whole list, quietly. Called whenever the print-notes window
    opens - never on every page load, because most days nobody prints John's
    sheet at all. */
function readPrintNotes() {
  if (printNotesReading) return printNotesReading;
  printNotesReading = (async () => {
    /* no Graph layer at all (a test harness, or a page that failed to load it):
       that is "the list cannot be read", not "not looked yet" - leaving it at
       null would keep the window saying "reading…" for ever */
    if (typeof CW === "undefined" || !CW || typeof CW.listItems !== "function") {
      PRINT_NOTE_OK = false; PRINT_NOTE_CONSENT = false;
      return PRINTNOTES;
    }
    try {
      const items = await CW.listItems(PRINT_NOTE_LIST, { fields: PRINT_NOTE_FIELDS });
      if (items == null) { PRINT_NOTE_OK = false; PRINT_NOTE_CONSENT = false; PRINTNOTES = {}; return PRINTNOTES; }
      PRINT_NOTE_OK = true; PRINT_NOTE_CONSENT = false;
      const next = {};
      items.forEach(it => {
        const f = it.fields || {};
        const job = String(f.Title == null ? "" : f.Title).trim().toUpperCase();
        if (!job) return;
        const at = String(f.At == null ? "" : f.At);
        /* two browsers can add the same job at once on a list without the
           unique-Title rule; the newest is what somebody most recently meant */
        const seen = next[job];
        if (seen && seen.at >= at) return;
        next[job] = { note: String(f.Note == null ? "" : f.Note),
                      who: String(f.By == null ? "" : f.By), at: at };
      });
      PRINTNOTES = next;
    } catch (e) {
      PRINT_NOTE_CONSENT = /permission needed/.test((e && e.message) || "");
      PRINT_NOTE_OK = false;
    }
    return PRINTNOTES;
  })().then(v => { printNotesReading = null; return v; },
            e => { printNotesReading = null; throw e; });
  return printNotesReading;
}
const printNoteInfo = id => PRINTNOTES[String(id).toUpperCase()] || null;
const printNoteOf = id => String((printNoteInfo(id) || {}).note || "");
/** Why notes cannot be saved right now, in words, or "" when they can. */
function printNoteTrouble() {
  if (PRINT_NOTE_OK === true) return "";
  if (PRINT_NOTE_OK === null) return "";                 // the first read has not answered yet
  return PRINT_NOTE_CONSENT ? PRINT_NOTE_CONSENT_MSG : PRINT_NOTE_MISSING;
}

/** The job numbers a John print covers: the ticked ones, or everything the
    view in front of the person is showing. The John print sheet has no filters
    of its own on purpose - the two views are the filter.

    From the John print sheet view "what I see now" means that sheet's own rows
    (searched), not the master list's; from the master it means `filtered()`. */
function xpJohnIds() {
  if (state.board === "john") {
    const shown = johnShown();
    if (XSTATE && XSTATE.f && XSTATE.f.scope === "ticked") return shown.filter(r => state.picked[r.id]).map(r => r.id);
    return shown.map(r => r.id);
  }
  const ctx = xpCtxNow();
  if (XSTATE && XSTATE.f && XSTATE.f.scope === "ticked") return ctx.all.filter(j => ctx.picked[j.id]).map(j => j.id);
  return (ctx.view || []).map(j => j.id);
}

/* the notes window's own state, while it is open */
let NSTATE = null;              // { ids, typed:{JOB:text}, failed:{JOB:why}, busy }
let NBUSY = false;

async function openNotesWindow() {
  const ids = xpJohnIds();
  if (!ids.length) { toast("There are no jobs to print", true); return null; }
  NSTATE = { ids: ids.slice(), typed: {}, failed: {}, busy: false };
  renderNotesWindow();
  /* the window opening is one of the two moments the list is read */
  try { await readPrintNotes(); } catch (e) {}
  if (NSTATE) renderNotesWindow();
  return NSTATE;
}
/** Has anybody typed something that is not already saved? */
function notesUnsaved() {
  if (!NSTATE) return false;
  return Object.keys(NSTATE.typed).some(k => String(NSTATE.typed[k]).trim() !== printNoteOf(k).trim());
}
function closeNotesWindow(ask) {
  /* clicking the scrim by accident with half a dozen notes typed into the
     window is a real way to lose work, so that one asks first; Cancel and
     Escape are deliberate and do not */
  if (ask && notesUnsaved() && typeof confirm === "function" &&
      !confirm("Close without printing? The notes you have typed will be lost.")) return false;
  NSTATE = null;
  const h = $("#nhost");
  if (h) h.remove();
  renderFab();
  return true;
}
/** What the box for one job currently holds: what was typed, or the note the
    list already has. */
function noteValue(id) {
  const k = String(id).toUpperCase();
  const typed = NSTATE && NSTATE.typed;
  return typed && typed[k] != null ? String(typed[k]) : printNoteOf(id);
}

function renderNotesWindow() {
  if (!NSTATE) return null;
  let host = $("#nhost");
  if (!host) { host = document.createElement("div"); host.id = "nhost"; document.body.appendChild(host); }
  const keep = $("#nbody") ? $("#nbody").scrollTop : 0;
  const trouble = printNoteTrouble();
  const checking = PRINT_NOTE_OK === null;
  const ctx = xpCtxNow();
  const rows = NSTATE.ids.slice();

  host.innerHTML = '<div class="scrim" id="nscrim"></div><div class="logwin nwin">' +
    '<div class="dhead"><div><div class="cond" style="font-size:25px;font-weight:700">Print notes — John print sheet</div>' +
      '<div style="font-size:12.5px;color:#a8a49a;margin-top:2px">' + rows.length + " job" + (rows.length === 1 ? "" : "s") +
      ", in the sheet’s order. Notes are kept by the dashboard — the Excel file is never changed.</div></div>" +
      '<button class="ghost" id="nclose">Cancel</button></div>' +
    (checking ? '<div class="nnote">Reading the notes already saved…</div>' : "") +
    (trouble ? '<div class="nnote err">' + esc(trouble) + "</div>" : "") +
    '<div class="logbody nbody" id="nbody">' +
      (rows.length ? rows.map(id => {
        const k = String(id).toUpperCase();
        const jr = johnRowFor(ctx, id);            // Production (2)'s own row
        const j = byId(id);                        // and the Production model, if it has it
        const flag = jr ? johnFlag(jr) : (j ? j.flag : "");
        const ink = jr ? inkFor(jr.inkHex, false) : (j ? flagInk(j) : "");
        /* the grey line is Production (2)'s note - the one this print will
           carry - with the Production comment underneath it, labelled, when the
           two say different things */
        const sheetNote = jr ? String(jr.notes || "") : "";
        const prod = j ? xpJohnNotes(j, "") : "";
        const extra = prod && prod.toLowerCase() !== sheetNote.toLowerCase() ? prod : "";
        const info = printNoteInfo(id);
        return '<div class="nrow" data-job="' + esc(id) + '">' +
          '<div class="nwho"><span class="tab" style="font-weight:600">' + esc(id) + "</span>" +
            (flag ? '<span class="badge flagchip" style="color:' + ink + '">' + esc(FLAGWORD[flag]) + "</span>" : "") +
            '<div class="ell" style="font-size:12.5px;color:var(--ink-2)">' +
              esc((jr && jr.cust) || (j && j.cust) || "—") + "</div>" +
            (jr ? "" : '<div class="nfail">not on John’s sheet</div>') +
            (NSTATE.failed[k] ? '<div class="nfail">not saved</div>' : "") + "</div>" +
          '<div class="nsheet">' +
            (sheetNote ? esc(sheetNote) : '<span style="color:var(--ink-4)">no note on John’s sheet</span>') +
            (extra ? '<div style="margin-top:4px;color:var(--ink-4)">Production comment: ' + esc(extra) + "</div>" : "") +
            (info && info.who ? '<div style="margin-top:4px;color:var(--ink-4)">last print note by ' +
              esc(shortWho(info.who)) + (info.at ? " · " + esc(stamp(info.at)) : "") + "</div>" : "") + "</div>" +
          '<textarea class="nbox" data-note="' + esc(id) + '" rows="2" maxlength="2000" ' +
            'placeholder="No print note yet">' + esc(noteValue(id)) + "</textarea>" +
          "</div>";
      }).join("") : '<div class="empty">No jobs.</div>') +
    "</div>" +
    '<div class="foot xfoot"><span id="ncount">' + rows.length + " job" + (rows.length === 1 ? "" : "s") +
      " · " + (XSTATE.format === "pdf" ? "PDF" : "Excel") + "</span>" +
      '<span style="display:flex;gap:8px"><button class="btn sec" id="ncancel">Cancel</button>' +
      '<button class="btn" id="nprint"' + (NSTATE.busy ? " disabled" : "") + ">" +
      (NSTATE.busy ? "Printing…" : "Print") + "</button></span></div>" +
    "</div>";

  if ($("#nbody")) $("#nbody").scrollTop = keep;
  host.querySelectorAll("[data-note]").forEach(box => {
    box.oninput = () => { NSTATE.typed[String(box.dataset.note).toUpperCase()] = box.value; };
  });
  if ($("#nscrim")) $("#nscrim").onclick = () => closeNotesWindow(true);
  if ($("#nclose")) $("#nclose").onclick = () => closeNotesWindow(false);
  if ($("#ncancel")) $("#ncancel").onclick = () => closeNotesWindow(false);
  if ($("#nprint")) $("#nprint").onclick = () => johnPrint();
  renderFab();
  return host;
}

/** Save the notes that changed, one after another. Not three at a time: every
    write to one list already goes through `serialised("list:" + name)` in
    graph.js, so lanes would queue behind each other anyway - and running them
    one at a time is what keeps a single consent prompt single. A failure is
    remembered against its job and nothing else: the print goes ahead. */
async function savePrintNotes(ids, typed) {
  const who = whoAmI(), at = new Date().toISOString();
  const list = (ids || []).slice();
  for (let i = 0; i < list.length; i++) {
    const id = list[i], k = String(id).toUpperCase();
    try {
      await CW.listUpsert(PRINT_NOTE_LIST, id, { Note: typed[k] || "", By: who, At: at });
      PRINTNOTES[k] = { note: typed[k] || "", who: who, at: at };
      if (NSTATE) delete NSTATE.failed[k];
    } catch (e) {
      if (NSTATE) NSTATE.failed[k] = friendly(e);
    }
  }
  return NSTATE ? Object.keys(NSTATE.failed).length : 0;
}

/** Build the John print sheet from the notes as they are typed right now, hand
    it to the browser, then write the one Dashboard Log line - last, so nothing
    claims to have been exported until the bytes have actually gone.

    `ids` is exactly what the notes window showed. It is passed in rather than
    worked out again: a poll can land between opening the window and pressing
    Print, and the file must be the jobs the person was looking at. */
async function johnDownload(ids, notes) {
  const rows = exportJohnRows(ids, notes, xpCtxNow());
  const when = new Date();
  const name = exportJohnFilename(XSTATE.format, when);
  const opts = { who: xpWhoName(), when: when, company: xpCompany() };
  if (XSTATE.format === "pdf") {
    const pm = await xpLoadPdf();
    pm.createPdf(buildJohnDoc(rows, opts)).download(name);
  } else {
    const wb = buildJohnWorkbook(rows, opts);
    const buf = await wb.xlsx.writeBuffer();
    downloadBlob(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), name);
  }
  noteChange("(export)", "Export", exportLogFrom(XSTATE.format, (ids || []).length, "john"),
             "John print sheet · " + (XSTATE.f.scope === "ticked" ? "ticked jobs" : "what I see now"));
  return name;
}

/** Print: save what changed, then build the file from what is on screen. */
async function johnPrint() {
  if (!NSTATE || NBUSY) return null;
  NBUSY = true; NSTATE.busy = true; renderNotesWindow();
  const ids = NSTATE.ids.slice();
  const typed = {}, changed = [];
  ids.forEach(id => {
    const k = String(id).toUpperCase();
    const v = String(noteValue(id)).trim();
    typed[k] = v;
    /* last time's failure marks are somebody else's news: this print is
       answered by this pass, so they start clear */
    delete NSTATE.failed[k];
    if (v !== printNoteOf(id).trim()) changed.push(id);
  });
  let name = null;
  try {
    if (changed.length) {
      /* One popup, not one per note. Every listUpsert asks for the list
         permission itself, and MSAL hands out a cached token silently once
         there is one - but the FIRST time, three writes starting together
         opened three consent dialogs on top of each other. So the click asks
         once, here, before any of them. A refusal is not fatal: the writes
         below will fail one by one, each row will say "not saved", and the
         print still goes. */
      if (PRINT_NOTE_OK !== true) {
        try { if (CW && CW.listConsent) await CW.listConsent(); } catch (e) {}
      }
      await savePrintNotes(changed, typed);
    }
    name = await johnDownload(ids, typed);
    const failed = NSTATE ? Object.keys(NSTATE.failed).length : 0;
    toast(failed ? "Downloading " + name + " — " + failed + " note" + (failed === 1 ? "" : "s") + " could not be saved"
                 : "Downloading " + name);
  } catch (e) {
    name = null;
    toast("Print failed: " + ((e && e.message) || String(e)).slice(0, 160), true);
  }
  NBUSY = false;
  if (NSTATE) {
    NSTATE.busy = false;
    /* the window stays open when something did not save, so the person can see
       which row it was; a clean print takes itself away */
    if (name && !Object.keys(NSTATE.failed).length) { closeNotesWindow(); return name; }
    renderNotesWindow();
  }
  if (state.sel) renderDrawer();
  return name;
}

/* ---------- the John print sheet view ---------------------------------------
   Production (2) drawn on its own terms, in the job list's place: its rows,
   its order, its sections, its fills and its text colours, plus a tick box per
   row so the wheel and the Export window work here exactly as they do on the
   master list. Nothing in this block reads the Production job model, and
   nothing in it writes anything anywhere. */

/** The rows the view is showing: Production (2)'s own, narrowed by the
    header search box. */
function johnShown() {
  const q = String(state.q || "").trim().toLowerCase();
  return (JOHNROWS || []).filter(r => !q ||
    (r.id + " " + r.cust + " " + r.area + " " + r.notes).toLowerCase().indexOf(q) >= 0);
}
/** [[section name, rows], …] in that sheet's own order. */
function johnGroups(rows) {
  const order = [], by = {};
  (rows || []).forEach(r => {
    const k = r.section || "No section";
    if (!by[k]) { by[k] = []; order.push(k); }
    by[k].push(r);
  });
  return order.map(k => [k, by[k]]);
}
/** The sheet's own ink, pushed until it can be read on the screen it is on.
    The FILE always keeps the hex exactly as the sheet has it: this is for the
    display only, where a dark blue on a dark background is unreadable. A row
    that carries one of the sheet's fills is left alone - those fills are all
    pale, so the sheet's own ink is already right on top of them. */
function inkFor(hex, onFill) {
  const h = String(hex || "").replace("#", "");
  if (!/^[0-9A-Fa-f]{6}$/.test(h)) return "";
  if (onFill) return "#" + h.toUpperCase();
  const dark = (document.documentElement.dataset || {}).theme === "dark";
  const n = parseInt(h, 16);
  let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  const lum = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  const mix = (c, t, k) => Math.round(c + (t - c) * k);
  if (dark && lum < 0.45) { const k = Math.min(0.62, (0.45 - lum) * 1.4); r = mix(r, 255, k); g = mix(g, 255, k); b = mix(b, 255, k); }
  if (!dark && lum > 0.62) { const k = Math.min(0.55, (lum - 0.62) * 1.6); r = mix(r, 0, k); g = mix(g, 0, k); b = mix(b, 0, k); }
  return "#" + [r, g, b].map(x => x.toString(16).padStart(2, "0")).join("").toUpperCase();
}

function johnRowHtml(r) {
  const picked = !!state.picked[r.id];
  const flag = johnFlag(r);
  const fill = /^[0-9A-Fa-f]{6}$/.test(String(r.fillHex || "")) ? "#" + r.fillHex.toUpperCase() : "";
  const ink = inkFor(r.inkHex, !!fill);
  /* every one of the sheet's fills is a pale one, so dark ink on top of it is
     right in both themes - and the row must not borrow the dark theme's
     near-white text and vanish */
  const style = (fill ? "background:" + fill + ";color:#17171A" : "");
  return '<div class="jrow' + (picked ? " picked" : "") + '" data-id="' + esc(r.id) + '"' +
    (style ? ' style="' + style + '"' : "") + ">" +
    '<span class="pickcell"><input type="checkbox" class="pick"' + (picked ? " checked" : "") + "></span>" +
    '<span class="tab" style="font-weight:600' + (ink ? ";color:" + ink : "") + '">' + esc(r.id) +
      (flag ? '<span class="badge flagchip" style="color:' + ink + '">' + esc(FLAGWORD[flag]) + "</span>" : "") +
    "</span>" +
    '<span class="tab">' + esc(xpJohnDate(r.ready)) + "</span>" +
    '<span class="ell">' + esc(r.cust || "—") + "</span>" +
    '<span class="tab">' + esc(r.phone || "—") + "</span>" +
    '<span class="ell">' + esc(r.area || "—") + "</span>" +
    '<span class="tab" style="text-align:center">' + (r.wnd || "") + "</span>" +
    '<span class="tab" style="text-align:center">' + (r.drs || "") + "</span>" +
    '<span class="ell" title="' + esc(r.notes) + '">' + esc(r.notes) + "</span>" +
    "</div>";
}

function johnViewHtml() {
  if (!(JOHNROWS || []).length) {
    return '<div class="empty">The workbook has no “Production (2)” sheet, so there is nothing to print for John.</div>';
  }
  const groups = johnGroups(johnShown());
  if (!groups.length) return '<div class="empty">No row on John’s sheet matches that search.</div>';
  return '<div class="jhead">' +
      "<span></span><span class=\"kick\">Job no</span><span class=\"kick\">Ready to print</span>" +
      '<span class="kick">Customer</span><span class="kick">Phone no</span><span class="kick">Area</span>' +
      '<span class="kick" style="text-align:center">Wnd</span><span class="kick" style="text-align:center">Drs</span>' +
      '<span class="kick">Notes from Brendan’s office</span></div>' +
    groups.map((g, gi) => {
      const name = g[0], rows = g[1];
      const on = rows.filter(r => state.picked[r.id]).length;
      /* keyed by its position, not its name: a section name carries an
         ampersand ("Collect & supply only") and an attribute is no place to
         have to think about that twice */
      return '<div class="grp jgrp" data-jg="' + gi + '">' +
        '<div class="ghead"><span class="gname">' + esc(name) + "</span>" +
        '<span class="gcount">' + rows.length + "</span>" +
        '<label class="selall gsel"><input type="checkbox" class="pick gall"' +
          (on === rows.length ? " checked" : "") + '><span>Select all ' + rows.length + "</span></label>" +
        "</div><div class=\"gbody\">" + rows.map(johnRowHtml).join("") + "</div></div>";
    }).join("");
}

/** Tick boxes only: this view has no drawer, no drag and nothing to write. */
function wireJohnView(host) {
  host.querySelectorAll(".jrow[data-id]").forEach(el => {
    const cb = el.querySelector(".pick");
    if (cb) cb.onchange = () => {
      if (cb.checked) state.picked[el.dataset.id] = 1; else delete state.picked[el.dataset.id];
      renderAll();
    };
  });
  const groups = johnGroups(johnShown());
  host.querySelectorAll(".jgrp").forEach(gEl => {
    const g = groups[Number(gEl.dataset.jg)];
    const shown = g ? g[1].map(r => r.id) : [];
    const gall = gEl.querySelector(".gall");
    if (!gall) return;
    const on = shown.filter(id => state.picked[id]).length;
    gall.indeterminate = on > 0 && on < shown.length;
    gall.onclick = e => e.stopPropagation();
    gall.onchange = () => pickMany(shown, gall.checked);
  });
}

/* ---------- render ---------- */
function filtered() {
  const q = state.q.trim().toLowerCase();
  const scope = state.scope;
  let list = live().filter(j => {
    if (state.view !== "flat" && !inView(j, state.view)) return false;
    if (scope && catOf(j) !== scope) return false;
    if (state.cat === "urgent") { if (!j.urg) return false; }
    else if (state.cat === "inprod") { if (["floor", "ready", "office"].indexOf(catOf(j)) < 0) return false; }
    else if (state.cat && catOf(j) !== state.cat) return false;
    if (state.hidden[catOf(j)]) return false;
    if (state.sheet && j.sheets.indexOf(state.sheet) < 0) return false;
    if (!q) return true;
    return (j.id + " " + j.cust + " " + j.area + " " + j.eir + " " + j.off + " " +
      j.notes.map(n => n.t).join(" ")).toLowerCase().indexOf(q) >= 0;
  });
  const bi = (a, b) => a.id.localeCompare(b.id);
  const s = state.sort;
  list.sort(s === "size" ? (a, b) => tot(comp(b)) - tot(comp(a))
    : s === "wait" ? (a, b) => ((a.dates.ready || "9") < (b.dates.ready || "9") ? -1 : 1)
    : s === "county" ? (a, b) => (a.area || "~").localeCompare(b.area || "~") || bi(a, b)
    : s === "urgent" ? (a, b) => ((b.urg ? 1 : 0) - (a.urg ? 1 : 0)) || bi(a, b)
    : s === "num" ? (a, b) => (jobNum(a) - jobNum(b)) || bi(a, b)
    : s === "cat" ? (a, b) => (CATORDER.indexOf(catOf(a)) - CATORDER.indexOf(catOf(b))) || bi(a, b)
    : bi);
  if (state.desc) list.reverse();
  return list;
}

function renderTiles() {
  const all = live(), n = f => all.filter(f).length;
  const defs = [
    { k: null, l: "All jobs", v: all.length, c: "--ink", s: all.reduce((a, j) => a + j.wnd, 0) + " wnd · " + all.reduce((a, j) => a + j.drs, 0) + " drs" },
    { k: "deliver", l: "Ready to deliver", v: n(j => catOf(j) === "deliver"), c: "--done" },
    { k: "inprod", l: "In production", v: n(j => ["floor", "ready", "office"].indexOf(catOf(j)) >= 0), c: "--single" },
    { k: "collect", l: "Collect &amp; supply", v: n(j => catOf(j) === "collect"), c: "--single" },
    { k: "wonttake", l: "Won&#39;t take", v: n(j => catOf(j) === "wonttake"), c: "--notsent" },
    { k: "secondhand", l: "Second hand", v: n(j => catOf(j) === "secondhand"), c: "--ink-3" },
    { k: "urgent", l: "Urgent", v: n(j => j.urg), c: "--urgent" }
  ];
  $("#tiles").innerHTML = defs.map((d, i) =>
    '<button class="tile" aria-pressed="' + (state.cat === d.k) + '" data-k="' + (d.k || "") + '"' +
    ' style="animation-delay:' + (i * 25) + 'ms;border-top-color:var(' + d.c + ')"><span class="kick">' + d.l + '</span>' +
    '<span class="n" style="color:var(' + d.c + ')">' + d.v + '</span>' +
    (d.s ? '<span style="font-size:11px;color:var(--ink-3)">' + d.s + '</span>' : '') + '</button>').join("");
  $("#tiles").querySelectorAll(".tile").forEach(b => b.onclick = () => {
    const k = b.dataset.k || null; state.cat = state.cat === k ? null : k; renderAll();
  });
}

function renderChips() {
  const c = $("#chips"); c.innerHTML = "";
  const all = live();
  const mk = (tag, cls, txt) => { const e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; };

  /* "Select all shown": the flat list's answer to the grouped view's per-group
     tick box. It only appears when something is actually narrowing the list - a
     tile, a category, a search, a hidden category - because with every job on
     screen "all shown" and "all" are the same thing and the box says nothing.
     One pass over the already-filtered list, one render: ticking two hundred
     jobs is two hundred assignments, not two hundred re-draws. */
  if (!state.board && state.view === "flat") {
    const shown = filtered();
    if (shown.length && shown.length < all.length) {
      const on = shown.filter(j => state.picked[j.id]).length;
      const lab = mk("label", "selall");
      const box = mk("input");
      box.type = "checkbox";
      box.className = "pick";
      box.checked = on === shown.length;
      box.indeterminate = on > 0 && on < shown.length;
      box.onchange = () => pickMany(shown.map(j => j.id), box.checked);
      lab.appendChild(box);
      lab.appendChild(mk("span", null, "Select all shown (" + shown.length + ")"));
      c.appendChild(lab);
    }
  }

  c.appendChild(mk("span", "kick", "View"));
  const vsel = mk("select", "txt");
  vsel.innerHTML = '<option value="flat">Flat list</option><option value="Abin">Abin — sheet order, grouped</option>' +
    viewNames().filter(v => v !== "Abin").map(v => '<option value="' + esc(v) + '">' + esc(v) + '</option>').join("");
  vsel.value = state.view;
  vsel.onchange = () => { state.view = vsel.value; state.picked = {}; renderAll(); };
  c.appendChild(vsel);

  /* The list, or a floor station's board — or the John print sheet — in its
     place. The other workbook sheets used to be here; they are still the
     export's own filter and still the chips on every row, but as a thing to
     switch the whole page to they only ever repeated what the row chips
     already said. */
  c.appendChild(mk("span", "kick", "Show"));
  const ssel = mk("select", "txt");
  ssel.innerHTML = '<option value="">All jobs (' + all.length + ')</option>' +
    BOARDS.map(s => '<option value="' + esc(s[0]) + '"' + (state.board === s[0] ? " selected" : "") +
      '>' + esc(s[1]) + "</option>").join("");
  ssel.id = "showsel";
  ssel.onchange = () => {
    state.board = ssel.value || null; state.picked = {}; renderAll();
    /* the tick was armed at the slow rate while nobody was looking at the
       floor; picking the board is exactly the moment to speed it up, or the
       first delta lands up to a minute later - and leaving it is the moment to
       slow it down again, which is why this runs for the John print sheet too */
    stationTick();
    /* the feeder normally fills STATION_ITEMS in on every load; if it skipped
       (nothing changed, or the permission was granted since) read it now. Only
       for a floor station: the John print sheet came out of the workbook
       download the page already made and has nothing at all to read. */
    if (state.board && state.board !== "john") stationReadIfNeeded(() => renderAll());
  };
  c.appendChild(ssel);

  /* the floor's log, next to the board it belongs to: who moved which counter
     and when, filtered by person, stage, job and day. Read-only. */
  const logb = mk("button", "chip", "Floor log");
  logb.id = "logbtn";
  logb.onclick = () => openStationLog("");
  c.appendChild(logb);

  const catsBtn = mk("button", "chip", "Categories…");
  catsBtn.onclick = () => renderCatMenu(catsBtn);
  c.appendChild(catsBtn);

  const nsel = Object.keys(state.picked).length;
  if (nsel) {
    const b = mk("button", "chip", nsel + " selected — move to…");
    b.style.cssText = "background:var(--accent);color:#fff;border-color:var(--accent)";
    b.onclick = () => renderMoveMenu(b);
    c.appendChild(b);
    if (isAdmin()) {                       // only the administrator changes alerts
      const ab = mk("button", "chip", "Alert to…");
      ab.onclick = () => renderAlertMenu(ab);
      c.appendChild(ab);
    }
    const cl = mk("button", "chip", "clear");
    cl.onclick = () => { state.picked = {}; renderAll(); };
    c.appendChild(cl);
  }

  const sp = mk("span", "kick", "Sort by"); sp.style.marginLeft = "auto"; c.appendChild(sp);
  const sort = mk("select", "txt");
  sort.innerHTML = SORTS.map(p => '<option value="' + p[0] + '"' + (state.sort === p[0] ? " selected" : "") + '>' + p[1] + "</option>").join("");
  sort.onchange = () => { state.sort = sort.value; state.desc = false; renderRows(); renderChips(); };
  c.appendChild(sort);
  const dir = mk("button", "chip", state.desc ? "▼ reversed" : "▲ normal");
  dir.title = "Click to flip the order";
  dir.onclick = () => { state.desc = !state.desc; renderRows(); renderChips(); };
  c.appendChild(dir);

  renderFab();                  // the same actions again, within thumb reach
}

/** Put a menu under the thing that opened it, and pull it back on screen when
    there is no room below - the selection wheel sits in the bottom corner, and
    a menu hung under that would open past the bottom of the window. Called
    after the menu is on the page, so its real height is known. */
function menuAt(m, anchor) {
  const r = anchor.getBoundingClientRect();
  const vw = window.innerWidth || 1200, vh = window.innerHeight || 800;
  const w = m.offsetWidth || 240, h = m.offsetHeight || 0;
  m.style.left = Math.max(8, Math.min(r.left, vw - w - 8)) + "px";
  let top = r.bottom + 6;
  if (h && top + h > vh - 8) top = Math.max(8, r.top - 6 - h);
  m.style.top = top + "px";
}

/** Which categories to show. Hiding is per person and remembered. */
function renderCatMenu(anchor) {
  const old = $("#catmenu"); if (old) { old.remove(); return; }
  const cats = [["deliver", "Ready to deliver"], ["floor", "On floor"], ["ready", "Waiting for floor"],
                ["office", "In office"], ["collect", "Collect & supply"], ["wonttake", "Won't take"],
                ["secondhand", "Second hand"]];
  const m = document.createElement("div"); m.id = "catmenu"; m.className = "menu";
  m.innerHTML = '<div class="kick" style="padding:4px 10px 8px">Show which categories</div>' +
    cats.map(cc => '<label class="mrow"><input type="checkbox" data-k="' + cc[0] + '"' +
      (state.hidden[cc[0]] ? "" : " checked") + '> ' + cc[1] +
      ' <span style="color:var(--ink-4)">' + live().filter(j => catOf(j) === cc[0]).length + '</span></label>').join("") +
    '<div style="display:flex;gap:6px;padding:8px 10px 4px;border-top:1px solid var(--line);margin-top:6px">' +
    '<button class="chip" id="mall">Show all</button><button class="chip" id="mnone">Hide all</button></div>';
  document.body.appendChild(m);
  menuAt(m, anchor);
  m.querySelectorAll("input").forEach(i => i.onchange = () => {
    if (i.checked) delete state.hidden[i.dataset.k]; else state.hidden[i.dataset.k] = 1;
    saveUi(); renderTiles(); renderRows();
  });
  $("#mall").onclick = () => { state.hidden = {}; saveUi(); m.remove(); renderAll(); };
  $("#mnone").onclick = () => { cats.forEach(cc => state.hidden[cc[0]] = 1); saveUi(); m.remove(); renderAll(); };
  setTimeout(() => document.addEventListener("click", function off(e) {
    if (!m.contains(e.target) && e.target !== anchor) { m.remove(); document.removeEventListener("click", off); }
  }), 0);
}

/** Move the ticked jobs into a group of a view, or into a brand new view. */
function renderMoveMenu(anchor) {
  const old = $("#movemenu"); if (old) { old.remove(); return; }
  const jobs = Object.keys(state.picked);
  const m = document.createElement("div"); m.id = "movemenu"; m.className = "menu";
  let opts = "";
  if (state.view === "Abin" || state.view === "flat") {
    opts += '<div class="kick" style="padding:4px 10px 6px">Move ' + jobs.length + ' job' + (jobs.length > 1 ? "s" : "") + ' in the Production sheet to</div>' +
      BLOCKNAMES.map((n, i) => '<button class="mrow" data-grp="' + i + '">' + esc(n) + '</button>').join("");
  }
  opts += '<div class="kick" style="padding:10px 10px 6px;border-top:1px solid var(--line);margin-top:6px">Or a dashboard-only category (Excel unchanged)</div>' +
    viewNames().filter(v => v !== "Abin").map(v => '<button class="mrow" data-view="' + esc(v) + '">' + esc(v) + '</button>').join("") +
    '<button class="mrow" id="newcat" style="color:var(--accent);font-weight:600">+ New category from selection…</button>';
  m.innerHTML = opts;
  document.body.appendChild(m);
  menuAt(m, anchor);
  m.querySelectorAll("[data-grp]").forEach(b => b.onclick = async () => {
    m.remove(); await moveJobsInSheet(jobs, Number(b.dataset.grp));
  });
  m.querySelectorAll("[data-view]").forEach(b => b.onclick = async () => {
    m.remove(); await assignMany(jobs, b.dataset.view, "");
  });
  $("#newcat").onclick = async () => {
    m.remove();
    const name = prompt("Name for the new category:");
    if (name && name.trim()) await assignMany(jobs, name.trim(), "");
  };
  setTimeout(() => document.addEventListener("click", function off(e) {
    if (!m.contains(e.target) && e.target !== anchor) { m.remove(); document.removeEventListener("click", off); }
  }), 0);
}

async function assignMany(jobs, view, group) {
  const who = whoAmI();
  setStatus("saving " + jobs.length + " to " + view + "…", "busy");
  VIEWS[view] = VIEWS[view] || {};
  jobs.forEach((id, i) => { VIEWS[view][id] = { group: group === "" ? "" : String(group), order: i }; pendView(view, id, group === "" ? "" : String(group), i); });
  state.picked = {}; renderAll();                     // instant
  let ok = 0;
  for (let i = 0; i < jobs.length; i++) {
    try { await CW.saveAssignment(view, jobs[i], group, i, who); ok++; } catch (e) { console.warn(e); }
  }
  noteChange(jobs.join(", "), "Moved to " + view + (group === "" ? "" : " / " + (BLOCKNAMES[group] || group)), "", view);
  toast(ok + " of " + jobs.length + " saved to " + view);
  setStatus("live");
}

/** How many of a job's checkpoints are part way through, for the list badge. */
function cpInProgress(j) {
  return cpItems(j).filter(x => { const s = itemState(j, x.key); return s && s.status === "process"; }).length;
}

/** The word in a job's status badge, in the list and at the head of the drawer:
    for a job still in production, where it has got to - "Cutting", "In glazing" -
    in place of the stage word. The section badges (Ready to deliver,
    Collect/Supply, Won't take, Second hand) are left exactly as they were.

    One exception, before the job reaches the floor. Phase 0 covers both "In
    office" and "Waiting" (ready to print, not yet sent to floor), because
    nothing on the sheet has moved yet in either case - so a job sitting at
    phase 0 keeps its stage word, which does tell those two apart. Once it is on
    the floor, or anything at all has moved, the phase is the better word. */
const statusWord = j => {
  const c = catOf(j);
  if (["floor", "ready", "office"].indexOf(c) < 0) return label(j).l;
  return (c === "floor" || effectivePhase(j) >= 1) ? phaseName(j) : label(j).l;
};

/* A row fades in, staggered, when somebody has asked for a list. A repaint the
   floor's poll asked for is not that: nobody chose it, it can land as often as
   every ten seconds, and seven hundred rows fading in and out under the cursor
   would be the feature making itself hated. Those draws come in silently. */
const rowDraw = i => ROWS_QUIET ? "animation:none" : "animation-delay:" + Math.min(i * 3, 200) + "ms";

/** What the floor has recorded on a job, as one badge in its own column on the
    row - measured at 1440px, the badge cell it used to share clipped it on
    every glass row, and a clipped "Glass 48/48" reads as "Glass 48/4". Everything
    it says is already in memory (STATION_ITEMS, kept current by the poll):
    this reads, it never asks for anything, and it never writes - the counters
    are the tablet's to move and the Excel file is not involved in any of it.

    One number for all three stages, because a row has room for one number:
    eight glasses to cut, hotmelt and glaze is twenty-four steps, and "8/24"
    says how far through the job the floor has got. The breakdown, and the fact
    that none of it is in the workbook, go in the title where there is room.

    Nothing at all for a job the floor has never been fed, or one with no glass
    on it: a row that has nothing to do with the floor must look exactly as it
    did before this chip existed. */
function glassChip(j) {
  if (!j || typeof ST === "undefined") return "";
  /* the floor's own reckoning of "has glass" first, because it is a few keys
     of the job in hand rather than a look at the floor's list - and because a
     job the floor never works on is the answer "nothing" either way, which is
     what the drawer's section already says about the same job */
  if (!ST.glassTotal(j)) return "";
  const g = stationForJob(j.id);
  if (!g || !g.total) return "";
  /* the THREE glass stages, deliberately, not the four. Tuff counts a different
     quantity, so adding it would make the denominator mean nothing: "8/24" is
     eight glasses through three stages, which is what the chip has always said. */
  const done = ST.STAGE_KEYS.reduce((n, k) => n + g[ST.STAGE_ROW[k]], 0);
  const steps = g.total * ST.STAGE_KEYS.length;
  const each = ST.STAGES.map(s => s[1].toLowerCase() + " " + g[ST.STAGE_ROW[s[0]]]).join(", ");
  const why = "The floor has recorded " + done + " of " + steps + " stage steps on " +
    ST.glassWords(g.total) + ": " + each +
    ". Recorded on the Glass station page — the Excel file is not involved.";
  /* gold on the same rule and in the same colour the board and the tablet use,
     so a job that reads finished on one screen reads finished on all three */
  return '<span class="badge" title="' + esc(why) + '" style="' +
    (g.finished ? "background:var(--done-bg);color:var(--done)"
                : "background:var(--surface-2);color:var(--ink-3)") +
    '">Glass ' + done + '/' + steps + '</span>';
}

function rowHtml(j, i, max) {
  const c = comp(j), T = tot(c), w = (T / max) * 110, st = label(j), green = !!j.done;
  const fab = j.prods.some(p => (p.st || []).indexOf("process") >= 0);
  const cpn = cpInProgress(j);
  const picked = !!state.picked[j.id];
  return '<div class="row' + (state.sel === j.id ? " on" : "") + (green ? " ready" : "") + (picked ? " picked" : "") +
    '" data-id="' + j.id + '" draggable="true" style="' + rowDraw(i) + '">' +
    '<span class="pickcell"><input type="checkbox" class="pick"' + (picked ? " checked" : "") + '></span>' +
    '<span class="tab jid" style="font-weight:600;color:' + (j.urg ? "var(--urgent)" : "var(--ink)") + '">' + esc(j.id) + '</span>' +
    '<span class="ell">' + esc(j.cust || "—") + '</span>' +
    '<span class="ell" style="color:var(--ink-2)">' + esc(j.area || "—") + '</span>' +
    '<span><span class="badge" style="background:var(--surface-2);color:var(' + st.c + ')">' + esc(statusWord(j)) + '</span></span>' +
    '<span class="tab" style="color:var(--ink-2)">' + j.wnd + " / " + j.drs + '</span>' +
    '<span style="display:flex;align-items:center;gap:8px"><span class="mini" style="width:110px">' +
      '<i style="width:' + (T ? c.f / T * w : 0) + 'px;background:var(--f)"></i>' +
      '<i style="width:' + (T ? c.s / T * w : 0) + 'px;background:var(--s)"></i>' +
      '<i style="width:' + (T ? c.t / T * w : 0) + 'px;background:var(--t)"></i></span>' +
      '<span class="tab" style="font-size:12px;color:var(--ink-3)">' + T + '</span></span>' +
    /* the floor's own column, straight from their list - nothing here reads or
       writes anything, on the sheet or off it. It has a cell of its own because
       the badge cell hides what overflows it, and a clipped "Glass 48/48" reads
       as "Glass 48/4": a real-looking number that is wrong. A job the floor has
       never been fed leaves this cell empty - never a nought, never a
       placeholder. */
    '<span>' + glassChip(j) + '</span>' +
    '<span style="display:flex;gap:4px;overflow:hidden">' +
      j.sheets.slice(0, 2).map(s => '<span class="stn">' + esc(s) + '</span>').join("") +
      /* the colour-code chip lives here, with the other badges, because this is
         the cell that already handles overflow - in the job-number column it
         pushed the number itself out of sight on a narrow screen */
      flagChip(j) +
      /* and the chip already says Urgent when the row is red; this badge is for
         the jobs whose comment is the only thing saying it */
      (j.urg && j.flag !== "urgent" ? '<span class="badge" style="background:var(--urgent-bg);color:var(--urgent)">Urgent</span>' : "") +
      (fab ? '<span class="badge" style="background:var(--fab-bg);color:var(--fab)">In fab</span>' : "") +
      (cpn ? '<span class="badge cpbadge">' + cpn + ' in progress</span>' : "") +
      (function () { const n = commentsFor(j.id).length;
        return n ? '<span class="badge" style="background:var(--accent-soft);color:var(--single)">' +
          n + (n > 1 ? " comments" : " comment") + '</span>' : ""; })() +
    '</span></div>';
}

function wireRows(scope) {
  scope.querySelectorAll(".row[data-id]").forEach(el => {
    el.onclick = e => {
      if (e.target.classList.contains("pick")) return;
      state.sel = el.dataset.id; state.edit = false; renderRows(); openDrawer();
    };
    const cb = el.querySelector(".pick");
    if (cb) cb.onchange = () => {
      if (cb.checked) state.picked[el.dataset.id] = 1; else delete state.picked[el.dataset.id];
      el.classList.toggle("picked", cb.checked); renderChips();
    };
    el.ondragstart = e => {
      const ids = Object.keys(state.picked).length ? Object.keys(state.picked) : [el.dataset.id];
      e.dataTransfer.setData("text/plain", ids.join(","));
      e.dataTransfer.effectAllowed = "move";
      el.classList.add("dragging");
    };
    el.ondragend = () => el.classList.remove("dragging");
  });
}

/** "Tue 14:02" for something that happened this week, "08/09 14:02" for
    anything older. The floor's log is read at a glance, so the day of the week
    beats a date for as long as the day of the week still means anything. */
const STDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
function stWhen(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return "";
  const p = n => (n < 10 ? "0" : "") + n;
  const hm = p(d.getHours()) + ":" + p(d.getMinutes());
  const age = Date.now() - d.getTime();
  if (age >= 0 && age < 6 * 86400000) return STDAYS[d.getDay()] + " " + hm;
  return p(d.getDate()) + "/" + p(d.getMonth() + 1) + " " + hm;
}

/** What the floor has recorded for one job, in the drawer: the three bars with
    who last moved each of them and when, and under it the job's own log lines,
    newest first. Read-only, all of it: this dashboard never writes a counter
    and never writes a log line, and there is nothing here to click but the
    link that opens the same log in full. */
const STATION_TIMELINE_MAX = 12;
function stationSectionHtml(j) {
  if (typeof ST === "undefined") return "";
  /* the floor's own reckoning of "has glass": DG plus TG. A job whose only
     glass is a kind the floor never works on is not on their board, and a
     section here saying "not fed to the floor yet" would be a lie about it. */
  if (!j || !ST.glassTotal(j)) return "";
  const head = '<div class="sect"><span class="kick">Glass station</span>';
  const note = t => head + '<div class="cphint">' + esc(t) + '</div></div>';
  if (STATION_OK === null) return note(STATION_CHECKING);
  if (STATION_OK !== true) return note(STATION_WHY || STATION_LIST_MISSING);
  const g = stationForJob(j.id);
  if (!g) return note("Not fed to the floor yet.");
  /* a job that has left production is off the floor's board but its record is
     still worth reading here - and calling that "not fed yet" would be wrong */
  const why = g.active
    ? "Recorded by the floor on the Glass station page" + (g.fedAt ? ", fed " + agoWords(g.fedAt) : "")
    : "Finished on the floor \u2014 this job is no longer on their board.";
  /* why the floor cannot move it, said here rather than left as a mystery: the
     office's own glass ticks are what locked it, and un-ticking one unlocks it */
  const lock = g.officeDone
    ? '<div class="cphint">This job’s glass is ticked off here, so it is read-only on the ' +
      'tablet. Un-tick one of the glass checkpoints above to let the floor move it again.</div>'
    : "";
  return head +
    (STATION_ERR ? '<div class="cphint" style="color:var(--urgent)">' + esc(STATION_ERR) + '</div>' : "") +
    '<div class="stbars">' + stStages(g).map(s => stStageHtml(s[1], g.bars[s[0]])).join("") + '</div>' +
    lock +
    stationTimelineHtml(j.id) +
    '<div class="cphint">' + esc(why) + ' Nothing in the Excel file is involved.</div></div>';
}

/* Parsing and sorting the log is the one thing on this page that is done six
   times a minute over a list that can run to thousands of lines. The answer
   only changes when STATION_LOG is replaced - and every path that changes it
   (a read, a delta merge) replaces the array rather than editing it - so the
   array itself is the cache key. */
let LOGROWS = null, LOGROWS_OF = false;
function logRowsNow() {
  if (LOGROWS_OF === STATION_LOG) return LOGROWS;
  LOGROWS_OF = STATION_LOG;
  LOGROWS = ST.logRows(STATION_LOG || []);
  return LOGROWS;
}

/** The job's own log lines, newest first, capped - with the way to the rest. */
function stationTimelineHtml(job) {
  if (STATION_LOG_OK === null) return '<div class="cphint">' + esc(STATION_CHECKING) + '</div>';
  if (STATION_LOG_OK !== true)
    return '<div class="cphint">' + esc(STATION_LOG_WHY || STATION_LOG_MISSING) + '</div>';
  /* exact, not a substring: R530 is not R5303, and a drawer quietly listing
     another job's work would be a lie about this one */
  const rows = ST.logFilter(logRowsNow(), { job: job, exact: true });
  if (!rows.length) return '<div class="cphint">Nothing recorded on the floor for this job yet.</div>';
  return '<div class="stlog">' + rows.slice(0, STATION_TIMELINE_MAX).map(stLogRowHtml).join("") + '</div>' +
    '<button class="ghost stfull" data-stfull="' + esc(job) + '">Full log' +
    (rows.length > STATION_TIMELINE_MAX ? " (" + rows.length + " lines)" : "") + '</button>';
}
/** One line of the floor's log, in the drawer and in the log window. There is
    no glass type in it: the floor counts glasses, not kinds of glass. */
function stLogRowHtml(r) {
  return '<div class="stlrow">' +
    '<span class="stlwho">' + esc(r.who || "\u2014") + '</span>' +
    '<span class="stlwhat">' + esc(ST.stageLabel(r.stage)) + ' \u00b7 ' +
      r.from + ' \u2192 ' + r.to + '</span>' +
    '<span class="stlwhen tab">' + esc(stWhen(r.at)) + '</span></div>';
}

/** The stages worth drawing for one record. Tuff is a fourth counter against a
    quantity most jobs do not have, and a "Tuff 0 of 0" line on every card would
    be four hundred rows of nothing: it appears only on the jobs that have tuff
    on them. The three glass stages are always shown. */
const stStages = g => ST.ALL_STAGES.filter(s => s[0] !== ST.TUFF_STAGE || (g && g.tuffTotal > 0));

/** One stage of one job: "Cutting 12 of 12" and, under it, who last moved it
    and when. No bar - the floor has none either, and a card that has gone gold
    has already said the only thing a bar was saying. */
function stStageHtml(label, b) {
  const full = b.total > 0 && b.done >= b.total;
  return '<div class="stbar' + (full ? " full" : "") + '">' +
    '<div class="stbhead"><span class="stbl">' + esc(label) + '</span>' +
    '<span class="cpnum tab">' + (b.total ? b.done + " of " + b.total : "\u2014") + '</span></div>' +
    (b.by || b.at ? '<div class="stwho">' + esc(b.by || "\u2014") +
      (b.at ? ' \u00b7 ' + esc(stWhen(b.at)) : "") + '</div>' : "") +
    '</div>';
}

/** The Glass station board, read-only, in the job list's place. Everything on
    it comes from the two SharePoint lists; nothing here can write anywhere. */
function stationBoardHtml() {
  if (typeof ST === "undefined") return '<div class="empty">The station board did not load.</div>';
  if (STATION_OK === null) return '<div class="empty">' + esc(STATION_CHECKING) + '</div>';
  if (STATION_OK !== true) return '<div class="empty" style="line-height:1.6">' + esc(STATION_WHY || STATION_LIST_MISSING) + '</div>';
  const board = ST.jobBoard(STATION_ITEMS || []);
  const log = STATION_LOG_OK === true ? logRowsNow() : [];
  /* a passing failure never takes the board away: it says so above whatever
     was last read, because a stale board beats a blank one */
  const trouble = STATION_ERR ? '<div class="sttrouble">' + esc(STATION_ERR) + '</div>' : "";
  if (!board.length)
    return trouble + '<div class="empty" style="line-height:1.6">No glass jobs on the floor\u2019s board yet. ' +
           'Jobs appear here once this dashboard has fed them across.</div>';
  /* the finished ones go to the bottom, gold, exactly as they do on the floor's
     own screen: the two boards are read side by side over the phone */
  const order = board.slice().sort((a, b) => (a.finished ? 1 : 0) - (b.finished ? 1 : 0));
  return trouble + order.map(g => {
    const last = ST.logLast(log, g.job);
    return '<div class="stcard' + (g.finished ? " done" : "") + '">' +
      '<div class="sthead">' +
        '<span class="cond tab stjob">' + esc(g.job) + '</span>' +
        '<span class="stcust">' + esc(g.customer || "\u2014") + '</span>' +
        '<span class="stcount tab">' + esc(ST.glassWords(g.total)) + '</span>' +
        '<span class="stfed">' + esc(g.fedAt ? "fed " + agoWords(g.fedAt) : "not fed yet") + '</span>' +
      '</div>' +
      '<div class="stbars">' + stStages(g).map(s => stStageHtml(s[1], g.bars[s[0]])).join("") + '</div>' +
      (last ? '<div class="stlast">last: ' + esc(last.who || "\u2014") + ' ' + esc(ST.stageLabel(last.stage)) +
        ' ' + last.from + '\u2192' + last.to + ', ' + esc(agoWords(last.at)) + '</div>' : "") +
    '</div>';
  }).join("");
}

/* ---- the log window ---------------------------------------------------------
   Every line the floor has recorded, newest first, filtered by person, stage,
   job and day, with a count per person and per stage for whatever is showing.
   Read-only in the strongest sense available: there is no code path in this
   file that writes or deletes a Station log item, and no export of it either -
   what leaves this dashboard is still governed by export.js and its standing
   test, and the floor's log is not part of it.                              */
const LOG_PAGE = 200;
let LOGF = { who: "", stage: "", job: "", day: "", show: LOG_PAGE };

function openStationLog(job) {
  LOGF = { who: "", stage: "", job: job || "", day: "", show: LOG_PAGE };
  renderStationLog();
  stationTick();                 // somebody is looking at the floor now: poll fast
}

/** The window's frame: the head, the filter bar and the two boxes the repaint
    fills. Built when the window opens and never again, so a poll landing
    mid-keystroke cannot take the box out from under somebody's fingers. */
function renderStationLog() {
  let host = $("#lhost");
  if (!host) { host = document.createElement("div"); host.id = "lhost"; document.body.appendChild(host); }
  /* both reads answer once and are shared: opening this window three times
     asks SharePoint once, exactly as the drawer does */
  stationLogReadIfNeeded(() => { if ($("#lhost")) paintStationLog(); });
  stationPeopleIfNeeded(() => { if ($("#lhost")) renderStationLog(); });

  const names = {};
  ST.stationPeople(STATION_PEOPLE || [], ST.STATION_NAME).forEach(p => { names[p.name] = 1; });
  (STATION_LOG_OK === true ? logRowsNow() : []).forEach(r => { if (r.who) names[r.who] = 1; });
  const opt = (v, label, now) => '<option value="' + esc(v) + '"' + (now === v ? " selected" : "") +
    '>' + esc(label) + '</option>';

  host.innerHTML = '<div class="scrim" id="lscrim"></div><div class="logwin">' +
    '<div class="dhead"><div><div class="cond" style="font-size:25px;font-weight:700">Glass station log</div>' +
      '<div style="font-size:12.5px;color:#a8a49a;margin-top:2px"><span id="lgcount"></span> \u00b7 ' +
      'who moved which counter, and when. ' +
      'Written by the tablet only \u2014 the Excel file is not involved.</div></div>' +
      '<button class="ghost" id="lclose">Close</button></div>' +
    '<div class="lgbar">' +
      '<select class="txt" id="lgwho">' + opt("", "Everyone", LOGF.who) +
        Object.keys(names).sort().map(nm => opt(nm, nm, LOGF.who)).join("") + '</select>' +
      '<select class="txt" id="lgstage">' + opt("", "Every stage", LOGF.stage) +
        ST.ALL_STAGES.map(st => opt(st[0], st[1], LOGF.stage)).join("") + '</select>' +
      '<input class="txt" id="lgjob" placeholder="Job number" value="' + esc(LOGF.job) + '">' +
      '<input class="txt" id="lgday" type="date" value="' + esc(LOGF.day) + '">' +
      '<button class="chip" id="lgclear">Clear filters</button>' +
    '</div>' +
    '<div class="lgcounts" id="lgcounts"></div>' +
    '<div class="logbody" id="lgbody"></div>' +
    '<div class="foot"><span>Read-only \u2014 nothing here changes the floor\u2019s record or the Excel file</span>' +
    '<span>Click a job number to open it</span></div></div>';

  $("#lscrim").onclick = closeWin(host);
  $("#lclose").onclick = closeWin(host);
  /* the filters repaint the rows and the counts only, so the box being typed
     into is never rebuilt and the caret stays where it was put */
  const set = (id, key) => { const el = $(id); if (el) el.onchange = () => { LOGF[key] = el.value; LOGF.show = LOG_PAGE; paintStationLog(); }; };
  set("#lgwho", "who"); set("#lgstage", "stage"); set("#lgday", "day");
  const jb = $("#lgjob");
  if (jb) jb.oninput = () => { LOGF.job = jb.value; LOGF.show = LOG_PAGE; paintStationLog(); };
  const cl = $("#lgclear");
  if (cl) cl.onclick = () => {
    LOGF = { who: "", stage: "", job: "", day: "", show: LOG_PAGE };
    renderStationLog();                                // the bar's own values have changed
  };
  paintStationLog();
  renderFab();                 // a window is open: the wheel steps aside
}

/** The rows and the counts, and nothing else on the window. */
function paintStationLog() {
  if (!$("#lhost")) return;
  const all = STATION_LOG_OK === true ? logRowsNow() : [];
  const rows = ST.logFilter(all, LOGF);
  const counts = ST.logCounts(rows);
  const head = $("#lgcount");
  if (head) head.textContent = rows.length + " line" + (rows.length === 1 ? "" : "s");

  const chip = c => '<span class="lgcount">' + esc(c.key) + ' <strong class="tab">' + c.units + '</strong>' +
    ' <span class="lgc2">' + c.lines + ' line' + (c.lines === 1 ? "" : "s") + '</span></span>';
  const cbox = $("#lgcounts");
  if (cbox) cbox.innerHTML =
    (counts.people.length ? '<div class="lgcrow"><span class="kick">Per person</span>' +
      counts.people.map(chip).join("") + '</div>' : "") +
    (counts.stages.length ? '<div class="lgcrow"><span class="kick">Per stage</span>' +
      counts.stages.map(c => chip({ key: ST.stageLabel(c.key), units: c.units, lines: c.lines })).join("") +
      '</div>' : "");

  /* a job number is only a way into the drawer when the job is still on the
     sheet: a line about a job that has been delivered would otherwise close
     this window and open nothing */
  const jobCell = j => (byId(j)
    ? '<button class="stn jump" data-j="' + esc(j) + '" style="border:0;cursor:pointer">' + esc(j) + '</button>'
    : '<span class="stn" title="not on the sheet any more">' + esc(j) + '</span>');
  const body = STATION_LOG_OK === null ? '<div class="empty">' + esc(STATION_CHECKING) + '</div>'
    : STATION_LOG_OK !== true ? '<div class="empty" style="line-height:1.6">' +
        esc(STATION_LOG_WHY || STATION_LOG_MISSING) + '</div>'
    : !rows.length ? '<div class="empty">No line on the floor\u2019s log matches that.</div>'
    : '<div class="lglist">' + rows.slice(0, LOGF.show).map(r =>
        '<div class="lgrow">' + jobCell(r.job) +
          '<span class="ell">' + esc(ST.stageLabel(r.stage)) + '</span>' +
          '<span class="tab">' + r.from + ' \u2192 ' + r.to + '</span>' +
          '<span class="ell">' + esc(r.who || "\u2014") + '</span>' +
          '<span class="tab lgwhen">' + esc(stWhen(r.at)) + '</span>' +
        '</div>').join("") +
      (rows.length > LOGF.show
        ? '<button class="ghost lgmore" id="lgmore">Show more (' + (rows.length - LOGF.show) + ' left)</button>'
        : "") + '</div>';
  const bbox = $("#lgbody");
  if (!bbox) return;
  bbox.innerHTML = body;
  const more = $("#lgmore");
  if (more) more.onclick = () => { LOGF.show += LOG_PAGE; paintStationLog(); };
  bbox.querySelectorAll(".jump").forEach(b => b.onclick = () => {
    const host = $("#lhost");
    if (!byId(b.dataset.j)) return;                    // nothing to open: the window stays
    if (host) host.remove();
    state.sel = b.dataset.j; state.edit = false; renderRows(); openDrawer();
  });
}

function renderRows() {
  const host = $("#rows");
  /* the master list's column header belongs to the job list only; any board
     (a floor station, the John print sheet) draws its own */
  const thead = document.querySelector(".thead"); if (thead) thead.hidden = !!state.board;
  /* John's own sheet takes the list's place: tickable, but with no drawer, no
     drag and nothing that writes anything anywhere */
  if (state.board === "john") {
    ROWS_GLASS = false; ROWS_CHIPS = ""; ROWS_DRAWN = []; ROWS_STALE = false;   // no job rows on screen
    host.innerHTML = '<div class="jboard">' + johnViewHtml() + "</div>";
    wireJohnView(host);
    const shown = johnShown().length, total = (JOHNROWS || []).length;
    $("#count").textContent = !total ? "No “Production (2)” sheet in the workbook"
      : shown === total ? "Showing all " + total + " rows on John’s sheet"
      : "Showing " + shown + " of " + total + " rows on John’s sheet";
    return;
  }
  /* a floor station's board takes the whole list's place: read-only, no
     selection, no drag targets, nothing that writes anything anywhere */
  if (state.board) {
    ROWS_GLASS = false; ROWS_CHIPS = ""; ROWS_DRAWN = []; ROWS_STALE = false;   // the board is its own reason to poll fast
    /* the card's "last: ..." line comes from the log list, which the feeder
       never reads - so the board asks for it once, here */
    stationLogReadIfNeeded(() => { if (state.board) renderRows(); });
    host.innerHTML = '<div class="stboard">' + stationBoardHtml() + '</div>';
    const n = (STATION_OK === true && typeof ST !== "undefined") ? ST.jobBoard(STATION_ITEMS || []).length : 0;
    $("#count").textContent = n ? "Showing " + n + " job" + (n > 1 ? "s" : "") + " on the Glass station board"
                                : "Glass station";
    return;
  }
  const list = filtered();
  const max = Math.max(1, ...list.map(j => tot(comp(j))));
  /* what the rows about to be drawn are worth to the poll, and what they will
     be saying once drawn - one pass over the list, one map lookup a row. These
     rows are current from here on, so nothing is owed. A rate that has just
     become worth having is armed now rather than up to a minute from now,
     which is the same thing choosing the board does. */
  const wasGlass = ROWS_GLASS;
  ROWS_DRAWN = list;
  ROWS_CHIPS = chipsNow();
  ROWS_GLASS = ROWS_CHIPS !== "";
  ROWS_STALE = false;
  if (ROWS_GLASS !== wasGlass) stationTick();

  if (state.view === "flat") {
    host.innerHTML = list.length ? list.map((j, i) => rowHtml(j, i, max)).join("")
      : '<div class="empty">No job matches that search or filter.</div>';
    wireRows(host);
  } else {
    /* grouped: one collapsible section per block, each with its own search */
    const groups = {};
    list.forEach(j => { const g = groupOf(j, state.view); (groups[g] = groups[g] || []).push(j); });
    const keys = Object.keys(groups).map(Number).sort((x, y) => x - y);
    const names = BLOCKNAMES.length ? BLOCKNAMES : [];
    /* what each group's tick box means: the jobs that group is showing right
       now, after its own search box has had its say */
    const shownBy = {};
    host.innerHTML = keys.length ? keys.map(g => {
      const name = names[g] || (state.view + " group " + g);
      const open = !state.collapsed[state.view + "|" + g];
      const q = (state.gq && state.gq[state.view + "|" + g]) || "";
      let rows = groups[g];
      if (q) rows = rows.filter(j => (j.id + " " + j.cust + " " + j.area).toLowerCase().indexOf(q.toLowerCase()) >= 0);
      shownBy[g] = rows.map(j => j.id);
      const nOn = rows.filter(j => state.picked[j.id]).length;
      return '<div class="grp" data-g="' + g + '">' +
        '<div class="ghead"><button class="gtog">' + (open ? "▾" : "▸") + '</button>' +
        '<span class="gname">' + esc(name) + '</span>' +
        '<span class="gcount">' + rows.length + (q ? " of " + groups[g].length : "") + '</span>' +
        /* one tick box per section: everything this group is showing, at once */
        (rows.length ? '<label class="selall gsel"><input type="checkbox" class="pick gall"' +
          (nOn === rows.length ? " checked" : "") + '><span>Select all ' + rows.length + '</span></label>' : "") +
        '<input class="gsearch txt" placeholder="Search in this group…" value="' + esc(q) + '">' +
        '</div>' + (open ? '<div class="gbody">' +
          (rows.length ? rows.map((j, i) => rowHtml(j, i, max)).join("")
                       : '<div class="empty" style="padding:22px">Nothing here.</div>') + '</div>' : "") +
        '</div>';
    }).join("") : '<div class="empty">Nothing in this view yet.</div>';

    host.querySelectorAll(".grp").forEach(gEl => {
      const g = gEl.dataset.g, key = state.view + "|" + g;
      gEl.querySelector(".gtog").onclick = () => {
        if (state.collapsed[key]) delete state.collapsed[key]; else state.collapsed[key] = 1;
        saveUi(); renderRows();
      };
      const gall = gEl.querySelector(".gall");
      if (gall) {
        const shown = shownBy[g] || [];
        const on = shown.filter(id => state.picked[id]).length;
        /* the third state a tick box has: some of this group is picked. It can
           only be set from script, never from markup, so it is set here. */
        gall.indeterminate = on > 0 && on < shown.length;
        gall.onclick = e => e.stopPropagation();
        gall.onchange = () => pickMany(shown, gall.checked);
      }
      const s = gEl.querySelector(".gsearch");
      s.oninput = () => { state.gq = state.gq || {}; state.gq[key] = s.value;
        const p = s.selectionStart; renderRows();
        const n2 = $("#rows").querySelector('.grp[data-g="' + g + '"] .gsearch');
        if (n2) { n2.focus(); n2.setSelectionRange(p, p); } };
      s.onclick = e => e.stopPropagation();
      gEl.ondragover = e => { e.preventDefault(); gEl.classList.add("dragover"); };
      gEl.ondragleave = () => gEl.classList.remove("dragover");
      gEl.ondrop = async e => {
        e.preventDefault(); gEl.classList.remove("dragover");
        const ids = (e.dataTransfer.getData("text/plain") || "").split(",").filter(Boolean);
        if (!ids.length) return;
        if (state.view === "Abin") await moveJobsInSheet(ids, Number(g));
        else await assignMany(ids, state.view, Number(g));
      };
      wireRows(gEl);
    });
  }

  $("#count").textContent = list.length === live().length
    ? "Showing all " + live().length + " jobs on the sheet"
    : "Showing " + list.length + " of " + live().length + " jobs";
}

function renderAll() { renderTiles(); renderChips(); renderRows(); }

/* ---------- the floating selection wheel ----------
   When jobs are ticked, a round button appears in the thumb's corner showing
   how many. Tapping it glides the button inwards to the middle of the screen's
   bottom-right corner and lays its actions out on a FULL circle around it, each
   a round icon with its name underneath. Every action here already exists in
   the chip bar above the list - this is a second way to reach them on a phone,
   never a second implementation. The ring is a CSS transition on transform and
   opacity with a per-option stagger; all the JavaScript does is work out where
   the middle of the ring may sit and add or remove the "open" class.

   Nothing about the wheel pops on or off: it rises from under the screen edge
   when the first job is ticked, sinks back when the last one is unticked, and
   the host is only taken off the page once that has finished. Every duration
   and easing lives in a CSS custom property (see --fab-t-* in index.html); the
   two the JavaScript has to wait for it reads straight back off :root, so the
   motion can be retuned in the stylesheet alone. */
const FAB_R = 100;                     // the ring's radius, in px
const FAB_OW = 76, FAB_OH = 68;        // one option's footprint: 52 px button + its label
const FAB_EDGE = 8;                    // the least air an option keeps from the screen edge
const FAB_HOME = 46;                   // the closed button's centre, in from the corner (18 + 56/2)
let FABOPEN = false, FABOFF = null;    // FABOFF: the outside-click listener, while open
let FABGO = null;                      // the leave in flight: { host, off, t }
let FABLAY = null;                     // a re-lay waiting for the close to finish
let FABN = -1;                         // the count the button is showing, for the pulse

/** One timing, read off :root, so the stylesheet stays the only place a
    duration is written down. The fallback is what index.html says today: it is
    used when there is no browser to ask (the tests) or the property is gone. */
function fabMs(k, dflt) {
  try {
    const cs = typeof window !== "undefined" && window.getComputedStyle
      ? window.getComputedStyle(document.documentElement) : null;
    const v = cs && cs.getPropertyValue ? String(cs.getPropertyValue(k)).trim() : "";
    if (/ms$/.test(v)) return parseFloat(v);
    if (/s$/.test(v)) return parseFloat(v) * 1000;
  } catch (e) { /* no computed style here: the fallback stands */ }
  return dflt;
}

/** Where the middle of the open ring sits, as a distance in from the right and
    from the bottom edge. The design asks for 150 px (130 px on a phone under
    400 px wide), but that is a wish rather than a rule: an option's own body
    reaches FAB_R + half its width further out again, so on a narrow screen a
    150 px inset would hang the right-hand option off the edge. The wish is
    therefore clamped to the band in which every option stays FAB_EDGE inside
    the viewport, and if even that band is empty the ring simply centres. */
function fabGeom(n, vw, vh) {
  const W = vw || 1200, H = vh || 800;
  const want = W < 400 ? 130 : 150;
  const fit = (span, pad) => (span < pad * 2 ? Math.round(span / 2)
                                             : Math.min(Math.max(want, pad), span - pad));
  return { r: FAB_R, n: n,
           cx: fit(W, FAB_R + FAB_OW / 2 + FAB_EDGE),
           cy: fit(H, FAB_R + FAB_OH / 2 + FAB_EDGE) };
}

/* Inline SVG, so the ring carries no icon library and nothing is fetched. Each
   is drawn on a 24-box in the current colour and is never rotated, so it stays
   upright wherever on the circle its option lands. */
const FAB_ICON = {
  alert: '<path d="M18 8a6 6 0 1 0-12 0c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.7 21a2 2 0 0 1-3.4 0"/>',
  export: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/>',
  move: '<path d="M5 9l-3 3 3 3"/><path d="M9 5l3-3 3 3"/><path d="M15 19l-3 3-3-3"/><path d="M19 9l3 3-3 3"/><path d="M2 12h20"/><path d="M12 2v20"/>',
  clear: '<path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>'
};
const fabSvg = body => '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" '
  + 'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">' + body + '</svg>';

function fabActions() {
  const out = [];
  if (isAdmin()) out.push({ k: "alert", l: "Alert to…" });   // only the admin changes alerts
  out.push({ k: "export", l: "Export" }, { k: "move", l: "Move to…" }, { k: "clear", l: "Untick all" });
  return out;
}
/** A window or the drawer is open. On a phone the wheel would sit right on top
    of their Download / Save buttons, so it takes itself out of the way. */
const fabCovered = () => !!($("#dhost") || $("#xhost") || $("#ahost") || $("#chost") || $("#vhost") || $("#lhost") || $("#nhost"));
const fabClass = () => "fabwrap" + (FABOPEN ? " open" : "") + (fabCovered() ? " over" : "");
/** Every window's Close and scrim go through here: the wheel hid itself while
    the window was open, so something has to tell it the window has gone. */
const closeWin = host => () => { host.remove(); renderFab(); };
/* custom properties, so where each option lands stays a CSS decision */
function fabVar(el, k, v) {
  if (el.style && el.style.setProperty) el.style.setProperty(k, v);
  else if (el.style) el.style[k] = v;
}

/** The last job has been unticked: the wheel sinks back under the screen edge
    and is only taken off the page when it has gone. The transition tells us
    when that is; a timer stands behind it for the cases where transitionend
    never comes (the wheel hidden under a window, a browser that skips it, no
    browser at all), so a host can never be orphaned. */
function fabDrop(host) {
  if (FABGO) return;                   // already on its way out
  fabClose();
  host.className = fabClass() + " leaving";
  const done = () => {
    if (!FABGO || FABGO.host !== host) return;
    clearTimeout(FABGO.t);
    if (host.removeEventListener) host.removeEventListener("transitionend", FABGO.off);
    FABGO = null;
    /* ticked again in the meantime? then it stays, and renderFab has it */
    if (!Object.keys(state.picked).length && host.remove) host.remove();
  };
  /* Only the host's OWN sink ends the leave. Unticked with the ring open,
     every child finishes something first and bubbles it up through here - the
     button gliding home, the options drawing in, the rings shrinking, the
     cross turning back - and any one of those taken for the end of the leave
     whips the whole wheel off the screen mid-motion. The host itself only ever
     transitions transform and opacity (.fabwrap.leaving), and under reduced
     motion only opacity, so either of those from the host is the real end. */
  const ended = e => {
    if (!e || e.target !== host) return;
    if (e.propertyName && e.propertyName !== "transform" && e.propertyName !== "opacity") return;
    done();
  };
  FABGO = { host: host, off: ended, t: setTimeout(done, fabMs("--fab-t-leave", 300) + 60) };
  if (host.addEventListener) host.addEventListener("transitionend", ended);
}
/** Ticked again before it finished leaving: it never left. */
function fabStay(host) {
  if (!FABGO) return;
  clearTimeout(FABGO.t);
  if (host.removeEventListener) host.removeEventListener("transitionend", FABGO.off);
  FABGO = null;
}

function fabClose() {
  FABOPEN = false;
  if (FABOFF) { document.removeEventListener("click", FABOFF); FABOFF = null; }
  const h = $("#fabhost"); if (h) h.className = fabClass();
  const b = $("#fabbtn"); if (b && b.setAttribute) b.setAttribute("aria-expanded", "false");
}
function fabToggle() {
  if (FABOPEN) { fabClose(); return false; }
  FABOPEN = true;
  const h = $("#fabhost"); if (h) h.className = fabClass();
  const b = $("#fabbtn"); if (b && b.setAttribute) b.setAttribute("aria-expanded", "true");
  FABOFF = e => {
    const host = $("#fabhost");
    if (!host || !host.contains || !host.contains(e.target)) fabClose();
  };
  /* next tick: the click that opened the ring must not close it again */
  setTimeout(() => { if (FABOPEN && FABOFF) document.addEventListener("click", FABOFF); }, 0);
  return true;
}

/** Carry out one option, on whatever is ticked right now. Returns those ids. */
function fabDo(k, anchor) {
  const ids = Object.keys(state.picked);
  fabClose();
  if (k === "clear") { state.picked = {}; renderAll(); }
  else if (k === "move") renderMoveMenu(anchor);
  else if (k === "alert") renderAlertMenu(anchor);
  else if (k === "export") {
    if (!XSTATE) XSTATE = xpNewState();
    XSTATE.f.scope = "ticked";          // the window opens on the ticked jobs
    renderExportWindow();
  }
  return ids;
}

/** Draw (or take away) the wheel. The buttons are only rebuilt when the set of
    options or the geometry changes, so re-rendering the chip bar mid-open does
    not restart the animation under the finger. */
function renderFab() {
  const n = Object.keys(state.picked).length;
  let host = $("#fabhost");
  if (!n) { if (host) fabDrop(host); FABOPEN = false; FABN = -1; return null; }
  if (host) fabStay(host);             // ticked again mid-leave: call it back
  if (!host) { host = document.createElement("div"); host.id = "fabhost"; FABN = -1; document.body.appendChild(host); }
  const acts = fabActions();
  const g = fabGeom(acts.length, window.innerWidth, window.innerHeight);
  /* the geometry is part of the signature, so turning the phone rebuilds the ring */
  const sig = acts.map(a => a.k).join(",") + "@" + g.r + "/" + g.cx + "/" + g.cy;
  /* Never rebuild the ring under a finger. If the screen changes shape while
     it is open, it closes with its own motion first and the new geometry is
     laid out afterwards, when there is nothing on the move to interrupt. */
  if (host.dataset.sig !== sig && FABOPEN) {
    fabClose();
    if (!FABLAY) FABLAY = setTimeout(() => { FABLAY = null; renderFab(); },
      fabMs("--fab-d-home", 200) + fabMs("--fab-t-glide", 450) + 40);
  } else if (host.dataset.sig !== sig) {
    host.dataset.sig = sig;
    host.innerHTML = "";
    FABN = -1;                         // a new button: it has no number on it yet
    /* the two rings that grow out behind the button as it lands */
    ["in", "out"].forEach(w => {
      const r = document.createElement("i");
      r.className = "fabring " + w;
      host.appendChild(r);
    });
    /* how far the button itself glides: from its corner home to the ring's middle */
    const dx = FAB_HOME - g.cx, dy = FAB_HOME - g.cy;
    fabVar(host, "--fdx", Math.round(dx) + "px");
    fabVar(host, "--fdy", Math.round(dy) + "px");
    acts.forEach((a, i) => {
      /* a full circle, evenly spaced, starting at the top and going clockwise:
         four options land 90° apart, three 120° apart */
      const ang = i * (2 * Math.PI / acts.length);
      const b = document.createElement("button");
      b.className = "fabopt"; b.id = "fab-" + a.k; b.dataset.fab = a.k;
      b.innerHTML = '<span class="fabi">' + fabSvg(FAB_ICON[a.k] || "") + '</span>'
                  + '<span class="fabl">' + esc(a.l) + '</span>';
      if (b.setAttribute) { b.setAttribute("type", "button"); b.setAttribute("aria-label", a.l); }
      fabVar(b, "--fx", Math.round(dx + Math.sin(ang) * g.r) + "px");
      fabVar(b, "--fy", Math.round(dy - Math.cos(ang) * g.r) + "px");
      /* the stagger is the CSS's: --i counts out on opening, --rev back in again */
      fabVar(b, "--i", String(i));
      fabVar(b, "--rev", String(acts.length - 1 - i));
      /* the ids it acted on come back out, which is what the test asserts on;
         an array is never false, so nothing about the click is cancelled */
      b.onclick = e => { if (e && e.stopPropagation) e.stopPropagation(); return fabDo(a.k, b); };
      host.appendChild(b);
    });
    const main = document.createElement("button");
    main.className = "fab"; main.id = "fabbtn";
    if (main.setAttribute) main.setAttribute("type", "button");
    main.onclick = e => { if (e && e.stopPropagation) e.stopPropagation(); fabToggle(); };
    host.appendChild(main);
  }
  const main = $("#fabbtn");
  if (main) {
    /* The count while it is closed; the plus - which the CSS turns 315° into a
       cross - while it is open. Only rewritten when the number really changes,
       and once it has changed at least once the button carries "pulse", so the
       fresh <span class="fabn"> the CSS animation hangs off gives the new
       number one small beat instead of swapping it in dead. */
    if (FABN !== n) {
      main.className = "fab" + (FABN >= 0 ? " pulse" : "");
      main.innerHTML = '<span class="fabn">' + n + '</span><span class="fabk">ticked</span>'
        + '<span class="fabx">' + fabSvg('<path d="M12 5v14"/><path d="M5 12h14"/>') + '</span>';
      FABN = n;
    }
    if (main.setAttribute) {
      main.setAttribute("aria-expanded", FABOPEN ? "true" : "false");
      main.setAttribute("aria-label", n + " job" + (n === 1 ? "" : "s") + " ticked — actions");
    }
  }
  host.className = fabClass();
  return host;
}

/* Turning the phone changes where the ring is allowed to sit. renderFab knows
   that from its own signature; this is only what tells it to look again, once
   the resizing has settled - and if the wheel is open at the time it closes on
   its own motion first, rather than jumping to the new geometry mid-flight. */
let FABRZ = null;
if (typeof window !== "undefined" && window.addEventListener) {
  window.addEventListener("resize", () => {
    clearTimeout(FABRZ);
    FABRZ = setTimeout(() => { if ($("#fabhost")) renderFab(); }, 120);
  });
}

/* ---------- drawer: checkpoints ---------- */
const shortWho = w => String(w || "").split("@")[0];

function cpSummaryHtml(j) {
  const b = { win: [0, 0], drs: [0, 0], glass: [0, 0], prod: [0, 0] };
  cpItems(j).forEach(x => {
    const s = itemState(j, x.key); if (!s) return;
    const k = x.group.indexOf("prod:") === 0 ? "prod" : x.group;
    b[k][0] += s.done || 0; b[k][1] += s.total;
  });
  const part = (l, p) => p[1] ? l + " " + p[0] + "/" + p[1] : "";
  return [part("Windows", b.win), part("Doors", b.drs), part("Glass", b.glass),
          part("Frames/Sashes/Transoms", b.prod)].filter(Boolean).join(" · ");
}

/** One countable line: label, count, bar, - + and a number box. */
function cpLineHtml(j, it, on, withAll) {
  const s = itemState(j, it.key); if (!s) return "";
  const dis = on ? "" : " disabled";
  const w = cpStored(j.id, it.key);
  const btn = (t, act, cls) => '<button class="' + cls + '" data-cp="' + esc(it.key) + '" data-act="' + act + '"' + dis + '>' + t + '</button>';
  return '<div class="cpline" data-cpline="' + esc(it.key) + '" data-cpgroup="' + esc(it.group) + '">' +
    '<span class="cplab">' + esc(cap(it.label)) + '</span>' +
    '<span class="cpnum tab">' + cpNumText(s) + '</span>' +
    '<span class="cpbar"><i style="width:' + cpBarPct(s) + '%;background:var(' + cpBarVar(s) + ')"></i></span>' +
    '<span class="cpctl">' + btn("&minus;", "dec", "cpbtn") +
      '<input class="cpin tab" data-cpin="' + esc(it.key) + '" inputmode="numeric" value="' +
        (s.done == null ? "" : s.done) + '"' + dis + '>' +
      btn("+", "inc", "cpbtn") +
      (withAll ? btn(s.status === "done" ? "Clear" : "All done", s.status === "done" ? "none" : "all", "cpall") : "") +
    '</span>' +
    (w && w.who ? '<span class="cpwho">' + esc(shortWho(w.who)) +
      (w.when ? ", " + esc(String(w.when).slice(11, 16)) : "") + '</span>' : "") +
    '</div>';
}
const cpNumText = s => s.done == null ? "in progress" : s.done + " of " + s.total;
const cpBarPct = s => s.done == null ? 50 : (s.total ? Math.round(s.done / s.total * 100) : 0);
const cpBarVar = s => s.status === "done" ? "--done" : "--fab";
const cpBusy = (j, group) => !!CPBUSY[j.id + "|" + group];
const cpGroupDone = (j, items) => items.every(x => { const st = itemState(j, x.key); return st && st.status === "done"; });

function cpGroupHtml(j, items, name, on, showBtn, perLineAll, allText) {
  const doneAll = cpGroupDone(j, items);
  const live = on && !cpBusy(j, items[0].group);
  return '<div class="cpgrp"><div class="cphead"><span class="cpgname">' + esc(name) + '</span>' +
    (showBtn ? '<button class="cpall" data-cpgrp="' + esc(items[0].group) + '" data-alltext="' + esc(allText) +
      '" data-act="' + (doneAll ? "none" : "all") + '"' + (live ? "" : " disabled") + '>' +
      (doneAll ? "Clear" : esc(allText)) + '</button>' : "") + '</div>' +
    items.map(x => cpLineHtml(j, x, live, perLineAll)).join("") + '</div>';
}

function cpSectionHtml(j, ed) {
  const items = cpItems(j);
  if (!items.length) return "";
  /* a gold row is finished work: every checkpoint on it reads as done, so
     offering "Clear" would only invite someone to punch white holes in it */
  const locked = !!j.done;
  const on = ed && !locked;
  const g = k => items.filter(x => x.group === k);
  const prodNames = [];
  items.forEach(x => { if (x.group.indexOf("prod:") === 0 && prodNames.indexOf(x.group) < 0) prodNames.push(x.group); });
  const hint = locked
    ? "Marked ready to deliver: all checkpoints are complete. Use Undo above to put the job back into production first."
    : (ed ? "" : "Click Edit above to tick work off. Only the colour goes into the Production sheet.");
  return '<div class="sect"><div style="display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap">' +
      '<span class="kick">Checkpoints</span>' +
      '<span class="cpsum">' + esc(cpSummaryHtml(j)) + '</span></div>' +
    (hint ? '<div class="cphint">' + esc(hint) + '</div>' : "") +
    g("win").map(x => cpLineHtml(j, x, on && !cpBusy(j, "win"), !locked)).join("") +
    g("drs").map(x => cpLineHtml(j, x, on && !cpBusy(j, "drs"), !locked)).join("") +
    (g("glass").length ? cpGroupHtml(j, g("glass"), "Glass", on, !locked, !locked, "All glass done") : "") +
    prodNames.map(n => cpGroupHtml(j, g(n), cap(n.slice(5)), on, !locked, false, "All done")).join("") +
    '</div>';
}

/** Update the open Checkpoints section from the current state, in place. A tap
    that rebuilt the section would take the button out from under the finger and
    swallow the next click. */
function cpPatchSection(j) {
  const host = $("#dhost");
  if (!host || !j || !host.querySelectorAll) return;
  const sum = host.querySelector(".cpsum");
  if (sum) sum.textContent = cpSummaryHtml(j);
  const on = state.edit && !j.done;
  host.querySelectorAll("[data-cpline]").forEach(el => {
    const s = itemState(j, el.dataset.cpline);
    if (!s) return;
    const live = on && !cpBusy(j, el.dataset.cpgroup);
    if (el.querySelectorAll) el.querySelectorAll("button,input").forEach(c => { c.disabled = !live; });
    const num = el.querySelector(".cpnum"); if (num) num.textContent = cpNumText(s);
    const bar = el.querySelector(".cpbar i");
    if (bar) { bar.style.width = cpBarPct(s) + "%"; bar.style.background = "var(" + cpBarVar(s) + ")"; }
    const inp = el.querySelector(".cpin");
    if (inp && inp !== document.activeElement) inp.value = s.done == null ? "" : s.done;
    const all = el.querySelector(".cpall[data-cp]");
    if (all) { all.textContent = s.status === "done" ? "Clear" : "All done"; all.dataset.act = s.status === "done" ? "none" : "all"; }
  });
  host.querySelectorAll("[data-cpgrp]").forEach(b => {
    const items = cpItems(j).filter(x => x.group === b.dataset.cpgrp);
    if (!items.length) return;
    const doneAll = cpGroupDone(j, items);
    b.textContent = doneAll ? "Clear" : (b.dataset.alltext || "All done");
    b.dataset.act = doneAll ? "none" : "all";
    b.disabled = !(on && !cpBusy(j, b.dataset.cpgrp));
  });
}

/** Wire the checkpoint controls of the open drawer. */
function wireCheckpoints(host, id) {
  const job = byId(id);
  if (!job || job.done) return;                    // read-only while the row is gold
  host.querySelectorAll("[data-cp]").forEach(el => el.onclick = () => {
    const j = byId(id); if (!j) return;
    const item = el.dataset.cp, s = itemState(j, item); if (!s) return;
    const cur = s.done == null ? 0 : s.done;
    const act = el.dataset.act;
    setItemProgress(j, item, act === "inc" ? cur + 1 : act === "dec" ? cur - 1 : act === "all" ? s.total : 0);
  });
  host.querySelectorAll("[data-cpin]").forEach(el => el.onchange = () => {
    const j = byId(id); if (!j) return;
    const item = el.dataset.cpin, s = itemState(j, item);
    /* an empty or unreadable box is not a zero: put the number back and write nothing */
    if (cpClamp(el.value, s ? s.total : 0) == null) { if (s) el.value = s.done == null ? "" : s.done; return; }
    setItemProgress(j, item, el.value);
  });
  host.querySelectorAll("[data-cpgrp]").forEach(el => el.onclick = () => {
    const j = byId(id); if (j) setGroupDone(j, el.dataset.cpgrp, el.dataset.act === "all");
  });
}

/* ---------- drawer ---------- */
/** The pipeline: seven steps, filled up to where the job has got to. The step
    it is on is named in words beside the heading and tagged "Now", so nothing
    here depends on colour alone. Steps 4-6 are shown whether or not anything
    has reached them.

    In Edit mode every step is a button: clicking one sets the phase by hand,
    clicking the hand-set one again clears it. That decision goes to the
    "Dashboard phases" SharePoint list and nowhere near the workbook. Steps the
    sheet has already gone past cannot be chosen - the sheet's own evidence is
    always the floor. */
function phasePipeHtml(j, ed) {
  const sheet = jobPhase(j);
  const cur = effectivePhase(j);
  const hand = handPhase(j);
  const info = handPhaseInfo(j);
  const wait = !!PHASEBUSY[j.id];
  /* the list could not be read because the SharePoint permission has not
     been granted yet: the administrator's click is what grants it */
  const consent = ed && PHASE_LIST_CONSENT;
  const missing = ed && PHASE_LIST_OK === false && !consent;
  /* the first read of the list has not answered yet: the steps are there, but
     nothing can be chosen until we know there is somewhere to put the answer */
  const checking = ed && PHASE_LIST_OK !== true && !missing && !consent;
  const clickable = ed && !missing;
  const step = (p, i) => {
    const cls = "phstep" + (i < cur ? " done" : i === cur ? " now" : "") + (i === hand ? " hand" : "");
    const inner = '<div class="phbar"></div><span class="phlab">' + esc(p) + '</span>' +
      (i === cur ? '<span class="phtag">Now</span>' : i === hand ? '<span class="phtag">Set</span>' : "");
    if (!clickable) return '<div class="' + cls + '"' + (i === cur ? ' aria-current="step"' : "") + '>' + inner + '</div>';
    const below = i < sheet;
    const title = checking ? PHASE_CHECKING
      : consent ? (isAdmin() ? "click to grant the SharePoint permission, then set this phase" : PHASE_NEED_CONSENT)
      : below ? PHASE_BELOW_SHEET
      : i === hand ? "click again to clear this and go back to what the sheet shows"
      : "set this job to " + p;
    return '<button type="button" class="' + cls + '" data-ph="' + i + '"' +
      (below || wait || checking || (consent && !isAdmin()) ? " disabled" : "") + ' title="' + esc(title) + '"' +
      (i === cur ? ' aria-current="step"' : "") + '>' + inner + '</button>';
  };
  let note = "";
  if (info) {
    note = '<div class="phset">Set by <strong>' + esc(info.who || "someone") + '</strong>' +
      (info.at ? ", " + esc(stamp(info.at)) : "") +
      (info.phase !== sheet ? ' \u00b7 sheet says: <strong>' + esc(PHASES[sheet]) + '</strong>' : "") + '</div>';
  } else if (hand != null) {
    note = '<div class="phset">Just set here \u2014 saving to the phases list\u2026</div>';
  }
  if (missing) note += '<div class="phset warn">' + esc(PHASE_LIST_MISSING) + '</div>';
  else if (consent) note += '<div class="phset warn">' + esc(PHASE_NEED_CONSENT) +
    (isAdmin() ? " Click any step to grant it." : "") + '</div>';
  else if (checking) note += '<div class="phset">Checking the <strong>Dashboard phases</strong> list…</div>';
  else if (clickable) note += '<div class="phset">Click a step to set the phase by hand. Shared with everyone through the ' +
    '<strong>Dashboard phases</strong> list in SharePoint \u2014 the Excel file is not touched.</div>';
  return '<div class="sect">' +
    '<div style="display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap">' +
      '<span class="kick">Phase</span>' +
      '<span class="phnow">Now: <strong>' + esc(PHASES[cur]) + '</strong></span></div>' +
    '<div class="pipe' + (clickable ? " live" : "") + '">' + PHASES.map(step).join("") + '</div>' + note + '</div>';
}
/** The five date steps, folded away under the pipeline. Which way it is folded
    is remembered exactly like a group in the list, under the key "dates". */
function datesSectionHtml(j, st) {
  const open = !state.collapsed.dates;
  let last = "none yet";
  STEPS.forEach(p => { if (j.dates[p[0]]) last = p[1] + " " + dshort(j.dates[p[0]]); });
  return '<div class="sect"><button class="dtog" id="datetog" aria-expanded="' + (open ? "true" : "false") + '">' +
      '<span class="chev">' + (open ? "▾" : "▸") + '</span><span class="kick">Dates</span>' +
      '<span class="dsum">' + esc(last) + '</span></button>' +
    (open ? '<div class="steps">' +
      STEPS.map(p => '<div class="step"><span class="tab" style="font-size:13px;font-weight:600;color:' +
        (j.dates[p[0]] ? "var(--ink)" : "var(--ink-4)") + '">' + dshort(j.dates[p[0]]) + '</span>' +
        '<div class="stepbar" style="background:' + (j.dates[p[0]] ? "var(" + st.c + ")" : "var(--line)") + '"></div>' +
        '<span style="font-size:11px;color:var(--ink-3)">' + p[1] + '</span></div>').join("") + '</div>' : "") +
    '</div>';
}

function openDrawer() {
  if (!$("#dhost")) { const d = document.createElement("div"); d.id = "dhost"; document.body.appendChild(d); }
  renderDrawer(); renderFab();
  stationTick();                 // a job with glass on screen polls at the fast rate
}
function closeDrawer() { state.sel = null; state.edit = false; const h = $("#dhost"); if (h) h.remove(); renderRows(); renderFab(); }

function renderDrawer() {
  const host = $("#dhost"); if (!host) return;
  const j = byId(state.sel); if (!j) return;
  /* the feeder skips a run when nothing has changed, so the station list can
     easily not have been read at all by the time somebody opens a job. Read it
     once, here, or the Glass station section sits on "checking…" for ever. */
  if (Object.keys(j.glass || {}).length) {
    stationReadIfNeeded(() => { if (state.sel === j.id) renderDrawer(); });
    stationLogReadIfNeeded(() => { if (state.sel === j.id) renderDrawer(); });
  }
  const st = label(j), ed = state.edit;
  const isCS = /^[CS]\d/.test(j.id);
  const readyName = isCS ? "Collect & supply only" : "Ready to fit";
  const hint = t => '<div style="font-size:11.5px;color:var(--ink-4);margin:-6px 0 12px;line-height:1.4">' + t + '</div>';
  /* the 12 s poll re-renders the drawer; typing a count into a box must not be
     wiped out from under the fingers half way through */
  const act = document.activeElement;
  const typing = act && act.dataset && act.dataset.cpin ? { item: act.dataset.cpin, val: act.value } : null;
  const alTyping = act && act.id === "alnew" ? act.value : null;   // a half-typed address, likewise
  /* and a half-written comment: the station poll redraws this drawer every ten
     seconds, and a comment box that empties itself mid-sentence is the worst
     kind of bug to have to explain */
  const cTyping = act && act.id === "cbox" ? act.value : null;
  host.innerHTML = '<div class="scrim" id="dscrim"></div><div class="drawer">' +
    '<div class="dhead"><div><div style="display:flex;align-items:baseline;gap:9px;flex-wrap:wrap">' +
      '<span class="cond tab" style="font-size:29px;font-weight:700">' + esc(j.id) + '</span>' +
      '<span class="badge" style="background:var(--brand-2);color:#d5d1c8">' + esc(statusWord(j)) + '</span>' +
      (j.urg && j.flag !== "urgent" ? '<span class="badge" style="background:var(--urgent);color:#fff">Urgent</span>' : "") +
      flagChip(j) +
      '</div><div style="font-size:13px;color:#d5d1c8;margin-top:4px">' + esc(j.cust || "—") + ' · ' + esc(j.area || "—") + '</div></div>' +
      '<div style="display:flex;gap:7px"><button class="ghost" id="editbtn">' + (ed ? "Done" : "Edit") + '</button>' +
      '<button class="ghost" id="dclose">Close</button></div></div>' +
    '<div class="dbody">' +
      (ed ? '<div class="editbar">Changes here are written <strong>straight into the Excel sheet</strong>. Everyone sees them.</div>' +
        (MOVING[j.id]
          ? '<button class="markbtn" disabled><span class="spin"></span> moving the row in Excel…</button>'
          : j.done
          ? '<button class="markbtn undo" id="markready">Undo — put back into production</button>' +
            hint("Clears the gold and moves the row to the bottom of <b>In production</b> in Excel. Dates are not touched.")
          : '<button class="markbtn" id="markready">✓ Mark as ready to deliver</button>' +
            hint("Turns the row gold and moves it to the bottom of <b>" + esc(readyName) + "</b> in Excel. Dates are not touched.")) : "") +
      phasePipeHtml(j, ed) +
      datesSectionHtml(j, st) +
      cpSectionHtml(j, ed) +
      '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px">' +
        [[j.wnd, "windows"], [j.drs, "doors"], [tot(comp(j)), "components"], [j.sheets.length, "sheets"]]
        .map(p => '<div style="background:var(--surface-2);border:1px solid var(--line-soft);border-radius:5px;padding:9px 11px">' +
          '<div class="cond" style="font-size:23px;font-weight:700;line-height:1">' + p[0] + '</div>' +
          '<div style="font-size:11px;color:var(--ink-3)">' + p[1] + '</div></div>').join("") + '</div>' +
      '<div class="sect"><span class="kick">Details</span><dl class="dl">' +
        '<dt>Office no</dt><dd>' + esc(j.off || "—") + '</dd><dt>Eircode</dt><dd>' + esc(j.eir || "—") + '</dd>' +
        '<dt>Phone</dt><dd>' + (j.ph3 ? "•••••• " + esc(j.ph3) : "—") + '</dd>' +
        '<dt>Window colour</dt><dd>' + esc(j.colour || "—") + '</dd>' +
        '<dt>On sheets</dt><dd>' + j.sheets.map(s => '<span class="stn" style="margin-right:4px">' + esc(s) + '</span>').join("") + '</dd></dl></div>' +
      (Object.keys(j.glass).length ? '<div class="sect"><span class="kick">Glass units</span><div style="display:flex;flex-wrap:wrap;gap:6px">' +
        Object.keys(j.glass).map(k => '<span style="font-size:12.5px;padding:5px 10px;border:1px solid var(--line);border-radius:4px;background:var(--surface-2)">' +
          esc(k.toUpperCase()) + ' <strong class="tab">' + j.glass[k] + '</strong></span>').join("") + '</div></div>' : "") +
      stationSectionHtml(j) +
      (function () {
        const cs = commentsFor(j.id);
        return '<div class="sect"><span class="kick">Comments (' + cs.length + ')</span>' +
          (cs.length ? cs.map(c => '<div class="cmt"><div style="display:flex;justify-content:space-between;' +
              'font-size:11px;color:var(--ink-3);margin-bottom:4px">' +
              '<strong style="color:var(--ink-2)">' + esc(c.who) + '</strong><span>' + esc(stamp(c.at)) + '</span></div>' +
              '<div style="font-size:13px;line-height:1.45;white-space:pre-wrap">' + esc(c.to) + '</div></div>').join("")
            : '<div style="font-size:13px;color:var(--ink-4)">No comments on this job yet.</div>') +
          (ed ? '<textarea id="cbox" rows="2" placeholder="Add a comment for this job…"></textarea>' +
                '<div style="display:flex;justify-content:space-between;align-items:center;gap:10px">' +
                '<span style="font-size:11.5px;color:var(--ink-4)">Saved to the Dashboard Log sheet and shared with everyone. The Production sheet is not changed.</span>' +
                '<button class="btn" id="cadd">Add comment</button></div>'
              : '<div style="font-size:12px;color:var(--ink-4)">Click <strong>Edit</strong> above to add a comment.</div>') +
          '</div>';
      })() +
      /* the print note, read only here: it is written in the print-notes
         window before a John print, and lives in a SharePoint list */
      (function () {
        const info = printNoteInfo(j.id);
        const note = info ? String(info.note || "").trim() : "";
        const body = PRINT_NOTE_OK === null
          ? '<div style="font-size:13px;color:var(--ink-4)">Print notes are read when the print-notes ' +
            'window opens, so there is nothing to show here yet.</div>'
          : note
          ? '<div class="cmt"><div style="display:flex;justify-content:space-between;font-size:11px;' +
              'color:var(--ink-3);margin-bottom:4px"><strong style="color:var(--ink-2)">' +
              esc(shortWho(info.who)) + '</strong><span>' + esc(stamp(info.at)) + '</span></div>' +
              '<div style="font-size:13px;line-height:1.45;white-space:pre-wrap">' + esc(note) + '</div></div>'
          : '<div style="font-size:13px;color:var(--ink-4)">No print note for this job.</div>';
        return '<div class="sect"><span class="kick">Print notes (John print sheet)</span>' + body +
          '<div style="font-size:11.5px;color:var(--ink-4);margin-top:6px">Written in the print-notes ' +
          'window before a John print. Kept in SharePoint, never in the Excel file.</div></div>';
      })() +
      alertsSectionHtml(j) +
      (j.notes.length ? '<div class="sect"><span class="kick">From the sheet</span>' +
        j.notes.map(n => '<div class="note"><div class="kick" style="margin-bottom:3px">' + esc(n.k) + ' · ' + esc(n.s) + '</div>' +
          '<div style="font-size:13px;line-height:1.45">' + esc(n.t) + '</div></div>').join("") + '</div>' : "") +
    '</div></div>';

  if (typing) {
    const box = host.querySelector('.cpin[data-cpin="' + typing.item + '"]');
    if (box) { box.value = typing.val; try { box.focus(); } catch (e) {} }
  }
  if (cTyping != null) {
    const box = host.querySelector("#cbox");
    if (box) { box.value = cTyping; try { box.focus(); } catch (e) {} }
  }
  if (alTyping) {
    const box = host.querySelector("#alnew");
    if (box) { box.value = alTyping; try { box.focus(); } catch (e) {} }
  }
  $("#dscrim").onclick = closeDrawer;
  $("#dclose").onclick = closeDrawer;
  $("#editbtn").onclick = () => { state.edit = !state.edit; renderDrawer(); };
  /* the one thing in the Glass station section that can be clicked: the same
     log, unfiltered from this job rather than cut off at twelve lines */
  const full = host.querySelector("[data-stfull]");
  if (full) full.onclick = () => openStationLog(full.dataset.stfull);
  const dtog = $("#datetog");
  if (dtog) dtog.onclick = () => { state.collapsed.dates = state.collapsed.dates ? 0 : 1; saveUi(); renderDrawer(); };
  /* the Alerts section is not part of Edit mode: it never touches the
     Production sheet, and only the administrator sees its controls at all */
  wireAlerts(host, j.id);

  if (ed) {
    /* the phase steps: one click sets the phase by hand, clicking the one it is
       already set to clears it. Nothing here goes near the workbook. */
    (host.querySelectorAll("[data-ph]") || []).forEach(el => {
      el.onclick = () => { const jj = byId(state.sel); if (jj) setPhaseByHand(jj, Number(el.dataset.ph)); };
    });
    const mr = $("#markready");
    if (mr) mr.onclick = async () => {
      const turningOn = !j.done;
      mr.disabled = true;
      mr.innerHTML = '<span class="spin"></span> writing to Excel…';
      const tw = performance.now();
      try {
        const row = await markReady(j, turningOn);
        console.log("[dashboard] write took " + Math.round(performance.now() - tw) + "ms");
        pend(j.id, { done: turningOn ? 1 : 0 });
        ALL = applyPending(ALL);
        noteChange(j.id, "Ready to deliver", turningOn ? "no" : "yes", turningOn ? "yes" : "no");
        toast(j.id + (turningOn ? " marked ready to deliver — row " + row + " is now gold in Excel"
                                : " put back into production"));
        renderAll(); renderDrawer();          // instant: don't wait ~35s for the file
        /* then the row goes to its section: C/S jobs to Collect & supply, others to
           Ready to fit; undo sends it to the bottom of In production */
        const target = sectionIdx(turningOn ? readyName : "In production");
        if (target >= 0 && byId(j.id) && byId(j.id).blk !== target) await moveJobsInSheet([j.id], target);
        else scheduleReconcile();                                // reconcile once the file catches up
      } catch (e) {
        toast(friendly(e), true);
        mr.disabled = false; renderDrawer();
      }
    };
    const cadd = $("#cadd");
    if (cadd) cadd.onclick = async () => {
      const box = $("#cbox"), t = box.value.trim();
      if (!t) return;
      cadd.disabled = true;
      await addComment(j.id, t);
      box.value = ""; cadd.disabled = false;
      renderDrawer(); renderRows();
    };
    wireCheckpoints(host, j.id);
  }
}

/* ---------- changes window ---------- */
let cf = { q: "", who: "", src: "" };
function stamp(v) {
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v)) {   // ISO from the log sheet
    return v.slice(8, 10) + "/" + v.slice(5, 7) + "  " + v.slice(11, 16);
  }
  if (typeof v === "string" && v.indexOf("/") > 0) return v;
  const d = new Date(v), p = n => (n < 10 ? "0" : "") + n;
  return isNaN(d) ? String(v) : p(d.getDate()) + "/" + p(d.getMonth() + 1) + "  " + p(d.getHours()) + ":" + p(d.getMinutes());
}
function ago(v) {
  const d = new Date(v);
  if (isNaN(d)) return "";
  const m = Math.round((Date.now() - d) / 60000);
  return m < 1 ? "just now" : m < 60 ? m + "m ago" : m < 1440 ? Math.round(m / 60) + "h ago" : Math.round(m / 1440) + "d ago";
}
function renderChanges() {
  let host = $("#chost");
  if (!host) { host = document.createElement("div"); host.id = "chost"; document.body.appendChild(host); }
  const people = [...new Set(CHANGES.map(c => c.who))].sort();
  const q = cf.q.trim().toLowerCase();
  const rows = CHANGES.filter(c =>
    (!cf.who || c.who === cf.who) && (!cf.src || c.src === cf.src) &&
    (!q || (c.job + " " + c.what + " " + c.who + " " + c.from + " " + c.to).toLowerCase().indexOf(q) >= 0));
  const opt = (v, l, cur) => '<option value="' + esc(v) + '"' + (cur === v ? " selected" : "") + '>' + esc(l) + '</option>';
  host.innerHTML = '<div class="scrim" id="cscrim"></div><div class="logwin">' +
    '<div class="dhead"><div><div class="cond" style="font-size:25px;font-weight:700">Changes</div>' +
      '<div style="font-size:12.5px;color:#a8a49a;margin-top:2px">Everything that has changed since this page was opened, newest first</div></div>' +
      '<div style="display:flex;gap:7px"><button class="ghost" id="cclear">Clear</button><button class="ghost" id="cclose">Close</button></div></div>' +
    '<div class="logfilters">' +
      '<input class="txt" id="cq" placeholder="Search job, person or field…" value="' + esc(cf.q) + '" style="flex:1;min-width:180px">' +
      '<select class="txt" id="cwho">' + opt("", "Everyone", cf.who) + people.map(p => opt(p, p, cf.who)).join("") + '</select>' +
      '<select class="txt" id="csrc">' + opt("", "All sources", cf.src) + opt("sheet", "Changed in Excel", cf.src) + opt("dashboard", "Changed here", cf.src) + '</select>' +
    '</div>' +
    '<div class="loghead"><span class="kick">When</span><span class="kick">Who</span><span class="kick">Job</span>' +
      '<span class="kick">What changed</span><span class="kick">From &rarr; to</span></div>' +
    '<div class="logbody">' +
    (rows.length ? rows.map(c =>
      '<div class="logrow"><span class="tab" style="font-size:12px;color:var(--ink-3)">' + stamp(c.at) +
        '<div style="font-size:10.5px;color:var(--ink-4)">' + ago(c.at) + '</div></span>' +
      '<span style="font-weight:600;font-size:12.5px">' + esc(c.who) + '</span>' +
      '<span><button class="stn jump" data-j="' + esc(c.job) + '" style="border:0;cursor:pointer">' + esc(c.job) + '</button></span>' +
      '<span><span class="badge" style="background:var(--surface-2);color:var(' + (c.src === "dashboard" ? "--accent" : "--single") + ');margin-right:6px">' +
        (c.shared ? "Logged" : c.src === "dashboard" ? "Here" : "Excel") + '</span>' + esc(c.what) + '</span>' +
      '<span style="font-size:12px">' + (c.from || c.to ?
        '<span style="color:var(--ink-4);text-decoration:line-through">' + esc(c.from || "blank") + '</span> &rarr; ' +
        '<span style="font-weight:600">' + esc(c.to || "blank") + '</span>' : "&mdash;") + '</span></div>').join("")
      : '<div class="empty">' + (CHANGES.length ? "Nothing matches those filters." :
          "Nothing has changed yet. Edit something in Excel or here, and it will be listed with what it was and what it became.") + '</div>') +
    '</div><div class="foot"><span>' + rows.length + ' of ' + CHANGES.length + '</span><span>Click a job number to open it</span></div></div>';
  $("#cscrim").onclick = closeWin(host);
  $("#cclose").onclick = closeWin(host);
  $("#cclear").onclick = () => {
    if (confirm("Hide the entries recorded in this browser? The permanent log in the Dashboard Log sheet is NOT touched and will reappear on the next refresh.")) {
      CHANGES = CHANGES.filter(c => c.shared); saveChanges(); updateChangeBtn(); renderChanges();
    }
  };
  const cq = $("#cq");
  cq.oninput = () => { cf.q = cq.value; const p = cq.selectionStart; renderChanges(); const n = $("#cq"); n.focus(); n.setSelectionRange(p, p); };
  $("#cwho").onchange = e => { cf.who = e.target.value; renderChanges(); };
  $("#csrc").onchange = e => { cf.src = e.target.value; renderChanges(); };
  host.querySelectorAll(".jump").forEach(b => b.onclick = () => { host.remove(); state.sel = b.dataset.j; state.edit = false; renderRows(); openDrawer(); });
  renderFab();                 // a window is open: the wheel steps aside
}


/* ---------- version history window ----------
   Lists SharePoint's own versions of the workbook, shows what each one changed
   (by parsing it with the same parser and diffing), and can roll back to one. */
let VERSIONS = [], VCACHE = {}, vsel = null, vmode = "since";
const vnum = v => String(v.id).replace(/\.0$/, "");
const vwho = v => (v.by || "").replace(/ ?[-\u2013] ?Costello.*$/i, "").replace(/ Costello Windows$/i, "");

async function versionJobs(v) {
  if (!VCACHE[v.id]) VCACHE[v.id] = parseWorkbook(await CW.downloadVersion(v.id));
  return VCACHE[v.id];
}

async function renderVersions() {
  let host = $("#vhost");
  if (!host) { host = document.createElement("div"); host.id = "vhost"; document.body.appendChild(host); }
  if (!VERSIONS.length) {
    host.innerHTML = '<div class="scrim" id="vscrim"></div><div class="logwin vwin"><div class="dhead"><div class="cond" style="font-size:25px;font-weight:700">Versions</div></div>' +
      '<div class="empty"><span class="spin dark"></span> Reading SharePoint version history\u2026</div></div>';
    $("#vscrim").onclick = closeWin(host);
    renderFab();
    try { VERSIONS = await CW.listVersions(80); }
    catch (e) {
      host.innerHTML = '<div class="scrim" id="vscrim"></div><div class="logwin vwin"><div class="empty">' + esc(friendly(e)) + '</div></div>';
      $("#vscrim").onclick = closeWin(host); renderFab(); return;
    }
  }
  host.innerHTML = '<div class="scrim" id="vscrim"></div><div class="logwin vwin">' +
    '<div class="dhead"><div><div class="cond" style="font-size:25px;font-weight:700">Versions</div>' +
      '<div style="font-size:12.5px;color:#a8a49a;margin-top:2px">SharePoint keeps every save. Pick one to see what changed, or roll back to it.</div></div>' +
      '<button class="ghost" id="vclose">Close</button></div>' +
    '<div class="vbody"><div class="vlist">' + VERSIONS.map((v, i) => {
        const d = new Date(v.at), p = n => (n < 10 ? "0" : "") + n;
        return '<button class="vrow' + (vsel === v.id ? " on" : "") + '" data-v="' + v.id + '">' +
          '<span class="tab" style="font-weight:600">v' + vnum(v) + '</span>' +
          '<span class="tab" style="color:var(--ink-3);font-size:12px">' + p(d.getDate()) + "/" + p(d.getMonth() + 1) + " " + p(d.getHours()) + ":" + p(d.getMinutes()) + '</span>' +
          '<span class="ell" style="font-size:12.5px">' + esc(vwho(v)) + '</span>' +
          (i === 0 ? '<span class="badge" style="background:var(--green-bg);color:var(--green)">current</span>' : "") + '</button>';
      }).join("") + '</div>' +
      '<div class="vdetail" id="vdetail">' + (vsel ? "" :
        '<div class="empty">Select a version on the left.<br><span style="font-size:12px">' + VERSIONS.length + ' most recent shown, newest first.</span></div>') +
      '</div></div></div>';
  $("#vscrim").onclick = closeWin(host);
  $("#vclose").onclick = closeWin(host);
  host.querySelectorAll(".vrow").forEach(b => b.onclick = () => { vsel = b.dataset.v; vmode = "since"; renderVersions(); });
  renderFab();                 // a window is open: the wheel steps aside
  if (vsel) renderVersionDetail();
}

async function renderVersionDetail() {
  const box = $("#vdetail"); if (!box) return;
  const v = VERSIONS.find(x => x.id === vsel); if (!v) return;
  const idx = VERSIONS.indexOf(v), isCurrent = idx === 0, prev = VERSIONS[idx + 1];
  const d = new Date(v.at);
  const head = '<div style="display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;margin-bottom:10px">' +
    '<span class="cond" style="font-size:24px;font-weight:700">v' + vnum(v) + '</span>' +
    '<span style="color:var(--ink-3)">' + d.toLocaleString([], { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) +
    ' \u00b7 ' + esc(vwho(v)) + ' \u00b7 ' + Math.round(v.size / 1024) + ' KB</span></div>';
  box.innerHTML = head + '<div class="empty" style="padding:30px"><span class="spin dark"></span> Downloading and comparing\u2026</div>';
  try {
    let entries, title, note;
    if (vmode === "since") {
      const vj = await versionJobs(v);
      entries = isCurrent ? [] : diffJobs(vj, ALL, "since", v.at);
      title = isCurrent ? "This is the current file" : "Changed since this version";
      note = isCurrent ? "" : "Rolling back to v" + vnum(v) + " would undo everything below.";
    } else if (!prev) {
      entries = []; title = "Oldest version loaded"; note = "Nothing older is loaded to compare with.";
    } else {
      const both = await Promise.all([versionJobs(v), versionJobs(prev)]);
      entries = diffJobs(both[1], both[0], vwho(v), v.at);
      title = "What this save changed"; note = "Compared with v" + vnum(prev) + ", the save immediately before it.";
    }
    box.innerHTML = head +
      '<div style="display:flex;gap:6px;margin-bottom:12px">' +
        '<button class="chip" aria-pressed="' + (vmode === "since") + '" id="vm1">Changed since</button>' +
        '<button class="chip" aria-pressed="' + (vmode === "save") + '" id="vm2">What this save changed</button></div>' +
      '<div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px;gap:12px">' +
        '<span class="kick">' + title + (entries.length ? " \u00b7 " + entries.length : "") + '</span>' +
        '<span style="font-size:12px;color:var(--ink-4);text-align:right">' + esc(note) + '</span></div>' +
      '<div class="vdiff">' + (entries.length ? entries.map(c =>
        '<div class="vline"><span class="stn">' + esc(c.job) + '</span><span>' + esc(c.what) + '</span>' +
        '<span style="font-size:12px"><span style="color:var(--ink-4);text-decoration:line-through">' + esc(c.from || "blank") +
        '</span> \u2192 <b>' + esc(c.to || "blank") + '</b></span></div>').join("")
        : '<div style="padding:18px;color:var(--ink-4);font-size:13px">' +
          (isCurrent ? "Nothing to compare \u2014 this is what you are looking at now." : "No differences in the job data.") + '</div>') + '</div>' +
      (isCurrent ? "" :
        '<div class="rollbox"><div><b>Roll back to this version</b>' +
        '<div style="font-size:12px;color:var(--ink-3);margin-top:3px;line-height:1.45">Makes v' + vnum(v) +
        ' the current file. Every sheet goes back to how it was at ' + d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) +
        ' \u2014 including Dashboard Log and Dashboard Views entries made since. The state you replace is kept as a new version, so this can itself be undone.</div></div>' +
        '<button class="btn danger" id="vroll">Roll back\u2026</button></div>');
    $("#vm1").onclick = () => { vmode = "since"; renderVersionDetail(); };
    $("#vm2").onclick = () => { vmode = "save"; renderVersionDetail(); };
    const rb = $("#vroll"); if (rb) rb.onclick = () => doRollback(v, entries.length);
  } catch (e) {
    box.innerHTML = head + '<div class="empty">' + esc(friendly(e)) + '</div>';
  }
}

async function doRollback(v, nChanges) {
  const vn = vnum(v);
  const typed = prompt("This replaces the live workbook with version " + vn + " and undoes " + nChanges +
    " change" + (nChanges === 1 ? "" : "s") + " made since.\n\nThe current state is kept as a version, so this can be undone.\n\nType " + vn + " to confirm:");
  if (typed === null) return;
  if (typed.trim() !== vn) { toast("Not rolled back \u2014 the number did not match.", true); return; }
  const cur = VERSIONS[0] ? vnum(VERSIONS[0]) : "?";
  setStatus("rolling back to v" + vn + "\u2026", "busy");
  try {
    await CW.restoreVersion(v.id);
    PENDING = {}; savePending();                 // held edits no longer describe the file
    VERSIONS = []; VCACHE = {}; vsel = null;
    noteChange("(workbook)", "Rolled back", "v" + cur, "v" + vn);
    const h = $("#vhost"); if (h) h.remove();
    toast("Rolled back to v" + vn + ". SharePoint is applying it \u2014 the view refreshes in about 40 seconds.");
    lastStamp = null;
    setTimeout(() => load("reloading after rollback\u2026", true), 40000);
  } catch (e) {
    setStatus("rollback failed", "err");
    toast(friendly(e), true);
  }
}

/* ---------- keep the page itself up to date ----------
   GitHub Pages caches index.html for around ten minutes, so a browser can sit
   on an old build after a deploy. Ask the server directly and offer a reload. */
async function checkBuild() {
  try {
    const el = $("#build");
    const running = el ? el.textContent.replace("build ", "").trim() : "";
    const r = await fetch("version.json?t=" + Date.now(), { cache: "no-store" });
    if (!r.ok) return;
    const latest = (await r.json()).build;
    if (running && latest && latest !== running && !$("#newver")) {
      const bar = document.createElement("div");
      bar.id = "newver"; bar.className = "toast";
      bar.style.cursor = "pointer";
      bar.innerHTML = "A newer version of the dashboard is available &nbsp;<b>Reload</b>";
      bar.onclick = () => location.reload(true);
      document.body.appendChild(bar);
    }
  } catch (e) { /* offline or blocked - not important */ }
}

/* ---------- theme ---------- */
function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  try { localStorage.setItem("cw_theme", t); } catch (e) {}
  $("#themebtn").textContent = t === "dark" ? "Light" : "Dark";
}

/* ---------- boot ---------- */
async function start() {
  /* belt and braces: set both the attribute and the style, so no stylesheet
     specificity accident can leave the overlay covering the dashboard */
  const gate = $("#gate");
  gate.hidden = true; gate.style.display = "none";
  $("#topbar").hidden = false; $("#topbar").style.display = "flex";
  $("#main").hidden = false; $("#main").style.display = "block";
  const a = CW.account;
  $("#whobtn").textContent = (a && (a.username || a.name)) || "signed in";
  $("#whobtn").onclick = () => {
    if (!confirm("Sign out of the dashboard?")) return;
    cpClearQueue();                  // owed taps belong to the person who made them
    CW.signOut();
  };
  /* findFile() is the first thing that touches the workbook, and it runs
     inside openSession() - before load() and its own catch. A station account
     signing in here would otherwise sit on "starting…" for ever, so the whole
     opening sequence answers to the same no-access gate. */
  NOACCESS = false;
  try {
    await CW.openSession();
    await load("first read…");
  } catch (e) {
    const m = (e && e.message) || String(e);
    if (NOACCESS_RE.test(m)) { showNoAccessGate(); return; }
    setStatus("read failed", "err");
    toast("Could not open the workbook: " + m.slice(0, 160), true);
    return;
  }
  if (NOACCESS) return;            // load() met it and has already put the gate up
  cpReplayQueue();                 // taps this browser owed from a previous visit
  bootReconcile();                 // a reload must not orphan a pending write
  setInterval(poll, 12000);
  setStationFoot(); setInterval(setStationFoot, 60000);   // "station feed: 3 min ago" keeps counting
  stationTick();                   // the floor's own two lists, kept current by delta
  checkBuild(); setInterval(checkBuild, 120000);
}

(async function boot() {
  let t = "light";
  try { t = localStorage.getItem("cw_theme") || (matchMedia("(prefers-color-scheme:dark)").matches ? "dark" : "light"); } catch (e) {}
  applyTheme(t);
  $("#themebtn").onclick = () => applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  $("#refreshbtn").onclick = () => load("refreshing…");
  $("#changebtn").onclick = () => renderChanges();
  $("#versbtn").onclick = () => renderVersions();
  $("#alertbtn").onclick = () => renderAlertsWindow();
  $("#exportbtn").onclick = () => renderExportWindow();
  /* the logo is looked for once, here, so the export path itself never has to
     go and get anything - and a missing file is simply a PDF without a logo */
  xpLoadLogo();
  updateChangeBtn();
  cpWatchExit();
  $("#q").addEventListener("input", e => { state.q = e.target.value; renderRows(); });
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && FABOPEN) { fabClose(); return; }
    /* the notes window sits on top of the Export window, so it takes Escape
       first: one press steps back to the Export window, the next closes that */
    if (e.key === "Escape" && $("#nhost")) { closeNotesWindow(); return; }
    if (e.key === "Escape" && $("#xhost")) { $("#xhost").remove(); renderFab(); return; }
    if (e.key === "Escape" && $("#lhost")) { $("#lhost").remove(); renderFab(); return; }
    if (e.key === "Escape" && $("#dhost")) closeDrawer();
    if (e.key === "/" && document.activeElement !== $("#q")) { e.preventDefault(); $("#q").focus(); }
  });

  $("#signinbtn").onclick = async () => {
    try { await CW.signIn(); await start(); }
    catch (e) { $("#gateerr").style.display = "block"; $("#gateerr").textContent = "Sign-in failed:\n" + e.message; }
  };

  try {
    const acct = await CW.initAuth();
    if (acct) await start();
  } catch (e) {
    $("#gateerr").style.display = "block";
    $("#gateerr").textContent = "Startup problem:\n" + e.message;
  }
})();
