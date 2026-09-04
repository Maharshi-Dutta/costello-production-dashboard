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
  p.at = now;
  if ("done" in patch) { p.done = patch.done; t.done = now; }
  if (patch.prod) { p.prods[patch.prod.name] = patch.prod.status; t["prod:" + patch.prod.name] = now; }
  if (patch.cp) {
    p.cp = p.cp || {};
    /* null means "stop holding this one" - used when a write failed and the
       count it replaced was itself unknown, so there is nothing to put back */
    for (const k in patch.cp) {
      if (patch.cp[k] == null) { delete p.cp[k]; delete t["cp:" + k]; }
      else { p.cp[k] = patch.cp[k]; t["cp:" + k] = now; }
    }
  }
  if ("blk" in patch) { if (patch.blk == null) { delete p.blk; delete t.blk; } else { p.blk = patch.blk; t.blk = now; } }
  savePending();
}
const blkCat = b => b === 0 ? "secondhand" : b === 1 ? "wonttake" : b === 2 ? "collect" : "active";
const pendEmpty = p => !("done" in p) && p.blk == null &&
  !Object.keys(p.prods || {}).length && !Object.keys(p.cp || {}).length;

/** `fresh` = this is a newly parsed workbook, so a held count can be compared
    with what the file now says and let go once the two agree. */
function applyPending(list, fresh) {
  const now = Date.now();
  let dropped = false;
  Object.keys(PENDING).forEach(id => {
    const p = PENDING[id], t = p.t = p.t || {};
    const old = k => now - (t[k] || p.at || 0) > PENDING_MS;
    if ("done" in p && old("done")) { delete p.done; delete t.done; dropped = true; }
    if (p.blk != null && old("blk")) { delete p.blk; delete t.blk; dropped = true; }
    Object.keys(p.prods || {}).forEach(k => { if (old("prod:" + k)) { delete p.prods[k]; delete t["prod:" + k]; dropped = true; } });
    Object.keys(p.cp || {}).forEach(k => { if (old("cp:" + k)) { delete p.cp[k]; delete t["cp:" + k]; dropped = true; } });
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
    if (fresh && p.cp) Object.keys(p.cp).forEach(k => {
      const st = itemState(j, k);
      if (st && st.done != null && st.done === p.cp[k]) { delete p.cp[k]; delete p.t["cp:" + k]; dropped = true; }
    });
    if (pendEmpty(p)) { delete PENDING[j.id]; dropped = true; return j; }
    const c = Object.assign({}, j);
    c.raw = j;
    if ("done" in p) c.done = p.done;
    if (p.cp && Object.keys(p.cp).length) { c.cp = cpWithHeld(j, p.cp); c.cpDone = Object.assign({}, p.cp); }
    if (p.blk != null) { c.blk = p.blk; c.cat = blkCat(p.blk); }   // moved in Excel; file still catching up
    c.prods = j.prods.map(y => Object.prototype.hasOwnProperty.call(p.prods, y.n)
      ? Object.assign({}, y, { st: p.prods[y.n] ? [p.prods[y.n]] : [] }) : y);
    c.stage = c.done ? "deliver" : (c.dates.floor ? "floor" : (c.dates.ready ? "ready" : "office"));
    return c;
  });
  if (dropped) savePending();
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
const saveChanges = () => { try { localStorage.setItem("cw_changes", JSON.stringify(CHANGES.slice(0, 400))); } catch (e) {} };
let state = { q: "", cat: null, sheet: null, sort: "id", desc: false, sel: null, edit: false,
              scope: "", view: "flat", picked: {}, collapsed: {}, hidden: {} };
let VIEWS = {};            // view name -> { job -> {group, order} }  (from the workbook)
let BLOCKNAMES = [];
try { state.hidden = JSON.parse(localStorage.getItem("cw_hidden") || "{}"); } catch (e) {}
try { state.collapsed = JSON.parse(localStorage.getItem("cw_collapsed") || "{}"); } catch (e) {}
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
  } catch (e) {
    setStatus("read failed", "err");
    toast("Could not read the workbook: " + e.message, true);
  }
  busy = false;
}

/** Poll cheaply: only re-download when SharePoint says the file actually changed. */
async function poll() {
  if (busy || !CW.account) return;
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

  c.appendChild(mk("span", "kick", "View"));
  const vsel = mk("select", "txt");
  vsel.innerHTML = '<option value="flat">Flat list</option><option value="Abin">Abin — sheet order, grouped</option>' +
    viewNames().filter(v => v !== "Abin").map(v => '<option value="' + esc(v) + '">' + esc(v) + '</option>').join("");
  vsel.value = state.view;
  vsel.onchange = () => { state.view = vsel.value; state.picked = {}; renderAll(); };
  c.appendChild(vsel);

  c.appendChild(mk("span", "kick", "Sheet"));
  const ssel = mk("select", "txt");
  ssel.innerHTML = '<option value="">All sheets (' + all.length + ')</option>' +
    SHEETNAMES.map(s => { const n = all.filter(j => j.sheets.indexOf(s) >= 0).length;
      return n ? '<option value="' + s + '"' + (state.sheet === s ? " selected" : "") + '>' + s + " (" + n + ")</option>" : ""; }).join("");
  ssel.onchange = () => { state.sheet = ssel.value || null; renderAll(); };
  c.appendChild(ssel);

  const catsBtn = mk("button", "chip", "Categories…");
  catsBtn.onclick = () => renderCatMenu(catsBtn);
  c.appendChild(catsBtn);

  const nsel = Object.keys(state.picked).length;
  if (nsel) {
    const b = mk("button", "chip", nsel + " selected — move to…");
    b.style.cssText = "background:var(--accent);color:#fff;border-color:var(--accent)";
    b.onclick = () => renderMoveMenu(b);
    c.appendChild(b);
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
}

/** Which categories to show. Hiding is per person and remembered. */
function renderCatMenu(anchor) {
  const old = $("#catmenu"); if (old) { old.remove(); return; }
  const cats = [["deliver", "Ready to deliver"], ["floor", "On floor"], ["ready", "Waiting for floor"],
                ["office", "In office"], ["collect", "Collect & supply"], ["wonttake", "Won't take"],
                ["secondhand", "Second hand"]];
  const m = document.createElement("div"); m.id = "catmenu"; m.className = "menu";
  const r = anchor.getBoundingClientRect();
  m.style.left = Math.max(8, r.left) + "px"; m.style.top = (r.bottom + 6) + "px";
  m.innerHTML = '<div class="kick" style="padding:4px 10px 8px">Show which categories</div>' +
    cats.map(cc => '<label class="mrow"><input type="checkbox" data-k="' + cc[0] + '"' +
      (state.hidden[cc[0]] ? "" : " checked") + '> ' + cc[1] +
      ' <span style="color:var(--ink-4)">' + live().filter(j => catOf(j) === cc[0]).length + '</span></label>').join("") +
    '<div style="display:flex;gap:6px;padding:8px 10px 4px;border-top:1px solid var(--line);margin-top:6px">' +
    '<button class="chip" id="mall">Show all</button><button class="chip" id="mnone">Hide all</button></div>';
  document.body.appendChild(m);
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
  const r = anchor.getBoundingClientRect();
  m.style.left = Math.max(8, r.left) + "px"; m.style.top = (r.bottom + 6) + "px";
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

function rowHtml(j, i, max) {
  const c = comp(j), T = tot(c), w = (T / max) * 110, st = label(j), green = !!j.done;
  const fab = j.prods.some(p => (p.st || []).indexOf("process") >= 0);
  const cpn = cpInProgress(j);
  const picked = !!state.picked[j.id];
  return '<div class="row' + (state.sel === j.id ? " on" : "") + (green ? " ready" : "") + (picked ? " picked" : "") +
    '" data-id="' + j.id + '" draggable="true" style="animation-delay:' + Math.min(i * 3, 200) + 'ms">' +
    '<span class="pickcell"><input type="checkbox" class="pick"' + (picked ? " checked" : "") + '></span>' +
    '<span class="tab jid" style="font-weight:600;color:' + (j.urg ? "var(--urgent)" : "var(--ink)") + '">' + esc(j.id) + '</span>' +
    '<span class="ell">' + esc(j.cust || "—") + '</span>' +
    '<span class="ell" style="color:var(--ink-2)">' + esc(j.area || "—") + '</span>' +
    '<span><span class="badge" style="background:var(--surface-2);color:var(' + st.c + ')">' + st.l + '</span></span>' +
    '<span class="tab" style="color:var(--ink-2)">' + j.wnd + " / " + j.drs + '</span>' +
    '<span style="display:flex;align-items:center;gap:8px"><span class="mini" style="width:110px">' +
      '<i style="width:' + (T ? c.f / T * w : 0) + 'px;background:var(--f)"></i>' +
      '<i style="width:' + (T ? c.s / T * w : 0) + 'px;background:var(--s)"></i>' +
      '<i style="width:' + (T ? c.t / T * w : 0) + 'px;background:var(--t)"></i></span>' +
      '<span class="tab" style="font-size:12px;color:var(--ink-3)">' + T + '</span></span>' +
    '<span style="display:flex;gap:4px;overflow:hidden">' +
      j.sheets.slice(0, 2).map(s => '<span class="stn">' + esc(s) + '</span>').join("") +
      (j.urg ? '<span class="badge" style="background:var(--urgent-bg);color:var(--urgent)">Urgent</span>' : "") +
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

function renderRows() {
  const list = filtered();
  const max = Math.max(1, ...list.map(j => tot(comp(j))));
  const host = $("#rows");

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
    host.innerHTML = keys.length ? keys.map(g => {
      const name = names[g] || (state.view + " group " + g);
      const open = !state.collapsed[state.view + "|" + g];
      const q = (state.gq && state.gq[state.view + "|" + g]) || "";
      let rows = groups[g];
      if (q) rows = rows.filter(j => (j.id + " " + j.cust + " " + j.area).toLowerCase().indexOf(q.toLowerCase()) >= 0);
      return '<div class="grp" data-g="' + g + '">' +
        '<div class="ghead"><button class="gtog">' + (open ? "▾" : "▸") + '</button>' +
        '<span class="gname">' + esc(name) + '</span>' +
        '<span class="gcount">' + rows.length + (q ? " of " + groups[g].length : "") + '</span>' +
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
function openDrawer() { if (!$("#dhost")) { const d = document.createElement("div"); d.id = "dhost"; document.body.appendChild(d); } renderDrawer(); }
function closeDrawer() { state.sel = null; state.edit = false; const h = $("#dhost"); if (h) h.remove(); renderRows(); }

function renderDrawer() {
  const host = $("#dhost"); if (!host) return;
  const j = byId(state.sel); if (!j) return;
  const st = label(j), ed = state.edit;
  const isCS = /^[CS]\d/.test(j.id);
  const readyName = isCS ? "Collect & supply only" : "Ready to fit";
  const hint = t => '<div style="font-size:11.5px;color:var(--ink-4);margin:-6px 0 12px;line-height:1.4">' + t + '</div>';
  /* the 12 s poll re-renders the drawer; typing a count into a box must not be
     wiped out from under the fingers half way through */
  const act = document.activeElement;
  const typing = act && act.dataset && act.dataset.cpin ? { item: act.dataset.cpin, val: act.value } : null;
  host.innerHTML = '<div class="scrim" id="dscrim"></div><div class="drawer">' +
    '<div class="dhead"><div><div style="display:flex;align-items:baseline;gap:9px;flex-wrap:wrap">' +
      '<span class="cond tab" style="font-size:29px;font-weight:700">' + esc(j.id) + '</span>' +
      '<span class="badge" style="background:var(--brand-2);color:#d5d1c8">' + st.l + '</span>' +
      (j.urg ? '<span class="badge" style="background:var(--urgent);color:#fff">Urgent</span>' : "") +
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
      '<div class="sect"><span class="kick">Progress</span><div class="steps">' +
        STEPS.map(p => '<div class="step"><span class="tab" style="font-size:13px;font-weight:600;color:' +
          (j.dates[p[0]] ? "var(--ink)" : "var(--ink-4)") + '">' + dshort(j.dates[p[0]]) + '</span>' +
          '<div class="stepbar" style="background:' + (j.dates[p[0]] ? "var(" + st.c + ")" : "var(--line)") + '"></div>' +
          '<span style="font-size:11px;color:var(--ink-3)">' + p[1] + '</span></div>').join("") + '</div></div>' +
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
      (j.notes.length ? '<div class="sect"><span class="kick">From the sheet</span>' +
        j.notes.map(n => '<div class="note"><div class="kick" style="margin-bottom:3px">' + esc(n.k) + ' · ' + esc(n.s) + '</div>' +
          '<div style="font-size:13px;line-height:1.45">' + esc(n.t) + '</div></div>').join("") + '</div>' : "") +
    '</div></div>';

  if (typing) {
    const box = host.querySelector('.cpin[data-cpin="' + typing.item + '"]');
    if (box) { box.value = typing.val; try { box.focus(); } catch (e) {} }
  }
  $("#dscrim").onclick = closeDrawer;
  $("#dclose").onclick = closeDrawer;
  $("#editbtn").onclick = () => { state.edit = !state.edit; renderDrawer(); };

  if (ed) {
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
  $("#cscrim").onclick = () => host.remove();
  $("#cclose").onclick = () => host.remove();
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
    $("#vscrim").onclick = () => host.remove();
    try { VERSIONS = await CW.listVersions(80); }
    catch (e) {
      host.innerHTML = '<div class="scrim" id="vscrim"></div><div class="logwin vwin"><div class="empty">' + esc(friendly(e)) + '</div></div>';
      $("#vscrim").onclick = () => host.remove(); return;
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
  $("#vscrim").onclick = () => host.remove();
  $("#vclose").onclick = () => host.remove();
  host.querySelectorAll(".vrow").forEach(b => b.onclick = () => { vsel = b.dataset.v; vmode = "since"; renderVersions(); });
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
  await CW.openSession();
  await load("first read…");
  cpReplayQueue();                 // taps this browser owed from a previous visit
  setInterval(poll, 12000);
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
  updateChangeBtn();
  cpWatchExit();
  $("#q").addEventListener("input", e => { state.q = e.target.value; renderRows(); });
  document.addEventListener("keydown", e => {
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
