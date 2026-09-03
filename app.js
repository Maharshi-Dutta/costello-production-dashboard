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
const PSTAT = { "": "—", process: "In fabrication", done: "Process done" };
const STEPS = [["sold", "Sold"], ["stamp", "Stamp"], ["ivana", "Ivana"], ["ready", "Ready to print"], ["floor", "Sent to floor"]];
const SORTS = [["id", "Job no"], ["cat", "Category"], ["urgent", "Urgent first"], ["size", "Biggest first"], ["wait", "Longest wait"], ["county", "County"]];
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
function pend(id, patch) {
  const p = PENDING[id] || (PENDING[id] = { at: 0, prods: {} });
  p.at = Date.now();
  if ("done" in patch) p.done = patch.done;
  if (patch.prod) p.prods[patch.prod.name] = patch.prod.status;
  savePending();
}
function applyPending(list) {
  const now = Date.now();
  let dropped = false;
  Object.keys(PENDING).forEach(id => { if (now - PENDING[id].at > PENDING_MS) { delete PENDING[id]; dropped = true; } });
  if (dropped) savePending();
  return list.map(j => {
    const p = PENDING[j.id];
    if (!p) return j;
    const c = Object.assign({}, j);
    if ("done" in p) c.done = p.done;
    c.prods = j.prods.map(x => Object.prototype.hasOwnProperty.call(p.prods, x.n)
      ? Object.assign({}, x, { st: p.prods[x.n] ? [p.prods[x.n]] : [] }) : x);
    c.stage = c.done ? "deliver" : (c.dates.floor ? "floor" : (c.dates.ready ? "ready" : "office"));
    return c;
  });
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
      const bs = (b.st || [])[0] || "", cs = (c.st || [])[0] || "";
      if (bs !== cs) add(n.id, cap(k) + " status", PSTAT[bs] || "none", PSTAT[cs] || "none");
    });
    Object.keys(ps).forEach(k => { if (!ns[k]) add(n.id, "Product removed - " + cap(k), ps[k].f + "/" + ps[k].s + "/" + ps[k].t, ""); });
    if (p.notes.length !== n.notes.length) {
      const old = p.notes.map(x => x.t), fresh = n.notes.filter(x => old.indexOf(x.t) < 0);
      fresh.forEach(x => add(n.id, "Note added", "", x.t.slice(0, 90)));
    }
  });
  prev.forEach(p => { if (!nm[p.id] && p.cat !== "past") add(p.id, "Removed from the sheet", CATNAME[catOf(p)] || p.cat, ""); });
  return out;
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
async function load(reason, force) {
  if (busy && !force) return;
  busy = true;
  setStatus(reason || "reading sheet…", "busy");
  const t0 = performance.now();
  try {
    const wb = await CW.downloadWorkbook();
    const tDown = performance.now() - t0;
    const prev = ALL;
    const parsed = parseWorkbook(wb);
    BLOCKNAMES = parsed.blockNames || [];
    ALL = applyPending(parsed);   // our own recent writes win over a stale file
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
      /* a dashboard edit also shows up as a cell difference - don't list it twice */
      const mine = new Set(CHANGES.filter(c => c.src === "dashboard" &&
        Date.now() - new Date(c.at).getTime() < 600000).map(c => c.job + "|" + c.what));
      d = d.filter(c => !mine.has(c.job + "|" + c.what));
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

/* ---------- writes ---------- */
const GOLD_HEX = "#FFE699", YELLOW_HEX = "#FFFF00";

async function markReady(job, on) {
  const row = await CW.rowForJob("Production", job.id);   // re-found every time: rows move
  const addr = "A" + row + ":CL" + row;
  if (on) await CW.setFill("Production", addr, GOLD_HEX);
  else await CW.clearFill("Production", addr);
  return row;
}

async function setProductStatus(job, prodName, status) {
  if (!PRODMAP || !PRODMAP.prod[prodName]) throw new Error("That product isn't on the Production sheet.");
  const cols = PRODMAP.prod[prodName];
  const row = await CW.rowForJob("Production", job.id);
  const idx = ["f", "s", "t"].map(k => cols[k]).filter(Boolean);
  const a = CW.A1(Math.min.apply(null, idx)) + row, b = CW.A1(Math.max.apply(null, idx)) + row;
  const addr = a + ":" + b;
  if (status === "process") await CW.setFill("Production", addr, YELLOW_HEX);
  else if (status === "done") await CW.setFill("Production", addr, GOLD_HEX);
  else await CW.clearFill("Production", addr);
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
    opts += '<div class="kick" style="padding:4px 10px 6px">Move ' + jobs.length + ' job' + (jobs.length > 1 ? "s" : "") + ' into</div>' +
      BLOCKNAMES.map((n, i) => '<button class="mrow" data-grp="' + i + '">' + esc(n) + '</button>').join("");
  }
  opts += '<div class="kick" style="padding:10px 10px 6px;border-top:1px solid var(--line);margin-top:6px">Or a category of your own</div>' +
    viewNames().filter(v => v !== "Abin").map(v => '<button class="mrow" data-view="' + esc(v) + '">' + esc(v) + '</button>').join("") +
    '<button class="mrow" id="newcat" style="color:var(--accent);font-weight:600">+ New category from selection…</button>';
  m.innerHTML = opts;
  document.body.appendChild(m);
  m.querySelectorAll("[data-grp]").forEach(b => b.onclick = async () => {
    m.remove(); await assignMany(jobs, state.view === "flat" ? "Abin" : state.view, Number(b.dataset.grp));
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
  jobs.forEach((id, i) => { VIEWS[view][id] = { group: group === "" ? "" : String(group), order: i }; });
  state.picked = {}; renderAll();                     // instant
  let ok = 0;
  for (let i = 0; i < jobs.length; i++) {
    try { await CW.saveAssignment(view, jobs[i], group, i, who); ok++; } catch (e) { console.warn(e); }
  }
  noteChange(jobs.join(", "), "Moved to " + view + (group === "" ? "" : " / " + (BLOCKNAMES[group] || group)), "", view);
  toast(ok + " of " + jobs.length + " saved to " + view);
  setStatus("live");
}

function rowHtml(j, i, max) {
  const c = comp(j), T = tot(c), w = (T / max) * 110, st = label(j), green = !!j.done;
  const fab = j.prods.some(p => (p.st || []).indexOf("process") >= 0);
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
        if (ids.length) await assignMany(ids, state.view, Number(g));
      };
      wireRows(gEl);
    });
  }

  $("#count").textContent = list.length === live().length
    ? "Showing all " + live().length + " jobs on the sheet"
    : "Showing " + list.length + " of " + live().length + " jobs";
}

function renderAll() { renderTiles(); renderChips(); renderRows(); }

/* ---------- drawer ---------- */
function openDrawer() { if (!$("#dhost")) { const d = document.createElement("div"); d.id = "dhost"; document.body.appendChild(d); } renderDrawer(); }
function closeDrawer() { state.sel = null; state.edit = false; const h = $("#dhost"); if (h) h.remove(); renderRows(); }

function renderDrawer() {
  const host = $("#dhost"); if (!host) return;
  const j = byId(state.sel); if (!j) return;
  const st = label(j), ed = state.edit;
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
        (j.done
          ? '<button class="markbtn undo" id="markready">Undo — put back into production</button>'
          : '<button class="markbtn" id="markready">✓ Mark as ready to deliver</button>') : "") +
      '<div class="sect"><span class="kick">Progress</span><div class="steps">' +
        STEPS.map(p => '<div class="step"><span class="tab" style="font-size:13px;font-weight:600;color:' +
          (j.dates[p[0]] ? "var(--ink)" : "var(--ink-4)") + '">' + dshort(j.dates[p[0]]) + '</span>' +
          '<div class="stepbar" style="background:' + (j.dates[p[0]] ? "var(" + st.c + ")" : "var(--line)") + '"></div>' +
          '<span style="font-size:11px;color:var(--ink-3)">' + p[1] + '</span></div>').join("") + '</div></div>' +
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
      (j.prods.length ? '<div class="sect"><div style="display:flex;justify-content:space-between;align-items:baseline">' +
        '<span class="kick">Components</span><span style="font-size:11px;color:var(--ink-3)">Frame · Sashes · Transom</span></div>' +
        '<div class="ptab" style="border-bottom:1.5px solid var(--ink)"><span class="kick">Product</span>' +
        '<span class="kick" style="text-align:right">F</span><span class="kick" style="text-align:right">S</span>' +
        '<span class="kick" style="text-align:right">T</span><span class="kick" style="text-align:right">' + (ed ? "Status" : "") + '</span></div>' +
        j.prods.map(p => {
          const s0 = (p.st || [])[0] || "";
          return '<div class="ptab" data-p="' + esc(p.n) + '"><span>' + esc(cap(p.n)) + '</span>' +
            '<span class="tab" style="text-align:right;color:var(--f);font-weight:600">' + (p.f || "—") + '</span>' +
            '<span class="tab" style="text-align:right;color:var(--s);font-weight:600">' + (p.s || "—") + '</span>' +
            '<span class="tab" style="text-align:right;color:var(--t);font-weight:600">' + (p.t || "—") + '</span>' +
            (ed ? '<select class="txt pst"><option value=""' + (s0 === "" ? " selected" : "") + '>—</option>' +
                  '<option value="process"' + (s0 === "process" ? " selected" : "") + '>In fabrication</option>' +
                  '<option value="done"' + (s0 === "done" ? " selected" : "") + '>Process done</option></select>'
                : '<span style="text-align:right;font-size:11.5px;color:var(--ink-3)">' + (PSTAT[s0] || "") + '</span>') +
            '</div>';
        }).join("") + '</div>' : "") +
      (Object.keys(j.glass).length ? '<div class="sect"><span class="kick">Glass units</span><div style="display:flex;flex-wrap:wrap;gap:6px">' +
        Object.keys(j.glass).map(k => '<span style="font-size:12.5px;padding:5px 10px;border:1px solid var(--line);border-radius:4px;background:var(--surface-2)">' +
          esc(k.toUpperCase()) + ' <strong class="tab">' + j.glass[k] + '</strong></span>').join("") + '</div></div>' : "") +
      (j.notes.length ? '<div class="sect"><span class="kick">From the sheet</span>' +
        j.notes.map(n => '<div class="note"><div class="kick" style="margin-bottom:3px">' + esc(n.k) + ' · ' + esc(n.s) + '</div>' +
          '<div style="font-size:13px;line-height:1.45">' + esc(n.t) + '</div></div>').join("") + '</div>' : "") +
    '</div></div>';

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
        setTimeout(() => load("checking…", true), 45000);   // reconcile once the file catches up
      } catch (e) {
        toast(friendly(e), true);
        mr.disabled = false; renderDrawer();
      }
    };
    host.querySelectorAll(".ptab[data-p]").forEach(rowEl => {
      const sel = rowEl.querySelector(".pst"); if (!sel) return;
      sel.onchange = async () => {
        const name = rowEl.dataset.p, val = sel.value;
        sel.disabled = true;
        try {
          const was = (j.prods.find(p => p.n === name) || {}).st || [];
          await setProductStatus(j, name, val);
          pend(j.id, { prod: { name: name, status: val } });
          ALL = applyPending(ALL);
          noteChange(j.id, cap(name) + " status", PSTAT[was[0] || ""] || "none", PSTAT[val] || "none");
          toast(cap(name) + " → " + (PSTAT[val] || "cleared"));
          renderAll(); renderDrawer();
          setTimeout(() => load("checking…", true), 45000);
        } catch (e) {
          toast(friendly(e), true);
          sel.disabled = false;
        }
      };
    });
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
  $("#whobtn").onclick = () => { if (confirm("Sign out of the dashboard?")) CW.signOut(); };
  await CW.openSession();
  await load("first read…");
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
  updateChangeBtn();
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
