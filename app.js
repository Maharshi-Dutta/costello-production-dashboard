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
let state = { q: "", cat: null, sheet: null, sort: "id", sel: null, edit: false };

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
function setStatus(text, kind) {
  $("#status").textContent = text;
  $("#livedot").className = "dot" + (kind ? " " + kind : "");
}

/* ---------- load ---------- */
async function load(reason) {
  if (busy) return;
  busy = true;
  setStatus(reason || "reading sheet…", "busy");
  try {
    const wb = await CW.downloadWorkbook();
    ALL = parseWorkbook(wb);
    const ps = wb.getWorksheet("Production");
    PRODMAP = ps ? mapSheet(ps) : null;
    const m = await CW.lastModified();
    lastStamp = m.at;
    const t = new Date(m.at);
    setStatus("live · updated " + t.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }));
    $("#srcinfo").textContent = "Production sheet · last edited by " + (m.by || "unknown");
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
  let list = live().filter(j => {
    if (state.cat === "urgent") { if (!j.urg) return false; }
    else if (state.cat === "inprod") { if (["floor", "ready", "office"].indexOf(catOf(j)) < 0) return false; }
    else if (state.cat && catOf(j) !== state.cat) return false;
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
  const lab = document.createElement("span"); lab.className = "kick"; lab.textContent = "Sheet"; c.appendChild(lab);
  const ssel = document.createElement("select"); ssel.className = "txt";
  ssel.innerHTML = '<option value="">All sheets (' + all.length + ')</option>' +
    SHEETNAMES.map(s => {
      const n = all.filter(j => j.sheets.indexOf(s) >= 0).length;
      return n ? '<option value="' + s + '"' + (state.sheet === s ? " selected" : "") + '>' + s + " (" + n + ")</option>" : "";
    }).join("");
  ssel.onchange = () => { state.sheet = ssel.value || null; renderAll(); };
  c.appendChild(ssel);
  const sp = document.createElement("span"); sp.style.marginLeft = "auto"; sp.className = "kick"; sp.textContent = "Sort by"; c.appendChild(sp);
  const sort = document.createElement("select"); sort.className = "txt";
  sort.innerHTML = SORTS.map(p => '<option value="' + p[0] + '"' + (state.sort === p[0] ? " selected" : "") + '>' + p[1] + "</option>").join("");
  sort.onchange = () => { state.sort = sort.value; renderRows(); };
  c.appendChild(sort);
}

function renderRows() {
  const list = filtered(), max = Math.max(1, ...list.map(j => tot(comp(j))));
  $("#rows").innerHTML = list.length ? list.map((j, i) => {
    const c = comp(j), T = tot(c), w = (T / max) * 110, st = label(j), green = !!j.done;
    const fab = j.prods.some(p => (p.st || []).indexOf("process") >= 0);
    return '<button class="row' + (state.sel === j.id ? " on" : "") + (green ? " ready" : "") +
      '" data-id="' + j.id + '" style="animation-delay:' + Math.min(i * 4, 260) + 'ms">' +
      '<span class="tab" style="font-weight:600;color:' + (j.urg ? "var(--urgent)" : "var(--ink)") + '">' + esc(j.id) + '</span>' +
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
        (j.sheets.length > 2 ? '<span class="stn">+' + (j.sheets.length - 2) + '</span>' : "") +
        (j.urg ? '<span class="badge" style="background:var(--urgent-bg);color:var(--urgent)">Urgent</span>' : "") +
        (fab ? '<span class="badge" style="background:var(--fab-bg);color:var(--fab)">In fab</span>' : "") +
      '</span>' +
      '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="var(--ink-4)" stroke-width="2.2" stroke-linecap="round"><path d="M9 5l7 7-7 7"/></svg></button>';
  }).join("") : '<div class="empty">No job matches that search or filter.</div>';
  $("#rows").querySelectorAll(".row").forEach(b => b.onclick = () => { state.sel = b.dataset.id; state.edit = false; renderRows(); openDrawer(); });
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
      try {
        const row = await markReady(j, turningOn);
        toast(j.id + (turningOn ? " marked ready to deliver (row " + row + " set gold)" : " put back into production"));
        await load("re-reading…");
      } catch (e) {
        toast("Write failed: " + e.message, true);
        mr.disabled = false; renderDrawer();
      }
    };
    host.querySelectorAll(".ptab[data-p]").forEach(rowEl => {
      const sel = rowEl.querySelector(".pst"); if (!sel) return;
      sel.onchange = async () => {
        const name = rowEl.dataset.p, val = sel.value;
        sel.disabled = true;
        try {
          await setProductStatus(j, name, val);
          toast(cap(name) + " → " + (PSTAT[val] || "cleared"));
          await load("re-reading…");
        } catch (e) {
          toast("Write failed: " + e.message, true);
          sel.disabled = false;
        }
      };
    });
  }
}

/* ---------- theme ---------- */
function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  try { localStorage.setItem("cw_theme", t); } catch (e) {}
  $("#themebtn").textContent = t === "dark" ? "Light" : "Dark";
}

/* ---------- boot ---------- */
async function start() {
  $("#gate").hidden = true;
  $("#topbar").hidden = false;
  $("#main").hidden = false;
  const a = CW.account;
  $("#whobtn").textContent = (a && (a.username || a.name)) || "signed in";
  $("#whobtn").onclick = () => { if (confirm("Sign out of the dashboard?")) CW.signOut(); };
  await CW.openSession();
  await load("first read…");
  setInterval(poll, 20000);
}

(async function boot() {
  let t = "light";
  try { t = localStorage.getItem("cw_theme") || (matchMedia("(prefers-color-scheme:dark)").matches ? "dark" : "light"); } catch (e) {}
  applyTheme(t);
  $("#themebtn").onclick = () => applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  $("#refreshbtn").onclick = () => load("refreshing…");
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
