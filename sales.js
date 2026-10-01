/* The Sales page - its screen. docs/specs/2026-09-30-sales-page.md

   Loaded by sales.html only, after app.js, which runs the whole office
   dashboard with window.CW_PAGE = "sales" and calls into here through a few
   hooks: salesWho (whoAmI), salesStart, salesOverlay (after each read),
   salesTiles, salesChips, salesRowHtml, salesDrawerHtml / salesWireDrawer and
   salesRepaint (after the Sales lists are read). The rules and every write are
   in sales-core.js (SALESC); this file draws and wires. It paints no fill, and
   nothing here calls setFill or clearFill.

   Top-level names all start with `sales`/`SALES_`: classic scripts share one
   global scope (HISTORY B29).                                               */
let SALES_WHO = "";                 // the name picked at this desk, once checked against `Sales people`
let SALES_PEOPLE = null;            // active names, or null until read / when the list is missing
let SALES_PEOPLE_OK = null;         // null: not read yet · false: missing or unreadable · true
let SALES_HOLD = {};                // JOB -> { fields..., at } - our writes, until the download shows them
let SALES_GONE = {};                // JOB -> when we deleted it
let SALES_BACKUPS = null, SALES_BACKUPS_OK = null;
let SALES_SIG = "";                 // what the rows last showed from the two lists
const SALES_HOLD_MS = 180000;
const SALES_WHO_KEY = "cw_salesperson";

function salesWho() { return SALES_WHO; }
const salesCtx = () => ({
  CW, who: SALES_WHO, note: noteChange,
  tmplFor: id => LASTWB ? templateForJob(LASTWB.getWorksheet("Production"), id) : null,
  identCol: k => (PRODMAP && PRODMAP.ident && PRODMAP.ident[k]) || null,
  hdrRow: () => (PRODMAP && PRODMAP.hdr && PRODMAP.hdr[0]) || 0
});
/** " disabled" while nobody is picked or any Sales write or move is in flight. */
const salesDis = () => (!SALES_WHO || SALESC.busy()) ? " disabled" : "";
/* a write starting or finishing anywhere on the page redraws every control that writes */
SALESC.onBusy = () => {
  if (state.sel && $("#dhost")) renderDrawer();
  salesDeletedPaint();
};
const salesSection = j => BLOCKNAMES[j.blk] || "";

/* ---- who is at the desk ---------------------------------------------------- */
async function salesReadPeople() {
  try {
    const items = await CW.listItems(SALESC.PEOPLE, { fields: SALESC.PEOPLE_FIELDS });
    SALES_PEOPLE_OK = items !== null;
    SALES_PEOPLE = items ? items.filter(it => SALESC.yes(it.fields.Active))
      .map(it => String(it.fields.Title || "").trim()).filter(Boolean) : null;
  } catch (e) { SALES_PEOPLE_OK = false; SALES_PEOPLE = null; console.warn("[sales] people: " + ((e && e.message) || e)); }
}
function salesSetWho(name) {
  SALES_WHO = name || "";
  try { if (name) localStorage.setItem(SALES_WHO_KEY, name); } catch (e) {}
  const b = $("#whobtn"); if (b) b.textContent = SALES_WHO || "Pick your name";
  const p = $("#spick"); if (p) p.remove();
  if (state.sel && $("#dhost")) renderDrawer();
}
function salesPicker() {
  let host = $("#spick");
  if (!host) { host = document.createElement("div"); host.id = "spick"; document.body.appendChild(host); }
  const names = SALES_PEOPLE || [];
  host.innerHTML = '<div class="gatebox"><h1>Sales</h1>' +
    (SALES_PEOPLE_OK !== true
      ? '<p>The “Sales people” list is not in SharePoint yet (or could not be read), so this page is read only. ' +
        'Ask the admin to add it.</p><button class="signin" id="spro">Look around, read only</button>'
      : !names.length ? '<p>Nobody is marked Active in the “Sales people” list yet, so this page is read only.</p>' +
        '<button class="signin" id="spro">Look around, read only</button>'
      : '<p>Who is using the page?</p><div class="spnames">' +
        names.map(n => '<button class="signin" data-spname="' + esc(n) + '">' + esc(n) + '</button>').join("") + '</div>' +
        (SALES_WHO ? '<button class="ghost" id="spro" style="margin-top:14px">Cancel</button>' : "")) +
    '</div>';
  host.querySelectorAll("[data-spname]").forEach(b => b.onclick = () => salesSetWho(b.dataset.spname));
  const ro = $("#spro"); if (ro) ro.onclick = () => host.remove();
}
async function salesStart() {
  const b = $("#whobtn");
  if (b) { b.textContent = "…"; b.onclick = () => salesPicker(); }
  const rq = $("#reqbtn"); if (rq) rq.onclick = () => salesRequestsWindow();
  const dl = $("#delbtn"); if (dl) dl.onclick = () => salesDeletedWindow();
  await salesReadPeople();
  let saved = "";
  try { saved = localStorage.getItem(SALES_WHO_KEY) || ""; } catch (e) {}
  if (saved && SALES_PEOPLE && SALES_PEOPLE.indexOf(saved) >= 0) salesSetWho(saved);
  else { SALES_WHO = ""; if (b) b.textContent = "Pick your name"; salesPicker(); }
}

/* ---- our own writes, shown at once and held until the download agrees ------ */
function salesHold(id, fields) {
  SALES_HOLD[id] = Object.assign(SALES_HOLD[id] || {}, fields, { at: Date.now() });
  ALL = salesOverlay(ALL); ALL.blockNames = BLOCKNAMES;
}
function salesOverlay(list) {
  const now = Date.now(), ids = {};
  (list || []).forEach(j => { ids[j.id] = 1; });
  Object.keys(SALES_GONE).forEach(id => { if (!ids[id] || now - SALES_GONE[id] > SALES_HOLD_MS) delete SALES_GONE[id]; });
  const put = (j, h) => {
    const c = Object.assign({}, j);
    Object.keys(h).forEach(k => { if (k !== "at") c[k] = h[k]; });
    if ("flag" in h) c.urg = h.flag === "urgent" || (j.urg && j.flag !== "urgent") ? 1 : 0;
    if ("ph" in h) { const d = String(h.ph).replace(/\D/g, ""); c.ph3 = d.length >= 3 ? d.slice(-3) : ""; }
    return c;
  };
  /* THE PRODUCTION SHEET ONLY (owner's rule, 2026-09-18): the Sales page never
     shows a job that is not on `Production`, and every customer field it shows
     or edits is that sheet's own cell (`j.main`, parser.js) - never a value
     another sheet filled into a blank. Wnd/Drs and the products likewise. */
  const HOLD_NONE = {};
  return (list || []).filter(j => j.main && !SALES_GONE[j.id]).map(j => {
    let h = SALES_HOLD[j.id] || HOLD_NONE;
    if (h !== HOLD_NONE && now - h.at > SALES_HOLD_MS) { delete SALES_HOLD[j.id]; h = HOLD_NONE; }
    const c = put(salesMainOnly(j), h);
    /* a move hold keeps its revert data: the parsed job under it gets the
       same Sales fields, so applyPending re-applying onto it keeps both */
    if (j.raw) c.raw = put(salesMainOnly(j.raw), h);
    return c;
  });
}
/** One job as the Sales page may show it: Production's own cells only - the
    shared rule (parser.js productionJob, amendment E), and the Sales page also
    names no other sheet. */
function salesMainOnly(j) {
  /* a job the load already made Production-only comes back as a plain copy,
     so a value applyPending holds over it (done, a move) survives (review 2026-10-01) */
  return productionJob(j);
}

/* ---- tiles, chips, rows ------------------------------------------------------ */
function salesTiles() {
  const all = live(), n = f => all.filter(f).length;
  const defs = [
    { k: null, l: "All jobs", v: all.length, c: "--ink" },
    { k: "urgent", l: "Urgent", v: n(j => j.urg), c: "--urgent" },
    { k: "salesready", l: "Ready to fit", v: n(j => j.done), c: "--done" },
    { k: "booked", l: "Booked", v: n(j => j.flag === "booked"), c: "--green" },
    { k: "inprod", l: "In production", v: n(j => ["floor", "ready", "office"].indexOf(catOf(j)) >= 0), c: "--single" }
  ];
  paintRows($("#tiles"), defs.map((d, i) =>
    '<button class="tile" aria-pressed="' + (state.cat === d.k) + '" data-k="' + (d.k || "") + '"' +
    ' style="animation-delay:' + (i * 25) + 'ms;border-top-color:var(' + d.c + ')"><span class="kick">' + d.l + '</span>' +
    '<span class="n" style="color:var(' + d.c + ')">' + d.v + '</span></button>').join(""));
  $("#tiles").querySelectorAll(".tile").forEach(b => b.onclick = () => {
    const k = b.dataset.k || null; state.cat = state.cat === k ? null : k; renderAll();
  });
}
function salesChips() {
  const c = $("#chips"); c.innerHTML = "";
  const mk = (tag, cls, txt) => { const e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; };
  /* amendment A: the office's list controls - Select all shown, the View
     (flat, or sheet order grouped by section with collapse and Select all per
     section), and which categories to show. No saved views or categories are
     written from here, so only the two built-in views are offered. */
  const all = live();
  if (!state.board && state.view === "flat") {
    const shown = filtered();
    if (shown.length && shown.length < all.length) {
      const on = shown.filter(j => state.picked[j.id]).length;
      const lab = mk("label", "selall"), box = mk("input");
      box.type = "checkbox"; box.className = "pick";
      box.checked = on === shown.length; box.indeterminate = on > 0 && on < shown.length;
      box.onchange = () => pickMany(shown.map(j => j.id), box.checked);
      lab.appendChild(box); lab.appendChild(mk("span", null, "Select all shown (" + shown.length + ")"));
      c.appendChild(lab);
    }
  }
  if (state.view !== "flat" && state.view !== "Abin") state.view = "flat";
  c.appendChild(mk("span", "kick", "View"));
  const vsel = mk("select", "txt");
  vsel.innerHTML = '<option value="flat">Flat list</option><option value="Abin">Sheet order, grouped</option>';
  vsel.value = state.view;
  vsel.onchange = () => { state.view = vsel.value; state.picked = {}; renderAll(); };
  c.appendChild(vsel);
  c.appendChild(mk("span", "kick", "Show"));
  const ssel = mk("select", "txt");
  ssel.id = "showsel";
  ssel.innerHTML = '<option value="">All jobs (' + live().length + ')</option>' +
    STATIONS.map(s => '<option value="' + esc(s[0]) + '"' + (state.board === s[0] ? " selected" : "") + '>' + esc(s[1]) + ' (view only)</option>').join("");
  ssel.onchange = () => {
    state.board = ssel.value || null; renderAll(); stationTick();
    if (state.board === "welding") weldReadIfNeeded(() => renderAll());
    else if (state.board === "glazing") glzReadIfNeeded(() => renderAll());
    else if (state.board === "fabrication") fabrReadIfNeeded("board", () => renderAll());
    else if (state.board) stationReadIfNeeded(() => renderAll());
  };
  c.appendChild(ssel);
  const catsBtn = mk("button", "chip", "Categories…");
  catsBtn.onclick = () => renderCatMenu(catsBtn);            // show/hide, remembered in this browser only
  c.appendChild(catsBtn);
  const nsel = Object.keys(state.picked).length;
  if (nsel) {
    const b = mk("button", "chip", nsel + " selected — move to…");
    b.style.cssText = "background:var(--accent);color:#fff;border-color:var(--accent)";
    b.onclick = () => salesMoveMenu(b);
    c.appendChild(b);
    const x = mk("button", "chip", "Export selected");
    x.onclick = () => fabDo("export");                       // the export window, opened on the ticked jobs
    c.appendChild(x);
    const cl = mk("button", "chip", "clear");
    cl.onclick = () => { state.picked = {}; renderAll(); };
    c.appendChild(cl);
  }
  const sp = mk("span", "kick", "Sort by"); sp.style.marginLeft = "auto"; c.appendChild(sp);
  const sort = mk("select", "txt");
  sort.innerHTML = SORTS.map(p => '<option value="' + p[0] + '"' + (state.sort === p[0] ? " selected" : "") + '>' + p[1] + "</option>").join("");
  sort.onchange = () => { state.sort = sort.value; state.desc = false; renderRows(); salesChips(); };
  c.appendChild(sort);
  const dir = mk("button", "chip", state.desc ? "▼ reversed" : "▲ normal");
  dir.onclick = () => { state.desc = !state.desc; renderRows(); salesChips(); };
  c.appendChild(dir);
  renderFab();                    // the selection wheel, as on the office page
}
/* ---- moving several jobs: a drag onto a section, or "move to" (amendment B) */
async function salesMoveMany(ids, idx) {
  const jobs = (ids || []).map(byId).filter(Boolean);
  const ticked = Object.assign({}, state.picked);
  let r = null;
  try {
    r = await SALESC.moveMany(salesCtx(), jobs, idx, BLOCKNAMES, moveJobsInSheet);
    if (r.refused.length) toast("Not moved - " + SALESC.REFUSED_NOTE + ": " + r.refused.join(", "), true);
    return r;
  } catch (e) { toast(friendly(e), true); return null; }
  finally {
    /* (review) the ticks go only when something moved; otherwise they stay for another try */
    state.picked = r && r.moved > 0 ? {} : ticked;
    renderAll();
  }
}
function salesMoveMenu(anchor) {
  const old = $("#movemenu"); if (old) { old.remove(); return; }
  const jobs = Object.keys(state.picked);
  const nReady = jobs.filter(id => SALESC.moveAllowed(byId(id))).length;
  const m = document.createElement("div"); m.id = "movemenu"; m.className = "menu";
  m.innerHTML = '<div class="kick" style="padding:4px 10px 6px">Move ' + nReady + ' of ' + jobs.length + ' ticked in the Production sheet to</div>' +
    (nReady < jobs.length ? '<div style="padding:0 10px 6px;font-size:11.5px;color:var(--ink-4)">' + (jobs.length - nReady) +
      ' not moved: ' + esc(SALESC.REFUSED_NOTE) + '</div>' : "") +
    SALESC.moveTargets(BLOCKNAMES, -1).map(t => '<button class="mrow" data-grp="' + t.idx + '"' +
      (nReady && SALES_WHO && !SALESC.busy() ? "" : " disabled") + '>' + esc(t.name) + '</button>').join("");
  document.body.appendChild(m);
  menuAt(m, anchor);
  m.querySelectorAll("[data-grp]").forEach(b => b.onclick = async () => { m.remove(); await salesMoveMany(jobs, Number(b.dataset.grp)); });
  setTimeout(() => document.addEventListener("click", function off(e) {
    if (!m.contains(e.target) && e.target !== anchor) { m.remove(); document.removeEventListener("click", off); }
  }), 0);
}
function salesPills(j) {
  const p = (t, bg, fg) => '<span class="badge" style="background:var(' + bg + ');color:var(' + fg + ')">' + esc(t) + '</span>';
  const out = [];
  if (j.urg) out.push(p("Urgent", "--urgent-bg", "--urgent"));
  if (j.done) out.push(p("Ready to fit", "--done-bg", "--done"));
  if (j.flag === "booked") out.push(p("Booked", "--green-bg", "--green"));
  if (!out.length) out.push(p(salesSection(j) || statusWord(j), "--surface-2", "--ink-2"));
  if (SALESC.openFor(SALES_REQS, j.id)) out.push(p("Request open", "--accent-soft", "--single"));
  return out.join("");
}
function salesRowHtml(j, i) {
  const d = SALES_DELIV[j.id];
  const ink = j.flag === "urgent" ? " s-urgent" : j.flag === "booked" ? " s-booked" : "";
  const picked = !!state.picked[j.id];
  return '<div class="row srow' + (state.sel === j.id ? " on" : "") + (j.done ? " ready" : "") + ink + (picked ? " picked" : "") +
    '" data-id="' + esc(j.id) + '" draggable="true" style="' + rowDraw(i) + '">' +
    '<span class="pickcell"><input type="checkbox" class="pick"' + (picked ? " checked" : "") + '></span>' +
    '<span class="tab jid" style="font-weight:600">' + esc(j.id) + '</span>' +
    '<span class="ell">' + esc(j.cust || "—") + '</span>' +
    '<span class="ell">' + esc(j.area || "—") + '</span>' +
    '<span class="ell">' + esc(j.eir || "—") + '</span>' +
    '<span class="ell tab">' + esc(j.ph || "—") + '</span>' +
    '<span class="ell tab">' + esc(j.off || "—") + '</span>' +
    wndDrsCell(j) +
    '<span style="display:flex;gap:4px;overflow:hidden">' + salesPills(j) + '</span>' +
    '<span class="tab">' + (d ? esc(SALESC.dayWords(d.date)) : "") + '</span></div>';
}

/* ---- the drawer's own sections, above the office's (read only) ones --------- */
function salesDrawerHtml(j) {
  const ro = !SALES_WHO, dis = salesDis(), k = s => esc(s + "|" + j.id);
  const d = SALES_DELIV[j.id], flag = SALESC.flagWord(j.flag);
  const reqs = SALESC.newestFirst(SALES_REQS.filter(r => String(r.fields.Job || "").toUpperCase() === j.id));
  const targets = SALESC.moveTargets(BLOCKNAMES, j.blk);
  return (ro ? '<div class="editbar">Pick your name (top right) to make changes. Until then this page is read only.</div>'
    : SALESC.busy() ? '<div class="editbar"><span class="spin"></span> Writing to Excel… changes wait until it has finished.</div>' : "") +
    '<div class="sect"><span class="kick">Customer</span><div class="sform">' +
      SALESC.FIELDS.map(F => '<label>' + esc(F.label) + '<input class="txt" data-keep="' + k(F.key) + '" data-sfield="' + esc(F.key) +
        '" value="' + esc(F.of(j) || "") + '"' + (F.text ? ' inputmode="text"' : "") + dis + '></label>').join("") +
      '</div><div><button class="btn" id="scsave"' + dis + '>Save customer details</button></div></div>' +
    '<div class="sect"><span class="kick">Status</span><div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">' +
      '<button class="chip stog urgent" id="surg" aria-pressed="' + (flag === "urgent") + '"' + dis + '>Urgent</button>' +
      '<button class="chip stog booked" id="sbook" aria-pressed="' + (flag === "booked") + '"' + dis + '>Booked for delivery</button>' +
      '<label style="font-size:12.5px">Delivery date <input class="txt" type="date" id="sdate" data-keep="' + k("date") + '" value="' +
        esc(d ? d.date : "") + '"' + dis + '></label>' +
      (d && d.by ? '<span style="font-size:11.5px;color:var(--ink-4)">set by ' + esc(d.by) + '</span>' : "") + '</div>' +
      '<div style="font-size:11.5px;color:var(--ink-4)">Only the text colour of the row changes in Excel, never its fill. ' +
        'The delivery date is kept in SharePoint, not in the Excel file.</div></div>' +
    '<div class="sect"><span class="kick">Move</span>' +
      (MOVING[j.id] ? '<div><span class="spin"></span> moving the row in Excel…</div>'
        : SALESC.moveAllowed(j) && targets.length
        ? '<div style="display:flex;gap:8px;flex-wrap:wrap"><select class="txt" id="smovesel"' + dis + '>' +
            targets.map(t => '<option value="' + t.idx + '">' + esc(t.name) + '</option>').join("") + '</select>' +
            '<button class="btn" id="smove"' + dis + '>Move job</button></div>'
        : '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center"><button class="btn" disabled>Move job</button>' +
            '<span style="font-size:12px;color:var(--ink-4)">the office has not marked this job ready</span>' +
            '<button class="chip" id="sreqmove"' + dis + '>Send a request</button></div>') +
      '<div><button class="chip sdanger" id="sdel"' + dis + '>Delete job…</button></div></div>' +
    '<div class="sect" id="sreqs"><span class="kick">Requests for this job (' + reqs.length + ')</span>' +
      reqs.map(r => salesRequestHtml(r, false)).join("") +
      '<div style="display:flex;gap:8px;align-items:flex-start;flex-wrap:wrap">' +
        '<select class="txt" id="sreqkind" data-keep="' + k("rk") + '"' + dis + '>' +
          SALESC.REQUEST_KINDS.map(p => '<option value="' + p[0] + '">' + esc(p[1]) + '</option>').join("") + '</select>' +
        '<textarea id="sreqtext" rows="2" data-keep="' + k("rt") + '" placeholder="Ask the office about this job…" style="flex:1;min-width:200px"' + dis + '></textarea>' +
        '<button class="btn" id="sreqsend"' + dis + '>Send</button></div></div>';
}
/** Run one Sales write with the drawer's buttons held, and say how it went. */
async function salesRun(host, fn, ok) {
  host.querySelectorAll(".dbody button, .dbody input, .dbody select, .dbody textarea").forEach(el => { el.disabled = true; });
  host.__html = null;               // the DOM was changed by hand: the next draw must not be skipped as identical
  setStatus("writing to Excel…", "busy");
  try { const r = await fn(); if (ok) toast(typeof ok === "function" ? ok(r) : ok); return r; }
  catch (e) { toast(friendly(e), true); return null; }
  finally {
    setStatus("live"); renderAll();
    const d = $("#dhost");
    if (d) { d.__html = null; renderDrawer(); }   // (review) a failed write must never leave the controls disabled
  }
}
function salesWireDrawer(host, j) {
  const on = (id, fn) => { const el = host.querySelector("#" + id); if (el) el.onclick = fn; };
  /* (review) rows are draggable on this page: a job number dropped on the
     drawer must never land in a customer field */
  if (!host.__nodrop && host.addEventListener) {
    host.__nodrop = 1;
    host.addEventListener("dragover", e => { e.preventDefault(); if (e.dataTransfer) e.dataTransfer.dropEffect = "none"; }, true);
    host.addEventListener("drop", e => { e.preventDefault(); e.stopPropagation(); }, true);
  }
  on("scsave", () => salesRun(host, async () => {
    const changed = [];
    host.querySelectorAll("[data-sfield]").forEach(el => {
      const F = SALESC.fieldOf(el.dataset.sfield);
      if (F && el.value.trim() !== String(F.of(j) || "").trim()) changed.push([F, el.value.trim()]);
    });
    if (!changed.length) throw new Error("Nothing has changed.");
    for (const [F, v] of changed) {
      const r = await SALESC.saveCustomer(salesCtx(), j, F.key, v);
      salesHold(j.id, { [F.key === "phone" ? "ph" : F.key]: v });
      if (r && r.warn) toast(r.warn, true);
    }
    return changed.length;
  }, n => j.id + ": " + n + " customer detail" + (n > 1 ? "s" : "") + " saved in Excel"));
  const colour = async which => {
    const want = SALESC.nextFlag(j.flag, which);
    /* Trade order (pink) / On hold (blue) is somebody's decision: ask first */
    const q = SALESC.replaceQuestion(j.flag, want);
    if (q && !(await salesConfirm(q, "Only the row's text colour changes; the old colour is kept in Sales job backups.", "Replace"))) return;
    salesRun(host, async () => {
      await SALESC.setColour(salesCtx(), j, want);
      salesHold(j.id, { flag: want, flagHex: want ? SALESC.COLOURS[want] : "" });
      return want;
    }, w => j.id + ": text is now " + SALESC.COLOUR_WORD[w].toLowerCase() + " in Excel");
  };
  on("surg", () => colour("urgent"));
  on("sbook", () => colour("booked"));
  const date = host.querySelector("#sdate");
  if (date) date.onchange = () => {
    const prev = SALES_DELIV[j.id] ? SALES_DELIV[j.id].date : "";
    salesRun(host, async () => {
      const to = await SALESC.setDelivery(salesCtx(), j, date.value, prev);
      if (to) SALES_DELIV[j.id] = { date: to, by: SALES_WHO, at: new Date().toISOString() }; else delete SALES_DELIV[j.id];
      return to;
    }, to => j.id + (to ? ": delivery " + SALESC.dayWords(to) : ": delivery date cleared"));
  };
  on("smove", async () => {
    const sel = host.querySelector("#smovesel"); if (!sel) return;
    /* in the same queue as every Sales write, so nothing else runs beside it */
    try { await SALESC.moveJob(salesCtx(), j, Number(sel.value), BLOCKNAMES, moveJobsInSheet); }
    catch (e) { toast(friendly(e), true); }
  });
  on("sreqmove", () => {
    const k = host.querySelector("#sreqkind"), t = host.querySelector("#sreqtext");
    if (k) k.value = "move";
    if (t) { if (!t.value.trim()) t.value = "Please move " + j.id + " to "; try { t.focus(); } catch (e) {} }
  });
  on("sreqsend", () => {
    const k = host.querySelector("#sreqkind"), t = host.querySelector("#sreqtext");
    const text = t ? t.value : "", kind = k ? k.value : "other";
    salesRun(host, async () => {
      const made = await SALESC.sendRequest(salesCtx(), j.id, kind, text);
      SALES_REQS.push({ id: made && made.id, fields: { Job: j.id, Kind: kind, Text: text.trim(), From: SALES_WHO,
                                                      At: new Date().toISOString(), Reply: "", ReplySeen: "No" } });
      if (t) t.value = "";
      return true;
    }, "Request sent to the office");
  });
  on("sdel", async () => {
    const section = salesSection(j);
    const yes = await salesConfirm("Delete " + j.id + " (" + (j.cust || "no customer") + ") from the Production sheet?",
      "The whole row is saved first and can be put back from Deleted jobs, at the bottom of " + (section || "its section") + ".",
      "Delete job");
    if (!yes) return;
    const r = await salesRun(host, () => SALESC.deleteJob(salesCtx(), j), j.id + " deleted - it can be restored from Deleted jobs");
    if (r && r.edgeError) toast(j.id + " was deleted, but the border above it could not be redrawn: " + r.edgeError, true);
    if (r) { SALES_GONE[j.id] = Date.now(); ALL = salesOverlay(ALL); ALL.blockNames = BLOCKNAMES; closeDrawer(); renderAll(); scheduleReconcile(); }
  });
}

/* ---- an in-page confirm, never window.confirm (spec decision 12) ---------- */
function salesConfirm(title, body, okText) {
  return new Promise(resolve => {
    const old = $("#sconfirm"); if (old) old.remove();
    const host = document.createElement("div"); host.id = "sconfirm";
    host.innerHTML = '<div class="scrim"></div><div class="sdialog"><div class="cond" style="font-size:22px;font-weight:700">' + esc(title) +
      '</div><p style="font-size:13px;line-height:1.5">' + esc(body) + '</p><div style="display:flex;gap:8px;justify-content:flex-end">' +
      '<button class="chip" data-sc="0">Cancel</button><button class="btn sdangerbtn" data-sc="1">' + esc(okText) + '</button></div></div>';
    document.body.appendChild(host);
    const done = v => { host.remove(); resolve(v); };
    host.querySelector(".scrim").onclick = () => done(false);
    host.querySelectorAll("[data-sc]").forEach(b => b.onclick = () => done(b.dataset.sc === "1"));
  });
}

/* ---- the two windows ------------------------------------------------------- */
function salesWindow(id, title, sub, body) {
  let host = $("#" + id);
  if (!host) { host = document.createElement("div"); host.id = id; document.body.appendChild(host); }
  host.innerHTML = '<div class="scrim"></div><div class="logwin">' +
    '<div class="dhead"><div><div class="cond" style="font-size:25px;font-weight:700">' + esc(title) + '</div>' +
      '<div style="font-size:12.5px;color:#a8a49a;margin-top:2px">' + esc(sub) + '</div></div>' +
      '<button class="ghost" data-close="1">Close</button></div>' +
    '<div class="logbody" style="padding:12px;display:flex;flex-direction:column;gap:10px">' + body + '</div></div>';
  const close = closeWin(host);
  host.querySelector(".scrim").onclick = close;
  host.querySelector("[data-close]").onclick = close;
  host.querySelectorAll(".jump").forEach(b => b.onclick = () => {
    if (!byId(b.dataset.j)) { toast(b.dataset.j + " is not on the sheet."); return; }
    host.remove(); state.sel = b.dataset.j; renderRows(); openDrawer();
  });
  renderFab();
  return host;
}
/** Every request with its reply, newest first; the replies shown are marked seen. */
function salesRequestsWindow() {
  const all = SALESC.newestFirst(SALES_REQS);
  salesWindow("sqhost", "Requests", SALESC.unreadReplies(all).length + " new repl" + (SALESC.unreadReplies(all).length === 1 ? "y" : "ies") +
    " from the office", SALES_REQS_OK !== true ? '<div class="empty">The “Sales requests” list is not in SharePoint, or could not be read.</div>'
    : all.length ? all.map(r => salesRequestHtml(r, false)).join("") : '<div class="empty">No requests yet.</div>');
  if (SALES_REQS_OK !== true || !SALES_WHO || !SALESC.unreadReplies(all).length) return;   // no write before a name
  const todo = SALESC.unreadReplies(all);
  SALESC.markSeen(salesCtx(), todo)
    .then(() => { todo.forEach(r => { r.fields.ReplySeen = "Yes"; }); salesReqCount(); })
    .catch(e => console.warn("[sales] ReplySeen: " + ((e && e.message) || e)));
}
function salesReqCount() {
  const b = $("#reqbtn"); if (!b) return;
  const n = SALESC.unreadReplies(SALES_REQS).length;
  b.textContent = "Requests" + (n ? " " + n : "");
  if (b.classList) b.classList.toggle("bell", n > 0);
}
const SALES_KIND = { delete: "Deleted", restore: "Restored", edit: "Customer", colour: "Text colour" };
function salesBackupWords(f) {
  if (f.Kind === "delete") return "Deleted from " + (f.Section || "?") + (SALESC.yes(f.Restored) ? " — restored by " + (f.RestoredBy || "?") + ", " + stamp(f.RestoredAt) : "");
  if (f.Kind === "restore") return "Restored to " + (f.Section || "?");
  return (SALES_KIND[f.Kind] || f.Kind) + (f.Field ? ": " + f.Field : "") + " — " + (f.From || "blank") + " → " + (f.To || "blank");
}
async function salesDeletedWindow() {
  salesWindow("sdhost", "Deleted jobs", "reading…", '<div class="empty">Reading the “Sales job backups” list…</div>');
  try {
    const items = await CW.listItems(SALESC.BACKUPS, { fields: SALESC.BACKUP_FIELDS });
    SALES_BACKUPS_OK = items !== null; SALES_BACKUPS = items || [];
  } catch (e) { SALES_BACKUPS_OK = false; SALES_BACKUPS = []; }
  salesDeletedPaint();
}
function salesDeletedPaint() {
  if (!$("#sdhost")) return;
  const items = (SALES_BACKUPS || []).slice().sort((a, b) => String(b.fields.At || "").localeCompare(String(a.fields.At || "")));
  const edits = job => items.filter(x => x.fields.Job === job && (x.fields.Kind === "edit" || x.fields.Kind === "colour"));
  const body = SALES_BACKUPS_OK !== true ? '<div class="empty">The “Sales job backups” list is not in SharePoint, or could not be read, so nothing can be deleted or restored.</div>'
    : !items.length ? '<div class="empty">Nothing has been deleted or changed from this page yet.</div>'
    : '<div class="sdtable"><div class="sdrow sdhead"><span class="kick">Job</span><span class="kick">What happened</span><span class="kick">By</span>' +
        '<span class="kick">When</span><span class="kick">Changes</span><span></span></div>' +
      items.map(x => { const f = x.fields, ch = f.Kind === "delete" ? edits(f.Job) : [];
        return '<div class="sdrow"><span><button class="stn jump" data-j="' + esc(f.Job) + '" style="border:0;cursor:pointer">' + esc(f.Job) + '</button></span>' +
          '<span>' + esc(salesBackupWords(f)) + '</span><span>' + esc(f.Who || "—") + '</span><span class="tab">' + esc(stamp(f.At)) + '</span>' +
          '<span style="font-size:11.5px;color:var(--ink-3)">' + ch.map(c => esc(salesBackupWords(c.fields))).join("<br>") + '</span>' +
          '<span>' + (f.Kind === "delete" && !SALESC.yes(f.Restored)
            ? '<button class="btn" data-srestore="' + esc(x.id) + '"' + salesDis() + '>Restore</button>' : "") + '</span></div>'; }).join("") +
      '</div>';
  const host = salesWindow("sdhost", "Deleted jobs", "Every delete, restore and customer change made on the Sales page", body);
  host.querySelectorAll("[data-srestore]").forEach(b => b.onclick = async () => {
    const it = (SALES_BACKUPS || []).find(x => x.id === b.dataset.srestore); if (!it) return;
    b.disabled = true; setStatus("restoring " + it.fields.Job + " in Excel…", "busy");
    try {
      const r = await SALESC.restoreJob(salesCtx(), it);
      Object.assign(it.fields, { Restored: "Yes", RestoredBy: SALES_WHO, RestoredAt: new Date().toISOString() });
      delete SALES_GONE[String(it.fields.Job).toUpperCase()];
      toast(it.fields.Job + " restored to the bottom of " + r.section + " (row " + r.row + ")");
      if (r.warn || r.edgeError) toast(r.warn || ("The border above the restored row could not be redrawn: " + r.edgeError), true);
      scheduleReconcile(5000);
    } catch (e) { toast(friendly(e), true); }
    setStatus("live");
    salesDeletedPaint();
  });
}

/* ---- after the two lists are read (every 30 s) ------------------------------ */
function salesRepaint() {
  salesReqCount();
  const sig = JSON.stringify(SALES_DELIV) + "|" + SALES_REQS.map(r => r.id + ":" + (r.fields.Reply ? 1 : 0)).join(",");
  if (sig === SALES_SIG) return;
  SALES_SIG = sig;
  if (!state.board) quietRows();
  if (state.sel && $("#dhost")) renderDrawer();      // typed text is kept (data-keep)
}
