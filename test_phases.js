/* Offline test of the two dashboard-only features added on 2026-09-07:

   - the phase pipeline: jobPhase(j) -> 0..6, worked out from the dates and the
     checkpoint colours already parsed, and the parser's new "cut" status,
     including where the Cut colour itself comes from (the sheet's own legend);
   - the floating selection wheel: it exists only while jobs are ticked, it
     offers the admin one more option than everyone else, and each of its
     options reaches the function that already did that job, on the ticked ids.

   Nothing here goes near a network or a real workbook: Graph is never loaded,
   fetch() throws if anything tries, and the DOM is a stub that remembers what
   was put into it so a test can look at what the page actually built.
   Run: node test_phases.js                                                   */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, innerWidth: 1280, innerHeight: 800,
                  addEventListener() {}, removeEventListener() {} };
global.performance = { now: () => Date.now() };
global.ExcelJS = ExcelJS;
let FETCHES = 0;
global.fetch = async () => { FETCHES++; throw new Error("neither feature may touch the network"); };

/* Elements remember their html, their children, their attributes and their
   handlers. An element with an id joins REG when it is appended, and leaves it
   again when it is removed - so $("#fabhost") is null exactly when the wheel is
   not on the page, which is half of what this file has to prove. */
const REG = {};
function stubEl(tag, id) {
  let html = "";
  const e = {
    tag: tag || "div", id: id || "", style: {}, dataset: {}, attrs: {}, kids: [],
    textContent: "", value: "", disabled: false, hidden: false, className: "", scrollTop: 0,
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); if (c.id) REG[c.id] = c; return c; },
    remove() { const drop = x => { if (x.id && REG[x.id] === x) delete REG[x.id]; x.kids.forEach(drop); }; drop(e); },
    contains(x) { return x === e || e.kids.some(k => k.contains && k.contains(x)); },
    setAttribute(k, v) { e.attrs[k] = String(v); }, getAttribute(k) { return e.attrs[k]; },
    /* listeners are remembered rather than dropped, so a test can end a
       transition the way a browser would and watch what the page does next.
       fire() takes the rest of the event, so a child's transition arriving at
       a parent's listener - which is what bubbling really looks like from the
       parent's side - can be told apart from the parent's own. */
    on: {},
    addEventListener(t, fn) { (e.on[t] = e.on[t] || []).push(fn); },
    removeEventListener(t, fn) { e.on[t] = (e.on[t] || []).filter(x => x !== fn); },
    fire(t, ev) {
      (e.on[t] || []).slice().forEach(fn => fn(Object.assign({ type: t, target: e }, ev || {})));
    },
    focus() {}, blur() {}, setSelectionRange() {},
    getBoundingClientRect: () => ({ left: 40, top: 500, right: 96, bottom: 556 }),
    querySelector: () => stubEl(), querySelectorAll: () => []
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; } });
  /* the wheel puts where each option lands into CSS custom properties */
  e.style.setProperty = (k, v) => { e.style[k] = String(v); };
  return e;
}
/* the selectors that stand for something the page creates and destroys: they
   answer null until it exists, the way a browser does */
const NULLABLE = ["#fabhost", "#fabbtn", "#dhost", "#xhost", "#ahost", "#chost", "#vhost",
                  "#lhost", "#nhost", "#catmenu", "#movemenu", "#alertmenu"];
const EL = {};
const el = sel => {
  const id = String(sel).charAt(0) === "#" ? String(sel).slice(1) : null;
  if (id && REG[id]) return REG[id];
  if (NULLABLE.indexOf(sel) >= 0) return null;
  return EL[sel] || (EL[sel] = stubEl("div", id || ""));
};
const HANDLERS = {};                 // the page's own document listeners, so a key can be pressed
global.document = {
  documentElement: stubEl(), body: stubEl(), head: stubEl(), activeElement: null,
  title: "Costello Production",
  createElement: t => stubEl(t),
  querySelector: el, querySelectorAll: () => [],
  addEventListener: (t, fn) => { (HANDLERS[t] = HANDLERS[t] || []).push(fn); },
  removeEventListener: (t, fn) => { HANDLERS[t] = (HANDLERS[t] || []).filter(x => x !== fn); }
};
const press = key => (HANDLERS.keydown || []).forEach(fn => fn({ key: key, preventDefault() {} }));

/* Graph is not loaded at all, so there is no route to the network even by
   accident; CW is the small part of it the dashboard calls without a write. */
const LOGS = [];
let WHO = "boss@example.test";
global.CW = {
  account: { get username() { return WHO; }, name: "The Boss" },
  appendLog: (who, job, what, from, to) => { LOGS.push({ who, job, what, from, to }); return Promise.resolve(true); },
  initAuth: async () => null, openSession: async () => {}, signOut() {}
};

/* ---------- load the page's own code, in the page's own order ---------- */
const run = f => vm.runInThisContext(fs.readFileSync(__dirname + "/" + f, "utf8"), { filename: f });
run("parser.js");
run("checkpoints.js");
global.CP = window.CP;
run("export.js");
run("app.js");

const TOASTS = [];
global.toast = (m, err) => TOASTS.push({ m: String(m), err: !!err });
const set = src => vm.runInThisContext(src);

/* ---------- fixtures ---------- */
const mkJob = o => Object.assign({
  id: "R0001", cust: "Ann", area: "Cork", eir: "", off: "", colour: "", ph3: "",
  wnd: 0, drs: 0, glass: {}, prods: [], notes: [], sheets: ["Production"], src: {},
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0,
  cp: { win: "", drs: "", glass: {}, prod: {} }
}, o || {});
/* a job carrying one product with all three sub-columns, plus windows/glass */
const withWork = (cpProd, extra) => mkJob(Object.assign({
  wnd: 4, drs: 0, glass: { tg: 6 },
  prods: [{ n: "7000 casement", f: 3, s: 2, t: 1, st: [] }],
  cp: { win: "", drs: "", glass: {}, prod: { "7000 casement": cpProd || {} } }
}, extra || {}));

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ---- 1. the pipeline itself ---- */
  assert.deepStrictEqual(PHASES, ["In office", "Sent to floor", "Cutting", "In fabrication",
    "In glazing", "Quality check", "Fitted / delivered"]);
  pass("seven phases, in the order the work happens");

  /* 0 - in office: no floor date and nothing coloured */
  assert.strictEqual(jobPhase(withWork({})), 0);
  assert.strictEqual(phaseName(withWork({})), "In office");
  /* 1 - sent to floor: the date alone is enough */
  assert.strictEqual(jobPhase(withWork({}, { dates: { floor: "2026-09-01" } })), 1);
  /* ...and so is a checkpoint that has moved, with no date at all */
  assert.strictEqual(jobPhase(withWork({}, { cp: { win: "done", drs: "", glass: {}, prod: {} } })), 1);
  pass("0 In office / 1 Sent to floor: the date, or any checkpoint that has moved");

  /* 2 - cutting: the sheet's Cut green on a product cell, nothing fabricated */
  const cutting = withWork({ f: "cut" }, { dates: { floor: "2026-09-01" } });
  assert.strictEqual(jobPhase(cutting), 2);
  assert.strictEqual(cpStatus(cutting, "prod:7000 casement:f"), "cut");
  /* the same job once a frame goes yellow: fabrication outranks cutting */
  assert.strictEqual(jobPhase(withWork({ f: "cut", s: "process" })), 3);
  pass("2 Cutting, and 3 In fabrication the moment anything goes yellow");

  /* 3 - windows or doors in process count as fabrication too */
  assert.strictEqual(jobPhase(withWork({}, { cp: { win: "process", drs: "", glass: {}, prod: {} } })), 3);
  /* glass on its own does not: that is the glazing end of the job */
  assert.strictEqual(jobPhase(withWork({}, { cp: { win: "", drs: "", glass: { tg: "process" }, prod: {} } })), 1);
  pass("windows/doors in process is fabrication; glass on its own is not");

  /* 4 - in glazing: every F/S/T done, the rest of the job not */
  const glazing = withWork({ f: "done", s: "done", t: "done" },
    { cp: { win: "process", drs: "", glass: { tg: "" }, prod: { "7000 casement": { f: "done", s: "done", t: "done" } } } });
  assert.strictEqual(jobPhase(glazing), 4, "products finished beats windows still in process");
  /* 5 - quality check: everything ticked, nobody has marked it ready yet */
  const qc = withWork({}, { cp: { win: "done", drs: "", glass: { tg: "done" },
    prod: { "7000 casement": { f: "done", s: "done", t: "done" } } } });
  assert.strictEqual(jobPhase(qc), 5);
  assert.strictEqual(jobPhase(mkJob({ done: 1 })), 6, "gold row");
  assert.strictEqual(jobPhase(mkJob({ cat: "past" })), 6, "no longer on the Production sheet");
  assert.strictEqual(jobPhase(Object.assign(mkJob({ done: 1 }), qc, { done: 1 })), 6,
    "ready to deliver wins over everything below it");
  pass("4 In glazing / 5 Quality check / 6 Fitted - delivered");

  /* a job with no countable work at all still reads sensibly */
  assert.strictEqual(jobPhase(mkJob({})), 0);
  assert.strictEqual(jobPhase(mkJob({ dates: { floor: "2026-09-01" } })), 1);
  /* products only, all done: that is everything done, so quality check */
  assert.strictEqual(jobPhase(mkJob({ prods: [{ n: "x", f: 2, s: 0, t: 0, st: [] }],
    cp: { win: "", drs: "", glass: {}, prod: { x: { f: "done" } } } })), 5);
  pass("a job with no items, and a job that is nothing but products, both read sensibly");

  /* ---- 2. the parser reads the Cut colour off the sheet's own legend ---- */
  const head = ws => {
    const put = (r, c, v) => { ws.getCell(r, c).value = v; };
    put(2, 4, "DATES ON CONTRACT"); put(3, 4, "SOLD"); put(2, 5, "CUSTOMER");
    put(2, 13, "QUANTITY"); put(3, 13, "WND"); put(3, 14, "DRS");
    put(2, 20, "7000 CASEMENT"); put(3, 20, "F"); put(3, 21, "S"); put(3, 22, "T");
    put(2, 51, "GLASS UNITS"); put(3, 51, "TG");
  };
  const put = (sheet, r, c, v, colour) => {
    const cell = sheet.getCell(r, c);
    if (v !== undefined && v !== null) cell.value = v;
    if (colour) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + colour } };
  };
  /** one Production sheet, one job, F/S/T coloured as asked, plus an optional
      legend cell on row 1 */
  function book(legend, legendFill, f, s, t) {
    const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet("Production");
    head(ws);
    if (legend) put(ws, 1, 2, legend, legendFill);
    put(ws, 6, 3, "R0001"); put(ws, 6, 5, "Ann");
    put(ws, 6, 13, 4); put(ws, 6, 51, 6, "FFFFFF");
    put(ws, 6, 20, 3, f); put(ws, 6, 21, 2, s); put(ws, 6, 22, 1, t);
    return parseWorkbook(wb).find(x => x.id === "R0001");
  }
  const sub = j => ["f", "s", "t"].map(k => (j.cp.prod["7000 casement"] || {})[k] || "");

  /* the legend says 92D050, so that - and only that - is the Cut colour */
  let j = book("Cut=", "92D050", "92D050", "00B050", "FFFF00");
  assert.deepStrictEqual(sub(j), ["cut", "", "process"],
    "the legend's own green is cut; another green the legend did not name is not");
  assert.deepStrictEqual(cutColours(new ExcelJS.Workbook().addWorksheet("x")).size, 3);
  pass("the Cut colour comes from the legend cell, and nothing else counts as cut");

  /* no legend at all: the greens the sheet has used are taken as cut */
  j = book(null, null, "00B050", "C6EFCE", "");
  assert.deepStrictEqual(sub(j), ["cut", "cut", ""], "both fallback greens read as cut");
  /* a heading that merely contains the word, in no colour at all, is not a key */
  j = book("PVC Cutting & Welding", null, "00B050", "", "");
  assert.deepStrictEqual(sub(j), ["cut", "", ""], "an uncoloured heading is passed over");
  /* nor is one filled yellow: that is the fabrication colour, not a legend */
  j = book("Frames cut =", "FFFF00", "92D050", "", "");
  assert.deepStrictEqual(sub(j), ["cut", "", ""], "a yellow 'cut' cell is not a colour key");
  pass("no readable legend: the sheet's known greens are used instead");

  /* cut ranks below process, and neither disturbs anything verify.js compares */
  j = book("Cut=", "92D050", "92D050", "FFFF00", "FFE699");
  assert.deepStrictEqual(sub(j), ["cut", "process", "done"]);
  assert.deepStrictEqual(j.prods[0].st.slice().sort(), ["done", "process"],
    "a cut cell adds nothing to the product status the Python reader also produces");
  assert.strictEqual(j.done, 0);
  assert.strictEqual(j.wnd, 4);
  assert.strictEqual(itemState(j, "prod:7000 casement:f").status, "",
    "cut is not progress you can tick: the checkpoint still reads as not started");
  assert.strictEqual(jobPhase(j), 3, "one yellow frame puts the whole job in fabrication");
  pass("cut sits below process, and leaves the fields verify.js compares untouched");

  /* ---- 3. the job row and the drawer ---- */
  set("BLOCKNAMES = ['Can sell as second hand', \"Ready, customer won't take\", " +
      "'Collect & supply only', 'Ready to fit', 'In production'];");
  const inProd = withWork({ f: "cut" }, { id: "R0001", dates: { floor: "2026-09-01" }, stage: "floor" });
  const ready = mkJob({ id: "R0002", done: 1, cat: "active", blk: 3, stage: "floor" });
  const collect = mkJob({ id: "R0003", cat: "collect", blk: 2 });
  global.__jobs = [inProd, ready, collect];
  set("ALL = __jobs; CHANGES = []; state.sel = null; state.picked = {}; state.view = 'flat';");

  assert.ok(rowHtml(inProd, 0, 10).indexOf(">Cutting<") >= 0,
    "a job in production wears its phase, not its stage");
  assert.strictEqual(rowHtml(inProd, 0, 10).indexOf(">On floor<"), -1);
  assert.ok(rowHtml(ready, 0, 10).indexOf(">Ready to deliver<") >= 0, "the gold badge is unchanged");
  assert.ok(rowHtml(collect, 0, 10).indexOf(">Collect/Supply<") >= 0, "so is the section badge");
  pass("the list row shows the phase for jobs in production, and the old badges elsewhere");

  /* Phase 0 covers both "ready to print, waiting for the floor" and "still in
     the office" - nothing on the sheet has moved in either case - so a job
     stuck at phase 0 keeps the stage word, which does tell the two apart. */
  const waiting = withWork({}, { id: "R0004", stage: "ready", dates: { ready: "2026-09-02" } });
  const office = withWork({}, { id: "R0005", stage: "office" });
  assert.strictEqual(jobPhase(waiting), 0);
  assert.strictEqual(jobPhase(office), 0);
  assert.ok(rowHtml(waiting, 0, 10).indexOf(">Waiting<") >= 0,
    "ready to print but not yet on the floor still reads as Waiting");
  assert.strictEqual(rowHtml(waiting, 0, 10).indexOf(">In office<"), -1,
    "and is not collapsed into In office by the phase word");
  assert.ok(rowHtml(office, 0, 10).indexOf(">In office<") >= 0);
  /* once anything on the sheet has moved, the phase is the better word for both */
  const waitingCut = withWork({ f: "cut" }, { id: "R0004", stage: "ready", dates: { ready: "2026-09-02" } });
  assert.ok(rowHtml(waitingCut, 0, 10).indexOf(">Cutting<") >= 0);
  const officeMoved = withWork({ f: "process" }, { id: "R0005", stage: "office" });
  assert.ok(rowHtml(officeMoved, 0, 10).indexOf(">In fabrication<") >= 0);
  pass("Waiting and In office stay apart until something on the sheet actually moves");

  set("state.sel = 'R0001';");
  openDrawer();
  let drawer = el("#dhost").innerHTML;
  PHASES.forEach(p => assert.ok(drawer.indexOf(p) >= 0, "the drawer shows the step: " + p));
  assert.ok(drawer.indexOf('Now: <strong>Cutting</strong>') >= 0, "and names the one it is on");
  assert.ok(drawer.indexOf('class="phstep now" aria-current="step"><div class="phbar"></div>' +
    '<span class="phlab">Cutting</span><span class="phtag">Now</span>') >= 0,
    "the current step is tagged in words, not by colour alone");
  assert.ok(drawer.indexOf('class="phstep done"') >= 0, "the steps behind it are filled");
  assert.ok(drawer.indexOf('color:#d5d1c8">Cutting</span>') >= 0,
    "the badge at the head of the drawer carries the phase as well");
  assert.strictEqual(drawer.indexOf("On floor"), -1, "and not the old stage word");
  pass("the drawer leads with the seven-step pipeline, current step named in words");

  set("state.sel = 'R0002';");
  openDrawer();
  assert.ok(el("#dhost").innerHTML.indexOf('color:#d5d1c8">Ready to deliver</span>') >= 0,
    "a job marked ready keeps its own badge at the head of the drawer");
  set("state.sel = 'R0001';");
  openDrawer();
  drawer = el("#dhost").innerHTML;

  assert.ok(drawer.indexOf('id="datetog" aria-expanded="false"') >= 0, "Dates starts folded away");
  assert.strictEqual(drawer.indexOf("Ready to print"), -1, "so the five date steps are not drawn");
  assert.ok(drawer.indexOf("Sent to floor 01/09") >= 0, "the fold shows how far the dates got");
  el("#datetog").onclick();
  drawer = el("#dhost").innerHTML;
  assert.ok(drawer.indexOf('id="datetog" aria-expanded="true"') >= 0);
  assert.ok(drawer.indexOf("Ready to print") >= 0 && drawer.indexOf(">Sold<") >= 0,
    "opened, all five date steps are there");
  assert.strictEqual(JSON.parse(mem.cw_collapsed).dates, 0, "and the choice is remembered");
  el("#datetog").onclick();
  assert.strictEqual(JSON.parse(mem.cw_collapsed).dates, 1);
  assert.ok(el("#dhost").innerHTML.indexOf("Checkpoints") >= 0, "the Checkpoints section is untouched");
  pass("the date steps fold under the pipeline, and which way is remembered like a group");

  closeDrawer();
  assert.strictEqual(el("#dhost"), null);

  /* ---- 4. the floating selection wheel ---- */
  set("CONFIG = { admin: 'boss@example.test' }; ALERTS = {}; VIEWS = { 'Fitting week': {} };");
  set("state.picked = {}; state.sel = null;");
  renderChips();
  assert.strictEqual(el("#fabhost"), null, "nothing ticked: no wheel at all");

  set("state.picked = { R0001: 1, R0003: 1 };");
  renderChips();
  let host = el("#fabhost");
  assert.ok(host, "two jobs ticked: the wheel appears");
  assert.ok(el("#fabbtn").innerHTML.indexOf('<span class="fabn">2</span><span class="fabk">ticked</span>') === 0,
    "and says how many");
  assert.ok(el("#fabbtn").innerHTML.indexOf('class="fabx"') > 0,
    "with the plus the CSS turns 315° into a cross while it is open");
  assert.strictEqual(el("#fabbtn").attrs["aria-label"], "2 jobs ticked — actions");
  assert.deepStrictEqual(host.kids.map(k => k.dataset.fab).filter(Boolean),
    ["alert", "export", "move", "clear"], "the admin gets all four options");
  const opt = k => host.kids.find(x => x.dataset.fab === k);
  /* a full circle of radius 100 around the ring's middle, first option straight
     up: on a 1280x800 screen that middle is 150 px in from the right and the
     bottom, which is 104 px in from where the closed button sits (18 + 56/2). */
  assert.deepStrictEqual(fabGeom(4, 1280, 800), { r: 100, n: 4, cx: 150, cy: 150 });
  assert.deepStrictEqual(host.kids.filter(x => x.dataset.fab).map(x => x.style["--fx"] + "/" + x.style["--fy"]),
    ["-104px/-204px", "-4px/-104px", "-104px/-4px", "-204px/-104px"]);
  assert.strictEqual(host.style["--fdx"], "-104px", "and the button itself glides in to that middle");
  assert.strictEqual(host.style["--fdy"], "-104px");
  /* the stagger is the CSS's own: --i counts them out, --rev counts them back */
  assert.deepStrictEqual(host.kids.filter(x => x.dataset.fab).map(x => x.style["--i"]), ["0", "1", "2", "3"]);
  assert.deepStrictEqual(host.kids.filter(x => x.dataset.fab).map(x => x.style["--rev"]), ["3", "2", "1", "0"]);
  assert.ok(el("#fab-export").innerHTML.indexOf("<svg") >= 0 &&
            el("#fab-export").innerHTML.indexOf('class="fabl">Export<') >= 0,
    "each option is an inline-SVG icon with its name written underneath it");
  assert.strictEqual(host.className, "fabwrap", "closed until it is tapped");
  pass("the wheel exists only while jobs are ticked, and opens on a full circle of radius 100");

  /* ---- every option on the circle, and all of it on the screen ----
     Each option is measured at its full 76 x 68 box (the 52 px button plus its
     label), centred on its point of the circle, so this is the worst case
     rather than a lucky one. */
  const ring = (vw, vh, n) => {
    window.innerWidth = vw; window.innerHeight = vh;
    set("state.picked = { R0001: 1, R0003: 1 };");
    renderChips();
    const h = el("#fabhost"), g = fabGeom(n, vw, vh);
    return { g: g, cx: vw - g.cx, cy: vh - g.cy,          // where the open ring's middle sits
      opts: h.kids.filter(x => x.dataset.fab).map(x => {
        const ox = vw - 46 + parseFloat(x.style["--fx"]);  // 46 = the closed button's centre
        const oy = vh - 46 + parseFloat(x.style["--fy"]);
        return { k: x.dataset.fab, x: ox, y: oy,
                 x1: ox - 38, x2: ox + 38, y1: oy - 34, y2: oy + 34 };
      }) };
  };
  const checkRing = (vw, vh, n) => {
    const r = ring(vw, vh, n), where = " at " + vw + "x" + vh + " with " + n + " options";
    assert.strictEqual(r.opts.length, n, "every option is on the ring" + where);
    r.opts.forEach((p, i) => {
      const d = Math.sqrt(Math.pow(p.x - r.cx, 2) + Math.pow(p.y - r.cy, 2));
      assert.ok(Math.abs(d - 100) <= 1, p.k + " is " + d.toFixed(1) + " px out, not 100" + where);
      /* clockwise from straight up, evenly spaced: 90° apart for four, 120° for three */
      const deg = ((Math.atan2(p.x - r.cx, r.cy - p.y) * 180 / Math.PI) + 360) % 360;
      const want = (i * (360 / n)) % 360;
      assert.ok(Math.min(Math.abs(deg - want), 360 - Math.abs(deg - want)) <= 1,
        p.k + " sits at " + deg.toFixed(1) + "°, not " + want + "°" + where);
      assert.ok(p.x1 >= 8, p.k + " runs off the left edge" + where);
      assert.ok(p.x2 <= vw - 8, p.k + " runs off the right edge" + where);
      assert.ok(p.y1 >= 8, p.k + " runs off the top" + where);
      assert.ok(p.y2 <= vh - 8, p.k + " runs off the bottom" + where);
    });
    return r;
  };
  [[1280, 800], [390, 844], [390, 640], [320, 640]].forEach(v => checkRing(v[0], v[1], 4));
  /* the wished-for inset is 150 px, but on a phone it has to give way, or the
     right-hand option would hang off the edge - the ring is never cropped */
  assert.strictEqual(fabGeom(4, 320, 640).cx, 146);
  assert.ok(fabGeom(4, 320, 640).cx < fabGeom(4, 1280, 800).cx,
    "a narrow phone pulls the ring's middle in rather than letting it hang off the screen");
  window.innerWidth = 1280; window.innerHeight = 800;
  renderChips();
  host = el("#fabhost");
  pass("every option sits on the circle and stays on screen, at 1280, 390 and 320 px wide");

  WHO = "someone@example.test";
  renderChips();
  host = el("#fabhost");
  assert.deepStrictEqual(host.kids.map(k => k.dataset.fab).filter(Boolean), ["export", "move", "clear"],
    "everyone else gets three: alerts are the administrator's");
  /* three options share the same circle, 120° apart, still starting at the top */
  checkRing(1280, 800, 3);
  checkRing(320, 640, 3);
  window.innerWidth = 1280; window.innerHeight = 800;
  WHO = "boss@example.test";
  renderChips();
  host = el("#fabhost");
  pass("only the administrator is offered Alert to…");

  /* opening, closing, and Escape */
  el("#fabbtn").onclick({ stopPropagation() {} });
  assert.strictEqual(el("#fabhost").className, "fabwrap open");
  assert.strictEqual(el("#fabbtn").attrs["aria-expanded"], "true");
  el("#fabbtn").onclick({ stopPropagation() {} });
  assert.strictEqual(el("#fabhost").className, "fabwrap", "tapping it again closes the fan");
  el("#fabbtn").onclick({ stopPropagation() {} });
  press("Escape");
  assert.strictEqual(el("#fabhost").className, "fabwrap", "and so does Escape");
  pass("the fan opens on a tap and closes on a second tap or on Escape");

  /* Move to… - the existing menu, built for the ticked jobs */
  el("#fabbtn").onclick({ stopPropagation() {} });
  let ids = opt("move").onclick({ stopPropagation() {} });
  assert.deepStrictEqual(ids, ["R0001", "R0003"]);
  assert.ok(el("#movemenu"), "the move menu is open");
  assert.ok(el("#movemenu").innerHTML.indexOf("Move 2 jobs in the Production sheet to") >= 0,
    "on the two ticked jobs");
  assert.ok(el("#movemenu").innerHTML.indexOf("Fitting week") >= 0, "with the dashboard-only views too");
  assert.strictEqual(el("#fabhost").className, "fabwrap", "and the fan closed behind it");
  el("#movemenu").remove();

  /* Alert to… - likewise */
  el("#fabbtn").onclick({ stopPropagation() {} });
  ids = opt("alert").onclick({ stopPropagation() {} });
  assert.deepStrictEqual(ids, ["R0001", "R0003"]);
  assert.ok(el("#alertmenu").innerHTML.indexOf("Email alerts for 2 jobs to") >= 0);
  el("#alertmenu").remove();
  pass("Move to… and Alert to… open the menus that already existed, on the ticked ids");

  /* Export - the window opens with the scope already on the ticked jobs */
  set("XSTATE = null;");
  el("#fabbtn").onclick({ stopPropagation() {} });
  opt("export").onclick({ stopPropagation() {} });
  assert.ok(el("#xhost"), "the Export window is open");
  assert.strictEqual(XSTATE.f.scope, "ticked", "with Scope preset to the ticked jobs");
  assert.ok(el("#xhost").innerHTML.indexOf('aria-pressed="true" data-xset="f.scope" data-xval="ticked"') >= 0,
    "and the Ticked jobs chip is the one showing as chosen");
  assert.ok(el("#xhost").innerHTML.indexOf("Ticked jobs (2)") >= 0, "counting both of them");
  assert.strictEqual(el("#fabhost").className, "fabwrap over",
    "a window is open, so on a phone the wheel takes itself out of the way");
  el("#xhost").remove();
  renderChips();
  assert.strictEqual(el("#fabhost").className, "fabwrap", "and comes back when the window closes");
  pass("Export opens the window with Scope already set to the ticked jobs");

  /* the drawer counts as a window too */
  set("state.sel = 'R0001';");
  openDrawer();
  assert.strictEqual(el("#fabhost").className, "fabwrap over");
  closeDrawer();
  assert.strictEqual(el("#fabhost").className, "fabwrap");

  /* ...and so do the three log-style windows, which all close through closeWin */
  renderAlertsWindow();
  assert.ok(el("#ahost"), "the Alerts window is open");
  assert.strictEqual(el("#fabhost").className, "fabwrap over");
  el("#aclose").onclick();
  assert.strictEqual(el("#ahost"), null, "Close really closes it");
  assert.strictEqual(el("#fabhost").className, "fabwrap", "and the wheel comes back");
  renderAlertsWindow();
  el("#ascrim").onclick();
  assert.strictEqual(el("#fabhost").className, "fabwrap", "the scrim does the same");

  renderChanges();
  assert.ok(el("#chost"), "the Changes window is open");
  assert.strictEqual(el("#fabhost").className, "fabwrap over");
  el("#cscrim").onclick();
  assert.strictEqual(el("#chost"), null);
  assert.strictEqual(el("#fabhost").className, "fabwrap");
  pass("the drawer and the Alerts / Changes windows all step the wheel aside, then give it back");

  /* ---- 4b. the motion: nothing pops, everything travels ----
     There is no browser here to actually run a transition in, so what is
     checked is everything a transition needs and everything the page decides:
     the class each stage turns on, the custom property it moves between, that
     the host outlives its own leaving animation, and that every duration and
     easing really is a --fab-* property in the stylesheet rather than a number
     buried in the JavaScript. */
  const CSS = fs.readFileSync(__dirname + "/index.html", "utf8");
  const FABCSS = CSS.slice(CSS.indexOf("/* ---- the floating selection radial menu"),
                           CSS.indexOf(".toast {"));
  assert.ok(FABCSS.length > 1000, "the wheel's stylesheet was found");
  const cssVal = k => {
    const m = FABCSS.match(new RegExp("\\" + k + "\\s*:\\s*([^;}]+)"));
    return m ? m[1].trim() : null;
  };
  /* every timing the motion uses is declared, and declared here */
  const declared = (FABCSS.match(/--fab-[tdse]-[a-z-]+\s*:/g) || []).map(s => s.split(":")[0].trim());
  const used = (FABCSS.match(/var\(--fab-[tdse]-[a-z-]+/g) || []).map(s => s.slice(4));
  used.forEach(v => assert.ok(declared.indexOf(v) >= 0, v + " is used but never declared"));
  assert.ok(declared.length >= 18, "the whole feel is tunable: " + declared.length + " properties");
  const timing = {
    "--fab-t-arrive": "420ms", "--fab-t-leave": "300ms",
    "--fab-t-glide": "450ms", "--fab-d-home": "200ms",
    "--fab-t-x": "500ms", "--fab-d-x": "80ms",
    "--fab-t-ring-in": "450ms", "--fab-t-ring-out": "550ms", "--fab-d-ring-out": "80ms",
    "--fab-t-opt": "450ms", "--fab-d-opt": "160ms", "--fab-s-opt": "70ms",
    "--fab-t-optback": "320ms", "--fab-s-optback": "45ms",
    "--fab-t-pulse": "250ms",
    "--fab-e-arrive": "cubic-bezier(.34,1.56,.64,1)",
    "--fab-e-spring": "cubic-bezier(.34,1.4,.64,1)"
  };
  Object.keys(timing).forEach(k =>
    assert.strictEqual(cssVal(k), timing[k], k + " is not what the motion asks for"));
  /* the JavaScript's fallbacks are the same numbers, so tuning the stylesheet
     alone can never leave the removal timer behind the animation it waits for */
  assert.strictEqual(fabMs("--fab-t-leave", 300) + "ms", timing["--fab-t-leave"]);
  assert.strictEqual(fabMs("--fab-t-glide", 450) + "ms", timing["--fab-t-glide"]);
  assert.strictEqual(fabMs("--fab-d-home", 200) + "ms", timing["--fab-d-home"]);
  /* nothing is transitioned that the compositor cannot carry on its own */
  (FABCSS.match(/transition:[^;}]+/g) || []).forEach(t =>
    assert.ok(!/(box-shadow|width|height|margin|inset|filter|background|border)/.test(t),
      "only transform and opacity may move: " + t));
  ["\\.fabwrap \\{[^}]*will-change", "\\.fab \\{[^}]*will-change", "\\.fabx \\{[^}]*will-change",
   "\\.fabring \\{[^}]*will-change", "\\.fabopt \\{[^}]*will-change"].forEach(re =>
    assert.ok(new RegExp(re).test(FABCSS.replace(/\s+/g, m => m.indexOf("\n") >= 0 ? " " : m)),
      "no will-change on " + re));
  /* the two stages that are keyframes rather than transitions */
  assert.ok(/@keyframes fabrise \{ from \{ transform:translateY\(120px\); opacity:0 \}/.test(FABCSS),
    "the wheel rises from 120 px under the edge");
  assert.ok(/@keyframes fabpulse \{ 0% \{ transform:scale\(1\) \} 45% \{ transform:scale\(1\.15\) \} 100% \{ transform:scale\(1\) \} \}/
    .test(FABCSS), "and the count beats 1 - 1.15 - 1");
  assert.ok(/animation:fabrise var\(--fab-t-arrive\) var\(--fab-e-arrive\)/.test(FABCSS));
  assert.ok(/\.fabwrap\.leaving \{ transform:translateY\(120px\); opacity:0/.test(FABCSS),
    "and sinks back the same way");
  /* asked for less motion: opacity only, and nowhere left mid-slide */
  const RM = FABCSS.slice(FABCSS.indexOf("prefers-reduced-motion"));
  assert.ok(/animation:none !important/.test(RM) &&
            /transition:opacity 120ms linear !important/.test(RM) &&
            /transition-delay:0s !important/.test(RM), "reduced motion keeps a short fade only");
  assert.ok(/\.fabwrap\.leaving \{ transform:none !important \}/.test(RM));
  pass("every stage of the motion is a --fab-* property in the stylesheet, on transform/opacity only");

  /* what the page itself builds and switches */
  host = el("#fabhost");
  assert.deepStrictEqual(host.kids.filter(x => /fabring/.test(x.className)).map(x => x.className),
    ["fabring in", "fabring out"], "two rings to grow out behind the button");
  assert.strictEqual(el("#fabbtn").className, "fab", "it arrives without a pulse");
  el("#fabbtn").onclick({ stopPropagation() {} });
  assert.strictEqual(el("#fabhost").className, "fabwrap open", "opening is one class");
  el("#fabbtn").onclick({ stopPropagation() {} });

  /* the count changing while it is on screen gives the number a beat */
  set("state.picked = { R0001: 1, R0003: 1, R0005: 1 };");
  renderChips();
  assert.strictEqual(el("#fabbtn").className, "fab pulse", "a new count beats");
  assert.ok(el("#fabbtn").innerHTML.indexOf('<span class="fabn">3</span>') === 0);
  /* and a render that changes nothing leaves the button entirely alone, so an
     open ring is never rebuilt under the finger by a passing chip-bar redraw */
  el("#fabbtn").innerHTML = "";
  renderChips();
  assert.strictEqual(el("#fabbtn").innerHTML, "", "the same count is not written again");
  set("state.picked = { R0001: 1, R0003: 1 };");
  renderChips();
  assert.ok(el("#fabbtn").innerHTML.indexOf('<span class="fabn">2</span>') === 0, "back to two");
  pass("the rings are built, opening is one class, and a changed count beats instead of jumping");

  /* the screen changes shape while the ring is open: it closes first */
  el("#fabbtn").onclick({ stopPropagation() {} });
  assert.strictEqual(el("#fabhost").className, "fabwrap open");
  const wasSig = el("#fabhost").dataset.sig;
  window.innerWidth = 390; window.innerHeight = 844;
  renderChips();
  assert.strictEqual(el("#fabhost").className, "fabwrap", "it closed rather than jumped");
  assert.strictEqual(el("#fabhost").dataset.sig, wasSig, "and the ring was NOT re-laid mid-motion");
  await new Promise(r => setTimeout(r, 760));
  assert.notStrictEqual(el("#fabhost").dataset.sig, wasSig, "the new geometry lands once it is still");
  assert.strictEqual(el("#fabhost").dataset.sig, "alert,export,move,clear@100/146/142");
  window.innerWidth = 1280; window.innerHeight = 800;
  renderChips();
  await new Promise(r => setTimeout(r, 760));
  assert.strictEqual(el("#fabhost").dataset.sig, "alert,export,move,clear@100/150/150");
  pass("turning the phone while it is open closes it with its motion, then re-lays the ring");

  /* leaving: it sinks first, and only then is it taken off the page - and the
     last job unticked with the ring OPEN is the hard case, because every child
     ends a transition of its own before the host has finished sinking and
     bubbles it straight up through the host's listener */
  el("#fabbtn").onclick({ stopPropagation() {} });
  assert.strictEqual(el("#fabhost").className, "fabwrap open");
  set("state.picked = {};");
  renderChips();
  host = el("#fabhost");
  assert.ok(host, "the last job unticked: the wheel is still there...");
  assert.strictEqual(host.className, "fabwrap leaving", "...closed, and on its way out");
  host.fire("transitionend", { target: opt("clear"), propertyName: "transform" });
  assert.strictEqual(el("#fabhost"), host, "an option drawing in is not the wheel leaving");
  host.fire("transitionend", { target: el("#fabbtn"), propertyName: "transform" });
  assert.strictEqual(el("#fabhost"), host, "nor is the button gliding home");
  host.fire("transitionend", { target: host.kids[0], propertyName: "opacity" });
  assert.strictEqual(el("#fabhost"), host, "nor a ring shrinking back");
  host.fire("transitionend", { target: host, propertyName: "width" });
  assert.strictEqual(el("#fabhost"), host, "nor anything the sink does not animate");
  host.fire("transitionend", { target: host, propertyName: "transform" });
  assert.strictEqual(el("#fabhost"), null, "gone the moment the host's own sinking ends");
  /* ...and gone anyway if transitionend never comes (hidden tab, no browser) */
  set("state.picked = { R0001: 1 };");
  renderChips();
  assert.ok(el("#fabhost"), "ticked again: it rises back");
  set("state.picked = {};");
  renderChips();
  assert.strictEqual(el("#fabhost").className, "fabwrap leaving");
  await new Promise(r => setTimeout(r, 420));
  assert.strictEqual(el("#fabhost"), null, "the timer behind the transition takes it away");

  /* ticked and unticked in a hurry: one host, and it never orphans */
  set("state.picked = { R0001: 1 };");
  renderChips();
  const first = el("#fabhost");
  set("state.picked = {};");
  renderChips();
  set("state.picked = { R0001: 1, R0003: 1 };");
  renderChips();
  assert.strictEqual(el("#fabhost"), first, "the very same host, called back mid-leave");
  assert.strictEqual(el("#fabhost").className, "fabwrap", "and no longer leaving");
  await new Promise(r => setTimeout(r, 420));
  assert.strictEqual(el("#fabhost"), first, "the cancelled leave never fires behind it");
  assert.ok(el("#fabbtn").innerHTML.indexOf('<span class="fabn">2</span>') === 0);
  pass("it sinks before it is removed, on transitionend or on the timer, and a fast re-tick calls it back");

  /* Untick all - the selection goes, and so does the wheel */
  el("#fabbtn").onclick({ stopPropagation() {} });
  assert.ok(opt("clear").innerHTML.indexOf('class="fabl">Untick all<') >= 0,
    'the old "Clear" now reads "Untick all"');
  ids = opt("clear").onclick({ stopPropagation() {} });
  assert.deepStrictEqual(ids, ["R0001", "R0003"], "it reports what it cleared");
  assert.deepStrictEqual(state.picked, {});
  assert.strictEqual(el("#fabhost").className, "fabwrap leaving", "the wheel is on its way out");
  host = el("#fabhost");
  host.fire("transitionend", { target: host, propertyName: "transform" });
  assert.strictEqual(el("#fabhost"), null, "nothing ticked, so no wheel");
  pass("Untick all empties the selection and takes the wheel away with it");

  /* ---- 5. neither feature writes or fetches anything ---- */
  assert.strictEqual(FETCHES, 0, "no network");
  assert.deepStrictEqual(LOGS, [], "nothing written to the log sheet");
  /* cw_theme and cw_hidden were already there before either feature; the only
     thing these two add to the browser is which way the Dates fold is turned */
  assert.deepStrictEqual(Object.keys(mem).sort(), ["cw_collapsed", "cw_hidden", "cw_theme"]);
  assert.deepStrictEqual(Object.keys(JSON.parse(mem.cw_collapsed)), ["dates"]);
  assert.strictEqual(JSON.parse(mem.cw_collapsed).dates, 1);
  pass("no network, no workbook write, and one UI preference in localStorage");

  console.log("\n" + n + " checks passed");
  process.exit(0);
})().catch(e => { console.error("FAIL", e); process.exit(1); });
