/* Offline test of the Checkpoints feature.

   REWRITTEN IN PART 2026-09-11, spec docs/specs/2026-09-11-status-list-is-truth.md
   (step 2). Checkpoint status used to be read out of the Excel colour, with a
   count from the `Dashboard Progress` sheet filling in the number, and a
   `PENDING` hold keeping the office's own click alive while the downloaded
   file - about 36 s behind - caught up. Status is now one SharePoint list,
   `Dashboard progress`, one row per JOB|ITEM, read directly and written first.

   So this suite covers: what an item's status IS (the record, and nothing
   else); the write order - record, then fill, then log line - and what happens
   when either half is refused; that no `PENDING` entry is ever made for a
   checkpoint; that a stale download changes nothing; the one-time import of
   today's colours; the safeguard that adopts a cell painted by hand in Excel;
   a missing list; the debounce that turns a run of + taps into one write; the
   unsent-tap queue; and the parser's cp statuses, which are still parsed and
   are still not status.
   Graph is a fake fetch() over an in-memory workbook and an in-memory list -
   nothing leaves the box.
   Run: node test_checkpoints.js                                             */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, addEventListener() {} };
global.performance = { now: () => Date.now() };
global.ExcelJS = ExcelJS;
function stubEl() {
  const e = {
    style: {}, dataset: {}, textContent: "", innerHTML: "", value: "", disabled: false, hidden: false,
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild() {}, remove() {}, addEventListener() {}, removeEventListener() {}, focus() {},
    setSelectionRange() {}, getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => stubEl(), querySelectorAll: () => []
  };
  return e;
}
global.document = {
  documentElement: stubEl(), body: stubEl(), createElement: () => stubEl(), activeElement: null,
  querySelector: () => stubEl(), querySelectorAll: () => [], addEventListener() {}, removeEventListener() {}
};

/* ---------- a tiny Excel behind a fake Graph ---------- */
const BOOK = {};
const CALLS = [];
let FAIL_FILL = 0;                 // 403: refused outright, so no retry sleeps in the test
const kk = (r, c) => r + "|" + c;
const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
function sh(name) { return BOOK[name] || (BOOK[name] = { v: {}, fill: {} }); }
function bounds(s) {
  let maxR = 0, maxC = 0;
  Object.keys(s.v).forEach(k => { const p = k.split("|"); maxR = Math.max(maxR, +p[0]); maxC = Math.max(maxC, +p[1]); });
  return { maxR, maxC };
}
function parseAddr(a) {
  let m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a);
  if (m) return { r1: +m[2], r2: +m[4], c1: colNum(m[1]), c2: colNum(m[3]) };
  m = /^([A-Z]+)(\d+)$/.exec(a);
  if (m) return { r1: +m[2], r2: +m[2], c1: colNum(m[1]), c2: colNum(m[1]) };
  m = /^([A-Z]+):([A-Z]+)$/.exec(a);
  if (m) return { cols: true, c1: colNum(m[1]), c2: colNum(m[2]), r1: 1, r2: 1 };
  throw new Error("bad address " + a);
}
const ok = body => ({ status: 200, body });

/* ---- the record, behind the same fake Graph ----
   `Dashboard progress` lives in the workbook's own site. Served rather than
   stubbed: the order of the writes and the exact body of each is most of what
   this suite is about. */
const WSITE = "workbook-site-0001";
const CPLIST_ID = "list-dashboard-progress";
const CPBASE = "/sites/" + WSITE + "/lists/" + CPLIST_ID + "/items";
let CPSRV = [];                     // the record, as the server holds it
let CPNEXT = 500;
let CPLIST_THERE = true;            // false: the owner has not made the list yet
let FAIL_LIST = 0;                  // n list writes to refuse with a 403
const LISTREQ = [];                 // every request this suite sends to that list
/* The ORDER of the three things a click does, as they really go on the wire:
   "record" (the list), "fill" (the Production cell) and "log" (Dashboard Log).
   That order is the whole of the 2026-09-11 change, so it is recorded here
   rather than inferred from two separate arrays afterwards. */
const SEQ = [];
const cpSrvRow = (job, item) => CPSRV.filter(x =>
  String(x.fields.Title || "").toUpperCase() === (job + "|" + item).toUpperCase())[0] || null;

function routeCpList(method, path, body) {
  if (path === "/sites/" + WSITE + "/lists?$select=id,displayName")
    return ok({ value: CPLIST_THERE ? [{ id: CPLIST_ID, displayName: "Dashboard progress" }]
                                    : [{ id: "other", displayName: "Site Assets" }] });
  if (path.indexOf(CPBASE) !== 0) return null;
  LISTREQ.push({ method: method, path: path, body: body });
  if (method !== "GET") SEQ.push("record");
  const rest = path.slice(CPBASE.length);
  if (method === "GET" && rest.indexOf("/delta") === 0)
    return ok({ value: CPSRV.map(x => ({ id: x.id, fields: x.fields })),
                "@odata.deltaLink": "https://graph.microsoft.com/v1.0" + CPBASE + "/delta?$skiptoken=1" });
  if (method === "GET" && (rest === "" || rest.charAt(0) === "?"))
    return ok({ value: CPSRV.map(x => ({ id: x.id, fields: x.fields })) });
  if (method === "POST" && rest === "") {
    if (FAIL_LIST) { FAIL_LIST--; return { status: 403, body: { error: { code: "AccessDenied" } } }; }
    const id = String(CPNEXT++);
    CPSRV.push({ id: id, fields: (body && body.fields) || {} });
    return ok({ id: id });
  }
  const mfd = /^\/([^/?]+)\/fields$/.exec(rest);
  if (mfd && method === "PATCH") {
    if (FAIL_LIST) { FAIL_LIST--; return { status: 403, body: { error: { code: "AccessDenied" } } }; }
    const it = CPSRV.filter(x => x.id === mfd[1])[0];
    if (!it) return { status: 404, body: { error: { code: "itemNotFound" } } };
    Object.keys(body || {}).forEach(k => { it.fields[k] = body[k]; });
    return ok({ id: it.id });
  }
  const mdl = /^\/([^/?]+)$/.exec(rest);
  if (mdl && method === "DELETE") { CPSRV = CPSRV.filter(x => x.id !== mdl[1]); return ok({}); }
  return { status: 404, body: { error: "no route " + method + " " + path } };
}

function route(method, path, body) {
  const cp = routeCpList(method, path, body);
  if (cp) return cp;
  if (path === "/x/workbook/createSession") return ok({ id: "S1" });
  if (path === "/x/workbook/worksheets") return ok({ value: Object.keys(BOOK).map(n => ({ name: n })) });
  if (path === "/x/workbook/worksheets/add") { sh(body.name); CALLS.push({ method, sheet: body.name, kind: "addSheet" }); return ok({}); }
  const ms = /\/worksheets\('([^']+)'\)(.*)$/.exec(path);
  if (!ms) return { status: 404, body: { error: "no route " + path } };
  const name = ms[1], rest = ms[2], s = sh(name);
  if (rest.indexOf("/usedRange") === 0) {
    const b = bounds(s), values = [];
    for (let r = 1; r <= b.maxR; r++) {
      const line = [];
      for (let c = 1; c <= Math.max(b.maxC, 1); c++) line.push(s.v[kk(r, c)] == null ? "" : s.v[kk(r, c)]);
      values.push(line);
    }
    CALLS.push({ method, sheet: name, kind: "usedRange" });
    return ok({ values: values, rowCount: b.maxR });
  }
  const mr = /^\/range\(address='([^']+)'\)(.*)$/.exec(rest);
  if (!mr) return { status: 404, body: { error: "no route " + path } };
  const A = parseAddr(mr[1]), tail = mr[2];
  /* one cell's fill, read live through the Excel API - the safeguard's
     confirming read (spec section 4a). The download can be 36 s behind; this
     never is, which is the whole reason it is the thing that decides. */
  if (method === "GET" && tail === "/format/fill") {
    CALLS.push({ method, sheet: name, kind: "readFill", addr: mr[1] });
    return ok({ color: s.fill[kk(A.r1, A.c1)] || "" });
  }
  if (method === "GET") {
    const values = [];
    for (let r = A.r1; r <= A.r2; r++) {
      const line = [];
      for (let c = A.c1; c <= A.c2; c++) line.push(s.v[kk(r, c)] == null ? "" : s.v[kk(r, c)]);
      values.push(line);
    }
    CALLS.push({ method, sheet: name, kind: "readRange", addr: mr[1] });
    return ok({ values: values });
  }
  if (method === "PATCH" && tail === "/format/fill") {
    if (FAIL_FILL && name === "Production") {
      FAIL_FILL--;
      CALLS.push({ method, sheet: name, kind: "fillRefused", addr: mr[1] });
      return { status: 403, body: { error: { code: "AccessDenied" } } };
    }
    for (let r = A.r1; r <= A.r2; r++) for (let c = A.c1; c <= A.c2; c++) s.fill[kk(r, c)] = body.color;
    CALLS.push({ method, sheet: name, kind: "fill", addr: mr[1], color: body.color });
    if (name === "Production") SEQ.push("fill");
    return ok({});
  }
  if (method === "PATCH" && tail === "") {
    (body.values || []).forEach((line, ri) => line.forEach((x, ci) => { s.v[kk(A.r1 + ri, A.c1 + ci)] = x; }));
    /* a PATCH with no values only sets number formats - not a row of data */
    CALLS.push({ method, sheet: name, kind: body.values ? "values" : "numberFormat", addr: mr[1], values: body.values });
    if (body.values && name === "Dashboard Log" && mr[1] !== "A1:F1") SEQ.push("log");
    return ok({});
  }
  if (method === "PATCH") { CALLS.push({ method, sheet: name, kind: "format", addr: mr[1] }); return ok({}); }
  return { status: 404, body: { error: "no route " + method + " " + path } };
}
let BATCHES = 0;
global.fetch = async (url, init) => {
  const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  let res;
  if (path === "/$batch") {
    BATCHES++;
    res = ok({ responses: body.requests.map(q => Object.assign({ id: q.id }, route(q.method, q.url, q.body))) });
  } else res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status, text: async () => JSON.stringify(res.body), arrayBuffer: async () => new ArrayBuffer(0) };
};

/* ---------- load the app the way the page does ---------- */
const run = f => vm.runInThisContext(fs.readFileSync(__dirname + "/" + f, "utf8"), { filename: f });
run("parser.js");
run("graph.js");
global.CW = window.CW;
CW._setToken(() => "t");
CW._setFile({ base: "/x/workbook", content: "/x/content", meta: "/x", siteId: WSITE });
run("checkpoints.js");
global.CP = window.CP;
/* station-core.js is pure logic and app.js reads two things out of it here:
   ST.COLOUR_TYPES, which the repaint keeps off until step 3, and
   ST.officeComplete, which is what the floor's lock is derived from. */
run("station-core.js");
global.ST = window.ST;
run("app.js");

const TOASTS = [];
global.toast = (m, err) => TOASTS.push({ m: String(m), err: !!err });   // no DOM: just remember them

const settle = ms => new Promise(r => setTimeout(r, ms == null ? 40 : ms));
const fills = () => CALLS.filter(c => c.kind === "fill" && c.sheet === "Production");
/* A1:F1 is the header row written once when a sheet is first created. */
const progressWrites = () => CALLS.filter(c => c.kind === "values" && c.sheet === "Dashboard Progress" && c.addr !== "A1:F1");
const logWrites = () => CALLS.filter(c => c.kind === "values" && c.sheet === "Dashboard Log" && c.addr !== "A1:F1");
const reset = () => { CALLS.length = 0; BATCHES = 0; TOASTS.length = 0; LISTREQ.length = 0; SEQ.length = 0; };
/** every write this suite sent to the record */
const listWrites = () => LISTREQ.filter(r => r.method !== "GET");
/** the record, as the page reads it */
const rec = (job, item) => cpRow(job, item);
/** Put the record where a test wants it - the server's copy and the page's,
    together - and tell the page the list answered. */
function setRecord(rows) {
  CPSRV = [];
  const items = (rows || []).map(r => {
    const id = String(CPNEXT++);
    const f = cpRowFields(r.job, r.item, r.done, r.total, r.status, r.who || "the admin",
                          r.when || "2026-09-01T09:00:00Z", r.source || "office");
    CPSRV.push({ id: id, fields: f });
    return { id: id, fields: JSON.parse(JSON.stringify(f)) };
  });
  global.__cpitems = items;
  vm.runInThisContext("CP_ITEMS = __cpitems; cpListRebuild(); CP_LIST_OK = true; " +
    "CP_LIST_WHY = ''; CP_LIST_CONSENT = false; STATION_FEEDS.progress.token = null;");
  /* the switch-on window is over unless a test says otherwise: an item with no
     row means nothing done, which is the state the dashboard lives in from the
     first day onwards (review finding M1) */
  settled(true);
}
/** Has the one-time import drained? Every test but the switch-on ones runs on
    the far side of it. */
function settled(yes) {
  vm.runInThisContext("CP_IMPORTED = " + (yes ? "'2026-09-11T09:00:00Z'" : "''") + "; cpImportCheck();");
}
/** The doors' own marker, which is a separate question from the general one:
    the doors shipped three days after the general import had drained. */
function doorsSettled(yes) {
  vm.runInThisContext("CP_DOORS_IMPORTED = " + (yes ? "'2026-09-14T09:00:00Z'" : "''") + ";");
  if (yes) localStorage.setItem("cw_cpimported_doors", "2026-09-14T09:00:00Z");
  else localStorage.removeItem("cw_cpimported_doors");
}
/** ... and the same for the colours this dashboard remembers having painted. */
function setPainted(map) {
  global.__painted = map || {};
  vm.runInThisContext("PAINTED = __painted; savePainted();");
}
const queued = () => JSON.parse(localStorage.getItem("cw_cpqueue") || "[]");
/** the stored count for one (job, item), read back out of the fake sheet */
function storedCount(job, item) {
  const s = BOOK["Dashboard Progress"]; if (!s) return null;
  for (let r = 2; r <= 200; r++)
    if (s.v[kk(r, 1)] === job && s.v[kk(r, 2)] === item) return s.v[kk(r, 3)];
  return null;
}

/* ---------- the job under test ---------- */
const PMAP = { qty: { wnd: 13, drs: 14 }, glass: { tg: 51, tuff: 52 },
               prod: { "7000 casement": { f: 20, s: 21, t: 22 } }, prodOrder: ["7000 casement"] };
const mkJob = (cp, extra) => Object.assign({
  id: "R0001", cust: "Ann", area: "Cork", wnd: 10, drs: 2, glass: { tg: 25, tuff: 11 },
  prods: [{ n: "7000 casement", f: 10, s: 8, t: 0, st: [] }], notes: [], sheets: ["Production"], src: {},
  doors: [], dates: {}, cat: "active", blk: 3, seq: 1, stage: "floor", done: 0, urg: 0,
  cp: JSON.parse(JSON.stringify(cp))
}, extra || {});

/* ---- the item the machinery tests tick -------------------------------------
   It used to be "win", the Windows quantity in column M. Since 2026-09-14
   (docs/specs/2026-09-11-doors-and-window-types.md) Windows is DERIVED from
   the job's window types and cannot be ticked at all, so the generic item is
   one of those types instead: the 7000 casement's frames, ten of them, in
   column T. Nothing about the machinery under test depends on which item it
   is - the debounce, the queue, the record-then-fill-then-log order, the
   safeguard and the repaint are the same for every counted item. */
const IT = "prod:7000 casement:f", ITCOL = 20, ITADDR = "T7";
/** A fixture whose generic item carries `word` as the colour in its cell. */
const mkIt = (word, extra) => mkJob({ win: "", drs: "", glass: {},
                                      prod: { "7000 casement": { f: word } } }, extra);
/** Move that colour on a fixture already made, the way a download would. */
const setFileCp = (j, word) => { (j.cp.prod["7000 casement"] = j.cp.prod["7000 casement"] || {}).f = word; };
/** The colours this browser remembers painting, for the generic item. */
const paintedIt = word => { const m = {}; m.R0001 = {}; m.R0001[IT] = word; return m; };
/* put one job on the fake Production sheet at row 7 and in the app's list */
function useJob(job, map) {
  BOOK["Production"] = { v: {}, fill: {} };
  BOOK["Production"].v[kk(7, 3)] = job.id;
  /* where the parser found it: row 7, which is where this fake sheet puts it.
     The safeguard and the repaint both read a cell of that row. */
  if (!(job.src && job.src.Production)) job.src = Object.assign({}, job.src, { Production: 7 });
  global.__j = job; global.__m = map === undefined ? PMAP : map;
  vm.runInThisContext("ALL = [__j]; PRODMAP = __m; CHANGES = []; state.sel = null; " +
    "PENDING = {}; savePending();");
  /* and a clean slate for the colours this browser remembers painting: a
     fixture is a different job from the one before it, and a leftover PAINTED
     would make the derived Windows/Doors cells look as though they had moved */
  setPainted({});
  return byId(job.id);
}
/** The record as the fixture's Excel colours would have it - what the one-time
    import writes on the first load after the switch-over. Nothing else in this
    suite reads a colour as status, and neither does the app. */
function recordFromColours(job, counts) {
  setRecord(cpItems(job).map(it => {
    const st = cpFileStatus(job, it.key);
    if (st !== "done" && st !== "process") return null;
    const n = (counts || {})[it.key];
    return { job: job.id, item: it.key, total: it.total, status: st, source: "import",
             done: st === "done" ? it.total : (n > 0 && n < it.total ? n : 0) };
  }).filter(Boolean));
}
/* age a held item so it looks older than the 180 s window */
const ageHold = (id, key, ms) => vm.runInThisContext(
  "PENDING['" + id + "'].t['" + key + "'] -= " + ms + "; PENDING['" + id + "'].at -= " + ms);

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ---- 1. THE RECORD DECIDES, and the Excel colour does not ---- */
  let j = mkJob({ win: "done", drs: "done", glass: { tg: "process", tuff: "done" },
                  prod: { "7000 casement": { f: "process", s: "done", t: "" } } });
  setRecord([
    { job: "R0001", item: IT, done: 10, total: 10, status: "done" },
    { job: "R0001", item: "glass:tg", done: 6, total: 25, status: "process" },
    { job: "R0001", item: "glass:tuff", done: 0, total: 11, status: "process" }
  ]);
  assert.deepStrictEqual(itemState(j, IT), { done: 10, total: 10, status: "done" });
  pass("a row saying done is done, whatever the cell is coloured");
  assert.deepStrictEqual(itemState(j, "glass:tg"), { done: 6, total: 25, status: "process" });
  pass("a row part way there carries its own count: 6 of 25, in process");
  assert.deepStrictEqual(itemState(j, "glass:tuff"), { done: null, total: 11, status: "process" });
  pass("in process with no count on the row: in progress, count unknown");
  /* the one that matters: a GOLD cell whose row says nothing. Before
     2026-09-11 it read as "done" off the colour. */
  assert.deepStrictEqual(itemState(j, "prod:7000 casement:s"), { done: 0, total: 8, status: "" });
  assert.strictEqual(cpFileStatus(j, "prod:7000 casement:s"), "done", "even though the file plainly says gold");
  pass("a gold cell the record says nothing about is NOT done: the colour is not status");
  assert.strictEqual(itemState(j, "prod:7000 casement:t"), null, "an item with no total does not exist");
  assert.strictEqual(cpStatus(j, IT), "done");
  assert.deepStrictEqual(cpItems(j).map(x => x.key),
    ["win", "drs", "glass:tg", "glass:tuff", IT, "prod:7000 casement:s"]);
  pass("only items with a total exist, in drawer order");

  /* ---- 1b. WINDOWS AND DOORS ARE DERIVED (2026-09-14, the doors spec) ------
     M and N stopped being ticks of their own. Windows says what the job's
     window types say and Doors says what its doors say, and neither reads a
     row of its own - so a gold M with nothing ticked under it is not done. */
  assert.deepStrictEqual(itemState(j, "win"), { done: null, total: 10, status: "process" },
    "one window type part way, one not started: Windows is in fabrication");
  assert.strictEqual(cpFileStatus(j, "win"), "done", "even though M is gold in the file");
  assert.deepStrictEqual(itemState(j, "drs"), { done: 0, total: 2, status: "" },
    "no door codes on the row at all: Doors is blank, whatever N is coloured");
  assert.strictEqual(cpFileStatus(j, "drs"), "done", "and N is gold in the file");
  pass("Windows and Doors are the aggregate of what is under them, not a row of their own");

  /* the same job with its five DOORS DONE cells filled in */
  const dj = mkJob({ win: "", drs: "", glass: {}, prod: {} },
    { wnd: 0, glass: {}, prods: [],
      drs: 3, doors: [{ slot: 1, code: "CD", status: "done" },
                      { slot: 2, code: "1 DOOR", status: "process" },
                      { slot: 4, code: "SFCD", status: "" }] });
  setRecord([
    { job: "R0001", item: "door:1", done: 1, total: 1, status: "done" },
    { job: "R0001", item: "door:2", done: 0, total: 1, status: "process" }
  ]);
  assert.deepStrictEqual(cpItems(dj).map(x => [x.key, x.label, x.total, x.group]),
    [["drs", "Doors", 3, "drs"], ["door:1", "CD", 1, "door"],
     ["door:2", "1 DOOR", 1, "door"], ["door:4", "SFCD", 1, "door"]]);
  pass("one item per door, total 1, its label the code the office typed, in its own group");
  assert.deepStrictEqual(itemState(dj, "door:1"), { done: 1, total: 1, status: "done" });
  assert.deepStrictEqual(itemState(dj, "door:2"), { done: null, total: 1, status: "process" });
  assert.deepStrictEqual(itemState(dj, "door:4"), { done: 0, total: 1, status: "" });
  assert.strictEqual(itemState(dj, "door:3"), null, "a slot with no text in it is not a door");
  assert.deepStrictEqual(itemState(dj, "drs"), { done: null, total: 3, status: "process" },
    "one done, one in fabrication, one not: the Doors line is in fabrication");
  /* every door done makes the line done; clearing one takes it back */
  setRecord([{ job: "R0001", item: "door:1", done: 1, total: 1, status: "done" },
             { job: "R0001", item: "door:2", done: 1, total: 1, status: "done" },
             { job: "R0001", item: "door:4", done: 1, total: 1, status: "done" }]);
  assert.deepStrictEqual(itemState(dj, "drs"), { done: 3, total: 3, status: "done" });
  setRecord([]);
  assert.deepStrictEqual(itemState(dj, "drs"), { done: 0, total: 3, status: "" },
    "and with nothing recorded under it, the line is blank again");
  pass("the Doors line follows its doors: all done, part way, nothing");
  assert.strictEqual(cpColumn("door:2", { doors: { 1: 72, 2: 73, 3: 74, 4: 75, 5: 76 } }), 73,
    "a door's cell is its own DOORS DONE column and nothing else");
  assert.strictEqual(cpLabel("door:2"), "Door 2");

  /* a row for a job or an item the sheet does not have is simply not asked
     about, and a Status the list does not know reads as blank */
  setRecord([{ job: "R0001", item: IT, done: 4, total: 10, status: "rubbish" }]);
  assert.deepStrictEqual(itemState(j, IT), { done: 0, total: 10, status: "" },
    "a Status nobody recognises is not a status");
  /* Total comes from the SHEET, never from the row: the sheet is still the
     master for how much work there is */
  setRecord([{ job: "R0001", item: IT, done: 99, total: 99, status: "done" }]);
  assert.deepStrictEqual(itemState(j, IT), { done: 10, total: 10, status: "done" },
    "the sheet's own quantity is the total, and done is clamped to it");
  pass("the record carries the status and the count; the sheet still carries the quantity");

  /* ---- 2. THERE IS NO HOLD, and a stale download changes nothing ----
     REPLACES eleven assertions about `cp` PENDING holds: that a held count
     beat Excel and the stored count; that holds expired one item at a time;
     that an expired hold the file still disagreed with was KEPT and a re-read
     demanded; and that a hold was let go the moment a fresh parse agreed.
     Every one of them existed to keep believing the office over a copy of the
     workbook ~36 s behind. Nothing in the checkpoint path reads that copy any
     more, so there is nothing to hold, nothing to expire and nothing to
     reconcile - which is the whole of the change and is what is proved here
     instead. */
  const staleGold = () => mkJob({ win: "done", drs: "", glass: { tg: "done", tuff: "done" },
                                  prod: { "7000 casement": { f: "done", s: "done" } } });
  useJob(staleGold());
  setRecord([]);                                          // the office has un-ticked everything
  for (let i = 0; i < 8; i++) {
    vm.runInThisContext("ALL = applyPending([__j], true);");   // the 36 s-behind file, again and again
    assert.deepStrictEqual(itemState(byId("R0001"), IT), { done: 0, total: 10, status: "" },
      "sample " + (i + 1) + ": still white");
    assert.deepStrictEqual(itemState(byId("R0001"), "glass:tg"), { done: 0, total: 25, status: "" },
      "sample " + (i + 1) + ": glass still white");
  }
  assert.strictEqual(Object.keys(vm.runInThisContext("PENDING")).length, 0,
    "and not one PENDING entry was made for any of it");
  pass("eight stale gold downloads in a row leave every checkpoint white, holding nothing");

  /* the same job seen by a second dashboard, or by this one after a refresh:
     the record is read back from SharePoint, so it is in the same place */
  vm.runInThisContext("CP_ITEMS = []; cpRowsSet({});");
  setRecord([{ job: "R0001", item: IT, done: 0, total: 10, status: "" }]);
  vm.runInThisContext("ALL = applyPending([__j], true);");
  assert.deepStrictEqual(itemState(byId("R0001"), IT), { done: 0, total: 10, status: "" },
    "a fresh page reads the record, not the colour, and is in exactly the same place");
  pass("a refresh changes nothing: the record does not live in this browser");

  /* ---- 3. five rapid taps make one write: record, then fill, then log ---- */
  useJob(mkIt("process"));
  setRecord([{ job: "R0001", item: IT, done: 1, total: 10, status: "process" }]);
  reset();
  for (let i = 0; i < 5; i++) {
    const cur = itemState(byId("R0001"), IT);
    setItemProgress(byId("R0001"), IT, (cur.done == null ? 0 : cur.done) + 1);
  }
  assert.deepStrictEqual(itemState(byId("R0001"), IT), { done: 6, total: 10, status: "process" },
    "the screen is already at 6 before anything is written");
  assert.strictEqual(CALLS.length, 0, "nothing is written while the taps are still coming");
  assert.strictEqual(queued().length, 1, "the unsent tap is remembered in case the tab closes");
  await settle(1400);
  /* two fills: this item's own cell, and the Windows cell above it, which
     since 2026-09-14 is a copy of the aggregate of the job's window types */
  assert.strictEqual(fills().length, 2, "the item's own fill, and the derived Windows cell");
  assert.strictEqual(fills()[0].addr, ITADDR, "written on the job's own row, in this item's column only");
  assert.strictEqual(fills()[0].color, "#FFFF00", "6 of 10 is yellow");
  assert.deepStrictEqual([fills()[1].addr, fills()[1].color], ["M7", "#FFFF00"],
    "and M goes yellow because a window type is part way through");
  assert.strictEqual(progressWrites().length, 0,
    "the Dashboard Progress SHEET is never written again by anything");
  const wrote3 = listWrites();
  assert.strictEqual(wrote3.length, 1, "exactly one write to the record");
  assert.strictEqual(wrote3[0].method, "PATCH", "a row that already exists is patched by its id");
  assert.deepStrictEqual([wrote3[0].body.Job, wrote3[0].body.Item, wrote3[0].body.Done,
                          wrote3[0].body.Total, wrote3[0].body.Status, wrote3[0].body.Source],
    ["R0001", IT, 6, 10, "process", "office"], "with the job, the item, the count and who set it");
  assert.ok(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/.test(wrote3[0].body.When),
    "and a full ISO stamp to the second: " + wrote3[0].body.When);
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Done, 6, "the record really says six");
  assert.strictEqual(logWrites().length, 1, "exactly one log line");
  assert.deepStrictEqual(logWrites()[0].values[0].slice(2, 6),
    ["R0001", "7000 CASEMENT frames", "1 of 10", "6 of 10 · Windows cell painted yellow"],
    "the from value is what it was before the burst of taps, and the ONE line " +
    "also says what happened to the Windows cell above it (amendment 3)");
  assert(!CALLS.some(c => c.sheet === "Production" && c.kind !== "fill" && c.kind !== "readRange"),
    "the Production sheet gets fills and a row lookup, nothing else");
  assert.strictEqual(queued().length, 0, "the queue is empty once the write has gone");
  assert.strictEqual(Object.keys(vm.runInThisContext("PENDING")).length, 0,
    "and no PENDING entry was ever made for the tick");
  /* THE ORDER, which is the whole change: the record is written first, and
     the sheet is coloured from it. */
  assert.deepStrictEqual(SEQ, ["record", "fill", "fill", "log"],
    "record first, then the colours, then the log line: " + JSON.stringify(SEQ));
  assert.strictEqual(SEQ[0], "record", "nothing has happened until the record says so");
  pass("5 rapid + taps: one record write, the item's fill and its Windows copy, one log line, and no hold");

  reset();
  setItemProgress(byId("R0001"), IT, 7);
  await settle(1400);
  assert.strictEqual(fills().length, 1, "a tap after the window is its own write");
  assert.strictEqual(listWrites().length, 1);
  assert.deepStrictEqual(logWrites()[0].values[0].slice(3, 6), ["7000 CASEMENT frames", "6 of 10", "7 of 10"]);
  pass("a tap after the debounce window makes a second write, counting from 6");

  /* ---- 3b. a refused RECORD write is a click that did not take ---- */
  reset(); FAIL_LIST = 1;
  setItemProgress(byId("R0001"), IT, 9);
  assert.strictEqual(itemState(byId("R0001"), IT).done, 9, "the row moves at the click");
  await settle(1400);
  assert.strictEqual(fills().length, 0, "nothing was painted: the record refused, so nothing happened");
  assert.strictEqual(logWrites().length, 0, "and nothing is logged for a click that did not take");
  assert(TOASTS.some(t => t.err), "the person is told");
  assert.strictEqual(itemState(byId("R0001"), IT).done, 7, "the screen goes back to seven");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Done, 7, "and the record still says seven");
  assert.strictEqual(queued().length, 0, "a failed write is not owed for ever");
  pass("record refused: nothing is painted, nothing is logged, the row goes back");

  /* ---- 3c. a refused FILL is a colour that is late, not a click undone ----
     The direction this change deliberately chose (spec section 7): the office
     decided, the record says so, and Excel is behind. Putting the record back
     would be the second lie. */
  reset(); FAIL_FILL = 1;
  assert.strictEqual(paintedOf("R0001", IT), "process",
    "the last colour this dashboard really painted into that cell was yellow");
  setItemProgress(byId("R0001"), IT, 10);          // 10 of 10 would be gold, if it landed
  await settle(1400);
  assert.strictEqual(fills().length, 0, "the fill was refused");
  assert.strictEqual(listWrites().length, 1, "but the record was written, and stands");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Done, 10, "the office's decision is on the record");
  assert.deepStrictEqual(itemState(byId("R0001"), IT), { done: 10, total: 10, status: "done" },
    "and on the screen");
  assert.strictEqual(logWrites().length, 0, "nothing is logged for a colour that did not land");
  assert(TOASTS.some(t => t.err && /colour could not be written/.test(t.m)),
    "the person is told the record took it and the sheet did not");
  assert.strictEqual(paintedOf("R0001", IT), "process",
    "and the dashboard does not claim to have painted a cell it did not paint");
  pass("fill refused: the record stands, Excel is behind, and PAINTED is not fooled");

  /* ---- 4. Dashboard Progress upsert and its queue ---- */
  CP.cpSetProgress({});
  reset();
  const r1 = await CW.saveProgress("R0001", IT, 7, 10, "abin");
  const r2 = await CW.saveProgress("R0001", IT, 8, 10, "abin");
  assert.strictEqual(r1, r2, "the same (Job, Item) is updated in place");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(r2, 3)], "8");
  const r3 = await CW.saveProgress("R0001", "drs", 1, 2, "abin");
  assert.strictEqual(r3, r2 + 1, "a new (Job, Item) pair appends");
  const both = await Promise.all([CW.saveProgress("R0002", IT, 1, 4, "abin"),
                                  CW.saveProgress("R0003", IT, 2, 6, "abin")]);
  assert.notStrictEqual(both[0], both[1], "two writes at the same moment take two different rows");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(both[0], 1)], "R0002");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(both[1], 1)], "R0003");
  pass("saveProgress upserts, and two concurrent calls cannot take the same row");
  const nb = BATCHES;
  const wrote = await CW.saveProgressMany("R0001", [
    { item: IT, done: 10, total: 10 }, { item: "drs", done: 2, total: 2 },
    { item: "glass:tg", done: 25, total: 25 }], "abin");
  assert.strictEqual(wrote, 3);
  assert.strictEqual(BATCHES - nb, 1, "saveProgressMany sends one batch");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(r2, 3)], "10", "win updated in its existing row");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(r3, 3)], "2", "drs updated in its existing row");
  pass("saveProgressMany writes every row in one batch");

  /* ---- 5. the colour of a count, and what is not a count ---- */
  assert.strictEqual(cpColour(0, 10), "#FFFFFF");
  assert.strictEqual(cpColour(1, 10), "#FFFF00");
  assert.strictEqual(cpColour(9, 10), "#FFFF00");
  assert.strictEqual(cpColour(10, 10), "#FFE699");
  assert.strictEqual(cpClamp(-4, 10), 0);
  assert.strictEqual(cpClamp(99, 10), 10);
  assert.strictEqual(cpClamp("4.6", 10), 5);
  assert.strictEqual(cpClamp("", 10), null, "a blank box is not a zero");
  assert.strictEqual(cpClamp("   ", 10), null);
  assert.strictEqual(cpClamp("rubbish", 10), null);
  pass("0 white, part yellow, all gold; blank or unreadable is not a count");

  useJob(mkIt("process"));
  setRecord([{ job: "R0001", item: IT, done: 4, total: 10, status: "process" }]);
  reset();
  setItemProgress(byId("R0001"), IT, "");
  setItemProgress(byId("R0001"), IT, "abc");
  await settle(1400);
  assert.strictEqual(CALLS.length, 0, "a blank or unreadable box writes nothing");
  assert.strictEqual(listWrites().length, 0, "not to the record either");
  assert.strictEqual(itemState(byId("R0001"), IT).done, 4, "and changes nothing");
  pass("clearing the number box writes nothing at all");

  /* ---- 5b. an item that is not on the sheet ---- */
  useJob(mkIt("process"), null);   // no PRODMAP
  setRecord([]);
  reset();
  setItemProgress(byId("R0001"), IT, 5);
  await settle(900);
  assert.strictEqual(CALLS.length, 0, "nothing is written when the column is unknown");
  assert.strictEqual(listWrites().length, 0, "and nothing is put on the record either");
  assert(TOASTS.some(t => t.err && /not a column/.test(t.m)), "and it says so");
  pass("no column on the Production sheet: no write, a plain message");

  /* ---- 5b2. and if the list is not there, nothing is written anywhere ----
     CLAUDE.md rule 3: no list is created by code. A missing one is a plain
     message and checkpoints are read-only. */
  useJob(mkIt("process"));
  setRecord([{ job: "R0001", item: IT, done: 3, total: 10, status: "process" }]);
  vm.runInThisContext("CP_LIST_OK = null; CP_LIST_WHY = ''; CP_ITEMS = null; cpRowsSet({}); " +
                      "cpListWarned = false; CW._resetListIds();");
  localStorage.removeItem("cw_listids");        // a list id is never cached when it is missing
  CPLIST_THERE = false;
  reset();
  await readCpList();
  assert.strictEqual(vm.runInThisContext("CP_LIST_OK"), false, "the list is not there");
  assert.ok(/Dashboard progress/.test(vm.runInThisContext("CP_LIST_WHY")),
    "and the reason names it: " + vm.runInThisContext("CP_LIST_WHY"));
  assert.ok(/not in SharePoint/.test(vm.runInThisContext("CP_LIST_WHY")));
  const shut = cpSectionHtml(byId("R0001"), true);
  assert.ok(shut.indexOf("Dashboard progress") >= 0, "the drawer says so, in the section itself");
  const shutControls = (shut.match(/<button/g) || []).length + (shut.match(/<input/g) || []).length;
  assert.strictEqual(shutControls, (shut.match(/ disabled/g) || []).length,
    "and every control on it is dead");
  reset();
  setItemProgress(byId("R0001"), IT, 5);
  await setGroupDone(byId("R0001"), "glass", true);
  await settle(1400);
  assert.strictEqual(CALLS.length, 0, "no workbook write of any kind");
  assert.strictEqual(listWrites().length, 0, "no list write of any kind");
  assert.strictEqual(queued().length, 0, "and nothing owed for later");
  assert.strictEqual(Object.keys(vm.runInThisContext("PENDING")).length, 0, "and nothing held");
  assert(TOASTS.some(t => t.err && /Dashboard progress/.test(t.m)), "the person is told which list");
  /* and the rest of the screen does not say it silently: with the record
     unreadable every job draws as though nothing were ticked, so the footer
     says so whether or not a drawer is open */
  const foot = stubEl(), wrap = stubEl();
  const realQF = document.querySelector;
  document.querySelector = sel => sel === "#stationfeed" ? foot
    : sel === "#stationfeedwrap" ? wrap : stubEl();
  setStationFoot();
  document.querySelector = realQF;
  assert.strictEqual(foot.textContent, "checkpoints read-only", "the footer says it too");
  assert.ok(/Dashboard progress/.test(foot.title), "and names the list in its tooltip: " + foot.title);
  CPLIST_THERE = true;
  vm.runInThisContext("CW._resetListIds();");
  localStorage.removeItem("cw_listids");
  pass("no list: checkpoints are read-only, the message names it, and nothing is written anywhere");

  /* ---- 5c. a whole group in one go ---- */
  useJob(mkIt(""));
  setRecord([]);
  reset();
  await setGroupDone(byId("R0001"), "glass", true);
  await settle(60);
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color).sort(),
    ["AY7 #FFE699", "AZ7 #FFE699"], "both glass cells, gold, on row 7 only");
  /* CHANGED 2026-09-11: one row per item, and SharePoint has no batch of list
     writes - so they go one after the other, and then ONE batch of fills. */
  const g5 = listWrites();
  assert.strictEqual(g5.length, 2, "one write to the record per item");
  assert.ok(g5.every(w => w.method === "POST"), "both rows are new, so both are created");
  assert.deepStrictEqual(g5.map(w => w.body.fields.Item).sort(), ["glass:tg", "glass:tuff"]);
  assert.deepStrictEqual(g5.map(w => w.body.fields.Status), ["done", "done"]);
  assert.strictEqual(progressWrites().length, 0, "and the sheet of that name is not written");
  assert.strictEqual(BATCHES, 1, "one batch, and it is the batch of fills");
  assert.strictEqual(logWrites().length, 1, "one log line for the group");
  assert.deepStrictEqual(logWrites()[0].values[0].slice(3, 6), ["Glass: all done", "", "TG 25, TUFF 11"]);
  assert.deepStrictEqual(SEQ, ["record", "record", "fill", "fill", "log"],
    "every row on the record before a single cell is coloured: " + JSON.stringify(SEQ));
  assert.strictEqual(Object.keys(vm.runInThisContext("PENDING")).length, 0, "and no hold for any of it");
  pass("All glass done: a row per item first, then one batch of fills, then one log line");

  /* ---- 5d. a job already marked ready to deliver ---- */
  const gold = mkJob({ win: "done", drs: "done", glass: { tg: "done", tuff: "done" },
                       prod: { "7000 casement": { f: "done", s: "done" } } }, { done: 1 });
  useJob(gold);
  setRecord([]);
  reset();
  const html = cpSectionHtml(byId("R0001"), true);
  assert(/Marked ready to deliver/.test(html), "it says why it cannot be edited");
  assert(!/data-cpgrp/.test(html), "no group buttons");
  assert(!/>Clear</.test(html), "nothing invites clearing a finished row");
  const controls = (html.match(/<button/g) || []).length + (html.match(/<input/g) || []).length;
  assert.strictEqual(controls, (html.match(/ disabled/g) || []).length, "every control is disabled");
  setItemProgress(byId("R0001"), IT, 3);
  await setGroupDone(byId("R0001"), "glass", false);
  await settle(900);
  assert.strictEqual(CALLS.length, 0, "and clicking writes nothing");
  assert.strictEqual(listWrites().length, 0, "not to the record either");
  pass("ready to deliver: the section is read-only and nothing can be written");

  /* ---- 5e. the open drawer is patched, not rebuilt ---- */
  useJob(mkJob({ win: "process", drs: "", glass: { tg: "done", tuff: "done" }, prod: {} }));
  setRecord([{ job: "R0001", item: IT, done: 6, total: 10, status: "process" },
             { job: "R0001", item: "glass:tg", done: 25, total: 25, status: "done" },
             { job: "R0001", item: "glass:tuff", done: 11, total: 11, status: "done" }]);
  const mkLine = item => {
    const num = { textContent: "" }, bar = { style: {} }, inp = { value: "" }, all = { textContent: "", dataset: {} };
    return { dataset: { cpline: item }, _num: num, _bar: bar, _inp: inp, _all: all,
             querySelector: sel => sel === ".cpnum" ? num : sel === ".cpbar i" ? bar
               : sel === ".cpin" ? inp : sel === ".cpall[data-cp]" ? all : null };
  };
  const sum = { textContent: "" }, line = mkLine(IT), tgLine = mkLine("glass:tg");
  const grpBtn = { textContent: "", dataset: { cpgrp: "glass", alltext: "All glass done" } };
  const host = { querySelector: sel => (sel === ".cpsum" ? sum : null),
                 querySelectorAll: sel => sel === "[data-cpline]" ? [line, tgLine] : sel === "[data-cpgrp]" ? [grpBtn] : [] };
  const realQ = document.querySelector;
  document.querySelector = sel => (sel === "#dhost" ? host : stubEl());
  cpPatchSection(byId("R0001"));
  assert.strictEqual(line._num.textContent, "6 of 10");
  assert.strictEqual(line._bar.style.width, "60%");
  assert.strictEqual(line._bar.style.background, "var(--fab)");
  assert.strictEqual(line._inp.value, 6);
  assert.strictEqual(line._all.textContent, "All done");
  assert.strictEqual(tgLine._bar.style.background, "var(--done)", "a finished item is gold");
  assert.strictEqual(grpBtn.textContent, "Clear", "the glass group is complete");
  assert.strictEqual(sum.textContent, "Windows 0/10 · Doors 0/2 · Glass 36/36 · Frames/Sashes/Transoms 6/18",
    "Windows counts nothing of its own now - it is in progress with no number of " +
    "its own - and the six frames are counted where they are actually recorded");
  document.activeElement = line._inp;                    // someone is typing in that box
  line._inp.value = "4";
  /* CHANGED 2026-09-11: what moves the row is the record, not a hold */
  vm.runInThisContext("cpRowNow('R0001', '" + IT + "', 8, 10, 'process', 'the admin');");
  cpPatchSection(byId("R0001"));
  assert.strictEqual(line._inp.value, "4", "the box being typed into is left alone");
  assert.strictEqual(line._num.textContent, "8 of 10", "everything else still updates");
  document.activeElement = null; document.querySelector = realQ;
  pass("a tap patches the open section in place instead of rebuilding it");

  /* ---- 6. unsent taps survive the tab closing ---- */
  useJob(mkIt("process"));
  setRecord([{ job: "R0001", item: IT, done: 1, total: 10, status: "process" }]);
  reset();
  setItemProgress(byId("R0001"), IT, 3);
  assert.strictEqual(queued().length, 1);
  CP.cpFireAll();                                  // what pagehide does
  await settle(60);
  assert.strictEqual(fills().length, 2,
    "the write goes out at once, not 800 ms later (this item, and its Windows copy)");
  assert.strictEqual(logWrites().length, 1);
  assert.strictEqual(queued().length, 0);
  pass("leaving the page sends the taps that were still waiting");

  reset();
  localStorage.setItem("cw_cpqueue", JSON.stringify([
    { key: "R0001|glass:tg", job: "R0001", item: "glass:tg", col: 51, from: 0, to: 12, total: 25, at: Date.now() - 60000 },
    { key: "R0009|win", job: "R0009", item: IT, col: ITCOL, from: 0, to: 2, total: 4, at: Date.now() - 7200000 }
  ]));
  assert.strictEqual(cpReplayQueue(), 1, "only the entry from within the hour is replayed");
  await settle(120);
  assert.strictEqual(fills().length, 1);
  assert.strictEqual(fills()[0].addr, "AY7");
  assert.strictEqual(fills()[0].color, "#FFFF00");
  assert.strictEqual(listWrites().length, 1, "one write to the record for the replayed tap");
  assert.strictEqual(cpSrvRow("R0001", "glass:tg").fields.Done, 12,
    "carrying the count the tap was made with");
  assert.deepStrictEqual(logWrites()[0].values[0].slice(3, 6), ["Glass TG", "0 of 25", "12 of 25"]);
  assert.strictEqual(queued().length, 0, "and it is not owed twice");
  pass("a tap left over from a previous visit is sent on the next load");

  /* ---- 6b. two group taps in a row ---- */
  useJob(mkIt(""));
  setRecord([]);
  reset();
  const p1 = setGroupDone(byId("R0001"), "glass", true);
  const midFlight = cpSectionHtml(byId("R0001"), true);
  assert(/data-cpgrp="glass"[^>]* disabled/.test(midFlight), "the group's button is dead while its write is in the air");
  const p2 = setGroupDone(byId("R0001"), "glass", false);   // the double tap
  await Promise.all([p1, p2]);
  await settle(60);
  /* CHANGED 2026-09-11: the counts go to the record rather than to a sheet, so
     the sequence is read off the wire - two record writes, two fills, then the
     same again - and the property is the one that always mattered: the two
     group writes never interleave. */
  assert.deepStrictEqual(SEQ, ["record", "record", "fill", "fill", "log",
                               "record", "record", "fill", "fill", "log"],
    "the two writes run one after the other, never interleaved: " + JSON.stringify(SEQ));
  assert.strictEqual(BOOK["Production"].fill[kk(7, 51)], "#FFFFFF", "Excel ends on the second tap");
  assert.strictEqual(cpSrvRow("R0001", "glass:tg").fields.Done, 0, "and so does the record");
  assert.strictEqual(cpSrvRow("R0001", "glass:tg").fields.Status, "");
  assert(!/data-cpgrp="glass"[^>]* disabled/.test(cpSectionHtml(byId("R0001"), true)), "the button comes back after");
  pass("two group taps: strictly one after the other, the second one wins in both places");

  /* ---- 6c. the column is worked out when the write goes, not when it was tapped ---- */
  useJob(mkIt(""));
  setRecord([]);
  reset();
  setItemProgress(byId("R0001"), IT, 3);
  global.__m2 = { qty: { wnd: 13, drs: 14 }, glass: { tg: 51, tuff: 52 },
                  prod: { "7000 casement": { f: 15, s: 16, t: 17 } }, prodOrder: ["7000 casement"] };
  vm.runInThisContext("PRODMAP = __m2");            // someone inserted two columns in Excel
  await settle(1400);
  assert.strictEqual(fills()[0].addr, "O7", "the fill goes to the column the sheet has now");
  assert.strictEqual(fills()[1].addr, "M7", "and the Windows copy to the column M is at now");
  pass("a column inserted between the tap and the write does not send the fill astray");

  useJob(mkIt(""));
  setRecord([]);
  reset();
  setItemProgress(byId("R0001"), IT, 4);
  vm.runInThisContext("PRODMAP = null");            // the sheet could not be read
  await settle(1400);
  assert.strictEqual(CALLS.length, 0, "nothing is written when the column cannot be worked out");
  assert.strictEqual(listWrites().length, 0, "and nothing reaches the record either");
  assert(TOASTS.some(t => t.err && /not a column/.test(t.m)));
  assert.strictEqual(itemState(byId("R0001"), IT).done, 0,
    "and the row goes back: the record must not say something the sheet has no cell for");
  pass("no PRODMAP when the write goes: nothing written anywhere, the row put back, a message");

  /* ---- 6d. what is replayed and what is not ---- */
  useJob(mkIt(""));
  setRecord([]);
  reset();
  const at = Date.now() - 30000;
  localStorage.setItem("cw_cpqueue", JSON.stringify([
    { key: "R0001|" + IT + "", job: "R0001", item: IT, col: ITCOL, from: 0, to: 3, total: 10, at: at, sent: 1 },
    /* NOT "drs": Windows and Doors are derived since 2026-09-14 and have no row
       of their own, so nothing can ever be queued against one */
    { key: "R0001|glass:tuff", job: "R0001", item: "glass:tuff", col: 52, from: 0, to: 1, total: 11, at: at, sent: 0 },
    { key: "R0001|glass:tg", job: "R0001", item: "glass:tg", col: 51, from: 0, to: 5, total: 25, at: at, who: "someone.else@example.test" }
  ]));
  assert.strictEqual(cpReplayQueue(), 1, "only the unsent entry belonging to this person");
  await settle(120);
  assert.strictEqual(fills().length, 1);
  assert.strictEqual(fills()[0].addr, "AZ7", "that one tuff tick, nothing else");
  pass("a write already sent is never replayed, nor one left by another account");

  /* ---- 6e. undoing ready to deliver paints white ---- */
  useJob(mkIt("", { done: 1 }));
  setRecord([]);
  reset();
  await markReady(byId("R0001"), false);
  assert.strictEqual(fills().length, 1);
  assert.strictEqual(fills()[0].addr, "A7:CL7");
  assert.strictEqual(fills()[0].color, "#FFFFFF", "white, not stripped: the sheet's cells are white, not blank");
  pass("undo of mark ready paints the row white instead of clearing the fill");

  /* ---- 6f. the list keeps its section names, and a group tick is not reported twice ---- */
  const withNames = [mkIt("")];
  withNames.blockNames = ["Second hand", "In production"];
  assert.deepStrictEqual(applyPending(withNames).blockNames, ["Second hand", "In production"]);
  pass("applyPending keeps the section names on the list");

  const t0 = Date.now();
  const iso = ms => new Date(t0 - ms).toISOString();
  const mineChanges = [
    { src: "dashboard", at: iso(60000), job: "R0001", what: "Glass: all done" },
    { src: "dashboard", at: iso(60000), job: "R0001", what: "7000 Casement: all done" },
    { src: "dashboard", at: iso(60000), job: "R0001", what: "Windows" },
    { src: "dashboard", at: iso(900000), job: "R0002", what: "Glass: all done" }   // too old to count
  ];
  const diffs = [
    { job: "R0001", what: "Glass TG" }, { job: "R0001", what: "Glass NOT TUFF" },
    { job: "R0001", what: "7000 CASEMENT frames" }, { job: "R0001", what: "Windows" },
    { job: "R0001", what: "Customer" }, { job: "R0002", what: "Glass TG" }
  ];
  assert.deepStrictEqual(dropMine(diffs, mineChanges, t0).map(c => c.job + " " + c.what),
    ["R0001 Customer", "R0002 Glass TG"]);
  pass("a group's log line also covers the per-item differences it caused");

  /* ---- 7. the parser reads the statuses out of the fills ---- */
  const wb = new ExcelJS.Workbook();
  const head = ws => {
    const put = (r, c, v) => { ws.getCell(r, c).value = v; };
    put(2, 4, "DATES ON CONTRACT"); put(3, 4, "SOLD"); put(2, 5, "CUSTOMER");
    put(2, 13, "QUANTITY"); put(3, 13, "WND"); put(3, 14, "DRS");
    put(2, 20, "7000 CASEMENT"); put(3, 20, "F"); put(3, 21, "S"); put(3, 22, "T");
    put(2, 51, "GLASS UNITS"); put(3, 51, "TG"); put(3, 52, "TUFF");
    /* DOORS DONE, five sub-columns numbered 1..5 (BT-BX on the real sheet) */
    put(2, 72, "DOORS DONE");
    for (let i = 0; i < 5; i++) put(3, 72 + i, String(i + 1));
  };
  const ws = wb.addWorksheet("Production"), ws2 = wb.addWorksheet("Production (2)");
  head(ws); head(ws2);
  const put = (sheet, r, c, v, colour) => {
    const cell = sheet.getCell(r, c);
    if (v !== undefined && v !== null) cell.value = v;
    if (colour) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + colour } };
  };
  /* R0001 on Production: WND yellow, DRS gold, TG gold, TUFF white, F yellow, S gold, T white */
  put(ws, 6, 3, "R0001"); put(ws, 6, 5, "Ann");
  put(ws, 6, 13, 10, "FFFF00"); put(ws, 6, 14, 2, "FFE699");
  put(ws, 6, 20, 12, "FFFF00"); put(ws, 6, 21, 8, "FFE699"); put(ws, 6, 22, 3, "FFFFFF");
  put(ws, 6, 51, 25, "FFE699"); put(ws, 6, 52, 11, "FFFFFF");
  /* and its doors: CD gold, "1 DOOR" yellow, SFCD uncoloured, one slot empty
     and one holding a stray word - ANY text is a door, whitespace alone is not */
  put(ws, 6, 72, "CD", "FFE699"); put(ws, 6, 73, "1 DOOR", "FFFF00");
  put(ws, 6, 74, "   "); put(ws, 6, 75, "sfcd"); put(ws, 6, 76, "n/a");
  /* the same job on Production (2), still yellow where Production has been
     cleared - and, since 2026-09-17, carrying a DIFFERENT number for the same
     product group. That is what the live workbook did: a column was inserted
     into `Production`, `Production (2)` is formulas that moved with it, and its
     headers stayed put - so every number on that sheet sits one column away
     from the header describing it. 99 stands in for that here. */
  put(ws2, 6, 3, "R0001"); put(ws2, 6, 13, 10, "FFFF00"); put(ws2, 6, 52, 11, "FFFF00");
  put(ws2, 6, 20, 12, "FFE699"); put(ws2, 6, 22, 77);
  /* R0002: the whole Production row is gold */
  put(ws, 7, 3, "R0002"); put(ws, 7, 13, 4); put(ws, 7, 20, 5); put(ws, 7, 51, 6);
  put(ws, 7, 14, 1); put(ws, 7, 72, "DD");
  for (let c = 1; c <= 80; c++) put(ws, 7, c, undefined, "FFE699");
  const jobs = parseWorkbook(wb);
  const a = jobs.find(x => x.id === "R0001"), b = jobs.find(x => x.id === "R0002");
  assert.strictEqual(a.cp.win, "process");
  assert.strictEqual(a.cp.drs, "done");
  assert.strictEqual(a.cp.glass.tg, "done");
  /* CHANGED 2026-09-11: the parser still reads every colour, and cpFileStatus
     is the only way to get at them. cpStatus is the RECORD and answers nothing
     at all for a job the record has never heard of - which is what the whole
     change is about. */
  assert.strictEqual(cpFileStatus(a, "glass:tuff"), "", "a white cell leaves the item untouched");
  assert.strictEqual(cpFileStatus(a, "prod:7000 casement:f"), "process");
  assert.strictEqual(cpFileStatus(a, "prod:7000 casement:s"), "done");
  assert.strictEqual(cpFileStatus(a, "prod:7000 casement:t"), "");
  assert.deepStrictEqual(a.prods[0].st.sort(), ["done", "process"], "the old product status still works");
  pass("parser: cp colours for M, N, the glass columns and F/S/T, still read, still parsed");

  /* ---- 7a. `prods` merges every sheet; `prodsMain` is Production alone ----
     Added 2026-09-17 after the live bug (docs/HISTORY.md B20). `Production (2)`
     says 77 transoms for a group `Production` says 3. The cross-sheet Math.max
     takes the 77 - which is what every existing consumer has always seen, and
     is deliberately left that way in this pass - while `prodsMain` is the main
     sheet's own answer and cannot inherit another sheet's misalignment. */
  assert.strictEqual(a.prods[0].n, "7000 casement");
  assert.strictEqual(a.prods[0].t, 77, "prods still takes the highest across the sheets");
  assert.ok(Array.isArray(a.prodsMain), "every job carries prodsMain");
  assert.strictEqual(a.prodsMain.length, 1, "one group on Production, whatever the copies say");
  assert.strictEqual(a.prodsMain[0].n, "7000 casement");
  assert.strictEqual(a.prodsMain[0].t, 3, "prodsMain is the Production sheet's own number");
  assert.strictEqual(a.prodsMain[0].f, 12);
  assert.strictEqual(a.prodsMain[0].s, 8);
  assert.deepStrictEqual(a.prodsMain[0].st.sort(), ["done", "process"],
    "and it carries the same status words, which are read from Production anyway");
  /* a job on Production with nothing on any copy is identical in both */
  assert.deepStrictEqual(b.prodsMain.map(p => [p.n, p.f, p.s, p.t]),
                         b.prods.map(p => [p.n, p.f, p.s, p.t]),
    "a job with no second-sheet row at all reads the same either way");
  pass("parser: prods still merges every sheet, prodsMain is the Production sheet and nothing else");
  /* ---- 7b. the DOORS DONE cells ---- */
  const m7 = mapSheet(ws);
  assert.deepStrictEqual(m7.doors, { 1: 72, 2: 73, 3: 74, 4: 75, 5: 76 },
    "the five sub-columns are mapped by their own numbers");
  assert.ok(!Object.keys(m7.prod).some(k => k.indexOf("doors done") >= 0),
    "and DOORS DONE is still kept out of the product groups, where it has no F/S/T");
  assert.deepStrictEqual(a.doors, [
    { slot: 1, code: "CD", status: "done" },
    { slot: 2, code: "1 DOOR", status: "process" },
    { slot: 4, code: "SFCD", status: "" },
    { slot: 5, code: "N/A", status: "" }
  ], "any non-empty text is a door, trimmed and upper-cased; whitespace alone is not");
  assert.strictEqual(cpFileStatus(a, "door:1"), "done");
  assert.strictEqual(cpFileStatus(a, "door:2"), "process");
  assert.strictEqual(cpFileStatus(a, "door:4"), "");
  assert.strictEqual(cpTotal(a, "door:3"), 0, "slot 3 held nothing but spaces");
  pass("parser: one door per DOORS DONE cell, its text the code and its fill the status");
  pass("a colour on Production (2) never overrides Production, so un-ticking sticks");
  setRecord([]);
  cpItems(a).forEach(x => assert.strictEqual(cpStatus(a, x.key), "",
    x.key + " has no record, so it is not done however the cell is coloured"));
  pass("and not one of those colours is status: with no record, nothing on that job is ticked");
  /* the import is what turns them into a record, once */
  recordFromColours(a);
  assert.deepStrictEqual(itemState(a, IT), { done: null, total: 12, status: "process" });
  /* Windows and Doors keep their own record rows (amendment 1) AND answer for
     what is under them - the higher of the two is what is shown. M is yellow
     in the file, so its own row says in fabrication; its window types say the
     same, so in fabrication is what it reads. */
  assert.deepStrictEqual(itemState(a, "win"), { done: null, total: 10, status: "process" });
  /* N is GOLD in the file, so its own row says done - and it stays done even
     though two of the four doors under it have not been started. That is the
     owner's rule: "already golden means it finished even if all the cell or
     component has no ticked." */
  assert.strictEqual(cpFileStatus(a, "drs"), "done");
  assert.deepStrictEqual(itemState(a, "drs"), { done: 2, total: 2, status: "done" },
    "a gold N stays done however little is ticked under it");
  assert.strictEqual(cpDerivedOf(a, "drs"), "process", "even though its doors only add up to in fabrication");
  /* and the aggregate still takes a line UP: clear N's own row and the doors
     carry it to in fabrication on their own */
  vm.runInThisContext("cpRowPut('R0001', 'drs', null);");
  assert.deepStrictEqual(itemState(a, "drs"), { done: null, total: 2, status: "process" },
    "with no row of its own, N reads what its doors say");
  assert.strictEqual(b.done, 1, "the gold row is a finished job");
  assert.strictEqual(b.cp.win, "done");
  assert.strictEqual(b.cp.glass.tg, "done");
  assert.strictEqual(b.cp.prod["7000 casement"].f, "done");
  recordFromColours(b);
  cpItems(b).forEach(x => assert.strictEqual(itemState(b, x.key).status, "done", x.key + " is done"));
  pass("a whole gold row, imported once, counts as done for every checkpoint on it");

  /* ---- 8. what the Changes list says about a colour that moved in Excel ---- */
  const before2 = JSON.parse(JSON.stringify(a)), after = JSON.parse(JSON.stringify(a));
  setFileCp(after, "done");
  const d = diffJobs([before2], [after], "someone", "2026-09-04T10:00:00Z");
  assert.deepStrictEqual(d.map(x => [x.what, x.from, x.to]),
    [["7000 CASEMENT frames", "in fabrication", "done"]]);
  /* and a door's cell moving in Excel is listed by the door's own name */
  const afterD = JSON.parse(JSON.stringify(a));
  afterD.doors[1].status = "done";
  assert.deepStrictEqual(diffJobs([JSON.parse(JSON.stringify(a))], [afterD], "someone", "2026-09-04T10:00:00Z")
    .map(x => [x.what, x.from, x.to]), [["Door 2", "in fabrication", "done"]]);
  pass("a colour change is listed by the item's own name, the way the log names it");

  /* ================= 9. THE ONE-TIME IMPORT (spec section 4) ================= */
  const impJob = mkJob({ win: "done", drs: "process", glass: { tg: "process", tuff: "" },
                         prod: { "7000 casement": { f: "done", s: "" } } },
                       { doors: [{ slot: 1, code: "CD", status: "process" }] });
  useJob(impJob);
  setRecord([]);
  /* the exact count behind the yellow TG, off the Dashboard Progress SHEET -
     the last thing that sheet is for */
  CP.cpSetProgress({ R0001: { "glass:tg": { done: 9, total: 25 } } });
  setPainted({});
  settled(false);                          // the switch-on window: nothing imported yet
  reset();
  const imported = await cpImportRun();
  await settle(150);                       // the log line is written without being awaited
  assert.strictEqual(imported, 5,
    "five cells say something: win, drs, glass:tg, the F frames and the job's one door");
  assert.strictEqual(fills().length, 0, "the import paints nothing: the colours are already there");
  const impRows = listWrites();
  assert.ok(impRows.every(w => w.method === "POST"), "every imported row is created, none patched");
  assert.ok(impRows.every(w => w.body.fields.Source === "import"), "and every one says where it came from");
  assert.deepStrictEqual(impRows.map(w => w.body.fields.Item).sort(),
    ["door:1", "drs", "glass:tg", IT, "win"]);
  /* Windows and Doors are ticked in their own right again (amendment 1), so
     they are imported like any other item - and the general pass covers the
     doors too, which is why it also settles the doors marker */
  assert.strictEqual(cpSrvRow("R0001", "win").fields.Done, 10, "M imports at the total");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Done, 10, "a gold cell imports at the total");
  assert.strictEqual(cpSrvRow("R0001", "glass:tg").fields.Done, 9,
    "a yellow cell imports the exact count the sheet had for it");
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Done, 0,
    "and a yellow cell with no count imports as in progress, count unknown");
  assert.strictEqual(itemState(byId("R0001"), "door:1").done, null, "which is how it reads back");
  assert.deepStrictEqual(itemState(byId("R0001"), "drs"), { done: null, total: 2, status: "process" },
    "so the derived Doors line reads in fabrication, from that one door");
  assert.strictEqual(cpSrvRow("R0001", "glass:tuff"), null,
    "a BLANK cell gets no row at all: no row IS nothing done, and importing every blank would " +
    "be a row per checkpoint of every job on the sheet");
  assert.strictEqual(paintedOf("R0001", IT), "done",
    "and the colour it imported is remembered as the one that is in the cell");
  const impLog = logWrites();
  assert.strictEqual(impLog.length, 1, "one summary line, not one per row");
  assert.ok(/Imported 5 checkpoints from the sheet's colours/.test(impLog[0].values[0][5]),
    "saying how many, in plain words: " + impLog[0].values[0][5]);
  pass("the import adopts today's colours once, with Source = import and one summary line");

  reset();
  assert.strictEqual(await cpImportRun(), 0, "a second load imports nothing");
  await settle(150);
  assert.strictEqual(listWrites().length, 0, "and writes nothing at all");
  assert.strictEqual(logWrites().length, 0, "and says nothing");
  assert.strictEqual(vm.runInThisContext("cpImportSettled()"), true,
    "and the switch-on window is declared over");
  assert.strictEqual(cpImportPending(), false, "so nothing falls back to the colour any more");
  pass("nothing is ever imported twice: a row that exists is never looked at again");

  /* the cap. A page opened on a sheet nobody has ticked in a month must not
     fire a thousand writes at once. */
  const many = [];
  for (let i = 1; i <= 40; i++) {
    many.push(mkJob({ win: "done", drs: "done", glass: { tg: "done", tuff: "done" },
                      prod: { "7000 casement": { f: "done", s: "done" } } },
                    { id: "R" + (2000 + i) }));
  }
  global.__many = many;
  vm.runInThisContext("ALL = __many; PRODMAP = __m;");
  setRecord([]);
  CP.cpSetProgress({});
  setPainted({});
  settled(false);
  reset();
  const capped = await cpImportRun();
  await settle(200);
  assert.strictEqual(capped, 60, "sixty rows in a load, and not the two hundred and forty on the sheet");
  assert.strictEqual(listWrites().length, 60, "sixty writes, no more");
  pass("the import is capped per load, so the first load after switch-on is not a write storm");
  vm.runInThisContext("if (cpImportAgainT) { clearTimeout(cpImportAgainT); cpImportAgainT = null; }");

  /* ---- M2: the follow-up must not fire into a closed door ----
     A pass of sixty writes can outlast the thirty-second timer. The timer then
     fired, met `if (cpImporting) return 0`, and armed nothing - so the import
     stalled until the next load(), which on a quiet workbook may never come. */
  vm.runInThisContext("cpImporting = true; if (cpImportAgainT) { clearTimeout(cpImportAgainT); } cpImportAgainT = null;");
  assert.strictEqual(await cpImportRun(), 0, "a pass is already running, so this one does nothing");
  assert.ok(vm.runInThisContext("!!cpImportAgainT"), "but it arms the next one before it goes");
  vm.runInThisContext("cpImporting = false; clearTimeout(cpImportAgainT); cpImportAgainT = null;");
  /* and a pass that throws re-arms too, because `owed` starts true */
  const realSend = vm.runInThisContext("cpSend");
  global.__boom = async () => { throw new Error("stub: the lane runner fell over"); };
  vm.runInThisContext("cpSend = __boom;");
  settled(false);
  useJob(mkIt("done"));
  setRecord([]); settled(false);
  await cpImportRun().catch(() => {});
  assert.ok(vm.runInThisContext("!!cpImportAgainT"), "a pass that throws still comes back");
  global.__real = realSend;
  vm.runInThisContext("cpSend = __real; clearTimeout(cpImportAgainT); cpImportAgainT = null;");
  pass("the import cannot stall: the follow-up is armed even when the pass is busy or throws");

  /* ============ 10. THE SAFEGUARD: a colour changed in Excel by hand ========
     Spec section 4a. The download FINDS a candidate; the Excel API - which
     reads the live file with no lag - CONFIRMS it. */
  const sg = mkIt("process", { src: { Production: 7 } });
  useJob(sg);
  setRecord([{ job: "R0001", item: IT, done: 4, total: 10, status: "process" }]);
  setPainted(paintedIt("process"));
  /* somebody has coloured M7 gold in Excel, and the download has caught up */
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFE699";
  setFileCp(sg, "done");
  vm.runInThisContext("LASTBY = 'the colleague';");
  reset();
  const adopted = await cpAdoptRun();
  await settle(150);
  assert.strictEqual(adopted, 1, "one cell adopted");
  assert.strictEqual(fills().length, 0, "nothing is painted: Excel already has the colour");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Status, "done", "the record now says done");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Done, 10);
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Source, "excel", "and where it came from");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Who, "the colleague",
    "with the file's last editor - never 'painted by', which Excel cannot tell us");
  assert.strictEqual(paintedOf("R0001", IT), "done", "and it is not adopted twice");
  const adoptLog = logWrites();
  assert.strictEqual(adoptLog.length, 1, "one log line");
  assert.deepStrictEqual(adoptLog[0].values[0][3], "7000 CASEMENT frames", "named the way every other change is");
  assert.ok(/adopted from Excel/.test(adoptLog[0].values[0][5]) &&
            /file last edited by the colleague/.test(adoptLog[0].values[0][5]),
    "saying it came from Excel and who last edited the file: " + adoptLog[0].values[0][5]);
  reset();
  assert.strictEqual(await cpAdoptRun(), 0, "and the same cell is not adopted again");
  assert.strictEqual(listWrites().length, 0);
  pass("a hand-paint the API confirms is adopted at once, with the row, the source and one log line");

  /* a STALE download: the file we downloaded says gold, the API says what we
     painted. Nothing happens, and nothing is said. */
  setRecord([{ job: "R0001", item: IT, done: 4, total: 10, status: "process" }]);
  setPainted(paintedIt("process"));
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFFF00";        // the LIVE cell is still yellow
  setFileCp(sg, "done");                                     // ... but the download says gold
  reset();
  assert.strictEqual(await cpAdoptRun(), 0, "nothing is adopted");
  await settle(150);
  assert.strictEqual(listWrites().length, 0, "nothing is written");
  assert.strictEqual(logWrites().length, 0, "nothing is logged");
  assert.strictEqual(TOASTS.length, 0, "and nothing is said: it is not a fault, it is a slow copy");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Status, "process", "the record is untouched");
  pass("a stale download is confirmed away by the API and ignored silently");

  /* the dashboard's own paint is never adopted as somebody else's */
  setRecord([{ job: "R0001", item: IT, done: 10, total: 10, status: "done" }]);
  setPainted(paintedIt("done"));
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFE699";
  setFileCp(sg, "done");
  reset();
  assert.strictEqual(await cpAdoptRun(), 0, "the file agrees with what we painted: nothing to adopt");
  assert.strictEqual(LISTREQ.length, 0, "and it is not even confirmed: no read is made at all");
  pass("the dashboard's own paint is never adopted as an outside change");

  /* CHANGED at the review (finding M4). This used to prove that the open
     drawer's job was read through the API on every ten-second poll WHATEVER
     the download said, so a hand-paint showed in ten seconds without waiting
     for a download at all. That cost about a hundred workbook operations a
     minute for as long as a drawer stayed open and adopted nothing almost
     always, so the download is now the DISCOVERY on both paths and the API
     read is only ever the CONFIRMATION. What the drawer still buys is the
     confirmation on the ten-second clock rather than on the next load. */
  setRecord([{ job: "R0001", item: IT, done: 4, total: 10, status: "process" }]);
  setPainted(paintedIt("process"));
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFE699";
  setFileCp(sg, "process");                                  // the download has NOT caught up
  vm.runInThisContext("state.sel = 'R0001'; cpDrawerJob = ''; cpDrawerAt = 0;");
  const realQ2 = document.querySelector;
  document.querySelector = () => stubEl();
  reset();
  assert.strictEqual(await cpAdoptRun(), 0,
    "the download shows nothing, so the ordinary pass has no candidate");
  assert.strictEqual(await cpAdoptDrawerJob(), 0,
    "and neither has the drawer: the API is not asked about a cell nothing suggests has moved");
  assert.strictEqual(CALLS.length, 0, "so a drawer left open reads nothing at all");
  setFileCp(sg, "done");                                     // now the download catches up
  vm.runInThisContext("cpDrawerJob = ''; cpDrawerAt = 0;");
  reset();
  assert.strictEqual(await cpAdoptDrawerJob(), 1,
    "and the drawer's job is confirmed on the ten-second clock, without waiting for a load");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Status, "done");
  document.querySelector = realQ2;
  vm.runInThisContext("state.sel = null;");
  pass("the drawer confirms on the ten-second clock, and asks the API for nothing else");

  /* twenty-five candidates are confirmed in two batches and adopted once each */
  const bulk = [];
  for (let i = 1; i <= 25; i++) {
    BOOK["Production"].v[kk(20 + i, 3)] = "R" + (3000 + i);
    BOOK["Production"].fill[kk(20 + i, ITCOL)] = "#FFE699";
    bulk.push(mkIt("done", { id: "R" + (3000 + i), src: { Production: 20 + i } }));
  }
  global.__bulk = bulk;
  vm.runInThisContext("ALL = __bulk; PRODMAP = __m; state.sel = null;");
  setRecord(bulk.map(x => ({ job: x.id, item: IT, done: 0, total: 10, status: "" })));
  setPainted(bulk.reduce((m, x) => { m[x.id] = {}; m[x.id][IT] = ""; return m; }, {}));
  reset();
  const nBatch = BATCHES;
  const bulkAdopted = await cpAdoptRun();
  assert.strictEqual(bulkAdopted, 25, "all twenty-five adopted");
  assert.strictEqual(BATCHES - nBatch, 3,
    "fifty reads - twenty-five fills and twenty-five row checks - in three batches of twenty");
  assert.strictEqual(listWrites().length, 25, "one write each, and not one twice");
  reset();
  assert.strictEqual(await cpAdoptRun(), 0, "and a second pass has nothing left to do");
  pass("a bulk hand-paint is confirmed in whole batches and adopted exactly once each");

  /* ============ 11. THE SWITCH-ON WINDOW (review finding M1) ==============
     The import takes 15-25 minutes on the owner's real sheet. Until it has
     drained, an item with no row must read as the SHEET's colour, or every
     consumer that used to read the colour reads a hole: the phase pipeline
     collapses every job to In office, the feeder writes OfficeDone "No" and
     then "Yes" again per job as rows land, the floor's seeds go to nought, and
     a job marked ready prints "all checkpoints are complete" over lines
     reading 0. */
  const winJob = () => mkJob({ win: "done", drs: "process",
                               glass: { tg: "done", tuff: "" },
                               prod: { "7000 casement": { f: "done", s: "process" } } });
  /* what every screen said BEFORE the switch-over: the record as the import
     will eventually write it */
  useJob(winJob());
  CP.cpSetProgress({ R0001: { "glass:tg": { done: 20, total: 25 } } });
  recordFromColours(byId("R0001"), { "glass:tg": 20 });
  settled(true);
  const beforePhase = jobPhase(byId("R0001"));
  const beforeStates = cpItems(byId("R0001")).map(x => JSON.stringify(itemState(byId("R0001"), x.key)));
  const beforeCounts = JSON.stringify(glassCounts(byId("R0001")));
  const beforeSummary = cpSummaryHtml(byId("R0001"));
  /* now the same job on the morning of the switch-over: not one row yet */
  useJob(winJob());
  setRecord([]);
  settled(false);
  assert.strictEqual(cpImportPending(), true, "the window is open");
  assert.strictEqual(jobPhase(byId("R0001")), beforePhase,
    "the phase pipeline reads exactly what it read before the switch-over");
  assert.deepStrictEqual(cpItems(byId("R0001")).map(x => JSON.stringify(itemState(byId("R0001"), x.key))),
    beforeStates, "and so does every item, count and all");
  assert.strictEqual(JSON.stringify(glassCounts(byId("R0001"))), beforeCounts,
    "and so does glassCounts, which is what the feeder's OfficeDone and seed are derived from");
  assert.strictEqual(ST.officeComplete(glassCounts(byId("R0001"))), true,
    "so the office's lock does not flap off and on as rows land");
  assert.strictEqual(cpSummaryHtml(byId("R0001")), beforeSummary,
    "and the drawer's summary line is the same, so a gold row does not read 0");
  /* an item that DOES have a row always answers from the row, window or not:
     an un-tick made during the window is still absolute */
  vm.runInThisContext("cpRowNow('R0001', '" + IT + "', 0, 10, '', 'the admin');");
  assert.deepStrictEqual(itemState(byId("R0001"), IT), { done: 0, total: 10, status: "" },
    "an un-tick inside the window is the record's word, not the sheet's gold");
  /* and once the import has drained, the fallback is dead */
  settled(true);
  setRecord([]);
  assert.strictEqual(cpImportPending(), false);
  assert.deepStrictEqual(itemState(byId("R0001"), IT), { done: 0, total: 10, status: "" },
    "a gold cell with no row is nothing done again");
  assert.strictEqual(jobPhase(byId("R0001")), 0, "and the colour is never consulted again");
  pass("the switch-on window reads the sheet's colours, an un-tick still wins, and it ends");

  /* ============ 12. PUTTING THE EXCEL COPY RIGHT (review finding M3) ======
     A refused fill leaves the record saying one thing and the sheet another,
     and the safeguard cannot see it: it compares the FILE with what we
     painted, and those two agree. So the record is compared with PAINTED as
     well, and a disagreement is painted again on the next load. */
  useJob(mkIt("process"));
  setRecord([{ job: "R0001", item: IT, done: 4, total: 10, status: "process" }]);
  setPainted(paintedIt("process"));
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFFF00";
  settled(true);
  reset(); FAIL_FILL = 1;
  setItemProgress(byId("R0001"), IT, 10);                 // gold, if it would land
  await settle(1400);
  assert.strictEqual(fills().length, 0, "the fill was refused");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Status, "done", "the record says done");
  assert.strictEqual(paintedOf("R0001", IT), "process", "the sheet still has yellow");
  reset();
  assert.strictEqual(await cpAdoptRun(), 0,
    "the safeguard cannot see it: the file and what we painted agree");
  assert.strictEqual(CALLS.filter(c => c.kind === "readFill").length, 0, "so it reads nothing");
  /* the next load puts it right - BOTH cells the refused write owed: this
     item's own, and the derived Windows copy above it that never landed either */
  const painted1 = await cpRepaintRun();
  assert.strictEqual(painted1, 2, "two cells repainted: the item, and its Windows copy");
  assert.strictEqual(fills().length, 2, "through the ordinary fill path");
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color).sort(),
    ["M7 #FFFF00", "T7 #FFE699"],
    "the job's own cell gold, and M yellow because its sashes are not started");
  assert.strictEqual(paintedOf("R0001", IT), "done", "and PAINTED catches up, so it stops");
  assert.strictEqual(logWrites().length, 0, "nothing is logged: this is a copy put right, not a change");
  reset();
  assert.strictEqual(await cpRepaintRun(), 0, "a second load has nothing to do");
  pass("a refused fill is repainted on the next load, and one success ends it");

  /* F1: THE DOWNLOAD IS NOT EVIDENCE THAT THE SHEET IS RIGHT.
     The exact sequence the reviewer found. The office un-ticks (record "", the
     fill lands, PAINTED ""). The floor then finishes the job and the writer's
     FILL is refused (record "done", PAINTED still ""). The download in hand is
     still the pre-un-tick gold - so it AGREES with the record. A shortcut that
     concluded "the file already shows it, remember that and skip" wrote
     PAINTED = "done" and the cell was never a candidate again: Excel white,
     record done, for ever. There is no such shortcut. */
  useJob(mkIt("", { wnd: 0 }));                // no Windows cell in play: one cell, one question
  setRecord([{ job: "R0001", item: IT, done: 10, total: 10, status: "done", source: "floor" }]);
  setPainted(paintedIt(""));                   // the office's white is what really landed
  setFileCp(byId("R0001"), "done");            // ... and the 36 s-old download still says gold
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFFFFF";
  reset();
  assert.strictEqual(await cpRepaintRun(), 1,
    "the cell is repainted: a stale download agreeing with the record proves nothing");
  assert.deepStrictEqual([fills()[0].addr, fills()[0].color], [ITADDR, "#FFE699"]);
  assert.strictEqual(paintedOf("R0001", IT), "done",
    "and PAINTED moves only now, because a fill actually landed");
  pass("the repaint never concludes from the download that the sheet is already right");

  /* a workbook that keeps refusing: once per load, never a loop inside one */
  setPainted(paintedIt("process"));
  reset(); FAIL_FILL = 9;
  const tried = await cpRepaintRun();
  assert.strictEqual(tried, 0, "nothing was painted");
  assert.strictEqual(CALLS.filter(c => c.kind === "fillRefused").length, 1,
    "and it was tried exactly once, not in a loop");
  assert.strictEqual(paintedOf("R0001", IT), "process", "PAINTED does not move on a refusal");
  FAIL_FILL = 0;
  assert.strictEqual(await cpRepaintRun(), 1, "the next load tries again, and this time it lands");
  pass("a persistent refusal is retried once per load and never loops inside one");

  /* CHANGED at step 3. The four glass columns used to be excluded, because the
     colour writer painted them without going through the record and two
     writers on one cell would have fought. The writer goes through the record
     now, so they are ordinary managed cells - which is what puts a FLOOR
     colour into Excel when the fill was refused the first time. */
  useJob(mkJob({ win: "", drs: "", glass: { tg: 25, tuff: 11 }, prod: {} }));
  setRecord([{ job: "R0001", item: "glass:tg", done: 25, total: 25, status: "done", source: "floor" }]);
  setPainted({ R0001: { "glass:tg": "" } });
  reset();
  assert.strictEqual(await cpRepaintRun(), 1, "a glass cell is repainted like any other");
  assert.deepStrictEqual([fills()[0].addr, fills()[0].color], ["AY7", "#FFE699"],
    "the floor's gold reaches the sheet on the load after a refused fill");
  /* except a cell carrying a colour this feature does not own */
  setPainted({ R0001: { "glass:tuff": "" } });
  setRecord([{ job: "R0001", item: "glass:tuff", done: 11, total: 11, status: "done", source: "floor" }]);
  byId("R0001").cp.glass = { tuff: "cut" };
  reset();
  assert.strictEqual(await cpRepaintRun(), 0, "the sheet's own Cut green is not painted over");
  assert.strictEqual(CALLS.length, 0);
  pass("the repaint covers the glass columns too, and still keeps off a colour it does not own");

  /* ============ 13. THE OPEN DRAWER COSTS NOTHING (review finding M4) =====
     The live check used to ask the Excel API about every managed item of the
     drawer's job on every ten-second poll - about a hundred workbook
     operations a minute for as long as the drawer stayed open, adopting
     nothing. The download is the DISCOVERY; the API read is the CONFIRMATION. */
  const drawerJob = mkJob({ win: "done", drs: "done", glass: { tg: 25, tuff: 11 },
                            prod: { "7000 casement": { f: "done", s: "done" } } });
  useJob(drawerJob);
  recordFromColours(byId("R0001"));
  settled(true);
  cpItems(byId("R0001")).forEach(x => paintedSet("R0001", x.key, cpFileStatus(byId("R0001"), x.key)));
  vm.runInThisContext("state.sel = 'R0001'; cpDrawerJob = ''; cpDrawerAt = 0;");
  const realQ3 = document.querySelector;
  document.querySelector = () => stubEl();
  reset();
  for (let i = 0; i < 6; i++) await cpAdoptDrawerJob();      // a minute of ten-second polls
  assert.strictEqual(CALLS.filter(c => c.kind === "readFill").length, 0,
    "sixty seconds with a drawer open and nothing changed: not one fill read");
  assert.strictEqual(CALLS.length, 0, "not one request of any kind");
  /* somebody paints a cell in Excel and the download shows it: now it reads,
     once, and not again for thirty seconds */
  setFileCp(byId("R0001"), "");                   // somebody cleared the frames cell in Excel
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFFFFF";
  reset();
  assert.strictEqual(await cpAdoptDrawerJob(), 1, "the change is adopted");
  const readsOnce = CALLS.filter(c => c.kind === "readFill").length;
  assert.ok(readsOnce >= 1, "having read the cell through the API");
  reset();
  assert.strictEqual(await cpAdoptDrawerJob(), 0, "and the next poll ten seconds later reads nothing");
  assert.strictEqual(CALLS.length, 0, "because the same job is not asked again inside thirty seconds");
  document.querySelector = realQ3;
  vm.runInThisContext("state.sel = null;");
  pass("an open drawer costs nothing while nothing changes, and asks at most every thirty seconds");

  /* ============ 14. A BLANK CELL PAINTED LATER (review finding M5) ========
     The import owns today's colours; a colour that APPEARS on a cell with no
     record afterwards is a hand-paint, and belongs to the safeguard - with the
     API confirming it and a log line naming it, not swallowed as
     Source = "import", Who = "the sheet", unconfirmed and unlogged. */
  useJob(mkIt(""));
  setRecord([]);
  setPainted({});
  settled(true);
  vm.runInThisContext("LASTBY = 'the colleague';");
  setFileCp(byId("R0001"), "done");                 // painted gold in Excel, by hand
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFE699";
  await settle(200);                                // let the last test's log line land first
  reset();
  assert.strictEqual(await cpImportRun(), 0, "the import does not touch it: its work is done");
  assert.strictEqual(await cpAdoptRun(), 1, "the safeguard does");
  await settle(150);
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Source, "excel", "recorded as an Excel change");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Who, "the colleague");
  assert.strictEqual(cpSrvRow("R0001", IT).fields.Done, 10);
  assert.ok(CALLS.some(c => c.kind === "readFill"), "and confirmed through the Excel API first");
  assert.strictEqual(logWrites().length, 1, "with a log line of its own");
  assert.ok(/adopted from Excel/.test(logWrites()[0].values[0][5]));
  pass("a blank cell painted in Excel afterwards is adopted by the safeguard, not by the import");

  /* ============ 15. A BACKGROUND WRITE NEVER ASKS FOR CONSENT (M6) ========
     listUpsert calls listConsent(), which may throw a popup at somebody. Fine
     from a click; never from a load or a ten-second timer. */
  let consents = 0;
  const realConsent = CW.listConsent;
  CW.listConsent = async () => { consents++; return "t"; };
  useJob(mkIt(""));
  setRecord([]);
  /* a cancelled burst left a row with only a local id and no item on the server */
  vm.runInThisContext("cpRowNow('R0001', '" + IT + "', 4, 10, 'process', 'the admin');");
  assert.ok(String(cpRow("R0001", IT).id).indexOf("local:") === 0, "the row has no real id");
  setPainted(paintedIt("process"));
  setFileCp(byId("R0001"), "done");
  BOOK["Production"].fill[kk(7, ITCOL)] = "#FFE699";
  settled(true);
  reset();
  const skipped = await cpAdoptRun();
  assert.strictEqual(skipped, 0, "the adoption is skipped rather than creating a row");
  assert.strictEqual(consents, 0, "and consent was never asked for from a background pass");
  assert.strictEqual(listWrites().length, 0, "nothing was written");
  /* the next click fixes it, and a click MAY ask */
  reset();
  setItemProgress(byId("R0001"), IT, 6);
  await settle(1400);
  assert.ok(listWrites().length >= 1, "the click writes the row properly");
  CW.listConsent = realConsent;
  pass("a background adoption never opens a consent window: it leaves the cell to the next click");

  /* ============ 16. A GROUP WRITE THAT FAILS HALF WAY (review finding M7) = */
  useJob(mkIt(""));
  setRecord([]);
  settled(true);
  reset(); FAIL_LIST = 1;                          // the FIRST of the two rows is refused
  await setGroupDone(byId("R0001"), "glass", true);
  await settle(80);
  assert.strictEqual(fills().length, 0, "nothing is painted when the record refused");
  assert.ok(TOASTS.some(t => t.err), "and the office is told");
  FAIL_LIST = 0;
  setRecord([]);
  reset();
  /* now let the first row land and refuse the second */
  global.__fail2 = 1;
  vm.runInThisContext("(function(){ const real = cpSaveRow; cpSaveRow = async function (o) {" +
    " if (o.item === 'glass:tuff' && __fail2) { __fail2 = 0; throw new Error('stub: refused'); }" +
    " return real(o); }; __realSave = real; })();");
  await setGroupDone(byId("R0001"), "glass", true);
  await settle(80);
  vm.runInThisContext("cpSaveRow = __realSave;");
  const half = TOASTS.filter(t => t.err).map(t => t.m).join(" | ");
  assert.ok(/Glass TG/.test(half) && /Glass TUFF/.test(half),
    "the message names both the item that landed and the one that did not: " + half);
  assert.ok(/was saved/.test(half) && /was not/.test(half),
    "and says which is which: " + half);
  assert.strictEqual(fills().length, 0, "and nothing is painted for a half-written group");
  pass("a group write that fails half way says which items took and which did not");

  /* ============ 17. DOORS: THE WRITE (2026-09-14, the doors spec) ==========
     A door is one thing with three states, so the write carries the WORD, not
     a count. The only new cell this feature may ever colour is that door's own
     DOORS DONE cell, and the code text in it is never touched. */
  const DMAP = { qty: { wnd: 13, drs: 14 }, glass: {}, prod: {}, prodOrder: [],
                 doors: { 1: 72, 2: 73, 3: 74, 4: 75, 5: 76 } };
  const doorJob = () => mkJob({ win: "", drs: "", glass: {}, prod: {} },
    { wnd: 0, glass: {}, prods: [], drs: 2,
      doors: [{ slot: 1, code: "CD", status: "" }, { slot: 3, code: "SS", status: "" }] });
  useJob(doorJob(), DMAP);
  BOOK["Production"].v[kk(7, 72)] = "CD";                 // the code, as the office typed it
  BOOK["Production"].v[kk(7, 74)] = "SS";
  setRecord([]);
  reset();
  setDoorStatus(byId("R0001"), "door:1", "process");
  assert.deepStrictEqual(itemState(byId("R0001"), "door:1"), { done: null, total: 1, status: "process" },
    "the drawer says in fabrication before anything has been written");
  await settle(1400);
  const dw = listWrites();
  assert.strictEqual(dw.length, 1, "exactly one write to the record");
  assert.deepStrictEqual([dw[0].body.fields.Job, dw[0].body.fields.Item, dw[0].body.fields.Done,
                          dw[0].body.fields.Total, dw[0].body.fields.Status, dw[0].body.fields.Source],
    ["R0001", "door:1", 0, 1, "process", "office"], "one row per door, total 1, the word as the status");
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color),
    ["BT7 #FFFF00", "N7 #FFFF00"],
    "that door's own cell yellow, and N yellow as the copy of the Doors aggregate");
  assert.strictEqual(BOOK["Production"].v[kk(7, 72)], "CD", "and the code text is untouched");
  assert.strictEqual(BOOK["Production"].v[kk(7, 74)], "SS");
  assert(!CALLS.some(c => c.sheet === "Production" && c.kind === "values"),
    "not one value is written to the Production sheet, ever");
  assert.deepStrictEqual(logWrites()[0].values[0].slice(2, 6),
    ["R0001", "Door 1", "not started", "in fabrication"],
    "and the log line says it in words, because a door has no count");
  assert.strictEqual(SEQ[0], "record", "record first, as for every checkpoint: " + JSON.stringify(SEQ));
  assert.strictEqual(SEQ.filter(x => x === "fill").length, 2, "two fills: the door, and the Doors copy");
  pass("a door set to In fabrication: one record row, its own cell yellow, the code text untouched");

  reset();
  setDoorStatus(byId("R0001"), "door:1", "done");
  await settle(1400);
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color), ["BT7 #FFE699"],
    "Done paints that cell gold - and N is already yellow, so it is not written again");
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Status, "done");
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Done, 1);
  assert.deepStrictEqual(itemState(byId("R0001"), "drs"), { done: null, total: 2, status: "process" },
    "one of the two doors done: the Doors line is still in fabrication");

  reset();
  setDoorStatus(byId("R0001"), "door:3", "done");
  await settle(1400);
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color).sort(),
    ["BV7 #FFE699", "N7 #FFE699"], "the second door gold, and N follows it to gold");
  assert.deepStrictEqual(itemState(byId("R0001"), "drs"), { done: 2, total: 2, status: "done" });
  pass("every door done takes the Doors cell gold, as a copy of the aggregate and nothing else");

  reset();
  setDoorStatus(byId("R0001"), "door:1", "");
  await settle(1400);
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color).sort(),
    ["BT7 #FFFFFF", "N7 #FFFF00"], "Clear paints white, and N comes back off gold");
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Status, "", "the record says so first");
  assert.deepStrictEqual(logWrites()[0].values[0].slice(3, 6),
    ["Door 1", "done", "not started · Doors cell painted yellow"],
    "one line, saying both what the door did and what the Doors cell above it did");
  pass("Clear is as boring as Done: the record, then white, then the line");

  reset();
  setDoorStatus(byId("R0001"), "door:1", "");
  await settle(900);
  assert.strictEqual(listWrites().length, 0, "setting a door to what it already says writes nothing");
  setDoorStatus(byId("R0001"), "door:2", "done");
  await settle(900);
  assert.strictEqual(listWrites().length, 0, "and there is no door in slot 2 to set");
  assert.strictEqual(fills().length, 0);
  pass("nothing is written for a no-op or for a slot with no door");

  /* ---- the Doors LINE is still the office's own tick (amendment 1) ---- */
  reset();
  setItemProgress(byId("R0001"), "drs", 2);           // "All done" on the Doors line itself
  await settle(1400);
  assert.strictEqual(listWrites().length, 1, "the Doors line has a record row of its own");
  assert.strictEqual(listWrites()[0].body.fields.Item, "drs");
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color), ["N7 #FFE699"],
    "and N is painted gold by that tick, once - the aggregate does not write it a second time");
  assert.deepStrictEqual(itemState(byId("R0001"), "drs"), { done: 2, total: 2, status: "done" });
  /* ... and it STAYS done with a door un-ticked under it: the higher wins */
  reset();
  setDoorStatus(byId("R0001"), "door:3", "");
  await settle(1400);
  assert.deepStrictEqual(itemState(byId("R0001"), "drs"), { done: 2, total: 2, status: "done" },
    "a gold N stays done however little is ticked under it (owner, 2026-09-14)");
  assert.strictEqual(fills().filter(f => f.addr === "N7").length, 0,
    "and N is not touched: nothing under the line may take it back");
  pass("the Doors line is ticked in its own right, and the doors under it never lower it");

  /* a door painted by hand in Excel is adopted, like any other managed cell -
     the owner's "vice versa" */
  useJob(doorJob(), DMAP);
  setRecord([{ job: "R0001", item: "door:1", done: 0, total: 1, status: "" }]);
  setPainted({ R0001: { "door:1": "" } });
  settled(true);
  byId("R0001").doors[0].status = "done";             // the download shows it gold
  BOOK["Production"].fill[kk(7, 72)] = "#FFE699";     // ... and so does the live API
  vm.runInThisContext("LASTBY = 'the colleague';");
  reset();
  assert.strictEqual(await cpAdoptRun(), 1, "the hand-painted door is adopted");
  await settle(150);
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Status, "done");
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Source, "excel");
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Who, "the colleague");
  assert.ok(/Door 1/.test(logWrites()[0].values[0][3]) &&
            /adopted from Excel/.test(logWrites()[0].values[0][5]),
    "with a log line naming the door: " + logWrites()[0].values[0].slice(3, 6).join(" | "));
  assert.strictEqual(fills().length, 0, "nothing is painted: Excel already has the colour");
  pass("a door coloured in Excel by hand is adopted into the record, with who and what");

  /* N painted gold by hand IS adopted: since amendment 1 the Doors line has a
     record row of its own, so somebody colouring it in Excel is saying
     something the dashboard can record - "already golden means it finished" */
  useJob(doorJob(), DMAP);
  setRecord([]);
  setPainted({ R0001: { drs: "" } });
  settled(true);
  byId("R0001").cp.drs = "done";                      // N painted gold by hand
  BOOK["Production"].fill[kk(7, 14)] = "#FFE699";
  vm.runInThisContext("LASTBY = 'the colleague';");
  reset();
  assert.strictEqual(await cpAdoptRun(), 1, "a hand-painted N is adopted like any other cell");
  await settle(150);
  assert.strictEqual(cpSrvRow("R0001", "drs").fields.Status, "done");
  assert.deepStrictEqual(itemState(byId("R0001"), "drs"), { done: 2, total: 2, status: "done" },
    "so the Doors line reads done, with not one of its doors ticked");
  assert.strictEqual(fills().length, 0, "and nothing is painted: Excel already has the colour");
  pass("a gold N painted in Excel is adopted, and the line reads done with nothing under it");

  /* ---- amendment 11: a WHOLE GOLD ROW is not a sheet full of hand-paints --- */
  useJob(doorJob(), DMAP);
  setRecord([]);
  setPainted({});
  settled(true);
  global.__gold = Object.assign({}, byId("R0001"), { done: 1 });
  vm.runInThisContext("ALL = [__gold];");
  /* the parser forces every checkpoint of a gold row to done, which is what
     "the job is finished" means - and is not 30 people painting 30 cells */
  cpItems(byId("R0001")).forEach(x => { if (x.key.indexOf("door:") === 0) byId("R0001").doors.forEach(d => { d.status = "done"; }); });
  byId("R0001").cp.drs = "done"; byId("R0001").cp.win = "done";
  reset();
  assert.strictEqual(await cpAdoptRun(), 0, "not one cell of a gold row is adopted");
  assert.strictEqual(CALLS.length, 0, "and not one request goes out for it");
  pass("a whole gold row is never read back as a sheet full of hand-paints");

  /* ====== 18. THE DOORS' OWN IMPORT (amendment 2, 2026-09-14) =============
     The doors arrived three days after the one-time import had drained in
     every browser, so `cw_cpimported` is already set and the general pass will
     never look at a door cell again. Without a pass of its own, every coded
     door cell already carrying a colour would be met by the SAFEGUARD and
     recorded as somebody's hand-paint under whoever last saved the file. */
  const impDoorJob = () => mkJob({ win: "", drs: "", glass: {}, prod: {} },
    { wnd: 0, glass: {}, prods: [], drs: 4,
      doors: [{ slot: 1, code: "CD", status: "done" }, { slot: 2, code: "CD", status: "done" },
              { slot: 3, code: "SS", status: "process" }, { slot: 4, code: "DD", status: "" }] });
  useJob(impDoorJob(), DMAP);
  setRecord([]);                                   // the general marker is set, as it is in the wild
  doorsSettled(false);
  const genMarker = localStorage.getItem("cw_cpimported");
  reset();
  assert.strictEqual(await cpImportRun(), 0,
    "the general import drained on 11 September and never runs again");
  assert.strictEqual(await cpAdoptRun(), 0,
    "and the safeguard keeps its hands off the door cells while their own import is owed");
  assert.strictEqual(CALLS.length, 0, "not one request of any kind goes out for them");
  const imprt = await cpDoorImportRun();
  await settle(150);
  assert.strictEqual(imprt, 3, "the three coloured door cells are imported; the blank one is not");
  const dRows = listWrites();
  assert.ok(dRows.every(w => w.method === "POST"), "every row is created, none patched");
  assert.ok(dRows.every(w => w.body.fields.Source === "import" && w.body.fields.Who === "the sheet"),
    "every one says it came from the sheet, not from a person");
  assert.deepStrictEqual(dRows.map(w => w.body.fields.Item).sort(), ["door:1", "door:2", "door:3"]);
  assert.strictEqual(cpSrvRow("R0001", "door:1").fields.Done, 1, "a gold door imports as done");
  assert.strictEqual(cpSrvRow("R0001", "door:3").fields.Status, "process",
    "and a yellow one as in fabrication");
  assert.strictEqual(cpSrvRow("R0001", "door:4"), null, "a blank cell gets no row, as ever");
  assert.strictEqual(fills().length, 0, "the import paints nothing: the colours are already there");
  assert.ok(/Imported 3 doors from the sheet's colours/.test(logWrites()[0].values[0][5]),
    "one summary line, in plain words: " + logWrites()[0].values[0][5]);
  assert.ok(!logWrites().some(l => /adopted from Excel/.test(l.values[0][5])),
    "and NOT one line calling it somebody's hand-paint");
  assert.strictEqual(paintedOf("R0001", "door:1"), "done",
    "the colour it imported is remembered as the one in the cell");
  assert.ok(!!localStorage.getItem("cw_cpimported_doors"), "the doors' own marker is set");
  assert.strictEqual(localStorage.getItem("cw_cpimported"), genMarker,
    "and the general marker was never touched");
  assert.strictEqual(cpImportPending(), false,
    "nor was the colour fallback reopened for anything else");
  reset();
  assert.strictEqual(await cpDoorImportRun(), 0, "a second load imports nothing");
  assert.strictEqual(listWrites().length, 0, "and writes nothing at all");
  pass("the doors get one import of their own, behind their own marker, as the sheet");

  /* a browser still working through the GENERAL import has the doors in that
     pass already: the two must not both POST a row for the same cell */
  useJob(impDoorJob(), DMAP);
  setRecord([]);
  settled(false); doorsSettled(false);
  reset();
  assert.strictEqual(await cpDoorImportRun(), 0,
    "the doors pass stands aside while the general one is still running");
  assert.strictEqual(listWrites().length, 0);
  assert.ok(cpImportPlan(ALL, cpStored, 60).some(p => p.item.indexOf("door:") === 0),
    "because the general pass covers the doors itself");
  await cpImportRun();
  await settle(150);
  assert.ok(!!localStorage.getItem("cw_cpimported_doors"),
    "and settling the general one settles the doors with it");
  pass("the two imports never both create a row for one cell");
  settled(true); doorsSettled(true);

  /* ====== 19. THE AGGREGATE NEVER PAINTS WHITE (amendment 1) ==============
     The reviewer's reproduction: a gold M cell, PAINTED "done", nothing ticked
     under it - and cpRepaintPlan planted #FFFFFF into it, because the line had
     no record of its own and its aggregate said nothing. The owner: "already
     golden means it finished even if all the cell or component has no ticked.
     u should not change anything in the excel sheet." */
  useJob(mkIt("", { wnd: 9, drs: 0, glass: {} }));
  setRecord([]);                                   // no record row for Windows at all
  setPainted({ R0001: { win: "done" } });          // and the cell is gold
  reset();
  assert.strictEqual(await cpRepaintRun(), 0, "the gold M cell is left exactly as it is");
  assert.strictEqual(fills().length, 0, "not one fill");
  assert.deepStrictEqual(cpAggregatePlan("R0001").map(c => c.item), [],
    "and the aggregate plans nothing for it either: it only ever paints upwards");
  /* the office's OWN clear of that line is the one thing that whitens it */
  setRecord([{ job: "R0001", item: "win", done: 0, total: 9, status: "" }]);
  reset();
  assert.strictEqual(await cpRepaintRun(), 1, "the office's own Clear is repaired like any other");
  assert.deepStrictEqual([fills()[0].addr, fills()[0].color], ["M7", "#FFFFFF"]);
  pass("a gold M with nothing ticked under it is never whitened; the office's own Clear still is");

  /* clearing the LAST door leaves N exactly as the office left it */
  useJob(doorJob(), DMAP);
  setRecord([{ job: "R0001", item: "door:1", done: 1, total: 1, status: "done" }]);
  setPainted({ R0001: { "door:1": "done", drs: "done" } });
  reset();
  setDoorStatus(byId("R0001"), "door:1", "");
  await settle(1400);
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color), ["BT7 #FFFFFF"],
    "the door's own cell goes white; N is not written at all");
  assert.strictEqual(paintedOf("R0001", "drs"), "done", "so N is still gold as far as we know");
  pass("clearing the last door leaves the Doors cell as it was");

  console.log("\n" + n + " checks passed");
  process.exit(0);                 // the 45 s reconcile timer would hold the process open
})().catch(e => { console.error("FAIL", e); process.exit(1); });
