/* Offline test of "Glass colours: the floor's work reaches the Production
   sheet" (docs/specs/2026-09-10-glass-colours-two-way.md).

   This is the first feature in which something done on the floor changes the
   master sheet, so the sharpest assertions here are about what goes on the
   wire, over the whole run:

     · every write is a FILL, and nothing else. Not one value, not one row, not
       one formula, not one number format, reaches the Production sheet;
     · every fill lands in DG, TG, TUFF or NOT TUFF of the job's own row. ARCH,
       ASTRAGAL, FANCY and EXTRA are never written, in either direction, under
       any state of the floor's row;
     · no sheet but Production is ever written at all - no Dashboard Progress
       row, no Dashboard Log line, nothing;
     · a cell already showing the right colour is never written again, which is
       what keeps a ten-second poll from being a write storm.

   The floor's list is not faked here at all: the writer reads it out of memory
   (STATION_ITEMS), which is exactly what it does in the browser, so this suite
   has no reason to reach SharePoint. The tablet's own half of the feature -
   the tuff counter and the office's lock - is in test_station.js, which proves
   over its own run that the tablet touches no workbook.

   Graph is a fake fetch() over an in-memory workbook; nothing leaves the box,
   every name in here is made up.
   Run: node test_glasscolour.js                                             */
const fs = require("fs"), vm = require("vm"), assert = require("assert");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null),
                        setItem: (k, v) => { mem[k] = String(v); },
                        removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, addEventListener() {} };
global.performance = { now: () => Date.now() };
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
/* one element per selector, kept, so a test can read back what the footer was
   told to say rather than only that nothing threw */
const EL = {};
const el = sel => EL[sel] || (EL[sel] = stubEl());
global.document = {
  documentElement: stubEl(), body: stubEl(), createElement: () => stubEl(), activeElement: null,
  querySelector: el, querySelectorAll: () => [], addEventListener() {}, removeEventListener() {}
};

/* ---------- a tiny Excel behind a fake Graph ----------
   The same shape as test_checkpoints.js': cell values and cell fills per
   sheet, addressed by A1, and every call remembered so the assertions can be
   about the requests rather than about the outcome. */
const BOOK = {};
const CALLS = [];
let FAIL_FILL = 0;                            // n fills to refuse with a 403
let DROP_BATCH_ID = 0;                        // n $batch replies to send one response short
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

/* ---- and the floor's own list, behind the same fake Graph ----
   The colour writer reads the floor's rows out of memory and never asks
   SharePoint for them, which is why this suite had no list at all until the
   office learned to clear one (docs/specs/2026-09-10-office-clears-the-floor.md).
   That clear is the ONE request the office makes to this list from here, so
   the list is served rather than stubbed: what matters is the exact body that
   goes on the wire, and the fact that nothing else ever does. */
const FSITE = "example.sharepoint.test,aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee,ffffffff-0000-1111-2222-333333333333";
const GLASS_LIST_ID = "list-glass-station";
/* ---- and, since 2026-09-11, the record of checkpoint status ----
   `Dashboard progress` lives in the WORKBOOK's site, not the floor's, and the
   drawer's own ticks are written to it before the colour is painted. Served
   rather than stubbed for the same reason the floor's list is: what matters is
   the exact body on the wire and the order it goes in. */
const WSITE = "workbook-site-0001";
const CPLIST_ID = "list-dashboard-progress";
const CPBASE = "/sites/" + WSITE + "/lists/" + CPLIST_ID + "/items";
let CPSRV = [];                               // the record, as the server holds it
let CPNEXT = 500;
const LISTS_PATH = "/sites/" + FSITE + "/lists";
const ITEMS_PATH = LISTS_PATH + "/" + GLASS_LIST_ID + "/items/";
const LISTWRITES = [];                        // every write this suite sends to that list
let FAIL_LIST = 0;                            // n list writes to refuse with a 403
let FAIL_CPLIST = 0;                          // ... and n writes to the RECORD, since step 3

function route(method, path, body) {
  if (path === LISTS_PATH + "?$select=id,displayName")
    return ok({ value: [{ id: GLASS_LIST_ID, displayName: "Glass station" }] });
  if (path === "/sites/" + WSITE + "/lists?$select=id,displayName")
    return ok({ value: [{ id: CPLIST_ID, displayName: "Dashboard progress" }] });
  if (path.indexOf(CPBASE) === 0) {
    const rest = path.slice(CPBASE.length);
    if (method === "GET" && (rest === "" || rest.charAt(0) === "?"))
      return ok({ value: CPSRV.map(x => ({ id: x.id, fields: x.fields })) });
    if (method === "POST" && rest === "") {
      if (FAIL_CPLIST) { FAIL_CPLIST--; return { status: 403, body: { error: { code: "AccessDenied" } } }; }
      const id = String(CPNEXT++);
      CPSRV.push({ id: id, fields: (body && body.fields) || {} });
      return ok({ id: id });
    }
    if (method === "GET" && rest.indexOf("/delta") === 0)
      return ok({ value: CPSRV.map(x => ({ id: x.id, fields: x.fields })),
                  "@odata.deltaLink": "https://graph.microsoft.com/v1.0" + CPBASE + "/delta?$skiptoken=1" });
    const mfd = /^\/([^/?]+)\/fields$/.exec(rest);
    if (mfd && method === "PATCH") {
      if (FAIL_CPLIST) { FAIL_CPLIST--; return { status: 403, body: { error: { code: "AccessDenied" } } }; }
      const it = CPSRV.filter(x => x.id === mfd[1])[0];
      if (!it) return { status: 404, body: { error: { code: "itemNotFound" } } };
      Object.keys(body || {}).forEach(k => { it.fields[k] = body[k]; });
      return ok({ id: it.id });
    }
    const mdl = /^\/([^/?]+)$/.exec(rest);
    if (mdl && method === "DELETE") { CPSRV = CPSRV.filter(x => x.id !== mdl[1]); return ok({}); }
    return { status: 404, body: { error: "no route " + method + " " + path } };
  }
  if (path.indexOf(ITEMS_PATH) === 0 && /\/fields$/.test(path)) {
    const id = path.slice(ITEMS_PATH.length).replace(/\/fields$/, "");
    if (method !== "PATCH") return { status: 405, body: { error: "no route " + method + " " + path } };
    if (FAIL_LIST) { FAIL_LIST--; return { status: 403, body: { error: { code: "AccessDenied" } } }; }
    LISTWRITES.push({ id: id, fields: body });
    return ok({ id: id, fields: body });
  }
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
    if (FAIL_FILL) { FAIL_FILL--; return { status: 403, body: { error: { code: "AccessDenied" } } }; }
    for (let r = A.r1; r <= A.r2; r++) for (let c = A.c1; c <= A.c2; c++) s.fill[kk(r, c)] = body.color;
    CALLS.push({ method, sheet: name, kind: "fill", addr: mr[1], color: body.color });
    return ok({});
  }
  if (method === "PATCH" && tail === "") {
    (body.values || []).forEach((line, ri) => line.forEach((x, ci) => { s.v[kk(A.r1 + ri, A.c1 + ci)] = x; }));
    CALLS.push({ method, sheet: name, kind: body.values ? "values" : "numberFormat", addr: mr[1], values: body.values });
    return ok({});
  }
  if (method === "PATCH") { CALLS.push({ method, sheet: name, kind: "format", addr: mr[1] }); return ok({}); }
  return { status: 404, body: { error: "no route " + method + " " + path } };
}
const ALLREQ = [];
global.fetch = async (url, init) => {
  const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  let res;
  if (path === "/$batch") {
    /* every request inside a batch is a request: the "over the whole run"
       assertions must see them, not the envelope */
    (body.requests || []).forEach(q => ALLREQ.push({ method: q.method, path: q.url, body: q.body }));
    let rs = body.requests.map(q => Object.assign({ id: q.id }, route(q.method, q.url, q.body)));
    /* Graph can answer a batch without answering every request in it. That is
       not a success for the ones it left out. */
    if (DROP_BATCH_ID) { DROP_BATCH_ID--; rs = rs.slice(1); }
    res = ok({ responses: rs });
  } else {
    ALLREQ.push({ method: init.method, path: path, body: body });
    res = route(init.method, path, body);
  }
  return { ok: res.status < 400, status: res.status, text: async () => JSON.stringify(res.body),
           arrayBuffer: async () => new ArrayBuffer(0) };
};

/* ---------- load the app the way index.html does ---------- */
const src = f => fs.readFileSync(__dirname + "/" + f, "utf8");
const run = f => vm.runInThisContext(src(f), { filename: f });
run("parser.js");
run("graph.js");
global.CW = window.CW;
CW._setToken(() => "t");
CW._setFile({ base: "/x/workbook", content: "/x/content", meta: "/x", siteId: WSITE });
CW._setStationSite(FSITE);                    // the floor's lists, already found
run("checkpoints.js");
global.CP = window.CP;
/* CHANGED 2026-09-11 (spec: status-list-is-truth, step 2). Throughout this
   suite, `CP.cpSetProgress({})` is how a test says "the office has no dated
   record of this job". Since the record moved from the `Dashboard Progress`
   sheet to the `Dashboard progress` list, that sentence has to clear both, or
   an office row left behind by the test before would keep the writer standing
   down (officeSettling) in the test after. Clearing it with a map still sets
   the sheet's counts exactly as it always did. */
const realSetProgress = CP.cpSetProgress;
CP.cpSetProgress = map => {
  realSetProgress(map);
  if (!map || !Object.keys(map).length) A("cpRowsSet({}); CP_ITEMS = [];");
};
run("station-core.js");
global.ST = window.ST;
run("app.js");

const TOASTS = [];
global.toast = (m, err) => TOASTS.push({ m: String(m), err: !!err });

const A = vm.runInThisContext.bind(vm);        // reach app.js' own let-bound state
const settle = ms => new Promise(r => setTimeout(r, ms == null ? 60 : ms));
const reset = () => { CALLS.length = 0; TOASTS.length = 0; };
/* fills on the PRODUCTION sheet, which is what every assertion here means by
   one. Since 2026-09-11 the writer also leaves a Dashboard Log line, and
   creating that sheet fills its header row once - a fill, on a sheet none of
   these tests are about. */
const fills = () => CALLS.filter(c => c.kind === "fill" && c.sheet === "Production");
/* the backoff record is a module-level const object, so a test clears it the
   same way test_checkpoints.js ages a hold: by reaching in */
const clearFail = () => A("Object.keys(GLASSC_FAIL).forEach(k => delete GLASSC_FAIL[k]); setGlassFoot();");
/* the office's record of the clears IT made, which outlives one scene the way
   PENDING and CHANGES do and has to be reset with them */
const forgetClears = () => A("if (typeof OFFICE_FLOOR_AT !== 'undefined') " +
  "Object.keys(OFFICE_FLOOR_AT).forEach(function (k) { delete OFFICE_FLOOR_AT[k]; });");
/** pretend this job's last failure was `ms` ago, so a backoff can be waited
    out without waiting it out */
const ageFail = (id, ms) => A("GLASSC_FAIL['" + id + "'].at -= " + ms);
const footWord = () => EL["#stationfeed"].textContent;
const footWhy = () => EL["#stationfeed"].title;
const fillOn = addr => { const f = fills().filter(c => c.addr === addr); return f.length ? f[f.length - 1].color : null; };

/* The four columns this feature may write, and the four it may never touch,
   at the addresses the live sheet uses: DG is AY, TG is AZ, TUFF is BA and
   NOT TUFF is BB. */
const COL = { dg: 51, tg: 52, tuff: 53, "not tuff": 54,
              arch: 55, astragal: 56, fancy: 57, extra: 58 };
const PMAP = { qty: { wnd: 13, drs: 14 }, glass: Object.assign({}, COL), prod: {}, prodOrder: [] };
const AT = { dg: "AY7", tg: "AZ7", tuff: "BA7", "not tuff": "BB7",
             arch: "BC7", astragal: "BD7", fancy: "BE7", extra: "BF7" };
const GOLD = "#FFE699", YELLOW = "#FFFF00", WHITE = "#FFFFFF";

const mkJob = o => Object.assign({
  id: "R7001", cust: "Customer One", area: "", eir: "", off: "", colour: "", ph3: "", ph: "",
  flag: "", flagHex: "", wnd: 0, drs: 0, glass: { dg: 4, tg: 4, tuff: 11, "not tuff": 8 },
  prods: [], notes: [], sheets: ["Production"], src: { Production: 7 },
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cat: "active", blk: 4, seq: 1, stage: "floor", done: 0, urg: 0,
  cp: { win: "", drs: "", glass: {}, prod: {} }
}, o || {});

/** One list row for the floor's board, in the shape the list really holds. */
const row = o => ({ id: "700", fields: Object.assign({
  Title: "R7001", Job: "R7001", Customer: "Customer One", GlassType: "GLASS",
  Total: 8, TuffTotal: 11, Seq: 1, Active: "Yes", OfficeDone: "No",
  Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0, DoneAt: "", DoneBy: ""
}, o || {}) });

/** Put one job on the fake Production sheet at row 7, with the fills its
    checkpoint statuses imply, and hand the app that job plus one floor row.
    STATION_ITEMS is always a NEW array: stationRecords() caches on its
    identity, exactly as the browser does. */
function scene(job, listRow, opts) {
  BOOK["Production"] = { v: {}, fill: {} };
  BOOK["Production"].v[kk(7, 3)] = job.id;
  global.__j = job; global.__m = PMAP; global.__items = listRow ? [listRow] : [];
  A("PENDING = {}; savePending(); ALL = [__j]; PRODMAP = __m; CHANGES = []; state.sel = null;" +
    "STATION_ITEMS = __items; STATION_OK = true; BLOCKNAMES = [];");
  clearFail();
  forgetClears();
  if (!opts || !opts.keepProgress) CP.cpSetProgress({});
  seedRecord(job);
  return byId(job.id);
}
/** CHANGED 2026-09-11 (spec: status-list-is-truth, step 2). Checkpoint status
    is the `Dashboard progress` record now, not the Excel colour, so a fixture
    that describes a job by its colours is put on the record here - exactly
    what the one-time import does on the first load after the switch-over. The
    server's copy is seeded with it too, so a click PATCHes a row that is
    really there. `Source` is "import", which deliberately does NOT count as
    the office having spoken: the office's stamp is still whatever the test
    gives it (a Dashboard Progress row, a Log line, or a click). */
function seedRecord(job) {
  CPSRV = [];
  A("PAINTED = {}; savePainted();");
  const items = [];
  cpItems(job).forEach(it => {
    const st = cpFileStatus(job, it.key);
    if (st !== "done" && st !== "process") return;
    const id = String(CPNEXT++);
    const f = cpRowFields(job.id, it.key, st === "done" ? it.total : 0, it.total, st,
                          "the sheet", "2026-09-01T09:00:00Z", "import");
    CPSRV.push({ id: id, fields: f });
    items.push({ id: id, fields: JSON.parse(JSON.stringify(f)) });
  });
  global.__cpitems = items;
  A("CP_ITEMS = __cpitems; cpListRebuild(); CP_LIST_OK = true; CP_LIST_WHY = ''; " +
    "STATION_FEEDS.progress.token = null; CP_IMPORTED = '2026-09-11T09:00:00Z'; cpImportCheck();");
}
/** The office's own dated record of one tick - what a `cp` hold used to be. */
function officeRow(job, item, done, total, when) {
  const st = done <= 0 ? "" : (done >= total ? "done" : "process");
  global.__cpone = { job: job, item: item, done: done, total: total, st: st, when: when };
  A("cpRowPut(__cpone.job, __cpone.item, cpRowsFrom([{ id: 'office-' + __cpone.item, " +
    "fields: cpRowFields(__cpone.job, __cpone.item, __cpone.done, __cpone.total, __cpone.st, " +
    "'the admin', __cpone.when, 'office') }])[cpTitle(__cpone.job, __cpone.item)]);");
}
/** The colours the sheet is now showing, by glass type. */
const sheetNow = () => {
  const out = {};
  Object.keys(AT).forEach(k => { out[k] = BOOK["Production"].fill[kk(7, COL[k])]; });
  return out;
};
/** An ISO stamp for a local wall-clock time, so a floor stamp and an office
    one can be made to describe the same instant to the millisecond. */
const isoAt = (h, mi, s) => new Date(2026, 8, 10, h, mi, s || 0).toISOString();
const officeAt = (h, mi) => "2026-09-10 " + (h < 10 ? "0" : "") + h + ":" + (mi < 10 ? "0" : "") + mi;

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* load() is stubbed for the whole run. It really does re-download and
     re-parse the workbook, and this suite's fake Graph serves no file to
     download - and since a hold that the file has not caught up with now asks
     to be re-read (section 12), timers that call it fire in the middle of
     other tests. What every test here is about is what the app DECIDES, so
     what a re-read would do is not the point; that a re-read was ASKED FOR is,
     and it is recorded. */
  global.__loads = [];
  A("load = async function (r) { __loads.push(String(r == null ? '' : r)); };");
  const loadsAsked = () => global.__loads.length;
  /* and the floor's ten-second poll is stopped for the same reason: this fake
     serves the Glass station list's ITEMS from nowhere, so the poll can only
     ever 404 and mark the list missing - at whatever moment its timer happens
     to land, which made this suite's own results depend on the clock. The poll
     is test_station.js' business; here the list is read out of memory. */
  A("stationTick = function () {};" +
    "if (stationPollT) { clearTimeout(stationPollT); stationPollT = null; }");

  /* ================= 1. the colour rule, on its own ================= */
  const rec = o => ST.jobRecord([row(o)], "R7001");
  const cols = o => ST.glassColours(rec(o));

  assert.deepStrictEqual(cols({}), { dg: "", tg: "", tuff: "", "not tuff": "" },
    "a job the floor has done nothing to has no colour in any of the four");
  assert.deepStrictEqual(cols({ Cut: 8 }),
    { dg: "", tg: "", tuff: "", "not tuff": "" }, "cut but not hotmelted is still nothing");
  assert.deepStrictEqual(cols({ Cut: 8, Hotmelt: 4 }),
    { dg: "", tg: "", tuff: "", "not tuff": "" }, "and half hotmelted is nothing either");
  assert.deepStrictEqual(cols({ Cut: 8, Hotmelt: 8 }),
    { dg: "yellow", tg: "yellow", tuff: "", "not tuff": "yellow" },
    "cutting AND hotmelting complete: DG, TG and NOT TUFF go yellow");
  assert.deepStrictEqual(cols({ Tuff: 11 }),
    { dg: "", tg: "", tuff: "yellow", "not tuff": "" },
    "and the tuff count on its own is the one thing that makes TUFF yellow");
  assert.deepStrictEqual(cols({ Tuff: 10 }),
    { dg: "", tg: "", tuff: "", "not tuff": "" }, "ten of eleven tuff is not complete");
  assert.deepStrictEqual(cols({ Cut: 8, Hotmelt: 8, Tuff: 11 }),
    { dg: "yellow", tg: "yellow", tuff: "yellow", "not tuff": "yellow" });
  pass("yellow: cutting and hotmelting for DG, TG and NOT TUFF; the tuff count for TUFF");

  /* GOLD REQUIRES GLAZING, for every column, with no exception */
  ST.COLOUR_TYPES.forEach(t => {
    assert.notStrictEqual(cols({ Cut: 8, Hotmelt: 8, Tuff: 11, Glazed: 7 })[t], "gold",
      t + " is not gold while one glass is still unglazed");
    assert.strictEqual(cols({ Glazed: 8 })[t], "gold",
      t + " is gold the moment glazing is complete");
  });
  assert.deepStrictEqual(cols({ Glazed: 8 }),
    { dg: "gold", tg: "gold", tuff: "gold", "not tuff": "gold" },
    "glazing is the gate for everything, whatever the other counters say");
  assert.deepStrictEqual(cols({ Total: 0, Glazed: 0, Cut: 0, Hotmelt: 0 }),
    { dg: "", tg: "", tuff: "", "not tuff": "" },
    "and a row with no glasses on it has finished nothing: 0 of 0 is not complete");
  pass("gold requires glazing, for every column, and no other counter can reach it");

  /* it walks back down as well as up */
  assert.strictEqual(cols({ Cut: 8, Hotmelt: 8, Glazed: 8 }).dg, "gold");
  assert.strictEqual(cols({ Cut: 8, Hotmelt: 8, Glazed: 0 }).dg, "yellow", "gold to yellow");
  assert.strictEqual(cols({ Cut: 8, Hotmelt: 0, Glazed: 0 }).dg, "", "yellow to nothing at all");
  assert.strictEqual(cols({ Cut: 0, Hotmelt: 0, Glazed: 0 }).dg, "");
  assert.strictEqual(cols({ Tuff: 11, Glazed: 8 }).tuff, "gold");
  assert.strictEqual(cols({ Tuff: 11, Glazed: 0 }).tuff, "yellow");
  assert.strictEqual(cols({ Tuff: 0, Glazed: 0 }).tuff, "");
  pass("the rule is a reading of the counters, so it reverses on its own: gold, yellow, blank");

  /* the floor's stamp, which is what the contest below is decided on */
  assert.strictEqual(ST.floorStamp(rec({})), "", "a row the floor never tapped has no stamp");
  assert.strictEqual(ST.floorStamp(rec({ DoneAt: isoAt(14, 5) })), isoAt(14, 5));
  assert.strictEqual(ST.floorStamp(rec({ DoneAt: isoAt(14, 5), CutAt: isoAt(15, 0) })), isoAt(15, 0),
    "and the newest of everything on the row wins, whichever column carries it");
  /* by PARSED TIME, not as text. Every column it reads is editable by hand in
     SharePoint, and a text maximum let one junk character win the comparison
     ("zzz" beats every real stamp), answer a time nothing could read, and
     switch the whole feature off for that job in silence. */
  assert.strictEqual(ST.floorStamp(rec({ DoneAt: isoAt(15, 0), CutAt: "zzz" })), isoAt(15, 0),
    "a stamp that will not parse is passed over rather than winning the comparison");
  assert.strictEqual(stampMs(ST.floorStamp(rec({ DoneAt: isoAt(15, 0), CutAt: "zzz" }))),
    stampMs(isoAt(15, 0)), "so the job is still coloured, which is what the text maximum stopped");
  assert.strictEqual(ST.floorStamp(rec({ DoneAt: "", CutAt: "not a date", GlazedAt: "  " })), "",
    "and a row on which NOTHING parses has no stamp at all, which fails closed as it should");
  assert.strictEqual(ST.floorStamp(rec({ DoneAt: isoAt(9, 0), GlazedAt: isoAt(11, 0) })), isoAt(11, 0),
    "the newest by time, whichever column it is in");
  pass("the floor's stamp is the newest READABLE thing on their own row, or nothing at all");

  /* ================= 2. the three stamp formats, and the missing seconds ==== */
  const at = (h, mi, s, ms) => new Date(2026, 8, 10, h, mi, s || 0, ms || 0).getTime();
  assert.strictEqual(stampMs(""), 0);
  assert.strictEqual(stampMs(null), 0);
  assert.strictEqual(stampMs("not a date"), 0);
  assert.strictEqual(stampMs(isoAt(14, 3)), at(14, 3),
    "the floor's ISO stamp is UTC, carries seconds, and is taken exactly");
  assert.strictEqual(stampMs("2026-09-10 14:03:07"), at(14, 3, 7),
    "a local stamp that carries seconds is taken exactly too");
  /* THE ONE THAT MATTERED. nowStamp() writes no seconds, so an office tick at
     14:03:40 is recorded as 14:03 - up to 59 s EARLY. Read literally, a floor
     tap at 14:03:20 would look like the later action and paint the office's own
     tick out. A stamp that only names a minute therefore covers that minute. */
  assert.strictEqual(stampMs("2026-09-10 14:03"), at(14, 3, 59, 999),
    "Dashboard Progress names a minute and no seconds, so it means the end of that minute");
  assert.strictEqual(stampMs("10/09/2026 14:03"), at(14, 3, 59, 999),
    "and the Log comes back out of the workbook the other way round, read the same way");
  assert.strictEqual(stampMs("2026-09-10 14:03"), stampMs("10/09/2026 14:03"),
    "so the same tick reads the same whichever of the two records it is read from");
  assert.ok(stampMs("2026-09-10 14:03") > stampMs(isoAt(14, 3, 59)),
    "anything inside the minute loses to it");
  assert.ok(stampMs("2026-09-10 14:03") < stampMs(isoAt(14, 4, 0)),
    "and the next minute beats it");
  pass("all three stamp formats read as the LATEST instant they can mean - which is what the missing seconds require");

  /* ================= 3. the writer: what reaches the sheet ================= */
  let j = scene(mkJob(), row({ Cut: 8, Hotmelt: 8, DoneAt: isoAt(14, 0), DoneBy: "Person A" }));
  reset();
  assert.strictEqual(await glassColourRun(), 1, "one job to paint");
  await settle();
  assert.deepStrictEqual(fills().map(c => c.addr).sort(), ["AY7", "AZ7", "BB7"],
    "DG, TG and NOT TUFF, and nothing else: TUFF is not yellow because tuff was not counted");
  fills().forEach(c => assert.strictEqual(c.color, YELLOW));
  assert.strictEqual(sheetNow().tuff, undefined, "TUFF was never written at all");
  pass("the floor cut and hotmelted a job, and three glass cells went yellow in the sheet");

  ["arch", "astragal", "fancy", "extra"].forEach(t =>
    assert.strictEqual(sheetNow()[t], undefined, t + " was not touched"));
  assert.ok(!fills().some(c => ["BC7", "BD7", "BE7", "BF7"].indexOf(c.addr) >= 0),
    "and no fill was even addressed at one of them");
  pass("ARCH, ASTRAGAL, FANCY and EXTRA are not written: the office ticks those by hand");

  /* IDEMPOTENCE. The office polls the floor's list every ten seconds. If the
     writer could rewrite a cell that is already right, this feature would be a
     write storm against the workbook rather than a colour. */
  reset();
  assert.strictEqual(await glassColourRun(), 0, "nothing to do: the write is still held");
  await settle();
  assert.strictEqual(CALLS.length, 0, "and not one request went out");
  for (let i = 0; i < 10; i++) await glassColourRun();
  await settle();
  assert.strictEqual(CALLS.length, 0, "ten more turns of the poll: still not one request");
  pass("a cell already showing the right colour is never written again - ten polls, no requests");

  /* CHANGED at step 3. What stood here proved the `gc` hold: the colour was
     kept in PENDING until the downloaded file agreed with it, survived a stale
     download, and was let go the moment the file caught up. There is no hold -
     the writer puts the floor's colour on the RECORD and the record is what
     every screen reads, so there is nothing to hold it against and no download
     in the loop at all. */
  assert.strictEqual(A("Object.keys(PENDING).length"), 0, "nothing is held anywhere");
  assert.deepStrictEqual(itemState(byId("R7001"), "glass:dg"),
    { done: null, total: 4, status: "process" },
    "the record says yellow, so the drawer and the row read yellow at once - with no " +
    "per-type count, because the floor counts one combined DG + TG number");
  const dgRec = cpRow("R7001", "glass:dg");
  assert.strictEqual(dgRec.source, "floor", "recorded as the floor's work");
  assert.ok(dgRec.who, "with a name on it: " + dgRec.who);
  assert.strictEqual(stampMs(dgRec.when), stampMs(ST.floorStamp(stationForJob("R7001"))),
    "and the floor's own action time, not now");
  assert.strictEqual(paintedOf("R7001", "glass:dg"), "process",
    "and the dashboard knows it painted that cell, so the safeguard leaves it alone");
  /* a stale download changes nothing, because nothing reads it */
  A("ALL = applyPending([__j], true)");
  assert.deepStrictEqual(itemState(byId("R7001"), "glass:dg"),
    { done: null, total: 4, status: "process" }, "a stale download changes nothing");
  global.__j2 = mkJob({ cp: { win: "", drs: "", glass: { dg: "process", tg: "process", "not tuff": "process" }, prod: {} } });
  A("ALL = applyPending([__j2], true)");
  assert.strictEqual(A("Object.keys(PENDING).length"), 0, "and the download catching up changes nothing either");
  reset();
  assert.strictEqual(await glassColourRun(), 0);
  assert.strictEqual(CALLS.length, 0, "still nothing to write: the record already says it");
  pass("the floor's colour is on the record at once, and no download decides anything about it");

  /* ================= 4. reversal, and no flicker ================= */
  j = scene(mkJob({ cp: { win: "", drs: "", glass: { dg: "done", tg: "done", tuff: "done", "not tuff": "done" }, prod: {} } }),
            row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11, DoneAt: isoAt(14, 0) }));
  reset();
  assert.strictEqual(await glassColourRun(), 0, "the sheet is already gold and the floor says gold");
  /* now the floor taps glazing back down: gold has to become yellow */
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 0, Tuff: 11, DoneAt: isoAt(15, 0) })];
  A("STATION_ITEMS = __items");
  reset();
  assert.strictEqual(await glassColourRun(), 1);
  await settle();
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AY7", YELLOW], ["AZ7", YELLOW], ["BA7", YELLOW], ["BB7", YELLOW]],
    "all four walk back from gold to yellow");
  /* the record now says yellow. The download is still 36 s behind and still
     says gold - which is exactly the moment a flicker used to happen. */
  A("ALL = applyPending(ALL, true)");
  assert.strictEqual(itemState(byId("R7001"), "glass:dg").status, "process",
    "the screen keeps the new colour rather than flickering back to the old one");
  assert.strictEqual(A("Object.keys(PENDING).length"), 0, "and nothing had to be held to do it");
  reset();
  assert.strictEqual(await glassColourRun(), 0, "and the writer does not paint it again either");
  assert.strictEqual(CALLS.length, 0);
  pass("a reversal reads back at once from the record: no flicker, and no second write");

  /* and all the way back to nothing */
  global.__items = [row({ Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0, DoneAt: isoAt(16, 0) })];
  A("PENDING = {}; savePending(); STATION_ITEMS = __items; ALL = applyPending([__j], true)");
  reset();
  assert.strictEqual(await glassColourRun(), 1);
  await settle();
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AY7", WHITE], ["AZ7", WHITE], ["BA7", WHITE], ["BB7", WHITE]],
    "a job tapped back to nought leaves its glass cells with no colour on them");
  assert.strictEqual(cpRow("R7001", "glass:dg").status, "",
    "recorded as nothing, which is a state like any other");
  assert.strictEqual(itemState(byId("R7001"), "glass:dg").status, "",
    "and the drawer reads it as untouched");
  assert.strictEqual(A("Object.keys(PENDING).length"), 0, "with nothing held for any of it");
  pass("gold, yellow, blank: white is written rather than the fill cleared, so the row keeps its own look");

  /* a cell that already has no colour is not painted white for the sake of it */
  global.__j3 = mkJob();                        // no cp statuses at all: blank cells
  global.__items = [row({ Cut: 1, DoneAt: isoAt(16, 0) })];
  A("PENDING = {}; savePending(); ALL = [__j3]; STATION_ITEMS = __items");
  seedRecord(global.__j3);                      // a clean record with it
  reset();
  assert.strictEqual(await glassColourRun(), 0, "the floor has done something, but nothing that shows yet");
  assert.strictEqual(CALLS.length, 0, "so no cell is painted white to say so");
  pass("a blank cell that should be blank is left alone: white and no fill are the same nothing");

  /* ================= 5. a job the floor has never tapped ================= */
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11, DoneAt: "" })];
  A("PENDING = {}; savePending(); ALL = [__j3]; STATION_ITEMS = __items");
  seedRecord(global.__j3);
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "the counters are the office's own seed echoed back, not the floor's work");
  assert.strictEqual(CALLS.length, 0);
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11, DoneAt: isoAt(14, 0) })];
  A("STATION_ITEMS = __items");
  assert.strictEqual(await glassColourRun(), 1, "and the first tap is what makes it the floor's to say");
  await settle();
  pass("a row with no floor stamp on it is never painted: the seed cannot come back as a colour");

  /* a job that was never fed to the floor at all */
  A("PENDING = {}; savePending(); STATION_ITEMS = []; ALL = [__j3]");
  reset();
  assert.strictEqual(await glassColourRun(), 0);
  assert.strictEqual(CALLS.length, 0);
  /* and a job marked ready to deliver: a gold row is finished work */
  global.__jd = mkJob({ done: 1 });
  global.__items = [row({ Glazed: 8, DoneAt: isoAt(14, 0) })];
  A("ALL = [__jd]; STATION_ITEMS = __items");
  assert.strictEqual(await glassColourRun(), 0, "a gold row is left whole");
  assert.strictEqual(CALLS.length, 0);
  pass("a job never fed to the floor, and a job marked ready, are both left entirely alone");

  /* ================= 5b. the office's own yellow survives the round trip ====
     OWNER DECISION, 2026-09-10. Found by the rehearsal against the real file:
     the one yellow glass cell in the whole workbook was being erased to blank
     by a feature that had learned nothing from doing it. A yellow item seeded
     the floor's row at nought (cpStored holds no count for one), the floor's
     reading was "blank", and the first tap on that job painted the office's
     own mark out. So an office yellow is now read as what yellow MEANS under
     this feature: cutting and hotmelting complete, glazing not.

     The test is the whole round trip - what the office ticked, through the
     seed, onto the floor's row, and back out as a colour. */
  const seedTrip = (counts, glass) => {
    const job = mkJob({ glass: glass });
    const sl = ST.glassSlice([job], ["a", "b", "c", "d", "In production"], () => counts);
    const fields = Object.assign({ Title: job.id, Job: job.id, Total: sl[0].total,
        TuffTotal: sl[0].tuffTotal, Active: "Yes", OfficeDone: "No" },
      ST.seedFields(sl[0]));
    /* the floor then taps something, which is the only thing that lets the
       colour writer look at the row at all */
    fields.DoneAt = isoAt(15, 0);
    return { seed: sl[0].seed,
             colour: ST.glassColours(ST.jobRecord([{ id: "700", fields: fields }], job.id)) };
  };
  const yellowOnly = seedTrip([{ type: "dg", total: 4, status: "process", done: 0 }], { dg: 4 });
  assert.deepStrictEqual(yellowOnly.seed, { cut: 4, hotmelt: 4, glazed: 0 },
    "an office yellow seeds cut and hotmelt to the total, and glazed to nought");
  assert.strictEqual(yellowOnly.colour.dg, "yellow",
    "and the cell round-trips to YELLOW - not blank, which is what it used to do");
  assert.notStrictEqual(yellowOnly.colour.dg, "gold",
    "and not gold either: glazing has not happened and the seed does not pretend it has");
  pass("an office yellow comes back as yellow: the mark the office made survives being fed to the floor");

  const goldOnly = seedTrip([{ type: "dg", total: 4, status: "done", done: 4 }], { dg: 4 });
  assert.deepStrictEqual(goldOnly.seed, { cut: 4, hotmelt: 4, glazed: 4 },
    "a gold item still seeds all three, exactly as before");
  assert.strictEqual(goldOnly.colour.dg, "gold", "and still round-trips to gold");
  const bothYellow = seedTrip(
    [{ type: "dg", total: 4, status: "process" }, { type: "tg", total: 4, status: "process" }],
    { dg: 4, tg: 4 });
  assert.deepStrictEqual(bothYellow.seed, { cut: 8, hotmelt: 8, glazed: 0 });
  assert.strictEqual(bothYellow.colour.dg, "yellow");
  assert.strictEqual(bothYellow.colour.tg, "yellow", "two yellow items, both still yellow");
  const goldAndYellow = seedTrip(
    [{ type: "dg", total: 4, status: "done" }, { type: "tg", total: 4, status: "process" }],
    { dg: 4, tg: 4 });
  assert.deepStrictEqual(goldAndYellow.seed, { cut: 8, hotmelt: 8, glazed: 0 });
  assert.strictEqual(goldAndYellow.colour.dg, "yellow",
    "one gold and one yellow comes back yellow for both - the floor's row holds one number " +
    "for DG and TG together, so it can only say one thing, and the honest one is the lower");
  pass("gold still round-trips to gold; a gold-and-yellow job says yellow, which is the lower of the two");

  /* the case that CANNOT be carried, said out loud rather than fudged: DG
     yellow and TG blank. Seeding cut and hotmelt to the total there would tell
     the FLOOR that glass which still needs cutting is cut. */
  const mixed = seedTrip(
    [{ type: "dg", total: 4, status: "process", done: 0 }, { type: "tg", total: 4, status: "" }],
    { dg: 4, tg: 4 });
  assert.deepStrictEqual(mixed.seed, { cut: 0, hotmelt: 0, glazed: 0 },
    "a part-yellow job seeds at the office's own count, and a countless yellow is nought");
  assert.strictEqual(mixed.colour.dg, "",
    "so that job's yellow is still lost once the floor taps - the known limit, and the price " +
    "of never telling the floor that uncut glass is cut");
  pass("a job with one type yellow and another blank keeps the office's count, and its yellow is not carried");

  /* and none of it can reach a row the floor has already tapped */
  const touched = ST.feedPlan(
    ST.glassSlice([mkJob({ glass: { dg: 4 } })], ["a", "b", "c", "d", "In production"],
                  () => [{ type: "dg", total: 4, status: "process" }]),
    [{ id: "700", fields: { Title: "R7001", Job: "R7001", Customer: "Customer One",
        GlassType: "GLASS", Total: 4, TuffTotal: 0, Seq: 1, Active: "Yes", OfficeDone: "No",
        Cut: 1, Hotmelt: 0, Glazed: 0, DoneAt: isoAt(15, 0), DoneBy: "Person A" } }],
    { at: isoAt(16, 0), by: "the office" });
  const touchedFields = (touched.patches[0] || { fields: {} }).fields;
  ST.SEED_FIELDS.forEach(k => assert.ok(!(k in touchedFields),
    "a row the floor has tapped is never re-seeded, and the yellow reading changes nothing about that: " + k));
  assert.ok(!("Tuff" in touchedFields), "and Tuff is still not a seeded column at all");
  pass("the yellow reading is a new way of DERIVING the three counters, not a wider set of them");

  /* ================= 6. last writer wins ================= */
  /* CHANGED at step 3. The office's side of the contest used to be a
     `Dashboard Progress` row read out of a workbook 36 s old, plus the Log.
     It is one row of the record now, written at the click, and `Source` says
     who wrote it. */
  const officeOnly = (when, source, status) => {
    CPSRV = [];
    clearFail();
    /* a fresh dashboard for each case: no record, and no memory of having
       painted anything */
    A("PAINTED = {}; savePainted();");
    A("CP_ITEMS = []; cpRowsSet({}); CP_LIST_OK = true; " +
      "CP_IMPORTED = '2026-09-11T09:00:00Z'; cpImportCheck();");
    if (!when) return;
    /* the server's copy as well as the page's, so a write against it is a
       PATCH of a row that is really there */
    const st = status === undefined ? "done" : status;
    const f = { Title: "R7001|glass:dg", Job: "R7001", Item: "glass:dg",
                Done: st === "done" ? 4 : 0, Total: 4,
                Status: st, Who: "the colleague", When: when, Source: source || "office" };
    CPSRV.push({ id: "rec-dg", fields: f });
    global.__one = [{ id: "rec-dg", fields: f }];
    A("CP_ITEMS = __one; cpListRebuild();");
  };
  /* the same fresh record, but without forgetting the backoff - the give-up
     test has to accumulate failures across passes */
  const officeOnlyKeepFail = () => {
    global.__keep = A("JSON.stringify(GLASSC_FAIL)");
    officeOnly("");
    A("Object.assign(GLASSC_FAIL, JSON.parse(__keep)); setGlassFoot();");
  };
  const contest = async (officeWhen, floorWhen, opts, source) => {
    /* one column only, so the answer is about the contest and not about the
       three other cells the same job would also be arguing over */
    const cp = { win: "", drs: "", glass: { dg: "done" }, prod: {} };
    global.__jc = mkJob({ glass: { dg: 4 }, cp: cp });
    global.__items = [row(Object.assign({ Cut: 8, Hotmelt: 8, Glazed: 0, DoneAt: floorWhen }, opts || {}))];
    A("PENDING = {}; savePending(); ALL = [__jc]; STATION_ITEMS = __items;" +
      "CHANGES = []; state.sel = null;");
    officeOnly(officeWhen, source);
    reset();
    return await glassColourRun();
  };

  assert.strictEqual(await contest(officeAt(14, 0), isoAt(15, 0)), 1,
    "the floor moved the job after the office ticked it: the floor's colour goes in");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW, "gold walked back to yellow");
  pass("last writer wins: a floor tap after an office tick reaches the sheet");

  assert.strictEqual(await contest(officeAt(16, 0), isoAt(15, 0)), 0,
    "the office ticked it after the floor's last tap, so the office's gold stands");
  await settle();
  assert.strictEqual(CALLS.length, 0, "and not one request went out to argue about it");
  pass("last writer wins the other way: an office tick after a floor tap is not overwritten");

  /* the same instant, which is what a tie is now. Both sides write a full
     ISO stamp at the moment they act, so a tie is a real tie rather than an
     artefact of one of them being written to the minute - and it goes to the
     office, for the owner's own reason: "any change from the dashboard is
     absolute". A tie resolving to no write is also the quiet answer. */
  assert.strictEqual(await contest(isoAt(15, 0, 20), isoAt(15, 0, 20)), 0,
    "the very same second: the tie goes to the office");
  await settle();
  assert.strictEqual(CALLS.length, 0, "not one request went out to argue about it");
  assert.strictEqual(await contest(isoAt(15, 0, 20), isoAt(15, 0, 21)), 1,
    "and one second later the floor has it");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  pass("a tie goes to the office, and a second either way decides it");

  /* a `When` with no seconds in it can still turn up on a row - somebody
     editing the list by hand in SharePoint - and stampMs still reads it as the
     latest instant it can mean, so it covers its whole minute. */
  assert.strictEqual(await contest(officeAt(15, 0), isoAt(15, 0, 59)), 0,
    "a minute-precision office row covers the last second of its own minute");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  assert.strictEqual(await contest(officeAt(15, 0), isoAt(15, 1, 0)), 1,
    "and the next minute beats it");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  pass("a stamp with no seconds still covers its own minute, and no more than it");

  /* one side with nothing at all */
  assert.strictEqual(await contest("", isoAt(15, 0)), 1,
    "no office row: there is nothing to argue with, so the floor's colour goes in");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  assert.strictEqual(await contest(isoAt(16, 0), ""), 0,
    "and a floor row with no stamp is never painted whatever the office has done");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  pass("an office row that does not exist loses; an undated floor row never writes at all");

  /* WHICH SOURCES CAN BLOCK, which is the other half of the rule. `office` is
     the drawer; `excel` is a colour somebody painted in the file by hand,
     which the safeguard confirmed and adopted - both are a person deciding
     something. `import` is the switch-over reading today's colours once and
     `floor` is this feature's own earlier work: neither is somebody acting
     after the floor did, so neither blocks. */
  assert.strictEqual(await contest(isoAt(16, 0), isoAt(15, 0), null, "excel"), 0,
    "a colour adopted from Excel after the floor's tap blocks it, exactly as the office does");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  assert.strictEqual(await contest(isoAt(16, 0), isoAt(15, 0), null, "import"), 1,
    "an import row is the sheet's old colour, not a decision: it never blocks");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  assert.strictEqual(await contest(isoAt(14, 0), isoAt(15, 0), null, "floor"), 1,
    "and neither does this feature's OWN older row");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  pass("only a person's decision blocks the floor: office and excel do, import and floor do not");

  /* ... but this feature's own NEWER row does, which is the one case where a
     floor row blocks a floor write (review finding F4). STATION_ITEMS can
     regress - a full read of the list dispatched before a delta can land after
     it - and for a moment this dashboard holds an older copy of the floor's
     row than the one it has already recorded. Writing that back would undo the
     floor's own newer tap. */
  assert.strictEqual(await contest(isoAt(16, 0), isoAt(15, 0), null, "floor"), 0,
    "a record of the floor's work NEWER than the list copy in hand is never written over");
  await settle();
  assert.strictEqual(CALLS.length, 0, "and not one request goes out to do it");
  assert.strictEqual(await contest(isoAt(15, 0), isoAt(15, 0), null, "floor"), 1,
    "the same instant is not a regression, so it writes");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  pass("a list read that has gone backwards cannot undo the floor's own newer work");

  /* THE DASHBOARD LOG DECIDES NOTHING NOW. It used to be half the office's
     record - glassLogStamps read every "Glass ..." line as the office having
     spoken - because there was nowhere better to look. There is now. */
  assert.strictEqual(typeof glassLogStamps, "undefined", "glassLogStamps is gone");
  global.__jl = mkJob({ glass: { dg: 4 }, cp: { win: "", drs: "", glass: { dg: "done" }, prod: {} } });
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 0, DoneAt: isoAt(15, 0) })];
  officeOnly("");
  A("PENDING = {}; savePending(); ALL = [__jl]; STATION_ITEMS = __items;" +
    "CHANGES = [{ at: '10/09/2026 23:00', who: 'the colleague', job: 'R7001', what: 'Glass DG'," +
    " from: '', to: '4 of 4', shared: true }];");
  reset();
  assert.strictEqual(await glassColourRun(), 1,
    "a Log line hours after the floor's tap does not hold the floor's colour off");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW, "because the record is the only thing that decides");
  A("CHANGES = [];");
  pass("the Dashboard Log is history again, and decides nothing");

  /* 6b - "there is no window to lose any more" - is section 14, at the end of
     this file: it needs the confirmation stub and the floor-clear helpers that
     section 10b declares. */

  /* ================= 7. a cell this feature does not own ================= */
  global.__jx = mkJob({ cp: { win: "", drs: "", glass: { dg: "cut" }, prod: {} } });
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11, DoneAt: isoAt(15, 0) })];
  A("PENDING = {}; savePending(); ALL = [__jx]; STATION_ITEMS = __items; CHANGES = [];");
  CP.cpSetProgress({});
  forgetClears();
  reset();
  assert.strictEqual(await glassColourRun(), 1, "the other three columns are painted");
  await settle();
  assert.deepStrictEqual(fills().map(c => c.addr).sort(), ["AZ7", "BA7", "BB7"],
    "and DG is left exactly as it was: the sheet's own Cut green is not this feature's to paint over");
  pass("a glass cell carrying a colour this feature does not own is left alone, in either direction");

  /* a column that is not on this sheet, and a glass the job does not have */
  global.__mx = { qty: {}, glass: { dg: 51, tg: 52 }, prod: {}, prodOrder: [] };
  global.__jy = mkJob({ glass: { dg: 4, tuff: 11 } });
  A("PENDING = {}; savePending(); ALL = [__jy]; PRODMAP = __mx; STATION_ITEMS = __items;");
  reset();
  assert.strictEqual(await glassColourRun(), 1);
  await settle();
  assert.deepStrictEqual(fills().map(c => c.addr), ["AY7"],
    "TG is a column but the job has none; TUFF the job has but the sheet does not");
  A("PRODMAP = __m");
  pass("a column missing from the sheet, and a glass the job has none of, are both skipped");

  /* ================= 8. a refused write, and the brake on it ================= */
  global.__jf = mkJob();
  global.__items = [row({ Cut: 8, Hotmelt: 8, DoneAt: isoAt(15, 0) })];
  const failScene = () => {
    A("PENDING = {}; savePending(); ALL = [__jf]; STATION_ITEMS = __items; CHANGES = [];");
  };
  failScene(); clearFail();
  officeOnly("");                          // an empty record: nothing to argue with
  reset();
  FAIL_FILL = 20;
  assert.strictEqual(await glassColourRun(), 1);
  FAIL_FILL = 0;
  /* CHANGED at step 3. What stood here proved the `gc` hold - held the moment
     the write was sent, let go when it failed. There is no hold. What happens
     to a refused FILL now is what the record was for: the row landed, so the
     office's screen already shows the floor's colour and only Excel is behind,
     and cpRepaintRun paints it on the next load. */
  assert.strictEqual(A("Object.keys(PENDING).length"), 0, "nothing was held for it");
  assert.strictEqual(cpRow("R7001", "glass:dg").status, "process",
    "the record took the floor's colour: the row landed before the fill was even attempted");
  assert.strictEqual(paintedOf("R7001", "glass:dg"), undefined,
    "and the dashboard does not claim to have painted a cell the workbook refused");
  assert.strictEqual(fills().length, 0, "and no fill landed in the sheet at all");
  /* F5: and it still reaches Changes. The log line used to sit AFTER the fill,
     so a refused fill threw first and a floor colour that had been RECORDED
     left no trace anywhere - and cpRepaintRun, painting the cell on a later
     load, said nothing either. The record is the event; the colour in the cell
     is its copy. */
  assert.ok(A("CHANGES").some(c => c.what === "Floor glass colours" && c.job === "R7001"),
    "the floor's colour is in the office's history even though the fill was refused");
  const refusedLine = A("CHANGES").filter(c => c.what === "Floor glass colours")[0];
  assert.ok(/yellow/.test(refusedLine.to), "saying what was recorded: " + refusedLine.to);
  pass("a refused fill leaves the record standing and PAINTED untouched: Excel is behind, nothing is lost");

  /* THE BRAKE. Without it a failing write goes out again on the very next
     pass, and stationPoll answers "moved" every ten seconds for as long as the
     floor keeps tapping - a write storm, which is the worst thing this feature
     has available. The trigger is ordinary: somebody opens the file
     exclusively in desktop Excel mid-shift, or SharePoint refuses the list.

     CHANGED at step 3: the write that is refused here is the RECORD's, not the
     fill's. A refused fill leaves the record standing, so the writer has
     nothing left to plan for that cell and there is nothing to brake - it is
     cpRepaintRun that comes back for the colour, once per load. A refused
     record row is planned again on every pass until it lands, which is exactly
     what needs a brake. */
  failScene(); clearFail(); officeOnly("");
  reset();
  FAIL_CPLIST = 20;
  assert.strictEqual(await glassColourRun(), 1, "it tried");
  FAIL_CPLIST = 0;
  assert.strictEqual(A("GLASSC_FAIL['R7001'].n"), 1, "the failure is counted against that job");
  failScene();
  reset();
  assert.strictEqual(await glassColourRun(), 0, "so the next pass does not send it again");
  assert.strictEqual(CALLS.length, 0, "not one request");
  assert.ok(footWord().indexOf("glass colours not saved") >= 0,
    "and the footer says so, rather than only the console");
  assert.ok(footWhy().indexOf("trying again shortly") > 0,
    "with what happens next in the tooltip");
  pass("a refused write is not retried on the next poll: the job serves a backoff, and the footer shows it");

  /* it does come back, once the backoff has run */
  ageFail("R7001", 120000);
  failScene();
  reset();
  FAIL_CPLIST = 20;
  assert.strictEqual(await glassColourRun(), 1, "a minute later it tries again");
  FAIL_CPLIST = 0;
  assert.strictEqual(A("GLASSC_FAIL['R7001'].n"), 2, "and counts the second failure");
  /* ... and the wait gets longer each time, so a workbook locked for an hour
     is asked four or five times rather than three hundred */
  ageFail("R7001", 120000);
  failScene();
  reset();
  assert.strictEqual(await glassColourRun(), 0, "two failures in, a minute is no longer long enough");
  assert.strictEqual(CALLS.length, 0);
  ageFail("R7001", 600000);
  failScene();
  reset();
  FAIL_CPLIST = 20;
  assert.strictEqual(await glassColourRun(), 1, "five minutes is");
  FAIL_CPLIST = 0;
  pass("the backoff grows: one minute, then five, then fifteen");

  /* and it stops. A workbook that has refused this five times is not going to
     take it on the sixth, and the person looking at the screen is told rather
     than left with a dashboard quietly hammering the file. */
  FAIL_CPLIST = 40;
  for (let i = 0; i < 6; i++) {
    A("if (GLASSC_FAIL['R7001']) GLASSC_FAIL['R7001'].at -= 3600000");
    failScene();
    officeOnlyKeepFail();               // the record back to empty, the backoff kept
    await glassColourRun();
  }
  FAIL_CPLIST = 0;
  assert.ok(A("GLASSC_FAIL['R7001'].n") >= 5, "five failures and more");
  A("GLASSC_FAIL['R7001'].at -= 86400000");        // a whole day: no backoff could still be running
  failScene();
  reset();
  assert.strictEqual(await glassColourRun(), 0, "it has stopped trying");
  assert.strictEqual(CALLS.length, 0, "however long is left to wait");
  assert.ok(footWhy().indexOf("given up on until this page is reloaded") > 0,
    "and the footer says exactly that, so a person can do something about it");
  pass("a write the workbook keeps refusing is given up on, and the footer says so");

  /* a success clears the record, so a passing squall costs nothing */
  clearFail();
  failScene();
  reset();
  assert.strictEqual(await glassColourRun(), 1);
  assert.strictEqual(fillOn("AY7"), YELLOW);
  assert.strictEqual(A("GLASSC_FAIL['R7001']"), undefined, "the record is gone");
  assert.strictEqual(footWord().indexOf("glass colours not saved"), -1, "and so is the footer's word");
  pass("one successful write forgets the whole backoff for that job");

  /* ================= 8b. the cap on one pass =================
     Rehearsed against the owner's real file: 141 fed rows, and switching this
     feature on after the floor has worked a week with no office dashboard open
     plans up to 362 cell fills in one burst. The cap is the feeder's own
     answer to the same problem, with the feeder's own numbers. */
  clearFail();
  const many = [], manyItems = [];
  for (let i = 1; i <= 40; i++) {                  // 40 jobs x 4 cells = 160 > GLASS_MAX
    const id = "R8" + (100 + i);
    many.push(mkJob({ id: id }));
    manyItems.push({ id: String(800 + i), fields: Object.assign({}, row().fields,
      { Title: id, Job: id, Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11, DoneAt: isoAt(15, 0) }) });
  }
  /* the fake sheet carries one job number, and rowForJob answers for it - which
     is all this needs: the question is how many fills go out, not where */
  BOOK["Production"] = { v: {}, fill: {} };
  many.forEach((j, i) => { BOOK["Production"].v[kk(7 + i, 3)] = j.id; });
  global.__many = many; global.__manyItems = manyItems;
  const manyScene = () => A("PENDING = {}; savePending(); ALL = __many; STATION_ITEMS = __manyItems;" +
    "CHANGES = []; PRODMAP = __m; if (glassAgainT) { clearTimeout(glassAgainT); glassAgainT = null; }");
  manyScene();
  reset();
  const firstPass = await glassColourRun();
  assert.strictEqual(firstPass, 15, "fifteen jobs of four cells is sixty, and sixty is the cap");
  assert.strictEqual(fills().length, 60, "so exactly sixty cell fills went out, not a hundred and sixty");
  assert.ok(A("!!glassAgainT"), "and one follow-up was armed for the rest");
  assert.strictEqual(A("GLASS_MAX"), 60, "the same number the feeder uses");
  pass("a pass larger than the cap sends the cap's worth and no more");

  /* exactly one timer, not one per job - the bug the feeder already had */
  const timerAfterOne = A("glassAgainT");
  A("glassColourAgain(); glassColourAgain();");
  assert.strictEqual(A("glassAgainT"), timerAfterOne,
    "asking again while one is waiting does not arm a second");
  pass("exactly one follow-up is ever armed, however many jobs were left behind");

  /* the rest goes out on the next passes. PENDING is NOT cleared between them:
     the fifteen already painted are held, so they are not re-planned, which is
     what makes the next pass take the next fifteen rather than the same ones. */
  A("if (glassAgainT) { clearTimeout(glassAgainT); glassAgainT = null; }");
  reset();
  const secondPass = await glassColourRun();
  assert.strictEqual(secondPass, 15, "the next pass takes the next sixty cells");
  assert.strictEqual(fills().length, 60);
  A("if (glassAgainT) { clearTimeout(glassAgainT); glassAgainT = null; }");
  reset();
  const thirdPass = await glassColourRun();
  assert.strictEqual(thirdPass, 10, "and the last ten jobs are inside the cap");
  assert.strictEqual(fills().length, 40);
  assert.strictEqual(A("glassAgainT"), null,
    "a pass that finished the work arms no follow-up at all");
  reset();
  assert.strictEqual(await glassColourRun(), 0, "and there is nothing left to do");
  assert.strictEqual(CALLS.length, 0);
  /* a job is never split across two passes: half a job's colours would be a
     lie on screen for thirty seconds and the other half would only be
     re-planned anyway */
  assert.strictEqual((firstPass + secondPass + thirdPass) * 4, 160,
    "forty whole jobs went out over three passes, and none of them in halves");
  pass("the remainder goes out on the next passes, no job is split, and the last pass schedules nothing");
  clearFail();
  A("PRODMAP = __m; if (glassAgainT) { clearTimeout(glassAgainT); glassAgainT = null; }");

  /* ================= 9. one pass at a time, one write per job ================= */
  BOOK["Production"] = { v: {}, fill: {} };                 // 8b filled it with forty other jobs
  BOOK["Production"].v[kk(7, 3)] = "R7001";
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11, DoneAt: isoAt(16, 0) })];
  A("PENDING = {}; savePending(); ALL = [__jf]; STATION_ITEMS = __items;" +
    "if (glassAgainT) { clearTimeout(glassAgainT); glassAgainT = null; }");
  reset();
  const pass1 = glassColourRun();                           // deliberately not awaited
  const pass2 = glassColourRun();
  assert.strictEqual(await pass2, 0,
    "a second pass starting while the first is still in the air sends nothing");
  assert.ok(A("!!glassAgainT"),
    "but it arms the follow-up rather than simply giving up, so a big catch-up still finishes");
  assert.strictEqual(await pass1, 1, "and the first pass did its work");
  assert.strictEqual(fills().filter(c => c.addr === "AY7").length, 1, "one fill per cell, once");
  A("if (glassAgainT) { clearTimeout(glassAgainT); glassAgainT = null; }");
  reset();
  assert.strictEqual(await glassColourRun(), 0, "and afterwards there is nothing left to write");
  assert.strictEqual(CALLS.length, 0);
  pass("two passes cannot overlap: the second sends nothing, arms the follow-up, and no cell is written twice");

  /* ================= 10. the office's own controls are untouched ================= */
  j = scene(mkJob({ glass: { dg: 4, tg: 4, tuff: 11, "not tuff": 8 } }),
            row({ Cut: 8, Hotmelt: 8, DoneAt: isoAt(14, 0) }));
  assert.deepStrictEqual(cpItems(j).map(x => x.key),
    ["glass:dg", "glass:tg", "glass:tuff", "glass:not tuff"],
    "every glass column is still a checkpoint the office can tick, one by one");
  reset();
  setItemProgress(j, "glass:dg", 2);
  await settle(1200);
  assert.strictEqual(fillOn("AY7"), YELLOW, "the office's own tick still writes the colour");
  /* CHANGED 2026-09-11: the exact count goes to the `Dashboard progress` LIST
     now, not to the sheet of the same name, which nothing writes any more. The
     property being proved is the one that mattered - a drawer tick still
     records the count somewhere the colour cannot, and the colour writer still
     records nothing at all. */
  assert.ok(!CALLS.some(c => c.kind === "values" && c.sheet === "Dashboard Progress"),
    "the Dashboard Progress SHEET is not written by anything any more");
  const dgRow = CPSRV.filter(x => x.fields.Title === "R7001|glass:dg")[0];
  assert.ok(dgRow, "the record carries a row for the item that was ticked");
  assert.deepStrictEqual([dgRow.fields.Done, dgRow.fields.Total, dgRow.fields.Status, dgRow.fields.Source],
    [2, 4, "process", "office"], "with the exact count, which the colour writer never writes");
  assert.ok(CALLS.some(c => c.kind === "values" && c.sheet === "Dashboard Log"),
    "and still logs it");
  pass("the drawer's per-type ticks are exactly what they were: colour, count on the record, log line");

  /* REWRITTEN TWICE, and this is what is left of it.

     It began as "when the office's `cp` hold and the writer's `gc` hold both
     cover one glass cell, the newer of the two is drawn". Step 2 took the `cp`
     hold away and it became "the writer's paint is never status, and an office
     click voids it". Step 3 has taken the `gc` hold away too: the writer has
     no un-landed paint to void, because it writes the record and the record is
     what everything reads.

     So what is left to prove is the rule that replaced all of it: an office
     click writes a row stamped NOW, and a floor colour whose stamp is older
     than that row never goes near the sheet. No hold, no mask, no void. */
  global.__jb = mkJob({ glass: { dg: 4 }, cp: { win: "", drs: "", glass: {}, prod: {} } });
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11,
                          DoneAt: new Date(Date.now() - 60000).toISOString() })];
  A("PENDING = {}; savePending(); ALL = [__jb]; STATION_ITEMS = __items; CHANGES = [];");
  CP.cpSetProgress({});
  /* the office clicks and its row lands: cleared, stamped now, Source office */
  officeOnly(new Date().toISOString(), "office", "");
  assert.strictEqual(cpRow("R7001", "glass:dg").source, "office");
  assert.strictEqual(itemState(byId("R7001"), "glass:dg").status, "",
    "the office's un-tick is on the screen at once");
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "and the floor's gold, tapped a minute earlier, does not go near it");
  assert.strictEqual(CALLS.length, 0, "not one request");
  assert.strictEqual(A("Object.keys(PENDING).length"), 0, "and nothing was held to achieve that");
  /* and the mirror: a floor tap AFTER the office's click still paints */
  global.__items2 = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11,
                           DoneAt: new Date(Date.now() + 2000).toISOString() })];
  A("STATION_ITEMS = __items2");
  reset();
  assert.strictEqual(await glassColourRun(), 1, "a later floor tap is the later word, and it paints");
  await settle();
  assert.strictEqual(fillOn("AY7"), GOLD);
  assert.strictEqual(cpRow("R7001", "glass:dg").source, "floor", "and the record says whose work it is");
  pass("an office click beats an older floor tap and is beaten by a later one - two fields of one row");

  /* ================= 10b. an office clear reaches the floor =================
     docs/specs/2026-09-10-office-clears-the-floor.md. The owner spent an
     afternoon believing the un-tick was broken. It was not: the workbook was
     cleared correctly every time. What was wrong is that the FLOOR's counters
     were never cleared, so the tablet stayed gold for ever and the only way
     back was somebody tapping minus forty-nine times.

     So the assertions here are about all three saying the same thing
     afterwards - the office, the workbook and the floor - and about the write
     that makes the third one true being exactly six fields and reachable from
     nothing but a clear. */
  const ASKED = [];
  let ANSWER = true;
  global.confirm = m => { ASKED.push(String(m)); return ANSWER; };
  const clearWrites = () => LISTWRITES.slice();
  /* FedBy is never on this path, so the name is the one the feeder would use -
     a display name, never an address (feedWho) */
  Object.defineProperty(CW, "account", { configurable: true, get: () => ({ name: "the admin" }) });
  const goldCp = { win: "", drs: "", glass: { dg: "done", tg: "done", tuff: "done", "not tuff": "done" }, prod: {} };
  const tapped = o => row(Object.assign({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11,
    DoneAt: isoAt(14, 0), DoneBy: "Person A",
    CutBy: "Person A", CutAt: isoAt(13, 0), HotmeltBy: "Person A", HotmeltAt: isoAt(13, 30),
    GlazedBy: "Person B", GlazedAt: isoAt(14, 0) }, o || {}));

  j = scene(mkJob({ cp: goldCp }), tapped());
  LISTWRITES.length = 0; ASKED.length = 0; ANSWER = true;
  reset();
  await setGroupDone(j, "glass", false);
  await settle(120);
  assert.strictEqual(ASKED.length, 1, "the office is asked once, before anything is written");
  assert.strictEqual(ASKED[0],
    "The floor has recorded 8 cut, 8 hotmelted, 8 glazed, 11 tuff on this job. Clearing the " +
    "glass here will set all of those back to zero. Clear it anyway?",
    "and told exactly what will be destroyed");
  assert.strictEqual(clearWrites().length, 1, "one write to the floor's list, and one only");
  assert.strictEqual(clearWrites()[0].id, "700", "onto that job's own row");
  assert.deepStrictEqual(clearWrites()[0].fields,
    { Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0, DoneBy: "the admin",
      DoneAt: clearWrites()[0].fields.DoneAt, OfficeDone: "No" },
    "the four counters at nought, the last touch, and the unlock - asserted as an exact key set");
  assert.strictEqual(A("stationForJob('R7001').officeDone"), false,
    "and this dashboard's own copy of the row shows the job unlocked at once, without waiting for a feed");
  assert.ok(/^\d{4}-\d\d-\d\dT.*Z$/.test(clearWrites()[0].fields.DoneAt),
    "DoneAt is a real ISO stamp, so the floor's row does not read as old news");
  pass("an office clear on a job the floor has tapped writes exactly the four noughts and the last touch");

  ["CutBy", "CutAt", "HotmeltBy", "HotmeltAt", "GlazedBy", "GlazedAt", "TuffBy", "TuffAt"]
    .forEach(k => assert.ok(!(k in clearWrites()[0].fields),
      "no " + k + ": that says who did that stage's work, and nobody did"));
  ["Job", "Customer", "Total", "TuffTotal", "Seq", "Active", "FedAt", "FedBy"]
    .forEach(k => assert.ok(!(k in clearWrites()[0].fields), "and no job fact either: no " + k));
  /* OfficeDone is the one office column that does ride along - the unlock, so
     the tablet frees the card on its next poll rather than on the next feed */
  assert.strictEqual(clearWrites()[0].fields.OfficeDone, "No");
  assert.ok(!ALLREQ.some(r => /Station log|Station people|list-station/.test(r.path)),
    "and not one request of any kind reached the Station log or the people list");
  assert.strictEqual(A("Object.keys(FLOORCLEAR_OK).length"), 0, "the office's permission is spent once");
  assert.strictEqual(A("Object.keys(FLOORCLEAR_OWED).length"), 0, "and nothing is left owed");
  pass("no Station log line, no per-stage By or At, no job fact but the unlock: seven fields and stopped");

  /* THE POINT OF THE WHOLE FEATURE: all three now say the same thing */
  assert.strictEqual(A("stationForJob('R7001').cut"), 0, "the floor's row reads nought");
  assert.deepStrictEqual(ST.glassColours(A("stationForJob('R7001')")),
    { dg: "", tg: "", tuff: "", "not tuff": "" }, "so the colour rule wants blank everywhere");
  assert.strictEqual(A("stationForJob('R7001').finished"), false, "and the card is not gold any more");
  assert.deepStrictEqual(sheetNow(), { dg: WHITE, tg: WHITE, tuff: WHITE, "not tuff": WHITE,
    arch: undefined, astragal: undefined, fancy: undefined, extra: undefined },
    "the workbook's own four are white, and the hand-ticked four were never touched");
  reset();
  assert.strictEqual(await glassColourRun(), 0, "and the colour writer has nothing left to do");
  await settle();
  assert.strictEqual(CALLS.length, 0, "no second round of writes: the two agree the moment the clear lands");
  /* the clear IS logged, as an office action, in the office's own log (spec §3)
     - and it is read back as ONE: an office stamp on the JOB. It has to be.
     Its other half, the DoneAt it writes on the floor's row, is read by
     floorStamp as a floor action; counting that half and not this one is what
     let the office's own clear out-rank the office and paint out the columns
     the click never mentioned. It is a job stamp and not a column one, because
     a clear is about the whole job's glass and names no column at all. */
  assert.ok(A("CHANGES").some(c => c.what === "Floor glass counters" && c.job === "R7001"),
    "the clearing of the floor's counters is in Dashboard Log, like every other office action");
  /* CHANGED at step 3: the line is history and nothing reads it as a stamp any
     more (glassLogStamps is gone). What decides after a clear is the record -
     an office row per glass item, cleared and stamped at the click. */
  ST.COLOUR_TYPES.forEach(t => {
    const r = cpRow("R7001", "glass:" + t);
    if (!r) return;
    assert.strictEqual(r.status, "", t + " is recorded as cleared");
    assert.strictEqual(r.source, "office", "and recorded as the office's doing");
  });
  pass("after a clear the office, the workbook and the floor all say nothing is done - with no write to settle it");

  /* DECLINING WRITES NOTHING AT ALL - not even the workbook half */
  j = scene(mkJob({ cp: goldCp }), tapped());
  LISTWRITES.length = 0; ASKED.length = 0; ANSWER = false;
  reset();
  await setGroupDone(j, "glass", false);
  await settle(120);
  assert.strictEqual(ASKED.length, 1, "the office is still asked");
  assert.strictEqual(CALLS.length, 0, "and answering no writes nothing to the workbook at all");
  assert.strictEqual(clearWrites().length, 0, "nor to the floor's list");
  assert.strictEqual(A("Object.keys(PENDING).length"), 0, "nothing is even held on screen");
  assert.strictEqual(byId("R7001").cp.glass.dg, "done", "the job is exactly as it was");
  assert.strictEqual(A("Object.keys(FLOORCLEAR_OK).length"), 0,
    "and no permission was left lying about for the next write to pick up");
  pass("declining the question writes nothing anywhere: not the floor's counters, not the workbook, not a hold");

  /* NOTHING IS ASKED, AND NOTHING IS WRITTEN, ON A ROW THE FLOOR NEVER TAPPED.
     Its counters are the office's own seed echoed back and the feeder puts
     them right on its next run - and writing DoneAt here would mark the row
     touched for ever and switch its seeding off, which is a worse bug than the
     one this fixes. */
  j = scene(mkJob({ cp: goldCp }), row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 0, DoneAt: "", DoneBy: "" }));
  LISTWRITES.length = 0; ASKED.length = 0; ANSWER = true;
  reset();
  await setGroupDone(j, "glass", false);
  await settle(120);
  assert.strictEqual(ASKED.length, 0, "nothing is asked: there is nothing of the floor's to lose");
  assert.strictEqual(clearWrites().length, 0, "and the office writes not one field of that row");
  assert.strictEqual(A("stationForJob('R7001').doneAt"), "",
    "the row is still untouched, so the feeder can still seed it - which is what will clear it");
  assert.deepStrictEqual(ST.seedFields({ total: 8, seed: ST.officeSeed({ total: 8 },
    [{ type: "dg", total: 4, status: "", done: 0 }, { type: "tg", total: 4, status: "", done: 0 }]) }),
    { Cut: 0, Hotmelt: 0, Glazed: 0 }, "and the seed it will write is nought, nought, nought");
  assert.ok(fills().length > 0, "while the workbook half went out exactly as before");
  pass("a row the floor never tapped is left to the feeder: no question, no write, no stamp planted on it");

  /* THE PER-ITEM CLEAR, on a job whose only glass is DG */
  j = scene(mkJob({ glass: { dg: 8 }, cp: { win: "", drs: "", glass: { dg: "done" }, prod: {} } }),
            tapped({ Total: 8, TuffTotal: 0, Tuff: 0 }));
  LISTWRITES.length = 0; ASKED.length = 0; ANSWER = true;
  reset();
  setItemProgress(j, "glass:dg", 0);
  await settle(1400);
  assert.strictEqual(ASKED.length, 1, "the per-item Clear asks too");
  assert.ok(ASKED[0].indexOf("8 cut, 8 hotmelted, 8 glazed") > 0, "naming the same work");
  assert.strictEqual(clearWrites().length, 1);
  assert.deepStrictEqual(Object.keys(clearWrites()[0].fields).sort(),
    ST.OFFICE_CLEAR_FIELDS.slice().sort(), "with the same six fields and no others");
  assert.strictEqual(fillOn("AY7"), WHITE, "and the workbook cell went white, as it always did");
  pass("the drawer's per-item Clear clears the floor as well, by the same one write");

  /* ... but only when it leaves the job's glass reading NOTHING done. Clearing
     DG on a job whose TG is still gold is not a clear of the job's glass: the
     floor's row carries one combined DG + TG number, and telling them nought
     while the office still says TG is finished would be a wrong instruction on
     a workshop screen. */
  j = scene(mkJob({ glass: { dg: 4, tg: 4 },
                    cp: { win: "", drs: "", glass: { dg: "done", tg: "done" }, prod: {} } }),
            tapped({ TuffTotal: 0, Tuff: 0 }));
  LISTWRITES.length = 0; ASKED.length = 0;
  reset();
  setItemProgress(j, "glass:dg", 0);
  await settle(1400);
  assert.strictEqual(ASKED.length, 0, "no question, because nothing of the floor's is being destroyed");
  assert.strictEqual(clearWrites().length, 0, "and no counter is touched");
  assert.strictEqual(fillOn("AY7"), WHITE, "the office's own half is exactly what it was");
  assert.strictEqual(A("stationForJob('R7001').cut"), 8, "the floor keeps its numbers");
  pass("a per-item clear that leaves the job's other glass ticked is not a clear of the job's glass");

  /* MARKING DONE IS UNTOUCHED: this is the un-tick path and nothing else */
  j = scene(mkJob({ cp: { win: "", drs: "", glass: {}, prod: {} } }), tapped({ Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0 }));
  LISTWRITES.length = 0; ASKED.length = 0;
  reset();
  await setGroupDone(j, "glass", true);
  await settle(120);
  assert.strictEqual(ASKED.length, 0, "marking done asks nobody anything");
  assert.strictEqual(clearWrites().length, 0, "and writes not one field of the floor's list");
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AY7", GOLD], ["AZ7", GOLD], ["BA7", GOLD], ["BB7", GOLD]],
    "it is the four fills it always was");
  setItemProgress(j, "glass:dg", 2);
  await settle(1400);
  assert.strictEqual(ASKED.length, 0, "and a partial tick asks nothing either");
  assert.strictEqual(clearWrites().length, 0);
  pass("marking done, and ticking part of a job off, behave exactly as they did: this is the un-tick path only");

  /* A REFUSED LIST WRITE. The workbook half has landed, so the counters must
     not simply be forgotten - the tablet would stay gold, which is the bug. */
  j = scene(mkJob({ cp: goldCp }), tapped());
  LISTWRITES.length = 0; ASKED.length = 0; ANSWER = true;
  FAIL_LIST = 1;
  reset();
  await setGroupDone(j, "glass", false);
  await settle(120);
  assert.strictEqual(clearWrites().length, 0, "the write was refused");
  assert.strictEqual(A("FLOORCLEAR_OWED['R7001']"), 1, "so it is owed");
  assert.ok(A("!!floorClearT"), "and a follow-up is armed rather than the counters being forgotten");
  assert.ok(fills().length > 0, "while the workbook half stands: the sheet really was cleared");
  A("clearTimeout(floorClearT); floorClearT = null;");
  reset();
  assert.strictEqual(await clearFloorGlass("R7001"), true, "the follow-up gets it through");
  assert.strictEqual(clearWrites().length, 1);
  assert.strictEqual(A("Object.keys(FLOORCLEAR_OWED).length"), 0, "and nothing is owed afterwards");
  pass("a refused clear is owed and retried, not lost - the one failure that would leave the tablet gold");

  /* and it cannot be reached by asking for it: the reason is re-derived */
  j = scene(mkJob({ cp: goldCp }), tapped());
  LISTWRITES.length = 0;
  assert.strictEqual(await clearFloorGlass("R7001"), false,
    "a call with no clear behind it writes nothing: the office's record still says done");
  assert.strictEqual(clearWrites().length, 0);
  A("delete FLOORCLEAR_OWED['R7001'];");
  pass("the write re-derives its own reason, so it is not reachable by calling it");


  /* THE OFFICE'S ANSWER BELONGS TO ONE WRITE, NOT TO THE JOB. A job has
     several checkpoint bursts at once - the burst key is job|item - and any of
     them can land first. Keyed on the JOB, whichever landed first would spend
     the answer, so an unrelated tick could clear the floor's counters BEFORE
     the glass write had landed; and if the glass write then failed the floor
     would be at nought with the sheet still gold, which is the mirror image of
     the bug this feature exists to remove. Here the unrelated burst is
     deliberately started FIRST, so it is the one that lands first. */
  j = scene(mkJob({ glass: { dg: 8, tuff: 5 },
                    cp: { win: "", drs: "", glass: { dg: "done" }, prod: {} } }),
            tapped({ Total: 8, TuffTotal: 5, Tuff: 0 }));
  LISTWRITES.length = 0; ASKED.length = 0; ANSWER = true;
  reset();
  const req0 = ALLREQ.length;
  setItemProgress(j, "glass:tuff", 3);              // an unrelated burst, started first
  assert.strictEqual(ASKED.length, 0, "a TUFF tick is not a clear of the job's glass: nothing is asked");
  setItemProgress(byId("R7001"), "glass:dg", 0);    // and then the clear itself
  assert.strictEqual(ASKED.length, 1, "the clear asks, once");
  await settle(1600);
  assert.strictEqual(clearWrites().length, 1, "one clear write, not two and not none");
  assert.deepStrictEqual(Object.keys(clearWrites()[0].fields).sort(),
    ST.OFFICE_CLEAR_FIELDS.slice().sort(), "and it is the clear's own six fields");
  const seq = ALLREQ.slice(req0).map(r => r.path);
  const iTuff = seq.findIndex(p => p.indexOf("BA7") >= 0);
  const iGlass = seq.findIndex(p => p.indexOf("AY7") >= 0);
  /* the FLOOR's list, named explicitly: since 2026-09-11 the drawer also
     writes the record - a list request of its own, and an earlier one */
  const iList = seq.findIndex(p => p.indexOf(GLASS_LIST_ID) >= 0 && p.indexOf("/items/") >= 0);
  assert.ok(iTuff >= 0 && iGlass >= 0 && iList >= 0, "all three writes really went out");
  assert.ok(iTuff < iList, "the unrelated write landed first, as it was meant to");
  assert.ok(iGlass < iList,
    "and the floor's counters were not cleared until the GLASS write itself had landed");
  assert.strictEqual(A("Object.keys(FLOORCLEAR_OK).length"), 0, "with no answer left over");
  pass("the office's answer is spent by the write it was given for, never by another burst on the same job");

  /* and the other half of the same fix: an answer left behind by a burst that
     never fired is inert - only the very write it was given for can spend it */
  j = scene(mkJob({ glass: { dg: 8, tuff: 5 },
                    cp: { win: "", drs: "", glass: { dg: "done" }, prod: {} } }),
            tapped({ Total: 8, TuffTotal: 5, Tuff: 0 }));
  LISTWRITES.length = 0; ASKED.length = 0;
  A("FLOORCLEAR_OK['R7001|glass:dg'] = 1;");       // left behind by a cancelled burst
  reset();
  setItemProgress(j, "glass:tuff", 2);
  await settle(1400);
  assert.strictEqual(clearWrites().length, 0,
    "a TUFF write cannot spend an answer that was given about the job's DG");
  assert.strictEqual(A("FLOORCLEAR_OK['R7001|glass:dg']"), 1, "which is still sitting there, unspent");
  A("delete FLOORCLEAR_OK['R7001|glass:dg'];");
  pass("a leftover answer is inert until the very write it was given for lands");

  /* CLEAR IT, THEN THINK BETTER OF IT, INSIDE THE SAME MINUTE. Now an ordinary
     workflow rather than a race, because the clear is a deliberate click and
     re-ticking is the obvious next move. The office's stamps carry no seconds
     (nowStamp), so they are read as covering their whole minute; the clear's
     DoneAt is a full ISO stamp and is read exactly. A clear at 15:00:10 and a
     re-tick recorded as 15:00 must therefore resolve to the OFFICE - otherwise
     the colour writer paints the floor's blank straight back over the re-tick
     and the office watches its own work undone. */
  j = scene(mkJob({ glass: { dg: 4, tg: 4 },
                    cp: { win: "", drs: "", glass: { dg: "done", tg: "done" }, prod: {} } }),
            row({ Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0,
                  DoneAt: isoAt(15, 0, 10), DoneBy: "the admin" }));
  /* CHANGED at step 3: the office's re-tick is a row in the record, stamped at
     the click, not a Dashboard Progress row read out of a workbook. Written at
     minute precision here on purpose, because stampMs still reads a
     seconds-less stamp as covering its whole minute and that is what decides
     this case. */
  CPSRV = [];
  global.__retick = ["dg", "tg"].map((t, n) => ({ id: "re" + n, fields:
    { Title: "R7001|glass:" + t, Job: "R7001", Item: "glass:" + t, Done: 4, Total: 4,
      Status: "done", Who: "the admin", When: officeAt(15, 0), Source: "office" } }));
  __retick.forEach(x => CPSRV.push({ id: x.id, fields: x.fields }));
  A("CP_ITEMS = __retick; cpListRebuild();");
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "the office re-ticked inside the same minute, so the office is the later word");
  await settle();
  assert.strictEqual(CALLS.length, 0, "and not one cell is painted back out");
  /* and it is not simply "the office always wins": a floor tap in the NEXT
     minute takes the same cells back */
  global.__items = [row({ Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0,
                          DoneAt: isoAt(15, 1, 0), DoneBy: "Person A" })];
  A("STATION_ITEMS = __items;");
  reset();
  assert.strictEqual(await glassColourRun(), 1, "the next minute beats the office's whole-minute stamp");
  await settle();
  assert.deepStrictEqual(fills().map(c => c.addr).sort(), ["AY7", "AZ7"],
    "and the floor's blank reaches the two glass columns it is about");
  pass("clear it, re-tick it inside the same minute, and the office keeps it - the minute a stamp covers decides it");

  Object.defineProperty(CW, "account", { configurable: true, get: () => null });
  delete global.confirm;


  /* ================= 10c and 10d: DELETED AT STEP 3 ========================
     Two sections, about three hundred and thirty lines, both of them about
     machinery that no longer exists.

     **10c, "the office's stamp is PER JOB".** glassOfficeStamp was asked per
     COLUMN, found no hold, no Progress row and no Log line for a column the
     office had not named in that action, answered a literal 0, and lost to any
     floor stamp at all - so pressing Clear on TG alone painted out the TUFF
     and NOT TUFF the office had ticked by hand. The fix was to add a per-JOB
     stamp gathered from four scattered records. There is no gathering now: the
     office's action on a column IS that column's row, written at the click,
     and a column the office never named simply keeps the row it already had.
     The case that section existed for is covered above, in "an office click
     beats an older floor tap and is beaten by a later one".

     **10d, "the office is absolute: the writer stands down".** officeSettling
     froze the writer for three minutes after any office action on a job,
     because inside that window every record the contest was decided on was a
     copy that had not caught up. There is no window and no copy: the office's
     row is on the record before a single request leaves. What that section
     proved - that a poll landing mid-clear cannot paint the floor's gold over
     an un-tick - is section 14 below, and it now passes without anything
     standing down at all.

     Both are named here rather than quietly dropped, because the bugs they
     were written for were real and expensive, and the next person to touch
     this needs to know they were answered rather than forgotten. */
  assert.strictEqual(typeof glassOfficeStamp, "undefined", "glassOfficeStamp is gone");
  assert.strictEqual(typeof glassOfficeJobStamp, "undefined", "and so is the per-job one");
  assert.strictEqual(typeof officeSettling, "undefined", "and the quarantine with them");
  assert.strictEqual(typeof OFFICE_FLOOR_AT, "undefined",
    "and this dashboard's private memory of its own clear");
  pass("the four mechanisms that argued about who spoke last are gone, and named");


  /* ================= 12. the hold machinery, and what replaced it ============
     REWRITTEN 2026-09-11 (spec: status-list-is-truth, step 2).

     What stood here was thirteen assertions over five cases (a-e) about the
     `cp` PENDING hold: that it was not let go into a download that still
     disagreed with it; that an EXPIRED one was kept and a re-read demanded;
     that after about a dozen disagreeing parses it was given up on out loud;
     that a tick was protected exactly as an un-tick was; that an
     expired-but-kept hold kept the colour writer standing down; and that a
     page booting with one armed a re-read for itself.

     Every one of those existed to make the dashboard go on believing the
     office over a copy of the workbook ~36 s behind. There is no copy in the
     checkpoint path any more: status is the `Dashboard progress` record, the
     record is written at the click, and the download is never asked. So the
     hold is gone, and with it the reconcile-for-holds, the expiry, the
     give-up and its red toast - for checkpoints. What is proved instead is
     the property all of that was trying to buy, directly:

       - an un-tick is white at once, stays white through a reload, and stays
         white when the stale gold download finally arrives;
       - a tick behaves the same way, which is the asymmetry the owner saw;
       - none of it involves PENDING at all.

     The `gc` holds - the colour writer's own un-landed paints - still ride the
     download and still have all of that machinery; (e) below proves a page
     that boots holding one still asks to be re-read. Step 3 retires them. */
  const holdOff = () => A("if (reconcileT) { clearTimeout(reconcileT); reconcileT = null; }" +
                          "holdForceAt = 0; busy = false;");
  const armed = () => A("!!reconcileT");
  /* the job as the stale download still has it: TG gold */
  const staleGold = () => mkJob({ glass: { tg: 49 },
    cp: { win: "", drs: "", glass: { tg: "done" }, prod: {} } });
  /* and as the file reads once it has caught up with the un-tick */
  const caughtUp = () => mkJob({ glass: { tg: 49 },
    cp: { win: "", drs: "", glass: {}, prod: {} } });

  /* (a) the un-tick, and then the same stale gold download over and over.
     R1-R8 of the scratchpad harness, in one assertion each. */
  j = scene(staleGold(), null);
  holdOff();
  A("PENDING = {}; savePending();");
  officeRow("R7001", "glass:tg", 0, 49, new Date().toISOString());
  assert.strictEqual(itemState(byId("R7001"), "glass:tg").status, "",
    "the un-tick is white at once");
  global.__stale = staleGold();
  for (let i = 0; i < 8; i++) {
    A("ALL = applyPending([__stale], true);");        // the 36 s-behind file, arriving again
    assert.strictEqual(itemState(byId("R7001"), "glass:tg").status, "",
      "sample " + (i + 1) + ": still white - nothing in the checkpoint path reads the download");
  }
  assert.strictEqual(A("PENDING['R7001']"), undefined,
    "and nothing is held: there is no copy to hold against");
  assert.ok(!armed(), "so no re-read is demanded for a checkpoint either");
  pass("an un-tick stays white through eight stale downloads, holding nothing and demanding nothing");

  /* (b) the reload. The record is re-read from SharePoint, not from
     localStorage, so a fresh page is in exactly the same place. */
  A("CP_ITEMS = []; cpRowsSet({});");                 // a brand-new page, knowing nothing
  global.__boot = [{ id: "900", fields: cpRowFields("R7001", "glass:tg", 0, 49, "",
                                                    "the admin", new Date().toISOString(), "office") }];
  A("CP_ITEMS = __boot; cpListRebuild();");
  A("ALL = applyPending([__stale], true);");
  assert.strictEqual(itemState(byId("R7001"), "glass:tg").status, "",
    "after a reload the record still says white, because the record is where it lives");
  pass("a refresh changes nothing: the record is read back from SharePoint, not from this browser");

  /* (c) TICKS ARE PROTECTED THE SAME WAY. */
  global.__white = caughtUp();
  officeRow("R7001", "glass:tg", 49, 49, new Date().toISOString());
  A("ALL = applyPending([__white], true);");
  assert.strictEqual(itemState(byId("R7001"), "glass:tg").status, "done",
    "the office's gold is on the screen at once, over a file that still says white");
  A("ALL = applyPending([__white], true);");
  assert.strictEqual(itemState(byId("R7001"), "glass:tg").status, "done",
    "and it does not flicker back on the next stale parse either");
  pass("a tick is protected exactly as an un-tick is, and by the same nothing-at-all");

  /* (d) THERE IS NO QUARANTINE, AND NONE IS NEEDED. Step 2 turned
     officeSettling's three-minute window onto the record's own stamp; step 3
     took the window away altogether. The office does not need a period of
     protection, because its row is on the record at the click and the contest
     is decided on that row - so it is protected for as long as it is the later
     word, and no longer, which is what the rule always meant. */
  holdOff();
  assert.strictEqual(typeof officeSettling, "undefined", "the quarantine is gone");
  A("cpRowsSet({}); CP_ITEMS = [];");
  pass("no quarantine: the office is protected by being the later word, not by a timer");

  /* (e) THE WRITER HAS NO UN-LANDED PAINT TO ORPHAN. Its colour is on the
     record before the fill goes out, and the record is not in this browser.
     bootReconcile still exists and still matters, but for the holds that DO
     ride the downloaded file: a section move, mark-ready, a product status. */
  holdOff();
  A("PENDING = {}; savePending(); pend('R7001', { blk: 2 });");
  A("if (reconcileT) { clearTimeout(reconcileT); reconcileT = null; }");
  A("bootReconcile();");
  assert.ok(armed(), "a page that starts up holding a section move asks to be re-read");
  holdOff();
  A("PENDING = {}; savePending(); bootReconcile();");
  assert.ok(!armed(), "a page holding nothing asks for nothing");
  /* and a glass colour can no longer be held at all */
  A("PENDING = {}; savePending(); pend('R7001', { gc: { tg: 'gold' } });");
  assert.strictEqual(A("!!(PENDING['R7001'] && PENDING['R7001'].gc)"), false,
    "pend() has no `gc` branch: a colour cannot be held, even by asking for it");
  assert.strictEqual(A("(localStorage.getItem('cw_pending') || '').indexOf('gc')"), -1,
    "and nothing about a colour reaches cw_pending");
  holdOff();
  A("PENDING = {}; savePending(); CHANGES = [];");
  pass("nothing can hold a colour any more, and the boot re-read serves the holds that remain");


  /* ================= 13. THE OWNER'S WHOLE COMPLAINT, AS ONE SEQUENCE =======
     This replaces "the office's clear voids the writer's paint" (about a
     hundred and seventy lines), which proved that a `gc` hold the office had
     out-ranked was DISCARDED rather than masked. There is no `gc` hold, so
     there is nothing to discard - and the thing that section was defending
     against is simply not reachable any more. What is worth proving instead is
     the whole run of events the owner actually reported, end to end, in order,
     with a stale download and a reload in the middle of it.

       the floor finishes a job   -> the record says done, Source floor
                                  -> Excel goes gold
                                  -> the tablet locks (OfficeDone)
       the office presses Clear   -> the record says nothing, Source office,
                                     stamped later
                                  -> Excel goes white
                                  -> the floor's counters go to nought
                                  -> the tablet unlocks
       a stale download arrives   -> nothing changes
       a stale list read arrives  -> nothing changes
       the floor taps again later -> the record says done again, Source floor
                                  -> Excel goes gold again

     Every one of those steps was, at some point in the two days before this,
     the one that went wrong. */
  const seqJob = () => mkJob({ glass: { dg: 4, tg: 4 },
                               cp: { win: "", drs: "", glass: {}, prod: {} } });
  const seqRow = (o) => row(Object.assign({ Total: 8, TuffTotal: 0, DoneBy: "Person B" }, o));
  const recWord = t => { const r = cpRow("R7001", "glass:" + t); return r ? r.status : null; };
  const recSrc = t => { const r = cpRow("R7001", "glass:" + t); return r ? r.source : null; };
  const glassOf = j => ST.COLOUR_TYPES.filter(t => (j.glass || {})[t] > 0);

  /* --- 1. THE FLOOR FINISHES THE JOB --- */
  const t0 = Date.now() - 120000;
  j = scene(seqJob(), seqRow({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 0,
                               DoneAt: new Date(t0).toISOString() }));
  officeOnly("");                                   // nothing on the record yet
  A("STATION_ITEMS = __items;");
  reset();
  assert.strictEqual(await glassColourRun(), 1, "the floor's finished job is planned");
  await settle();
  assert.deepStrictEqual(["dg", "tg"].map(recWord), ["done", "done"],
    "1. the record says the glass is done");
  assert.deepStrictEqual(["dg", "tg"].map(recSrc), ["floor", "floor"], "and that it is the floor's work");
  assert.strictEqual(cpRow("R7001", "glass:dg").who, "Person B", "with the floor person's name on it");
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AY7", GOLD], ["AZ7", GOLD]], "and Excel goes gold");
  assert.strictEqual(paintedOf("R7001", "glass:dg"), "done", "and the dashboard knows it painted it");
  /* the lock: OfficeDone is derived from the record, so the floor finishing a
     job is what locks it on the tablet */
  assert.strictEqual(ST.officeComplete(glassCounts(byId("R7001"))), true,
    "1. the job's glass reads finished, so the feeder will lock it on the tablet");
  pass("1. the floor finishes a job: the record, then Excel, then the lock");

  /* --- 2. THE OFFICE PRESSES CLEAR --- */
  LISTWRITES.length = 0; ASKED.length = 0; ANSWER = true;
  reset();
  await setGroupDone(byId("R7001"), "glass", false);
  await settle(80);
  assert.deepStrictEqual(["dg", "tg"].map(recWord), ["", ""], "2. the record says nothing is done");
  assert.deepStrictEqual(["dg", "tg"].map(recSrc), ["office", "office"], "and that the office said so");
  assert.ok(stampMs(cpRow("R7001", "glass:dg").when) > stampMs(new Date(t0).toISOString()),
    "stamped later than the floor's own tap, which is the whole of the contest");
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AY7", WHITE], ["AZ7", WHITE]], "and Excel goes white");
  assert.strictEqual(ST.officeComplete(glassCounts(byId("R7001"))), false,
    "2. the job no longer reads finished, so the tablet unlocks");
  const cleared = clearWrites();
  assert.strictEqual(cleared.length, 1, "and the floor's counters are cleared, in one write");
  assert.deepStrictEqual([cleared[0].fields.Cut, cleared[0].fields.Hotmelt,
                          cleared[0].fields.Glazed, cleared[0].fields.Tuff], [0, 0, 0, 0]);
  assert.strictEqual(cleared[0].fields.OfficeDone, "No", "with the unlock riding along");
  pass("2. the office clears it: the record, Excel, the floor's counters and the lock");

  /* --- 3. A STALE DOWNLOAD, AND A STALE LIST READ, CHANGE NOTHING --- */
  global.__stale = mkJob({ glass: { dg: 4, tg: 4 },
                           cp: { win: "", drs: "", glass: { dg: "done", tg: "done" }, prod: {} } });
  reset();
  for (let i = 0; i < 5; i++) {
    A("ALL = applyPending([__stale], true);");
    assert.deepStrictEqual(["dg", "tg"].map(t => itemState(byId("R7001"), "glass:" + t).status),
      ["", ""], "3. download " + (i + 1) + ": still white");
  }
  /* and the floor's list read back as it was BEFORE the clear - the tablet's
     own poll answering with what it had a moment ago */
  A("STATION_ITEMS = __items;");
  assert.strictEqual(await glassColourRun(), 0,
    "3. a stale list read does not put the gold back: the office spoke later");
  await settle();
  assert.strictEqual(CALLS.length, 0, "not one request went out");
  assert.deepStrictEqual(["dg", "tg"].map(recWord), ["", ""], "and the record still says nothing");
  pass("3. a stale download and a stale list read change nothing at all");

  /* --- 4. AND A GENUINE FLOOR TAP LATER PUTS IT BACK --- */
  global.__later = [seqRow({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 0,
                             DoneAt: new Date(Date.now() + 5000).toISOString() })];
  A("STATION_ITEMS = __later;");
  reset();
  assert.strictEqual(await glassColourRun(), 1, "4. a later tap is the later word");
  await settle();
  assert.deepStrictEqual(["dg", "tg"].map(recWord), ["done", "done"], "the record says done again");
  assert.deepStrictEqual(["dg", "tg"].map(recSrc), ["floor", "floor"], "and the floor's, again");
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AY7", GOLD], ["AZ7", GOLD]], "and Excel goes gold again");
  assert.strictEqual(ST.officeComplete(glassCounts(byId("R7001"))), true, "and the tablet locks again");
  assert.strictEqual(A("Object.keys(PENDING).length"), 0,
    "and not one of those four steps held anything, anywhere");
  pass("4. a genuine floor tap after the clear puts the gold back - and the whole sequence held nothing");

  /* ================= 14. there is no window to lose any more =================
     THE OWNER'S BUG, 2026-09-10: "when i am marking all done it working fine
     ... but when i am trying to remove the all done ... it refreshing back to
     all golden".

     Every record the contest used to be decided on was written by the WRITE,
     not by the click - the Dashboard Progress row at the start of it and the
     Dashboard Log line at the end. Between the two, the office HAD acted and
     nothing the planner could see said so, so the floor won by default, its
     gold landed after the office's white, and the cell was then the colour the
     floor wanted, so it was never written again: the un-tick gone, silently
     and for good. A whole quarantine (officeSettling) was built to survive
     that window.

     There is no window. cpRowNow puts the office's row in at the click, before
     a single request leaves, and that row is what the contest reads. */
  const untick = () => {
    j = scene(mkJob({ glass: { tg: 49, tuff: 20, "not tuff": 29 },
      cp: { win: "", drs: "", glass: { tg: "done", tuff: "done", "not tuff": "done" }, prod: {} } }),
      row({ Total: 49, TuffTotal: 20, Cut: 49, Hotmelt: 49, Glazed: 49, Tuff: 20,
            DoneBy: "Person A", DoneAt: new Date(Date.now() - 1000).toISOString() }));
    LISTWRITES.length = 0; ASKED.length = 0; ANSWER = true;
    reset();
    return j;
  };
  untick();
  assert.strictEqual(await glassColourRun(), 0,
    "before the un-tick there is nothing to do: the record and the floor agree");
  /* the office presses Clear, and the ten-second station poll lands while that
     write is still in the air */
  const clearing = setGroupDone(byId("R7001"), "glass", false);
  const raced = await glassColourRun();
  await clearing;
  await settle(300);
  assert.strictEqual(raced, 0,
    "a poll landing while the office's un-tick is in the air cannot plan the floor's gold over it");
  assert.ok(!fills().some(c => c.color === GOLD),
    "and no gold was written back: " + JSON.stringify(fills().map(c => c.addr + "=" + c.color)));
  ST.COLOUR_TYPES.forEach(t => {
    const r = cpRow("R7001", "glass:" + t);
    if (!r) return;
    assert.strictEqual(r.status, "", t + " is recorded as cleared");
    assert.strictEqual(r.source, "office", "by the office");
  });
  pass("an un-tick still in the air is already on the record: there is no window to lose");

  /* and the same one item at a time, where the window used to be longest of
     all: the stepper debounces for 800 ms before it writes anything */
  untick();
  setItemProgress(byId("R7001"), "glass:tg", 0);
  const racedItem = await glassColourRun();                     // inside the debounce
  await settle(1400);
  assert.strictEqual(racedItem, 0,
    "the same during the stepper's 800 ms debounce, when nothing has been written at all yet");
  assert.strictEqual(cpRow("R7001", "glass:tg").status, "", "TG is cleared on the record");
  assert.notStrictEqual(fillOn("AZ7"), GOLD, "and TG was not repainted gold under the office's hand");
  pass("a per-item un-tick is safe through its debounce as well, and for the same reason");
  A("PENDING = {}; savePending(); CHANGES = [];");


  /* ================= 11. the whole run, end to end ================= */
  const prodWrites = ALLREQ.filter(r => r.method !== "GET" && /worksheets\('Production'\)/.test(r.path));
  assert.ok(prodWrites.length > 0, "the Production sheet really was written during this run");
  prodWrites.forEach(r => assert.ok(/\/format\/fill$/.test(r.path),
    "every write to Production is a fill and nothing else: " + r.path));
  prodWrites.forEach(r => assert.deepStrictEqual(Object.keys(r.body || {}), ["color"],
    "and its body is a colour and nothing else: " + JSON.stringify(r.body)));
  assert.strictEqual(ALLREQ.filter(r => r.method === "DELETE").length, 0,
    "nothing was deleted, from anywhere");
  assert.strictEqual(ALLREQ.filter(r => /\/insert|\/delete|\/rows|worksheets\/add/.test(r.path))
                           .filter(r => /Production'/.test(r.path)).length, 0,
    "and no row was inserted or removed");
  pass("over the whole run, every write to the Production sheet was a single cell fill");

  const glassAddrs = prodWrites.map(r => (/address='([^']+)'/.exec(r.path) || [])[1]);
  /* one cell each time - never a range, never a whole row or column - and its
     column is one of the four. DG is AY, TG is AZ, TUFF is BA, NOT TUFF is BB;
     the four the office ticks by hand are BC to BF and are never addressed. */
  const ALLOWED = ["AY", "AZ", "BA", "BB"];
  glassAddrs.forEach(a => {
    assert.ok(/^[A-Z]{1,2}\d+$/.test(a), a + " is not a single cell");
    const col = (/^[A-Z]+/.exec(a) || [])[0];
    assert.ok(ALLOWED.indexOf(col) >= 0,
      a + " is not in one of the four columns this feature may write");
  });
  ["BC", "BD", "BE", "BF"].forEach(c => assert.ok(!glassAddrs.some(a => a.indexOf(c) === 0),
    "column " + c + " - one of the four the office ticks by hand - was never addressed"));
  assert.ok(glassAddrs.length > 100,
    "and that is over a real number of fills, not an empty log");
  pass("every fill of the whole run landed in DG, TG, TUFF or NOT TUFF, never in the four hand-ticked columns");

  /* which sheets were written at all. The colour writer touches Production and
     nothing else; the office's own tick in section 10 is what puts Dashboard
     Log on this list, and it is the dashboard's own sheet (CLAUDE.md rule 2).
     CHANGED 2026-09-11: `Dashboard Progress` is no longer on it. The exact
     count goes to the `Dashboard progress` LIST now and the sheet of that name
     is never written by anything again. */
  const sheetsWritten = {};
  ALLREQ.filter(r => r.method !== "GET").forEach(r => {
    const m = /worksheets\('([^']+)'\)/.exec(r.path);
    if (m) sheetsWritten[m[1]] = 1;
  });
  assert.deepStrictEqual(Object.keys(sheetsWritten).sort(),
    ["Dashboard Log", "Production"],
    "no sheet outside Production and the dashboard's own Log was written at all");
  assert.ok(!Object.keys(BOOK).some(s => /^Production \(2\)|PA Lam|Glass x|Glazing|Dashboard Config/.test(s)),
    "and no other sheet was even created");
  pass("no sheet but Production and the dashboard's own were touched, and Dashboard Config never at all");

  const bodies = ALLREQ.filter(r => r.body).map(r => JSON.stringify(r.body)).join(" ");
  ["eircode", "Eircode", "phone", "Phone", "county", "County", "price", "Price"].forEach(w =>
    assert.ok(bodies.indexOf(w) < 0, "no request body ever carried " + w));
  /* "@" alone is the text number format the dashboard's own sheets are written
     with, so the test that matters is an address rather than the character */
  assert.ok(!/[a-z0-9._-]+@[a-z0-9.-]+\.[a-z]{2,}/i.test(bodies),
    "and no email address was ever written anywhere");
  const mySrc = fs.readFileSync(__filename, "utf8");
  assert.ok(!/@[a-z0-9-]+\.(com|ie|net|org)/i.test(mySrc),
    "and no real address or domain in this file either");
  pass("nothing personal was ever written, and this file names no real person, address or domain");

  console.log("\n" + n + " checks passed");
  process.exit(0);                 // the 45 s reconcile timer would hold the process open
})().catch(e => { console.error("FAIL", e); process.exit(1); });
