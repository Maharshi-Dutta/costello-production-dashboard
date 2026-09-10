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
const LISTS_PATH = "/sites/" + FSITE + "/lists";
const ITEMS_PATH = LISTS_PATH + "/" + GLASS_LIST_ID + "/items/";
const LISTWRITES = [];                        // every write this suite sends to that list
let FAIL_LIST = 0;                            // n list writes to refuse with a 403

function route(method, path, body) {
  if (path === LISTS_PATH + "?$select=id,displayName")
    return ok({ value: [{ id: GLASS_LIST_ID, displayName: "Glass station" }] });
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
    res = ok({ responses: body.requests.map(q => Object.assign({ id: q.id }, route(q.method, q.url, q.body))) });
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
CW._setFile({ base: "/x/workbook", content: "/x/content", meta: "/x" });
CW._setStationSite(FSITE);                    // the floor's lists, already found
run("checkpoints.js");
global.CP = window.CP;
run("station-core.js");
global.ST = window.ST;
run("app.js");

const TOASTS = [];
global.toast = (m, err) => TOASTS.push({ m: String(m), err: !!err });

const A = vm.runInThisContext.bind(vm);        // reach app.js' own let-bound state
const settle = ms => new Promise(r => setTimeout(r, ms == null ? 60 : ms));
const reset = () => { CALLS.length = 0; TOASTS.length = 0; };
const fills = () => CALLS.filter(c => c.kind === "fill");
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
  return byId(job.id);
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

  /* the hold, and letting it go when the file catches up */
  assert.strictEqual(A("PENDING['R7001'].gc.dg"), "yellow", "the colour is held while the file catches up");
  assert.strictEqual(byId("R7001").cp.glass.dg, "process",
    "so the drawer and the row already read yellow, 36 seconds before the download does");
  /* the download still says nothing: the hold has to survive it */
  A("ALL = applyPending([__j], true)");
  assert.strictEqual(A("PENDING['R7001'] && PENDING['R7001'].gc.dg"), "yellow",
    "a stale download does not drop the hold");
  global.__j2 = mkJob({ cp: { win: "", drs: "", glass: { dg: "process", tg: "process", "not tuff": "process" }, prod: {} } });
  A("ALL = applyPending([__j2], true)");
  assert.strictEqual(A("PENDING['R7001']"), undefined,
    "and the download agreeing lets every one of them go at once");
  assert.strictEqual(byId("R7001").cp.glass.dg, "process", "with the sheet itself now saying it");
  reset();
  assert.strictEqual(await glassColourRun(), 0);
  assert.strictEqual(CALLS.length, 0, "and still nothing to write, now from the file rather than the hold");
  pass("a colour is held until the downloaded file agrees, then let go - and never re-written either way");

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
  /* the hold now says yellow. The download is still 36 s behind and still says
     gold - which is exactly the moment a flicker would happen. */
  A("ALL = applyPending(ALL, true)");
  assert.strictEqual(byId("R7001").cp.glass.dg, "process",
    "the screen keeps the new colour rather than flickering back to the old one");
  reset();
  assert.strictEqual(await glassColourRun(), 0, "and the writer does not paint it again either");
  assert.strictEqual(CALLS.length, 0);
  pass("a reversal is held like any other write: no flicker, and no second write while it is held");

  /* and all the way back to nothing */
  global.__items = [row({ Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0, DoneAt: isoAt(16, 0) })];
  A("PENDING = {}; savePending(); STATION_ITEMS = __items; ALL = applyPending([__j], true)");
  reset();
  assert.strictEqual(await glassColourRun(), 1);
  await settle();
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AY7", WHITE], ["AZ7", WHITE], ["BA7", WHITE], ["BB7", WHITE]],
    "a job tapped back to nought leaves its glass cells with no colour on them");
  assert.strictEqual(A("PENDING['R7001'].gc.dg"), "", "held as nothing, which is a state like any other");
  assert.strictEqual(byId("R7001").cp.glass.dg, "", "and the drawer reads it as untouched");
  pass("gold, yellow, blank: white is written rather than the fill cleared, so the row keeps its own look");

  /* a cell that already has no colour is not painted white for the sake of it */
  global.__j3 = mkJob();                        // no cp statuses at all: blank cells
  global.__items = [row({ Cut: 1, DoneAt: isoAt(16, 0) })];
  A("PENDING = {}; savePending(); ALL = [__j3]; STATION_ITEMS = __items");
  reset();
  assert.strictEqual(await glassColourRun(), 0, "the floor has done something, but nothing that shows yet");
  assert.strictEqual(CALLS.length, 0, "so no cell is painted white to say so");
  pass("a blank cell that should be blank is left alone: white and no fill are the same nothing");

  /* ================= 5. a job the floor has never tapped ================= */
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11, DoneAt: "" })];
  A("PENDING = {}; savePending(); ALL = [__j3]; STATION_ITEMS = __items");
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
  const contest = async (officeWhen, floorWhen, opts) => {
    /* one column only, so the answer is about the contest and not about the
       three other cells the same job would also be arguing over */
    const cp = { win: "", drs: "", glass: { dg: "done" }, prod: {} };
    global.__jc = mkJob({ glass: { dg: 4 }, cp: cp });
    global.__items = [row(Object.assign({ Cut: 8, Hotmelt: 8, Glazed: 0, DoneAt: floorWhen }, opts || {}))];
    A("PENDING = {}; savePending(); ALL = [__jc]; STATION_ITEMS = __items;" +
      "CHANGES = []; state.sel = null;");
    CP.cpSetProgress(officeWhen
      ? { R7001: { "glass:dg": { done: 4, total: 4, who: "the colleague", when: officeWhen } } }
      : {});
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

  /* the same minute, which is the case that really happens. The office's
     record has no seconds in it, so a tick made anywhere in 15:00 is written
     down as "15:00" and covers the whole of that minute: a floor tap at
     15:00:20 loses to it, because the office may well have acted at 15:00:40.
     Reading the stamp literally instead - which is what this suite used to
     assert - handed that job to the floor and painted the office's own tick
     back out, the one thing the owner forbade. */
  assert.strictEqual(await contest(officeAt(15, 0), isoAt(15, 0)), 0,
    "the office's minute covers the floor's tap at the top of it");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  assert.strictEqual(await contest(officeAt(15, 0), isoAt(15, 0, 20)), 0,
    "and twenty seconds into that minute as well: the office's tick could have been at 15:00:40");
  await settle();
  assert.strictEqual(CALLS.length, 0, "not one request went out to argue about it");
  assert.strictEqual(await contest(officeAt(15, 0), isoAt(15, 0, 59)), 0,
    "right up to the last second of it");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  /* and the boundary, which is what makes the minute a minute rather than a
     licence: the next one is decisive */
  assert.strictEqual(await contest(officeAt(15, 0), isoAt(15, 1, 0)), 1,
    "the floor tapping in the NEXT minute wins it");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  pass("the office's minute-precision stamp covers its whole minute; the next minute beats it");

  /* the round trip, which is where this used to change its mind. noteChange
     puts a full-second entry in CHANGES straight away; load() then drops it
     once the Log sheet carries the same line at minute precision. Both
     readings of one tick must answer the same side, or the tab that made the
     tick wins for half a minute and is then quietly overruled. */
  global.__jr = mkJob({ glass: { dg: 4 }, cp: { win: "", drs: "", glass: { dg: "done" }, prod: {} } });
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 0, DoneAt: isoAt(15, 0, 20) })];
  const asLocal = () => { A("PENDING = {}; savePending(); ALL = [__jr]; STATION_ITEMS = __items;"); };
  CP.cpSetProgress({});
  /* first: the fresh local entry, seconds and all */
  asLocal();
  A("CHANGES = [{ at: " + JSON.stringify(new Date(2026, 8, 10, 15, 0, 40).toISOString()) +
    ", who: 'x', job: 'R7001', what: 'Glass DG', from: '', to: '' }]");
  reset();
  assert.strictEqual(await glassColourRun(), 0, "the tab that made the tick keeps its own gold");
  /* then: the same tick, read back off the Log sheet with the seconds gone */
  asLocal();
  A("CHANGES = [{ at: '10/09/2026 15:00', who: 'x', job: 'R7001', what: 'Glass DG', from: '', to: '', shared: true }]");
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "and so does every other dashboard once the Log line is all that is left of it");
  assert.strictEqual(CALLS.length, 0);
  pass("one office tick reads the same before and after the round trip: no delayed change of mind");

  /* one side with no stamp at all */
  assert.strictEqual(await contest("", isoAt(15, 0)), 1,
    "an office tick from before this feature existed leaves no dated record, so it cannot be the later action");
  await settle();
  assert.strictEqual(fillOn("AY7"), YELLOW);
  assert.strictEqual(await contest(officeAt(16, 0), ""), 0,
    "and a floor row with no stamp is never painted whatever the office has done");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  pass("an undated office tick loses to a dated floor tap; an undated floor row never writes at all");

  /* the Dashboard Log is the other half of the office's own record */
  global.__jl = mkJob({ glass: { dg: 4 }, cp: { win: "", drs: "", glass: { dg: "done" }, prod: {} } });
  global.__items = [row({ Cut: 8, Hotmelt: 8, Glazed: 0, DoneAt: isoAt(15, 0) })];
  A("PENDING = {}; savePending(); ALL = [__jl]; STATION_ITEMS = __items;" +
    "CHANGES = [{ at: '10/09/2026 16:00', who: 'the colleague', job: 'R7001', what: 'Glass DG', from: '', to: '4 of 4', shared: true }];");
  CP.cpSetProgress({});
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "a Log line is a dated office action even with no Progress row beside it");
  assert.strictEqual(CALLS.length, 0);
  A("CHANGES = [{ at: '10/09/2026 16:00', who: 'x', job: 'R7001', what: 'Glass: all done', from: '', to: '' }]");
  assert.strictEqual(await glassColourRun(), 0, "and a whole-group tick counts for every one of the four");
  A("CHANGES = [{ at: '10/09/2026 16:00', who: 'x', job: 'R7001', what: 'Windows', from: '', to: '' }]");
  assert.strictEqual(await glassColourRun(), 1, "while a line about the windows says nothing about the glass");
  await settle();
  A("CHANGES = [{ at: '10/09/2026 16:00', who: 'x', job: 'R7002', what: 'Glass DG', from: '', to: '' }]");
  A("PENDING = {}; savePending(); ALL = [__jl];");
  assert.strictEqual(await glassColourRun(), 1, "and neither does one about another job");
  await settle();
  /* the index the run builds once, rather than walking the log per column */
  A("CHANGES = [{ at: '10/09/2026 16:00', who: 'x', job: 'r7001', what: 'Glass DG', from: '', to: '' }," +
    "{ at: '10/09/2026 17:00', who: 'x', job: 'R7001', what: 'Glass: all done', from: '', to: '' }," +
    "{ at: '10/09/2026 18:00', who: 'x', job: 'R7001', what: 'Windows', from: '', to: '' }]");
  const idx = glassLogStamps();
  assert.deepStrictEqual(Object.keys(idx), ["R7001"], "one entry per job, upper-cased");
  assert.deepStrictEqual(Object.keys(idx.R7001).sort(), ["*", "dg", "job"],
    "the column's own line under its own key, the whole-group line under a star, " +
    "the newest of ANY of them under job, and the windows line nowhere");
  assert.strictEqual(idx.R7001["*"], stampMs("10/09/2026 17:00"));
  assert.strictEqual(idx.R7001.job, stampMs("10/09/2026 17:00"),
    "the job stamp is the newest office action on this job's glass, whichever column it named");
  assert.strictEqual(idx.R7001.dg, stampMs("10/09/2026 16:00"), "and the column keeps its own");
  pass("Dashboard Log lines count as the office's own stamp - the right job, the right item, and no other");

  /* ================= 6b. an office tick that has not landed yet =============
     THE OWNER'S BUG, 2026-09-10: "when i am marking all done it working fine
     ... but when i am trying to remove the all done ... it refreshing back to
     all golden".

     Every record the contest above is decided on is written by the WRITE, not
     by the click: Dashboard Progress at the start of it, Dashboard Log at the
     end. Between the two - one round trip for a group, 800 ms of debounce plus
     a round trip for one item - the office HAS acted and nothing the planner
     can see says so, so the floor wins by default.

     It is not symmetrical, which is exactly what the owner reported. Marking
     glass DONE moves the cell towards what the floor already says, so there is
     nothing for the writer to plan and no race to lose. UN-marking moves it
     away, so the writer always has a plan, and its gold lands after the
     office's white (both go on the job's own cpChain). The cell is then the
     colour the floor wants, so `have === want` and it is never written again:
     the un-tick is gone, silently and for good.                             */
  const untickScene = () => {
    const jj = scene(mkJob({ cp: { win: "", drs: "", glass:
                       { dg: "done", tg: "done", tuff: "done", "not tuff": "done" }, prod: {} } }),
      row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11,
            DoneAt: new Date(Date.now() - 3600000).toISOString(), DoneBy: "Person A" }));
    /* the office's own record of having marked it done, hours ago - the state
       the owner is actually in when they come to un-tick it */
    CP.cpSetProgress({ R7001: {
      "glass:dg": { done: 4, total: 4, who: "the colleague", when: "2026-09-10 00:00" },
      "glass:tg": { done: 4, total: 4, who: "the colleague", when: "2026-09-10 00:00" },
      "glass:tuff": { done: 11, total: 11, who: "the colleague", when: "2026-09-10 00:00" },
      "glass:not tuff": { done: 8, total: 8, who: "the colleague", when: "2026-09-10 00:00" } } });
    ST.COLOUR_TYPES.forEach(t => { BOOK["Production"].fill[kk(7, COL[t])] = GOLD; });
    return jj;
  };

  untickScene();
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "before the un-tick there is nothing to do: the sheet is gold and so is the floor");

  /* the office clears the whole group, and the ten-second station poll lands
     while that write is still in the air */
  const untick = setGroupDone(byId("R7001"), "glass", false);   // deliberately not awaited
  const raced = await glassColourRun();
  await untick;
  await settle(300);
  assert.strictEqual(raced, 0,
    "a poll landing while the office's un-tick is in the air must not plan the floor's gold over it");
  assert.ok(!fills().some(c => c.color === GOLD),
    "and no gold was written back: " + JSON.stringify(fills().map(c => c.addr + "=" + c.color)));
  ST.COLOUR_TYPES.forEach(t => assert.strictEqual(sheetNow()[t], WHITE,
    t + " stays as the office left it, not repainted gold behind them"));
  pass("an un-tick still in the air is the office's action already: a poll cannot paint over it");

  /* the same thing one item at a time, where the window is longer still: the
     drawer's stepper debounces for 800 ms before it writes anything at all */
  untickScene();
  reset();
  setItemProgress(byId("R7001"), "glass:dg", 0);
  const racedItem = await glassColourRun();                     // inside the debounce
  await settle(1400);
  assert.strictEqual(racedItem, 0,
    "the same during the stepper's 800 ms debounce, when nothing has been written at all yet");
  assert.notStrictEqual(fillOn("AY7"), GOLD, "DG was not repainted gold under the office's hand");
  assert.strictEqual(sheetNow().dg, WHITE, "it is white, which is what the office asked for");
  pass("a per-item un-tick is safe through its debounce as well as through its write");

  /* CHANGED 2026-09-10 BY THE OWNER'S RESTATED RULE. This used to assert that
     a floor tap made two seconds after the office's click won, because the
     hold is stamped at the click and last-writer-wins decided it. It no longer
     does, and that is the point of the change: "an un-tick from the dashboard
     is absolute - no thinking, no arguing". Inside the settling window the
     writer stands down for that job whatever the floor's row says, because
     inside that window it cannot tell a genuine tap from a stale copy of the
     office's own change.

     The tap is not lost. Nothing here is one-shot: the writer re-plans from
     whatever the list says on every pass, so once the window has gone by, that
     tap - if it is still what the floor's row says - is painted then. Section
     10d test 3 is where that is asserted. */
  untickScene();
  reset();
  const untick2 = setGroupDone(byId("R7001"), "glass", false);
  await settle(5);
  global.__later = [row({ Cut: 8, Hotmelt: 8, Glazed: 8, Tuff: 11,
                          DoneAt: new Date(Date.now() + 2000).toISOString() })];
  A("STATION_ITEMS = __later");
  const afterTap = await glassColourRun();
  await untick2;
  await settle(300);
  assert.strictEqual(afterTap, 0,
    "a floor tap two seconds after the office clicked no longer wins: the office is absolute here");
  assert.strictEqual(A("officeSettling(byId('R7001'))"), true, "because the office is still settling");
  pass("inside the office's window the writer stands down, whatever the floor's row says");
  A("PENDING = {}; savePending(); CHANGES = [];");
  CP.cpSetProgress({});

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
  reset();
  FAIL_FILL = 20;
  const inFlight = glassColourRun();
  assert.strictEqual(A("PENDING['R7001'].gc.dg"), "yellow", "held the moment it was sent");
  assert.strictEqual(await inFlight, 1);
  FAIL_FILL = 0;
  assert.strictEqual(A("PENDING['R7001']"), undefined,
    "and the hold let go when the write failed: the sheet still says the old colour, and so does the screen");
  assert.strictEqual(byId("R7001").cp.glass.dg || "", "", "which is what it said before");
  pass("a refused fill drops its own hold rather than showing a colour the sheet does not have");

  /* THE BRAKE. Without it a failing write goes out again on the very next
     pass, and stationPoll answers "moved" every ten seconds for as long as the
     floor keeps tapping - a write storm against the live workbook, which is
     the worst thing this feature has available. The trigger is ordinary:
     somebody opens the file exclusively in desktop Excel mid-shift. */
  assert.strictEqual(A("GLASSC_FAIL['R7001'].n"), 1, "the failure is counted against that job");
  failScene();
  reset();
  assert.strictEqual(await glassColourRun(), 0, "so the next pass does not send it again");
  assert.strictEqual(CALLS.length, 0, "not one request");
  assert.ok(footWord().indexOf("glass colours not saved") >= 0,
    "and the footer says so, rather than only the console");
  assert.ok(footWhy().indexOf("Production sheet") > 0 && footWhy().indexOf("trying again shortly") > 0,
    "with what happened and what happens next in the tooltip");
  pass("a refused write is not retried on the next poll: the job serves a backoff, and the footer shows it");

  /* it does come back, once the backoff has run */
  ageFail("R7001", 120000);
  failScene();
  reset();
  FAIL_FILL = 20;
  assert.strictEqual(await glassColourRun(), 1, "a minute later it tries again");
  FAIL_FILL = 0;
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
  FAIL_FILL = 20;
  assert.strictEqual(await glassColourRun(), 1, "five minutes is");
  FAIL_FILL = 0;
  pass("the backoff grows: one minute, then five, then fifteen");

  /* and it stops. A workbook that has refused this five times is not going to
     take it on the sixth, and the person looking at the screen is told rather
     than left with a dashboard quietly hammering the file. */
  FAIL_FILL = 40;
  for (let i = 0; i < 6; i++) {
    A("if (GLASSC_FAIL['R7001']) GLASSC_FAIL['R7001'].at -= 3600000");
    failScene();
    await glassColourRun();
  }
  FAIL_FILL = 0;
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
  assert.ok(CALLS.some(c => c.kind === "values" && c.sheet === "Dashboard Progress"),
    "and still writes the exact count to Dashboard Progress, which the colour writer never does");
  assert.ok(CALLS.some(c => c.kind === "values" && c.sheet === "Dashboard Log"),
    "and still logs it");
  pass("the drawer's per-type ticks are exactly what they were: colour, count and log line");

  /* both sides holding the same cell at once: the newer hold is what shows */
  global.__jb = mkJob();
  A("PENDING = {}; savePending(); ALL = [__jb]; STATION_ITEMS = __items; CHANGES = [];");
  CP.cpSetProgress({});
  A("pend('R7001', { cp: { 'glass:dg': 2 } })");
  A("pend('R7001', { gc: { dg: 'gold' } })");
  A("ALL = applyPending([__jb])");
  assert.strictEqual(byId("R7001").cp.glass.dg, "done", "the floor's hold is the newer one, so it shows");
  assert.strictEqual((byId("R7001").cpDone || {})["glass:dg"], undefined,
    "and the count the office was holding is dropped, so the drawer does not read a number the cell denies");
  A("PENDING = {}; savePending();");
  A("pend('R7001', { gc: { dg: 'gold' } })");
  A("PENDING['R7001'].t['gc:dg'] -= 5000");
  A("pend('R7001', { cp: { 'glass:dg': 2 } })");
  A("ALL = applyPending([__jb])");
  assert.strictEqual(byId("R7001").cp.glass.dg, "process", "and the other way round, the office's tick shows");
  assert.strictEqual((byId("R7001").cpDone || {})["glass:dg"], 2, "with its count");
  A("PENDING = {}; savePending();");
  pass("when both sides are holding one cell, the newer hold is drawn - the same rule as the writer's");

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
  const clearIdx = A("(function(){ var keep = CHANGES; " +
    "CHANGES = CHANGES.filter(function (c) { return c.what === 'Floor glass counters'; }); " +
    "var r = JSON.stringify(glassLogStamps()); CHANGES = keep; return r; })()");
  assert.deepStrictEqual(Object.keys(JSON.parse(clearIdx)), ["R7001"],
    "read back on its own, the clear line is an office action on that job");
  assert.deepStrictEqual(Object.keys(JSON.parse(clearIdx).R7001), ["job"],
    "under the job, and under no single column - a clear names none");
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
  const iList = seq.findIndex(p => p.indexOf("/lists/") >= 0 && p.indexOf("/items/") >= 0);
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
  CP.cpSetProgress({ R7001: {
    "glass:dg": { done: 4, total: 4, who: "the admin", when: officeAt(15, 0) },
    "glass:tg": { done: 4, total: 4, who: "the admin", when: officeAt(15, 0) } } });
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
  CP.cpSetProgress({});
  pass("clear it, re-tick it inside the same minute, and the office keeps it - the minute a stamp covers decides it");

  Object.defineProperty(CW, "account", { configurable: true, get: () => null });
  delete global.confirm;


  /* ================= 10c. the office's stamp is PER JOB =====================
     OBSERVED IN PRODUCTION, 2026-09-10, and reproduced in a stubbed browser
     with the download lagged 36 s ("case C"). The office pressed Clear on TG
     alone. Three `[glass] painting 1 job` lines followed, and the drawer's own
     TUFF (20 of 20) and NOT TUFF (29 of 29) - ticked by hand, in the office,
     and never mentioned by that click - went to nought.

     Why: glassOfficeStamp was asked PER COLUMN. For `tuff` and `not tuff` it
     found no hold, no Dashboard Progress row and no Dashboard Log line, and
     answered a literal 0, so any floor stamp at all beat it. And the floor
     stamp it lost to was one THE OFFICE ITSELF HAD JUST WRITTEN:
     clearFloorGlass's own DoneAt, a fraction of a second after the click -
     while glassLogStamps deliberately refused to count the "Floor glass
     counters" line that records the same act as an office action. One half of
     the office's own clear was counted for the floor and the other half for
     nobody.

     The rule these tests pin: the office's control is per JOB - a clear is a
     job-level act, and the floor's row is one combined number - so an office
     action at time T on a job cannot lose to a floor stamp on a column the
     office happened not to name. A genuine floor tap AFTER it still wins. */
  /* Everything here is relative to NOW rather than to a fixed wall clock: a cp
     hold older than PENDING_MS is dropped by applyPending, so a click stamped
     in 2026 in a suite run at any other moment would not be a hold at all. T0
     is the click; the clear's own write lands after it, as it does in
     production. */
  const T0 = Date.now() - 5000;
  const isoMs = ms => new Date(ms).toISOString();
  const two = n => (n < 10 ? "0" : "") + n;
  /* what nowStamp() writes into Dashboard Progress: a minute, no seconds */
  const minStamp = ms => { const d = new Date(ms);
    return d.getFullYear() + "-" + two(d.getMonth() + 1) + "-" + two(d.getDate()) +
           " " + two(d.getHours()) + ":" + two(d.getMinutes()); };
  const CLICK = isoMs(T0);
  const officeMin = minStamp(T0);                      // what nowStamp() would write
  /* the office's own record of the clear, as clearFloorGlass keeps it. Guarded
     so that this suite fails on the ASSERTION rather than on a missing name
     when it is run against the build that has the bug. */
  const rememberClear = at => A("if (typeof OFFICE_FLOOR_AT !== 'undefined') " +
    "OFFICE_FLOOR_AT['R7001'] = " + JSON.stringify(at) + ";");
  /* the job of the reproduction: TG is its only DG/TG glass, and TUFF and NOT
     TUFF are ticked by hand in the office */
  const caseCJob = () => mkJob({ glass: { tg: 49, tuff: 20, "not tuff": 29 },
    cp: { win: "", drs: "", glass: { tg: "done", tuff: "done", "not tuff": "done" }, prod: {} } });
  /* the row as clearFloorGlass leaves it: four noughts, and the DoneAt it
     wrote itself */
  const clearedRow = doneAt => row({ Total: 49, TuffTotal: 20,
    Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0, DoneBy: "the admin", DoneAt: doneAt });
  /* the office is holding TG at nought, stamped at the click, as pend() does */
  const holdTg = () => {
    A("PENDING = {}; savePending(); pend('R7001', { cp: { 'glass:tg': 0 } });");
    A("PENDING['R7001'].t['cp:glass:tg'] = " + Date.parse(CLICK) +
      "; PENDING['R7001'].at = " + Date.parse(CLICK) + "; savePending();");
    A("ALL = applyPending(ALL, true);");
  };

  /* (1) CASE C, as recorded: the clear's DoneAt landed 0.8 s after the click,
     the office still holds TG at nought, and TUFF and NOT TUFF - which that
     click never mentioned - must not be touched. */
  j = scene(caseCJob(), clearedRow(isoMs(T0 + 800)));
  holdTg();
  rememberClear(isoMs(T0 + 800));
  reset();
  console.log("DBGC plan", JSON.stringify(A("glassColourPlan(byId('R7001'))")),
    "have", JSON.stringify(A("byId('R7001').cp.glass")),
    "floorAt", A("stampMs(ST.floorStamp(stationForJob('R7001')))"),
    "stamps", JSON.stringify(["tg","tuff","not tuff"].map(t => A("glassOfficeStamp('R7001','" + t + "')"))),
    "click", Date.parse(CLICK));
  assert.strictEqual(await glassColourRun(), 0,
    "case C: the office cleared TG, so nothing of this job's glass is painted from the floor");
  await settle();
  assert.strictEqual(CALLS.length, 0, "not one cell is written");
  assert.strictEqual(sheetNow().tuff, undefined, "TUFF is untouched");
  assert.strictEqual(sheetNow()["not tuff"], undefined, "and so is NOT TUFF");
  pass("case C: a per-row Clear on TG no longer wipes the TUFF and NOT TUFF the office ticked by hand");

  /* (2) PRODUCTION ORDERING: the write really lands AFTER the click - three
     seconds after it in the owner's console - so there is no tie to hide
     behind. The office must still win on all three columns. */
  j = scene(caseCJob(), clearedRow(isoMs(T0 + 3000)));
  holdTg();
  rememberClear(isoMs(T0 + 3000));
  reset();
  assert.ok(stampMs(isoMs(T0 + 3000)) > stampMs(CLICK),
    "the DoneAt really is later than the click: this is not a tie");
  assert.strictEqual(await glassColourRun(), 0,
    "and the office's own write still counts as the office's, so the office wins");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  pass("the clear's own DoneAt is the office's action, not the floor's - even three seconds after the click");

  /* (3) THE RELEASED-HOLD WINDOW, which is the route to the GOLD the owner
     reported. The downloaded file has caught up on the Production fill (the
     cell is white, so the cp hold is let go) but does not yet carry the
     Dashboard Progress row or the Dashboard Log line for that clear; and this
     dashboard's copy of the floor's list is a few seconds stale, still showing
     the FULL counters against the fresh DoneAt. Every per-column source
     answers nothing, the row looks finished, and without a job-level stamp the
     writer paints the job GOLD - over an un-tick the office has just made. */
  global.__jw = mkJob({ glass: { tg: 49, tuff: 20, "not tuff": 29 },
    cp: { win: "", drs: "", glass: {}, prod: {} } });
  j = scene(caseCJob(), row({ Total: 49, TuffTotal: 20,
    Cut: 49, Hotmelt: 49, Glazed: 49, Tuff: 20,          // the stale read: still full
    DoneBy: "the admin", DoneAt: isoMs(T0 + 3000) }));
  A("PENDING = {}; savePending(); ALL = [__jw]; CHANGES = [];");   // the file agrees: white
  CP.cpSetProgress({});                                            // no Progress row yet either
  rememberClear(isoMs(T0 + 3000));
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "the office's own clear is remembered, so a stale full-count read cannot paint the job gold");
  await settle();
  assert.strictEqual(CALLS.length, 0, "no gold, no fill of any kind");
  /* the same window after a reload, when the office's own memory of the clear
     is gone and only the Dashboard Log line survives: it is an office action
     on this job and must count as one. Built from scratch rather than carried
     on from the case above, so that a write the previous state had already
     made could not make this one look quiet. */
  await settle(200);
  j = scene(caseCJob(), row({ Total: 49, TuffTotal: 20,
    Cut: 49, Hotmelt: 49, Glazed: 49, Tuff: 20,
    DoneBy: "the admin", DoneAt: isoMs(T0 + 3000) }));
  global.__ch = [{ at: isoMs(T0 + 3300), who: "the admin", job: "R7001",
                   what: "Floor glass counters", from: "49 cut, 49 hotmelted, 49 glazed, 20 tuff",
                   to: "nothing", src: "dashboard" }];
  A("PENDING = {}; savePending(); ALL = [__jw]; CHANGES = __ch;");
  CP.cpSetProgress({});
  forgetClears();
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "and after a reload the log line alone still says the office did this");
  await settle();
  assert.strictEqual(CALLS.length, 0);
  A("CHANGES = [];");
  pass("the released-hold window cannot paint gold over a fresh un-tick, in this session or after a reload");

  /* (4) AND LAST-WRITER-WINS IS STILL LAST-WRITER-WINS: a real floor tap after
     the office's clear carries a later stamp than every office record on the
     job, and its colour is painted. Without this the fix would simply be "the
     office always wins", which is not the rule. */
  /* the office's clear is well behind us - past the settling window of section
     10d, which would otherwise stand the writer down whatever the stamps say -
     and the floor tapped a minute and a half after it */
  const OLD = Date.now() - 260000, TAP = Date.now() - 100000;
  j = scene(caseCJob(), row({ Total: 49, TuffTotal: 20,
    Cut: 49, Hotmelt: 49, Glazed: 0, Tuff: 20,
    DoneBy: "Person A", DoneAt: new Date(TAP).toISOString(),
    CutBy: "Person A", CutAt: new Date(TAP).toISOString() }));
  A("PENDING = {}; savePending(); CHANGES = [];");
  CP.cpSetProgress({ R7001: {
    "glass:tg": { done: 0, total: 49, who: "the admin", when: minStamp(OLD) } } });
  rememberClear(new Date(OLD).toISOString());
  reset();
  assert.strictEqual(await glassColourRun(), 1, "the floor tapped after the office: the floor wins");
  await settle();
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AZ7", YELLOW], ["BA7", YELLOW], ["BB7", YELLOW]],
    "and its yellow reaches every glass column this job has");
  CP.cpSetProgress({});
  forgetClears();
  A("PENDING = {}; savePending();");
  pass("a genuine floor tap after the office's clear still wins: the rule is last writer, not office first");


  /* ================= 10d. the office is absolute: the writer stands down =====
     THE OWNER'S RULE, restated 2026-09-10: "the hierarchy is Excel, then master
     dashboard, then glass. Any change from the dashboard is absolute. If the
     glass updates the ticks it comes golden instantly, correct, and should not
     change. An un-tick from the dashboard is absolute - no thinking, no
     arguing."

     So the writer does not argue. While an office change on a job's glass is
     still settling, it makes NO decision about that job at all - it does not
     compare stamps, it does not plan, it does not write. The reason it cannot
     be allowed to decide is simple: inside that window it cannot tell a
     genuine floor tap from a stale copy of the office's own change coming back
     round (the download is ~36 s behind, the floor's list a poll behind, and
     the clear's own DoneAt looks exactly like a tap). Stamps were refined
     twice; the owner still saw the gold come back. This stops the argument
     rather than trying to win it.

     Outside the window nothing changes: the floor's work still reaches the
     sheet, and last-writer-wins still decides it (test 3 below). */
  const N = Date.now();
  const win = ms => new Date(N - ms).toISOString();
  const qJob = () => mkJob({ glass: { tg: 49, tuff: 20, "not tuff": 29 },
    cp: { win: "", drs: "", glass: { tg: "done", tuff: "done", "not tuff": "done" }, prod: {} } });
  const fullRow = doneAt => row({ Total: 49, TuffTotal: 20,
    Cut: 49, Hotmelt: 49, Glazed: 49, Tuff: 20, DoneBy: "Person A", DoneAt: doneAt });
  /* the office's own records of an un-tick made `ago` milliseconds back: the
     hold at the click, the Log line the write left, and the clear it made on
     the floor's row. No Dashboard Progress row: within the first half minute
     the downloaded file does not carry one yet, which is the state this is
     about. */
  const officeUntick = (ago, hold) => {
    A("PENDING = {}; savePending();");
    if (hold) {
      A("pend('R7001', { cp: { 'glass:tg': 0 } });");
      A("PENDING['R7001'].t['cp:glass:tg'] = " + (N - ago) + "; PENDING['R7001'].at = " + (N - ago) + ";");
      A("savePending(); ALL = applyPending(ALL, true);");
    }
    global.__qch = [{ at: win(ago - 200), who: "the admin", job: "R7001", what: "Glass TG",
                      from: "49 of 49", to: "0 of 49", src: "dashboard" },
                    { at: win(ago - 350), who: "the admin", job: "R7001",
                      what: "Floor glass counters", from: "49 cut, 49 hotmelted, 49 glazed, 20 tuff",
                      to: "nothing", src: "dashboard" }];
    A("CHANGES = __qch;");
    A("if (typeof OFFICE_FLOOR_AT !== 'undefined') OFFICE_FLOOR_AT['R7001'] = " +
      JSON.stringify(win(ago - 300)) + ";");
  };

  /* (1) THE OBSERVED CASE. The office un-ticks TG. A moment later the floor's
     row reads FULL counters again with a DoneAt LATER than everything the
     office has - the tablet's own poll, a queued tap landing, or a list read
     that has not settled. Judged on stamps that is a floor tap and the writer
     paints the job gold again, which is exactly what the owner keeps seeing.
     Judged by the owner's rule there is nothing to judge: the office spoke,
     and the writer stands down until that has settled everywhere. */

  /* (a) at once, with the hold still on */
  j = scene(qJob(), fullRow(win(1000)));
  officeUntick(2000, true);
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "t+0: the office un-ticked this job, so the writer decides nothing about it");
  await settle();
  assert.strictEqual(CALLS.length, 0, "and writes nothing at all");
  /* deferred is not dropped: the pass arms the follow-up it already uses for a
     capped run, so the colour lands once the window has closed even if nothing
     else happens to move in the meantime */
  assert.ok(A("!!glassAgainT"), "but it arms the follow-up, so a deferred job is not left for ever");
  A("if (glassAgainT) { clearTimeout(glassAgainT); glassAgainT = null; }");

  /* (b) ten seconds later, the hold still on and the file still behind */
  j = scene(qJob(), fullRow(win(1000)));
  officeUntick(12000, true);
  reset();
  assert.strictEqual(await glassColourRun(), 0, "t+10 s: still the office's window, still nothing");
  await settle();
  assert.strictEqual(CALLS.length, 0);

  /* (c) the hold has been let go because the download agrees the cell is
     white - the stamp is the only thing left saying the office acted, and it
     has to be enough */
  global.__qw = mkJob({ glass: { tg: 49, tuff: 20, "not tuff": 29 },
    cp: { win: "", drs: "", glass: {}, prod: {} } });
  j = scene(qJob(), fullRow(win(1000)));
  officeUntick(45000, false);
  A("ALL = [__qw];");
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "the hold is gone but the office's action is still settling: still no decision");
  await settle();
  assert.strictEqual(CALLS.length, 0, "and above all no gold painted back over the un-tick");
  pass("an office un-tick is absolute: inside the window the writer plans nothing for that job at all");

  /* (2) THE TICK SIDE. "If the glass updates the ticks it comes golden
     instantly, correct, and should not change." The office marks the glass
     done; the floor's row is still at nought because the feeder has not run
     yet. There is no disagreement to resolve here - only a chance to paint
     white over the office's own gold, which must not happen. */
  j = scene(mkJob({ glass: { tg: 49, tuff: 20, "not tuff": 29 },
                    cp: { win: "", drs: "", glass: {}, prod: {} } }),
            row({ Total: 49, TuffTotal: 20, Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0,
                  DoneBy: "Person A", DoneAt: win(500) }));
  A("PENDING = {}; savePending(); pend('R7001', { cp: { 'glass:tg': 49, 'glass:tuff': 20, 'glass:not tuff': 29 } });");
  A("PENDING['R7001'].t['cp:glass:tg'] = " + (N - 1000) + "; PENDING['R7001'].at = " + (N - 1000) + ";");
  A("savePending(); ALL = applyPending(ALL, true); CHANGES = [];");
  assert.strictEqual(byId("R7001").cp.glass.tg, "done", "the office's gold is on the screen at once");
  reset();
  assert.strictEqual(await glassColourRun(), 0,
    "the office has just marked it done, so the writer says nothing about this job");
  await settle();
  assert.strictEqual(CALLS.length, 0, "no white over the office's gold");
  pass("marking done is absolute too: an office tick is never painted out by the floor's older numbers");

  /* (3) OUTSIDE THE WINDOW, LAST-WRITER-WINS IS UNTOUCHED. Three minutes after
     the office's last word on this job's glass, a genuine floor tap carries a
     later stamp and its colour reaches the sheet, exactly as before. Without
     this the guard would just be "the office always wins", which is not the
     rule and would strand the floor's work. */
  j = scene(qJob(), row({ Total: 49, TuffTotal: 20,
    Cut: 49, Hotmelt: 49, Glazed: 0, Tuff: 20,
    DoneBy: "Person A", DoneAt: win(1000), CutBy: "Person A", CutAt: win(1000) }));
  global.__och = [{ at: win(181000), who: "the admin", job: "R7001", what: "Glass TG",
                    from: "", to: "49 of 49", src: "dashboard" }];
  A("PENDING = {}; savePending(); CHANGES = __och;");
  A("if (typeof OFFICE_FLOOR_AT !== 'undefined') OFFICE_FLOOR_AT['R7001'] = " +
    JSON.stringify(win(181000)) + ";");
  reset();
  assert.strictEqual(await glassColourRun(), 1,
    "181 s on, the office's change has settled and the floor's tap is the later word");
  await settle();
  assert.deepStrictEqual(fills().map(c => [c.addr, c.color]).sort(),
    [["AZ7", YELLOW], ["BA7", YELLOW], ["BB7", YELLOW]],
    "and it is painted, on every glass column the job has");
  forgetClears();
  pass("outside the settling window the floor's work still reaches the sheet, decided as it always was");

  /* (4) AND THE WINDOW IS PER JOB. An un-tick on one job says nothing about
     another, and must not hold the floor's work off the whole sheet. */
  BOOK["Production"] = { v: {}, fill: {} };
  BOOK["Production"].v[kk(7, 3)] = "R7001";
  BOOK["Production"].v[kk(8, 3)] = "R7002";
  global.__jA = qJob();
  global.__jB = mkJob({ id: "R7002", cust: "Customer Two", src: { Production: 8 },
    glass: { tg: 10 }, cp: { win: "", drs: "", glass: { tg: "done" }, prod: {} } });
  global.__qitems = [fullRow(win(1000)),
    { id: "701", fields: { Title: "R7002", Job: "R7002", Customer: "Customer Two",
      GlassType: "GLASS", Total: 10, TuffTotal: 0, Seq: 2, Active: "Yes", OfficeDone: "No",
      Cut: 10, Hotmelt: 10, Glazed: 0, Tuff: 0, DoneBy: "Person A", DoneAt: win(1000) } }];
  A("PENDING = {}; savePending(); ALL = [__jA, __jB]; PRODMAP = __m; CHANGES = [];" +
    "STATION_ITEMS = __qitems; STATION_OK = true; BLOCKNAMES = []; state.sel = null;");
  CP.cpSetProgress({});
  clearFail(); forgetClears();
  officeUntick(2000, true);                    // the office un-ticked job A, and only job A
  A("ALL = applyPending(ALL, true);");
  reset();
  assert.strictEqual(await glassColourRun(), 1, "one job painted: the other one, not the one being cleared");
  await settle();
  assert.deepStrictEqual(fills().map(c => c.addr).sort(), ["AZ8"],
    "job B's own row and column, and nothing of job A's");
  pass("the window is one job's: an un-tick on one job never holds the floor's work off another");
  A("PENDING = {}; savePending(); CHANGES = [];");
  forgetClears();

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
     Progress and Dashboard Log on this list, and they are the dashboard's own
     sheets (CLAUDE.md rule 2). */
  const sheetsWritten = {};
  ALLREQ.filter(r => r.method !== "GET").forEach(r => {
    const m = /worksheets\('([^']+)'\)/.exec(r.path);
    if (m) sheetsWritten[m[1]] = 1;
  });
  assert.deepStrictEqual(Object.keys(sheetsWritten).sort(),
    ["Dashboard Log", "Dashboard Progress", "Production"],
    "no sheet outside Production and the dashboard's own two was written at all");
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
