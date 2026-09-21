/* Offline test of the glazing station (docs/specs/2026-09-21-glazing-station.md).

   The THIRD floor station, and the first one built from the "Adding a station"
   checklist rather than from a page. So the sharpest assertions here are not
   about glazing at all - they are about the boundaries the owner drew, and they
   are checked on the request log and on the bodies that go out on it:

     · over the entire run, not one request touches the workbook - no download,
       no /workbook path, no range, no worksheet. The ONE workbook write
       anywhere in this feature is the `Dashboard Log` line an office edit
       leaves, which is a dashboard-owned sheet (rule 2) and is asserted on its
       own;
     · the quantities come off the `Production` sheet and no other: a fixture
       whose cross-sheet numbers disagree with Production's proves it;
     · nothing the feeder sends ever carries one of the floor's own columns, and
       the feeder never issues a DELETE;
     · nothing the office's glazing edit sends carries a job fact, a Section or
       an Active - only the counter, its By/At and the last-touch pair - and it
       never writes a line of `Station log`;
     · no phone number and no eircode can reach the list: the customer and the
       COMMENT are stripped on the way in, and the strip is tested against both
       shapes;
     · the phase bar hears the floor and WRITES NOTHING for it - a whole section
       of this file asserts zero requests of any kind across every phase check.

   Graph is a fake fetch() over one in-memory SharePoint site; nothing leaves
   the box, every address in here is example.test and every person is made up.
   Run: node test_glazing.js                                                  */
const fs = require("fs"), vm = require("vm"), assert = require("assert");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null),
                        setItem: (k, v) => { mem[k] = String(v); },
                        removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, innerWidth: 1280, innerHeight: 800,
                  addEventListener() {}, removeEventListener() {} };
global.performance = { now: () => Date.now() };

const REG = {};
function stubEl(tag, id) {
  let html = "";
  const e = {
    tag: tag || "div", tagName: String(tag || "div").toUpperCase(),
    id: id || "", style: { setProperty() {} }, dataset: {}, attrs: {}, kids: [],
    textContent: "", value: "", disabled: false, hidden: false, className: "", title: "",
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); if (c.id) REG[c.id] = c; return c; },
    remove() { const drop = x => { if (x.id && REG[x.id] === x) delete REG[x.id]; x.kids.forEach(drop); }; drop(e); },
    contains(x) { return x === e || e.kids.some(k => k.contains && k.contains(x)); },
    setAttribute(k, v) { e.attrs[k] = String(v); }, getAttribute(k) { return e.attrs[k] || null; },
    on: {},
    addEventListener(t, fn) { (e.on[t] = e.on[t] || []).push(fn); },
    removeEventListener(t, fn) { e.on[t] = (e.on[t] || []).filter(x => x !== fn); },
    focus() {}, blur() {}, setSelectionRange() {},
    getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => null, querySelectorAll: () => []
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; } });
  return e;
}
const NULLABLE = ["#fabhost", "#fabbtn", "#dhost", "#xhost", "#ahost", "#chost", "#vhost",
                  "#lhost", "#dayhost", "#catmenu", "#movemenu", "#alertmenu",
                  "#weldopen", "#wq", "#wsect", "#glzopen", "#zq", "#zsect"];
const EL = {};
const el = sel => {
  const id = String(sel).charAt(0) === "#" ? String(sel).slice(1) : null;
  if (id && REG[id]) return REG[id];
  if (NULLABLE.indexOf(sel) >= 0) return null;
  return EL[sel] || (EL[sel] = stubEl("div", id || ""));
};
global.document = {
  documentElement: stubEl(), body: stubEl(), head: stubEl(),
  activeElement: null, title: "Costello Production",
  createElement: t => stubEl(t),
  querySelector: el, querySelectorAll: () => [], addEventListener() {}, removeEventListener() {}
};
document.body.appendChild = c => { if (c.id) REG[c.id] = c; return c; };

/* ---------- the fake Graph ----------
   One site: `Floor stations`, with the five lists this run touches. There is NO
   workbook route at all - a request for one is a test failure by construction,
   because there is nowhere for it to go. */
const G = "https://graph.microsoft.com/v1.0";
const SITE = "costellowindowsie.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";
const FSITE = "costellowindowsie.sharepoint.com,aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee,ffffffff-0000-1111-2222-333333333333";
const GLZ_ID = "list-glazing-station";
const WELD_ID = "list-welding-station";
const PEOPLE_ID = "list-station-people";
const LOG_ID = "list-station-log";
const CMT_ID = "list-station-comments";
const FLISTS_PATH = "/sites/" + FSITE + "/lists";
const OWNLISTS_PATH = "/sites/" + SITE + "/lists";
const HOST_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/FloorStations";
const OWN_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/ProductionProgress?$select=id,displayName";

let FLISTS = [{ id: GLZ_ID, displayName: "Glazing station" },
              { id: WELD_ID, displayName: "Welding station" },
              { id: PEOPLE_ID, displayName: "Station people" },
              { id: LOG_ID, displayName: "Station log" },
              { id: CMT_ID, displayName: "Station comments" }];
let ZITEMS = [], WITEMS = [], PEOPLEITEMS = [], LOGITEMS = [], CMTITEMS = [];
let NEXTID = 100;
let FAIL_ONCE = 0;

const REQ = [], ALLREQ = [];
const ok = body => ({ status: 200, body: body });
const tick = ms => new Promise(r => setTimeout(r, ms == null ? 1 : ms));
function item(fields, id) {
  return { id: String(id == null ? NEXTID++ : id), fields: Object.assign({}, fields) };
}
const storeFor = id => id === GLZ_ID ? ZITEMS : id === WELD_ID ? WITEMS
                     : id === PEOPLE_ID ? PEOPLEITEMS : id === LOG_ID ? LOGITEMS
                     : id === CMT_ID ? CMTITEMS : null;

function routeFlist(method, path, body) {
  const rest = path.slice(FLISTS_PATH.length);
  if (method === "GET" && rest.indexOf("?$select=id,displayName") === 0) return ok({ value: FLISTS });
  const mi = /^\/([^/?]+)\/items(.*)$/.exec(rest);
  if (!mi) return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + path } } };
  const id = mi[1], tail = mi[2];
  const store = storeFor(id);
  if (!store) return { status: 404, body: { error: { code: "itemNotFound" } } };
  if (method === "GET" && tail.indexOf("/delta") === 0)
    return ok({ value: store.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })),
                "@odata.deltaLink": G + FLISTS_PATH + "/" + id + "/items/delta?token=T1" });
  const mone = /^\/([^/?]+)(?:\?.*)?$/.exec(tail);
  if (method === "GET" && mone) {
    const hit = store.find(x => x.id === mone[1]);
    if (!hit) return { status: 404, body: { error: { code: "itemNotFound" } } };
    return ok({ id: hit.id, fields: Object.assign({}, hit.fields) });
  }
  if (method === "GET")
    return ok({ value: store.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) });
  if (FAIL_ONCE && method !== "GET") { FAIL_ONCE--; return { status: 403, body: { error: { code: "accessDenied" } } }; }
  if (method === "POST" && tail === "") {
    const made = item((body && body.fields) || {});
    store.push(made);
    return ok({ id: made.id, fields: made.fields });
  }
  const mf = /^\/([^/]+)\/fields$/.exec(tail);
  if (method === "PATCH" && mf) {
    const hit = store.find(x => x.id === mf[1]);
    if (!hit) return { status: 404, body: { error: { code: "itemNotFound" } } };
    Object.assign(hit.fields, body || {});
    return ok(Object.assign({}, hit.fields));
  }
  return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + method + " " + path } } };
}
function route(method, path, body) {
  if (path === HOST_LOOKUP) return ok({ id: FSITE, displayName: "Floor stations" });
  if (path.indexOf(OWN_LOOKUP) === 0) return ok({ id: SITE, displayName: "Production Progress" });
  if (path.indexOf(FLISTS_PATH) === 0) return routeFlist(method, path, body);
  if (path.indexOf(OWNLISTS_PATH) === 0) return ok({ value: [] });
  /* there is no workbook here, and there is no route to one */
  return { status: 599, body: { error: { code: "thisTestServesNoWorkbook", message: path } } };
}
global.fetch = async (url, init) => {
  const path = String(url).replace(G, "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  const rec = { method: init.method, path: path, body: body };
  REQ.push(rec); ALLREQ.push(rec);
  await tick();
  const res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status,
           text: async () => (res.body === "" ? "" : JSON.stringify(res.body)),
           arrayBuffer: async () => new ArrayBuffer(0) };
};

/* ---------- load the dashboard's own code, in the page's own order ---------- */
const src = f => fs.readFileSync(__dirname + "/" + f, "utf8");
const run = f => vm.runInThisContext(src(f), { filename: f });
run("parser.js");
run("graph.js");
global.CW = window.CW;
CW._setToken(() => "t");
CW._setFile({ siteId: SITE, base: "/x/workbook", content: "/x/content", meta: "/x" });
run("checkpoints.js");
global.CP = window.CP;
run("station-core.js");
global.ST = window.ST;
run("welding-core.js");
global.W = window.WELDC;
run("glazing-core.js");
global.Z = window.GLZC;
run("station-ui.js");
global.STU = window.STU;
run("export.js");
run("app.js");

const TOASTS = [];
global.toast = (m, isErr) => TOASTS.push({ m: String(m), err: !!isErr });
const A = vm.runInThisContext.bind(vm);
const consent = yes => { CW.hasListConsent = async () => !!yes; };
const reset = () => { REQ.length = 0; TOASTS.length = 0; };
const writes = () => REQ.filter(r => r.method !== "GET");

let n = 0;
const pass = msg => { n++; console.log("  ok  " + msg); };

/* ---------- fixtures ----------
   `wndMain`/`drsMain` are the `Production` sheet's own quantities, never merged
   with another sheet's (parser.js, 2026-09-21). A job that is only ever on
   `Production` has the same numbers in both, so they default to whatever
   wnd/drs were given. A fixture that wants the two to DISAGREE, the way the
   live workbook did on 2026-09-17, passes both. */
const NAMES = ["Can sell as second hand", "Ready to fit", "Collect & supply only",
               "Finished", "In production", "Not sent to floor"];
const mkJob = o => {
  const j = Object.assign({
    id: "R0001", cust: "Customer One", area: "Cork", eir: "", off: "", colour: "", ph3: "",
    wnd: 0, drs: 0, glass: {}, prods: [], prodsMain: [], notes: [], sheets: ["Production"],
    src: {}, doors: [],
    dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
    cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0,
    cp: { win: "", drs: "", glass: {}, prod: {} }
  }, o || {});
  if (!(o && "wndMain" in o)) j.wndMain = j.wnd;
  if (!(o && "drsMain" in o)) j.drsMain = j.drs;
  return j;
};

const JOBS = [
  /* six windows and two doors on Production, a customer and a comment that both
     carry things rule 3 forbids, in production */
  mkJob({ id: "R8001", cust: "Customer One 086 123 4567", wnd: 6, drs: 2, blk: 4, seq: 0,
          notes: [{ k: "comment", t: "rang on 0871234567, eircode D02X285, will collect Friday",
                    s: "Production" }] }),
  /* nothing to glaze: not fed at all */
  mkJob({ id: "R8002", cust: "Customer Two", wnd: 0, drs: 0, blk: 4, seq: 1 }),
  /* finished section - still fed, so the office can correct it */
  mkJob({ id: "R8003", cust: "Customer Three", wnd: 4, drs: 0, blk: 3, seq: 2 }),
  /* THE B20 SHAPE: Production says three windows, some other sheet says
     eleven. The slice must read three. */
  mkJob({ id: "R8004", cust: "Customer Four", wnd: 11, drs: 9, wndMain: 3, drsMain: 0,
          blk: 4, seq: 3 })
];
JOBS.blockNames = NAMES;

(async function main() {
  console.log("glazing station");

  /* ================= 1. the slice ================= */
  let slice = Z.glzSlice(JOBS, NAMES);
  assert.deepStrictEqual(slice.map(r => r.title), ["R8001", "R8003", "R8004"],
    "one row per job, in the office's own order");
  pass("the slice is one row per job, in the office's own order");

  const one = slice[0];
  assert.strictEqual(one.wnd, 6);
  assert.strictEqual(one.drs, 2);
  assert.strictEqual(one.total, 8, "Total is windows plus doors");
  assert.strictEqual(one.section, "In production");
  assert.strictEqual(slice[1].section, "Finished", "a finished job is still fed, with its section on it");
  assert.strictEqual(Z.glzFeederFields(one).Total, 8);
  pass("Total is the job's windows plus its doors, and every section is fed");

  assert.ok(!slice.some(r => r.job === "R8002"), "a job with nothing to glaze is not fed at all");
  assert.strictEqual(Z.glzSlice([mkJob({ id: "R8100", wnd: 0, drs: 0 })], NAMES).length, 0);
  pass("a job whose Total would be 0 never reaches the list");

  /* THE PRODUCTION SHEET AND NOTHING ELSE (owner's standing rule, 2026-09-18) */
  const four = slice.find(r => r.job === "R8004");
  assert.strictEqual(four.wnd, 3, "Windows is the Production sheet's own number, not the merged one");
  assert.strictEqual(four.drs, 0);
  assert.strictEqual(four.total, 3, "so Total is 3, not the 20 the cross-sheet numbers would give");
  assert.strictEqual(Z.glzWindowsOf({ wnd: 11, wndMain: 3 }), 3);
  assert.strictEqual(Z.glzDoorsOf({ drs: 9, drsMain: 0 }), 0);
  assert.strictEqual(Z.glzWindowsOf({ wnd: 11 }), 0,
    "a job with no Production quantity at all reads nought - it is not this station's business");
  pass("the quantities are read from `Production` alone: the cross-sheet pair is never consulted");

  /* the green set, and nothing beside it */
  const fed = Z.glzFeederFields(one);
  assert.deepStrictEqual(Object.keys(fed).sort(), Z.GLZ_FEEDER_FIELDS.slice().sort(),
    "the feeder writes exactly its own nine columns");
  ["Eircode", "Phone", "Area", "Colour", "Sold", "Office", "Glazed", "GlazedBy", "DoneBy"]
    .forEach(bad => assert.ok(!(bad in fed), "and never a " + bad));
  Z.GLZ_FLOOR_FIELDS.forEach(k => assert.ok(Z.GLZ_FEEDER_FIELDS.indexOf(k) < 0,
    "the feeder's columns and the floor's do not overlap: " + k));
  pass("the feeder's row is the nine job-fact columns, and never one of the floor's five");

  /* ================= 2. rule 3: the strip ================= */
  assert.strictEqual(one.customer, "Customer One …", "a phone number typed into a customer name goes");
  assert.strictEqual(one.comment, "rang on …, eircode …, will collect Friday",
    "a ten-digit phone number and an eircode-shaped token both go");
  assert.strictEqual(Z.glzStrip("0871234567"), "…");
  assert.strictEqual(Z.glzStrip("086 123 4567"), "…", "a phone number as people type one");
  assert.strictEqual(Z.glzStrip("+353 86 123 4567"), "…");
  /* the bracketed shape keeps its opening bracket - the strip's phone pattern
     starts at a digit - and loses every digit, which is the whole of rule 3 */
  assert.strictEqual(Z.glzStrip("(086) 123 4567"), "(…");
  assert.strictEqual(Z.glzStrip("D02 X285"), "…", "an eircode with its space");
  assert.strictEqual(Z.glzStrip("A65F4E2"), "…", "and without it");
  assert.strictEqual(Z.glzStrip("12345"), "12345", "five digits stay - they are a quantity");
  assert.strictEqual(Z.glzStrip("R5303"), "R5303", "a job number is not an eircode and stays");
  assert.strictEqual(Z.glzStrip("x".repeat(400)).length, Z.GLZ_COMMENT_MAX, "capped at 140");
  assert.strictEqual(Z.glzStrip(null), "");
  const sliceJson = JSON.stringify(slice);
  ["0871234567", "086 123 4567", "D02X285", "Cork", "@"].forEach(bad =>
    assert.strictEqual(sliceJson.indexOf(bad), -1, "the slice leaked " + bad));
  pass("rule 3: a phone number in three shapes and an eircode never reach a feed body");

  /* ================= 3. there is no seed ================= */
  assert.deepStrictEqual(Z.GLZ_SEED_FIELDS, [], "there is no office record of glazing to seed from");
  assert.deepStrictEqual(Z.glzSeedFields(one), {});
  assert.deepStrictEqual(Z.GLZ_FEEDER_WRITES.filter(k => Z.GLZ_FLOOR_FIELDS.indexOf(k) >= 0), [],
    "so the feeder's whole vocabulary holds not one floor column");
  pass("an untouched row starts at nought: the feeder has no seed and cannot express one");

  /* ================= 4. feedPlan with the glazing definition ================= */
  let plan = ST.feedPlan(slice, [], { at: "2026-09-21T09:00:00.000Z", by: "the office", def: Z.GLAZE });
  assert.strictEqual(plan.adds.length, 3);
  assert.strictEqual(plan.patches.length, 0);
  const add = plan.adds[0];
  assert.strictEqual(add.Title, "R8001");
  assert.strictEqual(add.Total, 8);
  assert.strictEqual(add.FedBy, "the office");
  Z.GLZ_FLOOR_FIELDS.forEach(k => assert.ok(!(k in add), "a new row carries no " + k));
  pass("feedPlan on the glazing definition adds a row per job and names no floor column");

  /* a row the floor has worked on: its facts may change, its counter may not */
  const live = [item({ Title: "R8001", Job: "R8001", Customer: "Old name", Section: "In production",
                       Seq: 0, Active: "Yes", Windows: 6, Doors: 2, Total: 8, Comment: "",
                       Glazed: 5, GlazedBy: "the glazer", GlazedAt: "2026-09-21T08:00:00.000Z",
                       DoneBy: "the glazer", DoneAt: "2026-09-21T08:00:00.000Z" }, "900")];
  plan = ST.feedPlan(slice, live, { at: "2026-09-21T09:00:00.000Z", by: "the office", def: Z.GLAZE });
  const p1 = plan.patches.find(p => p.id === "900");
  assert.ok(p1, "the changed customer is patched");
  assert.strictEqual(p1.fields.Customer, "Customer One …");
  Z.GLZ_FLOOR_FIELDS.forEach(k => assert.ok(!(k in p1.fields),
    "the feeder wrote a floor column on a row the floor has tapped: " + k));
  assert.strictEqual(live[0].fields.Glazed, 5, "the floor's own number is exactly where they left it");
  pass("the feeder patches job facts and can never touch the counter, By, At or last touch");

  /* a job that has left the sheet is deactivated, never deleted */
  plan = ST.feedPlan(slice.filter(r => r.job !== "R8001"), live,
                     { at: "2026-09-21T09:00:00.000Z", by: "the office", def: Z.GLAZE });
  const off = plan.patches.find(p => p.id === "900");
  assert.strictEqual(off.fields.Active, "No");
  assert.deepStrictEqual(Object.keys(off.fields).sort(), ["Active", "FedAt", "FedBy"],
    "and nothing else at all");
  pass("a job that has left the section is marked Active = No, and no plan can hold a delete");

  /* the hash notices a real change and ignores a rerun */
  const h1 = ST.sliceHash(slice, Z.GLAZE);
  assert.strictEqual(ST.sliceHash(Z.glzSlice(JOBS, NAMES), Z.GLAZE), h1, "an unchanged slice hashes the same");
  const moved = Z.glzSlice(JOBS.map(j => j.id === "R8001" ? mkJob(Object.assign({}, j, { wnd: 7, wndMain: 7 })) : j), NAMES);
  assert.notStrictEqual(ST.sliceHash(moved, Z.GLAZE), h1, "a changed quantity does not");
  pass("the slice hash moves when a quantity does, so a change reaches the floor on the next load");

  /* ================= 5. the board, and the three colour levels ================= */
  const rows = [
    item({ Title: "R8001", Job: "R8001", Customer: "Customer One", Section: "In production",
           Seq: 0, Active: "Yes", Windows: 6, Doors: 2, Total: 8, Comment: "mind the sill",
           Glazed: 0 }, "500"),
    item({ Title: "R8003", Job: "R8003", Customer: "Customer Three", Section: "Finished",
           Seq: 2, Active: "Yes", Windows: 4, Doors: 0, Total: 4, Glazed: 4,
           GlazedBy: "the glazer", GlazedAt: "2026-09-21T10:00:00.000Z",
           DoneBy: "the glazer", DoneAt: "2026-09-21T10:00:00.000Z" }, "503"),
    item({ Title: "R8004", Job: "R8004", Customer: "Customer Four", Section: "In production",
           Seq: 3, Active: "Yes", Windows: 3, Doors: 0, Total: 3, Glazed: 1 }, "504"),
    item({ Title: "R8009", Job: "R8009", Customer: "Gone", Section: "In production",
           Seq: 9, Active: "No", Windows: 2, Doors: 0, Total: 2, Glazed: 0 }, "509")
  ];
  const board = Z.glzBoard(rows);
  assert.deepStrictEqual(board.map(c => c.job), ["R8001", "R8004"],
    "the tablet sees In production and Active only");
  assert.deepStrictEqual(Z.glzOfficeBoard(rows).map(c => c.job), ["R8001", "R8003", "R8004"],
    "and the office sees every section, finished jobs included");
  assert.strictEqual(Z.glzColour(0, 8), "", "nothing glazed: no colour");
  assert.strictEqual(Z.glzColour(1, 3), "yellow", "started");
  assert.strictEqual(Z.glzColour(4, 4), "green", "every unit glazed");
  assert.strictEqual(Z.glzColour(0, 0), "", "a job with nothing to glaze has finished nothing");
  assert.strictEqual(board[0].colour, "");
  assert.strictEqual(board[1].colour, "yellow");
  assert.strictEqual(Z.glzOfficeBoard(rows)[1].colour, "green");
  assert.strictEqual(Z.glzOfficeBoard(rows)[1].finished, true);
  assert.strictEqual(board[0].finished, false);
  pass("the board narrows by Active and section, and the colour is none / yellow / green");

  /* a counter above its own quantity is clamped for display and nothing else */
  const shrunk = Z.glzRecord(item({ Title: "R8050", Job: "R8050", Total: 4, Glazed: 6 }, "550"));
  assert.strictEqual(shrunk.glazed, 4, "a shortened job reads 4 of 4, never 150%");
  assert.strictEqual(shrunk.left, 0);
  pass("a counter above the quantity is clamped on screen, and nothing is written back for it");

  /* ================= 6. "N left", the filter and the words ================= */
  assert.strictEqual(Z.glzLeft(board), 8 + 2, "left is the board's own, summed");
  assert.strictEqual(Z.glzLeftWords(10), "10 left");
  assert.strictEqual(Z.glzUnitWords(8), "8 units");
  assert.strictEqual(Z.glzUnitWords(1), "1 unit");
  assert.strictEqual(Z.glzQtyWords(board[0]), "6 windows · 2 doors");
  assert.strictEqual(Z.glzQtyWords({ wnd: 1, drs: 0 }), "1 window");
  assert.deepStrictEqual(Z.glzFilter(board, "R8004").map(c => c.job), ["R8004"]);
  assert.deepStrictEqual(Z.glzFilter(board, "customer one").map(c => c.job), ["R8001"]);
  assert.strictEqual(Z.glzFilter(board, "").length, 2, "an empty box is every card");
  pass("the header's number is the board's, and the box narrows the cards and never the number");

  /* one job's card for the drawer: active rows only */
  assert.strictEqual(Z.glzJobCard(rows, "R8001").total, 8);
  assert.strictEqual(Z.glzJobCard(rows, "R8009"), null,
    "a row that has left the sheet cannot make the drawer read a total no board agrees with");
  assert.strictEqual(Z.glzJobCard(rows, "nosuchjob"), null);
  pass("the drawer's card is an active row or nothing at all");

  /* ================= 7. the tap ================= */
  const rec = Z.glzRecord(rows[0]);
  assert.strictEqual(Z.glzApplyTap(rec, 1), 1);
  assert.strictEqual(Z.glzApplyTap(rec, -1), 0, "and never below nought");
  assert.strictEqual(Z.glzApplyTap(rec, "all"), 8);
  assert.strictEqual(Z.glzApplyTap(rec, "none"), 0);
  assert.strictEqual(Z.glzApplyTap({ total: 8, glazed: 8 }, 1), 8, "and never above the total");
  assert.strictEqual(Z.glzApplyTap({ total: 8, glazed: 3 }, "nonsense"), 3);
  pass("a tap is clamped to 0…Total, in both directions");

  const body = Z.glzFloorOnly(Z.glzTapFields(5, "the glazer", "2026-09-21T11:00:00.000Z"));
  assert.deepStrictEqual(Object.keys(body).sort(),
    ["DoneAt", "DoneBy", "Glazed", "GlazedAt", "GlazedBy"],
    "one tap writes exactly the counter, its By/At and the last-touch pair");
  assert.strictEqual(body.Glazed, 5);
  assert.strictEqual(body.GlazedBy, "the glazer");
  assert.strictEqual(body.DoneAt, "2026-09-21T11:00:00.000Z");
  /* the filter is the proof, not the builder: whatever is handed in, only five
     columns can come out, and a counter that is not a number is dropped */
  const smuggled = Z.glzFloorOnly({ Glazed: 5, GlazedBy: "x", GlazedAt: "y", DoneBy: "z", DoneAt: "w",
                                    Total: 999, Customer: "someone", Active: "No", Section: "Finished",
                                    Title: "R9999" });
  assert.deepStrictEqual(Object.keys(smuggled).sort(),
    ["DoneAt", "DoneBy", "Glazed", "GlazedAt", "GlazedBy"],
    "a job fact cannot get onto the wire through the filter");
  assert.deepStrictEqual(Z.glzFloorOnly({ Glazed: "8" }), {}, "a counter that is not a number is dropped");
  pass("the tap body is five fields, and the filter refuses everything else");

  /* ================= 8. offline queue replay ================= */
  /* The tablet keeps what it owes in localStorage, so a tablet closed mid-write
     still owes it after a reload. What matters is that the REPLAY carries the
     moment of the tap rather than the moment of the send, and that nothing
     anybody types into that storage can widen the write. */
  const owed = { id: "500", value: 5, who: "the glazer", at: "2026-09-21T11:00:00.000Z",
                 job: "R8001", title: "R8001", from: 0,
                 /* somebody with the tablet edits their own localStorage */
                 Total: 999, Customer: "someone else", Active: "No" };
  const afterReload = JSON.parse(JSON.stringify(owed));
  const replayed = Z.glzFloorOnly(Z.glzTapFields(afterReload.value, afterReload.who, afterReload.at));
  assert.deepStrictEqual(replayed, body, "a reload replays exactly the write the tap made");
  assert.strictEqual(replayed.GlazedAt, "2026-09-21T11:00:00.000Z",
    "with the time of the TAP on it, not the time of the send");
  assert.ok(!("Total" in replayed) && !("Customer" in replayed) && !("Active" in replayed),
    "and nothing edited into that storage can reach the list");
  pass("a queued tap survives a reload as the same five-field write, stamped when it was made");

  /* ================= 9. the rebase, and the office's later word ================= */
  const tap = { value: 5, from: 0, at: "2026-09-21T11:00:00.000Z" };
  assert.deepStrictEqual(Z.glzRebase(tap, { Glazed: 0, Total: 8 }), { action: "keep" },
    "nothing moved underneath: the tap stands");
  assert.deepStrictEqual(Z.glzRebase(tap, { Glazed: 2, Total: 8 }),
    { action: "rebase", value: 7, from: 2 },
    "the counter rose under it: the same movement, re-based");
  assert.deepStrictEqual(Z.glzRebase({ value: 5, from: 6, at: "2026-09-21T11:00:00.000Z" },
    { Glazed: 0, Total: 8, DoneAt: "2026-09-21T11:30:00.000Z" }), { action: "drop" },
    "it fell and the row was stamped after the tap: somebody said something later");
  assert.deepStrictEqual(Z.glzRebase({ value: 5, from: 6, at: "2026-09-21T11:00:00.000Z" },
    { Glazed: 0, Total: 8, DoneAt: "2026-09-21T10:00:00.000Z" }), { action: "keep" },
    "a fall with no later stamp leaves the floor's own statement alone");
  assert.deepStrictEqual(Z.glzRebase(tap, { Glazed: 7, Total: 8 }).value, 8,
    "and a re-base is clamped to the total like everything else");
  pass("a queued tap is re-based, dropped or kept on the stamps, exactly as the other two tablets do");

  /* ================= 10. the log line the TABLET writes ================= */
  const lf = ST.logFields(Z.glzLogEntry({ job: "r8001", from: 3, to: 6,
    who: "the glazer", at: "2026-09-21T12:00:00.000Z" }));
  assert.deepStrictEqual(lf, { Title: "R8001", Station: "Glazing", GlassType: "GLAZING",
    Stage: "glaze", From: 3, To: 6, Who: "the glazer", At: "2026-09-21T12:00:00.000Z" },
    "the shared log list's own eight columns, with Stage = glaze");
  const lrows = ST.logRows([item(lf, "1"),
    item({ Title: "R8001", Station: "Glass", GlassType: "GLASS", Stage: "cut", From: 0, To: 4,
           Who: "Person B", At: "2026-09-21T12:05:00.000Z" }, "2")], "Glazing");
  assert.strictEqual(lrows.length, 1, "and the glazing board reads its own station's lines only");
  assert.strictEqual(lrows[0].stage, "glaze");
  pass("a glazing log line is the shared list's eight columns with Stage = glaze, read back by station");

  /* ================= 11. people, and one stage ================= */
  const people = ST.stationPeople([
    item({ Title: "Person A", Station: "Glazing", Stages: "glaze", PIN: "", Active: "Yes" }, "1"),
    item({ Title: "Person B", Station: "Glass", Stages: "cut", PIN: "", Active: "Yes" }, "2"),
    item({ Title: "Person C", Station: "Glazing", Stages: "cut", PIN: "1234", Active: "Yes" }, "3"),
    item({ Title: "Person D", Station: "Glazing", Stages: "glaze", PIN: "", Active: "No" }, "4")
  ], "Glazing", Z.GLAZE.stages);
  assert.deepStrictEqual(people.map(p => p.name), ["Person A", "Person C"],
    "this station's people only, and only the active ones");
  assert.deepStrictEqual(people[0].stages, ["glaze"]);
  assert.deepStrictEqual(people[1].stages, [],
    "a stage this station does not have is not a stage they hold");
  assert.strictEqual(ST.canStage(people[0], "glaze"), true);
  assert.strictEqual(ST.canStage(people[1], "glaze"), false, "so they may move nothing");
  assert.strictEqual(ST.canStage(null, "glaze"), false, "and nobody signed in may move nothing too");
  pass("a person without `glaze` holds no stage here and cannot tap");

  /* ================= 12. the office's edit, on the real app.js ================= */
  consent(true);
  ZITEMS = rows.map(r => item(Object.assign({}, r.fields), r.id));
  A("GLZ_ITEMS = null; GLZ_OK = null; GLZ_WHY = ''; GLZ_SITEID = null; GRECS = null; GRECS_OF = false;");
  A("GLZ_FEED = { hash: '', at: 0 };");
  reset();
  const read = await A("readGlazing()");
  assert.ok(read, "the office reads the glazing list");
  assert.strictEqual(A("GLZ_OK"), true);
  assert.ok(REQ.every(r => r.path.indexOf("/drive") < 0 && r.path.indexOf("/workbook") < 0),
    "and not one request went anywhere near the workbook");
  const listCalls = REQ.filter(r => r.path.indexOf("/sites/") === 0 && r.path.indexOf("/lists") > 0);
  assert.ok(listCalls.length > 0, "there were list calls to check");
  assert.ok(listCalls.every(r => r.path.indexOf(FLISTS_PATH) === 0),
    "every list call names the Floor stations site: " +
    JSON.stringify(listCalls.map(r => r.path).filter(p => p.indexOf(FLISTS_PATH) !== 0)));
  pass("the office reads the Glazing station list out of the Floor stations site, and no workbook path");

  /* the Dashboard Log line an office action leaves is a WORKBOOK write - the
     `Dashboard Log` sheet is one of the dashboard's own (rule 2), and it is the
     same appendLog every other office action uses. It is stubbed here so the
     line can be asserted on directly. */
  const LOGGED = [];
  CW.appendLog = async (...a) => { LOGGED.push(a); };
  A("CHANGES = [];");
  reset();
  const edited = await A('glzOfficeEdit("500", "all")');
  assert.strictEqual(edited, true);
  const patches = writes().filter(w => w.method === "PATCH");
  assert.strictEqual(patches.length, 1, "exactly one PATCH went out");
  assert.deepStrictEqual(Object.keys(patches[0].body).sort(),
    ["DoneAt", "DoneBy", "Glazed", "GlazedAt", "GlazedBy"],
    "the office's edit writes exactly the counter, its By/At and the last-touch pair");
  assert.strictEqual(patches[0].body.Glazed, 8);
  assert.ok(patches[0].body.DoneBy && patches[0].body.DoneBy.length, "with the office person's name on it");
  assert.strictEqual(writes().filter(w => w.path.indexOf(LOG_ID) >= 0).length, 0,
    "RULE 2: the office writes no line of Station log, on this station or any other");
  assert.strictEqual(writes().filter(w => w.method === "DELETE").length, 0);
  assert.ok(writes().every(w => w.path.indexOf("/workbook") < 0 && w.path.indexOf("/drive") < 0),
    "and the office's glazing edit never touches the workbook either");
  pass("an office edit is one PATCH of five fields, no Station log line and no workbook write");

  const logged = A("CHANGES[0]");
  assert.ok(logged, "the change is in the office's own Changes list");
  assert.strictEqual(logged.job, "R8001");
  assert.strictEqual(logged.what, "Glazing: R8001");
  assert.strictEqual(logged.from, "0");
  assert.strictEqual(logged.to, "8");
  assert.strictEqual(LOGGED.length, 1, "exactly one line, in the Dashboard Log sheet");
  assert.deepStrictEqual(LOGGED[0].slice(1), ["R8001", "Glazing: R8001", "0", "8"]);
  pass("one Dashboard Log line per office change, naming the job and what moved");

  /* the board reads the new number at once, without waiting for a poll */
  const afterEdit = A('glzRecordsNow().byId["500"]');
  assert.strictEqual(afterEdit.glazed, 8);
  assert.strictEqual(afterEdit.colour, "green");
  assert.strictEqual(afterEdit.finished, true);
  pass("the office's own copy of the list is in step at once, so the board never reads a stale number");

  /* TWO CLICKS IN ONE TICK ARE ONE WRITE. The flag goes up before the first
     await, which is the whole of the welding board's review finding R3: set
     after the consent check, both clicks got past the guard, both read the same
     row and both PATCHed from the same base. */
  A("CHANGES = [];");
  reset();
  const two = await Promise.all([A('glzOfficeEdit("504", 1)'), A('glzOfficeEdit("504", 1)')]);
  assert.deepStrictEqual(two.sort(), [false, true], "the second click is refused while the first is in flight");
  assert.strictEqual(writes().filter(w => w.method === "PATCH").length, 1,
    "double-click on one row is one PATCH, not two");
  assert.strictEqual(A("CHANGES.length"), 1, "and one Dashboard Log line");
  pass("a double-click writes once: the busy flag goes up before the first await");

  /* a click that would change nothing writes nothing */
  reset();
  assert.strictEqual(await A('glzOfficeEdit("500", "all")'), false);
  assert.strictEqual(writes().length, 0, "All on a job that is already all is not a write");
  assert.strictEqual(await A('glzOfficeEdit("nosuchrow", 1)'), false);
  assert.strictEqual(writes().length, 0);
  pass("a click that could not move the number makes no request at all");

  /* THE BOARD IS UP TO TEN SECONDS OLD. A glazer taps All at 14:00:01; this
     screen last polled at 13:59:56 and is drawing 1 of 3; the office presses +
     at 14:00:05. Derived from the BOARD that writes 2 and destroys their work,
     with the office's later stamp on it. So the row is read immediately before
     the PATCH and the number comes from what it says. */
  const four504 = ZITEMS.find(x => x.id === "504");
  four504.fields.Glazed = 3;
  four504.fields.DoneBy = "the glazer";
  four504.fields.DoneAt = "2026-09-21T14:00:01.000Z";
  A("CHANGES = [];");
  reset();
  const bumped = await A('glzOfficeEdit("504", 1)');
  assert.ok(REQ.some(r => r.method === "GET" && /\/items\/504/.test(r.path)),
    "the row was read immediately before anything was written");
  assert.strictEqual(bumped, false, "+ on a row the floor has just finished writes nothing at all");
  assert.strictEqual(writes().length, 0, "not one PATCH went out");
  assert.strictEqual(four504.fields.Glazed, 3, "so the floor's 3 is exactly where they left it");
  assert.strictEqual(A("CHANGES.length"), 0, "and there is no Dashboard Log line about a non-event");
  assert.ok(TOASTS.some(t => /updated from the floor first/.test(t.m)),
    "the office is told the row moved under them: " + JSON.stringify(TOASTS.map(t => t.m)));
  pass("an office edit re-reads the row first, so a floor tap it had not seen is never overwritten");

  /* A FAILED WRITE RESTORES THE NUMBER. Nothing local is changed before the
     PATCH lands, so what goes back on screen is what the list still says. */
  four504.fields.Glazed = 1;
  four504.fields.DoneAt = "2026-09-21T14:00:01.000Z";
  A("GLZ_ITEMS = null; GLZ_OK = null; GRECS = null; GRECS_OF = false;");
  await A("readGlazing()");
  assert.strictEqual(A('glzRecordsNow().byId["504"].glazed'), 1);
  A("CHANGES = [];");
  reset();
  FAIL_ONCE = 1;
  assert.strictEqual(await A('glzOfficeEdit("504", 1)'), false, "the refused write is reported, not swallowed");
  FAIL_ONCE = 0;
  assert.strictEqual(A('glzRecordsNow().byId["504"].glazed'), 1,
    "and the board still reads the number the list really holds");
  assert.strictEqual(four504.fields.Glazed, 1);
  assert.strictEqual(A("CHANGES.length"), 0, "no Dashboard Log line for a write that never landed");
  assert.ok(TOASTS.some(t => t.err && /still reads 1/.test(t.m)),
    "and the office is told what it still reads: " + JSON.stringify(TOASTS.map(t => t.m)));
  assert.ok(!A("glzWriting['504']"), "the line is free to be clicked again");
  pass("a refused office write leaves the number where it was and says so");

  /* the board's own drawing: the card head opens the drawer only for a job that
     is still on the sheet, and every stepper names the row it belongs to */
  global.__jobs = JOBS;
  A("ALL = __jobs; BLOCKNAMES = " + JSON.stringify(NAMES) + "; ALL.blockNames = BLOCKNAMES;");
  A("GLZ_Q = ''; GLZ_SECT = ''; state.board = 'glazing';");
  const html = A("glzBoardHtml()");
  assert.ok(/data-zopen="R8001"/.test(html), "a job still on the sheet opens its drawer from the head");
  assert.ok(!/data-zopen="R8050"/.test(html));
  assert.ok(/data-zact="all"/.test(html) && /data-zact="none"/.test(html) &&
            /data-zact="1"/.test(html) && /data-zact="-1"/.test(html),
    "and every row carries − + All None");
  assert.ok(/Floor log/.test(html), "with the floor's own log under the board");
  assert.ok(html.indexOf("086 123 4567") < 0 && html.indexOf("D02X285") < 0,
    "and nothing rule 3 forbids anywhere on it");
  pass("the office board draws the steppers, the way into the drawer and the floor log");

  /* the drawer's read-only Glazing line */
  const line = A('glzDrawerLine(ALL.find(j => j.id === "R8001"))');
  assert.ok(/Glazing/.test(line) && /8 \/ 8/.test(line), "n of Total: " + line);
  assert.ok(/glzopen/.test(line), "with the way to the board");
  assert.ok(!/data-cp|cpline|cpgrp/.test(line), "and nothing that could tick a checkpoint");
  assert.strictEqual(A('glzDrawerLine(ALL.find(j => j.id === "R8002"))'), "",
    "a job with nothing to glaze gets no line at all");
  pass("the job drawer gains a read-only Glazing line, with no control that writes anything");

  /* ================= 13. the feeder, on the real app.js ================= */
  ZITEMS = []; LOGITEMS = [];
  A("state.board = null;");
  A("GLZ_ITEMS = null; GLZ_OK = null; GLZ_FEED = { hash: '', at: 0 }; GRECS = null; GRECS_OF = false;");
  reset();
  let r = await A("feedGlazing()");
  assert.ok(r && r.sent === 3, "three rows fed: R8001, R8003 and R8004");
  assert.deepStrictEqual(ZITEMS.map(x => x.fields.Title).sort(), ["R8001", "R8003", "R8004"]);
  assert.ok(writes().every(w => w.method === "POST"), "three new rows: three POSTs and nothing else");
  assert.strictEqual(writes().filter(w => w.method === "DELETE").length, 0, "the feeder never deletes");
  assert.ok(REQ.every(w => w.path.indexOf("/workbook") < 0 && w.path.indexOf("/drive") < 0),
    "and the feeder never touches the workbook");
  assert.strictEqual(ZITEMS.find(x => x.fields.Title === "R8004").fields.Total, 3,
    "R8004's Total is the Production sheet's 3, not the cross-sheet 20");
  assert.ok(!ZITEMS.some(x => x.fields.Title === "R8002"), "and a job with nothing to glaze is not there");
  const feedJson = JSON.stringify(ZITEMS);
  ["0871234567", "086 123 4567", "D02X285", "Cork", "@"].forEach(bad =>
    assert.strictEqual(feedJson.indexOf(bad), -1, "the fed list leaked " + bad));
  ZITEMS.forEach(x => Z.GLZ_FLOOR_FIELDS.forEach(k =>
    assert.ok(!(k in x.fields), "a fed row carries a floor column: " + k)));
  pass("the feeder pushes the sheet's glazing into the list, never deletes, never touches the workbook");

  reset();
  assert.strictEqual(await A("feedGlazing()"), null, "an unchanged slice inside ten minutes is skipped");
  assert.strictEqual(writes().length, 0);
  pass("an unchanged slice is not re-fed, so a dashboard left open writes nothing all afternoon");

  /* the floor taps a row; the feeder must never write its counter */
  const fedRow = ZITEMS.find(x => x.fields.Title === "R8001");
  fedRow.fields.Glazed = 2;
  fedRow.fields.DoneAt = "2026-09-21T13:00:00.000Z";
  fedRow.fields.DoneBy = "the glazer";
  A("GLZ_FEED = { hash: '', at: 0 };");
  global.__jobs2 = JOBS.map(j => j.id === "R8001" ? mkJob(Object.assign({}, j, { cust: "Renamed" })) : j);
  __jobs2.blockNames = NAMES;
  A("ALL = __jobs2; ALL.blockNames = BLOCKNAMES;");
  reset();
  await A("feedGlazing()");
  writes().forEach(w => Z.GLZ_FLOOR_FIELDS.forEach(k =>
    assert.ok(!(k in (w.body || {})),
      "the feeder wrote a floor column on a row the floor has tapped: " + JSON.stringify(w.body))));
  assert.strictEqual(ZITEMS.find(x => x.fields.Title === "R8001").fields.Glazed, 2,
    "the floor's own number is exactly where they left it");
  assert.strictEqual(ZITEMS.find(x => x.fields.Title === "R8001").fields.Customer, "Renamed",
    "while the job fact that really changed did reach the floor");
  A("ALL = __jobs; ALL.blockNames = BLOCKNAMES;");
  pass("once the floor has tapped a row, the feeder still updates its facts and never its counter");

  /* A MISSING LIST IS A STATE, NOT A CRASH - and it never takes the glass or
     the welding feed with it. No list is created by code, ever. */
  FLISTS = FLISTS.filter(l => l.id !== GLZ_ID);
  CW._resetListIds();
  try { delete mem.cw_listids; } catch (e) {}
  A("GLZ_OK = null; GLZ_WHY = ''; GLZ_FEED = { hash: '', at: 0 };");
  reset();
  assert.strictEqual(await A("feedGlazing()"), null);
  assert.strictEqual(A("GLZ_OK"), false);
  assert.ok(/Glazing station/.test(A("GLZ_WHY")), "and it says which list, plainly");
  assert.ok(/Ask the manager/.test(A("GLZ_WHY")));
  assert.strictEqual(writes().filter(w => w.method === "POST" && /lists/.test(w.path) &&
    !/items/.test(w.path)).length, 0, "and no list is created by code");
  assert.ok(/not in the/.test(A("glzBoardHtml()")), "the board shows the same quiet explanation");
  /* the other two feeds are untouched by it */
  A("WELD_OK = null; WELD_WHY = ''; WELD_FEED = { hash: '', at: 0 };");
  reset();
  await A("feedWelding()");
  assert.notStrictEqual(A("WELD_OK"), null, "the welding feed ran regardless");
  assert.ok(!/Glazing/.test(String(A("WELD_WHY") || "")), "and it is not describing glazing's problem");
  assert.strictEqual(A("STATION_OK"), null, "the glass station's own state was never touched");
  pass("a missing Glazing station list is explained, writes nothing, and stops no other station");
  FLISTS.push({ id: GLZ_ID, displayName: "Glazing station" });
  CW._resetListIds();

  /* ================= 14. the phase bar hears the floor ================= */
  /* Nothing in this section may write anything, anywhere: the floor's phase is
     computed on every render out of lists already in memory and stored nowhere
     - no list row, no workbook cell, no `Dashboard phases` write. */
  A("GLZ_ITEMS = null; GLZ_OK = null; GRECS = null; GRECS_OF = false;");
  A("WELD_ITEMS = null; WELD_OK = null; WRECS = null; WRECS_OF = false;");
  A("PHASES_SET = {}; PENDING = {};");
  A("CP_ITEMS = []; if (typeof cpListRebuild === 'function') cpListRebuild();");
  assert.strictEqual(A("jobPhase(ALL.find(j => j.id === 'R8001'))"), 0,
    "the sheet says nothing about this job yet");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 0,
    "and with no floor list read, neither does the floor");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "sheet");
  pass("with nothing read, the phase is the sheet's and the floor says nothing at all");

  /* welding alone: In fabrication */
  A("WELD_OK = true; WRECS = null; WRECS_OF = false;");
  A("WELD_ITEMS = " + JSON.stringify([item({ Title: "R8001|CASEMENT WINDOWS", Job: "R8001",
      Group: "CASEMENT WINDOWS", GroupSeq: 0, Section: "In production", Active: "Yes", Seq: 0,
      Frames: 4, Sashes: 4, FramesDone: 2, SashesDone: 0 }, "700")]) + ";");
  reset();
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 3,
    "any welding recorded is at least In fabrication");
  assert.strictEqual(A("PHASES[effectivePhase(ALL.find(j => j.id === 'R8001'))]"), "In fabrication");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "welding");
  pass("welding progress alone moves the phase to In fabrication");

  /* glazing outranks it */
  A("GLZ_OK = true; GRECS = null; GRECS_OF = false;");
  A("GLZ_ITEMS = " + JSON.stringify([item({ Title: "R8001", Job: "R8001", Section: "In production",
      Active: "Yes", Seq: 0, Windows: 6, Doors: 2, Total: 8, Glazed: 1 }, "800")]) + ";");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 4,
    "any glazing recorded is at least In glazing");
  assert.strictEqual(A("PHASES[effectivePhase(ALL.find(j => j.id === 'R8001'))]"), "In glazing");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "glazing");
  pass("glazing outranks welding: a job marked in glazing has been through fabrication");

  /* glazing with no welding at all */
  A("WELD_ITEMS = []; WRECS = null; WRECS_OF = false;");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 4,
    "glazing moves the phase with no welding recorded anywhere");
  pass("glazing progress moves the phase even with nothing recorded in welding");

  /* tapped back to nought: the voice is withdrawn */
  A("GLZ_ITEMS[0].fields.Glazed = 0; GRECS = null; GRECS_OF = false;");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 0,
    "a counter tapped back to nought withdraws that voice");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "sheet");
  A("WELD_ITEMS = " + JSON.stringify([item({ Title: "R8001|CASEMENT WINDOWS", Job: "R8001",
      Group: "CASEMENT WINDOWS", GroupSeq: 0, Section: "In production", Active: "Yes", Seq: 0,
      Frames: 4, Sashes: 4, FramesDone: 2, SashesDone: 0 }, "700")]) + "; WRECS = null; WRECS_OF = false;");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 3,
    "and falls back to the next highest voice, not to nothing");
  pass("a counter tapped back to nought withdraws the floor's voice and the phase falls back");

  /* the sheet, and a person, always win when they are further on */
  A("GLZ_ITEMS[0].fields.Glazed = 8; GRECS = null; GRECS_OF = false;");
  A("ALL.find(j => j.id === 'R8003').done = 1;");
  A("GLZ_ITEMS.push(" + JSON.stringify(item({ Title: "R8003", Job: "R8003", Section: "Finished",
      Active: "Yes", Seq: 2, Windows: 4, Doors: 0, Total: 4, Glazed: 1 }, "803")) +
    "); GRECS = null; GRECS_OF = false;");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8003'))"), 6,
    "the sheet is further on: the floor cannot drag a finished job backwards");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8003'))"), "sheet");
  A("ALL.find(j => j.id === 'R8003').done = 0;");
  A("PHASES_SET = { R8001: { phase: 5, name: 'Quality check', who: 'the office', at: '2026-09-21T09:00:00.000Z' } };");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 5,
    "a hand-set phase further on wins over the floor's");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "hand");
  A("PHASES_SET = { R8001: { phase: 2, name: 'Cutting', who: 'the office', at: '2026-09-21T09:00:00.000Z' } };");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 4,
    "and a hand-set phase behind the floor's is not a reason to go backwards");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "glazing");
  A("PHASES_SET = {};");
  pass("the effective phase is the highest of sheet, hand-set and floor - the floor only ever raises");

  /* the pure function, on its own */
  assert.strictEqual(CP.floorPhaseOf(null), null);
  assert.strictEqual(CP.floorPhaseOf({ welded: 0, glazed: 0 }), null, "nothing recorded says nothing");
  assert.strictEqual(CP.floorPhaseOf({ welded: 1, glazed: 0 }), 3);
  assert.strictEqual(CP.floorPhaseOf({ welded: 0, glazed: 1 }), 4);
  assert.strictEqual(CP.floorPhaseOf({ welded: 9, glazed: 1 }), 4, "glazing outranks welding");
  pass("floorPhaseOf is pure: two counters in, a phase or nothing out");

  /* NOTHING WAS WRITTEN FOR ANY OF IT */
  assert.strictEqual(writes().length, 0,
    "the phase bar wrote something: " + JSON.stringify(writes().map(w => w.method + " " + w.path)));
  assert.strictEqual(REQ.length, 0, "and did not even read anything - it is all already in memory");
  pass("the whole phase section made not one request: the floor's phase is stored nowhere");

  /* the screen's word, and the export's */
  A("GLZ_ITEMS[0].fields.Glazed = 8; GRECS = null; GRECS_OF = false;");
  assert.strictEqual(A("statusWord(ALL.find(j => j.id === 'R8001'))"), "In glazing",
    "the badge on the row and at the head of the drawer is the floor's phase");
  assert.strictEqual(A("statusWord(ALL.find(j => j.id === 'R8001'))"),
                     A("PHASES[effectivePhase(ALL.find(j => j.id === 'R8001'))]"),
    "which is exactly what effectivePhase answers - one reading, not two");
  /* the jobs export has no phase column at all, so no export can disagree with
     the screen about one. The check is on the field list, which is the whole of
     what an export may ever carry. */
  assert.ok(A("EXPORT_FIELD_KEYS").every(k => ["phase", "status", "stage"].indexOf(k) < 0),
    "the jobs export carries no phase, status or stage column: " + A("EXPORT_FIELD_KEYS").join(","));
  pass("the screen's status word is effectivePhase, and no export carries a phase that could differ");

  /* ================= 15. the station report, with no glazing in export.js ==== */
  const P = { from: "2026-09-21", to: "2026-09-27" };
  const ZDATA = {
    board: Z.glzOfficeBoard([
      item({ Title: "R8001", Job: "R8001", Customer: "Customer One 086 123 4567",
             Section: "In production", Seq: 0, Active: "Yes", Windows: 6, Doors: 2, Total: 8,
             Glazed: 8, GlazedBy: "the glazer", GlazedAt: "2026-09-21T10:00:00.000Z",
             DoneBy: "the glazer", DoneAt: "2026-09-21T10:00:00.000Z" }, "1"),
      item({ Title: "R8004", Job: "R8004", Customer: "Customer Four", Section: "In production",
             Seq: 3, Active: "Yes", Windows: 3, Doors: 0, Total: 3, Glazed: 1 }, "2")]),
    log: ST.logRows([
      item({ Title: "R8001", Station: "Glazing", GlassType: "GLAZING", Stage: "glaze",
             From: 0, To: 8, Who: "the glazer", At: "2026-09-21T10:00:00.000Z" }, "10"),
      item({ Title: "R8004", Station: "Glazing", GlassType: "GLAZING", Stage: "glaze",
             From: 0, To: 1, Who: "the glazer", At: "2026-09-22T11:00:00.000Z" }, "11")], "Glazing"),
    notes: ST.commentRows([
      item({ Title: "R8001|1", Job: "R8001", Station: "Glazing", Who: "the glazer",
             Text: "two beads short, ring the supplier on 086 123 4567",
             At: "2026-09-21T13:00:00.000Z" }, "20")]),
    days: [], target: null,
    who: "the office", when: new Date("2026-09-28T09:00:00.000Z"), build: "20260921-0900",
    /* one definition of "day" for the whole report, and this machine's zone
       must not decide which day a stamp falls on */
    dayOf: at => String(at).slice(0, 10)
  };
  const rep = A("stationReport")(Z.GLAZE, "glaze", ZDATA, P);
  assert.deepStrictEqual(rep.map(s => s.name), ["Summary", "Days", "Jobs", "Activity", "Notes"],
    "Summary, Days, Jobs, Activity and Notes, with no glazing-specific branch in export.js");
  const S1 = rep[0];
  assert.ok(S1.head.some(h => h[0] === "Station" && h[1] === "Glazing"));
  assert.ok(S1.head.some(h => h[0] === "Stage" && h[1] === "Glazing"));
  assert.ok(S1.head.some(h => h[0] === "Contact details" && String(h[1]).indexOf("None") === 0));
  assert.strictEqual(S1.rows.length, 1, "one ISO week in the period");
  assert.strictEqual(S1.rows[0][2], 9, "units recorded: 8 and 1");
  assert.strictEqual(S1.rows[0][3], 2, "two jobs touched");
  assert.strictEqual(S1.rows[0][4], 1, "one of them complete");
  assert.ok(S1.columns.indexOf("Target") < 0,
    "and no target or day-sheet columns: this station's definition has no daySheets");
  const jobs = rep[2];
  assert.deepStrictEqual(jobs.columns.slice(0, 6),
    ["Job", "Customer", "Section", "Windows", "Doors", "Total"]);
  assert.strictEqual(jobs.rows.length, 2, "one row per job for glazing");
  assert.strictEqual(jobs.rows[0][jobs.columns.indexOf("Glazed")], 8);
  assert.strictEqual(jobs.rows[0][jobs.columns.indexOf("Left")], 0);
  assert.strictEqual(jobs.rows[0][jobs.columns.indexOf("Complete")], "Yes");
  assert.strictEqual(jobs.rows[1][jobs.columns.indexOf("Complete")], "No");
  const act = rep[3];
  assert.strictEqual(act.rows.length, 2, "this stage's lines in the period");
  assert.ok(act.columns.indexOf("Part") < 0, "and no Part column: one report stage, one log stage");
  const flat = JSON.stringify(rep);
  ["086 123 4567", "0861234567", "D02X285"].forEach(bad =>
    assert.strictEqual(flat.indexOf(bad), -1, "the report leaked " + bad));
  assert.ok(/two beads short/.test(flat), "while the note itself is still readable");
  pass("stationReport(GLAZE, 'glaze', …) makes five sheets, with rule 3's strip over every piece of free text");

  /* the window offers it because the definition exists, and for no other reason */
  const opts = A("stationReportOptions()");
  assert.ok(opts.some(o => o.value === "glazing|glaze" && o.label === "Glazing · Glazing"),
    "the export window offers Glazing · Glazing: " + JSON.stringify(opts.map(o => o.value)));
  assert.ok(opts.some(o => o.value === "glass|cut") && opts.some(o => o.value === "welding|weld"),
    "beside the two that were already there");
  const zd = A('stationReportData(stationReportPick("glazing|glaze"))');
  assert.deepStrictEqual(zd.days, [], "no day sheets are read for a station that has none");
  assert.strictEqual(zd.target, null);
  pass("the station appears in the report window by being defined, and reads no day-sheet list");

  /* ================= 16. the three review findings (G1-G3) =================
     Each of these fails on the build that went to review. They are here rather
     than folded into the sections above because each is about a JOIN between
     two parts that were separately right - which is the only kind of fault an
     offline suite of pure functions can miss. */

  /* ---- G1: a welding row moving repaints the plain job list ----------------
     Section E's "repaint through the existing quiet path" reached
     redrawGlazing() and not redrawWelding(), which returned bare off its own
     board. So a job whose phase is the WELDING list's kept its old badge until
     the glass list happened to move or somebody reloaded the workbook. */
  A("state.board = null; state.sel = null;");
  A("__quiet = 0; __origQuiet = quietRows; quietRows = function () { __quiet++; };");
  /* the rows on screen, and what they were saying when they were drawn */
  A("ROWS_DRAWN = [ALL.find(j => j.id === 'R8001')]; ROWS_CHIPS = chipsNow(); ROWS_STALE = false;");
  A("__quiet = 0;");
  A("redrawWelding();");
  assert.strictEqual(A("__quiet"), 0, "nothing moved, so nothing is repainted");
  /* the welders finish the job while the office is looking at the plain list */
  A("WELD_ITEMS[0].fields.FramesDone = 4; WELD_ITEMS[0].fields.SashesDone = 4;");
  A("WELD_ITEMS = WELD_ITEMS.slice(); WRECS = null; WRECS_OF = false;");
  A("redrawWelding();");
  assert.strictEqual(A("__quiet"), 0,
    "a counter moving is not a repaint on its own - the phase did not change");
  /* ... and now one that really does change the phase: glazing withdrawn, so
     the job falls back to welding's word and the badge has to follow */
  A("GLZ_ITEMS[0].fields.Glazed = 0; GLZ_ITEMS = GLZ_ITEMS.slice(); GRECS = null; GRECS_OF = false;");
  A("__quiet = 0;");
  A("redrawWelding();");
  assert.strictEqual(A("__quiet"), 1,
    "a welding poll off the board repaints the rows when a job's floor phase changed");
  assert.strictEqual(A("PHASES[effectivePhase(ALL.find(j => j.id === 'R8001'))]"), "In fabrication",
    "and the word it repaints to is welding's");
  /* the glazing side has always done it; both are asserted so neither can
     regress on its own. The spy does not render, so what the rows are
     "currently saying" has to be brought up to date by hand first - that is
     what the real quietRows() does through renderRows(). */
  A("ROWS_CHIPS = chipsNow();");
  A("GLZ_ITEMS[0].fields.Glazed = 8; GLZ_ITEMS = GLZ_ITEMS.slice(); GRECS = null; GRECS_OF = false;");
  A("__quiet = 0;");
  A("redrawGlazing();");
  assert.strictEqual(A("__quiet"), 1, "and so does a glazing poll off the board");
  A("quietRows = __origQuiet;");
  pass("G1: a floor list moving off its own board repaints the job list when a phase changed");

  /* ---- G2: the floor only speaks for a job in a live section ---------------
     A floor row is never deleted - it goes Active = No when its job leaves the
     sheet, and its counter stays where the floor put it. Section F was
     deliberately not built, so NOTHING clears a finished glazing row when the
     office moves the job on by hand. Without the gate the job reads "In
     glazing" for ever, and it is the reading that outranks the sheet. */
  const R8001 = () => A("ALL.find(j => j.id === 'R8001')");
  A("ALL.find(j => j.id === 'R8001').blk = 4;");            // In production
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 4,
    "in production with a full glazing row: In glazing");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "glazing");
  /* the office moves it on by hand. The floor's row is untouched - still
     Active = Yes, still Glazed = Total - and must stop being heard. */
  A("ALL.find(j => j.id === 'R8001').blk = 1;");            // Ready to fit
  assert.strictEqual(A("glzRecordsNow().byJob['R8001'].glazed"), 8,
    "the floor's row is exactly as it was: nothing was cleared, deleted or reset");
  assert.strictEqual(A("floorPhaseRec(ALL.find(j => j.id === 'R8001'))"), null,
    "but the floor has no opinion about a job that has left production");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"),
                     A("jobPhase(ALL.find(j => j.id === 'R8001'))"),
    "so the phase is the sheet's own reading, exactly as before this station existed");
  assert.strictEqual(A("phaseSource(ALL.find(j => j.id === 'R8001'))"), "sheet");
  assert.notStrictEqual(A("PHASES[effectivePhase(ALL.find(j => j.id === 'R8001'))]"), "In glazing");
  /* welding is gated by the same line, not only glazing */
  A("GLZ_ITEMS[0].fields.Glazed = 0; GLZ_ITEMS = GLZ_ITEMS.slice(); GRECS = null; GRECS_OF = false;");
  assert.strictEqual(A("floorPhaseRec(ALL.find(j => j.id === 'R8001'))"), null,
    "the welding record is gated by the same test, not only the glazing one");
  /* a hand-set phase still works on such a job: the floor is silent, not the
     whole pipeline */
  A("PHASES_SET = { R8001: { phase: 5, name: 'Quality check', who: 'the office', at: '2026-09-21T09:00:00.000Z' } };");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 5,
    "and a hand-set phase is still heard on a job the floor has fallen silent about");
  A("PHASES_SET = {};");
  /* a job gone from the sheet altogether is covered by the same one line */
  A("ALL.find(j => j.id === 'R8001').blk = 4; ALL.find(j => j.id === 'R8001').cat = 'past';");
  A("GLZ_ITEMS[0].fields.Glazed = 8; GLZ_ITEMS = GLZ_ITEMS.slice(); GRECS = null; GRECS_OF = false;");
  assert.strictEqual(A("floorPhaseRec(ALL.find(j => j.id === 'R8001'))"), null,
    "a job that has left the sheet is covered by the same guard, not a second one");
  A("ALL.find(j => j.id === 'R8001').cat = 'active';");
  assert.strictEqual(A("effectivePhase(ALL.find(j => j.id === 'R8001'))"), 4,
    "and putting it back in production gives the floor its voice again");
  pass("G2: the floor speaks only for a job still in a live section, and falls silent rather than clearing anything");

  /* ---- G3: stationAfterFeed cannot skip the colour writer -----------------
     It sits in one promise chain as `.then(stationAfterFeed, () => {})` and the
     link AFTER it carries an onRejected of its own - so a throw in here is not
     a lost repaint, it is the rejection being swallowed by the next link's
     handler and glassColourRun() silently skipped for the whole load. The two
     links below are load()'s own, with a marker where the colour writer is. */
  A("__origSRIN = stationReadIfNeeded;");
  A("stationReadIfNeeded = function () { throw new Error('boom in the station read'); };");
  A("__glassRan = 0;");
  await A("Promise.resolve().then(stationAfterFeed, () => {}).then(() => { __glassRan++; }, () => {})");
  assert.strictEqual(A("__glassRan"), 1,
    "a throw inside stationAfterFeed must not stop the load reaching glassColourRun");
  A("stationReadIfNeeded = __origSRIN;");
  /* and the guard is not swallowing a real failure in silence */
  A("__warned = 0; __origWarn = console.warn; console.warn = function (m) { if (/after-feed/.test(String(m))) __warned++; };");
  A("stationReadIfNeeded = function () { throw new Error('boom'); };");
  A("stationAfterFeed();");
  assert.strictEqual(A("__warned"), 1, "and it says so in the console rather than going quiet");
  A("console.warn = __origWarn; stationReadIfNeeded = __origSRIN;");
  pass("G3: stationAfterFeed catches its own, so a load's glass colours are never skipped for a render");

  /* ================= 17. the gates ================= */
  const zsrc = src("glazing-core.js") + src("glazing.js") + src("glazing.html");
  ["setFill", "clearFill", "setValues", "appendLog", "saveProgress", "moveJobRow", "batchWrite",
   "/workbook", "downloadWorkbook", "parseWorkbook", "ExcelJS"].forEach(bad =>
    assert.strictEqual(zsrc.indexOf(bad), -1,
      "the glazing station files name `" + bad + "`, which they must never do"));
  pass("the three new station files contain no workbook word at all");

  /* and the same gate over the files this feature shares, unchanged */
  const shared = src("station-core.js") + src("station-ui.js") + src("welding-core.js") + src("welding.js");
  ["setFill", "clearFill", "setValues", "saveProgress", "moveJobRow", "/workbook"].forEach(bad =>
    assert.strictEqual(shared.indexOf(bad), -1, "a shared station file gained `" + bad + "`"));
  pass("and the station files it shares are still clean of one");

  /* no real name, address or domain anywhere in what this feature added */
  const mine = zsrc + src("test_glazing.js");
  assert.ok(!/@(?!example\.test)[a-z0-9-]+\.(com|ie|net|org|co\.uk)/i.test(mine),
    "an email address or a company domain reached the glazing files");
  pass("no real person, address or company domain in any of it");

  /* THE STANDING GATE, over every request this whole run made */
  assert.strictEqual(ALLREQ.filter(r => /\/workbook|\/drive|\/worksheets|\/range/.test(r.path)).length, 0,
    "something in this run asked for the workbook: " +
    JSON.stringify(ALLREQ.filter(r => /\/workbook|\/drive/.test(r.path)).map(r => r.path)));
  assert.strictEqual(ALLREQ.filter(r => r.method === "DELETE").length, 0,
    "something in this run deleted a list row");
  assert.ok(ALLREQ.length > 0, "there were requests to check");
  pass("over the whole run: not one workbook path and not one DELETE");

  console.log("\n" + n + " checks passed");
})().catch(e => { console.error("FAIL", e); process.exit(1); });
