/* Offline test of the welding station (docs/specs/2026-09-16-welding-station.md).

   The second floor station, and the first one built to a shape the third can be
   copied from. So the sharpest assertions here are not about welding at all -
   they are about the boundaries the owner drew, and they are checked on the
   request log and on the bodies that go out on it:

     · over the entire run, not one request touches the workbook - no download,
       no /workbook path, no range, no worksheet. The welding station never
       paints a cell, in either direction. The ONE workbook write anywhere in
       this feature is the `Dashboard Log` line an office edit leaves, which is
       a dashboard-owned sheet (rule 2) and is asserted on its own;
     · nothing the feeder sends ever carries one of the floor's own columns on a
       row the floor has tapped, and the feeder never issues a DELETE;
     · nothing the office's welding edit sends carries a job fact, a Section, an
       Active or anything but that part's counter, that part's By/At and the
       last-touch pair - and it never writes a line of `Station log`;
     · no phone number and no eircode can reach the list: the COMMENT is strip-
       ped on the way in and the strip is tested against both shapes;
     · T (transoms) and the four deny-listed product groups never reach the
       floor at all.

   Graph is a fake fetch() over one in-memory SharePoint site; nothing leaves
   the box, every address in here is example.test and every person is made up.
   Run: node test_welding.js                                                  */
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
                  "#lhost", "#dayhost", "#catmenu", "#movemenu", "#alertmenu", "#weldopen", "#wq", "#wsect"];
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
   One site: `Floor stations`, with the four lists the welding station uses.
   There is NO workbook route at all - a request for one is a test failure by
   construction, because there is nowhere for it to go. */
const G = "https://graph.microsoft.com/v1.0";
const SITE = "costellowindowsie.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";
const FSITE = "costellowindowsie.sharepoint.com,aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee,ffffffff-0000-1111-2222-333333333333";
const WELD_ID = "list-welding-station";
const PEOPLE_ID = "list-station-people";
const LOG_ID = "list-station-log";
const CMT_ID = "list-station-comments";
const FLISTS_PATH = "/sites/" + FSITE + "/lists";
const OWNLISTS_PATH = "/sites/" + SITE + "/lists";
const HOST_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/FloorStations";
const OWN_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/ProductionProgress?$select=id,displayName";

let SITE_EXISTS = true;
let FLISTS = [{ id: WELD_ID, displayName: "Welding station" },
              { id: PEOPLE_ID, displayName: "Station people" },
              { id: LOG_ID, displayName: "Station log" },
              { id: CMT_ID, displayName: "Station comments" }];
let WITEMS = [], PEOPLEITEMS = [], LOGITEMS = [], CMTITEMS = [];
let NEXTID = 100;
let FAIL_ONCE = 0;

const REQ = [], ALLREQ = [];
const ok = body => ({ status: 200, body: body });
const tick = ms => new Promise(r => setTimeout(r, ms == null ? 1 : ms));
function item(fields, id) {
  return { id: String(id == null ? NEXTID++ : id), fields: Object.assign({}, fields) };
}
const storeFor = id => id === WELD_ID ? WITEMS : id === PEOPLE_ID ? PEOPLEITEMS
                     : id === LOG_ID ? LOGITEMS : id === CMT_ID ? CMTITEMS : null;

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
  if (path === HOST_LOOKUP)
    return SITE_EXISTS ? ok({ id: FSITE, displayName: "Floor stations" })
                       : { status: 404, body: { error: { code: "itemNotFound" } } };
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
run("station-ui.js");
global.STU = window.STU;
run("app.js");

const TOASTS = [];
global.toast = (m, isErr) => TOASTS.push({ m: String(m), err: !!isErr });
const A = vm.runInThisContext.bind(vm);
const consent = yes => { CW.hasListConsent = async () => !!yes; };
const reset = () => { REQ.length = 0; TOASTS.length = 0; };
const writes = () => REQ.filter(r => r.method !== "GET");
const settle = ms => new Promise(r => setTimeout(r, ms == null ? 60 : ms));

let n = 0;
const pass = msg => { n++; console.log("  ok  " + msg); };

/* ---------- fixtures ----------
   Three jobs, exactly the three the spec names:
     R7001  two allowed groups (casement windows, pvc door), one denied
            (aluclad tilt & turn), a comment carrying a phone number and an
            eircode, and a "sent to floor" date
     R7002  one group with T only - nothing to weld here, ever
     R7003  in the Finished section, with a blank "sent to floor"          */
const NAMES = ["Can sell as second hand", "Ready to fit", "Collect & supply only",
               "Finished", "In production", "Not sent to floor"];
/* `prodsMain` is the `Production` sheet's own product counts, never merged with
   another sheet's (parser.js, 2026-09-17). A job that is only ever on
   `Production` - which is every fixture here unless it says otherwise - has the
   same numbers in both, so it defaults to whatever `prods` was given. A fixture
   that wants the two to DISAGREE, the way the live workbook did, passes both. */
const mkJob = o => {
  const j = Object.assign({
    id: "R0001", cust: "Customer One", area: "Cork", eir: "", off: "", colour: "", ph3: "",
    wnd: 0, drs: 0, glass: {}, prods: [], notes: [], sheets: ["Production"], src: {}, doors: [],
    dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
    cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0,
    cp: { win: "", drs: "", glass: {}, prod: {} }
  }, o || {});
  if (!(o && "prodsMain" in o)) j.prodsMain = j.prods;
  return j;
};
const prod = (name, f, s, t) => ({ n: name, f: f || 0, s: s || 0, t: t || 0, st: [] });

const JOBS = [
  mkJob({ id: "R7001", cust: "Customer One", wnd: 6, drs: 1, blk: 4, seq: 0,
          dates: { sold: null, stamp: null, ivana: null, ready: null, floor: "2026-09-12" },
          notes: [{ k: "comment", t: "rang on 0871234567, eircode D02X285, will collect Friday", s: "Production" }],
          prods: [prod("casement windows", 3, 8, 2), prod("aluclad tilt turn", 4, 4, 0),
                  prod("pvc door", 1, 1, 0)] }),
  mkJob({ id: "R7002", cust: "Customer Two", wnd: 2, drs: 0, blk: 4, seq: 1,
          prods: [prod("th w", 0, 0, 5)] }),
  mkJob({ id: "R7003", cust: "Customer Three", wnd: 4, drs: 0, blk: 3, seq: 2,
          prods: [prod("7000 casement", 4, 0, 0)] })
];
JOBS.blockNames = NAMES;

(async function main() {
  console.log("welding station");

  /* ================= 1. the slice ================= */
  let slice = W.weldSlice(JOBS, NAMES, () => "");
  const titles = slice.map(r => r.title);
  assert.deepStrictEqual(titles,
    ["R7001|CASEMENT WINDOWS", "R7001|PVC DOOR", "R7003|7000 CASEMENT"],
    "one row per job AND allowed group, in sheet order then group order");
  pass("the slice is one row per job and product group, in the office's own order");

  assert.ok(!titles.some(t => /ALUCLAD/.test(t)), "the deny-listed group is never fed");
  assert.strictEqual(W.weldAllowed("ALU CLAD WINDOWS"), false);
  assert.strictEqual(W.weldAllowed("aluclad tilt turn"), false, "matched after upper-casing and collapsing");
  assert.strictEqual(W.weldAllowed("Bifold"), false);
  assert.strictEqual(W.weldAllowed("composite"), false);
  assert.strictEqual(W.weldAllowed("CASEMENT WINDOWS"), true);
  assert.strictEqual(W.weldAllowed("PVC SMART"), true, "a group nobody has listed is allowed by default");
  pass("the four deny-listed groups never reach the floor, and a new green group needs no code change");

  assert.ok(!titles.some(t => /R7002/.test(t)), "a job whose only group is T is not fed at all");
  const cas = slice[0];
  assert.strictEqual(cas.frames, 3);
  assert.strictEqual(cas.sashes, 8);
  assert.strictEqual(cas.transoms, undefined, "there is no transom anywhere on a slice row");
  assert.ok(Object.keys(W.weldFeederFields(cas)).every(k => k !== "T" && !/transom/i.test(k)));
  assert.ok(JSON.stringify(slice).indexOf('"t"') < 0 || !slice.some(r => r.t),
    "and no row carries the sheet's T count under any name");
  pass("F and S only: T is never fed, and a group with nothing but T is not a row");

  assert.strictEqual(cas.job, "R7001");
  assert.strictEqual(cas.group, "CASEMENT WINDOWS");
  assert.strictEqual(cas.groupSeq, 0);
  assert.strictEqual(slice[1].groupSeq, 2, "the group's own column order on the sheet, gaps and all");
  assert.strictEqual(cas.customer, "Customer One");
  assert.strictEqual(cas.wnd, 6);
  assert.strictEqual(cas.drs, 1);
  assert.strictEqual(cas.sentToFloor, "2026-09-12");
  assert.strictEqual(cas.section, "In production");
  assert.strictEqual(slice[2].section, "Finished", "a finished job is still fed, with its section on it");
  assert.strictEqual(slice[2].sentToFloor, "", "a blank sent-to-floor stays blank");
  assert.strictEqual(W.weldTitle("r7001", "casement windows"), "R7001|CASEMENT WINDOWS");
  pass("the Title key is JOB|GROUP, and every job fact on the row is one the floor may see");

  /* the green set, and nothing beside it */
  const fed = W.weldFeederFields(cas);
  assert.deepStrictEqual(Object.keys(fed).sort(), W.WELD_FEEDER_FIELDS.slice().sort(),
    "the feeder writes exactly its own thirteen columns");
  ["Eircode", "Phone", "Area", "Colour", "Sold", "Stamp", "Office"].forEach(bad =>
    assert.ok(!(bad in fed), "and never a " + bad));
  pass("the feeder's row is the thirteen green columns and nothing else");

  /* ================= 2. rule 3: the comment strip ================= */
  assert.strictEqual(cas.comment, "rang on …, eircode …, will collect Friday",
    "a ten-digit phone number and an eircode-shaped token both go");
  assert.strictEqual(W.weldStripDigits("0871234567"), "…");
  assert.strictEqual(W.weldStripDigits("123456"), "…", "six digits go");
  assert.strictEqual(W.weldStripDigits("12345"), "12345", "five digits stay - they are a quantity");
  assert.strictEqual(W.weldStripDigits("cut 12 x 345 today"), "cut 12 x 345 today");
  assert.strictEqual(W.weldStripDigits("D02 X285"), "…", "an eircode with its space goes");
  assert.strictEqual(W.weldStripDigits("A65F4E2"), "…", "and without it");
  assert.strictEqual(W.weldStripDigits("D6W 1234"), "…", "including the D6W routing key");
  assert.strictEqual(W.weldStripDigits("R5303"), "R5303", "a job number is not an eircode and stays");
  assert.strictEqual(W.weldStripDigits("  two   spaces  "), "two spaces", "whitespace is collapsed");
  assert.strictEqual(W.weldStripDigits(null), "");
  assert.strictEqual(W.weldStripDigits("x".repeat(400)).length, W.WELD_COMMENT_MAX, "capped at 140");
  pass("rule 3: six or more digits and anything eircode-shaped are stripped out of every comment");

  const everything = JSON.stringify(W.weldSlice(JOBS, NAMES, () => ""));
  ["0871234567", "D02X285", "D02 X285", "Cork"].forEach(bad =>
    assert.strictEqual(everything.indexOf(bad), -1, "the slice leaked " + bad));
  pass("nothing in the whole slice carries a phone number, an eircode or a county");

  /* ================= 3. the seed ================= */
  const gold = (j, it) => (it === "prod:casement windows:f" ? "done" : "");
  slice = W.weldSlice(JOBS, NAMES, gold);
  assert.deepStrictEqual(W.weldSeedFields(slice[0]), { FramesDone: 3, SashesDone: 0 },
    "gold on the office's record seeds that component at its whole quantity, and nothing else");
  const yellow = () => "process";
  assert.deepStrictEqual(W.weldSeedFields(W.weldSlice(JOBS, NAMES, yellow)[0]),
    { FramesDone: 0, SashesDone: 0 }, "yellow seeds nothing (owner, answer 5)");
  assert.deepStrictEqual(W.weldSeedFields(W.weldSlice(JOBS, NAMES, () => "")[0]),
    { FramesDone: 0, SashesDone: 0 });
  const bothGold = () => "done";
  assert.deepStrictEqual(W.weldSeedFields(W.weldSlice(JOBS, NAMES, bothGold)[0]),
    { FramesDone: 3, SashesDone: 8 });
  /* a statusOf that throws must not take the feed down with it */
  assert.deepStrictEqual(W.weldSeedFields(W.weldSlice(JOBS, NAMES, () => { throw new Error("x"); })[0]),
    { FramesDone: 0, SashesDone: 0 }, "a record that will not answer seeds nothing rather than throwing");
  pass("the seed: gold is that component's whole quantity, yellow is nothing, and a throw is nothing");

  /* ================= 4. feedPlan with the welding definition ================= */
  const at = "2026-09-16T09:00:00.000Z";
  let plan = ST.feedPlan(slice, [], { at: at, by: "the office", def: W.WELD });
  assert.strictEqual(plan.adds.length, 3);
  assert.strictEqual(plan.patches.length, 0);
  assert.strictEqual(plan.adds[0].Title, "R7001|CASEMENT WINDOWS");
  assert.strictEqual(plan.adds[0].FramesDone, 3, "a row being created carries its seed");
  assert.strictEqual(plan.adds[0].FedBy, "the office");
  assert.ok(!("FramesBy" in plan.adds[0]) && !("DoneAt" in plan.adds[0]),
    "but never a By, an At or the last-touch pair: the office did not do the work");
  pass("feedPlan makes one add per slice row, seeded, with FedAt/FedBy and no stamps");

  const live = [item({ Title: "R7001|CASEMENT WINDOWS", Job: "R7001", Group: "CASEMENT WINDOWS",
    GroupSeq: 0, Customer: "Customer One", Comment: cas.comment, SentToFloor: "2026-09-12",
    Wnd: 6, Drs: 1, Frames: 3, Sashes: 8, Seq: 0, Section: "In production", Active: "Yes",
    FramesDone: 3, SashesDone: 0 }, "500")];
  plan = ST.feedPlan([slice[0]], live, { at: at, by: "the office", def: W.WELD });
  assert.strictEqual(plan.adds.length, 0);
  assert.strictEqual(plan.patches.length, 0);
  assert.strictEqual(plan.unchanged, 1, "a row that already says it is left completely alone");
  pass("a row the list already agrees with is not written at all");

  live[0].fields.Sashes = 9;                    // the office changed the quantity on the sheet
  plan = ST.feedPlan([slice[0]], live, { at: at, by: "the office", def: W.WELD });
  assert.deepStrictEqual(Object.keys(plan.patches[0].fields).sort(),
    ["FedAt", "FedBy", "Sashes"], "only the column that moved, plus the feed stamp");
  pass("a changed job fact is one patch of that column, and the feed stamp rides with it");

  /* the untouched-row guard, in the plan itself */
  live[0].fields.Sashes = 8;
  live[0].fields.FramesDone = 0;
  plan = ST.feedPlan([slice[0]], live, { at: at, by: "x", def: W.WELD });
  assert.strictEqual(plan.patches[0].fields.FramesDone, 3,
    "a row the floor has never tapped is still the office's to seed");
  live[0].fields.DoneAt = "2026-09-16T10:00:00.000Z";
  plan = ST.feedPlan([slice[0]], live, { at: at, by: "x", def: W.WELD });
  assert.strictEqual(plan.patches.length, 0,
    "and the moment DoneAt is set, the feeder never writes a counter on it again");
  assert.ok(!JSON.stringify(plan).match(/FramesDone|SashesDone/),
    "not in any shape, on any path");
  delete live[0].fields.DoneAt; live[0].fields.FramesDone = 3;
  pass("the untouched-row guard: a row the floor has tapped keeps every number on it");

  /* a job that leaves the sheet is marked inactive, never deleted */
  const gone = [item({ Title: "R9999|PVC DOOR", Job: "R9999", Group: "PVC DOOR", Frames: 1,
                       Sashes: 1, Active: "Yes", Seq: 9 }, "900")];
  plan = ST.feedPlan([], gone, { at: at, by: "x", def: W.WELD });
  assert.deepStrictEqual(plan.patches[0].fields, { Active: "No", FedAt: at, FedBy: "x" });
  assert.strictEqual(JSON.stringify(plan).indexOf("delete"), -1);
  assert.ok(!("deletes" in plan), "there is no delete in a plan, and nowhere to put one");
  plan = ST.feedPlan([], gone.map(x => item(Object.assign({}, x.fields, { Active: "No" }), x.id)),
                     { at: at, by: "x", def: W.WELD });
  assert.strictEqual(plan.patches.length, 0, "and it is only said once");
  pass("a job that leaves the sheet is marked Active = No and never, ever deleted");

  /* the hash, and that it notices everything worth re-feeding */
  const h0 = ST.sliceHash(slice, W.WELD);
  const moved = JSON.parse(JSON.stringify(slice));
  moved[0].comment = "something else";
  assert.notStrictEqual(ST.sliceHash(moved, W.WELD), h0, "a changed comment is a changed hash");
  const moved2 = JSON.parse(JSON.stringify(slice));
  moved2[0].seed.framesDone = 0;
  assert.notStrictEqual(ST.sliceHash(moved2, W.WELD), h0, "and so is a changed seed");
  const moved3 = JSON.parse(JSON.stringify(slice));
  moved3[0].sentToFloor = "2026-09-13";
  assert.notStrictEqual(ST.sliceHash(moved3, W.WELD), h0, "and a changed sent-to-floor date");
  assert.strictEqual(ST.sliceHash(JSON.parse(JSON.stringify(slice)), W.WELD), h0, "the same slice, the same hash");
  pass("the slice hash moves on every fact the floor would see, so nothing waits ten minutes for it");

  /* the glass definition is the default, and is untouched by any of this */
  const gslice = ST.glassSlice([mkJob({ id: "R8001", glass: { dg: 4, tg: 4 }, blk: 4 })], NAMES);
  assert.strictEqual(ST.sliceHash(gslice), ST.sliceHash(gslice, ST.GLASS),
    "omitting the definition IS the glass definition");
  assert.deepStrictEqual(ST.feedPlan(gslice, [], { at: at }),
                         ST.feedPlan(gslice, [], { at: at, def: ST.GLASS }),
    "and so does omitting it from a plan");
  assert.strictEqual(ST.GLASS.site, "own", "and glass is pinned to the site its lists are already in");
  assert.strictEqual(W.WELD.site, "floor", "while welding resolves the Floor stations site");
  pass("station-core's generalisation defaults to glass, so every older call means what it meant");

  /* ================= 5. the board, and the three colour levels ================= */
  assert.strictEqual(W.weldColour(0, 6), "");
  assert.strictEqual(W.weldColour(1, 6), "yellow");
  assert.strictEqual(W.weldColour(6, 6), "green");
  assert.strictEqual(W.weldColour(9, 6), "green", "a number past its total is still green, never 150%");
  assert.strictEqual(W.weldColour(0, 0), "", "nothing to weld has finished nothing");
  assert.strictEqual(W.weldRollUp(["green", "green"]), "green");
  assert.strictEqual(W.weldRollUp(["green", ""]), "yellow", "one line short is not a finished group");
  assert.strictEqual(W.weldRollUp(["", ""]), "");
  assert.strictEqual(W.weldRollUp([]), "");
  pass("the colour rule: none, yellow, green - at the line, the group and the card");

  const rows = [
    item({ Title: "R7001|CASEMENT WINDOWS", Job: "R7001", Group: "CASEMENT WINDOWS", GroupSeq: 0,
           Customer: "Customer One", Comment: "a word", SentToFloor: "2026-09-12", Wnd: 6, Drs: 1,
           Frames: 3, Sashes: 8, Seq: 0, Section: "In production", Active: "Yes",
           FramesDone: 3, SashesDone: 0, FramesBy: "Person A", FramesAt: "2026-09-16T11:00:00.000Z",
           DoneBy: "Person A", DoneAt: "2026-09-16T11:00:00.000Z" }, "500"),
    item({ Title: "R7001|PVC DOOR", Job: "R7001", Group: "PVC DOOR", GroupSeq: 2,
           Customer: "Customer One", Wnd: 6, Drs: 1, Frames: 1, Sashes: 1, Seq: 0,
           Section: "In production", Active: "Yes", FramesDone: 1, SashesDone: 1 }, "501"),
    item({ Title: "R7005|7000 CASEMENT", Job: "R7005", Group: "7000 CASEMENT", GroupSeq: 1,
           Customer: "Customer Five", Frames: 2, Sashes: 2, Seq: 1, Section: "In production",
           Active: "Yes", FramesDone: 2, SashesDone: 2 }, "502"),
    item({ Title: "R7006|PVC DOOR", Job: "R7006", Group: "PVC DOOR", GroupSeq: 0,
           Customer: "Customer Six", Frames: 1, Sashes: 0, Seq: 2, Section: "Finished",
           Active: "Yes", FramesDone: 0, SashesDone: 0 }, "503"),
    item({ Title: "R7007|PVC DOOR", Job: "R7007", Group: "PVC DOOR", GroupSeq: 0,
           Customer: "Customer Seven", Frames: 1, Sashes: 1, Seq: 3, Section: "In production",
           Active: "No", FramesDone: 0, SashesDone: 0 }, "504")
  ];
  let board = W.weldBoard(rows);
  assert.deepStrictEqual(board.map(c => c.job), ["R7001", "R7005"],
    "the tablet shows In production and Active = Yes, in the office's own order");
  const c1 = board[0];
  assert.strictEqual(c1.groups.length, 2, "one card per job, its groups inside it");
  assert.deepStrictEqual(c1.groups.map(g => g.group), ["CASEMENT WINDOWS", "PVC DOOR"],
    "in the sheet's own group order");
  assert.strictEqual(c1.customer, "Customer One");
  assert.strictEqual(c1.comment, "a word", "the job's facts are repeated on every row and read off any");
  assert.strictEqual(c1.sentToFloor, "2026-09-12");
  assert.strictEqual(c1.done, 5);
  assert.strictEqual(c1.total, 13);
  assert.strictEqual(c1.colour, "yellow");
  assert.strictEqual(c1.groups[0].colour, "yellow");
  assert.strictEqual(c1.groups[1].colour, "green");
  assert.strictEqual(c1.finished, false);
  assert.strictEqual(board[1].finished, true);
  assert.strictEqual(board[1].colour, "green");
  pass("one card per job, its groups in sheet order, coloured at all three levels");

  assert.deepStrictEqual(c1.groups[0].lines.map(l => l.label), ["Frames", "Sashes"]);
  assert.strictEqual(c1.groups[1].lines.length, 2);
  const noSash = W.weldRecord(item({ Title: "R1|PVC DOOR", Job: "R1", Group: "PVC DOOR",
                                     Frames: 2, Sashes: 0, FramesDone: 0 }, "1"));
  assert.strictEqual(noSash.lines.length, 1, "a line with nothing to weld is not drawn");
  assert.strictEqual(noSash.lines[0].part, "frames");
  assert.strictEqual(noSash.total, 2, "and it counts for nothing in the group's own total");
  pass("a line with nothing to weld is not drawn, and does not count towards anything");

  const off = W.weldOfficeBoard(rows);
  assert.deepStrictEqual(off.map(c => c.job), ["R7001", "R7005", "R7006"],
    "the office sees every section, the finished ones included, and not the inactive one");
  /* the drawer's line reads a job that has left the TABLET's board but is still
     on the sheet - R7006 is in the Finished section and still Active */
  assert.strictEqual(W.weldJobCard(rows, "R7006").job, "R7006",
    "the drawer reads a job that is off the floor's board but still on the sheet");
  assert.strictEqual(W.weldJobCard(rows, "R7006").total, 1);
  /* ... and NOT one whose groups have left the sheet. Active = No rows are
     never deleted, so an old group would otherwise pile up behind a job and
     make the drawer's "Welding 9 / 16" a total no board anywhere agrees with. */
  assert.strictEqual(W.weldJobCard(rows, "R7007"), null,
    "a job whose groups are all Active = No is not counted in the drawer's line");
  assert.strictEqual(W.weldJobCard(rows, "NOSUCH"), null);
  /* the mixed case, which is the one that actually bites: one live group and
     one that has left the sheet. Only the live one counts. */
  const mixed = rows.concat([item({ Title: "R7005|PVC DOOR", Job: "R7005", Group: "PVC DOOR",
    GroupSeq: 5, Customer: "Customer Five", Frames: 9, Sashes: 9, Seq: 1,
    Section: "In production", Active: "No", FramesDone: 0, SashesDone: 0 }, "505")]);
  assert.strictEqual(W.weldJobCard(mixed, "R7005").total, 4,
    "an inactive group of a live job is left out of the drawer's total too");
  assert.strictEqual(W.weldJobCard(mixed, "R7005").groups.length, 1);
  pass("the office's board is every section; the tablet's is In production only");

  /* a duplicate Title: the oldest wins, and nothing is deleted over it */
  const dupes = rows.concat([item({ Title: "R7005|7000 CASEMENT", Job: "R7005",
    Group: "7000 CASEMENT", Frames: 2, Sashes: 2, Seq: 1, Section: "In production",
    Active: "Yes", FramesDone: 0, SashesDone: 0 }, "999")]);
  const dboard = W.weldBoard(dupes);
  assert.strictEqual(dboard.find(c => c.job === "R7005").done, 4,
    "a duplicate row is ignored in favour of the oldest id, exactly as the phases list does");
  pass("two rows carrying one Title: the oldest is the one both boards read");

  /* ================= 6. "N left" and the two filters ================= */
  assert.strictEqual(W.weldLeft(board), 8, "frames left plus sashes left over the cards given");
  assert.strictEqual(W.weldLeftWords(8), "8 left");
  assert.strictEqual(W.weldLeft([]), 0);
  pass("the header's N left is the board's own, added over every card on it");

  /* weldLeftByPart: the same board number as weldLeft, split into the two
     header capsules (Frames left, Sashes left). A Super door group draws no
     Frames line at all (WELD_GROUP_PARTS feeds it sashes only, Frames: 0 on
     the row) and must contribute 0 to frames left, not be skipped from the
     sum or throw it off. */
  const byPartRows = [
    item({ Title: "R9001|PVC DOOR", Job: "R9001", Group: "PVC DOOR", GroupSeq: 0,
      Frames: 5, Sashes: 5, FramesDone: 2, SashesDone: 1, Seq: 1,
      Section: "In production", Active: "Yes" }, "901"),
    item({ Title: "R9002|SUPER DOOR", Job: "R9002", Group: "SUPER DOOR", GroupSeq: 0,
      Frames: 0, Sashes: 4, FramesDone: 0, SashesDone: 1, Seq: 2,
      Section: "In production", Active: "Yes" }, "902"),
  ];
  const byPartBoard = W.weldBoard(byPartRows);
  const byPart = W.weldLeftByPart(byPartBoard);
  assert.deepStrictEqual(byPart, { frames: 3, sashes: 7 },
    "frames left is 5-2 plus 0 from the Super door group with no frames line; sashes left is (5-1)+(4-1)");
  assert.strictEqual(byPart.frames + byPart.sashes, W.weldLeft(byPartBoard),
    "the two capsules must add back to the one number weldLeft gives");
  pass("weldLeftByPart splits the header's N left into frames and sashes, and a Super door contributes 0 frames");

  assert.deepStrictEqual(W.weldSentFilter(board, false).map(c => c.job), ["R7001", "R7005"]);
  assert.deepStrictEqual(W.weldSentFilter(board, true).map(c => c.job), ["R7001"],
    "the chip hides a job with a blank sent-to-floor date");
  assert.deepStrictEqual(W.weldFilter(board, "7005").map(c => c.job), ["R7005"]);
  assert.deepStrictEqual(W.weldFilter(board, "customer one").map(c => c.job), ["R7001"]);
  assert.deepStrictEqual(W.weldFilter(board, "pvc door").map(c => c.job), ["R7001"],
    "the box finds a group as well as a job and a customer");
  assert.strictEqual(W.weldFilter(board, "").length, 2, "an empty box is every card");
  pass("the search box and the sent-to-floor chip narrow the board and never become it");

  /* ================= 7. the tap ================= */
  const rec = W.weldRecord(item({ Title: "R7001|CASEMENT WINDOWS", Job: "R7001",
    Group: "CASEMENT WINDOWS", Frames: 6, Sashes: 8, FramesDone: 3, SashesDone: 0 }, "600"));
  assert.strictEqual(W.weldApplyTap(rec, "frames", 1), 4);
  assert.strictEqual(W.weldApplyTap(rec, "frames", -1), 2);
  assert.strictEqual(W.weldApplyTap(rec, "frames", "all"), 6);
  assert.strictEqual(W.weldApplyTap(rec, "frames", "none"), 0);
  assert.strictEqual(W.weldApplyTap(rec, "frames", 99), 6, "clamped to the quantity, never past it");
  assert.strictEqual(W.weldApplyTap(rec, "sashes", -1), 0, "and never below nought");
  /* a finished line is still tappable, and reducing it is what brings a card
     back off the Finished group */
  const full = W.weldRecord(item({ Title: "R1|PVC DOOR", Job: "R1", Group: "PVC DOOR",
    Frames: 1, Sashes: 1, FramesDone: 1, SashesDone: 1 }, "601"));
  assert.strictEqual(W.weldApplyTap(full, "frames", -1), 0);
  assert.strictEqual(W.weldApplyTap(full, "frames", 1), 1, "and cannot be pushed past its quantity");
  assert.strictEqual(W.weldApplyTap(rec, "transoms", 1), null, "there is no transom to tap");
  assert.strictEqual(W.weldApplyTap(rec, "", 1), null);
  pass("a tap is a clamp and nothing else - a finished line can still be reduced");

  const tf = W.weldTapFields("frames", 4, "Person A", "2026-09-16T12:00:00.000Z");
  assert.deepStrictEqual(tf, { FramesDone: 4, FramesBy: "Person A",
    FramesAt: "2026-09-16T12:00:00.000Z", DoneBy: "Person A", DoneAt: "2026-09-16T12:00:00.000Z" });
  assert.strictEqual(W.weldTapFields("nonsense", 1, "x", "y"), null);
  const dirty = W.weldFloorOnly(Object.assign({}, tf,
    { Job: "R7001", Customer: "Somebody", Active: "No", Section: "Finished", Frames: 99 }));
  assert.deepStrictEqual(dirty, tf, "floorOnly drops every job fact, however it got in");
  assert.deepStrictEqual(W.weldFloorOnly({ FramesDone: "4" }), {},
    "a counter that is not a number is not a counter");
  assert.deepStrictEqual(W.weldFloorOnly({ FramesBy: 7 }), {}, "and a stamp that is not text is not a stamp");
  pass("one tap writes that part's counter, its By/At and the last-touch pair, and can carry nothing else");

  /* ================= 7b. the remake count (2026-09-23) ================= */
  const rm = W.weldRecord(item({ Title: "R7001|CASEMENT WINDOWS", Job: "R7001",
    Group: "CASEMENT WINDOWS", Frames: 6, Sashes: 8, FramesDone: 6, SashesDone: 0,
    FramesRemade: 2 }, "602"));
  assert.strictEqual(rm.remade.frames, 2);
  assert.strictEqual(rm.remade.sashes, 0, "a blank column is nought, not NaN");
  assert.strictEqual(rm.lines[0].remade, 2, "the line carries it for the drawing");
  assert.strictEqual(rm.frames, 6); assert.strictEqual(rm.colour, "yellow");
  assert.strictEqual(rm.left, 8, "and it moves no done count, no colour and no left");
  assert.strictEqual(W.weldApplyTap(rm, "frames-remake", 1), 3);
  assert.strictEqual(W.weldApplyTap(rm, "frames-remake", -1), 1);
  assert.strictEqual(W.weldApplyTap(rm, "frames-remake", 99), 101, "no ceiling: it can pass the quantity");
  assert.strictEqual(W.weldApplyTap(rm, "sashes-remake", -1), 0, "and never below nought");
  assert.strictEqual(W.weldApplyTap(rm, "frames-remake", "all"), 2, "All means nothing to a remake count");
  assert.strictEqual(W.weldApplyTap(rm, "frames-remake", "none"), 0);
  const rtf = W.weldTapFields("sashes-remake", 3, "Person A", "2026-09-23T12:00:00.000Z");
  assert.deepStrictEqual(rtf, { SashesRemade: 3 },
    "a remake tap writes the remake count ALONE: no By/At, no DoneBy/DoneAt (a remake on a fresh row " +
    "must not close the feeder's untouched-row seed), never a done count");
  assert.deepStrictEqual(W.weldFloorOnly(Object.assign({ Frames: 9, Active: "No" }, rtf)), rtf,
    "and the whitelist lets exactly that through");
  assert.deepStrictEqual(W.weldFloorOnly({ FramesRemade: "3" }), {}, "a remake count that is not a number is dropped");
  const rq = { part: "frames-remake", value: 3, from: 2, at: "2026-09-23T12:00:00.000Z" };
  assert.deepStrictEqual(W.weldRebase(rq, { FramesRemade: 5, Frames: 6 }),
    { action: "rebase", value: 6, from: 5 }, "re-based on a rise like any counter, and NOT clamped to Frames");
  assert.deepStrictEqual(W.weldRebase(rq, { FramesRemade: 1, Frames: 6, DoneAt: "2026-09-23T12:30:00.000Z" }),
    { action: "drop" });
  assert.ok(W.WELD_FLOOR_FIELDS.indexOf("FramesRemade") >= 0 && W.WELD_COUNTER_FIELDS.indexOf("SashesRemade") >= 0);
  assert.ok(W.WELD_FEEDER_WRITES.indexOf("FramesRemade") < 0 && W.WELD_SEED_FIELDS.indexOf("FramesRemade") < 0,
    "the feeder never writes and never seeds a remake count");
  assert.strictEqual(W.weldPartOf("sashes-remake"), "sashes");
  assert.ok(W.weldIsRemake("frames-remake") && !W.weldIsRemake("frames"));
  assert.notStrictEqual(W.weldCardSig({ job: "R1", groups: [rm] }),
    W.weldCardSig({ job: "R1", groups: [Object.assign({}, rm, { remade: { frames: 3, sashes: 0 } })] }),
    "a changed remake count redraws the card");
  const rep = W.weldReportJobs({ board: W.weldOfficeBoard([item({ Title: "R7001|CASEMENT WINDOWS",
    Job: "R7001", Group: "CASEMENT WINDOWS", Frames: 6, Sashes: 8, FramesRemade: 2, SashesRemade: 1,
    Section: "In production", Active: "Yes" }, "603")]) });
  assert.strictEqual(rep.rows[0][rep.columns.indexOf("Frames remade")], 2);
  assert.strictEqual(rep.rows[0][rep.columns.indexOf("Sashes remade")], 1);
  assert.deepStrictEqual(W.WELD.reportLogStages(), ["frames", "sashes"],
    "and the report's Summary/Days sum real work only - a remake line is never a unit recorded");
  pass("the remake count: its own counter key, no ceiling, no effect on progress, floor-only, in the report");

  /* ================= 8. the rebase, and the office's later word ================= */
  const q = { part: "frames", value: 4, from: 3, at: "2026-09-16T12:00:00.000Z" };
  assert.deepStrictEqual(W.weldRebase(q, { FramesDone: 3, Frames: 6 }), { action: "keep" },
    "nothing moved under it: the tap stands");
  assert.deepStrictEqual(W.weldRebase(q, { FramesDone: 5, Frames: 6 }),
    { action: "rebase", value: 6, from: 5 }, "the office raised it: +1 on five is six, not four");
  assert.deepStrictEqual(W.weldRebase(q, { FramesDone: 6, Frames: 6 }),
    { action: "rebase", value: 6, from: 6 }, "clamped to the quantity on the way");
  assert.deepStrictEqual(W.weldRebase(q, { FramesDone: 1, Frames: 6,
    DoneAt: "2026-09-16T12:30:00.000Z" }), { action: "drop" },
    "THE SPEC'S CASE: a queued floor tap older than an office edit is dropped");
  assert.deepStrictEqual(W.weldRebase(q, { FramesDone: 1, Frames: 6,
    DoneAt: "2026-09-16T11:30:00.000Z" }), { action: "keep" },
    "but an OLDER stamp is not a later word, and the floor's tap stands");
  assert.deepStrictEqual(W.weldRebase(q, { FramesDone: 1, Frames: 6, DoneAt: "yesterday" }),
    { action: "keep" }, "a stamp nothing can parse is not a time, and decides nothing");
  assert.deepStrictEqual(W.weldRebase(q, { FramesDone: 1, Frames: 6 }), { action: "keep" },
    "and neither is a fall with no stamp at all");
  assert.deepStrictEqual(W.weldRebase({ part: "nonsense" }, {}), { action: "keep" });
  pass("a queued tap is re-based on a rise, dropped under a later office edit, and kept otherwise");

  /* ================= 9. the office's edit, on the real app.js ================= */
  consent(true);
  WITEMS = rows.map(r => item(Object.assign({}, r.fields), r.id));
  A("WELD_ITEMS = null; WELD_OK = null; WELD_WHY = ''; WELD_SITEID = null; WRECS = null; WRECS_OF = false;");
  A("WELD_FEED = { hash: '', at: 0 };");
  reset();
  const read = await A("readWelding()");
  assert.ok(read, "the office reads the welding list");
  assert.strictEqual(A("WELD_OK"), true);
  assert.ok(REQ.every(r => r.path.indexOf("/drive") < 0 && r.path.indexOf("/workbook") < 0),
    "and not one request went anywhere near the workbook");
  /* EVERY, not some: "every list call names the Floor stations site" is only
     worth asserting if it is asked of all of them. `some` passed while any
     number of calls went elsewhere. */
  const listCalls = REQ.filter(r => r.path.indexOf("/sites/") === 0 && r.path.indexOf("/lists") > 0);
  assert.ok(listCalls.length > 0, "there were list calls to check");
  assert.ok(listCalls.every(r => r.path.indexOf(FLISTS_PATH) === 0),
    "every list call names the Floor stations site: " +
    JSON.stringify(listCalls.map(r => r.path).filter(p => p.indexOf(FLISTS_PATH) !== 0)));
  pass("the office reads the Welding station list out of the Floor stations site, and no workbook path");

  /* the Dashboard Log line an office action leaves is a WORKBOOK write - the
     `Dashboard Log` sheet is one of the dashboard's own (rule 2), and it is the
     same appendLog every other office action uses. It is stubbed here so the
     line can be asserted on directly, and so a fire-and-forget request at a
     sheet this harness deliberately does not serve cannot make the run flaky. */
  const LOGGED = [];
  CW.appendLog = async (...a) => { LOGGED.push(a); };
  reset();
  const edited = await A('weldOfficeEdit("500", "sashes", "all")');
  assert.strictEqual(edited, true);
  const patch = writes().find(w => w.method === "PATCH");
  assert.ok(patch, "one PATCH went out");
  assert.deepStrictEqual(Object.keys(patch.body).sort(),
    ["DoneAt", "DoneBy", "SashesAt", "SashesBy", "SashesDone"],
    "the office's edit writes exactly that part's counter, its By/At and the last-touch pair");
  assert.strictEqual(patch.body.SashesDone, 8);
  assert.ok(patch.body.DoneBy && patch.body.DoneBy.length, "with the office person's name on it");
  assert.strictEqual(writes().filter(w => w.path.indexOf(LOG_ID) >= 0).length, 0,
    "RULE 2: the office writes no line of Station log, on this station or any other");
  assert.strictEqual(writes().filter(w => w.method === "DELETE").length, 0);
  assert.ok(writes().every(w => w.path.indexOf("/workbook") < 0 && w.path.indexOf("/drive") < 0),
    "and the office's welding edit never touches the workbook either");
  pass("an office edit is one PATCH of five fields, no Station log line and no workbook write");

  /* ... and it leaves a Dashboard Log line, which is where an office action goes */
  const logged = A("CHANGES[0]");
  assert.ok(logged, "the change is in the office's own Changes list");
  assert.strictEqual(logged.job, "R7001");
  assert.strictEqual(logged.what, "Welding: R7001 CASEMENT WINDOWS sashes");
  assert.strictEqual(logged.from, "0");
  assert.strictEqual(logged.to, "8");
  assert.strictEqual(LOGGED.length, 1, "exactly one line, in the Dashboard Log sheet");
  assert.deepStrictEqual(LOGGED[0].slice(1),
    ["R7001", "Welding: R7001 CASEMENT WINDOWS sashes", "0", "8"]);
  pass("one Dashboard Log line per office change, naming the job, the group and the part");

  /* the board reads the new number at once, without waiting for a poll */
  A("WRECS = null; WRECS_OF = false;");
  const after = A('weldRecordsNow().byId["500"]');
  assert.strictEqual(after.sashes, 8);
  assert.strictEqual(after.colour, "green");
  pass("the office's own copy of the list is in step at once, so the board never reads a stale number");

  /* a click that would change nothing writes nothing */
  reset();
  assert.strictEqual(await A('weldOfficeEdit("500", "sashes", "all")'), false);
  assert.strictEqual(writes().length, 0, "All on a line that is already all is not a write");
  assert.strictEqual(await A('weldOfficeEdit("nosuchrow", "frames", 1)'), false);
  assert.strictEqual(writes().length, 0);
  pass("a click that could not move the number makes no request at all");

  /* the remake count is read on the board and never written by the office */
  reset();
  assert.strictEqual(await A('weldOfficeEdit("500", "frames-remake", 1)'), false,
    "the office cannot move a remake count");
  assert.strictEqual(writes().length, 0, "and no request goes out for the try");
  WITEMS.find(x => x.id === "500").fields.FramesRemade = 3;
  A("WELD_ITEMS = null; WRECS = null; WRECS_OF = false;");
  await A("readWelding()");
  A("WELD_OPEN = { R7001: 1 };");
  const remHtml = A("weldBoardHtml()");
  assert.ok(/worem tab has[^>]*>↻ 3</.test(remHtml), "the open row shows the floor's 3 remakes on its Frames line");
  assert.ok(!/data-wpart="frames-remake"/.test(remHtml), "and offers no button for it");
  A("WELD_OPEN = {};");
  pass("the office reads the remake count on its board and has no way to write one");

  /* ---- the office's board is up to ten seconds old (review finding 4) ----
     A welder taps All (49 of 49) at 14:00:01. This screen last polled at
     13:59:56 and is drawing 3 of 49. The office presses + at 14:00:05. Derived
     from the BOARD that writes 4 - forty-six taps destroyed, with the office's
     later stamp on it so nothing can argue them back. So the row is read
     immediately before the PATCH and the number is derived from what it says. */
  A("WELD_ITEMS = null; WELD_OK = null; WELD_SITEID = null; WRECS = null; WRECS_OF = false;");
  WITEMS = [item({ Title: "R7900|CASEMENT WINDOWS", Job: "R7900", Group: "CASEMENT WINDOWS",
    GroupSeq: 0, Customer: "Customer Nine", Frames: 49, Sashes: 0, Seq: 0,
    Section: "In production", Active: "Yes", FramesDone: 3, SashesDone: 0,
    DoneBy: "Person A", DoneAt: "2026-09-16T13:59:50.000Z" }, "790")];
  await A("readWelding()");
  assert.strictEqual(A('weldRecordsNow().byId["790"].frames'), 3, "the board is drawing 3 of 49");
  /* the floor finishes the job while this screen is still showing 3 */
  WITEMS[0].fields.FramesDone = 49;
  WITEMS[0].fields.FramesBy = "Person A";
  WITEMS[0].fields.FramesAt = "2026-09-16T14:00:01.000Z";
  WITEMS[0].fields.DoneBy = "Person A";
  WITEMS[0].fields.DoneAt = "2026-09-16T14:00:01.000Z";
  A("CHANGES = [];");
  reset(); LOGGED.length = 0; TOASTS.length = 0;
  const bumped = await A('weldOfficeEdit("790", "frames", 1)');
  assert.ok(REQ.some(r => r.method === "GET" && /\/items\/790/.test(r.path)),
    "the row was read immediately before anything was written");
  assert.strictEqual(bumped, false,
    "and + on a row the floor has just finished writes nothing at all");
  assert.strictEqual(writes().length, 0, "not one PATCH went out");
  assert.strictEqual(WITEMS[0].fields.FramesDone, 49,
    "so the floor's 49 is exactly where they left it - the stale board would have written 4");
  assert.strictEqual(A("CHANGES.length"), 0, "and there is no Dashboard Log line about a non-event");
  assert.ok(TOASTS.some(t => /updated from the floor first/.test(t.m)),
    "the office is told the row moved under them: " + JSON.stringify(TOASTS.map(t => t.m)));
  pass("an office edit re-reads the row first, so a floor tap it had not seen is never overwritten");

  /* and when the office's click CAN still move it, the number it moves is the
     fresh one: the floor is at 20, the board still says 3, + writes 21 */
  WITEMS[0].fields.FramesDone = 20;
  WITEMS[0].fields.DoneAt = "2026-09-16T14:00:03.000Z";
  A("CHANGES = [];");
  reset(); TOASTS.length = 0;
  assert.strictEqual(await A('weldOfficeEdit("790", "frames", 1)'), true);
  const got = writes().find(w => w.method === "PATCH");
  assert.strictEqual(got.body.FramesDone, 21,
    "the step is taken from the row's real number, not from the board's stale one");
  assert.strictEqual(A("CHANGES[0].from"), "20",
    "and the Dashboard Log line says what really changed, not what the board was showing");
  assert.strictEqual(A("CHANGES[0].to"), "21");
  pass("and an office step is taken from the row's real number, not the board's");

  /* the same race, the other way: the floor took it DOWN, so + is 1 not 4 */
  WITEMS[0].fields.FramesDone = 0;
  WITEMS[0].fields.DoneAt = "2026-09-16T14:00:09.000Z";
  A("CHANGES = [];");
  reset(); TOASTS.length = 0;
  await A('weldOfficeEdit("790", "frames", 1)');
  const down = writes().find(w => w.method === "PATCH");
  assert.strictEqual(down.body.FramesDone, 1,
    "+ on a row the floor has just cleared is one, not one more than the stale number");
  pass("and a counter the floor took DOWN is stepped from where it really is");

  /* a click the floor has already made is not a second write */
  WITEMS[0].fields.FramesDone = 49;
  WITEMS[0].fields.DoneAt = "2026-09-16T14:00:20.000Z";
  A("CHANGES = [];");
  reset(); TOASTS.length = 0;
  assert.strictEqual(await A('weldOfficeEdit("790", "frames", "all")'), false,
    "All on a row the floor has already finished writes nothing");
  assert.strictEqual(writes().length, 0);
  assert.strictEqual(A("CHANGES.length"), 0, "and leaves no Dashboard Log line either");
  assert.ok(TOASTS.some(t => /updated from the floor first/.test(t.m)),
    "but does say why nothing happened");
  pass("a click the floor has already made writes nothing, and says so rather than going quiet");

  /* a read that will not answer is never a reason to write blind */
  A("WELD_ITEMS = null; WELD_OK = null; WRECS = null; WRECS_OF = false;");
  await A("readWelding()");
  const keepItems = WITEMS.slice();
  WITEMS.length = 0;                        // the row is gone: the GET 404s
  reset(); TOASTS.length = 0;
  assert.strictEqual(await A('weldOfficeEdit("790", "frames", 1)'), false);
  assert.strictEqual(writes().length, 0, "nothing was written on a row that could not be read");
  assert.ok(TOASTS.some(t => t.err), "and the office is told, rather than left guessing");
  WITEMS.push.apply(WITEMS, keepItems);
  pass("a row that cannot be re-read is not written to at all");


  /* ================= 10. the feeder, on the real app.js ================= */
  WITEMS = []; LOGITEMS = [];
  global.__jobs = JOBS;
  A("ALL = __jobs; BLOCKNAMES = " + JSON.stringify(NAMES) + "; ALL.blockNames = BLOCKNAMES;");
  A("WELD_ITEMS = null; WELD_OK = null; WELD_FEED = { hash: '', at: 0 }; WRECS = null; WRECS_OF = false;");
  A("CP_ITEMS = []; if (typeof cpListRebuild === 'function') cpListRebuild();");
  reset();
  let r = await A("feedWelding()");
  assert.ok(r && r.sent === 3, "three rows fed: two groups of R7001 and one of R7003");
  assert.deepStrictEqual(WITEMS.map(x => x.fields.Title).sort(),
    ["R7001|CASEMENT WINDOWS", "R7001|PVC DOOR", "R7003|7000 CASEMENT"]);
  assert.ok(writes().every(w => w.method === "POST"), "three new rows: three POSTs and nothing else");
  assert.strictEqual(writes().filter(w => w.method === "DELETE").length, 0, "the feeder never deletes");
  assert.ok(REQ.every(w => w.path.indexOf("/workbook") < 0 && w.path.indexOf("/drive") < 0),
    "and the feeder never touches the workbook");
  assert.ok(WITEMS.every(x => !/aluclad/i.test(String(x.fields.Title))), "the denied group is not on the list");
  assert.ok(WITEMS.every(x => !/R7002/.test(String(x.fields.Title))), "and neither is the T-only job");
  const feedJson = JSON.stringify(WITEMS);
  ["0871234567", "D02X285", "Cork", "@"].forEach(bad =>
    assert.strictEqual(feedJson.indexOf(bad), -1, "the fed list leaked " + bad));
  pass("the feeder pushes the sheet's welding into the list, never deletes, never touches the workbook");

  reset();
  assert.strictEqual(await A("feedWelding()"), null, "an unchanged slice inside ten minutes is skipped");
  assert.strictEqual(writes().length, 0);
  pass("an unchanged slice is not re-fed, so a dashboard left open writes nothing all afternoon");

  /* the floor taps a row; the feeder must never write its counters again */
  const mine = WITEMS.find(x => x.fields.Title === "R7001|CASEMENT WINDOWS");
  mine.fields.FramesDone = 2;
  mine.fields.DoneAt = "2026-09-16T13:00:00.000Z";
  mine.fields.DoneBy = "Person A";
  A("WELD_FEED = { hash: '', at: 0 };");
  reset();
  await A("feedWelding()");
  writes().forEach(w => {
    assert.ok(!("FramesDone" in (w.body || {})) && !("SashesDone" in (w.body || {})) &&
              !("FramesRemade" in (w.body || {})) && !("SashesRemade" in (w.body || {})),
      "the feeder wrote a counter on a row the floor has tapped: " + JSON.stringify(w.body));
    assert.ok(!("DoneBy" in (w.body || {})) && !("DoneAt" in (w.body || {})),
      "and it must never write a By, an At or the last-touch pair");
  });
  assert.strictEqual(WITEMS.find(x => x.fields.Title === "R7001|CASEMENT WINDOWS").fields.FramesDone, 2,
    "the floor's own number is exactly where they left it");
  pass("once the floor has tapped a row, the feeder never writes a counter on it again");

  /* a missing list is a state, not a crash - and it never takes the glass feed
     or the checkpoint writes with it */
  FLISTS = FLISTS.filter(l => l.id !== WELD_ID);
  CW._resetListIds();
  try { delete mem.cw_listids; } catch (e) {}
  A("WELD_OK = null; WELD_WHY = ''; WELD_FEED = { hash: '', at: 0 };");
  reset();
  assert.strictEqual(await A("feedWelding()"), null);
  assert.strictEqual(A("WELD_OK"), false);
  assert.ok(/Welding station/.test(A("WELD_WHY")), "and it says which list, plainly");
  assert.ok(/Ask the manager/.test(A("WELD_WHY")));
  assert.strictEqual(writes().length, 0, "writing nothing at all");
  assert.ok(/not in the/.test(A("weldBoardHtml()")), "the board shows the same quiet explanation");
  FLISTS.push({ id: WELD_ID, displayName: "Welding station" });
  CW._resetListIds();
  pass("a missing Welding station list is explained and writes nothing; no list is ever created by code");

  /* ================= 11. the log line the TABLET writes ================= */
  const lf = ST.logFields(W.weldLogEntry({ job: "r7001", group: "casement windows", part: "frames",
    from: 3, to: 6, who: "Person A", at: "2026-09-16T12:00:00.000Z" }));
  assert.deepStrictEqual(lf, { Title: "R7001", Station: "Welding", GlassType: "CASEMENT WINDOWS",
    Stage: "frames", From: 3, To: 6, Who: "Person A", At: "2026-09-16T12:00:00.000Z" },
    "the shared log list's own eight columns, with the product group in GlassType");
  const lrows = ST.logRows([item(lf, "1"),
    item({ Title: "R7001", Station: "Glass", GlassType: "GLASS", Stage: "cut", From: 0, To: 4,
           Who: "Person B", At: "2026-09-16T12:05:00.000Z" }, "2")], "Welding");
  assert.strictEqual(lrows.length, 1, "and the welding board reads its own station's lines only");
  assert.strictEqual(lrows[0].type, "CASEMENT WINDOWS");
  assert.strictEqual(ST.logFields(W.weldLogEntry({ job: "R7001", group: "CASEMENT WINDOWS",
    part: "frames-remake", from: 2, to: 3, who: "Person A", at: "2026-09-23T12:00:00.000Z" })).Stage,
    "frames-remake", "a remake tap's line says which counter it moved");
  pass("a welding log line is the shared list's eight columns, read back by station");

  /* ================= 12. people, and one stage ================= */
  const people = ST.stationPeople([
    item({ Title: "Person A", Station: "Welding", Stages: "weld", PIN: "", Active: "Yes" }, "1"),
    item({ Title: "Person B", Station: "Glass", Stages: "cut", PIN: "", Active: "Yes" }, "2"),
    item({ Title: "Person C", Station: "Welding", Stages: "cut", PIN: "1234", Active: "Yes" }, "3"),
    item({ Title: "Person D", Station: "Welding", Stages: "weld", PIN: "", Active: "No" }, "4")
  ], "Welding", W.WELD.stages);
  assert.deepStrictEqual(people.map(p => p.name), ["Person A", "Person C"],
    "this station's people only, and only the active ones");
  assert.deepStrictEqual(people[0].stages, ["weld"]);
  assert.deepStrictEqual(people[1].stages, [],
    "a stage this station does not have is not a stage they hold");
  assert.strictEqual(ST.canStage(people[0], "weld"), true);
  assert.strictEqual(ST.canStage(people[1], "weld"), false, "so they may move nothing");
  const glassPeople = ST.stationPeople([
    item({ Title: "Person B", Station: "Glass", Stages: "cut,hotmelt", PIN: "", Active: "Yes" }, "2")
  ], "Glass");
  assert.deepStrictEqual(glassPeople[0].stages, ["cut", "hotmelt"],
    "and omitting the stage list is still the glass station's four");
  pass("the people list is filtered by station and by that station's own stages");

  /* ================= 13. the gates ================= */
  const wsrc = src("welding-core.js") + src("welding.js");
  const uisrc = src("station-ui.js");
  const coresrc = src("station-core.js"), stsrc = src("station.js");
  ["setFill", "clearFill", "setValues", "appendLog", "saveProgress", "moveJobRow", "batchWrite",
   "/workbook", "downloadWorkbook", "parseWorkbook", "ExcelJS"].forEach(bad => {
    assert.strictEqual(wsrc.indexOf(bad), -1, "welding-core.js / welding.js must not mention " + bad);
    assert.strictEqual(uisrc.indexOf(bad), -1, "station-ui.js must not mention " + bad);
    assert.strictEqual(coresrc.indexOf(bad), -1, "station-core.js must not mention " + bad);
    assert.strictEqual(stsrc.indexOf(bad), -1, "station.js must not mention " + bad);
  });
  [wsrc, uisrc, coresrc, stsrc].forEach(s => {
    assert.strictEqual(s.indexOf('"DELETE"'), -1);
    assert.strictEqual(s.indexOf("listDelete"), -1);
  });
  pass("no station file can write a cell, a colour, a log sheet or a delete: not one of the words is in them");

  const page = src("welding.html");
  ["exceljs", "parser.js", "app.js", "checkpoints.js", "export.js", "index.html"].forEach(bad =>
    assert.strictEqual(page.toLowerCase().indexOf(bad), -1, "welding.html must not load or link " + bad));
  ["graph.js", "station-core.js", "station-ui.js", "welding-core.js", "welding.js", "msal-browser"]
    .forEach(good => assert.ok(page.indexOf(good) > 0, "welding.html loads " + good));
  assert.ok(/<html lang="en" data-theme="dark">/.test(page), "the page opens dark, like the glass one");
  assert.ok(/:root\[data-theme="light"\]/.test(page), "with the light set behind the switch");
  /* 2026-09-17, after the owner rejected the first layout: the steppers are
     56 px TALL and at least 64 px WIDE, and both buttons say so in the same
     words, because they are tapped with a work glove on. The real sizes are
     measured in the browser check at five widths; this is the standing grep
     that stops the rule being quietly dropped out of the stylesheet. */
  assert.ok(/\.sbtn \{ width:64px; height:56px/.test(page),
    "the - and + buttons are 56 px tall and 64 wide");
  /* a FIXED width on All/None, not a minimum: "None" is the wider word, and a
     minimum moved the whole block two pixels between a finished row and an
     unfinished one - under the thumb that is reaching for it */
  assert.ok(/\.sall \{ width:84px; height:56px/.test(page),
    "and All / None is a fixed width, so the column never moves between rows");
  /* one column, always: the two-column grid is what the owner rejected */
  assert.ok(!/\.grp \{[^}]*grid-template-columns:1fr 1fr/.test(page) &&
            page.indexOf("grid-template-columns:1fr 1fr") < 0,
    "and the board has no two-column break left in it at any width");
  assert.ok(/\.toprow/.test(page) && (page.match(/class="toprow"/g) || []).length === 2,
    "the header is two explicit rows, not one that wraps");
  assert.ok(page.indexOf("id=\"sentchip\"") > 0, "the sent-to-floor chip is in the header");
  pass("welding.html loads five scripts and cannot reach the workbook through any of them");

  /* rule 4: no real person, address or company domain anywhere in what was
     added. The SharePoint hostname constant in graph.js is the one exception
     and is not in any of these files. */
  const added = wsrc + uisrc + page + src("test_welding.js");
  /* built out of two pieces so this very line cannot be what the check finds */
  assert.strictEqual(added.indexOf("@" + "costello"), -1, "a company domain is in the new files");
  /* a dot in the domain is what tells an address from a font weight
     ("wght@600" in the Google Fonts link is not somebody's email) */
  const mails = added.match(/[A-Za-z0-9._%-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+/g) || [];
  mails.forEach(m => assert.ok(/example\.test$/.test(m),
    "a real-looking address is in the new files: " + m));
  pass("rule 4: no real name, address or company domain in any file this feature added");

  /* ================= 14. the site pin, on graph.js itself =================
     The review's first finding, and the one the owner ruled on: "own" must
     mean THE WORKBOOK'S OWN SITE, resolved by path, with the `Floor stations`
     site never looked up on that channel - not on an empty cache, not after a
     forget, not when somebody taps Try again. The first build of the pin
     returned the cached site and otherwise fell through to the legacy route,
     which PREFERS Floor stations; three ordinary things empty that cache (a
     404 on any list call, localStorage cleared, a new browser profile) and
     every one of them moved the glass page to a site with no glass list in it,
     permanently. These are the checks that would have caught it. */
  const floorLookups = () => REQ.filter(r => r.path.indexOf(HOST_LOOKUP) === 0);
  const ownLookups = () => REQ.filter(r => r.path.indexOf(OWN_LOOKUP.split("?")[0]) === 0);

  SITE_EXISTS = true;                       // the Floor stations site DOES exist
  CW._resetPinnedSites(); CW._setStationSite(null); CW._resetListIds();
  try { delete mem.cw_listids; } catch (e) {}
  reset();
  assert.strictEqual(await CW.stationSite("own"), SITE,
    "(i) with nothing cached, “own” resolves the workbook's own site by path");
  assert.strictEqual(floorLookups().length, 0,
    "and never once looks for Floor stations, though it is right there: " +
    JSON.stringify(floorLookups().map(r => r.path)));
  assert.strictEqual(ownLookups().length, 1, "one lookup, by path");
  assert.strictEqual(CW._pinnedSiteInfo("own").id, SITE, "and it is remembered");
  reset();
  assert.strictEqual(await CW.stationSite("own"), SITE);
  assert.strictEqual(REQ.length, 0, "and then it is not looked up again at all");
  pass("(i) an empty cache resolves the workbook's own site and never reaches for Floor stations");

  /* (ii) after a forget - which is what a 404 on any glass list call does */
  reset();
  CW.forgetStationSite(false, "own");
  assert.strictEqual(CW._pinnedSiteInfo("own").id, null, "the forget really cleared it");
  assert.strictEqual(await CW.stationSite("own"), SITE,
    "(ii) and the very next resolve is STILL the workbook's own site");
  assert.strictEqual(floorLookups().length, 0,
    "with Floor stations still never looked for: " + JSON.stringify(floorLookups().map(r => r.path)));
  reset();
  CW.forgetStationSite(true, "own");                 // what Try again does
  assert.strictEqual(await CW.stationSite("own"), SITE, "and so is the one after Try again");
  assert.strictEqual(floorLookups().length, 0);
  pass("(ii) a forget, and a Try again, cannot move the glass page to Floor stations");

  /* (iii) the other direction: "floor" has no fallback of any kind */
  SITE_EXISTS = false;
  CW._resetPinnedSites();
  reset();
  assert.strictEqual(await CW.stationSite("floor"), null,
    "(iii) a missing Floor stations site is null on the “floor” channel");
  assert.strictEqual(ownLookups().length, 0,
    "and it never falls back to the workbook's own site: " + JSON.stringify(ownLookups().map(r => r.path)));
  reset();
  assert.strictEqual(await CW.stationSite("floor"), null);
  assert.strictEqual(REQ.length, 0, "the miss is held for a minute, so a poll cannot hammer it");
  SITE_EXISTS = true;
  CW._resetPinnedSites();
  reset();
  assert.strictEqual(await CW.stationSite("floor"), FSITE, "and it is found once it exists");
  assert.strictEqual(ownLookups().length, 0);
  pass("(iii) “floor” resolves Floor stations or nothing, and never the workbook's own site");

  /* the two channels are separate answers to separate questions */
  assert.notStrictEqual(CW._pinnedSiteInfo("own").id, CW._pinnedSiteInfo("floor").id,
    "the two pinned channels hold two different sites at the same time");
  assert.strictEqual(ST.GLASS.site, "own");
  assert.strictEqual(WELDC.WELD.site, "floor");
  pass("glass and welding hold two different sites at once, and neither can answer the other's");

  /* ================= 15. rule 3: a phone number as people type one ==========
     The review's second finding. `\\d{6,}` alone caught a ten-digit block and
     nothing else, and a phone number in a workshop comment is almost never
     typed as one. Every shape below reached the floor's list unchanged. */
  [["086 123 4567", "spaces"],
   ["+353 86 123 4567", "the international form"],
   ["mob 087-123-4567", "dashes"],
   ["ring 086.123.4567", "dots"],
   ["tel: (086) 123 4567", "brackets"],
   ["08712 34567", "a stray space in the middle"],
   ["0871234567", "and the plain contiguous one"]].forEach(pair => {
    const out = W.weldStripDigits(pair[0]);
    assert.ok(out.indexOf("…") >= 0, "a phone number written with " + pair[1] +
      " was not stripped: " + JSON.stringify(out));
    assert.ok(!/\d{3}/.test(out), "and it left three digits of one behind: " + JSON.stringify(out));
  });
  pass("a phone number is stripped however it is typed: spaces, dashes, dots, brackets, +353");

  /* the fourth shape (review addendum, 2026-09-16): digit groups joined by
     `/`, `,`, `:` or `_`, none of which is in WELD_PHONE_RE's separator
     class - only stripped once the match holds nine digits or more */
  [["086/123/4567", "slashes"],
   ["086,123,4567", "commas"],
   ["086:123:4567", "colons"],
   ["phone 086_123_4567", "underscores"],
   ["mob 086‑123‑4567", "the non-breaking hyphen (U+2011)"]].forEach(pair => {
    const out = W.weldStripDigits(pair[0]);
    assert.ok(out.indexOf("…") >= 0, "a phone number written with " + pair[1] +
      " was not stripped: " + JSON.stringify(out));
    assert.ok(!/\d{3}/.test(out), "and it left three digits of one behind: " + JSON.stringify(out));
  });
  pass("a phone number survives none of slash, comma, colon, underscore or a Unicode hyphen either");

  /* ... and what must SURVIVE, or the strip is worse than useless */
  [["R5303", "a job number"],
   ["12.09", "the sheet's own short date"],
   ["12/09/2026", "and its long one"],
   ["Wnd 6", "a quantity"],
   ["4000 CASEMENT", "a product group with a number in its name"],
   ["7000", "and one that is only a number"],
   ["cut 12 x 345 today", "two small numbers in a sentence"],
   ["Customer One", "a plain name"],
   ["ref 12/34/56", "a slash-joined run short of nine digits"],
   ["spacer 4/20/4", "a small slash-joined size"],
   ["glass 4-16-4", "a dash-joined glass size"],
   ["2 x 1200 x 900", "an x-separated size"],
   ["frames 1-6", "a short dash-joined range"],
   ["bays 1, 2, 3, 4, 5, 6", "a comma-and-space list"]].forEach(pair => {
    assert.strictEqual(W.weldStripDigits(pair[0]), pair[0],
      pair[1] + " must survive the strip: " + JSON.stringify(pair[0]));
  });
  pass("a job number, the sheet's dates, quantities, product names and sizes all survive it");

  /* the CUSTOMER column gets the same guard (review finding 10) */
  const phoneJob = mkJob({ id: "R7100", cust: "Customer One 086 123 4567", blk: 4, seq: 0,
    notes: [{ k: "comment", t: "site contact 087-123-4567", s: "Production" }],
    prods: [prod("casement windows", 2, 2, 0)] });
  const phoneSlice = W.weldSlice([phoneJob], NAMES, () => "");
  assert.strictEqual(phoneSlice[0].customer, "Customer One …",
    "a number typed into a customer name is stripped like any other");
  assert.strictEqual(phoneSlice[0].comment, "site contact …");
  /* the two FREE-TEXT columns, which are the only ones a phone number can
     travel in - the job number and the quantities are numbers by construction */
  const ff = W.weldFeederFields(phoneSlice[0]);
  assert.ok(!/\d{3}/.test(ff.Customer + " " + ff.Comment),
    "and no run of three digits survives in either free-text column: " +
    JSON.stringify([ff.Customer, ff.Comment]));
  pass("the Customer column is stripped too: a site contact typed into a name never reaches the floor");

  /* the cap still applies, and is per column */
  assert.strictEqual(W.weldStripDigits("x".repeat(400)).length, W.WELD_COMMENT_MAX);
  assert.strictEqual(W.weldStripDigits("x".repeat(400), W.WELD_CUSTOMER_MAX).length, W.WELD_CUSTOMER_MAX);
  pass("the strip caps the comment at 140 and the customer at 70");

  /* ================= 16. one Title, one row (review finding 12) ============= */
  const twice = mkJob({ id: "R7200", cust: "Customer Four", blk: 4, seq: 0,
    prods: [prod("casement windows", 2, 2, 0), prod("CASEMENT  WINDOWS", 9, 9, 0)] });
  const warned = [];
  const realWarn = console.warn;
  console.warn = (...a) => { warned.push(a.join(" ")); };
  const twiceSlice = W.weldSlice([twice], NAMES, () => "");
  console.warn = realWarn;
  assert.strictEqual(twiceSlice.length, 1, "one Title makes one row, never two");
  assert.strictEqual(twiceSlice[0].frames, 2, "and it is the first one, not the last");
  assert.ok(warned.some(w => /R7200\|CASEMENT WINDOWS/.test(w) && /twice/.test(w)),
    "and the one that was dropped is said out loud: " + JSON.stringify(warned));
  pass("a product group that appears twice on one job is fed once, and the second is warned about");

  {
  /* ================= 17. the Production sheet, and nothing else ==============
     The live bug of 2026-09-17 (docs/HISTORY.md B20), in the shape it really
     had. The owner inserted a column into `Production`; `Production (2)` is
     formulas that moved with it and headers that did not, so on that sheet
     every number sits one column right of the header describing it. The
     parser's cross-sheet Math.max then gave R5053 four product groups it has
     not got, and the feeder wrote 300 wrong rows into `Welding station`.

     This is that job: `prodsMain` is what `Production` says, `prods` is what
     the merge made of it. The slice must read the first and never the second. */
  const R5053 = mkJob({
    id: "R5053", cust: "Customer Five", wnd: 23, drs: 4, blk: 4, seq: 7,
    dates: { sold: null, stamp: null, ivana: null, ready: null, floor: "15.09" },
    /* what the `Production` sheet itself says */
    prodsMain: [prod("polaris 85mm casement", 19, 16, 6), prod("polaris 85 tilt turn", 4, 6, 6),
                prod("pvc door", 3, 4, 0), prod("super door", 1, 2, 0), prod("composite", 0, 2, 0)],
    /* ... and what Math.max across the sheets made of it: two counts inflated
       from the neighbouring column, and five groups that are not on this job */
    prods: [prod("polaris 85mm casement", 19, 19, 16), prod("polaris 85 tilt turn", 4, 6, 6),
            prod("pvc door", 3, 4, 0), prod("super door", 4, 2, 0), prod("composite", 0, 2, 0),
            prod("4000 casement", 6, 0, 0), prod("4000 tilt turn", 6, 0, 0),
            prod("sidelights", 2, 0, 0), prod("bifold", 2, 0, 0), prod("pvc smart", 2, 0, 0)]
  });
  const fixed = W.weldSlice([R5053], NAMES, () => "");
  assert.deepStrictEqual(fixed.map(r => r.group),
    ["POLARIS 85MM CASEMENT", "POLARIS 85 TILT TURN", "PVC DOOR", "SUPER DOOR"],
    "four rows, and not one of the groups the merge invented");
  assert.deepStrictEqual(fixed.map(r => r.frames + "/" + r.sashes),
    ["19/16", "4/6", "3/4", "0/2"], "with the Production sheet's own numbers");
  ["4000 CASEMENT", "4000 TILT TURN", "SIDELIGHTS", "PVC SMART", "BIFOLD"].forEach(bad =>
    assert.ok(!fixed.some(r => r.group === bad), bad + " is not on this job and is not fed"));
  assert.ok(!fixed.some(r => r.group === "COMPOSITE"), "and COMPOSITE is deny-listed as before");
  pass("the slice is the Production sheet's own product list, not the cross-sheet maximum");

  /* the one that would have been silent: the count, not the group */
  assert.strictEqual(fixed[0].sashes, 16,
    "16 sashes, which is what Production says - not the 19 the neighbouring column had");
  pass("a count inflated from the next column over does not reach the floor either");

  /* a job that is not on `Production` at all is not this station's business */
  const elsewhere = mkJob({ id: "R5099", blk: 4, seq: 8, prodsMain: [],
    prods: [prod("casement windows", 4, 4, 0)] });
  assert.deepStrictEqual(W.weldSlice([elsewhere], NAMES, () => ""), [],
    "a job with nothing on Production is not fed, whatever another sheet says");
  const noneAtAll = mkJob({ id: "R5098", blk: 4, seq: 9 });
  delete noneAtAll.prodsMain;
  assert.deepStrictEqual(W.weldSlice([noneAtAll], NAMES, () => ""), [],
    "and neither is one from a parser that never filled prodsMain in");
  pass("a job that is not on the Production sheet is never fed to the welding floor");

  /* ---- the per-group component override ---- */
  assert.deepStrictEqual(W.weldPartsFor("SUPER DOOR"), ["sashes"]);
  assert.deepStrictEqual(W.weldPartsFor("super door"), ["sashes"], "matched after weldKey()");
  assert.deepStrictEqual(W.weldPartsFor("PVC DOOR"), ["frames", "sashes"], "everything else has both");
  assert.deepStrictEqual(W.weldPartsFor("ANYTHING NEW"), ["frames", "sashes"]);
  const sd = fixed.find(r => r.group === "SUPER DOOR");
  assert.strictEqual(sd.frames, 0, "a Super door's frames number is IGNORED, not carried");
  assert.strictEqual(W.weldFeederFields(sd).Frames, 0, "and the row it writes says 0");
  assert.strictEqual(sd.sashes, 2);
  /* ... and a line with nothing to weld is not drawn on either board */
  const sdRow = W.weldRecord(item(Object.assign({ Title: "R5053|SUPER DOOR" }, W.weldFeederFields(sd)), "1"));
  assert.deepStrictEqual(sdRow.lines.map(l => l.part), ["sashes"],
    "so the tablet and the office board draw the Sashes line and no Frames line");
  assert.strictEqual(sdRow.total, 2, "and the group's total is the sashes alone");
  /* a Super door with ONLY frames on the sheet has nothing to weld, so no row */
  const onlyFrames = mkJob({ id: "R5097", blk: 4, seq: 10,
    prodsMain: [prod("super door", 3, 0, 0)] });
  assert.deepStrictEqual(W.weldSlice([onlyFrames], NAMES, () => ""), [],
    "a Super door with frames and no sashes is not a row at all");
  pass("SUPER DOOR is welded in sashes only: its frames number is dropped, and draws no line");

  /* ================= 18. what the feeder does about the 300 wrong rows ======
     The correction has to reach a list that already holds them. Nothing is
     ever deleted: the four groups that are not on the job go Active = No, and
     the two wrong counts are patched back. */
  const seq7 = { Job: "R5053", Customer: "Customer Five", Comment: "", SentToFloor: "15.09",
                 Wnd: 23, Drs: 4, Seq: 7, Section: "In production", Active: "Yes" };
  const held = [
    /* the four real groups, two of them carrying the inflated numbers */
    item(Object.assign({ Title: "R5053|POLARIS 85MM CASEMENT", Group: "POLARIS 85MM CASEMENT",
      GroupSeq: 0, Frames: 19, Sashes: 19 }, seq7), "801"),
    item(Object.assign({ Title: "R5053|POLARIS 85 TILT TURN", Group: "POLARIS 85 TILT TURN",
      GroupSeq: 1, Frames: 4, Sashes: 6 }, seq7), "802"),
    item(Object.assign({ Title: "R5053|PVC DOOR", Group: "PVC DOOR",
      GroupSeq: 2, Frames: 3, Sashes: 4 }, seq7), "803"),
    item(Object.assign({ Title: "R5053|SUPER DOOR", Group: "SUPER DOOR",
      GroupSeq: 3, Frames: 4, Sashes: 2 }, seq7), "804"),
    /* and the four the merge invented */
    item(Object.assign({ Title: "R5053|4000 CASEMENT", Group: "4000 CASEMENT",
      GroupSeq: 5, Frames: 6, Sashes: 0 }, seq7), "805"),
    item(Object.assign({ Title: "R5053|4000 TILT TURN", Group: "4000 TILT TURN",
      GroupSeq: 6, Frames: 6, Sashes: 0 }, seq7), "806"),
    item(Object.assign({ Title: "R5053|SIDELIGHTS", Group: "SIDELIGHTS",
      GroupSeq: 7, Frames: 2, Sashes: 0 }, seq7), "807"),
    item(Object.assign({ Title: "R5053|PVC SMART", Group: "PVC SMART",
      GroupSeq: 9, Frames: 2, Sashes: 0 }, seq7), "808")
  ];
  const fix = ST.feedPlan(fixed, held, { at: at, by: "the office", def: W.WELD });
  assert.strictEqual(fix.adds.length, 0, "nothing is created: all four real rows are already there");
  const byId = {};
  fix.patches.forEach(p => { byId[p.id] = p.fields; });
  const wentOff = fix.patches.filter(p => p.fields.Active === "No").map(p => p.id).sort();
  assert.deepStrictEqual(wentOff, ["805", "806", "807", "808"],
    "the four groups that are not on the job go Active = No");
  wentOff.forEach(id => assert.deepStrictEqual(Object.keys(byId[id]).sort(), ["Active", "FedAt", "FedBy"],
    "and nothing else about them is touched: " + JSON.stringify(byId[id])));
  assert.strictEqual(byId["801"].Sashes, 16, "the inflated sashes count is patched back");
  assert.strictEqual(byId["804"].Frames, 0, "and the Super door's frames go to nought");
  assert.ok(!byId["802"] && !byId["803"], "the two rows that were always right are not written at all");
  assert.strictEqual(fix.unchanged, 2);
  assert.strictEqual(JSON.stringify(fix).indexOf("delete"), -1, "and nothing anywhere is deleted");
  assert.ok(!JSON.stringify(fix.patches).match(/FramesDone|SashesDone|DoneAt|DoneBy/),
    "no counter and no stamp is touched by the correction: the floor's work stays where it is");
  pass("the correction: four rows go inactive, two counts are patched back, nothing is created or deleted");
  }

  /* ---- the standing gate, over every request this whole run made ---- */
  ALLREQ.forEach(r => {
    assert.strictEqual(r.path.indexOf("/workbook"), -1, "a request reached the workbook: " + r.path);
    assert.strictEqual(r.path.indexOf("/drive"), -1, "a request reached the file: " + r.path);
    assert.strictEqual(r.path.indexOf("/worksheets"), -1, "a request reached a worksheet: " + r.path);
    assert.notStrictEqual(r.method, "DELETE", "something issued a DELETE: " + r.path);
  });
  assert.ok(ALLREQ.length > 10, "and there were real requests to check, not none");
  pass("over " + ALLREQ.length + " requests: not one workbook path, not one worksheet, not one DELETE");

  console.log("\n" + n + " checks passed");
})().catch(e => { console.error("\nFAIL " + (e && e.stack || e)); process.exit(1); });
