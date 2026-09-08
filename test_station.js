/* Offline test of the glass station, second iteration: the slice the master
   dashboard feeds into the "Glass station" SharePoint list, the plan of writes
   that gets it there, who on the floor may move which counter, the log line
   every counter write leaves behind, the delta poll that puts a floor tap on
   the office screen inside ten seconds, and the board diff that lets a tablet
   redraw six times a minute without losing an open card.

   The whole point of the feature is that the floor gets a screen WITHOUT
   getting the workbook, and that neither side can write the other's columns.
   So the sharpest assertions in here are about the request log and the bodies
   that go out on it:

     · over the entire run, not one request touches the workbook - no download,
       no /workbook path, no range, no worksheet, no Dashboard Log;
     · nothing the feeder sends ever carries one of the floor's own columns,
       and the feeder never issues a DELETE;
     · nothing the station sends ever carries Job, Customer, GlassType, Total,
       Seq, Active, FedAt or FedBy;
     · nothing anywhere sends a DELETE to the Station log;
     · no phone number, eircode, county, price, comment, product or product
       count is in the slice, the lists, or any of the three screens.

   Graph is a fake fetch() over two in-memory SharePoint sites; nothing leaves
   the box, every address in here is example.test and every person is made up.
   Run: node test_station.js                                                 */
const fs = require("fs"), vm = require("vm"), assert = require("assert");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null),
                        setItem: (k, v) => { mem[k] = String(v); },
                        removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, innerWidth: 1280, innerHeight: 800,
                  addEventListener() {}, removeEventListener() {} };
global.performance = { now: () => Date.now() };

/* Elements the page creates and destroys answer null until they exist, the way
   a browser does - otherwise "is the log window open?" is always yes. */
const REG = {};
function stubEl(tag, id) {
  let html = "";
  const e = {
    tag: tag || "div", id: id || "", style: { setProperty() {} }, dataset: {}, attrs: {}, kids: [],
    textContent: "", value: "", disabled: false, hidden: false, className: "", title: "",
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); if (c.id) REG[c.id] = c; return c; },
    remove() { const drop = x => { if (x.id && REG[x.id] === x) delete REG[x.id]; x.kids.forEach(drop); }; drop(e); },
    contains(x) { return x === e || e.kids.some(k => k.contains && k.contains(x)); },
    setAttribute(k, v) { e.attrs[k] = String(v); }, getAttribute(k) { return e.attrs[k] || null; },
    on: {},
    addEventListener(t, fn) { (e.on[t] = e.on[t] || []).push(fn); },
    removeEventListener(t, fn) { e.on[t] = (e.on[t] || []).filter(x => x !== fn); },
    fire(t, ev) { (e.on[t] || []).slice().forEach(fn => fn(Object.assign({ type: t, target: e }, ev || {}))); },
    focus() {}, blur() {}, setSelectionRange() {},
    getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => null, querySelectorAll: () => []
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; } });
  return e;
}
const NULLABLE = ["#fabhost", "#fabbtn", "#dhost", "#xhost", "#ahost", "#chost", "#vhost",
                  "#lhost", "#catmenu", "#movemenu", "#alertmenu"];
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
   Two sites: the one the workbook lives in (which the station account cannot
   see at all) and the small Floor stations site, which holds three lists. This
   fake serves no workbook at all: a request for one is a test failure by
   construction, because there is no route for it. */
const G = "https://graph.microsoft.com/v1.0";
const SITE = "costellowindowsie.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";
const FSITE = "costellowindowsie.sharepoint.com,aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee,ffffffff-0000-1111-2222-333333333333";
const GLASS_ID = "list-glass-station";
const PEOPLE_ID = "list-station-people";
const LOG_ID = "list-station-log";
const FLISTS_PATH = "/sites/" + FSITE + "/lists";
const OWNLISTS_PATH = "/sites/" + SITE + "/lists";
const HOST_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/FloorStations";
/* the interim arrangement: the same three lists, in the workbook's own site,
   looked up by path with the $select that marks it as the station's lookup
   rather than findFile()'s */
const SITE_BY_PATH = "costellowindowsie.sharepoint.com:/sites/ProductionProgress";
const OWN_LOOKUP = "/sites/" + SITE_BY_PATH + "?$select=id,displayName";

let SITE_EXISTS = true;                 // has the manager made the Floor stations site yet?
let FLISTS = [{ id: "list-site-assets", displayName: "Site Assets" },
              { id: GLASS_ID, displayName: "Glass station" },
              { id: PEOPLE_ID, displayName: "Station people" },
              { id: LOG_ID, displayName: "Station log" }];
let ITEMS = [];                         // the items of the Glass station list
let PEOPLEITEMS = [];                   // the items of the Station people list
let LOGITEMS = [];                      // the items of the Station log list
let PAGE = 999;                         // how many items one page of a plain read holds
let NEXTID = 100;
let FAIL_ONCE = 0;                      // n writes to refuse with a 403
let FAIL_503 = 0;                       // n writes to answer 503 (call() retries those)
let FAIL_LOG = 0;                       // n Station log POSTs to refuse
let UNIQUE_TITLE = false;               // the list's "enforce unique values" rule on Title
let SITE_403 = false;                   // the Floor stations lookup refused rather than missing
let OWN_HAS_LISTS = false;              // are the three lists in the workbook's own site?
/* the interim copies. Separate stores with separate item ids, because that is
   the whole point: an item id only means anything in the list it came from. */
let OWNITEMS = [], OWNPEOPLE = [], OWNLOG = [];
let OWN_404 = false;                    // ... and can this account even see that site?

/* the delta feed, scripted: what a token call answers next, how big a page of
   the initial enumeration is, and whether the endpoint is refusing at all */
let DELTA_FAIL = 0;                     // 0, 400 or 410
let DELTA_FAIL_LEFT = 0;                // how many delta calls to refuse with it
let DELTA_PAGE = 999;
let DELTA_SEQ = 0;
const DELTA_NEXT = {};                  // listId -> [batch, batch, ...] for token calls

const REQ = [], ALLREQ = [];
let inflight = 0, peakInflight = 0;
const ok = body => ({ status: 200, body: body });
const tick = ms => new Promise(r => setTimeout(r, ms == null ? 1 : ms));

function item(fields, id) {
  return { id: String(id == null ? NEXTID++ : id), fields: Object.assign({}, fields) };
}
function storeFor(id, base) {
  const own = base === OWNLISTS_PATH;
  return id === GLASS_ID ? (own ? OWNITEMS : ITEMS)
       : id === PEOPLE_ID ? (own ? OWNPEOPLE : PEOPLEITEMS)
       : id === LOG_ID ? (own ? OWNLOG : LOGITEMS) : null;
}

function routeDelta(listId, tail, base) {
  if (DELTA_FAIL && DELTA_FAIL_LEFT > 0) {
    DELTA_FAIL_LEFT--;
    const code = DELTA_FAIL === 410 ? "resyncChangesApplyDifferences" : "invalidRequest";
    return { status: DELTA_FAIL, body: { error: { code: code, message: "delta token no good" } } };
  }
  const store = storeFor(listId, base) || [];
  const link = () => G + base + "/" + listId + "/items/delta?token=T" + (++DELTA_SEQ);
  const tok = /[?&]token=([^&]*)/.exec(tail);
  if (tok) {
    const batch = (DELTA_NEXT[listId] || []).shift() || [];
    return ok({ value: batch, "@odata.deltaLink": link() });
  }
  const page = Number((/[?&]page=(\d+)/.exec(tail) || [])[1] || 1);
  const from = (page - 1) * DELTA_PAGE, slice = store.slice(from, from + DELTA_PAGE);
  const out = { value: slice.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) };
  if (from + DELTA_PAGE < store.length)
    out["@odata.nextLink"] = G + base + "/" + listId + "/items/delta?page=" + (page + 1);
  else out["@odata.deltaLink"] = link();
  return ok(out);
}

function routeFlist(method, path, body, base) {
  const rest = path.slice(base.length);
  if (method === "GET" && rest.indexOf("?$select=id,displayName") === 0) return ok({ value: FLISTS });
  const mi = /^\/([^/?]+)\/items(.*)$/.exec(rest);
  if (!mi) return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + path } } };
  const id = mi[1], tail = mi[2];
  const store = storeFor(id, base);
  if (!store) return { status: 404, body: { error: { code: "itemNotFound" } } };
  if (method === "GET" && tail.indexOf("/delta") === 0) return routeDelta(id, tail, base);
  if (method === "GET") {
    const m = /[?&]page=(\d+)/.exec(tail);
    const page = m ? Number(m[1]) : 1;
    const from = (page - 1) * PAGE, slice = store.slice(from, from + PAGE);
    const out = { value: slice.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) };
    if (from + PAGE < store.length)
      out["@odata.nextLink"] = G + base + "/" + id + "/items?expand=fields&page=" + (page + 1);
    return ok(out);
  }
  if (FAIL_503 && method !== "GET") { FAIL_503--; return { status: 503, body: { error: { code: "serviceNotAvailable" } } }; }
  if (FAIL_LOG && id === LOG_ID && method !== "GET") {
    FAIL_LOG--; return { status: 403, body: { error: { code: "accessDenied" } } };
  }
  if (FAIL_ONCE && method !== "GET" && id !== LOG_ID) {
    FAIL_ONCE--; return { status: 403, body: { error: { code: "accessDenied" } } };
  }
  if (method === "POST" && tail === "") {
    const title = String((body && body.fields && body.fields.Title) || "").trim().toUpperCase();
    if (UNIQUE_TITLE && id === GLASS_ID && store.some(x => String(x.fields.Title || "").trim().toUpperCase() === title))
      return { status: 400, body: { error: { code: "invalidRequest",
        message: "The value of the Title column must be unique." } } };
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
  if (path === HOST_LOOKUP) {
    if (SITE_403) return { status: 403, body: { error: { code: "accessDenied" } } };
    return SITE_EXISTS ? ok({ id: FSITE, displayName: "Floor stations" })
                       : { status: 404, body: { error: { code: "itemNotFound" } } };
  }
  if (path.indexOf(OWN_LOOKUP) === 0) {
    /* the workbook's own site, looked up by path by stationSite()'s fallback */
    return OWN_404 ? { status: 404, body: { error: { code: "itemNotFound" } } }
                   : ok({ id: SITE, displayName: "Production Progress" });
  }
  if (path.indexOf(FLISTS_PATH) === 0) return routeFlist(method, path, body, FLISTS_PATH);
  /* the workbook's own site holds the same three lists in the interim
     arrangement, and nothing at all otherwise. There is no workbook route at
     all: anything that reaches for the file gets a 599 no retry rule matches */
  if (path.indexOf(OWNLISTS_PATH) === 0)
    return OWN_HAS_LISTS ? routeFlist(method, path, body, OWNLISTS_PATH) : ok({ value: [] });
  return { status: 599, body: { error: { code: "thisTestServesNoWorkbook", message: path } } };
}

global.fetch = async (url, init) => {
  const path = String(url).replace(G, "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  const rec = { method: init.method, path: path, body: body, headers: init.headers || {} };
  REQ.push(rec); ALLREQ.push(rec);
  inflight++; if (inflight > peakInflight) peakInflight = inflight;
  await tick();                       // real overlap, so a concurrency cap is observable
  inflight--;
  const res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status,
           text: async () => (res.body === "" ? "" : JSON.stringify(res.body)),
           arrayBuffer: async () => new ArrayBuffer(0) };
};

/* ---------- load the master dashboard's own code, in the page's own order ---------- */
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
run("app.js");

const TOASTS = [];
global.toast = (m, isErr) => TOASTS.push({ m: String(m), err: !!isErr });

/* ---------- the station page, in a context of its own ----------
   glass.html is a different page from index.html: it never loads app.js, so
   the two share nothing but graph.js and station-core.js. Giving station.js
   its own vm context says exactly that, and keeps its $ / esc from colliding
   with the dashboard's. */
const stationDoc = global.document;
let BUILD_SERVED = "20260908-0951";     // what version.json answers to the tablet
let RELOADED = false;                   // did the tablet reload itself?
const stationFetch = async (url, init) => {
  if (String(url).indexOf("version.json") === 0)
    return { ok: true, status: 200, json: async () => ({ build: BUILD_SERVED }) };
  return global.fetch(url, init);
};
function newStation() {
  const sb = {
    console: console, setTimeout: setTimeout, clearTimeout: clearTimeout,
    setInterval: () => 0, clearInterval: clearInterval,
    localStorage: global.localStorage, document: stationDoc,
    window: { location: { origin: "http://localhost" } },
    location: { reload: () => { RELOADED = true; } }, fetch: stationFetch,
    CW: CW, ST: ST, confirm: () => false, prompt: () => null, alert: () => {}
  };
  sb.globalThis = sb;
  vm.createContext(sb);
  vm.runInContext(src("station.js"), sb, { filename: "station.js" });
  return sb;
}
const sandbox = newStation();
const S = code => vm.runInContext(code, sandbox);

/* ---------- helpers ---------- */
const reset = () => { REQ.length = 0; TOASTS.length = 0; peakInflight = 0; };
const forget = () => { delete mem.cw_listids; CW._resetListIds(); CW._setStationSite(null); };
const paths = () => REQ.map(r => r.method + " " + r.path);
const writes = () => REQ.filter(r => r.method !== "GET");
const settle = ms => new Promise(r => setTimeout(r, ms == null ? 40 : ms));
const A = vm.runInThisContext.bind(vm);          // reach app.js' own let-bound state
const boardHtml = () => S("Object.keys(NODES).map(k => NODES[k].innerHTML).join('')");
/* the log window is a frame built once plus two boxes repainted on every
   filter and every poll, so a test reads the boxes rather than the frame */
const logFrame = () => (REG["lhost"] ? REG["lhost"].innerHTML : "");
const logBody = () => EL["#lgbody"].innerHTML;
const logCounts = () => EL["#lgcounts"].innerHTML;
const logHead = () => EL["#lgcount"].textContent;

const mkJob = o => Object.assign({
  id: "R0001", cust: "Customer One", area: "Cork", eir: "", off: "", colour: "", ph3: "",
  wnd: 0, drs: 0, glass: {}, prods: [], notes: [], sheets: ["Production"], src: {},
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0,
  cp: { win: "", drs: "", glass: {}, prod: {} }
}, o || {});
const NAMES = ["Can sell as second hand", "Ready, customer won't take",
               "Collect & supply only", "Ready to fit", "In production"];
function useJobs(list) {
  global.__jobs = list;
  A("ALL = __jobs; BLOCKNAMES = " + JSON.stringify(NAMES) + "; ALL.blockNames = BLOCKNAMES;");
}
/* the feeder is meant to run with the list permission already granted; this
   makes that answer yes or no without a sign-in library in the room */
const consent = yes => { CW.hasListConsent = async () => !!yes; };

/* three made-up people, the way the owner keeps them in SharePoint */
const person = (name, stages, pin, active, station) =>
  item({ Title: name, Station: station || "Glass", Stages: stages,
         PIN: pin == null ? "" : pin, Active: active == null ? "Yes" : active });

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ================= 1. the shape of the thing ================= */
  assert.deepStrictEqual(ST.STAGES.map(s => s[0]), ["cut", "hotmelt", "glazed"]);
  assert.deepStrictEqual(ST.STAGES.map(s => s[1]), ["Glass cut", "Hotmelt", "Glazing"]);
  assert.deepStrictEqual(ST.STAGE_FIELD, { cut: "Cut", hotmelt: "Hotmelt", glazed: "Glazed" });
  assert.deepStrictEqual(ST.STAGE_BY, { cut: "CutBy", hotmelt: "HotmeltBy", glazed: "GlazedBy" });
  assert.deepStrictEqual(ST.STAGE_AT, { cut: "CutAt", hotmelt: "HotmeltAt", glazed: "GlazedAt" });
  assert.strictEqual(ST.stageLabel("hotmelt"), "Hotmelt");
  pass("the three stages are Glass cut, Hotmelt and Glazing, each with its own counter and By/At pair");

  assert.strictEqual(ST.STATION_LIST, "Glass station");
  assert.strictEqual(ST.PEOPLE_LIST, "Station people");
  assert.strictEqual(ST.LOG_LIST, "Station log");
  assert.strictEqual(ST.STATION_NAME, "Glass");
  assert.strictEqual(ST.STATION_SITE, CW.STATION_SITE_NAME, "graph.js and station-core.js name one site");
  pass("three lists in one Floor stations site, and both files name that site the same way");

  /* the stage rename really is a rename: nothing new says Toughened anywhere */
  const NEWSRC = src("station-core.js") + src("station.js") + src("glass.html");
  assert.ok(!/[Tt]oughen/.test(NEWSRC), "no new file may still say Toughened");
  assert.ok(!/TOUGH_TYPES/.test(NEWSRC + src("app.js")), "and the type list it needed is gone with it");
  assert.strictEqual(ST.TOUGH_TYPES, undefined, "station-core.js does not export it either");
  assert.strictEqual(ST.STAGE_FIELD.tough, undefined);
  pass("Toughening is gone: no constant, no column, no word left in the new code");

  assert.deepStrictEqual(ST.FLOOR_FIELDS,
    ["Cut", "Hotmelt", "Glazed", "CutBy", "CutAt", "HotmeltBy", "HotmeltAt",
     "GlazedBy", "GlazedAt", "DoneBy", "DoneAt"]);
  assert.deepStrictEqual(ST.STATION_FIELDS,
    ["Title", "Job", "Customer", "GlassType", "Total", "Seq", "Active", "FedAt", "FedBy"]
      .concat(ST.FLOOR_FIELDS));
  assert.deepStrictEqual(ST.LOG_FIELDS,
    ["Title", "Station", "GlassType", "Stage", "From", "To", "Who", "At"]);
  assert.deepStrictEqual(ST.PEOPLE_FIELDS, ["Title", "Station", "Stages", "PIN", "Active"]);
  pass("the floor owns exactly eleven columns, and the three lists have the columns the brief names");

  ["Phone", "Eircode", "County", "Area", "Price", "Comment", "Notes", "Product", "Windows", "Doors"]
    .forEach(bad => ST.STATION_FIELDS.concat(ST.LOG_FIELDS, ST.PEOPLE_FIELDS).forEach(f =>
      assert.ok(f.toLowerCase().indexOf(bad.toLowerCase()) < 0, "no list may carry " + bad)));
  pass("no column of any of the three lists is a phone, eircode, county, price, comment or product");

  assert.strictEqual(ST.PERSON_LOCK_MS, 600000, "ten minutes, one constant");
  assert.strictEqual(ST.REFRESH_MS, 10000, "and ten seconds, likewise");
  pass("the lock and the poll are one constant each, at the numbers the owner asked for");

  /* ================= 2. the slice ================= */
  const jobs = [
    mkJob({ id: "r5303", cust: "Customer One", glass: { tg: 6, dg: 2, arch: 0 }, blk: 4, seq: 0 }),
    mkJob({ id: "R5304", cust: "Customer Two", glass: { tuff: 3 }, blk: 4, seq: 1 }),
    mkJob({ id: "R5305", cust: "Customer Three", glass: { dg: 4 }, blk: 3, seq: 2 }),        // Ready to fit
    mkJob({ id: "R5306", cust: "Customer Four", glass: {}, blk: 4, seq: 3 }),                 // no glass at all
    mkJob({ id: "R5307", cust: "Customer Five", glass: { tg: 1 }, blk: -1, cat: "past", seq: 4 })
  ];
  let slice = ST.glassSlice(jobs, NAMES);
  assert.deepStrictEqual(slice.map(r => r.title), ["R5303|DG", "R5303|TG", "R5304|TUFF"],
    "one row per job AND type, in sheet order then title, job numbers upper-cased");
  assert.strictEqual(slice.filter(r => r.job === "R5306").length, 0, "a job with no glass produces nothing");
  assert.strictEqual(slice.filter(r => r.type === "ARCH").length, 0, "and a type with a zero count is not a row");
  pass("the slice is one row per job and glass type, and nothing for a job with no glass");

  assert.ok(slice.every(r => r.active === true), "every row in the slice is a job on the floor");
  assert.strictEqual(slice.filter(r => r.job === "R5305").length, 0, "Ready to fit is not fed at all");
  assert.strictEqual(slice.filter(r => r.job === "R5307").length, 0, "and neither is a job off the sheet");
  assert.strictEqual(ST.glassSlice([mkJob({ id: "R1", glass: { tg: 2 }, blk: 2 })], NAMES).length, 0,
    "Collect & supply only is not the floor's work either");
  assert.strictEqual(ST.inProduction(mkJob({ blk: 4 }), ["a", "b", "c", "d", "IN PRODUCTION (2)"]), true,
    "the section name is matched on its prefix, ignoring case");
  assert.strictEqual(ST.inProduction(mkJob({ blk: 9 }), NAMES), false, "a job in no section is not in production");
  pass("only jobs In production are fed: no row is ever created for a job the floor will never see");

  const longName = "Customer" + "x".repeat(120);
  const sl2 = ST.glassSlice([mkJob({ id: "R1", cust: longName, glass: { tg: 1 } })], NAMES);
  assert.strictEqual(sl2[0].customer.length, 70, "a customer name is cut to 70 characters");
  ["eir", "area", "ph3", "off", "notes", "prods", "wnd", "drs", "colour", "cp", "dates"]
    .forEach(k => assert.ok(!(k in sl2[0]), "the slice must not carry " + k));
  assert.deepStrictEqual(Object.keys(sl2[0]).sort(),
    ["active", "customer", "job", "seq", "title", "total", "type"]);
  pass("a slice row is seven fields: no phone, eircode, county, product, comment or date among them");

  /* ================= 3. the plan ================= */
  const AT = "2026-09-08T09:00:00.000Z", BY = "office@example.test";
  let plan = ST.feedPlan(slice, [], { at: AT, by: BY });
  assert.strictEqual(plan.adds.length, 3);
  assert.strictEqual(plan.patches.length, 0);
  assert.deepStrictEqual(plan.adds[0], { Title: "R5303|DG", Job: "R5303", Customer: "Customer One",
    GlassType: "DG", Total: 2, Seq: 0, Active: "Yes", FedAt: AT, FedBy: BY });
  assert.deepStrictEqual(plan.adds.map(a => a.Title), ["R5303|DG", "R5303|TG", "R5304|TUFF"],
    "Seq then Title, always");
  pass("an empty list is filled by adds, in Seq then Title order");

  ITEMS = plan.adds.map((f, i) => item(f, 200 + i));
  ITEMS[1].fields.Cut = 6; ITEMS[1].fields.Hotmelt = 4; ITEMS[1].fields.Glazed = 1;
  ITEMS[1].fields.CutBy = "Person A"; ITEMS[1].fields.CutAt = "2026-09-08T10:00:00.000Z";
  ITEMS[1].fields.DoneBy = "Person A"; ITEMS[1].fields.DoneAt = "2026-09-08T10:00:00.000Z";
  plan = ST.feedPlan(slice, ITEMS, { at: AT, by: BY });
  assert.deepStrictEqual(plan, { adds: [], patches: [], unchanged: 3 },
    "nothing changed on the sheet, so nothing is written");
  pass("a list that already agrees with the sheet gets no writes at all");

  const jobs2 = jobs.slice();
  jobs2[0] = mkJob({ id: "R5303", cust: "Customer One Ltd", glass: { tg: 8, dg: 2 }, blk: 4, seq: 0 });
  slice = ST.glassSlice(jobs2, NAMES);
  plan = ST.feedPlan(slice, ITEMS, { at: "2026-09-08T11:00:00.000Z", by: BY });
  assert.strictEqual(plan.adds.length, 0);
  assert.strictEqual(plan.patches.length, 2, "one for the name, one for the name and the total");
  assert.deepStrictEqual(plan.patches.find(p => p.id === "201").fields,
    { Customer: "Customer One Ltd", Total: 8, FedAt: "2026-09-08T11:00:00.000Z", FedBy: BY },
    "only the fields that actually differ, plus FedAt/FedBy");
  assert.strictEqual(plan.unchanged, 1);
  pass("a patch carries only the feeder fields that changed, and says when and by whom");

  const allPlanFields = p => p.adds.concat(p.patches.map(x => x.fields));
  const noFloor = p => allPlanFields(p).every(f => ST.FLOOR_FIELDS.every(k => !(k in f)));
  assert.ok(noFloor(plan), "a patch never names one of the floor's columns");
  assert.strictEqual(ITEMS[1].fields.Cut, 6, "and the counters that were there are still there");
  assert.strictEqual(ITEMS[1].fields.CutBy, "Person A", "and so is who put them there");
  pass("no plan the feeder can make contains a counter, a By, an At, or the last-touch pair");

  const jobs3 = jobs2.filter(j => j.id !== "R5304");
  plan = ST.feedPlan(ST.glassSlice(jobs3, NAMES), ITEMS, { at: AT, by: BY });
  assert.deepStrictEqual(plan.patches.find(p => p.id === "202").fields,
    { Active: "No", FedAt: AT, FedBy: BY });
  assert.ok(noFloor(plan), "and even then, not a counter in sight");
  assert.ok(!("delete" in plan) && !("deletes" in plan) && !("removes" in plan),
    "there is no such thing as a delete in a feed plan");
  pass("a job that leaves production is marked Active = No, and its item is never removed");

  ITEMS.find(x => x.id === "202").fields.Active = "No";
  plan = ST.feedPlan(ST.glassSlice(jobs3, NAMES), ITEMS, { at: AT, by: BY });
  assert.strictEqual(plan.patches.filter(p => p.id === "202").length, 0);
  pass("an item already marked inactive is not written again on the next run");

  const ragged = [{ id: "300", fields: { Title: "R9001|TG" } }, { id: "301" }, null,
                  { id: "302", fields: { Title: "" } }];
  plan = ST.feedPlan([{ title: "R9001|TG", job: "R9001", customer: "Customer Seven", type: "TG",
                        total: 2, seq: 7, active: true }], ragged, { at: AT, by: BY });
  assert.strictEqual(plan.adds.length, 0, "the ragged item is still the item for that Title");
  assert.deepStrictEqual(plan.patches[0].fields,
    { Job: "R9001", Customer: "Customer Seven", GlassType: "TG", Total: 2, Seq: 7, Active: "Yes", FedAt: AT, FedBy: BY });
  pass("an item missing every column but Title is filled in rather than crashing anything");

  const twins = [{ id: "410", fields: { Title: "R7|TG", Job: "R7", Customer: "Old", GlassType: "TG",
                                        Total: 1, Seq: 1, Active: "Yes" } },
                 { id: "409", fields: { Title: "R7|TG", Job: "R7", Customer: "Older", GlassType: "TG",
                                        Total: 1, Seq: 1, Active: "Yes" } }];
  plan = ST.feedPlan([{ title: "R7|TG", job: "R7", customer: "New", type: "TG",
                        total: 1, seq: 1, active: true }], twins, { at: AT, by: BY });
  assert.strictEqual(plan.patches.length, 1);
  assert.strictEqual(plan.patches[0].id, "409", "the oldest item wins, exactly as the phases list does");
  pass("two items carrying one Title: the older is maintained, the other left alone, neither deleted");

  /* ================= 4. the hash ================= */
  const h1 = ST.sliceHash(ST.glassSlice(jobs, NAMES));
  assert.strictEqual(h1, ST.sliceHash(ST.glassSlice(jobs.slice().reverse(), NAMES)),
    "the same jobs in another order are the same slice");
  assert.notStrictEqual(h1, ST.sliceHash(ST.glassSlice(jobs2, NAMES)), "a changed count changes it");
  assert.notStrictEqual(h1, ST.sliceHash(ST.glassSlice(jobs.map(j => Object.assign({}, j, { seq: j.seq + 10 })), NAMES)),
    "and so does a changed order");
  assert.strictEqual(ST.sliceHash([]), "0-811c9dc5");
  pass("the slice hash is stable, order-independent, and moves when anything the floor sees moves");

  /* ================= 5. the board ================= */
  const bItems = [
    item({ Title: "R5303|TG", Job: "R5303", Customer: "Customer One", GlassType: "TG", Total: 6, Seq: 2,
           Active: "Yes", Cut: 6, Hotmelt: 2, Glazed: 0, FedAt: "2026-09-08T09:00:00.000Z",
           CutBy: "Person A", CutAt: "2026-09-08T12:00:00.000Z",
           HotmeltBy: "Person B", HotmeltAt: "2026-09-08T13:00:00.000Z" }, "500"),
    item({ Title: "R5303|ARCH", Job: "R5303", Customer: "Customer One", GlassType: "ARCH", Total: 2, Seq: 2,
           Active: "Yes", Cut: 1, Glazed: 0, FedAt: "2026-09-08T09:05:00.000Z",
           CutBy: "Person C", CutAt: "2026-09-08T14:00:00.000Z" }, "501"),
    item({ Title: "R5301|DG", Job: "R5301", Customer: "Customer Six", GlassType: "DG", Total: 4, Seq: 1,
           Active: "Yes", Cut: 4, Hotmelt: 4, Glazed: 4 }, "502"),
    item({ Title: "R5299|TG", Job: "R5299", Customer: "Customer Eight", GlassType: "TG", Total: 3, Seq: 0,
           Active: "No", Cut: 3 }, "503")
  ];
  let board = ST.jobBoard(bItems);
  assert.deepStrictEqual(board.map(g => g.job), ["R5301", "R5303"], "Seq order, and inactive is off the board");
  const r5303 = board[1];
  assert.deepStrictEqual(r5303.bars.cut.done + "/" + r5303.bars.cut.total, "7/8");
  assert.deepStrictEqual(r5303.bars.hotmelt.done + "/" + r5303.bars.hotmelt.total, "2/8",
    "hotmelt applies to every type, so ARCH is in that bar's total");
  assert.deepStrictEqual(r5303.bars.glazed.done + "/" + r5303.bars.glazed.total, "0/8");
  assert.deepStrictEqual(r5303.rows.map(r => r.type), ["ARCH", "TG"]);
  assert.strictEqual(r5303.fedAt, "2026-09-08T09:05:00.000Z", "the newest feed stamp on the job");
  pass("the board groups by job in Seq order, and every stage counts every glass type");

  assert.strictEqual(r5303.bars.cut.by, "Person C", "the latest cut on the job, whichever type it was");
  assert.strictEqual(r5303.bars.cut.at, "2026-09-08T14:00:00.000Z");
  assert.strictEqual(r5303.bars.hotmelt.by, "Person B");
  assert.strictEqual(r5303.bars.glazed.by, "", "nothing glazed yet, so nobody to name");
  assert.strictEqual(r5303.rows[1].by.cut, "Person A", "and the row keeps its own");
  pass("each bar carries who last moved it and when, taken from the newest At on the job");

  assert.strictEqual(board[0].finished, true, "everything full");
  assert.strictEqual(r5303.finished, false);
  const empty = ST.jobBoard([item({ Title: "R1|ARCH", Job: "R1", GlassType: "ARCH", Total: 0,
                                    Seq: 1, Active: "Yes" }, "600")]);
  assert.strictEqual(empty[0].finished, true, "an empty bar is not something anyone can finish");
  pass("finished means every bar full, and a bar with no units in it never holds a job back");

  const shrunk = ST.jobBoard([item({ Title: "R2|TG", Job: "R2", GlassType: "TG", Total: 4, Seq: 1,
                                     Active: "Yes", Cut: 9, Hotmelt: 6, Glazed: -3 }, "700")]);
  assert.strictEqual(shrunk[0].rows[0].cut, 4);
  assert.strictEqual(shrunk[0].rows[0].hotmelt, 4);
  assert.strictEqual(shrunk[0].rows[0].glazed, 0);
  assert.deepStrictEqual(shrunk[0].bars.cut.done + "/" + shrunk[0].bars.cut.total, "4/4");
  pass("a count above the total, or below zero, is clamped for display and never written back");

  const dupes = ST.jobBoard([item({ Title: "R3|TG", Job: "R3", GlassType: "TG", Total: 2, Seq: 1,
                                    Active: "Yes", Cut: 2 }, "811"),
                             item({ Title: "R3|TG", Job: "R3", GlassType: "TG", Total: 2, Seq: 1,
                                    Active: "Yes", Cut: 0 }, "810")]);
  assert.strictEqual(dupes[0].rows.length, 1, "one row, not two");
  assert.strictEqual(dupes[0].rows[0].id, "810", "the oldest item is the one that is shown");
  pass("two items for one Title show as one row on the board, the oldest one");

  assert.deepStrictEqual(ST.jobBoard([{ id: "900" }, null, item({ Active: "Yes" }, "901"),
                                      item({ Title: "R4|TG", Active: "yes", Job: "R4",
                                             GlassType: "TG", Total: 1 }, "902")]).map(g => g.job), ["R4"]);
  assert.deepStrictEqual(ST.jobBoard(null), []);
  pass("an item missing its fields does not crash the board, and a null list is an empty board");

  const finishedRec = ST.jobRecord(bItems, "R5299");
  assert.strictEqual(finishedRec.active, false);
  assert.strictEqual(finishedRec.bars.cut.done, 3);
  assert.strictEqual(ST.jobRecord(bItems, "R5303").active, true);
  assert.strictEqual(ST.jobRecord(bItems, "R0000"), null);
  pass("jobRecord finds a job whatever its Active is, so a finished job keeps its record");

  /* ================= 6. the tap ================= */
  const trow = { id: "1", type: "TG", total: 6, cut: 3, hotmelt: 0, glazed: 0 };
  assert.strictEqual(ST.applyTap(trow, "cut", 1), 4);
  assert.strictEqual(ST.applyTap(trow, "cut", -1), 2);
  assert.strictEqual(ST.applyTap(trow, "cut", "all"), 6);
  assert.strictEqual(ST.applyTap(trow, "cut", "none"), 0);
  assert.strictEqual(ST.applyTap({ total: 6, cut: 6 }, "cut", 1), 6, "never above the total");
  assert.strictEqual(ST.applyTap({ total: 6, cut: 0 }, "cut", -1), 0, "never below zero");
  assert.strictEqual(ST.applyTap(trow, "tough", 1), null, "and there is no such stage any more");
  assert.strictEqual(ST.applyTap(trow, "cut", "wat"), 3, "a delta that is not a number changes nothing");
  pass("a tap only ever clamps between 0 and the total, and only for a stage that exists");

  assert.strictEqual(ST.applyTap({ total: 4, glazed: 0 }, "glazed", "all"), 4);
  assert.strictEqual(ST.applyTap({ total: 4, hotmelt: 0 }, "hotmelt", 1), 1,
    "glazing before hotmelt is recorded, not refused: the counters are a record, not a workflow");
  pass("no order is enforced between the three stages");

  /* ================= 7. who may record what ================= */
  const pItems = [person("Person A", "cut"), person("Person B", "hotmelt,cut", "1234"),
                  person("Person C", "", "9999"), person("Person D", "glazed", "", "No"),
                  person("Person E", "cut", "", "Yes", "Frames"),
                  person("", "cut"), person("Person F", " Glazed ; HOTMELT , cut , cut ")];
  const ppl = ST.stationPeople(pItems, "Glass");
  assert.deepStrictEqual(ppl.map(p => p.name), ["Person A", "Person B", "Person C", "Person F"],
    "only Active Yes, only this station, only rows with a name - in name order");
  assert.deepStrictEqual(ppl[0].stages, ["cut"]);
  assert.deepStrictEqual(ppl[1].stages, ["hotmelt", "cut"], "one person, two stages");
  assert.deepStrictEqual(ppl[2].stages, [], "and one with none at all");
  assert.deepStrictEqual(ppl[3].stages, ["glazed", "hotmelt", "cut"],
    "spacing, case, semicolons and a repeat are all read forgivingly");
  pass("the people list is read per station, active only, with one or more stages each");

  assert.strictEqual(ST.canStage(ppl[0], "cut"), true);
  assert.strictEqual(ST.canStage(ppl[0], "hotmelt"), false);
  assert.strictEqual(ST.canStage(ppl[0], "CUT"), true, "the stage key is matched case-insensitively");
  assert.strictEqual(ST.canStage(ppl[1], "hotmelt"), true);
  assert.strictEqual(ST.canStage(ppl[1], "glazed"), false);
  assert.strictEqual(ST.canStage(ppl[2], "cut"), false, "a person with no stages may move nothing");
  assert.strictEqual(ST.canStage(null, "cut"), false);
  pass("a person may move the stages they hold and no others, whether they hold one, two or none");

  assert.strictEqual(ST.pinOk(ppl[0], ""), true, "an empty PIN column asks for nothing");
  assert.strictEqual(ST.pinOk(ppl[0], "0000"), true, "and nothing typed at it can be wrong");
  assert.strictEqual(ST.pinOk(ppl[1], "1234"), true);
  assert.strictEqual(ST.pinOk(ppl[1], "1235"), false);
  assert.strictEqual(ST.pinOk(ppl[1], ""), false);
  assert.strictEqual(ST.pinOk(ppl[1], " 1234 "), true, "trimmed, because a pad can add nothing else");
  pass("a PIN is asked for only where the column is filled in, and then it must match exactly");

  const T0 = 1757325600000;
  assert.strictEqual(ST.personExpired(T0, T0 + 599999, 600000), false, "under ten minutes: still them");
  assert.strictEqual(ST.personExpired(T0, T0 + 600000, 600000), true, "ten minutes: locked");
  assert.strictEqual(ST.personExpired(T0, T0 + 900000, 600000), true);
  assert.strictEqual(ST.personExpired(0, T0, 600000), true, "and nobody chosen is nobody");
  assert.strictEqual(ST.personExpired(null, T0), true);
  pass("the chosen name locks itself after ten quiet minutes, and a missing one is already locked");

  /* ================= 8. the writes one tap makes ================= */
  const tf = ST.tapFields("cut", 5, "Person A", "2026-09-08T15:00:00.000Z");
  assert.deepStrictEqual(Object.keys(tf).sort(), ["Cut", "CutAt", "CutBy", "DoneAt", "DoneBy"]);
  assert.strictEqual(tf.Cut, 5);
  assert.strictEqual(tf.CutBy, "Person A");
  assert.strictEqual(tf.CutAt, "2026-09-08T15:00:00.000Z");
  assert.strictEqual(tf.DoneBy, "Person A");
  assert.strictEqual(ST.tapFields("hotmelt", 2, "x", "y").Hotmelt, 2);
  assert.strictEqual(ST.tapFields("tough", 2, "x", "y"), null, "there is no such stage to write");
  pass("a stage's PATCH is that stage's counter, that stage's By/At, and the last-touch pair - five keys");

  const dirty = { Cut: 3, Hotmelt: "4", CutBy: "Person A", CutAt: "2026-09-08T15:00:00.000Z",
                  Job: "R1", Customer: "Someone", Total: 99, Active: "No", Title: "R1|TG",
                  FedBy: "x", Glazed: NaN, GlazedBy: "" };
  assert.deepStrictEqual(ST.floorOnly(dirty),
    { Cut: 3, CutBy: "Person A", CutAt: "2026-09-08T15:00:00.000Z" },
    "a job fact, a counter that is not a number and an empty By are all dropped");
  assert.deepStrictEqual(ST.floorOnly(null), {});
  pass("the floor's whitelist drops every job fact, however it got into the bag");

  const lf = ST.logFields({ job: "r5303", type: "tg", stage: "Cut", from: 5, to: 8,
                            who: "Person A", at: "2026-09-08T15:00:00.000Z" });
  assert.deepStrictEqual(lf, { Title: "R5303", Station: "Glass", GlassType: "TG", Stage: "cut",
    From: 5, To: 8, Who: "Person A", At: "2026-09-08T15:00:00.000Z" });
  assert.strictEqual(ST.logFields({ job: "R1" }).Station, "Glass", "the station names itself");
  assert.strictEqual(ST.logFields({ job: "R1", from: "3", to: "4" }).From, 3, "numbers are numbers");
  pass("a log line is job, station, type, stage, From, To, who and when - and nothing else");

  /* ================= 9. reading the log back ================= */
  const lItems = [
    item({ Title: "R5303", Station: "Glass", GlassType: "TG", Stage: "cut", From: 0, To: 5,
           Who: "Person A", At: "2026-09-08T09:00:00.000Z" }, "1"),
    item({ Title: "R5303", Station: "Glass", GlassType: "TG", Stage: "cut", From: 5, To: 8,
           Who: "Person A", At: "2026-09-08T11:00:00.000Z" }, "2"),
    item({ Title: "R5301", Station: "Glass", GlassType: "DG", Stage: "hotmelt", From: 0, To: 4,
           Who: "Person B", At: "2026-09-07T16:30:00.000Z" }, "3"),
    item({ Title: "R5303", Station: "Frames", GlassType: "TG", Stage: "cut", From: 0, To: 2,
           Who: "Person Z", At: "2026-09-08T12:00:00.000Z" }, "4"),
    item({ Title: "", Station: "Glass", Stage: "cut", At: "2026-09-08T12:00:00.000Z" }, "5"),
    item({ Title: "R5303", GlassType: "TG", Stage: "glazed", From: 0, To: 1,
           Who: "Person C", At: "2026-09-08T10:00:00.000Z" }, "6")
  ];
  let lrows = ST.logRows(lItems, "Glass");
  assert.deepStrictEqual(lrows.map(r => r.id), ["2", "6", "1", "3"],
    "newest first; another station's line is not ours");
  assert.deepStrictEqual(ST.logRows([
    item({ Title: "R1", Stage: "cut", At: "2026-09-08T11:00:00.000Z" }, "40"),
    item({ Title: "R1", Stage: "cut", At: "2026-09-08T11:00:00.000Z" }, "41")]).map(r => r.id),
    ["41", "40"], "two taps inside one second: the later item id is the later tap");
  assert.strictEqual(lrows[0].job, "R5303");
  assert.strictEqual(lrows[0].from + "->" + lrows[0].to, "5->8");
  assert.strictEqual(lrows[3].station, "Glass", "a line with no Station is taken as this one");
  pass("the log reads back newest first, this station only, and a line with no job is skipped");

  assert.deepStrictEqual(ST.logFilter(lrows, { who: "Person A" }).map(r => r.id), ["2", "1"]);
  assert.deepStrictEqual(ST.logFilter(lrows, { stage: "glazed" }).map(r => r.id), ["6"]);
  assert.deepStrictEqual(ST.logFilter(lrows, { job: "r5301" }).map(r => r.id), ["3"]);
  assert.deepStrictEqual(ST.logFilter(lrows, { day: "2026-09-07" }).map(r => r.id), ["3"]);
  assert.deepStrictEqual(ST.logFilter(lrows, { who: "Person A", stage: "cut", day: "2026-09-08" })
    .map(r => r.id), ["2", "1"], "the four filters are an AND");
  assert.strictEqual(ST.logFilter(lrows, {}).length, 4, "and no filter at all is everything");
  pass("the log window's four filters each narrow it, and together they narrow it further");

  /* AMENDMENT 2: the window's box is somebody typing part of a number; the
     drawer is a job saying "this is mine", and R530 is not R5303 */
  assert.strictEqual(ST.logFilter(lrows, { job: "530" }).length, 4, "a substring finds them all");
  assert.strictEqual(ST.logFilter(lrows, { job: "530", exact: true }).length, 0,
    "the same substring, exactly, is no job at all");
  assert.deepStrictEqual(ST.logFilter(lrows, { job: "R5303", exact: true }).map(r => r.id), ["2", "6", "1"]);
  assert.deepStrictEqual(ST.logFilter(lrows, { job: "r5301", exact: true }).map(r => r.id), ["3"],
    "and it is still case-insensitive, because a job number is upper-cased either way");
  pass("logFilter matches a job as a substring for the window and exactly for the drawer");

  /* AMENDMENT 13: the office only ever reads the last ninety days of it */
  assert.strictEqual(ST.LOG_DAYS, 90);
  const NOW = Date.parse("2026-09-08T12:00:00.000Z");
  assert.strictEqual(ST.logSince(90, NOW).slice(0, 10), "2026-06-10");
  assert.ok(ST.logSince(90, NOW) < "2026-09-08");
  pass("the log has a ninety-day horizon, worked out in one place");

  const counts = ST.logCounts(lrows);
  assert.deepStrictEqual(counts.people.map(c => c.key + ":" + c.units + "/" + c.lines),
    ["Person A:8/2", "Person B:4/1", "Person C:1/1"]);
  assert.deepStrictEqual(counts.stages.map(c => c.key + ":" + c.units + "/" + c.lines),
    ["cut:8/2", "hotmelt:4/1", "glazed:1/1"]);
  assert.deepStrictEqual(ST.logCounts([]), { people: [], stages: [] });
  pass("the counts at the top of the log window are units and lines, per person and per stage");

  assert.strictEqual(ST.logLast(lrows, "R5303").id, "2", "the newest line for one job");
  assert.strictEqual(ST.logLast(lrows, "r5301").id, "3");
  assert.strictEqual(ST.logLast(lrows, "R0000"), null);
  pass("the board card's one-line 'last: ...' is the newest log line for that job");

  /* ================= 10. what changed between two boards ================= */
  const bA = ST.jobBoard(bItems);
  let d = ST.boardDiff(bA, bA);
  assert.deepStrictEqual([d.added, d.changed, d.removed, d.order], [[], [], [], false],
    "the same board twice: nothing at all to redraw");
  pass("boardDiff on an unchanged board names no card, so a poll that found nothing draws nothing");

  const bump = bItems.map(x => item(Object.assign({}, x.fields), x.id));
  bump.find(x => x.id === "500").fields.Glazed = 3;
  d = ST.boardDiff(bA, ST.jobBoard(bump));
  assert.deepStrictEqual(d.changed, ["R5303"], "only the job whose counter moved");
  assert.deepStrictEqual([d.added, d.removed], [[], []]);
  assert.strictEqual(d.order, false, "and it did not move on the screen");
  pass("one counter moving names exactly one card, so five other cards are never touched");

  const plusOne = bump.concat([item({ Title: "R5400|TG", Job: "R5400", Customer: "New",
    GlassType: "TG", Total: 2, Seq: 5, Active: "Yes" }, "504")]);
  d = ST.boardDiff(ST.jobBoard(bump), ST.jobBoard(plusOne));
  assert.deepStrictEqual(d.added, ["R5400"]);
  assert.strictEqual(d.order, true, "a new card is a change of order too");
  d = ST.boardDiff(ST.jobBoard(plusOne), ST.jobBoard(bump));
  assert.deepStrictEqual(d.removed, ["R5400"]);
  pass("a job arriving on the board, and one leaving it, are each named once");

  const done = bItems.map(x => item(Object.assign({}, x.fields), x.id));
  done.find(x => x.id === "500").fields.Glazed = 6;
  done.find(x => x.id === "500").fields.Hotmelt = 6;
  done.find(x => x.id === "501").fields.Cut = 2;
  done.find(x => x.id === "501").fields.Hotmelt = 2;
  done.find(x => x.id === "501").fields.Glazed = 2;
  const bDone = ST.jobBoard(done);
  assert.strictEqual(bDone.find(g => g.job === "R5303").finished, true);
  d = ST.boardDiff(bA, bDone);
  assert.deepStrictEqual(d.changed, ["R5303"]);
  assert.strictEqual(d.order, true, "becoming finished moves the card to the other group");
  pass("a card that becomes finished is both a change and a move, so it lands in the Finished group");

  const nameOnly = bItems.map(x => item(Object.assign({}, x.fields, { FedAt: "2026-09-09T00:00:00.000Z" }), x.id));
  d = ST.boardDiff(bA, ST.jobBoard(nameOnly));
  assert.deepStrictEqual(d.changed, [], "FedAt is not on the tablet's card, so it is not a redraw");
  const custMoved = bItems.map(x => item(Object.assign({}, x.fields,
    x.fields.Job === "R5303" ? { Customer: "Customer One Ltd" } : {}), x.id));
  assert.deepStrictEqual(ST.boardDiff(bA, ST.jobBoard(custMoved)).changed, ["R5303"]);
  pass("only what a card actually draws counts as a change: a customer rename does, a feed stamp does not");

  assert.deepStrictEqual(ST.boardDiff(null, bA).added, ["R5301", "R5303"], "the first paint is all adds");
  assert.deepStrictEqual(ST.boardDiff(bA, []).removed, ["R5301", "R5303"]);
  pass("a first paint and an emptied board are both described rather than crashed on");

  /* ================= 11. merging a delta feed ================= */
  const base = [{ id: "1", fields: { Title: "A", Cut: 1 } }, { id: "2", fields: { Title: "B", Cut: 2 } }];
  let merged = ST.mergeDelta(base, [{ id: "2", fields: { Title: "B", Cut: 5 } }]);
  assert.strictEqual(merged.length, 2);
  assert.strictEqual(merged[1].fields.Cut, 5, "an updated item replaces the one that was there");
  assert.strictEqual(merged[0].fields.Cut, 1, "and the others are exactly as they were");
  pass("a delta update replaces the item it names and touches nothing else");

  merged = ST.mergeDelta(base, [{ id: "3", fields: { Title: "C" } }]);
  assert.deepStrictEqual(merged.map(x => x.id), ["1", "2", "3"], "an item nobody had is appended");
  merged = ST.mergeDelta(base, [{ id: "1", removed: true }]);
  assert.deepStrictEqual(merged.map(x => x.id), ["2"], "and a removed item is dropped");
  merged = ST.mergeDelta(base, [{ id: "9", removed: true }]);
  assert.deepStrictEqual(merged.map(x => x.id), ["1", "2"], "a removal for something we never had is harmless");
  pass("a delta feed can add and remove as well as update, and a removal of nothing changes nothing");

  merged = ST.mergeDelta(base, [{ id: "2", fields: { Cut: 3 } }, { id: "2", fields: { Cut: 7 } }]);
  assert.strictEqual(merged[1].fields.Cut, 7, "the same item twice in one feed: the last one wins");
  merged = ST.mergeDelta(base, [{ id: "2", fields: { Cut: 3 } }, { id: "2", removed: true }]);
  assert.deepStrictEqual(merged.map(x => x.id), ["1"], "changed and then deleted is deleted");
  pass("an item that appears more than once in one delta feed is taken from its last occurrence");

  assert.deepStrictEqual(ST.mergeDelta(null, null), []);
  assert.deepStrictEqual(ST.mergeDelta(base, [null, {}, { fields: {} }]).map(x => x.id), ["1", "2"]);
  pass("a malformed delta entry is skipped rather than crashing a poll six times a minute");

  /* AMENDMENT 7: a change that carries no fields must not blank the row */
  merged = ST.mergeDelta(base, [{ id: "2" }]);
  assert.strictEqual(merged.length, 2);
  assert.deepStrictEqual(merged[1].fields, { Title: "B", Cut: 2 },
    "a change with no fields bag keeps the fields that were already there");
  merged = ST.mergeDelta(base, [{ id: "2", fields: {} }]);
  assert.deepStrictEqual(merged[1].fields, { Title: "B", Cut: 2 }, "and so does an empty one");
  merged = ST.mergeDelta(base, [{ id: "4" }]);
  assert.deepStrictEqual(merged[2], { id: "4", fields: {} },
    "while an unknown item with no fields is still added, empty, rather than dropped");
  pass("a delta change with no fields keeps the row it names instead of emptying it");

  /* ================= 12. the graph list layer ================= */
  forget(); reset();
  assert.strictEqual(await CW.stationSite(), FSITE);
  assert.deepStrictEqual(paths(), ["GET " + HOST_LOOKUP], "one lookup, by the hostname the workbook uses");
  assert.strictEqual(JSON.parse(mem.cw_stationsite).id, FSITE);
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE);
  assert.strictEqual(REQ.length, 0, "and then it is remembered");
  pass("the Floor stations site is resolved once by path and cached in cw_stationsite");

  /* ---- (1) the interim arrangement: no Floor stations site yet ----
     Creating a site needs an administrator the owner has not got, so the three
     lists live in the workbook's own site for now and the station account is a
     member of it. Nothing on either screen knows or says which site it is. */
  SITE_EXISTS = false; OWN_HAS_LISTS = true; forget(); reset();
  assert.strictEqual(await CW.stationSite(), SITE, "the workbook's own site is used instead");
  assert.deepStrictEqual(paths(), ["GET " + HOST_LOOKUP, "GET " + OWN_LOOKUP],
    "looked for the Floor stations site, then resolved the workbook's own site by path");
  assert.ok(REQ.every(r => r.path.indexOf("/drive") < 0 && r.path.indexOf("/workbook") < 0),
    "and never once through findFile(): not a /drive/ or /workbook path in either lookup");
  assert.strictEqual(CW._stationSiteInfo().own, true, "it knows it is on the fallback");
  assert.strictEqual(JSON.parse(mem.cw_stationsite).id, SITE, "which is what gets remembered");
  assert.strictEqual(JSON.parse(mem.cw_stationsite).own, true, "together with which site it is");
  reset();
  assert.strictEqual(await CW.stationSite(), SITE);
  assert.strictEqual(REQ.length, 0, "and then it is remembered, like any other answer");
  pass("with no Floor stations site, the lists are read from the workbook's own site, resolved by path");

  /* every list call really does go there, and the boards read normally */
  OWNITEMS = bItems.map(x => item(Object.assign({}, x.fields), "7" + x.id)); reset();
  const ownOpts = { siteId: await CW.stationSite(), fields: ST.STATION_FIELDS };
  assert.strictEqual(await CW.listId("Glass station", ownOpts), GLASS_ID);
  const got0 = await CW.listItems("Glass station", ownOpts);
  assert.strictEqual(got0.length, 4, "and the list reads out of it");
  assert.deepStrictEqual(got0.map(x => x.id).sort(), ["7500", "7501", "7502", "7503"],
    "with that site's own item ids, which are nothing to do with the other site's");
  assert.ok(REQ.every(r => r.path.indexOf(OWNLISTS_PATH) === 0 || r.path.indexOf("/sites/") === 0));
  assert.ok(REQ.some(r => r.path.indexOf(OWNLISTS_PATH) === 0), "against the workbook's own site");
  assert.ok(REQ.every(r => r.path.indexOf("/drive") < 0 && r.path.indexOf("/workbook") < 0),
    "and still not one path near the file");
  pass("on the fallback every list call addresses the workbook's site, and none of them touches the file");

  /* neither site: still null, still held for a minute */
  OWN_404 = true; forget(); reset();
  assert.strictEqual(await CW.stationSite(), null, "nowhere to read from is null, not a throw");
  assert.strictEqual(mem.cw_stationsite, undefined, "and nothing is cached as a fact");
  reset();
  assert.strictEqual(await CW.stationSite(), null);
  assert.strictEqual(REQ.length, 0, "the miss is held for a minute, so a poll cannot hammer it");
  OWN_404 = false; SITE_EXISTS = true; OWN_HAS_LISTS = false; forget();
  pass("with neither site reachable the answer is null, held for a minute, never cached as an answer");

  /* ---- (2) the Floor stations site, once it exists, wins ---- */
  forget(); reset();
  assert.strictEqual(await CW.stationSite(), FSITE, "the separate site is preferred");
  assert.deepStrictEqual(paths(), ["GET " + HOST_LOOKUP], "and found in one lookup");
  assert.strictEqual(CW._stationSiteInfo().own, false, "with no fallback involved");
  assert.strictEqual(JSON.parse(mem.cw_stationsite).own, false);
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE);
  assert.strictEqual(REQ.length, 0, "and never looked for again: there is nothing left to look for");
  pass("when the Floor stations site is there it is used, and the fallback is never reached for");

  const opts = { siteId: FSITE, fields: ST.STATION_FIELDS };
  reset();
  assert.strictEqual(await CW.listId("Glass station", opts), GLASS_ID);
  assert.strictEqual(await CW.listId("Station people", opts), PEOPLE_ID);
  assert.strictEqual(await CW.listId("Station log", opts), LOG_ID);
  assert.deepStrictEqual(JSON.parse(mem.cw_listids)[FSITE + "|Station log"], LOG_ID,
    "a list in another site is cached under a key that says which site");
  pass("all three of the floor's lists are found in the site they were given, and cached per site");

  ITEMS = bItems.slice(); PAGE = 2; reset();
  let got = await CW.listItems("Glass station", opts);
  assert.strictEqual(got.length, 4);
  const reads = REQ.filter(r => r.method === "GET" && r.path.indexOf("/items") > 0);
  assert.strictEqual(reads.length, 2, "one request per page, following @odata.nextLink");
  assert.ok(reads[0].path.indexOf("expand=fields(select=" + ST.STATION_FIELDS.join(",") + ")") > 0,
    "asking for the station's own columns by name");
  PAGE = 999;
  pass("listItems reads every page of another site's list, asking for the station columns");

  reset();
  const made = await CW.listAdd("Glass station", { Title: "R8|TG", Job: "R8" }, opts);
  assert.strictEqual(writes().length, 1, "a plain POST: no read-before-write, no dedupe pass");
  reset();
  await CW.listPatch("Glass station", made.id, { Cut: 2 }, opts);
  assert.deepStrictEqual(paths(), ["PATCH " + FLISTS_PATH + "/" + GLASS_ID + "/items/" + made.id + "/fields"]);
  assert.ok(REQ.every(r => !r.headers["workbook-session-id"]), "no workbook session on a list call");
  ITEMS = ITEMS.filter(x => x.id !== made.id);
  pass("listAdd and listPatch are one request each, straight at the item, with no workbook session header");

  /* ---- listDelta ---- */
  ITEMS = bItems.slice(); DELTA_PAGE = 999; reset();
  let dr = await CW.listDelta("Glass station", opts);
  assert.strictEqual(dr.items.length, 4, "the first call, with no token, enumerates the list");
  assert.ok(dr.next && dr.next.indexOf("token=") > 0, "and hands back the deltaLink for next time");
  assert.ok(REQ[REQ.length - 1].path.indexOf("/items/delta?expand=fields(select=") > 0,
    "at the delta endpoint, with the station's own columns");
  assert.ok(REQ[REQ.length - 1].path.indexOf("$top=") > 0);
  pass("listDelta with no token enumerates the list and returns the token for the next call");

  DELTA_PAGE = 2; reset();
  dr = await CW.listDelta("Glass station", opts);
  assert.strictEqual(dr.items.length, 4, "every page of the enumeration is followed");
  assert.strictEqual(REQ.filter(r => r.path.indexOf("/delta") > 0).length, 2);
  assert.ok(dr.next, "and only the last page carries the deltaLink");
  DELTA_PAGE = 999;
  pass("listDelta follows @odata.nextLink to the end and takes the deltaLink off the last page");

  DELTA_NEXT[GLASS_ID] = [[{ id: "500", fields: { Title: "R5303|TG", Cut: 6, Glazed: 4 } },
                           { id: "503", "@removed": { reason: "deleted" } }]];
  reset();
  const dr2 = await CW.listDelta("Glass station", { siteId: FSITE, fields: ST.STATION_FIELDS, token: dr.next });
  assert.strictEqual(dr2.items.length, 2);
  assert.strictEqual(dr2.items[0].removed, false);
  assert.strictEqual(dr2.items[1].removed, true, "a deleted item arrives carrying @removed");
  assert.strictEqual(REQ.length, 1, "and a token call is one request, not an enumeration");
  assert.ok(REQ[0].path.indexOf("token=") > 0);
  pass("a token call answers only what moved, and says which of it was deleted");

  DELTA_FAIL = 410; DELTA_FAIL_LEFT = 1; reset();
  await assert.rejects(() => CW.listDelta("Glass station", { siteId: FSITE, fields: ST.STATION_FIELDS, token: dr.next }),
    e => CW.isDeltaRestart(e), "410 Gone with a resync code is the recognisable restart error");
  DELTA_FAIL = 400; DELTA_FAIL_LEFT = 1;
  await assert.rejects(() => CW.listDelta("Glass station", { siteId: FSITE, fields: ST.STATION_FIELDS, token: dr.next }),
    e => CW.isDeltaRestart(e), "and so is any other 4xx from the delta endpoint");
  DELTA_FAIL = 0; DELTA_FAIL_LEFT = 0;
  assert.strictEqual(CW.isDeltaRestart(new Error("Failed to fetch")), false,
    "but a passing network failure is NOT: it must not cost a full read");
  pass("a stale token (410 resync) and a refused delta both throw one recognisable error");

  assert.strictEqual(await CW.listDelta("No such list", opts), null,
    "a list that is not there answers null from delta as well as from listItems");
  pass("listDelta answers null for a list that does not exist, the same way listItems does");

  /* AMENDMENT 7: the deltaLink is not obliged to carry the $expand, and a
     token call without one answers items with no fields - which reads on the
     other end as every row going blank */
  DELTA_NEXT[GLASS_ID] = [[]];
  reset();
  const bare = G + FLISTS_PATH + "/" + GLASS_ID + "/items/delta?token=TBARE";
  assert.ok(bare.indexOf("expand=") < 0, "the token this test hands back carries no expand");
  await CW.listDelta("Glass station", { siteId: FSITE, fields: ST.STATION_FIELDS, token: bare });
  assert.ok(REQ[0].path.indexOf("expand=fields(select=" + ST.STATION_FIELDS.join(",") + ")") > 0,
    "so listDelta puts it back before asking");
  assert.ok(REQ[0].path.indexOf("token=TBARE") > 0, "without losing the token");
  pass("a deltaLink that lost its $expand gets it back, so a poll never answers rows with no fields");

  /* AMENDMENT 8: a stale token and a list that will not serve a delta at all
     are the same error to isDeltaRestart and different to isDeltaResync */
  const restart = m => { try { throw new Error("delta must be restarted: " + m); } catch (e) { return e; } };
  assert.strictEqual(CW.isDeltaResync(restart("GET /x -> 410 gone")), true);
  assert.strictEqual(CW.isDeltaResync(restart("GET /x -> 409 resyncChangesApplyDifferences")), true);
  assert.strictEqual(CW.isDeltaResync(restart("GET /x -> 400 invalidRequest")), false,
    "a refusal is not a resync: that list will not serve a delta at all");
  assert.strictEqual(CW.isDeltaResync(restart("GET /x -> 403 accessDenied")), false);
  assert.strictEqual(CW.isDeltaResync(new Error("Failed to fetch")), false,
    "and a passing network failure is neither");
  pass("a stale delta token is told apart from a list that refuses delta outright");

  /* ---- which permission each request asks for ---- */
  const WBBASE = "/sites/" + SITE + "/drive/items/01ABCDEF/workbook";
  assert.deepStrictEqual(CW._scopeFor(WBBASE + "/worksheets('Production')/range(address='A1:K600')"), CW.SCOPES,
    "a range address is nothing but colons, and is still a workbook write");
  assert.deepStrictEqual(CW._scopeFor(WBBASE + "/worksheets('Dashboard Log')/range(address='A403:F403')"), CW.SCOPES);
  assert.deepStrictEqual(CW._scopeFor("/sites/" + SITE + "/drive/items/01ABCDEF/content"), CW.SCOPES);
  assert.deepStrictEqual(CW._scopeFor("/sites/costellowindowsie.sharepoint.com:/sites/ProductionProgress"), CW.SCOPES,
    "the workbook's own site, looked up by path at sign-in: SCOPES, or nobody signs in");
  assert.deepStrictEqual(CW._scopeFor(HOST_LOOKUP), CW.LIST_SCOPES);
  assert.deepStrictEqual(CW._scopeFor(OWN_LOOKUP), CW.LIST_SCOPES,
    "and so is the station's own lookup of the workbook's site: quiet, never a consent window");
  /* AMENDMENT 7: a list path is a list call whatever else is in it. A site id
     carrying "/drive" in it - or any future list path that does - must not be
     read as a workbook call and sent with the wrong scopes, non-quietly. */
  assert.deepStrictEqual(CW._scopeFor("/sites/host,a/drive/b/lists/L/items?expand=fields"), CW.LIST_SCOPES,
    "/lists is decided before /drive/, so a list call can never be mistaken for the file");
  assert.deepStrictEqual(CW._scopeFor("/sites/" + SITE + "/drive/items/01A/workbook/worksheets"), CW.SCOPES,
    "while a real workbook path is still a workbook path");
  assert.deepStrictEqual(CW._scopeFor(FLISTS_PATH + "/" + LOG_ID + "/items/delta?expand=fields"), CW.LIST_SCOPES,
    "and the delta endpoint is a list call like any other");
  assert.ok(CW.SCOPES.indexOf("Sites.ReadWrite.All") < 0, "the workbook scopes still do not include it");
  pass("the scope is chosen on the shape of the path: a range address is a workbook write, delta is a list call");

  const SEEN = [];
  CW._setToken((sc, quiet) => { SEEN.push({ sc: sc, quiet: quiet }); return "t"; });
  forget(); reset();
  await CW.stationSite();
  await CW.listDelta("Glass station", opts);
  assert.ok(SEEN.length >= 2);
  assert.ok(SEEN.every(x => x.sc.indexOf("Sites.ReadWrite.All") >= 0));
  assert.ok(SEEN.every(x => x.quiet === true), "quietly: a background poll can never open a popup");
  CW._setToken(() => "t");
  pass("a delta poll asks for the list scope quietly, so a ten-second poll can never throw up a window");

  /* ================= 13. the feeder, in app.js ================= */
  ITEMS = []; LOGITEMS = []; forget(); consent(true);
  useJobs([mkJob({ id: "R5303", cust: "Customer One", glass: { tg: 6 }, blk: 4, seq: 0 }),
           mkJob({ id: "R5304", cust: "Customer Two", glass: { dg: 2 }, blk: 4, seq: 1 })]);
  A("STATION_FEED = { hash: '', at: 0 }");
  reset();
  let fed = await feedStation();
  assert.strictEqual(fed.sent, 2);
  assert.strictEqual(ITEMS.length, 2);
  assert.ok(writes().every(w => w.method === "POST"), "two new jobs: two POSTs and nothing else");
  assert.strictEqual(A("STATION_OK"), true);
  pass("the feeder pushes the sheet's glass into the list on a load");

  reset();
  assert.strictEqual(await feedStation(), null, "the same slice, a moment later: skipped");
  assert.strictEqual(REQ.length, 0, "not even a read");
  pass("an unchanged slice inside ten minutes does not go near SharePoint");

  A("STATION_FEED.at = Date.now() - 700000");          // eleven minutes ago
  reset();
  await feedStation();
  assert.ok(REQ.length > 0, "past ten minutes it reads again");
  assert.strictEqual(writes().length, 0, "and finds there is nothing to write");
  pass("after ten minutes the feeder looks again, and writes nothing when nothing changed");

  ITEMS[0].fields.Cut = 6; ITEMS[0].fields.Hotmelt = 6; ITEMS[0].fields.Glazed = 3;
  ITEMS[0].fields.CutBy = "Person A"; ITEMS[0].fields.CutAt = "2026-09-08T12:00:00.000Z";
  ITEMS[0].fields.DoneBy = "Person A"; ITEMS[0].fields.DoneAt = "2026-09-08T12:00:00.000Z";
  useJobs([mkJob({ id: "R5303", cust: "Customer One", glass: { tg: 6 }, blk: 3, seq: 0 }),   // Ready to fit now
           mkJob({ id: "R5304", cust: "Customer Two", glass: { dg: 2 }, blk: 4, seq: 1 })]);
  A("STATION_FEED = { hash: '', at: 0 }");
  reset();
  await feedStation();
  assert.strictEqual(writes().length, 1);
  assert.deepStrictEqual(Object.keys(writes()[0].body).sort(), ["Active", "FedAt", "FedBy"]);
  assert.strictEqual(ITEMS[0].fields.Cut, 6, "the floor's counters are exactly where they were");
  assert.strictEqual(ITEMS[0].fields.CutBy, "Person A", "and so is who put them there");
  pass("a job leaving production is patched Active = No, and the floor's record is not touched");

  const feederBodies = ALLREQ.filter(r => r.method === "POST" || r.method === "PATCH")
    .map(r => r.body && (r.body.fields || r.body)).filter(Boolean);
  assert.ok(feederBodies.length > 0);
  assert.strictEqual(ALLREQ.filter(r => r.method === "DELETE").length, 0,
    "the feeder has issued no DELETE, and there is no code path that could");
  pass("across every feed so far, not one DELETE has been sent");

  consent(false);
  A("STATION_FEED = { hash: '', at: 0 }");
  reset();
  assert.strictEqual(await feedStation(), null);
  assert.strictEqual(REQ.length, 0, "no token, no read, no write, no window in anybody's face");
  consent(true);
  pass("without the list permission already granted, the feeder does nothing at all");

  SITE_EXISTS = false; forget(); A("STATION_FEED = { hash: '', at: 0 }");
  reset();
  assert.strictEqual(await feedStation(), null);
  assert.strictEqual(writes().length, 0);
  assert.strictEqual(A("STATION_OK"), false);
  assert.strictEqual(TOASTS.length, 0, "and it never toasts: the dashboard carries on regardless");
  SITE_EXISTS = true; forget();
  pass("a missing Floor stations site is a plain state on the board, not an error anybody sees");

  ITEMS = []; forget();
  const many = [];
  for (let i = 0; i < 70; i++) many.push(mkJob({ id: "R" + (6000 + i), cust: "C" + i, glass: { tg: 2 }, blk: 4, seq: i }));
  useJobs(many);
  A("STATION_FEED = { hash: '', at: 0 }");
  reset();
  fed = await feedStation();
  assert.strictEqual(fed.sent, 60, "sixty writes, and the rest waits for the next load");
  assert.ok(peakInflight <= 3, "at most three writes in flight at once, saw " + peakInflight);
  assert.strictEqual(A("STATION_FEED.hash"), "", "a run that did not finish is not remembered as done");
  fed = await feedStation();
  assert.strictEqual(fed.sent, 10, "the next run finishes the job");
  pass("a big first run is capped at sixty writes, three at a time, and finishes on the next load");

  assert.ok(A("stationAgainT") !== null, "the capped run left a follow-up due");
  const t1 = A("stationAgainT");
  stationFeedAgain();
  assert.strictEqual(A("stationAgainT"), t1, "and asking again does not stack a second timer");
  A("clearTimeout(stationAgainT); stationAgainT = null;");
  pass("a run cut short by the cap schedules exactly one follow-up run, not a timer per call");

  useJobs([mkJob({ id: "R6000", cust: "CHANGED", glass: { tg: 2 }, blk: 4, seq: 0 })]);
  A("STATION_FEED = { hash: '', at: 0 }");
  FAIL_ONCE = 99;
  reset();
  fed = await feedStation();
  assert.ok(fed.failed > 0);
  assert.strictEqual(TOASTS.length, 0, "a failed feed never toasts");
  assert.ok(/refused/.test(A("STATION_FEED_ERR")));
  A("clearTimeout(stationAgainT); stationAgainT = null;");
  FAIL_ONCE = 0;
  pass("a refused feed sets a word in the footer, never a toast, and never blocks the dashboard");

  ITEMS = []; forget();
  useJobs([mkJob({ id: "R6400", cust: "Small", glass: { tg: 1 }, blk: 4, seq: 0 })]);
  A("STATION_FEED = { hash: '', at: 0 }; STATION_FEED_ERR = 'something went wrong earlier';");
  await feedStation();
  assert.strictEqual(A("STATION_FEED_ERR"), "", "a run that succeeded clears the failure");
  A("STATION_FEED_ERR = 'stale'");
  assert.strictEqual(await feedStation(), null, "and so does a run that skips");
  assert.strictEqual(A("STATION_FEED_ERR"), "");
  pass("the footer's failure word is cleared by the next run that skips or succeeds");

  ITEMS = []; forget();
  useJobs([mkJob({ id: "R6500", cust: "Twin", glass: { tg: 1 }, blk: 4, seq: 0 })]);
  A("STATION_FEED = { hash: '', at: 0 }");
  const both = await Promise.all([feedStation(), feedStation()]);
  assert.strictEqual(both.filter(Boolean).length, 1, "one of the two turned round at the door");
  assert.strictEqual(ITEMS.length, 1, "and the row was created once, not twice");
  pass("the busy flag goes up before the first await, so two loads cannot feed the list twice");

  UNIQUE_TITLE = true;
  ITEMS = [item({ Title: "R7100|TG", Job: "R7100", Customer: "Old", GlassType: "TG",
                  Total: 1, Seq: 1, Active: "Yes", Cut: 1, CutBy: "Person A" }, "980")];
  reset();
  await stationAdd({ Title: "R7100|TG", Job: "R7100", Customer: "New", GlassType: "TG",
                     Total: 1, Seq: 1, Active: "Yes", FedAt: AT, FedBy: "x" },
                   { siteId: FSITE, fields: ST.STATION_FIELDS });
  assert.strictEqual(ITEMS.length, 1, "no second item for that Title");
  assert.strictEqual(ITEMS[0].fields.Customer, "New", "the refused POST became a patch of the row that won");
  assert.strictEqual(ITEMS[0].fields.Cut, 1, "and the floor's count on it is untouched");
  UNIQUE_TITLE = false;
  pass("a POST refused because the row already exists is read back and patched, not lost");

  const acctDesc = Object.getOwnPropertyDescriptor(CW, "account");
  Object.defineProperty(CW, "account", { configurable: true, get: () => ({ name: "Person G", username: "g@example.test" }) });
  assert.strictEqual(feedWho(), "Person G");
  Object.defineProperty(CW, "account", { configurable: true, get: () => ({ username: "g@example.test" }) });
  assert.strictEqual(feedWho(), "g", "with no display name, the local part - never the whole address");
  Object.defineProperty(CW, "account", { configurable: true, get: () => null });
  assert.strictEqual(feedWho(), "dashboard");
  Object.defineProperty(CW, "account", acctDesc);
  pass("FedBy carries a name or a local part, so no address ever reaches the floor's list");

  A("STATION_FEED = { hash: '', at: 0 }; STATION_FEED_ERR = '';");
  setStationFoot();
  assert.strictEqual(EL["#stationfeedwrap"].hidden, true, "nothing to say: the separator goes too");
  A("STATION_FEED = { hash: 'h', at: Date.now() };");
  setStationFoot();
  assert.strictEqual(EL["#stationfeedwrap"].hidden, false);
  assert.ok(EL["#stationfeed"].textContent.indexOf("station feed:") === 0);
  pass("the footer's separator appears with the station word and disappears with it");

  forget(); reset();
  assert.strictEqual(await CW.stationSite(), FSITE);
  CW.forgetStationSite();
  assert.strictEqual(mem.cw_stationsite, undefined, "forgotten in this browser too, not just in memory");
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE, "and resolved again on the next call");
  assert.strictEqual(REQ.length, 1);
  pass("the cached Floor stations id can be forgotten, and is resolved again straight away");

  /* ---- (3) the fallback in use, and then the site appears ---- */
  SITE_EXISTS = false; OWN_HAS_LISTS = true; forget(); reset();
  assert.strictEqual(await CW.stationSite(), SITE, "on the fallback to begin with");
  await CW.listId("Glass station", { siteId: SITE, fields: ST.STATION_FIELDS });
  assert.ok(JSON.parse(mem.cw_listids)[SITE + "|Glass station"], "with the list found in that site");
  const genBefore = CW._stationSiteInfo().gen;
  SITE_EXISTS = true;                                  // the manager makes the site
  reset();
  assert.strictEqual(await CW.stationSite(), SITE, "nothing changes until the re-check is due");
  assert.strictEqual(REQ.length, 0, "and it is not looked for six times a minute in the meantime");
  CW._stationSiteLookedAt(Date.now() - 700000);        // eleven minutes ago
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE, "past ten minutes it looks, finds it, and moves");
  assert.strictEqual(CW._stationSiteInfo().own, false);
  assert.strictEqual(CW._stationSiteInfo().gen, genBefore + 1, "and says the lists have moved");
  assert.strictEqual(JSON.parse(mem.cw_listids)[SITE + "|Glass station"], undefined,
    "the list ids found in the old site go with it: a list id only means anything in its own site");
  assert.strictEqual(await CW.listId("Glass station", { siteId: FSITE, fields: ST.STATION_FIELDS }), GLASS_ID,
    "and the lists are found again in the new one");
  OWN_HAS_LISTS = false; forget();
  pass("a Floor stations site that appears later is picked up at the next ten-minute look, and the lists move with it");

  /* ---- AMENDMENT 1: the first blocker ----
     forgetStationSite() nulls the cached id, so a later rememberStationSite()
     has nothing to compare against and would say nothing about the move. Both
     screens would then keep polling a delta token issued in the OLD site while
     writing into the new one. Every forget is a move, whether or not anything
     is yet known about where to next. */
  SITE_EXISTS = false; OWN_HAS_LISTS = true; forget(); reset();
  assert.strictEqual(await CW.stationSite(), SITE, "on the fallback, with a token about to be taken");
  const genA = CW._stationSiteInfo().gen;
  await CW.listId("Glass station", { siteId: SITE, fields: ST.STATION_FIELDS });
  CW.forgetStationSite();                              // what a 404 on a list call does
  assert.strictEqual(CW._stationSiteInfo().gen, genA + 1,
    "forgetting the site is itself a move: anything held about the old one is stale");
  assert.strictEqual(JSON.parse(mem.cw_listids)[SITE + "|Glass station"], undefined,
    "and the list ids found in it went with it");
  SITE_EXISTS = true;
  reset();
  assert.strictEqual(await CW.stationSite(), null,
    "AMENDMENT 4: forgetting does not restart the cadence, so it is not looked for again at once");
  assert.strictEqual(REQ.length, 0);
  CW._stationSiteLookedAt(Date.now() - 70000);         // and when the look is due
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE, "the real site is found");
  assert.strictEqual(CW._stationSiteInfo().gen, genA + 1,
    "the forget already announced the move; resolving somewhere new does not need to announce it twice");
  reset();
  const afterMove = await CW.listItems("Glass station", { siteId: await CW.stationSite(), fields: ST.STATION_FIELDS });
  assert.ok(afterMove, "and the list reads");
  assert.ok(REQ.every(r => r.path.indexOf(OWNLISTS_PATH) < 0),
    "with not one request left addressing the site it came from");
  assert.ok(REQ.some(r => r.path.indexOf(FLISTS_PATH) === 0), "every one of them names the new site");
  OWN_HAS_LISTS = false; forget();
  pass("forgetting the site counts as a move, so a page that then finds the real one drops its old tokens");

  /* AMENDMENT 5: and the office throws away everything else it was holding */
  A("STATION_SITE_GEN = 0; STATION_FEED = { hash: 'stale', at: Date.now() };");
  A("STATION_PEOPLE = [{ id: '1', fields: { Title: 'Person A', Station: 'Glass', Stages: 'cut', Active: 'Yes' } }];");
  A("STATION_FEEDS.items.token = 'old'; STATION_FEEDS.log.token = 'old';");
  stationSiteMoved(99);
  assert.strictEqual(A("STATION_FEED.hash"), "", "the 'already fed, nothing changed' hash is gone");
  assert.strictEqual(A("STATION_PEOPLE"), null, "so is the roster the log window filters by");
  assert.strictEqual(A("STATION_FEEDS.items.token"), null, "and both delta tokens with them");
  assert.strictEqual(A("STATION_FEEDS.log.token"), null);
  assert.strictEqual(A("STATION_SITE_GEN"), 99);
  A("STATION_SITE_GEN = CW.stationSiteMoves(); STATION_FEED = { hash: '', at: 0 };");
  pass("a move clears the feeder's hash, the people cache and both tokens, so the new site is fed and re-read");

  /* ---- (4) a refusal on the Floor stations lookup ---- */
  SITE_403 = true; OWN_HAS_LISTS = true; forget(); reset();
  assert.strictEqual(await CW.stationSite(), SITE,
    "cannot see it is not the same as there is none, but it falls back just the same");
  assert.strictEqual(CW._stationSiteInfo().own, true);
  CW._stationSiteLookedAt(Date.now() - 70000);         // seventy seconds ago
  SITE_403 = false;
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE,
    "and a refusal is looked at again within the minute, not held for ten");
  assert.strictEqual(CW._stationSiteInfo().own, false);
  SITE_403 = false; SITE_EXISTS = true; OWN_HAS_LISTS = false; forget();
  pass("a refused Floor stations lookup falls back too, and is never cached as a settled miss");

  /* ---- AMENDMENT 4: what may and may not forget the site ----
     Only "there is no such thing here" is a reason to doubt the cached site. A
     refusal or a dead connection says nothing about where the lists are, and
     dropping the site over one restarts the whole re-check cadence - or, on a
     403, flips a dashboard that is happily on the real site to the fallback. */
  forget();
  assert.strictEqual(await CW.stationSite(), FSITE);
  const genB = CW._stationSiteInfo().gen;
  A("stationWarned = true;");
  stationTrouble(new Error("GET /sites/x/lists/y/items -> 403 {\"error\":{\"code\":\"accessDenied\"}}"));
  assert.strictEqual(CW._stationSiteInfo().id, FSITE, "a refused list call keeps the site");
  stationTrouble(new Error("Failed to fetch"));
  assert.strictEqual(CW._stationSiteInfo().id, FSITE, "and so does a dead connection");
  assert.strictEqual(CW._stationSiteInfo().gen, genB, "neither of them is a move");
  stationTrouble(new Error("GET /sites/x/lists/y/items -> 404 {\"error\":{\"code\":\"itemNotFound\"}}"));
  assert.strictEqual(CW._stationSiteInfo().id, null, "only a 404 drops it");
  assert.strictEqual(CW._stationSiteInfo().gen, genB + 1, "and that is a move");
  pass("a refusal or a dropped connection never forgets the site; only a genuine 404 does");

  /* and a 403 on the re-check cannot flip a page that is on the real site */
  forget(); reset();
  assert.strictEqual(await CW.stationSite(), FSITE);
  SITE_403 = true; OWN_HAS_LISTS = true;
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE, "the real site is not re-checked at all");
  assert.strictEqual(REQ.length, 0, "so a refusal on that lookup cannot happen, let alone flip anything");
  SITE_403 = false; OWN_HAS_LISTS = false; forget();
  pass("a page already on the real site never looks again, so no refusal can move it to the fallback");

  /* ---- AMENDMENT 3: nothing is written without a resolved site ---- */
  reset();
  await assert.rejects(() => CW.listItems("Glass station", { siteId: null, fields: ST.STATION_FIELDS }),
    /no site resolved/, "a list call naming an empty site is an error, not a licence");
  assert.ok(REQ.every(r => r.path.indexOf("/drive") < 0),
    "and above all it never falls through to findFile() and asks for the drive");
  assert.strictEqual(REQ.length, 0, "it does not go out at all");
  pass("a list call with an explicit empty site throws rather than reaching for the workbook");

  /* a browser that was running the previous build has a bare id in storage */
  forget();
  mem.cw_stationsite = FSITE;
  reset();
  assert.strictEqual(await CW.stationSite(), FSITE, "an id left by the previous build still works");
  assert.strictEqual(REQ.length, 0, "without a lookup");
  assert.strictEqual(CW._stationSiteInfo().own, false, "and is taken as the real site, which is what it was");
  forget();
  pass("a cw_stationsite left by the previous build is read as it was meant, with no lookup and no reset");

  /* and the feeder works the same on either site, saying nothing about which */
  SITE_EXISTS = false; OWN_HAS_LISTS = true; forget(); consent(true);
  ITEMS = []; LOGITEMS = []; OWNITEMS = []; OWNLOG = [];
  useJobs([mkJob({ id: "R8100", cust: "Customer One", glass: { tg: 3 }, blk: 4, seq: 0 })]);
  A("STATION_FEED = { hash: '', at: 0 }; STATION_ITEMS = null; STATION_OK = null; stationReading = null;");
  A("STATION_FEEDS.items.token = null; STATION_FEEDS.items.off = 0; STATION_SITE_GEN = 0;");
  reset();
  const fedOwn = await feedStation();
  assert.strictEqual(fedOwn.sent, 1, "the feed went out");
  assert.strictEqual(OWNITEMS.length, 1, "and the row is on that site's own copy of the list");
  assert.strictEqual(ITEMS.length, 0, "and nothing at all was written to the other site");
  assert.ok(writes().every(w => w.path.indexOf(OWNLISTS_PATH) === 0),
    "every write addressed the workbook's own site");
  assert.ok(REQ.every(r => r.path.indexOf("/drive") < 0 && r.path.indexOf("/workbook") < 0));
  A("state.board = 'glass'");
  renderRows();
  const ownHtml = EL["#rows"].innerHTML;
  assert.ok(ownHtml.indexOf("R8100") > 0, "and the office board reads it back");
  ["Floor stations", "FloorStations", "ProductionProgress", "workbook site", "same site"]
    .forEach(w => assert.ok(ownHtml.indexOf(w) < 0, "nothing on screen says which site: " + w));
  A("state.board = null;");
  SITE_EXISTS = true; OWN_HAS_LISTS = false; forget();
  ITEMS = []; LOGITEMS = []; OWNITEMS = []; OWNLOG = [];
  A("STATION_SITE_GEN = CW.stationSiteMoves(); STATION_FEED = { hash: '', at: 0 };");
  pass("the feeder and the office board work the same on the fallback, and no wording anywhere names a site");

  /* ================= 14. the office: board, drawer, log window ================= */
  ITEMS = bItems.slice(); LOGITEMS = lItems.slice(); forget();
  A("STATION_ITEMS = null; STATION_OK = null; STATION_WHY = ''; STATION_FEEDS.items.token = null;");
  A("STATION_LOG = null; STATION_LOG_OK = null; STATION_PEOPLE = null;");
  await readStation();
  assert.strictEqual(A("STATION_OK"), true);
  await readStationLog();
  assert.strictEqual(A("STATION_LOG_OK"), true);
  assert.strictEqual(A("STATION_LOG.length"), lItems.length);
  pass("the office reads both of the floor's lists, and each answers for itself");

  A("state.board = 'glass'");
  renderRows();
  let html = EL["#rows"].innerHTML;
  assert.ok(html.indexOf("R5303") > 0 && html.indexOf("R5301") > 0, "every active job is on it");
  assert.ok(html.indexOf("R5299") < 0, "and the inactive one is not");
  ["Glass cut", "Hotmelt", "Glazing"].forEach(w => assert.ok(html.indexOf(w) > 0, w + " has a bar"));
  ["Cork", "eircode", "Comment", "Export", "Delete"].forEach(w =>
    assert.ok(html.indexOf(w) < 0, "the board must not show " + w));
  assert.ok(html.indexOf("<button") < 0, "the office's board is read-only: not one button on it");
  pass("the Glass station board replaces the job list, read-only, with the three new bars per job");

  assert.ok(/last: Person A cut TG 5→8/.test(html),
    "each card carries one line from the log: who moved what, and by how much");
  assert.ok(html.indexOf("Person C") > 0, "and the bars name who last moved each of them");
  pass("a board card says who last touched the job and what they did, straight out of the log");

  A("STATION_OK = false; STATION_WHY = STATION_NEED_CONSENT;");
  renderRows();
  assert.ok(EL["#rows"].innerHTML.indexOf("permission") > 0, "an ungranted permission is explained");
  A("STATION_OK = false; STATION_WHY = STATION_SITE_MISSING;");
  renderRows();
  assert.ok(EL["#rows"].innerHTML.indexOf("nothing in the Excel file is involved") > 0);
  A("STATION_OK = true; state.board = null;");
  pass("a missing site, a missing list and an ungranted permission each say so plainly");

  const drawerJob = mkJob({ id: "R5303", cust: "Customer One", glass: { tg: 6, arch: 2 } });
  html = stationSectionHtml(drawerJob);
  assert.ok(html.indexOf("Glass station") > 0 && html.indexOf("Hotmelt") > 0);
  assert.ok(/Person C/.test(html), "who last cut it");
  assert.ok(/Person A/.test(html), "and the log line under it");
  assert.ok(/5 → 8/.test(html), "with the numbers it moved between");
  assert.ok(html.indexOf("data-stfull") > 0, "and the way to the whole log");
  pass("the drawer shows each stage with who and when, and the job's own log lines under it");

  assert.strictEqual(stationSectionHtml(mkJob({ id: "R9999", glass: {} })), "",
    "a job with no glass has no station section at all");
  assert.ok(/Not fed to the floor yet/.test(stationSectionHtml(mkJob({ id: "R0404", glass: { tg: 1 } }))));
  html = stationSectionHtml(mkJob({ id: "R5299", glass: { tg: 3 } }));
  assert.ok(/Finished on the floor/.test(html), "a job off the floor's board keeps its record");
  assert.ok(html.indexOf("Not fed") < 0);
  pass("a job with no glass, one never fed, and one finished on the floor each read correctly");

  assert.ok(/Nothing recorded on the floor for this job yet/.test(stationTimelineHtml("R9999")));
  A("STATION_LOG_OK = false; STATION_LOG_WHY = STATION_LOG_MISSING;");
  assert.ok(/Station log/.test(stationTimelineHtml("R5303")), "a missing log list says which list it is");
  assert.ok(stationSectionHtml(drawerJob).indexOf("Hotmelt") > 0,
    "and the bars are still there: the two lists fail separately");
  A("STATION_LOG_OK = true;");
  pass("a missing Station log list is its own state and never takes the bars away with it");

  const manyLog = [];
  for (let i = 0; i < 20; i++) manyLog.push(item({ Title: "R5303", Station: "Glass", GlassType: "TG",
    Stage: "cut", From: i, To: i + 1, Who: "Person A",
    At: "2026-09-0" + (i < 9 ? "1" : "2") + "T0" + (i % 9) + ":00:00.000Z" }, "9" + i));
  A("STATION_LOG = " + JSON.stringify(manyLog) + ";");
  html = stationTimelineHtml("R5303");
  assert.strictEqual((html.match(/stlrow/g) || []).length, 12, "twelve lines in the drawer, no more");
  assert.ok(/Full log \(20 lines\)/.test(html), "and the link says how many there are altogether");
  pass("the drawer's timeline is capped at twelve lines, with the count on the way to the rest");

  /* one line from somebody who is no longer on the people list, so the window's
     person filter can be seen to offer the list AND everyone who has recorded */
  const winLog = lItems.concat([item({ Title: "R5310", Station: "Glass", GlassType: "DG",
    Stage: "glazed", From: 0, To: 2, Who: "Person H", At: "2026-09-05T08:00:00.000Z" }, "70")]);
  A("STATION_LOG = " + JSON.stringify(winLog) + "; STATION_PEOPLE = " + JSON.stringify(pItems) + ";");
  openStationLog("");
  let lw = logFrame();
  assert.ok(lw.indexOf("Glass station log") > 0);
  assert.ok(lw.indexOf("Person A") > 0 && lw.indexOf("Person B") > 0 && lw.indexOf("Person C") > 0);
  assert.ok(lw.indexOf("Person Z") < 0, "another station's line is not this station's log");
  assert.ok(lw.indexOf("Person F") > 0, "somebody on the list who has recorded nothing is still a filter");
  assert.ok(lw.indexOf("Person H") > 0, "and so is somebody who has recorded something and left the list");
  assert.ok(lw.indexOf("Person D") < 0, "but a deactivated person with nothing recorded is not offered");
  assert.ok(lw.indexOf("id=\"lgwho\"") > 0 && lw.indexOf("id=\"lgstage\"") > 0 &&
            lw.indexOf("id=\"lgjob\"") > 0 && lw.indexOf("id=\"lgday\"") > 0, "four filters");
  assert.ok(/>Clear filters</.test(lw), "and a button that says what it clears");
  pass("the log window opens with every line of this station's log and its four filters");

  assert.ok(/Per person/.test(logCounts()) && /Per stage/.test(logCounts()));
  assert.ok(/Person A <strong class="tab">8<\/strong>/.test(logCounts()), "eight units for Person A");
  assert.ok(/Glass cut <strong class="tab">8<\/strong>/.test(logCounts()), "and eight for the cutting stage");
  const whole = lw + logCounts() + logBody();
  assert.ok(whole.indexOf("Delete") < 0 && whole.indexOf("Export") < 0 && whole.indexOf("Download") < 0,
    "there is no delete and no export on the floor's log");
  pass("the log window counts units and lines per person and per stage, and offers no way to change them");

  /* AMENDMENT 3: a filter repaints the rows and the counts and leaves the bar
     - and the box being typed into - exactly where it was */
  const frameWas = logFrame();
  A("LOGF = { who: 'Person A', stage: '', job: '', day: '', show: 200 };");
  paintStationLog();
  assert.strictEqual(logHead(), "2 lines", "filtered to one person: two lines");
  assert.ok(logBody().indexOf("R5301") < 0, "and the other person's job is gone with them");
  assert.strictEqual(logFrame(), frameWas, "while the filter bar itself was not rebuilt");
  A("LOGF = { who: '', stage: '', job: '', day: '2026-09-07', show: 200 };");
  paintStationLog();
  assert.strictEqual(logHead(), "1 line", "a day narrows it to that day");
  A("LOGF = { who: '', stage: '', job: '530', day: '', show: 200 };");
  paintStationLog();
  assert.strictEqual(logHead(), "4 lines", "a part of a job number still finds the jobs");
  A("LOGF = { who: '', stage: '', job: 'R5303', day: '', show: 200 };");
  paintStationLog();
  assert.strictEqual(logHead(), "3 lines", "and a whole one narrows it to that job");
  assert.strictEqual(logFrame(), frameWas, "through every one of those, the bar is untouched");
  pass("a filter or a poll repaints the rows and the counts only, never the filter bar");

  A("STATION_LOG = " + JSON.stringify(manyLog) + "; LOGF = { who:'', stage:'', job:'', day:'', show: 5 };");
  paintStationLog();
  assert.strictEqual((logBody().match(/lgrow/g) || []).length, 5, "only what the page says it is showing");
  assert.ok(/Show more \(15 left\)/.test(logBody()));
  A("LOGF.show = 200;");
  paintStationLog();
  assert.strictEqual((logBody().match(/lgrow/g) || []).length, 20);
  pass("the log window pages at what it says, and Show more brings the rest in");

  /* AMENDMENT 14: a job that has left the sheet is not a way into a drawer */
  useJobs([mkJob({ id: "R5303", cust: "Customer One", glass: { tg: 6 }, blk: 4, seq: 0 })]);
  A("STATION_LOG = " + JSON.stringify(winLog) + "; LOGF = { who:'', stage:'', job:'', day:'', show: 200 };");
  paintStationLog();
  assert.ok(/<button class="stn jump" data-j="R5303"/.test(logBody()),
    "a job still on the sheet is a button into its drawer");
  assert.ok(/<span class="stn" title="not on the sheet any more">R5310<\/span>/.test(logBody()),
    "one that has left it is plain text, so the window does not close on nothing");
  assert.ok(logBody().indexOf('data-j="R5310"') < 0);
  REG["lhost"].remove();
  pass("a log line for a job that is no longer on the sheet cannot close the window and open nothing");

  /* ---- the poll ---- */
  A("STATION_LOG = " + JSON.stringify(lItems) + "; STATION_LOG_OK = true; STATION_OK = true;");
  A("STATION_FEEDS.items.token = null; STATION_FEEDS.log.token = null;");
  ITEMS = bItems.slice(); LOGITEMS = lItems.slice();
  reset();
  assert.strictEqual(await stationPoll(), true, "the first pass enumerates both lists");
  assert.ok(A("STATION_FEEDS.items.token") && A("STATION_FEEDS.log.token"), "and keeps both tokens");
  reset();
  DELTA_NEXT[GLASS_ID] = [[]]; DELTA_NEXT[LOG_ID] = [[]];
  assert.strictEqual(await stationPoll(), false, "a pass where nothing moved answers no");
  assert.strictEqual(REQ.filter(r => r.path.indexOf("/delta") > 0).length, 2,
    "two requests, one per list, and not a full read between them");
  pass("the office poll asks each list what has changed since last time, and nothing more");

  DELTA_NEXT[GLASS_ID] = [[{ id: "500", fields: Object.assign({}, ITEMS[0].fields, { Glazed: 6 }) }]];
  DELTA_NEXT[LOG_ID] = [[]];
  reset();
  assert.strictEqual(await stationPoll(), true);
  assert.strictEqual(A("ST.jobBoard(STATION_ITEMS).find(g => g.job === 'R5303').bars.glazed.done"), 6,
    "a floor tap is merged into the office's own copy");
  assert.strictEqual(TOASTS.length, 0, "and a poll never toasts, however often it runs");
  pass("a counter moved on the floor reaches the office board through the poll, silently");

  DELTA_FAIL = 410; DELTA_FAIL_LEFT = 1; reset();
  assert.strictEqual(await stationPoll(), true, "a stale token is not a failure");
  assert.strictEqual(A("STATION_FEEDS.items.token"), null, "the token is thrown away");
  assert.ok(REQ.some(r => r.path.indexOf("/items?expand") > 0), "and the whole list is read instead");
  assert.strictEqual(A("STATION_OK"), true, "the board never went away");
  DELTA_FAIL = 0;
  A("STATION_FEEDS.items.token = null; STATION_FEEDS.log.token = null;");
  await stationPoll();
  pass("a 410 resync on the poll costs one full read and a new token, and nobody sees anything");

  A("state.board = null; state.sel = null;");
  assert.strictEqual(stationWatching(), false, "nobody looking at the floor: the slow rate");
  A("state.board = 'glass'");
  assert.strictEqual(stationWatching(), true);
  A("state.board = null; state.sel = 'R5303';");
  useJobs([mkJob({ id: "R5303", cust: "Customer One", glass: { tg: 6 }, blk: 4, seq: 0 }),
           mkJob({ id: "R9000", cust: "No glass", glass: {}, blk: 4, seq: 1 })]);
  assert.strictEqual(stationWatching(), true, "a drawer for a job with glass counts");
  A("state.sel = 'R9000'");
  assert.strictEqual(stationWatching(), false, "a drawer for a job without glass does not");
  assert.ok(A("STATION_FAST_MS") === 10000 && A("STATION_SLOW_MS") === 60000);
  A("state.sel = null;");
  pass("the poll runs every ten seconds while somebody is looking at the floor, and every minute otherwise");

  const realItems = CW.listItems, realDelta = CW.listDelta;
  A("STATION_OK = true; STATION_ERR = ''; stationWarned = true;");
  CW.listDelta = async () => { throw new Error("Failed to fetch"); };
  await stationPoll();
  assert.strictEqual(A("STATION_OK"), true, "the last board is still on screen");
  assert.strictEqual(A("STATION_ERR"), A("STATION_UNREACHABLE"));
  CW.listDelta = async () => { throw new Error("GET /sites/x/lists/y/items -> 404 {\"error\":{\"code\":\"itemNotFound\"}}"); };
  await stationPoll();
  assert.strictEqual(A("STATION_OK"), false, "a genuine 404 does clear it");
  CW.listDelta = realDelta; CW.listItems = realItems;
  A("STATION_OK = true; STATION_ERR = ''; STATION_WHY = '';");
  forget();
  pass("a passing failure on the poll keeps the board with a line above it; only a 404 clears it");

  assert.ok(NOACCESS_RE.test("GET /sites/x/drive/root/children -> 403 {\"error\":{\"code\":\"accessDenied\"}}"));
  assert.ok(!NOACCESS_RE.test("PATCH /worksheets('Production')/range(address='A403:K403') -> 423 locked"),
    "row 403 of the sheet is not a permission problem");
  const realOpen = CW.openSession;
  CW.openSession = async () => { throw new Error("GET /sites/x/drive/root/children -> 403 {\"error\":{\"code\":\"accessDenied\"}}"); };
  A("NOACCESS = false;");
  await start();
  assert.strictEqual(A("NOACCESS"), true);
  assert.ok(EL["#gateerr"].innerHTML.indexOf('href="glass.html"') > 0, "with the way to the station page");
  CW.openSession = realOpen;
  A("NOACCESS = false;");
  pass("a station account opening the dashboard is told so at once, from wherever the refusal comes");

  /* ---- AMENDMENT 1: the blocker ----
     A delta is in the air for as long as SharePoint takes to answer it, and a
     feed landing in that window nulls the token underneath it. Before the fix
     the answer - two changed rows and no token in hand - was read as a fresh
     enumeration and STATION_ITEMS collapsed to those two rows, with nothing to
     put it right for ten minutes. */
  ITEMS = bItems.slice(); LOGITEMS = lItems.slice(); forget();
  A("STATION_ITEMS = null; STATION_OK = null; STATION_WHY = ''; stationReading = null;");
  A("STATION_FEEDS.items.token = null; STATION_FEEDS.items.off = 0;");
  A("STATION_FEEDS.log.token = null; STATION_FEEDS.log.off = 0; STATION_LOG_OK = false;");
  await readStation();
  assert.strictEqual(A("STATION_ITEMS.length"), 4, "four rows on the board to begin with");
  DELTA_NEXT[GLASS_ID] = [[]];
  await stationPoll();
  assert.ok(A("STATION_FEEDS.items.token"), "and a delta token in hand");

  DELTA_NEXT[GLASS_ID] = [[{ id: "500", fields: Object.assign({}, bItems[0].fields, { Cut: 5 }) },
                           { id: "501", fields: Object.assign({}, bItems[1].fields, { Cut: 2 }) }]];
  const quickDelta = CW.listDelta;
  CW.listDelta = async (name, o) => { await settle(120); return quickDelta(name, o); };
  const inFlight = stationPoll();
  await settle(40);
  A("stationResetFeed('items')");            // exactly what a feed does, mid-flight
  await inFlight;
  CW.listDelta = quickDelta;
  assert.strictEqual(A("STATION_ITEMS.length"), 4,
    "the board did NOT collapse to the two rows that delta happened to name");
  assert.strictEqual(A("STATION_FEEDS.items.token"), null, "and the reset stood: the token is still gone");
  pass("a delta answer that arrives after a feed has replaced the list is thrown away, not merged");

  /* AMENDMENT 1: and the feeder stands off while a poll is merging */
  A("stationPolling = true;");
  useJobs([mkJob({ id: "R7700", cust: "Customer One", glass: { tg: 2 }, blk: 4, seq: 0 })]);
  A("STATION_FEED = { hash: '', at: 0 };");
  reset();
  assert.strictEqual(await feedStation(), null, "the feed turned round at the door");
  assert.strictEqual(REQ.length, 0, "not one request went out");
  assert.ok(A("stationAgainT") !== null, "and it comes back in half a minute instead");
  A("clearTimeout(stationAgainT); stationAgainT = null; stationPolling = false;");
  pass("the feeder waits rather than pulling the list out from under a poll that is merging");

  /* AMENDMENT 8: a list that refuses a delta outright is not asked six times a
     minute; a stale token is a different thing and costs one read */
  A("STATION_FEEDS.items.token = null; STATION_FEEDS.items.off = 0;");
  await stationPoll();                                   // a token again
  DELTA_FAIL = 400; DELTA_FAIL_LEFT = 1; reset();
  await stationPoll();
  assert.ok(A("STATION_FEEDS.items.off") > 0, "the refusal marked the list off");
  assert.ok(REQ.some(r => r.path.indexOf("/items?expand") > 0), "and it read the list the plain way");
  DELTA_FAIL = 0; DELTA_FAIL_LEFT = 0;
  reset();
  await stationPoll();
  assert.strictEqual(REQ.filter(r => r.path.indexOf("/items/delta") > 0).length, 0,
    "the next poll does not ask that list for a delta at all");
  assert.ok(REQ.some(r => r.path.indexOf("/items?expand") > 0), "it just reads it");
  assert.strictEqual(A("STATION_ITEMS.length"), 4, "and the board is intact through all of it");
  A("STATION_FEEDS.items.off = Date.now() - 400000;");   // more than five minutes ago
  reset();
  await stationPoll();
  assert.ok(REQ.some(r => r.path.indexOf("/items/delta") > 0), "after five minutes it tries delta again");
  assert.strictEqual(A("STATION_FEEDS.items.off"), 0, "and having been served one, stops marking it off");
  pass("a list that refuses a delta is polled the plain way for five minutes, then tried again");

  /* AMENDMENT 10: the office has no business holding anybody's PIN */
  A("STATION_PEOPLE = null; stationPeopleReading = null;");
  PEOPLEITEMS = [person("Person A", "cut", "1234")];
  reset();
  await readStationPeople();
  const peopleRead = REQ.find(r => r.path.indexOf(PEOPLE_ID) > 0);
  assert.ok(peopleRead, "it did read the people list");
  assert.ok(peopleRead.path.indexOf("expand=fields(select=Title,Station,Stages,Active)") > 0,
    "asking for four columns by name");
  assert.ok(peopleRead.path.indexOf("PIN") < 0, "and never for the PIN");
  assert.deepStrictEqual(ST.PEOPLE_FIELDS_OFFICE, ["Title", "Station", "Stages", "Active"]);
  assert.ok(ST.PEOPLE_FIELDS.indexOf("PIN") >= 0, "while the tablet, which compares it, still asks");
  pass("the office reads the people list without the PIN column; only the tablet asks for it");

  /* AMENDMENT 13: ninety days, and the parse done once per change */
  const ancient = item({ Title: "R1000", Station: "Glass", GlassType: "TG", Stage: "cut",
    From: 0, To: 1, Who: "Person A", At: "2020-01-01T00:00:00.000Z" }, "800");
  const undated = item({ Title: "R1001", Station: "Glass", GlassType: "TG", Stage: "cut",
    From: 0, To: 1, Who: "Person A", At: "" }, "801");
  assert.deepStrictEqual(stationLogRecent([ancient, undated], Date.parse("2026-09-08T12:00:00.000Z"))
    .map(x => x.id), ["801"], "the old line is left behind, the undated one is kept");
  LOGITEMS = lItems.concat([ancient, undated]);
  A("STATION_LOG = null; STATION_LOG_OK = null; stationLogReading = null;");
  await readStationLog();
  assert.ok(A("STATION_LOG.map(x => x.id).indexOf('800')") < 0, "a line from 2020 is not read into memory");
  assert.ok(A("STATION_LOG.map(x => x.id).indexOf('801')") >= 0);
  pass("the office reads only the last ninety days of the floor's log, keeping undated lines");

  A("LOGROWS = null; LOGROWS_OF = false;");
  const rowsA = A("logRowsNow()");
  assert.strictEqual(A("logRowsNow()") === rowsA, true, "asked twice, parsed once");
  A("STATION_LOG = STATION_LOG.slice();");
  assert.strictEqual(A("logRowsNow()") === rowsA, false, "and parsed again the moment the list changes");
  pass("the log is parsed and sorted once per change, not six times a minute");

  /* AMENDMENT 15: the poll's rate follows what is on screen, from the moment
     it goes on screen - not from whenever the next tick happened to be armed */
  LOGITEMS = lItems.slice();
  A("state.board = null; state.sel = null; STATION_TICK_MS = 0;");
  A("if (stationPollT) { clearTimeout(stationPollT); stationPollT = null; }");
  stationTick();
  assert.strictEqual(A("STATION_TICK_MS"), 60000, "nobody looking at the floor: the slow rate");
  useJobs([mkJob({ id: "R5303", cust: "Customer One", glass: { tg: 6 }, blk: 4, seq: 0 }),
           mkJob({ id: "R9000", cust: "Customer Nine", glass: {}, blk: 4, seq: 1 })]);
  renderChips();
  const showsel = REG["showsel"];
  assert.ok(showsel, "the Show dropdown is on the page");
  showsel.value = "glass";
  showsel.onchange();
  assert.strictEqual(A("state.board"), "glass");
  assert.strictEqual(A("STATION_TICK_MS"), 10000,
    "picking the board re-arms the tick fast, so the first delta lands within ten seconds");
  pass("choosing the Glass station board speeds the poll up at once rather than up to a minute later");

  A("state.board = null; STATION_TICK_MS = 0; state.sel = 'R9000';");
  stationTick();
  assert.strictEqual(A("STATION_TICK_MS"), 60000, "a drawer for a job with no glass changes nothing");
  A("state.sel = 'R5303';");
  openDrawer();
  assert.strictEqual(A("STATION_TICK_MS"), 10000, "but opening one for a job with glass does");
  const dh = $("#dhost"); if (dh) dh.remove();
  A("state.sel = null; STATION_TICK_MS = 0;");
  stationTick();
  assert.strictEqual(A("STATION_TICK_MS"), 60000);
  openStationLog("");
  assert.strictEqual(A("STATION_TICK_MS"), 10000, "and so does opening the log window");
  if (REG["lhost"]) REG["lhost"].remove();
  A("if (stationPollT) { clearTimeout(stationPollT); stationPollT = null; }");
  pass("opening the drawer or the log window re-arms the fast tick the same way the dropdown does");

  /* AMENDMENT 2: the drawer lists its own job and no other */
  A("STATION_OK = true; STATION_LOG_OK = true; LOGROWS = null; LOGROWS_OF = false;");
  A("STATION_LOG = " + JSON.stringify(lItems.concat([item({ Title: "R530", Station: "Glass",
      GlassType: "TG", Stage: "cut", From: 0, To: 9, Who: "Person H",
      At: "2026-09-08T13:00:00.000Z" }, "72")])) + ";");
  A("STATION_ITEMS = " + JSON.stringify(bItems) + ";");
  const tl = stationTimelineHtml("R530");
  assert.ok(/Person H/.test(tl), "the short job number's own line is there");
  assert.ok(tl.indexOf("Person B") < 0 && (tl.match(/stlrow/g) || []).length === 1,
    "and not one line belonging to R5303, which merely starts with it");
  assert.strictEqual((stationTimelineHtml("R5303").match(/stlrow/g) || []).length, 3,
    "while R5303 still gets its own three");
  pass("the drawer's timeline matches its job exactly, so R530 never lists R5303's work");

  assert.deepStrictEqual(STATIONS, [["glass", "Glass station"]],
    "one array next to SHEETNAMES is where the next station goes");
  assert.strictEqual(SHEETNAMES[0], "Production", "SHEETNAMES is untouched, for the export and the row chips");
  pass("the dropdown is driven by one STATIONS array, and SHEETNAMES is left exactly as it was");

  /* ================= 15. the tablet ================= */
  ITEMS = [item({ Title: "R5303|TG", Job: "R5303", Customer: "Customer One", GlassType: "TG", Total: 6,
                  Seq: 1, Active: "Yes", Cut: 0, Hotmelt: 0, Glazed: 0 }, "900"),
           item({ Title: "R5303|ARCH", Job: "R5303", Customer: "Customer One", GlassType: "ARCH", Total: 2,
                  Seq: 1, Active: "Yes", Cut: 0, Hotmelt: 0, Glazed: 0 }, "901")];
  PEOPLEITEMS = [person("Person A", "cut"), person("Person B", "hotmelt,glazed", "1234"),
                 person("Person C", "")];
  LOGITEMS = [];
  delete mem.cw_stationq; delete mem.cw_stationlogq; delete mem.cw_person;
  S("QUEUE = {}; LOGQ = {}; SITEID = " + JSON.stringify(FSITE) + "; TOKEN = null; OPEN = { R5303: 1 };");
  assert.strictEqual(await S("readPeople()"), true);
  assert.deepStrictEqual(S("PEOPLE.map(p => p.name)"), ["Person A", "Person B", "Person C"]);
  assert.strictEqual(S("PERSON"), null, "nobody has picked a name yet");
  assert.ok(EL["#board"].innerHTML.indexOf("Who are you?") > 0, "so the picker is what is on screen");
  assert.ok(EL["#board"].innerHTML.indexOf("Person A") > 0);
  pass("the tablet reads the people list and asks who is there before it shows anything else");

  assert.ok(EL["#board"].innerHTML.indexOf("Glass cut") > 0, "the picker says which stages each holds");
  assert.ok(/no stages yet/.test(EL["#board"].innerHTML), "including a person who holds none");
  pass("the picker is one big button per active person, saying which stages they hold");

  await S("readList()");
  assert.strictEqual(S("READY"), true);
  assert.ok(S("TOKEN") !== null, "the first read came back with a delta token for the poll");
  S("pickPerson(PEOPLE.find(p => p.name === 'Person A'))");
  assert.strictEqual(S("PERSON.name"), "Person A");
  assert.ok(mem.cw_person.indexOf("Person A") > 0, "kept on this tablet, with the time of the tap");
  assert.strictEqual(EL["#whois"].textContent, "Person A · Glass cut", "and the header says so");
  pass("picking a name with no PIN goes straight in, and the header shows the name and the stages");

  let bh = boardHtml();
  assert.ok(bh.indexOf("R5303") > 0 && bh.indexOf("Customer One") > 0);
  ["Glass cut", "Hotmelt", "Glazing"].forEach(w => assert.ok(bh.indexOf(w) > 0));
  ["Cork", "eircode", "@", "Export", "Delete", "index.html"].forEach(w =>
    assert.ok(bh.indexOf(w) < 0, "the station page must not show " + w));
  assert.ok(bh.indexOf('data-stage="hotmelt"') > 0, "hotmelt applies to every type, ARCH included");
  pass("the board draws job number, customer, glass chips and the three bars - nothing else");

  assert.ok(/data-stage="cut" data-act="1">\+<\/button>/.test(bh), "Person A may move cutting");
  assert.ok(/data-stage="hotmelt"[^>]*disabled aria-disabled="true"/.test(bh),
    "and the stages they do not hold are drawn disabled rather than hidden");
  assert.ok(/data-stage="glazed"[^>]*disabled/.test(bh));
  assert.ok(bh.indexOf("not yours") > 0, "with a word saying why");
  assert.ok(/Glazing/.test(bh) && /0 \/ 2/.test(bh), "the numbers are still there to read");
  pass("a stage this person does not hold is greyed and disabled, values still shown");

  reset();
  S("tap('900', 'hotmelt', 1)");
  await settle(40);
  assert.strictEqual(writes().length, 0, "not one request left the tablet");
  assert.strictEqual(S("Object.keys(QUEUE).length"), 0, "and nothing was even queued");
  assert.strictEqual(ITEMS[0].fields.Hotmelt, 0);
  S("tap('900', 'glazed', 'all')");
  await settle(40);
  assert.strictEqual(writes().length, 0, "'All' follows the same rule as the steppers");
  pass("tap() itself refuses a stage the person does not hold - the disabled button is only the drawing");

  reset();
  S("tap('900', 'cut', 1)");
  assert.ok(boardHtml().indexOf("1 / 6") > 0, "the new number is on screen before anything is sent");
  await settle(80);
  let patched = writes().filter(w => w.method === "PATCH");
  assert.strictEqual(patched.length, 1, "one PATCH, once");
  assert.deepStrictEqual(Object.keys(patched[0].body).sort(), ["Cut", "CutAt", "CutBy", "DoneAt", "DoneBy"]);
  assert.strictEqual(patched[0].body.Cut, 1);
  assert.strictEqual(patched[0].body.CutBy, "Person A");
  assert.strictEqual(patched[0].body.DoneBy, "Person A");
  assert.strictEqual(ITEMS[0].fields.Cut, 1);
  assert.strictEqual(ITEMS[0].fields.Total, 6, "and the job's facts are exactly as the office left them");
  pass("a tap shows at once and sends one PATCH carrying that stage's counter, its By/At and the last touch");

  const FEEDERCOLS = ["Job", "Customer", "GlassType", "Total", "Seq", "Active", "FedAt", "FedBy", "Title"];
  assert.ok(FEEDERCOLS.every(k => !(k in patched[0].body)), "the station never writes a job fact");
  assert.strictEqual(ALLREQ.filter(r => r.method === "DELETE").length, 0);
  pass("the station's counter write carries no job fact - not even Title - and it has sent no DELETE");

  const logPosts = writes().filter(w => w.method === "POST" && w.path.indexOf(LOG_ID) > 0);
  assert.strictEqual(logPosts.length, 1, "one log line for one sent write");
  assert.deepStrictEqual(logPosts[0].body.fields,
    { Title: "R5303", Station: "Glass", GlassType: "TG", Stage: "cut", From: 0, To: 1,
      Who: "Person A", At: logPosts[0].body.fields.At });
  assert.ok(/^\d{4}-\d\d-\d\dT/.test(logPosts[0].body.fields.At));
  assert.strictEqual(LOGITEMS.length, 1, "and it really is on the list");
  pass("every counter write leaves one Station log line: job, type, stage, From, To, who and when");

  const orderOfWrites = writes().map(w => w.method + (w.path.indexOf(LOG_ID) > 0 ? " log" : " counter"));
  assert.strictEqual(orderOfWrites.indexOf("PATCH counter") < orderOfWrites.indexOf("POST log"), true,
    "the counter first, the line about it second");
  pass("the log line is written only after the counter PATCH it describes has succeeded");

  S("QUEUE = {}; LOGQ = {}; flushing = true;");            // hold the queue shut while the taps land
  reset();
  S("tap('900', 'cut', 1); tap('900', 'cut', 1); tap('900', 'cut', 1);");
  assert.strictEqual(S("Object.keys(QUEUE).length"), 1, "one entry for the row and stage, not three");
  assert.strictEqual(S("QUEUE['900|cut'].value"), 4);
  S("flushing = false;");
  await S("flushQueue()");
  await settle(60);
  assert.strictEqual(writes().filter(w => w.method === "PATCH").length, 1,
    "three taps, one write: the floor's latest count, not a replay of every press");
  const merged1 = writes().filter(w => w.method === "POST" && w.path.indexOf(LOG_ID) > 0);
  assert.strictEqual(merged1.length, 1, "and one log line, not three");
  assert.strictEqual(merged1[0].body.fields.From + "->" + merged1[0].body.fields.To, "1->4",
    "saying where it started and where it got to");
  pass("a run of quick taps merges into one write and one log line saying 1 to 4");

  /* AMENDMENT 5: From is the number the office could last see, taken when the
     entry went into the queue - never re-derived at flush time, where a lost
     response or a poll that has already merged this very change would make
     From equal To and swallow the line entirely. */
  ITEMS[0].fields.Cut = 2;
  S("QUEUE = {}; LOGQ = {};");
  await S("readList()");
  S("flushing = true; tap('900', 'cut', 1);");     // nothing leaves until the queue is let go
  assert.strictEqual(S("QUEUE['900|cut'].from"), 2, "captured at the tap, from the list");
  S("tap('900', 'cut', 1); tap('900', 'cut', 1);");
  assert.strictEqual(S("QUEUE['900|cut'].from"), 2, "and kept across the taps that merge into it");
  assert.strictEqual(S("QUEUE['900|cut'].value"), 5);
  /* the list agrees with the new number before the write goes - a poll got
     there first. Re-deriving From here would give 5, and the line would say
     5 to 5 and be dropped on the floor. */
  ITEMS[0].fields.Cut = 5;
  S("ITEMS[0].fields.Cut = 5; flushing = false;");
  reset();
  await S("flushQueue()");
  await settle(60);
  const savedLine = writes().find(w => w.method === "POST" && w.path.indexOf(LOG_ID) > 0);
  assert.ok(savedLine, "the log line was still written");
  assert.strictEqual(savedLine.body.fields.From + "->" + savedLine.body.fields.To, "2->5",
    "saying where the floor started and where it got to, not 5 to 5");
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("the log line's From is taken when the tap is queued, so nothing can swallow the line later");

  ITEMS[0].fields.Cut = 4;
  S("QUEUE = {}; LOGQ = {};");
  await S("readList()");
  FAIL_503 = 1;
  reset();
  S("tap('900', 'cut', 1)");
  await settle(3400);                            // call() backs off ~2 s before the retry
  assert.strictEqual(ITEMS[0].fields.Cut, 5, "the retry landed");
  assert.strictEqual(S("Object.keys(QUEUE).length"), 0, "and the queue is clear again");
  pass("a 503 on a counter write is retried by call() itself, and the tap is not lost");

  S("QUEUE = {}; LOGQ = {};");
  FAIL_ONCE = 99;
  reset();
  S("tap('900', 'cut', 1)");
  await settle(120);
  assert.strictEqual(S("Object.keys(QUEUE).length"), 1, "still owed");
  assert.strictEqual(S("QUEUE['900|cut'].err"), 1);
  assert.strictEqual(S("Object.keys(LOGQ).length"), 0, "and no log line for a write that never landed");
  S("render()");
  assert.ok(boardHtml().indexOf("not saved yet") > 0, "the row says so, quietly");
  assert.ok(mem.cw_stationq.indexOf('"value":6') > 0, "the owed write is in localStorage for next time");
  FAIL_ONCE = 0;
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("a counter write that fails is kept, marked on the row, and writes no log line about it");

  /* a log line that will not go: the counter is not held up by it */
  ITEMS[0].fields.Cut = 0;
  S("QUEUE = {}; LOGQ = {};");
  await S("readList()");
  FAIL_LOG = 99;
  reset();
  S("tap('900', 'cut', 2)");
  await settle(160);
  assert.strictEqual(ITEMS[0].fields.Cut, 2, "the counter went through");
  assert.strictEqual(S("Object.keys(QUEUE).length"), 0, "and is not owed any more");
  assert.strictEqual(S("Object.keys(LOGQ).length"), 1, "while the log line is still owed");
  FAIL_LOG = 0;
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  reset();
  await S("flushQueue()");
  await settle(60);
  assert.strictEqual(S("Object.keys(LOGQ).length"), 0, "and goes out on the next pass");
  assert.ok(LOGITEMS.some(x => x.fields.From === 0 && x.fields.To === 2));
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("a refused log line is retried with the queue and never blocks or undoes a counter");

  /* a tap during an in-flight write */
  S("QUEUE = {}; LOGQ = {}; if (retryT) { clearTimeout(retryT); retryT = null; }");
  ITEMS[0].fields.Cut = 0;
  await S("readList()");
  const realPatch = CW.listPatch;
  CW.listPatch = async (name, id, fields, o) => { await settle(140); return realPatch(name, id, fields, o); };
  reset();
  S("tap('900', 'cut', 1)");                     // starts a write that takes 140 ms
  await settle(40);
  S("tap('900', 'cut', 1)");                     // lands in the middle of it
  assert.strictEqual(S("flushing"), true, "the second tap really did arrive during a write");
  await settle(900);
  CW.listPatch = realPatch;
  assert.strictEqual(S("Object.keys(QUEUE).length"), 0, "nothing is left owed");
  assert.strictEqual(ITEMS[0].fields.Cut, 2, "and the second tap was sent, without anyone nudging it");
  assert.strictEqual(writes().filter(w => w.method === "PATCH").length, 2);
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("a tap made during an in-flight write is drained in the same pass, never stranded");

  S("QUEUE = {}; LOGQ = {}; READY = false; if (retryT) { clearTimeout(retryT); retryT = null; }");
  S("QUEUE['900|cut'] = { id: '900', stage: 'cut', value: 3, who: 'Person A', at: '2026-09-08T07:00:00.000Z', job: 'R5303', type: 'TG', err: 0 };");
  await S("flushQueue()");
  assert.ok(S("retryT !== null"), "no site yet, so the queue has a retry due rather than being forgotten");
  S("clearTimeout(retryT); retryT = null; READY = true;");
  pass("a queue that cannot be sent yet - no site, no list - still has a retry armed");

  /* who and when belong to the tap */
  S("QUEUE = {}; LOGQ = {};");
  ITEMS[0].fields.Cut = 0;
  await S("readList()");
  S("tap('900', 'cut', 1)");
  assert.strictEqual(S("QUEUE['900|cut'].who"), "Person A", "captured at the tap");
  assert.ok(/^\d{4}-\d\d-\d\dT/.test(S("QUEUE['900|cut'].at")));
  await settle(80);
  S("QUEUE = { '900|cut': { id: '900', stage: 'cut', value: 4, who: 'Person B', at: '2026-09-08T06:00:00.000Z', job: 'R5303', type: 'TG', err: 0 } };");
  S("PERSON = PEOPLE.find(p => p.name === 'Person A');");
  reset();
  await S("flushQueue()");
  await settle(60);
  const late = writes().find(w => w.method === "PATCH");
  assert.strictEqual(late.body.CutBy, "Person B", "a tap sent hours later is still the person who made it");
  assert.strictEqual(late.body.CutAt, "2026-09-08T06:00:00.000Z");
  const lateLog = writes().find(w => w.method === "POST" && w.path.indexOf(LOG_ID) > 0);
  assert.strictEqual(lateLog.body.fields.Who, "Person B");
  assert.strictEqual(lateLog.body.fields.At, "2026-09-08T06:00:00.000Z");
  pass("who and when are taken at the moment of the tap, on the counter and on the log line alike");

  /* the queue comes back out of storage anyone can edit */
  mem.cw_stationq = JSON.stringify({
    "900|cut": { id: "900", stage: "cut", value: 2, who: "Person A", at: "2026-09-08T06:00:00.000Z",
                 job: "R5303", type: "TG",
                 fields: { Total: 999, Customer: "Not this", Active: "No", Title: "R5303|TG" } },
    "junk": { stage: "cut", value: 1 },
    "nostage": { id: "901", stage: "tough", value: 1 },
    "novalue": { id: "901", stage: "cut", value: "wat" }
  });
  mem.cw_stationlogq = JSON.stringify({
    "a": { fields: { Title: "R5303", Station: "Glass", GlassType: "TG", Stage: "cut", From: 0, To: 2,
                     Who: "Person A", At: "2026-09-08T06:00:00.000Z", Customer: "Not this", Total: 9 } },
    "b": { fields: { Stage: "cut" } }
  });
  const sbW = newStation();
  const SW = code => vm.runInContext(code, sbW);
  await settle(20);
  assert.strictEqual(SW("Object.keys(QUEUE).join(',')"), "900|cut",
    "an entry with no item id, a stage that does not exist, or no number is dropped");
  assert.strictEqual(SW("QUEUE['900|cut'].value"), 2);
  assert.strictEqual(SW("'fields' in QUEUE['900|cut']"), false, "and a smuggled job fact does not survive");
  assert.strictEqual(SW("Object.keys(LOGQ).join(',')"), "a", "a log line with no job is dropped too");
  assert.strictEqual(SW("Object.keys(LOGQ['a'].fields).sort().join(',')"),
    "At,From,GlassType,Stage,Station,Title,To,Who",
    "and the one that survives is rebuilt as the eight columns of the log list");
  pass("both owed-write queues are filtered to their own columns on the way out of localStorage");

  ITEMS[0].fields.Cut = 0;
  SW("SITEID = " + JSON.stringify(FSITE) + "; READY = true; ITEMS = " + JSON.stringify(ITEMS) + ";");
  SW("PERSON = { name: 'Person A', stages: ['cut'], pin: '' };");
  reset();
  await SW("flushQueue()");
  await settle(80);
  const sent900 = writes().find(w => w.method === "PATCH");
  assert.deepStrictEqual(Object.keys(sent900.body).sort(), ["Cut", "CutAt", "CutBy", "DoneAt", "DoneBy"],
    "so the PATCH is five keys, whatever was in localStorage");
  assert.strictEqual(ITEMS[0].fields.Cut, 2, "and the owed write really went");
  SW("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("owed writes survive a reload of the tablet and go out on the next visit, filtered again");

  /* ---- PIN ---- */
  S("QUEUE = {}; LOGQ = {}; switchPerson();");
  assert.strictEqual(S("PERSON"), null);
  assert.strictEqual(mem.cw_person, undefined, "Switch person forgets who it was");
  assert.ok(EL["#board"].innerHTML.indexOf("Who are you?") > 0, "and the picker comes back");
  pass("Switch person clears the name, the stored one with it, and asks again");

  S("PINFOR = PEOPLE.find(p => p.name === 'Person B'); PINTYPED = ''; PINBAD = false; render();");
  let pin = EL["#board"].innerHTML;
  assert.ok(pin.indexOf("Enter your PIN") > 0 && pin.indexOf("Person B") > 0);
  assert.ok(pin.indexOf('data-pin="1"') > 0 && pin.indexOf('data-pin="ok"') > 0 &&
            pin.indexOf('data-pin="back"') > 0, "a numeric pad with a way back");
  S("PINTYPED = '9999'; submitPin();");
  assert.strictEqual(S("PERSON"), null, "a wrong PIN lets nobody in");
  assert.strictEqual(S("PINBAD"), true);
  assert.strictEqual(S("PINTYPED"), "", "and clears what was typed");
  assert.ok(EL["#board"].innerHTML.indexOf("Try again") > 0, "saying so, with no lockout");
  assert.ok(EL["#board"].innerHTML.indexOf("shake") > 0);
  pass("a person with a PIN is asked for it, and a wrong one shakes and says Try again");

  S("PINTYPED = '1234'; submitPin();");
  assert.strictEqual(S("PERSON.name"), "Person B");
  assert.strictEqual(S("PINFOR"), null);
  assert.deepStrictEqual(S("PERSON.stages"), ["hotmelt", "glazed"]);
  bh = boardHtml();
  assert.ok(/data-stage="hotmelt" data-act="1">\+<\/button>/.test(bh), "now hotmelt is theirs");
  assert.ok(/data-stage="cut"[^>]*disabled/.test(bh), "and cutting is not");
  pass("the right PIN lets that person in, and the steppers follow the stages they hold");

  reset();
  S("tap('900', 'hotmelt', 2)");
  await settle(80);
  const hp = writes().find(w => w.method === "PATCH");
  assert.deepStrictEqual(Object.keys(hp.body).sort(), ["DoneAt", "DoneBy", "Hotmelt", "HotmeltAt", "HotmeltBy"]);
  assert.strictEqual(hp.body.HotmeltBy, "Person B");
  const hl = writes().find(w => w.method === "POST" && w.path.indexOf(LOG_ID) > 0);
  assert.strictEqual(hl.body.fields.Stage, "hotmelt");
  assert.strictEqual(hl.body.fields.Who, "Person B");
  pass("the second person's own stage writes its own columns and its own log line");

  /* ---- the lock ---- */
  S("LAST_TAP = Date.now() - 601000; lockIfIdle();");
  assert.strictEqual(S("PERSON"), null, "ten quiet minutes and the name means nobody");
  assert.ok(EL["#board"].innerHTML.indexOf("Who are you?") > 0);
  S("PERSON = PEOPLE.find(p => p.name === 'Person A'); LAST_TAP = Date.now() - 60000; lockIfIdle();");
  assert.ok(S("PERSON") !== null, "one quiet minute is not ten");
  S("LAST_TAP = Date.now() - 601000;");
  ITEMS[0].fields.Cut = 0;
  await S("readList()");
  S("touch(); tap('900', 'cut', 1);");
  await settle(60);
  S("lockIfIdle()");
  assert.ok(S("PERSON") !== null, "and a tap pushes the lock back");
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("the chosen name locks itself after ten minutes without a tap, and every tap postpones it");

  /* AMENDMENT 9: a stepper already at the total, and opening a card, are both
     somebody working the screen - the lock must not fire under their hand */
  ITEMS[0].fields.Cut = 6;
  S("QUEUE = {}; LOGQ = {}; PERSON = PEOPLE.find(p => p.name === 'Person A');");
  await S("readList()");
  S("LAST_TAP = Date.now() - 599000;");
  reset();
  S("tap('900', 'cut', 1)");
  await settle(40);
  assert.strictEqual(writes().length, 0, "nothing to write: the counter was already at the total");
  assert.ok(S("Date.now() - LAST_TAP") < 5000, "but the lock was pushed back all the same");
  S("LAST_TAP = Date.now() - 599000; OPEN = {}; render();");
  EL["#board"].fire("click", { target: { dataset: { toggle: "R5303" } } });
  assert.strictEqual(S("!!OPEN.R5303"), true, "the card opened");
  assert.ok(S("Date.now() - LAST_TAP") < 5000, "and that counted as activity too");
  S("OPEN = { R5303: 1 };");
  pass("a tap at the clamp and opening a card both push the ten-minute lock back");

  /* AMENDMENT 11: a stored stamp from the future is not a licence */
  mem.cw_person = JSON.stringify({ name: "Person B", at: Date.now() + 86400000 });
  S("PERSON = null; LAST_TAP = 0; loadPerson();");
  assert.ok(S("LAST_TAP") <= Date.now(), "a stamp from tomorrow is clamped to now");
  S("LAST_TAP = LAST_TAP - 601000; lockIfIdle();");
  assert.strictEqual(S("PERSON"), null, "so the lock can still reach it");
  pass("a clock change or an edited stamp cannot hold the tablet unlocked past ten minutes");

  mem.cw_person = JSON.stringify({ name: "Person B", at: Date.now() - 601000 });
  S("PERSON = null; loadPerson();");
  assert.strictEqual(S("PERSON"), null, "a stale name in storage is not a name");
  mem.cw_person = JSON.stringify({ name: "Person B", at: Date.now() });
  S("PERSON = null; loadPerson();");
  assert.strictEqual(S("PERSON.name"), "Person B", "a fresh one is picked up again");
  mem.cw_person = JSON.stringify({ name: "Person B", at: Date.now(), stages: ["cut", "hotmelt", "glazed"] });
  S("PERSON = null; loadPerson();");
  assert.deepStrictEqual(S("PERSON.stages"), ["hotmelt", "glazed"],
    "and the stages come from the LIST, never from storage anybody can edit");
  pass("the chosen name survives a reload only while it is fresh, and its stages always come from SharePoint");

  /* ---- the ten-second poll on the tablet ---- */
  ITEMS = [item({ Title: "R5303|TG", Job: "R5303", Customer: "Customer One", GlassType: "TG", Total: 6,
                  Seq: 1, Active: "Yes", Cut: 1 }, "900")];
  S("QUEUE = {}; LOGQ = {}; TOKEN = null; ITEMS = [];");
  await S("readList()");
  assert.ok(S("TOKEN") !== null);
  DELTA_NEXT[GLASS_ID] = [[{ id: "900", fields: { Title: "R5303|TG", Job: "R5303", Customer: "Customer One",
    GlassType: "TG", Total: 6, Seq: 1, Active: "Yes", Cut: 4 } }]];
  reset();
  await S("pollList()");
  assert.strictEqual(REQ.filter(r => r.path.indexOf("/delta") > 0).length, 1, "one request per poll");
  assert.ok(boardHtml().indexOf("4 / 6") > 0, "and the other tablet's tap is on this one's screen");
  pass("the tablet polls by delta and merges what someone else changed into its own board");

  DELTA_FAIL = 410; DELTA_FAIL_LEFT = 1; reset();
  await S("pollList()");
  assert.strictEqual(S("READY"), true, "a stale token never takes the board away");
  assert.ok(REQ.some(r => r.path.indexOf("/items/delta?expand") > 0),
    "it throws the token away and enumerates the list again from scratch");
  assert.ok(S("TOKEN") !== null, "and comes back with a new token");
  DELTA_FAIL = 0; DELTA_FAIL_LEFT = 0;
  pass("a 410 resync on the tablet is one fresh enumeration, with nothing shown to anybody");

  /* a tenant or a list where delta is refused outright: the tablet still works,
     it simply polls the long way round */
  DELTA_FAIL = 400; DELTA_FAIL_LEFT = 99;
  S("TOKEN = null;"); reset();
  assert.strictEqual(await S("readList()"), true, "the board still reads");
  assert.strictEqual(S("TOKEN"), null, "with no token, because there is no delta to be had");
  assert.ok(REQ.some(r => r.path.indexOf("/items?expand") > 0), "it fell back to the plain read");
  reset();
  await S("pollList()");
  assert.ok(REQ.some(r => r.path.indexOf("/items?expand") > 0), "and the poll keeps doing that");
  assert.strictEqual(S("READY"), true);
  DELTA_FAIL = 0; DELTA_FAIL_LEFT = 0;
  S("TOKEN = null;"); await S("readList()");
  pass("a list that will not serve a delta at all falls back to a plain read and keeps working");

  /* ---- drawing only what moved ---- */
  S("OPEN = { R5303: 1 }; render();");
  const nodeBefore = S("NODES['R5303']");
  S("ITEMS[0].fields.Cut = 5; render();");
  assert.strictEqual(S("NODES['R5303']") === nodeBefore, true,
    "the card element itself is not thrown away and rebuilt");
  assert.ok(boardHtml().indexOf("5 / 6") > 0, "only its contents were redrawn");
  assert.ok(boardHtml().indexOf("stepl") > 0, "and the card that was open is still open");
  pass("a refresh redraws the cards that moved and keeps the nodes - so an open card stays open");

  const seen = S("JSON.stringify(BOARD_PREV.map(g => g.job))");
  S("render()");
  assert.strictEqual(S("JSON.stringify(BOARD_PREV.map(g => g.job))"), seen);
  assert.strictEqual(S("NODES['R5303']") === nodeBefore, true, "a refresh that found nothing changes nothing");
  pass("a poll that found nothing new leaves every node exactly where it was");

  S("ITEMS[0].fields.Cut = 6; ITEMS[0].fields.Hotmelt = 6; ITEMS[0].fields.Glazed = 6; render();");
  assert.strictEqual(S("BOARD_PREV[0].finished"), true);
  assert.strictEqual(S("FIN.kids.length"), 1, "the finished card moved into the Finished group");
  assert.strictEqual(S("FINHEAD.hidden"), false);
  assert.ok(/Finished/.test(S("FINHEAD.textContent")));
  assert.strictEqual(S("FIN.hidden"), true, "which is collapsed until somebody taps it");
  S("FINOPEN = true; render();");
  assert.strictEqual(S("FIN.hidden"), false);
  S("FINOPEN = false;");
  pass("a job whose bars are all full drops into a Finished group, collapsed, with a count on it");

  /* ---- the tablet's theme ---- */
  delete mem.cw_stationtheme;
  assert.strictEqual(S("themeNow()"), "dark", "the tablet starts dark, as the owner asked");
  S("applyTheme('light')");
  assert.strictEqual(mem.cw_stationtheme, "light", "the choice is the device's, and it is remembered");
  assert.strictEqual(S("themeNow()"), "light");
  assert.strictEqual(document.documentElement.dataset.theme, "light");
  S("applyTheme('dark')");
  assert.strictEqual(document.documentElement.dataset.theme, "dark");
  const gs = src("glass.html");
  assert.ok(/<html lang="en" data-theme="dark">/.test(gs), "and the page itself opens dark");
  assert.ok(/:root\[data-theme="light"\]/.test(gs), "with the light set behind the switch");
  assert.ok(/id="themebtn"/.test(gs));
  pass("the tablet is dark by default with a switch in the header, saved on the device");

  assert.ok(/id="switchbtn"/.test(gs) && />Switch person</.test(gs), "the Switch person button is in the header");
  assert.ok(gs.indexOf("min-height:64px") > 0, "the picker's buttons are 64 px");
  assert.ok(gs.indexOf(".pk { min-height:56px") > 0, "and the PIN pad's are 56");
  assert.ok(/\.sbtn \{ width:48px; height:48px/.test(gs), "the steppers stay at 48");
  assert.ok(/@media \(min-width:700px\)\{[\s\S]*grid-template-columns:1fr 1fr/.test(gs),
    "two columns of cards from 700 px");
  assert.ok(/\.bar \{ display:grid; gap:4px \}/.test(gs), "and the bars are stacked at every width");
  pass("the tablet's layout is the stacked one, two columns from 700 px, everything thumb-sized");

  /* ---- errors on the tablet ---- */
  ITEMS = [item({ Title: "R5303|TG", Job: "R5303", Customer: "Customer One", GlassType: "TG", Total: 6,
                  Seq: 1, Active: "Yes", Cut: 1 }, "900")];
  S("QUEUE = {}; LOGQ = {}; SITEID = " + JSON.stringify(FSITE) + "; TOKEN = null; DELTA_OFF = 0;");
  await S("readList()");
  assert.strictEqual(S("READY"), true);
  CW.listDelta = async () => { throw new Error("Failed to fetch"); };
  CW.listItems = async () => { throw new Error("Failed to fetch"); };
  await S("readList()");
  assert.strictEqual(S("READY"), true, "the board is still on screen");
  assert.strictEqual(S("SOFT"), "cannot reach SharePoint — retrying");
  const boom = m => { CW.listDelta = async () => { throw new Error(m); };
                      CW.listItems = async () => { throw new Error(m); }; };
  boom("interaction_required");
  await S("readList()");
  assert.strictEqual(S("READY"), false);
  assert.ok(EL["#board"].innerHTML.indexOf("Tap Sign out, then Sign in again") > 0,
    "an expired sign-in says what to do, not ask the office for a permission");
  boom("GET /lists/x/items -> 404 {\"error\":{\"code\":\"itemNotFound\"}}");
  await S("readList()");
  assert.ok(EL["#board"].innerHTML.indexOf("Ask the office") > 0, "a genuine 404 does clear it");
  CW.listDelta = realDelta; CW.listItems = realItems;
  S("SITEID = " + JSON.stringify(FSITE) + "; TOKEN = null;");
  await S("readList()");
  pass("the tablet keeps its board through a passing failure, and only a 404 or an expired sign-in takes it");

  /* AMENDMENT 6: a tablet that cannot show the board always has something to
     tap, and a people list that failed once is retried on the ten-second clock */
  S("READY = false; PROBLEM = ''; PEOPLE_READ = false; render();");
  assert.ok(EL["#board"].innerHTML.indexOf("Reading the board") > 0);
  assert.ok(EL["#board"].innerHTML.indexOf('id="again"') > 0,
    "Try again is offered even when the trouble has no name yet");
  S("PROBLEM = 'reauth'; render();");
  assert.ok(EL["#board"].innerHTML.indexOf('id="reauth"') > 0, "an expired sign-in gets its own button");
  assert.ok(EL["#board"].innerHTML.indexOf('id="again"') < 0, "and only that one");
  S("PROBLEM = ''; PEOPLE = []; PEOPLE_READ = false; SITEID = " + JSON.stringify(FSITE) + ";");
  reset();
  await S("tickOnce()");
  assert.ok(REQ.some(r => r.path.indexOf(PEOPLE_ID) > 0), "the ten-second clock retries the people read");
  assert.strictEqual(S("PEOPLE_READ"), true, "and this time it answered");
  assert.deepStrictEqual(S("PEOPLE.map(p => p.name)"), ["Person A", "Person B", "Person C"]);
  reset();
  await S("tickOnce()");
  assert.ok(!REQ.some(r => r.path.indexOf(PEOPLE_ID) > 0), "and having answered, it is not asked again");
  pass("Try again is always offered, and a people list that failed once is retried on the poll's own clock");

  /* ---- AMENDMENT 2: the second blocker ----
     A queued tap holds a SharePoint item id, and an item id only means
     anything in the list it came from. If the floor's lists move between the
     tap and the write, PATCHing that number against the new list lands on
     whichever row happens to carry it - almost certainly another job - and the
     log then records that as work somebody did on it. The entry is stamped
     with its site and re-matched by Title, which is unique and does not move. */
  ITEMS = [item({ Title: "R5303|TG", Job: "R5303", Customer: "Customer One", GlassType: "TG",
                  Total: 6, Seq: 1, Active: "Yes", Cut: 1 }, "9500"),
           item({ Title: "R7777|DG", Job: "R7777", Customer: "Customer Two", GlassType: "DG",
                  Total: 4, Seq: 2, Active: "Yes", Cut: 0 }, "900")];
  OWNITEMS = [item({ Title: "R5303|TG", Job: "R5303", Customer: "Customer One", GlassType: "TG",
                     Total: 6, Seq: 1, Active: "Yes", Cut: 1 }, "900")];
  LOGITEMS = []; OWNLOG = [];
  S("QUEUE = {}; LOGQ = {}; TOKEN = null; DELTA_OFF = 0; OPEN = { R5303: 1 };");
  S("SITEID = " + JSON.stringify(SITE) + ";");        // the tablet is on the fallback
  S("PERSON = { name: 'Person A', stages: ['cut', 'hotmelt', 'glazed'], pin: '' };");
  S("PEOPLE_READ = true; READY = true; PROBLEM = ''; SOFT = ''; LAST_TAP = Date.now();");
  S("ITEMS = " + JSON.stringify(OWNITEMS) + "; render();");
  S("tap('900', 'cut', 2)");
  assert.strictEqual(S("QUEUE['900|cut'].site"), SITE, "the tap is stamped with the site it was made in");
  assert.strictEqual(S("QUEUE['900|cut'].job"), "R5303");
  S("flushing = true;");                              // hold it until the lists have moved
  await settle(20);
  /* the lists move: the same job is item 9500 over there, and 900 is another job */
  S("SITEID = " + JSON.stringify(FSITE) + "; TOKEN = null; flushing = false;");
  await S("readList()");
  reset();
  await S("flushQueue()");
  await settle(80);
  const moved = writes().filter(w => w.method === "PATCH");
  assert.strictEqual(moved.length, 1, "one write");
  assert.ok(moved[0].path.indexOf("/items/9500/fields") > 0,
    "against the row that IS this job in the new list, found again by its Title");
  assert.ok(moved[0].path.indexOf("/items/900/fields") < 0,
    "and never against the number the tap was queued with, which is another job over here");
  assert.strictEqual(ITEMS.find(x => x.id === "9500").fields.Cut, 3, "the right row moved");
  assert.strictEqual(ITEMS.find(x => x.id === "900").fields.Cut, 0, "and the other job was not touched");
  const movedLog = writes().find(w => w.method === "POST" && w.path.indexOf(LOG_ID) > 0);
  assert.strictEqual(movedLog.body.fields.Title, "R5303", "the log names the job the floor meant");
  assert.strictEqual(movedLog.body.fields.From + "->" + movedLog.body.fields.To, "1->3");
  pass("a tap queued in one site is re-matched by Title after a move, never written by item id");

  /* and a queued tap whose row is not in the new list at all is dropped */
  S("QUEUE = {}; LOGQ = {};");
  S("QUEUE['4242|cut'] = { id: '4242', stage: 'cut', value: 2, who: 'Person A', " +
    "at: '2026-09-08T07:00:00.000Z', job: 'R9999', type: 'TG', site: " + JSON.stringify(SITE) + ", from: 0, err: 0 };");
  reset();
  await S("flushQueue()");
  await settle(60);
  assert.strictEqual(writes().length, 0, "nothing was written anywhere");
  assert.strictEqual(S("Object.keys(QUEUE).length"), 0, "and the entry is gone rather than retried for ever");
  assert.strictEqual(S("Object.keys(LOGQ).length"), 0, "with no log line invented for it");
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("a queued tap whose row did not come across is dropped, not written onto whatever row has that id");

  /* ---- AMENDMENT 3: the tablet never writes without a resolved site ---- */
  ITEMS = [item({ Title: "R5303|TG", Job: "R5303", Customer: "Customer One", GlassType: "TG",
                  Total: 6, Seq: 1, Active: "Yes", Cut: 0 }, "900")];
  S("QUEUE = {}; LOGQ = {}; SITEID = " + JSON.stringify(FSITE) + "; TOKEN = null; DELTA_OFF = 0;");
  await S("readList()");
  S("SITEID = null;");                                // the window after a 404 forgot the site
  reset();
  S("tap('900', 'cut', 1)");
  await settle(80);
  assert.strictEqual(REQ.length, 0, "not one request left the tablet");
  assert.ok(REQ.every(r => r.path.indexOf("/drive") < 0 && r.path.indexOf("/workbook") < 0),
    "and above all nothing that looks like the workbook: no /drive/root/children");
  assert.strictEqual(S("Object.keys(QUEUE).length"), 1, "the tap is still owed");
  assert.ok(S("retryT !== null"), "with a retry due");
  S("QUEUE['900|cut'].site = " + JSON.stringify(FSITE) + ";");
  S("clearTimeout(retryT); retryT = null; SITEID = " + JSON.stringify(FSITE) + ";");
  reset();
  await S("flushQueue()");
  await settle(80);
  assert.strictEqual(ITEMS[0].fields.Cut, 1, "and it goes the moment there is a site again");
  S("if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("a tap made while the site is unresolved waits, and never sends the tablet at the workbook");

  /* AMENDMENT 4, on the tablet: a refusal does not throw the site away */
  S("SITEID = " + JSON.stringify(FSITE) + ";");
  S("trouble(new Error('GET /x -> 403 accessDenied'))");
  assert.strictEqual(S("SITEID"), FSITE, "a refused call keeps the site the tablet is reading");
  S("trouble(new Error('Failed to fetch'))");
  assert.strictEqual(S("SITEID"), FSITE, "and so does a workshop with no wifi");
  S("trouble(new Error('GET /x -> 404 itemNotFound'))");
  assert.strictEqual(S("SITEID"), null, "only a genuine 404 drops it");
  S("SITEID = " + JSON.stringify(FSITE) + "; PROBLEM = ''; SOFT = ''; READY = true;");
  forget(); CW._setStationSite(FSITE);
  pass("the tablet keeps the site it is reading through a refusal or a dead connection, and drops it on a 404");

  /* ---- AMENDMENT 6: no message anywhere names a site ---- */
  const SITENAMES = ["Floor stations", "FloorStations", "ProductionProgress", "Production Progress"];
  const officeWords = [A("STATION_SITE_MISSING"), A("STATION_LIST_MISSING"), A("STATION_LOG_MISSING"),
                       A("STATION_NEED_CONSENT"), A("STATION_SIGN_IN_AGAIN"), A("STATION_UNREACHABLE"),
                       A("STATION_CHECKING")].join(" | ");
  const tabletWords = ["reauth", "consent", "people", "list", "site", ""].map(pb => {
    S("PROBLEM = " + JSON.stringify(pb) + ";");
    return S("words()");
  }).join(" | ");
  SITENAMES.forEach(nm => {
    assert.ok(officeWords.indexOf(nm) < 0, "the office must not name the site: " + nm);
    assert.ok(tabletWords.indexOf(nm) < 0, "and neither must the tablet: " + nm);
  });
  assert.ok(/Glass station/.test(officeWords) && /Glass station/.test(tabletWords),
    "the LIST is still named, because that is the thing somebody has to go and make");
  assert.ok(/floor/i.test(officeWords) && /floor/i.test(tabletWords),
    "and it is called the floor's site, which is true of either arrangement");
  S("PROBLEM = '';");
  pass("neither page's wording names a SharePoint site, so the interim arrangement shows nowhere on screen");

  /* AMENDMENT 8: the tablet stops asking a refusing list for a delta too */
  S("QUEUE = {}; LOGQ = {}; DELTA_OFF = 0; TOKEN = null;");
  DELTA_FAIL = 400; DELTA_FAIL_LEFT = 1;
  reset();
  await S("readList()");
  assert.ok(S("DELTA_OFF") > 0, "the refusal marked delta off");
  assert.strictEqual(S("READY"), true, "and the board read the plain way regardless");
  DELTA_FAIL = 0; DELTA_FAIL_LEFT = 0;
  reset();
  await S("pollList()");
  assert.strictEqual(REQ.filter(r => r.path.indexOf("/items/delta") > 0).length, 0,
    "the next poll does not ask it for one");
  S("DELTA_OFF = Date.now() - 400000;");
  reset();
  await S("readList()");
  assert.ok(REQ.some(r => r.path.indexOf("/items/delta") > 0), "after five minutes it tries again");
  assert.strictEqual(S("DELTA_OFF"), 0, "and having been served, stops marking it off");
  pass("the tablet marks a list that refuses delta off for five minutes rather than asking every ten seconds");

  FLISTS = FLISTS.filter(l => l.displayName !== "Station people");
  forget(); S("SITEID = null;");
  assert.strictEqual(await S("readPeople()"), false);
  assert.ok(EL["#board"].innerHTML.indexOf("Station people") > 0,
    "a missing list is named, so the office knows which one to make");
  FLISTS.push({ id: PEOPLE_ID, displayName: "Station people" });
  forget(); S("SITEID = " + JSON.stringify(FSITE) + "; PROBLEM = '';");
  await S("readPeople()");
  pass("a list that has not been made yet is named plainly rather than thrown");

  S("BUILD_NOW = null; QUEUE = {}; LOGQ = {};");
  RELOADED = false;
  await S("checkBuild()");
  assert.strictEqual(S("BUILD_NOW"), "20260908-0951", "the first look is only a baseline");
  BUILD_SERVED = "20260908-1200";
  await S("checkBuild()");
  assert.strictEqual(RELOADED, true, "a changed build reloads the tablet");
  RELOADED = false;
  BUILD_SERVED = "20260908-1300";
  S("LOGQ = { a: { key: 'a', err: 0, fields: { Title: 'R1' } } };");
  await S("checkBuild()");
  assert.strictEqual(RELOADED, false, "but never while a log line is still owed either");
  S("QUEUE = {}; LOGQ = {}; if (retryT) { clearTimeout(retryT); retryT = null; }");
  pass("the tablet reloads itself when the build changes, and never with anything still owed");

  /* ================= 16. what the station page cannot do ================= */
  const stationSrc = src("station.js"), glassSrc = src("glass.html"), coreSrc = src("station-core.js");
  ["setFill", "clearFill", "setValues", "appendLog", "saveProgress", "moveJobRow", "batchWrite",
   "downloadWorkbook", "/workbook", "parseWorkbook", "ExcelJS"]
    .forEach(bad => assert.ok(stationSrc.indexOf(bad) < 0, "station.js must not mention " + bad));
  assert.ok(stationSrc.indexOf('"DELETE"') < 0 && stationSrc.indexOf("listDelete") < 0,
    "and it has no way to delete anything");
  ["exceljs", "parser.js", "app.js", "checkpoints.js", "export.js", "index.html"]
    .forEach(bad => assert.ok(glassSrc.toLowerCase().indexOf(bad) < 0, "glass.html must not load or link " + bad));
  assert.ok(/src="graph\.js/.test(glassSrc) && /src="station-core\.js/.test(glassSrc) &&
            /src="station\.js/.test(glassSrc) && /msal-browser/.test(glassSrc),
    "it loads msal, graph.js, station-core.js and station.js, and that is the lot");
  pass("the station page cannot reach the workbook: no exceljs, no parser, no app.js, no /workbook, no delete");

  ["setFill", "clearFill", "setValues", "appendLog", "saveProgress", "moveJobRow", "batchWrite", "/workbook"]
    .forEach(bad => assert.ok(coreSrc.indexOf(bad) < 0, "station-core.js must not mention " + bad));
  assert.ok(!/@[a-z0-9-]+\.(com|ie|net|org)/i.test(stationSrc + coreSrc + glassSrc),
    "no real email address or domain in any of the new files");
  assert.ok(!/@[a-z0-9-]+\.(com|ie|net|org)/i.test(src("test_station.js")),
    "and none in this file either - every address here is example.test");
  pass("nothing new in the repo names a real person, address or domain");

  const appSrc = src("app.js");
  assert.ok(appSrc.indexOf("CW.listAdd(ST.LOG_LIST") < 0 && appSrc.indexOf("listPatch(ST.LOG_LIST") < 0,
    "the master never writes a line of the floor's log");
  assert.ok(appSrc.indexOf("listAdd(ST.PEOPLE_LIST") < 0 && appSrc.indexOf("listPatch(ST.PEOPLE_LIST") < 0,
    "nor a row of the people list");
  pass("the office reads the log and the people list and has no code path that writes either");

  /* ================= 17. the whole run, end to end ================= */
  const touchedBook = ALLREQ.filter(r => /\/workbook|\/content|\/versions|createSession|\/drive|\/range\(|worksheets/.test(r.path));
  assert.deepStrictEqual(touchedBook.map(r => r.method + " " + r.path), [],
    "over every request of the whole run, not one touched the workbook");
  /* the interim arrangement puts the lists in the workbook's SITE, which is
     not the workbook: the difference is that no path ever reaches its drive */
  assert.ok(ALLREQ.some(r => r.path.indexOf(SITE_BY_PATH) >= 0),
    "the workbook's site really was looked up by path during this run");
  assert.ok(ALLREQ.every(r => r.path.indexOf("/drive") < 0),
    "and not one request of the whole run went near a drive");
  assert.strictEqual(ALLREQ.filter(r => r.method === "DELETE").length, 0,
    "and not one was a delete, from either side");
  pass("across every test above, the workbook was never read or written and nothing was ever deleted");

  const bodies = ALLREQ.filter(r => r.body).map(r => JSON.stringify(r.body)).join(" ");
  ["eircode", "Eircode", "phone", "Phone", "county", "County", "price", "Price", "comment", "Comment"]
    .forEach(w => assert.ok(bodies.indexOf(w) < 0, "no request body ever carried " + w));
  assert.ok(bodies.indexOf("Cork") < 0, "not even the area the job is in");
  assert.ok(ALLREQ.length > 150, "and that is over a real number of requests, not an empty log");
  pass("nothing personal was ever sent to the floor: no phone, eircode, county, price or comment");

  const floorWrites = ALLREQ.filter(r => r.method === "PATCH" && r.path.indexOf(GLASS_ID) > 0 &&
    r.body && ST.FLOOR_FIELDS.some(k => k in r.body));
  assert.ok(floorWrites.length > 0);
  floorWrites.forEach(w => Object.keys(w.body).forEach(k =>
    assert.ok(ST.FLOOR_FIELDS.indexOf(k) >= 0, "the floor sent " + k + ", which is not its column")));
  const feedPatches = ALLREQ.filter(r => (r.method === "PATCH" || r.method === "POST") &&
    r.path.indexOf(GLASS_ID) > 0 && r.body && !ST.FLOOR_FIELDS.some(k => k in (r.body.fields || r.body)));
  feedPatches.forEach(w => Object.keys(w.body.fields || w.body).forEach(k =>
    assert.ok(ST.FLOOR_FIELDS.indexOf(k) < 0, "the feeder sent " + k + ", which is the floor's column")));
  pass("every write of the whole run stayed on its own side of the list: floor columns or job facts, never both");

  console.log("\n" + n + " checks passed");
  process.exit(0);
})().catch(e => { console.error("FAIL", e); process.exit(1); });
