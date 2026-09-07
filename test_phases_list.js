/* Offline test of the hand-set phases: the SharePoint list they live in, the
   Graph helpers that read and write it, and what the dashboard does with them.

   The point of the feature is that a phase set by hand is shared with everyone
   WITHOUT the production workbook being touched, so the sharpest assertions in
   here are about the request log: setting a phase, clearing one and reading
   them all make list requests and nothing else. The one workbook request the
   whole flow makes is the Dashboard Log line that every dashboard edit already
   writes - the shared audit trail, its own long-standing feature - and even
   that is checked to be the Dashboard Log sheet and nothing else. No request
   ever names Production, downloads the file, or writes a range anywhere else.

   Graph is a fake fetch() over an in-memory list and a two-sheet workbook;
   nothing leaves the box.
   Run: node test_phases_list.js                                             */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, innerWidth: 1280, innerHeight: 800,
                  addEventListener() {}, removeEventListener() {} };
global.performance = { now: () => Date.now() };
global.ExcelJS = ExcelJS;

function stubEl() {
  let html = "";
  const e = {
    style: { setProperty() {} }, dataset: {}, textContent: "", value: "", disabled: false,
    hidden: false, className: "", kids: [],
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); }, remove() {}, addEventListener() {}, removeEventListener() {},
    setAttribute() {}, getAttribute() { return null; }, focus() {}, blur() {}, setSelectionRange() {},
    getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => stubEl(), querySelectorAll: () => []
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; } });
  return e;
}
const EL = {};
const el = sel => EL[sel] || (EL[sel] = stubEl());
global.document = {
  documentElement: stubEl(), body: stubEl(), head: stubEl(), createElement: () => stubEl(),
  activeElement: null, title: "Costello Production",
  querySelector: el, querySelectorAll: () => [], addEventListener() {}, removeEventListener() {}
};

/* ---------- the fake Graph ----------
   A SharePoint site with two lists, and a workbook with two dashboard sheets.
   Every request is recorded, method and path, so a test can prove what the
   feature did and - far more importantly - what it did not. */
const SITE = "costellowindowsie.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";
const PHASE_LIST_ID = "list-dashboard-phases";
const LISTS_PATH = "/sites/" + SITE + "/lists";
const G = "https://graph.microsoft.com/v1.0";

let LISTS = [{ id: "list-site-assets", displayName: "Site Assets" },
             { id: PHASE_LIST_ID, displayName: "Dashboard phases" }];
let ITEMS = [];                    // the items of the phases list
let PAGE = 999;                    // how many items one page of the read holds
let NEXTID = 100;
let FAIL_WRITE = 0;                // n list writes to refuse with a 403 (no retry sleeps)
let UNIQUE_TITLE = false;          // the list's "enforce unique values" rule on Title
let INJECT_ON_POST = null;         // the other browser, arriving between our read and our POST
let NEXTLINK_ORIGIN = null;        // where @odata.nextLink claims to live

const REQ = [];                    // every request since the last reset()
const ALLREQ = [];                 // and every request of the whole run, never cleared
const LOGSHEET = [["When", "Who", "Job", "What changed", "From", "To"]];
const ok = body => ({ status: 200, body: body });

function item(job, phase, who, at) {
  return { id: String(NEXTID++), fields: { Title: job, Phase: phase, PhaseName: PHASES_WORDS[phase],
                                           SetBy: who || "boss@example.test", SetAt: at || "2026-09-06T09:00:00.000Z" } };
}
const PHASES_WORDS = ["In office", "Sent to floor", "Cutting", "In fabrication",
                      "In glazing", "Quality check", "Fitted / delivered"];

function routeList(method, path, body) {
  const rest = path.slice(LISTS_PATH.length);
  if (method === "GET" && rest.indexOf("?$select=id,displayName") === 0) return ok({ value: LISTS });
  const mi = /^\/([^/?]+)\/items(.*)$/.exec(rest);
  if (!mi) return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + path } } };
  const id = mi[1], tail = mi[2];
  if (id !== PHASE_LIST_ID) return { status: 404, body: { error: { code: "itemNotFound" } } };
  if (method === "GET") {
    const m = /[?&]page=(\d+)/.exec(tail);
    const page = m ? Number(m[1]) : 1;
    const from = (page - 1) * PAGE, slice = ITEMS.slice(from, from + PAGE);
    const out = { value: slice.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) };
    if (from + PAGE < ITEMS.length)
      out["@odata.nextLink"] = (NEXTLINK_ORIGIN || G) + LISTS_PATH + "/" + id +
                               "/items?expand=fields&page=" + (page + 1);
    return ok(out);
  }
  if (FAIL_WRITE && method !== "GET") { FAIL_WRITE--; return { status: 403, body: { error: { code: "accessDenied" } } }; }
  if (method === "POST" && tail === "") {
    /* the other desk got here first: their item lands between our read and our POST */
    if (INJECT_ON_POST) { const f = INJECT_ON_POST; INJECT_ON_POST = null; f(); }
    const title = String((body && body.fields && body.fields.Title) || "").trim().toUpperCase();
    if (UNIQUE_TITLE && ITEMS.some(x => String(x.fields.Title || "").trim().toUpperCase() === title))
      return { status: 400, body: { error: { code: "invalidRequest",
        message: "The value of the Title column must be unique." } } };
    const made = { id: String(NEXTID++), fields: Object.assign({}, body && body.fields) };
    ITEMS.push(made);
    return ok({ id: made.id, fields: made.fields });
  }
  const mf = /^\/([^/]+)\/fields$/.exec(tail);
  if (method === "PATCH" && mf) {
    const hit = ITEMS.find(x => x.id === mf[1]);
    if (!hit) return { status: 404, body: { error: { code: "itemNotFound" } } };
    Object.assign(hit.fields, body || {});
    return ok(Object.assign({}, hit.fields));
  }
  const md = /^\/([^/]+)$/.exec(tail);
  if (method === "DELETE" && md) {
    const i = ITEMS.findIndex(x => x.id === md[1]);
    if (i < 0) return { status: 404, body: { error: { code: "itemNotFound" } } };
    ITEMS.splice(i, 1);
    return { status: 204, body: "" };
  }
  return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + method + " " + path } } };
}

/* just enough workbook for the Dashboard Log line every edit already writes */
function routeBook(method, path, body) {
  if (path === "/x/workbook/worksheets") return ok({ value: [{ name: "Production" }, { name: "Dashboard Log" }] });
  const m = /^\/x\/workbook\/worksheets\('([^']+)'\)(.*)$/.exec(path);
  if (!m) return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + path } } };
  const sheet = m[1], rest = m[2];
  if (sheet !== "Dashboard Log") return { status: 403, body: { error: { code: "accessDenied", message: "this test forbids " + sheet } } };
  if (rest.indexOf("/usedRange") === 0) return ok({ rowCount: LOGSHEET.length, values: LOGSHEET });
  const mr = /^\/range\(address='([^']+)'\)$/.exec(rest);
  if (method === "PATCH" && mr && body && body.values) { LOGSHEET.push(body.values[0]); return ok({}); }
  return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + method + " " + path } } };
}

function route(method, path, body) {
  if (path.indexOf(LISTS_PATH) === 0) return routeList(method, path, body);
  if (path.indexOf("/x/") === 0) return routeBook(method, path, body);
  return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + path } } };
}
global.fetch = async (url, init) => {
  const path = String(url).replace(G, "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  const rec = { method: init.method, path: path, body: body, headers: init.headers || {} };
  REQ.push(rec); ALLREQ.push(rec);
  const res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status,
           text: async () => (res.body === "" ? "" : JSON.stringify(res.body)),
           arrayBuffer: async () => new ArrayBuffer(0) };
};

/* ---------- load the page's own code, in the page's own order ---------- */
const run = f => vm.runInThisContext(fs.readFileSync(__dirname + "/" + f, "utf8"), { filename: f });
run("parser.js");
run("graph.js");
global.CW = window.CW;
CW._setToken(() => "t");
CW._setFile({ siteId: SITE, base: "/x/workbook", content: "/x/content", meta: "/x" });
CW._setSession("SESSION-1");             // a warm workbook session, to prove list calls skip it
run("checkpoints.js");
global.CP = window.CP;
run("app.js");

const TOASTS = [];
global.toast = (m, isErr) => TOASTS.push({ m: String(m), err: !!isErr });
let WHO = "boss@example.test";
global.whoAmI = () => WHO;

/* ---------- helpers ---------- */
const reset = () => { REQ.length = 0; TOASTS.length = 0; };
const forget = () => { delete mem.cw_listids; CW._resetListIds(); };
const paths = () => REQ.map(r => r.method + " " + r.path);
/* CW_TRACE=1 node test_phases_list.js prints the exact Graph calls a set and a
   clear make, which is what anyone reviewing "does this touch the workbook?"
   actually wants to see. */
const trace = (what, list) => { if (process.env.CW_TRACE) console.log("\n  --- " + what + " ---\n    " + list.join("\n    ") + "\n"); };
const listReqs = () => REQ.filter(r => r.path.indexOf(LISTS_PATH) === 0);
const bookReqs = () => REQ.filter(r => r.path.indexOf(LISTS_PATH) !== 0);
const settle = ms => new Promise(r => setTimeout(r, ms == null ? 30 : ms));
const clearHolds = () => { vm.runInThisContext("PENDING = {}"); savePending(); };
const setPhases = map => vm.runInThisContext("PHASES_SET = " + JSON.stringify(map || {}));

const mkJob = o => Object.assign({
  id: "R0001", cust: "Ann", area: "Cork", eir: "", off: "", colour: "", ph3: "",
  wnd: 0, drs: 0, glass: {}, prods: [], notes: [], sheets: ["Production"], src: {},
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0,
  cp: { win: "", drs: "", glass: {}, prod: {} }
}, o || {});
/* one product with all three sub-columns, plus windows and glass: enough for
   the sheet's own reading to land anywhere between 0 and 5 */
const withWork = (cpProd, extra) => mkJob(Object.assign({
  wnd: 4, drs: 0, glass: { tg: 6 },
  prods: [{ n: "7000 casement", f: 3, s: 2, t: 1, st: [] }],
  cp: { win: "", drs: "", glass: {}, prod: { "7000 casement": cpProd || {} } }
}, extra || {}));
function useJobs(list) {
  global.__jobs = list;
  vm.runInThisContext("ALL = __jobs; BLOCKNAMES = ['Can sell as second hand','Ready, customer won\\'t take'," +
    "'Collect & supply only','Ready to fit','In production']; ALL.blockNames = BLOCKNAMES; " +
    "CHANGES = []; state.sel = null; state.picked = {};");
}

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ---- 1. the scope ---- */
  assert.ok(SCOPES.indexOf("Sites.ReadWrite.All") >= 0, "the list needs its own scope");
  assert.ok(SCOPES.indexOf("Files.ReadWrite.All") >= 0, "and the workbook still needs its own");
  assert.ok(SCOPES.indexOf("User.Read") >= 0);
  pass("Sites.ReadWrite.All is asked for at sign-in, alongside the workbook scope");

  /* ---- 2. the list id: found once, then remembered ---- */
  forget(); reset();
  assert.strictEqual(await CW.listId("Dashboard phases"), PHASE_LIST_ID);
  assert.deepStrictEqual(paths(), ["GET " + LISTS_PATH + "?$select=id,displayName"],
    "one lookup, asking for the two fields it needs and nothing else");
  assert.deepStrictEqual(JSON.parse(mem.cw_listids), { "Dashboard phases": PHASE_LIST_ID },
    "and it is written to cw_listids, the only new key this feature adds");
  reset();
  assert.strictEqual(await CW.listId("Dashboard phases"), PHASE_LIST_ID);
  assert.strictEqual(REQ.length, 0, "the second call asks nobody");
  CW._resetListIds();
  reset();
  assert.strictEqual(await CW.listId("Dashboard phases"), PHASE_LIST_ID);
  assert.strictEqual(REQ.length, 0, "and a fresh page reads it back out of localStorage");
  pass("the list id is looked up once by display name and cached in memory and in localStorage");

  forget(); reset();
  assert.strictEqual(await CW.listId("dashboard PHASES"), PHASE_LIST_ID,
    "SharePoint display names are not case-sensitive and neither is this");
  pass("the display name is matched case-insensitively");

  /* ---- 3. a missing list: null, no cache, no writes ---- */
  const KEEP = LISTS;
  LISTS = [{ id: "list-site-assets", displayName: "Site Assets" }];
  forget(); reset();
  assert.strictEqual(await CW.listId("Dashboard phases"), null);
  assert.strictEqual(await CW.listItems("Dashboard phases"), null, "null, not [] - there is no list to be empty");
  assert.strictEqual(mem.cw_listids, undefined, "a missing list is never cached: it appears the moment it is made");
  assert.strictEqual(await CW.listDelete("Dashboard phases", "R0001"), false);
  await assert.rejects(() => CW.listUpsert("Dashboard phases", "R0001", { Phase: 3 }), /not in SharePoint/);
  assert.strictEqual(REQ.filter(r => r.method !== "GET").length, 0, "and not one write went out");
  LISTS = KEEP; forget();
  pass("a list that is not there answers null, is not cached, and refuses every write");

  /* ---- 4. reading every item, page by page ---- */
  ITEMS = [item("R0001", 3), item("R0002", 5), item("R0003", 1), item("R0004", 6), item("R0005", 2)];
  PAGE = 2;
  reset();
  let got = await CW.listItems("Dashboard phases");
  assert.strictEqual(got.length, 5, "all five, across three pages");
  assert.deepStrictEqual(got.map(x => x.fields.Title), ["R0001", "R0002", "R0003", "R0004", "R0005"]);
  const reads = listReqs().filter(r => r.method === "GET" && r.path.indexOf("/items") > 0);
  assert.strictEqual(reads.length, 3, "one request per page, following @odata.nextLink");
  assert.ok(reads[0].path.indexOf("$top=999") > 0, "and it asks for a big page in the first place");
  assert.ok(reads[0].path.indexOf("expand=fields(select=Title,Phase,PhaseName,SetBy,SetAt)") > 0,
    "asking for the five columns by name");
  PAGE = 999;
  pass("every item is read, following @odata.nextLink to the last page");

  /* ---- 5. list requests never carry the workbook session ---- */
  assert.ok(reads.every(r => !r.headers["workbook-session-id"]),
    "a workbook session header on a /sites/.../lists call is meaningless and can 400");
  pass("the warm workbook session is not sent on list requests");

  /* ---- 6. upsert: PATCH when the job is there, POST when it is not ---- */
  ITEMS = [item("R0001", 3)];
  const wasId = ITEMS[0].id;
  reset();
  let r = await CW.listUpsert("Dashboard phases", "R0001", { Phase: 5, PhaseName: "Quality check", SetBy: WHO, SetAt: "2026-09-07T10:00:00.000Z" });
  assert.strictEqual(r.created, false);
  assert.strictEqual(r.id, wasId);
  const patch = REQ.filter(x => x.method === "PATCH");
  assert.strictEqual(patch.length, 1);
  assert.strictEqual(patch[0].path, LISTS_PATH + "/" + PHASE_LIST_ID + "/items/" + wasId + "/fields");
  assert.deepStrictEqual(patch[0].body, { Title: "R0001", Phase: 5, PhaseName: "Quality check",
                                          SetBy: WHO, SetAt: "2026-09-07T10:00:00.000Z" });
  assert.strictEqual(ITEMS.length, 1, "one item per job, still");
  assert.strictEqual(ITEMS[0].fields.Phase, 5);
  reset();
  r = await CW.listUpsert("Dashboard phases", "r0002", { Phase: 2, PhaseName: "Cutting", SetBy: WHO, SetAt: "x" });
  assert.strictEqual(r.created, true);
  const post = REQ.filter(x => x.method === "POST");
  assert.strictEqual(post.length, 1);
  assert.strictEqual(post[0].path, LISTS_PATH + "/" + PHASE_LIST_ID + "/items");
  assert.deepStrictEqual(post[0].body, { fields: { Title: "R0002", Phase: 2, PhaseName: "Cutting", SetBy: WHO, SetAt: "x" } },
    "the job number goes in upper-cased, whatever case it arrived in");
  assert.strictEqual(ITEMS.length, 2);
  pass("upsert PATCHes the item that is there and POSTs one when it is not");

  /* two clicks at once still leave one item: the writes are serialised per list */
  ITEMS = [];
  reset();
  await Promise.all([CW.listUpsert("Dashboard phases", "R0009", { Phase: 1 }),
                     CW.listUpsert("Dashboard phases", "R0009", { Phase: 4 })]);
  assert.strictEqual(ITEMS.length, 1, "not two rows for the same job");
  assert.strictEqual(ITEMS[0].fields.Phase, 4, "and the second click is the one that stands");
  assert.strictEqual(REQ.filter(x => x.method === "POST").length, 1);
  assert.strictEqual(REQ.filter(x => x.method === "PATCH").length, 1);
  pass("two writes to the same list at once cannot make a duplicate item");

  /* ---- 7a. two browsers, one job: the list keeps exactly one item ---- */
  /* the unique-values rule on Title refuses the second POST. That is not a
     failure - it means the item is there now, so read it back and update it. */
  ITEMS = []; UNIQUE_TITLE = true;
  INJECT_ON_POST = () => ITEMS.push(item("R0011", 2, "other@example.test", "2026-09-07T11:00:00.000Z"));
  reset();
  r = await CW.listUpsert("Dashboard phases", "R0011",
    { Phase: 5, PhaseName: "Quality check", SetBy: WHO, SetAt: "2026-09-07T11:05:00.000Z" });
  assert.strictEqual(r.created, false, "no new item: the one already there was updated");
  assert.strictEqual(ITEMS.length, 1, "one item for the job, not two");
  assert.strictEqual(ITEMS[0].fields.Phase, 5, "and it carries what this browser meant to write");
  assert.strictEqual(ITEMS[0].fields.SetBy, WHO);
  assert.deepStrictEqual(REQ.map(x => x.method), ["GET", "POST", "GET", "PATCH"],
    "read, POST refused, read back, PATCH the one that is there");
  pass("a POST refused by the unique-Title rule becomes a PATCH of the item that won the race");

  /* a POST refused for a real reason, with nothing left behind, still fails */
  ITEMS = []; UNIQUE_TITLE = false; FAIL_WRITE = 1;
  reset();
  await assert.rejects(() => CW.listUpsert("Dashboard phases", "R0012", { Phase: 3 }), /403/,
    "a refusal that leaves no item is a real failure and is thrown");
  assert.strictEqual(ITEMS.length, 0);
  pass("a write that fails for any other reason is still a failure, not a silent no-op");

  /* a list WITHOUT the unique rule lets both through: the older item wins and
     the newer one is deleted, so the next read cannot see two */
  ITEMS = [];
  INJECT_ON_POST = () => ITEMS.push(item("R0013", 1, "other@example.test", "2026-09-07T11:00:00.000Z"));
  reset();
  const older = NEXTID;                                  // the injected item takes this id
  r = await CW.listUpsert("Dashboard phases", "R0013",
    { Phase: 4, PhaseName: "In glazing", SetBy: WHO, SetAt: "2026-09-07T11:05:00.000Z" });
  assert.strictEqual(ITEMS.length, 1, "the duplicate was cleared away");
  assert.strictEqual(ITEMS[0].id, String(older), "the oldest id is the one that survives");
  assert.strictEqual(ITEMS[0].fields.Phase, 4, "with the fields this browser meant to write");
  assert.strictEqual(r.deduped, 1);
  assert.strictEqual(REQ.filter(x => x.method === "DELETE").length, 1);
  pass("without the unique rule both POSTs land, and the read-back keeps the oldest and deletes the rest");

  /* ---- 7b. a nextLink that does not start with the Graph base ---- */
  ITEMS = [item("R0001", 1), item("R0002", 2), item("R0003", 3), item("R0004", 4), item("R0005", 5)];
  PAGE = 2;
  NEXTLINK_ORIGIN = "https://eur-01.graph.microsoft.com/v1.0";
  reset();
  got = await CW.listItems("Dashboard phases");
  assert.strictEqual(got.length, 5, "every page was still followed");
  assert.ok(REQ.every(x => x.path.indexOf("http") < 0),
    "the origin is cut off, not handed to call() whole: " + REQ.map(x => x.path).join(" | "));
  NEXTLINK_ORIGIN = null; PAGE = 999;
  pass("an @odata.nextLink on another origin has it stripped before the next request");

  /* ---- 7. delete ---- */
  ITEMS = [item("R0001", 4), item("R0002", 2)];
  const goneId = ITEMS[0].id;
  reset();
  assert.strictEqual(await CW.listDelete("Dashboard phases", "R0001"), true);
  const del = REQ.filter(x => x.method === "DELETE");
  assert.strictEqual(del.length, 1);
  assert.strictEqual(del[0].path, LISTS_PATH + "/" + PHASE_LIST_ID + "/items/" + goneId);
  assert.deepStrictEqual(ITEMS.map(x => x.fields.Title), ["R0002"]);
  assert.strictEqual(await CW.listDelete("Dashboard phases", "R7777"), false, "nothing to remove is not an error");
  pass("clearing a phase DELETEs that job's item and leaves the rest alone");

  /* ---- 8. reading the list into PHASES_SET ---- */
  ITEMS = [item("R0001", 4, "ann@example.test", "2026-09-07T08:30:00.000Z"),
           item("r0002", 2, "bob@example.test", "2026-09-07T08:31:00.000Z"),
           { id: "900", fields: { Title: "", Phase: 3 } },                 // no job: skipped
           { id: "901", fields: { Title: "R0003", Phase: "" } },           // no phase: skipped
           { id: "902", fields: { Title: "R0004", Phase: 99 } }];          // out of range: clamped
  await readPhases();
  assert.deepStrictEqual(Object.keys(PHASES_SET).sort(), ["R0001", "R0002", "R0004"]);
  assert.strictEqual(PHASES_SET.R0002.phase, 2, "the job number is upper-cased on the way in");
  assert.strictEqual(PHASES_SET.R0001.who, "ann@example.test");
  assert.strictEqual(PHASES_SET.R0004.phase, 6, "a nonsense number is clamped to a real phase");
  assert.strictEqual(PHASE_LIST_OK, true);
  pass("the list is read into PHASES_SET, half-filled rows skipped and job numbers normalised");

  /* two items for one job: the newest SetAt is the one anybody sees, whichever
     order they come back in, and the next write clears the older one away */
  const dupOld = { id: "500", fields: { Title: "R0001", Phase: 2, PhaseName: "Cutting",
                                        SetBy: "old@example.test", SetAt: "2026-09-07T08:00:00.000Z" } };
  const dupNew = { id: "501", fields: { Title: "R0001", Phase: 5, PhaseName: "Quality check",
                                        SetBy: "new@example.test", SetAt: "2026-09-07T09:00:00.000Z" } };
  ITEMS = [dupOld, dupNew];
  await readPhases();
  assert.strictEqual(PHASES_SET.R0001.phase, 5, "the newest SetAt wins");
  assert.strictEqual(PHASES_SET.R0001.who, "new@example.test");
  ITEMS = [dupNew, dupOld];
  await readPhases();
  assert.strictEqual(PHASES_SET.R0001.phase, 5, "and it wins whichever order they arrive in");
  assert.strictEqual(Object.keys(PHASES_SET).length, 1, "one job, one phase, whatever the list holds");
  pass("two items for one job read as one: the newest SetAt is the one everybody sees");

  /* and the very next write leaves the list with one item again */
  clearHolds();
  useJobs([withWork({ f: "cut", s: "process" }, { id: "R0001", dates: { floor: "2026-09-01" } })]);
  reset();
  assert.strictEqual(await setPhaseByHand(byId("R0001"), 6), true);
  await settle();
  assert.strictEqual(ITEMS.length, 1, "the duplicate is gone");
  assert.strictEqual(ITEMS[0].id, "500", "the oldest id survives");
  assert.strictEqual(ITEMS[0].fields.Phase, 6, "carrying the phase that was just set");
  assert.strictEqual(REQ.filter(x => x.method === "DELETE").length, 1, "the extra item was deleted");
  assert.deepStrictEqual(bookReqs().filter(x => x.path.indexOf("Dashboard Log") < 0 &&
                                                x.path !== "/x/workbook/worksheets"), [],
    "and tidying the list still went nowhere near the workbook");
  clearHolds(); setPhases({}); CHANGES.length = 0; ITEMS = [];
  pass("the next write to that job collapses the duplicates, oldest id kept");

  /* a read that fails keeps the last map and says so once */
  const savedItems = ITEMS;
  LISTS = [{ id: "list-site-assets", displayName: "Site Assets" }];
  reset(); forget();
  await readPhases();
  assert.strictEqual(PHASE_LIST_OK, false, "the list is gone");
  LISTS = KEEP; forget();
  ITEMS = savedItems;
  await readPhases();
  assert.strictEqual(PHASE_LIST_OK, true);
  const before = JSON.parse(JSON.stringify(PHASES_SET));
  ITEMS = null;                       // the route will throw: a read that goes wrong
  reset();
  await readPhases();
  await readPhases();
  assert.deepStrictEqual(PHASES_SET, before, "the last known phases stay on screen");
  assert.strictEqual(TOASTS.filter(t => t.err).length, 1, "one toast for the page, not one per refresh");
  ITEMS = savedItems;
  vm.runInThisContext("phaseWarned = false;");
  pass("a read that fails keeps the last known phases and complains once");

  /* ---- 9. effectivePhase: the sheet is always the floor ---- */
  clearHolds();
  const sheet3 = withWork({ f: "cut", s: "process" }, { id: "R0001", dates: { floor: "2026-09-01" } });
  assert.strictEqual(jobPhase(sheet3), 3, "the sheet's own reading of this job");
  setPhases({});
  assert.strictEqual(effectivePhase(sheet3), 3);
  assert.strictEqual(phaseName(sheet3), "In fabrication");
  setPhases({ R0001: { phase: 5, name: "Quality check", who: "ann@example.test", at: "2026-09-07T08:30:00.000Z" } });
  assert.strictEqual(effectivePhase(sheet3), 5, "a hand-set phase ahead of the sheet is used");
  assert.strictEqual(phaseName(sheet3), "Quality check", "and phaseName follows it");
  setPhases({ R0001: { phase: 1, name: "Sent to floor", who: "ann@example.test", at: "" } });
  assert.strictEqual(effectivePhase(sheet3), 3, "a hand-set phase behind the sheet is ignored: the sheet has evidence");
  assert.strictEqual(phaseName(sheet3), "In fabrication");
  setPhases({});
  assert.strictEqual(effectivePhase(sheet3), 3, "clearing it returns the job to the sheet");
  const gold = mkJob({ id: "R0009", done: 1 });
  setPhases({ R0009: { phase: 0, name: "In office", who: "x", at: "" } });
  assert.strictEqual(effectivePhase(gold), 6, "nobody can drag a finished job backwards by hand");
  setPhases({});
  pass("effectivePhase is the higher of the sheet's own reading and the hand-set one");

  /* the row badge follows it too */
  const onFloor = mkJob({ id: "R0010", stage: "floor", cat: "active", dates: { floor: "2026-09-01" } });
  assert.strictEqual(statusWord(onFloor), "Sent to floor");
  setPhases({ R0010: { phase: 4, name: "In glazing", who: "x", at: "" } });
  assert.strictEqual(statusWord(onFloor), "In glazing", "the word in the list row is the effective phase");
  setPhases({});
  pass("the job row's badge shows the hand-set phase like any other");

  /* ---- 10. the hold: survives a stale read, released when the list agrees ---- */
  clearHolds(); setPhases({});
  useJobs([sheet3]);
  ITEMS = [item("R0001", 3)];
  await readPhases();
  assert.strictEqual(PHASES_SET.R0001.phase, 3);
  pend("R0001", { phase: 5 });
  vm.runInThisContext("ALL = applyPending(ALL, true);");
  assert.strictEqual(effectivePhase(byId("R0001")), 5, "our own click stands while the list is behind");
  assert.ok(PENDING.R0001 && PENDING.R0001.phase === 5, "and the hold is still there");
  ITEMS = [item("R0001", 5)];
  await readPhases();
  vm.runInThisContext("ALL = applyPending(ALL, true);");
  assert.strictEqual(PENDING.R0001, undefined, "the list agrees: the hold is let go");
  assert.strictEqual(effectivePhase(byId("R0001")), 5, "and what everyone can see says the same thing");
  pass("a phase hold survives a stale read and is released the moment the list agrees");

  /* a hold on a clear is let go when the item really is gone */
  pend("R0001", { phase: -1 });
  vm.runInThisContext("ALL = applyPending(ALL, true);");
  assert.strictEqual(effectivePhase(byId("R0001")), 3, "cleared here: straight back to the sheet");
  assert.ok(PENDING.R0001, "still held, because the list still shows the old value");
  ITEMS = [];
  await readPhases();
  vm.runInThisContext("ALL = applyPending(ALL, true);");
  assert.strictEqual(PENDING.R0001, undefined, "the item is gone: the hold goes with it");
  assert.strictEqual(effectivePhase(byId("R0001")), 3);
  pass("a hold on a cleared phase is released once the item really is gone");

  /* and either way it expires on its own after 180 s */
  pend("R0001", { phase: 6 });
  vm.runInThisContext("PENDING.R0001.t.phase = Date.now() - 200000; PENDING.R0001.at = Date.now() - 200000;");
  vm.runInThisContext("ALL = applyPending(ALL, true);");
  assert.strictEqual(PENDING.R0001, undefined, "180 s and it is gone, like every other hold");
  assert.strictEqual(effectivePhase(byId("R0001")), 3);
  pass("a phase hold expires after 180 s like every other pending change");

  /* ---- 11. setting a phase end to end ---- */
  clearHolds(); setPhases({}); ITEMS = []; LOGSHEET.length = 1;
  useJobs([sheet3]);
  await readPhases();
  reset();
  assert.strictEqual(await setPhaseByHand(byId("R0001"), 5), true);
  await settle();
  const setPaths = paths();
  trace("setting R0001 to phase 5", setPaths);
  assert.deepStrictEqual(listReqs().map(x => x.method), ["GET", "POST", "GET"],
    "the id is cached, so: read the list, add the item, read it back to be sure it is the only one");
  const made = REQ.find(x => x.method === "POST");
  assert.strictEqual(made.path, LISTS_PATH + "/" + PHASE_LIST_ID + "/items");
  assert.strictEqual(made.body.fields.Title, "R0001");
  assert.strictEqual(made.body.fields.Phase, 5);
  assert.strictEqual(made.body.fields.PhaseName, "Quality check");
  assert.strictEqual(made.body.fields.SetBy, WHO);
  assert.ok(/^\d{4}-\d{2}-\d{2}T/.test(made.body.fields.SetAt), "SetAt is an ISO stamp");
  assert.strictEqual(effectivePhase(byId("R0001")), 5);
  assert.ok(PENDING.R0001 && PENDING.R0001.phase === 5, "held for 180 s while SharePoint catches up");
  pass("clicking a step writes one list item and holds the new phase on screen at once");

  /* the log line, and where it went */
  assert.strictEqual(CHANGES.length, 1);
  assert.deepStrictEqual({ job: CHANGES[0].job, what: CHANGES[0].what, from: CHANGES[0].from, to: CHANGES[0].to },
    { job: "R0001", what: "Phase", from: "In fabrication", to: "Quality check" });
  assert.deepStrictEqual(LOGSHEET.slice(1).map(x => [x[2], x[3], x[4], x[5]]),
    [["R0001", "Phase", "In fabrication", "Quality check"]]);
  pass("one log line per change, named from and to, in the shared Dashboard Log");

  /* ---- 12. what the workbook was asked for ---- */
  const stray = bookReqs().filter(x => x.path.indexOf("Dashboard Log") < 0 && x.path !== "/x/workbook/worksheets");
  assert.deepStrictEqual(stray, [], "the workbook was touched for the audit log and nothing else:\n" + setPaths.join("\n"));
  assert.deepStrictEqual(REQ.filter(x => /Production|\/content|\/versions|createSession|\/drive\//.test(x.path)), [],
    "nothing downloaded the file, opened a session, or named the Production sheet");
  pass("setting a phase makes list requests and one Dashboard Log line - no other workbook request at all");

  /* ---- 13. clearing: click the step it is already on ---- */
  reset();
  clearHolds();
  await readPhases();
  assert.strictEqual(PHASES_SET.R0001.phase, 5, "everyone can see it now");
  const liveId = ITEMS[0].id;
  assert.strictEqual(await setPhaseByHand(byId("R0001"), 5), true);
  await settle();
  assert.deepStrictEqual(REQ.filter(x => x.method === "DELETE").map(x => x.path),
    [LISTS_PATH + "/" + PHASE_LIST_ID + "/items/" + liveId],
    "the item is deleted by its own id, not blanked");
  assert.strictEqual(REQ.filter(x => x.method === "PATCH" && x.path.indexOf("/lists/") > 0).length, 0,
    "and no field was written on the way out");
  trace("clearing R0001 (clicking the step it is on)", paths());
  assert.strictEqual(ITEMS.length, 0, "the list is empty again");
  assert.strictEqual(effectivePhase(byId("R0001")), 3, "and the job is back on the sheet's own reading");
  assert.deepStrictEqual({ what: CHANGES[0].what, from: CHANGES[0].from, to: CHANGES[0].to },
    { what: "Phase", from: "Quality check", to: "sheet" });
  assert.deepStrictEqual(bookReqs().filter(x => x.path.indexOf("Dashboard Log") < 0 && x.path !== "/x/workbook/worksheets"), [],
    "clearing touches no more of the workbook than setting did");
  pass("clicking the hand-set step again deletes the item and logs “Phase: X → sheet”");

  /* ---- 14. a write that is refused ---- */
  clearHolds(); ITEMS = []; setPhases({});
  useJobs([sheet3]);
  reset();
  FAIL_WRITE = 1;
  assert.strictEqual(await setPhaseByHand(byId("R0001"), 4), false);
  await settle();
  assert.strictEqual(ITEMS.length, 0, "nothing was written");
  assert.strictEqual(PENDING.R0001, undefined, "and the hold was dropped, so the screen tells the truth");
  assert.strictEqual(effectivePhase(byId("R0001")), 3);
  assert.strictEqual(TOASTS.filter(t => t.err).length, 1);
  assert.ok(/edit rights/i.test(TOASTS.filter(t => t.err)[0].m), "friendly(e), not a Graph error string");
  assert.strictEqual(CHANGES.length, 0, "and nothing was logged that did not happen");
  pass("a refused write says so in plain words, drops the hold and logs nothing");

  /* ---- 15. the sheet's own phase is the floor, in the code as well as the UI ---- */
  clearHolds(); setPhases({}); ITEMS = [];
  reset();
  assert.strictEqual(await setPhaseByHand(byId("R0001"), 1), false, "below the sheet's own phase 3");
  await settle();
  assert.strictEqual(REQ.length, 0, "not one request went out");
  assert.ok(/past this step/.test(TOASTS[TOASTS.length - 1].m));
  pass("a step the sheet has already gone past cannot be set, even calling the function directly");

  /* ---- 16. what the pipeline actually renders ---- */
  setPhases({});
  let html = phasePipeHtml(sheet3, false);
  assert.strictEqual((html.match(/<button/g) || []).length, 0, "read-only outside Edit mode, as before");
  assert.ok(html.indexOf("Now: <strong>In fabrication</strong>") > 0);

  html = phasePipeHtml(sheet3, true);
  const buttons = html.match(/<button[^>]*>/g) || [];
  assert.strictEqual(buttons.length, 7, "seven steps, seven buttons");
  for (let i = 0; i < 3; i++) {
    const b = buttons[i];
    assert.ok(b.indexOf(' data-ph="' + i + '"') > 0);
    assert.ok(b.indexOf(" disabled") > 0, "step " + i + " is below the sheet's own phase, so it is disabled");
    assert.ok(b.indexOf("the sheet already shows this job past this step") > 0, "and it says why");
  }
  for (let i = 3; i < 7; i++) {
    assert.ok(buttons[i].indexOf(" disabled") < 0, "step " + i + " can be chosen");
    assert.ok(buttons[i].indexOf("set this job to ") > 0);
  }
  pass("in Edit mode every step is a button; the ones the sheet has passed are disabled and say why");

  /* the tap target is big enough to hit on a phone */
  const css = fs.readFileSync(__dirname + "/index.html", "utf8");
  const rule = /button\.phstep\s*{[^}]*}/.exec(css);
  assert.ok(rule, "there is a rule for the step buttons");
  const mh = /min-height:\s*(\d+)px/.exec(rule[0]);
  assert.ok(mh && Number(mh[1]) >= 40, "a step is at least 40 px tall, not a 10 px sliver");
  pass("a step button is at least 40 px tall - these are tapped, not clicked");

  /* the note: who set it, and what the sheet says if the two differ */
  setPhases({ R0001: { phase: 5, name: "Quality check", who: "ann@example.test", at: "2026-09-07T08:30:00.000Z" } });
  html = phasePipeHtml(sheet3, true);
  assert.ok(html.indexOf("Set by <strong>ann@example.test</strong>") > 0, "who set it");
  assert.ok(html.indexOf("07/09") > 0, "and when");
  assert.ok(html.indexOf("sheet says: <strong>In fabrication</strong>") > 0, "and what the sheet itself still says");
  assert.ok(/data-ph="5"[^>]*title="click again to clear/.test(html), "the step it is on offers to clear it");
  setPhases({ R0001: { phase: 3, name: "In fabrication", who: "ann@example.test", at: "2026-09-07T08:30:00.000Z" } });
  html = phasePipeHtml(sheet3, true);
  assert.ok(html.indexOf("Set by <strong>ann@example.test</strong>") > 0);
  assert.strictEqual(html.indexOf("sheet says:"), -1, "nothing to point out when the two agree");
  setPhases({});
  pass("a hand-set phase says who set it and when, and what the sheet says when they differ");

  /* ---- 16b. before the first read has answered ---- */
  setPhases({});
  vm.runInThisContext("PHASE_LIST_OK = null;");
  html = phasePipeHtml(sheet3, true);
  const waiting = html.match(/<button[^>]*>/g) || [];
  assert.strictEqual(waiting.length, 7);
  assert.strictEqual(waiting.filter(b => b.indexOf(" disabled") > 0).length, 7,
    "nothing can be chosen before we know there is somewhere to put the answer");
  assert.ok(waiting.every(b => b.indexOf('title="checking') > 0), "and every one says why");
  assert.ok(html.indexOf("Checking the <strong>Dashboard phases</strong> list") > 0);
  /* clicking one anyway reads the list first rather than guessing */
  ITEMS = []; clearHolds();
  useJobs([sheet3]);
  reset();
  assert.strictEqual(await setPhaseByHand(byId("R0001"), 5), true);
  await settle();
  assert.strictEqual(REQ[0].method, "GET", "the first thing it does is read the list");
  assert.strictEqual(PHASE_LIST_OK, true);
  assert.strictEqual(ITEMS.length, 1);
  ITEMS = []; clearHolds(); setPhases({}); CHANGES.length = 0;
  pass("until the first read answers, every step is disabled and says “checking…”");

  /* ---- 17. no list, no controls, no writes ---- */
  LISTS = [{ id: "list-site-assets", displayName: "Site Assets" }];
  forget(); clearHolds(); setPhases({});
  await readPhases();
  assert.strictEqual(PHASE_LIST_OK, false);
  html = phasePipeHtml(sheet3, true);
  assert.strictEqual((html.match(/<button/g) || []).length, 0, "no clickable steps without somewhere to put the answer");
  assert.ok(html.indexOf("is not in SharePoint yet") > 0, "a plain message instead");
  assert.ok(html.indexOf("nothing in the Excel file is involved") > 0, "and it says the workbook has nothing to do with it");
  reset();
  assert.strictEqual(await setPhaseByHand(sheet3, 5), false);
  await settle();
  assert.strictEqual(REQ.length, 0, "and calling it anyway does nothing at all");
  assert.ok(/not in SharePoint yet/.test(TOASTS[TOASTS.length - 1].m));
  LISTS = KEEP; forget();
  await readPhases();
  pass("a missing list shows a plain message and writes nothing anywhere");

  /* ---- 18. nothing new in this browser ---- */
  const keys = Object.keys(mem).sort();
  /* cw_theme and cw_changes are the page's own, from long before this feature */
  const OLD_KEYS = ["cw_pending", "cw_changes", "cw_theme"];
  assert.deepStrictEqual(keys.filter(k => OLD_KEYS.indexOf(k) < 0), ["cw_listids"],
    "the 180 s hold and cw_listids: nothing else was added. Keys: " + keys.join(", "));
  assert.strictEqual(String(mem.cw_listids), JSON.stringify({ "Dashboard phases": PHASE_LIST_ID }),
    "cw_listids holds ids, never a phase or a job");
  pass("nothing is kept in this browser beyond the 180 s hold and cw_listids");

  /* ---- 19. the whole run, end to end ---- */
  assert.deepStrictEqual(ALLREQ.filter(x => /Production|\/content|\/versions|createSession|\/drive\/|\/range\(/.test(x.path) &&
                                            x.path.indexOf("Dashboard Log") < 0).map(x => x.method + " " + x.path), [],
    "over every request of the whole run, not one touched the workbook outside the Dashboard Log");
  assert.ok(ALLREQ.length > 60, "and that is over a real number of requests, not an empty log");
  pass("across every test above, the workbook was never read, written or asked about");

  console.log("\n" + n + " checks passed");
  process.exit(0);                 // the 45 s reconcile timer would hold the process open
})().catch(e => { console.error("FAIL", e); process.exit(1); });
