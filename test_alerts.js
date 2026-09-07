/* Offline test of the job-alerts feature: reading the Dashboard Config and
   Dashboard Alerts sheets, who counts as the administrator, address
   validation, the upsert (no duplicate rows), removal by blanking the row,
   the 180 s hold and how it is released, the log lines, what is allowed to
   reach localStorage, the finished-job rule, and the fact that nothing outside
   Dashboard Alerts / Dashboard Log is ever written.
   Graph is a fake fetch() over an in-memory workbook - nothing leaves the box.
   Every address here is deliberately fake (*@example.test).
   Run: node test_alerts.js                                                  */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, addEventListener() {} };
global.performance = { now: () => Date.now() };
global.ExcelJS = ExcelJS;
/* The elements remember what was written into them - the HTML of a window, the
   children appended to the chip bar - so a test can look at what the code
   actually put on screen, not merely that it did not throw. Setting innerHTML
   forgets the children again, the way a browser does. */
function stubEl() {
  let html = "";
  const e = {
    style: {}, dataset: {}, textContent: "", value: "", disabled: false, hidden: false, className: "", kids: [],
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); }, remove() {}, addEventListener() {}, removeEventListener() {},
    focus() {}, blur() {}, setSelectionRange() {},
    getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => stubEl(), querySelectorAll: () => []
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; } });
  return e;
}
/* one element per selector, kept between calls, so #ahost and #chips can be
   read back after a render */
const EL = {};
const el = sel => EL[sel] || (EL[sel] = stubEl());
global.document = {
  documentElement: stubEl(), body: stubEl(), createElement: () => stubEl(), activeElement: null,
  querySelector: el, querySelectorAll: () => [], addEventListener() {}, removeEventListener() {}
};

/* ---------- a tiny Excel behind a fake Graph ---------- */
const BOOK = {};
const CALLS = [];
let FAIL_ALERT_WRITE = 0;          // 403: refused outright, so no retry sleeps in the test
const kk = (r, c) => r + "|" + c;
const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
function sh(name) { return BOOK[name] || (BOOK[name] = { v: {} }); }
function bounds(s) {
  let maxR = 0, maxC = 0;
  Object.keys(s.v).forEach(k => { const p = k.split("|"); if (s.v[k] !== "") { maxR = Math.max(maxR, +p[0]); maxC = Math.max(maxC, +p[1]); } });
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
function route(method, path, body) {
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
  if (method === "PATCH" && tail === "" && body && body.values) {
    if (FAIL_ALERT_WRITE && name === "Dashboard Alerts" && mr[1] !== "A1:D1") {
      FAIL_ALERT_WRITE--; return { status: 403, body: { error: { code: "AccessDenied" } } };
    }
    body.values.forEach((line, ri) => line.forEach((x, ci) => { s.v[kk(A.r1 + ri, A.c1 + ci)] = x; }));
    CALLS.push({ method, sheet: name, kind: "values", addr: mr[1], values: body.values });
    return ok({});
  }
  if (method === "PATCH") { CALLS.push({ method, sheet: name, kind: tail ? "format" : "numberFormat", addr: mr[1] }); return ok({}); }
  return { status: 404, body: { error: "no route " + method + " " + path } };
}
global.fetch = async (url, init) => {
  const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  const res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status, text: async () => JSON.stringify(res.body), arrayBuffer: async () => new ArrayBuffer(0) };
};

/* A hold left behind by a previous visit, already older than the 180 s window.
   It is in localStorage before app.js is ever parsed, exactly as it would be
   when the tab is re-opened the next morning. */
mem.cw_penda = JSON.stringify({
  "R9999|stale@example.test": { act: "add", who: "someone@example.test", when: "2026-09-01 08:00", at: Date.now() - 200000 }
});

/* ---------- load the app the way the page does ---------- */
const run = f => vm.runInThisContext(fs.readFileSync(__dirname + "/" + f, "utf8"), { filename: f });
run("parser.js");
run("graph.js");
global.CW = window.CW;
CW._setToken(() => "t");
CW._setFile({ base: "/x/workbook", content: "/x/content", meta: "/x" });
run("checkpoints.js");
global.CP = window.CP;
run("app.js");

/* who is signed in - the real one asks MSAL, which is not here */
let WHO = "admin@example.test";
global.whoAmI = () => WHO;
const TOASTS = [];
global.toast = (m, err) => TOASTS.push({ m: String(m), err: !!err });

const ADMIN = "admin@example.test", COLLEAGUE = "colleague@example.test", OTHER = "other@example.test";
const OUTSIDE = "someone@elsewhere.test";
const settle = ms => new Promise(r => setTimeout(r, ms == null ? 40 : ms));
const reset = () => { CALLS.length = 0; TOASTS.length = 0; };
const alertWrites = () => CALLS.filter(c => c.kind === "values" && c.sheet === "Dashboard Alerts" && c.addr !== "A1:D1");
const logWrites = () => CALLS.filter(c => c.kind === "values" && c.sheet === "Dashboard Log" && c.addr !== "A1:F1");
/* every sheet this feature wrote to, apart from the two it is allowed to touch */
const strayWrites = () => CALLS.filter(c => c.sheet && c.sheet !== "Dashboard Alerts" && c.sheet !== "Dashboard Log");
/** the (Job, Email) lines that are actually on the fake Alerts sheet */
function sheetRows() {
  const s = BOOK["Dashboard Alerts"]; if (!s) return [];
  const out = [];
  for (let r = 2; r <= 200; r++) {
    const j = s.v[kk(r, 1)], e = s.v[kk(r, 2)];
    if (j || e) out.push({ row: r, job: j, email: e, who: s.v[kk(r, 3)], when: s.v[kk(r, 4)] });
  }
  return out;
}
const held = () => JSON.parse(localStorage.getItem("cw_penda") || "{}");
/* the holds left over from an earlier step are not what the next one is about */
const clearHolds = () => { vm.runInThisContext("PENDA = {}"); savePendA(); };
/** every localStorage value that mentions an address, by key */
const leaks = addr => Object.keys(mem).filter(k => String(mem[k]).toLowerCase().indexOf(addr) >= 0);

/* ---------- a workbook to read, and jobs to look at ---------- */
function mkWorkbook(cfg, alerts) {
  const wb = new ExcelJS.Workbook();
  if (cfg) {
    const ws = wb.addWorksheet("Dashboard Config");
    ws.getRow(1).getCell(1).value = "Key"; ws.getRow(1).getCell(2).value = "Value";
    cfg.forEach((p, i) => { ws.getRow(i + 2).getCell(1).value = p[0]; ws.getRow(i + 2).getCell(2).value = p[1]; });
  }
  if (alerts) {
    const ws = wb.addWorksheet("Dashboard Alerts");
    ["Job", "Email", "Added by", "When"].forEach((h, c) => { ws.getRow(1).getCell(c + 1).value = h; });
    alerts.forEach((p, i) => p.forEach((v, c) => { ws.getRow(i + 2).getCell(c + 1).value = v; }));
  }
  return wb;
}
const mkJob = extra => Object.assign({
  id: "R0001", cust: "Ann", area: "Cork", eir: "", off: "", colour: "", ph3: "", wnd: 10, drs: 2,
  glass: {}, prods: [], notes: [], sheets: ["Production"], src: {},
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cp: { win: "", drs: "", glass: {}, prod: {} }, cat: "active", blk: 4, seq: 1,
  stage: "floor", done: 0, urg: 0
}, extra || {});
function useJobs(list, names) {
  global.__jobs = list; global.__names = names || ["Can sell as second hand", "Ready, customer won't take",
    "Collect & supply only", "Ready to fit", "In production"];
  vm.runInThisContext("ALL = __jobs; BLOCKNAMES = __names; ALL.blockNames = __names; CHANGES = []; state.sel = null; state.picked = {};");
}

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ---- 0. a hold left over from a previous visit ---- */
  assert.deepStrictEqual(PENDA, {}, "the stale hold did not survive the parse - no load() needed");
  assert.strictEqual(String(mem.cw_penda || "").indexOf("stale@example.test"), -1,
    "and it was written out of localStorage there and then");
  vm.runInThisContext("PENDA['R9998|stale@example.test'] = " +
    "{ act:'add', who:'x', when:'', at: Date.now() - 200000 }");
  savePendA();
  assert.strictEqual(String(mem.cw_penda).indexOf("stale@example.test"), -1, "and no write ever persists a stale one");
  assert.deepStrictEqual(PENDA, {});
  pass("a hold older than 180 s is dropped on parse and never written back out");

  /* ---- 1. reading the two sheets ---- */
  readAlertSheets(mkWorkbook([["admin", ADMIN], ["dashboardUrl", "https://example.test/dash"]],
    [["R0001", COLLEAGUE, "Someone", "2026-09-01 09:00"], ["R0001", COLLEAGUE, "Someone", "2026-09-01 09:01"],
     ["r0002", "COLLEAGUE@EXAMPLE.TEST", "Someone", "2026-09-02 09:00"], ["", "", "", ""]]));
  assert.strictEqual(CONFIG.admin, ADMIN);
  assert.strictEqual(CONFIG.dashboardurl, "https://example.test/dash", "keys are matched case-insensitively");
  pass("Dashboard Config is read into CONFIG, key by key");
  assert.deepStrictEqual(Object.keys(ALERTS).sort(), ["R0001", "R0002"]);
  assert.deepStrictEqual(ALERTS.R0001.map(x => x.email), [COLLEAGUE], "a line typed twice by hand is one subscription");
  assert.deepStrictEqual(ALERTS.R0002.map(x => x.email), [COLLEAGUE], "job and address are normalised on the way in");
  assert.strictEqual(ALERTS.R0001[0].who, "Someone");
  assert.strictEqual(ALERTS.R0001[0].when, "2026-09-01 09:00");
  pass("Dashboard Alerts is read into ALERTS, blank lines skipped, addresses lower-cased");

  WHO = ADMIN;
  assert.strictEqual(isAdmin(), true, "the configured address is the administrator");
  WHO = COLLEAGUE;
  assert.strictEqual(isAdmin(), false);
  WHO = ADMIN.toUpperCase();
  assert.strictEqual(isAdmin(), true, "the comparison is case-insensitive both ways");
  WHO = ADMIN;
  pass("isAdmin() compares the signed-in address with the one in the sheet");

  readAlertSheets(mkWorkbook(null, null));
  assert.deepStrictEqual(CONFIG, {});
  assert.deepStrictEqual(ALERTS, {});
  assert.strictEqual(isAdmin(), false, "no Config sheet: nobody is the administrator");
  assert.strictEqual(alertEmailCheck(COLLEAGUE).ok, false, "and nothing can be added");
  readAlertSheets(mkWorkbook([["something else", "x"]], []));
  assert.strictEqual(isAdmin(), false, "a Config sheet with no admin key: still nobody");
  pass("a missing sheet or a missing key means nobody is the administrator");

  /* ---- 2. validation ---- */
  readAlertSheets(mkWorkbook([["admin", ADMIN]], []));
  assert.deepStrictEqual(alertEmailCheck("  COLLEAGUE@Example.TEST  "), { ok: true, email: COLLEAGUE });
  pass("an address is trimmed and lower-cased");
  ["", "   ", "abin", "abin@", "@example.test", "abin example.test", "abin@example",
   "a b@example.test", "colleague@example.test, other@example.test"].forEach(bad =>
    assert.strictEqual(alertEmailCheck(bad).ok, false, JSON.stringify(bad) + " is not an address"));
  pass("anything that is not an email address is refused");
  const wrong = alertEmailCheck(OUTSIDE);
  assert.strictEqual(wrong.ok, false);
  assert.ok(/organisation/i.test(wrong.msg), "the message says why");
  assert.strictEqual(alertEmailCheck("COLLEAGUE@EXAMPLE.TEST").email.split("@")[1], adminDomain(),
    "the only domain allowed is the administrator's own, taken from the sheet");
  pass("an address outside the administrator's domain is refused");

  /* ---- 3. admin gating: a non-admin never writes, even called directly ---- */
  BOOK["Dashboard Alerts"] = { v: {} };
  readAlertSheets(mkWorkbook([["admin", ADMIN]], []));
  WHO = COLLEAGUE;
  reset();
  let r = await addJobAlert("R0001", COLLEAGUE);
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.msg, ALERT_ADMIN_ONLY);
  r = await removeJobAlert("R0001", COLLEAGUE);
  assert.strictEqual(r.ok, false);
  const many = await alertMany(["R0001", "R0002"], COLLEAGUE);
  assert.strictEqual(many.ok, false);
  await settle();
  assert.strictEqual(CALLS.length, 0, "not one request went out");
  assert.deepStrictEqual(held(), {}, "and nothing was held either");
  assert.deepStrictEqual(ALERTS, {}, "and the list is untouched");
  pass("a non-admin cannot add, remove or bulk-add, even calling the functions directly");

  /* ---- 4. adding: one row, one log line, nothing else touched ---- */
  WHO = ADMIN;
  useJobs([mkJob(), mkJob({ id: "R0002", cust: "Bob", blk: 4 })]);
  reset();
  r = await addJobAlert("r0001", "  COLLEAGUE@EXAMPLE.TEST ");
  await settle();
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.email, COLLEAGUE);
  assert.deepStrictEqual(sheetRows().map(x => [x.job, x.email, x.who]), [["R0001", COLLEAGUE, ADMIN]]);
  assert.strictEqual(alertWrites().length, 1, "exactly one line written to the Alerts sheet");
  assert.strictEqual(alertWrites()[0].addr, "A2:D2", "the first subscription goes on the first free row");
  assert.deepStrictEqual(logWrites().map(w => w.values[0].slice(2, 6)),
    [["R0001", "Alert", "", COLLEAGUE]], 'the log line is Alert, "" -> the address');
  assert.strictEqual(strayWrites().length, 0, "no other sheet was written to");
  assert.deepStrictEqual(ALERTS.R0001.map(x => x.email), [COLLEAGUE], "the chip is there straight away");
  pass("add: one upserted row, one log line, and nothing outside the two sheets");

  /* ---- 5. the upsert: no second row for the same job and address ---- */
  reset();
  r = await addJobAlert("R0001", COLLEAGUE.toUpperCase());
  await settle();
  assert.strictEqual(r.already, true, "the dashboard already knows about it");
  assert.strictEqual(CALLS.length, 0, "so nothing is written at all");
  await CW.addAlert("r0001", "COLLEAGUE@EXAMPLE.TEST", "Someone else");   // the sheet-level upsert
  await settle();
  assert.strictEqual(sheetRows().length, 1, "still one row");
  assert.strictEqual(alertWrites()[0].addr, "A2:D2", "the same row is written again, not a new one");
  assert.strictEqual(sheetRows()[0].who, "Someone else");
  assert.strictEqual(logWrites().length, 0, "a bare CW.addAlert does not log - the dashboard does that");
  pass("adding the same address twice updates one row instead of adding a second");

  reset();
  await addJobAlert("R0001", OTHER);
  await addJobAlert("R0002", COLLEAGUE);
  await settle();
  assert.deepStrictEqual(sheetRows().map(x => x.row + ":" + x.job + " " + x.email),
    ["2:R0001 " + COLLEAGUE, "3:R0001 " + OTHER, "4:R0002 " + COLLEAGUE]);
  assert.deepStrictEqual(alertAddresses(), [COLLEAGUE, OTHER], "the quick-add list is every address in use");
  pass("a second address on a job, and the same address on a second job, get their own rows");

  /* ---- 6. removing blanks the row ---- */
  reset();
  r = await removeJobAlert("R0001", "COLLEAGUE@EXAMPLE.TEST");
  await settle();
  assert.strictEqual(r.ok, true);
  assert.strictEqual(sheetRows().length, 2, "the line is gone from the sheet");
  const s = BOOK["Dashboard Alerts"];
  assert.deepStrictEqual([1, 2, 3, 4].map(c => s.v[kk(2, c)]), ["", "", "", ""], "row 2 is blank, not deleted");
  assert.deepStrictEqual(logWrites().map(w => w.values[0].slice(2, 6)), [["R0001", "Alert", COLLEAGUE, ""]],
    'the log line is Alert, the address -> ""');
  assert.deepStrictEqual((ALERTS.R0001 || []).map(x => x.email), [OTHER]);
  assert.strictEqual(strayWrites().length, 0);
  pass("remove: the row is blanked (A:D) and logged the other way round");

  reset();
  await addJobAlert("R0003", COLLEAGUE);
  await settle();
  assert.strictEqual(sheetRows().filter(x => x.job === "R0003")[0].row, 2, "the blanked row is used again");
  pass("a later add re-uses the row a removal blanked instead of growing the sheet");

  /* ---- 7. the hold ---- */
  reset(); clearHolds();
  await addJobAlert("R0009", COLLEAGUE);
  await settle();
  assert.ok(held()["R0009|" + COLLEAGUE], "the add is held while SharePoint catches up");
  /* the downloadable file is still ~36 s behind: it knows nothing about it */
  readAlertSheets(mkWorkbook([["admin", ADMIN]], [["R0001", OTHER, ADMIN, "2026-09-04 10:00"]]));
  assert.deepStrictEqual((ALERTS.R0009 || []).map(x => x.email), [COLLEAGUE], "a stale read does not undo the add");
  assert.ok(held()["R0009|" + COLLEAGUE], "so the hold stays");
  /* now the file agrees */
  readAlertSheets(mkWorkbook([["admin", ADMIN]],
    [["R0001", OTHER, ADMIN, "2026-09-04 10:00"], ["R0009", COLLEAGUE, ADMIN, "2026-09-04 10:01"]]));
  assert.deepStrictEqual(held(), {}, "the hold is let go the moment the file agrees");
  assert.deepStrictEqual(ALERTS.R0009.map(x => x.email), [COLLEAGUE], "and the sheet speaks for itself");
  pass("an add is held through a stale read and released when the file agrees");

  await removeJobAlert("R0001", OTHER);
  await settle();
  assert.ok(held()["R0001|" + OTHER], "a removal is held the same way");
  readAlertSheets(mkWorkbook([["admin", ADMIN]],
    [["R0001", OTHER, ADMIN, "2026-09-04 10:00"], ["R0009", COLLEAGUE, ADMIN, "2026-09-04 10:01"]]));
  assert.deepStrictEqual((ALERTS.R0001 || []).map(x => x.email), [], "the stale file still lists it: keep hiding it");
  readAlertSheets(mkWorkbook([["admin", ADMIN]], [["R0009", COLLEAGUE, ADMIN, "2026-09-04 10:01"]]));
  assert.deepStrictEqual(held(), {}, "gone from the file: the hold is released");
  pass("a removal is held through a stale read and released when the file agrees");

  await addJobAlert("R0010", COLLEAGUE);
  await settle();
  vm.runInThisContext("PENDA['R0010|" + COLLEAGUE + "'].at -= 200000");   // older than 180 s
  readAlertSheets(mkWorkbook([["admin", ADMIN]], [["R0009", COLLEAGUE, ADMIN, "2026-09-04 10:01"]]));
  assert.deepStrictEqual(held(), {}, "an expired hold is dropped");
  assert.strictEqual(ALERTS.R0010, undefined, "and what the file says takes over");
  pass("a hold that nothing confirms is let go after 180 s");

  /* ---- 8. a failed write puts the screen back ---- */
  reset();
  FAIL_ALERT_WRITE = 1;
  r = await addJobAlert("R0011", COLLEAGUE);
  await settle();
  assert.strictEqual(r.ok, false);
  assert.deepStrictEqual((ALERTS.R0011 || []).map(x => x.email), [], "the chip is taken away again");
  assert.deepStrictEqual(held(), {}, "and the hold with it");
  assert.strictEqual(TOASTS.filter(t => t.err).length, 1, "the person is told");
  assert.strictEqual(logWrites().length, 0, "a write that failed is not logged");
  assert.strictEqual(sheetRows().filter(x => x.job === "R0011").length, 0);
  pass("a refused write drops the hold, puts the list back and says so");

  /* ---- 9. what may be kept in this browser ---- */
  reset();
  vm.runInThisContext("CHANGES = []");
  delete mem.cw_changes; clearHolds();
  await addJobAlert("R0012", COLLEAGUE);
  await settle();
  assert.deepStrictEqual(leaks(COLLEAGUE), ["cw_penda"], "only the hold mentions the address");
  assert.strictEqual(mem.cw_alerts, undefined, "ALERTS itself is never stored");
  assert.ok(CHANGES.some(c => c.what === "Alert" && c.to === COLLEAGUE), "the Changes list still shows it in memory");
  readAlertSheets(mkWorkbook([["admin", ADMIN]], [["R0012", COLLEAGUE, ADMIN, "2026-09-04 10:02"]]));
  saveChanges();
  assert.deepStrictEqual(leaks(COLLEAGUE), [], "once the hold is released nothing here knows the address");
  pass("an address reaches localStorage only inside the 180 s hold, never anywhere else");

  /* ---- 10. several ticked jobs at once ---- */
  BOOK["Dashboard Alerts"] = { v: {} };
  readAlertSheets(mkWorkbook([["admin", ADMIN]], []));
  useJobs([mkJob(), mkJob({ id: "R0002", cust: "Bob" }), mkJob({ id: "R0003", cust: "Cal" })]);
  reset();
  const bulk = await alertMany(["R0001", "R0002", "R0003"], "COLLEAGUE@Example.test");
  await settle();
  assert.strictEqual(bulk.added, 3);
  assert.deepStrictEqual(sheetRows().map(x => x.job), ["R0001", "R0002", "R0003"]);
  assert.strictEqual(logWrites().length, 3, "one log line per job");
  assert.deepStrictEqual(logWrites().map(w => w.values[0][2]), ["R0001", "R0002", "R0003"]);
  assert.strictEqual(strayWrites().length, 0);
  pass("Alert to… applies one address to every ticked job, one log line each");

  /* ---- 11. finished jobs ---- */
  const inProd = mkJob({ id: "R0020", blk: 4 });
  assert.strictEqual(sectionIdx("In production"), 4);
  assert.strictEqual(alertFinished(inProd), false, "a job in In production is still emailed");
  assert.strictEqual(alertFinished(mkJob({ id: "R0021", blk: 3 })), true, "moved to Ready to fit: finished");
  assert.strictEqual(alertFinished(mkJob({ id: "R0022", blk: 2 })), true, "collect & supply: finished");
  assert.strictEqual(alertFinished(mkJob({ id: "R0023", blk: 4, done: 1 })), true, "a gold row is finished");
  assert.strictEqual(alertFinished(null), true, "no longer on the sheet at all: finished");
  vm.runInThisContext("BLOCKNAMES = []");
  assert.strictEqual(alertFinished(inProd), false, "with no sections read yet, nothing is called finished");
  useJobs([inProd]);
  pass("the finished rule: only a job still in In production, and not gold, keeps getting emails");

  /* ---- 12. what the drawer and the windows put on screen ---- */
  clearHolds();
  readAlertSheets(mkWorkbook([["admin", ADMIN]],
    [["R0020", COLLEAGUE, ADMIN, "2026-09-04 10:00"], ["R0021", OTHER, ADMIN, "2026-09-04 10:05"]]));
  WHO = ADMIN;
  let html = alertsSectionHtml(inProd);
  assert.ok(html.indexOf(COLLEAGUE) >= 0, "the subscribed address is a chip");
  assert.ok(html.indexOf('id="alnew"') >= 0 && html.indexOf('type="email"') >= 0, "the admin gets an input");
  assert.ok(html.indexOf('data-al-x="' + COLLEAGUE + '"') >= 0, "and a remove button on the chip");
  assert.ok(html.indexOf('data-al-quick="' + OTHER + '"') >= 0, "and quick-add chips for addresses used elsewhere");
  assert.ok(html.indexOf('data-al-quick="' + COLLEAGUE + '"') < 0, "but not for one already on this job");
  assert.ok(html.indexOf(ALERT_NOTE) >= 0, "the schedule is spelled out once, from one constant");
  assert.ok(html.indexOf("Finished") < 0, "this job is not finished");
  ALCONF[alKey("R0020", COLLEAGUE)] = 1;
  assert.ok(alertsSectionHtml(inProd).indexOf('data-al-yes="' + COLLEAGUE + '"') >= 0,
    "removing asks inline rather than with a browser dialog");
  delete ALCONF[alKey("R0020", COLLEAGUE)];
  WHO = COLLEAGUE;
  html = alertsSectionHtml(inProd);
  assert.ok(html.indexOf(COLLEAGUE) >= 0, "everyone sees who is subscribed");
  assert.ok(html.indexOf('id="alnew"') < 0 && html.indexOf("data-al-x") < 0, "but only the admin gets the controls");
  assert.ok(html.indexOf(ALERT_ADMIN_ONLY) >= 0, "and is told why");
  WHO = ADMIN;
  assert.ok(alertsSectionHtml(mkJob({ id: "R0021", blk: 3 })).indexOf("Finished — no more emails for this job.") >= 0);
  pass("the drawer section: chips for everyone, controls for the admin, a finished notice");

  const by = alertsByAddress();
  assert.deepStrictEqual(Object.keys(by).sort(), [COLLEAGUE, OTHER]);
  assert.deepStrictEqual(by[COLLEAGUE].map(x => x.job), ["R0020"]);
  renderAlertsWindow();                       // the DOM is a stub: this only has to not throw
  renderAlertMenu(stubEl());
  useJobs([inProd, mkJob({ id: "R0021", blk: 3 })]);
  state.sel = "R0020"; renderDrawer(); state.sel = null;
  pass("the Alerts window, the Alert to… menu and the drawer all render");

  /* ---- 13. the sheet is created the first time it is needed ---- */
  delete BOOK["Dashboard Alerts"];
  CW._resetSheetMemo();
  reset();
  await CW.addAlert("R0100", COLLEAGUE, ADMIN);
  await settle();
  const made = CALLS.filter(c => c.sheet === "Dashboard Alerts");
  assert.ok(made.some(c => c.kind === "addSheet"), "the sheet is added");
  assert.deepStrictEqual((made.find(c => c.kind === "values" && c.addr === "A1:D1") || {}).values,
    [["Job", "Email", "Added by", "When"]], "with its header row");
  assert.strictEqual(made.filter(c => c.kind === "format" && c.addr === "A1:D1").length, 2, "bold white on the dark fill");
  assert.ok(made.some(c => c.kind === "numberFormat" && c.addr === "A2:D2000"),
    'A2:D2000 is pre-formatted as text, so Excel never re-reads an address or a stamp');
  assert.deepStrictEqual(made.filter(c => c.kind === "format" && /^[A-D]:[A-D]$/.test(c.addr)).map(c => c.addr),
    ["A:A", "B:B", "C:C", "D:D"], "and every column gets its width");
  assert.deepStrictEqual(sheetRows().map(x => [x.row, x.job, x.email]), [[2, "R0100", COLLEAGUE]],
    "the first subscription lands under the header");
  pass("the Dashboard Alerts sheet is created on first use, formatted like the other dashboard sheets");

  /* ---- 14. two writes at once ---- */
  BOOK["Dashboard Alerts"] = { v: {} };
  reset();
  await Promise.all([CW.addAlert("R0201", COLLEAGUE, ADMIN), CW.addAlert("R0202", OTHER, ADMIN)]);
  assert.deepStrictEqual(sheetRows().map(x => x.row + " " + x.job + " " + x.email),
    ["2 R0201 " + COLLEAGUE, "3 R0202 " + OTHER], "two rows, not one on top of the other");
  await Promise.all([CW.addAlert("R0203", COLLEAGUE, ADMIN), CW.addAlert("R0203", "COLLEAGUE@EXAMPLE.TEST", ADMIN)]);
  assert.strictEqual(sheetRows().filter(x => x.job === "R0203").length, 1,
    "the same pair sent twice at once is still one row");
  pass("concurrent writes are queued per sheet: distinct pairs get distinct rows, the same pair one");

  /* ---- 15. a pair typed twice by hand in Excel ---- */
  const dup = BOOK["Dashboard Alerts"] = { v: {} };
  ["Job", "Email", "Added by", "When"].forEach((h, c) => { dup.v[kk(1, c + 1)] = h; });
  [[2, "R0300", COLLEAGUE], [3, "R0301", OTHER], [4, "R0300", COLLEAGUE]].forEach(p => {
    dup.v[kk(p[0], 1)] = p[1]; dup.v[kk(p[0], 2)] = p[2]; dup.v[kk(p[0], 3)] = ADMIN; dup.v[kk(p[0], 4)] = "2026-09-04 10:00";
  });
  await CW.removeAlert("R0300", COLLEAGUE);
  assert.deepStrictEqual(sheetRows().map(x => x.row + " " + x.job), ["3 R0301"], "both copies are gone");
  assert.deepStrictEqual([1, 2, 3, 4].map(c => dup.v[kk(4, c)]), ["", "", "", ""], "the second one is blanked too");
  pass("removing blanks every duplicate line, and leaves everyone else's alone");

  /* ---- 16. a non-admin is offered nothing ---- */
  BOOK["Dashboard Alerts"] = { v: {} };
  readAlertSheets(mkWorkbook([["admin", ADMIN]], [["R0001", COLLEAGUE, ADMIN, "2026-09-04 10:00"]]));
  useJobs([mkJob(), mkJob({ id: "R0002", cust: "Bob" })]);
  state.picked = { R0001: 1 };
  WHO = COLLEAGUE;
  renderChips();
  const chipText = () => EL["#chips"].kids.map(k => String(k.textContent));
  assert.strictEqual(chipText().indexOf("Alert to…"), -1, "no Alert to… in the selection bar");
  renderAlertsWindow();
  let win = EL["#ahost"].innerHTML;
  assert.ok(win.indexOf(COLLEAGUE) >= 0, "everyone can still see who is subscribed");
  assert.ok(win.indexOf("data-a-ask") < 0 && win.indexOf("data-a-yes") < 0, "but there is no remove button");
  assert.ok(win.indexOf(ALERT_ADMIN_ONLY) >= 0, "and the window says why");
  WHO = ADMIN;
  renderChips();
  assert.ok(chipText().indexOf("Alert to…") >= 0, "the admin gets the selection-bar action");
  renderAlertsWindow();
  assert.ok(EL["#ahost"].innerHTML.indexOf('data-a-ask="R0001|' + COLLEAGUE + '"') >= 0, "and a Remove button");
  ALCONF[alKey("R0001", COLLEAGUE)] = 1;
  renderAlertsWindow();
  win = EL["#ahost"].innerHTML;
  assert.ok(win.indexOf('data-a-yes="R0001|' + COLLEAGUE + '"') >= 0 && win.indexOf("data-a-no=") >= 0,
    "which asks inline, the same way the drawer chip does");
  delete ALCONF[alKey("R0001", COLLEAGUE)];
  state.picked = {};
  pass("a non-admin sees the list and no controls anywhere; the window asks before removing");

  /* ---- 17. the box empties itself after a save ---- */
  clearHolds(); reset();
  state.sel = "R0001"; renderDrawer();
  EL["#alnew"].value = "  OTHER@Example.TEST ";
  EL["#aladd"].onclick();
  await settle(80);
  assert.strictEqual(EL["#alnew"].value, "", "the address is a chip now, so the box is empty again");
  assert.deepStrictEqual(alertsFor("R0001").map(x => x.email), [COLLEAGUE, OTHER]);
  state.sel = null;
  pass("saving from the drawer clears the box instead of leaving the address in it");

  /* ---- 18. a refused bulk add stops at the first job ---- */
  BOOK["Dashboard Alerts"] = { v: {} };
  clearHolds();
  readAlertSheets(mkWorkbook([["admin", ADMIN]], []));
  useJobs([mkJob(), mkJob({ id: "R0002", cust: "Bob" }), mkJob({ id: "R0003", cust: "Cal" })]);
  reset();
  FAIL_ALERT_WRITE = 1;
  const stopped = await alertMany(["R0001", "R0002", "R0003"], COLLEAGUE);
  await settle();
  assert.strictEqual(stopped.ok, false);
  assert.strictEqual(stopped.added, 0);
  assert.ok(stopped.msg, "the message comes back so the menu can show it inline");
  assert.strictEqual(sheetRows().length, 0, "the other two jobs were never attempted");
  assert.strictEqual(TOASTS.filter(t => t.err).length, 1, "one toast, not one per job");
  assert.strictEqual(EL["#livedot"].className, "dot err", 'the status line is not put back to "live" after a failure');
  reset();
  const done = await alertMany(["R0001", "R0002"], COLLEAGUE);
  await settle();
  assert.strictEqual(done.ok, true);
  assert.strictEqual(done.added, 2);
  assert.strictEqual(EL["#livedot"].className, "dot", "a clean run does put it back");
  pass("Alert to… stops at the first refusal, says so once, and leaves the status showing it");

  /* ---- 19. an Alert line read back from the log is not cached here ---- */
  delete mem.cw_changes;
  vm.runInThisContext("CHANGES = [" +
    "{at:'2026-09-04 10:00', who:'x', job:'R0001', what:'Alert', from:'', to:'" + COLLEAGUE + "', src:'dashboard', shared:true}," +
    "{at:'2026-09-04 10:01', who:'x', job:'R0001', what:'Comment', from:'', to:'ring the customer', src:'dashboard', shared:true}]");
  saveChanges();
  assert.strictEqual(String(mem.cw_changes).indexOf(COLLEAGUE), -1,
    "the address does not go into the cached Changes list either");
  assert.ok(String(mem.cw_changes).indexOf("Comment") >= 0, "everything else still is");
  assert.strictEqual(CHANGES.length, 2, "and both lines are still on screen in this session");
  clearHolds();                    // the live holds from the run above are the only other place
  assert.deepStrictEqual(leaks(COLLEAGUE), [], "with them gone, nothing in this browser knows the address");
  pass("log lines about alerts are shown but never cached in this browser");

  console.log("\n" + n + " checks passed");
  process.exit(0);                 // the 45 s reconcile timer would hold the process open
})().catch(e => { console.error("FAIL", e); process.exit(1); });
