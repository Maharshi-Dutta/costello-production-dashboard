/* Offline test of the end-of-day sheet, the weekly target and the station
   report. Spec: docs/specs/2026-09-21-day-sheets-and-station-reports.md.

   What this file is here to hold, and why each of them matters:

     · ONE SHEET PER PERSON PER DAY. The Title is the key, the tablet refuses a
       second save and SharePoint refuses it again; a refusal that means "it is
       already there" is read as saved, never shown as a red error;
     · the tablet APPENDS AND NOTHING ELSE. Over the whole run it makes no
       PATCH and no DELETE of either new list, never writes `Station targets`,
       and - on a stage whose definition has no day sheet - never asks for
       either list at all. Asserted over every request, not by comment;
     · the week helper is the one both screens and the report use: Monday to
       Sunday, `2026-W39`, right across a year boundary, and LOCAL at 23:30;
     · the office's two writes are the whole of the new rule-2 exception: an
       upsert of the target and a PATCH of a sheet's counts and note. Each
       leaves one `Dashboard Log` line and no `Station log` line;
     · `stationReport` has no station in it: glass and welding both go through
       their own definitions, and rule 3 strips the free text of anything that
       could be a phone number or an eircode on its way into the file.

   Graph is a fake fetch() over one site; every person here is invented and
   there is no real address of any kind. Run: node test_daysheets.js         */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null),
                        setItem: (k, v) => { mem[k] = String(v); },
                        removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, innerWidth: 1280, innerHeight: 800,
                  addEventListener() {}, removeEventListener() {} };
global.performance = { now: () => Date.now() };
global.ExcelJS = ExcelJS;
global.Blob = function Blob(parts, o) { this.parts = parts; this.type = (o || {}).type || ""; };
global.URL.createObjectURL = () => "blob:test";
global.URL.revokeObjectURL = () => {};

const REG = {};
function stubEl(tag, id) {
  let html = "";
  const e = {
    tag: tag || "div", tagName: String(tag || "div").toUpperCase(),
    id: id || "", style: { setProperty() {} }, dataset: {}, attrs: {}, kids: [],
    textContent: "", value: "", disabled: false, hidden: false, className: "", title: "",
    href: "", download: "",
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); if (c.id) REG[c.id] = c; return c; },
    remove() { const drop = x => { if (x.id && REG[x.id] === x) delete REG[x.id]; x.kids.forEach(drop); }; drop(e); },
    contains(x) { return x === e || e.kids.some(k => k.contains && k.contains(x)); },
    setAttribute(k, v) { e.attrs[k] = String(v); }, getAttribute(k) { return e.attrs[k] || null; },
    on: {},
    addEventListener(t, fn) { (e.on[t] = e.on[t] || []).push(fn); },
    removeEventListener(t, fn) { e.on[t] = (e.on[t] || []).filter(x => x !== fn); },
    fire(t, ev) { (e.on[t] || []).slice().forEach(fn => fn(Object.assign({ type: t, target: e }, ev || {}))); },
    focus() {}, blur() {}, click() {}, setSelectionRange() {},
    getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => null, querySelectorAll: () => []
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; } });
  return e;
}
const NULLABLE = ["#fabhost", "#fabbtn", "#dhost", "#xhost", "#ahost", "#chost", "#vhost",
                  "#lhost", "#dayhost", "#nhost", "#catmenu", "#movemenu", "#alertmenu"];
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
   One site, six lists, and NO workbook route at all: anything that reaches for
   the file gets a 599 no retry rule matches, so a stray workbook call is a
   failure by construction rather than by assertion. */
const G = "https://graph.microsoft.com/v1.0";
const SITE = "costellowindowsie.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";
const GLASS_ID = "list-glass-station", PEOPLE_ID = "list-station-people";
const LOG_ID = "list-station-log", NOTES_ID = "list-station-comments";
const DAY_ID = "list-station-day-sheets", TARGET_ID = "list-station-targets";
const LISTS_PATH = "/sites/" + SITE + "/lists";
const OWN_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/ProductionProgress";
const HOST_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/FloorStations";

let DAY_EXISTS = true, TARGET_EXISTS = true;
let ITEMS = [], PEOPLEITEMS = [], LOGITEMS = [], NOTEITEMS = [], DAYITEMS = [], TARGETITEMS = [];
let NEXTID = 700;
let FAIL_DAY_POST = 0;                  // n POSTs to refuse outright
const lists = () => [{ id: GLASS_ID, displayName: "Glass station" },
                     { id: PEOPLE_ID, displayName: "Station people" },
                     { id: LOG_ID, displayName: "Station log" },
                     { id: NOTES_ID, displayName: "Station comments" }]
  .concat(DAY_EXISTS ? [{ id: DAY_ID, displayName: "Station day sheets" }] : [])
  .concat(TARGET_EXISTS ? [{ id: TARGET_ID, displayName: "Station targets" }] : []);

const REQ = [], ALLREQ = [];
const ok = body => ({ status: 200, body: body });
const item = (fields, id) => ({ id: String(id == null ? NEXTID++ : id), fields: Object.assign({}, fields) });
const storeFor = id => id === GLASS_ID ? ITEMS : id === PEOPLE_ID ? PEOPLEITEMS
                     : id === LOG_ID ? LOGITEMS : id === NOTES_ID ? NOTEITEMS
                     : id === DAY_ID ? DAYITEMS : id === TARGET_ID ? TARGETITEMS : null;

let DELTA_SEQ = 0;
function routeDelta(listId) {
  const store = storeFor(listId) || [];
  return ok({ value: store.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })),
              "@odata.deltaLink": G + LISTS_PATH + "/" + listId + "/items/delta?token=T" + (++DELTA_SEQ) });
}
function route(method, path, body) {
  if (path.indexOf(HOST_LOOKUP) === 0) return { status: 404, body: { error: { code: "itemNotFound" } } };
  if (path.indexOf(OWN_LOOKUP) === 0) return ok({ id: SITE, displayName: "Production Progress" });
  if (path.indexOf(LISTS_PATH) !== 0)
    return { status: 599, body: { error: { code: "thisTestServesNoWorkbook", message: path } } };
  const rest = path.slice(LISTS_PATH.length);
  if (method === "GET" && rest.indexOf("?$select=id,displayName") === 0) return ok({ value: lists() });
  const mi = /^\/([^/?]+)\/items(.*)$/.exec(rest);
  if (!mi) return { status: 404, body: { error: { code: "itemNotFound" } } };
  const id = mi[1], tail = mi[2];
  const store = storeFor(id);
  if (!store) return { status: 404, body: { error: { code: "itemNotFound" } } };
  if (method === "GET" && tail.indexOf("/delta") === 0) return routeDelta(id);
  if (method === "GET" && (tail === "" || tail.charAt(0) === "?"))
    return ok({ value: store.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) });
  if (method === "POST" && tail === "") {
    const fields = (body && body.fields) || {};
    /* a refusal no retry rule in graph.js will sit and wait on - the workshop
       wifi, in effect, without a real timer in the middle of the test */
    if (id === DAY_ID && FAIL_DAY_POST) { FAIL_DAY_POST--; return { status: 403, body: { error: { code: "accessDenied" } } }; }
    /* ENFORCE-UNIQUE-VALUES on Title, which is what the owner turns on for
       both new lists: a second save of the same key is refused here exactly as
       SharePoint refuses it. */
    if ((id === DAY_ID || id === TARGET_ID) &&
        store.some(x => String(x.fields.Title) === String(fields.Title)))
      return { status: 400, body: { error: { code: "invalidRequest",
                                             message: "The value must be unique" } } };
    const made = item(fields);
    store.push(made);
    return ok({ id: made.id, fields: made.fields });
  }
  const mp = /^\/([^/?]+)\/fields$/.exec(tail);
  if (method === "PATCH" && mp) {
    const hit = store.find(x => x.id === mp[1]);
    if (!hit) return { status: 404, body: { error: { code: "itemNotFound" } } };
    Object.assign(hit.fields, body || {});
    return ok(Object.assign({ id: hit.id }, hit.fields));
  }
  const mg = /^\/([^/?]+)(\?|$)/.exec(tail);
  if (method === "GET" && mg) {
    const hit = store.find(x => x.id === mg[1]);
    if (!hit) return { status: 404, body: { error: { code: "itemNotFound" } } };
    return ok({ id: hit.id, fields: Object.assign({}, hit.fields) });
  }
  if (method === "DELETE" && mg) {
    const at = store.findIndex(x => x.id === mg[1]);
    if (at >= 0) store.splice(at, 1);
    return { status: 204, body: "" };
  }
  return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + method + " " + path } } };
}
global.fetch = async (url, init) => {
  const path = String(url).replace(G, "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  const rec = { method: init.method, path: path, body: body };
  REQ.push(rec); ALLREQ.push(rec);
  const res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status,
           text: async () => (res.body === "" ? "" : JSON.stringify(res.body)),
           arrayBuffer: async () => new ArrayBuffer(0) };
};

/* ---------- the office's own code, in the page's own order ---------- */
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
global.WELDC = window.WELDC;
run("export.js");
run("app.js");

const TOASTS = [];
global.toast = (m, isErr) => TOASTS.push({ m: String(m), err: !!isErr });
const LOGGED = [];
CW.appendLog = async (who, job, what, from, to) => { LOGGED.push({ who, job, what, from, to }); return true; };

/* ---------- a tablet, in a context of its own ---------- */
const stationDoc = global.document;
const TIMERS = [];
const stationFetch = async (url, init) => {
  if (String(url).indexOf("version.json") === 0)
    return { ok: true, status: 200, json: async () => ({ build: "20260921-0900" }) };
  return global.fetch(url, init);
};
function newStation(stage) {
  mem.cw_stationstage = stage;
  const sb = {
    console: console,
    /* captured rather than scheduled: an owed write arms a retry, and a live
       five-second timer would keep node running after the last assertion */
    setTimeout: fn => { TIMERS.push(fn); return TIMERS.length; }, clearTimeout: () => {},
    setInterval: () => 0, clearInterval: () => {},
    localStorage: global.localStorage, document: stationDoc,
    window: { location: { origin: "http://localhost" } },
    location: { reload() {} }, fetch: stationFetch,
    CW: CW, ST: ST, confirm: () => true, prompt: () => null, alert: () => {}
  };
  sb.globalThis = sb;
  vm.createContext(sb);
  vm.runInContext(src("station.js"), sb, { filename: "station.js" });
  return sb;
}
const A = vm.runInThisContext.bind(vm);
const reset = () => { REQ.length = 0; TOASTS.length = 0; LOGGED.length = 0; };
const dayReq = () => REQ.filter(r => r.path.indexOf("/" + DAY_ID) >= 0);
const targetReq = () => REQ.filter(r => r.path.indexOf("/" + TARGET_ID) >= 0);
const consent = yes => { CW.hasListConsent = async () => !!yes; };
/* saveDay() sets its own flush going and does not await it - the tablet must
   never block a hand on the screen - so the test waits for it to finish rather
   than starting a second one that the `flushing` guard would turn straight back. */
const settle = async () => { for (let i = 0; i < 30; i++) await new Promise(r => setImmediate(r)); };

/* [column, the question the tablet asks, the heading a table puts over it] -
   written out here, from the spec, so every assertion is against the BRIEF and
   not against the definition's own copy of it. */
const COUNTS = [["Clear", "Clear glass sheets cut", "Clear"],
                ["KGlass", "K-glass sheets cut", "K-glass"],
                ["Satin", "Satin sheets cut", "Satin"],
                ["Obscure", "Other obscure sheets cut", "Obscure"]];
/* The columns the spec's Data model table names, written out here so every
   assertion is against the BRIEF and not against the code's own copy of it. */
const SPEC_DAY_COLUMNS = ["Title", "Station", "Stage", "Day", "Who", "Clear", "KGlass", "Satin",
                          "Obscure", "Note", "WeekTarget", "SavedAt", "EditedBy", "EditedAt"];
const SPEC_TARGET_COLUMNS = ["Title", "WeeklyTarget", "SetBy", "SetAt"];

const daySheetRow = (day, who, c, o, id) => item(Object.assign(
  { Title: "Glass|cut|" + day + "|" + who, Station: "Glass", Stage: "cut", Day: day, Who: who,
    Clear: c[0], KGlass: c[1], Satin: c[2], Obscure: c[3], Note: "",
    SavedAt: day + "T17:02:00.000Z" }, o || {}), id);

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ================= 1. the shape of the thing ================= */
  assert.strictEqual(ST.DAY_LIST, "Station day sheets");
  assert.strictEqual(ST.TARGET_LIST, "Station targets");
  assert.deepStrictEqual(ST.dayFieldsFor(ST.GLASS, "cut").slice().sort(),
    SPEC_DAY_COLUMNS.slice().sort(), "the columns the spec's Data model table names");
  assert.deepStrictEqual(ST.TARGET_FIELDS, SPEC_TARGET_COLUMNS);
  ["Phone", "PHONE NO.", "Eircode", "County", "Price", "Cell", "Colour"].forEach(k =>
    assert.ok(SPEC_DAY_COLUMNS.indexOf(k) < 0, k + " is not a column of the day sheets list"));
  pass("two lists, the spec's own columns, and not one of them a contact detail");

  /* the day sheet is switched on by the DEFINITION - which is the whole of
     "built so another page can switch it on later" */
  assert.ok(ST.daySheetOf(ST.GLASS, "cut"), "cutting has one");
  assert.strictEqual(ST.daySheetOf(ST.GLASS, "hotmelt"), null, "hotmelting has not");
  assert.strictEqual(ST.daySheetOf(ST.GLASS, "tuff"), null);
  assert.strictEqual(ST.daySheetOf(WELDC.WELD, "weld"), null, "and neither has welding, yet");
  assert.deepStrictEqual(ST.daySheetStages(ST.GLASS), ["cut"]);
  assert.deepStrictEqual(ST.daySheetOf(ST.GLASS, "cut").counts, COUNTS,
    "four counts, in the order the paper sheet asks for them");
  pass("which stage has a day sheet comes off the station definition, not off a page");

  /* ================= 2. the week helper ================= */
  const w = ST.isoWeek("2026-09-21");
  assert.strictEqual(w.key, "2026-W39");
  assert.strictEqual(w.monday, "2026-09-21", "Monday starts the week");
  assert.strictEqual(w.sunday, "2026-09-27", "and Sunday ends it");
  assert.strictEqual(w.weekday, "Monday");
  assert.strictEqual(ST.isoWeek("2026-09-27").monday, "2026-09-21", "Sunday belongs to the week before");
  assert.strictEqual(ST.isoWeek("2026-09-26").weekday, "Saturday", "Saturdays are worked and counted");
  assert.strictEqual(ST.isoWeek("2026-09-28").key, "2026-W40", "the next Monday starts the next week");
  pass("the ISO week runs Monday to Sunday and is labelled the way the office reads it");

  /* the year boundary, which is the one every hand-rolled week helper gets
     wrong: these two days are the same working week */
  assert.strictEqual(ST.isoWeek("2026-12-31").key, ST.isoWeek("2027-01-01").key,
    "2026-12-31 and 2027-01-01 are the same week");
  assert.strictEqual(ST.isoWeek("2026-12-31").monday, "2026-12-28");
  assert.strictEqual(ST.isoWeek("2027-01-03").key, ST.isoWeek("2026-12-31").key, "Sunday too");
  assert.notStrictEqual(ST.isoWeek("2027-01-04").key, ST.isoWeek("2026-12-31").key,
    "and 2027-01-04 starts the next one");
  assert.strictEqual(ST.isoWeek("2027-01-04").key, "2027-W01");
  assert.strictEqual(ST.isoWeek("").key, "", "rubbish in, nothing out - never a throw");
  assert.strictEqual(ST.isoWeek("not a day").key, "");
  pass("the week helper is right across a year boundary and silent on rubbish");

  /* LOCAL, not UTC: a sheet saved at 23:30 belongs to the day the person
     worked. A UTC date would put half the world's evenings on tomorrow. */
  const late = new Date(2026, 8, 21, 23, 30, 0);          // 21 Sep 2026, 23:30 local
  assert.strictEqual(ST.dayKey(late), "2026-09-21", "23:30 is still today, wherever this runs");
  const early = new Date(2026, 8, 21, 0, 15, 0);
  assert.strictEqual(ST.dayKey(early), "2026-09-21", "and so is a quarter past midnight");
  assert.strictEqual(ST.dayKey("2026-09-21"), "2026-09-21", "a day that is already a day is not re-parsed");
  assert.strictEqual(ST.dayKey("2026-09-21T23:30:00.000Z"), "2026-09-21");
  pass("the day a sheet belongs to is the local calendar day, at 23:30 as at noon");

  /* ================= 3. the row a save builds ================= */
  const f1 = ST.dayFields(ST.GLASS, "cut", {
    day: "2026-09-21", who: "Person A",
    counts: { Clear: "12", KGlass: 4, Satin: "", Obscure: "0" },
    note: "  air for cutting not working  ", weekTarget: 250, at: "2026-09-21T17:02:00.000Z" });
  assert.deepStrictEqual(Object.keys(f1).sort(),
    ["Title", "Station", "Stage", "Day", "Who", "Clear", "KGlass", "Satin", "Obscure",
     "Note", "WeekTarget", "SavedAt"].sort(), "exactly these columns, no more");
  assert.strictEqual(f1.Title, "Glass|cut|2026-09-21|Person A", "Title is the key");
  assert.strictEqual(f1.Clear, 12, "a typed number is a number");
  assert.strictEqual(f1.Satin, 0, "blank on paper is nought");
  assert.strictEqual(f1.Obscure, 0);
  assert.strictEqual(f1.Note, "air for cutting not working", "trimmed, otherwise as typed");
  assert.strictEqual(f1.WeekTarget, 250, "the target in force when it was saved rides along");
  assert.strictEqual(f1.SavedAt, "2026-09-21T17:02:00.000Z");
  assert.ok(!("EditedBy" in f1) && !("EditedAt" in f1), "the office's two columns are the office's");
  pass("the save's row: the key, the station, the stage, the day, the person, four counts and a note");

  /* no target set - the column is not written at all, rather than written as a
     nought that would read later as "the target was nothing" */
  assert.ok(!("WeekTarget" in ST.dayFields(ST.GLASS, "cut",
    { day: "2026-09-21", who: "Person A", counts: {} })), "no target set writes no WeekTarget");
  assert.strictEqual(ST.dayFields(ST.GLASS, "cut",
    { day: "2026-09-21", who: "Person A", counts: {}, note: "x".repeat(900) }).Note.length,
    ST.DAY_NOTE_MAX, "the note is capped by the row builder, not only by the box");
  pass("a sheet with no target and a very long note both come out right");

  /* a count that is not a whole number is refused OUTRIGHT - the row is not
     built, so there is nothing to queue and nothing to send */
  [-1, "-1", 1.5, "1.5", "twelve", "12a", "1e3"].forEach(bad =>
    assert.strictEqual(ST.dayFields(ST.GLASS, "cut",
      { day: "2026-09-21", who: "Person A", counts: { Clear: bad } }), null,
      JSON.stringify(bad) + " is refused at the form"));
  assert.strictEqual(ST.dayCount(""), 0);
  assert.strictEqual(ST.dayCount("007"), 7);
  assert.strictEqual(ST.dayCount(-1), null);
  assert.strictEqual(ST.dayCount(1.5), null);
  assert.strictEqual(ST.dayFields(ST.GLASS, "hotmelt", { day: "2026-09-21", who: "A", counts: {} }), null,
    "and a stage with no day sheet builds no row at all");
  pass("a negative, a decimal and a word are all refused at the form, not rounded");

  /* ================= 4. reading them back, and the weeks ================= */
  const rows = ST.dayRows([
    daySheetRow("2026-09-21", "Person A", [12, 4, 0, 2], { WeekTarget: 250 }),
    daySheetRow("2026-09-22", "Person A", [10, 0, 0, 0], { WeekTarget: 250, Note: "machine down" }),
    daySheetRow("2026-09-26", "Person B", [8, 8, 0, 0], { WeekTarget: 300 }),
    daySheetRow("2026-09-14", "Person A", [30, 0, 0, 0], { WeekTarget: 200 }),
    item({ Title: "x", Station: "Welding", Stage: "weld", Day: "2026-09-21", Who: "Person C" })
  ], COUNTS, { station: "Glass", stage: "cut" });
  assert.strictEqual(rows.length, 4, "another station's row is not this station's");
  assert.deepStrictEqual(rows.map(r => r.day),
    ["2026-09-26", "2026-09-22", "2026-09-21", "2026-09-14"], "newest day first");
  assert.strictEqual(rows[2].total, 18, "the total is the four counts added");
  assert.strictEqual(rows[2].week, "2026-W39");
  assert.strictEqual(rows[0].weekday, "Saturday", "a worked Saturday reads as one");
  pass("the list reads back as one row per person per day, newest first, this station's only");

  /* ===== D5: ONE ROW PER TITLE, whatever the list holds =====
     Enforce-unique-values on Title is the owner's step in SharePoint, and the
     code must not depend on it having been done: with the rule left off (or
     two rows typed in by hand) every duplicate doubled that person's day in
     the office window, in the week subtotal, in the board's line, in the
     tablet's own "this week" and in the report - silently, and in a direction
     that flatters the floor. Oldest item id wins, as everywhere else here. */
  const dupes = ST.dayRows([
    daySheetRow("2026-09-21", "Person A", [12, 0, 0, 0], { WeekTarget: 250 }, "10"),
    daySheetRow("2026-09-21", "Person A", [99, 0, 0, 0], { WeekTarget: 250 }, "11"),
    daySheetRow("2026-09-22", "Person A", [5, 0, 0, 0], {}, "12")
  ], COUNTS, {});
  assert.strictEqual(dupes.length, 2, "two rows carrying one Title are one day, not two");
  assert.strictEqual(dupes.find(r => r.day === "2026-09-21").counts.Clear, 12,
    "and it is the OLDEST id's row that is read, as everywhere else in this app");
  assert.strictEqual(ST.dayWeekTotal(dupes, "Person A", "2026-W39"), 17,
    "so the week reads 17, not 116");
  assert.strictEqual(ST.dayWeeks(dupes, COUNTS, 250)[0].total, 17, "and neither does the subtotal");
  /* a row with no Title at all cannot have come off a tablet and is nobody's
     duplicate: it keeps its own identity rather than colliding on "" */
  const untitled = ST.dayRows([
    item({ Station: "Glass", Stage: "cut", Day: "2026-09-21", Who: "Person A", Clear: 3 }, "20"),
    item({ Station: "Glass", Stage: "cut", Day: "2026-09-22", Who: "Person A", Clear: 4 }, "21")
  ], COUNTS, {});
  assert.strictEqual(untitled.length, 2, "two hand-made rows with no Title are still two rows");
  pass("D5: a duplicate row is read once, oldest id winning, everywhere a total is added up");

  assert.strictEqual(ST.dayWeekTotal(rows, "Person A", "2026-W39"), 28,
    "one person's week: 18 and 10");
  assert.strictEqual(ST.dayWeekTotal(rows, "", "2026-W39"), 44, "and the whole floor's, 28 and 16");
  assert.strictEqual(ST.dayWeekTotal(rows, "Person A", "2026-W38"), 30, "last week is its own");
  const weeks = ST.dayWeeks(rows, COUNTS, 400);
  assert.deepStrictEqual(weeks.map(g => g.week), ["2026-W39", "2026-W38"], "newest week first");
  assert.strictEqual(weeks[0].total, 44);
  assert.strictEqual(weeks[0].counts.Clear, 30);
  assert.strictEqual(weeks[0].target, 300,
    "the week's target is the one on its most recently saved sheet, not today's");
  assert.strictEqual(weeks[0].diff, -256);
  assert.strictEqual(weeks[1].target, 200, "so an old week stays true after the target changes");
  const noStored = ST.dayWeeks(ST.dayRows([daySheetRow("2026-09-21", "Person A", [5, 0, 0, 0])],
                                          COUNTS, {}), COUNTS, 250);
  assert.strictEqual(noStored[0].target, 250, "a week with no stored target falls back to the live one");
  assert.strictEqual(ST.dayWeeks([], COUNTS, null).length, 0);
  pass("week subtotals use the target that was in force in that week, and fall back to the live one");

  /* the office's filters */
  const byWeekday = ST.dayRows([
    daySheetRow("2026-09-21", "Person A", [1, 0, 0, 0]),
    daySheetRow("2026-09-22", "Person A", [2, 0, 0, 0]),
    daySheetRow("2026-09-28", "Person B", [3, 0, 0, 0])
  ], COUNTS, { weekday: 0 });
  assert.deepStrictEqual(byWeekday.map(r => r.day), ["2026-09-28", "2026-09-21"], "Mondays only");
  assert.strictEqual(ST.dayRows([
    daySheetRow("2026-09-21", "Person A", [1, 0, 0, 0]),
    daySheetRow("2026-09-28", "Person B", [3, 0, 0, 0])
  ], COUNTS, { week: "2026-W39" }).length, 1, "one week only");
  assert.strictEqual(ST.dayRows([
    daySheetRow("2026-09-21", "Person A", [1, 0, 0, 0]),
    daySheetRow("2026-09-28", "Person B", [3, 0, 0, 0])
  ], COUNTS, { who: "person b" }).length, 1, "one person only, whatever the case");
  assert.strictEqual(ST.dayRows([
    daySheetRow("2026-09-21", "Person A", [1, 0, 0, 0]),
    daySheetRow("2026-09-28", "Person B", [3, 0, 0, 0])
  ], COUNTS, { from: "2026-09-22", to: "2026-09-30" }).length, 1, "and a date range");
  pass("the window's filters: weekday, week, person and a date range");

  /* ================= 5. the target ================= */
  const tf = ST.targetFields(250, "the office", "2026-09-21T09:00:00.000Z");
  assert.deepStrictEqual(Object.keys(tf).sort(), ["WeeklyTarget", "SetBy", "SetAt"].sort(),
    "three columns; listUpsert puts the Title on");
  assert.strictEqual(tf.WeeklyTarget, 250);
  assert.strictEqual(ST.targetTitle("Glass", "cut"), "Glass|cut");
  const got = ST.targetOf([item({ Title: "Glass|cut", WeeklyTarget: 250, SetBy: "the office",
                                  SetAt: "2026-09-21T09:00:00.000Z" }, "1"),
                           item({ Title: "Welding|weld", WeeklyTarget: 900 }, "2")], "Glass", "cut");
  assert.strictEqual(got.target, 250);
  assert.strictEqual(got.by, "the office");
  assert.strictEqual(ST.targetOf([], "Glass", "cut"), null, "no row at all is no target, not nought");
  assert.strictEqual(ST.targetOf([item({ Title: "Glass|cut", WeeklyTarget: "" })], "Glass", "cut"), null);
  pass("the target is one row per station-stage, and a missing one is “no target”, never nought");

  /* ================= 6. the tablet, end to end ================= */
  consent(true);
  CW._resetListIds(); delete mem.cw_listids;
  CW._setStationSite(SITE, false);
  ITEMS = [item({ Title: "R0001", Job: "R0001", Customer: "Customer One", GlassType: "GLASS",
                  Total: 6, TuffTotal: 0, Seq: 1, Active: "Yes", OfficeDone: "No",
                  Cut: 0, Hotmelt: 0, Tuff: 0 }, "900")];
  PEOPLEITEMS = [item({ Title: "Person A", Station: "Glass", Stages: "cut,tuff", PIN: "", Active: "Yes" }),
                 item({ Title: "Person B", Station: "Glass", Stages: "hotmelt", PIN: "", Active: "Yes" })];
  DAYITEMS = []; TARGETITEMS = [item({ Title: "Glass|cut", WeeklyTarget: 250,
                                       SetBy: "the office", SetAt: "2026-09-20T09:00:00.000Z" })];
  delete mem.cw_stationq; delete mem.cw_stationlogq; delete mem.cw_stationdayq;
  delete mem.cw_daysheetdraft;

  const cut = newStation("cut");
  const S = code => vm.runInContext(code, cut);
  S("SITEID = " + JSON.stringify(SITE) + "; QUEUE = {}; LOGQ = {}; DAYQ = {}; TOKEN = null;");
  assert.strictEqual(await S("readPeople()"), true);
  await S("readList()");
  S("pickPerson(PEOPLE.find(p => p.name === 'Person A'))");
  assert.strictEqual(S("PERSON.name"), "Person A");
  reset();
  assert.strictEqual(await S("readDay()"), true, "the cutting tablet reads both lists");
  assert.strictEqual(S("TARGET"), 250, "and knows the office's target");
  assert.ok(dayReq().length === 1 && targetReq().length === 1, "one read each, no more");
  assert.ok(dayReq().every(r => r.method === "GET") && targetReq().every(r => r.method === "GET"));
  pass("the cutting tablet reads the day sheets and the target, once each, read-only");

  /* the form: the draft survives a redraw and a reload, and the week line adds
     what is typed to what is already saved */
  const TODAY = ST.dayKey(new Date());
  DAYITEMS = [daySheetRow(ST.isoWeek(TODAY).monday, "Person A", [20, 0, 0, 0], { WeekTarget: 250 })];
  await S("readDay()");
  S("openDaySheet(); draftNow().counts.Clear = '12'; draftNow().counts.KGlass = '3'; saveDraft();");
  assert.strictEqual(S("draftTotal()"), 15, "what is typed now adds up");
  const wk = S("JSON.stringify(weekWords(draftTotal()))");
  assert.strictEqual(JSON.parse(wk).done, 35, "this week is the saved 20 plus the typed 15");
  assert.ok(JSON.parse(wk).words.indexOf("35 of 250") >= 0, "“35 of 250 sheets”");
  assert.ok(String(mem.cw_daysheetdraft).indexOf('"Clear":"12"') >= 0,
    "and the draft is in localStorage, so a poll or a reload cannot lose it");
  pass("the form's running totals: today, and this week including what is typed but not saved");

  /* ===== D4: A DRAFT BELONGS TO THE DAY IT WAS STARTED =====
     The first build keyed the draft on "today", so a sheet typed at 23:55 and
     saved at 00:01 was filed under tomorrow - the wrong day, the wrong ISO
     week, and possibly the wrong target - or, if the tick landed first, thrown
     away entirely. Both are tested here with a draft dated YESTERDAY, which is
     the same case the clock produces once a night. */
  const YESTERDAY = ST.dayKey(new Date(Date.parse(TODAY + "T12:00:00Z") - 86400000));
  const DAYBEFORE = ST.dayKey(new Date(Date.parse(TODAY + "T12:00:00Z") - 2 * 86400000));
  mem.cw_daysheetdraft = JSON.stringify({ day: YESTERDAY, stage: "cut", who: "Person A",
                                          counts: { Clear: "9" }, note: "late finish", target: 200 });
  S("DAYDRAFT = null;");
  assert.strictEqual(S("draftNow().day"), YESTERDAY,
    "a draft started yesterday is still yesterday's - not thrown away, and not re-dated");
  assert.strictEqual(S("draftNow().counts.Clear"), "9", "with what was typed into it");
  assert.strictEqual(S("sheetDay()"), YESTERDAY, "and that is the day the sheet on screen is about");
  assert.ok(S("daySheetHtml()").indexOf(YESTERDAY) >= 0, "the header says which day it will save under");
  assert.ok(S("daySheetHtml()").indexOf("daycarry") >= 0, "and says so in as many words");
  assert.strictEqual(S("weekWords(0).target"), 200,
    "measured against the target that was in force then, not the one set since");
  reset();
  assert.strictEqual(await S("saveDay()"), true);
  await settle();
  const carried = dayReq().filter(r => r.method === "POST");
  assert.strictEqual(carried.length, 1);
  assert.strictEqual(carried[0].body.fields.Day, YESTERDAY, "and Save files it under THAT day");
  assert.strictEqual(carried[0].body.fields.Title, "Glass|cut|" + YESTERDAY + "|Person A");
  assert.strictEqual(carried[0].body.fields.WeekTarget, 200, "with that day's target on it");
  assert.strictEqual(mem.cw_daysheetdraft, undefined, "the draft is let go once it is saved");
  pass("D4: a draft started yesterday saves under yesterday, its week and its target");

  /* ... and the day after that it is gone: a sheet nobody saved is not left
     sitting there to be filed under a day three weeks late */
  mem.cw_daysheetdraft = JSON.stringify({ day: DAYBEFORE, stage: "cut", who: "Person A",
                                          counts: { Clear: "99" }, note: "old" });
  S("DAYDRAFT = null;");
  assert.strictEqual(S("draftNow().day"), TODAY, "a draft two days old is dropped");
  assert.strictEqual(S("draftNow().counts.Clear"), undefined);
  pass("D4: a draft survives to the end of the following day and no longer");

  /* ===== D4: an untouched form is not four noughts =====
     with today's own sheet not yet saved, or the form would be the read-only
     one - today can be the Monday the fixture row above sits on */
  DAYITEMS = [];
  await S("readDay()");
  reset();
  S("clearDraft(); DAYDRAFT = null;");
  assert.strictEqual(S("draftOk()"), false, "Save is dead on a form nobody has typed into");
  assert.ok(S("daySheetHtml()").indexOf('id="daysave"') >= 0);
  assert.ok(/id="daysave"[^>]*disabled/.test(S("daySheetHtml()")), "and it is drawn disabled");
  assert.strictEqual(await S("saveDay()"), false, "and it refuses even if the click gets through");
  assert.strictEqual(dayReq().filter(r => r.method === "POST").length, 0, "so nothing is sent");
  S("draftNow().note = 'machine down all day'; saveDraft();");
  assert.strictEqual(S("draftOk()"), true,
    "a day with nothing cut but a line about why is a real entry and saves");
  assert.strictEqual(S("draftTotal()"), 0);
  S("clearDraft(); DAYDRAFT = null; draftNow().counts.Clear = '0'; saveDraft();");
  assert.strictEqual(S("draftOk()"), true, "and so is a deliberate nought typed into a box");
  pass("D4: an untouched form cannot be saved; a note alone, or a typed nought, can");

  /* ===== the count cap ===== */
  S("clearDraft(); DAYDRAFT = null; draftNow().counts.Clear = '10000'; saveDraft();");
  assert.strictEqual(S("draftOk()"), false, "ten thousand sheets in a day is a stuck finger");
  assert.strictEqual(await S("saveDay()"), false);
  assert.strictEqual(ST.dayCount(9999), 9999, "the cap itself is fine");
  assert.strictEqual(ST.dayCount(10000), null);
  assert.strictEqual(ST.dayFields(ST.GLASS, "cut",
    { day: TODAY, who: "Person A", counts: { Clear: 10000 } }), null,
    "and the row builder refuses it too, not only the form");
  pass("a count over " + ST.DAY_COUNT_MAX + " is refused at the form and in the row builder");

  /* no target set at all says so, rather than "of 0" */
  S("clearDraft(); DAYDRAFT = null; TARGET = null;");
  assert.ok(S("weekWords(0).words").indexOf("no target set") >= 0, "the words, not “of 0”");
  assert.ok(S("weekWords(0).words").indexOf("of 0") < 0);
  S("TARGET = 250; clearDraft(); DAYDRAFT = null;");
  pass("with no target the tablet says “no target set”, never “of 0”");

  /* the save: one POST, the right row, and nothing else touched. The fixture
     week starts clean first - today can BE the Monday the row above sits on,
     and then this person would already have saved today. */
  reset();
  DAYITEMS = [];
  await S("readDay()");
  S("DAYDRAFT = null; draftNow().counts.Clear = '12'; draftNow().counts.KGlass = '3'; saveDraft();");
  assert.strictEqual(await S("saveDay()"), true);
  await settle();
  const posts = dayReq().filter(r => r.method === "POST");
  assert.strictEqual(posts.length, 1, "exactly one POST");
  const sent = posts[0].body.fields;
  assert.strictEqual(sent.Title, "Glass|cut|" + TODAY + "|Person A");
  assert.strictEqual(sent.Clear, 12);
  assert.strictEqual(sent.KGlass, 3);
  assert.strictEqual(sent.Satin, 0, "the boxes left blank are noughts");
  assert.strictEqual(sent.WeekTarget, 250, "with the target that was in force");
  assert.strictEqual(mem.cw_daysheetdraft, undefined, "and the draft is cleared");
  assert.strictEqual(targetReq().filter(r => r.method !== "GET").length, 0,
    "the tablet never writes the target list");
  pass("one tap of Save is one POST of one row, and the draft is let go");

  /* it cannot be saved twice: the page refuses, and the list would too */
  reset();
  assert.strictEqual(await S("saveDay()"), false, "the page refuses a second save");
  await settle();
  assert.strictEqual(dayReq().filter(r => r.method === "POST").length, 0, "so no POST is made");
  assert.ok(S("!!daySheetSaved(sheetDay())"), "the sheet reads as saved");
  assert.ok(S("daySheetHtml()").indexOf(ST.DAY_SAVED_WORDS) >= 0,
    "and says to ask the office about a mistake");
  pass("one sheet per person per day: the second attempt makes no request at all");

  /* a replayed queue item that already landed, or a second tablet saving the
     same key: SharePoint's unique refusal is read as ALREADY SAVED */
  reset();
  S("DAYQ[" + JSON.stringify("Glass|cut|" + TODAY + "|Person A") + "] = { fields: " +
    JSON.stringify(Object.assign({}, sent)) + ", err: 0 };");
  await S("flushDay()");
  assert.strictEqual(S("Object.keys(DAYQ).length"), 0, "the owed row is let go");
  assert.strictEqual(S("!!daySheetSaved(sheetDay())"), true, "and the saved sheet is what is shown");
  assert.ok(S("daySheetHtml()").indexOf("waiting to send") < 0, "never a red error");
  assert.strictEqual(DAYITEMS.filter(x => x.fields.Who === "Person A" && x.fields.Day === TODAY).length, 1,
    "and the list still holds exactly one row for that person and day");
  pass("a unique-value refusal is read as already saved, not as a failure");

  /* offline at save: the row is owed, and says "waiting to send" */
  reset();
  DAYITEMS = [];
  await S("readDay()");
  FAIL_DAY_POST = 2;
  S("DAYDRAFT = null; draftNow().counts.Clear = '7'; saveDraft();");
  assert.strictEqual(await S("saveDay()"), true);
  await settle();
  assert.strictEqual(S("Object.keys(DAYQ).length"), 1, "still owed");
  assert.ok(S("daySheetHtml()").indexOf("waiting to send") >= 0, "and it says so, rather than “saved”");
  FAIL_DAY_POST = 0;
  await S("flushQueue()");
  await settle();
  assert.strictEqual(S("Object.keys(DAYQ).length"), 0, "and it goes when the connection is back");
  assert.strictEqual(S("!!daySheetSaved(sheetDay())"), true);
  pass("a save made offline is owed, says “waiting to send”, and lands when the wifi is back");

  /* ===== D2: A DRAFT IS ONE PERSON'S WRITING =====
     The tablet is passed between three people. The draft was keyed on the day
     and the stage only and the sheet was left open across a change of person,
     so the next name picked was handed the last person's half-typed numbers on
     a form whose Save would have filed them under the NEW name - on a row
     neither of them could correct afterwards. */
  reset();
  DAYITEMS = [];
  await S("readDay()");
  S("clearDraft(); DAYDRAFT = null; openDaySheet(); draftNow().counts.Clear = '11'; saveDraft();");
  assert.strictEqual(S("DAYOPEN"), true, "Person A has the sheet open with 11 in it");
  assert.strictEqual(S("draftNow().who"), "Person A", "and the draft says whose writing it is");
  S("switchPerson()");
  assert.strictEqual(S("DAYOPEN"), false, "switching person closes the sheet");
  S("pickPerson(PEOPLE.find(p => p.name === 'Person A'))");   // a second cutter, same tablet
  vm.runInContext("PERSON = { name: 'Person E', stages: ['cut'] };", cut);
  assert.strictEqual(S("draftNow().counts.Clear"), undefined,
    "the next person is handed a blank form, never somebody else's numbers");
  assert.strictEqual(S("draftNow().who"), "Person E");
  /* ... and Person A's writing is not lost: it is theirs, and it comes back */
  vm.runInContext("PERSON = { name: 'Person A', stages: ['cut'] }; DAYDRAFT = null;", cut);
  assert.strictEqual(S("draftNow().counts.Clear"), "11", "and theirs comes back when they do");
  /* the ten-minute idle lock goes through switchPerson, so it closes it too */
  S("openDaySheet(); LAST_TAP = Date.now() - ST.PERSON_LOCK_MS - 1000; lockIfIdle();");
  assert.strictEqual(S("DAYOPEN"), false, "and a tablet left alone closes the sheet as well");
  vm.runInContext("PERSON = PEOPLE.find(p => p.name === 'Person A'); DAYDRAFT = null;", cut);
  S("clearDraft(); DAYDRAFT = null;");
  pass("D2: a draft belongs to one person; switching person, or the idle lock, closes the sheet");

  /* ===== D7: unreachable is not missing =====
     A sheet typed out at the end of a shift on a tablet that cannot reach
     SharePoint must be QUEUED, exactly like every other offline write on this
     page. The first build refused it - because DAY_OK is false for "no such
     list" and for "the wifi is out" alike - and the day was lost. */
  reset();
  S("DAY_OK = false; DAY_MISSING = false; DAY_WHY = ST.DAY_UNREACHABLE;");
  S("clearDraft(); DAYDRAFT = null; draftNow().counts.Clear = '6'; saveDraft();");
  assert.strictEqual(await S("saveDay()"), true, "a list that could not be READ still takes the save");
  assert.strictEqual(S("Object.keys(DAYQ).length"), 1, "it is owed, like any other offline write");
  await settle();
  assert.strictEqual(S("Object.keys(DAYQ).length"), 0, "and goes the moment the list answers again");
  /* ... and a list that genuinely is not there still refuses, with the words */
  reset();
  DAYITEMS = [];
  await S("readDay()");
  S("DAY_OK = false; DAY_MISSING = true; DAY_WHY = ST.DAY_MISSING_FLOOR;");
  S("clearDraft(); DAYDRAFT = null; draftNow().counts.Clear = '6'; saveDraft();");
  assert.strictEqual(await S("saveDay()"), false, "a list that is not there refuses the save");
  assert.strictEqual(S("Object.keys(DAYQ).length"), 0, "nothing is queued for a list that cannot exist");
  assert.ok(S("DAYBAD").indexOf("Station day sheets") >= 0, "and the form says which list it is");
  S("DAY_OK = true; DAY_MISSING = false; DAY_WHY = ''; DAYBAD = ''; clearDraft(); DAYDRAFT = null;");
  pass("D7: a save is queued when the list is unreachable and refused only when it is missing");

  /* ===== D6: a row SharePoint keeps refusing =====
     "waiting to send" invites somebody to stand there waiting for wifi that is
     working perfectly well. After three refusals the words change - and the
     row stays queued, because the tablet cannot show it and dropping it is the
     only way the day would really be lost. */
  reset();
  DAYITEMS = [];
  await S("readDay()");
  FAIL_DAY_POST = 9;
  S("clearDraft(); DAYDRAFT = null; draftNow().counts.Clear = '5'; saveDraft();");
  assert.strictEqual(await S("saveDay()"), true);
  await settle();
  assert.strictEqual(S("Object.keys(DAYQ).length"), 1);
  assert.ok(S("daySheetHtml()").indexOf("waiting to send") >= 0, "once refused it is still the wifi");
  await S("flushDay()"); await settle();
  await S("flushDay()"); await settle();
  assert.ok(S("Object.values(DAYQ)[0].refused >= 3"), "three refusals of the same row");
  assert.ok(S("daySheetHtml()").indexOf("could not be saved — tell the office") >= 0,
    "and now it says so, instead of pointing at the wifi");
  assert.ok(S("daySheetHtml()").indexOf("waiting to send") < 0);
  assert.strictEqual(S("Object.keys(DAYQ).length"), 1, "and it is still queued, never thrown away");
  /* it survives a reload with its count: a refused row must not read as brand
     new every time the tablet is restarted */
  const reloaded = newStation("cut");
  assert.strictEqual(vm.runInContext("Object.values(DAYQ)[0].refused >= 3", reloaded), true,
    "the refusal count is rebuilt from localStorage with the row");
  FAIL_DAY_POST = 0;
  await S("flushDay()"); await settle();
  assert.strictEqual(S("Object.keys(DAYQ).length"), 0, "and it still goes the moment it can");
  pass("D6: three refusals change the words, keep the row, and survive a reload");

  /* ===== D6: an owed day sheet does not pin the tablet on an old build =====
     It is in localStorage and rebuilt on load, so a reload loses nothing - and
     a new build is exactly what a list that keeps refusing may need. */
  reset();
  DAYITEMS = [];
  await S("readDay()");                 // today's sheet is not on the list again
  FAIL_DAY_POST = 9;
  S("clearDraft(); DAYDRAFT = null; draftNow().counts.Clear = '4'; saveDraft();");
  await S("saveDay()"); await settle();
  assert.strictEqual(S("Object.keys(DAYQ).length"), 1, "a sheet is owed");
  assert.strictEqual(S("owingWrites()"), false, "but no COUNTER or log line is");
  assert.strictEqual(S("owingAnything()"), true, "the retry timer still knows about it");
  let reloads = 0;
  cut.location.reload = () => { reloads++; };
  S("BUILD_NOW = '20260821-0900';");     // the build this tablet is running
  await S("checkBuild()");               // version.json says 20260921-0900
  assert.strictEqual(reloads, 1, "a new build reloads the tablet with a day sheet owed");
  /* ... and the owed sheet is still there afterwards, because it is in storage */
  const afterReload = newStation("cut");
  assert.strictEqual(vm.runInContext("Object.keys(DAYQ).length", afterReload), 1,
    "the owed sheet is rebuilt on the next load: the reload cost nothing");
  FAIL_DAY_POST = 0;
  await S("flushDay()"); await settle();
  pass("D6: an owed day sheet no longer pins the tablet on an old build, and survives the reload");

  /* THE WHOLE RUN so far, over both new lists: append and read, and nothing
     else - no PATCH, no DELETE, and never a write of the target list */
  const tabletTouched = ALLREQ.filter(r => r.path.indexOf("/" + DAY_ID) >= 0 || r.path.indexOf("/" + TARGET_ID) >= 0);
  assert.ok(tabletTouched.length > 0, "the tablet did touch them, or this proves nothing");
  assert.ok(tabletTouched.every(r => r.method === "GET" || r.method === "POST"),
    "the tablet only ever reads or appends: " + tabletTouched.map(r => r.method).join(","));
  assert.ok(ALLREQ.filter(r => r.path.indexOf("/" + TARGET_ID) >= 0).every(r => r.method === "GET"),
    "and `Station targets` is read-only on the tablet");
  assert.strictEqual(ALLREQ.filter(r => /\/workbook|\/drive/.test(r.path)).length, 0,
    "and nothing on this page has been anywhere near the workbook");
  pass("over the whole tablet run: reads and appends only, no PATCH, no DELETE, no workbook");

  /* a stage with no day sheet: no button, and NOT ONE REQUEST for either list */
  reset();
  const hot = newStation("hotmelt");
  const H = code => vm.runInContext(code, hot);
  H("SITEID = " + JSON.stringify(SITE) + "; QUEUE = {}; LOGQ = {}; DAYQ = {}; TOKEN = null;");
  await H("readPeople()");
  await H("readList()");
  H("pickPerson(PEOPLE.find(p => p.name === 'Person B'))");
  assert.strictEqual(H("PAGE_STAGE"), "hotmelt");
  assert.strictEqual(H("daySheet()"), null, "the hotmelting tablet has no day sheet");
  assert.strictEqual(await H("readDay()"), false, "asking it to read one does nothing");
  assert.strictEqual(await H("saveDay()"), false, "and there is nothing to save");
  H("render()");
  assert.strictEqual(dayReq().length, 0, "not one request for the day sheets list");
  assert.strictEqual(targetReq().length, 0, "and none for the target list either");
  pass("a stage whose definition has no day sheet draws no button and asks for neither list");

  /* ================= 7. the office ================= */
  reset();
  A("STATION_ITEMS = null; STATION_OK = null; DAY_ITEMS = null; DAY_OK = null; TARGET_ITEMS = null;");
  DAYITEMS = [daySheetRow("2026-09-21", "Person A", [12, 4, 0, 2], { WeekTarget: 250 }),
              daySheetRow("2026-09-22", "Person A", [10, 0, 0, 0], { WeekTarget: 250, Note: "machine down" }),
              daySheetRow("2026-09-14", "Person A", [30, 0, 0, 0], { WeekTarget: 200 })];
  assert.ok(await A("readDaySheets()"), "the office reads them");
  assert.strictEqual(A("DAY_OK"), true);
  assert.strictEqual(A("dayTargetNow().target"), 250);
  assert.strictEqual(A("dayStage()"), "cut", "the stage comes off the definition");
  pass("the office reads the day sheets and the target from the station's own site");

  /* the target: one upsert, one Dashboard Log line, no Station log line */
  reset();
  assert.strictEqual(await A("saveDayTarget('300')"), true);
  const tPosts = targetReq().filter(r => r.method === "POST" || r.method === "PATCH");
  assert.strictEqual(tPosts.length, 1, "one write of the target row");
  const tBody = tPosts[0].body.fields || tPosts[0].body;
  assert.strictEqual(tBody.WeeklyTarget, 300);
  assert.strictEqual(tBody.Title, "GLASS|CUT", "listUpsert puts its own upper-cased key on");
  assert.ok(!("Clear" in tBody) && !("Day" in tBody), "and nothing else can get into it");
  assert.strictEqual(LOGGED.length, 1, "exactly one Dashboard Log line");
  assert.strictEqual(LOGGED[0].what, "Cutting weekly target");
  assert.strictEqual(LOGGED[0].from, "250");
  assert.strictEqual(LOGGED[0].to, "300");
  assert.strictEqual(REQ.filter(r => r.path.indexOf("/" + LOG_ID) >= 0 && r.method === "POST").length, 0,
    "and never a line of the floor's own Station log");
  assert.strictEqual(await A("saveDayTarget('300')"), false, "setting it to what it already is says nothing");
  pass("the office's target: one upsert, one log line, and no Station log line ever");

  /* ===== D10: a target is at least one =====
     Clearing the box and tapping Save wrote a target of NOUGHT, which is not
     "no target": every week then reads "+44" as if it had beaten it, and the
     tablet says "35 of 0". Removing a target is not built, so an empty box is
     a mistake - refused, with nothing written at all. */
  reset();
  assert.strictEqual(await A("saveDayTarget('')"), false, "an empty box saves nothing");
  assert.strictEqual(await A("saveDayTarget('0')"), false, "and neither does a nought");
  assert.strictEqual(await A("saveDayTarget('-5')"), false);
  assert.strictEqual(await A("saveDayTarget('2.5')"), false);
  assert.strictEqual(await A("saveDayTarget('lots')"), false);
  assert.strictEqual(targetReq().filter(r => r.method !== "GET").length, 0,
    "not one write of the target list between them");
  assert.strictEqual(LOGGED.length, 0, "and nothing in Dashboard Log either");
  assert.strictEqual(TOASTS.length, 5, "each one says so");
  assert.ok(TOASTS.every(t => t.err && /one or more/.test(t.m)));
  assert.strictEqual(A("dayTargetNow().target"), 300, "the target that was there is untouched");
  assert.strictEqual(await A("saveDayTarget('1')"), true, "one is a target");
  pass("D10: a target must be a whole number of one or more, or nothing is written");

  /* a correction: exactly the allowed columns, one log line, no delete */
  reset();
  const rowId = DAYITEMS[0].id;
  assert.strictEqual(await A("saveDayEdit(" + JSON.stringify(rowId) +
    ", { counts: { Clear: '15', KGlass: '4', Satin: '0', Obscure: '2' }, note: 'corrected' })"), true);
  const patches = dayReq().filter(r => r.method === "PATCH");
  assert.strictEqual(patches.length, 1, "one PATCH");
  assert.deepStrictEqual(Object.keys(patches[0].body).sort(),
    ["Clear", "KGlass", "Satin", "Obscure", "Note", "EditedBy", "EditedAt"].sort(),
    "exactly the columns the spec allows the office to write");
  assert.strictEqual(patches[0].body.Clear, 15);
  assert.strictEqual(patches[0].body.Note, "corrected");
  assert.ok(patches[0].body.EditedBy, "stamped with who corrected it");
  assert.strictEqual(dayReq().filter(r => r.method === "DELETE").length, 0, "and nothing is ever deleted");
  assert.strictEqual(LOGGED.length, 1);
  assert.strictEqual(LOGGED[0].what, "Day sheet corrected");
  assert.ok(LOGGED[0].from.indexOf("Person A") >= 0 && LOGGED[0].from.indexOf("18") >= 0,
    "the log line names the day, the person and the old total");
  assert.strictEqual(LOGGED[0].to, "21", "and the new one");
  pass("an office correction PATCHes exactly the allowed columns, logs it, and deletes nothing");

  /* the builder is the proof: it takes counts, a note, a name and a time, so
     there is no argument that could carry a Day, a Who or a Title in */
  const off = ST.dayOfficeFields(COUNTS, { counts: { Clear: 1, KGlass: 2, Satin: 3, Obscure: 4 },
                                           note: "x", who: "the office", at: "2026-09-21T10:00:00.000Z" });
  assert.deepStrictEqual(Object.keys(off).sort(),
    ["Clear", "KGlass", "Satin", "Obscure", "Note", "EditedBy", "EditedAt"].sort());
  assert.strictEqual(ST.dayOfficeFields(COUNTS, { counts: { Clear: -2 } }), null,
    "and a bad count refuses the whole correction");
  pass("the correction's body can only be counts, a note and who corrected it");

  /* the two windows are BUILT here, not only reasoned about: a typo in a
     string of markup is a window that throws when somebody opens it, and
     nothing else in this suite would ever call these. */
  reset();
  A("openDaySheets();");
  /* opening a window speeds the floor's poll up, which arms a timer that
     re-arms itself: disarmed here, or node would never exit */
  A("if (stationPollT) { clearTimeout(stationPollT); stationPollT = null; }");
  const win = REG["dayhost"] ? REG["dayhost"].innerHTML : "";
  assert.ok(win.indexOf("day sheets") >= 0, "the window opens");
  assert.ok(win.indexOf('id="dstarget"') >= 0, "with the target control");
  assert.ok(win.indexOf('id="dsdow"') >= 0 && win.indexOf('id="dsweek"') >= 0 &&
            win.indexOf('id="dswho"') >= 0, "and the weekday, week and person filters");
  A("DAYF.from = '2026-09-01'; DAYF.to = '2026-09-30'; paintDaySheets();");
  const html = sel => String(A("($(" + JSON.stringify(sel) + ") || {}).innerHTML || ''"));
  const body = html("#dsbody");
  assert.ok(body.indexOf("2026-W39") >= 0, "the rows are grouped into weeks");
  assert.ok(body.indexOf("Week total") >= 0, "with a subtotal row under each");
  assert.ok(body.indexOf("target 250") >= 0, "measured against the week's own stored target");
  assert.ok(body.indexOf("data-dsedit") >= 0, "and every row can be corrected");
  A("DAYF.from = '2026-09-01'; DAYF.to = '2026-09-30'; DAYF.week = '2026-W38'; paintDaySheets();");
  const one = html("#dsbody");
  assert.ok(one.indexOf("2026-W38") >= 0 && one.indexOf("2026-W39") < 0, "the week filter narrows it");
  A("DAYF.week = ''; DAYF.weekday = '0'; paintDaySheets();");
  assert.ok(html("#dsbody").indexOf("2026-09-21") >= 0, "and the weekday filter picks Mondays");
  pass("the day sheets window builds: the target control, the filters, the weeks and their subtotals");

  /* ===== the heading row =====
     Without it the four counts are bare numbers - "12 4 0 2 18" with nothing
     to say which column is K-glass (seen in a screenshot). The labels come off
     the STATION DEFINITION, in the same order as the cells under them, so a
     station with three counts or five gets its own headings and app.js keeps
     no list of its own. */
  A("DAYF.weekday = ''; DAYF.from = '2026-09-01'; DAYF.to = '2026-09-30'; paintDaySheets();");
  const listed = html("#dsbody");
  const headAt = listed.indexOf('class="dsrow dshead"');
  assert.ok(headAt >= 0, "there is one heading row");
  assert.strictEqual(listed.indexOf('class="dsrow dshead"', headAt + 1), -1, "exactly one");
  assert.ok(headAt < listed.indexOf('class="dsrow"'), "at the top of the list, before any sheet");
  const headHtml = listed.slice(headAt, listed.indexOf("</div>", listed.indexOf("dsacts", headAt)));
  /* the definition's OWN short labels, in the definition's own order */
  const shorts = ST.daySheetOf(ST.GLASS, "cut").counts.map(c => ST.dayCountShort(c));
  assert.deepStrictEqual(shorts, ["Clear", "K-glass", "Satin", "Obscure"],
    "the definition carries a short label for every count");
  let at = -1;
  shorts.forEach(label => {
    const found = headHtml.indexOf(">" + label + "<", at);
    assert.ok(found > at, 'the heading row is missing "' + label + '", or it is out of order');
    at = found;
  });
  assert.ok(headHtml.indexOf(">Day<") >= 0 && headHtml.indexOf(">Who<") >= 0 &&
            headHtml.indexOf(">Total<") >= 0 && headHtml.indexOf(">Note<") >= 0,
    "with Day, Who, Total and Note around them");
  /* and it really is the definition's list, not a copy in app.js. A quoted
     label is a copy; the word in a comment explaining the fix is not. */
  assert.ok(!/["']K-glass["']/.test(src("app.js")),
    "no label is written down in app.js: they can only have come from the definition");
  assert.ok(!/["']Other obscure/.test(src("app.js")));
  assert.strictEqual(ST.dayCountShort(["X", "The long question"]), "The long question",
    "a definition that gives no short label falls back to the long one");
  pass("the day sheets table has one heading row, from the definition, in the cells' own order");

  /* ===== the clock is LOCAL, on both screens =====
     `savedAt` is an ISO stamp; slicing it out gives the UTC clock, so a sheet
     saved at 12:28 Irish time read 11:28 in the office all summer. This is
     zone-independent: it asks what the helper answers (proved local against
     Date's own getHours, which is local by definition) and that neither screen
     slices the stamp any more. */
  const clockAt = "2026-09-21T11:28:00.000Z";
  const asDate = new Date(clockAt), pad = n => (n < 10 ? "0" : "") + n;
  assert.strictEqual(ST.stClock(clockAt), pad(asDate.getHours()) + ":" + pad(asDate.getMinutes()),
    "ST.stClock is the local clock, wherever this runs");
  assert.strictEqual(ST.stClock("not a time"), "", "and says nothing about what is not a time");
  assert.strictEqual(ST.stClock(""), "");
  assert.ok(A("stWhen(" + JSON.stringify(clockAt) + ")").indexOf(ST.stClock(clockAt)) >= 0,
    "the office's own stWhen is built on it");
  const savedRow = A("dayRowsNow()[0]");
  assert.ok(html("#dsbody").indexOf(ST.stClock(savedRow.savedAt)) >= 0,
    "and a row's “saved at” is that clock");
  [src("app.js"), src("station.js")].forEach(f =>
    assert.ok(!/savedAt\s*\)?\s*\.slice\(11/.test(f) && !/savedAt\).slice\(11, 16\)/.test(f),
      "neither screen slices the ISO stamp for a clock any more"));
  assert.ok(src("station.js").indexOf("ST.stClock(saved.savedAt)") >= 0,
    "the tablet's “saved 14:02” is the same helper");
  /* the window is left OPEN: the edit-mode test below carries on with it */
  pass("“saved at” is the local clock on both screens, from one helper, never a slice of the stamp");

  /* ===== D3: a half-typed correction is not wiped by the poll =====
     The floor's poll re-reads this list every twenty seconds while the window
     is open, and each read repainted the whole table body - so a correction
     being typed was thrown away, mid-sentence, about three times a minute. */
  A("DAYF.weekday = ''; DAYF.from = '2026-09-01'; DAYF.to = '2026-09-30'; paintDaySheets();");
  const editId = A("dayRowsNow()[0].id");
  A("DAYEDIT = " + JSON.stringify(editId) + "; paintDaySheets(true);");
  const editing = html("#dsbody");
  assert.ok(editing.indexOf("data-dsc=") >= 0, "the row is open with boxes in it");
  assert.ok(editing.indexOf("data-dssave") >= 0);
  /* somebody is typing. A poll lands: the read happens, the paint does not */
  assert.strictEqual(A("DAY_PAINT_OWED"), false);
  A("paintDaySheets();");
  assert.strictEqual(html("#dsbody"), editing, "the body is exactly as it was - not rebuilt");
  assert.strictEqual(A("DAY_PAINT_OWED"), true, "and the repaint is owed, not lost");
  /* three more polls, and a new row arriving underneath: still untouched */
  DAYITEMS = DAYITEMS.concat([daySheetRow("2026-09-23", "Person B", [7, 0, 0, 0], {}, "77")]);
  assert.ok(await A("readDaySheets()"), "the read itself still happens on every poll");
  A("paintDaySheets(); paintDaySheets(); paintDaySheets();");
  assert.strictEqual(html("#dsbody"), editing, "still exactly as it was");
  /* cancel, and everything that arrived while it was open is there */
  A("DAYEDIT = ''; paintDaySheets();");
  assert.strictEqual(A("DAY_PAINT_OWED"), false, "the owed paint is taken");
  assert.ok(html("#dsbody").indexOf("Person B") >= 0,
    "and the row that arrived during the edit is on screen");
  assert.ok(html("#dsbody").indexOf("data-dsc=") < 0, "with no row in edit mode any more");
  A("$('#dayhost').remove();");
  pass("D3: no repaint of the table while a correction is being typed; the paint is owed, not lost");

  reset();
  A("state.board = 'glass'; STATION_OK = true;");
  A("openStationReport('glass');");
  const xw = REG["xhost"] ? REG["xhost"].innerHTML : "";
  assert.ok(xw.indexOf("Station report") >= 0, "the export window opens on the new template");
  assert.ok(xw.indexOf('value="glass|cut"') >= 0 && xw.indexOf('value="welding|weld"') >= 0,
    "offering every station-stage the definitions name");
  assert.ok(xw.indexOf('value="glass|tuff"') < 0, "and not a counter that is not a stage to report on");
  assert.strictEqual(A("XSTATE.station"), "glass|cut", "with the board's own station picked");
  assert.ok(xw.indexOf("This week") >= 0 && xw.indexOf("Custom dates") >= 0, "and the four periods");
  assert.ok(xw.indexOf('id="xdl"') >= 0, "with a Download button");
  A("XSTATE.period = 'custom'; XSTATE.pfrom = '2026-09-21'; XSTATE.pto = '2026-09-27'; renderExportWindow();");
  assert.ok(REG["xhost"].innerHTML.indexOf('id="xpfrom"') >= 0, "custom dates get their own two boxes");
  assert.ok(typeof A("xpUpdateCount()") === "number", "and the count answers without throwing");
  A("$('#xhost').remove(); state.board = null;");
  pass("the export window's Station report: every station-stage, the periods, and the count");

  /* a missing list is a quiet, explained state on BOTH screens, and the board
     is exactly as it was */
  reset();
  DAY_EXISTS = false; TARGET_EXISTS = false;
  CW._resetListIds(); delete mem.cw_listids;
  A("DAY_ITEMS = null; DAY_OK = null; DAY_SOFT = 0; DAY_SOFT_MS = 0; TARGET_OK = null;");
  assert.strictEqual(await A("readDaySheets()"), null);
  assert.strictEqual(A("DAY_OK"), false);
  assert.ok(String(A("DAY_WHY")).indexOf("Station day sheets") >= 0, "it says which list, plainly");
  assert.ok(String(A("DAY_WHY")).indexOf("nothing in the Excel file is involved") >= 0,
    "and that the Excel file has nothing to do with it");
  A("STATION_OK = true; STATION_ITEMS = " + JSON.stringify(ITEMS) + ";");
  const board = A("stationBoardHtml()");
  assert.ok(board.indexOf("R0001") >= 0, "the board is untouched by it");
  assert.ok(board.indexOf("stweek") < 0, "and simply has no week line on it");
  assert.strictEqual(TOASTS.length, 0, "nothing is shouted at anybody");
  pass("a missing list on the office side is a quiet state; the board is exactly as it was");

  reset();
  const cut2 = newStation("cut");
  const S2 = code => vm.runInContext(code, cut2);
  S2("SITEID = " + JSON.stringify(SITE) + "; QUEUE = {}; LOGQ = {}; DAYQ = {};");
  assert.strictEqual(await S2("readDay()"), false);
  assert.strictEqual(S2("DAY_OK"), false);
  assert.ok(String(S2("DAY_WHY")).indexOf("Station day sheets") >= 0);
  assert.ok(String(S2("daySheetHtml()")).indexOf("Station day sheets") < 0 ||
            S2("!!PERSON") === false, "the tablet says so inside the sheet, not over the board");
  DAY_EXISTS = true; TARGET_EXISTS = true;
  CW._resetListIds(); delete mem.cw_listids;
  pass("and on the tablet: the sheet says which list is missing, the board is never taken away");

  /* ================= 8. the station report ================= */
  const P = { from: "2026-09-21", to: "2026-09-27", label: "This week" };
  const GDATA = {
    board: ST.jobBoard([
      item({ Title: "R0001", Job: "R0001", Customer: "Customer One", Total: 6, TuffTotal: 0,
             Seq: 1, Active: "Yes", OfficeDone: "No", Cut: 6, Hotmelt: 2, Tuff: 0,
             CutBy: "Person A", CutAt: "2026-09-21T10:00:00.000Z",
             DoneBy: "Person A", DoneAt: "2026-09-21T10:00:00.000Z" }, "1"),
      /* D8: the office types a site contact into a customer name often enough
         that welding's feeder has stripped it since the day it shipped. The
         glass report exported it as typed - a phone number in an exported
         file, which is the one thing rule 3 forbids outright. */
      item({ Title: "R0002", Job: "R0002", Customer: "Customer Two 086 123 4567", Total: 4,
             TuffTotal: 0, Seq: 2, Active: "Yes", OfficeDone: "No",
             Cut: 1, Hotmelt: 0, Tuff: 0 }, "2")]),
    log: ST.logRows([
      item({ Title: "R0001", Station: "Glass", Stage: "cut", From: 0, To: 6,
             Who: "Person A", At: "2026-09-21T10:00:00.000Z" }, "10"),
      item({ Title: "R0002", Station: "Glass", Stage: "cut", From: 0, To: 1,
             Who: "Person A", At: "2026-09-22T11:00:00.000Z" }, "11"),
      item({ Title: "R0002", Station: "Glass", Stage: "hotmelt", From: 0, To: 3,
             Who: "Person B", At: "2026-09-22T12:00:00.000Z" }, "12"),
      item({ Title: "R0003", Station: "Glass", Stage: "cut", From: 0, To: 9,
             Who: "Person A", At: "2026-01-02T12:00:00.000Z" }, "13")]),
    notes: ST.commentRows([
      item({ Title: "R0001|1", Job: "R0001", Station: "Glass", Who: "Person A",
             Text: "two units short, ring the supplier on 086 123 4567",
             At: "2026-09-21T13:00:00.000Z" }, "20")]),
    days: ST.dayRows([
      daySheetRow("2026-09-21", "Person A", [12, 4, 0, 2], { WeekTarget: 250 }),
      daySheetRow("2026-09-22", "Person A", [10, 0, 0, 0],
                  { WeekTarget: 250, Note: "air off 9:20-11:21, call D02 X285 or 087-123-4567",
                    EditedBy: "the office", EditedAt: "2026-09-23T09:00:00.000Z" })], COUNTS, {}),
    target: 250, who: "the office", when: new Date("2026-09-28T09:00:00.000Z"), build: "20260921-0900"
  };
  const rep = stationReport(ST.GLASS, "cut", GDATA, P);
  assert.deepStrictEqual(rep.map(s => s.name), ["Summary", "Days", "Jobs", "Activity", "Notes"],
    "all five sheets for a stage with day sheets");
  const S1 = rep[0];
  assert.ok(S1.head.some(h => h[0] === "Station" && h[1] === "Glass"));
  assert.ok(S1.head.some(h => h[0] === "Stage" && h[1] === "Cutting"));
  assert.ok(S1.head.some(h => h[0] === "Contact details" && String(h[1]).indexOf("None") === 0),
    "and it says in words that there are no contact details in the file");
  assert.strictEqual(S1.rows.length, 1, "one week in the period");
  assert.strictEqual(S1.rows[0][0], "2026-W39");
  assert.strictEqual(S1.rows[0][2], 7, "units recorded at this stage: 6 and 1, not the hotmelting 3");
  assert.strictEqual(S1.rows[0][3], 2, "two jobs touched");
  assert.strictEqual(S1.rows[0][4], 1, "one of them complete at this stage");
  assert.deepStrictEqual(S1.columns.slice(5), ["Clear glass sheets cut", "K-glass sheets cut",
    "Satin sheets cut", "Other obscure sheets cut", "Sheets total", "Target", "Difference"]);
  assert.strictEqual(S1.rows[0][S1.columns.indexOf("Sheets total")], 28);
  assert.strictEqual(S1.rows[0][S1.columns.indexOf("Target")], 250);
  assert.strictEqual(S1.rows[0][S1.columns.indexOf("Difference")], -222);
  pass("the Summary counts the week: units, jobs, jobs complete, the counts, the target and the gap");

  const days = rep[1];
  assert.deepStrictEqual(days.rows.map(r => r[0]), ["2026-09-21", "2026-09-22"],
    "a row per day in the period, oldest first, and January is not in it");
  assert.strictEqual(days.rows[0][1], "Monday");
  assert.strictEqual(days.rows[0][2], 6, "units recorded that day at this stage");
  assert.ok(String(days.rows[0][3]).indexOf("Person A 6") >= 0, "and who recorded them");
  assert.ok(String(days.rows[1][days.columns.indexOf("Corrected by the office")]).indexOf("the office") >= 0,
    "a sheet the office corrected says so");
  const jobs = rep[2];
  assert.deepStrictEqual(jobs.columns.slice(0, 3), ["Job", "Customer", "Status"]);
  assert.strictEqual(jobs.rows.length, 2, "one row per job for glass");
  assert.deepStrictEqual(jobs.rows[0].slice(0, 2), ["R0001", "Customer One"]);
  assert.strictEqual(jobs.rows[0][jobs.columns.indexOf("Done at this stage")], 6);
  assert.strictEqual(jobs.rows[0][jobs.columns.indexOf("Left")], 0);
  assert.strictEqual(jobs.rows[0][jobs.columns.indexOf("Complete")], "Yes");
  assert.strictEqual(jobs.rows[1][jobs.columns.indexOf("Complete")], "No");
  const act = rep[3];
  assert.strictEqual(act.rows.length, 2, "this stage's lines in the period, and no other's");
  assert.ok(act.rows[0][0] < act.rows[1][0], "oldest first");
  assert.ok(act.head.some(h => String(h[1]).indexOf("Dashboard Log") >= 0),
    "with the note that the office's own edits live in Dashboard Log");
  /* D12: glass lines carry no product group, so there is no column of the
     station's own name repeated down the sheet - and its one report stage is
     made of one log stage, so there is no Part column either */
  assert.deepStrictEqual(act.columns, ["When", "Who", "Job", "From", "To", "Units"]);
  pass("Days, Jobs and Activity: the period only, this stage only, in the office's own words");

  /* ===== D8: rule 3 over the customer name ===== */
  assert.strictEqual(jobs.rows[1][1], "Customer Two …",
    "a phone number typed into a customer name is stripped out of the Jobs sheet");
  assert.ok(!JSON.stringify(rep).match(/086 ?123 ?4567/),
    "and is nowhere in the report at all");
  pass("D8: the Jobs sheet strips the customer name, as welding's feeder already did");

  /* RULE 3: the two free-text columns, three written shapes of a phone number
     and an eircode - all of them stripped on the way into the file */
  const noteCell = days.rows[1][days.columns.indexOf("Note")];
  assert.ok(noteCell.indexOf("087-123-4567") < 0, "a dashed phone number is stripped from a day note");
  assert.ok(noteCell.indexOf("D02 X285") < 0, "and an eircode with it");
  assert.ok(noteCell.indexOf("air off") >= 0, "while what the person actually said survives");
  const notesSheet = rep[4];
  const text = notesSheet.rows[0][3];
  assert.ok(text.indexOf("086 123 4567") < 0, "a spaced phone number is stripped from a floor note");
  assert.ok(text.indexOf("two units short") >= 0);
  ["086 123 4567", "087-123-4567", "0861234567", "D02 X285", "+353 86 123 4567"].forEach(shape => {
    const one = stationReport(ST.GLASS, "cut", Object.assign({}, GDATA, {
      notes: ST.commentRows([item({ Title: "R1|1", Job: "R1", Station: "Glass", Who: "Person A",
                                    Text: "ring " + shape, At: "2026-09-21T13:00:00.000Z" }, "30")]) }), P);
    const sheetRows = (one.find(s => s.name === "Notes") || { rows: [] }).rows;
    assert.ok(String(sheetRows[0][3]).indexOf(shape) < 0, shape + " reached the report unstripped");
  });
  pass("rule 3: a phone number in any of four written shapes and an eircode are stripped from the report");

  /* ===== D9: ONE DEFINITION OF "DAY" FOR THE WHOLE REPORT =====
     A day sheet is filed under the tablet's LOCAL date; log lines and notes
     were bucketed by slicing their ISO stamp, which is the UTC date. In Irish
     summer time (UTC+1) everything recorded between midnight and one in the
     morning therefore landed on the day before - in a different row of Days
     from the sheet it was cut for, and, on a Monday, in a different ISO week
     of the Summary.

     The zone is INJECTED rather than taken from the machine: this suite has to
     give the same answer on a laptop in Ireland and on a build server in UTC,
     and TZ cannot be changed from inside a running node. `dayOf` is the same
     hook a caller reporting for a fixed zone would use. */
  const IRISH_SUMMER = at => {               // UTC+1, the zone the workshop is in from March to October
    const t = Date.parse(String(at));
    if (!isFinite(t)) return "";
    return new Date(t + 3600000).toISOString().slice(0, 10);
  };
  const MIDNIGHT = "2026-09-20T23:30:00.000Z";       // 00:30 on Monday 21st, Irish summer time
  assert.strictEqual(IRISH_SUMMER(MIDNIGHT), "2026-09-21", "the fixture's own arithmetic, stated");
  assert.strictEqual(xpIsoDate(MIDNIGHT.slice(0, 10)), "2026-09-20",
    "... and the UTC slice the first build used says the day before");
  const ZONED = Object.assign({}, GDATA, {
    dayOf: IRISH_SUMMER,
    log: ST.logRows([item({ Title: "R0001", Station: "Glass", Stage: "cut", From: 0, To: 6,
                            Who: "Person A", At: MIDNIGHT }, "40")]),
    notes: ST.commentRows([item({ Title: "R0001|1", Job: "R0001", Station: "Glass", Who: "Person A",
                                  Text: "started at midnight", At: MIDNIGHT }, "41")]),
    days: ST.dayRows([daySheetRow("2026-09-21", "Person A", [12, 0, 0, 0], { WeekTarget: 250 })],
                     COUNTS, {})
  });
  const zrep = stationReport(ST.GLASS, "cut", ZONED, P);
  const zdays = zrep.find(s => s.name === "Days");
  assert.strictEqual(zdays.rows.length, 1,
    "the half-past-midnight line and that day's sheet are ONE row of Days, not two");
  assert.strictEqual(zdays.rows[0][0], "2026-09-21", "and it is the day the person worked");
  assert.strictEqual(zdays.rows[0][2], 6, "with the units on it");
  assert.strictEqual(zdays.rows[0][zdays.columns.indexOf("Total")], 12, "and the sheet's counts");
  assert.strictEqual(zrep[0].rows.length, 1, "one week in the Summary, not two");
  assert.strictEqual(zrep[0].rows[0][0], "2026-W39",
    "the week the day belongs to locally - a Monday at 00:30 is not last week");
  assert.strictEqual(zrep[0].rows[0][2], 6);
  assert.ok(zrep.find(s => s.name === "Notes"), "and a note at the same moment is in the period too");
  /* the period's own edges use the same day: a line at 00:30 on the Monday a
     period starts is IN it, where the UTC slice would have shut it out */
  const edge = stationReport(ST.GLASS, "cut", ZONED, { from: "2026-09-21", to: "2026-09-21" });
  assert.ok(edge.find(s => s.name === "Activity"), "a line at 00:30 on the first day is in the period");
  /* with no zone injected the report uses the machine's own local date, which
     is what ST.dayKey answers - the two must be the same function */
  const plain = stationReport(ST.GLASS, "cut",
    Object.assign({}, ZONED, { dayOf: null }), { from: "2020-01-01", to: "2030-01-01" });
  assert.strictEqual(plain.find(s => s.name === "Days").rows[0][0], ST.dayKey(new Date(MIDNIGHT)),
    "the default is ST.dayKey of the stamp: the local day, wherever this runs");
  pass("D9: log lines, notes and day sheets are all bucketed by the same local day");

  /* a stage with no day sheet: the same report, without the day-sheet columns */
  const hotRep = stationReport(ST.GLASS, "hotmelt", GDATA, P);
  const hs = hotRep[0];
  assert.deepStrictEqual(hs.columns, ["Week", "Week starting", "Units recorded", "Jobs touched",
                                      "Jobs complete at this stage"],
    "no counts, no total, no target: this stage has no day sheet");
  assert.ok(hs.head.some(h => h[0] === "Stage" && h[1] === "Hotmelting"));
  assert.strictEqual(hs.rows[0][2], 3, "and it counts its own stage's units");
  assert.ok(hotRep.find(s => s.name === "Days").columns.indexOf("Note") < 0,
    "the Days sheet loses its day-sheet columns too");
  pass("glass hotmelting gets the same report with the day-sheet columns left out");

  /* WELDING goes through the same function with no branch in it: its own
     definition answers for its board, one row per job AND product group.

     THE LOG FIXTURE IS BUILT THE WAY THE TABLET BUILDS IT (D1) - through
     ST.logFields(WELDC.weldLogEntry(...)) - and never by hand. Hand-writing
     `Stage: "weld"` here was the whole reason the first build's suite was
     green while the welding report came out with no Activity sheet and an
     empty Days: the tablet logs one line per PART. A fixture that agrees with
     the code instead of with the tablet proves nothing at all. */
  const weldLine = (job, group, part, from, to, who, at, id) =>
    item(ST.logFields(WELDC.weldLogEntry({ job: job, group: group, part: part,
                                           from: from, to: to, who: who, at: at })), id);
  assert.strictEqual(weldLine("R0001", "CASEMENT WINDOWS", "frames", 0, 10,
                              "Person C", "2026-09-21T10:00:00.000Z", "60").fields.Stage, "frames",
    "the tablet logs the PART as the stage - this is what the report has to match");
  const WDATA = {
    board: WELDC.weldOfficeBoard([
      item({ Title: "R0001|CASEMENT WINDOWS", Job: "R0001", Group: "CASEMENT WINDOWS", GroupSeq: 0,
             Customer: "Customer One", Section: "In production", Seq: 1, Active: "Yes",
             Frames: 10, Sashes: 8, FramesDone: 10, SashesDone: 8,
             DoneBy: "Person C", DoneAt: "2026-09-21T10:00:00.000Z" }, "50"),
      item({ Title: "R0001|SUPER DOOR", Job: "R0001", Group: "SUPER DOOR", GroupSeq: 1,
             Customer: "Customer One", Section: "In production", Seq: 1, Active: "Yes",
             Frames: 0, Sashes: 2, FramesDone: 0, SashesDone: 1 }, "51")]),
    log: ST.logRows([
      weldLine("R0001", "CASEMENT WINDOWS", "frames", 0, 10, "Person C", "2026-09-21T10:00:00.000Z", "60"),
      weldLine("R0001", "CASEMENT WINDOWS", "sashes", 0, 8, "Person C", "2026-09-22T11:00:00.000Z", "61"),
      weldLine("R0001", "SUPER DOOR", "sashes", 0, 1, "Person D", "2026-09-22T12:00:00.000Z", "62")
    ], "Welding"),
    notes: [], days: [], target: null, who: "the office", when: new Date("2026-09-28T09:00:00.000Z")
  };
  const wrep = stationReport(WELDC.WELD, "weld", WDATA, P);
  assert.deepStrictEqual(wrep.map(s => s.name), ["Summary", "Days", "Jobs", "Activity"],
    "Days and Activity are there - the welding report was empty without D1");
  const wj = wrep.find(s => s.name === "Jobs");
  assert.strictEqual(wj.rows.length, 2, "one row per job AND product group");
  assert.deepStrictEqual(wj.rows.map(r => r[3]), ["CASEMENT WINDOWS", "SUPER DOOR"]);
  assert.strictEqual(wj.columns.indexOf("Frames done") >= 0 && wj.columns.indexOf("Sashes") >= 0, true,
    "frames and sashes, as the board shows them");
  assert.strictEqual(wj.rows[0][wj.columns.indexOf("Frames done")], 10);
  assert.strictEqual(wj.rows[1][wj.columns.indexOf("Sashes done")], 1);
  const ws = wrep[0];
  assert.deepStrictEqual(ws.columns, ["Week", "Week starting", "Units recorded", "Jobs touched",
                                      "Jobs complete at this stage"], "and no day-sheet columns at all");
  assert.ok(ws.head.some(h => h[0] === "Station" && h[1] === "Welding"));
  /* the units really are counted, and both parts are in them: 10 frames + 8
     sashes + 1 sash = 19 over the week, on two days */
  assert.strictEqual(ws.rows.length, 1, "one week");
  assert.strictEqual(ws.rows[0][2], 19, "every part's units, not one part's and not none");
  const wd = wrep.find(s => s.name === "Days");
  assert.deepStrictEqual(wd.rows.map(r => r[0]), ["2026-09-21", "2026-09-22"]);
  assert.strictEqual(wd.rows[0][2], 10, "Monday: the frames");
  assert.strictEqual(wd.rows[1][2], 9, "Tuesday: 8 sashes and 1 more");
  assert.ok(String(wd.rows[1][3]).indexOf("Person C 8") >= 0 &&
            String(wd.rows[1][3]).indexOf("Person D 1") >= 0, "and who did which");
  const wa = wrep.find(s => s.name === "Activity");
  assert.strictEqual(wa.rows.length, 3, "a line per log line in the period");
  /* D12: this station's lines carry a product group, so the column is there -
     and because its one report stage is made of two log stages, the part is
     named too */
  assert.deepStrictEqual(wa.columns, ["When", "Who", "Job", "Group", "Part", "From", "To", "Units"]);
  assert.strictEqual(wa.rows[0][3], "CASEMENT WINDOWS");
  assert.deepStrictEqual(wa.rows.map(r => r[4]), ["frames", "sashes", "sashes"]);
  pass("welding: its own definition's log stages, its per-group Jobs sheet, its units by day and person");

  /* nothing at all in the period: a Summary that says so, and no throw */
  const empty = stationReport(ST.GLASS, "cut",
    { board: [], log: [], notes: [], days: [], target: null, when: new Date() },
    { from: "2020-01-01", to: "2020-01-07" });
  assert.strictEqual(empty.length, 1, "one sheet: the Summary");
  assert.strictEqual(empty[0].rows.length, 0);
  assert.ok(empty[0].head.some(h => h[0] === "Nothing recorded"),
    "and it says in words that the station recorded nothing");
  pass("an empty period produces a Summary that says so, and nothing crashes");

  /* the named periods, measured from a fixed day so this needs no clock */
  const NOW = new Date(2026, 8, 23, 12, 0, 0);            // Wednesday 23 Sep 2026
  assert.deepStrictEqual([stationPeriod("week", NOW).from, stationPeriod("week", NOW).to],
    ["2026-09-21", "2026-09-23"], "this week runs from Monday to today");
  assert.deepStrictEqual([stationPeriod("last", NOW).from, stationPeriod("last", NOW).to],
    ["2026-09-14", "2026-09-20"], "last week is Monday to Sunday");
  assert.strictEqual(stationPeriod("month", NOW).from, "2026-09-01");
  assert.deepStrictEqual([stationPeriod("custom", NOW, "2026-01-01", "2026-01-31").from,
                          stationPeriod("custom", NOW, "2026-01-01", "2026-01-31").to],
    ["2026-01-01", "2026-01-31"]);
  pass("this week, last week, this month and custom dates all mean what they say");

  /* the file: a real workbook, read back, with the same sheets in it */
  const wb = buildStationWorkbook(rep, { when: new Date("2026-09-28T09:00:00.000Z") });
  const buf = await wb.xlsx.writeBuffer();
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(buf);
  assert.deepStrictEqual(back.worksheets.map(s => s.name), ["Summary", "Days", "Jobs", "Activity", "Notes"]);
  const cells = [];
  back.eachSheet(s => s.eachRow(r => r.eachCell(c => cells.push(String(c.value == null ? "" : c.value)))));
  assert.ok(cells.some(v => v === "R0001"), "the jobs really are in the file");
  ["086 123 4567", "087-123-4567", "0861234567", "D02 X285"].forEach(shape =>
    assert.ok(!cells.some(v => v.indexOf(shape) >= 0), "the workbook leaked " + shape));
  /* the naming half, asked of the COLUMNS: no heading of this file says either
     word. What a person typed into a note is their text and is not censored -
     it is the values and the columns that are kept out. */
  rep.forEach(s => (s.columns || []).forEach(c => {
    const low = String(c).toLowerCase();
    assert.ok(low.indexOf("phone") < 0, 'the report has a column called "' + c + '"');
    assert.ok(low.indexOf("eircode") < 0, 'the report has a column called "' + c + '"');
  }));
  pass("the workbook is built, read back, and carries no phone number, eircode or column named after one");

  assert.strictEqual(exportLogFrom("xlsx", 2, "station", "Glass · Cutting · 2026-09-21 to 2026-09-27"),
    "Excel · 2 jobs · Station report · Glass · Cutting · 2026-09-21 to 2026-09-27",
    "the log line names the template, the station, the stage and the period");
  assert.strictEqual(stationReportFilename(ST.GLASS, "cut", P, "xlsx"),
    "Station report - Glass Cutting - 2026-09-21 to 2026-09-27.xlsx");
  pass("the file's name, and the Dashboard Log line that says what was exported");

  /* ================= 9. the gates ================= */
  const gate = /setFill|clearFill|setValues|appendLog|saveProgress|moveJobRow|batchWrite|\/workbook/;
  ["station-core.js", "station-ui.js", "station.js", "glass.html"].forEach(f =>
    assert.ok(!gate.test(src(f)), f + " must not be able to touch the workbook"));
  assert.strictEqual(ALLREQ.filter(r => /\/workbook|\/drive/.test(r.path)).length, 0,
    "and over the whole run, not one request went near the file");
  pass("the workbook gate: not a word of it in the tablet's files, not a request in the run");

  console.log("\n" + n + " checks passed");
})().catch(e => { console.error("FAIL", e); process.exit(1); });
