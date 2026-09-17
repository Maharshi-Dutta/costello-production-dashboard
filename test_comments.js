/* Offline test of the station comment channel - a note typed on a floor tablet
   against a job, read by the office in that job's drawer.
   Spec: docs/specs/2026-09-15-station-comments.md.

   The five things the brief asks to be proved, and why each of them matters:

     · the composer builds the right row - six columns, the job upper-cased, the
       station's own name, the person signed in, the note and an ISO stamp;
     · the TABLET shows a station its OWN notes only: a glass tablet must not
       show a cutting note, or the floor reads somebody else's problem as its
       own. The OFFICE DRAWER shows every station's, because that is the one
       place a job's notes come together;
     · Title is a row id, never a key: two notes may carry the same one and both
       survive. Nothing in this feature upserts, patches or deletes a row -
       asserted over every request the run makes, not by comment;
     · a missing `Station comments` list is a quiet, explained state on BOTH
       pages - the tablet keeps its board and the drawer keeps its job;
     · the tablet touches no workbook: over the whole run, not one request goes
       near /workbook or /drive.

   Graph is a fake fetch() over one floor site; nothing leaves the box, every
   person in here is made up and there is no address of any kind.
   Run: node test_comments.js                                                */
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
   One floor site with four lists. There is NO workbook route at all: anything
   that reaches for the file gets a 599 that no retry rule matches, so a stray
   workbook call is a failure by construction rather than by assertion. */
const G = "https://graph.microsoft.com/v1.0";
const SITE = "costellowindowsie.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";
const FSITE = "costellowindowsie.sharepoint.com,aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee,ffffffff-0000-1111-2222-333333333333";
const GLASS_ID = "list-glass-station";
const PEOPLE_ID = "list-station-people";
const LOG_ID = "list-station-log";
const NOTES_ID = "list-station-comments";
const FLISTS_PATH = "/sites/" + FSITE + "/lists";
const HOST_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/FloorStations";
/* FIXTURE, 2026-09-16: the WORKBOOK'S OWN SITE serves the same four lists.
   Since ST.GLASS.site became "own" (the hard pin - graph.js, "a station that
   is PINNED to one site"), the glass page and the office resolve the
   workbook's own site by path and never look at `Floor stations`. The stores
   are the same objects, because this suite is about one list read from both
   sides - which site it is read from is test_station.js' question, not this
   one's. */
const OWNLISTS_PATH = "/sites/" + SITE + "/lists";
const OWN_LOOKUP = "/sites/costellowindowsie.sharepoint.com:/sites/ProductionProgress";

let NOTES_EXISTS = true;                // has the office made the Station comments list yet?
let ITEMS = [], PEOPLEITEMS = [], LOGITEMS = [], NOTEITEMS = [];
let NEXTID = 500;
let FAIL_NOTES = 0;                     // n writes to the comments list to refuse
let FAIL_NOTES_GET = 0;                 // n reads of it to refuse - the workshop wifi, in effect
const lists = () => [{ id: GLASS_ID, displayName: "Glass station" },
                     { id: PEOPLE_ID, displayName: "Station people" },
                     { id: LOG_ID, displayName: "Station log" }]
  .concat(NOTES_EXISTS ? [{ id: NOTES_ID, displayName: "Station comments" }] : []);

const REQ = [], ALLREQ = [];
const ok = body => ({ status: 200, body: body });
const item = (fields, id) => ({ id: String(id == null ? NEXTID++ : id), fields: Object.assign({}, fields) });
const storeFor = id => id === GLASS_ID ? ITEMS : id === PEOPLE_ID ? PEOPLEITEMS
                     : id === LOG_ID ? LOGITEMS : id === NOTES_ID ? NOTEITEMS : null;

let DELTA_SEQ = 0;
const DELTA_NEXT = {};                  // listId -> [batch, ...] for token calls

function routeDelta(listId, tail, base) {
  const store = storeFor(listId) || [];
  /* the deltaLink comes back against the site it was asked of, so a token can
     never carry a page from one site to the other */
  const link = () => G + (base || FLISTS_PATH) + "/" + listId + "/items/delta?token=T" + (++DELTA_SEQ);
  if (/[?&]token=/.test(tail)) {
    const batch = (DELTA_NEXT[listId] || []).shift() || [];
    return ok({ value: batch, "@odata.deltaLink": link() });
  }
  return ok({ value: store.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })),
              "@odata.deltaLink": link() });
}

function route(method, path, body) {
  if (path === HOST_LOOKUP) return ok({ id: FSITE, displayName: "Floor stations" });
  if (path.indexOf(OWN_LOOKUP) === 0) return ok({ id: SITE, displayName: "Production Progress" });
  if (path.indexOf(OWNLISTS_PATH) === 0) return routeLists(method, path, body, OWNLISTS_PATH);
  if (path.indexOf(FLISTS_PATH) === 0) return routeLists(method, path, body, FLISTS_PATH);
  return { status: 599, body: { error: { code: "thisTestServesNoWorkbook", message: path } } };
}
function routeLists(method, path, body, base) {
  {
    const rest = path.slice(base.length);
    if (method === "GET" && rest.indexOf("?$select=id,displayName") === 0) return ok({ value: lists() });
    const mi = /^\/([^/?]+)\/items(.*)$/.exec(rest);
    if (!mi) return { status: 404, body: { error: { code: "itemNotFound" } } };
    const id = mi[1], tail = mi[2];
    const store = storeFor(id);
    if (!store) return { status: 404, body: { error: { code: "itemNotFound" } } };
    if (FAIL_NOTES_GET && id === NOTES_ID && method === "GET") {
      FAIL_NOTES_GET--; return { status: 403, body: { error: { code: "accessDenied" } } };
    }
    if (method === "GET" && tail.indexOf("/delta") === 0) return routeDelta(id, tail, base);
    if (method === "GET") return ok({ value: store.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) });
    if (FAIL_NOTES && id === NOTES_ID && method !== "GET") {
      FAIL_NOTES--; return { status: 403, body: { error: { code: "accessDenied" } } };
    }
    if (method === "POST" && tail === "") {
      /* NO unique rule on Title, deliberately: this list is a log. Two rows may
         carry the same Title and the list takes both. */
      const made = item((body && body.fields) || {});
      store.push(made);
      return ok({ id: made.id, fields: made.fields });
    }
    return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + method + " " + path } } };
  }
  return { status: 599, body: { error: { code: "thisTestServesNoWorkbook", message: path } } };
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
run("app.js");

const TOASTS = [];
global.toast = (m, isErr) => TOASTS.push({ m: String(m), err: !!isErr });

/* ---------- the tablet, in a context of its own ---------- */
const stationDoc = global.document;
const stationFetch = async (url, init) => {
  if (String(url).indexOf("version.json") === 0)
    return { ok: true, status: 200, json: async () => ({ build: "20260915-0900" }) };
  return global.fetch(url, init);
};
function newStation() {
  const sb = {
    console: console, setTimeout: setTimeout, clearTimeout: clearTimeout,
    setInterval: () => 0, clearInterval: clearInterval,
    localStorage: global.localStorage, document: stationDoc,
    window: { location: { origin: "http://localhost" } },
    location: { reload() {} }, fetch: stationFetch,
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
const A = vm.runInThisContext.bind(vm);
const reset = () => { REQ.length = 0; TOASTS.length = 0; };
const notesReq = () => REQ.filter(r => r.path.indexOf("/" + NOTES_ID + "/") >= 0);
const consent = yes => { CW.hasListConsent = async () => !!yes; };
const boardHtml = () => S("Object.keys(NODES).map(k => NODES[k].innerHTML).join('')");
const cardHtml = job => S("NODES[" + JSON.stringify(job) + "] ? NODES[" + JSON.stringify(job) + "].innerHTML : ''");

const mkJob = o => Object.assign({
  id: "R0001", cust: "Customer One", area: "Cork", eir: "", off: "", colour: "", ph3: "",
  wnd: 0, drs: 0, glass: { dg: 4 }, prods: [], notes: [], sheets: ["Production"], src: {},
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0,
  cp: { win: "", drs: "", glass: {}, prod: {} }
}, o || {});

/* The columns the spec's "Data model" table names, written out here so every
   assertion about a row is against the BRIEF and not against the code's own
   copy of it. A change to ST.COMMENT_FIELDS has to be a deliberate change to
   this line too. */
const SPEC_COLUMNS = ["Title", "Job", "Station", "Who", "Text", "At"];

const note = (job, station, who, text, at, id) =>
  item({ Title: String(job).toUpperCase() + "|" + Date.parse(at), Job: String(job).toUpperCase(),
         Station: station, Who: who, Text: text, At: at }, id);

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ================= 1. the shape of the thing ================= */
  assert.strictEqual(ST.COMMENT_LIST, "Station comments",
    "named like Station people and Station log: any floor station, not only glass");
  /* SPEC_COLUMNS is written out here, from the spec's own Data model table, and
     everything else in this file is checked against IT rather than against
     ST.COMMENT_FIELDS - otherwise the assertion is the code agreeing with
     itself and would follow a mistake anywhere it went. */
  assert.deepStrictEqual(ST.COMMENT_FIELDS, SPEC_COLUMNS,
    "the columns the spec's Data model table names, in that order");
  assert.strictEqual(typeof ST.stationComments, "function");
  assert.strictEqual(typeof ST.commentRows, "function");
  assert.strictEqual(typeof ST.commentFields, "function");
  pass("one list, six columns, and the channel itself is a function of station-core.js");

  /* the columns of this list carry nothing personal about a customer and
     nothing at all about the workbook: no phone, no eircode, no county, no
     price, no cell, no colour */
  ["Phone", "PHONE NO.", "Eircode", "County", "Price", "Range", "Cell", "Colour", "Fill"]
    .forEach(k => assert.ok(SPEC_COLUMNS.indexOf(k) < 0, k + " is not a column of this list"));
  pass("no phone, eircode, county, price, cell or colour column exists on the notes list");

  /* ================= 2. the row the composer builds ================= */
  const f1 = ST.commentFields({ job: "r0001", station: "Glass", who: "Person A",
                                text: "  two units short on this one  ", at: "2026-09-15T08:30:00.000Z" });
  assert.deepStrictEqual(Object.keys(f1).sort(), SPEC_COLUMNS.slice().sort(),
    "exactly the six columns, no more");
  assert.strictEqual(f1.Job, "R0001", "the job number, upper-cased, as everywhere else");
  assert.strictEqual(f1.Station, "Glass");
  assert.strictEqual(f1.Who, "Person A");
  assert.strictEqual(f1.Text, "two units short on this one", "trimmed, and otherwise exactly what was typed");
  assert.strictEqual(f1.At, "2026-09-15T08:30:00.000Z");
  assert.strictEqual(f1.Title, "R0001|" + Date.parse("2026-09-15T08:30:00.000Z"),
    "Title is JOB|unix-ms: a row id, not a key");
  pass("the composer's row: job, station, person, the note, an ISO stamp, and a JOB|unix-ms id");

  /* a second station writes the same shape with one word different - which is
     the whole design: a new station is a new VALUE, not a schema change */
  const f2 = ST.commentFields({ job: "R0001", station: "Cutting", who: "Person B", text: "cut list is wrong" });
  assert.strictEqual(f2.Station, "Cutting");
  assert.ok(f2.At && /^\d{4}-\d{2}-\d{2}T/.test(f2.At), "a missing stamp is filled in at send");
  assert.strictEqual(ST.commentFields({ job: "R0001", text: "x" }).Station, ST.STATION_NAME,
    "and with no station named at all it is this page's own");
  pass("Station is free text: a second station is a new value, not a new column");

  /* the note is capped, in the row builder as well as in the box, so a long
     paste cannot arrive through a draft replayed from localStorage */
  const long = ST.commentFields({ job: "R1", text: "x".repeat(ST.COMMENT_MAX + 500) });
  assert.strictEqual(long.Text.length, ST.COMMENT_MAX);
  pass("a note is capped at " + ST.COMMENT_MAX + " characters by the row builder, not only by the box");

  /* ================= 3. reading the list back ================= */
  const MIXED = [
    note("R0001", "Glass", "Person A", "two units short", "2026-09-15T08:00:00Z", "1"),
    note("R0001", "Cutting", "Person B", "cut list is wrong", "2026-09-15T09:00:00Z", "2"),
    note("R0002", "Glass", "Person C", "another job entirely", "2026-09-15T09:30:00Z", "3"),
    note("R0001", "Glass", "Person C", "sorted, ignore the last one", "2026-09-15T10:00:00Z", "4")
  ];
  const mineGlass = ST.commentRows(MIXED, { job: "R0001", station: "Glass" });
  assert.deepStrictEqual(mineGlass.map(r => r.id), ["1", "4"],
    "the glass tablet sees its own two notes on this job and not the cutting one");
  assert.deepStrictEqual(mineGlass.map(r => r.text),
    ["two units short", "sorted, ignore the last one"], "oldest first, so it reads as a conversation");
  assert.deepStrictEqual(ST.commentRows(MIXED, { job: "R0001", station: "Cutting" }).map(r => r.id), ["2"],
    "and a cutting tablet sees exactly the other one");
  pass("the tablet's thread is its own station's notes on its own job, oldest first");

  const drawer = ST.commentRows(MIXED, { job: "R0001" });
  assert.deepStrictEqual(drawer.map(r => r.id), ["1", "2", "4"],
    "the drawer does NOT filter by station: every station's notes on the job");
  assert.deepStrictEqual(drawer.map(r => r.station), ["Glass", "Cutting", "Glass"],
    "each one tagged with the station that wrote it");
  assert.deepStrictEqual(drawer.map(r => r.who), ["Person A", "Person B", "Person C"]);
  assert.ok(drawer.every(r => r.job === "R0001"), "and nothing from another job");
  pass("the office drawer's thread crosses every station and stays one job's");

  /* two notes inside one second, which a tablet can certainly produce: the item
     id breaks the tie, so the order is the order they were written in */
  const sameSecond = ST.commentRows([
    note("R0003", "Glass", "Person A", "second", "2026-09-15T11:00:00Z", "8"),
    note("R0003", "Glass", "Person A", "first", "2026-09-15T11:00:00Z", "7")
  ], { job: "R0003" });
  assert.deepStrictEqual(sameSecond.map(r => r.text), ["first", "second"]);
  pass("two notes in one second come back in the order they were written");

  /* how long ago each line says it was - drawn on every thread line, on both
     screens, so it is worth pinning down at the boundaries */
  const T0 = Date.parse("2026-09-15T12:00:00Z");
  const ago = (iso, now) => ST.commentAgo(iso, now == null ? T0 : now);
  assert.strictEqual(ago("2026-09-15T12:00:00Z"), "just now");
  assert.strictEqual(ago("2026-09-15T11:59:16Z"), "just now", "44 s is still just now");
  assert.strictEqual(ago("2026-09-15T11:59:15Z"), "1 min ago", "45 s is a minute");
  assert.strictEqual(ago("2026-09-15T11:30:00Z"), "30 min ago");
  assert.strictEqual(ago("2026-09-15T10:31:00Z"), "89 min ago", "89 minutes is still minutes");
  assert.strictEqual(ago("2026-09-15T10:29:00Z"), "2 h ago", "and 91 of them are hours");
  assert.strictEqual(ago("2026-09-13T12:00:00Z"), "2 days ago");
  assert.strictEqual(ago("2026-09-10T12:00:00Z"), "5 days ago");
  assert.strictEqual(ago("2026-09-15T12:05:00Z"), "just now",
    "a stamp from the future is never a negative age");
  assert.strictEqual(ago(""), "", "and nothing at all says nothing, rather than NaN");
  assert.strictEqual(ago("not a date"), "");
  assert.strictEqual(ST.commentAgo(null), "");
  pass("a thread line's “how long ago” is right at every boundary, and silent on rubbish");

  /* rubbish in the list is skipped rather than drawn: an empty note is not a
     note, and a row with no job belongs to nothing */
  assert.strictEqual(ST.commentRows([note("R1", "Glass", "A", "   ", "2026-09-15T11:00:00Z", "9")]).length, 0);
  assert.strictEqual(ST.commentRows([item({ Title: "", Job: "", Text: "orphan", At: "" }, "10")]).length, 0);
  assert.strictEqual(ST.commentRows(null).length, 0);
  assert.strictEqual(ST.commentRows([null, undefined]).length, 0);
  pass("an empty note, a row with no job, and no list at all are all handled without a throw");

  /* the Job column is what is read, with the Title's own prefix as the fallback:
     a row typed into SharePoint by hand can easily have one and not the other */
  const byTitle = ST.commentRows([item({ Title: "R0007|1", Text: "typed in by hand", At: "2026-09-15T11:00:00Z" }, "11")],
    { job: "R0007" });
  assert.strictEqual(byTitle.length, 1);
  assert.strictEqual(byTitle[0].station, "", "and a row with no station is nobody's in particular");
  pass("a hand-made row with only a Title still reaches the right job's drawer");

  /* ================= 4. the channel, end to end, on a tablet ================= */
  consent(true);
  CW._resetListIds(); delete mem.cw_listids;
  CW._setStationSite(FSITE, false);
  ITEMS = [item({ Title: "R0001", Job: "R0001", Customer: "Customer One", GlassType: "GLASS",
                  Total: 4, TuffTotal: 0, Seq: 1, Active: "Yes", OfficeDone: "No",
                  Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0 }, "900")];
  PEOPLEITEMS = [item({ Title: "Person A", Station: "Glass", Stages: "cut,hotmelt,glazed",
                        PIN: "", Active: "Yes" })];
  NOTEITEMS = [note("R0001", "Glass", "Person B", "yesterday: two units short",
                    "2026-09-14T16:00:00Z", "40"),
               note("R0001", "Cutting", "Person C", "a note from another station",
                    "2026-09-14T17:00:00Z", "41")];

  S("QUEUE = {}; LOGQ = {}; SITEID = " + JSON.stringify(FSITE) + "; TOKEN = null; QUERY = '';");
  assert.strictEqual(await S("readPeople()"), true);
  await S("readList()");
  S("pickPerson(PEOPLE.find(p => p.name === 'Person A'))");
  assert.strictEqual(S("PERSON.name"), "Person A");
  assert.strictEqual(await S("NOTES.read()"), true, "the channel read the list");
  S("render()");

  /* the card carries the composer, closed, with the count of THIS station's
     notes on it - one, not the two rows in the list */
  let card = cardHtml("R0001");
  assert.ok(/data-cmt="R0001"/.test(card), "a note button on the job's own card");
  assert.ok(/Notes \(1\)/.test(card), "counting this station's notes, not the cutting one");
  assert.ok(!/data-cmbox/.test(card), "and closed, so a card nobody is writing on is the height it was");
  pass("the tablet's card carries a note button counting its own station's notes");

  S("NOTES.toggle('R0001'); render();");
  card = cardHtml("R0001");
  assert.ok(/data-cmbox="R0001"/.test(card), "opened: there is a box");
  assert.ok(/data-cmsend="R0001"/.test(card), "... and a send button");
  assert.ok(card.indexOf("yesterday: two units short") >= 0, "the shift before is on screen");
  assert.ok(card.indexOf("a note from another station") < 0,
    "and the cutting station's note is not: a glass tablet does not see it");
  pass("the composer opens into this station's own thread and a box");

  /* nothing on the tablet can edit or remove a note: this list is append-only
     and there is no affordance for either, not even a disabled one */
  assert.ok(!/data-cmdel|data-cmedit/.test(card), "no delete and no edit hook on a note");
  assert.ok(!/>Delete<|>Edit<|>Remove</.test(card), "and no button that says so");
  pass("append-only on the tablet: no edit, no delete, not even a stub");

  /* ---- sending one ---- */
  reset();
  S("NOTES.setDraft('R0001', '  the frame is 20mm out  ');");
  await S("sendNote('R0001')");
  const posts = notesReq().filter(r => r.method === "POST");
  assert.strictEqual(posts.length, 1, "one note is one POST");
  assert.deepStrictEqual(Object.keys(posts[0].body.fields).sort(), SPEC_COLUMNS.slice().sort());
  assert.strictEqual(posts[0].body.fields.Job, "R0001");
  assert.strictEqual(posts[0].body.fields.Station, "Glass", "the tablet's own station, never another's");
  assert.strictEqual(posts[0].body.fields.Who, "Person A", "the person signed in at the picker");
  assert.strictEqual(posts[0].body.fields.Text, "the frame is 20mm out", "trimmed");
  assert.strictEqual(NOTEITEMS.length, 3, "and the list has it");
  assert.strictEqual(S("NOTES.draftOf('R0001')"), "", "the box is emptied once it is away");
  pass("sending appends exactly one row, from this station, as the person signed in");

  /* every request this feature has made to that list is a GET or a POST. No
     PATCH, no DELETE, no read-before-write: it is append-only in the code, not
     only in the brief */
  assert.ok(notesReq().every(r => r.method === "GET" || r.method === "POST"),
    "nothing has patched or deleted a note");
  S("render()");
  assert.ok(cardHtml("R0001").indexOf("the frame is 20mm out") >= 0,
    "and the new note is on screen without waiting for a read");
  pass("append-only on the wire: only GET and POST ever reach the notes list");

  /* Title is a row id, not a key. Two notes made in the same millisecond carry
     the same Title and the list takes both - nothing upserts and nothing
     dedupes, so nothing may assume it is unique. */
  reset();
  const fixed = 1789000000000;
  const twin = ST.stationComments({ station: "Glass", now: () => fixed,
    listItems: (name, o) => CW.listItems(name, o),
    listAdd: (name, fields, o) => CW.listAdd(name, fields, o),
    opts: () => ({ siteId: FSITE, fields: ST.COMMENT_FIELDS }) });
  await twin.read();
  twin.setDraft("R0009", "one");
  await twin.send("R0009", "Person A");
  twin.setDraft("R0009", "two");
  await twin.send("R0009", "Person A");
  const twins = NOTEITEMS.filter(x => x.fields.Job === "R0009");
  assert.strictEqual(twins.length, 2, "both rows are in the list");
  assert.strictEqual(twins[0].fields.Title, twins[1].fields.Title,
    "carrying the same Title, because Title is a row id and not a key");
  assert.deepStrictEqual(twin.rows("R0009").map(r => r.text), ["one", "two"],
    "and both are read back, in order");
  assert.ok(REQ.filter(r => r.method !== "GET").every(r => r.method === "POST"),
    "two notes are two POSTs and nothing else - no upsert, no dedupe, no delete");
  pass("Title uniqueness is never assumed: two notes may share one and both survive");

  /* ---- a refused write keeps the typing ---- */
  reset();
  FAIL_NOTES = 1;
  S("NOTES.setDraft('R0001', 'this one will not go');");
  await S("sendNote('R0001')");
  assert.strictEqual(S("NOTES.draftOf('R0001')"), "this one will not go",
    "the floor typed something they meant somebody to read: it stays in the box");
  assert.ok(/not sent/.test(cardHtml("R0001")), "and the card says so");
  assert.strictEqual(S("PROBLEM"), "", "a refused note never takes the board away");
  assert.strictEqual(S("READY"), true);
  FAIL_NOTES = 0;
  await S("sendNote('R0001')");
  assert.strictEqual(S("NOTES.draftOf('R0001')"), "", "and tapping Send again gets it away");
  pass("a note that will not send keeps the typing, says so, and leaves the board alone");

  /* ---- TYPING THE NEXT NOTE WHILE THE LAST ONE IS STILL IN THE AIR ----
     Review finding 3. The Send button is disabled during the POST; the BOX is
     not. On workshop wifi that is a second or two in which somebody starts the
     next note - and emptying the box when the POST comes back would throw that
     away silently, which is the one thing this feature must never do. */
  reset();
  let release = null;
  const slow = ST.stationComments({ station: "Glass",
    listItems: (name, o) => CW.listItems(name, o),
    listAdd: async (name, fields, o) => {
      await new Promise(r => { release = r; });        // held open until the test lets go
      return await CW.listAdd(name, fields, o);
    },
    opts: () => ({ siteId: FSITE, fields: ST.COMMENT_FIELDS }) });
  await slow.read();
  slow.setDraft("R0005", "note A");
  const inFlight = slow.send("R0005", "Person A");
  await new Promise(r => setTimeout(r, 5));
  assert.strictEqual(slow.state.sending["R0005"], 1, "the note is in the air");
  slow.setDraft("R0005", "note B, typed while A was sending");   // the box is still live
  release();
  await inFlight;
  assert.strictEqual(slow.draftOf("R0005"), "note B, typed while A was sending",
    "what was typed since the tap is still there - it was NOT wiped by the send coming back");
  assert.deepStrictEqual(slow.rows("R0005").map(r => r.text), ["note A"],
    "and the note that actually went is the one that was in the box when Send was tapped");
  /* ... and the box SAYS why there is still writing in it. Without a word here
     it looks exactly like a note that failed to send, and the obvious thing to
     do about that - tap Send again - would post the first note twice. */
  slow.toggle("R0005");
  assert.ok(slow.html("R0005").indexOf(ST.COMMENT_KEPT) >= 0,
    "the composer says the first note went and this one is new");
  assert.ok(slow.html("R0005").indexOf(ST.COMMENT_UNSENT) < 0, "and does not say it failed");
  slow.setDraft("R0005", "note B, still typing");
  assert.ok(slow.html("R0005").indexOf(ST.COMMENT_KEPT) < 0,
    "the hint goes as soon as they carry on typing - it has been read by then");
  slow.toggle("R0005");
  /* the ordinary case is unchanged: nothing typed since, so the box empties */
  slow.setDraft("R0006", "the only note");
  const plain = slow.send("R0006", "Person A");
  await new Promise(r => setTimeout(r, 5));
  release();
  await plain;
  assert.strictEqual(slow.draftOf("R0006"), "", "an untouched box is emptied once the note is away");
  pass("a note typed while the last one is still sending is never wiped by it landing");

  /* typing is remembered without redrawing the card: the draft is deliberately
     not part of the card's signature, or the ten-second poll would take the
     caret out of the box six times a minute */
  const sigBefore = S("NOTES.sig('R0001')");
  S("NOTES.setDraft('R0001', 'half a sen');");
  assert.strictEqual(S("NOTES.sig('R0001')"), sigBefore,
    "a half-typed note does not change what the board thinks the card is showing");
  S("NOTES.setDraft('R0001', '');");
  pass("typing a note never makes the board redraw the card under the fingers");

  /* ================= 5. one channel, any station ================= */
  /* the point of the whole feature: a second station page passes its own name
     and gets the same thing, with no new list and no new drawer code */
  const cutting = ST.stationComments({ station: "Cutting",
    listItems: (name, o) => CW.listItems(name, o),
    listAdd: (name, fields, o) => CW.listAdd(name, fields, o),
    opts: () => ({ siteId: FSITE, fields: ST.COMMENT_FIELDS }) });
  assert.strictEqual(cutting.station, "Cutting");
  assert.strictEqual(cutting.list, ST.COMMENT_LIST, "the same list: a station is a value, not a list");
  await cutting.read();
  assert.deepStrictEqual(cutting.rows("R0001").map(r => r.text), ["a note from another station"],
    "and it sees its own station's notes on the job, not the glass tablet's");
  reset();
  cutting.setDraft("R0001", "cut list is wrong");
  await cutting.send("R0001", "Person C");
  assert.strictEqual(notesReq().filter(r => r.method === "POST")[0].body.fields.Station, "Cutting");
  pass("a second station passes one word and has the feature: same list, its own thread");

  /* ================= 6. the office drawer ================= */
  consent(true);
  A("STATION_NOTES = null; STATION_NOTES_OK = null; STATION_NOTES_WHY = ''; NOTES_SEEN = null;");
  A("CHANGES = [];");
  assert.ok(await A("readStationNotes()"), "the office read the same list");
  assert.strictEqual(A("STATION_NOTES_OK"), true);
  const j = mkJob({ id: "R0001" });
  global.__j = j;
  let html = A("floorNotesHtml(__j)");
  assert.ok(html.indexOf("Floor notes") >= 0, "the drawer has a Floor notes section");
  assert.ok(html.indexOf("yesterday: two units short") >= 0, "with the glass tablet's note");
  assert.ok(html.indexOf("a note from another station") >= 0, "and the cutting station's");
  assert.ok(html.indexOf("cut list is wrong") >= 0, "and the one cutting has just sent");
  assert.ok(html.indexOf("Person B") >= 0 && html.indexOf("Person C") >= 0, "tagged with who wrote each");
  assert.ok(html.indexOf(">Glass<") >= 0 && html.indexOf(">Cutting<") >= 0, "and with which station");
  assert.ok(html.indexOf("another job entirely") < 0, "and nothing from another job");
  pass("the drawer's Floor notes section shows every station's notes on that job, tagged");

  /* oldest first here too, and the office's half is READ-ONLY: no box, no
     button, nothing to edit and nothing to delete */
  const order = ["yesterday: two units short", "a note from another station"];
  assert.ok(html.indexOf(order[0]) < html.indexOf(order[1]), "oldest first");
  assert.ok(!/<textarea|<button|<input/.test(html), "no reply box and no control of any kind");
  assert.ok(!/data-cmdel|data-cmedit|>Delete<|>Edit</.test(html), "nothing to edit or remove a note with");
  pass("the office's half is read-only: no reply box, no edit, no delete");

  /* the rule that has held since 2026-09-09: per-job detail goes in the card,
     never in a new column on the job row */
  const appsrc = src("app.js");
  assert.ok(appsrc.indexOf("floorNotesHtml(j)") >= 0, "the drawer calls it");
  assert.ok(!/rowHtml[\s\S]{0,4000}floorNotesHtml/.test(appsrc),
    "and nothing on the job row does - no new column (owner, 2026-09-09)");
  pass("Floor notes is drawer-only: the job row gains nothing");

  /* ---- the Changes line ---- */
  A("CHANGES = [];");
  const before = A("CHANGES.length");
  assert.strictEqual(A("noteFloorNotes()"), 0,
    "a second pass over notes already seen says nothing: only what is new is news");
  NOTEITEMS.push(note("R0001", "Glass", "Person A", "the glass is here now", "2026-09-15T12:00:00Z", "77"));
  A("STATION_NOTES = null;");
  await A("readStationNotes()");
  assert.strictEqual(A("CHANGES.length"), before + 1, "one line for one new note");
  const line = A("CHANGES[0]");
  assert.strictEqual(line.what, "New floor note on R0001, from Glass, Person A",
    "in the owner's own words");
  assert.strictEqual(line.job, "R0001");
  assert.strictEqual(line.who, "Person A");
  assert.strictEqual(line.to, "the glass is here now", "with what was said");
  assert.strictEqual(line.src, "floor", "and marked as the floor's, not this dashboard's and not Excel's");
  pass("a new note is one Changes line: “New floor note on <job>, from <station>, <who>”");

  /* and the FIRST time this BROWSER ever looks at the list, nothing is
     announced: a page opening on two hundred old notes must not post two
     hundred lines to somebody who has never seen the feature */
  delete mem.cw_notesseen;
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; CHANGES = [];");
  await A("readStationNotes()");
  assert.strictEqual(A("CHANGES.length"), 0, "the first look is the baseline, not the news");
  assert.ok(mem.cw_notesseen, "and it is written down, so the NEXT page load knows what it has seen");
  pass("the first time a browser ever looks at the list, none of it is announced");

  /* ---- THE RELOAD (review finding 1) ----
     The set of seen ids has to outlive the page, because CHANGES does. Close
     the browser overnight with a note waiting, open it in the morning: without
     this, the waiting note is seeded as "already seen" and the only signal the
     spec gives the office never fires. */
  NOTEITEMS.push(note("R0001", "Glass", "Person B", "left overnight, nobody saw this",
                      "2026-09-15T22:00:00Z", "88"));
  /* a fresh page: in-memory state gone, localStorage kept, exactly as a reload
     leaves it - the Changes list comes back from cw_changes too */
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; NOTES_SEEN_ORDER = [];");
  A("CHANGES = JSON.parse(localStorage.getItem('cw_changes') || '[]');");
  await A("readStationNotes()");
  assert.strictEqual(A("CHANGES.filter(c => c.noteId === '88').length"), 1,
    "the note that arrived while the page was shut is announced when it opens");
  assert.strictEqual(A("CHANGES[0].what"), "New floor note on R0001, from Glass, Person B");
  pass("a note that arrived while the dashboard was closed IS announced on the next load");

  /* ... and only once, however many times the page is reloaded after that */
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; NOTES_SEEN_ORDER = [];");
  A("CHANGES = JSON.parse(localStorage.getItem('cw_changes') || '[]');");
  await A("readStationNotes()");
  assert.strictEqual(A("CHANGES.filter(c => c.noteId === '88').length"), 1,
    "a second reload does not announce it again");
  pass("and it is announced once, not once per reload");

  /* the ids of notes that are GONE are capped, so the key cannot grow without
     end, and the map and the list are trimmed together */
  A("STATION_NOTES = [];");                      // none of these ids is in the feed
  A("NOTES_SEEN = {}; NOTES_SEEN_ORDER = [];");
  A("for (let i = 0; i < NOTES_SEEN_MAX + 40; i++) notesSeenMark('x' + i); saveNotesSeen();");
  assert.strictEqual(A("NOTES_SEEN_ORDER.length"), A("NOTES_SEEN_MAX"), "the list is capped");
  assert.strictEqual(A("Object.keys(NOTES_SEEN).length"), A("NOTES_SEEN_MAX"),
    "and the map with it, so nothing is remembered that is no longer written down");
  assert.strictEqual(A("NOTES_SEEN['x0']"), undefined, "the oldest went");
  assert.strictEqual(A("NOTES_SEEN['x' + (NOTES_SEEN_MAX + 39)]"), 1, "the newest stayed");
  pass("the seen-notes memory is capped, and its map and its list are trimmed together");

  /* ---- PAST THE CAP, THROUGH THE REAL noteFloorNotes (review bug A) ----
     The cap above is about ids of notes that are gone. This is the case that
     actually broke: a LIST longer than the cap, walked by noteFloorNotes every
     pass. The oldest ids fell out of the memory, were announced again as new,
     and evicted a different block, which was announced again next pass - for
     ever, at exactly (rows - cap) false lines per pass, six times a minute for
     five minutes whenever a list refuses a delta. A synthetic
     notesSeenMark/saveNotesSeen test cannot see it, because its ids are never
     in STATION_NOTES: this one drives the real thing. */
  const bulk = total => {
    NOTEITEMS = [];
    for (let i = 0; i < total; i++)
      NOTEITEMS.push(note("R" + (1000 + (i % 40)), "Glass", "Person A", "note " + i,
        new Date(Date.parse("2026-09-01T00:00:00Z") + i * 60000).toISOString(), String(20000 + i)));
  };
  const spuriousAt = async total => {
    bulk(total);
    delete mem.cw_notesseen;
    A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; NOTES_SEEN_ORDER = []; CHANGES = [];");
    await A("readStationNotes()");                 // the baseline: none of it is news
    assert.strictEqual(A("CHANGES.length"), 0, total + " rows: the first look announces none of them");
    /* three more passes over a list that has not changed by so much as a row */
    return A("noteFloorNotes()") + A("noteFloorNotes()") + A("noteFloorNotes()");
  };
  assert.strictEqual(await spuriousAt(A("NOTES_SEEN_MAX") - 1), 0);
  assert.strictEqual(await spuriousAt(A("NOTES_SEEN_MAX")), 0);
  assert.strictEqual(await spuriousAt(A("NOTES_SEEN_MAX") + 1), 0,
    "one row past the cap announced one ghost per pass before this was fixed");
  assert.strictEqual(await spuriousAt(A("NOTES_SEEN_MAX") + 100), 0,
    "and a hundred rows past it announced a hundred ghosts per pass, every pass");
  /* ... because an id still in the feed can never be evicted */
  A("STATION_NOTES = [];");
  assert.ok(A("NOTES_SEEN_ORDER.length") >= A("NOTES_SEEN_MAX"),
    "the memory is allowed over the cap to hold every row the feed still has");
  pass("a list longer than the seen-ids cap announces nothing twice, however many passes run");

  /* and the feed itself holds a window, so that memory stays bounded: the same
     ninety days the log gets, for the same reason */
  const old = note("R0001", "Glass", "Person A", "from last spring", "2025-01-01T09:00:00Z", "30001");
  const fresh = note("R0001", "Glass", "Person A", "this week", new Date().toISOString(), "30002");
  NOTEITEMS = [old, fresh];
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; NOTES_SEEN_ORDER = [];");
  delete mem.cw_notesseen;
  await A("readStationNotes()");
  assert.deepStrictEqual(A("STATION_NOTES.map(x => x.id)"), ["30002"],
    "a note older than the window is not carried around for ever");
  assert.strictEqual(A("stationNotesRecent([]).length"), 0);
  assert.strictEqual(A("stationNotesRecent([{ id: '1', fields: { At: '' } }]).length"), 1,
    "and a row with no stamp at all is kept, exactly as the log keeps one");
  pass("the notes feed holds the same ninety-day window the log does");

  /* ---- nothing about a note is written into the workbook ----
     NOT a string match near a call: the real functions are replaced with
     counters and noteFloorNotes is run for a note it has never seen, so a
     regression that logged the note under any name at all would be caught. */
  delete mem.cw_notesseen;
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; NOTES_SEEN_ORDER = []; CHANGES = [];");
  await A("readStationNotes()");                  // baseline: this browser has now seen them
  const realNoteChange = global.noteChange, realAppendLog = CW.appendLog;
  let noteChangeCalls = 0, appendLogCalls = 0;
  global.noteChange = function () { noteChangeCalls++; return realNoteChange.apply(null, arguments); };
  CW.appendLog = function () { appendLogCalls++; return realAppendLog.apply(CW, arguments); };
  reset();
  NOTEITEMS.push(note("R0001", "Glass", "Person A", "this must not reach the log sheet",
                      "2026-09-15T23:00:00Z", "99"));
  A("STATION_NOTES = null;");
  await A("readStationNotes()");
  assert.strictEqual(A("CHANGES.filter(c => c.noteId === '99').length"), 1, "the note was announced");
  assert.strictEqual(noteChangeCalls, 0, "and noteChange - which writes the Dashboard Log - was never called");
  assert.strictEqual(appendLogCalls, 0, "nor CW.appendLog underneath it");
  assert.deepStrictEqual(REQ.filter(r => r.method !== "GET").map(r => r.method + " " + r.path), [],
    "announcing a note sent no write of any kind, anywhere");
  global.noteChange = realNoteChange; CW.appendLog = realAppendLog;
  pass("announcing a note calls neither noteChange nor appendLog, and writes nothing at all");

  /* ================= 7. a missing list, on both pages ================= */
  NOTES_EXISTS = false;
  CW._resetListIds(); delete mem.cw_listids;

  /* ---- the tablet ---- */
  reset();
  const orphan = ST.stationComments({ station: "Glass",
    listItems: (name, o) => CW.listItems(name, o),
    listAdd: (name, fields, o) => CW.listAdd(name, fields, o),
    opts: () => ({ siteId: FSITE, fields: ST.COMMENT_FIELDS }) });
  assert.strictEqual(await orphan.read(), false, "a missing list is not a read");
  assert.strictEqual(orphan.state.ok, false);
  assert.ok(/Station comments/.test(orphan.state.why), "the state says which list is missing");
  assert.ok(/Ask the office/.test(orphan.state.why), "and who to ask");
  assert.ok(/Excel file is involved/.test(orphan.state.why), "and that the workbook has nothing to do with it");
  orphan.toggle("R0001");
  const orphanHtml = orphan.html("R0001");
  assert.ok(orphanHtml.indexOf("Station comments") >= 0, "the composer says it plainly");
  assert.ok(!/data-cmbox|data-cmsend/.test(orphanHtml), "and offers no box to type into and no Send");
  orphan.setDraft("R0001", "nowhere to put this");
  assert.strictEqual(await orphan.send("R0001", "Person A"), null, "a send writes nothing at all");
  assert.strictEqual(REQ.filter(r => r.method !== "GET").length, 0, "not one write left the tablet");
  pass("a missing list on the tablet: an explained line, no box, and nothing written anywhere");

  /* and the board is untouched by it - which is the point: a note channel that
     is not set up yet must not stop the floor recording work */
  S("SITEID = " + JSON.stringify(FSITE) + ";");
  assert.strictEqual(await S("NOTES.read()"), false);
  S("if (!NOTES.isOpen('R0001')) NOTES.toggle('R0001');");   // forced open, not "if it happens to be"
  S("render()");
  assert.strictEqual(S("NOTES.isOpen('R0001')"), true, "the composer really is open for this check");
  assert.strictEqual(S("PROBLEM"), "", "the board has no problem");
  assert.strictEqual(S("READY"), true, "and is still drawn");
  assert.ok(/data-stage="cut"/.test(boardHtml()), "with its steppers exactly as before");
  assert.ok(boardHtml().indexOf("Station comments") >= 0,
    "and the explanation is inside the open composer, where somebody looking for it will be");
  assert.ok(!/data-cmbox|data-cmsend/.test(boardHtml()), "with no box and no Send on the card either");
  pass("a missing list never takes the tablet's board away");

  /* ---- the office ---- */
  A("STATION_NOTES = null; STATION_NOTES_OK = null; STATION_NOTES_WHY = ''; NOTES_SEEN = null;");
  assert.strictEqual(await A("readStationNotes()"), null);
  assert.strictEqual(A("STATION_NOTES_OK"), false);
  html = A("floorNotesHtml(__j)");
  assert.ok(html.indexOf("Floor notes") >= 0, "the section is still there");
  assert.ok(html.indexOf("Station comments") >= 0, "saying which list is missing");
  assert.ok(/Ask the manager/.test(html), "and who to ask");
  assert.ok(html.indexOf("Excel file is involved") >= 0);
  assert.ok(!/<textarea|<button|<input/.test(html), "and nothing to click");
  pass("a missing list in the drawer: the section says which list, plainly, and nothing breaks");

  /* the notes list failing must not touch the board's own state, ever: they are
     separate lists and separate failures */
  A("STATION_OK = true; STATION_WHY = ''; STATION_ERR = '';");
  A("STATION_NOTES_OK = null; STATION_NOTES = null; NOTES_SEEN = null;");
  await A("readStationNotes()");
  assert.strictEqual(A("STATION_OK"), true, "the Glass station board is untouched by it");
  assert.strictEqual(A("STATION_WHY"), "");
  assert.strictEqual(A("STATION_ERR"), "");
  pass("the notes list is its own failure: the floor's board never suffers for it");

  /* ---- A MISSING LIST IS "ASK AGAIN LATER", NOT A SETTLED ANSWER ----
     Review finding 4, and it is tomorrow's case, not a hypothetical: the list
     does not exist yet. The owner creates it mid-morning, and a dashboard
     opened before that must pick it up on its own. */
  NOTES_EXISTS = false;
  CW._resetListIds(); delete mem.cw_listids;
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SOFT = 0; NOTES_SOFT_MS = 0;");
  await A("readStationNotes()");
  assert.strictEqual(A("STATION_NOTES_OK"), false);
  assert.strictEqual(A("STATION_NOTES_WHY"), ST.COMMENT_MISSING_OFFICE);
  assert.ok(A("NOTES_SOFT") > 0, "a missing list is marked to be asked about again");
  assert.strictEqual(A("NOTES_SOFT_MS"), A("NOTES_MISS_RETRY_MS"),
    "on the longer clock, because a list appearing is rarer than the wifi dropping");
  A("stationNotesReadIfNeeded()");
  await new Promise(r => setTimeout(r, 10));
  assert.strictEqual(A("STATION_NOTES_OK"), false, "a caller a moment later does not hammer it");
  /* the owner creates the list; five minutes later the open dashboard finds it,
     with nobody reloading anything and no drawer being opened */
  NOTES_EXISTS = true;
  CW._resetListIds(); delete mem.cw_listids;
  A("NOTES_SOFT = Date.now() - (NOTES_MISS_RETRY_MS + 1000);");
  A("stationNotesReadIfNeeded()");
  await new Promise(r => setTimeout(r, 30));
  assert.strictEqual(A("STATION_NOTES_OK"), true,
    "the morning the list is created, a dashboard already open picks it up on its own");
  assert.strictEqual(A("NOTES_SOFT"), 0);
  pass("a list that does not exist yet is asked about again, and found the moment it is made");

  /* a passing failure is the same idea on a shorter clock: the workshop wifi
     dropping once at start-up must not switch the channel off for the day */
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SOFT = 0; NOTES_SOFT_MS = 0;");
  FAIL_NOTES_GET = 1;
  await A("readStationNotes()");
  assert.strictEqual(A("STATION_NOTES_OK"), false);
  assert.strictEqual(A("STATION_NOTES_WHY"), ST.COMMENT_UNREACHABLE, "said as a passing failure");
  assert.strictEqual(A("NOTES_SOFT_MS"), A("NOTES_RETRY_MS"), "and asked about again in a minute");
  A("stationNotesReadIfNeeded()");
  await new Promise(r => setTimeout(r, 5));
  assert.strictEqual(A("STATION_NOTES_OK"), false, "a caller a moment later does not hammer it");
  A("NOTES_SOFT = Date.now() - (NOTES_RETRY_MS + 1000);");
  A("stationNotesReadIfNeeded()");
  await new Promise(r => setTimeout(r, 20));
  assert.strictEqual(A("STATION_NOTES_OK"), true, "a minute later it tries again, and gets it");
  assert.strictEqual(A("NOTES_SOFT"), 0);
  pass("a passing failure is retried a minute later, on its own shorter clock");

  /* ---- THE NOTES LIST'S FAILURES ARE ITS OWN, INSIDE THE POLL ----
     Review finding 2. stationDelta only swallows the 4xx that means "restart
     the delta"; a 503 is re-thrown. If that landed in stationPoll's own catch
     it would put an error on a board that read perfectly well AND return before
     glassColourRun(), so a floor tap made that cycle would never reach the
     Production sheet. This drives the real stationPoll. */
  A("STATION_OK = true; STATION_ERR = ''; STATION_WHY = '';");
  A("STATION_ITEMS = []; STATION_LOG = []; STATION_LOG_OK = true;");
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SOFT = 0; NOTES_SOFT_MS = 0;");
  await A("readStationNotes()");
  assert.strictEqual(A("STATION_NOTES_OK"), true, "the channel is up before the failure is staged");
  A("STATION_FEEDS.items.token = null; STATION_FEEDS.log.token = null; STATION_FEEDS.notes.token = null;");
  await A("stationPoll()");                        // a clean pass, so every feed holds a token
  let glassRuns = 0;
  const realGlassRun = global.glassColourRun;
  global.glassColourRun = async function () { glassRuns++; return realGlassRun ? undefined : undefined; };
  FAIL_NOTES_GET = 9;                              // the notes delta refuses, over and over
  ITEMS.push(item({ Title: "R0002", Job: "R0002", Customer: "Customer Two", GlassType: "GLASS",
                    Total: 2, Seq: 2, Active: "Yes", Cut: 1, Hotmelt: 0, Glazed: 0 }, "901"));
  DELTA_NEXT[GLASS_ID] = [[{ id: "901", fields: ITEMS[ITEMS.length - 1].fields }]];
  const moved = await A("stationPoll()");
  assert.strictEqual(moved, true, "the board's own delta still counts as movement");
  assert.strictEqual(A("STATION_OK"), true, "the board is still readable");
  assert.strictEqual(A("STATION_ERR"), "", "and says nothing about a floor it read perfectly well");
  assert.strictEqual(glassRuns, 1,
    "and the floor's tap still reached glassColourRun - a flapping notes list cannot stall the colours");
  assert.ok(A("NOTES_SOFT") > 0, "the notes channel took the failure, alone, and backed off");
  assert.strictEqual(A("STATION_NOTES_OK"), true,
    "keeping the notes it had already read on screen - one 503 must not blank a job's notes");
  /* and it stops asking: a refusing list is not polled six times a minute for
     the rest of the afternoon */
  reset();
  ITEMS[ITEMS.length - 1].fields.Cut = 2;                      // the floor taps again
  DELTA_NEXT[GLASS_ID] = [[{ id: "901", fields: ITEMS[ITEMS.length - 1].fields }]];
  await A("stationPoll()");
  assert.deepStrictEqual(notesReq().map(r => r.method + " " + r.path), [],
    "the very next poll does not go near the notes list at all");
  assert.strictEqual(glassRuns, 2, "while everything else about the poll carries on as normal");
  global.glassColourRun = realGlassRun;
  FAIL_NOTES_GET = 0;
  pass("a notes-list failure inside the poll never reaches the board, its error line or the glass colours");

  /* ---- THE LIST VANISHES AFTER A GOOD READ (review bug B) ----
     Not a throw, and that is the point: the list is renamed, deleted, or left
     behind in the old site when the others are moved across - the migration
     STATIONS.md describes, and this is the one list that postdates those notes.
     `listItems` answers null, nothing is thrown, and read as "nothing moved"
     the channel stayed healthy for ever: a fresh enumeration of a list that is
     not there every ten seconds, a drawer rendering rows nobody could refresh,
     and no Changes line ever again. */
  NOTES_EXISTS = true;
  CW._resetListIds(); delete mem.cw_listids; delete mem.cw_notesseen;
  NOTEITEMS = [note("R0001", "Glass", "Person A", "still here", new Date().toISOString(), "40001")];
  A("STATION_OK = true; STATION_ERR = ''; STATION_LOG_OK = true;");
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; NOTES_SEEN_ORDER = [];");
  A("NOTES_SOFT = 0; NOTES_SOFT_MS = 0;");
  await A("readStationNotes()");
  assert.strictEqual(A("STATION_NOTES_OK"), true, "a good read first");
  A("STATION_FEEDS.items.token = null; STATION_FEEDS.log.token = null; STATION_FEEDS.notes.token = null;");
  await A("stationPoll()");
  /* now it is gone, and the plain-read path is the one in use - which is what a
     deltaOff window puts every poll down for five minutes at a time */
  NOTES_EXISTS = false;
  CW._resetListIds(); delete mem.cw_listids;
  A("STATION_FEEDS.notes.off = Date.now();");
  reset();
  await A("stationPoll()");
  assert.strictEqual(A("STATION_NOTES_OK"), false, "the channel notices it has gone");
  assert.strictEqual(A("STATION_NOTES_WHY"), ST.COMMENT_MISSING_OFFICE, "and says which list, plainly");
  assert.ok(A("NOTES_SOFT") > 0, "and arms its clock instead of asking again in ten seconds");
  assert.strictEqual(A("NOTES_SOFT_MS"), A("NOTES_MISS_RETRY_MS"));
  assert.strictEqual(A("STATION_OK"), true, "the board is untouched by it, as ever");
  assert.strictEqual(A("STATION_ERR"), "");
  /* three more polls: not one of them goes looking for the list */
  for (let i = 0; i < 3; i++) {
    reset();
    await A("stationPoll()");
    assert.deepStrictEqual(notesReq().map(r => r.method + " " + r.path), [],
      "poll " + (i + 2) + " after it vanished asks for nothing");
  }
  /* and the drawer stops pretending: the rows it cannot refresh are replaced by
     the plain explanation, rather than sitting there looking current */
  const gone = A("floorNotesHtml(__j)");
  assert.ok(gone.indexOf("Station comments") >= 0, "the drawer says the list is missing");
  assert.ok(gone.indexOf("still here") < 0, "rather than going on showing rows nobody can refresh");
  pass("a list that vanishes after a good read degrades, arms its clock and stops asking");

  /* a job with no notes at all, and the list perfectly readable */
  NOTES_EXISTS = true;
  CW._resetListIds(); delete mem.cw_listids;
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null;");
  await A("readStationNotes()");
  global.__k = mkJob({ id: "R9999", glass: {} });
  const quiet = A("floorNotesHtml(__k)");
  assert.ok(quiet.indexOf("No notes from the floor") >= 0, "says so plainly rather than showing nothing");
  assert.ok(quiet.indexOf("Floor notes") >= 0);
  pass("a job nobody has written about says so, and a job with no glass still has the section");

  /* ================= 8. the owner's additions, amendment E =================
     E1 the Components F/S/T cell emptied; E2 an unread-notes icon on the job
     row, seen per screen in `cw_notesread`; E3 the same notes, and the same
     icon, on the station board's cards.                                     */

  /* the fixture for all of it: two stations writing about one job, a second
     job nobody has written about, and a note with quotes and a tag in it */
  NOTES_EXISTS = true;
  CW._resetListIds(); delete mem.cw_listids;
  const E_A = "2026-09-15T08:58:00Z", E_B = "2026-09-15T09:30:00Z";
  NOTEITEMS = [
    note("R0001", "Glass", "Person A", "two DG units short, need more stock", E_A, "300"),
    note("R0001", "Cutting", "Person B", 'the cut list says <b>4</b> & "it is not"', E_B, "301")
  ];
  delete mem.cw_notesread;
  A("NOTES_READ = null; NOTES_UNREAD = null;");
  A("STATION_NOTES = null; STATION_NOTES_OK = null; NOTES_SEEN = null; NOTES_SEEN_ORDER = []; CHANGES = [];");
  await A("readStationNotes()");
  global.__j = mkJob({ id: "R0001" });
  const rowOf = j => A("rowHtml(" + j + ", 0, 10)");
  const iconOf = job => A("notesIconHtml(" + JSON.stringify(job) + ")");

  /* ---- E2: what is unread, and when ---- */
  assert.strictEqual(A("notesUnreadFor('R0001').length"), 2,
    "a screen with no stamp at all has read none of them: no seeding on a fresh browser");
  assert.strictEqual(A("notesUnreadFor('R9999').length"), 0, "and a job nobody wrote about has none");
  const seenAt = at => A("NOTES_READ = { R0001: { at: " + JSON.stringify(at) + ", ids: [] } }; NOTES_UNREAD = null;");
  seenAt(E_A);
  assert.deepStrictEqual(A("notesUnreadFor('R0001').map(r => r.id)"), ["301"],
    "a stamp older than one note leaves exactly that note unread");
  seenAt(E_B);
  assert.strictEqual(A("notesUnreadFor('R0001').length"), 0, "a stamp on the newest note leaves none");
  pass("unread means “newer than what this screen has opened”: none, one, or all of them");

  /* ---- the stamps are MOMENTS, not strings (fix pass, 2026-09-16) ----
     the tablet writes milliseconds and a row typed in by hand has none, so the
     two shapes meet in this comparison all the time. As text ".100Z" sorts
     BEFORE "Z" - which made a note a tenth of a second later read as already
     seen, for ever, on every screen that had the undated stamp. */
  NOTEITEMS.push(note("R0001", "Glass", "Person A", "a tenth of a second later",
                      "2026-09-15T09:30:00.100Z", "310"));
  A("STATION_NOTES = null; STATION_NOTES_OK = null;");
  await A("readStationNotes()");
  seenAt("2026-09-15T09:30:00Z");
  assert.deepStrictEqual(A("notesUnreadFor('R0001').map(r => r.id)"), ["310"],
    "a note with milliseconds, a tenth of a second after the stamp without them, is unread");
  assert.strictEqual(ST.atCmp("2026-09-15T09:30:00.100Z", "2026-09-15T09:30:00Z"), 1,
    "because the stamps are compared as the moments they name");
  assert.ok("2026-09-15T09:30:00.100Z" < "2026-09-15T09:30:00Z",
    "which as plain text they are not - this is the whole bug");
  assert.strictEqual(ST.atCmp("2026-09-15T09:30:00.100Z", "2026-09-15T09:30:00.100Z"), 0, "equal is equal");
  assert.strictEqual(ST.atCmp("about noon", "yesterday"), -1,
    "and two stamps no clock can read fall back to the text comparison they always had");
  assert.strictEqual(ST.atCmp("yesterday", "about noon"), 1);
  assert.strictEqual(ST.atCmp("", ""), 0);
  /* and the same comparison decides the ORDER the two screens draw them in */
  assert.deepStrictEqual(ST.commentRows([
    note("R5", "Glass", "Person A", "later", "2026-09-15T09:30:00.100Z", "2"),
    note("R5", "Glass", "Person A", "earlier", "2026-09-15T09:30:00Z", "1")
  ], { job: "R5" }).map(r => r.text), ["earlier", "later"],
    "oldest first means oldest in time, whichever of them carries milliseconds");
  A("markNotesRead('R0001'); NOTES_UNREAD = null;");
  assert.strictEqual(A("notesUnreadFor('R0001').length"), 0, "and opening the job still settles it");
  NOTEITEMS.pop();
  pass("a stamp with milliseconds and one without are compared as moments, not as text");

  /* a note typed into SharePoint by hand with no stamp at all: unread until
     the job is opened, seen from then on */
  NOTEITEMS.push(item({ Title: "R0001|x", Job: "R0001", Station: "Glass", Who: "Person C",
                        Text: "no stamp on this one", At: "" }, "302"));
  A("STATION_NOTES = null; STATION_NOTES_OK = null;");
  await A("readStationNotes()");
  A("NOTES_READ = {}; NOTES_UNREAD = null;");
  assert.strictEqual(A("notesUnreadFor('R0001').length"), 3, "with no stamp for the job, all three are unread");
  assert.strictEqual(A("markNotesRead('R0001')"), true, "opening the job writes an entry");
  assert.strictEqual(A("notesUnreadFor('R0001').length"), 0, "and the one with no At counts as seen after it");
  assert.strictEqual(A("NOTES_READ.R0001.at"), E_B,
    "the stamp is the newest DATED note, never a clock reading - a clock would bury the next undated one");
  assert.deepStrictEqual(A("NOTES_READ.R0001.ids"), ["302"], "and the undated note is remembered by its id");
  assert.strictEqual(A("markNotesRead('R0001')"), false, "a second open changes nothing and repaints nothing");
  assert.strictEqual(A("markNotesRead('R9999')"), false, "and a job with no notes never gets an entry at all");
  assert.strictEqual(A("NOTES_READ.R9999"), undefined);

  /* a SECOND undated note, typed in after the job was opened. Under the first
     shape of this key the stamp was "now" and this note was born already seen -
     the office would never learn it existed. */
  NOTEITEMS.push(item({ Title: "R0001|y", Job: "R0001", Station: "Glass", Who: "Person C",
                        Text: "and another with no stamp", At: "" }, "303"));
  A("STATION_NOTES = null; STATION_NOTES_OK = null;");
  await A("readStationNotes()");
  assert.deepStrictEqual(A("notesUnreadFor('R0001').map(r => r.id)"), ["303"],
    "the new undated note is unread, and the one already seen stays seen");
  assert.strictEqual(A("markNotesRead('R0001')"), true, "opening the job takes it in");
  assert.deepStrictEqual(A("NOTES_READ.R0001.ids"), ["302", "303"], "by id, beside the first");
  assert.strictEqual(A("notesUnreadFor('R0001').length"), 0);
  NOTEITEMS.pop(); NOTEITEMS.pop();
  pass("a note with no At is unread until the job is opened - every one of them, not just the first");

  /* the shape the first build of this key wrote, in a browser that ran it */
  mem.cw_notesread = JSON.stringify({ R0001: E_B, R0002: { at: E_A }, R0003: 7, R0004: null });
  A("NOTES_READ = null; NOTES_UNREAD = null;");
  A("notesReadRestore()");
  assert.deepStrictEqual(A("NOTES_READ.R0001"), { at: E_B, ids: [] },
    "a bare string reads as that stamp with no ids - no throw, no reset, nothing re-announced");
  assert.deepStrictEqual(A("NOTES_READ.R0002"), { at: E_A, ids: [] }, "and a half-written entry is filled in");
  assert.deepStrictEqual(A("NOTES_READ.R0003"), { at: "", ids: [] }, "rubbish reads as nothing seen");
  assert.deepStrictEqual(A("NOTES_READ.R0004"), { at: "", ids: [] });
  mem.cw_notesread = "{not json";
  A("NOTES_READ = null; notesReadRestore();");
  assert.deepStrictEqual(A("NOTES_READ"), {}, "and an unreadable key leaves everything unread, never a throw");
  delete mem.cw_notesread;
  pass("the seen key from the first build of this feature is read, not thrown over");

  /* ---- E2: the icon on the row, and the seen state behind it ---- */
  delete mem.cw_notesread;
  A("NOTES_READ = null; NOTES_UNREAD = null; STATION_NOTES = null; STATION_NOTES_OK = null;");
  await A("readStationNotes()");
  let row = rowOf("__j");
  assert.ok(/data-notes="R0001"/.test(row), "the row carries the icon while something is unread");
  assert.ok(/\u{1F4AC} 2/u.test(row), "with the count of what is unread on it");
  const badgeCellOf = r => r.slice(r.indexOf('overflow:hidden'), r.lastIndexOf("</div>"));
  assert.ok(badgeCellOf(row).indexOf("data-notes") > 0,
    "in the badge cell with the other chips - never a column of its own (owner, 2026-09-09)");
  assert.strictEqual((A("chipsNow.toString()").match(/notesUnreadFresh/g) || []).length, 1,
    "and what the rows are saying includes it, so the quiet repaint path notices a change");
  pass("a job with unread notes draws a speech bubble and a count in its badge cell");

  /* opening the drawer, by any route, is what marks them seen */
  /* the drawer also asks after `Dashboard progress`, which lives in the
     WORKBOOK's site and this test serves no workbook at all: told it has
     already looked, it leaves that list alone and the run stays quiet */
  A("CP_LIST_OK = false; ALL = [__j]; state.sel = 'R0001'; state.edit = false;");
  A("openDrawer()");
  const stored = JSON.parse(mem.cw_notesread || "{}");
  assert.strictEqual(stored.R0001.at, E_B, "the newest At of that job goes into cw_notesread, under the job");
  assert.deepStrictEqual(stored.R0001.ids, [], "with no ids, because every note on it carries a stamp");
  assert.deepStrictEqual(Object.keys(stored), ["R0001"], "one entry per job, not one per note");
  assert.strictEqual(iconOf("R0001"), "", "and the row's icon html is empty afterwards");
  assert.ok(rowOf("__j").indexOf("data-notes") < 0);
  pass("opening a job stores the newest note's At under cw_notesread and the icon goes");

  /* and a note arriving after that brings it back */
  NOTEITEMS.push(note("R0001", "Glass", "Person A", "and one more thing", "2026-09-15T11:00:00Z", "303"));
  A("STATION_NOTES = null; STATION_NOTES_OK = null;");
  await A("readStationNotes()");
  assert.deepStrictEqual(A("notesUnreadFor('R0001').map(r => r.id)"), ["303"], "only the new one is unread");
  assert.ok(rowOf("__j").indexOf('data-notes="R0001"') > 0, "so the icon is back on the row");
  assert.ok(/\u{1F4AC} 1/u.test(rowOf("__j")), "saying one, not four");
  pass("a note arriving after the job was opened makes the icon non-empty again");

  /* the tooltip: the native title attribute, one line a note, and escaped */
  const title = (iconOf("R0001").match(/title="([^"]*)"/) || [])[1] || "";
  A("NOTES_READ = {}; NOTES_UNREAD = null;");
  const fullTitle = (iconOf("R0001").match(/title="([^"]*)"/) || [])[1] || "";
  assert.ok(title.indexOf("and one more thing") >= 0, "the unread note is in it");
  assert.strictEqual(fullTitle.split("\n").length, 3, "one line per unread note");
  assert.ok(fullTitle.indexOf("Glass") >= 0 && fullTitle.indexOf("Cutting") >= 0, "each line says which station");
  assert.ok(fullTitle.indexOf("Person A") >= 0 && fullTitle.indexOf("Person B") >= 0, "and who wrote it");
  assert.ok(fullTitle.indexOf(A("stWhen(" + JSON.stringify(E_A) + ")")) >= 0, "and when, in the list's own words");
  assert.ok(fullTitle.indexOf("&lt;b&gt;4&lt;/b&gt;") >= 0 && fullTitle.indexOf("&quot;it is not&quot;") >= 0,
    "and the note itself, escaped - a tag or a quote in a note cannot break out of the attribute");
  assert.ok(fullTitle.indexOf("<b>") < 0 && fullTitle.indexOf('"it is not"') < 0);
  pass("the tooltip is the native title: station, who, when and the text of every unread note, escaped");

  /* the icon must not tick the row's pick box and must not start a drag */
  const wireSrc = A("wireRows.toString()") + A("wireNotesIcons.toString()");
  assert.ok(/stopPropagation/.test(wireSrc), "the icon's click never reaches the row");
  assert.ok(/inNotesIcon\(e\.target\)/.test(A("wireRows.toString()")),
    "and the row's own click and dragstart both step aside for it");
  assert.strictEqual((A("wireRows.toString()").match(/inNotesIcon/g) || []).length, 2,
    "both of them, not just the click");
  assert.ok(/draggable="false"/.test(iconOf("R0001")), "and the badge itself is not draggable");
  assert.ok(/id="floornotes"/.test(A("floorNotesHtml(__j)")),
    "the Floor notes section carries the anchor the icon scrolls to");
  assert.ok(/scrollIntoView/.test(A("openJobNotes.toString()")), "which is what the click does with it");
  pass("the icon opens the job at its Floor notes, and neither picks the row nor drags it");

  /* and the handlers themselves, fired: a stub DOM cannot bubble an event, so
     the icon and the row it sits in are wired separately and the icon's own
     click is fired at both to prove what each of them does with it */
  const seen = { stop: 0, prevent: 0, picked: 0, dragged: 0 };
  const fakeIcon = { dataset: { notes: "R0001" }, classList: { contains: () => false },
                     closest: s => (s === "[data-notes]" ? fakeIcon : null) };
  const fakeRow = { dataset: { id: "R0001" }, classList: { add() {}, remove() {}, toggle() {} },
                    querySelector: () => ({ set onchange(f) { seen.picked++; }, checked: false }) };
  global.__scope = { querySelectorAll: s => (s === "[data-notes]" ? [fakeIcon] : [fakeRow]) };
  A("ALL = [__j]; state.sel = null;");
  A("wireRows(__scope)");
  const ev = { target: fakeIcon, stopPropagation: () => seen.stop++, preventDefault: () => seen.prevent++,
               get dataTransfer() { seen.dragged++; throw new Error("the row started a drag"); } };
  fakeRow.onclick(ev);
  assert.strictEqual(A("state.sel"), null, "the ROW's click handler ignores a click on the icon");
  assert.strictEqual(fakeRow.ondragstart(ev), false, "and its dragstart refuses to start on the icon");
  assert.strictEqual(seen.dragged, 0, "without ever reaching the drag's own payload");
  fakeIcon.onclick(ev);
  assert.strictEqual(seen.stop, 1, "the icon's click stops there and never reaches the row");
  assert.strictEqual(A("state.sel"), "R0001", "and opens that job");
  assert.strictEqual(JSON.parse(mem.cw_notesread || "{}").R0001 !== undefined, true,
    "which marks its notes seen, like any other way into the drawer");
  pass("fired for real: the icon opens the job, the row under it neither selects nor drags");

  /* ---- E1: the Components cell is empty, and nothing else moved ---- */
  row = rowOf("__j");
  assert.ok(row.indexOf('class="mini"') < 0, "no three-colour components bar on the row any more");
  assert.ok(row.indexOf("var(--f)") < 0 && row.indexOf("var(--s)") < 0 && row.indexOf("var(--t)") < 0,
    "and none of the three colours it was drawn in");
  const page = src("index.html");
  const headFrom = page.indexOf('<div class="thead">');
  const headBlock = page.slice(headFrom, page.indexOf("</div>", headFrom));
  assert.ok(/<span class="kick"><\/span>/.test(headBlock), "the header cell is there and says nothing");
  assert.ok(headBlock.indexOf("Components") < 0, "the words Components F/S/T are gone from the head");
  assert.strictEqual((headBlock.match(/<span/g) || []).length, 9,
    "the column itself stays: nine head cells, as before, so nothing on the row moves");
  assert.ok(/grid-template-columns:28px 92px minmax\(110px,1\.4fr\) minmax\(80px,1fr\) 124px 66px 116px 96px/.test(page),
    "and the grid keeps the 116px track, so the space is reserved rather than reclaimed");
  /* the drawer's own tile is untouched - the count only left the ROW */
  const dsrc = src("app.js");
  assert.ok(/\[tot\(comp\(j\)\), "components"\]/.test(dsrc), "the drawer still counts the components");
  pass("E1: the row's components bar and the header text are gone, the column and the drawer tile stay");

  /* ---- E3: the notes on the station board's cards ---- */
  ITEMS = [item({ Title: "R0001", Job: "R0001", Customer: "Customer One", GlassType: "GLASS",
                  Total: 4, TuffTotal: 0, Seq: 1, Active: "Yes", OfficeDone: "No",
                  Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0 }, "900"),
           item({ Title: "R0002", Job: "R0002", Customer: "Customer Two", GlassType: "GLASS",
                  Total: 2, TuffTotal: 0, Seq: 2, Active: "Yes", OfficeDone: "No",
                  Cut: 0, Hotmelt: 0, Glazed: 0, Tuff: 0 }, "901")];
  A("STATION_ITEMS = " + JSON.stringify(ITEMS) + "; STATION_OK = true; STATION_WHY = ''; STATION_ERR = '';");
  A("STATION_LOG = []; STATION_LOG_OK = true;");
  A("NOTES_READ = {}; NOTES_UNREAD = null;");
  const board = A("stationBoardHtml()");
  const cardOf = job => {
    const at = board.indexOf(">" + job + "<");
    const next = board.indexOf('<div class="stcard', at);
    return board.slice(board.lastIndexOf('<div class="stcard', at), next < 0 ? board.length : next);
  };
  const c1 = cardOf("R0001"), c2 = cardOf("R0002");
  assert.ok(c1.indexOf("two DG units short") >= 0, "the glass station's note is on the card");
  assert.ok(c1.indexOf("the cut list says") >= 0, "and the cutting station's: every station, not only Glass");
  assert.ok(c1.indexOf("two DG units short") < c1.indexOf("the cut list says"), "oldest first, as in the drawer");
  assert.ok(c1.indexOf(">Glass<") >= 0 && c1.indexOf(">Cutting<") >= 0, "each tagged with who said it");
  assert.ok(c1.indexOf("class=\"stnotes\"") >= 0);
  assert.ok(c2.indexOf("stnotes") < 0 && c2.indexOf("cmt") < 0,
    "a card nobody has written about carries no note markup at all");
  assert.ok(c1.indexOf('data-notes="R0001"') >= 0, "the unread icon is in the card head too");
  assert.ok(c2.indexOf("data-notes") < 0);
  /* the same line helper as the drawer, so the two can never drift apart */
  assert.strictEqual((dsrc.match(/floorNoteLineHtml/g) || []).length, 3,
    "one helper, defined once and called by the drawer and the card");
  pass("E3: the board's cards carry that job's notes, oldest first, from every station");

  /* the three channel states say nothing on the board: no banner, no hint */
  A("STATION_NOTES_OK = false; STATION_NOTES_WHY = ST.COMMENT_MISSING_OFFICE; NOTES_UNREAD = null;");
  const broken = A("stationBoardHtml()");
  assert.ok(broken.indexOf("Station comments") < 0, "a missing list is not announced on the board");
  assert.ok(broken.indexOf("cmt") < 0, "and no stale note is left on a card");
  assert.ok(broken.indexOf('class="stcard') > 0, "the board itself is exactly as it was");
  A("STATION_NOTES_OK = true; STATION_NOTES_WHY = '';");
  pass("the board says nothing about a channel that is checking, missing or unreachable");

  /* and the office still only ever READS this list: every request made while
     nothing but app.js was running is a GET */
  reset();
  A("STATION_NOTES = null; STATION_NOTES_OK = null;");
  await A("readStationNotes()");
  A("notesUnreadFresh(); markNotesRead('R0001'); rowHtml(__j, 0, 10); stationBoardHtml(); floorNotesHtml(__j);");
  assert.deepStrictEqual(REQ.filter(r => r.method !== "GET").map(r => r.method + " " + r.path), [],
    "nothing the office does to a note is a write of any kind");
  pass("the office's whole half of this feature is reads: not one POST, PATCH or DELETE");

  /* opening a drawer above armed the floor's poll, as it does on a real screen.
     This is a test run and not a screen, so it is disarmed here or node never
     gets to the end of its own event loop. */
  A("closeDrawer(); ALL = []; if (stationPollT) clearTimeout(stationPollT); stationPollT = null;");

  /* ================= 9. over the whole run ================= */
  const workbooky = ALLREQ.filter(r => /\/workbook|\/drive|\/content|Dashboard Log|Dashboard%20Log/.test(r.path));
  assert.deepStrictEqual(workbooky.map(r => r.method + " " + r.path), [],
    "not one request of this whole run went near the workbook");
  assert.deepStrictEqual(ALLREQ.filter(r => r.method === "DELETE").map(r => r.path), [],
    "and nothing was ever deleted, anywhere");
  assert.deepStrictEqual(ALLREQ.filter(r => r.method === "PATCH" && r.path.indexOf(NOTES_ID) >= 0).map(r => r.path), [],
    "and no note was ever patched");
  pass("over the whole run: no workbook, no delete, no note ever edited");

  /* the hard rule, checked against the source of the station files themselves:
     nothing on the floor's side of this feature can write a cell, a fill, a
     sheet or a log row of the workbook, because none of those calls is in it */
  const stationSrc = src("station.js") + src("station-core.js") + src("glass.html");
  ["setFill", "clearFill", "setValues", "appendLog", "saveProgress", "moveJobRow",
   "batchWrite", "/workbook"].forEach(bad =>
    assert.ok(stationSrc.indexOf(bad) < 0, "no " + bad + " in any station file"));
  pass("the station files contain no workbook call of any kind: the grep gate, as a test");

  /* nothing this feature touches carries a customer's phone number or eircode,
     in either direction */
  const bodies = JSON.stringify(ALLREQ.map(r => r.body || {}));
  assert.ok(!/ph3|eircode|Eircode|PHONE/.test(bodies), "nothing personal was ever sent");
  pass("no phone number and no eircode went out on any request of this run");

  console.log("\n" + n + " checks passed");
})().catch(e => { console.error("\nFAILED: " + (e && e.stack || e)); process.exit(1); });
