/* Offline test of the 2026-09-09 work: the Production sheet's row colour code,
   selecting a whole section or everything shown, the John print sheet view
   (Production (2) on its own terms) and the print-notes window behind it.

   The sharp assertions in here are about what is NOT done. The print notes are
   the dashboard's, not the sheet's: they live in the SharePoint list "Dashboard
   print notes" and go there through listUpsert and nothing else. So this file
   watches every request the whole run makes and proves that

     · no workbook sheet is written except the one Dashboard Log line every
       export has always written,
     · nothing PATCHes a range, downloads the file or names Production,
     · every write to "Dashboard print notes" came from listUpsert, and
     · the first save of a session opens ONE consent dialog, not one per note.

   The John print sheet is also the one export in the app that carries a phone
   number, so the file checks that the phone in it is Production (2)'s own, that
   the eircode is in nothing, and that a Default export is still clean.

   Graph is a fake fetch() over an in-memory site, and a token override stands
   in for MSAL - refusing quietly until somebody consents, exactly as
   acquireTokenSilent then acquireTokenPopup do. Nothing leaves the box.
   Every name, county and number below is invented.  Run: node test_john.js  */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, innerWidth: 1280, innerHeight: 800,
                  addEventListener() {}, removeEventListener() {} };
global.performance = { now: () => Date.now() };
global.ExcelJS = ExcelJS;

const SEQ = [];                        // what a download actually did, in order
global.URL.createObjectURL = () => { SEQ.push("URL.createObjectURL"); return "blob:test"; };
global.URL.revokeObjectURL = () => { SEQ.push("URL.revokeObjectURL"); };
global.Blob = function Blob(parts, o) { this.parts = parts; this.type = (o || {}).type || ""; };
let PDFDEF = null;
global.pdfMake = {
  addVirtualFileSystem() {},
  createPdf(def) { PDFDEF = def; SEQ.push("pdfMake.createPdf");
    return { download: name => { SEQ.push("pdf.download " + name); } }; }
};
let CONFIRMS = 0, CONFIRM_ANSWER = true;
global.confirm = () => { CONFIRMS++; return CONFIRM_ANSWER; };

/* Elements remember their html, children, handlers and attributes; an element
   with an id joins REG when it is appended and leaves when it is removed, so
   $("#nhost") is null exactly when the notes window is not on the page. */
const REG = {};
function stubEl(tag, id) {
  let html = "";
  const e = {
    tag: tag || "div", id: id || "", style: {}, dataset: {}, attrs: {}, kids: [],
    textContent: "", value: "", disabled: false, hidden: false, className: "", href: "",
    download: "", src: "", scrollTop: 0, checked: false, indeterminate: false, type: "",
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); if (c.id) REG[c.id] = c; return c; },
    remove() { const drop = x => { if (x.id && REG[x.id] === x) delete REG[x.id]; x.kids.forEach(drop); }; drop(e); },
    contains(x) { return x === e || e.kids.some(k => k.contains && k.contains(x)); },
    setAttribute(k, v) { e.attrs[k] = String(v); }, getAttribute(k) { return e.attrs[k]; },
    addEventListener() {}, removeEventListener() {},
    focus() {}, blur() {}, setSelectionRange() {},
    click() { SEQ.push("anchor.click " + e.download); },
    getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    /* The selectors the page looks for inside a host it has just written. Each
       is answered out of the html that was actually written - one element per
       group, row or note box the render produced - so the wiring under test
       runs on exactly what the page drew, not on a hopeful stand-in. */
    querySelector(sel) { return e.find(sel)[0] || stubEl(); },
    querySelectorAll(sel) { return e.find(sel); },
    find(sel) {
      const s = String(sel);
      const scan = (re, key, field) => {
        const out = [];
        let m;
        while ((m = re.exec(html))) { const d = {}; d[field] = m[1]; out.push(e.child(key + m[1], d)); }
        return out;
      };
      if (s === ".grp") return scan(/class="grp" data-g="(\d+)"/g, "grp:", "g");
      if (s === ".jgrp") return scan(/class="grp jgrp" data-jg="([^"]+)"/g, "jgrp:", "jg");
      if (s === ".jrow[data-id]") return scan(/class="jrow[^"]*" data-id="([^"]+)"/g, "jrow:", "id");
      if (s === "[data-note]") return scan(/data-note="([^"]+)"/g, "note:", "note");
      if (e.parts && e.parts[s]) return [e.parts[s]];
      if (/^\./.test(s) || /^\[/.test(s)) { e.parts = e.parts || {}; e.parts[s] = e.parts[s] || stubEl("div"); return [e.parts[s]]; }
      return [];
    },
    child(key, data) {
      e.made = e.made || {};
      if (!e.made[key]) { const c = stubEl("div"); Object.assign(c.dataset, data || {}); e.made[key] = c; }
      return e.made[key];
    }
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; e.made = {}; e.parts = {}; } });
  e.style.setProperty = (k, v) => { e.style[k] = String(v); };
  return e;
}
const NULLABLE = ["#fabhost", "#fabbtn", "#dhost", "#xhost", "#ahost", "#chost", "#vhost",
                  "#lhost", "#nhost", "#catmenu", "#movemenu", "#alertmenu"];
const EL = {};
const el = sel => {
  const id = String(sel).charAt(0) === "#" ? String(sel).slice(1) : null;
  if (id && REG[id]) return REG[id];
  if (NULLABLE.indexOf(sel) >= 0) return null;
  return EL[sel] || (EL[sel] = stubEl("div", id || ""));
};
global.document = {
  documentElement: stubEl(), body: stubEl(), head: stubEl(), activeElement: null,
  title: "Costello Production",
  createElement: t => stubEl(t),
  querySelector: el, querySelectorAll: () => [],
  addEventListener() {}, removeEventListener() {}
};

/* ---------- the fake Graph ---------- */
const SITE = "costellowindowsie.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";
const NOTE_LIST_ID = "list-dashboard-print-notes";
const LISTS_PATH = "/sites/" + SITE + "/lists";
const G = "https://graph.microsoft.com/v1.0";

let LISTS = [{ id: "list-site-assets", displayName: "Site Assets" },
             { id: NOTE_LIST_ID, displayName: "Dashboard print notes" }];
let ITEMS = [];
let NEXTID = 100;
let FAIL_JOBS = {};                 // job -> refuse its write once, with a 403

const REQ = [], ALLREQ = [];
const LOGSHEET = [["When", "Who", "Job", "What changed", "From", "To"]];
const ok = body => ({ status: 200, body: body });
const note = (job, text, who, at) => ({ id: String(NEXTID++),
  fields: { Title: job, Note: text, By: who || "office@example.test", At: at || "2026-09-08T09:00:00.000Z" } });

function routeList(method, path, body) {
  const rest = path.slice(LISTS_PATH.length);
  if (method === "GET" && rest.indexOf("?$select=id,displayName") === 0) return ok({ value: LISTS });
  const mi = /^\/([^/?]+)\/items(.*)$/.exec(rest);
  if (!mi) return { status: 404, body: { error: { code: "itemNotFound", message: "no route " + path } } };
  const id = mi[1], tail = mi[2];
  if (id !== NOTE_LIST_ID) return { status: 404, body: { error: { code: "itemNotFound" } } };
  if (method === "GET") return ok({ value: ITEMS.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) });
  const wanted = String((body && (body.Title || (body.fields && body.fields.Title))) || "").toUpperCase();
  if (wanted && FAIL_JOBS[wanted]) { delete FAIL_JOBS[wanted]; return { status: 403, body: { error: { code: "accessDenied" } } }; }
  if (method === "POST" && tail === "") {
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
/* just enough workbook for the one Dashboard Log line an export writes */
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

/* MSAL, as far as this feature can tell it apart: a silent request for the
   list permission fails until somebody has consented, and a request that is
   allowed to prompt opens the dialog - which is what POPUPS counts. */
let POPUPS = 0, LIST_GRANTED = true;
CW._setToken((scopes, quiet) => {
  const needsList = (scopes || []).indexOf("Sites.ReadWrite.All") >= 0;
  if (needsList && !LIST_GRANTED) {
    if (quiet) throw new Error("permission needed: consent_required");
    POPUPS++; LIST_GRANTED = true;             // the dialog appeared, and it was accepted
  }
  return "token";
});
CW._setFile({ siteId: SITE, base: "/x/workbook", content: "/x/content", meta: "/x" });
CW._setSession("SESSION-1");
run("checkpoints.js");
global.CP = window.CP;
run("export.js");
run("app.js");

const TOASTS = [];
global.toast = (m, isErr) => TOASTS.push({ m: String(m), err: !!isErr });
global.whoAmI = () => "office@example.test";
const set = src => vm.runInThisContext(src);
const reset = () => { REQ.length = 0; TOASTS.length = 0; SEQ.length = 0; CONFIRMS = 0; };
const paths = () => REQ.map(r => r.method + " " + r.path);
const settle = ms => new Promise(r => setTimeout(r, ms == null ? 20 : ms));

/* ---------- the fixture workbook ----------
   Two sheets: Production, coloured the way the office colours it, and
   Production (2) - John's own sheet, which deliberately disagrees with it. */
const PHONE = "021 555 0742", JPHONE = "065 555 0742", EIR = "ZZ99XY7";

function header(ws, notesCol) {
  ws.getCell(2, 1).value = "COMMENT";
  ws.getCell(2, 2).value = "OFFICE NO";
  ws.getCell(2, 3).value = "JOB NO";
  ws.getCell(2, 4).value = "DATES ON CONTRACT";
  ws.getCell(3, 4).value = "Sold";
  ws.getCell(3, 5).value = "Stamp";
  ws.getCell(3, 6).value = "Ivana";
  ws.getCell(3, 7).value = "Ready to print";
  ws.getCell(3, 8).value = "Sent to floor";
  ws.getCell(2, 9).value = "CUSTOMER";
  ws.getCell(2, 10).value = "PHONE NO";
  ws.getCell(2, 11).value = "AREA";
  ws.getCell(2, 12).value = "EIRCODE";
  ws.getCell(2, 13).value = "QUANTITY";
  ws.getCell(3, 13).value = "Wnd";
  ws.getCell(3, 14).value = "Drs";
  ws.getCell(2, notesCol).value = "NOTES FROM BRENDAN'S OFFICE";
}
/** One job row. `ink` is the font colour of the job-number cell, `custInk` of
    the customer cell; either may be an argb string or { theme, tint }. */
function jobRow(ws, r, id, cust, ink, custInk, extra) {
  const e = extra || {};
  ws.getCell(r, 1).value = e.comment || "";
  ws.getCell(r, 2).value = "OF-" + r;
  ws.getCell(r, 3).value = id;
  ws.getCell(r, 7).value = e.ready || "";
  ws.getCell(r, 9).value = cust;
  ws.getCell(r, 10).value = e.phone || PHONE;
  ws.getCell(r, 11).value = e.area || "Cork";
  ws.getCell(r, 12).value = EIR;
  ws.getCell(r, 13).value = e.wnd == null ? 4 : e.wnd;
  ws.getCell(r, 14).value = e.drs == null ? 1 : e.drs;
  if (e.notesCol) ws.getCell(r, e.notesCol).value = e.brendan || "";
  const paint = (c, col) => { if (col) ws.getCell(r, c).font = { color: (typeof col === "string" ? { argb: "FF" + col } : col) }; };
  paint(3, ink); paint(9, custInk);
  if (e.link) ws.getCell(r, 3).value = { text: id, hyperlink: "https://example.test/" + id };
  const paintFill = (cols, hex) => cols.forEach(c => {
    ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
  });
  /* a whole gold row is how the sheet says "ready to deliver"; yellow on the
     two count cells is how it says "in fabrication" */
  if (e.gold) paintFill([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], "FFE699");
  if (e.fill) paintFill([13, 14], e.fill);
}
function divider(ws, r, text) { ws.getCell(r, 3).value = ""; ws.getCell(r, 4).value = text; }

function fixtureWorkbook() {
  const wb = new ExcelJS.Workbook();

  /* ---- Production: the job model and the row colour code ---- */
  const ws = wb.addWorksheet("Production");
  header(ws, 15);
  const P = (r, id, cust, ink, custInk, extra) =>
    jobRow(ws, r, id, cust, ink, custInk, Object.assign({ notesCol: 15 }, extra || {}));
  P(4, "S2001", "Customer Second", null, null);                                  // block 0
  divider(ws, 5, "CUSTOMERS WON'T TAKE");
  P(6, "W2002", "Customer Wont", null, null);                                    // block 1
  divider(ws, 7, "COLLECT OR SUPPLY ONLY");
  P(8, "C2003", "Customer Collect", null, null);                                 // block 2
  divider(ws, 9, "NOT SENT TO FLOOR");
  P(10, "R3001", "Customer One", "FF0000", null,                                 // block 3: red
    { ready: "2026-02-10", brendan: "Two keys", comment: "Ring first" });
  P(11, "R3002", "Customer Two", { theme: 9 }, null, { gold: 1 });               // a theme green, row done
  divider(ws, 12, "NOT SENT TO FLOOR");
  P(13, "R4001", "Customer Three", "FF3300", null, { fill: "FFFF00" });          // block 4: a near-red
  P(14, "R4002", "Customer Four", "FF3399", null);                               // pink
  P(15, "R4003", "Customer Five", "0070C0", null);                               // blue
  P(16, "R4004", "Customer Six", "000000", "FF3399");                            // black job cell, pink customer
  P(17, "R4005", "Customer Seven", null, null,
    { comment: "URGENT - customer waiting" });                                    // the word, not the colour
  P(18, "R4006", "Customer Eight", "808080", null);                              // grey: no code at all
  P(19, "R4007", "Customer Nine", { theme: 10 }, null, { link: 1 });             // a hyperlink: not a flag

  /* ---- Production (2): John's own sheet ----
     Fewer rows, its own order, its own sections, its own colours - and, for
     R3001, its own phone, area, ready date and note. Its notes column is way
     out at BY, past a run of hidden columns, exactly as the real one is. */
  const w2 = wb.addWorksheet("Production (2)");
  header(w2, 77);
  for (let c = 16; c <= 76; c++) w2.getColumn(c).hidden = true;
  const J = (r, id, cust, ink, extra) =>
    jobRow(w2, r, id, cust, ink, null, Object.assign({ notesCol: 77 }, extra || {}));
  divider(w2, 4, "COLLECT OR SUPPLY ONLY");
  J(5, "C2003", "Customer Collect", null, { ready: "2026-05-04", brendan: "Trade counter" });
  divider(w2, 6, "NOT SENT TO FLOOR");
  J(7, "R4002", "Customer Four", "FF3399", { ready: "2026-03-20", gold: 1, brendan: "Pink on John's sheet" });
  J(8, "R3001", "Customer One", "FF0000",
    { ready: "2026-02-14", phone: JPHONE, area: "West Clare", brendan: "Two keys and the side gate",
      fill: "FFFF00", wnd: 9, drs: 2 });
  J(9, "R4003", "Customer Five", "00B0F0", { ready: "2026-04-01" });
  J(10, "R4006", "Customer Eight", null, { ready: "" });
  return wb;
}

function useJobs(list, names, john) {
  global.__jobs = list; global.__names = names; global.__john = john || [];
  set("ALL = __jobs; BLOCKNAMES = __names; ALL.blockNames = __names; JOHNROWS = __john; CHANGES = []; " +
      "state.picked = {}; state.q = ''; state.cat = null; state.sheet = null; state.board = null; " +
      "state.sort = 'id'; state.desc = false; state.view = 'flat'; state.hidden = {}; state.gq = {}; " +
      "state.collapsed = {}; state.sel = null; ALERTS = {}; VIEWS = {};");
}

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };
  const WB = fixtureWorkbook();

  /* ---- 1. the row colour code, read off a real workbook ---- */
  const jobs = parseWorkbook(WB);
  const by = {};
  jobs.forEach(j => { by[j.id] = j; });
  assert.strictEqual(by.R3001.flag, "urgent", "FF0000 is the red the office writes urgent in");
  assert.strictEqual(by.R3001.flagHex, "FF0000", "and the hex that said so is kept for the exports");
  assert.strictEqual(by.R4001.flag, "urgent", "FF3300 is a near-red: hue ranges, not exact matches");
  assert.strictEqual(by.R3002.flag, "booked", "a THEME colour resolves through the palette the fills use");
  assert.strictEqual(by.R3002.flagHex, "70AD47");
  assert.strictEqual(by.R4002.flag, "trade", "FF3399 is the pink for a trade order");
  assert.strictEqual(by.R4003.flag, "hold", "0070C0 is the blue for on hold");
  assert.strictEqual(by.R4004.flag, "trade",
    "a black job-number cell falls back to the customer cell, which is where this one is coloured");
  assert.strictEqual(by.R4006.flag, "", "grey text is ordinary text and says nothing");
  assert.strictEqual(by.S2001.flag, "", "and so is a cell with no font colour at all");
  pass("the colour code is read from the row's text: red, green, pink, blue, and nothing for grey");

  assert.strictEqual(by.R4007.flag, "", "a hyperlinked cell is Excel's own blue and is never a flag");
  assert.strictEqual(fontOf({ hyperlink: "https://example.test/x", font: { color: { argb: "FF0070C0" } } }), "",
    "a link says nothing about the job, whatever colour it happens to be");
  assert.strictEqual(fontOf({ font: { color: { theme: 10 } } }), "", "theme 10 is the hyperlink colour");
  assert.strictEqual(fontOf({ font: { color: { theme: 11 } } }), "", "theme 11 the followed-link colour");
  assert.strictEqual(flagOf(fontOf({ font: { color: { theme: 8 } } })), "hold",
    "while a real theme blue is still read, so this is not a blanket refusal");
  pass("a hyperlink, and the two hyperlink theme colours, are never read as a colour code");

  assert.strictEqual(blackInk(""), true, "no ink at all is 'the sheet said nothing'");
  assert.strictEqual(blackInk("000000"), true);
  assert.strictEqual(blackInk("808080"), false, "grey is somebody's decision, not black");
  assert.strictEqual(blackInk("FFC000"), false, "and so is orange");
  assert.strictEqual(by.R4006.flag, "", "so a grey job number never asks the customer cell instead");
  pass("the customer-cell fallback happens only when the job number itself is plain black");

  assert.strictEqual(by.R3001.urg, 1, "a red row is urgent even with no word in the comment");
  assert.strictEqual(by.R4005.urg, 1, "and the word still is, with no colour at all");
  assert.strictEqual(by.R4005.flag, "", "without pretending the sheet coloured that row");
  assert.strictEqual(by.R4003.urg, 0, "a blue row is not urgent");
  pass("urgency is the word in a comment OR the sheet's red row text, and neither invents the other");

  assert.strictEqual(flagOf("C00000"), "urgent", "the darker red the office also uses");
  assert.strictEqual(flagOf("00B050"), "booked");
  assert.strictEqual(flagOf("92D050"), "booked", "and the lighter green");
  assert.strictEqual(flagOf("FF00FF"), "trade");
  assert.strictEqual(flagOf("00B0F0"), "hold");
  assert.strictEqual(flagOf("000000"), "", "black");
  assert.strictEqual(flagOf("FFFFFF"), "", "white");
  assert.strictEqual(flagOf("595959"), "", "grey");
  assert.strictEqual(flagOf("FFC000"), "", "orange is not one of the four and is never guessed at");
  assert.strictEqual(flagOf(""), "", "and nothing at all is nothing at all");
  pass("the hue ranges cover the colours the sheet really uses and refuse the ones it does not");

  /* ---- 2. Production (2), read on its own terms ---- */
  const JOHN = parseJohnSheet(WB);
  assert.deepStrictEqual(JOHN.map(r => r.id), ["C2003", "R4002", "R3001", "R4003", "R4006"],
    "that sheet's own rows, in that sheet's own order - not the Production model's");
  assert.deepStrictEqual(JOHN.sections, ["Collect & supply only", "Ready to fit"],
    "and its own sections, from its own divider rows - not the Production sheet's five");
  assert.deepStrictEqual(JOHN.map(r => r.section),
    ["Collect & supply only", "Ready to fit", "Ready to fit", "Ready to fit", "Ready to fit"]);
  const jrow = id => JOHN.find(r => r.id === id);
  assert.strictEqual(jrow("R3001").phone, JPHONE, "the phone is that sheet's, not the Production model's");
  assert.strictEqual(jrow("R3001").area, "West Clare", "and so is the area");
  assert.strictEqual(jrow("R3001").ready, "2026-02-14", "and the ready date");
  assert.strictEqual(jrow("R3001").notes, "Two keys and the side gate",
    "the notes column is found by its header, seventy columns out past the hidden ones");
  assert.strictEqual(jrow("R3001").wnd, 9);
  assert.strictEqual(jrow("R3001").drs, 2);
  assert.strictEqual(jrow("R3001").fillHex, "FFFF00", "the row's own fill");
  assert.strictEqual(jrow("R3001").inkHex, "FF0000", "and the row's own ink");
  assert.strictEqual(jrow("R4002").fillHex, "FFE699", "a gold row on John's sheet is gold here");
  assert.strictEqual(jrow("R4002").inkHex, "FF3399");
  assert.strictEqual(jrow("R4003").inkHex, "00B0F0");
  assert.strictEqual(jrow("R4006").inkHex, "", "an uncoloured row carries no ink");
  assert.strictEqual(jrow("R4006").fillHex, "", "and no fill");
  assert.strictEqual(jrow("C2003").seq, 0, "seq is that sheet's own row order");
  assert.strictEqual(jrow("R4006").seq, 4);
  assert.strictEqual(JSON.stringify(JOHN).indexOf(EIR), -1, "and no eircode is read into it at all");
  pass("parseJohnSheet reads Production (2) on its own terms: its rows, order, sections, colours");

  assert.strictEqual(by.R3001.ph, PHONE, "the Production model still holds Production's own phone");
  assert.strictEqual(by.R3001.area, "Cork", "and Production's own area");
  assert.strictEqual(by.R3001.cust, "Customer One");
  assert.deepStrictEqual(parseJohnSheet(new ExcelJS.Workbook()), [],
    "a workbook with no such sheet answers an empty list, never null");
  pass("nothing from Production (2) is merged into the job model, and its absence is not an error");

  /* ---- 3. selecting a whole section, and everything shown ---- */
  const SECTIONS = jobs.blockNames;
  const live5 = jobs.filter(j => j.cat !== "past");
  useJobs(live5, SECTIONS, JOHN);
  assert.strictEqual(live().length, 12, "twelve jobs on the sheet");

  renderChips();
  const chipText = () => el("#chips").kids.map(k => (k.textContent || "") +
    k.kids.map(x => x.textContent || "").join("")).join(" | ");
  assert.strictEqual(chipText().indexOf("Select all shown"), -1,
    "with every job on screen there is nothing to select that is not already all of it");

  set("state.cat = 'urgent';");
  renderChips();
  assert.ok(chipText().indexOf("Select all shown (3)") >= 0,
    "a tile that narrows the list puts the box in the toolbar with the count on it");
  const selBox = el("#chips").kids.filter(k => k.className === "selall")
    .map(k => k.kids[0]).filter(Boolean)[0];
  assert.ok(selBox, "and it is a real tick box");
  assert.strictEqual(selBox.checked, false);
  selBox.checked = true;
  selBox.onchange();
  assert.deepStrictEqual(Object.keys(state.picked).sort(), ["R3001", "R4001", "R4005"],
    "ticking it picks every job the filter is showing, and nothing else");
  renderChips();
  const selBox2 = el("#chips").kids.filter(k => k.className === "selall").map(k => k.kids[0])[0];
  assert.strictEqual(selBox2.checked, true, "and it comes back ticked");
  selBox2.checked = false;
  selBox2.onchange();
  assert.deepStrictEqual(Object.keys(state.picked), [], "unticking clears exactly those again");
  pass("flat list: “Select all shown (n)” appears only when a filter narrows the list, and picks it");

  set("state.picked = { R3001: 1 };");
  renderChips();
  const partial = el("#chips").kids.filter(k => k.className === "selall").map(k => k.kids[0])[0];
  assert.strictEqual(partial.checked, false);
  assert.strictEqual(partial.indeterminate, true, "some but not all: the box is indeterminate");
  pass("a part-picked list shows the box's third state rather than lying either way");

  set("state.cat = null; state.picked = {}; state.view = 'Abin';");
  renderRows();
  const rowsHtml = el("#rows").innerHTML;
  assert.ok(rowsHtml.indexOf("Select all 7") >= 0, "In production shows its seven");
  assert.ok(rowsHtml.indexOf("Select all 2") >= 0, "Ready to fit shows its two");
  assert.ok(rowsHtml.indexOf("Select all 1") >= 0, "and a section of one says one");
  const grp = g => el("#rows").querySelectorAll(".grp").filter(x => x.dataset.g === String(g))[0];
  const gall = g => grp(g).querySelector(".gall");
  assert.strictEqual(gall(4).checked, false);
  gall(4).checked = true;
  gall(4).onchange();
  assert.deepStrictEqual(Object.keys(state.picked).sort(),
    ["R4001", "R4002", "R4003", "R4004", "R4005", "R4006", "R4007"],
    "ticking a section picks every job in it and touches no other section");
  renderRows();
  assert.ok(el("#rows").innerHTML.indexOf('class="pick gall" checked') >= 0, "the section's box is ticked now");
  gall(3).checked = true;
  gall(3).onchange();
  assert.strictEqual(Object.keys(state.picked).length, 9, "a second section adds to the first");
  gall(4).checked = false;
  gall(4).onchange();
  assert.deepStrictEqual(Object.keys(state.picked).sort(), ["R3001", "R3002"],
    "and unticking one takes only its own jobs away");
  pass("grouped view: a tick box per section that picks and unpicks exactly that section");

  set("state.picked = {}; state.gq = { 'Abin|4': 'R400' };");
  renderRows();
  assert.ok(el("#rows").innerHTML.indexOf("Select all 7") >= 0,
    "the search matches seven of them, so the box still says seven");
  set("state.gq = { 'Abin|4': 'Customer Four' };");
  renderRows();
  assert.ok(el("#rows").innerHTML.indexOf("Select all 1") >= 0, "a narrower search, a smaller box");
  gall(4).checked = true;
  gall(4).onchange();
  assert.deepStrictEqual(Object.keys(state.picked), ["R4002"],
    "and it picks what the group is showing, not what it holds");
  pass("a group's own search box narrows what “select all” means, exactly as it narrows the rows");

  /* a folded section keeps its box: the owner asked for that expressly - it is
     how you take a whole section you are not reading */
  set("state.picked = {}; state.gq = {}; state.collapsed = { 'Abin|4': 1 };");
  renderRows();
  assert.ok(el("#rows").innerHTML.indexOf("Select all 7") >= 0,
    "a collapsed section still offers its tick box, still with the count on it");
  assert.ok(el("#rows").innerHTML.indexOf('<button class="gtog">▸') >= 0, "with its rows folded away");
  gall(4).checked = true;
  gall(4).onchange();
  assert.strictEqual(Object.keys(state.picked).length, 7, "and it still picks all of them");
  set("state.collapsed = {}; state.picked = {};");
  pass("a collapsed section keeps its “Select all n”, count and all");

  let draws = 0;
  set("state.picked = {}; state.gq = {}; state.view = 'flat'; state.cat = null;");
  global.__realRenderAll = renderAll;
  global.__countRenderAll = function () { draws++; return global.__realRenderAll(); };
  set("renderAll = global.__countRenderAll;");
  const many = live().map(j => j.id);
  pickMany(many, true);
  assert.strictEqual(Object.keys(state.picked).length, 12, "all twelve are ticked");
  assert.strictEqual(draws, 1, "in one draw, not twelve");
  set("renderAll = global.__realRenderAll;");
  pickMany(many, false);
  assert.deepStrictEqual(Object.keys(state.picked), []);
  pass("picking a whole section is one pass and one render, whatever the section holds");

  assert.ok(rowHtml(by.R4002, 0, 10).indexOf("Trade order") >= 0, "the colour chip is on the list row");
  assert.ok(rowHtml(by.R4002, 0, 10).indexOf('class="stn">Production<') <
            rowHtml(by.R4002, 0, 10).indexOf("Trade order"),
    "and it sits in the last cell with the other badges, where overflow is already handled");
  assert.strictEqual(rowHtml(by.R4006, 0, 10).indexOf("flagchip"), -1, "an uncoloured row gets no chip");
  pass("the colour chip is a word, in the badges cell, and only when the sheet coloured the row");

  /* ---- 4. the John print sheet view ---- */
  set("state.board = 'john'; state.picked = {}; state.q = '';");
  renderRows();
  const jhtml = () => el("#rows").innerHTML;
  assert.ok(jhtml().indexOf("Notes from Brendan") >= 0, "the view's own eight columns");
  assert.ok(jhtml().indexOf("Phone no") >= 0, "phone included: this is John's paper");
  ["C2003", "R4002", "R3001", "R4003", "R4006"].forEach(id =>
    assert.ok(jhtml().indexOf('data-id="' + id + '"') >= 0, "the view is missing " + id));
  assert.strictEqual(jhtml().indexOf('data-id="R3002"'), -1,
    "and it shows nothing that is not on Production (2), whatever the Production model says");
  assert.ok(jhtml().indexOf(JPHONE) >= 0, "with that sheet's own phone number");
  assert.ok(jhtml().indexOf("West Clare") >= 0, "and its own area");
  assert.ok(jhtml().indexOf("Two keys and the side gate") >= 0, "and its own note");
  assert.strictEqual(jhtml().indexOf(EIR), -1, "never an eircode");
  assert.ok(jhtml().indexOf("Collect &amp; supply only") >= 0 && jhtml().indexOf("Ready to fit") >= 0,
    "its own two sections, as divider rows");
  assert.ok(jhtml().indexOf("background:#FFFF00") >= 0, "each row wearing its own fill");
  assert.ok(jhtml().indexOf("Trade order") >= 0, "and its colour word, so colour is never the only signal");
  assert.strictEqual(el("#count").textContent, "Showing all 5 rows on John’s sheet");
  pass("the John print sheet view draws Production (2) on its own terms, colours, sections and all");

  /* the sections are keyed by position, so the test looks them up the same way
     the page does - by where they are in Production (2)'s own order */
  const jall = i => el("#rows").querySelectorAll(".jgrp")[i].querySelector(".gall");
  const COLLECT = 0, READY = 1;
  assert.ok(jhtml().indexOf("Select all 4") >= 0, "Ready to fit offers its four");
  jall(READY).checked = true;
  jall(READY).onchange();
  assert.deepStrictEqual(Object.keys(state.picked).sort(), ["R3001", "R4002", "R4003", "R4006"],
    "ticking a section of John's sheet picks exactly that section");
  jall(COLLECT).checked = true;
  jall(COLLECT).onchange();
  assert.strictEqual(Object.keys(state.picked).length, 5, "and the other section adds to it");
  jall(READY).checked = false;
  jall(READY).onchange();
  assert.deepStrictEqual(Object.keys(state.picked), ["C2003"], "unticking takes only its own away");
  const jr1 = el("#rows").querySelectorAll(".jrow[data-id]").filter(x => x.dataset.id === "R4003")[0];
  jr1.querySelector(".pick").checked = true;
  jr1.querySelector(".pick").onchange();
  assert.deepStrictEqual(Object.keys(state.picked).sort(), ["C2003", "R4003"], "and a single row ticks on its own");
  pass("the view's tick boxes and per-section “Select all n” feed the same state.picked as the master list");

  set("state.q = 'clare'; state.picked = {};");
  renderRows();
  assert.ok(jhtml().indexOf('data-id="R3001"') >= 0, "the search box narrows the view");
  assert.strictEqual(jhtml().indexOf('data-id="R4003"'), -1, "to the rows that match");
  assert.ok(jhtml().indexOf("Select all 1") >= 0, "and “Select all” follows the search");
  assert.strictEqual(el("#count").textContent, "Showing 1 of 5 rows on John’s sheet");
  set("state.q = '';");
  renderRows();
  pass("the header search box narrows the John print sheet view, count and select-all together");

  /* the view sits in the same slot as the floor's board but has nothing to do
     with the floor: it must not pin the ten-second poll on */
  assert.strictEqual(stationWatching(), false,
    "John's own sheet is no reason to poll the floor six times a minute");
  set("state.board = 'glass';");
  assert.strictEqual(stationWatching(), true, "the floor's own board still is");
  set("state.board = 'john';");
  pass("the John print sheet view never speeds up the floor poll: nothing on it comes from the floor");

  /* ---- 5. the Export window, from the view and from the master ---- */
  set("XSTATE = null; state.picked = { R3001: 1, R4002: 1 };");
  renderExportWindow();
  assert.strictEqual(XSTATE.template, "john", "opened from the view, the window is already on John's template");
  const xhtml = () => el("#xhost").innerHTML;
  assert.ok(xhtml().indexOf("Continue to notes") >= 0, "and the button says what happens next");
  assert.strictEqual(xhtml().indexOf(">Download<"), -1, "there is no Download button on this template");
  assert.strictEqual(xhtml().indexOf('data-xtog="fields"'), -1, "no field picker: the layout is fixed");
  assert.strictEqual(xhtml().indexOf('data-xset="layout"'), -1, "and no card/table choice either");
  assert.strictEqual(xhtml().indexOf('data-xset="f.scope" data-xval="sections"'), -1,
    "and no Sections scope: the two views are the filter");
  assert.ok(xhtml().indexOf("All of John’s sheet") >= 0, "the scope is worded for the view it was opened from");
  set("XSTATE.f.scope = 'view';");
  assert.deepStrictEqual(xpJohnIds(), ["C2003", "R4002", "R3001", "R4003", "R4006"],
    "and “all of John's sheet” means that sheet's rows, in its order");
  set("XSTATE.f.scope = 'ticked';");
  assert.deepStrictEqual(xpJohnIds(), ["R4002", "R3001"], "ticked means the ticked rows of that sheet");

  /* preselected, not locked: the Template row still works from in there */
  set("XSTATE.f.scope = 'view'; xpSet('template', 'default'); renderExportWindow();");
  assert.strictEqual(XSTATE.template, "default",
    "the John view preselects the template on the way in and then leaves the choice alone");
  assert.ok(xhtml().indexOf('data-xtog="fields"') >= 0, "so the Default field picker comes back");
  set("XSTATE.template = 'john'; renderExportWindow();");

  set("state.board = null; state.view = 'flat'; XSTATE = null; renderExportWindow();");
  assert.strictEqual(XSTATE.template, "default", "from the master list the window opens on Default, as always");
  assert.ok(xhtml().indexOf('data-xtog="fields"') >= 0, "with its field picker back");
  assert.ok(xhtml().indexOf("No phone numbers and no eircodes are ever included") >= 0);
  set("XSTATE.template = 'john'; XSTATE.f.scope = 'ticked'; renderExportWindow();");
  assert.deepStrictEqual(xpJohnIds().sort(), ["R3001", "R4002"],
    "and John's template can be chosen from the master, on the jobs ticked there");
  pass("Export opens on John's template from the view, on Default from the master, and both print it");

  /* ---- 6. the notes window ---- */
  ITEMS = [note("R3001", "Leave at the side gate"), note("R4002", "Ring the site foreman")];
  set("state.picked = { R3001: 1, R4001: 1, R4002: 1 }; XSTATE.f.scope = 'ticked'; XSTATE.format = 'xlsx';");
  reset();
  await openNotesWindow();
  await settle();
  assert.ok(el("#nhost"), "the notes window is open");
  const nhtml = () => el("#nhost").innerHTML;
  assert.ok(nhtml().indexOf("Print notes — John print sheet") >= 0, "named for what it prints");
  assert.ok(nhtml().indexOf('data-note="R3001"') >= 0 && nhtml().indexOf('data-note="R4001"') >= 0 &&
            nhtml().indexOf('data-note="R4002"') >= 0, "one row per job in the print");
  assert.ok(nhtml().indexOf("Leave at the side gate") >= 0, "pre-filled with the note the list already holds");
  assert.ok(nhtml().indexOf("Two keys and the side gate") >= 0,
    "beside Production (2)'s own note, which is shown but cannot be typed into");
  assert.ok(nhtml().indexOf("Production comment: Ring first") >= 0,
    "with the Production comment underneath it, labelled, because the two differ");
  assert.ok(nhtml().indexOf("not on John’s sheet") >= 0, "R4001 is not on that sheet and the row says so");
  assert.ok(nhtml().indexOf('maxlength="2000"') >= 0, "a note has a length somebody cannot get past");
  assert.ok(nhtml().indexOf("Urgent") >= 0, "the colour chip is on the row too");
  assert.strictEqual(paths().filter(p => p.indexOf("/x/") >= 0).length, 0,
    "opening the window read a list and no workbook at all");
  assert.ok(paths().some(p => p.indexOf("expand=fields(select=Title,Note,By,At)") > 0),
    "asking the list for its four columns by name");
  pass("the notes window: one row per job, pre-filled from the list and from John's own sheet");

  /* the scrim asks before throwing typed notes away; Cancel does not */
  const boxes0 = {};
  el("#nhost").querySelectorAll("[data-note]").forEach(b => { boxes0[b.dataset.note] = b; });
  boxes0.R4001.value = "typed but not printed";
  boxes0.R4001.oninput();
  CONFIRM_ANSWER = false;
  assert.strictEqual(el("#nscrim").onclick(), false, "the scrim asks first");
  assert.strictEqual(CONFIRMS, 1);
  assert.ok(el("#nhost"), "and a no keeps the window and the typing");
  CONFIRM_ANSWER = true;
  boxes0.R4001.value = "";
  boxes0.R4001.oninput();
  CONFIRMS = 0;
  el("#nscrim").onclick();
  assert.strictEqual(CONFIRMS, 0, "with nothing typed it does not ask at all");
  assert.strictEqual(el("#nhost"), null);
  pass("clicking the scrim asks before discarding typed notes, and does not nag when there are none");

  /* ---- 7. one consent dialog, not one per note ---- */
  LIST_GRANTED = false; POPUPS = 0;
  set("PRINT_NOTE_OK = null; PRINTNOTES = {};");
  ITEMS = [];
  reset();
  await openNotesWindow();
  await settle();
  const boxes = {};
  el("#nhost").querySelectorAll("[data-note]").forEach(b => { boxes[b.dataset.note] = b; });
  ["R3001", "R4001", "R4002"].forEach((id, i) => {
    boxes[id].value = "Note number " + (i + 1);
    boxes[id].oninput();
  });
  const name = await johnPrint();
  await settle();
  assert.strictEqual(POPUPS, 1,
    "three changed notes, ONE consent dialog - the click asks once before any of them are written");
  assert.strictEqual(name, "John print sheet " + xpIsoDate(new Date()) + ".xlsx", "the file is named for the day");
  assert.ok(SEQ.some(x => x.indexOf("anchor.click John print sheet") === 0), "and it really was handed over");
  assert.strictEqual(ITEMS.length, 3, "one list item per note");
  assert.deepStrictEqual(ITEMS.map(x => x.fields.Note).sort(),
    ["Note number 1", "Note number 2", "Note number 3"]);
  assert.strictEqual(el("#nhost"), null, "a clean print closes the window");
  pass("the first save of a session opens one consent dialog for the whole print, not one per note");

  /* ---- 8. only what changed is written ---- */
  reset();
  await openNotesWindow();
  await settle();
  const boxes2 = {};
  el("#nhost").querySelectorAll("[data-note]").forEach(b => { boxes2[b.dataset.note] = b; });
  boxes2.R4002.value = "Changed this one only";
  boxes2.R4002.oninput();
  await johnPrint();
  await settle();
  const wrote = REQ.filter(r => r.method !== "GET" && r.path.indexOf(LISTS_PATH) === 0);
  assert.strictEqual(wrote.length, 1, "one write, for the one note that changed");
  assert.strictEqual(wrote[0].method, "PATCH", "and it PATCHed the item that was already there");
  assert.strictEqual(ITEMS.length, 3, "no new item for the two that were typed back the same");
  assert.strictEqual(POPUPS, 1, "and no second consent dialog, ever");
  const logs = REQ.filter(r => r.path.indexOf("/x/") === 0);
  assert.ok(logs.length > 0, "the export wrote its log line");
  const line = LOGSHEET[LOGSHEET.length - 1];
  assert.strictEqual(line[3], "Export");
  assert.strictEqual(line[4], "Excel · 3 jobs · John print sheet · with phone numbers",
    "the log line says in words that this file has phone numbers in it");
  assert.strictEqual(line[5], "John print sheet · ticked jobs");
  pass("only the notes that changed are written, and one Dashboard Log line names the template");

  /* ---- 9. the print is the jobs the window showed ---- */
  reset();
  await openNotesWindow();
  await settle();
  assert.deepStrictEqual(NSTATE.ids.slice().sort(), ["R3001", "R4001", "R4002"]);
  /* a poll lands between opening the window and pressing Print, and the whole
     job list is replaced under it */
  useJobs(live5.filter(j => j.id === "R4003"), SECTIONS, JOHN);
  set("state.picked = { R4003: 1 }; XSTATE.f.scope = 'ticked'; XSTATE.format = 'pdf';");
  await johnPrint();
  await settle();
  const printed = PDFDEF.content[1].table.body
    .filter(r => r.length > 1).map(r => r[0].text).filter(t => /^[A-Z]\d/.test(String(t)));
  assert.deepStrictEqual(printed.sort(), ["R3001", "R4001", "R4002"],
    "the file is the jobs the window showed, not whatever the list holds now");
  assert.strictEqual(printed.indexOf("R4003"), -1);
  useJobs(live5, SECTIONS, JOHN);
  set("XSTATE.f.scope = 'ticked'; XSTATE.format = 'xlsx';");
  pass("Print prints exactly the jobs the notes window listed, whatever changed underneath it");

  /* ---- 10. a note that will not save still prints, and its mark clears ---- */
  reset();
  set("state.picked = { R3001: 1, R4001: 1, R4002: 1 };");
  FAIL_JOBS.R4001 = 1;
  await openNotesWindow();
  await settle();
  const boxes3 = {};
  el("#nhost").querySelectorAll("[data-note]").forEach(b => { boxes3[b.dataset.note] = b; });
  boxes3.R4001.value = "This one will not save";
  boxes3.R4001.oninput();
  const name3 = await johnPrint();
  await settle();
  assert.ok(name3, "the file was still built");
  assert.ok(SEQ.some(x => x.indexOf("anchor.click John print sheet") === 0), "and handed over");
  assert.ok(el("#nhost"), "the window stays open so the row that failed can be seen");
  assert.ok(el("#nhost").innerHTML.indexOf("not saved") >= 0, "and it says so on that row");
  assert.ok(TOASTS.some(t => /could not be saved/.test(t.m)), "with a toast that does not pretend otherwise");
  /* press Print again with nothing changed: last time's mark is not this
     print's news, so it goes */
  reset();
  const again = await johnPrint();
  await settle();
  assert.ok(again, "the second print goes too");
  assert.strictEqual(el("#nhost"), null,
    "and the window closes, because the failure mark was cleared before this pass rather than left to rot");
  pass("a note that will not save shows on its row, does not stop the print, and does not stick");

  /* ---- 11. the file itself: Production (2)'s data, phone yes, eircode no ---- */
  set("state.picked = {}; XSTATE.f.scope = 'view'; state.board = 'john';");
  const rows = exportJohnRows(xpJohnIds(), { R3001: "Leave at the side gate" }, xpCtxNow());
  assert.deepStrictEqual(rows.filter(r => r.kind === "section").map(r => r.section),
    ["Collect & supply only", "Ready to fit"], "Production (2)'s own sections, as divider rows");
  const john = rows.filter(r => r.kind === "job");
  assert.deepStrictEqual(john.map(r => r.id), ["C2003", "R4002", "R3001", "R4003", "R4006"],
    "and its own row order");
  const jj = id => john.find(r => r.id === id);
  assert.strictEqual(jj("R3001").phone, JPHONE, "the phone printed is that sheet's");
  assert.strictEqual(jj("R3001").area, "West Clare", "and the area");
  assert.strictEqual(jj("R3001").ready, "14-Feb", "and the ready date, as dd-MMM");
  assert.strictEqual(jj("R3001").notes, "Two keys and the side gate · Leave at the side gate",
    "and its note, with the dashboard's print note after it");
  assert.strictEqual(jj("R3001").fill, "#FFFF00");
  assert.strictEqual(jj("R3001").colour, "#FF0000");
  assert.strictEqual(jj("R3001").flag, "urgent");
  assert.strictEqual(jj("R4006").fill, "", "an uncoloured row prints uncoloured");
  assert.strictEqual(jj("R4006").colour, "");
  assert.ok(john.every(r => !r.missing));
  assert.strictEqual(JSON.stringify(rows).indexOf(EIR), -1, "no eircode reaches the file");
  assert.ok(JSON.stringify(rows).indexOf(JPHONE) >= 0, "and the phone number really is in it");

  const mixed = exportJohnRows(["R3001", "R4001"], {}, xpCtxNow());
  const gone = mixed.filter(r => r.kind === "job").find(r => r.id === "R4001");
  assert.ok(gone, "a job that is not on Production (2) is still printed");
  assert.strictEqual(gone.missing, true);
  assert.strictEqual(gone.fill, "", "with no fill");
  assert.strictEqual(gone.colour, "", "and no colour: that sheet said nothing about it");
  assert.ok(gone.notes.indexOf("not on John's sheet") >= 0, "and its Notes say so");
  assert.strictEqual(gone.cust, "Customer Three", "the rest of it comes from the Production model");
  /* a job with no section at all - off the sheet's blocks entirely - must not
     land under "section NaN" */
  const orphan = exportJohnRows(["R0000"], {}, Object.assign(xpCtxNow(),
    { all: [{ id: "R0000", cust: "Customer Nobody", blk: -1, notes: [], dates: {} }] }));
  assert.deepStrictEqual(orphan.map(r => r.kind === "section" ? "== " + r.section : r.id),
    ["== No section", "R0000"], "a job with no section of its own gets a plain heading, never NaN");
  pass("the file is Production (2)'s data, and a job it does not have prints uncoloured and marked");

  const wb2 = buildJohnWorkbook(rows, { who: "office@example.test", when: new Date() });
  const buf = await wb2.xlsx.writeBuffer();
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(buf);
  const ws2 = back.getWorksheet("John print sheet");
  const all = [];
  ws2.eachRow({ includeEmpty: false }, r => r.eachCell({ includeEmpty: false }, c => all.push(String(c.value))));
  assert.strictEqual(all.join(" | ").indexOf(EIR), -1, "the John workbook has no eircode anywhere in it");
  assert.ok(all.some(x => x === JPHONE), "and it does have that sheet's phone number");
  assert.ok(all.some(x => x === "Phone no"), "under a column that says so");
  pass("the John workbook carries Production (2)'s phone number and no eircode at all");

  const defRows = exportRows(live(), exportAllFields(), xpCtxNow());
  const defFlat = JSON.stringify(defRows);
  assert.strictEqual(defFlat.indexOf(EIR), -1, "no eircode in a Default row");
  assert.strictEqual(defFlat.indexOf(PHONE), -1, "and no phone number either");
  assert.strictEqual(defFlat.indexOf(JPHONE), -1, "not Production (2)'s number either");
  assert.strictEqual(defFlat.indexOf("0742"), -1, "not even the last digits of one");
  const defCols = exportColumns(defRows, exportAllFields()).map(c => c.name.toLowerCase());
  assert.strictEqual(defCols.filter(x => x.indexOf("phone") >= 0 || x.indexOf("eircode") >= 0).length, 0,
    "and no Default column is even called either of those things");
  assert.ok(defCols.indexOf("flag") >= 0, "the colour code is a Default column, though");
  assert.strictEqual(defRows.find(r => r.id === "R4002").flag, "trade",
    "carrying the word the sheet's colour stands for");
  pass("the Default template still carries no phone number and no eircode of any kind");

  /* ---- 12. the drawer's read-only line ---- */
  set("state.sel = 'R3001'; state.board = null;");
  const dhost = stubEl("div", "dhost");
  REG.dhost = dhost;
  renderDrawer();
  assert.ok(dhost.innerHTML.indexOf("Print notes (John print sheet)") >= 0, "the drawer has a Print notes line");
  assert.ok(dhost.innerHTML.indexOf("Note number 1") >= 0, "showing the note");
  assert.ok(dhost.innerHTML.indexOf("office") >= 0, "and who put it there");
  assert.strictEqual(dhost.innerHTML.indexOf('<textarea class="nbox"'), -1,
    "read only: there is no box to type in from the drawer");
  assert.ok(dhost.innerHTML.indexOf("never in the Excel file") >= 0, "and it says where the note lives");
  set("PRINT_NOTE_OK = null;");
  renderDrawer();
  assert.ok(dhost.innerHTML.indexOf("read when the print-notes window opens") >= 0,
    "and before the list has been read it says exactly when it will be");
  set("PRINT_NOTE_OK = true;");
  delete REG.dhost;
  set("state.sel = null;");
  pass("the drawer shows the print note under Comments, read only, with who and when");

  /* ---- 13. the sweep: nothing here writes a workbook sheet ---- */
  const workbookReqs = ALLREQ.filter(r => r.path.indexOf("/x/") === 0);
  assert.ok(workbookReqs.length > 0, "the run really did talk to the workbook, or this proves nothing");
  assert.deepStrictEqual(workbookReqs.filter(r => r.path.indexOf("Dashboard%20Log") < 0 &&
                                                 r.path.indexOf("Dashboard Log") < 0 &&
                                                 r.path !== "/x/workbook/worksheets")
                                     .map(r => r.method + " " + r.path), [],
    "every workbook request of the whole run was the Dashboard Log sheet");
  assert.deepStrictEqual(ALLREQ.filter(r => /Production|\/content|\/versions|createSession|\/drive\//.test(r.path))
                               .map(r => r.method + " " + r.path), [],
    "nothing named Production, downloaded the file, or opened a session");
  const noteWrites = ALLREQ.filter(r => r.method !== "GET" && r.path.indexOf(LISTS_PATH) === 0);
  assert.ok(noteWrites.length > 0);
  assert.ok(noteWrites.every(r => r.path.indexOf("/" + NOTE_LIST_ID + "/items") > 0),
    "every list write went to “Dashboard print notes” and to no other list");
  assert.ok(noteWrites.every(r => r.method === "POST" || r.method === "PATCH"),
    "and every one of them was listUpsert's own POST or PATCH - nothing here deletes");
  const src = fs.readFileSync(__dirname + "/app.js", "utf8");
  const block = src.slice(src.indexOf("async function savePrintNotes"),
                          src.indexOf("async function johnDownload"));
  assert.ok(block.indexOf("CW.listUpsert(PRINT_NOTE_LIST") > 0, "notes are saved with listUpsert");
  assert.strictEqual(/CW\.(setValues|setFill|clearFill|appendLog|batchWrite|moveJobRow|saveProgress|listDelete)/.test(block), false,
    "and with nothing else at all");
  const johnBlock = src.slice(src.indexOf("const PRINT_NOTE_LIST"), src.indexOf("/* ---------- render ---------- */"));
  assert.strictEqual(/\/workbook/.test(johnBlock), false, "no path in this whole block names the workbook");
  assert.strictEqual(/setFill|clearFill|setValues|batchWrite|moveJobRow/.test(johnBlock), false,
    "and not one workbook write of any kind appears in it");
  assert.strictEqual((johnBlock.match(/noteChange\(/g) || []).length, 1,
    "one log line per print, written in one place");
  pass("across the whole run: one Dashboard Log line, one list, listUpsert only, no workbook write");

  /* ---- 14. nothing new in this browser ---- */
  const keys = Object.keys(mem).sort();
  const OLD = ["cw_pending", "cw_changes", "cw_theme", "cw_hidden", "cw_collapsed", "cw_pendv", "cw_penda",
               "cw_stationfeed", "cw_exportpresets"];
  assert.deepStrictEqual(keys.filter(k => OLD.indexOf(k) < 0), ["cw_listids"],
    "nothing was added to this browser but the list-id cache. Keys: " + keys.join(", "));
  const everything = keys.map(k => String(mem[k])).join(" | ");
  assert.strictEqual(everything.indexOf(PHONE), -1, "and no phone number was left behind in it");
  assert.strictEqual(everything.indexOf(JPHONE), -1);
  assert.strictEqual(everything.indexOf(EIR), -1, "nor an eircode");
  pass("nothing about a job, and no contact detail, is kept in this browser");

  console.log("\n" + n + " checks passed");
  process.exit(0);                 // the reconcile timer would hold the process open
})().catch(e => { console.error("FAIL", e); process.exit(1); });
