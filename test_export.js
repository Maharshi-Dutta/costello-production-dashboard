/* Offline test of the export feature: choosing the jobs (every filter on its
   own and together, the three scopes, the text search), the rows, the
   filename, the filter summary, the presets, the real Excel workbook read back
   through ExcelJS, the pdfmake document definition, the Export window, the
   single log line a download writes - and, at the end, real PDFs rendered by
   the vendored pdfmake and read back to prove no job was dropped.

   The standing rules this file is here to hold:
   - a job's phone digits and eircode never appear in an exported file (values,
     absolutely); and no column heading, sheet name or field label ever says
     "phone" or "eircode". What somebody typed into a comment is their text -
     one fixture comment says "phone" on purpose, to prove free text is not
     censored while the two columns stay out.
   - no fetch(). Nothing in the export path may reach the network; the only
     write is the Dashboard Log line, and even that goes through a stub here.

   Every name, county and address below is invented. Run: node test_export.js */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, addEventListener() {}, removeEventListener() {} };
global.ExcelJS = ExcelJS;

const SEQ = [];                 // what a Download actually did, in order
let FETCHES = 0;
global.fetch = async () => { FETCHES++; throw new Error("the export path must never use the network"); };

function stubEl(tag) {
  let html = "";
  const e = {
    tag: tag || "div", style: {}, dataset: {}, textContent: "", value: "", disabled: false,
    hidden: false, className: "", href: "", download: "", src: "", scrollTop: 0, removed: false, kids: [],
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); }, remove() { e.removed = true; },
    addEventListener() {}, removeEventListener() {},
    focus() {}, blur() {}, click() { SEQ.push("anchor.click " + e.download); },
    getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => stubEl(), querySelectorAll: () => []
  };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); e.kids.length = 0; } });
  return e;
}
const EL = {};
/* The print-notes window is never opened in this file, so its selector
   answers null the way a browser does for something that is not on the
   page - otherwise the page's Escape handler, which closes the topmost
   window first, would think it was there. */
const el = sel => sel === "#nhost" ? null : (EL[sel] || (EL[sel] = stubEl()));
const MADE = [];
const HANDLERS = {};            // the page's own document-level listeners, so a key can be pressed
global.document = {
  documentElement: stubEl(), body: stubEl(), head: stubEl(), activeElement: null,
  title: "Costello Production",
  createElement: t => { const e = stubEl(t); MADE.push(e); return e; },
  querySelector: el, querySelectorAll: () => [],
  addEventListener: (t, fn) => { (HANDLERS[t] = HANDLERS[t] || []).push(fn); },
  removeEventListener() {}
};
/* the browser's two halves of a download, watched rather than performed */
global.URL.createObjectURL = () => { SEQ.push("URL.createObjectURL"); return "blob:test"; };
global.URL.revokeObjectURL = () => { SEQ.push("URL.revokeObjectURL"); };
global.Blob = function Blob(parts, o) { this.parts = parts; this.type = (o || {}).type || ""; this.size = (parts[0] || "").length || 0; };
/* pdfmake, watched the same way while the window is under test; the real
   vendored one is loaded over it at the end of the run */
let PDFDEF = null;
const PDFSTUB = {
  addVirtualFileSystem() {},
  createPdf(def) {
    SEQ.push("pdfMake.createPdf");
    PDFDEF = def;
    return { download: name => { SEQ.push("pdf.download " + name); } };
  }
};
global.pdfMake = PDFSTUB;

/* Graph is not loaded at all - CW is a stub, so there is no route to the
   network even by accident. appendLog records the one line an export writes. */
const LOGS = [];
global.CW = {
  account: { username: "exporter@example.test", name: "Pat Exporter" },
  appendLog: (who, job, what, from, to) => { LOGS.push({ who, job, what, from, to }); return Promise.resolve(true); },
  initAuth: async () => null, openSession: async () => {}, signOut() {}
};

/* ---------- load the page's own code, in the page's own order ---------- */
const run = f => vm.runInThisContext(fs.readFileSync(__dirname + "/" + f, "utf8"), { filename: f });
run("parser.js");
run("checkpoints.js");
global.CP = window.CP;
run("export.js");
run("app.js");

const TOASTS = [];
global.toast = (m, err) => { TOASTS.push({ m: String(m), err: !!err }); SEQ.push("toast"); };
global.whoAmI = () => "exporter@example.test";

/* ---------- the fixtures: invented jobs, distinctive fake contact values ---- */
const PH3 = "742", EIR = "ZZ99XY7";
/* the whole number, carried by the John template and by nothing else. It ends
   in the same three digits as PH3, so every standing "no phone" assertion
   below catches a Default export that leaked either of them. */
const PHONE = "021 555 0" + PH3;
/* the number Production (2) holds for the same job - deliberately not the
   model's, so a John print that quietly used the wrong sheet is caught */
const JPHONE = "065 555 0" + PH3;
const SECTIONS = ["Can sell as second hand", "Ready, customer won't take",
                  "Collect & supply only", "Ready to fit", "In production"];
const mkJob = o => Object.assign({
  id: "R0000", cust: "", area: "", off: "", colour: "", wnd: 0, drs: 0,
  ph3: PH3, eir: EIR, ph: PHONE,            // present on every job; only John carries the phone
  flag: "", flagHex: "",                    // the colour of the row's own text on Production
  glass: {}, prods: [], notes: [], sheets: ["Production"], src: {},
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cp: { win: "", drs: "", glass: {}, prod: {} },
  cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0
}, o);

const A = mkJob({
  id: "R1001", cust: "Customer One", area: "Cork", off: "OF-101", colour: "White",
  wnd: 12, drs: 4, glass: { tg: 8, dg: 6 },
  prods: [{ n: "7000 casement", f: 24, s: 18, t: 6, st: ["process"] }],
  cp: { win: "done", drs: "process", glass: { tg: "done", dg: "" },
        prod: { "7000 casement": { f: "done", s: "process" } } },
  dates: { sold: "2026-01-05", stamp: "2026-01-20", ivana: null, ready: "2026-02-10", floor: "2026-03-01" },
  notes: [{ k: "comment", t: "Ring before delivery", s: "Production" },
          { k: "brendan", t: "Check the cill height", s: "Production" }],
  sheets: ["Production", "Glass"], blk: 4, seq: 2, stage: "floor",
  flag: "trade", flagHex: "FF3399"
});
const B = mkJob({
  id: "R1002", cust: "Customer Two", area: "Kerry", off: "OF-102", colour: "Grey",
  wnd: 5, drs: 0, glass: { dg: 3 },
  prods: [{ n: "sliding", f: 5, s: 5, t: 0, st: ["done"] }],
  cp: { win: "done", drs: "", glass: { dg: "done" }, prod: { sliding: { f: "done", s: "done" } } },
  dates: { sold: "2026-02-01", stamp: null, ivana: null, ready: null, floor: null },
  notes: [{ k: "specials", t: "Urgent - customer waiting", s: "Production" }],
  blk: 3, seq: 1, done: 1, urg: 1, stage: "office", flag: "urgent", flagHex: "FF0000"
});
/* R1003 carries a date the sheet holds as words, not a date */
const C = mkJob({
  id: "R1003", cust: "Customer Three", area: "Cork", off: "OF-103",
  drs: 2, sheets: ["Production", "PVC Doors"], blk: 4, seq: 3, stage: "office",
  dates: { sold: "Before 12 Nov", stamp: null, ivana: null, ready: null, floor: null }
});
const PAST = mkJob({ id: "R0900", cust: "Customer Nine", cat: "past", blk: -1 });

/* the job that used to disappear: taller than a page on its own */
const TALLPRODS = [];
for (let i = 1; i <= 20; i++) TALLPRODS.push({ n: "profile " + i, f: i, s: 20 - i, t: (i % 5) + 1, st: [] });
const TALL = mkJob({
  id: "R1004", cust: "Customer Four", area: "Clare", off: "OF-104", colour: "Anthracite",
  wnd: 44, drs: 9, glass: { tg: 30, dg: 22, tuff: 11 }, prods: TALLPRODS,
  dates: { sold: "2026-01-02", stamp: "2026-01-09", ivana: "2026-01-16", ready: "2026-01-23", floor: "2026-02-02" },
  notes: (function () { const out = []; for (let i = 1; i <= 30; i++)
    out.push({ k: "comment", t: "Note number " + i + " about the fabrication of this job", s: "Production" }); return out; })(),
  blk: 4, seq: 4, stage: "floor", flag: "hold", flagHex: "00B0F0"
});

function useJobs(list) {
  global.__jobs = list; global.__names = SECTIONS;
  vm.runInThisContext("ALL = __jobs; BLOCKNAMES = __names; ALL.blockNames = __names; " +
    "CHANGES = []; state.picked = {}; state.q = ''; state.cat = null; state.sheet = null; " +
    "state.sort = 'id'; state.desc = false; state.view = 'flat'; state.hidden = {}; ALERTS = {};");
}
useJobs([A, B, C, PAST]);
/* dashboard comments (one of them deliberately says "phone", to prove free text
   is exported as written) and one alert, so hasComments / hasAlerts have
   something to find. The alert address is fake and only ever exported as a count. */
/* CHANGES is newest-first, the way the app keeps it; commentsFor() turns it
   back the other way up, so the older line comes out first */
vm.runInThisContext("CHANGES = [" +
  "{at:'2026-03-03T11:00:00.000Z', who:'office@example.test', job:'R1001', what:'Comment', from:'', to:'Customer says the phone line is out, call the site', src:'dashboard'}," +
  "{at:'2026-03-02T09:15:00.000Z', who:'office@example.test', job:'R1001', what:'Comment', from:'', to:'Chase the glass supplier', src:'dashboard'}];" +
  "ALERTS = { R1002: [{ email:'watcher@example.test', who:'admin@example.test', when:'2026-03-01' }] };");

const CTX = () => ({ all: [A, B, C], view: [A, B, C], picked: {}, sections: SECTIONS,
                     comments: id => commentsFor(id), alerts: id => alertsFor(id) });
const ids = list => list.map(j => j.id);
const F = o => Object.assign(exportDefaults(), o || {});

/* ---------- recursive scans ---------- */
/** Every string in a value, keys included, following arrays and objects. */
function strings(v, out) {
  out = out || [];
  if (v == null) return out;
  const t = typeof v;
  if (t === "string") { out.push(v); return out; }
  if (t !== "object") return out;
  if (Array.isArray(v)) { v.forEach(x => strings(x, out)); return out; }
  if (v instanceof Date) { out.push(v.toISOString()); return out; }
  Object.keys(v).forEach(k => { out.push(k); strings(v[k], out); });
  return out;
}
/** The absolute half: a job's phone digits and eircode appear nowhere, ever. */
function assertNoValues(list, where) {
  list.forEach(s => {
    const low = String(s).toLowerCase();
    assert.strictEqual(low.indexOf(PH3.toLowerCase()), -1, where + ' leaked the phone digits: "' + s + '"');
    assert.strictEqual(low.indexOf(EIR.toLowerCase()), -1, where + ' leaked the eircode: "' + s + '"');
  });
}
/** The naming half: no heading, sheet name or field label names either column.
    Free text is not covered - a comment may say anything the office wrote. */
function assertNoHeaderWords(names, where) {
  names.forEach(s => {
    const low = String(s).toLowerCase();
    assert.strictEqual(low.indexOf("phone"), -1, where + ' has a column called "' + s + '"');
    assert.strictEqual(low.indexOf("eircode"), -1, where + ' has a column called "' + s + '"');
  });
}
async function roundTrip(wb) {
  const buf = await wb.xlsx.writeBuffer();
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(buf);
  return back;
}
function sheetStrings(wb) {
  const out = [];
  wb.eachSheet(ws => {
    out.push(ws.name);
    ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, cell => {
      strings(cell.value, out);
      if (cell.numFmt) out.push(cell.numFmt);
    }));
    Object.keys(ws.tables || {}).forEach(k => { out.push(k); strings((ws.tables[k] || {}).table, out); });
  });
  return out;
}
const headerRow = (ws, r) => {
  const out = [];
  ws.getRow(r).eachCell({ includeEmpty: false }, c => out.push(String(c.value)));
  return out;
};
function tileValue(def, label) {
  let found = null;
  (function walk(v) {
    if (found !== null || !v || typeof v !== "object") return;
    if (Array.isArray(v)) { v.forEach(walk); return; }
    const st = v.table && v.table.body && v.table.body[0] && v.table.body[0][0] && v.table.body[0][0].stack;
    if (st && st[1] && st[1].text === label) { found = Number(st[0].text); return; }
    Object.keys(v).forEach(k => walk(v[k]));
  })(def.content);
  return found;
}
function collect(v, pred, out) {
  out = out || [];
  if (!v || typeof v !== "object") return out;
  if (Array.isArray(v)) { v.forEach(x => collect(x, pred, out)); return out; }
  if (pred(v)) out.push(v);
  Object.keys(v).forEach(k => collect(v[k], pred, out));
  return out;
}

/* ---------- reading a rendered PDF back ----------
   pdfkit embeds a subset of Roboto and writes text as glyph numbers, so a
   rendered page cannot simply be searched for a job number. Each font does
   carry a /ToUnicode CMap, though, and with `compress: false` the whole file
   is plain: this maps the glyphs back to characters, which is what makes
   "every job is in the file" a real assertion rather than a hopeful one. */
function pdfText(buf) {
  const s = buf.toString("latin1");
  /* An object body, found by its header. The embedded font files are binary and
     can happen to contain the bytes "18 0 obj", so the header is anchored to a
     line start and the body has to actually look like a dictionary: picking up
     a false one silently decodes half the page with the wrong font. */
  const objBody = num => {
    const re = new RegExp("\\n" + num + " 0 obj\\n([\\s\\S]*?)\\nendobj", "g");
    let m;
    while ((m = re.exec(s))) if (/^\s*<</.test(m[1])) return m[1];
    return "";
  };
  /* a destination can be more than one code unit - Roboto's "fi" ligature maps
     back to <0066 0069>, and skipping it shifts every later glyph by one */
  const hex2 = h => { const x = String(h).replace(/\s+/g, ""); let o = "";
    for (let i = 0; i + 3 < x.length; i += 4) o += String.fromCharCode(parseInt(x.substr(i, 4), 16)); return o; };
  function cmapOf(objText) {
    const map = {};
    const st = objText.indexOf("stream"), en = objText.lastIndexOf("endstream");
    const inner = st >= 0 ? objText.slice(st + 6, en < 0 ? objText.length : en) : objText;
    let bm;
    const bc = /beginbfchar([\s\S]*?)endbfchar/g;
    while ((bm = bc.exec(inner))) {
      const rr = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F\s]+)>/g; let r;
      while ((r = rr.exec(bm[1]))) map[parseInt(r[1], 16)] = hex2(r[2]);
    }
    const br = /beginbfrange([\s\S]*?)endbfrange/g;
    while ((bm = br.exec(inner))) {
      const rr = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(?:\[([\s\S]*?)\]|<([0-9a-fA-F]+)>)/g; let r;
      while ((r = rr.exec(bm[1]))) {
        const lo = parseInt(r[1], 16), hi = parseInt(r[2], 16);
        if (r[3] != null) (r[3].match(/<[0-9a-fA-F\s]+>/g) || []).forEach((it, i) => { map[lo + i] = hex2(it.slice(1, -1)); });
        else { const base = parseInt(r[4], 16); for (let c = lo; c <= hi; c++) map[c] = String.fromCharCode(base + (c - lo)); }
      }
    }
    return map;
  }
  const fonts = {};
  let fm; const fre = /\/Font\s*<<([\s\S]*?)>>/g;
  while ((fm = fre.exec(s))) {
    let pm; const pre = /\/(F\d+)\s+(\d+) 0 R/g;
    while ((pm = pre.exec(fm[1]))) {
      if (fonts[pm[1]]) continue;
      const body = objBody(pm[2]);
      const tu = /\/ToUnicode\s+(\d+) 0 R/.exec(body);
      fonts[pm[1]] = tu ? cmapOf(objBody(tu[1])) : {};
    }
  }
  let out = "", cur = {};
  const sre = /stream\r?\n([\s\S]*?)endstream/g; let sm;
  while ((sm = sre.exec(s))) {
    const body = sm[1];
    if (body.indexOf("BT") < 0 || body.indexOf(" Tf") < 0) continue;    // not a page's content
    const tre = /\/(F\d+)\s+[\d.]+\s+Tf|<([0-9a-fA-F]+)>|(ET)/g; let tm;
    while ((tm = tre.exec(body))) {
      if (tm[1]) cur = fonts[tm[1]] || {};
      else if (tm[2]) { const h = tm[2];
        for (let i = 0; i + 3 < h.length; i += 4) { const c = parseInt(h.substr(i, 4), 16); out += (cur[c] == null ? "" : cur[c]); } }
      else out += "\n";
    }
  }
  return out;
}
const pdfPages = buf => (buf.toString("latin1").match(/\/Type\s*\/Page[^s]/g) || []).length;
/* pdfmake lays each word out as its own text run, so the decoded page comes
   back one word per line; flatten it before looking for a phrase. */
const flat = t => String(t).replace(/\s+/g, " ");
/* a narrow table column wraps its heading mid-word ("Commen ts"), so a phrase
   that has to survive the layout is compared with the spaces taken out */
const squash = t => String(t).replace(/\s+/g, "");
/** How many text runs each page carries. This is what catches the bug the
    reviewer found: a page pdfmake could not fit a card onto still comes out,
    carrying the running header and footer and nothing else. */
function pdfRunsPerPage(buf) {
  const s = buf.toString("latin1"), out = [];
  const sre = /stream\r?\n([\s\S]*?)endstream/g; let m;
  while ((m = sre.exec(s))) {
    const b = m[1];
    if (b.indexOf("BT") < 0 || b.indexOf(" Tf") < 0) continue;
    out.push((b.match(/TJ/g) || []).length);
  }
  return out;
}

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };
  const noNet = where => assert.strictEqual(FETCHES, 0, "something called fetch() during " + where);

  /* ---- 1. the three scopes ---- */
  assert.deepStrictEqual(ids(exportScope(CTX(), F({ scope: "view" }))), ["R1001", "R1002", "R1003"]);
  const ticked = Object.assign(CTX(), { picked: { R1002: 1 } });
  assert.deepStrictEqual(ids(exportScope(ticked, F({ scope: "ticked" }))), ["R1002"]);
  assert.deepStrictEqual(ids(exportScope(CTX(), F({ scope: "sections", sections: [4] }))), ["R1001", "R1003"]);
  assert.deepStrictEqual(ids(exportScope(CTX(), F({ scope: "sections", sections: [3, 4] }))), ["R1001", "R1002", "R1003"]);
  assert.deepStrictEqual(ids(exportScope(CTX(), F({ scope: "ticked" }))), [], "nothing ticked means nothing to export");
  pass("scope: what I see now, the ticked jobs, and the chosen sections");

  /* ---- 2. every filter on its own ---- */
  const only = f => ids(exportFilter([A, B, C], F(f), CTX()));
  assert.deepStrictEqual(only({ ready: true }), ["R1002"]);
  assert.deepStrictEqual(only({ ready: false }), ["R1001", "R1003"]);
  assert.deepStrictEqual(only({ urgent: true }), ["R1002"]);
  pass("ready to deliver and urgent");

  assert.deepStrictEqual(only({ cp: { win: "done" } }), ["R1001", "R1002"]);
  assert.deepStrictEqual(only({ cp: { drs: "process" } }), ["R1001"]);
  assert.deepStrictEqual(only({ cp: { drs: "none" } }), ["R1003"], "doors on the sheet with no colour yet");
  assert.deepStrictEqual(only({ cp: { glass: "done" } }), ["R1002"]);
  assert.deepStrictEqual(only({ cp: { glass: "process" } }), ["R1001"], "one unit done and one not is in progress");
  assert.deepStrictEqual(only({ cp: { prod: "done" } }), ["R1002"]);
  assert.deepStrictEqual(only({ cp: { prod: "process" } }), ["R1001"]);
  assert.deepStrictEqual(only({ cp: { win: "done", glass: "done" } }), ["R1002"], "two checkpoint filters together");
  pass("checkpoint state, per group, taken from the sheet's own colours");

  assert.deepStrictEqual(only({ dates: { sold: { from: "2026-01-01", to: "2026-01-31" } } }), ["R1001"]);
  assert.deepStrictEqual(only({ dates: { sold: { from: "2026-01-01" } } }), ["R1001", "R1002"],
    'R1003\'s "Before 12 Nov" is not a date anybody can compare, so a range leaves it out');
  assert.deepStrictEqual(only({ dates: { sold: { to: "2026-01-31" } } }), ["R1001"]);
  assert.deepStrictEqual(only({ dates: { sold: {} } }), ["R1001", "R1002", "R1003"], "with no range set it stays in");
  assert.deepStrictEqual(only({ dates: { floor: { from: "2026-01-01" } } }), ["R1001"], "a job with no such date is out");
  assert.deepStrictEqual(only({ dates: { ready: { from: "2026-02-01", to: "2026-02-28" } } }), ["R1001"]);
  pass("the three date ranges, each end optional, text in a date cell counted as unknown");

  assert.deepStrictEqual(only({ county: ["cork"] }), ["R1001", "R1003"], "county matches whatever the case");
  assert.deepStrictEqual(only({ county: ["Cork", "Kerry"] }), ["R1001", "R1002", "R1003"]);
  assert.deepStrictEqual(only({ products: ["7000 casement"] }), ["R1001"]);
  assert.deepStrictEqual(only({ glassTypes: ["dg"] }), ["R1001", "R1002"]);
  assert.deepStrictEqual(only({ glassTypes: ["tg"] }), ["R1001"]);
  assert.deepStrictEqual(only({ sheets: ["PVC Doors"] }), ["R1003"]);
  assert.deepStrictEqual(only({ sheets: ["Glass", "PVC Doors"] }), ["R1001", "R1003"]);
  pass("county, product, glass type and sheet");

  assert.deepStrictEqual(only({ hasComments: true }), ["R1001"]);
  assert.deepStrictEqual(only({ hasComments: false }), ["R1002", "R1003"]);
  assert.deepStrictEqual(only({ hasAlerts: true }), ["R1002"]);
  assert.deepStrictEqual(only({ hasAlerts: false }), ["R1001", "R1003"]);
  pass("has comments / has an email alert");

  assert.deepStrictEqual(only({ q: "customer two" }), ["R1002"]);
  assert.deepStrictEqual(only({ q: "ring before" }), ["R1001"], "the sheet's own notes are searched");
  assert.deepStrictEqual(only({ q: "OF-103" }), ["R1003"]);
  assert.deepStrictEqual(only({ q: "R1001" }), ["R1001"]);
  assert.deepStrictEqual(only({ q: "cork" }), ["R1001", "R1003"]);
  assert.deepStrictEqual(only({ q: EIR }), [],
    "the eircode is not in the haystack: the export search cannot even look at it");
  assert.deepStrictEqual(only({ q: PH3 }), [], "and neither are the phone digits");
  pass("the text search - the dashboard's, minus the eircode");

  /* ---- 3. filters combined, and the filter object is never aliased ---- */
  assert.deepStrictEqual(only({ ready: false, county: ["Cork"], cp: { drs: "process" } }), ["R1001"]);
  assert.deepStrictEqual(only({ ready: true, county: ["Cork"] }), [], "nothing matches both");
  assert.deepStrictEqual(ids(exportJobs(Object.assign(CTX(), { picked: { R1001: 1, R1002: 1 } }),
    F({ scope: "ticked", urgent: true }))), ["R1002"], "the filter applies on top of the scope");
  const mine = { cp: { win: "done" }, dates: { sold: { from: "2026-01-01" } }, county: ["Cork"], sections: [4] };
  const copy = xpFilter(mine);
  copy.cp.win = "none"; copy.dates.sold.from = "1999-01-01"; copy.county.push("Kerry"); copy.sections.push(9);
  assert.strictEqual(mine.cp.win, "done", "xpFilter copies the nested objects instead of pointing at them");
  assert.strictEqual(mine.dates.sold.from, "2026-01-01");
  assert.deepStrictEqual(mine.county, ["Cork"]);
  assert.deepStrictEqual(mine.sections, [4]);
  pass("filters combine, sit on top of the scope, and never mutate the caller's filter");

  /* ---- 4. sorting ---- */
  const sorted = f => ids(exportSort([A, B, C], F(f)));
  assert.deepStrictEqual(sorted({ sort: "id" }), ["R1001", "R1002", "R1003"]);
  assert.deepStrictEqual(sorted({ sort: "id", desc: true }), ["R1003", "R1002", "R1001"]);
  assert.deepStrictEqual(sorted({ sort: "urgent" }), ["R1002", "R1001", "R1003"]);
  assert.deepStrictEqual(sorted({ sort: "county" }), ["R1001", "R1003", "R1002"]);
  assert.deepStrictEqual(sorted({ sort: "size" }), ["R1001", "R1002", "R1003"]);
  assert.deepStrictEqual(sorted({ sort: "wait" }), ["R1001", "R1002", "R1003"],
    "longest wait: the one with a Ready to print date comes before the two without");
  assert.deepStrictEqual(ids(exportSort([mkJob({ id: "S200", stage: "office" }), mkJob({ id: "C30", stage: "office" })],
    F({ sort: "num" }))), ["C30", "S200"], "job number ignoring the letter");
  assert.deepStrictEqual(ids(exportSort([C, B], F({ sort: "cat" }))), ["R1002", "R1003"],
    "category order: ready to deliver before in-office");
  pass("the dashboard's own sort keys, and the reversed flag");
  noNet("choosing the jobs");

  /* ---- 5. the rows ---- */
  const rowsAll = exportRows([A, B, C], exportAllFields(), CTX());
  assert.strictEqual(rowsAll.length, 3);
  assert.strictEqual(rowsAll[0].id, "R1001");
  assert.strictEqual(rowsAll[0].section, "In production");
  assert.strictEqual(rowsAll[1].section, "Ready to fit");
  assert.strictEqual(rowsAll[0].wnd, 12);
  assert.strictEqual(rowsAll[1].alerts, 1, "alerts come out as a count, never as an address");
  assert.deepStrictEqual(rowsAll[0].prods[0].st, { f: "done", s: "process", t: "" });
  assert.deepStrictEqual(rowsAll[0].glass.map(g => g.type + ":" + g.qty + ":" + g.st), ["tg:8:done", "dg:6:"]);
  assert.strictEqual(rowsAll[2].dates.sold, "Before 12 Nov", "a text date is kept whole, not sliced or dropped");
  assert.deepStrictEqual(rowsAll[0].comments.map(c => c.who), ["office", "office"],
    "a comment names a person, never the address the log holds");
  assert.deepStrictEqual(rowsAll[0].notes.map(x => x.src), ["Sheet comment", "Brendan's notes"]);
  assert.strictEqual(JSON.stringify(rowsAll).indexOf("@"), -1, "no address of any kind reaches a row");
  pass("rows carry the job in plain data, people by name, alerts as a count only");

  const few = exportRows([A], ["job", "cust"], CTX());
  assert.deepStrictEqual(Object.keys(few[0]), ["id", "cust"], "an unticked field has no key at all");
  assert.deepStrictEqual(Object.keys(exportRows([A], {}, CTX())[0]), ["id"],
    "no fields ticked leaves only the job number, which is the row's key rather than a field");
  assert.deepStrictEqual(Object.keys(exportRows([A], ["job", "ph3", "eir", "phone"], CTX())[0]), ["id"],
    "a field name that is not in the picker is dropped, so nothing can be smuggled in");
  assertNoValues(strings(rowsAll), "a row");
  assertNoHeaderWords(EXPORT_FIELDS.map(p => p[1]).concat(EXPORT_FIELD_KEYS), "the field picker");
  pass("only the ticked fields appear, and nothing outside the picker can be asked for");

  /* ---- 6. checkpoints per group ---- */
  const cpA = exportCheckpoints(A);
  assert.deepStrictEqual(cpA.map(g => g.label), ["Windows", "Doors", "Glass", "7000 casement"]);
  assert.deepStrictEqual(cpA[0], { group: "win", label: "Windows", done: 12, total: 12, unknown: false,
    items: [{ item: "win", label: "Windows", done: 12, total: 12, status: "done" }], status: "done" });
  assert.strictEqual(cpA[2].total, 14, "glass groups both unit types");
  assert.strictEqual(cpA[2].done, 8);
  assert.strictEqual(cpA[2].status, "process");
  assert.strictEqual(cpA[1].unknown, true, "doors are in fabrication with no count entered");
  pass("checkpoints roll up into the same groups the drawer shows");

  /* ---- 7. filename, summary, log line ---- */
  const when = new Date(2026, 8, 7, 14, 30);
  assert.strictEqual(exportFilename("xlsx", F({ scope: "sections", sections: [4], sectionNames: ["In production"] }), when),
    "Production export - In production - 2026-09-07.xlsx");
  assert.strictEqual(exportFilename("pdf", F({ scope: "sections", sections: [3] }), when, SECTIONS),
    "Production export - Ready to fit - 2026-09-07.pdf",
    "the section's name comes from the sheet when the filter only carries its number");
  assert.strictEqual(exportFilename("pdf", F({ scope: "sections", sections: [3, 4] }), when, SECTIONS),
    "Production export - Ready to fit + In production - 2026-09-07.pdf");
  assert.strictEqual(exportScopeLabel(F({ scope: "sections", sections: [0, 1, 2, 3, 4] }), SECTIONS),
    "Can sell as second hand + 4 more", "a long list of sections is cut down to something readable");
  assert.strictEqual(exportFilename("pdf", F({ ready: true }), when),
    "Production export - Ready to deliver - 2026-09-07.pdf");
  assert.strictEqual(exportFilename("xlsx", F({}), when), "Production export - All jobs - 2026-09-07.xlsx");
  assert.strictEqual(exportFilename("xlsx", F({ scope: "ticked" }), when), "Production export - Ticked jobs - 2026-09-07.xlsx");
  assert.strictEqual(exportFilename("xlsx", F({ q: 'ring the site' }), when), "Production export - All jobs - 2026-09-07.xlsx",
    "what somebody typed into the search box does not belong in a filename");
  assert.strictEqual(exportFilename("xlsx", F({ q: 'a/b:c*?"<>|d' }), when).indexOf("/"), -1,
    "characters Windows will not take in a filename are taken out");
  pass("the filename says what is in the file, by name, and the date it was made");

  assert.strictEqual(filtersSummaryLines(F({ scope: "sections", sections: [3] }), SECTIONS)[0],
    "Scope: Ready to fit", "the summary names the section rather than numbering it");
  assert.strictEqual(filtersSummaryLines(F({ scope: "sections", sections: [3] }))[0], "Scope: section 3",
    "and says so plainly when there is no list of names to look it up in");
  const sum = filtersSummary(F({ scope: "sections", sections: [4], sectionNames: ["In production"],
    ready: false, urgent: true, cp: { glass: "done" }, county: ["Cork"],
    dates: { sold: { from: "2026-01-01", to: "2026-01-31" } }, q: "cill", sort: "urgent", desc: true, groupBySection: true }));
  ["Scope: In production", "In production only", "Urgent only", "Glass: done",
   "Sold between 01-Jan-2026 and 31-Jan-2026", "County: Cork", 'Search: "cill"',
   "Sorted by Urgent first (reversed)", "Grouped by section"].forEach(bit =>
    assert.ok(sum.indexOf(bit) >= 0, "the summary should say: " + bit + "\n" + sum));
  assert.deepStrictEqual(filtersSummaryLines(F({})), ["Scope: what I see now", "Sorted by Job no (A-Z)"],
    "with nothing set it still says what the scope and the order were");
  assert.strictEqual(exportLogFrom("pdf", 12), "PDF · 12 jobs");
  assert.strictEqual(exportLogFrom("xlsx", 1), "Excel · 1 job");
  assertNoValues(strings(sum), "the filter summary");
  pass("the filter summary reads as a sentence and is the same text everywhere");

  /* ---- 8. presets ---- */
  delete mem.cw_exportpresets;
  assert.deepStrictEqual(presetsLoad(), {});
  presetSave("Weekly delivery run", { format: "pdf", layout: "table",
    fields: ["job", "cust", "ready"], f: F({ ready: true, sort: "county" }) });
  let ps = presetsLoad();
  assert.deepStrictEqual(Object.keys(ps), ["Weekly delivery run"]);
  assert.strictEqual(ps["Weekly delivery run"].format, "pdf");
  assert.strictEqual(ps["Weekly delivery run"].layout, "table");
  assert.deepStrictEqual(ps["Weekly delivery run"].fields, { job: true, cust: true, ready: true });
  assert.strictEqual(ps["Weekly delivery run"].f.ready, true);
  assert.strictEqual(ps["Weekly delivery run"].f.sort, "county");
  const stored = String(mem.cw_exportpresets);
  ["R1001", "R1002", "Customer One", "Customer Two", "Cork", PH3, EIR].forEach(bad =>
    assert.strictEqual(stored.indexOf(bad), -1, "a preset must not hold job data: " + bad));
  presetSave("Weekly delivery run", { format: "xlsx", fields: ["job"], f: F({}) });
  assert.strictEqual(presetsLoad()["Weekly delivery run"].format, "xlsx", "saving again over the same name replaces it");
  presetSave("Second", { f: F({}) });
  presetDelete("Weekly delivery run");
  assert.deepStrictEqual(Object.keys(presetsLoad()), ["Second"]);
  presetDelete("Second");
  assert.deepStrictEqual(presetsLoad(), {});
  presetSave("", { f: F({}) });
  assert.deepStrictEqual(presetsLoad(), {}, "a preset with no name is not saved");
  mem.cw_exportpresets = JSON.stringify({ Dodgy: { f: { q: "x" }, jobs: [A], fields: ["job", "eir"] } });
  const cleaned = presetsLoad().Dodgy;
  assert.strictEqual(cleaned.jobs, undefined);
  assert.deepStrictEqual(cleaned.fields, { job: true });
  delete mem.cw_exportpresets;
  pass("presets save, load and delete, and can only ever hold choices");
  noNet("the presets");

  /* ---- 9. the Excel workbook ---- */
  const fields = exportAllFields();
  const rows = exportRows([A, B, C], fields, CTX());
  const opts = { fields: fields, filters: F({ ready: null, sort: "id" }), who: "Pat Exporter",
                 when: when, company: "Costello Production", sections: SECTIONS, build: "20260907-0925" };
  const back = await roundTrip(buildWorkbook(rows, opts));
  assert.deepStrictEqual(back.worksheets.map(w => w.name), ["Jobs", "Comments", "Checkpoints", "Export info"]);
  const jobs = back.getWorksheet("Jobs");
  const jobsHead = headerRow(jobs, 1);
  assert.deepStrictEqual(jobsHead, [
    "Comment", "Office no", "Job no", "Sold", "Stamp", "Ivana", "Ready to print", "Sent to floor",
    "Customer", "Area", "Section", "Ready to deliver", "Urgent", "WND", "DRS", "Windows colour",
    "7000 CASEMENT F", "7000 CASEMENT S", "7000 CASEMENT T", "SLIDING F", "SLIDING S", "SLIDING T",
    "TG", "DG", "Alerts", "Comments", "Flag"
  ], "the Production sheet's own order, without PHONE NO. and without EIRCODE; " +
     "Flag last, because it is not one of the sheet's columns but the colour of the row's own text");
  assert.deepStrictEqual(Object.keys(jobs.tables), ["Jobs_1"], "a real Excel table, not just cells");
  assert.strictEqual(jobs.views[0].state, "frozen");
  assert.strictEqual(jobs.views[0].ySplit, 1, "the header row stays put when you scroll");
  assert.strictEqual(jobs.getColumn(1).width, 34, "the comment column is wide enough to read");
  pass("Jobs sheet: the sheet's column order, a real table, a frozen header");

  assert.ok(jobs.getCell("D2").value instanceof Date, "Sold is a real date cell, not text");
  assert.strictEqual(jobs.getCell("D2").numFmt, "dd-mmm-yyyy");
  assert.strictEqual(jobs.getCell("D2").value.toISOString().slice(0, 10), "2026-01-05");
  assert.strictEqual(jobs.getCell("F2").value, null, "a date nobody has filled in stays empty");
  assert.strictEqual(jobs.getCell("D4").value, "Before 12 Nov",
    "and what the office wrote in words is written into the cell as the text it is");
  assert.strictEqual(jobs.getCell("D4").numFmt, undefined, "with no date format pretending otherwise");
  assert.strictEqual(jobs.getCell("N2").value, 12, "WND is a number");
  assert.strictEqual(typeof jobs.getCell("N2").value, "number");
  assert.strictEqual(jobs.getCell("C2").value, "R1001");
  pass("dates are dates, words are words, and numbers are numbers");

  const fillOfCell = c => (c.fill && c.fill.fgColor || {}).argb || "";
  assert.strictEqual(fillOfCell(jobs.getCell("C3")), "FFE8F5E3", "R1002 is ready to deliver: its job number is green");
  assert.strictEqual(jobs.getCell("L3").value, "Yes", "and the word is right beside the colour");
  assert.strictEqual(fillOfCell(jobs.getCell("C2")), "", "R1001 is not ready, so no green");
  assert.strictEqual((jobs.getCell("C3").font || {}).bold, true);
  assert.strictEqual(((jobs.getCell("C3").font || {}).color || {}).argb, "FFC62828", "and it is urgent, so red and bold");
  assert.strictEqual(jobs.getCell("M3").value, "Yes", "with the Urgent column saying so in words");
  assert.strictEqual(fillOfCell(jobs.getCell("Q2")), "FFFFE699", "7000 CASEMENT F is done: gold");
  assert.strictEqual(jobs.getCell("Q2").value, 24, "with the number in the cell");
  assert.strictEqual(fillOfCell(jobs.getCell("R2")), "FFFFFF00", "S is in fabrication: yellow");
  assert.strictEqual(fillOfCell(jobs.getCell("S2")), "", "T has no colour on the sheet yet");
  assert.strictEqual(fillOfCell(jobs.getCell("W2")), "FFFFE699", "TG glass is done");
  assert.strictEqual(jobs.getCell("W2").value, 8);
  assert.strictEqual(fillOfCell(jobs.getCell("X2")), "", "DG is not");
  pass("the fills mirror the sheet's meaning and always sit next to a number or a word");

  const cmt = back.getWorksheet("Comments");
  assert.deepStrictEqual(headerRow(cmt, 1), ["Job", "Source", "Who", "When", "Text"]);
  const cmtRows = [];
  cmt.eachRow({ includeEmpty: false }, (r, i) => { if (i > 1) cmtRows.push(r.values.slice(1).map(v => v == null ? "" : String(v))); });
  assert.deepStrictEqual(cmtRows[0].slice(0, 2), ["R1001", "Sheet comment"]);
  assert.deepStrictEqual(cmtRows[1].slice(0, 2), ["R1001", "Brendan's notes"]);
  assert.deepStrictEqual(cmtRows[2].slice(0, 3), ["R1001", "Dashboard", "office"], "a name, not an address");
  assert.strictEqual(cmtRows[2][4], "Chase the glass supplier");
  assert.strictEqual(cmtRows[3][4], "Customer says the phone line is out, call the site",
    "free text is exported exactly as it was written - the rule is about the two columns, not the words");
  assert.deepStrictEqual(cmtRows[4].slice(0, 2), ["R1002", "Specials"]);
  pass("Comments sheet: the sheet's own notes and the dashboard's, each with its source");

  const chk = back.getWorksheet("Checkpoints");
  assert.deepStrictEqual(headerRow(chk, 1), ["Job", "Group", "Item", "Done", "Total", "Status"]);
  assert.deepStrictEqual(chk.getRow(2).values.slice(1), ["R1001", "Windows", "Windows", 12, 12, "done"]);
  assert.strictEqual(fillOfCell(chk.getCell("F2")), "FFFFE699");
  assert.strictEqual(chk.getRow(3).values[6], "in fabrication");
  assert.strictEqual(fillOfCell(chk.getCell("F3")), "FFFFFF00");
  pass("Checkpoints sheet: one line per item, with the word as well as the colour");

  const info = back.getWorksheet("Export info");
  const infoText = sheetStrings({ eachSheet: cb => cb(info) }).join(" | ");
  ["Costello Production", "Production export", "Exported by", "Pat Exporter",
   "2026-09-07 14:30", "Jobs in this file", "20260907-0925", "Filters",
   "Scope: what I see now"].forEach(bit =>
    assert.ok(infoText.indexOf(bit) >= 0, "Export info should say: " + bit));
  assert.strictEqual(infoText.indexOf("exporter@example.test"), -1, "an exported file names a person, not an address");
  assert.strictEqual(info.getCell("B7").value, 3, "the job count is the number of jobs in the file");
  pass("Export info: who, when, what and which filters - always written");

  assertNoValues(sheetStrings(back), "the workbook");
  assertNoHeaderWords(back.worksheets.map(w => w.name).concat(jobsHead)
    .concat(headerRow(cmt, 1)).concat(headerRow(chk, 1)), "the workbook's headings");
  pass("no phone number or eircode value anywhere, and no column named after either");

  /* only the ticked fields become sheets and columns */
  const slim = await roundTrip(buildWorkbook(
    exportRows([A, B, C], ["job", "cust", "ready"], CTX()),
    { fields: ["job", "cust", "ready"], filters: F({}), who: "Pat Exporter", when: when }));
  assert.deepStrictEqual(slim.worksheets.map(w => w.name), ["Jobs", "Export info"],
    "no comments field, no Comments sheet; no checkpoints field, no Checkpoints sheet");
  assert.deepStrictEqual(headerRow(slim.getWorksheet("Jobs"), 1), ["Job no", "Customer", "Ready to deliver"]);
  pass("unticking a field removes its columns and its sheet");

  /* grouped by section: one table each, with the section's name above it */
  const grouped = await roundTrip(buildWorkbook(rows, Object.assign({}, opts, { filters: F({ groupBySection: true }) })));
  const gj = grouped.getWorksheet("Jobs");
  assert.deepStrictEqual(Object.keys(gj.tables).sort(), ["Jobs_1", "Jobs_2"], "one table per section");
  assert.strictEqual(gj.getCell("A1").value, "In production — 2 jobs");
  assert.strictEqual(gj.getCell("A1").font.bold, true);
  assert.strictEqual(gj.getCell("C3").value, "R1001");
  assert.strictEqual(gj.getCell("C4").value, "R1003");
  assert.strictEqual(gj.getCell("A6").value, "Ready to fit — 1 job");
  assert.strictEqual(gj.getCell("C8").value, "R1002");
  assert.notStrictEqual(((gj.views || [])[0] || {}).state, "frozen",
    "row 1 is a section heading now, so freezing it would pin the wrong line");
  pass("group by section: one table per section under a bold heading, and nothing frozen");
  noNet("building the workbook");

  /* ---- 10. nothing ticked, and the checkpoints-only workbook ---- */
  const nothing = exportRows([A, B, C], {}, CTX());
  assert.strictEqual(exportBuildable(nothing, {}, "xlsx").ok, false);
  assert.strictEqual(exportBuildable(nothing, {}, "xlsx").why, "no fields ticked");
  assert.strictEqual(exportBuildable(nothing, {}, "pdf", "cards").ok, false);
  assert.strictEqual(exportBuildable(nothing, {}, "pdf", "table").ok, false);
  assert.strictEqual(exportBuildable([], exportAllFields(), "xlsx").ok, false, "and no jobs is no export either");
  assert.strictEqual(exportBuildable([], exportAllFields(), "xlsx").why, "no jobs match");
  assert.strictEqual(exportBuildable(rows, exportAllFields(), "pdf", "cards").ok, true);
  const cpOnly = exportRows([A, B, C], ["cp"], CTX());
  assert.strictEqual(exportColumns(cpOnly, ["cp"]).length, 0, "checkpoints alone give the Jobs sheet no columns");
  assert.strictEqual(exportBuildable(cpOnly, ["cp"], "xlsx").ok, true, "but the workbook is still worth making");
  const cpBook = await roundTrip(buildWorkbook(cpOnly, { fields: ["cp"], filters: F({}), who: "Pat Exporter", when: when }));
  assert.deepStrictEqual(cpBook.worksheets.map(w => w.name), ["Checkpoints", "Export info"],
    "the Jobs sheet is left out rather than added with no columns at all");
  assert.strictEqual(cpBook.getWorksheet("Checkpoints").getRow(2).values[1], "R1001");
  pass("with no fields there is nothing to build; with only checkpoints there still is");

  /* ---- 11. the PDF document definition ---- */
  const pdfOpts = Object.assign({}, opts, { layout: "cards", logo: null });
  const def = buildDocDefinition(rows, pdfOpts);
  assert.strictEqual(def.pageSize, "A4");
  assert.strictEqual(def.pageOrientation, "portrait");
  assert.strictEqual(def.defaultStyle.font, "Roboto");
  assert.strictEqual(tileValue(def, "JOBS"), 3);
  assert.strictEqual(tileValue(def, "WINDOWS"), 17, "12 + 5");
  assert.strictEqual(tileValue(def, "DOORS"), 6, "4 + 0 + 2");
  assert.strictEqual(tileValue(def, "COMPONENTS F+S+T"), 58, "24+18+6 and 5+5+0");
  assert.strictEqual(tileValue(def, "READY TO DELIVER"), 1);
  assert.strictEqual(tileValue(def, "URGENT"), 1);
  const coverText = strings(def.content).join(" | ");
  assert.ok(coverText.indexOf("Scope: what I see now") >= 0, "the cover carries the filter summary");
  assert.ok(coverText.indexOf("Jobs by section") >= 0);
  pass("cover: the big numbers add up, and the filters are written out");

  /* the shape that used to lose a job: only the head is unbreakable */
  const heads = collect(def.content, v => v.unbreakable === true);
  assert.strictEqual(heads.length, 3, "one head band per job");
  const breaks = collect(def.content, v => v.pageBreak === "before");
  assert.strictEqual(breaks.length, 1, "one page break in the whole document - the cover keeps its own page");
  assert.strictEqual(breaks[0], heads[0], "and it is on the first job, not between every pair");
  const headText = strings(heads[0]).join(" | ");
  ["R1001", "Customer One", "Cork", "In production", "IN PRODUCTION", "WINDOWS"].forEach(bit =>
    assert.ok(headText.indexOf(bit) >= 0, "the head band should show: " + bit));
  assert.strictEqual(headText.indexOf("Ring before delivery"), -1,
    "the comments are outside the unbreakable part, so a long job can flow over the page");
  const rules = collect(def.content, v => Array.isArray(v.canvas) && v.canvas.length === 1 &&
                                          v.canvas[0].type === "line" && v.canvas[0].x1 === 0 && v.canvas[0].y1 === 0);
  assert.strictEqual(rules.length, 2, "a rule between jobs, and none before the first");
  assert.ok(rules[0].canvas[0].x2 <= XP_PAGE_W, "which stays inside the page rather than overhanging it");
  const cardNodes = xpCardNodes(rows[0], exportAllFields(), true);
  assert.ok(cardNodes.length > 8, "a job is a run of nodes now, not one block");
  assert.strictEqual(cardNodes.filter(x => x.unbreakable).length, 1);
  const allCardText = strings(def.content).join(" | ");
  ["Ring before delivery", "Chase the glass supplier", "7000 CASEMENT", "05-Jan-2026",
   "not yet", "Windows colour: White", "Before 12 Nov"].forEach(bit =>
    assert.ok(allCardText.indexOf(bit) >= 0, "the job should still show: " + bit));
  assert.ok(strings(heads[1]).join(" | ").indexOf("URGENT") >= 0, "the urgent job says so in a word");
  pass("job cards: an unbreakable head, a body that flows, one page break, a rule between jobs");

  /* the stacked bar: proportional widths, and every segment labelled */
  const segs = xpBarSegments([{ key: "f", n: 24, colour: "#2A78D6" }, { key: "s", n: 18, colour: "#C25525" },
                              { key: "t", n: 6, colour: "#14805A" }], 300);
  assert.deepStrictEqual(segs.map(s => s.w), [150, 112.5, 37.5], "widths are the numbers, to scale");
  assert.strictEqual(segs.reduce((a, s) => a + s.w, 0), 300);
  const zero = xpBarSegments([{ key: "f", n: 5, colour: "#1" }, { key: "s", n: 5, colour: "#2" },
                              { key: "t", n: 0, colour: "#3" }], 300);
  assert.deepStrictEqual(zero.map(s => s.key), ["f", "s"], "a zero is left out rather than drawn as a sliver");
  const tiny = xpBarSegments([{ key: "f", n: 200, colour: "#1" }, { key: "s", n: 1, colour: "#2" }], 300, 16);
  assert.strictEqual(tiny[1].w, 16, "a very small segment still gets room for its number");
  assert.ok(tiny[0].w < 300 - 16 + 0.001 && tiny[0].w > 250, "and the big one gives up exactly that much");

  const canvases = collect(def.content, v => Array.isArray(v.canvas));
  const barShapes = canvases.map(c => c.canvas).find(sh => sh.length === 3 && sh.every(x => x.type === "rect"));
  assert.ok(barShapes, "the product bar is drawn with canvas rectangles");
  assert.deepStrictEqual(barShapes.map(s => s.w), [150, 112.5, 37.5]);
  assert.deepStrictEqual(barShapes.map(s => s.x), [0, 150, 262.5], "stacked, not overlapping");
  assert.deepStrictEqual(barShapes.map(s => s.color), ["#2A78D6", "#C25525", "#14805A"]);
  ["F 24", "S 18", "T 6"].forEach(l => assert.ok(strings(def.content).indexOf(l) >= 0,
    "each segment is labelled with its own number: " + l));
  pass("the F/S/T bar is canvas rectangles sized to the numbers, every segment numbered");

  const glassBlocks = collect(def.content, v => Array.isArray(v.canvas) && v.canvas.length === 1 && v.canvas[0].h === 20);
  assert.ok(glassBlocks.length >= 2, "one block per glass type");
  assert.ok(glassBlocks[0].canvas[0].w > glassBlocks[1].canvas[0].w, "8 units is a wider block than 6");
  assert.strictEqual(glassBlocks[0].canvas[0].color, "#FFE699", "and gold, because it is done");
  const bars = collect(def.content, v => Array.isArray(v.canvas) && v.canvas.length === 2 && v.canvas[0].h === 8);
  assert.ok(bars.length >= 4, "a progress bar for windows, doors, glass and each product");
  assert.strictEqual(bars[0].canvas[1].w, bars[0].canvas[0].w, "windows are 12 of 12: the bar is full");
  assert.ok(allCardText.indexOf("12 of 12  ·  done") >= 0, '"done of total" is written out too');
  assert.ok(allCardText.indexOf("8 of 14") >= 0, "glass is 8 of 14");
  pass("glass blocks are sized by quantity and checkpoints get a bar and a count");

  const hdr = def.header(2, 5), ftr = def.footer(2, 5);
  const hdrText = strings(hdr).join(" | ");
  assert.ok(hdrText.indexOf("Costello Production") >= 0, "the company name comes from the page, not from the code");
  assert.ok(hdrText.indexOf("Production export") >= 0);
  assert.ok(hdrText.indexOf("07-Sep-2026") >= 0);
  assert.strictEqual(collect(hdr, v => !!v.image).length, 0, "no logo file, no image node - and no error");
  const ftrText = strings(ftr).join(" | ");
  assert.ok(ftrText.indexOf("Page 2 of 5") >= 0);
  assert.ok(ftrText.indexOf("Exported by Pat Exporter") >= 0, "by name, not by address");
  const withLogo = buildDocDefinition(rows, Object.assign({}, pdfOpts, { logo: "data:image/png;base64,iVBORw0KGgo=" }));
  const logoNodes = collect(withLogo.header(1, 1), v => !!v.image);
  assert.strictEqual(logoNodes.length, 1, "when the file is there it goes in the header");
  assert.strictEqual(logoNodes[0].image, "data:image/png;base64,iVBORw0KGgo=");
  pass("header and footer are pdfmake functions returning plain content; the logo is optional");

  /* ---- 12. the table layout is a summary, not the whole sheet ---- */
  const wide = mkJob({ id: "R1005", cust: "Customer Five", area: "Clare", blk: 4, wnd: 3, drs: 1,
    glass: { tg: 1, dg: 2, tuff: 3, "not tuff": 4, arch: 5, astragal: 6, fancy: 7, extra: 8 },
    prods: (function () { const p = []; for (let i = 1; i <= 21; i++) p.push({ n: "profile " + i, f: i, s: i, t: i, st: [] }); return p; })() });
  const wideRows = exportRows([wide], fields, CTX());
  assert.strictEqual(exportColumns(wideRows, fields).length, 21 * 3 + 8 + 19,
    "the real sheet really is that wide in the workbook: 90 columns, Flag included");
  const tcols = exportTableColumns(fields);
  assert.deepStrictEqual(tcols.map(c => c.name), ["Job no", "Customer", "County", "Section", "Ready", "Urgent",
    "WND", "DRS", "Components F/S/T", "Glass units", "Sold", "Ready to print", "Sent to floor", "Comments"],
    "the table layout stays at fourteen, whatever the products do");
  const tw = exportTableWidths(tcols);
  assert.strictEqual(tw.length, 14);
  assert.ok(Math.min.apply(null, tw) >= 28, "and no column is narrower than 28pt (" + Math.min.apply(null, tw) + ")");
  assert.ok(tw.reduce((a, w) => a + w, 0) + 14 * 8 <= XP_LAND_W, "the row fits across A4 landscape");
  assert.throws(() => exportTableWidths(new Array(40).fill({ w: 30 })), /cannot fit 40 columns/,
    "and asking for far too many says so rather than drawing 5pt columns");
  assertNoHeaderWords(tcols.map(c => c.name), "the PDF table's headings");

  const tableDef = buildDocDefinition(wideRows.concat(rows), Object.assign({}, pdfOpts, { layout: "table" }));
  assert.strictEqual(tableDef.pageOrientation, "landscape");
  const tbl = collect(tableDef.content, v => v.table && v.table.headerRows === 1 && v.table.body.length > 3)[0];
  assert.deepStrictEqual(tbl.table.body[0].map(c => c.text), tcols.map(c => c.name));
  assert.deepStrictEqual(tbl.table.widths, tw);
  assert.strictEqual(tbl.table.body[1][8].text, "231/231/231", "the F/S/T totals in one cell instead of 63 columns");
  assert.strictEqual(tbl.table.body[1][9].text, "36", "and every glass unit added up in one more");
  assert.strictEqual(tbl.table.body[4][10].text, "Before 12 Nov", "a text date still reads as what it says");
  assert.strictEqual(tbl.table.body[2][1].fillColor, "#FAF9F7", "zebra rows, set per cell so the definition stays plain data");
  assert.strictEqual(tbl.table.body[1][1].fillColor, undefined, "the row above it is not shaded");
  assert.strictEqual(tbl.table.body[3][0].fillColor, "#E8F5E3", "a meaning-carrying fill wins over the zebra");
  assert.strictEqual(tbl.table.body[3][0].color, "#C62828", "and the urgent job number is still red");
  assert.strictEqual(tbl.table.body[3][0].bold, true);
  const groupedTable = buildDocDefinition(rows, Object.assign({}, pdfOpts, { layout: "table", filters: F({ groupBySection: true }) }));
  assert.ok(strings(groupedTable.content).join(" | ").indexOf("In production — 2 jobs") >= 0,
    "grouped: a section heading row across the table");
  pass("the table layout: fourteen readable columns whatever the sheet does, landscape, zebra, headings");

  const everyString = strings(def.content)
    .concat(strings(def.header(1, 1))).concat(strings(def.footer(1, 1)))
    .concat(strings(def.info)).concat(strings(tableDef.content));
  assertNoValues(everyString, "the document definition");
  assert.ok(everyString.length > 200, "and the scan really did walk the whole thing");
  pass("no phone number and no eircode anywhere in either PDF layout");
  noNet("building the PDF");

  /* ---- 13. the Export window ---- */
  vm.runInThisContext("XSTATE = null;");
  renderExportWindow();
  const win = el("#xhost").innerHTML;
  ["Export", "Excel workbook", "PDF", "What I see now", "Ticked jobs (0)", "Sections…",
   "Ready to deliver", "Urgent", "This week", "Last 7 days", "Last 30 days",
   "Sort &amp; group", "Presets", "Download"].forEach(bit =>
    assert.ok(win.indexOf(bit) >= 0, "the window should show: " + bit));
  assert.strictEqual(win.indexOf("&amp;amp;"), -1, "and nothing is escaped twice");
  assert.ok(win.indexOf('data-xtog="fields" value="job"') >= 0, "the field picker is there");
  assert.strictEqual(win.indexOf('value="eir"'), -1, "eircode is not offered as a field");
  assert.strictEqual(el("#xcount").textContent, "3 jobs match", "the live count");
  assert.strictEqual(el("#xdl").disabled, false);
  assert.ok(win.indexOf('data-xopen="filters" open') >= 0, "the filters start expanded");
  assert.ok(win.indexOf('data-xopen="presets">') >= 0, "and the presets start folded away");
  XSTATE.open.presets = true;
  renderExportWindow();
  assert.ok(el("#xhost").innerHTML.indexOf('data-xopen="presets" open') >= 0,
    "a group the person opened stays open when a chip redraws the window");
  pass("the window renders, with the count and an enabled Download");

  XSTATE.f.ready = true; xpUpdateCount();
  assert.strictEqual(el("#xcount").textContent, "1 job match");
  XSTATE.f.county = ["Cork"]; xpUpdateCount();
  assert.strictEqual(el("#xcount").textContent, "0 jobs match");
  assert.strictEqual(el("#xdl").disabled, true, "nothing to download, so the button is off");
  XSTATE.f.ready = null; XSTATE.f.county = []; xpUpdateCount();
  assert.strictEqual(el("#xcount").textContent, "3 jobs match");
  XSTATE.fields = {}; xpUpdateCount();
  assert.strictEqual(el("#xcount").textContent, "3 jobs match — no fields ticked",
    "with nothing ticked the count says why Download is off, instead of ExcelJS throwing later");
  assert.strictEqual(el("#xdl").disabled, true);
  XSTATE.fields = { cp: true }; xpUpdateCount();
  assert.strictEqual(el("#xdl").disabled, false, "checkpoints alone is a workbook worth making");
  XSTATE.format = "pdf"; XSTATE.layout = "table"; xpUpdateCount();
  assert.strictEqual(el("#xdl").disabled, true, "but there is no table to draw from them");
  assert.ok(el("#xcount").textContent.indexOf("no columns for the table") > 0);
  XSTATE.format = "xlsx"; XSTATE.layout = "cards"; XSTATE.fields = exportAllFields(); xpUpdateCount();
  pass("the count follows every change and Download turns itself off with the reason beside it");

  xpSet("f.cp.win", "done"); assert.strictEqual(XSTATE.f.cp.win, "done");
  xpSet("f.cp.win", ""); assert.strictEqual(XSTATE.f.cp.win, null, '"Any" clears it');
  xpSet("f.ready", "false"); assert.strictEqual(XSTATE.f.ready, false);
  xpSet("f.ready", "");
  xpToggleList("sections", "4", true);
  assert.deepStrictEqual(XSTATE.f.sections, [4]);
  assert.deepStrictEqual(XSTATE.f.sectionNames, ["In production"], "the names follow, for the summary and the filename");
  xpToggleList("sections", "4", false);
  assert.deepStrictEqual(XSTATE.f.sections, []);
  xpToggleList("fields", "cust", false);
  assert.strictEqual(XSTATE.fields.cust, undefined);
  xpToggleList("fields", "cust", true);
  assert.strictEqual(XSTATE.fields.cust, true);
  xpDatePreset("sold", "30");
  assert.ok(XSTATE.f.dates.sold.from && XSTATE.f.dates.sold.to, "a date preset fills both ends");
  assert.strictEqual(XSTATE.f.dates.sold.to, xpIsoDate(new Date()));
  xpDatePreset("sold", "week");
  const monday = new Date();
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
  assert.strictEqual(XSTATE.f.dates.sold.from, xpIsoDate(monday), "this week starts on Monday");
  assert.strictEqual(XSTATE.f.dates.sold.to, xpIsoDate(new Date()));
  xpDatePreset("sold", "clear");
  assert.deepStrictEqual(XSTATE.f.dates.sold, { from: null, to: null });
  pass("the controls set exactly what they say they set, this week included");

  delete mem.cw_exportpresets;
  renderExportWindow();
  el("#xpname").value = "Cork run";
  el("#xpsave").onclick();
  assert.deepStrictEqual(Object.keys(presetsLoad()), ["Cork run"]);
  XSTATE.f.urgent = true;
  el("#xpsel").value = "Cork run";
  el("#xpload").onclick();
  assert.strictEqual(XSTATE.f.urgent, null, "loading a preset puts the choices back as they were saved");
  el("#xpsel").value = "Cork run";
  el("#xpdel").onclick();
  assert.deepStrictEqual(presetsLoad(), {});
  pass("Save as… / Load / Delete, from the window");

  el("#xhost").removed = false;
  assert.ok((HANDLERS.keydown || []).length, "the page listens for keys");
  HANDLERS.keydown.forEach(fn => fn({ key: "Escape", preventDefault() {} }));
  assert.strictEqual(el("#xhost").removed, true, "Escape closes the Export window");
  el("#xhost").removed = false;
  noNet("the window");
  pass("Escape closes the window");

  /* ---- 14. what a Download actually does ---- */
  LOGS.length = 0; SEQ.length = 0; TOASTS.length = 0;
  vm.runInThisContext("XSTATE = null;");
  renderExportWindow();
  XSTATE.format = "xlsx";
  const xlsxName = await el("#xdl").onclick();
  assert.strictEqual(xlsxName, "Production export - All jobs - " + xpIsoDate(new Date()) + ".xlsx");
  assert.deepStrictEqual(SEQ, ["URL.createObjectURL", "anchor.click " + xlsxName, "toast"],
    "the workbook is built in the tab and handed over through an object URL");
  assert.strictEqual(FETCHES, 0, "and nothing went near the network");
  assert.strictEqual(LOGS.length, 1, "exactly one log line per export");
  assert.deepStrictEqual(LOGS[0], { who: "exporter@example.test", job: "(export)", what: "Export",
    from: "Excel · 3 jobs", to: filtersSummary(XSTATE.f, SECTIONS) });
  assert.ok(TOASTS[0].m.indexOf("Downloading") === 0);
  const anchor = MADE.filter(e => e.tag === "a").pop();
  assert.strictEqual(anchor.download, xlsxName);
  assert.strictEqual(anchor.href, "blob:test");
  pass("Download (Excel): build, hand over, log one line, say so");

  LOGS.length = 0; SEQ.length = 0; TOASTS.length = 0;
  XSTATE.format = "pdf";
  XSTATE.f.urgent = true;
  const pdfName = await xpDownload();
  assert.strictEqual(pdfName, "Production export - Urgent - " + xpIsoDate(new Date()) + ".pdf");
  assert.deepStrictEqual(SEQ, ["pdfMake.createPdf", "pdf.download " + pdfName, "toast"],
    "pdfMake.createPdf(def).download(name) - one click, no print dialog");
  assert.strictEqual(FETCHES, 0);
  assert.deepStrictEqual(LOGS[0], { who: "exporter@example.test", job: "(export)", what: "Export",
    from: "PDF · 1 job", to: filtersSummary(XSTATE.f, SECTIONS) });
  assert.ok(LOGS[0].to.indexOf("Urgent only") >= 0, "the log line says which filters were on");
  assert.strictEqual(collect(PDFDEF.content, v => v.unbreakable === true).length, 1, "one job, for the one urgent job");
  assert.ok(strings(PDFDEF.content).join(" | ").indexOf("Exported by") < 0);
  assert.ok(strings(PDFDEF.header(1, 1)).join(" | ").indexOf("Costello Production") >= 0);
  assertNoValues(strings(PDFDEF.content).concat(strings(PDFDEF.footer(1, 1))), "the downloaded PDF");
  assert.strictEqual(strings(PDFDEF.footer(1, 1)).join(" | ").indexOf("@"), -1,
    "the footer names the person who made the file, not their address");
  pass("Download (PDF): the same, through pdfmake, with the engine loaded on demand");

  /* pdfmake and its fonts are 2.2 MB and are not in the page: the first PDF
     asks for them, in that order, and says so on the button while it waits */
  LOGS.length = 0; SEQ.length = 0; TOASTS.length = 0;
  delete global.pdfMake;
  vm.runInThisContext("XPDFLOAD = null;");
  const before = MADE.length;
  const pending = xpDownload();
  const scripts = () => MADE.slice(before).filter(e => e.tag === "script");
  assert.strictEqual(el("#xdl").textContent, "Loading PDF engine…", "the button says what it is waiting for");
  assert.deepStrictEqual(scripts().map(e => e.src), ["vendor/pdfmake.min.js"]);
  scripts()[0].onload();
  await new Promise(r => setTimeout(r, 0));
  assert.deepStrictEqual(scripts().map(e => e.src), ["vendor/pdfmake.min.js", "vendor/vfs_fonts.js"],
    "the fonts go in after pdfmake, never before it");
  assert.strictEqual(scripts()[1].async, false, "and in that order, not whenever they happen to arrive");
  global.pdfMake = PDFSTUB;                 // in a browser the first script would have defined it
  scripts()[1].onload();
  const lazyName = await pending;
  assert.ok(lazyName && lazyName.indexOf(".pdf") > 0);
  assert.deepStrictEqual(SEQ, ["pdfMake.createPdf", "pdf.download " + lazyName, "toast"]);
  assert.strictEqual(el("#xdl").textContent, "Download", "and the button goes back to itself afterwards");
  assert.strictEqual(LOGS.length, 1);
  assert.strictEqual(FETCHES, 0, "the scripts are page assets, not something fetched from anywhere");
  vm.runInThisContext("XPDFLOAD = null;");
  pass("the PDF engine is injected on the first PDF, pdfmake before its fonts");

  LOGS.length = 0; SEQ.length = 0;
  XSTATE.f.q = "nothing matches this";
  xpUpdateCount();
  assert.strictEqual(el("#xdl").disabled, true);
  assert.strictEqual(await xpDownload(), null, "and pressing it anyway does nothing at all");
  assert.deepStrictEqual(SEQ, []);
  assert.deepStrictEqual(LOGS, [], "no jobs, no log line");
  XSTATE.f.q = ""; XSTATE.fields = {};
  assert.strictEqual(await xpDownload(), null, "nor with jobs but no fields");
  assert.deepStrictEqual(SEQ, []);
  assert.deepStrictEqual(LOGS, []);
  pass("with nothing selected, and with no fields ticked, there is no file and no log line");

  /* ---- 15. real PDFs, rendered by the vendored pdfmake and read back ----
     The bug this guards: a job taller than a page used to be dropped whole,
     leaving an empty page behind it. Nothing short of rendering finds that. */
  /* the bundle only claims the global name when nothing holds it yet
     (`void 0 === g.pdfMake && (g.pdfMake = p)`), so the stub stands aside */
  delete global.pdfMake;
  run("vendor/pdfmake.min.js");
  run("vendor/vfs_fonts.js");
  assert.ok(global.pdfMake && global.pdfMake !== PDFSTUB, "the real pdfmake is loaded now");
  assert.strictEqual(typeof global.pdfMake.createPdf, "function");
  const engine = xpPdfReady();
  const render = def2 => new Promise((ok, no) => {
    def2.compress = false;                 // so the text streams can be read back
    try { engine.createPdf(def2).getBuffer(b => ok(b), {}); } catch (e) { no(e); }
  });

  const bigJobs = [TALL, A, B, C];
  const bigRows = exportRows(bigJobs, fields, Object.assign(CTX(), { all: bigJobs, view: bigJobs }));
  const cardsBuf = await render(buildDocDefinition(bigRows, pdfOpts));
  const cardsText = flat(pdfText(cardsBuf));
  assert.ok(cardsText.indexOf("Production export") >= 0,
    "the reader really does decode the page text (it found the cover title)");
  ["R1004", "R1001", "R1002", "R1003"].forEach(id =>
    assert.ok(cardsText.indexOf(id) >= 0, "job " + id + " is missing from the rendered PDF"));
  assert.ok(cardsText.indexOf("Note number 30 about") >= 0,
    "and the thirtieth comment of the tall job is in there too, not cut off");
  assert.ok(cardsText.indexOf("PROFILE 20") >= 0, "with all twenty of its products");
  assert.ok(pdfPages(cardsBuf) >= 3, "over several pages");
  const runs = pdfRunsPerPage(cardsBuf);
  assert.strictEqual(runs.length, pdfPages(cardsBuf), "every page has content of its own");
  console.log("      (text runs per page: " + runs.join(", ") + ")");
  assert.ok(Math.min.apply(null, runs) > 30,
    "no page is just the header and the footer with a dropped card behind it: " + runs.join(", "));
  pass("cards: a job with 20 products and 30 comments renders and nothing is dropped");

  const tableBuf = await render(buildDocDefinition(bigRows, Object.assign({}, pdfOpts, { layout: "table" })));
  const tableText = flat(pdfText(tableBuf));
  ["R1004", "R1001", "R1002", "R1003"].forEach(id =>
    assert.ok(tableText.indexOf(id) >= 0, "the table layout is missing job " + id));
  ["Jobno", "Customer", "ComponentsF/S/T", "Glassunits", "Readytoprint", "Senttofloor", "Comments"]
    .forEach(h => assert.ok(squash(tableText).indexOf(h) >= 0, "the table layout is missing the column " + h));
  assert.ok(squash(tableText).indexOf("210/190/60") >= 0,
    "the twenty-product job's F/S/T totals fit in one cell instead of sixty-three columns");
  assert.ok(squash(tableText).indexOf("Before12Nov") >= 0, "and a text date still reads as what it says");
  pass("table: every job on one landscape row, with readable headings");

  const groupBuf = await render(buildDocDefinition(bigRows,
    Object.assign({}, pdfOpts, { filters: F({ groupBySection: true }) })));
  const groupText = flat(pdfText(groupBuf));
  ["R1004", "R1001", "R1002", "R1003"].forEach(id =>
    assert.ok(groupText.indexOf(id) >= 0, "grouped: job " + id + " is missing"));
  const gTableBuf = await render(buildDocDefinition(bigRows,
    Object.assign({}, pdfOpts, { layout: "table", filters: F({ groupBySection: true }) })));
  const gTableText = flat(pdfText(gTableBuf));
  ["Inproduction", "Readytofit", "R1004", "R1002"].forEach(bit =>
    assert.ok(squash(gTableText).indexOf(bit) >= 0, "grouped table: missing " + bit));
  assertNoValues([cardsText, tableText, groupText, gTableText], "a rendered PDF");
  pass("grouped: both layouts render with every job and every section heading present");

  /* ---- 15b. the John print sheet ----
     A fixed layout printed from "Production (2)" and from nothing else. It is
     the one export in the app that carries a phone number, sanctioned by the
     owner on 2026-09-09; the eircode is in none of it, and the Default
     template still carries neither.

     Production (2) deliberately disagrees with the Production model here - a
     different phone, area, ready date and note on R1001 - because the whole
     point of the amendment is that John's paper comes from John's sheet. */
  const JOHN2 = [
    { id: "R1002", section: "Ready to fit", ready: "2026-04-02", cust: "Customer Two",
      phone: PHONE, area: "Kerry", wnd: 5, drs: 0, notes: "Collect on Friday",
      fillHex: "FFE699", inkHex: "FF0000", seq: 0 },
    { id: "R1001", section: "In production", ready: "2026-02-10", cust: "Customer One",
      phone: JPHONE, area: "West Clare", wnd: 12, drs: 4, notes: "Check the cill height",
      fillHex: "FFFF00", inkHex: "FF3399", seq: 1 },
    { id: "R1003", section: "In production", ready: "", cust: "Customer Three",
      phone: PHONE, area: "Cork", wnd: 0, drs: 2, notes: "", fillHex: "", inkHex: "", seq: 2 }
  ];
  JOHN2.sections = ["Ready to fit", "In production"];
  const JCTX = () => Object.assign(CTX(), { all: [A, B, C, TALL], john: JOHN2 });

  const JN = { R1001: "Leave at the side gate", R9999: "a note for a job not in this print" };
  const johnRows = exportJohnRows(["R1003", "R1001", "R1002"], JN, JCTX());
  assert.deepStrictEqual(johnRows.map(r => r.kind === "section" ? "== " + r.section : r.id),
    ["== Ready to fit", "R1002", "== In production", "R1001", "R1003"],
    "Production (2)'s own order: section by section, and inside a section by its own row order");
  const jr = johnRows.filter(r => r.kind === "job");
  assert.deepStrictEqual(jr.map(r => r.phone), [PHONE, JPHONE, PHONE],
    "every row carries a phone number - this template and no other");
  assert.strictEqual(jr[1].phone, JPHONE, "and it is Production (2)'s number, not the Production model's");
  assert.strictEqual(jr[1].area, "West Clare", "the area comes from Production (2) too");
  assert.strictEqual(jr[1].ready, "10-Feb", "Ready to print is dd-MMM, from that sheet's date");
  assert.strictEqual(jr[0].ready, "02-Apr");
  assert.strictEqual(jr[2].ready, "", "and a row with no ready date has an empty cell");
  assert.strictEqual(jr[1].notes, "Check the cill height · Leave at the side gate",
    "Notes: Production (2)'s Brendan note, then the dashboard's print note");
  assert.strictEqual(jr[2].notes, "", "a row with neither has an empty Notes cell");
  assert.strictEqual(jr[0].fill, "#FFE699", "R1002 wears its own gold");
  assert.strictEqual(jr[1].fill, "#FFFF00", "R1001 its own yellow");
  assert.strictEqual(jr[2].fill, "", "and R1003 has no fill on that sheet");
  assert.strictEqual(jr[0].colour, "#FF0000", "the text colour is that sheet's own hex");
  assert.strictEqual(jr[0].flag, "urgent", "with the word it stands for");
  assert.strictEqual(jr[1].colour, "#FF3399");
  assert.strictEqual(jr[1].flag, "trade");
  assert.strictEqual(jr[2].colour, "", "a black row stays black");
  assert.ok(jr.every(r => !r.missing));
  assert.strictEqual(xpJohnDate("Before 12 Nov"), "Before 12 Nov",
    "what the office wrote in words in a date cell is printed as written");
  const johnStrings = strings(johnRows);
  johnStrings.forEach(x => assert.strictEqual(String(x).toLowerCase().indexOf(EIR.toLowerCase()), -1,
    'the John rows leaked the eircode: "' + x + '"'));
  assert.ok(johnStrings.some(x => String(x).indexOf(JPHONE) >= 0),
    "and they really do carry the phone number, or this test is proving nothing");
  pass("John rows come from Production (2): its order, its dividers, its colours, its phone and note");

  /* a job the sheet does not have */
  const withMissing = exportJohnRows(["R1001", "R1004"], {}, JCTX());
  const gone = withMissing.filter(r => r.kind === "job").find(r => r.id === "R1004");
  assert.ok(gone, "a job that is not on Production (2) is still printed");
  assert.strictEqual(gone.missing, true);
  assert.strictEqual(gone.fill, "", "with no fill");
  assert.strictEqual(gone.colour, "", "and no colour at all: that sheet said nothing about it");
  assert.ok(gone.notes.indexOf(XP_JOHN_MISSING) >= 0, "and its Notes say so: " + gone.notes);
  assert.strictEqual(gone.cust, "Customer Four", "the rest of it comes from the Production model");
  assert.deepStrictEqual(withMissing.filter(r => r.kind === "section").map(r => r.section),
    ["In production"], "and it lands in its own Production section, not a new one");
  assert.deepStrictEqual(exportJohnRows([], {}, JCTX()), [], "nothing chosen, nothing to print");
  assert.deepStrictEqual(exportJohnRows(["R1001"], {}, Object.assign(CTX(), { john: [] }))
    .filter(r => r.kind === "job").map(r => r.missing), [true],
    "and with no Production (2) sheet at all, every job prints as one that is not on it");
  pass("a job Production (2) does not have is printed from the model, uncoloured and marked");

  const jwb = await roundTrip(buildJohnWorkbook(johnRows, { who: "Pat Exporter", when: when }));
  assert.deepStrictEqual(jwb.worksheets.map(w => w.name), ["John print sheet"], "one sheet, named for the print");
  const jws = jwb.getWorksheet("John print sheet");
  const jcells = r => [1, 2, 3, 4, 5, 6, 7, 8].map(c => String(jws.getCell(r, c).value == null ? "" : jws.getCell(r, c).value));
  assert.deepStrictEqual(jcells(1),
    ["Job no", "Ready to print", "Customer", "Phone no", "Area", "QUANTITY", "QUANTITY", "Notes"],
    "the sheet's own two-row header, with QUANTITY over the two counts");
  assert.strictEqual(jcells(2)[5], "Wnd");
  assert.strictEqual(jcells(2)[6], "Drs");
  assert.deepStrictEqual([1, 2, 3, 4, 5, 6, 7, 8].map(c => jws.getColumn(c).width), [11, 10, 20, 14, 12, 6, 5, 60],
    "and the widths of the sheet it is a copy of");
  assert.strictEqual(jcells(3)[0], "Ready to fit", "row 3 is the first section divider");
  assert.strictEqual((jws.getCell(3, 1).font || {}).bold, true, "bold");
  assert.strictEqual(((jws.getCell(3, 1).fill || {}).fgColor || {}).argb, "FFEFECE7", "on light grey");
  assert.deepStrictEqual(jcells(4).slice(0, 5), ["R1002", "02-Apr", "Customer Two", PHONE, "Kerry"]);
  assert.strictEqual(jws.getCell(4, 6).value, 5, "Wnd is a number");
  assert.strictEqual(jws.getCell(4, 7).value, null, "and a zero count is left blank, as the sheet leaves it");
  assert.strictEqual(((jws.getCell(4, 1).fill || {}).fgColor || {}).argb, "FFFFE699", "R1002's row is gold");
  assert.strictEqual(((jws.getCell(4, 1).font || {}).color || {}).argb, "FFFF0000", "and its text is red");
  assert.strictEqual(((jws.getCell(6, 1).font || {}).color || {}).argb, "FFFF3399", "R1001's is pink");
  assert.strictEqual(((jws.getCell(6, 1).fill || {}).fgColor || {}).argb, "FFFFFF00", "on yellow");
  assert.ok((jws.getCell(4, 1).border || {}).bottom, "every cell is bordered");
  assert.strictEqual(jws.pageSetup.orientation, "landscape");
  assert.strictEqual(jws.pageSetup.fitToWidth, 1, "fitted to one page across");
  assert.strictEqual(jws.pageSetup.printTitlesRow, "1:2", "with the header repeated on every printed page");
  assert.strictEqual(jws.views[0].ySplit, 2, "and frozen under the two header rows on screen");
  const jstr = sheetStrings(jwb);
  jstr.forEach(x => assert.strictEqual(String(x).toLowerCase().indexOf(EIR.toLowerCase()), -1,
    "the John workbook leaked the eircode"));
  assert.ok(jstr.some(x => String(x) === JPHONE), "and it carries Production (2)'s phone number");
  pass("the John workbook: one sheet, the two-row header, the sheet's widths, colours and print setup");

  const missWb = await roundTrip(buildJohnWorkbook(withMissing, { who: "Pat Exporter", when: when }));
  const missWs = missWb.getWorksheet("John print sheet");
  let missRow = 0;
  missWs.eachRow({ includeEmpty: false }, (row, i) => { if (String(row.getCell(1).value) === "R1004") missRow = i; });
  assert.ok(missRow, "the missing job is a row of its own in the file");
  assert.strictEqual(((missWs.getCell(missRow, 1).font || {}).color || {}).argb, "FF6D6A62",
    "printed grey, so the eye can see its line came from somewhere else");
  assert.strictEqual((missWs.getCell(missRow, 1).font || {}).italic, true);
  assert.strictEqual(((missWs.getCell(missRow, 1).fill || {}).fgColor || {}).argb, undefined,
    "and with no fill of any kind");
  pass("a job not on Production (2) prints grey and italic, with no colour borrowed from anywhere");

  assert.strictEqual(exportJohnFilename("xlsx", when), "John print sheet 2026-09-07.xlsx");
  assert.strictEqual(exportJohnFilename("pdf", when), "John print sheet 2026-09-07.pdf");
  assert.strictEqual(exportLogFrom("xlsx", 3, "john"), "Excel · 3 jobs · John print sheet · with phone numbers");
  assert.strictEqual(exportLogFrom("pdf", 1, "john"), "PDF · 1 job · John print sheet · with phone numbers");
  assert.strictEqual(exportLogFrom("xlsx", 3), "Excel · 3 jobs",
    "and the Default template's log line is exactly what it always was");
  assert.strictEqual(exportLogFrom("xlsx", 3, "default"), "Excel · 3 jobs");
  pass("the file name, and a log line that says in words that this file has phone numbers in it");

  const johnBuf = await render(buildJohnDoc(johnRows,
    { who: "Pat Exporter", when: when, company: "Costello Production" }));
  const johnText = flat(pdfText(johnBuf));
  ["R1001", "R1002", "R1003"].forEach(id =>
    assert.ok(johnText.indexOf(id) >= 0, "the John PDF is missing job " + id));
  ["Readytofit", "Inproduction", "Jobno", "Readytoprint", "Phoneno", "QUANTITY", "Wnd", "Drs", "Notes"]
    .forEach(bit => assert.ok(squash(johnText).indexOf(bit) >= 0, "the John PDF is missing " + bit));
  assert.ok(squash(johnText).indexOf(squash(JPHONE)) >= 0, "and the phone number is on the page");
  assert.ok(squash(johnText).indexOf("Printed") >= 0, "with the printed-by footer");
  assert.strictEqual(johnText.toLowerCase().indexOf(EIR.toLowerCase()), -1, "the John PDF leaked the eircode");
  pass("the John PDF renders on A4 landscape with every job, every divider and the header band");

  /* the Default template, with the John one sitting in the same file: still
     neither a phone number nor an eircode, anywhere */
  const cleanRows = exportRows([A, B, C, TALL], exportAllFields(), CTX());
  assertNoValues(strings(cleanRows), "a Default row");
  assertNoValues(sheetStrings(await roundTrip(buildWorkbook(cleanRows, opts))), "the Default workbook");
  assertNoValues([flat(pdfText(await render(buildDocDefinition(cleanRows, pdfOpts))))], "the Default PDF");
  assert.strictEqual(exportColumns(cleanRows, exportAllFields()).filter(c => c.key === "phone").length, 0);
  assert.strictEqual(JSON.stringify(cleanRows).indexOf(JPHONE), -1,
    "and Production (2)'s numbers do not reach a Default row either");
  pass("the Default template still carries no phone number and no eircode of any kind");

  /* the Flag column and chip in the Default export */
  const flagCell = exportCell(cleanRows[0], { key: "flag", kind: "flag" });
  assert.strictEqual(flagCell.text, "Trade order", "the word, always");
  assert.strictEqual(flagCell.colour, "#FF3399", "in the sheet's own colour");
  assert.strictEqual(exportCell(cleanRows[2], { key: "flag", kind: "flag" }).text, "",
    "and nothing at all for a black row");
  assert.strictEqual(exportCell({ flag: "hold" }, { key: "flag", kind: "flag" }).colour, "#1565C0",
    "a flag with no hex still prints in a readable stand-in colour");
  const withFlag = exportRows([A], ["job", "flag"], CTX());
  assert.deepStrictEqual(Object.keys(withFlag[0]), ["id", "flag", "flagHex"]);
  assert.deepStrictEqual(Object.keys(exportRows([A], ["job"], CTX())[0]), ["id"],
    "and no flag at all when the field is not ticked");
  pass("Flag is a Default field like any other: a word, its colour, and only when it is ticked");

  /* ---- 16. the standing assertions ---- */
  assert.strictEqual(FETCHES, 0, "no fetch() happened anywhere in this run");
  const keys = Object.keys(mem);
  assert.deepStrictEqual(keys.filter(k => k.indexOf("export") >= 0), ["cw_exportpresets"],
    "the only thing the export keeps in this browser is the presets someone named");
  const everything = keys.map(k => String(mem[k])).join(" | ");
  assertNoValues([everything], "localStorage");
  pass("no network, and nothing about a job left behind in the browser");

  console.log("\n" + n + " checks passed");
  process.exit(0);
})().catch(e => { console.error("FAIL", e); process.exit(1); });
