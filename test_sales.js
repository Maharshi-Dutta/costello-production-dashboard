/* Offline test of the Sales page's rules and writes (sales-core.js + the Sales
   primitives in graph.js) against a fake Production sheet served through a
   fake fetch(), and fake SharePoint list calls whose failures can be forced.
   docs/specs/2026-09-30-sales-page.md, "Tests to deliver".
   Run: node test_sales.js                                                    */
const fs = require("fs"), vm = require("vm"), assert = require("assert");

global.ExcelJS = { Workbook: function () {} };
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, CW_PAGE: "sales" };
vm.runInThisContext(fs.readFileSync(__dirname + "/parser.js", "utf8"), { filename: "parser.js" });
vm.runInThisContext(fs.readFileSync(__dirname + "/graph.js", "utf8"), { filename: "graph.js" });
vm.runInThisContext(fs.readFileSync(__dirname + "/sales-core.js", "utf8"), { filename: "sales-core.js" });
const CW = window.CW, S = window.SALESC;
CW._setToken(() => "t");
CW._setFile({ base: "/x/workbook", content: "/x/content", meta: "/x", siteId: "SITE" });

/* ---------- a tiny Excel (the same shape test_move.js uses) ---------- */
const N = 90;
const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
function mkRow(vals, fill, height) {
  const r = { v: new Array(N).fill(""), fill: new Array(N).fill(fill || ""), height: height || 30,
              font: [], nf: new Array(N).fill("General") };
  for (let c = 0; c < N; c++) r.font.push({ bold: c === 2, italic: false, size: 12, name: "Calibri", color: "#000000", underline: "None" });
  (vals || []).forEach((x, i) => { r.v[i] = x; });
  r.nf[3] = "d-mmm";
  return r;
}
/* a real-width job row: all 90 columns written, formulas, fills and fonts varying */
function wideRow(id) {
  const r = mkRow(["Unglazed, asap - see office", 9001, id, 46020, "", "", "", "", "Person A", "0871234567", "Cork"], "#FFE699", 30);
  for (let c = 11; c < N; c++) r.v[c] = c % 3 === 0 ? "=SUM(L5:M5)*" + c : c % 3 === 1 ? c * 7 : "word " + c;
  for (let c = 0; c < N; c++) { r.fill[c] = ["#FFE699", "#FFFF00", "#FFFFFF", "#D9D2E9"][c % 4];
    r.font[c] = { bold: c % 2 === 0, italic: c % 5 === 0, size: 11 + (c % 3), name: "Calibri", color: ["#000000", "#FF0000", "#00B050"][c % 3], underline: "None" };
    r.nf[c] = c % 4 === 0 ? "d-mmm" : c % 4 === 1 ? "0.00" : "General"; }
  return r;
}
let SHEET, LOG, FAIL_PATCH_VALUES, ON_READ, READS, FAIL_BORDERS, FETCHES = 0;
function reset() {
  SHEET = [null,
    mkRow(["", "", "", "", "", "", "", "", "", "", "8843"]),
    mkRow(["COMMENT", "OFFICE NO.", "JOB NO.", "", "", "", "", "", "CUSTOMER", "PHONE NO.", "AREA", "EIRCODE", "WINDOWS COLOUR"]),
    mkRow(["", "", "", "SOLD"]),
    mkRow(["Below is orders ready and customers won't take"]),
    mkRow(["", 7785, "R0001", 46000, "", "", "", "", "Person A", "086", "Cork"], "#FFE699", 30.75),
    mkRow(["Below is collect or supply only orders"]),
    mkRow(["asap", 8001, "C0003", 46010, "", "", "", "", "Person B", "085", "Kerry"], "#FFE699", 30),
    mkRow(["Not sent to floor(No section) ="]),
    mkRow(["Unglazed", 9001, "R0005", 46020, "", "", "", "", "Person A", "089", "Cork"], "#FFE699", 30),
    mkRow(["on hold", 9002, "R0006", 46021, "", "", "", "", "Person B", "086", "Dublin"], "#FFE699", 30),
    mkRow(["Not sent to floor(No section) ="]),
    mkRow(["", 9101, "R0007", 46030, "", "", "", "", "Person A", "087", "Laois"], "#FFFFFF", 30),
    mkRow(["URGENT", 9102, "R0008", 46031, "", "", "", "", "Person B", "086", "Cork"], "#FFFFFF", 30),
    mkRow(["", "", 5293])];
  LOG = []; FAIL_PATCH_VALUES = 0; ON_READ = null; READS = 0; FAIL_BORDERS = 0;
}
const typeOf = v => v === "" || v == null ? "Empty" : typeof v === "number" ? "Double" : typeof v === "boolean" ? "Boolean" : "String";
function parseAddr(a) {
  let m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a);
  if (m) return { r1: +m[2], r2: +m[4], c1: colNum(m[1]), c2: colNum(m[3]) };
  m = /^([A-Z]+)(\d+)$/.exec(a);
  if (m) return { r1: +m[2], r2: +m[2], c1: colNum(m[1]), c2: colNum(m[1]) };
  m = /^(\d+):(\d+)$/.exec(a);
  if (m) return { r1: +m[1], r2: +m[2], c1: 1, c2: N, whole: true };
  throw new Error("bad address " + a);
}
const ok = body => ({ status: 200, body });
function route(method, url, body) {
  LOG.push({ m: method, u: url.replace(/^\/x\/workbook\/worksheets\('Production'\)\//, ""), b: body });
  const m = /range\(address='([^']+)'\)(.*)$/.exec(url);
  if (!m) return { status: 404, body: { error: "no route " + url } };
  const A = parseAddr(m[1]), rest = m[2];
  const rows = () => { const out = []; for (let r = A.r1; r <= A.r2; r++) out.push(SHEET[r] || mkRow([])); return out; };
  if (method === "GET") {
    if (rest.startsWith("/format/fill")) { const s = new Set(); rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) s.add(r.fill[c - 1]); }); return ok({ color: s.size === 1 ? [...s][0] : null }); }
    if (rest.startsWith("/format/font")) {
      const keys = ["bold", "italic", "size", "name", "color", "underline"], out = {};
      keys.forEach(k => { const s = new Set(); rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) s.add(r.font[c - 1][k]); }); out[k] = s.size === 1 ? [...s][0] : null; });
      return ok(out);
    }
    if (rest.startsWith("/format")) return ok({ rowHeight: SHEET[A.r1] ? SHEET[A.r1].height : 15 });
    /* the Sales locate's read of A1:K600: the sheet can be changed just before the nth one */
    if (A.r1 === 1 && A.r2 === 600 && A.c1 === 1 && A.c2 === 11) { READS++; if (ON_READ) ON_READ(READS); }
    const o = { values: [], valueTypes: [], numberFormat: [], formulas: [] };
    rows().forEach(r => {
      const v = [], t = [], nf = [], f = [];
      for (let c = A.c1; c <= A.c2; c++) { const x = r.v[c - 1]; const isF = typeof x === "string" && x.charAt(0) === "="; v.push(isF ? 0 : x); t.push(isF ? "Double" : typeOf(x)); nf.push(r.nf[c - 1]); f.push(x); }
      o.values.push(v); o.valueTypes.push(t); o.numberFormat.push(nf); o.formulas.push(f);
    });
    return ok(o);
  }
  if (method === "POST" && rest === "/insert") { assert(A.whole); SHEET.splice(A.r1, 0, mkRow([], "#FFE699", 30.75)); return ok({}); }
  if (method === "POST" && rest === "/delete") { assert(A.whole); SHEET.splice(A.r1, 1); return { status: 204, body: {} }; }
  if (method === "POST" && rest === "/format/fill/clear") { rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) r.fill[c - 1] = ""; }); return ok({}); }
  if (method === "PATCH") {
    const typed = x => (typeof x === "string" && x.charAt(0) === "'") ? x.slice(1) : (typeof x === "string" && /^\d+$/.test(x)) ? Number(x) : x;
    if (rest === "") {
      if (FAIL_PATCH_VALUES) { FAIL_PATCH_VALUES--; return { status: 400, body: { error: { code: "Boom" } } }; }
      const r = SHEET[A.r1];
      (body.formulas || body.values)[0].forEach((x, i) => { r.v[A.c1 - 1 + i] = typed(x); });   // Excel parses typed text
      if (body.numberFormat) body.numberFormat[0].forEach((x, i) => { r.nf[A.c1 - 1 + i] = x; });
      return ok({});
    }
    if (rest === "/format/fill") { rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) r.fill[c - 1] = body.color; }); return ok({}); }
    if (rest === "/format/font") { rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) r.font[c - 1] = Object.assign({}, r.font[c - 1], body); }); return ok({}); }
    if (rest.startsWith("/format/borders/") && FAIL_BORDERS === A.r1) return { status: 400, body: { error: { code: "BorderRefused" } } };
    if (rest === "/format" || rest.startsWith("/format/borders/")) { rows().forEach(r => { if (body.rowHeight) r.height = body.rowHeight; }); return ok({}); }
  }
  return { status: 404, body: { error: "no route " + method + " " + url } };
}
global.fetch = async (url, init) => {
  FETCHES++;
  const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  let res;
  if (path === "/$batch") res = ok({ responses: body.requests.map(q => Object.assign({ id: q.id }, route(q.method, q.url, q.body))) });
  else res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status, text: async () => JSON.stringify(res.body), arrayBuffer: async () => new ArrayBuffer(0) };
};

/* ---------- fake lists, on top of the real workbook primitives ---------- */
let ITEMS, NEXT, FAIL_ADD, BAD_READBACK, NOTES;
function resetLists() { ITEMS = {}; NEXT = 1; FAIL_ADD = false; BAD_READBACK = false; NOTES = []; }
const listAPI = {
  listAdd: async (name, fields) => {
    if (FAIL_ADD) throw new Error("list refused");
    const id = String(NEXT++); ITEMS[id] = { list: name, fields: Object.assign({}, fields) }; return { id };
  },
  listItem: async (name, id) => ITEMS[id] ? { id, fields: Object.assign({}, ITEMS[id].fields,
    BAD_READBACK ? { Row: String(ITEMS[id].fields.Row).slice(0, -5) } : {}) } : null,
  listPatch: async (name, id, fields) => { Object.assign(ITEMS[id].fields, fields); return {}; },
  listUpsert: async (name, title, fields) => { ITEMS["u" + title] = { list: name, fields: Object.assign({ Title: title }, fields) }; return {}; }
};
const fakeCW = Object.assign({}, CW, listAPI);
const fakeTmpl = () => ({ row: 0, height: 30, cells: Array.from({ length: N }, () => ({ h: null, v: null, wrap: false,
  bottom: { style: "thin", color: "#000000" }, left: null, right: null, fk: "a", flk: "g" })) });
const ctx = () => ({ CW: fakeCW, who: "Person A", note: (job, what, from, to) => NOTES.push([job, what, from, to]),
  tmplFor: () => fakeTmpl(), identCol: k => ({ cust: 9, phone: 10, area: 11, off: 2, eir: 12, colour: 13 })[k],
  hdrRow: () => 2, now: () => "2026-09-30T10:00:0" + (NEXT % 10) + ".000Z" });
const dupRow = id => { const r = JSON.parse(JSON.stringify(SHEET[rowOf(id)])); SHEET.splice(13, 0, r); };
const shiftDown = () => SHEET.splice(4, 0, mkRow(["someone inserted a row in Excel"]));
const jobsInOrder = () => SHEET.slice(1).filter(r => /^[A-Z]{1,2}\d{3,5}$/.test(String(r.v[2]))).map(r => r.v[2]);
const rowOf = id => SHEET.findIndex(r => r && r.v[2] === id);
const snapshot = () => JSON.stringify(SHEET.slice(1).map(r => [r.v, r.fill, r.font, r.nf]));
const wrote = () => LOG.filter(l => l.m !== "GET");
const job = id => ({ id, cust: "Person A", ph: "086", area: "Cork", eir: "", off: "9001", colour: "", flag: "", done: 0 });

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ---- 1. the move gate and the allowed targets ---- */
  const NAMES = ["Can sell as second hand", "Ready, customer won't take", "Collect & supply only", "Ready to fit", "In production"];
  assert.strictEqual(S.moveAllowed({ done: 0 }), false);
  assert.strictEqual(S.moveAllowed({ done: 1 }), true);
  assert.deepStrictEqual(S.moveTargets(NAMES, 3).map(t => t.name), ["Ready, customer won't take", "Collect & supply only"]);
  assert.deepStrictEqual(S.moveTargets(NAMES, 4).map(t => t.idx), [1, 2, 3]);
  pass("move only once the office marked it ready; never to In production or second hand");

  /* ---- 2. urgent / booked exclusive, and the colour values ---- */
  assert.strictEqual(S.nextFlag("", "urgent"), "urgent");
  assert.strictEqual(S.nextFlag("booked", "urgent"), "urgent", "urgent replaces booked");
  assert.strictEqual(S.nextFlag("urgent", "booked"), "booked", "booked replaces urgent");
  assert.strictEqual(S.nextFlag("urgent", "urgent"), "", "pressing it again switches it off");
  assert.strictEqual(S.nextFlag("trade", "booked"), "booked");
  assert.deepStrictEqual(S.COLOURS, { urgent: "#FF0000", booked: "#00B050", "": "#000000" });
  reset(); resetLists();
  await S.setColour(ctx(), Object.assign(job("R0006"), { flag: "" }), "urgent");
  const fontW = wrote().filter(l => /format\/font$/.test(l.u));
  assert.strictEqual(fontW.length, 1); assert.strictEqual(fontW[0].u, "range(address='A10:CL10')/format/font");
  assert.deepStrictEqual(fontW[0].b, { color: "#FF0000" }, "font colour only - nothing else in the body");
  assert(!wrote().some(l => /fill/.test(l.u)), "no fill is ever touched");
  assert.strictEqual(ITEMS["1"].fields.Kind, "colour"); assert.strictEqual(ITEMS["1"].fields.To, "Urgent");
  assert.strictEqual(ITEMS["1"].fields.Section, "Ready to fit", "the live section, found by the locate");
  assert(ITEMS["1"].fields.Row.length > 100, "the colour backup carries the captured row");
  assert.deepStrictEqual(NOTES, [["R0006", "Text colour", "Black", "Urgent"]]);
  reset(); resetLists(); FAIL_ADD = true;
  await assert.rejects(S.setColour(ctx(), job("R0006"), "booked"));
  assert(!wrote().length, "no backup, no paint");
  await assert.rejects(CW.salesSetFont("R0006", 10, "#FF3399"), /Not a Sales text colour/);
  pass("urgent and booked exclude; the row's font colour is the only thing painted, after its backup");

  /* ---- 2b. Trade order / On hold rows: the real word is logged, and asked about ---- */
  assert.strictEqual(S.replaceQuestion("trade", "urgent"), "This row's text is pink (Trade order). Replace with red (Urgent)?");
  assert.strictEqual(S.replaceQuestion("hold", "booked"), "This row's text is blue (On hold). Replace with green (Booked)?");
  assert.strictEqual(S.replaceQuestion("urgent", "booked"), "", "red/green/black need no question");
  assert.strictEqual(S.replaceQuestion("", "urgent"), "");
  reset(); resetLists();
  await S.setColour(ctx(), Object.assign(job("R0006"), { flag: "trade" }), S.nextFlag("trade", "urgent"));
  assert.deepStrictEqual(NOTES.pop(), ["R0006", "Text colour", "Trade order", "Urgent"]);
  assert.strictEqual(ITEMS["1"].fields.From, "Trade order");
  await S.setColour(ctx(), Object.assign(job("R0006"), { flag: "hold" }), "");
  assert.deepStrictEqual(NOTES.pop(), ["R0006", "Text colour", "On hold", "Black"]);
  pass("a pink or blue row: From names Trade order / On hold, and the page asks before replacing it");

  /* ---- 3. Row JSON size ---- */
  reset(); SHEET.splice(10, 0, wideRow("R0042"));
  const cap = await CW.captureRow(rowOf("R0042"), null);
  const json = S.rowJson(cap, "In production");
  assert(json.length < S.ROW_MAX, "a real-width row fits: " + json.length);
  const back = S.parseRow(json);
  ["values", "types", "numberFormat", "formulas", "fills", "fonts"].forEach(k => assert.deepStrictEqual(back.cap[k], cap[k], k + " round-trips"));
  assert.strictEqual(back.sectionName, "In production");
  const huge = Object.assign({}, cap, { values: cap.values.map(() => "x".repeat(800)) });
  assert.throws(() => S.rowJson(huge, "In production"), /too large/);
  pass("a 90-column row with formulas, fills and fonts fits (" + json.length + " chars); an oversize one is refused");

  /* ---- 4. delete order ---- */
  reset(); resetLists(); let before = snapshot(); FAIL_ADD = true;
  await assert.rejects(S.deleteJob(ctx(), job("R0006")), /list refused/);
  assert.strictEqual(snapshot(), before); assert(!wrote().some(l => /delete/.test(l.u)));
  resetLists(); BAD_READBACK = true; LOG = [];
  await assert.rejects(S.deleteJob(ctx(), job("R0006")), /read back/);
  assert.strictEqual(snapshot(), before); assert(!wrote().length);
  pass("delete: nothing goes when the backup fails or reads back different");

  /* ---- 4b. (review 1) one row backed up, the same row deleted - or nothing ---- */
  reset(); resetLists(); dupRow("R0006"); before = snapshot();
  await assert.rejects(S.deleteJob(ctx(), job("R0006")), /2 times/);
  await assert.rejects(S.setColour(ctx(), job("R0006"), "urgent"), /2 times/);
  await assert.rejects(S.saveCustomer(ctx(), job("R0006"), "cust", "x"), /2 times/);
  assert.strictEqual(snapshot(), before); assert(!wrote().length); assert.strictEqual(Object.keys(ITEMS).length, 0,
    "a job on the sheet twice: nothing backed up, written or deleted");
  reset(); resetLists(); before = snapshot();
  ON_READ = n => { if (n === 2) shiftDown(); };         // rows move between the backup and the delete
  await assert.rejects(S.deleteJob(ctx(), job("R0006")), /has moved from row 10 to row 11/);
  assert(!wrote().some(l => /delete$/.test(l.u)), "the row moved after its backup: nothing deleted");
  assert.strictEqual(rowOf("R0006"), 11, "and it is still there");
  reset(); resetLists();
  ON_READ = n => { if (n === 2) dupRow("R0006"); };     // a duplicate appears between backup and delete
  await assert.rejects(S.deleteJob(ctx(), job("R0006")), /2 times/);
  assert(!wrote().some(l => /delete$/.test(l.u)));
  reset(); resetLists();
  ON_READ = n => { if (n === 2) shiftDown(); };
  await assert.rejects(S.setColour(ctx(), job("R0006"), "urgent"), /has moved/);
  assert(!wrote().some(l => /format\/font$/.test(l.u)), "no paint onto a row that moved after its backup");
  reset(); resetLists();
  ON_READ = n => { if (n === 2) shiftDown(); };
  await assert.rejects(S.saveCustomer(ctx(), job("R0006"), "cust", "x"), /has moved/);
  assert(!wrote().length);
  /* restore's own duplicate check reads the same 600 rows */
  reset(); resetLists(); SHEET.push(...Array.from({ length: 440 }, () => mkRow([]))); SHEET.push(mkRow(["", "", "R0099"]));
  const capX = await CW.captureRow(rowOf("R0005"), null);
  ITEMS.far = { fields: { Row: S.rowJson(capX, "In production") } };
  await assert.rejects(S.restoreJob(ctx(), { id: "far", fields: { Kind: "delete", Job: "R0099", Restored: "No" } }), /already on the Production sheet/,
    "a copy of the job at row " + rowOf("R0099") + " (past row 400) still counts");
  pass("delete/colour/customer refuse a job on the sheet twice, or a row that moved after its backup");

  /* ---- 4c. (review 3) logged straight after the delete; a border failure is reported, not a failed delete ---- */
  reset(); resetLists(); FAIL_BORDERS = 9;           // the row above R0006: only restoreBottomEdge writes there
  const d2 = await S.deleteJob(ctx(), job("R0006"));
  assert.strictEqual(rowOf("R0006"), -1, "deleted");
  assert.deepStrictEqual(NOTES.pop(), ["R0006", "Deleted", "Ready to fit", ""], "and logged");
  assert(/BorderRefused/.test(d2.edgeError), "the border step's failure comes back as a report");
  FAIL_BORDERS = 9;
  const r2 = await S.restoreJob(ctx(), { id: d2.itemId, fields: ITEMS[d2.itemId].fields });
  assert(/BorderRefused/.test(r2.edgeError)); assert.deepStrictEqual(NOTES.pop(), ["R0006", "Restored", "", "Ready to fit"]);
  pass("delete and restore are logged the moment the row goes or comes back; a border failure is only reported");

  reset(); resetLists(); FAIL_BORDERS = 0;
  const orig = JSON.parse(JSON.stringify(SHEET[rowOf("R0006")]));
  const del = await S.deleteJob(ctx(), job("R0006"));
  assert.strictEqual(rowOf("R0006"), -1, "the row is gone");
  assert.deepStrictEqual(jobsInOrder(), ["R0001", "C0003", "R0005", "R0007", "R0008"]);
  const di = wrote().findIndex(l => /delete$/.test(l.u));
  assert(di >= 0 && wrote().slice(0, di).every(l => !/insert|delete/.test(l.u)), "one row delete, nothing structural before it");
  assert.strictEqual(ITEMS[del.itemId].fields.Kind, "delete");
  assert.deepStrictEqual(NOTES.pop(), ["R0006", "Deleted", "Ready to fit", ""]);
  assert.strictEqual(ITEMS[del.itemId].fields.Section, "Ready to fit");
  pass("delete: the backed-up row is the deleted row, logged with its live section");

  /* ---- 5. restore ---- */
  const item = { id: del.itemId, fields: Object.assign({}, ITEMS[del.itemId].fields) };
  before = snapshot();
  await assert.rejects(S.restoreJob(ctx(), { id: "x", fields: { Kind: "delete", Job: "R0005" } }), /could not be read/);
  ITEMS.fake = { fields: { Row: S.rowJson(cap, "In production") } };
  await assert.rejects(S.restoreJob(ctx(), { id: "fake", fields: { Kind: "delete", Job: "R0005", Restored: "No" } }), /already on the Production sheet/);
  ITEMS.gone = { fields: { Row: ITEMS[del.itemId].fields.Row.replace('"Ready to fit"', '"No such section"') } };
  await assert.rejects(S.restoreJob(ctx(), { id: "gone", fields: { Kind: "delete", Job: "R0006", Restored: "No" } }), /nowhere safe/);
  assert.strictEqual(snapshot(), before, "a refused restore changes nothing");
  FAIL_PATCH_VALUES = 1;
  await assert.rejects(S.restoreJob(ctx(), item), /Boom|500/);
  assert.strictEqual(snapshot(), before, "a failed write: the inserted row is taken out again");
  assert.strictEqual(ITEMS[del.itemId].fields.Restored, "No");
  const put = await S.restoreJob(ctx(), item);
  assert.deepStrictEqual(jobsInOrder(), ["R0001", "C0003", "R0005", "R0006", "R0007", "R0008"], "back at the bottom of its section");
  assert.strictEqual(put.row, rowOf("R0006"));
  const now6 = SHEET[rowOf("R0006")];
  assert.deepStrictEqual(now6.v, orig.v); assert.deepStrictEqual(now6.fill, orig.fill); assert.deepStrictEqual(now6.nf, orig.nf);
  assert.strictEqual(ITEMS[del.itemId].fields.Restored, "Yes"); assert.strictEqual(ITEMS[del.itemId].fields.RestoredBy, "Person A");
  assert(Object.keys(ITEMS).some(k => ITEMS[k].fields.Kind === "restore"));
  assert.deepStrictEqual(NOTES.pop(), ["R0006", "Restored", "", "Ready to fit"]);
  assert.throws(() => S.restoreJob(ctx(), { id: del.itemId, fields: ITEMS[del.itemId].fields }), /already been restored/);
  pass("restore: refused when on the sheet or the section is gone; a failed copy is removed; the row comes back whole");

  /* ---- 6. customer cells ---- */
  reset(); resetLists();
  await S.saveCustomer(ctx(), job("R0005"), "phone", "0871234567");
  const pw = wrote().filter(l => l.m === "PATCH");
  assert.strictEqual(pw.length, 1); assert.strictEqual(pw[0].u, "range(address='J9')");
  /* (review 8) the same route writeRow takes for a text cell: `formulas`, with the apostrophe */
  assert.deepStrictEqual(pw[0].b, { formulas: [["'0871234567"]] }, "written as text, through formulas");
  assert.strictEqual(SHEET[9].v[9], "0871234567", "the leading 0 is kept");
  const viaWriteRow = []; ["0871234567", "007", "Cork", "TRUE", "A65 F4E2"].forEach(v =>
    viaWriteRow.push(cellOut({ formulas: [v], values: [v], types: ["String"] }, 0)));
  assert.deepStrictEqual(viaWriteRow, ["'0871234567", "'007", "Cork", "'TRUE", "A65 F4E2"],
    "writeRow's own text cells: the same marker, so the two routes agree");
  /* (review 9) the Dashboard Log gets the last three characters only */
  assert.deepStrictEqual(NOTES.pop(), ["R0005", "Phone no", "•••086", "•••567"]);
  const ed = Object.keys(ITEMS).map(k => ITEMS[k].fields).find(f => f.Kind === "edit");
  assert.strictEqual(ed.Field, "Phone no"); assert.strictEqual(ed.From, "086"); assert.strictEqual(ed.To, "0871234567",
    "the whole value stays in the backups list only");
  LOG = []; await S.saveCustomer(ctx(), Object.assign(job("R0005"), { eir: "V94 N6PC" }), "eir", "A65 F4E2");
  assert.deepStrictEqual(wrote()[0].b, { formulas: [["A65 F4E2"]] });
  assert.deepStrictEqual(NOTES.pop(), ["R0005", "Eircode", "•••6PC", "•••4E2"]);
  /* (review 6) every one of the six is text: an office no keeps its leading zeros */
  LOG = []; await S.saveCustomer(ctx(), job("R0005"), "off", "007");
  assert.deepStrictEqual(wrote()[0].b, { formulas: [["'007"]] }); assert.strictEqual(SHEET[9].v[1], "007");
  assert.deepStrictEqual(NOTES.pop(), ["R0005", "Office no", "9001", "007"], "not a secret: logged whole");
  LOG = []; await S.saveCustomer(ctx(), job("R0005"), "cust", "=HYPERLINK(1)");
  assert.deepStrictEqual(wrote()[0].b, { formulas: [["'=HYPERLINK(1)"]] }, "a customer name is never a formula");
  /* (review 7) the column's header is read first */
  SHEET[2].v[9] = "MOBILE"; LOG = [];
  await assert.rejects(S.saveCustomer(ctx(), job("R0005"), "phone", "0861111111"), /no longer reads 'phone no'/);
  assert(!wrote().length, "a moved column: nothing written");
  SHEET[2].v[9] = "PHONE NO.";
  await assert.rejects(S.saveCustomer(Object.assign(ctx(), { hdrRow: () => 0 }), job("R0005"), "phone", "0861111111"), /header is not known/);
  assert.throws(() => S.saveCustomer(ctx(), job("R0005"), "sold", "x"), /not a customer field/);
  pass("customer cells: all six written as text via formulas; header checked; phone/eircode logged masked");

  /* ---- 7. not the Sales page: nothing can run ---- */
  window.CW_PAGE = undefined; reset(); resetLists();
  const f0 = FETCHES;
  const tries = [
    () => S.saveCustomer(ctx(), job("R0005"), "cust", "x"), () => S.setColour(ctx(), job("R0005"), "urgent", "x"),
    () => S.deleteJob(ctx(), job("R0005"), "x"), () => S.restoreJob(ctx(), item),
    () => S.setDelivery(ctx(), job("R0005"), "2026-10-07"), () => S.sendRequest(ctx(), "R0005", "status", "when?"),
    () => S.markSeen(ctx(), []), () => S.moveJob(ctx(), { id: "R0005", done: 1 }, 2, NAMES, async () => 1),
    () => CW.salesLocate("R0005"), () => CW.salesSetCell("R0005", 9, 9, "x", { row: 2, label: "customer" }),
    () => CW.salesSetFont("R0005", 9, "#FF0000"),
    () => CW.salesDeleteRow("R0005", 9), () => CW.salesInsertRow("R0006", cap, "In production")
  ];
  for (const t of tries) {
    let threw = false;
    try { await t(); } catch (e) { threw = /Only the Sales page/.test(e.message); }
    assert(threw, "refused: " + t);
  }
  assert.strictEqual(FETCHES, f0, "not one fetch"); assert.strictEqual(Object.keys(ITEMS).length, 0); assert.strictEqual(NOTES.length, 0);
  window.CW_PAGE = "sales";
  pass("off the Sales page every new write throws before Graph is asked anything");

  /* ---- 8. requests: unread count and ReplySeen ---- */
  const R = [
    { id: "1", fields: { Job: "R0005", At: "2026-09-30T09:00:00Z", Reply: "", ReplySeen: "No" } },
    { id: "2", fields: { Job: "R0006", At: "2026-09-30T08:00:00Z", Reply: "Monday", ReplySeen: "No" } },
    { id: "3", fields: { Job: "R0007", At: "2026-09-30T07:00:00Z", Reply: "Done", ReplySeen: "Yes" } }];
  assert.strictEqual(S.unanswered(R).length, 1);
  assert.strictEqual(S.unreadReplies(R).length, 1);
  assert.strictEqual(S.openFor(R, "r0005"), true); assert.strictEqual(S.openFor(R, "R0006"), false);
  assert.deepStrictEqual(S.officeOrder(R.slice().reverse()).map(r => r.id), ["1", "2", "3"], "unanswered first, then newest");
  resetLists(); ITEMS["2"] = { fields: Object.assign({}, R[1].fields) }; ITEMS["3"] = { fields: Object.assign({}, R[2].fields) };
  const patched = []; const c2 = Object.assign(ctx(), { CW: Object.assign({}, fakeCW, { listPatch: async (l, id, f) => { patched.push([l, id, f]); } }) });
  assert.strictEqual(await S.markSeen(c2, R), 1);
  assert.deepStrictEqual(patched, [["Sales requests", "2", { ReplySeen: "Yes" }]], "only the unread reply is marked");
  resetLists();
  await S.sendRequest(ctx(), "r0005", "move", "  please move to ready  ");
  const rq = ITEMS["1"].fields;
  assert.strictEqual(rq.Job, "R0005"); assert.strictEqual(rq.Kind, "move"); assert.strictEqual(rq.Text, "please move to ready");
  assert.strictEqual(rq.From, "Person A"); assert.strictEqual(rq.ReplySeen, "No"); assert(/^R0005\|/.test(rq.Title));
  await S.setDelivery(ctx(), job("R0005"), "2026-10-07", "");
  assert.strictEqual(ITEMS.uR0005.fields.DeliveryDate, "2026-10-07");
  assert.strictEqual(S.dayWords("2026-10-07"), "Wed 07 Oct");
  assert.deepStrictEqual(S.deliveryMap([{ id: "9", fields: { Title: "r0005", DeliveryDate: "2026-10-07", SetBy: "Person B" } },
    { id: "8", fields: { Title: "R0006", DeliveryDate: "" } }]), { R0005: { date: "2026-10-07", by: "Person B", at: "", id: "9" } });
  assert.throws(() => S.setDelivery(ctx(), job("R0005"), "07/10/2026"), /YYYY-MM-DD/);
  assert.throws(() => S.sendRequest(Object.assign(ctx(), { who: "" }), "R0005", "status", "x"), /Pick your name/);
  assert.throws(() => S.markSeen(Object.assign(ctx(), { who: "" }), R), /Pick your name/, "no write at all before a name");
  pass("requests: unread count, ReplySeen marks only the unread; delivery date and request fields");

  /* ---- 9. (review 2) one queue: the move waits for a write in flight, and busy() holds the controls ---- */
  const events = [], order = [];
  S.onBusy = b => events.push(b);
  let release; const gate = new Promise(r => { release = r; });
  const w1 = S.serial(async () => { order.push("write starts"); await gate; order.push("write ends"); });
  assert.strictEqual(S.busy(), true, "busy from the moment a write is asked for");
  let moved = 0;
  const mv = S.moveJob(ctx(), { id: "R0001", done: 1 }, 2, NAMES, async (ids, idx) => { order.push("move " + ids + "->" + idx); moved++; return 1; });
  await new Promise(r => setTimeout(r, 20));
  assert.deepStrictEqual(order, ["write starts"], "the move has not started while the write is in flight");
  assert.strictEqual(S.busy(), true);
  release(); await w1; await mv;
  assert.deepStrictEqual(order, ["write starts", "write ends", "move R0001->2"]);
  await new Promise(r => setTimeout(r, 0));
  assert.strictEqual(S.busy(), false, "and free again once both are done");
  assert.deepStrictEqual(events, [true, false], "the page hears busy once and free once");
  assert.throws(() => S.moveJob(ctx(), { id: "R0005", done: 0 }, 2, NAMES, async () => 1), /not marked/);
  assert.throws(() => S.moveJob(ctx(), { id: "R0001", done: 1 }, 4, NAMES, async () => 1), /cannot be moved to In production/);
  assert.throws(() => S.moveJob(ctx(), { id: "R0001", done: 1 }, 0, NAMES, async () => 1), /Can sell as second hand/);
  assert.throws(() => S.moveJob(Object.assign(ctx(), { who: "" }), { id: "R0001", done: 1 }, 2, NAMES, async () => 1), /Pick your name/);
  assert.strictEqual(moved, 1);
  const failing = S.serial(async () => { throw new Error("boom"); });
  await assert.rejects(failing, /boom/);
  await new Promise(r => setTimeout(r, 0));
  assert.strictEqual(S.busy(), false, "a failed write frees the queue too");
  S.onBusy = null;
  pass("the Sales move runs in the same queue as every write; busy() covers the whole time anything is in flight");

  /* ---- 9b. (amendment B) a drag or "move to" on several jobs: ready ones move, the rest are refused ---- */
  const calls = [];
  const mv2 = async (ids, idx) => { calls.push([ids.slice(), idx]); return ids.length; };
  const mix = [{ id: "R0001", done: 1 }, { id: "R0005", done: 0 }, { id: "C0003", done: 1 }, { id: "R0006", done: 0 }];
  let res = await S.moveMany(ctx(), mix, 3, NAMES, mv2);
  assert.deepStrictEqual(calls, [[["R0001", "C0003"], 3]], "only the jobs the office marked ready are handed to the move");
  assert.deepStrictEqual(res, { moved: 2, refused: ["R0005", "R0006"] }, "the refused ones are named for the toast");
  assert.strictEqual(S.REFUSED_NOTE, "the office has not marked this job ready - send a request");
  calls.length = 0;
  res = await S.moveMany(ctx(), [{ id: "R0005", done: 0 }], 3, NAMES, mv2);
  assert.deepStrictEqual(calls, [], "nothing ready: the move is never called");
  assert.deepStrictEqual(res, { moved: 0, refused: ["R0005"] });
  assert.throws(() => S.moveMany(ctx(), mix, 4, NAMES, mv2), /In production from here/, "the same targets as the drawer's Move");
  assert.throws(() => S.moveMany(ctx(), mix, 0, NAMES, mv2), /second hand from here/);
  assert.throws(() => S.moveMany(Object.assign(ctx(), { who: "" }), mix, 3, NAMES, mv2), /Pick your name/);
  /* ... and in the same queue: a write in flight holds the multi-move back */
  let rel2; const g2 = new Promise(r => { rel2 = r; }); const ord2 = [];
  const w2 = S.serial(async () => { await g2; ord2.push("write"); });
  const m2 = S.moveMany(ctx(), mix, 2, NAMES, async ids => { ord2.push("move " + ids.join("+")); return ids.length; });
  await new Promise(r => setTimeout(r, 10));
  assert.deepStrictEqual(ord2, []); rel2(); await w2; await m2;
  assert.deepStrictEqual(ord2, ["write", "move R0001+C0003"]);
  window.CW_PAGE = undefined;
  assert.throws(() => S.moveMany(ctx(), mix, 3, NAMES, mv2), /Only the Sales page/);
  window.CW_PAGE = "sales";
  pass("several jobs at once: ready ones move through the queue, the rest are refused and named");

  /* ---- 10. (review 12) the office's reply: refused if someone has replied since ---- */
  window.CW_PAGE = undefined;
  const RQ = { a: { Reply: "", ReplyBy: "" }, b: { Reply: "Tuesday", ReplyBy: "colleague@example.test" } }, rp = [];
  const offCW = { listItem: async (l, id) => RQ[id] ? { id, fields: Object.assign({}, RQ[id]) } : null,
                  listPatch: async (l, id, f) => { rp.push([l, id, f]); Object.assign(RQ[id], f); } };
  const got = await S.officeReply(offCW, "a", "  Ready Thursday ", "office@example.test", "2026-09-30T12:00:00Z");
  assert.deepStrictEqual(rp, [["Sales requests", "a", { Reply: "Ready Thursday", ReplyBy: "office@example.test", ReplyAt: "2026-09-30T12:00:00Z" }]],
    "Reply, ReplyBy, ReplyAt and nothing else");
  assert.strictEqual(got.Reply, "Ready Thursday");
  await assert.rejects(S.officeReply(offCW, "b", "Wednesday", "office@example.test"), /Already answered by colleague/);
  await assert.rejects(S.officeReply(offCW, "a", "again", "office@example.test"), /Already answered/, "a second reply to the same request is refused");
  await assert.rejects(S.officeReply(offCW, "zz", "x", "office@example.test"), /no longer in the list/);
  await assert.rejects(S.officeReply(offCW, "a", "   ", "office@example.test"), /Write the reply/);
  assert.strictEqual(rp.length, 1, "only the first reply was written");
  window.CW_PAGE = "sales";
  await assert.rejects(S.officeReply(offCW, "a", "x", "Person A"), /office page/);
  pass("office reply: read first, refused if already answered, writes only the three reply fields");

  /* ---- 11. (review 13) the Sales hold keeps a move hold's revert data ---- */
  vm.runInThisContext(fs.readFileSync(__dirname + "/sales.js", "utf8"), { filename: "sales.js" });
  vm.runInThisContext('SALES_HOLD.R0001 = { flag: "booked", flagHex: "#00B050", at: Date.now() }');
  const parsedJob = { id: "R0001", blk: 1, flag: "", urg: 0, main: { cust: "Person A", ph: "" } };
  const moving = Object.assign({}, parsedJob, { blk: 3, raw: parsedJob });    // applyPending's copy: moved, file not caught up
  const out = vm.runInThisContext("salesOverlay")([moving])[0];
  assert.strictEqual(out.flag, "booked"); assert.strictEqual(out.blk, 3, "the move hold is still shown");
  assert(out.raw, "the revert data is kept");
  assert.strictEqual(out.raw.blk, 1, "a failed move can still go back to where the file has it");
  assert.strictEqual(out.raw.flag, "booked", "and going back keeps the Sales colour hold");
  pass("salesOverlay keeps applyPending's raw (revert) job, with the Sales fields on it too");

  /* ---- 12. the Production sheet only (owner's rule 2026-09-18; review of the demo) ---- */
  const XL = require("exceljs");
  const wbP = new XL.Workbook();
  const head = ws => { const p = (r, c, v) => { ws.getCell(r, c).value = v; };
    p(2, 2, "OFFICE NO."); p(2, 4, "DATES ON CONTRACT"); p(3, 4, "SOLD"); p(2, 5, "CUSTOMER"); p(2, 6, "PHONE NO.");
    p(2, 7, "AREA"); p(2, 8, "EIRCODE"); p(2, 13, "QUANTITY"); p(3, 13, "WND"); p(3, 14, "DRS"); };
  const wsP = wbP.addWorksheet("Production"), wsP2 = wbP.addWorksheet("Production (2)");
  head(wsP); head(wsP2);
  const cell = (ws, r, c, v) => { ws.getCell(r, c).value = v; };
  /* R0001 on Production with a BLANK phone and area; Production (2) has both, and another name */
  cell(wsP, 6, 3, "R0001"); cell(wsP, 6, 5, "Person A"); cell(wsP, 6, 8, "A65 F4E2"); cell(wsP, 6, 13, 3);
  cell(wsP2, 6, 3, "R0001"); cell(wsP2, 6, 5, "Person B"); cell(wsP2, 6, 6, "0871111111"); cell(wsP2, 6, 7, "Cork"); cell(wsP2, 6, 13, 9);
  /* R0009 is on Production (2) only */
  cell(wsP2, 7, 3, "R0009"); cell(wsP2, 7, 5, "Person B"); cell(wsP2, 7, 6, "0862222222");
  const parsed = parseWorkbook(wbP);
  const p1 = parsed.find(x => x.id === "R0001"), p9 = parsed.find(x => x.id === "R0009");
  assert.strictEqual(p1.ph, "0871111111", "(the office's merged value is untouched: Production (2) filled the blank)");
  assert.deepStrictEqual(p1.main, { cust: "Person A", area: "", eir: "A65 F4E2", off: "", colour: "", ph: "" },
    "the Production-only copy: blank where Production is blank");
  assert.strictEqual(p9.main, null, "a job not on Production has no Production copy");
  const shown = vm.runInThisContext("salesOverlay")(parsed);
  assert.deepStrictEqual(shown.map(x => x.id), ["R0001"], "the Sales list: a job only on Production (2) is absent");
  const s1 = shown[0];
  assert.strictEqual(s1.ph, "", "Sales shows the blank phone, not Production (2)'s number");
  assert.strictEqual(s1.area, ""); assert.strictEqual(s1.cust, "Person A"); assert.strictEqual(s1.eir, "A65 F4E2");
  assert.strictEqual(s1.wnd, 3, "Wnd off Production alone (Production (2) says 9)");
  assert.strictEqual(S.fieldOf("phone").of(s1), "", "and the customer form's 'from' is the Production cell");
  pass("Sales reads the Production sheet only: blanks stay blank, other sheets' jobs never appear");

  /* ---- 13. ready (gold) from Production only: a gold row on John's sheet is not "ready" ---- */
  const wbG = new XL.Workbook();
  const gP = wbG.addWorksheet("Production"), gP2 = wbG.addWorksheet("Production (2)");
  head(gP); head(gP2);
  const gold = (ws, r) => { for (let c = 1; c <= 60; c++) ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE699" } }; };
  /* R0002: plain on Production, the whole row gold on Production (2) */
  cell(gP, 6, 3, "R0002"); cell(gP, 6, 5, "Person A"); cell(gP, 6, 13, 2);
  cell(gP2, 6, 3, "R0002"); cell(gP2, 6, 5, "Person A"); cell(gP2, 6, 13, 2); gold(gP2, 6);
  /* R0003: gold on Production itself */
  cell(gP, 7, 3, "R0003"); cell(gP, 7, 5, "Person B"); cell(gP, 7, 13, 1); gold(gP, 7);
  const pg = parseWorkbook(wbG);
  const gj2 = pg.find(x => x.id === "R0002"), gj3 = pg.find(x => x.id === "R0003");
  assert.strictEqual(gj2.done, 1, "(the raw parse: gold on any sheet)");
  assert.strictEqual(gj2.doneMain, 0, "but not gold on Production");
  assert.notStrictEqual(gj2.cp.win, "done", "and a gold row on John's sheet paints no checkpoint 'done'");
  assert.strictEqual(gj3.doneMain, 1); assert.strictEqual(gj3.cp.win, "done");
  const pj = productionOnly(pg), pj2 = pj.find(x => x.id === "R0002"), pj3 = pj.find(x => x.id === "R0003");
  assert.strictEqual(pj2.done, 0, "the pages see R0002 as not ready");
  assert.strictEqual(pj3.done, 1, "and R0003, gold on Production, as ready");
  assert.strictEqual(S.moveAllowed(pj2), false, "so the Sales move gate refuses it");
  assert.throws(() => S.moveJob(ctx(), pj2, 3, NAMES, async () => 1), /not marked/);
  const mm = await S.moveMany(ctx(), [pj2, pj3], 3, NAMES, async ids => ids.length);
  assert.deepStrictEqual(mm, { moved: 1, refused: ["R0002"] }, "gold only on Production (2): refused; gold on Production: moves");
  pass("ready is a gold row on Production only: gold on Production (2) is not ready, and Sales cannot move it");

  console.log("\n" + n + " checks passed");
  process.exit(0);
})().catch(e => { console.error("FAIL", e); process.exit(1); });
