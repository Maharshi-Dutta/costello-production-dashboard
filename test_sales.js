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
let SHEET, LOG, FAIL_PATCH_VALUES, MISMATCH_AT, CREADS, FETCHES = 0;
function reset() {
  SHEET = [null,
    mkRow(["", "", "", "", "", "", "", "", "", "", "8843"]), mkRow(["COMMENT", "OFFICE NO.", "JOB NO."]), mkRow(["", "", "", "SOLD"]),
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
  LOG = []; FAIL_PATCH_VALUES = 0; MISMATCH_AT = 0; CREADS = 0;
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
    /* the one-cell check of column C: the nth such read can be made to lie */
    if (A.c1 === 3 && A.c2 === 3 && A.r1 === A.r2) { CREADS++;
      if (MISMATCH_AT && CREADS === MISMATCH_AT) return ok({ values: [["R9999"]] }); }
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
  now: () => "2026-09-30T10:00:0" + (NEXT % 10) + ".000Z" });
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
  await S.setColour(ctx(), Object.assign(job("R0006"), { flag: "" }), "urgent", "In production");
  const fontW = wrote().filter(l => /format\/font$/.test(l.u));
  assert.strictEqual(fontW.length, 1); assert.strictEqual(fontW[0].u, "range(address='A10:CL10')/format/font");
  assert.deepStrictEqual(fontW[0].b, { color: "#FF0000" }, "font colour only - nothing else in the body");
  assert(!wrote().some(l => /fill/.test(l.u)), "no fill is ever touched");
  assert.strictEqual(ITEMS["1"].fields.Kind, "colour"); assert.strictEqual(ITEMS["1"].fields.To, "Urgent");
  assert(ITEMS["1"].fields.Row.length > 100, "the colour backup carries the captured row");
  assert.deepStrictEqual(NOTES, [["R0006", "Text colour", "Black", "Urgent"]]);
  reset(); resetLists(); FAIL_ADD = true;
  await assert.rejects(S.setColour(ctx(), job("R0006"), "booked", "In production"));
  assert(!wrote().length, "no backup, no paint");
  pass("urgent and booked exclude; the row's font colour is the only thing painted, after its backup");

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
  await assert.rejects(S.deleteJob(ctx(), job("R0006"), "In production"), /list refused/);
  assert.strictEqual(snapshot(), before); assert(!wrote().some(l => /delete/.test(l.u)));
  resetLists(); BAD_READBACK = true; LOG = [];
  await assert.rejects(S.deleteJob(ctx(), job("R0006"), "In production"), /read back/);
  assert.strictEqual(snapshot(), before); assert(!wrote().length);
  resetLists(); LOG = []; CREADS = 0; MISMATCH_AT = 1;
  await assert.rejects(S.deleteJob(ctx(), job("R0006"), "In production"), /no longer on row/);
  assert.strictEqual(Object.keys(ITEMS).length, 0, "a row that is not this job: nothing saved, nothing deleted");
  resetLists(); LOG = []; CREADS = 0; MISMATCH_AT = 2;       // passes the first check, fails the one before the delete
  await assert.rejects(S.deleteJob(ctx(), job("R0006"), "In production"), /no longer on row/);
  assert.strictEqual(snapshot(), before); assert(!wrote().length, "column C changed before the delete: nothing deleted");
  resetLists(); LOG = []; MISMATCH_AT = 0;
  const orig = JSON.parse(JSON.stringify(SHEET[rowOf("R0006")]));
  const del = await S.deleteJob(ctx(), job("R0006"), "In production");
  assert.strictEqual(rowOf("R0006"), -1, "the row is gone");
  assert.deepStrictEqual(jobsInOrder(), ["R0001", "C0003", "R0005", "R0007", "R0008"]);
  const di = wrote().findIndex(l => /delete$/.test(l.u));
  assert(di >= 0 && wrote().slice(0, di).every(l => !/insert|delete/.test(l.u)), "one row delete, nothing structural before it");
  assert.strictEqual(ITEMS[del.itemId].fields.Kind, "delete");
  assert.deepStrictEqual(NOTES.pop(), ["R0006", "Deleted", "In production", ""]);
  pass("delete: nothing goes when the backup fails, reads back different, or column C no longer matches");

  /* ---- 5. restore ---- */
  const item = { id: del.itemId, fields: Object.assign({}, ITEMS[del.itemId].fields) };
  before = snapshot();
  await assert.rejects(S.restoreJob(ctx(), { id: "x", fields: { Kind: "delete", Job: "R0005" } }), /could not be read/);
  ITEMS.fake = { fields: { Row: S.rowJson(cap, "In production") } };
  await assert.rejects(S.restoreJob(ctx(), { id: "fake", fields: { Kind: "delete", Job: "R0005", Restored: "No" } }), /already on the Production sheet/);
  ITEMS.gone = { fields: { Row: ITEMS[del.itemId].fields.Row.replace('"In production"', '"No such section"') } };
  await assert.rejects(S.restoreJob(ctx(), { id: "gone", fields: { Kind: "delete", Job: "R0006", Restored: "No" } }), /nowhere safe/);
  assert.strictEqual(snapshot(), before, "a refused restore changes nothing");
  FAIL_PATCH_VALUES = 1;
  await assert.rejects(S.restoreJob(ctx(), item), /Boom|500/);
  assert.strictEqual(snapshot(), before, "a failed write: the inserted row is taken out again");
  assert.strictEqual(ITEMS[del.itemId].fields.Restored, "No");
  const put = await S.restoreJob(ctx(), item);
  assert.deepStrictEqual(jobsInOrder(), ["R0001", "C0003", "R0005", "R0007", "R0008", "R0006"], "back at the bottom of its section");
  assert.strictEqual(put.row, rowOf("R0006"));
  const now6 = SHEET[rowOf("R0006")];
  assert.deepStrictEqual(now6.v, orig.v); assert.deepStrictEqual(now6.fill, orig.fill); assert.deepStrictEqual(now6.nf, orig.nf);
  assert.strictEqual(ITEMS[del.itemId].fields.Restored, "Yes"); assert.strictEqual(ITEMS[del.itemId].fields.RestoredBy, "Person A");
  assert(Object.keys(ITEMS).some(k => ITEMS[k].fields.Kind === "restore"));
  assert.deepStrictEqual(NOTES.pop(), ["R0006", "Restored", "", "In production"]);
  assert.throws(() => S.restoreJob(ctx(), { id: del.itemId, fields: ITEMS[del.itemId].fields }), /already been restored/);
  pass("restore: refused when on the sheet or the section is gone; a failed copy is removed; the row comes back whole");

  /* ---- 6. customer cells ---- */
  reset(); resetLists(); CREADS = 0; MISMATCH_AT = 1;
  await assert.rejects(S.saveCustomer(ctx(), job("R0005"), "cust", "Person B"), /no longer on row/);
  assert(!wrote().length, "a row that is not this job: no write");
  MISMATCH_AT = 0; LOG = [];
  await S.saveCustomer(ctx(), job("R0005"), "phone", "0871234567");
  const pw = wrote().filter(l => l.m === "PATCH");
  assert.strictEqual(pw.length, 1); assert.strictEqual(pw[0].u, "range(address='J9')");
  assert.deepStrictEqual(pw[0].b, { values: [["'0871234567"]] }, "written as text");
  assert.strictEqual(SHEET[9].v[9], "0871234567", "the leading 0 is kept");
  assert.deepStrictEqual(NOTES.pop(), ["R0005", "Phone no", "086", "0871234567"]);
  const ed = Object.keys(ITEMS).map(k => ITEMS[k].fields).find(f => f.Kind === "edit");
  assert.strictEqual(ed.Field, "Phone no"); assert.strictEqual(ed.From, "086"); assert.strictEqual(ed.To, "0871234567");
  LOG = []; await S.saveCustomer(ctx(), job("R0005"), "eir", "A65 F4E2");
  assert.deepStrictEqual(wrote()[0].b, { values: [["A65 F4E2"]] });
  LOG = []; await S.saveCustomer(ctx(), job("R0005"), "cust", "=HYPERLINK(1)");
  assert.deepStrictEqual(wrote()[0].b, { values: [["'=HYPERLINK(1)"]] }, "a customer name is never a formula");
  assert.throws(() => S.saveCustomer(ctx(), job("R0005"), "sold", "x"), /not a customer field/);
  pass("customer cells: refused on a row mismatch; phone and eircode kept as text");

  /* ---- 7. not the Sales page: nothing can run ---- */
  window.CW_PAGE = undefined; reset(); resetLists();
  const f0 = FETCHES;
  const tries = [
    () => S.saveCustomer(ctx(), job("R0005"), "cust", "x"), () => S.setColour(ctx(), job("R0005"), "urgent", "x"),
    () => S.deleteJob(ctx(), job("R0005"), "x"), () => S.restoreJob(ctx(), item),
    () => S.setDelivery(ctx(), job("R0005"), "2026-10-07"), () => S.sendRequest(ctx(), "R0005", "status", "when?"),
    () => S.markSeen(ctx(), []),
    () => CW.salesLocate("R0005"), () => CW.salesSetCell("R0005", 9, 9, "x"), () => CW.salesSetFont("R0005", 9, "#FF0000"),
    () => CW.salesDeleteRow("R0005"), () => CW.salesInsertRow("R0006", cap, "In production")
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
  pass("requests: unread count, ReplySeen marks only the unread; delivery date and request fields");

  console.log("\n" + n + " checks passed");
  process.exit(0);
})().catch(e => { console.error("FAIL", e); process.exit(1); });
