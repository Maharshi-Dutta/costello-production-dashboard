/* Offline test of moveJobRow against a fake Production sheet served through a
   fake fetch(). Covers: move down, move up, no-op, write failure (sheet put
   back), the sheet changing underneath (sheet put back), empty top section.
   Run: node test_move.js                                                    */
const fs = require("fs"), vm = require("vm"), assert = require("assert");

global.ExcelJS = { Workbook: function () {} };
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" } };
vm.runInThisContext(fs.readFileSync(__dirname + "/parser.js", "utf8"), { filename: "parser.js" });
vm.runInThisContext(fs.readFileSync(__dirname + "/graph.js", "utf8"), { filename: "graph.js" });
const CW = window.CW;
CW._setToken(() => "t");
CW._setFile({ base: "/x/workbook", content: "/x/content", meta: "/x" });

/* ---------- a tiny Excel ---------- */
const N = 90;
const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
function mkRow(vals, fill, height, bold) {
  const r = { v: new Array(N).fill(""), fill: new Array(N).fill(fill || ""), height: height || 30,
              font: [], nf: new Array(N).fill("General"), align: new Array(N).fill(null), borders: new Array(N).fill(null) };
  for (let c = 0; c < N; c++) r.font.push({ bold: !!bold && c === 2, italic: false, size: 12, name: "Calibri", color: c === 8 ? "#FF3399" : "#000000", underline: "None" });
  (vals || []).forEach((x, i) => { r.v[i] = x; });
  r.nf[3] = "d-mmm";
  return r;
}
let SHEET, LOG, FAIL_PATCH_VALUES, SHIFT_ON_VERIFY;
function reset() {
  SHEET = [null,
    mkRow(["", "", "", "", "", "", "", "", "", "", "8843"]), mkRow(["COMMENT", "OFFICE NO.", "JOB NO."]), mkRow(["", "", "", "SOLD"]),
    mkRow(["Below is orders ready and customers won't take"]),
    mkRow(["", 7785, "R0001", 46000, "", "", "", "", "Ann", "086", "Cork"], "#FFE699", 30.75, true),
    mkRow(["store it", 7786, "R0002", 46001, "", "", "", "", "Bob", "087", "Clare"], "#FFE699", 30.75, true),
    mkRow(["Below is collect or supply only orders"]),
    mkRow(["asap", 8001, "C0003", 46010, "", "", "", "", "Cal", "085", "Kerry"], "#FFE699", 30, true),
    mkRow(["07/04 will collect", 8002, "S0004", 46011, "", "", "", "", "Dee", "083", "Mayo"], "#FFE699", 30, true),
    mkRow(["Not sent to floor(No section) ="]),
    mkRow(["Unglazed", 9001, "R0005", 46020, "", "", "", "", "Eve", "089", "Cork"], "#FFE699", 30, true),
    mkRow(["on hold", 9002, "R0006", 46021, "", "", "", "", "Fay", "086", "Dublin"], "#FFE699", 30, true),
    mkRow(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "=SUM(P11:P13)"]),
    mkRow(["Not sent to floor(No section) ="]),
    mkRow(["", 9101, "R0007", 46030, "", "", "", "", "Gus", "087", "Laois"], "#FFFFFF", 30, true),
    mkRow(["URGENT", 9102, "R0008", 46031, "", "", "", "", "Hal", "086", "Cork"], "#FFFFFF", 30, true),
    mkRow(["", 9103, "R0009", 46032, "", "", "", "", "Ivy", "085", "Kerry"], "", 22.15, true),
    mkRow(["", 9104, "R0010", 46033, "", "", "", "", "Jo", "083", "Clare"], "", 22.15, true),
    mkRow(["", "", 5293]), mkRow(["", "", 5182])];
  LOG = []; FAIL_PATCH_VALUES = 0; SHIFT_ON_VERIFY = false;
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
  LOG.push(method + " " + url.replace(/^\/x\/workbook\/worksheets\('Production'\)\//, ""));
  if (url === "/x/workbook/createSession") return ok({ id: "S1" });
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
    const o = { values: [], valueTypes: [], numberFormat: [], formulas: [] };
    rows().forEach(r => {
      const v = [], t = [], nf = [], f = [];
      for (let c = A.c1; c <= A.c2; c++) { const x = r.v[c - 1]; const isF = typeof x === "string" && x.charAt(0) === "="; v.push(isF ? 0 : x); t.push(isF ? "Double" : typeOf(x)); nf.push(r.nf[c - 1]); f.push(x); }
      o.values.push(v); o.valueTypes.push(t); o.numberFormat.push(nf); o.formulas.push(f);
    });
    return ok(o);
  }
  if (method === "POST" && rest === "/insert") { assert(A.whole, "insert must address a whole row"); SHEET.splice(A.r1, 0, mkRow([], "#FFE699", 30.75)); return ok({}); }
  if (method === "POST" && rest === "/delete") { assert(A.whole, "delete must address a whole row"); SHEET.splice(A.r1, 1); return { status: 204, body: {} }; }
  if (method === "POST" && rest === "/format/fill/clear") { rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) r.fill[c - 1] = ""; }); return ok({}); }
  if (method === "PATCH") {
    if (rest === "") {
      if (FAIL_PATCH_VALUES) { FAIL_PATCH_VALUES--; return { status: 500, body: { error: { code: "Boom" } } }; }
      const r = SHEET[A.r1];
      body.formulas[0].forEach((x, i) => { r.v[A.c1 - 1 + i] = (typeof x === "string" && x.charAt(0) === "'") ? x.slice(1)
        : (typeof x === "string" && /^\d+$/.test(x)) ? Number(x) : x; });   // Excel parses typed text
      body.numberFormat[0].forEach((x, i) => { r.nf[A.c1 - 1 + i] = x; });
      if (SHIFT_ON_VERIFY) { SHEET.splice(2, 0, mkRow(["someone inserted a row in Excel"])); SHIFT_ON_VERIFY = false; }
      return ok({});
    }
    if (rest === "/format/fill") { rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) r.fill[c - 1] = body.color; }); return ok({}); }
    if (rest === "/format/font") { rows().forEach(r => { for (let c = A.c1; c <= A.c2; c++) r.font[c - 1] = Object.assign({}, body); }); return ok({}); }
    if (rest === "/format") { rows().forEach(r => { if (body.rowHeight) r.height = body.rowHeight; for (let c = A.c1; c <= A.c2; c++) if ("wrapText" in body) r.align[c - 1] = body; }); return ok({}); }
    if (rest.startsWith("/format/borders/")) { const side = rest.split("/").pop(); rows().forEach(r => { r.sides = r.sides || {}; r.sides[side] = (r.sides[side] || 0) + 1; for (let c = A.c1; c <= A.c2; c++) r.borders[c - 1] = (r.borders[c - 1] || 0) + 1; }); return ok({}); }
  }
  return { status: 404, body: { error: "no route " + method + " " + url } };
}
global.fetch = async (url, init) => {
  const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  let res;
  if (path === "/$batch") res = ok({ responses: body.requests.map(q => Object.assign({ id: q.id }, route(q.method, q.url, q.body))) });
  else res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status, text: async () => JSON.stringify(res.body), arrayBuffer: async () => new ArrayBuffer(0) };
};

const jobsInOrder = () => SHEET.slice(1).filter(r => /^[A-Z]{1,2}\d{3,5}$/.test(String(r.v[2]))).map(r => r.v[2]);
const snapshot = () => JSON.stringify(SHEET.slice(1).map(r => [r.v, r.fill, r.height, r.font, r.nf]));
const rowOf = id => SHEET.findIndex(r => r && r.v[2] === id);
const fakeTmpl = () => ({ row: 0, height: 30, cells: Array.from({ length: N }, (_, i) => ({ h: i === 1 ? "left" : null, v: null, wrap: i === 0,
  top: { style: "thin", color: "#000000" }, bottom: { style: i < 77 ? "medium" : "thin", color: "#000000" }, left: { style: "thin", color: "#000000" }, right: { style: "thin", color: "#000000" },
  fk: i < 77 ? "a" : "b", flk: i < 77 ? "g" : "" })) });

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  reset();
  const before = snapshot(); const origRow = JSON.parse(JSON.stringify(SHEET[rowOf("R0005")]));
  let r = await CW.moveJobRow("R0005", 4, () => fakeTmpl(), () => {});
  assert.deepStrictEqual(jobsInOrder(), ["R0001", "R0002", "C0003", "S0004", "R0006", "R0007", "R0008", "R0009", "R0010", "R0005"]);
  assert.strictEqual(r.moved, true); assert.strictEqual(r.from, "Ready to fit"); assert.strictEqual(r.to, "In production");
  assert.strictEqual(r.row, rowOf("R0005"));
  const moved = SHEET[rowOf("R0005")];
  assert.deepStrictEqual(moved.v, origRow.v); assert.deepStrictEqual(moved.fill, origRow.fill); assert.deepStrictEqual(moved.font, origRow.font);
  assert.strictEqual(moved.height, origRow.height); assert.deepStrictEqual(moved.nf, origRow.nf);
  assert.strictEqual(SHEET.length - 1, 20, "row count unchanged");
  const totals = SHEET.slice(1).filter(r => String(r.v[15]).startsWith("=SUM"));
  assert.strictEqual(totals.length, 1, "totals row still there, once");
  assert(!LOG.some(l => l.indexOf("/format/borders/EdgeTop") >= 0), "the top edge is never written");
  assert(moved.sides && moved.sides.EdgeBottom && moved.sides.EdgeLeft && moved.sides.EdgeRight && moved.sides.InsideVertical, "landing row: bottom + verticals from the template");
  assert(moved.align[0] && moved.align[0].wrapText === true && moved.align[1].horizontalAlignment === "Left", "alignment from the template");
  assert.strictEqual(moved.height, 30, "height from the template, not the API's rounded value");
  const above = SHEET[rowOf("R0005") - 1];
  assert.strictEqual(above.v[2], "R0010", "landed under the section's last job");
  assert(above.sides && above.sides.EdgeBottom >= 1 && Object.keys(above.sides).length === 1, "row above the landing row: only its bottom edge re-asserted");
  const aboveOld = SHEET[rowOf("R0006") - 1];   // R0006 moved up into R0005's old slot; the row above that slot was the divider - no write
  assert(!aboveOld.sides, "a divider above the vacated slot is left alone");
  const calls = LOG.length;
  pass("move down (Ready to fit -> In production): order, content, formats, row count; " + calls + " requests");

  r = await CW.moveJobRow("R0005", 4, null, () => {});
  assert.strictEqual(r.moved, false); pass("already in the section: nothing written");

  reset(); const first = SHEET[rowOf("R0009")];
  const orig9 = JSON.parse(JSON.stringify(first));
  r = await CW.moveJobRow("R0009", 3, null, () => {});
  assert.deepStrictEqual(jobsInOrder(), ["R0001", "R0002", "C0003", "S0004", "R0005", "R0006", "R0009", "R0007", "R0008", "R0010"]);
  const m9 = SHEET[rowOf("R0009")];
  assert.deepStrictEqual(m9.v, orig9.v); assert.deepStrictEqual(m9.fill, orig9.fill); assert.strictEqual(m9.height, 22.15);
  assert.strictEqual(String(SHEET[14].v[15]), "=SUM(P11:P13)", "totals row shifted down intact");
  pass("move up (In production -> Ready to fit), lands just above the totals row, no template");

  reset(); const cs = JSON.parse(JSON.stringify(SHEET[rowOf("S0004")]));
  await CW.moveJobRow("S0004", 4, null, () => {}); await CW.moveJobRow("S0004", 2, null, () => {});
  assert.deepStrictEqual(jobsInOrder(), ["R0001", "R0002", "C0003", "S0004", "R0005", "R0006", "R0007", "R0008", "R0009", "R0010"]);
  assert.strictEqual(SHEET[rowOf("S0004")].v[0], "07/04 will collect", "date-like text survives both moves");
  assert.deepStrictEqual(SHEET[rowOf("S0004")].v, cs.v);
  pass("C/S job to In production and back to Collect & supply: text starting with digits kept as text");

  reset(); const b0 = snapshot(); FAIL_PATCH_VALUES = 9;
  await assert.rejects(CW.moveJobRow("R0002", 4, null, () => {}), /500|Boom/);
  assert.strictEqual(snapshot(), b0); pass("write fails: inserted row removed again, sheet identical");

  reset(); SHIFT_ON_VERIFY = true; const jobs0 = jobsInOrder();
  await assert.rejects(CW.moveJobRow("R0002", 4, null, () => {}), /changed under/);
  assert.deepStrictEqual(jobsInOrder(), jobs0); assert.strictEqual(SHEET.length - 1, 21, "only the outside insert remains");
  pass("sheet changes underneath: nothing deleted, our copy removed");

  reset();
  await assert.rejects(CW.moveJobRow("R0007", 0, null, () => {}), /nowhere safe/);
  assert.strictEqual(snapshot(), before); pass("empty top section: refused before any write");

  console.log("\n" + n + " checks passed");
})().catch(e => { console.error("FAIL", e); process.exit(1); });
