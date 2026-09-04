/* Offline test of the Checkpoints feature: the merge rule between Excel's
   colour and the dashboard's own counts, held (pending) counts and how they
   expire, the debounce that turns a run of + taps into one write, the write
   order and what happens when a write fails, the Dashboard Progress upsert and
   its queue, the colour choice, the unsent-tap queue, and the parser's cp
   statuses.
   Graph is a fake fetch() over an in-memory workbook - nothing leaves the box.
   Run: node test_checkpoints.js                                             */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const ExcelJS = require("exceljs");

/* ---------- browser shims ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, addEventListener() {} };
global.performance = { now: () => Date.now() };
global.ExcelJS = ExcelJS;
function stubEl() {
  const e = {
    style: {}, dataset: {}, textContent: "", innerHTML: "", value: "", disabled: false, hidden: false,
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild() {}, remove() {}, addEventListener() {}, removeEventListener() {}, focus() {},
    setSelectionRange() {}, getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => stubEl(), querySelectorAll: () => []
  };
  return e;
}
global.document = {
  documentElement: stubEl(), body: stubEl(), createElement: () => stubEl(), activeElement: null,
  querySelector: () => stubEl(), querySelectorAll: () => [], addEventListener() {}, removeEventListener() {}
};

/* ---------- a tiny Excel behind a fake Graph ---------- */
const BOOK = {};
const CALLS = [];
let FAIL_FILL = 0;                 // 403: refused outright, so no retry sleeps in the test
const kk = (r, c) => r + "|" + c;
const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
function sh(name) { return BOOK[name] || (BOOK[name] = { v: {}, fill: {} }); }
function bounds(s) {
  let maxR = 0, maxC = 0;
  Object.keys(s.v).forEach(k => { const p = k.split("|"); maxR = Math.max(maxR, +p[0]); maxC = Math.max(maxC, +p[1]); });
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
  if (method === "GET") {
    const values = [];
    for (let r = A.r1; r <= A.r2; r++) {
      const line = [];
      for (let c = A.c1; c <= A.c2; c++) line.push(s.v[kk(r, c)] == null ? "" : s.v[kk(r, c)]);
      values.push(line);
    }
    CALLS.push({ method, sheet: name, kind: "readRange", addr: mr[1] });
    return ok({ values: values });
  }
  if (method === "PATCH" && tail === "/format/fill") {
    if (FAIL_FILL && name === "Production") { FAIL_FILL--; return { status: 403, body: { error: { code: "AccessDenied" } } }; }
    for (let r = A.r1; r <= A.r2; r++) for (let c = A.c1; c <= A.c2; c++) s.fill[kk(r, c)] = body.color;
    CALLS.push({ method, sheet: name, kind: "fill", addr: mr[1], color: body.color });
    return ok({});
  }
  if (method === "PATCH" && tail === "") {
    (body.values || []).forEach((line, ri) => line.forEach((x, ci) => { s.v[kk(A.r1 + ri, A.c1 + ci)] = x; }));
    /* a PATCH with no values only sets number formats - not a row of data */
    CALLS.push({ method, sheet: name, kind: body.values ? "values" : "numberFormat", addr: mr[1], values: body.values });
    return ok({});
  }
  if (method === "PATCH") { CALLS.push({ method, sheet: name, kind: "format", addr: mr[1] }); return ok({}); }
  return { status: 404, body: { error: "no route " + method + " " + path } };
}
let BATCHES = 0;
global.fetch = async (url, init) => {
  const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
  const body = init && init.body ? JSON.parse(init.body) : null;
  let res;
  if (path === "/$batch") {
    BATCHES++;
    res = ok({ responses: body.requests.map(q => Object.assign({ id: q.id }, route(q.method, q.url, q.body))) });
  } else res = route(init.method, path, body);
  return { ok: res.status < 400, status: res.status, text: async () => JSON.stringify(res.body), arrayBuffer: async () => new ArrayBuffer(0) };
};

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

const TOASTS = [];
global.toast = (m, err) => TOASTS.push({ m: String(m), err: !!err });   // no DOM: just remember them

const settle = ms => new Promise(r => setTimeout(r, ms == null ? 40 : ms));
const fills = () => CALLS.filter(c => c.kind === "fill" && c.sheet === "Production");
/* A1:F1 is the header row written once when a sheet is first created. */
const progressWrites = () => CALLS.filter(c => c.kind === "values" && c.sheet === "Dashboard Progress" && c.addr !== "A1:F1");
const logWrites = () => CALLS.filter(c => c.kind === "values" && c.sheet === "Dashboard Log" && c.addr !== "A1:F1");
const reset = () => { CALLS.length = 0; BATCHES = 0; TOASTS.length = 0; };
const queued = () => JSON.parse(localStorage.getItem("cw_cpqueue") || "[]");
/** the stored count for one (job, item), read back out of the fake sheet */
function storedCount(job, item) {
  const s = BOOK["Dashboard Progress"]; if (!s) return null;
  for (let r = 2; r <= 200; r++)
    if (s.v[kk(r, 1)] === job && s.v[kk(r, 2)] === item) return s.v[kk(r, 3)];
  return null;
}

/* ---------- the job under test ---------- */
const PMAP = { qty: { wnd: 13, drs: 14 }, glass: { tg: 51, tuff: 52 },
               prod: { "7000 casement": { f: 20, s: 21, t: 22 } }, prodOrder: ["7000 casement"] };
const mkJob = (cp, extra) => Object.assign({
  id: "R0001", cust: "Ann", area: "Cork", wnd: 10, drs: 2, glass: { tg: 25, tuff: 11 },
  prods: [{ n: "7000 casement", f: 12, s: 8, t: 0, st: [] }], notes: [], sheets: ["Production"], src: {},
  dates: {}, cat: "active", blk: 3, seq: 1, stage: "floor", done: 0, urg: 0,
  cp: JSON.parse(JSON.stringify(cp))
}, extra || {});
/* put one job on the fake Production sheet at row 7 and in the app's list */
function useJob(job, map) {
  BOOK["Production"] = { v: {}, fill: {} };
  BOOK["Production"].v[kk(7, 3)] = job.id;
  global.__j = job; global.__m = map === undefined ? PMAP : map;
  vm.runInThisContext("ALL = [__j]; PRODMAP = __m; CHANGES = []; state.sel = null;");
  return byId(job.id);
}
/* age a held item so it looks older than the 180 s window */
const ageHold = (id, key, ms) => vm.runInThisContext(
  "PENDING['" + id + "'].t['" + key + "'] -= " + ms + "; PENDING['" + id + "'].at -= " + ms);

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ---- 1. the merge rule ---- */
  CP.cpSetProgress({ R0001: { win: { done: 3, total: 10 }, "glass:tg": { done: 6, total: 25 },
                              drs: { done: 6, total: 2 }, "prod:7000 casement:f": { done: 6, total: 12 } } });
  let j = mkJob({ win: "done", drs: "", glass: { tg: "process", tuff: "process" },
                  prod: { "7000 casement": { f: "process", s: "", t: "" } } });
  assert.deepStrictEqual(itemState(j, "win"), { done: 10, total: 10, status: "done" });
  pass("gold cell beats the stored 3 of 10: 10 of 10, done");
  assert.deepStrictEqual(itemState(j, "glass:tg"), { done: 6, total: 25, status: "process" });
  pass("yellow cell + stored 6 of 25: 6, in process");
  assert.deepStrictEqual(itemState(j, "glass:tuff"), { done: null, total: 11, status: "process" });
  pass("yellow cell with nothing stored: in progress, count unknown");
  assert.deepStrictEqual(itemState(j, "drs"), { done: 0, total: 2, status: "" });
  pass("white cell beats a stored count: 0 of 2");
  assert.strictEqual(itemState(j, "prod:7000 casement:t"), null, "an item with no total does not exist");
  assert.deepStrictEqual(itemState(j, "prod:7000 casement:f"), { done: 6, total: 12, status: "process" });
  pass("product F/S/T items behave the same");
  assert.deepStrictEqual(cpItems(j).map(x => x.key),
    ["win", "drs", "glass:tg", "glass:tuff", "prod:7000 casement:f", "prod:7000 casement:s"]);
  pass("only items with a total exist, in drawer order");

  /* ---- 2. held counts: they win, they expire one by one ---- */
  pend("R0001", { cp: { win: 4 } });
  let held = applyPending([j])[0];
  assert.strictEqual(held.cpDone.win, 4);
  assert.strictEqual(held.cp.win, "process", "a partial held count re-colours the cell yellow");
  assert.deepStrictEqual(itemState(held, "win"), { done: 4, total: 10, status: "process" });
  pass("a held count overrides Excel and the stored count");
  pend("R0001", { cp: { win: 10 } });
  assert.strictEqual(applyPending([j])[0].cp.win, "done");
  pend("R0001", { cp: { win: 0 } });
  held = applyPending([j])[0];
  assert.strictEqual(held.cp.win, "");
  assert.deepStrictEqual(itemState(held, "win"), { done: 0, total: 10, status: "" });
  pass("held 0 / total re-colour to white / gold");
  pend("R0001", { cp: { "glass:tg": 9 } });
  ageHold("R0001", "cp:win", 200000);                    // only the windows tick is old
  held = applyPending([j])[0];
  assert.strictEqual(held.cpDone.win, undefined, "the old hold is gone");
  assert.strictEqual(held.cpDone["glass:tg"], 9, "the newer one is untouched");
  assert.deepStrictEqual(itemState(held, "win"), { done: 10, total: 10, status: "done" });
  pass("holds expire one item at a time: a new tick does not extend an old one");
  pend("R0001", { cp: { "glass:tg": null } });
  assert.strictEqual(applyPending([j])[0].cpDone, undefined);
  pass("a null hold drops that item entirely");
  pend("R0001", { cp: { "glass:tg": 6 } });
  applyPending([j], true);                                // the file already says 6
  assert.strictEqual(applyPending([j])[0].cpDone, undefined, "the hold is let go once the file agrees");
  pass("a hold is released as soon as a fresh parse agrees with it");

  /* ---- 3. five rapid taps make one write ---- */
  CP.cpSetProgress({ R0001: { win: { done: 1, total: 10 } } });
  useJob(mkJob({ win: "process", drs: "", glass: {}, prod: {} }));
  reset();
  for (let i = 0; i < 5; i++) {
    const cur = itemState(byId("R0001"), "win");
    setItemProgress(byId("R0001"), "win", (cur.done == null ? 0 : cur.done) + 1);
  }
  assert.deepStrictEqual(itemState(byId("R0001"), "win"), { done: 6, total: 10, status: "process" },
    "the screen is already at 6 before anything is written");
  assert.strictEqual(CALLS.length, 0, "nothing is written while the taps are still coming");
  assert.strictEqual(queued().length, 1, "the unsent tap is remembered in case the tab closes");
  await settle(1400);
  assert.strictEqual(fills().length, 1, "exactly one fill written");
  assert.strictEqual(fills()[0].addr, "M7", "written on the job's own row, in the WND column only");
  assert.strictEqual(fills()[0].color, "#FFFF00", "6 of 10 is yellow");
  assert.strictEqual(progressWrites().length, 1, "exactly one progress row written");
  assert.deepStrictEqual(progressWrites()[0].values[0].slice(0, 4), ["R0001", "win", "6", "10"]);
  assert.strictEqual(logWrites().length, 1, "exactly one log line");
  assert.deepStrictEqual(logWrites()[0].values[0].slice(2, 6), ["R0001", "Windows", "1 of 10", "6 of 10"],
    "the from value is what it was before the burst of taps");
  assert(!CALLS.some(c => c.sheet === "Production" && c.kind !== "fill" && c.kind !== "readRange"),
    "the Production sheet gets fills and a row lookup, nothing else");
  assert.strictEqual(queued().length, 0, "the queue is empty once the write has gone");
  const order = CALLS.filter(c => c.kind === "fill" || (c.kind === "values" && c.sheet === "Dashboard Progress"));
  assert.strictEqual(order[0].sheet, "Dashboard Progress", "our own count is stored before the sheet is coloured");
  pass("5 rapid + taps: one fill, one progress row, one log line 1 of 10 -> 6 of 10");

  reset();
  setItemProgress(byId("R0001"), "win", 7);
  await settle(1400);
  assert.strictEqual(fills().length, 1, "a tap after the window is its own write");
  assert.strictEqual(progressWrites().length, 1);
  assert.deepStrictEqual(logWrites()[0].values[0].slice(3, 6), ["Windows", "6 of 10", "7 of 10"]);
  pass("a tap after the debounce window makes a second write, counting from 6");

  /* ---- 3b. when the write fails, the screen goes back ---- */
  reset(); FAIL_FILL = 1;
  setItemProgress(byId("R0001"), "win", 9);
  await settle(1400);
  assert.strictEqual(fills().length, 0, "the fill was refused");
  assert.strictEqual(progressWrites().length, 2, "the count was written, then put back");
  assert.deepStrictEqual(progressWrites()[1].values[0].slice(0, 4), ["R0001", "win", "7", "10"]);
  assert.strictEqual(logWrites().length, 0, "nothing is logged for a write that did not happen");
  assert(TOASTS.some(t => t.err), "the person is told");
  assert.deepStrictEqual(itemState(byId("R0001"), "win").done, 7, "the screen shows 7 again");
  assert.strictEqual(queued().length, 0, "a failed write is not owed for ever");
  pass("fill refused: the count is restored, nothing is logged, the screen goes back");

  reset(); FAIL_FILL = 1;
  CP.cpSetProgress({});                                   // nothing stored: "in progress, count unknown"
  useJob(mkJob({ win: "process", drs: "", glass: {}, prod: {} }));
  assert.strictEqual(itemState(byId("R0001"), "win").done, null);
  setItemProgress(byId("R0001"), "win", 1);
  await settle(1400);
  assert.strictEqual(itemState(byId("R0001"), "win").done, null,
    "back to in progress, count unknown - not held at 0");
  pass("fill refused on an unknown count: the hold is dropped, not set to zero");

  /* ---- 4. Dashboard Progress upsert and its queue ---- */
  CP.cpSetProgress({});
  reset();
  const r1 = await CW.saveProgress("R0001", "win", 7, 10, "abin");
  const r2 = await CW.saveProgress("R0001", "win", 8, 10, "abin");
  assert.strictEqual(r1, r2, "the same (Job, Item) is updated in place");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(r2, 3)], "8");
  const r3 = await CW.saveProgress("R0001", "drs", 1, 2, "abin");
  assert.strictEqual(r3, r2 + 1, "a new (Job, Item) pair appends");
  const both = await Promise.all([CW.saveProgress("R0002", "win", 1, 4, "abin"),
                                  CW.saveProgress("R0003", "win", 2, 6, "abin")]);
  assert.notStrictEqual(both[0], both[1], "two writes at the same moment take two different rows");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(both[0], 1)], "R0002");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(both[1], 1)], "R0003");
  pass("saveProgress upserts, and two concurrent calls cannot take the same row");
  const nb = BATCHES;
  const wrote = await CW.saveProgressMany("R0001", [
    { item: "win", done: 10, total: 10 }, { item: "drs", done: 2, total: 2 },
    { item: "glass:tg", done: 25, total: 25 }], "abin");
  assert.strictEqual(wrote, 3);
  assert.strictEqual(BATCHES - nb, 1, "saveProgressMany sends one batch");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(r2, 3)], "10", "win updated in its existing row");
  assert.strictEqual(BOOK["Dashboard Progress"].v[kk(r3, 3)], "2", "drs updated in its existing row");
  pass("saveProgressMany writes every row in one batch");

  /* ---- 5. the colour of a count, and what is not a count ---- */
  assert.strictEqual(cpColour(0, 10), "#FFFFFF");
  assert.strictEqual(cpColour(1, 10), "#FFFF00");
  assert.strictEqual(cpColour(9, 10), "#FFFF00");
  assert.strictEqual(cpColour(10, 10), "#FFE699");
  assert.strictEqual(cpClamp(-4, 10), 0);
  assert.strictEqual(cpClamp(99, 10), 10);
  assert.strictEqual(cpClamp("4.6", 10), 5);
  assert.strictEqual(cpClamp("", 10), null, "a blank box is not a zero");
  assert.strictEqual(cpClamp("   ", 10), null);
  assert.strictEqual(cpClamp("rubbish", 10), null);
  pass("0 white, part yellow, all gold; blank or unreadable is not a count");

  CP.cpSetProgress({ R0001: { win: { done: 4, total: 10 } } });
  useJob(mkJob({ win: "process", drs: "", glass: {}, prod: {} }));
  reset();
  setItemProgress(byId("R0001"), "win", "");
  setItemProgress(byId("R0001"), "win", "abc");
  await settle(1400);
  assert.strictEqual(CALLS.length, 0, "a blank or unreadable box writes nothing");
  assert.strictEqual(itemState(byId("R0001"), "win").done, 4, "and changes nothing");
  pass("clearing the number box writes nothing at all");

  /* ---- 5b. an item that is not on the sheet ---- */
  useJob(mkJob({ win: "process", drs: "", glass: {}, prod: {} }), null);   // no PRODMAP
  reset();
  setItemProgress(byId("R0001"), "win", 5);
  await settle(900);
  assert.strictEqual(CALLS.length, 0, "nothing is written when the column is unknown");
  assert(TOASTS.some(t => t.err && /not a column/.test(t.m)), "and it says so");
  pass("no column on the Production sheet: no write, a plain message");

  /* ---- 5c. a whole group in one go ---- */
  CP.cpSetProgress({});
  useJob(mkJob({ win: "", drs: "", glass: {}, prod: {} }));
  reset();
  await setGroupDone(byId("R0001"), "glass", true);
  await settle(60);
  assert.deepStrictEqual(fills().map(f => f.addr + " " + f.color).sort(),
    ["AY7 #FFE699", "AZ7 #FFE699"], "both glass cells, gold, on row 7 only");
  assert.strictEqual(progressWrites().length, 2, "both counts stored");
  assert.strictEqual(BATCHES, 2, "one batch of fills, one batch of progress rows");
  assert.strictEqual(logWrites().length, 1, "one log line for the group");
  assert.deepStrictEqual(logWrites()[0].values[0].slice(3, 6), ["Glass: all done", "", "TG 25, TUFF 11"]);
  const first = CALLS.filter(c => c.kind === "fill" || (c.kind === "values" && c.sheet === "Dashboard Progress"))[0];
  assert.strictEqual(first.sheet, "Dashboard Progress", "the counts are stored before the sheet is coloured");
  pass("All glass done: one batch of fills, both counts in one batch, one log line");

  /* ---- 5d. a job already marked ready to deliver ---- */
  const gold = mkJob({ win: "done", drs: "done", glass: { tg: "done", tuff: "done" },
                       prod: { "7000 casement": { f: "done", s: "done" } } }, { done: 1 });
  useJob(gold);
  reset();
  const html = cpSectionHtml(byId("R0001"), true);
  assert(/Marked ready to deliver/.test(html), "it says why it cannot be edited");
  assert(!/data-cpgrp/.test(html), "no group buttons");
  assert(!/>Clear</.test(html), "nothing invites clearing a finished row");
  const controls = (html.match(/<button/g) || []).length + (html.match(/<input/g) || []).length;
  assert.strictEqual(controls, (html.match(/ disabled/g) || []).length, "every control is disabled");
  setItemProgress(byId("R0001"), "win", 3);
  await setGroupDone(byId("R0001"), "glass", false);
  await settle(900);
  assert.strictEqual(CALLS.length, 0, "and clicking writes nothing");
  pass("ready to deliver: the section is read-only and nothing can be written");

  /* ---- 5e. the open drawer is patched, not rebuilt ---- */
  CP.cpSetProgress({ R0001: { win: { done: 6, total: 10 } } });
  useJob(mkJob({ win: "process", drs: "", glass: { tg: "done", tuff: "done" }, prod: {} }));
  const mkLine = item => {
    const num = { textContent: "" }, bar = { style: {} }, inp = { value: "" }, all = { textContent: "", dataset: {} };
    return { dataset: { cpline: item }, _num: num, _bar: bar, _inp: inp, _all: all,
             querySelector: sel => sel === ".cpnum" ? num : sel === ".cpbar i" ? bar
               : sel === ".cpin" ? inp : sel === ".cpall[data-cp]" ? all : null };
  };
  const sum = { textContent: "" }, line = mkLine("win"), tgLine = mkLine("glass:tg");
  const grpBtn = { textContent: "", dataset: { cpgrp: "glass", alltext: "All glass done" } };
  const host = { querySelector: sel => (sel === ".cpsum" ? sum : null),
                 querySelectorAll: sel => sel === "[data-cpline]" ? [line, tgLine] : sel === "[data-cpgrp]" ? [grpBtn] : [] };
  const realQ = document.querySelector;
  document.querySelector = sel => (sel === "#dhost" ? host : stubEl());
  cpPatchSection(byId("R0001"));
  assert.strictEqual(line._num.textContent, "6 of 10");
  assert.strictEqual(line._bar.style.width, "60%");
  assert.strictEqual(line._bar.style.background, "var(--fab)");
  assert.strictEqual(line._inp.value, 6);
  assert.strictEqual(line._all.textContent, "All done");
  assert.strictEqual(tgLine._bar.style.background, "var(--done)", "a finished item is gold");
  assert.strictEqual(grpBtn.textContent, "Clear", "the glass group is complete");
  assert.strictEqual(sum.textContent, "Windows 6/10 · Doors 0/2 · Glass 36/36 · Frames/Sashes/Transoms 0/20");
  document.activeElement = line._inp;                    // someone is typing in that box
  line._inp.value = "4";
  pend("R0001", { cp: { win: 8 } });
  vm.runInThisContext("ALL = applyPending(ALL)");
  cpPatchSection(byId("R0001"));
  assert.strictEqual(line._inp.value, "4", "the box being typed into is left alone");
  assert.strictEqual(line._num.textContent, "8 of 10", "everything else still updates");
  document.activeElement = null; document.querySelector = realQ;
  pend("R0001", { cp: { win: null } });
  pass("a tap patches the open section in place instead of rebuilding it");

  /* ---- 6. unsent taps survive the tab closing ---- */
  CP.cpSetProgress({ R0001: { win: { done: 1, total: 10 } } });
  useJob(mkJob({ win: "process", drs: "", glass: {}, prod: {} }));
  reset();
  setItemProgress(byId("R0001"), "win", 3);
  assert.strictEqual(queued().length, 1);
  CP.cpFireAll();                                  // what pagehide does
  await settle(60);
  assert.strictEqual(fills().length, 1, "the write goes out at once, not 800 ms later");
  assert.strictEqual(logWrites().length, 1);
  assert.strictEqual(queued().length, 0);
  pass("leaving the page sends the taps that were still waiting");

  reset();
  localStorage.setItem("cw_cpqueue", JSON.stringify([
    { key: "R0001|glass:tg", job: "R0001", item: "glass:tg", col: 51, from: 0, to: 12, total: 25, at: Date.now() - 60000 },
    { key: "R0009|win", job: "R0009", item: "win", col: 13, from: 0, to: 2, total: 4, at: Date.now() - 7200000 }
  ]));
  assert.strictEqual(cpReplayQueue(), 1, "only the entry from within the hour is replayed");
  await settle(120);
  assert.strictEqual(fills().length, 1);
  assert.strictEqual(fills()[0].addr, "AY7");
  assert.strictEqual(fills()[0].color, "#FFFF00");
  assert.deepStrictEqual(progressWrites()[0].values[0].slice(0, 4), ["R0001", "glass:tg", "12", "25"]);
  assert.deepStrictEqual(logWrites()[0].values[0].slice(3, 6), ["Glass TG", "0 of 25", "12 of 25"]);
  assert.strictEqual(queued().length, 0, "and it is not owed twice");
  pass("a tap left over from a previous visit is sent on the next load");

  /* ---- 6b. two group taps in a row ---- */
  CP.cpSetProgress({});
  useJob(mkJob({ win: "", drs: "", glass: {}, prod: {} }));
  reset();
  const p1 = setGroupDone(byId("R0001"), "glass", true);
  const midFlight = cpSectionHtml(byId("R0001"), true);
  assert(/data-cpgrp="glass"[^>]* disabled/.test(midFlight), "the group's button is dead while its write is in the air");
  const p2 = setGroupDone(byId("R0001"), "glass", false);   // the double tap
  await Promise.all([p1, p2]);
  await settle(60);
  const seq = CALLS.filter(c => (c.kind === "fill" && c.sheet === "Production") ||
                                (c.kind === "values" && c.sheet === "Dashboard Progress" && c.addr !== "A1:F1"))
    .map(c => c.kind === "fill" ? "fill " + c.color : "count " + c.values[0][2]);
  assert.deepStrictEqual(seq, ["count 25", "count 11", "fill #FFE699", "fill #FFE699",
                               "count 0", "count 0", "fill #FFFFFF", "fill #FFFFFF"],
    "the two writes run one after the other, never interleaved");
  assert.strictEqual(BOOK["Production"].fill[kk(7, 51)], "#FFFFFF", "Excel ends on the second tap");
  assert.strictEqual(storedCount("R0001", "glass:tg"), "0", "and so does the stored count");
  assert(!/data-cpgrp="glass"[^>]* disabled/.test(cpSectionHtml(byId("R0001"), true)), "the button comes back after");
  pass("two group taps: strictly one after the other, the second one wins in both places");

  /* ---- 6c. the column is worked out when the write goes, not when it was tapped ---- */
  CP.cpSetProgress({});
  useJob(mkJob({ win: "", drs: "", glass: {}, prod: {} }));
  reset();
  setItemProgress(byId("R0001"), "win", 3);
  global.__m2 = { qty: { wnd: 15, drs: 16 }, glass: { tg: 51, tuff: 52 }, prod: {}, prodOrder: [] };
  vm.runInThisContext("PRODMAP = __m2");            // someone inserted two columns in Excel
  await settle(1400);
  assert.strictEqual(fills().length, 1);
  assert.strictEqual(fills()[0].addr, "O7", "the fill goes to the column the sheet has now");
  pass("a column inserted between the tap and the write does not send the fill astray");

  useJob(mkJob({ win: "", drs: "", glass: {}, prod: {} }));
  reset();
  setItemProgress(byId("R0001"), "win", 4);
  vm.runInThisContext("PRODMAP = null");            // the sheet could not be read
  await settle(1400);
  assert.strictEqual(CALLS.length, 0, "nothing is written when the column cannot be worked out");
  assert(TOASTS.some(t => t.err && /not a column/.test(t.m)));
  assert.strictEqual(itemState(byId("R0001"), "win").done, 0, "and the tick is not left hanging");
  pass("no PRODMAP when the write goes: nothing written, the hold dropped, a message");

  /* ---- 6d. what is replayed and what is not ---- */
  CP.cpSetProgress({});
  useJob(mkJob({ win: "", drs: "", glass: {}, prod: {} }));
  reset();
  const at = Date.now() - 30000;
  localStorage.setItem("cw_cpqueue", JSON.stringify([
    { key: "R0001|win", job: "R0001", item: "win", col: 13, from: 0, to: 3, total: 10, at: at, sent: 1 },
    { key: "R0001|drs", job: "R0001", item: "drs", col: 14, from: 0, to: 1, total: 2, at: at, sent: 0 },
    { key: "R0001|glass:tg", job: "R0001", item: "glass:tg", col: 51, from: 0, to: 5, total: 25, at: at, who: "someone.else@costellowindows.ie" }
  ]));
  assert.strictEqual(cpReplayQueue(), 1, "only the unsent entry belonging to this person");
  await settle(120);
  assert.strictEqual(fills().length, 1);
  assert.strictEqual(fills()[0].addr, "N7", "the doors tick, nothing else");
  pass("a write already sent is never replayed, nor one left by another account");

  /* ---- 6e. undoing ready to deliver paints white ---- */
  useJob(mkJob({ win: "", drs: "", glass: {}, prod: {} }, { done: 1 }));
  reset();
  await markReady(byId("R0001"), false);
  assert.strictEqual(fills().length, 1);
  assert.strictEqual(fills()[0].addr, "A7:CL7");
  assert.strictEqual(fills()[0].color, "#FFFFFF", "white, not stripped: the sheet's cells are white, not blank");
  pass("undo of mark ready paints the row white instead of clearing the fill");

  /* ---- 6f. the list keeps its section names, and a group tick is not reported twice ---- */
  const withNames = [mkJob({ win: "", drs: "", glass: {}, prod: {} })];
  withNames.blockNames = ["Second hand", "In production"];
  assert.deepStrictEqual(applyPending(withNames).blockNames, ["Second hand", "In production"]);
  pass("applyPending keeps the section names on the list");

  const t0 = Date.now();
  const iso = ms => new Date(t0 - ms).toISOString();
  const mineChanges = [
    { src: "dashboard", at: iso(60000), job: "R0001", what: "Glass: all done" },
    { src: "dashboard", at: iso(60000), job: "R0001", what: "7000 Casement: all done" },
    { src: "dashboard", at: iso(60000), job: "R0001", what: "Windows" },
    { src: "dashboard", at: iso(900000), job: "R0002", what: "Glass: all done" }   // too old to count
  ];
  const diffs = [
    { job: "R0001", what: "Glass TG" }, { job: "R0001", what: "Glass NOT TUFF" },
    { job: "R0001", what: "7000 CASEMENT frames" }, { job: "R0001", what: "Windows" },
    { job: "R0001", what: "Customer" }, { job: "R0002", what: "Glass TG" }
  ];
  assert.deepStrictEqual(dropMine(diffs, mineChanges, t0).map(c => c.job + " " + c.what),
    ["R0001 Customer", "R0002 Glass TG"]);
  pass("a group's log line also covers the per-item differences it caused");

  /* ---- 7. the parser reads the statuses out of the fills ---- */
  const wb = new ExcelJS.Workbook();
  const head = ws => {
    const put = (r, c, v) => { ws.getCell(r, c).value = v; };
    put(2, 4, "DATES ON CONTRACT"); put(3, 4, "SOLD"); put(2, 5, "CUSTOMER");
    put(2, 13, "QUANTITY"); put(3, 13, "WND"); put(3, 14, "DRS");
    put(2, 20, "7000 CASEMENT"); put(3, 20, "F"); put(3, 21, "S"); put(3, 22, "T");
    put(2, 51, "GLASS UNITS"); put(3, 51, "TG"); put(3, 52, "TUFF");
  };
  const ws = wb.addWorksheet("Production"), ws2 = wb.addWorksheet("Production (2)");
  head(ws); head(ws2);
  const put = (sheet, r, c, v, colour) => {
    const cell = sheet.getCell(r, c);
    if (v !== undefined && v !== null) cell.value = v;
    if (colour) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + colour } };
  };
  /* R0001 on Production: WND yellow, DRS gold, TG gold, TUFF white, F yellow, S gold, T white */
  put(ws, 6, 3, "R0001"); put(ws, 6, 5, "Ann");
  put(ws, 6, 13, 10, "FFFF00"); put(ws, 6, 14, 2, "FFE699");
  put(ws, 6, 20, 12, "FFFF00"); put(ws, 6, 21, 8, "FFE699"); put(ws, 6, 22, 3, "FFFFFF");
  put(ws, 6, 51, 25, "FFE699"); put(ws, 6, 52, 11, "FFFFFF");
  /* the same job on Production (2), still yellow where Production has been cleared */
  put(ws2, 6, 3, "R0001"); put(ws2, 6, 13, 10, "FFFF00"); put(ws2, 6, 52, 11, "FFFF00");
  put(ws2, 6, 20, 12, "FFE699");
  /* R0002: the whole Production row is gold */
  put(ws, 7, 3, "R0002"); put(ws, 7, 13, 4); put(ws, 7, 20, 5); put(ws, 7, 51, 6);
  for (let c = 1; c <= 60; c++) put(ws, 7, c, undefined, "FFE699");
  const jobs = parseWorkbook(wb);
  const a = jobs.find(x => x.id === "R0001"), b = jobs.find(x => x.id === "R0002");
  assert.strictEqual(a.cp.win, "process");
  assert.strictEqual(a.cp.drs, "done");
  assert.strictEqual(a.cp.glass.tg, "done");
  assert.strictEqual(cpStatus(a, "glass:tuff"), "", "a white cell leaves the item untouched");
  assert.strictEqual(cpStatus(a, "prod:7000 casement:f"), "process");
  assert.strictEqual(cpStatus(a, "prod:7000 casement:s"), "done");
  assert.strictEqual(cpStatus(a, "prod:7000 casement:t"), "");
  assert.deepStrictEqual(a.prods[0].st.sort(), ["done", "process"], "the old product status still works");
  pass("parser: cp statuses for M, N, the glass columns and F/S/T");
  pass("a colour on Production (2) never overrides Production, so un-ticking sticks");
  CP.cpSetProgress({});
  assert.deepStrictEqual(itemState(a, "win"), { done: null, total: 10, status: "process" });
  assert.deepStrictEqual(itemState(a, "drs"), { done: 2, total: 2, status: "done" });
  assert.strictEqual(b.done, 1, "the gold row is a finished job");
  assert.strictEqual(b.cp.win, "done");
  assert.strictEqual(b.cp.glass.tg, "done");
  assert.strictEqual(b.cp.prod["7000 casement"].f, "done");
  cpItems(b).forEach(x => assert.strictEqual(itemState(b, x.key).status, "done", x.key + " is done"));
  pass("a whole gold row counts as done for every checkpoint on it");

  /* ---- 8. what the Changes list says about a tick ---- */
  const before2 = JSON.parse(JSON.stringify(a)), after = JSON.parse(JSON.stringify(a));
  after.cp.win = "done";
  const d = diffJobs([before2], [after], "someone", "2026-09-04T10:00:00Z");
  assert.deepStrictEqual(d.map(x => [x.what, x.from, x.to]), [["Windows", "in fabrication", "done"]]);
  pass("a colour change is listed by the item's own name, the way the log names it");

  console.log("\n" + n + " checks passed");
  process.exit(0);                 // the 45 s reconcile timer would hold the process open
})().catch(e => { console.error("FAIL", e); process.exit(1); });
