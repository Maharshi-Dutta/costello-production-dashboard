/* Offline test of the doors and of the Windows/Doors folds.
   Spec: docs/specs/2026-09-11-doors-and-window-types.md (owner, 2026-09-11;
   approved to build 2026-09-14).

   What this file is about, and it is mostly about what is NOT done:

     · a door is one DOORS DONE cell, its TEXT the type code and its FILL the
       status. Any non-empty text is a door - the live sheet carries CD, DD,
       SS, SD, PVC, DOOR, ACSS, ACSD, SFCD and "1 DOOR" - so there is no list
       of known codes here to fall behind what the office types;
     · the drawer draws one row per door with In fabrication / Done / Clear and
       highlights the one it is on, inside a Doors fold that remembers whether
       it was left open; the window types live in a Windows fold beside it;
     · Windows and Doors keep their own tick AND answer for what is under them
       (amendment 1, after the owner withdrew the first build's assumption):
       the HIGHER of the two is shown, the paint only ever goes upwards, and
       their numbers on the home list carry that colour with the words in the
       hover. A doors quantity that disagrees with the coded cells is said in
       plain words in both places and changes no paint;
     · what Excel says reaches the dashboard on the next load, through the
       imports and the safeguard, and never as status read off a cell.

   `global.fetch` throws here. Nothing in this file may touch the network, and
   a write that slipped out of a click would be caught by that rather than by
   an assertion somebody remembered to write.
   Every name and code below is invented.   Run: node test_doors.js         */
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
  const e = {
    style: {}, dataset: {}, attrs: {}, textContent: "", innerHTML: "", value: "", disabled: false,
    hidden: false, open: false, className: "",
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild() {}, remove() {}, addEventListener() {}, removeEventListener() {}, focus() {},
    setAttribute(k, v) { e.attrs[k] = String(v); }, getAttribute(k) { return e.attrs[k]; },
    setSelectionRange() {}, getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => null, querySelectorAll: () => []
  };
  return e;
}
global.document = {
  documentElement: stubEl(), body: stubEl(), createElement: () => stubEl(), activeElement: null,
  querySelector: () => stubEl(), querySelectorAll: () => [], addEventListener() {}, removeEventListener() {}
};
/* NO NETWORK. Not a stub that records: a throw, so nothing can be "nearly" offline. */
let FETCHED = 0;
global.fetch = async () => { FETCHED++; throw new Error("test_doors made a network call"); };

const run = f => vm.runInThisContext(fs.readFileSync(__dirname + "/" + f, "utf8"), { filename: f });
run("parser.js");
run("graph.js");
global.CW = window.CW;
run("checkpoints.js");
global.CP = window.CP;
run("station-core.js");
global.ST = window.ST;
run("app.js");
const TOASTS = [];
global.toast = (m, err) => TOASTS.push({ m: String(m), err: !!err });

/* ---------- the fixture ---------- */
const PMAP = { qty: { wnd: 13, drs: 14 }, glass: { tg: 51 },
               prod: { "polaris 85mm casement": { f: 20, s: 21, t: 22 },
                       "4000 casement": { f: 23, s: 24, t: 25 } },
               prodOrder: ["polaris 85mm casement", "4000 casement"],
               doors: { 1: 72, 2: 73, 3: 74, 4: 75, 5: 76 } };
const mkJob = extra => Object.assign({
  id: "R5001", cust: "Ann", area: "Cork", eir: "", off: "", colour: "", ph3: "", ph: "",
  flag: "", flagHex: "", wnd: 9, drs: 3, glass: { tg: 4 },
  prods: [{ n: "polaris 85mm casement", f: 5, s: 4, t: 0, st: [] },
          { n: "4000 casement", f: 4, s: 3, t: 0, st: [] }],
  doors: [{ slot: 1, code: "CD", status: "" }, { slot: 2, code: "CD", status: "" },
          { slot: 4, code: "SS", status: "" }],
  notes: [], sheets: ["Production"], src: { Production: 7 },
  dates: { sold: null, stamp: null, ivana: null, ready: null, floor: null },
  cp: { win: "", drs: "", glass: {}, prod: {} },
  cat: "active", blk: 4, seq: 1, stage: "office", done: 0, urg: 0
}, extra || {});

/** Put the fixture in front of the page, with the record answering. */
function useJob(job, rows) {
  global.__j = job; global.__m = PMAP;
  global.__rows = (rows || []).map((r, i) => ({ id: "r" + i,
    fields: cpRowFields(r.job || job.id, r.item, r.done, r.total, r.status,
                        r.who || "the admin", r.when || "2026-09-11T09:00:00Z", r.source || "office") }));
  vm.runInThisContext("ALL = [__j]; PRODMAP = __m; CHANGES = []; state.sel = null; state.edit = false; " +
    "PENDING = {}; savePending(); CP_ITEMS = __rows; cpListRebuild(); " +
    "CP_LIST_OK = true; CP_LIST_WHY = ''; CP_IMPORTED = '2026-09-11T09:00:00Z'; cpImportCheck();");
  return byId(job.id);
}
/* the drawer's HTML, read back with small regexes - the same trick the other
   suites use: what is asserted is the string the page really writes */
const between = (html, start, end) => {
  const a = html.indexOf(start);
  if (a < 0) return "";
  const b = html.indexOf(end, a + start.length);
  return html.slice(a, b < 0 ? html.length : b);
};
const doorRows = html => {
  const out = [];
  const re = /<div class="cpdoor" data-cpline="(door:\d)"[\s\S]*?<span class="cplab cpcode">([^<]*)<\/span><span class="cpdw">([^<]*)<\/span>([\s\S]*?)<\/div>/g;
  let m;
  while ((m = re.exec(html))) {
    const btns = [];
    const bre = /<button class="cpdoorb( on)?" data-door="[^"]*" data-act="([^"]*)"([^>]*)>([^<]*)<\/button>/g;
    let b;
    while ((b = bre.exec(m[4]))) btns.push({ on: !!b[1], act: b[2], attrs: b[3], label: b[4] });
    out.push({ item: m[1], code: m[2], word: m[3], btns: btns });
  }
  return out;
};
/** Every door button in a drawer's html, as elements a wiring pass can click. */
const doorButtons = html => {
  const out = [];
  const re = /data-door="([^"]*)" data-act="([^"]*)"/g;
  let m;
  while ((m = re.exec(html))) {
    const el = stubEl();
    el.dataset.door = m[1]; el.dataset.act = m[2];
    out.push(el);
  }
  return out;
};

(async () => {
  let n = 0; const pass = t => { n++; console.log("  ok  " + t); };

  /* ================ 1. the drawer: one row per door ==================== */
  let j = useJob(mkJob(), [
    { item: "door:1", done: 1, total: 1, status: "done" },
    { item: "door:2", done: 0, total: 1, status: "process" }
  ]);
  vm.runInThisContext("state.edit = true;");
  let html = cpSectionHtml(j, true);
  const rows = doorRows(html);
  assert.deepStrictEqual(rows.map(r => [r.item, r.code, r.word]),
    [["door:1", "CD", "done"], ["door:2", "CD", "in fabrication"], ["door:4", "SS", "not started"]]);
  pass("one row per door, in slot order, each with its own code and the word it stands at");
  assert.deepStrictEqual(rows[1].btns.map(b => b.label), ["In fabrication", "Done", "Clear"]);
  assert.deepStrictEqual(rows[1].btns.map(b => b.act), ["process", "done", ""]);
  pass("two buttons and a Clear, in that order, on every door");
  assert.deepStrictEqual(rows.map(r => (r.btns.filter(b => b.on)[0] || {}).act), ["done", "process", ""],
    "the one it is on is highlighted, and only that one");
  assert.ok(rows[0].btns.filter(b => b.on)[0].attrs.indexOf('aria-pressed="true"') >= 0,
    "and says so to a screen reader as well as in colour");
  pass("the button the door is on is the one marked, in colour and in aria-pressed");
  /* two doors of the same job carry the same code, deliberately: the live
     sheet has CD in four cells of one job. They are told apart by their slot. */
  assert.strictEqual(rows[0].code, rows[1].code, "two doors of one job may read the same");
  assert.notStrictEqual(rows[0].item, rows[1].item, "and are still two different doors");
  pass("two cells carrying the same code are two doors, told apart by their slot");

  /* the buttons are thumb sized in the stylesheet, not by hope */
  const css = fs.readFileSync(__dirname + "/index.html", "utf8");
  const doorCss = between(css, ".cpdoorb {", "}");
  assert.ok(/min-height:40px/.test(doorCss) && /min-width:40px/.test(doorCss),
    "a door button is at least 40px both ways: " + doorCss);
  pass("door buttons are thumb sized, because the drawer is used on a phone");

  /* ================ 2. the Windows and Doors folds ===================== */
  assert.ok(/<details class="cpfold" data-cpfold="win"/.test(html), "a Windows fold");
  assert.ok(/<details class="cpfold" data-cpfold="drs"/.test(html), "a Doors fold");
  assert.ok(!/data-cpfold="win" open/.test(html) && !/data-cpfold="drs" open/.test(html),
    "closed by default: opening them is the reader's decision, not the page's");
  pass("Windows and Doors are folds, closed until somebody opens them");
  const winFold = between(html, 'data-cpfold="win"', 'class="cpline" data-cpline="drs"');
  assert.ok(/data-cpgrp="prod:polaris 85mm casement"/.test(winFold) &&
            /data-cpgrp="prod:4000 casement"/.test(winFold),
    "every window type of the job is inside the Windows fold");
  assert.ok(!/data-cpgrp="prod:/.test(between(html, 'data-cpfold="drs"', 'class="cpgrp"')),
    "and none of them is left loose beside it");
  assert.ok(/<span class="cpfoldn">2 window types<\/span>/.test(winFold),
    "the fold says how many are inside it");
  assert.ok(/<span class="cpfoldn">3 doors<\/span>/.test(html));
  pass("the window types live under Windows, the doors under Doors, each counted on its head");
  /* AMENDMENT 1 (2026-09-14, the owner withdrew assumption (a)): both lines
     are still the office's own tick, with their own stepper and their own
     All done / Clear. The fold below each of them carries no control at all,
     because a click anywhere in a <summary> opens and closes it. */
  assert.ok(/<div class="cpline" data-cpline="win"/.test(html) &&
            /<div class="cpline" data-cpline="drs"/.test(html),
    "Windows and Doors each have a line of their own");
  assert.ok(/data-cp="win" data-act="inc"/.test(html) && /data-cpin="win"/.test(html) &&
            /data-cp="win" data-act="all"/.test(html),
    "with the stepper and the All done they always had");
  assert.ok(/data-cp="drs" data-act="inc"/.test(html) && /data-cpin="drs"/.test(html));
  const winHead = between(winFold, "<summary", "</summary>");
  assert.ok(!/<button/.test(winHead) && !/<input/.test(winHead),
    "and the fold head under it has nothing to click but itself: " + winHead);
  pass("Windows and Doors are ticked in their own right; the folds below them are headings");

  /* the open/closed state is remembered in this browser, and nowhere else */
  const foldEl = Object.assign(stubEl(), { dataset: { cpfold: "drs" }, open: true });
  const host = { querySelectorAll: sel => sel === "[data-cpfold]" ? [foldEl] : [],
                 querySelector: () => null };
  wireCheckpoints(host, "R5001");
  assert.strictEqual(typeof foldEl.ontoggle, "function", "the fold reports when it is opened");
  pass("a fold's open/closed state is wired to be remembered");
  cpFoldSet("drs", true);
  assert.ok(/data-cpfold="drs" open/.test(cpSectionHtml(byId("R5001"), true)),
    "and the next drawer opens it again");
  assert.strictEqual(JSON.parse(localStorage.getItem("cw_cpopen")).drs, 1,
    "kept in this browser only, in one key, and nothing about a job is in it");
  cpFoldSet("drs", false);
  pass("what is remembered is one open/closed flag per fold, in this browser");

  /* ================ 3. a click writes the word, not a count ============ */
  useJob(mkJob(), []);
  vm.runInThisContext("state.edit = true;");
  const btns = doorButtons(cpSectionHtml(byId("R5001"), true));
  const host2 = { querySelectorAll: sel => sel === "[data-door]" ? btns : [], querySelector: () => null };
  wireCheckpoints(host2, "R5001");
  const fab = btns.filter(b => b.dataset.door === "door:4" && b.dataset.act === "process")[0];
  assert.ok(fab && typeof fab.onclick === "function", "the In fabrication button is wired");
  fab.onclick();
  assert.deepStrictEqual(itemState(byId("R5001"), "door:4"), { done: null, total: 1, status: "process" },
    "the record moves at the click, before anything goes out");
  const owed = CP.cpPending("R5001|door:4");
  assert.ok(owed, "one write is owed");
  assert.strictEqual(owed.status, "process", "and it carries the WORD, because a door has no count");
  assert.strictEqual(owed.was, "", "and the word it started from, for the log line");
  assert.strictEqual(owed.col, 75, "aimed at that door's own DOORS DONE column and no other");
  assert.strictEqual(owed.total, 1);
  CP.cpCancelBurst("R5001|door:4");
  assert.strictEqual(FETCHED, 0, "and not one request has gone out yet");
  pass("a door click records the word at once and owes exactly one write, at that cell");

  /* the aggregate follows it, with no tick and no row of its own */
  assert.deepStrictEqual(itemState(byId("R5001"), "drs"), { done: null, total: 3, status: "process" },
    "one door in fabrication makes the Doors line read in fabrication");
  assert.strictEqual(cpRow("R5001", "drs"), null,
    "and the door's own write created no row for the Doors line: the line is carried");
  pass("a door in fabrication carries the Doors line without writing a row for it");

  /* ================ 4. the home list: colour on the numbers =========== */
  /* nothing done: both numbers plain, and nothing to hover on the doors */
  useJob(mkJob(), []);
  let cell = wndDrsCell(byId("R5001"));
  assert.ok(/>9<\/span> \/ <span[^>]*>3</.test(cell), "the two numbers, unchanged: " + cell);
  assert.strictEqual((cell.match(/var\(--ink-2\)/g) || []).length, 3,
    "both of them plain while nothing has been ticked");
  assert.ok(/Doors: CD not started · CD not started · SS not started/.test(cell),
    "and the hover names every door and where it stands: " + cell);
  pass("nothing ticked: two plain numbers, the breakdown in the hover");

  useJob(mkJob(), [
    { item: "door:1", done: 1, total: 1, status: "done" },
    { item: "door:2", done: 0, total: 1, status: "process" },
    { item: "door:4", done: 1, total: 1, status: "done" },
    { item: "prod:polaris 85mm casement:f", done: 2, total: 5, status: "process" }
  ]);
  cell = wndDrsCell(byId("R5001"));
  assert.strictEqual((cell.match(/var\(--fab\)/g) || []).length, 2,
    "windows and doors are both part way, so both numbers go yellow: " + cell);
  assert.ok(/Windows: Polaris 85mm Casement in progress · 4000 Casement not started/.test(cell),
    "the windows breakdown is one phrase per window type: " + cell);
  assert.ok(/Doors: CD done · CD in fabrication · SS done/.test(cell),
    "and the doors breakdown is one phrase per door: " + cell);
  pass("part way: both numbers yellow, and the hover says which types and which doors");

  useJob(mkJob(), [
    { item: "door:1", done: 1, total: 1, status: "done" },
    { item: "door:2", done: 1, total: 1, status: "done" },
    { item: "door:4", done: 1, total: 1, status: "done" }
  ]);
  cell = wndDrsCell(byId("R5001"));
  assert.ok(/var\(--done\)/.test(cell) && /font-weight:600/.test(cell),
    "every door done: the Doors number goes gold: " + cell);
  assert.ok(/var\(--ink-2\)/.test(cell), "and the Windows number is still plain");
  pass("every door done takes the Doors number gold on its own, without the windows");

  /* a job with no doors and no window types, and no doors quantity to
     disagree with: exactly as the cell looked before this existed */
  useJob(mkJob({ doors: [], prods: [], drs: 0 }), []);
  cell = wndDrsCell(byId("R5001"));
  assert.strictEqual(cell.indexOf("title="), -1, "nothing under either line: nothing to hover");
  assert.ok(/>9<\/span> \/ <span[^>]*>0</.test(cell), "and the two numbers are exactly as they were");
  pass("a job with nothing under either line shows two plain numbers and no hover");

  /* ---- the quantity warning (amendment 6; owner: "that mean the quantity is
     wrong and should give a warning") ---- */
  useJob(mkJob({ doors: [{ slot: 1, code: "CD", status: "" }, { slot: 2, code: "SS", status: "" }] }), []);
  assert.strictEqual(cpDoorWarning(byId("R5001")), "quantity says 3 · 2 doors listed");
  assert.ok(wndDrsCell(byId("R5001")).indexOf("quantity says 3 · 2 doors listed") >= 0,
    "the home list's hover carries the same words");
  vm.runInThisContext("state.edit = true;");
  const warned = cpSectionHtml(byId("R5001"), true);
  assert.ok(/<div class="cpline" data-cpline="drs"[\s\S]*?<span class="cpwarn">quantity says 3 · 2 doors listed<\/span>/.test(warned),
    "and so does the drawer's Doors line, in the warning style");
  pass("a doors quantity that disagrees with the coded cells is said in both places");

  /* quantity 0 with codes present: the same warning, and N stays unpainted -
     it is not a checkpoint at all when the sheet says there are no doors
     (amendment 10) */
  useJob(mkJob({ drs: 0 }), []);
  assert.strictEqual(cpDoorWarning(byId("R5001")), "quantity says 0 · 3 doors listed");
  assert.strictEqual(cpTotal(byId("R5001"), "drs"), 0, "so there is no Doors item");
  assert.strictEqual(itemState(byId("R5001"), "drs"), null);
  assert.deepStrictEqual(cpAggregatePlan("R5001").map(c => c.item), [],
    "and nothing plans a paint for N: quantity 0 leaves that cell alone");
  assert.ok(cpSectionHtml(byId("R5001"), true).indexOf("quantity says 0 · 3 doors listed") >= 0,
    "the warning still shows, on the fold head where the line would have been");
  pass("quantity 0 with codes typed: warned in words, and the N cell is never painted");

  /* the doors are told apart by their SLOT, so re-typing a code keeps the
     status that slot was at (amendment 8) */
  useJob(mkJob(), [{ item: "door:2", done: 1, total: 1, status: "done" }]);
  assert.deepStrictEqual(itemState(byId("R5001"), "door:2"), { done: 1, total: 1, status: "done" });
  byId("R5001").doors[1].code = "SFCD";                    // somebody re-types that cell
  assert.deepStrictEqual(itemState(byId("R5001"), "door:2"), { done: 1, total: 1, status: "done" },
    "the slot is the door: a new code in the same cell keeps that door's status");
  assert.strictEqual(doorRows(cpSectionHtml(byId("R5001"), true))[1].code, "SFCD",
    "and the drawer shows the new code");
  pass("a door is its slot: re-typing the code changes the label, not the status");

  /* the whole row still has the columns it had: no new one (owner's rule A6) */
  useJob(mkJob(), []);
  const row = rowHtml(byId("R5001"), 0, 10);
  assert.ok(row.indexOf(wndDrsCell(byId("R5001"))) >= 0, "the job row carries that same cell");
  assert.strictEqual(row.indexOf("Doors:"), row.lastIndexOf("Doors:"),
    "and says Doors once, in the hover, not in a column of its own");
  pass("no new column on the job row: the numbers that were there carry the colour");

  /* ================ 5. Excel -> the dashboard ========================== */
  /* a workbook with DOORS DONE filled in, parsed the way a load parses it */
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Production");
  const put = (r, c, v, colour) => {
    const cell2 = ws.getCell(r, c);
    if (v !== undefined && v !== null) cell2.value = v;
    if (colour) cell2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + colour } };
  };
  put(2, 4, "DATES ON CONTRACT"); put(3, 4, "SOLD"); put(2, 5, "CUSTOMER");
  put(2, 13, "QUANTITY"); put(3, 13, "WND"); put(3, 14, "DRS");
  put(2, 72, "DOORS DONE");
  for (let i = 0; i < 5; i++) put(3, 72 + i, String(i + 1));
  put(6, 3, "R5002"); put(6, 5, "Ann"); put(6, 13, 4); put(6, 14, 4);
  put(6, 72, "CD", "FFE699"); put(6, 73, " dd ", "FFFF00");
  put(6, 74, "ACSS"); put(6, 76, "PVC", "FFE699");
  const parsed = parseWorkbook(wb).find(x => x.id === "R5002");
  assert.deepStrictEqual(parsed.doors, [
    { slot: 1, code: "CD", status: "done" },
    { slot: 2, code: "DD", status: "process" },
    { slot: 3, code: "ACSS", status: "" },
    { slot: 5, code: "PVC", status: "done" }
  ], "four doors read off the row, trimmed and upper-cased, the empty slot skipped");
  assert.deepStrictEqual(mapSheet(ws).doors, { 1: 72, 2: 73, 3: 74, 4: 75, 5: 76 });
  pass("a parsed workbook carries one door per DOORS DONE cell, with its code and its fill");

  /* ... and none of that is status until the record says so */
  parsed.src = { Production: 6 };
  useJob(parsed, []);
  cpItems(byId("R5002")).forEach(x => assert.strictEqual(cpStatus(byId("R5002"), x.key), "",
    x.key + ": a colour in the file is not status"));
  pass("and not one of those colours is status on its own: the record decides, as for everything else");

  /* the one-time import is what turns today's colours into the record - once */
  const plan = cpImportPlan([byId("R5002")], null, 60);
  assert.deepStrictEqual(plan.map(p => p.item + " " + p.status + " " + p.done).sort(),
    ["door:1 done 1", "door:2 process 0", "door:5 done 1"],
    "the import takes the doors that say something and leaves the blank one alone");
  assert.ok(!plan.some(p => p.item === "win" || p.item === "drs"),
    "and never M or N, which are derived and have no row to import into");
  global.__imported = plan.map((p, i) => ({ id: "i" + i,
    fields: cpRowFields(p.job, p.item, p.done, p.total, p.status, "the sheet", "2026-09-11T08:00:00Z", "import") }));
  vm.runInThisContext("CP_ITEMS = __imported; cpListRebuild();");
  assert.deepStrictEqual(doorRows(cpSectionHtml(byId("R5002"), true)).map(r => r.code + " " + r.word),
    ["CD done", "DD in fabrication", "ACSS not started", "PVC done"]);
  assert.deepStrictEqual(itemState(byId("R5002"), "drs"), { done: null, total: 4, status: "process" },
    "and the Doors line reads what its doors read");
  pass("what Excel says reaches the drawer on the next load, through the record and not off the cell");

  /* a gold door through a THEME colour, not a literal FFE699 (amendment 2).
     The live sheet's gold door cells are accent 4 at tint .6, which is what
     the whole workbook's gold is - `fillOf` resolves it, and it has to resolve
     to the same word here or the doors import would miss them. */
  const tw = new ExcelJS.Workbook(), tws = tw.addWorksheet("Production");
  const tput = (r, c, v) => { tws.getCell(r, c).value = v; };
  tput(2, 4, "DATES ON CONTRACT"); tput(3, 4, "SOLD"); tput(2, 5, "CUSTOMER");
  tput(2, 13, "QUANTITY"); tput(3, 13, "WND"); tput(3, 14, "DRS");
  tput(2, 72, "DOORS DONE");
  for (let i = 0; i < 5; i++) tput(3, 72 + i, String(i + 1));
  tput(6, 3, "R5033"); tput(6, 13, 2); tput(6, 14, 3);
  ["CD", "CD", "SS"].forEach((code, i) => { tws.getCell(6, 72 + i).value = code; });
  /* slots 1 and 2 gold through the theme, slot 3 yellow as a literal */
  [0, 1].forEach(i => { tws.getCell(6, 72 + i).fill =
    { type: "pattern", pattern: "solid", fgColor: { theme: 7, tint: 0.5999938962981048 } }; });
  tws.getCell(6, 74).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  const themed = parseWorkbook(tw).find(x => x.id === "R5033");
  assert.deepStrictEqual(themed.doors.map(d => d.code + " " + d.status),
    ["CD done", "CD done", "SS process"],
    "a theme-gold door reads as done, exactly as every other theme-gold cell does");
  assert.deepStrictEqual(cpImportPlan([themed], null, 60, "door").map(p => p.item + " " + p.status),
    ["door:1 done", "door:2 done", "door:3 process"],
    "so the doors import finds them rather than leaving them to the safeguard");
  pass("a door coloured through the theme's accent tint resolves like any other gold cell");

  /* ---- amendment 5: the column fallback stops at the next labelled group -- */
  const fb = new ExcelJS.Workbook(), fws = fb.addWorksheet("Production");
  const fput = (r, c, v) => { fws.getCell(r, c).value = v; };
  fput(2, 4, "DATES ON CONTRACT"); fput(3, 4, "SOLD");
  fput(2, 13, "QUANTITY"); fput(3, 13, "WND"); fput(3, 14, "DRS");
  fput(2, 72, "DOORS DONE");                    // no sub-headers at all under it
  fput(2, 75, "WINDOWS FABRICATED"); fput(3, 75, "1");
  fput(6, 3, "R5044");
  assert.deepStrictEqual(mapSheet(fws).doors, { 1: 72, 2: 73, 3: 74 },
    "three columns before the next labelled group is three door slots, not five");
  /* ... and where the group has room, it takes five and no more */
  const fb2 = new ExcelJS.Workbook(), fws2 = fb2.addWorksheet("Production");
  const f2 = (r, c, v) => { fws2.getCell(r, c).value = v; };
  f2(2, 4, "DATES ON CONTRACT"); f2(3, 4, "SOLD");
  f2(2, 13, "QUANTITY"); f2(3, 13, "WND"); f2(3, 14, "DRS");
  f2(2, 72, "DOORS DONE");
  f2(2, 80, "WINDOWS FABRICATED"); f2(3, 80, "1");
  f2(6, 3, "R5044");
  assert.deepStrictEqual(mapSheet(fws2).doors, { 1: 72, 2: 73, 3: 74, 4: 75, 5: 76 },
    "five slots, and the three unlabelled columns after them are left alone");
  pass("the door-column fallback is bounded by the group it is in, and by five");

  /* ---- amendment 9: a number or a date in that cell is not a door -------- */
  ["4", "12", "0", "3.5", "-2", "12/05/2026", "2026-09-14", "07/04"]
    .forEach(t => assert.strictEqual(doorCode(t), "", JSON.stringify(t) + " is not a door type"));
  ["CD", "1 DOOR", "SFCD", "ACSS", "PVC", "2 CD", "DD-3"]
    .forEach(t => assert.ok(doorCode(t), JSON.stringify(t) + " is"));
  pass("a cell holding only a number or a date is a quantity or a date, not a door");

  /* ---- amendment 3: Windows and Doors are never a Changes line of their own */
  useJob(mkJob(), []);
  const was = JSON.parse(JSON.stringify(byId("R5001")));
  const now = JSON.parse(JSON.stringify(byId("R5001")));
  now.cp.win = "process"; now.cp.drs = "done";         // the aggregate paint landed in the file
  now.doors[0].status = "done";                        // and the door that caused it
  assert.deepStrictEqual(diffJobs([was], [now], "someone", "2026-09-14T10:00:00Z")
    .map(x => x.what), ["Door 1"],
    "one line, for the door - not three, with the copy of a copy reported twice over");
  pass("a Windows or Doors cell moving in the file is never news of its own");

  /* ---- amendment 7: the badge counts the work, not the line above it ----- */
  useJob(mkJob(), [
    { item: "door:1", done: 0, total: 1, status: "process" },
    { item: "prod:polaris 85mm casement:f", done: 2, total: 5, status: "process" }
  ]);
  assert.deepStrictEqual(["win", "drs"].map(k => cpStatus(byId("R5001"), k)), ["process", "process"],
    "both lines read in fabrication, from what is under them");
  assert.strictEqual(cpInProgress(byId("R5001")), 2,
    "and the row's badge counts the two real pieces of work, not those two as well");
  pass("the in-progress badge counts the children only, never a line and its children");

  /* ================ 6. the standing proofs ============================= */
  assert.strictEqual(FETCHED, 0, "the whole run made no network call at all");
  const files = ["parser.js", "checkpoints.js", "app.js", "export.js", "test_doors.js"]
    .map(f => fs.readFileSync(__dirname + "/" + f, "utf8")).join("\n");
  assert.strictEqual(/[A-Za-z0-9._%-]+@(?!example\.test)[A-Za-z0-9.-]+\.[A-Za-z]{2,}/.test(files), false,
    "no real address anywhere in the files this feature touches");
  /* the tablet is not involved in any of this: the doors are the office's, and
     the floor's files carry no door key, no DOORS DONE column and no j.doors */
  const station = fs.readFileSync(__dirname + "/station.js", "utf8") +
                  fs.readFileSync(__dirname + "/station-core.js", "utf8");
  ["door:", "DOORS DONE", ".doors", "setDoorStatus", "cpDoorAt", "cpDerived"].forEach(bit =>
    assert.strictEqual(station.indexOf(bit), -1, "the floor's files must not mention " + bit));
  pass("no network, nothing personal, and the tablet is not in this at all");

  console.log("\n" + n + " checks passed");
  process.exit(0);
})().catch(e => { console.error("FAIL", e); process.exit(1); });
