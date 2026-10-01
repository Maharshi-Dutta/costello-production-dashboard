/* Offline test of the fabrication station, Part A
   (docs/specs/2026-09-25-fabrication-station.md).

   Two halves. The first is the pure core (fabrication-core.js over
   station-core.js): the allow-list, the slice off `Production` alone, the door
   labels, F/S/T with T, Active/OnSheet, the tabs, the eligibility gate, the
   clamp, the write bodies, and the colour painter's rule (fbCellWant).

   The second loads the office's own code (index.html's chain) with a stubbed
   workbook channel and runs the painter end to end: what it plans, what it
   sends, and - the two "verify and report" checks of the brief - that the
   checkpoint writer never whitens a lavender cell whose record is blank and
   that the hand-paint adopter never takes a lavender/purple cell for a hand
   change. No network; every person is made up.
   Run: node test_fabrication.js                                             */
const fs = require("fs"), vm = require("vm"), assert = require("assert");

/* ---------- browser shims (the minimum app.js needs to load) ---------- */
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null),
                        setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" }, innerWidth: 1280, innerHeight: 800,
                  addEventListener() {}, removeEventListener() {} };
global.performance = { now: () => Date.now() };
function stubEl(tag) {
  let html = "";
  const e = { tagName: String(tag || "div").toUpperCase(), style: { setProperty() {} }, dataset: {}, kids: [],
    textContent: "", value: "", hidden: false, className: "",
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    appendChild(c) { e.kids.push(c); return c; }, remove() {}, contains() { return false; },
    setAttribute() {}, getAttribute() { return null; }, addEventListener() {}, removeEventListener() {},
    focus() {}, blur() {}, getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0 }),
    querySelector: () => null, querySelectorAll: () => [] };
  Object.defineProperty(e, "innerHTML", { get: () => html, set: v => { html = String(v); } });
  return e;
}
/* any element the page asks for exists, except the ones whose absence means
   "that window is not open" */
const EL = {}, NULLABLE = ["#fabhost", "#fabbtn", "#dhost", "#xhost", "#ahost", "#chost", "#vhost", "#lhost", "#dayhost"];
global.document = { documentElement: stubEl(), body: stubEl(), head: stubEl(), activeElement: null,
  title: "x", createElement: t => stubEl(t),
  querySelector: s => (NULLABLE.indexOf(s) >= 0 ? null : (EL[s] = EL[s] || stubEl())), querySelectorAll: () => [],
  addEventListener() {}, removeEventListener() {} };
/* nothing in this run may reach the network: every call is a failure */
const NET = [];
global.fetch = async (url, init) => { NET.push(String(url)); return { ok: false, status: 599, text: async () => "{}" }; };

const src = f => fs.readFileSync(__dirname + "/" + f, "utf8");
const run = f => vm.runInThisContext(src(f), { filename: f });
run("parser.js");
run("graph.js");
global.CW = window.CW;
run("checkpoints.js");
run("station-core.js");
global.ST = window.ST;
run("welding-core.js");
run("glazing-core.js");
run("fabrication-core.js");
const F = window.FABC;
run("export.js");
run("app.js");
const A = code => vm.runInThisContext(code);

let n = 0;
const pass = m => { n++; console.log("  ok  " + m); };

(async () => {
  /* ================= 1. the allow-list, through the key normaliser ========= */
  ["casement windows", "polaris 85 tilt turn", "aluclad tilt turn", "4000 tilt turn", "pvc french wds",
   "sidelights", "pvc door", "super door", "bifold", "pvc smart", "alu clad windows"]
    .forEach(h => assert.ok(F.fbAllowed(h), h + " is on the owner's list"));
  /* the parser's headers the owner's names do not match - reported, not guessed */
  /* owner 2026-09-25: the sheet's own headers are fed */
  ["arch angles", "th w", "composite"].forEach(h => assert.ok(F.fbAllowed(h), h + " is fed"));
  assert.ok(F.fbIsDoorGroup("composite"), "composite is a door group (door labels)");
  assert.ok(!F.fbAllowed("composite pvc door") && !F.fbAllowed("arch angles th w"), "the owner's merged names are not headers");
  assert.ok(F.fbAllowed("ALUCLAD TILT & TURN") && F.fbAllowed("  Casement   Windows "), "the & and spacing do not matter");
  assert.ok(F.fbIsDoorGroup("pvc door") && !F.fbIsDoorGroup("casement windows"));
  pass("allow-list matched through fbKey; unmatched headers are left out");

  /* ================= 2. the slice: Production alone, F/S/T, doors ========== */
  const names = ["Ready to fit", "In production"];
  const jobs = [
    { id: "R9001", cust: "Person A 0861234567", cat: "prod", blk: 1, seq: 2,
      notes: [{ k: "comment", t: "ring 087 1234567 first" }],
      doors: [{ slot: 1, code: "CD" }, { slot: 2, code: "CD" }, { slot: 3, code: "DD" }],
      prods: [{ n: "casement windows", f: 9, s: 9, t: 9 }, { n: "bifold", f: 1, s: 4, t: 0 }],
      prodsMain: [{ n: "casement windows", f: 4, s: 6, t: 2 }, { n: "pvc door", f: 3, s: 4, t: 0 },
                  { n: "composite", f: 0, s: 2, t: 0 }, { n: "sidelights", f: 0, s: 0, t: 0 }] },
    { id: "R9002", cust: "Person B", cat: "prod", blk: 0, seq: 1, doors: [],
      prodsMain: [{ n: "7000 casement", f: 0, s: 0, t: 3 }] },
    { id: "R9003", cust: "Gone", cat: "past", blk: -1, prodsMain: [{ n: "casement windows", f: 1, s: 1, t: 0 }] }
  ];
  const slice = F.fbSlice(jobs, names);
  const t = slice.map(r => r.title);
  assert.deepStrictEqual(t, ["R9002|7000 CASEMENT", "R9001|CASEMENT WINDOWS", "R9001|PVC DOOR", "R9001|COMPOSITE"],
    "office order; composite fed (owner 2026-09-25); a zero group not fed; the past job not fed; bifold (prods only) not fed");
  assert.strictEqual(slice[3].sashes, 2, "composite: the CD sash count");
  assert.strictEqual(slice[3].doors, "2 CD, 1 DD", "composite is a door group: door labels");
  const cw = slice[1], pd = slice[2], t7 = slice[0];
  assert.deepStrictEqual([cw.frames, cw.sashes, cw.transoms], [4, 6, 2], "counts from prodsMain, never prods");
  assert.strictEqual(t7.transoms, 3, "a group with only T is fed");
  assert.strictEqual(pd.doors, "2 CD, 1 DD", "door codes as labels, on a door group");
  assert.strictEqual(cw.doors, "", "and not on a window group");
  assert.ok(cw.active && !t7.active && t7.onSheet, "Active = In production; OnSheet = on the sheet");
  assert.ok(!/086|087|1234567/.test(cw.customer + cw.comment), "rule 3: no phone number reaches the list");
  const ff = F.fbFeederFields(cw);
  assert.deepStrictEqual(Object.keys(ff).sort(), F.FB_FEEDER_FIELDS.slice().sort());
  pass("the slice reads Production only, carries T, labels doors, and strips contact numbers");

  /* ================= 3. the feed plan: no floor column, no delete ========= */
  const have = [{ id: "5", fields: Object.assign({ Title: cw.title }, ff, { Frames: 1, FramesDone: 3, DoneAt: "2026-09-25T10:00:00Z" }) },
                { id: "6", fields: { Title: "R8000|CASEMENT WINDOWS", Active: "Yes", OnSheet: "Yes", FramesDone: 2 } }];
  const plan = ST.feedPlan(slice, have, { at: "T", by: "the office", def: F.FAB });
  const named = plan.adds.concat(plan.patches.map(p => p.fields));
  named.forEach(f => Object.keys(f).forEach(k =>
    assert.ok(F.FB_FEEDER_WRITES.indexOf(k) >= 0, "the feeder wrote " + k)));
  const gone = plan.patches.find(p => p.id === "6");
  assert.deepStrictEqual([gone.fields.Active, gone.fields.OnSheet], ["No", "No"], "a row off the sheet: both flags off");
  assert.strictEqual(plan.adds.length, 3);
  pass("the feeder writes job facts only, and marks a gone row inactive rather than deleting it");

  /* ================= 4. cards, colours, tabs ================= */
  const items = [
    { id: "1", fields: Object.assign({ Title: cw.title }, F.fbFeederFields(cw), { FramesDone: 4, SashesDone: 99, TransomsDone: 1 }) },
    { id: "2", fields: Object.assign({ Title: pd.title }, F.fbFeederFields(pd), { FramesDone: 3, SashesDone: 4 }) },
    { id: "3", fields: Object.assign({ Title: t7.title }, F.fbFeederFields(t7)) },
    { id: "4", fields: { Title: "R7000|PVC DOOR", Job: "R7000", Group: "PVC DOOR", Frames: 1, Active: "No", OnSheet: "No" } }
  ];
  const board = F.fbOfficeBoard(items);
  assert.deepStrictEqual(board.map(c => c.job), ["R9002", "R9001"], "off-sheet rows are on no board");
  const c1 = board[1];
  assert.strictEqual(c1.groups[0].sashes, 6, "display clamp: 99 of 6 reads 6");
  assert.strictEqual(c1.groups[0].lines.length, 3, "three lines where T > 0");
  assert.strictEqual(c1.groups[0].colour, "lavender", "started");
  assert.strictEqual(c1.groups[1].colour, "purple", "a group all done");
  assert.strictEqual(c1.colour, "lavender");
  assert.strictEqual(F.fbColour(0, 5), "");
  const tabs = F.fbTabs(items);
  assert.deepStrictEqual([tabs.floor.map(c => c.job), tabs.finished.map(c => c.job)], [["R9001"], ["R9002"]],
    "On floor = In production not finished; Finished = every other on-sheet card");
  items[0].fields.SashesDone = 6; items[0].fields.TransomsDone = 2;
  assert.deepStrictEqual(F.fbTabs(items).finished.map(c => c.job), ["R9002", "R9001"], "a finished job moves to Finished");
  assert.deepStrictEqual(F.fbSearchTab(F.fbTabs(items), "9002", "floor"), { tab: "finished", other: 0 });
  assert.strictEqual(F.fbJobCard(items, "r9001").groups.length, 2);
  pass("cards, lavender/purple roll-up, the two tabs and the search switch");

  /* ================= 5. eligibility ================= */
  const people = F.fbPeople([
    { id: "1", fields: { Title: "Person A", Station: "Fabrication", Active: "Yes", Stages: "Casement  windows; pvc door, NOT A GROUP" } },
    { id: "2", fields: { Title: "Person B", Station: "Welding", Active: "Yes", Stages: "casement windows" } },
    { id: "3", fields: { Title: "Person C", Station: "Fabrication", Active: "Yes", Stages: "" } }]);
  assert.deepStrictEqual(people.map(p => p.name), ["Person A", "Person C"], "this station's people only");
  const pa = people[0];
  assert.ok(F.fbEligible(pa, "CASEMENT WINDOWS") && F.fbEligible(pa, "pvc door"));
  assert.ok(!F.fbEligible(pa, "SIDELIGHTS") && !F.fbEligible(people[1], "CASEMENT WINDOWS") && !F.fbEligible(null, "PVC DOOR"));
  const tsrc = src("fabrication.js");
  assert.ok(/!mayTap\(rec, part\) \? "assign"[\s\S]{0,200}if \(why\) \{[\s\S]{0,200}return;[\s\S]{0,300}queueTap/.test(tsrc),
    "tap() refuses a line the gate refuses before queueing (section 13 runs it)");
  assert.ok(/const mayTap = \(rec, part\) => !rec\.sheetDone && F\.fbCanTap\(/.test(tsrc), "and the gate is the core's fbCanTap");
  pass("eligibility: a person moves only the groups named in Stages; the tablet's tap() enforces it");

  /* per group AND part (owner, 2026-09-28) */
  const P = t => F.fbParseStages(t);
  const all3 = { frames: true, sashes: true, transoms: true };
  assert.deepStrictEqual(P("PVC DOOR"), { "pvc door": all3 }, "a group alone is every part (today's meaning)");
  assert.deepStrictEqual(P("PVC SMART:sashes"), { "pvc smart": { sashes: true } });
  assert.deepStrictEqual(P(" pvc door : Frames + TRANSOMS "), { "pvc door": { frames: true, transoms: true } },
    "case and spaces do not matter; + joins parts");
  assert.deepStrictEqual(P("PVC DOOR:frames, pvc door:sashes"), { "pvc door": { frames: true, sashes: true } },
    "the same group twice is the union");
  const allS = P("ALL:sashes");
  assert.strictEqual(Object.keys(allS).length, F.FB_GROUP_KEYS.length, "ALL is every fed group");
  assert.ok(Object.keys(allS).every(k => JSON.stringify(allS[k]) === '{"sashes":true}'));
  assert.strictEqual(Object.keys(P("ALL")).length, F.FB_GROUP_KEYS.length);
  assert.deepStrictEqual(P("NOT A GROUP, PVC DOOR:glass+frames, SIDELIGHTS:nonsense,,:frames"),
    { "pvc door": { frames: true } }, "unknown groups and part words are ignored; an entry left with no part grants nothing");
  assert.deepStrictEqual(P("arch & angles:frames"), { "arch angles": { frames: true } }, "the & and the colon both survive keying");
  /* N2: a comma after a part continues the group, as a hand types it */
  assert.deepStrictEqual(P("CASEMENT WINDOWS:frames,sashes"), { "casement windows": { frames: true, sashes: true } });
  assert.deepStrictEqual(P("CASEMENT WINDOWS:frames; sashes + transoms, PVC DOOR:sashes"),
    { "casement windows": all3, "pvc door": { sashes: true } }, "; too, and a new GROUP: starts a new group");
  assert.deepStrictEqual(P("frames, PVC SMART"), { "pvc smart": all3 }, "part words with no group before them are ignored");
  const pp = F.fbPeople([{ id: "7", fields: { Title: "Person D", Station: "Fabrication", Active: "Yes",
                                             Stages: "PVC SMART:sashes, PVC DOOR:frames+transoms" } }])[0];
  assert.deepStrictEqual(pp.stages.sort(), ["pvc door", "pvc smart"]);
  assert.ok(F.fbEligible(pp, "PVC SMART", "sashes") && !F.fbEligible(pp, "PVC SMART", "frames"), "per part");
  assert.ok(F.fbEligible(pp, "PVC SMART") && !F.fbEligible(pp, "SIDELIGHTS"), "group level: any part of it");
  assert.ok(F.fbEligible(pp, "pvc door", "Transoms") && !F.fbEligible(pp, "pvc door", "sashes"));
  assert.ok(F.fbCanTap(pp, "PVC DOOR", "frames", {}, false, "X") && !F.fbCanTap(pp, "PVC DOOR", "sashes", {}, false, "X"),
    "the Part A fallback is per part too");
  const hold = F.fbAssignIndex(F.fbAssignRows([{ id: "1", fields: { Job: "X", Group: "PVC DOOR", Part: "sashes",
    Person: "Person D", Qty: 2, Status: "Assigned" } }]));
  assert.ok(!F.fbCanTap(pp, "PVC DOOR", "sashes", hold, true, "X"), "holding an assignment on a part they do not do: still locked");
  assert.ok(/words: "not your part"/.test(tsrc) && /Not your part/.test(tsrc), "the tablet says 'not your part'");
  pass("Stages per group and part: plain, :parts, +, ALL, duplicates, junk; the gate per part");

  /* ================= 6. clamp, write bodies, rebase ================= */
  const g0 = F.fbRecord(items[0]);
  assert.strictEqual(F.fbApplyTap(g0, "transoms", 5), 2, "+5 clamps to the total");
  assert.strictEqual(F.fbApplyTap(g0, "frames", -9), 0);
  assert.strictEqual(F.fbApplyTap(g0, "sashes", "none"), 0);
  assert.strictEqual(F.fbApplyTap(g0, "frames", "all"), 4);
  assert.strictEqual(F.fbApplyTap(g0, "glass", 1), null, "not a part");
  const body = F.fbFloorOnly(Object.assign(F.fbTapFields("transoms", 2, "Person A", "2026-09-25T11:00:00Z"),
    { Job: "X", Active: "No", Urgent: "group", Frames: 99 }));
  assert.deepStrictEqual(Object.keys(body).sort(), ["DoneAt", "DoneBy", "TransomsAt", "TransomsBy", "TransomsDone"]);
  assert.strictEqual(F.fbOfficeFields, F.fbTapFields, "the office writes the same body");
  const e = { part: "frames", from: 1, value: 2, at: "2026-09-25T10:00:00Z" };
  assert.deepStrictEqual(F.fbRebase(e, { FramesDone: 3, Frames: 4 }), { action: "rebase", value: 4, from: 3 });
  assert.deepStrictEqual(F.fbRebase(e, { FramesDone: 0, DoneAt: "2026-09-25T10:05:00Z" }), { action: "drop" });
  assert.deepStrictEqual(F.fbLogEntry({ job: "r9001", group: "pvc door", part: "Frames", from: 1, to: 2 }).type, "PVC DOOR");
  pass("clamp, the five-field write body, Urgent unwritable, rebase and the log line");

  /* ================= 7. the painter's rule ================= */
  assert.strictEqual(F.fbCellWord("FFD9D2E9"), "process");
  assert.strictEqual(F.fbCellWord("#b4a7d6"), "done");
  assert.strictEqual(F.fbCellWord(""), "");
  assert.strictEqual(F.fbCellWord("FFFFFF"), "");
  ["FFFF00", "FFE699", "00B050", "FF0000"].forEach(h => assert.strictEqual(F.fbCellWord(h), "other"));
  assert.strictEqual(F.fbCellWant(2, 4, "", ""), "process", "blank record: paints lavender");
  assert.strictEqual(F.fbCellWant(4, 4, "", "process"), "done", "blank record: purple when all done");
  ["process", "done", "cut"].forEach(r => assert.strictEqual(F.fbCellWant(4, 4, r, ""), null, r + " record: never"));
  assert.strictEqual(F.fbCellWant(0, 4, "", "done"), "", "own colour, done 0: white");
  assert.strictEqual(F.fbCellWant(0, 4, "", "process"), "");
  assert.strictEqual(F.fbCellWant(0, 4, "", "other"), null, "foreign colour, done 0: nothing");
  assert.strictEqual(F.fbCellWant(3, 4, "", "other"), null, "foreign colour: never painted over");
  assert.strictEqual(F.fbCellWant(0, 4, "", ""), null, "white and nothing done: nothing");
  assert.strictEqual(F.fbCellWant(2, 4, "", "process"), null, "already right: nothing");
  assert.ok(A("cpWordForHex('" + F.FB_PROCESS_HEX + "') === null && cpWordForHex('" + F.FB_DONE_HEX + "') === null"),
    "neither colour is a checkpoint colour");
  pass("the painter's rule: never over the office, never over a foreign colour, white only over its own");

  /* ================= 8. the painter end to end, in the office's code ========= */
  /* The workbook, twice over: FILLS is the DOWNLOAD (what the plan sees), LIVE
     is the cell as the Excel API reads it right now (what the write re-checks),
     HDRV the header cells. A write lands in LIVE. */
  const WRITES = [], LOGS = [], LIVE = {}, COLS = { P: 16, Q: 17, R: 18 };
  const HDRV = { P2: "CASEMENT WINDOWS", P3: "F", Q3: "S", R3: "T" };
  let GETFAIL = false, HOOK = null;
  CW.serialised = async (s, fn) => fn();
  CW.rowForJob = async () => { if (HOOK) HOOK(); return 7; };
  CW.findFile = async () => ({ base: "/x/workbook", siteId: "s" });
  CW.batchGet = async urls => {
    if (GETFAIL) throw new Error("batch refused");
    return urls.map(u => {
      const a = /address='([A-Z]+)(\d+)'/.exec(u);
      if (/format\/fill/.test(u)) return { color: "#" + (LIVE[COLS[a[1]]] || "FFFFFF") };
      return { values: [[HDRV[a[1] + a[2]] || ""]] };
    });
  };
  CW.batchWrite = async reqs => reqs.forEach(r => {
    WRITES.push(r);
    LIVE[COLS[/address='([A-Z]+)/.exec(r.url)[1]]] = r.body.color.slice(1);
  });
  global.__LOGS = LOGS;
  A("noteChange = function (job, what, from, to) { __LOGS.push({ job: job, what: what, from: from, to: to }); };" +
    "scheduleReconcile = function () {};");
  const FILLS = {};                                  // column -> hex in the downloaded sheet
  global.__WB = { getWorksheet: () => ({ getRow: () => ({ getCell: c =>
    ({ fill: FILLS[c] ? { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + FILLS[c] } } : null }) }) }) };
  const job = { id: "R9001", src: { Production: 7 }, done: 0, cat: "prod", blk: 1,
    prods: [{ n: "casement windows", f: 4, s: 6, t: 2 }], prodsMain: [{ n: "casement windows", f: 4, s: 6, t: 2 }],
    cp: { win: "", drs: "", glass: {}, prod: {} }, doors: [], glass: {} };
  global.__JOB = job;
  const row = { id: "1", fields: { Title: "R9001|CASEMENT WINDOWS", Job: "R9001", Group: "CASEMENT WINDOWS",
    Frames: 4, Sashes: 6, Transoms: 2, FramesDone: 2, SashesDone: 6, TransomsDone: 0,
    Active: "Yes", OnSheet: "Yes", DoneAt: "2026-09-25T10:00:00Z", DoneBy: "Person A" } };
  global.__ROWS = [row];
  A("CP_LIST_OK = true; CP_IMPORTED = 'x'; cpSetImportPending(false); cpRowsSet({});" +
    "PRODMAP = { prod: { 'casement windows': { f: 16, s: 17, t: 18 } }, glass: {}, hdr: [2, 3] };" +
    "LASTWB = __WB; ALL = [__JOB]; FABR_OK = true; FABR_ITEMS = __ROWS.slice(); FABR_PAINTED = {};");
  const runP = () => A("fabrColourRun()");
  assert.strictEqual(await runP(), 1, "one job painted");
  assert.deepStrictEqual(WRITES.map(w => [w.url.replace(/^.*address='/, "").replace(/'.*$/, ""), w.body.color]),
    [["P7", "#D9D2E9"], ["Q7", "#B4A7D6"]], "F lavender, S purple, T (done 0, white) untouched");
  assert.ok(WRITES.every(w => w.method === "PATCH" && /worksheets\('Production'\)\/range\(address='[A-Z]+7'\)\/format\/fill$/.test(w.url)),
    "fills only, on the job's own row");
  assert.deepStrictEqual(LOGS.map(l => l.what), ["Fabrication colours"], "one Dashboard Log line per paint");
  assert.strictEqual(await runP(), 0, "nothing to do again: the paint is remembered while the download lags");

  const fresh = (fields) => {                        // download white, live white, nothing remembered
    WRITES.length = 0; LOGS.length = 0; HOOK = null; GETFAIL = false;
    Object.keys(LIVE).forEach(k => delete LIVE[k]); Object.keys(FILLS).forEach(k => delete FILLS[k]);
    global.__F = fields || {};
    A("FABR_PAINTED = {}; cpRowsSet({}); Object.keys(FABR_FAIL).forEach(k => delete FABR_FAIL[k]);" +
      "FABR_ITEMS = [{ id: '1', fields: Object.assign({}, __ROWS[0].fields, __F) }];");
  };
  const cells = () => WRITES.map(w => /address='([A-Z]+\d+)/.exec(w.url)[1]);

  /* B1: the office's record turns yellow on F between the plan and the write */
  fresh();
  HOOK = () => A("cpRowPut('R9001', 'prod:casement windows:f', { status: 'process', done: 1, total: 4 });");
  assert.strictEqual(await runP(), 1);
  assert.deepStrictEqual(cells(), ["Q7"], "B1: F is re-decided against the record as it is NOW and left");
  /* B1: a colour painted in Excel after the download (foreign), and a cell
     already right live: no PATCH at all, and so no log line */
  fresh();
  LIVE[16] = "FFFF00"; LIVE[17] = "B4A7D6";
  await runP();
  assert.deepStrictEqual([WRITES.length, LOGS.length], [0, 0],
    "B1: live yellow is never painted over; live already purple is not rewritten or logged");
  /* B1: the live read fails: nothing is written for that job this pass */
  fresh();
  GETFAIL = true;
  await runP();
  assert.deepStrictEqual([WRITES.length, LOGS.length], [0, 0], "B1: a failed live read paints nothing");
  /* M3: a header that no longer names this sub-column is left alone */
  fresh();
  HDRV.P3 = "S";
  await runP();
  assert.deepStrictEqual(cells(), ["Q7"], "M3: the F cell's header says S now: not painted");
  HDRV.P3 = "F";
  /* M4: the sheet's own count gates the paint, not only the list's Total */
  fresh();
  job.prodsMain = [{ n: "casement windows", f: 0, s: 6, t: 2 }];
  await runP();
  assert.deepStrictEqual(cells(), ["Q7"], "M4: F is 0 on the sheet: not painted though the list says 4");
  job.prodsMain = [{ n: "casement windows", f: 4, s: 6, t: 2 }];
  /* the floor takes sashes back to 0: its own purple goes white */
  fresh({ SashesDone: 0, FramesDone: 0 });
  LIVE[17] = "B4A7D6"; FILLS[17] = "B4A7D6";
  assert.strictEqual(await runP(), 1);
  assert.deepStrictEqual(WRITES.map(w => w.body.color), ["#FFFFFF"], "own colour, done 0: white");
  /* an untouched row paints nothing; a gold row is whole */
  fresh({ DoneAt: "" });
  assert.strictEqual(await runP(), 0, "DoneAt empty: nobody touched it, nothing painted");
  A("FABR_ITEMS = __ROWS.slice();");
  job.done = 1;
  assert.strictEqual(await runP(), 0, "a gold row is left whole");
  job.done = 0;
  pass("the painter end to end: one $batch of fills, one log line, never over the office, white only over its own");

  /* ================= 9. verify-and-report (a) and (b) ================= */
  /* (a) the checkpoint writer (cpRepaintPlan) never whitens a lavender cell
     whose record is blank - no row at all, or a row saying "" that this
     browser has never painted (the adopter, run before it in load(), records
     PAINTED from the file first, and lavender parses as "") */
  FILLS[16] = "D9D2E9";
  /* the parser reads both colours as no checkpoint colour at all - the same as
     white - so j.cp for a lavender cell is "" (which is what __JOB carries) */
  assert.ok(A("cpOf('D9D2E9') === '' && cpOf('B4A7D6') === ''"), "the parser reads lavender and purple as blank");
  A("PAINTED = {}; cpRowsSet({});");
  assert.deepStrictEqual(A("cpAdoptCandidates([__JOB]).length"), 0, "(b) no row: not a hand change");
  assert.deepStrictEqual(A("cpRepaintPlan([__JOB]).length"), 0, "(a) no row: never repainted");
  A("cpRowPut('R9001', 'prod:casement windows:f', { status: '', done: 0, total: 4 });");
  assert.deepStrictEqual(A("cpAdoptCandidates([__JOB]).length"), 0, "(b) a blank row: not a hand change");
  assert.deepStrictEqual(A("cpRepaintPlan([__JOB]).length"), 0,
    "(a) a blank row, after the adopter has run: not repainted white");
  pass("verify-and-report: (a) and (b) hold on the office's own code");

  /* M1: a blank record row, and this browser has never painted the cell (the
     adopter has not run): the repaint still leaves lavender / purple alone... */
  A("PAINTED = {}; cpRowsSet({}); cpRowPut('R9001', 'prod:casement windows:f', { status: '', done: 0, total: 4 });");
  FILLS[16] = "D9D2E9";
  assert.strictEqual(A("cpRepaintPlan([__JOB]).length"), 0, "M1: lavender over a blank record is not whitened");
  FILLS[16] = "B4A7D6";
  assert.strictEqual(A("cpRepaintPlan([__JOB]).length"), 0, "M1: nor purple");
  /* ... and the guard is that narrow: any other colour is still put right */
  FILLS[16] = "FFFF00";
  assert.strictEqual(A("cpRepaintPlan([__JOB]).length"), 1, "M1: a yellow cell over a blank record is still repainted");
  pass("M1: the checkpoint repaint never whitens a fabrication colour over a blank record");

  /* ================= 11. Part B, the core ================= */
  const ai = (id, f) => ({ id: String(id), fields: Object.assign({ Job: "R9001", Group: "CASEMENT WINDOWS" }, f) });
  let arows = F.fbAssignRows([
    ai(1, { Part: "frames", Person: "Person A", Qty: 3, Status: "Assigned" }),
    ai(2, { Part: "frames", Person: "Person B", Qty: 1, Status: "Requested" }),
    ai(3, { Part: "sashes", Person: "Person A", Qty: 2, Status: "Refused" }),
    ai(4, { Part: "sashes", Person: "Person A", Qty: 6, Status: "Removed" }),
    ai(5, { Part: "", Person: "Person A", Qty: 1, Status: "Assigned" })]);
  let idx = F.fbAssignIndex(arows);
  assert.strictEqual(arows.length, 4, "a row with no part is not an assignment");
  const L = F.fbLineOf(idx, "r9001", "casement windows", "frames");
  assert.deepStrictEqual([L.assigned.length, L.requested.length], [1, 1]);
  assert.deepStrictEqual(F.fbLineOf(idx, "R9001", "CASEMENT WINDOWS", "sashes"), { assigned: [], requested: [] },
    "refused and removed rows are history");
  const over = F.fbSplitCheck(idx, "R9001", "CASEMENT WINDOWS", "frames", 4, 2);
  assert.ok(!over.ok && over.free === 1 && /Only 1 of 4/.test(over.msg), "split: a sum over the total is refused, said plainly");
  assert.ok(F.fbSplitCheck(idx, "R9001", "CASEMENT WINDOWS", "frames", 4, 1).ok, "exactly the total is fine");
  assert.ok(F.fbSplitCheck(idx, "R9001", "CASEMENT WINDOWS", "frames", 4, 1, "1").ok &&
            !F.fbSplitCheck(idx, "R9001", "CASEMENT WINDOWS", "frames", 4, 5, "1").ok, "an approval excludes its own row");
  assert.ok(!F.fbSplitCheck(idx, "R9001", "CASEMENT WINDOWS", "frames", 4, 0).ok, "0 is not an assignment");
  const pB = { name: "Person B", stages: ["casement windows"] };
  assert.ok(F.fbCanTap(pa, "CASEMENT WINDOWS", "frames", idx, true, "R9001"), "assigned: may tap");
  assert.ok(!F.fbCanTap(pB, "CASEMENT WINDOWS", "frames", idx, true, "R9001"), "requested: locked until approved");
  assert.ok(!F.fbCanTap(pa, "CASEMENT WINDOWS", "sashes", idx, true, "R9001"), "eligible but unassigned: locked");
  assert.ok(F.fbCanTap(pa, "CASEMENT WINDOWS", "sashes", {}, false, "R9001"), "no assignments list: Part A's gate");
  assert.ok(!F.fbCanTap(pa, "SIDELIGHTS", "frames", {}, false, "R9001"), "and never an ineligible group");
  /* P1: the gate fails closed while the list's state is unknown */
  assert.ok(!F.fbCanTap(pa, "CASEMENT WINDOWS", "frames", idx, null, "R9001"), "unknown (never read): locked, even holding it");
  assert.ok(!F.fbCanTap(pa, "CASEMENT WINDOWS", "sashes", {}, undefined, "R9001"), "and undefined is unknown too");
  /* P2: All / None only for somebody holding the whole line */
  assert.ok(!F.fbActAllowed(pa, idx, true, "R9001", "CASEMENT WINDOWS", "frames", 10, "all"), "3 of 10: no All");
  assert.ok(!F.fbActAllowed(pa, idx, true, "R9001", "CASEMENT WINDOWS", "frames", 10, "none"), "and no None");
  assert.ok(F.fbActAllowed(pa, idx, true, "R9001", "CASEMENT WINDOWS", "frames", 10, 1), "+ is fine");
  assert.ok(F.fbActAllowed(pa, idx, true, "R9001", "CASEMENT WINDOWS", "frames", 3, "all"), "3 of 3: All is fine");
  assert.ok(F.fbActAllowed(pa, {}, false, "R9001", "CASEMENT WINDOWS", "frames", 10, "all"), "no list: Part A, All allowed");
  assert.deepStrictEqual(F.fbUnapproveFields(4, "o", "T"), { Status: "Requested", Qty: 4, DecidedBy: "o", DecidedAt: "T" });
  assert.ok(F.fbRequested(idx, pB, "R9001", "CASEMENT WINDOWS", "frames"));
  assert.strictEqual(F.fbMine(idx, pa, "R9001", "CASEMENT WINDOWS", "frames"), 3, "yours: 3");
  const rq = F.fbRequestFields("r9001", "casement windows", "Transoms", "Person A", 2, "T");
  assert.deepStrictEqual(Object.keys(rq).sort(), ["Group", "Job", "Part", "Person", "Qty", "RequestedAt", "RequestedBy", "Status", "Title"]);
  assert.ok(rq.Status === "Requested" && rq.Person === "Person A" && /^R9001\|CASEMENT WINDOWS\|transoms\|[a-z0-9]{6}$/.test(rq.Title));
  assert.strictEqual(F.fbAssignFields("R9001", "X", "frames", "Person A", 2, "the office", "T").Status, "Assigned");
  assert.deepStrictEqual(Object.keys(F.fbApproveFields(2, "o", "T")).sort(), ["DecidedAt", "DecidedBy", "Qty", "Status"]);
  assert.deepStrictEqual(F.fbDecideFields("refused", "o", "T").Status, "Refused");
  /* the tablet's only write to the assignments list is a Requested row */
  assert.strictEqual((tsrc.match(/FB_ASSIGN_LIST/g) || []).length, 2, "the tablet names the list twice: its read, its one POST");
  assert.ok(/CW\.listAdd\(F\.FB_ASSIGN_LIST, body/.test(tsrc) && /F\.fbRequestFields\(/.test(tsrc) &&
            !/listPatch\(F\.FB_ASSIGN_LIST/.test(tsrc), "a POST of fbRequestFields, and never a PATCH");
  pass("Part B core: rows, split refused over the total, locked until approved, Part A fallback, the Take body");

  /* urgent, and the feeder never touches it */
  assert.strictEqual(F.fbUrgentToggle("", "job", true), "job");
  assert.strictEqual(F.fbUrgentToggle("job,frames", "job", false), "frames");
  assert.deepStrictEqual(F.fbUrgentOf("Group; sashes, nonsense"), { group: true, sashes: true });
  const ucards = F.fbOfficeBoard([
    { id: "1", fields: { Title: "A1|X", Job: "A1", Group: "X", Frames: 1, Seq: 1, OnSheet: "Yes" } },
    { id: "2", fields: { Title: "B2|X", Job: "B2", Group: "X", Frames: 1, Seq: 2, OnSheet: "Yes", Urgent: "transoms" } }]);
  assert.deepStrictEqual(F.fbUrgentFirst(ucards).map(c => c.job), ["B2", "A1"], "urgent sorts to the top");
  const uplan = ST.feedPlan([cw], [{ id: "9", fields: Object.assign({ Title: cw.title, Urgent: "job,frames" },
    F.fbFeederFields(cw), { Customer: "changed" }) }], { at: "T", by: "o", def: F.FAB });
  assert.ok(uplan.patches.length === 1 && !("Urgent" in uplan.patches[0].fields), "the feeder never writes Urgent");
  assert.ok(F.FB_FEEDER_WRITES.indexOf("Urgent") < 0 && F.FB_FLOOR_FIELDS.indexOf("Urgent") < 0);
  pass("urgent: the toggle, urgent first, and the feeder never overwrites it");

  /* notifications: a baseline, then only what is new or changed */
  const store = {};
  let ncards = F.fbOfficeBoard([{ id: "1", fields: { Title: "R9001|CASEMENT WINDOWS", Job: "R9001",
    Group: "CASEMENT WINDOWS", Frames: 4, OnSheet: "Yes" } }]);
  let nts = F.fbNotices(arows, ncards, pa, idx);
  assert.strictEqual(nts.length, 3, "Person A: one assigned, one refused, one removed; never their own request");
  let seen = F.fbSeenFor(store, "Person A", nts);
  assert.strictEqual(F.fbUnseen(nts, seen).length, 0, "first sight is a baseline, not a storm");
  arows = F.fbAssignRows([ai(1, { Part: "frames", Person: "Person A", Qty: 2, Status: "Assigned" }),
                          ai(6, { Part: "sashes", Person: "Person A", Qty: 6, Status: "Assigned" })]);
  idx = F.fbAssignIndex(arows);
  ncards = F.fbOfficeBoard([{ id: "1", fields: { Title: "R9001|CASEMENT WINDOWS", Job: "R9001",
    Group: "CASEMENT WINDOWS", Frames: 4, Sashes: 6, OnSheet: "Yes", Urgent: "group" } }]);
  nts = F.fbNotices(arows, ncards, pa, idx);
  assert.strictEqual(F.fbUnseen(nts, F.fbSeenFor(store, "Person A", nts)).length, 3,
    "a changed Qty, a new assignment and an urgent flag on a held line are three notices");
  F.fbUnseen(nts, store["Person A"]).forEach(n => { store["Person A"][n.key] = 1; });
  assert.strictEqual(F.fbUnseen(nts, store["Person A"]).length, 0, "seen once tapped");
  assert.strictEqual(F.fbNotices(arows, ncards, pB, idx).length, 0, "nothing for somebody holding nothing");
  /* P7: keys no longer current are pruned, and a prune that removes nothing says so */
  const before = Object.keys(store["Person A"]).length;
  assert.ok(F.fbPruneSeen(store["Person A"], nts), "the three baseline keys of the old rows go");
  assert.ok(before > Object.keys(store["Person A"]).length && Object.keys(store["Person A"]).length === nts.length);
  assert.ok(!F.fbPruneSeen(store["Person A"], nts), "nothing to prune: no change, so no save");
  assert.ok(/if \(F\.fbPruneSeen\(seen, all\) \|\| fresh\) saveSeen\(\);/.test(src("fabrication.js")),
    "the tablet saves cw_fabseen only when it changed");
  const fsrcN = src("fabrication.js");
  assert.ok(/cw_fabseen/.test(fsrcN) && /AudioContext/.test(fsrcN) && !/\.mp3|\.wav|\.ogg/.test(fsrcN),
    "seen per person in cw_fabseen; a WebAudio beep, no audio file");
  assert.ok(/function beep\(\) \{[\s\S]{0,700}\} catch \(e\) \{\}/.test(fsrcN), "and a blocked AudioContext is silent, not an error");
  pass("notices: baseline, new and changed assignments and urgent flags, seen state");

  /* ================= 12. Part B, the office's writes ================= */
  let AITEMS = [], ADDS = [], PATCHES = [], TOASTS = [], ASSIGN_THERE = true, NEXT = 100;
  CW.hasListConsent = async () => true;
  CW.stationSite = async () => "site";
  CW.listItems = async (name) => name === "Fabrication assignments"
    ? (ASSIGN_THERE ? AITEMS.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })) : null) : [];
  CW.listAdd = async (name, fields) => { ADDS.push({ name: name, fields: fields });
    const it = { id: String(NEXT++), fields: Object.assign({}, fields) }; if (name === "Fabrication assignments") AITEMS.push(it); return it; };
  CW.listPatch = async (name, id, fields) => { PATCHES.push({ name: name, id: String(id), fields: fields });
    const hit = AITEMS.find(x => x.id === String(id)); if (name === "Fabrication assignments" && hit) Object.assign(hit.fields, fields); return {}; };
  const SERVER = {};                                  // row id -> fields the list has that this screen has not seen
  CW.listItem = async (name, id) => {
    const it = A("FABR_ITEMS").find(x => String(x.id) === String(id));
    return it ? { id: it.id, fields: Object.assign({}, it.fields, SERVER[id] || {}) } : null;
  };
  global.__TOASTS = TOASTS;
  A("toast = function (m, bad) { __TOASTS.push(String(m)); };");
  global.__R2 = [row, { id: "2", fields: Object.assign({}, row.fields, { Title: "R9001|PVC DOOR", Group: "PVC DOOR",
    Frames: 3, Sashes: 4, Transoms: 0, DoneAt: "" }) }];
  A("FABR_SITEID = 'site'; FABR_OK = true; FABR_ITEMS = __R2.slice(); FABR_ASSIGN_OK = true; FABR_ASSIGN = [];" +
    "FABR_PEOPLE = FABC.fbPeople([{ id: '1', fields: { Title: 'Person A', Station: 'Fabrication', Active: 'Yes', Stages: 'casement windows' } }," +
    " { id: '2', fields: { Title: 'Person B', Station: 'Fabrication', Active: 'Yes', Stages: 'casement windows:frames' } }," +
    " { id: '3', fields: { Title: 'Person C', Station: 'Fabrication', Active: 'Yes', Stages: 'casement windows' } }]);");
  /* the Assign picker offers only people who do that group AND part */
  const pick = part => A("fabrAssignHtml(fabrRecordsNow().byId['1'], fabrRecordsNow().byId['1'].lines.find(l => l.part === '" + part + "'))");
  assert.ok(/Person B/.test(pick("frames")) && !/Person B/.test(pick("sashes")) && /Person A/.test(pick("sashes")),
    "picker filtered per part: Person B does frames only");
  LOGS.length = 0;
  assert.ok(await A("fabrAssign('1', 'frames', 'Person A', 3)"));
  assert.strictEqual(ADDS.length, 1);
  assert.deepStrictEqual([ADDS[0].name, ADDS[0].fields.Status, ADDS[0].fields.Qty, ADDS[0].fields.Person],
    ["Fabrication assignments", "Assigned", 3, "Person A"]);
  await A("fabrAssign('1', 'frames', 'Person A', 2)");
  assert.strictEqual(ADDS.length, 1, "3 + 2 over a total of 4: refused, nothing written");
  assert.ok(TOASTS.some(t => /Only 1 of 4/.test(t)), "and said so");
  await A("fabrAssign('2', 'frames', 'Person A', 1)");
  assert.strictEqual(ADDS.length, 1, "Person A does not do PVC DOOR: refused");
  /* approve a Take */
  AITEMS.push({ id: "50", fields: F.fbRequestFields("R9001", "CASEMENT WINDOWS", "frames", "Person B", 1, "T") });
  await A("fabrApprove('50', 2)");
  assert.strictEqual(PATCHES.length, 0, "approving 2 more over the total is refused");
  await A("fabrApprove('50', 1)");
  assert.deepStrictEqual(PATCHES.map(p => [p.id, p.fields.Status, p.fields.Qty]), [["50", "Assigned", 1]]);
  await A("fabrDecide('50', 'removed')");
  assert.strictEqual(PATCHES[1].fields.Status, "Removed");
  AITEMS.push({ id: "51", fields: F.fbRequestFields("R9001", "CASEMENT WINDOWS", "sashes", "Person B", 6, "T") });
  /* Person B does frames only: a request for sashes cannot be approved */
  const pBeforeB = PATCHES.length;
  await A("fabrApprove('51', 1)");
  assert.strictEqual(PATCHES.length, pBeforeB, "approve refused for a part the person does not do");
  assert.ok(TOASTS.some(t => /Person B does not do CASEMENT WINDOWS sashes/.test(t)), "and said so");
  const aBeforeB = ADDS.length;
  await A("fabrAssign('1', 'sashes', 'Person B', 1)");
  assert.strictEqual(ADDS.length, aBeforeB, "and Assign refuses it too");
  await A("fabrDecide('51', 'refused')");
  assert.deepStrictEqual(Object.keys(PATCHES[2].fields).sort(), ["DecidedAt", "DecidedBy", "Status"],
    "refuse PATCHes Status and the decision stamps, nothing else");
  /* urgent: job level writes every row of the job, one log line */
  const pBefore = PATCHES.length;
  await A("fabrUrgent('job', 'R9001')");
  const up = PATCHES.slice(pBefore);
  assert.deepStrictEqual(up.map(p => [p.name, p.id, p.fields.Urgent]),
    [["Fabrication station", "1", "job"], ["Fabrication station", "2", "job"]], "job-level urgent on every row");
  assert.ok(up.every(p => Object.keys(p.fields).length === 1), "Urgent and nothing else");
  await A("fabrUrgent('part', 'R9001', '1', 'sashes')");
  assert.strictEqual(PATCHES[PATCHES.length - 1].fields.Urgent, "job,sashes");
  assert.deepStrictEqual(LOGS.map(l => l.what.split(":")[0]),
    ["Fabrication assign", "Fabrication approve", "Fabrication remove", "Fabrication refuse",
     "Fabrication urgent", "Fabrication urgent"], "one Dashboard Log line per office write, and none for a refusal");
  assert.ok(ADDS.concat(PATCHES).every(w => w.name === "Fabrication assignments" || w.name === "Fabrication station"),
    "the office never writes Station people or Station log");

  /* P4: urgent is toggled from the row as it is NOW, not the board's copy */
  SERVER["1"] = { Urgent: "job,frames" };            // another screen set these since this one polled
  A("FABR_ITEMS = FABR_ITEMS.map(it => it.id === '1' ? { id: '1', fields: Object.assign({}, it.fields, { Urgent: '' }) } : it);");
  await A("fabrUrgent('group', 'R9001', '1')");
  assert.strictEqual(PATCHES[PATCHES.length - 1].fields.Urgent, "job,group,frames",
    "P4: the other screen's words survive; a stale cache would have written just 'group'");
  delete SERVER["1"];

  /* P3: two office screens give the same line away at the same moment */
  AITEMS.length = 0; LOGS.length = 0; TOASTS.length = 0;
  const addPlain = CW.listAdd;
  CW.listAdd = async (name, fields) => {
    const made = await addPlain(name, fields);
    /* the other screen's write lands in between: 3 of 4 already given */
    AITEMS.push({ id: "77", fields: F.fbAssignFields("R9001", "CASEMENT WINDOWS", "frames", "Person C", 3, "the colleague", "T") });
    return made;
  };
  await A("fabrAssign('1', 'frames', 'Person A', 2)");
  const myAdd = ADDS[ADDS.length - 1];
  const undo = PATCHES[PATCHES.length - 1];
  assert.deepStrictEqual([undo.name, undo.fields.Status], ["Fabrication assignments", "Removed"],
    "P3: the over-the-total assignment this screen wrote is taken back");
  assert.strictEqual(undo.id, AITEMS.find(x => x.fields.Title === myAdd.fields.Title).id, "only its own row");
  assert.strictEqual(AITEMS.find(x => x.id === "77").fields.Status, "Assigned", "the other screen's row is left");
  assert.ok(TOASTS.some(t => /Someone else assigned this line at the same moment — not saved/.test(t)));
  assert.deepStrictEqual(LOGS.map(l => l.what.split(":")[0]), ["Fabrication assign", "Fabrication assign undone"]);
  CW.listAdd = addPlain;
  /* ... and the same for an approval: it goes back to Requested */
  AITEMS.length = 0; LOGS.length = 0;
  AITEMS.push({ id: "60", fields: F.fbRequestFields("R9001", "CASEMENT WINDOWS", "frames", "Person B", 2, "T") });
  const patchPlain = CW.listPatch;
  let raced = false;
  CW.listPatch = async (name, id, fields) => {
    const r = await patchPlain(name, id, fields);
    if (!raced && fields.Status === "Assigned") {
      raced = true;
      AITEMS.push({ id: "78", fields: F.fbAssignFields("R9001", "CASEMENT WINDOWS", "frames", "Person C", 3, "c", "T") });
    }
    return r;
  };
  await A("fabrApprove('60', 2)");
  assert.deepStrictEqual([AITEMS.find(x => x.id === "60").fields.Status, AITEMS.find(x => x.id === "60").fields.Qty],
    ["Requested", 2], "P3: the approval is put back to Requested with the quantity it asked for");
  assert.deepStrictEqual(LOGS.map(l => l.what.split(":")[0]), ["Fabrication approve", "Fabrication approve undone"]);
  CW.listPatch = patchPlain;
  /* no race: nothing is undone */
  AITEMS.length = 0; LOGS.length = 0;
  await A("fabrAssign('1', 'frames', 'Person A', 2)");
  assert.deepStrictEqual(LOGS.map(l => l.what.split(":")[0]), ["Fabrication assign"], "no race, no undo");
  pass("P3 split race undone on this screen's own row only; P4 urgent from the fresh row");

  /* P6: the people are re-read on the five-minute cadence */
  assert.ok(/FABR_PEOPLE && Date\.now\(\) - fabrPeopleAt > 300000 && \(await readFabricationPeople\(\)\)/.test(src("app.js")),
    "P6: fabrPoll re-reads Station people every five minutes");
  /* the list missing: the board says so, quietly */
  ASSIGN_THERE = false;
  await A("readFabricationAssign()");
  assert.strictEqual(A("FABR_ASSIGN_OK"), false);
  assert.ok(A("fabrBoardHtml()").indexOf("“Fabrication assignments” list is not in") >= 0, "the explained state");
  assert.ok(A("fabrBoardHtml()").indexOf("fareqs") < 0);
  pass("Part B office: assign, split refused, approve, remove, refuse, urgent, one log line each, missing list explained");

  /* ================= 13. the tablet page itself (P1, P2, P5) =================
     fabrication.js in a context of its own (it shares top-level names with
     app.js, as every page does), with station-core, station-ui and the core,
     a fake CW that records writes, and stub elements. */
  const TW = [];
  const tel = () => { let h = ""; const e = { style: {}, dataset: {}, hidden: false, textContent: "", value: "",
    className: "", kids: [], appendChild(c) { e.kids.push(c); return c; }, remove() {}, setAttribute() {},
    getAttribute() { return null; }, addEventListener() {}, querySelector: () => null, querySelectorAll: () => [],
    focus() {} }; Object.defineProperty(e, "innerHTML", { get: () => h, set: v => { h = String(v); } }); return e; };
  const TEL = {};
  const tmem = {};
  const tctx = vm.createContext({
    console: { log() {}, warn() {} }, setTimeout: () => 0, clearTimeout() {}, setInterval: () => 0, clearInterval() {},
    Date, JSON, Math, Object, Array, String, Number, Promise, isFinite, RegExp, Error,
    localStorage: { getItem: k => (k in tmem ? tmem[k] : null), setItem: (k, v) => { tmem[k] = String(v); }, removeItem: k => { delete tmem[k]; } },
    document: { documentElement: tel(), activeElement: null, createElement: () => tel(),
                querySelector: s => (TEL[s] = TEL[s] || tel()), querySelectorAll: () => [] },
    CW: { initAuth: async () => null, signIn: async () => null,
          listPatch: async (name, id, fields) => { TW.push({ what: "patch:" + name, id: id, fields: fields }); return {}; },
          listAdd: async (name, fields) => { TW.push({ what: "add:" + name, fields: fields }); return { id: "x" }; },
          listItems: async () => [], listDelta: async () => { throw new Error("no"); }, isDeltaRestart: () => true,
          stationSite: async () => "s", stationSiteMoves: () => 0, isMissing: () => false }
  });
  tctx.window = tctx;
  ["station-core.js", "station-ui.js", "fabrication-core.js", "fabrication.js"].forEach(f =>
    vm.runInContext(src(f), tctx, { filename: f }));
  const T = code => vm.runInContext(code, tctx);
  await new Promise(r => setTimeout(r, 5));
  T("PEOPLE = FABC.fbPeople([{ id: '1', fields: { Title: 'Person A', Station: 'Fabrication', Active: 'Yes', Stages: 'casement windows' } }]);" +
    "PERSON = PEOPLE[0]; PEOPLE_READ = true; READY = true; SITEID = 's'; LAST_TAP = Date.now();" +
    "ITEMS = [{ id: '1', fields: { Title: 'R9001|CASEMENT WINDOWS', Job: 'R9001', Group: 'CASEMENT WINDOWS', Frames: 10, " +
    "  FramesDone: 1, Sashes: 2, Active: 'Yes', OnSheet: 'Yes', Section: 'In production' } }]; boardNow();");
  const setAssign = (state, rows) => T("ASSIGN_OK = " + JSON.stringify(state) + "; ASSIGN_ROWS = FABC.fbAssignRows(" +
    JSON.stringify(rows || []) + "); IDX = FABC.fbAssignIndex(ASSIGN_ROWS);");
  const give = qty => [{ id: "9", fields: { Job: "R9001", Group: "CASEMENT WINDOWS", Part: "frames", Person: "Person A",
                                            Qty: qty, Status: "Assigned" } }];
  /* P1: three states */
  setAssign(null, give(10));
  assert.strictEqual(T("mayTap(recordById('1'), 'frames')"), false, "P1 unknown: locked even though it is held");
  assert.strictEqual(T("lineState(recordById('1'), 'frames').words"), "checking assignments…");
  T("tap('1', 'frames', 1)");
  assert.strictEqual(T("Object.keys(QUEUE).length"), 0, "P1 unknown: a tap queues nothing");
  assert.strictEqual(T("HINT['R9001'].why"), "checking");
  setAssign(false);
  assert.strictEqual(T("mayTap(recordById('1'), 'frames')"), true, "P1 positively missing: Part A, eligible may tap");
  setAssign(true, []);
  assert.strictEqual(T("mayTap(recordById('1'), 'frames')"), false, "P1 read, nothing held: locked");
  setAssign(true, give(2));
  assert.strictEqual(T("mayTap(recordById('1'), 'frames')"), true, "P1 read, held: may tap");
  /* after one good read, a failed read keeps the last good index */
  tctx.CW.listItems = async () => { throw new Error("503 Service Unavailable"); };
  await T("readAssign()");
  assert.strictEqual(T("ASSIGN_OK"), true, "a later failed read keeps the last good state");
  assert.strictEqual(T("mayTap(recordById('1'), 'frames')"), true, "and the last good index");
  /* and a first read that fails leaves it unknown, i.e. locked */
  setAssign(null);
  await T("readAssign()");
  assert.strictEqual(T("ASSIGN_OK"), null, "a first read that fails is not 'missing': still unknown");
  pass("P1 the tablet gate: unknown locks, missing is Part A, read uses the index, a blip keeps the last good read");

  /* P2: holding 2 of 10, All and None are refused with the hint; + works */
  setAssign(true, give(2));
  T("QUEUE = {}; HINT = {}; tap('1', 'frames', 'all')");
  assert.strictEqual(T("Object.keys(QUEUE).length"), 0, "P2: All refused for a part of the line");
  assert.strictEqual(T("HINT['R9001'].why"), "whole");
  T("tap('1', 'frames', 'none')");
  assert.strictEqual(T("Object.keys(QUEUE).length"), 0, "P2: None refused too");
  assert.ok(/All\/None only when the whole line is yours/.test(T("cardInner(boardNow().tabs.floor[0])")));
  assert.ok(/class="sall dead"/.test(T("stepHtml(recordById('1'), recordById('1').lines[0])")), "All is drawn dead");
  T("tap('1', 'frames', 1)");
  assert.strictEqual(T("QUEUE['1|frames'].value"), 2, "P2: + still moves the line");
  setAssign(true, give(10));
  T("QUEUE = {}; tap('1', 'frames', 'all')");
  assert.strictEqual(T("QUEUE['1|frames'].value"), 10, "P2: holding the whole line, All is allowed");
  pass("P2: All/None only when the whole line is the tapper's; − and + always");

  /* P5: a queued tap is dropped at flush once the tapper no longer holds the line.
     Let the flushes the P2 taps started finish first. */
  const settle = () => new Promise(r => setTimeout(r, 30));
  await settle();
  setAssign(true, give(2));
  T("ITEMS[0].fields.FramesDone = 1; READY = false; QUEUE = {}; LOST = {}; boardNow(); tap('1', 'frames', 1); READY = true;");
  assert.strictEqual(T("Object.keys(QUEUE).length"), 1);
  setAssign(true, []);                                // the office removed the assignment
  TW.length = 0;
  await T("flushQueue()"); await settle();
  assert.strictEqual(TW.filter(w => /patch:Fabrication station/.test(w.what)).length, 0, "P5: not sent");
  assert.strictEqual(T("Object.keys(QUEUE).length"), 0, "and not kept");
  assert.ok(/no longer assigned to Person A/.test(T("cardInner(boardNow().tabs.floor[0])")), "and said so on the card");
  /* unknown gate at flush: held, neither sent nor dropped */
  setAssign(true, give(2));
  T("ITEMS[0].fields.FramesDone = 1; READY = false; QUEUE = {}; LOST = {}; boardNow(); tap('1', 'frames', 1); READY = true;");
  setAssign(null, give(2));
  await T("flushQueue()"); await settle();
  assert.strictEqual(T("Object.keys(QUEUE).length"), 1, "P5: an unknown gate holds the tap");
  setAssign(true, give(2));
  TW.length = 0;
  await T("flushQueue()"); await settle();
  assert.strictEqual(TW.filter(w => /patch:Fabrication station/.test(w.what)).length, 1, "and sends it once known and held");
  pass("P5: a tap queued before the assignment was removed is dropped at flush and said so");

  /* ================= 14. finished on the sheet: green (owner, 2026-09-28) ===== */
  ["Ready to fit", "Ready, customer won't take", "Collect & supply only"].forEach(s =>
    assert.ok(F.fbSheetDone(s), s + " is finished on the sheet"));
  ["In production", "Can sell as second hand", "", "Ready to deliver"].forEach(s =>
    assert.ok(!F.fbSheetDone(s), (s || "blank") + " is not"));
  const shRow = (id, section) => ({ id: id, fields: { Title: "J" + id + "|CASEMENT WINDOWS", Job: "J" + id,
    Group: "CASEMENT WINDOWS", Frames: 10, Sashes: 4, Transoms: 0, FramesDone: 3, SashesDone: 0,
    Section: section, Active: /^In production/.test(section) ? "Yes" : "No", OnSheet: "Yes", Seq: Number(id) } });
  const shItems = [shRow("1", "Ready to fit"), shRow("2", "Ready, customer won't take"),
                   shRow("3", "Collect & supply only"), shRow("4", "In production"), shRow("5", "Can sell as second hand")];
  const shCards = F.fbOfficeBoard(shItems);
  const byJ = {}; shCards.forEach(c => { byJ[c.job] = c; });
  ["J1", "J2", "J3"].forEach(j => {
    const c = byJ[j], g = c.groups[0];
    assert.ok(c.sheetDone && c.finished && c.colour === "sheet", j + ": finished, green");
    assert.deepStrictEqual(g.lines.map(l => [l.done, l.total, l.colour]), [[10, 10, "sheet"], [4, 4, "sheet"]],
      j + ": every line drawn full");
    assert.deepStrictEqual([g.raw.frames, g.raw.sashes], [3, 0], j + ": the list's own counts kept, untouched");
  });
  ["J4", "J5"].forEach(j => {
    const c = byJ[j];
    assert.ok(!c.sheetDone && c.colour === "lavender" && c.groups[0].lines[0].done === 3, j + ": unchanged");
  });
  const shTabs = F.fbTabs(shItems);
  assert.deepStrictEqual(shTabs.floor.map(c => c.job), ["J4"], "In production is still On floor");
  assert.ok(["J1", "J2", "J3", "J5"].every(j => shTabs.finished.some(c => c.job === j)), "the three are on the Finished tab");
  const rep = F.fbReportJobs({ board: [byJ.J1] });
  assert.strictEqual(rep.rows[0][rep.columns.indexOf("Frames done")], 3, "the report carries the list's own count");
  /* the tablet: locked, "finished on sheet", a tap refused with the hint */
  T("ITEMS = " + JSON.stringify(shItems) + "; boardNow(); QUEUE = {}; HINT = {};");
  setAssign(false);
  assert.strictEqual(T("lineState(recordById('1'), 'frames').words"), "finished on sheet");
  assert.strictEqual(T("mayTap(recordById('1'), 'frames')"), false, "not tappable, even under Part A");
  T("tap('1', 'frames', 1)");
  assert.deepStrictEqual([T("Object.keys(QUEUE).length"), T("HINT['J1'].why")], [0, "sheet"]);
  assert.strictEqual(T("mayTap(recordById('4'), 'frames')"), true, "In production: tappable as before");
  assert.ok(/finished on sheet/.test(T("cardInner(boardNow().tabs.finished.find(c => c.job === 'J1'))")));
  /* N1: a signed-out (idle-locked) tablet still owes a tap on job J4; the job
     moves to Ready to fit before the flush. The drawn records are stale (they
     are rebuilt only while somebody is signed in) - the list decides. */
  await settle();
  T("READY = false; QUEUE = {}; LOST = {}; boardNow(); tap('4', 'frames', 1); READY = true;");
  assert.strictEqual(T("Object.keys(QUEUE).length"), 1);
  T("PERSON = null; ITEMS = ITEMS.map(it => it.id === '4' ? { id: '4', fields: Object.assign({}, it.fields, " +
    "{ Section: 'Ready to fit', Active: 'No' }) } : it);");
  TW.length = 0;
  await T("flushQueue()"); await settle();
  assert.strictEqual(TW.filter(w => /patch:Fabrication station/.test(w.what)).length, 0, "N1: not sent");
  assert.deepStrictEqual([T("Object.keys(QUEUE).length"), T("LOST['J4'].parts[0].why")], [0, "sheet"],
    "N1: dropped, and kept on the card as 'finished on the sheet'");
  T("PERSON = PEOPLE[0];");
  /* the office: no steppers, the word, and the painter skips the job */
  global.__SH = shItems;
  A("FABR_ITEMS = __SH.map(x => ({ id: x.id, fields: Object.assign({}, x.fields, { DoneAt: '2026-09-28T10:00:00Z' }) }));");
  const offRow = A("FABR_OPEN = { J1: 1 }; fabrRowHtml(fabrRecordsNow().byJob['J1'])");
  assert.ok(/wosheet">finished on sheet/.test(offRow) && !/data-fwact/.test(offRow), "office: the word, no steppers");
  assert.strictEqual(await A("fabrOfficeEdit('1', 'frames', 1)"), false, "office edit refused");
  /* N3: the board's copy says In production, the row as it is NOW says Ready
     to fit: refused on the fresh read, nothing written */
  SERVER["4"] = { Section: "Ready to fit", Active: "No" };
  const pN3 = PATCHES.length;
  A("FRECS_OF = false;");
  assert.strictEqual(await A("fabrOfficeEdit('4', 'frames', 1)"), false, "N3: refused on the fresh row");
  assert.strictEqual(PATCHES.length, pN3, "N3: nothing written");
  delete SERVER["4"];
  global.__SHJ = { id: "J1", src: { Production: 7 }, done: 0, cat: "prod", blk: 0,
    prods: [{ n: "casement windows", f: 10, s: 4, t: 0 }], prodsMain: [{ n: "casement windows", f: 10, s: 4, t: 0 }],
    cp: { win: "", drs: "", glass: {}, prod: {} }, doors: [], glass: {} };
  Object.keys(FILLS).forEach(k => delete FILLS[k]);   // the download: white cells
  A("CP_LIST_OK = true; cpSetImportPending(false); cpRowsSet({}); FABR_PAINTED = {};");
  assert.strictEqual(A("fabrColourPlan(__SHJ)"), null, "the painter skips a job finished on the sheet");
  A("FABR_ITEMS = FABR_ITEMS.map(it => it.id === '1' ? { id: '1', fields: Object.assign({}, it.fields, { Section: 'In production' }) } : it);");
  assert.ok(A("fabrColourPlan(__SHJ)"), "... and the same job In production would be painted (the skip is the reason)");
  pass("finished on the sheet: the three sections green, full, locked; In production and second hand unchanged; painter skips");

  /* ================= 9b. the Floor log reads every station's LOG (review, 2026-09-28) =====
     This suite loads all four cores and app.js, so it is where "opening the
     Floor log from the job list reads welding's, glazing's and fabrication's
     logs" can be proven. The readers are replaced by recorders (no network in
     this run): a station whose site is not resolved yet reads its board first,
     which is what finds the site, and then its log. */
  const CALLS = [];
  A("WELD_SITEID = null; GLZ_SITEID = null; FABR_SITEID = null;" +
    "WELD_LOG_OK = null; GLZ_LOG_OK = null; FABR_LOG_OK = null; WELD_OK = null; GLZ_OK = null; FABR_OK = null;");
  global.__rec = (name, then, site) => { CALLS.push(name); if (site) A(site); if (then) then(); };
  A("stationLogReadIfNeeded = t => __rec('glass-log', t);" +
    "weldReadIfNeeded = t => __rec('weld-board', t, \"WELD_SITEID = 'S'\");" +
    "weldLogReadIfNeeded = t => __rec('weld-log', t);" +
    "glzReadIfNeeded = t => __rec('glz-board', t, \"GLZ_SITEID = 'S'\");" +
    "glzLogReadIfNeeded = t => __rec('glz-log', t);" +
    "fabrReadIfNeeded = (w, t) => __rec('fabr-' + w, t, w === 'board' ? \"FABR_SITEID = 'S'\" : '');");
  A("state.board = null;");
  const lhAt = NULLABLE.indexOf("#lhost"); NULLABLE.splice(lhAt, 1);   // the window opens this time
  A("openStationLog('')");
  ["glass-log", "weld-board", "weld-log", "glz-board", "glz-log", "fabr-board", "fabr-log"].forEach(c =>
    assert.ok(CALLS.indexOf(c) >= 0, "opening the Floor log from the job list calls " + c));
  assert.ok(CALLS.indexOf("weld-log") > CALLS.indexOf("weld-board"), "the log after the board has found the site");
  CALLS.length = 0;
  A("openStationLog('')");                     // sites known now: straight to the logs
  assert.ok(CALLS.indexOf("weld-log") >= 0 && CALLS.indexOf("weld-board") < 0, "a known site goes straight to its log");
  /* review fix 2: a station still being read says so, one quiet line each */
  A("STATION_LOG_OK = true; STATION_LOG = []; WELD_LOG_OK = null; GLZ_LOG_OK = true; GLZ_LOG = [];" +
    "FABR_LOG_OK = null; FABR_OK = false; FABR_WHY = 'the fabrication site is not there';");
  A("paintStationLog()");
  const lg = EL["#lgbody"].innerHTML;
  assert.ok(lg.indexOf("Welding: still reading…") >= 0, "a station still being read says so");
  assert.ok(lg.indexOf("Fabrication: the fabrication site is not there") >= 0,
    "a station whose site failed says its board's reason, never “still reading” for ever");
  NULLABLE.splice(lhAt, 0, "#lhost");
  A("if (stationPollT) { clearTimeout(stationPollT); stationPollT = null; }");   // openStationLog armed the tick
  pass("the Floor log reads each station's log (board first when its site is unknown), and says who is still reading");

  /* ================= 15. the view filter: Everything | My work | Assigned (2026-10-01) ===== */
  const vRow = (id, job, group, f, s, t, section) => ({ id: id, fields: { Title: job + "|" + group, Job: job,
    Group: group, Frames: f, Sashes: s, Transoms: t, FramesDone: 0, Section: section || "In production",
    Active: section ? "No" : "Yes", OnSheet: "Yes", Seq: Number(id) } });
  const vItems = [vRow("1", "V1", "PVC SMART", 2, 2, 0), vRow("2", "V1", "CASEMENT WINDOWS", 4, 4, 1),
                  vRow("3", "V2", "CASEMENT WINDOWS", 3, 0, 0), vRow("4", "V3", "PVC SMART", 1, 1, 0, "Ready to fit")];
  const vCards = F.fbOfficeBoard(vItems);
  const smart = F.fbPeople([{ id: "1", fields: { Title: "Person S", Station: "Fabrication", Active: "Yes",
    Stages: "PVC SMART:sashes" } }])[0];
  const allSash = F.fbPeople([{ id: "2", fields: { Title: "Person T", Station: "Fabrication", Active: "Yes",
    Stages: "ALL:sashes" } }])[0];
  const shape = cs => cs.map(c => c.job + "[" + c.groups.map(g => g.group + ":" + g.lines.map(l => l.part).join("+")).join(",") + "]");
  assert.deepStrictEqual(F.fbViewFilter(vCards, smart, "all", {}), vCards, "Everything is the board as it is");
  assert.deepStrictEqual(shape(F.fbViewFilter(vCards, smart, "mine", {})),
    ["V1[PVC SMART:sashes]", "V3[PVC SMART:sashes]"],
    "My work, PVC SMART:sashes: only that part; V2 (no eligible part) gone; the green V3 kept, it has an eligible part");
  assert.deepStrictEqual(shape(F.fbViewFilter(vCards, allSash, "mine", {})),
    ["V1[CASEMENT WINDOWS:sashes,PVC SMART:sashes]", "V3[PVC SMART:sashes]"],
    "My work, ALL:sashes: the sashes of every group; V2 has no sashes and goes");
  const mineV1 = F.fbViewFilter(vCards, allSash, "mine", {})[0];
  assert.deepStrictEqual([mineV1.total, mineV1.groups[0].total], [6, 4], "counts are re-added from what is drawn");
  assert.strictEqual(vCards[0].total, 13, "and the board's own card is untouched");
  const vIdx = F.fbAssignIndex(F.fbAssignRows([
    { id: "1", fields: { Job: "V1", Group: "CASEMENT WINDOWS", Part: "frames", Person: "Person T", Qty: 2, Status: "Assigned" } },
    { id: "2", fields: { Job: "V1", Group: "PVC SMART", Part: "sashes", Person: "Person T", Qty: 1, Status: "Requested" } },
    { id: "3", fields: { Job: "V2", Group: "CASEMENT WINDOWS", Part: "frames", Person: "Person T", Qty: 3, Status: "Removed" } },
    { id: "4", fields: { Job: "V1", Group: "CASEMENT WINDOWS", Part: "sashes", Person: "Person S", Qty: 4, Status: "Assigned" } }]));
  assert.deepStrictEqual(shape(F.fbViewFilter(vCards, allSash, "assigned", vIdx)), ["V1[CASEMENT WINDOWS:frames]"],
    "Assigned to me: only the Assigned row - not the Requested, not the Removed, not someone else's");
  assert.deepStrictEqual(F.fbViewFilter(vCards, smart, "nonsense", vIdx), vCards, "an unknown mode is Everything");
  assert.deepStrictEqual([F.fbViewFor({ "Person S": "mine" }, "Person S"), F.fbViewFor({}, "Person S"),
                          F.fbViewFor({ "Person S": "junk" }, "Person S")], ["mine", "all", "all"]);

  /* the tablet page: tabs, search and memory follow the view; the gate does not */
  T("ITEMS = " + JSON.stringify(vItems) + "; PEOPLE = FABC.fbPeople([{ id: '1', fields: { Title: 'Person S', " +
    "Station: 'Fabrication', Active: 'Yes', Stages: 'PVC SMART:sashes' } }, { id: '2', fields: { Title: 'Person T', " +
    "Station: 'Fabrication', Active: 'Yes', Stages: 'ALL:sashes' } }]); QUERY = ''; TAB = 'floor';");
  setAssign(false);
  T("pickPerson(PEOPLE[0]);");
  assert.strictEqual(T("VIEW"), "all", "default after sign-in: Everything");
  assert.deepStrictEqual(JSON.parse(T("JSON.stringify([boardNow().tabs.floor.length, boardNow().tabs.finished.length])")), [2, 1]);
  T("setView('mine'); render();");
  assert.deepStrictEqual(JSON.parse(T("JSON.stringify([boardNow().tabs.floor.length, boardNow().tabs.finished.length])")), [1, 1],
    "My work: the tab counts follow");
  assert.ok(/On floor · 1/.test(T("document.querySelector('#tabfloor').textContent")), "the tab button says so");
  assert.strictEqual(T("F.fbSearchTab(boardNow().tabs, 'V2', 'floor').tab"), "floor",
    "search follows: V2 is not Person S's work, so it moves no tab");
  assert.strictEqual(T("mayTap(recordById('2'), 'frames')"), false, "the gate is unchanged: still not theirs");
  assert.strictEqual(T("recordById('2').group"), "CASEMENT WINDOWS", "and the gate reads the whole board's records");
  T("pickPerson(PEOPLE[1]);");
  assert.strictEqual(T("VIEW"), "all", "remembered PER PERSON: Person T starts at Everything");
  T("setView('assigned'); render();");
  T("pickPerson(PEOPLE[0]);");
  assert.strictEqual(T("VIEW"), "mine", "Person S comes back to My work");
  assert.deepStrictEqual(JSON.parse(T("localStorage.getItem('cw_fabview')")), { "Person S": "mine", "Person T": "assigned" });
  /* empty states */
  T("pickPerson(PEOPLE[1]); TAB = 'floor'; render();");
  assert.ok(/Assignments are not set up yet/.test(T("document.querySelector('#board').innerHTML")),
    "Assigned to me with no list: said so");
  setAssign(null);
  T("render();");
  assert.ok(/Assignments not loaded yet/.test(T("document.querySelector('#board').innerHTML")), "and while unknown");
  setAssign(true, []);
  T("render();");
  assert.ok(/Nothing for you here — switch to Everything to see all jobs\./.test(T("document.querySelector('#board').innerHTML")),
    "nothing assigned: the plain line");
  /* notifications are not filtered: an assignment on a group the view hides still notifies */
  assert.ok(/now\.all\.floor\.concat\(now\.all\.finished\)/.test(src("fabrication.js")), "notices read the whole board");
  pass("view filter: Everything / My work / Assigned to me, per person, tabs and search follow, gate unchanged");

  /* ================= 10. the gates ================= */
  const fsrc = src("fabrication-core.js") + src("fabrication.js") + src("fabrication.html");
  ["setFill", "clearFill", "setValues", "appendLog", "saveProgress", "moveJobRow", "batchWrite",
   "/workbook", "downloadWorkbook", "parseWorkbook", "ExcelJS", "listDelete", '"DELETE"'].forEach(bad =>
    assert.strictEqual(fsrc.indexOf(bad), -1, "a fabrication station file names `" + bad + "`"));
  const page = src("fabrication.html");
  ["graph.js", "station-core.js", "station-ui.js", "fabrication-core.js", "fabrication.js"]
    .forEach(s => assert.ok(page.indexOf(s) > 0, "fabrication.html loads " + s));
  ["exceljs", "parser.js", "app.js", "checkpoints.js"].forEach(s =>
    assert.strictEqual(page.toLowerCase().indexOf(s), -1, "fabrication.html must not load " + s));
  assert.ok(/:root\[data-theme="light"\]/.test(page) && /\.sbtn \{ width:64px; height:56px/.test(page));
  assert.ok(src("index.html").indexOf("fabrication-core.js") > 0, "the office page loads the core");
  const mine = fsrc + src("test_fabrication.js");
  assert.ok(!/@(?!example\.test)[a-z0-9-]+\.(com|ie|net|org|co\.uk)/i.test(mine), "no address or domain");
  assert.strictEqual(NET.length, 0, "nothing in this run reached the network");
  pass("no station file names a workbook write or a delete; no network; no real address");

  console.log("\n" + n + " checks passed");
})().catch(e => { console.error(e); process.exit(1); });
