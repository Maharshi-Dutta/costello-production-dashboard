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

  /* M2: under a view a card is placed by the lines it SHOWS */
  const vDone = vItems.map(it => it.id === "1" ? { id: "1", fields: Object.assign({}, it.fields, { SashesDone: 2 }) } : it);
  const vTabs = F.fbTabs(vDone);
  assert.deepStrictEqual(vTabs.floor.map(c => c.job), ["V1", "V2"], "Everything: V1 is on the floor (its casement is not done)");
  const vMine = F.fbViewTabs(vTabs, smart, "mine", {});
  assert.deepStrictEqual([vMine.floor.map(c => c.job), vMine.finished.map(c => c.job)], [[], ["V1", "V3"]],
    "My work: V1's only shown line (PVC SMART sashes) is done, so V1 is Finished; the green V3 stays Finished");
  const vMineT = F.fbViewTabs(vTabs, allSash, "mine", {});
  assert.deepStrictEqual(vMineT.floor.map(c => c.job), ["V1"], "ALL:sashes: V1's casement sashes are not done, so On floor");
  assert.deepStrictEqual(F.fbViewTabs(vTabs, smart, "all", {}).floor.map(c => c.job), ["V1", "V2"], "Everything unchanged");
  const vUrg = F.fbTabs(vItems.map(it => it.id === "3" ? { id: "3", fields: Object.assign({}, it.fields, { Urgent: "job" }) } : it));
  assert.deepStrictEqual(F.fbViewTabs(vUrg, allSash, "mine", {}).floor.map(c => c.job), ["V1"], "the sort keeps working");
  pass("M2: under a view, a card's tab follows the lines it shows; green stays Finished; Everything unchanged");

  /* M1: the empty state is decided from the CURRENT tab */
  const boardHtml = () => T("document.querySelector('#board').innerHTML");
  setAssign(false);
  T("ITEMS = " + JSON.stringify(vDone) + "; pickPerson(PEOPLE[0]); setView('mine'); TAB = 'floor'; QUERY = ''; render();");
  assert.ok(/Nothing for you here — 2 for you in Finished ›/.test(boardHtml()) && /id="gotab"/.test(boardHtml()),
    "M1(b): nothing of theirs on the floor, 2 in Finished: offered as a tap");
  T("PEOPLE.push(FABC.fbPeople([{ id: '9', fields: { Title: 'Person U', Station: 'Fabrication', Active: 'Yes', " +
    "Stages: 'SIDELIGHTS' } }])[0]); pickPerson(PEOPLE[PEOPLE.length - 1]); setView('mine'); TAB = 'floor'; render();");
  assert.ok(/Nothing for you here — switch to Everything to see all jobs\./.test(boardHtml()),
    "M1(c): nothing of theirs in either tab");
  T("ITEMS = ITEMS.filter(it => it.id !== '4'); TAB = 'finished'; render();");
  assert.ok(/No finished jobs yet\./.test(boardHtml()), "M1(a): the whole board's tab is empty: the plain wording");
  T("TAB = 'floor'; setView('all'); pickPerson(PEOPLE[0]);");
  pass("M1: empty state from the current tab: plain / other tab offered / switch to Everything");

  /* ================= 16. glass status, door glazing, gold (2026-10-01) ================= */
  /* A. the job's glass, in one word */
  const GS = (g, has) => F.fbGlassStatus(g, has === undefined ? true : has);
  assert.strictEqual(GS(null, false), "", "no glass: blank");
  assert.strictEqual(GS(null), "none", "glass and no row: none");
  assert.strictEqual(GS({ total: 20, cut: 0, hotmelt: 0 }), "none");
  assert.strictEqual(GS({ total: 20, cut: 12, hotmelt: 5 }), "part:12/20:5/20");
  assert.strictEqual(GS({ total: 20, cut: 20, hotmelt: 20 }), "done", "cut and hotmelt complete, no tuff: done");
  assert.strictEqual(GS({ total: 20, cut: 20, hotmelt: 20, tuffTotal: 4, tuff: 1 }), "part:20/20:20/20:tuff 1/4",
    "the Tuff rule: a job with Tuff is not done until Tuff is");
  assert.strictEqual(GS({ total: 20, cut: 20, hotmelt: 20, tuffTotal: 4, tuff: 4 }), "done");
  assert.strictEqual(GS({ total: 20, cut: 3, hotmelt: 0, officeDone: true }), "done", "the office's record says done");
  /* A-1: OfficeDone speaks for DG and TG only - a job that owes tuff is not done */
  assert.strictEqual(GS({ total: 20, cut: 3, hotmelt: 0, officeDone: true, tuffTotal: 4, tuff: 1 }), "part:20/20:20/20:tuff 1/4",
    "A-1: the office's done with tuff owed is not done");
  assert.strictEqual(GS({ total: 20, officeDone: true, tuffTotal: 4, tuff: 4 }), "done", "A-1: ... and is, once tuff is complete");
  assert.strictEqual(GS({ total: 0, tuffTotal: 4, tuff: 2, officeDone: true }), "part:0/0:0/0:tuff 2/4", "A-1: the no-DG/TG branch too");
  assert.strictEqual(GS({ total: 0, tuffTotal: 4, tuff: 4 }), "done");
  assert.deepStrictEqual(F.fbGlassChip("done"), { kind: "done", words: "Glass ✓ done" });
  assert.deepStrictEqual(F.fbGlassChip("part:12/20:5/20"), { kind: "part", words: "Glass: cut 12/20 · hotmelt 5/20" });
  assert.strictEqual(F.fbGlassChip("part:20/20:20/20:tuff 1/4").words, "Glass: cut 20/20 · hotmelt 20/20 · tuff 1/4");
  assert.deepStrictEqual([F.fbGlassChip("none").kind, F.fbGlassChip("").kind, F.fbGlassChip("junk").kind], ["none", "", ""]);
  pass("glass status: blank, none, part, done incl. the Tuff rule and the office's done; the chip words");

  /* B. door codes -> group; the slice; Stages; the gate */
  assert.deepStrictEqual(["SS", "CD", "SFCD", "BF", "ACSD", "ACSS", "PVC", "DD", "SD", "1 DOOR", "xyz", ""].map(F.fbDoorGlazeGroup),
    ["PVC SMART", "", "", "", "", "", "PVC DOOR", "PVC DOOR", "PVC DOOR", "PVC DOOR", "PVC DOOR", ""],
    "SS is PVC SMART; CD / SFCD / BF / AC* never; every other code, known or not, is PVC DOOR");
  const gJob = { id: "R9100", cust: "Person A", cat: "prod", blk: 1, seq: 1,
    doors: [{ slot: 1, code: "PVC" }, { slot: 2, code: "CD" }, { slot: 3, code: "SS" }, { slot: 4, code: "DD" }],
    prodsMain: [{ n: "pvc door", f: 2, s: 2, t: 0 }, { n: "casement windows", f: 1, s: 1, t: 0 }] };
  const gSlice = F.fbSlice([gJob], names, () => "part:1/2:0/2");
  const gPvc = gSlice.find(r => r.group === "PVC DOOR"), gCas = gSlice.find(r => r.group === "CASEMENT WINDOWS");
  assert.deepStrictEqual([gPvc.glazeTotal, gCas.glazeTotal], [2, 0], "PVC + DD on the PVC DOOR row; no PVC SMART row, so SS is fed nowhere");
  assert.ok(gSlice.every(r => r.glass === "part:1/2:0/2"), "Glass on every row of the job");
  assert.deepStrictEqual([F.fbFeederFields(gPvc).GlazeTotal, F.fbFeederFields(gPvc).Glass], [2, "part:1/2:0/2"]);
  assert.notStrictEqual(ST.sliceHash(gSlice, F.FAB), ST.sliceHash(F.fbSlice([gJob], names, () => "done"), F.FAB),
    "a glass change re-feeds");
  ["GlazeDone", "GlazeBy", "GlazeAt"].forEach(k => assert.ok(F.FB_FEEDER_WRITES.indexOf(k) < 0 && F.FB_FLOOR_FIELDS.indexOf(k) >= 0,
    k + " is the floor's, never the feeder's"));
  assert.deepStrictEqual(P("PVC DOOR"), { "pvc door": all3 }, "a bare group is still frames+sashes+transoms only");
  assert.deepStrictEqual(P("PVC DOOR:glazing"), { "pvc door": { glazing: true } });
  assert.deepStrictEqual(P("ALL:glazing"), { "pvc door": { glazing: true }, "pvc smart": { glazing: true } },
    "ALL:glazing is the two glazed groups");
  assert.deepStrictEqual(P("CASEMENT WINDOWS:glazing"), {}, "glazing exists on PVC DOOR and PVC SMART only");
  const glazer = F.fbPeople([{ id: "1", fields: { Title: "Person G", Station: "Fabrication", Active: "Yes", Stages: "ALL:glazing" } }])[0];
  const fabber = F.fbPeople([{ id: "2", fields: { Title: "Person F", Station: "Fabrication", Active: "Yes", Stages: "PVC DOOR" } }])[0];
  [true, false, null].forEach(st => {
    assert.ok(F.fbCanTap(glazer, "PVC DOOR", "glazing", {}, st, "R9100"), "role only: no assignment needed (state " + st + ")");
    assert.ok(!F.fbCanTap(fabber, "PVC DOOR", "glazing", {}, st, "R9100"), "a bare PVC DOOR person may not glaze");
  });
  assert.ok(!F.fbCanTap(glazer, "PVC DOOR", "frames", {}, false, "R9100"), "and the glazer may not fabricate");
  assert.ok(F.fbActAllowed(glazer, {}, true, "R9100", "PVC DOOR", "glazing", 2, "all"), "All / None allowed on glazing");
  assert.strictEqual(F.fbRequestFields("R9100", "PVC DOOR", "glazing", "Person G", 1, "T"), null, "no Take on glazing");
  /* its own count: not in the fabrication totals, colour or tabs */
  const gItems = [{ id: "1", fields: Object.assign({ Title: gPvc.title }, F.fbFeederFields(gPvc),
                    { FramesDone: 2, SashesDone: 2, GlazeDone: 1, GlazeBy: "Person G", GlazeAt: "2026-10-01T09:00:00Z" }) },
                  { id: "2", fields: Object.assign({ Title: gCas.title }, F.fbFeederFields(gCas), { FramesDone: 1, SashesDone: 1 }) }];
  const gCard = F.fbOfficeBoard(gItems)[0], gRec = gCard.groups.find(g => g.group === "PVC DOOR");
  assert.deepStrictEqual([gRec.total, gRec.done, gRec.colour, gRec.lines.length, gRec.extra.length], [4, 4, "purple", 2, 1],
    "the group's fabrication total, done and colour do not count glazing");
  assert.deepStrictEqual([gRec.extra[0].part, gRec.extra[0].done, gRec.extra[0].total], ["glazing", 1, 2]);
  assert.ok(gCard.finished && F.fbTabs(gItems).finished.length === 1, "finished for the tabs exactly as before, glazing owed or not");
  assert.strictEqual(F.fbApplyTap(gRec, "glazing", "all"), 2);
  assert.deepStrictEqual(Object.keys(F.fbFloorOnly(F.fbTapFields("glazing", 2, "Person G", "T"))).sort(),
    ["DoneAt", "DoneBy", "GlazeAt", "GlazeBy", "GlazeDone"]);
  assert.strictEqual(F.fbLogEntry({ job: "R9100", group: "PVC DOOR", part: "glazing", from: 1, to: 2 }).stage, "glazing");
  assert.deepStrictEqual(F.FB_PARTS, ["frames", "sashes", "transoms"], "the painter's and the assignments' parts are unchanged");
  /* the views: My work shows it to the glazer and places the card by it; Assigned never */
  const gMine = F.fbViewTabs(F.fbTabs(gItems), glazer, "mine", {});
  assert.deepStrictEqual([gMine.floor.map(c => c.job), gMine.floor[0].groups.map(g => g.group + ":" + g.lines.length + "+" + g.extra.length)],
    [["R9100"], ["PVC DOOR:0+1"]], "My work for a glazer: only the glazing line, and the job is On floor while it is owed");
  assert.strictEqual(F.fbViewFilter(F.fbOfficeBoard(gItems), glazer, "assigned", {}).length, 0, "Assigned to me never shows glazing");
  assert.strictEqual(F.fbViewFilter(F.fbOfficeBoard(gItems), fabber, "mine", {})[0].groups[0].extra.length, 0);
  const sdItems = gItems.map(it => ({ id: it.id, fields: Object.assign({}, it.fields, { Section: "Ready to fit", Active: "No" }) }));
  const sdRec = F.fbOfficeBoard(sdItems)[0].groups.find(g => g.group === "PVC DOOR");
  assert.deepStrictEqual([sdRec.extra[0].done, sdRec.extra[0].colour, sdRec.raw.glazing], [2, "sheet", 1], "finished on sheet: drawn full, raw kept");
  pass("door glazing: code table, slice, Stages (named only), role-only gate, its own count, views, finished-on-sheet");

  /* Glass ready first */
  const gfCards = F.fbOfficeBoard([
    { id: "1", fields: { Title: "A1|X", Job: "A1", Group: "X", Frames: 1, Seq: 1, OnSheet: "Yes", Active: "Yes", Glass: "none" } },
    { id: "2", fields: { Title: "B2|X", Job: "B2", Group: "X", Frames: 1, Seq: 2, OnSheet: "Yes", Active: "Yes", Glass: "done" } },
    { id: "3", fields: { Title: "C3|X", Job: "C3", Group: "X", Frames: 1, Seq: 3, OnSheet: "Yes", Active: "Yes", Urgent: "job" } },
    { id: "4", fields: { Title: "D4|X", Job: "D4", Group: "X", Frames: 1, FramesDone: 1, Seq: 4, OnSheet: "Yes", Active: "Yes", Glass: "done" } }]);
  assert.deepStrictEqual(F.fbGlassFirst(gfCards, true).map(c => c.job), ["C3", "B2", "A1", "D4"],
    "urgent stays first, then glass-ready and unfinished; a finished job is not lifted");
  assert.deepStrictEqual(F.fbGlassFirst(gfCards, false).map(c => c.job), ["C3", "A1", "B2", "D4"], "off: urgent first only");
  assert.deepStrictEqual(gfCards.map(c => !!c.glassStart), [false, true, false, false], "the start line: glass done, nothing fabricated");
  pass("Glass ready first: after urgent, unfinished glass-done jobs; the start line");

  /* C. the gold plan (pure) */
  const gDoors = [{ slot: 1, code: "PVC" }, { slot: 2, code: "CD" }, { slot: 3, code: "SS" }, { slot: 4, code: "DD" }];
  const GP = (o, states) => F.fbGoldPlan(Object.assign({ doors: gDoors, group: "PVC DOOR", total: 2, done: 2, fabModified: 1000,
    stateOf: s => (states || {})[s] || {} }, o));
  assert.deepStrictEqual(GP(), [1, 4], "full: the PVC DOOR doors; CD never; SS is PVC SMART's");
  assert.deepStrictEqual(GP({ done: 1 }), [], "partial paints nothing");
  assert.deepStrictEqual(GP({ done: 0 }), [], "a count that dropped raises nothing - and there is nothing in a plan that clears");
  assert.deepStrictEqual(GP({}, { 1: { status: "done" } }), [4], "already done: left");
  assert.deepStrictEqual(GP({}, { 1: { pending: true }, 4: { foreign: true } }), [], "a write in the air, or a colour not ours: skipped");
  assert.deepStrictEqual(GP({}, { 1: { status: "", office: true, modified: 2000 }, 4: { status: "process", office: true, modified: 500 } }), [4],
    "an office row modified (server clock) after the glazing stands - its clear is not gilded again; an older one is raised");
  assert.deepStrictEqual(GP({}, { 1: { status: "", office: true, modified: 500 } }), [1, 4], "an office clear OLDER than the glazing is raised");
  /* C-1 fail-safe: a server stamp missing on either side */
  assert.deepStrictEqual(GP({ fabModified: 0 }, { 1: { status: "", office: true, modified: 500 }, 4: { status: "process", office: true, modified: 500 } }), [4],
    "no fabrication stamp: an office row is only ever raised from process, never from a clear");
  assert.deepStrictEqual(GP({}, { 1: { status: "", office: true, modified: 0 }, 4: { status: "", office: false, modified: 0 } }), [4],
    "no record stamp: the office's clear stands; a row that is not the office's (import, floor) is raised");
  assert.deepStrictEqual(GP({ group: "PVC SMART", total: 1, done: 1 }), [3]);
  pass("gold plan: full only, mapped doors only, raise only, never over a later office row, CD never");

  /* C. end to end, through the checkpoint record path */
  const FILLCALLS = [];
  CW.setFill = async (sheet, addr, color) => { FILLCALLS.push({ sheet: sheet, addr: addr, color: color }); };
  const goldJob = { id: "R9100", src: { Production: 7 }, done: 0, cat: "prod", blk: 1, drs: 4, wnd: 0,
    doors: gDoors.map(d => Object.assign({ status: "" }, d)), prods: [], prodsMain: [], glass: {},
    cp: { win: "", drs: "", glass: {}, prod: {} } };
  global.__GJ = goldJob;
  global.__GI = [{ id: "1", fields: { Title: "R9100|PVC DOOR", Job: "R9100", Group: "PVC DOOR", Frames: 2, Sashes: 2,
    GlazeTotal: 2, GlazeDone: 1, GlazeBy: "Person G", GlazeAt: "2026-10-01T09:00:00Z", Section: "In production",
    Active: "Yes", OnSheet: "Yes" } }];
  const setGlaze = n => A("FABR_ITEMS = [{ id: '1', fields: Object.assign({}, __GI[0].fields, { GlazeDone: " + n + " }) }];");
  A("ALL = [__GJ]; BLOCKNAMES = ['Ready to fit', 'In production']; PRODMAP = { prod: {}, glass: {}, qty: { drs: 14 }, hdr: [2, 3]," +
    " doors: { 1: 72, 2: 73, 3: 74, 4: 75 } }; CP_LIST_OK = true; cpSetImportPending(false); cpRowsSet({}); CP_ITEMS = [];" +
    " PAINTED = {}; FABR_OK = true; Object.keys(FABR_GOLD_FAIL).forEach(k => delete FABR_GOLD_FAIL[k]);");
  /* the job list's own row renderer wants a whole parsed job; it is not under test */
  A("renderRows = function () {}; renderDrawer = function () {};");
  ADDS.length = 0; PATCHES.length = 0; LOGS.length = 0;
  setGlaze(1);
  assert.strictEqual(await A("fabrGoldRun()"), 0, "partial: nothing");
  assert.deepStrictEqual([ADDS.length, FILLCALLS.length], [0, 0]);
  setGlaze(2);
  assert.strictEqual(await A("fabrGoldRun()"), 2, "full: the two PVC DOOR doors raised");
  const recs = ADDS.filter(a => a.name === "Dashboard progress");
  assert.deepStrictEqual(recs.map(a => [a.fields.Title, a.fields.Status, a.fields.Source, a.fields.Who]),
    [["R9100|door:1", "done", "fabrication", "Person G"], ["R9100|door:4", "done", "fabrication", "Person G"]],
    "the RECORD first: a Dashboard progress row per door, Source fabrication, Who the glazer");
  const doorFills = FILLCALLS.filter(f => /^(BT|BW)7$/.test(f.addr));
  assert.deepStrictEqual(doorFills.map(f => [f.sheet, f.addr]), [["Production", "BT7"], ["Production", "BW7"]],
    "then the fill, from the record, on the job's own row, those two door cells");
  assert.ok(doorFills.every(f => f.color === A("CP_WORD_HEX.done")), "gold");
  assert.ok(!FILLCALLS.some(f => /^(BU|BV)7$/.test(f.addr)), "the CD cell and the SS cell are never touched");
  assert.deepStrictEqual(LOGS.map(l => l.what), ["Door 1", "Door 4"], "logged as an office tick on a door is");
  assert.ok(/door glazing complete/.test(LOGS[0].to));
  assert.deepStrictEqual(A("[paintedOf('R9100', 'door:1'), cpStatus(__GJ, 'door:1')]").join(), "done,done", "painted and recorded");
  const nAdds = ADDS.length, nFills = FILLCALLS.length;
  assert.strictEqual(await A("fabrGoldRun()"), 0, "already done: nothing again");
  setGlaze(0);
  assert.strictEqual(await A("fabrGoldRun()"), 0, "the count dropped: nothing is cleared");
  assert.deepStrictEqual([ADDS.length, FILLCALLS.length, A("cpStatus(__GJ, 'door:1')")], [nAdds, nFills, "done"]);
  assert.ok(recs.every(a => a.fields.When !== "2026-10-01T09:00:00Z" && /^20\d\d-/.test(a.fields.When)),
    "C-2: the record row is stamped with the office's own clock at the gild, never the tablet's GlazeAt");
  assert.ok(/glazed /.test(LOGS[0].to), "... which stays in the log words only");

  /* C-1: who spoke last is the SERVER's clock (Modified on both rows), never the tablet's */
  const fabRow = (glazeAt, modified) => A("Object.keys(FABR_GLAZE_MOD).forEach(k => delete FABR_GLAZE_MOD[k]);" +
    "FABR_ITEMS = [{ id: '1', fields: Object.assign({}, __GI[0].fields, { GlazeDone: 2, GlazeAt: '" + glazeAt +
    "', Modified: '" + modified + "' }) }];");
  const officeClear = (modified, status) => A("cpRowPut('R9100', 'door:1', { id: '900', job: 'R9100', item: 'door:1', status: '" +
    (status || "") + "', done: 0, total: 1, source: 'office', when: '2026-10-01T10:00:00Z', modified: '" + modified + "' });");
  fabRow("2026-10-01T09:00:00Z", "2026-10-01T09:00:05Z");
  officeClear("2026-10-01T10:00:00Z");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "the office cleared the door after the glazing was recorded: it stands");
  fabRow("2026-10-03T23:00:00Z", "2026-10-01T09:00:05Z");            // a FAST tablet: GlazeAt two days ahead
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "C-1 fast tablet clock: the office's later clear is still not gilded over");
  fabRow("2026-09-20T08:00:00Z", "2026-10-01T11:00:00Z");            // a SLOW tablet: GlazeAt days behind, glazed after the clear
  assert.deepStrictEqual(Array.from(A("fabrGoldPlan(__GJ).map(p => p.item)")), ["door:1"],
    "C-1 slow tablet clock: glazing recorded after the office's clear is not blocked");
  /* a feeder patch of the same row later (its Glass word) must not make an old glazing look new */
  fabRow("2026-10-01T09:00:00Z", "2026-10-01T09:00:05Z");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null);
  A("FABR_ITEMS = [{ id: '1', fields: Object.assign({}, FABR_ITEMS[0].fields, { Glass: 'done', Modified: '2026-10-01T12:00:00Z' }) }];");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "the glazing's server stamp is kept from when GlazeAt last changed");
  /* fail-safe: a server stamp missing */
  fabRow("2026-10-01T09:00:00Z", "");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "no fabrication stamp: an office clear is never gilded over");
  officeClear("", "process");
  assert.strictEqual(A("fabrGoldPlan(__GJ).length"), 1, "... but an office 'in fabrication' is raised");
  setGlaze(2);
  /* ... a write in the air on that item, a gold row, and a finished-on-sheet job are skipped */
  A("cpRowPut('R9100', 'door:1', null);");
  assert.strictEqual(A("fabrGoldPlan(__GJ).length"), 1, "with no row it would be raised");
  A("CPBURST['R9100|door:1'] = { key: 'R9100|door:1' };");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "cpPending: skipped");
  A("delete CPBURST['R9100|door:1'];");
  goldJob.done = 1;
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "a gold row is whole");
  goldJob.done = 0;
  A("FABR_ITEMS = [{ id: '1', fields: Object.assign({}, __GI[0].fields, { GlazeDone: 2, Section: 'Ready to fit', Active: 'No' }) }];");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "finished on sheet: skipped");
  setGlaze(2);
  A("cpSetImportPending(true);");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "import pending: skipped");
  A("cpSetImportPending(false); CP_LIST_OK = false;");
  assert.strictEqual(A("fabrGoldPlan(__GJ)"), null, "the record list unreadable: skipped");
  A("CP_LIST_OK = true;");
  assert.ok(src("fabrication.js").indexOf("Dashboard progress") < 0 && src("fabrication-core.js").indexOf("cpSaveRow") < 0,
    "the tablet never writes the record");

  /* C-2: a SECOND office screen whose record is stale. The list already holds
     both doors done (the first screen gilded them); this screen has not heard.
     Inside the chain it refreshes the record first, so it adds no second row
     and no second log line. */
  const serverRows = ["door:1", "door:4"].map((it, i) => ({ id: String(700 + i), fields: { Title: "R9100|" + it, Job: "R9100",
    Item: it, Done: 1, Total: 1, Status: "done", Who: "Person G", When: "2026-10-01T09:10:00Z", Source: "fabrication",
    Modified: "2026-10-01T09:10:00Z" } }));
  let deltaAsked = 0;
  const realDelta = CW.listDelta;
  CW.listDelta = async name => {
    if (name !== "Dashboard progress") throw new Error("no delta in this test");
    deltaAsked++;
    return { items: serverRows.map(x => ({ id: x.id, fields: Object.assign({}, x.fields) })), next: "t" };
  };
  A("cpRowsSet({}); CP_ITEMS = []; stationResetFeed('progress'); PAINTED = {};" +
    "Object.keys(FABR_GLAZE_MOD).forEach(k => delete FABR_GLAZE_MOD[k]);");
  setGlaze(2);
  const a2 = ADDS.length, p2 = PATCHES.length, l2 = LOGS.length, f2 = FILLCALLS.length;
  assert.strictEqual(A("fabrGoldPlan(__GJ).length"), 2, "the stale screen would plan both doors");
  assert.strictEqual(await A("fabrGoldRun()"), 0, "C-2: after the refresh inside the chain there is nothing left to raise");
  assert.ok(deltaAsked > 0, "the record was refreshed before the re-plan");
  assert.deepStrictEqual([ADDS.length, PATCHES.length, LOGS.length, FILLCALLS.length], [a2, p2, l2, f2],
    "C-2: no second row, no PATCH, no second log line, no fill");
  /* (d) a row that is there is PATCHed, never added beside */
  serverRows[0].fields.Status = "process"; serverRows[0].fields.Source = "import";
  A("cpRowsSet({}); CP_ITEMS = []; stationResetFeed('progress');");
  assert.strictEqual(await A("fabrGoldRun()"), 1);
  assert.deepStrictEqual([ADDS.length, PATCHES.length - p2, PATCHES[PATCHES.length - 1].id], [a2, 1, "700"],
    "C-2: the existing row is PATCHed by its id; nothing is added");
  CW.listDelta = realDelta;
  pass("gold end to end: record first (Source fabrication), then the fill; raise only; office clear, pending, gold row, sheet-done skipped");

  /* the painter still knows nothing of glazing */
  A("FABR_PAINTED = {}; cpRowsSet({}); PRODMAP.prod = { 'pvc door': { f: 62, s: 63 } };");
  goldJob.prodsMain = [{ n: "pvc door", f: 2, s: 2, t: 0 }]; goldJob.prods = goldJob.prodsMain;
  A("FABR_ITEMS = [{ id: '1', fields: Object.assign({}, __GI[0].fields, { FramesDone: 1, GlazeDone: 2, DoneAt: '2026-10-01T09:00:00Z' }) }];");
  assert.deepStrictEqual(Array.from(A("(fabrColourPlan(__GJ) || []).map(p => p.item)")), ["prod:pvc door:f"],
    "the lavender/purple painter plans F/S/T cells only");

  /* the office board: the chip, the glazing line with steppers and no Assign, an office edit */
  A("FABR_ASSIGN_OK = true; FABR_ASSIGN = []; FABR_ITEMS = [{ id: '1', fields: Object.assign({}, __GI[0].fields, { Glass: 'done', GlazeDone: 1 }) }];");
  const gRow = A("FABR_OPEN = { R9100: 1 }; fabrRowHtml(fabrRecordsNow().byJob['R9100'])");
  assert.ok(/gchip g-done/.test(gRow) && /Glass ✓ done/.test(gRow), "the chip on the board");
  assert.ok(/Glass is ready and nothing is fabricated yet — start this job/.test(gRow), "the start line");
  assert.ok(/Door glazing/.test(gRow) && /data-fwpart="glazing" data-fwact="1"/.test(gRow), "the glazing line has steppers");
  assert.ok(!/data-fasg="1\|glazing"/.test(gRow) && /data-fasg="1\|frames"/.test(gRow), "and no Assign: a role, not an assignment");
  CW.listItem = async (name, id) => { const it = A("FABR_ITEMS").find(x => String(x.id) === String(id)); return it ? { id: it.id, fields: Object.assign({}, it.fields) } : null; };
  PATCHES.length = 0; LOGS.length = 0;
  assert.strictEqual(await A("fabrOfficeEdit('1', 'glazing', 'all')"), true);
  assert.deepStrictEqual(Object.keys(PATCHES[0].fields).sort(), ["DoneAt", "DoneBy", "GlazeAt", "GlazeBy", "GlazeDone"],
    "the office edit: the counter, GlazeBy/At, DoneBy/At");
  assert.ok(/^Fabrication: R9100 PVC DOOR glazing$/.test(LOGS[0].what) && !ADDS.some(a => a.name === "Station log"),
    "one Dashboard Log line, never Station log");
  await new Promise(r => setTimeout(r, 30));           // the gold run the edit started
  pass("office: chip, start line, glazing line with steppers and no Assign, the five-field edit, painter unchanged");

  /* A. the office's glass word, and a list without the new columns */
  global.__GLJ = { id: "R9200", glass: { dg: 10, tg: 10 } };
  A("STATION_OK = true; STATION_ITEMS = [{ id: '1', fields: { Title: 'R9200', Job: 'R9200', Total: 20, Cut: 12, Hotmelt: 5, Active: 'Yes' } }];");
  assert.strictEqual(A("fabrGlassOf(__GLJ)"), "part:12/20:5/20", "derived from the glass row this page already holds");
  assert.strictEqual(A("fabrGlassOf({ id: 'R9201', glass: {} })"), "", "no glass: blank");
  A("STATION_OK = false;");
  assert.strictEqual(A("fabrGlassOf(__GLJ)"), "", "the glass list unreadable: blank, never a 'not started' nobody said");
  const sent = [];
  global.__SEND = async body => {
    if ("Glass" in body) throw new Error("PATCH x -> 400 {\"error\":\"Field 'Glass' is not recognized\"}");
    sent.push(body); return {};
  };
  await A("fabrFeedWrite({ Customer: 'x', Glass: 'done', GlazeTotal: 2, FedAt: 'T', FedBy: 'o' }, __SEND)");
  assert.deepStrictEqual([Object.keys(sent[0]).sort(), A("fabrColsMissing()").join()],
    [["Customer", "FedAt", "FedBy", "GlazeTotal"], "Glass"], "a 400 on the new column: sent without it, and that column remembered");
  await A("fabrFeedWrite({ Glass: 'done', FedAt: 'T', FedBy: 'o' }, __SEND)");
  assert.strictEqual(sent.length, 1, "afterwards a write that only carried it is not sent at all");
  assert.ok(/has no “Glass” column yet/.test(A("fabrBoardHtml()")), "and the board says which column is missing");
  global.__SEND2 = async () => { throw new Error("PATCH x -> 503 busy"); };
  await assert.rejects(A("fabrFeedWrite({ GlazeTotal: 2, FedAt: 'T' }, __SEND2)"), /503/, "anything but a 400 is just a failed write");
  assert.strictEqual(A("fabrColsMissing()").join(), "Glass", "and marks no column missing");
  A("delete FABR_NOCOL.Glass;");
  /* F-1: a 400 that does not NAME the column is an ordinary failed write */
  global.__SEND3 = async () => { throw new Error("POST x -> 400 {\"error\":\"The list item could not be added: duplicate Title\"}"); };
  await assert.rejects(A("fabrFeedWrite({ Customer: 'x', Glass: 'done', GlazeTotal: 2, FedAt: 'T' }, __SEND3)"), /duplicate Title/,
    "F-1: a 400 about something else is not read as a missing column");
  assert.strictEqual(A("fabrColsMissing()").join(), "", "F-1: and strips nothing");
  let tries = 0;
  global.__SEND4 = async body => { tries++; if ("GlazeTotal" in body) throw new Error("PATCH x -> 400 Field 'GlazeTotal' is not recognized"); return {}; };
  await A("fabrFeedWrite({ Customer: 'x', Glass: 'done', GlazeTotal: 2, FedAt: 'T' }, __SEND4)");
  assert.deepStrictEqual([A("fabrColsMissing()").join(), tries], ["GlazeTotal", 2], "F-1: only the NAMED column is tried without, and remembered");
  A("delete FABR_NOCOL.GlazeTotal;");

  /* F-2: the fabrication feed waits for the glass list - it never writes a
     blank Glass that it would have to write again */
  const fabAdds = () => ADDS.filter(a => a.name === "Fabrication station");
  global.__FJ = { id: "R9300", cust: "Person A", cat: "prod", blk: 1, seq: 1, glass: { dg: 10, tg: 10 }, doors: [], notes: [],
    prods: [{ n: "casement windows", f: 2, s: 2, t: 0 }], prodsMain: [{ n: "casement windows", f: 2, s: 2, t: 0 }],
    cp: { win: "", drs: "", glass: {}, prod: {} } };
  CW.listItems = async () => [];
  A("ALL = [__FJ]; BLOCKNAMES = ['Ready to fit', 'In production']; FABR_FEED = { hash: '', at: 0 }; fabrBusy = false; fabrPolling = false;" +
    "STATION_OK = null; STATION_ITEMS = null;");
  const fa0 = fabAdds().length;
  for (let i = 0; i < 5; i++) { await A("feedFabrication()"); A("clearTimeout(fabrAgainT); fabrAgainT = null;"); }
  assert.strictEqual(fabAdds().length, fa0, "F-2: glass list not read yet: nothing is fed, however many turns pass");
  A("STATION_OK = true; STATION_ITEMS = [{ id: '1', fields: { Title: 'R9300', Job: 'R9300', Total: 20, Cut: 20, Hotmelt: 4, Active: 'Yes' } }];");
  await A("feedFabrication()");
  A("clearTimeout(fabrAgainT); fabrAgainT = null;");
  assert.deepStrictEqual(fabAdds().slice(fa0).map(a => a.fields.Glass), ["part:20/20:4/20"],
    "F-2: fed once, with the glass word - never blank first");
  A("STATION_OK = false; FABR_FEED = { hash: '', at: 0 };");
  ADDS.length = 0;
  await A("feedFabrication()");
  A("clearTimeout(fabrAgainT); fabrAgainT = null;");
  assert.deepStrictEqual(fabAdds().map(a => a.fields.Glass), [""], "F-2: a glass list that is not there feeds blank");
  assert.ok(/fabrSoonT = setTimeout\([\s\S]{0,120}, 60000\)/.test(src("app.js")), "F-2: the glass-follow debounce is a minute");
  pass("glass word from the glass row; a missing new column is stripped, remembered and said, the feed carries on");

  /* D. welding and glazing boards are told after a feed that wrote rows */
  global.__RD = { w: 0, g: 0 };
  A("redrawWelding = function () { __RD.w++; }; redrawGlazing = function () { __RD.g++; };" +
    "WELD_FEED = { hash: '', at: 0 }; GLZ_FEED = { hash: '', at: 0 }; ALL = [__JOB]; BLOCKNAMES = ['Ready to fit', 'In production'];");
  job.wndMain = 3; job.blk = 1; job.done = 0;
  CW.listItemsFor = async () => [];
  const rw = await A("feedWelding()"), rz = await A("feedGlazing()");
  assert.ok(rw && rw.sent > 0 && global.__RD.w > 0, "D: feedWelding redraws the welding board after writing rows");
  assert.ok(rz && rz.sent > 0 && global.__RD.g > 0, "D: feedGlazing redraws the glazing board after writing rows");
  pass("D: redrawWelding / redrawGlazing after a feed that wrote rows");

  /* the tablet page: the glazing line, the chip, the toggle */
  T("ITEMS = " + JSON.stringify([
      { id: "1", fields: Object.assign({}, gItems[0].fields, { Glass: "done", FramesDone: 0, SashesDone: 0 }) },
      { id: "2", fields: { Title: "R9050|CASEMENT WINDOWS", Job: "R9050", Group: "CASEMENT WINDOWS", Frames: 1, Seq: 0,
                           Section: "In production", Active: "Yes", OnSheet: "Yes", Glass: "part:1/2:0/2" } }]) +
    "; PEOPLE = FABC.fbPeople([{ id: '1', fields: { Title: 'Person G', Station: 'Fabrication', Active: 'Yes', Stages: 'ALL:glazing' } }," +
    " { id: '2', fields: { Title: 'Person F', Station: 'Fabrication', Active: 'Yes', Stages: 'PVC DOOR' } }]);" +
    " QUERY = ''; TAB = 'floor'; QUEUE = {}; LOST = {}; HINT = {};");
  setAssign(null);
  T("pickPerson(PEOPLE[1]); setView('all'); setGlassFirst(false); render();");     // Person G
  assert.deepStrictEqual(JSON.parse(T("JSON.stringify(boardNow().tabs.floor.map(c => c.job))")), ["R9050", "R9100"], "board order");
  assert.deepStrictEqual(JSON.parse(T("JSON.stringify([lineState(recordById('1'), 'glazing'), lineState(recordById('1'), 'frames').words])")),
    [{ can: true, words: "" }, "not your part"], "the glazer: the glazing line is theirs, with no assignment and no Take");
  T("READY = false; tap('1', 'glazing', 'all'); READY = true;");
  assert.strictEqual(T("QUEUE['1|glazing'].value"), 2, "a tap queues with the assignments list unknown: role only; All allowed");
  T("QUEUE = {}; boardNow(); pickPerson(PEOPLE[0]);");                           // Person F: bare PVC DOOR
  assert.strictEqual(T("lineState(recordById('1'), 'glazing').words"), "not your part", "a non-glazing person: locked");
  T("READY = false; tap('1', 'glazing', 1); READY = true;");
  assert.deepStrictEqual([T("Object.keys(QUEUE).length"), T("HINT['R9100'].why")], [0, "part"]);
  const gHtml = T("cardInner(boardNow().tabs.floor.find(c => c.job === 'R9100'))");
  assert.ok(/gchip g-done">Glass ✓ done/.test(gHtml) && /Glass is ready and nothing is fabricated yet — start this job/.test(gHtml) &&
            /Door glazing/.test(gHtml), "the chip, the start line and the glazing line on the card");
  assert.ok(/gchip g-part">Glass: cut 1\/2 · hotmelt 0\/2/.test(T("cardInner(boardNow().tabs.floor.find(c => c.job === 'R9050'))")));
  T("setGlassFirst(true);");
  assert.deepStrictEqual(JSON.parse(T("JSON.stringify(boardNow().tabs.floor.map(c => c.job))")), ["R9100", "R9050"], "Glass ready first");
  assert.strictEqual(T("localStorage.getItem('cw_fabglassfirst')"), "1", "remembered per tablet");
  T("setView('mine');");
  assert.deepStrictEqual(JSON.parse(T("JSON.stringify(boardNow().tabs.floor.map(c => c.job))")), ["R9100"], "and it works under My work");
  T("setGlassFirst(false); setView('all');");
  pass("tablet: glazing line for a glazer and a non-glazer, chip states, start line, Glass ready first");

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
