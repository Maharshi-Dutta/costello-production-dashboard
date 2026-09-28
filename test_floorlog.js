/* Offline test of the one Floor log for every station (2026-09-28).
   Spec: docs/specs/2026-09-28-day-sheets-welding-hotmelt-floor-log.md.

   The pure half, in station-core.js: which `Dashboard Log` lines are the
   office's own edits of a floor counter (one list, matched on the exact words
   the boards write), the two time-stamp shapes that list carries, and the
   merge of four stations' `Station log` lines with them - newest first by the
   moment each names, narrowed by station. The window itself (one station's
   list missing still shows the rest) is exercised in test_station.js, which
   already runs app.js.

   Every person here is invented. Run: node test_floorlog.js               */
const fs = require("fs"), vm = require("vm"), assert = require("assert");
global.window = {};
vm.runInThisContext(fs.readFileSync(__dirname + "/station-core.js", "utf8"), { filename: "station-core.js" });
const ST = global.window.ST;

let n = 0;
const pass = t => { n++; console.log("  ok  " + t); };

/* ---- 1. which Dashboard Log lines are office floor edits ---- */
const E = (what, job) => ST.officeFloorEdit(what, job);
assert.deepStrictEqual(E("Glass cutting", "R5301"), { station: "Glass", stage: "cut" });
assert.deepStrictEqual(E("Glass hotmelting", "R5301"), { station: "Glass", stage: "hotmelt" });
assert.deepStrictEqual(E("Glass tuff", "R5301"), { station: "Glass", stage: "tuff" });
assert.deepStrictEqual(E("Floor glass counters", "R5301"), { station: "Glass", stage: "" });
assert.deepStrictEqual(E("Welding: R5303 CASEMENT WINDOWS frames", "R5303"), { station: "Welding", stage: "frames" });
assert.deepStrictEqual(E("Glazing: R5304", "R5304"), { station: "Glazing", stage: "glaze" });
assert.deepStrictEqual(E("Glazing astragal: R5304", "R5304"), { station: "Glazing", stage: "astragal" });
assert.deepStrictEqual(E("Fabrication: R5305 SUPER DOOR transoms", "R5305"), { station: "Fabrication", stage: "transoms" });
assert.deepStrictEqual(E("Day sheet corrected", "(Welding weld)"), { station: "Welding", stage: "day sheet" });
assert.deepStrictEqual(E("Day sheet corrected", "(Glass hotmelt)"), { station: "Glass", stage: "day sheet" });
assert.deepStrictEqual(E("Day sheet corrected", "something else"), { station: "", stage: "day sheet" },
  "a correction whose station cannot be read is still an office edit, of no known station");
/* the words the builders really write, not a copy of them */
vm.runInThisContext(fs.readFileSync(__dirname + "/welding-core.js", "utf8"), { filename: "welding-core.js" });
vm.runInThisContext(fs.readFileSync(__dirname + "/glazing-core.js", "utf8"), { filename: "glazing-core.js" });
vm.runInThisContext(fs.readFileSync(__dirname + "/fabrication-core.js", "utf8"), { filename: "fabrication-core.js" });
const WELDC = global.window.WELDC, GLZC = global.window.GLZC, FABC = global.window.FABC;
assert.strictEqual(E(WELDC.weldLogWords("R1", "PVC DOOR", "sashes"), "R1").station, "Welding");
assert.strictEqual(E(GLZC.glzLogWords("R1", "astragal"), "R1").stage, "astragal");
assert.strictEqual(E(GLZC.glzLogWords("R1"), "R1").stage, "glaze");
assert.strictEqual(E(FABC.fbLogWords("R1", "CASEMENT", "frames"), "R1").station, "Fabrication");
["Export", "Phase", "Section", "Comment", "Fabrication colours", "Floor glass colours",
 "Fabrication urgent: R1 X", "Fabrication assign: R1 X frames", "Cutting weekly target", "Alert", ""]
  .forEach(w => assert.strictEqual(E(w, "R1"), null, "“" + w + "” is not an office edit of a floor counter"));
pass("the office's floor edits are matched off one list, on the exact words each board writes");

/* ---- 2. the time stamps ---- */
assert.strictEqual(ST.logWhenMs("21/09/2026 14:05"), new Date(2026, 8, 21, 14, 5).getTime(), "dd/mm/yyyy hh:mm, local");
assert.strictEqual(ST.logWhenMs("2026-09-21 14:05"), new Date(2026, 8, 21, 14, 5).getTime(), "Dashboard Log's own text");
assert.strictEqual(ST.logWhenMs("2026-09-21T13:05:00.000Z"), Date.UTC(2026, 8, 21, 13, 5), "an ISO stamp");
assert.strictEqual(ST.logWhenMs("03/09/2026 08:00"), new Date(2026, 8, 3, 8, 0).getTime(), "3 September, not 9 March");
assert.ok(isNaN(ST.logWhenMs("yesterday")) && isNaN(ST.logWhenMs("")), "rubbish is NaN, never a throw");
pass("Dashboard Log's dd/mm/yyyy hh:mm and yyyy-mm-dd hh:mm are local times; ISO is ISO");

/* ---- 3. the merge ---- */
const at = (d, h, m) => new Date(2026, 8, d, h, m).toISOString();
const line = (id, job, station, stage, from, to, who, when) =>
  ({ id: id, fields: { Title: job, Station: station, GlassType: "X", Stage: stage, From: from, To: to, Who: who, At: when } });
const glass = ST.logRows([line("1", "R1", "Glass", "cut", 0, 4, "Person A", at(21, 9, 0)),
                          line("2", "R1", "Glass", "hotmelt", 0, 2, "Person B", at(21, 11, 0))], "Glass");
const weld = ST.logRows([line("3", "R2", "Welding", "frames", 0, 6, "Person C", at(21, 10, 0)),
                         line("4", "R2", "Welding", "frames-remake", 0, 1, "Person C", at(21, 10, 30))], "Welding");
const glz = ST.logRows([line("5", "R3", "Glazing", "glaze", 0, 3, "Person D", at(21, 8, 0))], "Glazing");
const fab = ST.logRows([line("6", "R4", "Fabrication", "transoms", 0, 2, "Person E", at(21, 12, 0))], "Fabrication");
const CHANGES = [
  { at: "21/09/2026 10:15", who: "the office", job: "R1", what: "Glass cutting", from: "4", to: "5", shared: true },
  { at: "2026-09-21 12:30", who: "the colleague", job: "R2", what: "Welding: R2 PVC DOOR frames", from: "6", to: "0" },
  { at: new Date(2026, 8, 21, 7, 0).toISOString(), who: "the office", job: "(Glass cut)", what: "Day sheet corrected",
    from: "2026-09-20 Person A: 12", to: "14" },
  { at: "21/09/2026 13:00", who: "the office", job: "(nowhere)", what: "Day sheet corrected", from: "a", to: "b" },
  { at: "21/09/2026 13:30", who: "the office", job: "R9", what: "Phase", from: "1", to: "2" }
];
const office = ST.officeLogRows(CHANGES);
assert.strictEqual(office.length, 4, "four office floor edits; the Phase line is not one");
assert.ok(office.every(r => r.office), "each marked office");
assert.strictEqual(office[2].job, "", "a day-sheet correction is not about a job");
assert.strictEqual(office[0].from, "4", "from/to kept as written");
const src = [{ station: "Glass", rows: glass }, { station: "Welding", rows: weld },
             { station: "Glazing", rows: glz }, { station: "Fabrication", rows: fab }];
const all = ST.floorLogMerge(src, office, "");
assert.strictEqual(all.length, 10, "six floor lines and four office lines");
const order = all.map(r => (r.office ? "o:" : "") + r.station + ":" + (r.stage || "-"));
assert.deepStrictEqual(order, ["o::day sheet", "o:Welding:frames", "Fabrication:transoms",
  "Glass:hotmelt", "Welding:frames-remake", "o:Glass:cut", "Welding:frames", "Glass:cut",
  "Glazing:glaze", "o:Glass:day sheet"],
  "newest first by the moment each names - ISO and local text stamps sorted together");
const onlyWeld = ST.floorLogMerge(src, office, "Welding");
assert.deepStrictEqual(onlyWeld.map(r => r.stage), ["frames", "frames-remake", "frames"],
  "the station filter: that station's lines and its office edits");
assert.ok(ST.floorLogMerge(src, office, "Glass").every(r => r.station === "Glass"),
  "and an office line of no known station shows under All only");
assert.strictEqual(ST.logFilter(onlyWeld, { stage: "frames-remake" }).length, 1, "the stage filter within a station");
assert.strictEqual(ST.logFilter(all, { stage: "frames" }).length, 2,
  "under All the stage filter takes the floor's line and the office's edit alike");
const noWeld = ST.floorLogMerge([src[0], { station: "Welding", rows: [] }, src[2], src[3]], office, "");
assert.strictEqual(noWeld.filter(r => !r.office).length, 4, "one station with nothing read: the rest still show");
pass("four stations and the office merge newest first, and narrow by station and by stage");

/* ---- 4. the words per station ---- */
assert.strictEqual(ST.floorStageLabel("Glass", "hotmelt"), "Hotmelting");
assert.strictEqual(ST.floorStageLabel("Glass", "glazed"), "Glazing", "the pre-2026-09-21 glass word still reads");
assert.strictEqual(ST.floorStageLabel("Welding", "sashes-remake"), "Sashes remade");
assert.strictEqual(ST.floorStageLabel("Glazing", "astragal"), "Astragal");
assert.strictEqual(ST.floorStageLabel("Fabrication", "transoms"), "Transoms");
assert.deepStrictEqual(ST.FLOOR_LOG_STATIONS, ["Glass", "Welding", "Glazing", "Fabrication"]);
pass("each station's stage filter has its own words");

console.log("\n" + n + " checks passed");
