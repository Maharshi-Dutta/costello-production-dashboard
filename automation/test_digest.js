/* Offline test of the alerts digest.

   Two things are exercised with the same fixtures:
     1. automation/digest-core.js  - required directly;
     2. automation/alerts-digest.ts - the generated Office Script, run through
        new Function() after its TypeScript annotations are stripped, driven by
        an in-memory fake `workbook` that implements only the three methods the
        script uses (getWorksheet / getUsedRange / getValues).
   The two must agree, and the .ts must still contain the current core.

   Nothing leaves the box: no network, no Graph, no Excel, no real addresses.
   Every address here is a reserved test domain (example.test / other.test).

   Run:  node automation/test_digest.js                                      */

const fs = require("fs");
const path = require("path");
const assert = require("assert");

const DIR = __dirname;
const core = require(path.join(DIR, "digest-core.js"));

/* ---------- fixtures ------------------------------------------------------
   A miniature Production sheet with all five sections, laid out like the real
   one: job numbers in column C (index 2), the header pair on rows 2 and 3, and
   the two notes columns out past K where a fixed A1:K600 read would miss them. */

const ADMIN = "Admin@Example.test";            // mixed case on purpose
const SOMEONE = "someone@example.test";
const ANOTHER = "another@example.test";
const LONELY = "lonely@example.test";
const OUTSIDER = "outsider@other.test";        // wrong domain: must never be mailed
const URL = "https://dashboard.example.test/";

const C_JOB = 2, C_CUST = 4, C_COMMENT = 5, C_BRENDAN = 12, C_SPECIALS = 13;

function blank() { const r = []; for (let i = 0; i < 14; i++) r.push(""); return r; }
function jobRow(id, o) {
  o = o || {};
  const r = blank();
  r[C_JOB] = id;
  r[C_CUST] = o.cust || "";
  r[C_COMMENT] = o.comment || "";
  r[C_BRENDAN] = o.brendan || "";
  r[C_SPECIALS] = o.specials || "";
  return r;
}
function dividerRow(text) { const r = blank(); r[0] = text; return r; }

const HDR_GROUP = blank();
HDR_GROUP[1] = "OFFICE NO"; HDR_GROUP[C_JOB] = "JOB NO"; HDR_GROUP[3] = "DATES ON CONTRACT";
HDR_GROUP[C_CUST] = "CUSTOMER"; HDR_GROUP[C_COMMENT] = "COMMENT";
HDR_GROUP[C_BRENDAN] = "Notes from Brendan"; HDR_GROUP[C_SPECIALS] = "Notes of specials";
const HDR_SUB = blank();
HDR_SUB[3] = "SOLD";

const PRODUCTION = [
  blank(),                                              // 1  title strip
  HDR_GROUP,                                            // 2  group headers
  HDR_SUB,                                              // 3  sub headers
  jobRow("A1000", { cust: "Cust A" }),                  // 4  can sell as second hand
  dividerRow("CUSTOMERS WON'T TAKE"),                   // 5
  jobRow("B2000", { cust: "Cust B" }),                  // 6  ready, customer won't take
  dividerRow("COLLECT OR SUPPLY ONLY"),                 // 7
  jobRow("C3000", { cust: "Cust C" }),                  // 8  collect & supply only
  dividerRow("NOT SENT TO FLOOR"),                      // 9
  jobRow("D4000", { cust: "Cust D", comment: "ready" }),// 10 ready to fit  -> finished
  dividerRow("NOT SENT TO FLOOR"),                      // 11
  jobRow("E5000", { cust: "Cust E", comment: 'Waiting on <glass> & trim "urgent"',
                    brendan: "Chase the supplier", specials: "Arch head" }),   // 12
  jobRow("F6000", { cust: "Cust F" }),                  // 13 no comments at all
  jobRow("G7000", { cust: "Cust G", comment: "Sills cut" }),                   // 14
  jobRow("H8000", { cust: "Cust H", comment: "not sent to floor until Friday" })// 15 must NOT split the sheet
];

const CONFIG = [
  ["Key", "Value"],
  ["admin", ADMIN],
  ["Dashboard URL", URL],
  ["", ""],
  ["somethingElse", "ignored"]
];

const ALERTS = [
  ["Job", "Email", "Added by", "When"],
  ["A1000", SOMEONE, "the admin", "2026-09-01"],        // second-hand block -> finished
  ["D4000", SOMEONE, "the admin", "2026-09-01"],        // ready to fit      -> finished
  ["E5000", SOMEONE, "the admin", "2026-09-01"],
  ["H8000", SOMEONE, "the admin", "2026-09-01"],
  ["E5000", SOMEONE, "the admin", "2026-09-02"],        // duplicate pair, must count once
  ["F6000", ANOTHER, "the admin", "2026-09-01"],
  ["G7000", OUTSIDER, "the admin", "2026-09-01"],       // wrong domain
  ["D4000", LONELY, "the admin", "2026-09-01"],         // only a finished job -> no mail
  ["", "", "", ""],                                     // a removed row is blanked, not deleted
  ["Z9999", SOMEONE, "the admin", "2026-09-01"]         // no longer on the Production sheet
];

const LOG = [
  ["When", "Who", "Job", "What changed", "From", "To"],
  ["2026-09-01 09:00", "user one", "E5000", "Comment", "", "Frames delivered"],
  ["2026-09-02 09:10", "user two", "E5000", "Section", "Ready to fit", "In production"],
  ["2026-09-02 09:20", "user two", "E5000", "Comment", "", "Glass ordered"],
  ["2026-09-02 09:30", "user two", "E5000", "Comment", "", "   "],   // empty once trimmed
  ["2026-09-03 08:00", "user one", "F6000", "Alert", "", ANOTHER],   // not a comment
  ["2026-09-03 08:05", "user one", "", "Comment", "", "no job number"]
];

const NOW = new Date(2026, 8, 5, 8, 0, 0);      // 5 September 2026, 08:00 local

/* ---------- the same workbook, but not written from A1 --------------------
   A real sheet can start anywhere: leading blank rows, and on Production a
   blank column A and B (the job number is in column C, so that is as far right
   as its used range can begin - C3). getUsedRange() then hands back only that
   block, and readSheet in the Office Script pads it back to true coordinates.
   These fixtures are laid out at their proper coordinates; usedRangeAt() slices
   off the empty part the way Excel does, and checks that it really was empty. */

const HDR_GROUP_C3 = blank();
HDR_GROUP_C3[C_JOB] = "JOB NO"; HDR_GROUP_C3[3] = "DATES ON CONTRACT";
HDR_GROUP_C3[C_CUST] = "CUSTOMER"; HDR_GROUP_C3[C_COMMENT] = "COMMENT";
HDR_GROUP_C3[C_BRENDAN] = "Notes from Brendan"; HDR_GROUP_C3[C_SPECIALS] = "Notes of specials";
function dividerRowD(text) { const r = blank(); r[3] = text; return r; }   // divider text in column D

const PRODUCTION_C3 = [
  [], [],                                               // 1-2  blank
  HDR_GROUP_C3,                                         // 3    group headers, from column C
  HDR_SUB,                                              // 4    sub headers
  dividerRowD("NOT SENT TO FLOOR"),                     // 5    -> ready to fit
  jobRow("D4000", { cust: "Cust D" }),                  // 6    finished
  dividerRowD("NOT SENT TO FLOOR"),                     // 7    -> in production
  jobRow("E5000", { cust: "Cust E", comment: "Waiting on <glass>",
                    brendan: "Chase the supplier", specials: "Arch head" }),   // 8
  jobRow("F6000", { cust: "Cust F" })                   // 9    no comments
];
const CONFIG_A3 = [[], [], ["Key", "Value"], ["admin", ADMIN], ["Dashboard URL", URL]];
const ALERTS_A3 = [[], [], ["Job", "Email", "Added by", "When"],
                   ["D4000", SOMEONE, "the admin", "2026-09-01"],
                   ["E5000", SOMEONE, "the admin", "2026-09-01"],
                   ["F6000", SOMEONE, "the admin", "2026-09-01"],
                   ["E5000", OUTSIDER, "the admin", "2026-09-01"]];
const LOG_A3 = [[], [], ["When", "Who", "Job", "What changed", "From", "To"],
                ["2026-09-02 09:20", "user two", "E5000", "Comment", "", "Glass ordered"]];

const OFFSET_BOOK = {
  "Dashboard Config": usedRangeAt(CONFIG_A3, 2, 0),
  "Dashboard Alerts": usedRangeAt(ALERTS_A3, 2, 0),
  "Production": usedRangeAt(PRODUCTION_C3, 2, 2),      // C3
  "Dashboard Log": usedRangeAt(LOG_A3, 2, 0)
};

/* ---------- the generated Office Script, made runnable ------------------- */

const TS_PATH = path.join(DIR, "alerts-digest.ts");
const TS = fs.readFileSync(TS_PATH, "utf8");

/** Strip the TypeScript from the generated script so Node can run it. The type
    names are not guessed: build-script.js writes the list it used into the
    "@types:" line, so this can never drift out of date with the generator. */
function stripTypes(ts) {
  const m = /\/\* @types: ([^*]+)\*\//.exec(ts);
  assert.ok(m, "the generated script has no @types line");
  const names = m[1].trim().split("|").map(function (n) { return n.replace(/\./g, "\\."); });
  const start = ts.indexOf("/* @types:"), end = ts.indexOf("/* @types-end */");
  assert.ok(start >= 0 && end > start, "the generated script has no type block");
  let js = ts.slice(0, start) + ts.slice(end + "/* @types-end */".length);
  const annotation = new RegExp(":\\s*(?:" + names.join("|") + ")(?:\\s*\\[\\s*\\])*", "g");
  js = js.replace(annotation, "");
  assert.ok(js.indexOf("interface ") < 0, "an interface survived the strip");
  assert.ok(js.indexOf("/*:") < 0, "an annotation comment survived the strip");
  return js;
}
const SCRIPT_MAIN = new Function(stripTypes(TS) + "\nreturn main;")();

/** The used range of a sheet whose content really starts at (top, left) - the
    rows above it and the columns to its left being empty. Excel hands back only
    the used block plus its 0-based position, exactly as modelled here, and
    readSheet in the script has to pad it back out. */
function usedRangeAt(full, top, left) {
  const values = [];
  for (let r = 0; r < top; r++)                                  // the discarded part must really be empty
    (full[r] || []).forEach(v => assert.strictEqual(String(v == null ? "" : v), "",
                                                    "row " + (r + 1) + " is not empty"));
  for (let r = top; r < full.length; r++) {
    const row = full[r] || [], out = [];
    for (let c = 0; c < left && c < row.length; c++)
      assert.strictEqual(String(row[c] == null ? "" : row[c]), "",
                         "column " + (c + 1) + " is not empty");
    for (let c = left; c < row.length; c++) out.push(row[c]);
    values.push(out);
  }
  return { top: top, left: left, values: values };
}

/** The fake workbook. `sheets` maps a sheet name to its values matrix (starting
    at A1), or to a { top, left, values } used range from usedRangeAt(). A name
    that is absent gives undefined (no such sheet) and a null value gives a sheet
    whose used range is undefined (an empty sheet). */
function fakeWorkbook(sheets) {
  return {
    getWorksheet: function (name) {
      if (!Object.prototype.hasOwnProperty.call(sheets, name)) return undefined;
      const entry = sheets[name];
      return {
        getUsedRange: function () {
          if (entry === null) return undefined;
          const box = Array.isArray(entry) ? { top: 0, left: 0, values: entry } : entry;
          return {
            getValues: function () { return box.values; },
            getRowIndex: function () { return box.top; },
            getColumnIndex: function () { return box.left; }
          };
        }
      };
    }
  };
}
const FULL_BOOK = {
  "Dashboard Config": CONFIG, "Dashboard Alerts": ALERTS,
  "Production": PRODUCTION, "Dashboard Log": LOG
};

/* ---------- runner -------------------------------------------------------- */

let passed = 0, failed = 0;
function t(name, fn) {
  try { fn(); passed++; console.log("  ok    " + name); }
  catch (e) { failed++; console.log("  FAIL  " + name + "\n        " + e.message); }
}
const digestFor = (list, email) => list.filter(d => d.email === email)[0];

/* ---------- the tests ----------------------------------------------------- */

console.log("digest core");

t("sections come out with the same names and order as parser.js", () => {
  const B = core.blocksFromValues(PRODUCTION);
  assert.deepStrictEqual(B.names, ["Can sell as second hand", "Ready, customer won't take",
                                   "Collect & supply only", "Ready to fit", "In production"]);
  const inProd = B.blocks.filter(b => b.name === "In production")[0];
  assert.deepStrictEqual(inProd.jobs.map(j => j.id), ["E5000", "F6000", "G7000", "H8000"]);
});

t("a job whose comment says 'not sent to floor' does not split the sheet", () => {
  const B = core.blocksFromValues(PRODUCTION);
  assert.strictEqual(B.names.length, 5);
  assert.ok(core.productionJobRows(PRODUCTION)["H8000"] !== undefined);
});

t("only jobs in 'In production' are live; everything else is finished", () => {
  const rows = core.productionJobRows(PRODUCTION);
  ["A1000", "B2000", "C3000", "D4000", "Z9999"].forEach(id =>
    assert.strictEqual(rows[id], undefined, id + " should be finished"));
  assert.strictEqual(rows["E5000"], 12);
});

t("the comment columns are found by header text, past column K", () => {
  const cols = core.commentColumns(PRODUCTION);
  assert.deepStrictEqual(cols, { comment: C_COMMENT, brendan: C_BRENDAN, specials: C_SPECIALS });
  assert.deepStrictEqual(core.headerRowsFromValues(PRODUCTION), [1, 2]);
});

t("config keys are read loosely and are the only source of admin / url", () => {
  const c = core.readConfig(CONFIG);
  assert.strictEqual(c.admin, ADMIN);
  assert.strictEqual(c.dashboardUrl, URL);
  assert.deepStrictEqual(core.readConfig([]), { admin: "", dashboardUrl: "" });
});

t("subscriptions: header and blanked rows fall out, duplicates count once", () => {
  const subs = core.readSubscriptions(ALERTS);
  assert.strictEqual(subs.length, 8);
  assert.ok(subs.every(s => s.email.indexOf("@") > 0));
  assert.strictEqual(subs.filter(s => s.job === "E5000" && s.email === SOMEONE).length, 1);
});

t("log comments: only column D = Comment, blanks dropped, oldest first", () => {
  const byJob = core.logComments(LOG);
  assert.deepStrictEqual(byJob["E5000"], ["Frames delivered", "Glass ordered"]);
  assert.strictEqual(byJob["F6000"], undefined);
});

console.log("\ndigests");

const D = core.buildDigests(CONFIG, ALERTS, PRODUCTION, LOG, NOW);

t("one digest per address, wrong domain skipped, empty address omitted", () => {
  assert.deepStrictEqual(D.map(d => d.email), [ANOTHER, SOMEONE]);
  assert.strictEqual(digestFor(D, OUTSIDER), undefined, "outside the admin's domain");
  assert.strictEqual(digestFor(D, LONELY), undefined, "nothing unfinished to send");
});

t("comments come from all four sources, in sheet-then-log order", () => {
  const html = digestFor(D, SOMEONE).html;
  const items = html.slice(html.indexOf("E5000")).split("<li>").slice(1, 6)
                    .map(s => s.slice(0, s.indexOf("</li>")));
  assert.deepStrictEqual(items, ["Waiting on &lt;glass&gt; &amp; trim &quot;urgent&quot;",
                                 "Chase the supplier", "Arch head",
                                 "Frames delivered", "Glass ordered"]);
});

t("comment text is HTML-escaped", () => {
  const html = digestFor(D, SOMEONE).html;
  assert.ok(html.indexOf("<glass>") < 0, "raw angle brackets got through");
  assert.ok(html.indexOf("&lt;glass&gt; &amp; trim &quot;urgent&quot;") > 0);
});

t("a job with no comment anywhere still appears, as (no comments)", () => {
  const d = digestFor(D, ANOTHER);
  assert.ok(d.html.indexOf("F6000") > 0);
  assert.ok(d.html.indexOf("(no comments)") > 0);
});

t("subject counts the jobs and dates the run", () => {
  assert.strictEqual(digestFor(D, SOMEONE).subject, "Job alerts: 2 jobs - Sat 5 Sep");
  assert.strictEqual(digestFor(D, ANOTHER).subject, "Job alerts: 1 job - Sat 5 Sep");
  assert.strictEqual(core.subjectFor(11, new Date(2026, 0, 1)), "Job alerts: 11 jobs - Thu 1 Jan");
});

t("finished jobs are not listed", () => {
  const html = digestFor(D, SOMEONE).html;
  ["A1000", "D4000", "Z9999"].forEach(id =>
    assert.ok(html.indexOf(id) < 0, id + " should not be in the mail"));
  assert.ok(html.indexOf("E5000") > 0 && html.indexOf("H8000") > 0);
});

t("the footer links to the dashboard URL from the config sheet", () => {
  assert.ok(digestFor(D, SOMEONE).html.indexOf('<a href="' + URL + '">') > 0);
  const noUrl = core.buildDigests([["admin", ADMIN]], ALERTS, PRODUCTION, LOG, NOW);
  assert.ok(noUrl[0].html.indexOf("<a href") < 0, "no URL configured, so no link");
  const bad = core.buildDigests([["admin", ADMIN], ["dashboardUrl", "javascript:alert(1)"]],
                                ALERTS, PRODUCTION, LOG, NOW);
  assert.ok(bad[0].html.indexOf("javascript:") < 0, "only http(s) links are used");
});

t("no admin address means no domain rule and therefore no mail", () => {
  assert.deepStrictEqual(core.buildDigests([["dashboardUrl", URL]], ALERTS, PRODUCTION, LOG, NOW), []);
  assert.deepStrictEqual(core.buildDigests([["admin", "not-an-address"]], ALERTS, PRODUCTION, LOG, NOW), []);
});

t("missing or empty sheets are handled without throwing", () => {
  assert.deepStrictEqual(core.buildDigests([], [], [], [], NOW), []);
  assert.deepStrictEqual(core.buildDigests(CONFIG, ALERTS, [], [], NOW), []);   // no Production -> nothing live
  const noLog = core.buildDigests(CONFIG, ALERTS, PRODUCTION, [], NOW);
  assert.strictEqual(noLog.length, 2);
  assert.ok(digestFor(noLog, SOMEONE).html.indexOf("Frames delivered") < 0);
});

t("junk rows are skipped, never thrown on", () => {
  const junk = [[null, undefined, 12345], ["", ""], [{}, [], "x"]];
  assert.doesNotThrow(() => core.buildDigests(junk, junk, junk, junk, NOW));
  assert.doesNotThrow(() => core.buildDigests(CONFIG, junk, junk, junk, NOW));
});

console.log("\ngenerated Office Script (automation/alerts-digest.ts)");

t("the .ts still contains the current core, verbatim", () => {
  const src = fs.readFileSync(path.join(DIR, "digest-core.js"), "utf8");
  const a = src.indexOf("/* --- BEGIN SHARED CORE"), b = src.indexOf("/* --- END SHARED CORE --- */");
  const block = src.slice(src.indexOf("*/", a) + 2, b).trim()
                   .replace(/\s*\/\*:([^*]+)\*\//g, (_, ty) => ": " + ty.trim());
  assert.ok(TS.indexOf(block) > 0, "alerts-digest.ts is stale - run node automation/build-script.js");
});

t("the .ts uses nothing Office Scripts rejects", () => {
  [/\brequire\s*\(/, /\bmodule\s*\./, /^\s*import\s/m, /^\s*export\s/m,
   /\bprocess\s*\./, /\b__dirname\b/, /\bObject\.fromEntries\b/, /\bvar\s/].forEach(re =>
    assert.ok(!re.test(TS), "generated script contains " + re));
  assert.ok(/function main\(workbook: ExcelScript\.Workbook\): string \{/.test(TS));
});

t("no address, name or domain is hard-coded in the shipped files", () => {
  ["digest-core.js", "alerts-digest.ts", "build-script.js", "SETUP.md"].forEach(f => {
    const text = fs.readFileSync(path.join(DIR, f), "utf8");
    const hits = text.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
    assert.deepStrictEqual(hits, [], f + " contains an address: " + hits.join(", "));
  });
});

t("main() over a fake workbook returns the same digests as the core", () => {
  const out = SCRIPT_MAIN(fakeWorkbook(FULL_BOOK));
  assert.strictEqual(typeof out, "string");
  const parsed = JSON.parse(out);
  /* the script stamps the mail with its own run date, which the fixtures cannot
     know, so the subject is checked by shape and everything else by value */
  const strip = list => list.map(d => ({ email: d.email, html: d.html }));
  assert.deepStrictEqual(strip(parsed), strip(D));
  parsed.forEach(d => assert.ok(/^Job alerts: \d+ jobs? - \w{3} \d{1,2} \w{3}$/.test(d.subject),
                                "odd subject: " + d.subject));
  assert.deepStrictEqual(parsed.map(d => d.email), [ANOTHER, SOMEONE]);
  assert.ok(parsed[1].html.indexOf("&lt;glass&gt;") > 0);
});

t("main() returns [] when the sheets are missing", () => {
  assert.strictEqual(SCRIPT_MAIN(fakeWorkbook({})), "[]");
  assert.strictEqual(SCRIPT_MAIN(fakeWorkbook({ "Production": PRODUCTION })), "[]");
});

t("main() survives empty sheets (a used range that does not exist)", () => {
  assert.strictEqual(SCRIPT_MAIN(fakeWorkbook({
    "Dashboard Config": null, "Dashboard Alerts": null,
    "Production": null, "Dashboard Log": null
  })), "[]");
  const out = JSON.parse(SCRIPT_MAIN(fakeWorkbook({
    "Dashboard Config": CONFIG, "Dashboard Alerts": ALERTS,
    "Production": PRODUCTION, "Dashboard Log": null
  })));
  assert.strictEqual(out.length, 2);
});

t("main() reads each sheet's used range exactly once", () => {
  const sheets = [], ranges = [], reads = [];
  const book = {
    getWorksheet: function (name) {
      sheets.push(name);
      const values = FULL_BOOK[name];
      if (!values) return undefined;
      return {
        getUsedRange: function () {
          ranges.push(name);
          return {
            getValues: function () { reads.push(name); return values; },
            getRowIndex: function () { return 0; },
            getColumnIndex: function () { return 0; }
          };
        }
      };
    }
  };
  SCRIPT_MAIN(book);
  const order = ["Dashboard Config", "Dashboard Alerts", "Production", "Dashboard Log"];
  assert.deepStrictEqual(sheets, order);
  assert.deepStrictEqual(ranges, order);
  assert.deepStrictEqual(reads, order);
});

t("main() pads a used range that does not start at A1 (Production at C3)", () => {
  const out = JSON.parse(SCRIPT_MAIN(fakeWorkbook(OFFSET_BOOK)));
  assert.strictEqual(out.length, 1);
  const d = out[0];
  assert.strictEqual(d.email, SOMEONE, "config keys were found in the padded matrix");
  assert.ok(/^Job alerts: 2 jobs - /.test(d.subject), "odd subject: " + d.subject);
  /* sections: the second "not sent to floor" divider still opens In production,
     so D4000 above it is finished and the two below it are not */
  assert.ok(d.html.indexOf("D4000") < 0, "a finished job got in");
  assert.ok(d.html.indexOf("E5000") > 0 && d.html.indexOf("F6000") > 0);
  /* comment columns: all three Production columns plus the log, still found by
     header text on the padded row 3 */
  assert.ok(d.html.indexOf("Waiting on &lt;glass&gt;") > 0);
  assert.ok(d.html.indexOf("Chase the supplier") > 0);
  assert.ok(d.html.indexOf("Arch head") > 0);
  assert.ok(d.html.indexOf("Glass ordered") > 0);
  assert.ok(d.html.indexOf("(no comments)") > 0, "F6000 should still say (no comments)");
  /* the config sheet was padded too: the dashboard URL still came through */
  assert.ok(d.html.indexOf('<a href="' + URL + '">') > 0);
});

t("without the padding the offset workbook would read as empty", () => {
  /* proves the padding is what does the work above, not the fixtures */
  const raw = name => OFFSET_BOOK[name].values;
  assert.deepStrictEqual(
    core.buildDigests(raw("Dashboard Config"), raw("Dashboard Alerts"),
                      raw("Production"), raw("Dashboard Log"), NOW), []);
});

console.log("\n" + passed + " passed, " + failed + " failed");
process.exit(failed ? 1 : 0);
