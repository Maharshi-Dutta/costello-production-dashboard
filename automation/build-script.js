/* build-script.js - generates automation/alerts-digest.ts from digest-core.js.

   Office Scripts cannot import anything: the whole script has to be one file.
   So the shared logic lives in digest-core.js (which Node can require and the
   test can exercise) and this generator inlines it into a TypeScript file that
   can be pasted straight into Excel on the web -> Automate -> New script.

   What it does, and nothing more:
     1. takes the block between the CORE markers in digest-core.js;
     2. wraps it in a header (banner + type declarations) and a footer
        (main() and the one function that touches ExcelScript);
     3. turns the  /* :Type * /  comments into real TypeScript annotations;
     4. refuses to write a file containing anything Office Scripts rejects
        (require, module, import/export, Node globals, Object.fromEntries).

   Run:  node automation/build-script.js
   Never edit alerts-digest.ts by hand - it is overwritten by this script.   */

const fs = require("fs");
const path = require("path");

const DIR = __dirname;
const SRC = path.join(DIR, "digest-core.js");
const OUT = path.join(DIR, "alerts-digest.ts");

const BEGIN = "/* --- BEGIN SHARED CORE";
const END = "/* --- END SHARED CORE --- */";

/* Types declared in the generated header. Everything else the annotations
   mention is collected from the source below, so the list the test uses to
   strip the annotations again can never drift out of date. */
const DECLARED = ["Cell", "JobRef", "Block", "Conf", "Cols", "Sub", "Digest", "JobDigest"];

const TYPES_BLOCK = [
  "/* @types-begin */",
  "type Cell = string | number | boolean;",
  "interface JobRef { row: number; id: string; }",
  "interface Block { idx: number; name: string; divider: number; jobs: JobRef[]; last: number; }",
  "interface Conf { admin: string; dashboardUrl: string; }",
  "interface Cols { comment: number; brendan: number; specials: number; }",
  "interface Sub { job: string; email: string; who: string; when: string; }",
  "interface JobDigest { id: string; comments: string[]; }",
  "interface Digest { email: string; subject: string; html: string; }",
  "/* @types-end */"
].join("\n");

const BANNER = [
  "/* Dashboard alerts digest - an Office Script for Excel on the web.",
  "",
  "   GENERATED FILE - DO NOT EDIT HERE.",
  "     source     automation/digest-core.js",
  "     generator  automation/build-script.js",
  "     regenerate node automation/build-script.js",
  "",
  "   It reads four sheets, works out which subscribed jobs are still in",
  "   production and what has been said about them, and returns the mails to",
  "   send as JSON. It sends nothing itself and writes nothing to the",
  "   workbook - the Power Automate flow does the sending (see SETUP.md).",
  "",
  "   No address, name or domain is written into this file: the admin address",
  "   (which decides the only allowed recipient domain) and the dashboard URL",
  "   are read from the Dashboard Config sheet every run.",
  "",
  "   Returns: JSON text, an array of { email, subject, html }.",
  "*/"
].join("\n");

const FOOTER = [
  "/* ---- the only part that talks to Excel ---------------------------------",
  "   Each sheet is read once, whole: getUsedRange().getValues(). Per-cell reads",
  "   through the API would be hundreds of round trips and would time the script",
  "   out. A sheet that does not exist yet (or is empty) comes back as an empty",
  "   matrix, and the core logic treats that as \"nothing to say\".",
  "",
  "   The used range begins at the first used cell, which is not necessarily A1,",
  "   so the block is padded back out with empty rows in front and empty cells on",
  "   the left of every row. Every index the core then sees is a true sheet",
  "   coordinate - row 0 is row 1, column 0 is column A. Without this, a sheet",
  "   with a blank column A would shift every job number out of column C and the",
  "   whole run would silently come back empty. The padding lives here and not in",
  "   digest-core.js so the core stays pure - it only ever sees a plain matrix",
  "   addressed from A1.                                                       */",
  "function readSheet(workbook /*:ExcelScript.Workbook*/, name /*:string*/) /*:Cell[][]*/ {",
  "  const sheet = workbook.getWorksheet(name);",
  "  if (!sheet) return [];                     // getWorksheet returns undefined when there is no such sheet",
  "  const used = sheet.getUsedRange();",
  "  if (!used) return [];                      // an entirely empty sheet has no used range",
  "  const values = used.getValues();",
  "  if (!values) return [];",
  "  const top = used.getRowIndex() || 0;       // 0-based, so row 1 is 0",
  "  const left = used.getColumnIndex() || 0;   // 0-based, so column A is 0",
  "  const out /*:Cell[][]*/ = [];",
  "  for (let r = 0; r < top; r++) out.push([]);",
  "  for (let r = 0; r < values.length; r++) {",
  "    const row /*:Cell[]*/ = [];",
  "    for (let c = 0; c < left; c++) row.push(\"\");",
  "    const src = values[r] || [];",
  "    for (let c = 0; c < src.length; c++) row.push(src[c]);",
  "    out.push(row);",
  "  }",
  "  return out;",
  "}",
  "",
  "function main(workbook /*:ExcelScript.Workbook*/) /*:string*/ {",
  "  const config = readSheet(workbook, \"Dashboard Config\");",
  "  const alerts = readSheet(workbook, \"Dashboard Alerts\");",
  "  const production = readSheet(workbook, \"Production\");",
  "  const log = readSheet(workbook, \"Dashboard Log\");",
  "  const digests = buildDigests(config, alerts, production, log, new Date());",
  "  return JSON.stringify(digests);",
  "}"
].join("\n");

/* Things the Office Scripts compiler or runtime rejects. */
const FORBIDDEN = [
  [/\brequire\s*\(/, "require()"],
  [/\bmodule\s*\./, "module."],
  [/^\s*import\s/m, "import"],
  [/^\s*export\s/m, "export"],
  [/\bprocess\s*\./, "process."],
  [/\b__dirname\b/, "__dirname"],
  [/\bObject\.fromEntries\b/, "Object.fromEntries"],
  [/\bvar\s/, "var"]
];

function fail(msg) { console.error("build-script: " + msg); process.exit(1); }

const src = fs.readFileSync(SRC, "utf8");
const a = src.indexOf(BEGIN), b = src.indexOf(END);
if (a < 0 || b < 0) fail("could not find the CORE markers in digest-core.js");
const core = src.slice(src.indexOf("*/", a) + 2, b).trim();
if (!/function buildDigests/.test(core)) fail("the core block does not contain buildDigests");

let out = [BANNER, "", TYPES_BLOCK, "", FOOTER, "",
           "/* ---- shared core, copied verbatim from automation/digest-core.js ---- */",
           "", core, ""].join("\n");

/* Collect every type name the annotations use, so the test can strip exactly
   these again, then turn the annotations into real TypeScript. */
const used = {};
DECLARED.forEach(function (t) { used[t] = 1; });
out.replace(/\/\*:([^*]+)\*\//g, function (_, t) {
  t.replace(/[A-Za-z_$][\w$.]*/g, function (id) { used[id] = 1; return id; });
  return _;
});
out = out.replace(/\s*\/\*:([^*]+)\*\//g, function (_, t) { return ": " + t.trim(); });

if (out.indexOf("/*:") >= 0) fail("an annotation comment survived the rewrite");
FORBIDDEN.forEach(function (p) { if (p[0].test(out)) fail("generated file contains " + p[1]); });
if (!/function main\(workbook: ExcelScript\.Workbook\): string \{/.test(out))
  fail("main() did not come out with its types");

const names = Object.keys(used).sort();
out = out.replace("/* @types-begin */", "/* @types: " + names.join("|") + " */\n/* @types-begin */");

fs.writeFileSync(OUT, out, "utf8");
console.log("build-script: wrote " + path.relative(process.cwd(), OUT).replace(/\\/g, "/") +
            "  (" + out.split("\n").length + " lines, " + out.length + " bytes)");
console.log("build-script: types " + names.join(" "));
console.log("build-script: checks passed - no require/module/import/export/process/var, no Object.fromEntries");
