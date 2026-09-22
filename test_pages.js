/* Every page's script chain parses as ONE script.

   The pages load classic <script src> tags, and classic scripts share one
   global lexical scope: a top-level `const X` in graph.js and another
   `const X` in a page's own file is a SyntaxError at load time, the second file
   never runs, and the page sits on its sign-in gate doing nothing (the glazing
   tablet, 2026-09-22: `Identifier 'G' has already been declared`). `node --check`
   per file cannot see it, and the browser rig stubs graph.js away. So: read each
   page's <script src> list, concatenate the repo's own files in that order, and
   parse the result once. Parse only - nothing runs, no DOM, no network.

   Run: node test_pages.js                                                     */
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const HERE = __dirname;
const PAGES = fs.readdirSync(HERE).filter(f => /\.html$/i.test(f));
let pass = 0, fail = 0;

for (const page of PAGES) {
  const html = fs.readFileSync(path.join(HERE, page), "utf8");
  const srcs = [];
  html.replace(/<script[^>]*\ssrc=["']([^"']+)["']/gi, (m, s) => { srcs.push(s.replace(/[?#].*$/, "")); return m; });
  const own = srcs.filter(s => !/^https?:/i.test(s) && !/^vendor\//i.test(s));
  if (!own.length) continue;
  const joined = own.map(s => fs.readFileSync(path.join(HERE, s), "utf8")).join("\n;\n");
  try {
    new vm.Script(joined, { filename: page + " (" + own.join(" + ") + ")" });
    pass++;
  } catch (e) {
    fail++;
    console.log("FAIL " + page + ": " + e.message + "  [" + own.join(" + ") + "]");
  }
}
console.log("pages: " + pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
