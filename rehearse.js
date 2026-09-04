/* Rehearsal harness: runs the dashboard's real graph.js + parser.js in Node
   against a chosen workbook (normally the OneDrive rehearsal copy), with a
   token supplied from the scratch Python helper instead of the sign-in popup.

   node rehearse.js <ref.json> <token> move <JOB> <sectionIdx>
   node rehearse.js <ref.json> <token> blocks
   node rehearse.js <ref.json> <token> capture <JOB>
   node rehearse.js <ref.json> <token> failmove <JOB> <sectionIdx>   (write is made to fail: sheet must be unchanged)
*/
const fs = require("fs"), vm = require("vm");
const [,, refPath, tok, cmd, ...args] = process.argv;
const ref = JSON.parse(fs.readFileSync(refPath, "utf8"));

global.ExcelJS = require("exceljs");
const mem = {};
global.localStorage = { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } };
global.window = { location: { origin: "http://localhost" } };
vm.runInThisContext(fs.readFileSync(__dirname + "/parser.js", "utf8"), { filename: "parser.js" });
vm.runInThisContext(fs.readFileSync(__dirname + "/graph.js", "utf8"), { filename: "graph.js" });
const CW = window.CW;
CW._setToken(() => tok);
CW._setFile({ base: ref.base, content: ref.content, meta: ref.meta || ref.content.replace(/\/content$/, "") });

(async () => {
  const t0 = Date.now();
  await CW.openSession();
  if (cmd === "blocks") {
    const B = await CW.liveBlocks();
    B.blocks.forEach(b => console.log(b.idx, b.name.padEnd(28), "divider", b.divider, "jobs", b.jobs.length, "last", b.last,
      b.jobs.length ? b.jobs[0].id + ".." + b.jobs[b.jobs.length - 1].id : ""));
  } else if (cmd === "capture") {
    const L = await CW.locateJob(args[0].toUpperCase());
    const wb = await CW.downloadWorkbook();
    const tmpl = templateForJob(wb.getWorksheet("Production"), args[0]);
    const cap = await CW.captureRow(L.hit.row, tmpl);
    console.log("reads:", cap.reads, "fallback cells:", cap.fallback);
    console.log(JSON.stringify({ row: L.hit.row, height: cap.height, values: cap.values.slice(0, 12), types: cap.types.slice(0, 12),
      fills: cap.fills.slice(0, 16), fonts: cap.fonts.slice(0, 4) }, null, 1));
  } else if (cmd === "move" || cmd === "failmove") {
    const wb = await CW.downloadWorkbook();
    const tmpl = templateForJob(wb.getWorksheet("Production"), args[0]);
    console.log("template from download: row", tmpl && tmpl.row, "height", tmpl && tmpl.height);
    if (cmd === "failmove") {
      /* sabotage the values write once, to prove the sheet is put back */
      const origFetch = global.fetch; let armed = true;
      global.fetch = async (url, init) => {
        if (armed && init && init.method === "PATCH" && /range\(address='A\d+:CL\d+'\)$/.test(decodeURIComponent(String(url)))) {
          armed = false; throw new Error("SABOTAGE: simulated network failure during the row write");
        }
        return origFetch(url, init);
      };
    }
    try {
      const tmplFor = x => templateForJob(wb.getWorksheet("Production"), x);
      const r = await CW.moveJobRow(args[0], Number(args[1]), tmplFor, t => console.log("  step:", t));
      console.log("RESULT", JSON.stringify(r), "in", Date.now() - t0, "ms");
    } catch (e) { console.log("FAILED (expected for failmove):", e.message, "in", Date.now() - t0, "ms"); }
  }
})().catch(e => { console.error("ERROR", e); process.exit(1); });
