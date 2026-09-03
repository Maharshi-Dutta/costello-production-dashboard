const ExcelJS = require('exceljs');
const fs = require('fs');
const { parseWorkbook } = require('./parser.js');

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:/Costello Windows/Excel Dashboard/live.xlsx');
  const jobs = parseWorkbook(wb);
  const py = JSON.parse(fs.readFileSync('d:/Costello Windows/Excel Dashboard/design-canvas/jobs.json', 'utf8'));

  const count = (arr, f) => arr.filter(f).length;
  const cats = a => { const o = {}; a.forEach(j => o[j.cat] = (o[j.cat] || 0) + 1); return o; };
  console.log('                      JS (browser)      PYTHON (verified)');
  const rows = [
    ['total jobs',        jobs.length,                    py.length],
    ['on Production',     count(jobs, j => j.cat !== 'past'), count(py, j => j.cat !== 'past')],
    ['gold / done',       count(jobs, j => j.done),        count(py, j => j.done)],
    ['urgent',            count(jobs, j => j.urg),         count(py, j => j.urg)],
    ['with notes',        count(jobs, j => j.notes.length), count(py, j => j.notes.length)],
    ['in fabrication',    count(jobs, j => j.prods.some(p => p.st.indexOf('process') >= 0)),
                          count(py,   j => j.prods.some(p => (p.st||[]).indexOf('process') >= 0))],
    ['total windows',     jobs.reduce((a, j) => a + j.wnd, 0), py.reduce((a, j) => a + j.wnd, 0)],
    ['total doors',       jobs.reduce((a, j) => a + j.drs, 0), py.reduce((a, j) => a + j.drs, 0)],
  ];
  let allMatch = true;
  for (const [k, a, b] of rows) {
    const ok = a === b; if (!ok) allMatch = false;
    console.log('  ' + k.padEnd(20) + String(a).padStart(8) + String(b).padStart(18) + '   ' + (ok ? 'match' : ' <-- DIFFERS'));
  }
  console.log('\n  categories JS  :', JSON.stringify(cats(jobs)));
  console.log('  categories PY  :', JSON.stringify(cats(py)));

  // per-job deep check
  const pym = {}; py.forEach(j => pym[j.id] = j);
  let diff = 0, shown = 0;
  for (const j of jobs) {
    const p = pym[j.id]; if (!p) { diff++; continue; }
    if (j.done !== p.done || j.cat !== p.cat || j.wnd !== p.wnd || j.drs !== p.drs) {
      diff++; if (shown++ < 6) console.log('   DIFF', j.id, JSON.stringify({js:{d:j.done,c:j.cat}, py:{d:p.done,c:p.cat}}));
    }
  }
  console.log('\n  per-job mismatches:', diff, 'of', jobs.length);
  console.log(allMatch && diff === 0 ? '\n  ==> PARSERS AGREE EXACTLY' : '\n  ==> DIFFERENCES FOUND');

  const c5255 = jobs.find(j => j.id === 'C5255');
  console.log('\n  C5255 :', JSON.stringify({ done: c5255.done, cat: c5255.cat, cust: c5255.cust }));
})();
