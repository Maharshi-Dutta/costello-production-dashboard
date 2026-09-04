/* Parses the Costello production workbook into the job model the dashboard uses.
   Runs unchanged in Node (for testing) and in the browser (via ExcelJS UMD).
   Mirrors the Python reader exactly: per-sheet header detection, divider-driven
   categories, and fill-based status. */

const JOB_RE = /^[A-Z]{1,2}\d{3,5}$/;
const URG_RE = /urgent|asap|a\.s\.a\.p/i;

const SHEETS = ['Production', 'Production (2)', 'PA Lam', 'Glass x 02', 'PVC Windows x 2',
  'Glazing', 'Cut & Weld PVC', 'PVC Doors', 'Smart slides+Dave Mc', 'THWS', 'Office call logs'];
const LABEL = {
  'Production': 'Production', 'Production (2)': 'Production (2)', 'PA Lam': 'PA Lam',
  'Glass x 02': 'Glass', 'PVC Windows x 2': 'Wds Prep', 'Glazing': 'Glazing',
  'Cut & Weld PVC': 'Cut & Weld', 'PVC Doors': 'PVC Doors',
  'Smart slides+Dave Mc': 'Smart Slides', 'THWS': 'THWS', 'Office call logs': 'Call Log'
};
const DATES = { 'sold': 'sold', 'stamp': 'stamp', 'ivana': 'ivana', 'ready to print': 'ready', 'sent to floor': 'floor' };
const IDENT = { 'comment': 'cm', 'office no': 'off', 'customer': 'cust', 'phone no': 'phone', 'area': 'area', 'eircode': 'eir', 'windows colour': 'colour' };
const GLASSC = ['dg', 'tg', 'tuff', 'not tuff', 'arch', 'astragal', 'fancy', 'extra'];

/* Excel's default theme palette, indexed as fills reference it.
   accent4 (index 7) at tint .6 is #FFE699 -- the "process done" gold. */
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

function applyTint(hex, tint) {
  if (!tint) return hex;
  const n = parseInt(hex, 16);
  let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  const f = c => tint > 0 ? Math.round(c * (1 - tint) + 255 * tint) : Math.round(c * (1 + tint));
  r = f(r); g = f(g); b = f(b);
  return [r, g, b].map(x => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Normalised fill colour of a cell as RRGGBB, or '' when unfilled. */
function fillOf(cell) {
  const f = cell && cell.fill;
  if (!f || f.type !== 'pattern' || !f.pattern || f.pattern === 'none') return '';
  const fg = f.fgColor;
  if (!fg) return '';
  if (fg.argb) return String(fg.argb).slice(-6).toUpperCase();
  if (fg.theme !== undefined) {
    const base = THEME[fg.theme] || 'FFFFFF';
    return applyTint(base, fg.tint || 0);
  }
  return '';
}

const GOLD = new Set(['FFE699', 'FFC000']);
const YELLOW = new Set(['FFFF00']);

/* A job shows up on several sheets; when two of them disagree about a cell's
   colour the further-along one wins, because a colour is only ever added as
   work is finished - it is never taken back to mean "less done". */
const CPRANK = { '': 0, process: 1, done: 2 };
const cpOf = fill => YELLOW.has(fill) ? 'process' : (GOLD.has(fill) ? 'done' : '');
const cpBump = (o, k, st) => { if (CPRANK[st] > CPRANK[o[k] || '']) o[k] = st; };
const norm = v => String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

function cellText(cell) {
  const v = cell ? cell.value : null;
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    if (v.richText) return v.richText.map(t => t.text).join('');
    return '';
  }
  return String(v);
}
function cellDate(cell) {
  const v = cell ? cell.value : null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = cellText(cell).trim();
  return s ? s.slice(0, 16) : null;
}
function num(cell) {
  /* strict: the whole cell must be a number. Cells like "3x White / 1x Silver Grey"
     are colour notes, not quantities -- parseFloat would wrongly read them as 3. */
  const s = cellText(cell).trim();
  if (!s || !/^[-+]?\d*\.?\d+$/.test(s)) return 0;
  const f = parseFloat(s);
  return isFinite(f) ? (f % 1 === 0 ? f : Math.round(f * 10) / 10) : 0;
}

/** Find the (group,sub) header row pair by locating 'DATES ON CONTRACT'. */
function headerRows(ws) {
  for (let r = 1; r <= 6; r++) {
    for (let c = 1; c <= Math.min(ws.columnCount || 40, 40); c++) {
      const n = norm(cellText(ws.getRow(r).getCell(c)));
      if (n.indexOf('dates on contract') >= 0 || n === 'dates') return [r, r + 1];
    }
  }
  return [2, 3];
}

function mapSheet(ws) {
  const [g, s] = headerRows(ws);
  const m = { hdr: [g, s], ident: {}, dates: {}, qty: {}, prod: {}, prodOrder: [], glass: {}, notes: {} };
  let group = null;
  const maxC = ws.columnCount || 150;
  for (let c = 1; c <= maxC; c++) {
    const lab = norm(cellText(ws.getRow(g).getCell(c)));
    const sub = norm(cellText(ws.getRow(s).getCell(c)));
    if (lab) group = lab;
    for (const k in IDENT) if (lab === k && m.ident[IDENT[k]] === undefined) m.ident[IDENT[k]] = c;
    if (DATES[sub] && m.dates[DATES[sub]] === undefined) m.dates[DATES[sub]] = c;
    if (group && group.indexOf('quantity') >= 0) {
      if (sub === 'wnd') m.qty.wnd = c;
      else if (sub === 'drs') m.qty.drs = c;
    }
    if (group && group.indexOf('glass unit') >= 0 && GLASSC.indexOf(sub) >= 0) m.glass[sub] = c;
    if (lab.indexOf('brendan') >= 0) m.notes.brendan = c;
    if (lab.indexOf('notes of specials') >= 0) m.notes.specials = c;
    const skip = ['doors done', 'windows fabricated', 'doors fabricated', 'dates'];
    if (group && sub && group.indexOf('glass unit') < 0 && group.indexOf('quantity') < 0 &&
        !skip.some(k => group.indexOf(k) >= 0) && ['f', 's', 't'].indexOf(sub) >= 0) {
      if (!m.prod[group]) { m.prod[group] = {}; m.prodOrder.push(group); }
      m.prod[group][sub] = c;
    }
  }
  return m;
}

/** The Production sheet's own divider rows define its blocks. Everything here is
    derived live from the sheet, so adding or moving a divider changes the blocks
    automatically - nothing about the structure is stored anywhere. */
/** The Production sheet is organised by position: divider rows split it into
    sections. This reads that structure from a plain values matrix (rows of
    cell text, columns A..K at least) so the same rule serves the downloaded
    file and a live read through the Excel API.
    A divider is a row whose text matches AND whose job-number cell is not a
    job number - so a job whose comment happens to say "not sent to floor"
    can never be mistaken for a section break.                             */
function blocksFromValues(matrix) {
  const names = ["Can sell as second hand"];
  const blocks = [{ idx: 0, name: names[0], divider: 0, jobs: [], last: 0 }];
  let cur = blocks[0];
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i] || [], r = i + 1;
    const j = String(row[2] == null ? "" : row[2]).trim().toUpperCase();
    const isJob = JOB_RE.test(j);
    let line = "";
    for (let c = 0; c < 11; c++) line += " " + String(row[c] == null ? "" : row[c]);
    line = norm(line);
    let hit = null;
    if (!isJob) {
      if (line.indexOf("can sell as second hand") >= 0) hit = null;          // labels the block above
      else if (line.indexOf("customers won t take") >= 0 || line.indexOf("customers wont take") >= 0)
        hit = "Ready, customer won't take";
      else if (line.indexOf("collect or supply only") >= 0) hit = "Collect & supply only";
      else if (line.indexOf("not sent to floor") >= 0)
        /* display names only - the sheet's own divider text is never changed */
        hit = names.indexOf("Ready to fit") < 0 ? "Ready to fit" : "In production";
    }
    if (hit) { names.push(hit); cur = { idx: names.length - 1, name: hit, divider: r, jobs: [], last: r }; blocks.push(cur); continue; }
    if (isJob) { cur.jobs.push({ row: r, id: j }); cur.last = r; }
  }
  return { names, blocks };
}

function productionBlocks(ws) {
  const maxR = ws.rowCount || 0, matrix = [];
  for (let r = 1; r <= maxR; r++) {
    const row = ws.getRow(r), line = [];
    for (let c = 1; c <= 11; c++) line.push(cellText(row.getCell(c)));
    matrix.push(line);
  }
  const B = blocksFromValues(matrix);
  const cat = {}, blk = {}, order = {};
  let seen = 0;
  B.blocks.forEach(b => b.jobs.forEach(j => {
    blk[j.id] = b.idx; order[j.id] = seen++;
    cat[j.id] = b.idx === 0 ? "secondhand" : b.idx === 1 ? "wonttake" : b.idx === 2 ? "collect" : "active";
  }));
  return { cat, blk, names: B.names, order };
}

/** The formatting of a job's row that the Excel API cannot read cheaply per
    cell (alignment, wrap, borders) plus grouping keys, from the last download.
    Found by job number, never by row, because rows move. Plain data only.  */
function templateForJob(ws, jobId) {
  if (!ws) return null;
  const want = String(jobId || "").trim().toUpperCase();
  const maxR = ws.rowCount || 0;
  for (let r = 1; r <= maxR; r++) {
    const row = ws.getRow(r);
    if (cellText(row.getCell(3)).trim().toUpperCase() !== want) continue;
    const cells = [];
    for (let c = 1; c <= 90; c++) {
      const cell = row.getCell(c), a = cell.alignment || {}, b = cell.border || {};
      const f = cell.font || {}, fl = cell.fill || {};
      const ck = col => col ? (col.argb || ("t" + col.theme + "/" + (col.tint || 0))) : "";
      const side = s => (s && s.style) ? { style: s.style, color: (s.color && s.color.argb) ? "#" + s.color.argb.slice(-6) : "#000000" } : null;
      cells.push({ h: a.horizontal || null, v: a.vertical || null, wrap: !!a.wrapText,
                   bottom: side(b.bottom), left: side(b.left), right: side(b.right),
                   /* keys only, used to group cells that share one font / one fill so the live
                      read can fetch each group in a single request */
                   fk: [f.bold ? 1 : 0, f.italic ? 1 : 0, f.size || "", f.name || "", ck(f.color)].join("|"),
                   flk: fl.pattern === "solid" ? ck(fl.fgColor) : "" });
    }
    return { row: r, height: row.height || null, cells };
  }
  return null;
}

function parseWorkbook(wb) {
  const prodSheet = wb.getWorksheet('Production');
  const B = prodSheet ? productionBlocks(prodSheet) : { cat: {}, blk: {}, names: [], order: {} };
  const PRODCAT = B.cat;
  const jobs = {};

  for (const name of SHEETS) {
    const ws = wb.getWorksheet(name);
    if (!ws) continue;
    const m = mapSheet(ws);
    /* Checkpoint colours come from the Production sheet alone: it is the only
       sheet this dashboard writes, so taking the furthest-along colour across
       all the sheets would make un-ticking something snap straight back from
       whatever the copy on Production (2) or PA Lam still says. */
    const cpHere = name === 'Production';
    const maxC = Math.min(ws.columnCount || 150, 150);
    const maxR = ws.rowCount || 0;
    for (let r = 1; r <= maxR; r++) {
      const row = ws.getRow(r);
      const jid = cellText(row.getCell(3)).trim().toUpperCase();
      if (!JOB_RE.test(jid)) continue;

      let goldSpine = 0, goldAll = 0;
      for (let c = 1; c <= maxC; c++) {
        const f = fillOf(row.getCell(c));
        if (GOLD.has(f)) { goldAll++; if (c <= 14) goldSpine++; }
      }
      const rowDone = goldSpine >= 10 || goldAll >= 40;

      let j = jobs[jid];
      if (!j) { j = jobs[jid] = { id: jid, sheets: [], src: {}, prods: {}, glass: {}, status: {}, notes: [],
                                  cp: { win: '', drs: '', glass: {}, prod: {} } }; }
      j.src[LABEL[name]] = r;
      if (rowDone) j.done = 1;
      if (j.sheets.indexOf(LABEL[name]) < 0) j.sheets.push(LABEL[name]);

      for (const k in m.ident) { const v = cellText(row.getCell(m.ident[k])).trim(); if (v && !j[k]) j[k] = v.slice(0, 70); }
      for (const k in m.dates) { const v = cellDate(row.getCell(m.dates[k])); if (v && !j['d_' + k]) j['d_' + k] = v; }
      for (const k in m.qty) {
        const c = row.getCell(m.qty[k]), v = num(c);
        if (v && !j[k]) j[k] = v;
        if (cpHere) cpBump(j.cp, k === 'wnd' ? 'win' : 'drs', cpOf(fillOf(c)));
      }
      for (const k in m.glass) {
        const c = row.getCell(m.glass[k]), v = num(c);
        if (v) j.glass[k] = Math.max(j.glass[k] || 0, v);
        if (cpHere) cpBump(j.cp.glass, k, cpOf(fillOf(c)));
      }

      for (const pname of m.prodOrder) {
        const cols = m.prod[pname];
        /* not every product has all three sub-columns (Composite is S only) */
        const at = c => (c ? num(row.getCell(c)) : 0);
        const f = at(cols.f), s2 = at(cols.s), t = at(cols.t);
        if (f || s2 || t) {
          const cur = j.prods[pname] || [0, 0, 0];
          j.prods[pname] = [Math.max(cur[0], f), Math.max(cur[1], s2), Math.max(cur[2], t)];
        }
        for (const sub of ['f', 's', 't']) {
          if (!cols[sub]) continue;
          const fill = fillOf(row.getCell(cols[sub]));
          if (YELLOW.has(fill)) { (j.status[pname] = j.status[pname] || {})['process'] = 1; }
          else if (GOLD.has(fill) && !rowDone) { (j.status[pname] = j.status[pname] || {})['done'] = 1; }
          /* only make an entry when there is a colour, so cp.prod stays small */
          const cst = cpHere ? cpOf(fill) : '';
          if (cst) cpBump(j.cp.prod[pname] = j.cp.prod[pname] || {}, sub, cst);
        }
      }
      for (const k in m.notes) {
        const v = cellText(row.getCell(m.notes[k])).trim();
        if (v && !j.notes.some(n => n.t === v.slice(0, 140))) j.notes.push({ k, t: v.slice(0, 140), s: LABEL[name] });
      }
      if (m.ident.cm) {
        const v = cellText(row.getCell(m.ident.cm)).trim();
        if (v && !j.notes.some(n => n.t === v.slice(0, 140))) j.notes.push({ k: 'comment', t: v.slice(0, 140), s: LABEL[name] });
      }
    }
  }

  const result = Object.keys(jobs).sort().map(id => {
    const j = jobs[id];
    const cmtxt = j.notes.map(n => n.t).join(' ');
    const ph = String(j.phone || '').replace(/\D/g, '');
    const prods = Object.keys(j.prods).map(k => ({ n: k, f: j.prods[k][0], s: j.prods[k][1], t: j.prods[k][2], st: Object.keys(j.status[k] || {}) }));
    /* a whole gold row says the job is finished, so every checkpoint on it is */
    if (j.done) {
      if (j.wnd) j.cp.win = 'done';
      if (j.drs) j.cp.drs = 'done';
      Object.keys(j.glass).forEach(k => { j.cp.glass[k] = 'done'; });
      prods.forEach(p => {
        const o = j.cp.prod[p.n] = j.cp.prod[p.n] || {};
        ['f', 's', 't'].forEach(s => { if (p[s]) o[s] = 'done'; });
      });
    }
    return {
      id, cust: j.cust || '', area: j.area || '', eir: j.eir || '', off: j.off || '',
      colour: j.colour || '', ph3: ph.length >= 3 ? ph.slice(-3) : '',
      wnd: j.wnd || 0, drs: j.drs || 0,
      dates: { sold: j.d_sold || null, stamp: j.d_stamp || null, ivana: j.d_ivana || null, ready: j.d_ready || null, floor: j.d_floor || null },
      prods: prods, cp: j.cp,
      glass: j.glass, notes: j.notes, sheets: j.sheets, src: j.src,
      urg: URG_RE.test(cmtxt) ? 1 : 0, done: j.done || 0,
      cat: PRODCAT[id] || 'past',
      blk: (B.blk[id] === undefined ? -1 : B.blk[id]),
      seq: (B.order[id] === undefined ? 99999 : B.order[id]),
      stage: j.d_floor ? 'floor' : (j.d_ready ? 'ready' : 'office')
    };
  });
  result.blockNames = B.names;
  return result;
}

if (typeof module !== 'undefined') module.exports = { parseWorkbook, mapSheet, fillOf, JOB_RE, blocksFromValues, templateForJob };
if (typeof window !== 'undefined') { window.parseWorkbook = parseWorkbook; window.mapSheet = mapSheet;
  window.blocksFromValues = blocksFromValues; window.templateForJob = templateForJob; }
