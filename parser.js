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
function productionBlocks(ws) {
  const cat = {}, blk = {}, names = [], order = {};
  let idx = 0, seen = 0;
  names[0] = "Can sell as second hand";
  const maxR = ws.rowCount || 0;
  for (let r = 1; r <= maxR; r++) {
    const row = ws.getRow(r);
    const j = cellText(row.getCell(3)).trim().toUpperCase();
    let line = "";
    for (let c = 1; c <= 11; c++) line += " " + cellText(row.getCell(c));
    line = norm(line);
    let hit = null;
    if (line.indexOf("can sell as second hand") >= 0) hit = null;          // labels the block above
    else if (line.indexOf("customers won t take") >= 0 || line.indexOf("customers wont take") >= 0)
      hit = "Ready, customer won't take";
    else if (line.indexOf("collect or supply only") >= 0) hit = "Collect & supply only";
    else if (line.indexOf("not sent to floor") >= 0)
      hit = names.indexOf("Not sent to floor") < 0 ? "Not sent to floor" : "In production";
    if (hit) { idx++; names[idx] = hit; continue; }
    if (JOB_RE.test(j)) { blk[j] = idx; order[j] = seen++;
      cat[j] = idx === 0 ? "secondhand" : idx === 1 ? "wonttake" : idx === 2 ? "collect" : "active"; }
  }
  return { cat, blk, names, order };
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
      if (!j) { j = jobs[jid] = { id: jid, sheets: [], src: {}, prods: {}, glass: {}, status: {}, notes: [] }; }
      j.src[LABEL[name]] = r;
      if (rowDone) j.done = 1;
      if (j.sheets.indexOf(LABEL[name]) < 0) j.sheets.push(LABEL[name]);

      for (const k in m.ident) { const v = cellText(row.getCell(m.ident[k])).trim(); if (v && !j[k]) j[k] = v.slice(0, 70); }
      for (const k in m.dates) { const v = cellDate(row.getCell(m.dates[k])); if (v && !j['d_' + k]) j['d_' + k] = v; }
      for (const k in m.qty) { const v = num(row.getCell(m.qty[k])); if (v && !j[k]) j[k] = v; }
      for (const k in m.glass) { const v = num(row.getCell(m.glass[k])); if (v) j.glass[k] = Math.max(j.glass[k] || 0, v); }

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
    return {
      id, cust: j.cust || '', area: j.area || '', eir: j.eir || '', off: j.off || '',
      colour: j.colour || '', ph3: ph.length >= 3 ? ph.slice(-3) : '',
      wnd: j.wnd || 0, drs: j.drs || 0,
      dates: { sold: j.d_sold || null, stamp: j.d_stamp || null, ivana: j.d_ivana || null, ready: j.d_ready || null, floor: j.d_floor || null },
      prods: Object.keys(j.prods).map(k => ({ n: k, f: j.prods[k][0], s: j.prods[k][1], t: j.prods[k][2], st: Object.keys(j.status[k] || {}) })),
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

if (typeof module !== 'undefined') module.exports = { parseWorkbook, mapSheet, fillOf, JOB_RE };
if (typeof window !== 'undefined') { window.parseWorkbook = parseWorkbook; window.mapSheet = mapSheet; }
