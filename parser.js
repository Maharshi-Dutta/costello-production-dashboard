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

/** Normalised font colour of a cell as RRGGBB, or '' when the cell says nothing
    about its own colour (no font, no colour, or a colour this reader cannot
    resolve). Theme colours go through the same palette and tint the fills use.

    Two colours are deliberately not read at all. A hyperlink is drawn blue and
    underlined by Excel itself, and theme 10 and 11 ARE those two hyperlink
    colours - so a cell carrying a link, or painted in either of them, says
    nothing about the job and would otherwise arrive here as "on hold". */
function fontOf(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (cell.hyperlink || (v && typeof v === 'object' && v.hyperlink)) return '';
  const f = cell.font;
  const col = f && f.color;
  if (!col) return '';
  if (col.argb) return String(col.argb).slice(-6).toUpperCase();
  if (col.theme !== undefined) {
    if (col.theme === 10 || col.theme === 11) return '';       // hlink / followed hlink
    return applyTint(THEME[col.theme] || '000000', col.tint || 0);
  }
  return '';                                  // indexed / unknown: say nothing rather than guess
}

/* ---- the row colour code --------------------------------------------------
   The Production sheet says four things with the colour of a job row's TEXT,
   and says them nowhere else on the sheet:

     word        colour the office uses      hue range read here
     ----------  --------------------------  -------------------
     urgent      red (FF0000, C00000, …)     345-360 and 0-20
     booked      green (00B050, 92D050, …)   75-170        (an exact delivery date is agreed)
     trade       pink / magenta (FF3399,     275-345
                 FF00FF, …)
     hold        blue (00B0F0, 0070C0, …)    170-260

   Ranges, not exact hex, because the colours are picked by hand out of Excel's
   palette and drift: one row's "red" is FF0000 and the next is FF3300. A cell
   with almost no colour in it (grey, black, white) is ordinary text and means
   nothing, so saturation and lightness gate the hue before it is read at all.
   Orange, yellow and violet fall through deliberately: the sheet does not use
   them as a code, and guessing at one would invent a status.               */
function hsl(hex) {
  if (!/^[0-9A-F]{6}$/i.test(String(hex || ''))) return null;
  const n = parseInt(hex, 16);
  const r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  const l = (max + min) / 2;
  if (!d) return { h: 0, s: 0, l: l };
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = ((g - b) / d) % 6;
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  h *= 60;
  if (h < 0) h += 360;
  return { h: h, s: s, l: l };
}
/** The flag word a text colour stands for, or '' for "no code here". */
function flagOf(hex) {
  const c = hsl(hex);
  if (!c) return '';
  if (c.s < 0.25 || c.l < 0.12 || c.l > 0.92) return '';   // grey, black, white: ordinary text
  const h = c.h;
  if (h >= 345 || h < 20) return 'urgent';
  if (h >= 75 && h < 170) return 'booked';
  if (h >= 170 && h < 260) return 'hold';
  if (h >= 275 && h < 345) return 'trade';
  return '';
}
/** Is this ink "the sheet said nothing here" - unset, or the black the rows
    are written in? Grey, orange and everything else are somebody's decision,
    even when they are not one of the four words, so they are left alone: the
    customer cell is only asked when the job number itself is plain black. */
function blackInk(hex) {
  if (!hex) return true;
  const c = hsl(hex);
  return !!c && c.s < 0.2 && c.l < 0.2;
}

const GOLD = new Set(['FFE699', 'FFC000']);
const YELLOW = new Set(['FFFF00']);
/* "Cut" is a green the sheet paints on a product cell once that part has been
   cut. The sheet says which green in its own legend - a cell reading "Cut=" on
   the header rows, filled in that colour - so repainting the legend carries
   straight through to here. The list below is the greens the sheet has used,
   for a workbook whose legend cannot be read. */
const CUT_FALLBACK = ['00B050', '92D050', 'C6EFCE'];
let CUT = new Set(CUT_FALLBACK);

/** The sheet's own Cut colour, read off the legend cell on rows 1-2 of the
    Production sheet. A candidate that is white, yellow or gold is not a colour
    key at all - it is a heading that happens to say "cut" - so it is passed
    over, and the fallback greens are used instead. */
function cutColours(ws) {
  if (ws) {
    const maxC = Math.min(ws.columnCount || 60, 60);
    /* two passes: the legend is written "Cut=", so a cell carrying the equals
       sign is preferred over any other header that merely contains the word */
    for (const wantEq of [true, false]) {
      for (let r = 1; r <= 2; r++) {
        for (let c = 1; c <= maxC; c++) {
          const cell = ws.getRow(r).getCell(c), txt = cellText(cell);
          if (norm(txt).indexOf('cut') < 0) continue;
          if (wantEq && txt.indexOf('=') < 0) continue;
          const f = fillOf(cell);
          if (f && f !== 'FFFFFF' && !YELLOW.has(f) && !GOLD.has(f)) return new Set([f]);
        }
      }
    }
  }
  return new Set(CUT_FALLBACK);
}

/* A job shows up on several sheets; when two of them disagree about a cell's
   colour the further-along one wins, because a colour is only ever added as
   work is finished - it is never taken back to mean "less done". */
const CPRANK = { '': 0, cut: 1, process: 2, done: 3 };
const cpOf = fill => YELLOW.has(fill) ? 'process'
  : (GOLD.has(fill) ? 'done' : (CUT.has(fill) ? 'cut' : ''));
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

/* ---- "Production (2)": John's own sheet -------------------------------------
   Production (2) is not a copy of Production that happens to lag: it is a
   separate sheet, kept by the office for the paper John works from, with its
   own rows, its own order, its own sections and its own colouring. So it is
   read on its own terms here and NOTHING from it is merged into the job model
   above - a job's status, checkpoints and colour code still come from
   Production alone, and John's print comes from this and nothing else.

   The identity columns are the fixed ones of that sheet (C, G, I, J, K, M, N);
   only the notes column is found by its header, because "Notes from Brendan's
   office" has moved along the row before now.                               */
const JOHN_SHEET = 'Production (2)';
const JOHN_COLS = { id: 3, ready: 7, cust: 9, phone: 10, area: 11, wnd: 13, drs: 14 };

/** The row's own fill: the job-number cell if it carries one, otherwise the
    first of the eight printed cells that does - the office paints a whole row,
    but not always every cell of it. */
function johnFill(row) {
  const first = fillOf(row.getCell(JOHN_COLS.id));
  if (first && first !== 'FFFFFF') return first;
  const cols = [JOHN_COLS.ready, JOHN_COLS.cust, JOHN_COLS.phone, JOHN_COLS.area,
                JOHN_COLS.wnd, JOHN_COLS.drs];
  for (const c of cols) {
    const f = fillOf(row.getCell(c));
    if (f && f !== 'FFFFFF') return f;
  }
  return '';
}

/** Production (2) as the rows John prints, in that sheet's own order.
    [] when the workbook has no such sheet - never null, so a caller can
    always iterate it. Each row:
      { id, section, ready, cust, phone, area, wnd, drs, notes,
        fillHex, inkHex, seq }                                              */
function parseJohnSheet(wb) {
  const ws = wb && wb.getWorksheet ? wb.getWorksheet(JOHN_SHEET) : null;
  if (!ws) return [];
  const maxR = ws.rowCount || 0;

  /* the notes column, by its header, anywhere in the first six rows. The whole
     width has to be searched: on the real sheet it is column BY, seventy-odd
     columns out past the ones that are printed. */
  let notesCol = 0;
  const maxC = Math.min(ws.columnCount || 150, 150);
  for (let r = 1; r <= Math.min(6, maxR) && !notesCol; r++) {
    for (let c = 1; c <= maxC; c++) {
      if (norm(cellText(ws.getRow(r).getCell(c))).indexOf('brendan') >= 0) { notesCol = c; break; }
    }
  }

  /* this sheet's own sections, from its own divider rows */
  const matrix = [];
  for (let r = 1; r <= maxR; r++) {
    const row = ws.getRow(r), line = [];
    for (let c = 1; c <= 11; c++) line.push(cellText(row.getCell(c)));
    matrix.push(line);
  }
  const B = blocksFromValues(matrix);
  const sectionOf = {};
  B.blocks.forEach(b => b.jobs.forEach(j => { sectionOf[j.row] = B.names[b.idx] || ''; }));

  const out = [];
  for (let r = 1; r <= maxR; r++) {
    const row = ws.getRow(r);
    const id = cellText(row.getCell(JOHN_COLS.id)).trim().toUpperCase();
    if (!JOB_RE.test(id)) continue;
    let ink = fontOf(row.getCell(JOHN_COLS.id));
    if (blackInk(ink)) {
      const alt = fontOf(row.getCell(JOHN_COLS.cust));
      if (flagOf(alt)) ink = alt;
    }
    out.push({
      id: id,
      section: sectionOf[r] || '',
      ready: cellDate(row.getCell(JOHN_COLS.ready)) || cellText(row.getCell(JOHN_COLS.ready)).trim(),
      cust: cellText(row.getCell(JOHN_COLS.cust)).trim().slice(0, 70),
      phone: cellText(row.getCell(JOHN_COLS.phone)).trim().slice(0, 40),
      area: cellText(row.getCell(JOHN_COLS.area)).trim().slice(0, 70),
      wnd: num(row.getCell(JOHN_COLS.wnd)),
      drs: num(row.getCell(JOHN_COLS.drs)),
      notes: notesCol ? cellText(row.getCell(notesCol)).trim().slice(0, 300) : '',
      fillHex: johnFill(row),
      inkHex: flagOf(ink) ? ink : '',
      seq: out.length
    });
  }
  /* the section names this sheet uses, in the order it uses them */
  out.sections = [];
  out.forEach(x => { if (x.section && out.sections.indexOf(x.section) < 0) out.sections.push(x.section); });
  return out;
}

function parseWorkbook(wb) {
  const prodSheet = wb.getWorksheet('Production');
  CUT = cutColours(prodSheet);          // the sheet's legend decides, once per workbook
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

      /* The colour code is read from the Production sheet alone: Production (2)
         is a copy of it and is repainted by hand, so it is never asked. The
         job-number cell is the one that is always filled in; when its font is
         black (or has no colour of its own) the customer cell is asked instead,
         because the two are coloured by hand and do not always agree. */
      if (cpHere && !j.flag) {
        let hex = fontOf(row.getCell(3)), fl = flagOf(hex);
        if (!fl && blackInk(hex) && m.ident.cust) {
          const h2 = fontOf(row.getCell(m.ident.cust)), f2 = flagOf(h2);
          if (f2) { hex = h2; fl = f2; }
        }
        if (fl) { j.flag = fl; j.flagHex = hex; }
      }

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
      /* The whole phone number, as the sheet holds it. It exists for exactly
         one reader: the John print template, which the owner sanctioned on
         2026-09-09 to carry it (docs/specs/2026-09-09-john-template.md, §3).
         Nothing else may read `ph` - the dashboard shows `ph3`, and every other
         export path reads neither. The eircode is never carried anywhere. */
      ph: String(j.phone || '').trim().slice(0, 40),
      /* the colour of the row's own text on Production: '' / urgent / booked /
         trade / hold, with the hex that said so, for the exports */
      flag: j.flag || '', flagHex: j.flagHex || '',
      wnd: j.wnd || 0, drs: j.drs || 0,
      dates: { sold: j.d_sold || null, stamp: j.d_stamp || null, ivana: j.d_ivana || null, ready: j.d_ready || null, floor: j.d_floor || null },
      prods: prods, cp: j.cp,
      glass: j.glass, notes: j.notes, sheets: j.sheets, src: j.src,
      /* the word in a comment, or the sheet's own red text: either says urgent */
      urg: (URG_RE.test(cmtxt) || j.flag === 'urgent') ? 1 : 0, done: j.done || 0,
      cat: PRODCAT[id] || 'past',
      blk: (B.blk[id] === undefined ? -1 : B.blk[id]),
      seq: (B.order[id] === undefined ? 99999 : B.order[id]),
      stage: j.d_floor ? 'floor' : (j.d_ready ? 'ready' : 'office')
    };
  });
  result.blockNames = B.names;
  return result;
}

if (typeof module !== 'undefined') module.exports = { parseWorkbook, parseJohnSheet, mapSheet, fillOf, fontOf, flagOf, blackInk, JOB_RE, URG_RE, blocksFromValues, templateForJob, cutColours };
if (typeof window !== 'undefined') { window.parseWorkbook = parseWorkbook; window.mapSheet = mapSheet;
  window.fontOf = fontOf; window.flagOf = flagOf; window.blackInk = blackInk;
  window.parseJohnSheet = parseJohnSheet;
  window.blocksFromValues = blocksFromValues; window.templateForJob = templateForJob;
  window.cutColours = cutColours; }
