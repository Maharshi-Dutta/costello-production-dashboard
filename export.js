/* Export to Excel or PDF - the pure half.

   Everything here is data in, data out: choosing the jobs, turning them into
   plain rows, naming the file, wording the filter summary, and building either
   an ExcelJS workbook or a pdfmake document definition. None of it touches the
   DOM or the network, so test_export.js can load this file and exercise the
   whole export in Node. The window that drives it lives in app.js.

   Two rules run through the whole file and are worth stating once:
   - PHONE NO. and EIRCODE never leave this file. `j.ph3` and `j.eir` are never
     read, no column or PDF text is ever built from them, and the free-text
     search below deliberately leaves `eir` out of the haystack even though the
     dashboard's own search includes it. There is no field, no option and no
     preset that can turn them back on. (What somebody typed into a Comment is
     their text and is exported as written; the rule is about the two columns.)
   - No network. The export works entirely on the jobs already in memory; the
     only write the feature makes at all is the single Dashboard Log line that
     app.js appends afterwards.

   Loaded after checkpoints.js (it uses cpItems / cpStatus / itemState) and
   before app.js. Plain ES2017, no modules. */

/* ---- the dashboard's colour language ----------------------------------
   Every one of these is only ever used next to a number or a word, because
   the office reads these on paper and some of the readers are colour-blind. */
const XP_GOLD = "#FFE699";        // a checkpoint that is done
const XP_YELLOW = "#FFFF00";      // in fabrication  (dark text on top of it)
const XP_GREEN = "#2E7D32";       // ready to deliver
const XP_RED = "#C62828";         // urgent
const XP_INK = "#17171A";
const XP_INK_2 = "#4D4A44";
const XP_GREY = "#6D6A62";
const XP_LINE = "#DCD9D2";
const XP_PAPER = "#FAF9F7";
const XP_F = "#2A78D6", XP_S = "#C25525", XP_T = "#14805A";   // frames / sashes / transoms
const XP_GREEN_BG = "#E8F5E3";

/* the same colours as Excel wants them (solid ARGB) */
const xpArgb = hex => "FF" + String(hex).replace("#", "").toUpperCase();
const XL_GOLD = xpArgb(XP_GOLD), XL_YELLOW = xpArgb(XP_YELLOW);

/* ---- the page, in points ----
   A4 is 595.28 x 841.89. With a 28pt margin each side that leaves 539 across
   a portrait page and 786 across a landscape one. Every width below is
   measured against these, so nothing is ever drawn off the edge of the paper. */
const XP_PAGE_W = 539;            // portrait, inside the margins
const XP_LAND_W = 786;            // landscape, inside the margins
const XP_HEAD_W = 511;            // inside a job's header band (539 less its padding)

const XP_PRESET_KEY = "cw_exportpresets";
const XP_SUBLABEL = { f: "F", s: "S", t: "T" };
const XP_DATE_STEPS = [["sold", "Sold"], ["stamp", "Stamp"], ["ivana", "Ivana"],
                       ["ready", "Ready to print"], ["floor", "Sent to floor"]];
const XP_NOTE_SOURCE = { comment: "Sheet comment", brendan: "Brendan's notes", specials: "Specials" };
const XP_STATUS_WORD = { "": "not started", none: "not started", process: "in fabrication", done: "done" };
/* the dashboard's own sort keys, so "Sort by" means the same in both places */
const XP_SORTS = [["id", "Job no (A-Z)"], ["num", "Job number (ignore letter)"], ["cat", "Category"],
                  ["urgent", "Urgent first"], ["size", "Biggest first"], ["wait", "Longest wait"],
                  ["county", "County"]];

/** The picker, in the order it is shown and in the order the columns come out. */
const EXPORT_FIELDS = [
  ["job", "Job no"], ["cust", "Customer"], ["county", "County"], ["section", "Section"],
  ["ready", "Ready to deliver"], ["urgent", "Urgent"], ["dates", "Dates (five)"],
  ["wnd", "Windows"], ["drs", "Doors"], ["colour", "Windows colour"], ["off", "Office no"],
  ["prods", "Products F/S/T"], ["glass", "Glass units"], ["cp", "Checkpoints"],
  ["comments", "Comments"], ["alerts", "Alerts"], ["notes", "Notes from the sheet"]
];
const EXPORT_FIELD_KEYS = EXPORT_FIELDS.map(p => p[0]);

/** Every field on - the default, and what "all" gives you. */
function exportAllFields() {
  const o = {};
  EXPORT_FIELD_KEYS.forEach(k => { o[k] = true; });
  return o;
}
/** Accepts an array of keys or a map, always hands back a map. Unknown keys
    are dropped, so nothing outside EXPORT_FIELDS can ever be asked for. */
function xpFieldSet(fields) {
  if (!fields) return exportAllFields();
  const o = {};
  if (Array.isArray(fields)) fields.forEach(k => { if (EXPORT_FIELD_KEYS.indexOf(k) >= 0) o[k] = true; });
  else EXPORT_FIELD_KEYS.forEach(k => { if (fields[k]) o[k] = true; });
  return o;
}

/* ---- small helpers ---- */
const xpStr = v => String(v == null ? "" : v);
const xpLow = v => xpStr(v).trim().toLowerCase();
const xpDay = v => xpStr(v).slice(0, 10);
const xpUpper = v => xpStr(v).toUpperCase();
const xpIsIso = v => /^\d{4}-\d{2}-\d{2}$/.test(xpDay(v));
/** What the sheet holds in a date cell, kept whole. Not every one of them is a
    date: the office writes things like "Before 12 Nov" in there, and throwing
    that away would lose the only thing the cell was saying. */
const xpDateVal = v => xpStr(v).trim() || null;
/** An exported file names a person, not an address. */
const xpShortWho = v => xpStr(v).split("@")[0].trim();
function xpPad(n) { return (n < 10 ? "0" : "") + n; }
/** Local calendar date as YYYY-MM-DD - the day the person pressed Download. */
function xpIsoDate(when) {
  if (!when) return xpIsoDate(new Date());
  if (typeof when === "string") return xpDay(when);
  return when.getFullYear() + "-" + xpPad(when.getMonth() + 1) + "-" + xpPad(when.getDate());
}
function xpStamp(when) {
  const d = (when && typeof when !== "string") ? when : new Date(when || Date.now());
  return xpIsoDate(d) + " " + xpPad(d.getHours()) + ":" + xpPad(d.getMinutes());
}
/** dd-mmm-yyyy, the same shape the Excel cells are formatted in. */
const XP_MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function xpNiceDate(iso) {
  if (!xpIsIso(iso)) return xpStr(iso);
  const s = xpDay(iso);
  return s.slice(8, 10) + "-" + XP_MON[parseInt(s.slice(5, 7), 10) - 1] + "-" + s.slice(0, 4);
}
/** An ISO day as a real Date at noon UTC, so no timezone can shift the day. */
function xpDateCell(iso) {
  if (!xpIsIso(iso)) return null;
  const s = xpDay(iso);
  return new Date(Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10), 12, 0, 0));
}
const xpJobNum = j => { const m = /(\d+)/.exec(xpStr(j.id)); return m ? parseInt(m[1], 10) : 0; };
/* the dashboard's catOf(), repeated here because export.js loads before app.js
   and this half has to stay loadable on its own in Node */
function xpCat(j) {
  if (j.cat === "collect") return "collect";
  if (j.cat === "wonttake") return "wonttake";
  if (j.cat === "secondhand") return "secondhand";
  if (j.done) return "deliver";
  return j.stage;
}
const XP_CATORDER = ["deliver", "collect", "wonttake", "secondhand", "floor", "ready", "office"];
const xpComp = j => (j.prods || []).reduce((a, p) => ({ f: a.f + (p.f || 0), s: a.s + (p.s || 0), t: a.t + (p.t || 0) }), { f: 0, s: 0, t: 0 });
const xpTot = c => c.f + c.s + c.t;

/** Checkpoint status of one item, straight from the sheet's colour. Falls back
    to "" when checkpoints.js is not loaded, so this file never hard-fails. */
function xpItemStatus(j, item) {
  return typeof cpStatus === "function" ? (cpStatus(j, item) || "") : "";
}
function xpItems(j) {
  return typeof cpItems === "function" ? cpItems(j) : [];
}
/** The colour of a whole checkpoint group, for the filter:
    "done" (every item done), "process" (some colour), "none" (no colour at
    all), or null when the job has no such group to speak of. */
function xpGroupStatus(j, key) {
  const items = xpItems(j).filter(it => key === "prod"
    ? xpStr(it.group).indexOf("prod:") === 0
    : it.group === key);
  if (!items.length) return null;
  const sts = items.map(it => xpItemStatus(j, it.key));
  if (sts.every(s => s === "done")) return "done";
  if (sts.some(s => s !== "")) return "process";
  return "none";
}

/* ---------------------------------------------------------------------------
   1. choosing the jobs
   --------------------------------------------------------------------------- */

/** Normalised context: the lists and lookups the filters need, all supplied by
    the caller. Nothing here goes looking for globals. */
function xpContext(ctx) {
  ctx = ctx || {};
  return {
    all: ctx.all || [],
    view: ctx.view || null,
    picked: ctx.picked || {},
    sections: ctx.sections || [],
    comments: typeof ctx.comments === "function" ? ctx.comments : () => [],
    alerts: typeof ctx.alerts === "function" ? ctx.alerts : () => []
  };
}

/** A filter object with every key present, so nothing downstream has to guess. */
function exportDefaults() {
  return {
    scope: "view", sections: [], sectionNames: [],
    ready: null, urgent: null,
    cp: { win: null, drs: null, glass: null, prod: null },
    dates: { sold: { from: null, to: null }, ready: { from: null, to: null }, floor: { from: null, to: null } },
    county: [], products: [], glassTypes: [], sheets: [],
    hasComments: null, hasAlerts: null,
    q: "", sort: "id", desc: false, groupBySection: false
  };
}
/** A complete filter built from a partial one. Every nested object and array is
    copied, never aliased: the caller's own filter (the live one in the Export
    window) must not change because something in here tidied a default in. */
function xpFilter(f) {
  const d = exportDefaults();
  if (!f) return d;
  const out = Object.assign({}, d, f);
  out.cp = Object.assign({}, d.cp, f.cp || {});
  out.dates = {};
  Object.keys(d.dates).forEach(k => {
    out.dates[k] = Object.assign({ from: null, to: null }, (f.dates || {})[k] || {});
  });
  out.sections = (f.sections || []).map(Number);
  out.sectionNames = (f.sectionNames || []).slice();
  ["county", "products", "glassTypes", "sheets"].forEach(k => { out[k] = (f[k] || []).slice(); });
  return out;
}

/** What "What I see now" / "Ticked jobs" / "Sections…" actually mean. */
function exportScope(ctx, f) {
  const c = xpContext(ctx), s = xpFilter(f);
  if (s.scope === "ticked") return c.all.filter(j => c.picked[j.id]);
  if (s.scope === "sections") return c.all.filter(j => s.sections.indexOf(Number(j.blk)) >= 0);
  return (c.view || c.all).slice();
}

/** A date range test. A cell the sheet holds as text ("Before 12 Nov") is not a
    date anybody can compare, so it counts as unknown: it drops out the moment a
    range is set, and is left alone when none is. */
const xpInRange = (v, r) => {
  if (!r || (!r.from && !r.to)) return true;
  const d = xpDay(v);
  if (!xpIsIso(d)) return false;
  if (r.from && d < xpDay(r.from)) return false;
  if (r.to && d > xpDay(r.to)) return false;
  return true;
};

/** The jobs matching `f`. Applied on top of whatever list is handed in, so the
    caller decides the scope; `scope:"sections"` is honoured here as well, for
    the times the whole live list is passed straight in. */
function exportFilter(jobs, f, ctx) {
  const s = xpFilter(f), c = xpContext(ctx);
  const county = s.county.map(xpLow).filter(Boolean);
  const prods = s.products.map(xpLow).filter(Boolean);
  const glass = s.glassTypes.map(xpLow).filter(Boolean);
  const sheets = s.sheets.filter(Boolean);
  const q = xpLow(s.q);
  return (jobs || []).filter(j => {
    if (s.scope === "sections" && s.sections.length && s.sections.indexOf(Number(j.blk)) < 0) return false;
    if (s.ready === true && !j.done) return false;
    if (s.ready === false && j.done) return false;
    if (s.urgent === true && !j.urg) return false;
    if (s.urgent === false && j.urg) return false;
    for (const k in s.cp) {
      const want = s.cp[k];
      if (want == null || want === "") continue;
      if (xpGroupStatus(j, k) !== want) return false;
    }
    const d = j.dates || {};
    if (!xpInRange(d.sold, s.dates.sold)) return false;
    if (!xpInRange(d.ready, s.dates.ready)) return false;
    if (!xpInRange(d.floor, s.dates.floor)) return false;
    if (county.length && county.indexOf(xpLow(j.area)) < 0) return false;
    if (prods.length && !(j.prods || []).some(p => prods.indexOf(xpLow(p.n)) >= 0)) return false;
    if (glass.length && !glass.some(g => (j.glass || {})[g] > 0)) return false;
    if (sheets.length && !sheets.some(x => (j.sheets || []).indexOf(x) >= 0)) return false;
    if (s.hasComments === true && !c.comments(j.id).length) return false;
    if (s.hasComments === false && c.comments(j.id).length) return false;
    if (s.hasAlerts === true && !c.alerts(j.id).length) return false;
    if (s.hasAlerts === false && c.alerts(j.id).length) return false;
    if (q) {
      /* the dashboard's own search minus the eircode - the one deliberate
         difference, because no part of the export path may read `eir` */
      const hay = (j.id + " " + (j.cust || "") + " " + (j.area || "") + " " + (j.off || "") + " " +
                   (j.notes || []).map(n => n.t).join(" ")).toLowerCase();
      if (hay.indexOf(q) < 0) return false;
    }
    return true;
  });
}

/** The dashboard's sort keys, applied to an already-filtered list. */
function exportSort(jobs, f) {
  const s = xpFilter(f), list = (jobs || []).slice();
  const bi = (a, b) => xpStr(a.id).localeCompare(xpStr(b.id));
  const k = s.sort;
  list.sort(k === "size" ? (a, b) => xpTot(xpComp(b)) - xpTot(xpComp(a))
    : k === "wait" ? (a, b) => (((a.dates || {}).ready || "9") < ((b.dates || {}).ready || "9") ? -1 : 1)
    : k === "county" ? (a, b) => xpStr(a.area || "~").localeCompare(xpStr(b.area || "~")) || bi(a, b)
    : k === "urgent" ? (a, b) => ((b.urg ? 1 : 0) - (a.urg ? 1 : 0)) || bi(a, b)
    : k === "num" ? (a, b) => (xpJobNum(a) - xpJobNum(b)) || bi(a, b)
    : k === "cat" ? (a, b) => (XP_CATORDER.indexOf(xpCat(a)) - XP_CATORDER.indexOf(xpCat(b))) || bi(a, b)
    : bi);
  if (s.desc) list.reverse();
  return list;
}

/** Scope, then filter, then sort - the whole choosing step in one call. */
function exportJobs(ctx, f) {
  return exportSort(exportFilter(exportScope(ctx, f), f, ctx), f);
}

/* ---------------------------------------------------------------------------
   2. the rows
   --------------------------------------------------------------------------- */

/** Checkpoints for one job, gathered into the groups the drawer shows:
    Windows, Doors, Glass, and one per product. `done` can be unknown for an
    item that is in fabrication with no count entered - that is carried as
    `unknown` rather than pretending it is zero. */
function exportCheckpoints(j) {
  const order = [], byGroup = {};
  xpItems(j).forEach(it => {
    const st = (typeof itemState === "function" ? itemState(j, it.key) : null) ||
               { done: 0, total: it.total, status: "" };
    let g = byGroup[it.group];
    if (!g) { g = byGroup[it.group] = { group: it.group, label: it.groupLabel, done: 0, total: 0, unknown: false, items: [] }; order.push(g); }
    g.items.push({ item: it.key, label: it.label, done: st.done, total: st.total, status: st.status });
    g.total += st.total || 0;
    if (st.done == null) g.unknown = true; else g.done += st.done;
  });
  order.forEach(g => {
    g.status = (g.total > 0 && g.done >= g.total) ? "done" : ((g.done > 0 || g.unknown) ? "process" : "");
  });
  return order;
}

/** Plain row objects holding only the ticked fields. This is the one place a
    job object is read; everything after this point sees rows, never jobs, so
    a field that is not ticked cannot leak into a workbook or a PDF. */
function exportRows(jobs, fields, ctx) {
  const F = xpFieldSet(fields), c = xpContext(ctx);
  return (jobs || []).map(j => {
    /* the job number is the row's key, not one of its fields: the Comments and
       Checkpoints sheets are meaningless without it. Whether a "Job no" column
       appears in the Jobs sheet or on a card is still the picker's decision. */
    const r = { id: xpStr(j.id) };
    if (F.cust) r.cust = xpStr(j.cust);
    if (F.county) r.area = xpStr(j.area);
    if (F.section) r.section = xpStr(c.sections[j.blk] || "");
    if (F.ready) r.ready = !!j.done;
    if (F.urgent) r.urg = !!j.urg;
    if (F.dates) {
      const d = j.dates || {};
      r.dates = { sold: xpDateVal(d.sold), stamp: xpDateVal(d.stamp), ivana: xpDateVal(d.ivana),
                  ready: xpDateVal(d.ready), floor: xpDateVal(d.floor) };
    }
    if (F.wnd) r.wnd = j.wnd || 0;
    if (F.drs) r.drs = j.drs || 0;
    if (F.colour) r.colour = xpStr(j.colour);
    if (F.off) r.off = xpStr(j.off);
    if (F.prods) {
      r.prods = (j.prods || []).map(p => ({
        n: xpStr(p.n), f: p.f || 0, s: p.s || 0, t: p.t || 0,
        st: { f: xpItemStatus(j, "prod:" + p.n + ":f"), s: xpItemStatus(j, "prod:" + p.n + ":s"),
              t: xpItemStatus(j, "prod:" + p.n + ":t") }
      }));
    }
    if (F.glass) {
      r.glass = Object.keys(j.glass || {}).filter(k => j.glass[k] > 0)
        .map(k => ({ type: xpStr(k), qty: j.glass[k], st: xpItemStatus(j, "glass:" + k) }));
    }
    if (F.cp) r.cp = exportCheckpoints(j);
    if (F.comments) {
      /* a name, not an address: an exported file goes to people outside the
         office often enough that it should not carry everyone's email */
      r.comments = (c.comments(j.id) || []).map(x => ({
        src: "Dashboard", who: xpShortWho(x.who), when: xpStr(x.at).slice(0, 16), text: xpStr(x.to)
      }));
    }
    if (F.notes) {
      r.notes = (j.notes || []).map(n => ({
        src: XP_NOTE_SOURCE[n.k] || "Sheet comment", kind: xpStr(n.k), sheet: xpStr(n.s),
        who: "", when: "", text: xpStr(n.t)
      }));
    }
    if (F.alerts) r.alerts = (c.alerts(j.id) || []).length;   // the count only; never an address
    return r;
  });
}

/* ---------------------------------------------------------------------------
   3. wording: the filename, the summary, the log line
   --------------------------------------------------------------------------- */

const xpSortLabel = k => (XP_SORTS.find(p => p[0] === k) || ["", k])[1];

/** The chosen sections by name. The window fills in `sectionNames`; a preset or
    a saved filter carries only the numbers, so the sheet's own list of section
    names is used instead - "Scope: sections 3" helps nobody. */
function xpSectionNames(f, names) {
  const s = xpFilter(f);
  if (s.sectionNames.length) return s.sectionNames.slice();
  return s.sections.map(i => xpStr((names || [])[i]) || ("section " + i));
}

/** The short "what this is" used in the filename. The free-text search is
    deliberately left out of it: a filename is not the place for whatever
    somebody happened to type into the box. */
function exportScopeLabel(f, names) {
  const s = xpFilter(f);
  if (s.scope === "ticked") return "Ticked jobs";
  if (s.scope === "sections") {
    const list = xpSectionNames(f, names);
    if (!list.length) return "Sections";
    const joined = list.join(" + ");
    return joined.length <= 60 ? joined : (list[0] + " + " + (list.length - 1) + " more");
  }
  if (s.ready === true) return "Ready to deliver";
  if (s.ready === false) return "In production";
  if (s.urgent === true) return "Urgent";
  if (s.county.length) return s.county.join(" + ");
  return "All jobs";
}

/** One line per filter that is actually set - the wording used in the log, the
    Export info sheet and the PDF cover, so all three always say the same thing. */
function filtersSummaryLines(f, names) {
  const s = xpFilter(f), out = [];
  out.push("Scope: " + (s.scope === "ticked" ? "ticked jobs"
    : s.scope === "sections" ? (xpSectionNames(f, names).join(", ") || "no section chosen")
    : "what I see now"));
  if (s.ready === true) out.push("Ready to deliver only");
  if (s.ready === false) out.push("In production only (not ready to deliver)");
  if (s.urgent === true) out.push("Urgent only");
  const cpName = { win: "Windows", drs: "Doors", glass: "Glass", prod: "Products" };
  Object.keys(cpName).forEach(k => {
    const v = s.cp[k];
    if (v) out.push(cpName[k] + ": " + (XP_STATUS_WORD[v] || v));
  });
  const dname = { sold: "Sold", ready: "Ready to print", floor: "Sent to floor" };
  Object.keys(dname).forEach(k => {
    const r = s.dates[k];
    if (!r || (!r.from && !r.to)) return;
    if (r.from && r.to) out.push(dname[k] + " between " + xpNiceDate(r.from) + " and " + xpNiceDate(r.to));
    else if (r.from) out.push(dname[k] + " from " + xpNiceDate(r.from));
    else out.push(dname[k] + " up to " + xpNiceDate(r.to));
  });
  if (s.county.length) out.push("County: " + s.county.join(", "));
  if (s.products.length) out.push("Products: " + s.products.map(xpUpper).join(", "));
  if (s.glassTypes.length) out.push("Glass: " + s.glassTypes.map(xpUpper).join(", "));
  if (s.sheets.length) out.push("On sheet: " + s.sheets.join(", "));
  if (s.hasComments === true) out.push("With comments");
  if (s.hasComments === false) out.push("Without comments");
  if (s.hasAlerts === true) out.push("With email alerts");
  if (s.hasAlerts === false) out.push("Without email alerts");
  if (xpStr(s.q).trim()) out.push('Search: "' + xpStr(s.q).trim() + '"');
  out.push("Sorted by " + xpSortLabel(s.sort) + (s.desc ? " (reversed)" : ""));
  if (s.groupBySection) out.push("Grouped by section");
  return out;
}
/** The same thing on one line, for the log row and the file. */
function filtersSummary(f, names) { return filtersSummaryLines(f, names).join("; "); }

/** Windows will not take \ / : * ? " < > | in a name, and a very long one is
    unhelpful anyway. */
function xpSafeName(s) {
  return xpStr(s).replace(/[\\/:*?"<>|\r\n\t]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 70);
}
function exportFilename(format, f, when, names) {
  const ext = xpLow(format) === "pdf" ? ".pdf" : ".xlsx";
  const label = xpSafeName(exportScopeLabel(f, names));
  return xpSafeName("Production export" + (label ? " - " + label : "")) + " - " + xpIsoDate(when) + ext;
}
/** The "from" half of the Dashboard Log line. */
function exportLogFrom(format, n) {
  return (xpLow(format) === "pdf" ? "PDF" : "Excel") + " · " + n + " job" + (n === 1 ? "" : "s");
}

/* ---------------------------------------------------------------------------
   4. presets - names only, never job data
   --------------------------------------------------------------------------- */

/** What a preset is allowed to hold. A preset is a set of choices; it never
    carries a job, a customer or anything read out of the workbook. */
function xpCleanPreset(p) {
  p = p || {};
  const f = xpFilter(p.f);
  f.sectionNames = [];                         // names come back from BLOCKNAMES at load time
  return {
    format: xpLow(p.format) === "pdf" ? "pdf" : "xlsx",
    layout: p.layout === "table" ? "table" : "cards",
    fields: xpFieldSet(p.fields),
    f: f
  };
}
function presetsLoad() {
  let raw = {};
  try { raw = JSON.parse(localStorage.getItem(XP_PRESET_KEY) || "{}") || {}; } catch (e) { raw = {}; }
  const out = {};
  Object.keys(raw).forEach(k => { out[k] = xpCleanPreset(raw[k]); });
  return out;
}
function presetsSaveAll(map) {
  try { localStorage.setItem(XP_PRESET_KEY, JSON.stringify(map || {})); } catch (e) {}
  return map || {};
}
function presetSave(name, p) {
  const n = xpStr(name).trim().slice(0, 40);
  if (!n) return presetsLoad();
  const all = presetsLoad();
  all[n] = xpCleanPreset(p);
  return presetsSaveAll(all);
}
function presetDelete(name) {
  const all = presetsLoad();
  delete all[xpStr(name).trim()];
  return presetsSaveAll(all);
}

/* ---------------------------------------------------------------------------
   5. the columns
   --------------------------------------------------------------------------- */

/** Product names and glass types actually present in this export, first-seen
    order, so the columns match what is really in the file. */
function xpPresentProducts(rows) {
  const seen = [], has = {};
  (rows || []).forEach(r => (r.prods || []).forEach(p => { if (!has[p.n]) { has[p.n] = 1; seen.push(p.n); } }));
  return seen;
}
function xpPresentGlass(rows) {
  const seen = [], has = {};
  (rows || []).forEach(r => (r.glass || []).forEach(g => { if (!has[g.type]) { has[g.type] = 1; seen.push(g.type); } }));
  return seen;
}

/** The Production sheet's own column order, minus PHONE NO. and EIRCODE, minus
    anything not ticked. `kind` says how a cell is written and coloured. This is
    the wide, detailed set: it is what the workbook gets, and it can easily run
    to eighty columns on the real sheet. */
function exportColumns(rows, fields) {
  const F = xpFieldSet(fields), cols = [];
  const add = (key, name, width, kind, extra) => cols.push(Object.assign({ key, name, width, kind: kind || "text" }, extra || {}));
  if (F.notes) add("comment", "Comment", 34, "text");
  if (F.off) add("off", "Office no", 12, "text");
  if (F.job) add("id", "Job no", 11, "job");
  if (F.dates) XP_DATE_STEPS.forEach(p => add("date:" + p[0], p[1], 13, "date"));
  if (F.cust) add("cust", "Customer", 24, "text");
  if (F.county) add("area", "Area", 15, "text");
  if (F.section) add("section", "Section", 22, "text");
  if (F.ready) add("ready", "Ready to deliver", 16, "yesno");
  if (F.urgent) add("urg", "Urgent", 9, "yesno");
  if (F.wnd) add("wnd", "WND", 7, "num");
  if (F.drs) add("drs", "DRS", 7, "num");
  if (F.colour) add("colour", "Windows colour", 18, "text");
  if (F.prods) xpPresentProducts(rows).forEach(n => ["f", "s", "t"].forEach(sub =>
    add("prod:" + n + ":" + sub, xpUpper(n) + " " + XP_SUBLABEL[sub], 11, "prod", { prod: n, sub: sub })));
  if (F.glass) xpPresentGlass(rows).forEach(t =>
    add("glass:" + t, xpUpper(t), 9, "glass", { glass: t }));
  if (F.alerts) add("alerts", "Alerts", 8, "num");
  if (F.comments) add("comments", "Comments", 10, "num");
  /* an Excel table refuses two columns with the same header */
  const used = {};
  cols.forEach(c => {
    let n = c.name, i = 1;
    while (used[xpLow(n)]) { i++; n = c.name + " (" + i + ")"; }
    used[xpLow(n)] = 1; c.name = n;
  });
  return cols;
}

/** The PDF table's own, deliberately short list.

    The wide set above is right for a spreadsheet, where a column can be
    narrow and still be read, and wrong for a sheet of A4: the real workbook
    has twenty-one products and eight glass types, which is eighty columns and
    about four points each. So the table layout summarises instead - the F/S/T
    totals in one cell, the glass units in one cell - and the per-product
    detail stays in the workbook and on the job cards, where there is room for
    it. The field picker still decides which of these appear. */
const XP_TABLE_COLS = [
  { key: "id", name: "Job no", w: 46, field: "job", kind: "job" },
  { key: "cust", name: "Customer", w: 92, field: "cust" },
  { key: "area", name: "County", w: 54, field: "county" },
  { key: "section", name: "Section", w: 76, field: "section" },
  { key: "ready", name: "Ready", w: 34, field: "ready", kind: "yesno" },
  { key: "urg", name: "Urgent", w: 34, field: "urgent", kind: "yesno" },
  { key: "wnd", name: "WND", w: 30, field: "wnd", kind: "num" },
  { key: "drs", name: "DRS", w: 30, field: "drs", kind: "num" },
  { key: "fst", name: "Components F/S/T", w: 70, field: "prods" },
  { key: "glasstot", name: "Glass units", w: 44, field: "glass", kind: "num" },
  { key: "date:sold", name: "Sold", w: 50, field: "dates", kind: "date" },
  { key: "date:ready", name: "Ready to print", w: 50, field: "dates", kind: "date" },
  { key: "date:floor", name: "Sent to floor", w: 50, field: "dates", kind: "date" },
  { key: "ncomments", name: "Comments", w: 36, field: "comments", kind: "num" }
];
const XP_TABLE_MIN = 20;          // below this a column cannot be read at all
function exportTableColumns(fields) {
  const F = xpFieldSet(fields);
  return XP_TABLE_COLS.filter(c => F[c.field]).map(c => Object.assign({}, c));
}
/** The widths that fit, or an explanation of why they do not. */
function exportTableWidths(cols) {
  const pad = 8, usable = XP_LAND_W - cols.length * pad;
  const sum = cols.reduce((a, c) => a + c.w, 0);
  const scale = sum > usable ? usable / sum : 1;
  const widths = cols.map(c => Math.floor(c.w * scale * 10) / 10);
  if (widths.some(w => w < XP_TABLE_MIN)) {
    throw new Error("the table layout cannot fit " + cols.length +
      " columns across a page - untick some fields or use the job-card layout");
  }
  return widths;
}

/** The value of one column for one row, plus the colour it should carry.
    Returns { v, text, fill, bold, colour } - `v` typed for Excel (Date /
    number / string), `text` the same thing for the PDF. */
function exportCell(row, col) {
  const out = { v: null, text: "", fill: "", colour: "", bold: false };
  const k = col.key;
  if (col.kind === "date") {
    const raw = (row.dates || {})[k.slice(5)] || null;
    /* a real date becomes a real date cell; anything else the office wrote in
       there is kept as the text it is, rather than thrown away as a blank */
    out.v = xpIsIso(raw) ? xpDateCell(raw) : (raw || null);
    out.text = xpIsIso(raw) ? xpNiceDate(raw) : xpStr(raw);
    return out;
  }
  if (col.kind === "prod") {
    const p = (row.prods || []).find(x => x.n === col.prod);
    const n = p ? (p[col.sub] || 0) : 0;
    out.v = n || null; out.text = n ? String(n) : "";
    const st = p ? (p.st || {})[col.sub] : "";
    out.fill = st === "done" ? XP_GOLD : st === "process" ? XP_YELLOW : "";
    return out;
  }
  if (col.kind === "glass") {
    const g = (row.glass || []).find(x => x.type === col.glass);
    const n = g ? g.qty : 0;
    out.v = n || null; out.text = n ? String(n) : "";
    out.fill = g && g.st === "done" ? XP_GOLD : (g && g.st === "process" ? XP_YELLOW : "");
    return out;
  }
  if (col.kind === "yesno") {
    const on = k === "ready" ? !!row.ready : !!row.urg;
    out.v = on ? "Yes" : "No"; out.text = out.v;
    if (k === "ready" && on) out.fill = XP_GREEN_BG;
    if (k === "urg" && on) { out.colour = XP_RED; out.bold = true; }
    return out;
  }
  if (col.kind === "job") {
    out.v = xpStr(row.id); out.text = out.v;
    if (row.ready) out.fill = XP_GREEN_BG;
    if (row.urg) { out.colour = XP_RED; out.bold = true; }
    return out;
  }
  if (k === "fst") {
    const c = (row.prods || []).reduce((a, p) => ({ f: a.f + p.f, s: a.s + p.s, t: a.t + p.t }), { f: 0, s: 0, t: 0 });
    out.v = c.f + "/" + c.s + "/" + c.t; out.text = out.v;
    return out;
  }
  if (col.kind === "num") {
    const n = k === "alerts" ? (row.alerts || 0)
      : k === "comments" || k === "ncomments" ? ((row.comments || []).length)
      : k === "glasstot" ? ((row.glass || []).reduce((a, g) => a + g.qty, 0))
      : (row[k] || 0);
    out.v = n; out.text = String(n);
    return out;
  }
  if (k === "comment") {
    const t = (row.notes || []).filter(n => n.kind === "comment").map(n => n.text).join(" / ");
    out.v = t || null; out.text = t;
    return out;
  }
  const v = xpStr(row[k]);
  out.v = v || null; out.text = v;
  return out;
}

/** Every comment line for the Comments sheet / the PDF card, sheet notes and
    dashboard comments together, in that order. */
function exportCommentLines(row) {
  return (row.notes || []).map(n => ({ src: n.src, who: "", when: "", text: n.text }))
    .concat((row.comments || []).map(x => ({ src: x.src, who: x.who, when: x.when, text: x.text })));
}

/** Rows grouped for "group by section": [[name, rows], ...]. */
function exportGroups(rows, f) {
  const s = xpFilter(f);
  if (!s.groupBySection) return [["", rows || []]];
  const order = [], by = {};
  (rows || []).forEach(r => {
    const k = xpStr(r.section) || "No section";
    if (!by[k]) { by[k] = []; order.push(k); }
    by[k].push(r);
  });
  return order.map(k => [k, by[k]]);
}

/** Is there anything to build at all? Ticking no fields used to make ExcelJS
    throw on a table with no columns and left the PDF as a cover page with
    nothing behind it, so the window asks this before enabling Download. */
function exportBuildable(rows, fields, format, layout) {
  const F = xpFieldSet(fields);
  if (!(rows || []).length) return { ok: false, why: "no jobs match" };
  if (xpLow(format) === "pdf") {
    if (xpLow(layout) === "table") {
      const cols = exportTableColumns(F);
      if (!cols.length) return { ok: false, why: "no columns for the table - tick a field" };
      try { exportTableWidths(cols); } catch (e) { return { ok: false, why: e.message }; }
      return { ok: true, why: "" };
    }
    return EXPORT_FIELD_KEYS.some(k => F[k]) ? { ok: true, why: "" } : { ok: false, why: "no fields ticked" };
  }
  /* Excel can still be worth making with no Jobs columns at all, as long as one
     of the other sheets has something in it */
  if (exportColumns(rows, F).length || F.cp) return { ok: true, why: "" };
  return { ok: false, why: "no fields ticked" };
}

/* ---------------------------------------------------------------------------
   6. the Excel workbook
   --------------------------------------------------------------------------- */

function xpExcelLib() {
  if (typeof ExcelJS !== "undefined" && ExcelJS) return ExcelJS;
  if (typeof window !== "undefined" && window.ExcelJS) return window.ExcelJS;
  throw new Error("ExcelJS is not loaded");
}
const xpSolid = argb => ({ type: "pattern", pattern: "solid", fgColor: { argb: argb } });
const xpTableName = (base, i) => (xpStr(base).replace(/[^A-Za-z0-9]+/g, "_").replace(/^(\d)/, "T$1") || "Table") + "_" + i;

/** An ExcelJS workbook. opts: { fields, filters, who, when, sections, build,
    company }. Sheets appear only when their field is ticked; Export info always. */
function buildWorkbook(rows, opts) {
  opts = opts || {};
  const F = xpFieldSet(opts.fields), f = xpFilter(opts.filters);
  const Lib = xpExcelLib();
  const wb = new Lib.Workbook();
  wb.creator = "Costello production dashboard";
  wb.created = (opts.when && typeof opts.when !== "string") ? opts.when : new Date();

  /* ---- Jobs ---- */
  const cols = exportColumns(rows, F);
  if (cols.length) {
    /* the header row is only worth freezing when it is at the top: grouped by
       section, row 1 is a section heading and freezing it pins the wrong line */
    const ws = wb.addWorksheet("Jobs", f.groupBySection ? {} : { views: [{ state: "frozen", ySplit: 1 }] });
    const groups = exportGroups(rows, f);
    let at = 1, ti = 0;
    groups.forEach(g => {
      const name = g[0], list = g[1];
      if (name) {
        /* one table per section, with the section's name in bold above it */
        const cell = ws.getCell(at, 1);
        cell.value = name + " — " + list.length + " job" + (list.length === 1 ? "" : "s");
        cell.font = { bold: true, size: 12 };
        at += 1;
      }
      ti++;
      ws.addTable({
        name: xpTableName("Jobs", ti),
        ref: "A" + at,
        headerRow: true,
        style: { theme: "TableStyleMedium2", showRowStripes: true, showFirstColumn: false, showLastColumn: false },
        columns: cols.map(c => ({ name: c.name, filterButton: true })),
        /* an Excel table has to have a body: an export of nothing still opens */
        rows: list.length ? list.map(r => cols.map(c => exportCell(r, c).v)) : [cols.map(() => null)]
      });
      /* the colours the dashboard uses, always sitting next to the value they
         describe: nothing here is colour on its own */
      list.forEach((r, ri) => {
        const row = ws.getRow(at + 1 + ri);
        cols.forEach((c, ci) => {
          const cell = row.getCell(ci + 1), spec = exportCell(r, c);
          if (c.kind === "date" && spec.v instanceof Date) cell.numFmt = "dd-mmm-yyyy";
          if (spec.fill) cell.fill = xpSolid(xpArgb(spec.fill));
          if (spec.colour || spec.bold) cell.font = { bold: !!spec.bold, color: { argb: xpArgb(spec.colour || XP_INK) } };
        });
      });
      at += 1 + list.length + 1;      // header + body + a blank line before the next section
    });
    cols.forEach((c, i) => { ws.getColumn(i + 1).width = c.width; });
  }

  /* ---- Comments ---- */
  if (F.comments || F.notes) {
    const cs = wb.addWorksheet("Comments", { views: [{ state: "frozen", ySplit: 1 }] });
    const body = [];
    (rows || []).forEach(r => exportCommentLines(r).forEach(x =>
      body.push([xpStr(r.id), x.src, x.who, x.when, x.text])));
    cs.addTable({
      name: "Comments_1", ref: "A1", headerRow: true,
      style: { theme: "TableStyleMedium2", showRowStripes: true },
      columns: [{ name: "Job" }, { name: "Source" }, { name: "Who" }, { name: "When" }, { name: "Text" }],
      rows: body.length ? body : [["", "", "", "", ""]]
    });
    [12, 18, 24, 18, 90].forEach((w, i) => { cs.getColumn(i + 1).width = w; });
  }

  /* ---- Checkpoints ---- */
  if (F.cp) {
    const ks = wb.addWorksheet("Checkpoints", { views: [{ state: "frozen", ySplit: 1 }] });
    const body = [];
    (rows || []).forEach(r => (r.cp || []).forEach(g => g.items.forEach(it =>
      body.push([xpStr(r.id), g.label, it.label, it.done == null ? "" : it.done, it.total,
                 XP_STATUS_WORD[it.status] || "not started"]))));
    ks.addTable({
      name: "Checkpoints_1", ref: "A1", headerRow: true,
      style: { theme: "TableStyleMedium2", showRowStripes: true },
      columns: [{ name: "Job" }, { name: "Group" }, { name: "Item" }, { name: "Done" }, { name: "Total" }, { name: "Status" }],
      rows: body.length ? body : [["", "", "", "", "", ""]]
    });
    body.forEach((line, i) => {
      const cell = ks.getRow(2 + i).getCell(6);
      const st = line[5];
      if (st === XP_STATUS_WORD.done) cell.fill = xpSolid(XL_GOLD);
      else if (st === XP_STATUS_WORD.process) { cell.fill = xpSolid(XL_YELLOW); cell.font = { color: { argb: xpArgb(XP_INK) } }; }
    });
    [12, 20, 20, 9, 9, 16].forEach((w, i) => { ks.getColumn(i + 1).width = w; });
  }

  /* ---- Export info: always, so a printed file can say where it came from ---- */
  const is = wb.addWorksheet("Export info");
  const lines = [["Exported by", xpShortWho(opts.who) || "unknown"],
                 ["Date and time", xpStamp(opts.when)],
                 ["Format", "Excel"],
                 ["Jobs in this file", (rows || []).length],
                 ["Dashboard build", xpStr(opts.build) || "unknown"],
                 ["Fields", EXPORT_FIELDS.filter(p => F[p[0]]).map(p => p[1]).join(", ")]];
  is.getCell("A1").value = xpStr(opts.company) || "Production";
  is.getCell("A1").font = { bold: true, size: 14 };
  is.getCell("A2").value = "Production export";
  is.getCell("A2").font = { color: { argb: xpArgb(XP_GREY) } };
  let r0 = 4;
  lines.forEach(p => {
    is.getCell(r0, 1).value = p[0]; is.getCell(r0, 1).font = { bold: true };
    is.getCell(r0, 2).value = p[1];
    r0++;
  });
  r0++;
  is.getCell(r0, 1).value = "Filters"; is.getCell(r0, 1).font = { bold: true };
  r0++;
  filtersSummaryLines(f, opts.sections).forEach(line => { is.getCell(r0, 2).value = line; r0++; });
  r0++;
  /* worded without naming the two columns, so the standing test can assert
     that neither word appears in any header this file writes */
  is.getCell(r0, 2).value = "No contact details are included in this file.";
  is.getCell(r0, 2).font = { color: { argb: xpArgb(XP_GREY) }, italic: true };
  is.getColumn(1).width = 22; is.getColumn(2).width = 82;
  return wb;
}

/* ---------------------------------------------------------------------------
   7. the PDF document definition (plain data, apart from header/footer)
   --------------------------------------------------------------------------- */

const XP_BAR_W = 300;             // the width a stacked F/S/T bar is drawn in
const XP_BAR_H = 13;
const XP_MIN_SEG = 16;            // a segment narrower than this cannot hold its number

/** Segment widths for a stacked bar: proportional to the values, but never so
    narrow that the number printed under it has nowhere to sit. Zero-valued
    segments are dropped rather than drawn as a sliver. */
function xpBarSegments(values, width, minSeg) {
  const w = width || XP_BAR_W, min = minSeg == null ? XP_MIN_SEG : minSeg;
  const live = values.filter(v => v.n > 0);
  const total = live.reduce((a, v) => a + v.n, 0);
  if (!live.length || !total) return [];
  let segs = live.map(v => ({ key: v.key, n: v.n, colour: v.colour, w: w * v.n / total }));
  const small = segs.filter(s => s.w < min), big = segs.filter(s => s.w >= min);
  if (small.length && big.length) {
    const owed = small.reduce((a, s) => a + (min - s.w), 0);
    const pool = big.reduce((a, s) => a + s.w, 0);
    small.forEach(s => { s.w = min; });
    big.forEach(s => { s.w = Math.max(min, s.w - owed * (s.w / pool)); });
  } else if (small.length) {
    segs.forEach(s => { s.w = w / segs.length; });
  }
  return segs;
}

const xpCpColour = st => st === "done" ? XP_GOLD : st === "process" ? XP_YELLOW : XP_LINE;
/* yellow is far too bright to put white text on; the dashboard uses dark ink
   on it everywhere and so does the PDF */
const xpCpInk = st => st === "process" ? XP_INK : (st === "done" ? XP_INK_2 : XP_GREY);

function xpTile(n, label, colour) {
  return {
    width: "*",
    table: { widths: ["*"], body: [[{
      stack: [{ text: String(n), fontSize: 26, bold: true, color: colour || XP_INK, margin: [0, 2, 0, 0] },
               { text: label, fontSize: 8, color: XP_GREY, characterSpacing: 0.6 }],
      fillColor: XP_PAPER, margin: [10, 8, 10, 8], alignment: "center"
    }]] },
    layout: "noBorders"
  };
}

/** One job, as a run of content nodes.

    Only the head - the job number, the customer, the badges and the three big
    numbers - is `unbreakable`, and it is small enough to fit on any page.
    Everything after it flows and may break wherever it needs to. That matters:
    a job with twenty products and thirty comments is taller than an A4 page,
    and pdfmake silently drops an unbreakable block it cannot fit anywhere,
    leaving an empty page where the job should have been. Jobs are separated by
    a rule and some air instead of by a forced page break. */
function xpCardNodes(row, F, first) {
  const out = [];
  if (!first) out.push({
    canvas: [{ type: "line", x1: 0, y1: 0, x2: XP_PAGE_W - 4, y2: 0, lineWidth: 0.8, lineColor: XP_LINE }],
    margin: [0, 16, 0, 12]
  });

  /* --- the head band: never split --- */
  const head = [];
  const badge = (text, bg, ink) => ({ text: text, fontSize: 7.5, bold: true, color: ink || "#FFFFFF", background: bg });
  const top = [];
  if (F.job) top.push({ width: "auto", text: xpStr(row.id), fontSize: 20, bold: true,
                        color: row.urg ? XP_RED : XP_INK, margin: [0, 0, 10, 0] });
  const who = [];
  if (F.cust) who.push({ text: xpStr(row.cust) || "—", fontSize: 11, bold: true });
  const sub = [];
  if (F.county && row.area) sub.push(xpStr(row.area));
  if (F.section && row.section) sub.push(xpStr(row.section));
  if (sub.length) who.push({ text: sub.join("  ·  "), fontSize: 8, color: XP_GREY });
  if (who.length) top.push({ width: "*", stack: who });
  const flags = [];
  if (F.ready) flags.push(row.ready ? badge("  READY TO DELIVER  ", XP_GREEN) : badge("  IN PRODUCTION  ", XP_GREY));
  if (F.urgent && row.urg) flags.push(badge("  URGENT  ", XP_RED));
  if (flags.length) top.push({ width: "auto", stack: flags, alignment: "right" });
  if (top.length) head.push({ columns: top, columnGap: 6 });

  const tiles = [];
  if (F.wnd) tiles.push(xpTile(row.wnd || 0, "WINDOWS", XP_INK));
  if (F.drs) tiles.push(xpTile(row.drs || 0, "DOORS", XP_INK));
  if (F.prods) tiles.push(xpTile((row.prods || []).reduce((a, p) => a + p.f + p.s + p.t, 0), "COMPONENTS F+S+T", XP_INK));
  if (tiles.length) {
    head.push({ canvas: [{ type: "line", x1: 0, y1: 4, x2: XP_HEAD_W, y2: 4, lineWidth: 0.7, lineColor: XP_LINE }], margin: [0, 2, 0, 6] });
    head.push({ columns: tiles, columnGap: 6 });
  }
  if (head.length) out.push({
    unbreakable: true, margin: [0, 0, 0, 8],
    table: { widths: ["*"], body: [[{ stack: head, fillColor: XP_PAPER, margin: [10, 9, 10, 10] }]] },
    layout: "noBorders"
  });

  /* --- products: one stacked bar per product, every segment numbered --- */
  if (F.prods && (row.prods || []).length) {
    out.push({ text: "Products", fontSize: 7.5, color: XP_GREY, bold: true, margin: [0, 2, 0, 3] });
    row.prods.forEach(p => {
      const segs = xpBarSegments([
        { key: "f", n: p.f, colour: XP_F }, { key: "s", n: p.s, colour: XP_S }, { key: "t", n: p.t, colour: XP_T }
      ], XP_BAR_W);
      let x = 0;
      const shapes = segs.map(s => { const node = { type: "rect", x: x, y: 0, w: s.w, h: XP_BAR_H, color: s.colour }; x += s.w; return node; });
      const labels = segs.map(s => ({ width: s.w, text: XP_SUBLABEL[s.key] + " " + s.n, fontSize: 6.5, color: XP_GREY }));
      out.push({
        columns: [
          { width: 110, text: xpUpper(p.n), fontSize: 8, margin: [0, 1, 0, 0] },
          { width: "auto", stack: [
            { canvas: shapes.length ? shapes : [{ type: "rect", x: 0, y: 0, w: XP_BAR_W, h: XP_BAR_H, color: XP_LINE }] },
            { columns: labels.length ? labels : [{ width: "*", text: "none", fontSize: 6.5, color: XP_GREY }], columnGap: 0, margin: [0, 1, 0, 0] }
          ] },
          { width: "*", text: (p.f + p.s + p.t) + " total", fontSize: 8, alignment: "right", margin: [0, 1, 0, 0] }
        ], columnGap: 8, margin: [0, 0, 0, 3]
      });
    });
  }

  /* --- glass: a block per type, sized by quantity, labelled with it --- */
  if (F.glass && (row.glass || []).length) {
    const maxQty = row.glass.reduce((a, g) => Math.max(a, g.qty), 0) || 1;
    out.push({ text: "Glass units", fontSize: 7.5, color: XP_GREY, bold: true, margin: [0, 3, 0, 3] });
    out.push({
      columns: row.glass.map(g => {
        const w = 26 + 54 * (g.qty / maxQty);
        return { width: w + 8, stack: [
          { canvas: [{ type: "rect", x: 0, y: 0, w: w, h: 20, r: 2, color: xpCpColour(g.st) }] },
          { text: xpUpper(g.type) + " " + g.qty, fontSize: 7.5, bold: true, margin: [0, 2, 0, 0] },
          { text: XP_STATUS_WORD[g.st] || "not started", fontSize: 6, color: XP_GREY }
        ] };
      }).concat([{ width: "*", text: "" }]), columnGap: 0, margin: [0, 0, 0, 4]
    });
  }

  /* --- checkpoints: a progress bar per group with "done of total" --- */
  if (F.cp && (row.cp || []).length) {
    out.push({ text: "Checkpoints", fontSize: 7.5, color: XP_GREY, bold: true, margin: [0, 3, 0, 3] });
    row.cp.forEach(g => {
      const w = 220, pct = g.total ? Math.max(0, Math.min(1, g.done / g.total)) : 0;
      const done = g.unknown && !g.done ? "in progress" : g.done + " of " + g.total;
      out.push({
        columns: [
          { width: 110, text: g.label, fontSize: 8 },
          { width: "auto", canvas: [
            { type: "rect", x: 0, y: 2, w: w, h: 8, r: 2, color: XP_LINE },
            { type: "rect", x: 0, y: 2, w: Math.max(pct * w, pct > 0 ? 2 : 0), h: 8, r: 2, color: xpCpColour(g.status) }
          ] },
          { width: "*", text: done + "  ·  " + (XP_STATUS_WORD[g.status] || "not started"),
            fontSize: 7.5, color: xpCpInk(g.status), margin: [6, 0, 0, 0] }
        ], columnGap: 6, margin: [0, 0, 0, 2]
      });
    });
  }

  /* --- the five dates as a timeline: a filled dot means it happened --- */
  if (F.dates) {
    out.push({ text: "Dates", fontSize: 7.5, color: XP_GREY, bold: true, margin: [0, 4, 0, 3] });
    out.push({
      columns: XP_DATE_STEPS.map(p => {
        const raw = (row.dates || {})[p[0]], real = xpIsIso(raw);
        return { width: "*", stack: [
          { canvas: [
            { type: "line", x1: 0, y1: 5, x2: 92, y2: 5, lineWidth: 1.4, lineColor: real ? XP_GREEN : XP_LINE },
            { type: "ellipse", x: 5, y: 5, r1: 4, r2: 4, color: real ? XP_GREEN : XP_LINE }
          ] },
          { text: p[1], fontSize: 6, color: XP_GREY, margin: [0, 2, 0, 0] },
          /* text the office wrote in a date cell is shown as written */
          { text: real ? xpNiceDate(raw) : (raw ? xpStr(raw) : "not yet"), fontSize: 7.5,
            bold: real, color: real ? XP_INK : XP_GREY }
        ] };
      }), columnGap: 2, margin: [0, 0, 0, 4]
    });
  }

  /* --- colour, office number, alerts --- */
  const facts = [];
  if (F.colour && row.colour) facts.push("Windows colour: " + xpStr(row.colour));
  if (F.off && row.off) facts.push("Office no: " + xpStr(row.off));
  if (F.alerts) facts.push("Email alerts: " + (row.alerts || 0));
  if (facts.length) out.push({ text: facts.join("   ·   "), fontSize: 7.5, color: XP_GREY, margin: [0, 2, 0, 3] });

  /* --- comments and sheet notes: all of them, however many there are --- */
  const lines = exportCommentLines(row);
  if ((F.comments || F.notes) && lines.length) {
    out.push({ text: "Comments", fontSize: 7.5, color: XP_GREY, bold: true, margin: [0, 2, 0, 3] });
    lines.forEach(x => {
      const meta = [x.src, x.who, x.when].filter(Boolean).join(" · ");
      out.push({ text: [{ text: meta + "  ", fontSize: 6.5, color: XP_GREY }, { text: xpStr(x.text), fontSize: 8 }], margin: [0, 0, 0, 2] });
    });
  }
  return out;
}

/** The landscape one-row-per-job layout, on the short column set. Zebra shading
    is set per cell so the definition stays plain data - no layout function to
    inspect around. */
function xpTableDoc(rows, F, f) {
  const cols = exportTableColumns(F);
  if (!cols.length) throw new Error("no fields are ticked for the table layout");
  const widths = exportTableWidths(cols);
  const head = cols.map(c => ({ text: c.name, bold: true, fontSize: 8, color: "#FFFFFF", fillColor: XP_INK }));
  const body = [head];
  let i = 0;
  exportGroups(rows, f).forEach(g => {
    if (g[0]) {
      const cell = { text: g[0] + " — " + g[1].length + " job" + (g[1].length === 1 ? "" : "s"),
                     bold: true, fontSize: 9, fillColor: "#EFECE7", colSpan: cols.length, margin: [2, 3, 2, 3] };
      body.push([cell].concat(cols.slice(1).map(() => ({ text: "" }))));
    }
    g[1].forEach(r => {
      i++;
      const zebra = i % 2 === 0 ? XP_PAPER : null;
      body.push(cols.map(c => {
        const spec = exportCell(r, c);
        const cell = { text: spec.text, fontSize: 8 };
        if (spec.fill) cell.fillColor = spec.fill;
        else if (zebra) cell.fillColor = zebra;
        if (spec.colour) cell.color = spec.colour;
        if (spec.bold) cell.bold = true;
        return cell;
      }));
    });
  });
  return [{
    table: { headerRows: 1, dontBreakRows: true, widths: widths, body: body },
    layout: "noBorders", fontSize: 8
  }];
}

/** A pdfmake document definition. opts: { layout:"cards"|"table", fields,
    filters, who, when, company, logo (data URL or null), sections, build }.
    Everything is plain data except header and footer, which pdfmake requires
    to be functions and which return plain content. */
function buildDocDefinition(rows, opts) {
  opts = opts || {};
  const F = xpFieldSet(opts.fields), f = xpFilter(opts.filters);
  const list = rows || [];
  const company = xpStr(opts.company) || "Production";
  const when = opts.when || new Date();
  const dateText = xpNiceDate(xpIsoDate(when));
  const who = xpShortWho(opts.who) || "unknown";
  const cards = xpLow(opts.layout) !== "table";
  const logo = xpStr(opts.logo);

  /* ---- cover ---- */
  const content = [];
  content.push({ text: "Production export", fontSize: 24, bold: true, margin: [0, 6, 0, 0] });
  content.push({ text: company + "  ·  " + dateText + "  ·  exported by " + who,
                 fontSize: 9, color: XP_GREY, margin: [0, 2, 0, 12] });

  const tiles = [xpTile(list.length, "JOBS", XP_INK)];
  if (F.wnd) tiles.push(xpTile(list.reduce((a, r) => a + (r.wnd || 0), 0), "WINDOWS", XP_INK));
  if (F.drs) tiles.push(xpTile(list.reduce((a, r) => a + (r.drs || 0), 0), "DOORS", XP_INK));
  if (F.prods) tiles.push(xpTile(list.reduce((a, r) => a + (r.prods || []).reduce((b, p) => b + p.f + p.s + p.t, 0), 0),
                                 "COMPONENTS F+S+T", XP_INK));
  if (F.ready) tiles.push(xpTile(list.filter(r => r.ready).length, "READY TO DELIVER", XP_GREEN));
  if (F.urgent) tiles.push(xpTile(list.filter(r => r.urg).length, "URGENT", XP_RED));
  for (let i = 0; i < tiles.length; i += 3) {
    const line = tiles.slice(i, i + 3);
    while (line.length < 3) line.push({ width: "*", text: "" });     // keep the widths even
    content.push({ columns: line, columnGap: 6, margin: [0, 0, 0, 6] });
  }
  content[content.length - 1].margin = [0, 0, 0, 14];

  content.push({ text: "What is in this file", fontSize: 8, bold: true, color: XP_GREY, margin: [0, 0, 0, 4] });
  content.push({ ul: filtersSummaryLines(f, opts.sections), fontSize: 9, margin: [0, 0, 0, 12] });

  if (F.section) {
    const order = [], by = {};
    list.forEach(r => { const k = xpStr(r.section) || "No section";
      if (!by[k]) { by[k] = 0; order.push(k); } by[k]++; });
    if (order.length) {
      content.push({ text: "Jobs by section", fontSize: 8, bold: true, color: XP_GREY, margin: [0, 0, 0, 4] });
      content.push({
        table: { headerRows: 1, widths: ["*", 60],
          body: [[{ text: "Section", bold: true, fontSize: 8 }, { text: "Jobs", bold: true, fontSize: 8, alignment: "right" }]]
            .concat(order.map(k => [{ text: k, fontSize: 9 }, { text: String(by[k]), fontSize: 9, alignment: "right" }])) },
        layout: "lightHorizontalLines", margin: [0, 0, 0, 12]
      });
    }
  }
  content.push({ text: "No contact details are included in this file.", fontSize: 7.5, color: XP_GREY, italics: true });

  /* ---- the jobs ---- */
  if (cards) {
    list.forEach((r, i) => {
      const nodes = xpCardNodes(r, F, i === 0);
      /* the only page break in the whole run: the cover keeps its own page */
      if (i === 0 && nodes.length) nodes[0] = Object.assign({ pageBreak: "before" }, nodes[0]);
      nodes.forEach(nd => content.push(nd));
    });
  } else if (list.length) {
    xpTableDoc(list, F, f).forEach((node, i) => {
      if (i === 0) node.pageBreak = "before";
      content.push(node);
    });
  }

  const def = {
    pageSize: "A4",
    pageOrientation: cards ? "portrait" : "landscape",
    pageMargins: [28, logo ? 58 : 50, 28, 34],
    info: { title: "Production export " + xpIsoDate(when), author: company },
    defaultStyle: { font: "Roboto", fontSize: 9, color: XP_INK },
    content: content,
    /* pdfmake's own two callbacks; both return plain content nodes */
    header: function (currentPage, pageCount) {
      const bar = [];
      /* the logo slot: filled when app.js found assets/logo.png, and simply
         absent when it did not - the name then stands on its own */
      if (logo) bar.push({ width: "auto", image: logo, fit: [96, 26], margin: [0, 0, 8, 0] });
      bar.push({ width: "auto", text: company, fontSize: 12, bold: true, margin: [0, logo ? 6 : 2, 0, 0] });
      bar.push({ width: "*", text: "Production export", fontSize: 9, color: XP_GREY, alignment: "center", margin: [0, 6, 0, 0] });
      bar.push({ width: "auto", text: dateText, fontSize: 9, color: XP_GREY, alignment: "right", margin: [0, 6, 0, 0] });
      return { margin: [28, 16, 28, 0], columns: bar, columnGap: 6 };
    },
    footer: function (currentPage, pageCount) {
      return {
        margin: [28, 8, 28, 0],
        columns: [
          { width: "*", text: "Exported by " + who, fontSize: 7.5, color: XP_GREY },
          { width: "auto", text: "Page " + currentPage + " of " + pageCount, fontSize: 7.5, color: XP_GREY, alignment: "right" }
        ]
      };
    }
  };
  if (opts.compress === false) def.compress = false;    // tests read the text streams
  return def;
}

/* ---------------------------------------------------------------------------
   8. handing the file over (browser only)
   --------------------------------------------------------------------------- */

/** An anchor with an object URL, clicked and then cleaned up. No network, no
    server, no third party: the bytes were made in this tab. */
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { try { a.remove(); } catch (e) {} URL.revokeObjectURL(url); }, 4000);
  return name;
}

/** vendor/vfs_fonts.js hands its Roboto files to pdfMake itself the moment it
    loads, as long as pdfmake.min.js loaded first - app.js injects them in that
    order the first time somebody exports a PDF, so normally there is nothing
    to do here beyond checking they arrived.

    pdfmake 0.2.23 keeps both the font list and the file system in closures, so
    neither can be read back off pdfMake to check: addVirtualFileSystem() is
    therefore called again (harmless - it is the same object), and pdfMake.fonts
    is filled in because createPdf falls back to it when setFonts() was never
    called. Naming Roboto here also means a different font file dropped into
    vendor/ later cannot silently change the typeface. */
function xpPdfReady() {
  if (typeof pdfMake === "undefined" || !pdfMake) throw new Error("pdfmake is not loaded");
  if (typeof vfs !== "undefined" && vfs && typeof pdfMake.addVirtualFileSystem === "function") {
    pdfMake.addVirtualFileSystem(vfs);
  }
  if (!pdfMake.fonts) {
    pdfMake.fonts = { Roboto: { normal: "Roboto-Regular.ttf", bold: "Roboto-Medium.ttf",
                                italics: "Roboto-Italic.ttf", bolditalics: "Roboto-MediumItalic.ttf" } };
  }
  return pdfMake;
}

const XP_API = {
  EXPORT_FIELDS, EXPORT_FIELD_KEYS, XP_SORTS, XP_DATE_STEPS, XP_PRESET_KEY, XP_STATUS_WORD,
  XP_TABLE_COLS, XP_TABLE_MIN, XP_LAND_W, XP_PAGE_W,
  exportAllFields, exportDefaults, exportScope, exportFilter, exportSort, exportJobs,
  exportRows, exportCheckpoints, exportColumns, exportTableColumns, exportTableWidths,
  exportCell, exportCommentLines, exportGroups, exportBuildable,
  exportFilename, exportScopeLabel, exportLogFrom, filtersSummary, filtersSummaryLines,
  buildWorkbook, buildDocDefinition, downloadBlob, xpBarSegments, xpPdfReady, xpCardNodes,
  presetsLoad, presetSave, presetDelete, xpFieldSet, xpFilter, xpIsoDate, xpNiceDate, xpStamp,
  xpShortWho, xpIsIso, xpSectionNames,
  colours: { gold: XP_GOLD, yellow: XP_YELLOW, green: XP_GREEN, red: XP_RED, grey: XP_GREY }
};
if (typeof window !== "undefined") window.XP = XP_API;
if (typeof module !== "undefined" && module.exports) module.exports = XP_API;
