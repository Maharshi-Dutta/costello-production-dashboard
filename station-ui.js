/* The floor tablet's SHELL - the parts of a station page that have nothing to
   do with which station it is.

   New on 2026-09-16 with the welding station. Every floor page needs the same
   five things and none of them is about glass or welding: the device's own
   light/dark choice, "how long ago was that" in words, the sign-in gate, the
   "Who are you?" picker with its PIN pad, and one HTML escape. Written once
   here so the third station (PA Lam) is a definition and a renderer and not a
   third copy of a page.

   NOTHING IN HERE TOUCHES A WORKBOOK. Apart from the end-of-day sheet
   (2026-09-28), which reads and appends to the floor's own day-sheet list,
   and the read-only list reader and the poll clock at the bottom
   (2026-10-02), it is strings, a localStorage key for a theme, and two wiring
   helpers; nothing else here touches Graph or a list. Every identifier is
   prefixed `stu`/`STU` so that a page which loads this file beside another
   station's script cannot collide with it - browser scripts share one global
   lexical scope, and a collision there is a page that does not start.

   A NOTE ON glass.html: station.js calls into this file for the end-of-day
   sheet only (since 2026-09-28); its theme, gate and picker are still its
   own. The test harnesses load this file into station.js's vm context ahead
   of it, as the page does. See docs/STATIONS.md, "Adding a station".       */

const STU_THEME_KEY = "cw_stationtheme";      // this device's choice, not this person's

const stuEsc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

/** "just now" / "40 s ago" / "12 min ago" / "3 h ago", from a millisecond
    stamp. Short on purpose: it sits in a header read at arm's length. */
function stuAgo(at) {
  const t = Number(at);
  if (!isFinite(t) || !t) return "";
  const s = Math.max(0, Math.round((Date.now() - t) / 1000));
  if (s < 10) return "just now";
  if (s < 90) return s + " s ago";
  if (s < 5400) return Math.round(s / 60) + " min ago";
  return Math.round(s / 3600) + " h ago";
}

/* ---- the device's own theme -------------------------------------------------
   The owner asked for dark by default on a tablet - a workshop screen at arm's
   length - with a switch for whoever prefers the other one. It belongs to the
   device and not to the person, so it lives in localStorage and survives
   somebody switching person, signing out or the page reloading itself. */
function stuThemeNow() {
  try { return localStorage.getItem(STU_THEME_KEY) === "light" ? "light" : "dark"; }
  catch (e) { return "dark"; }
}
/** Set the theme, remember it, and put the right word on the button. */
function stuApplyTheme(t, btn) {
  const dark = t !== "light";
  try { document.documentElement.dataset.theme = dark ? "dark" : "light"; } catch (e) {}
  try { localStorage.setItem(STU_THEME_KEY, dark ? "dark" : "light"); } catch (e) {}
  const b = btn || document.querySelector("#themebtn");
  if (b) b.textContent = dark ? "Light" : "Dark";
  return dark ? "dark" : "light";
}

/* ---- the sign-in gate -------------------------------------------------------
   Three nodes every station page has: #gate (signed out), #top (the header)
   and #main (the board). One of the first two is on screen and never both. */
function stuGate(on, err) {
  const q = s => document.querySelector(s);
  const set = (sel, hide, disp) => {
    const el = q(sel);
    if (!el) return;
    el.hidden = hide;
    if (el.style) el.style.display = hide ? "none" : disp;
  };
  const g = q("#gate");
  if (g) { g.hidden = !on; if (g.style) g.style.display = on ? "flex" : "none"; }
  set("#top", on, "flex");
  set("#main", on, "block");
  const e = q("#gateerr");
  if (e) {
    if (e.style) e.style.display = err ? "block" : "none";
    e.textContent = err || "";
  }
}

/* ---- "Who are you?" ---------------------------------------------------------
   Full screen, one big button per person, and a numeric pad for anybody whose
   PIN column is filled in. A wrong PIN shakes and says "Try again": there is no
   lockout, because locking a shared tablet out of the only screen the floor has
   would stop the work rather than protect it.

   Station-neutral in two places, and they are the two that used to be glass's:
   `empty` is the sentence shown when nobody is set up for THIS station, and
   `labelOf` turns a stage key into the word this station calls it. A station
   with one stage passes a labelOf that answers one word.                    */
function stuPickerHtml(o) {
  o = o || {};
  const people = o.people || [];
  const labelOf = typeof o.labelOf === "function" ? o.labelOf : (k => String(k));
  if (o.pinFor) {
    const dots = String(o.pinTyped || "").replace(/./g, "•");
    const key = t => '<button class="pk" data-pin="' + stuEsc(t) + '">' + stuEsc(t) + '</button>';
    return '<div class="picker">' +
      '<div class="pickh">' + stuEsc(o.pinFor.name) + '</div>' +
      '<div class="picksub">Enter your PIN</div>' +
      '<div class="pindots' + (o.pinBad ? " shake" : "") + '">' + stuEsc(dots || "····") + '</div>' +
      (o.pinBad ? '<div class="pinbad">Try again</div>' : "") +
      '<div class="pinpad">' + ["1", "2", "3", "4", "5", "6", "7", "8", "9"].map(key).join("") +
        '<button class="pk" data-pin="back">←</button>' + key("0") +
        '<button class="pk go" data-pin="ok">OK</button></div>' +
      '<button class="pcancel" data-pin="cancel">Back to the names</button>' +
      '</div>';
  }
  if (!people.length)
    return '<div class="picker"><div class="pickh">Who are you?</div>' +
      '<div class="msg">' + stuEsc(o.empty || "Nobody is set up for this station yet. " +
        "Ask the office to add you to the “Station people” list.") + '</div></div>';
  return '<div class="picker">' +
    '<div class="pickh">Who are you?</div>' +
    '<div class="pgrid">' + people.map(p =>
      '<button class="pbtn" data-person="' + stuEsc(p.name) + '">' +
        '<span class="pname">' + stuEsc(p.name) + '</span>' +
        '<span class="pstages">' + stuEsc((p.stages || []).length
          ? p.stages.map(labelOf).join(" · ") : "no stages yet") + '</span>' +
      '</button>').join("") + '</div></div>';
}

/** Wire the picker that stuPickerHtml just drew. `onPerson` gets the name that
    was tapped; `onKey` gets one of the pad's keys ("0".."9", "back", "ok",
    "cancel"). Neither knows anything about a list. */
function stuWirePicker(host, o) {
  if (!host || !host.querySelectorAll) return;
  o = o || {};
  host.querySelectorAll("[data-person]").forEach(el => {
    el.onclick = () => { if (typeof o.onPerson === "function") o.onPerson(el.dataset.person); };
  });
  host.querySelectorAll("[data-pin]").forEach(el => {
    el.onclick = () => { if (typeof o.onKey === "function") o.onKey(el.dataset.pin); };
  });
}

/* ---- the end-of-day sheet ---------------------------------------------------
   Shipped on the glass cutting page 2026-09-21 (docs/specs/2026-09-21-day-
   sheets-and-station-reports.md); MOVED HERE FROM station.js on 2026-09-28 so
   the welding page and the hotmelting page have the same sheet without a copy
   of it (docs/specs/2026-09-28-day-sheets-welding-hotmelt-floor-log.md).

   The person's paper sheet, on the tablet: their name, the day, the counts the
   STATION DEFINITION names for this stage (`def.daySheets[stage]`), a line
   about anything that got in the way, and - for a stage that has one - this
   week against the office's target. A stage with no sheet gets no button,
   reads neither list and sends no request for either.

   THE ONE PART OF THIS FILE THAT TALKS TO A LIST. It reads `Station day
   sheets` (and `Station targets`, only for a stage with a target; and, only
   for a stage whose sheet shows the tablet's own count, `Station log`) and
   makes ONE POST per person per stage per day to `Station day sheets`. No
   PATCH, no DELETE, nothing else written, and nothing near a workbook: once
   saved, a mistake is the office's to correct.

   One controller per page, made by stuDaySheet(cfg):
     def        the station definition (ST.GLASS, WELDC.WELD)
     stage()    this page's stage key
     who()      the person signed in, "" for nobody
     siteId()   the site the station's lists live in, null until resolved
     touch()    somebody is working the screen (the page's own stamp)
     render()   redraw the page
     flush()    set the page's own write queue going (it calls .flush())
     keys       { q, draft } - localStorage keys, one pair PER PAGE so two
                pages on one tablet never share a draft or an owed sheet
     stageWords()   the stage in words, for the sheet's heading
     pendingLog()   optional: this tablet's own log lines not yet sent, as
                    `Station log` field bags - counted with the rest
     tag        the console prefix                                          */
const STU_DAY_REFUSE_MAX = 3;
function stuDaySheet(cfg) {
  const def = cfg.def;
  const tag = cfg.tag || "[station]";
  const q$ = s => document.querySelector(s);
  const D = {
    shown: false,        // the sheet is on screen, in the board's place
    rows: [],            // this stage's saved sheets, as last read
    ok: null,            // null not looked · false missing or unreachable · true read it
    why: "",
    /* "there is no such list" and "I could not reach the list" are two
       different answers and only the first is a reason to refuse a save
       (review, 2026-09-21). `ok` is false for both; this says which. */
    missing: false,
    target: null,        // the weekly target in force, or null
    draft: null,         // { day, stage, who, counts, note, target } typed but not saved
    bad: "",             // what is wrong with what is typed, in words
    q: {},               // saved sheets this tablet still owes
    logItems: null,      // Station log, for a stage that counts its own taps
    logToken: null,
    logState: ""         // "" not asked · "reading" · "ok" · "failed" - for this opening
  };
  const esc = stuEsc;
  const stage = () => String(cfg.stage() || "");
  const who = () => String(cfg.who() || "");
  const siteId = () => cfg.siteId();
  D.sheet = () => (typeof ST.daySheetOf === "function" ? ST.daySheetOf(def, stage()) : null);
  const hasTarget = () => !!D.sheet() && D.sheet().target !== false;
  const counts = () => (D.sheet() || { counts: [] }).counts;
  const dayOpts = () => ({ siteId: siteId(), fields: ST.dayFieldsFor(def, stage()) });
  const today = () => ST.dayKey(new Date());

  /** Whatever is in storage, rebuilt through the row builder - so an edited
      localStorage can put no column on the wire that a save could not. */
  function cleanQ(raw) {
    const out = {};
    Object.keys(raw || {}).forEach(k => {
      const e = raw[k], f = (e && e.fields) || null;
      if (!f || !f.Title) return;
      const ds = ST.daySheetOf(def, f.Stage);
      if (!ds) return;
      const c = {};
      ds.counts.forEach(x => { c[x[0]] = f[x[0]]; });
      const rebuilt = ST.dayFields(def, f.Stage, { day: f.Day, who: f.Who, counts: c, note: f.Note,
        weekTarget: f.WeekTarget, counted: f[ST.DAY_COUNTED_FIELD], at: f.SavedAt });
      /* how often SharePoint has refused it survives the reload with it */
      if (rebuilt) out[rebuilt.Title] = { fields: rebuilt, err: 0,
                                          refused: Math.max(0, Math.round(Number(e.refused) || 0)) };
    });
    return out;
  }
  try { D.q = cleanQ(JSON.parse(localStorage.getItem(cfg.keys.q) || "{}")); } catch (e) { D.q = {}; }
  D.persist = () => { try { localStorage.setItem(cfg.keys.q, JSON.stringify(D.q)); } catch (e) {} };
  D.owing = () => Object.keys(D.q).length;

  /** The day sheets (and the target, and the log, where this stage has them),
      quietly. A missing list is a state, never an error: the board must not be
      takeable away by a sheet nobody has opened. */
  D.read = async function () {
    if (!D.sheet() || !siteId()) return false;
    let got = false;
    try {
      const items = await CW.listItems(ST.DAY_LIST, dayOpts());
      if (items == null) { D.ok = false; D.missing = true; D.why = ST.DAY_MISSING_FLOOR; return false; }
      D.rows = ST.dayRows(items, counts(), { station: def.name, stage: stage() });
      D.ok = true; D.missing = false; D.why = ""; got = true;
    } catch (e) {
      /* a throw is the wifi, a bad gateway, a refused token - never "there is
         no such list". `missing` is left alone: a save must still be queued
         through a bad afternoon. */
      if (D.ok !== true) { D.ok = false; D.why = ST.DAY_UNREACHABLE; }
      console.warn(tag + " the day sheets could not be read:", (e && e.message) || e);
      return false;
    }
    /* the target is a nicety, and only for a stage that has one: a stage with
       `target: false` never asks for the list, so a site without it says
       nothing about it */
    if (hasTarget()) {
      try {
        const t = await CW.listItems(ST.TARGET_LIST, { siteId: siteId(), fields: ST.TARGET_FIELDS });
        const hit = t == null ? null : ST.targetOf(t, def.name, stage());
        D.target = hit ? hit.target : null;
      } catch (e) { /* keep the last one */ }
    }
    /* NOT the log (review, 2026-09-28): this runs at start and after every
       flush, and a full `Station log` read belongs only to opening the sheet */
    return got;
  };
  /** `Station log`, for the tablet's own count: the delta feed, so opening the
      sheet a second time asks only for what is new. */
  D.readLog = async function () {
    if (!siteId()) return false;
    const opts = { siteId: siteId(), fields: ST.LOG_FIELDS };
    try {
      let d;
      try { d = await CW.listDelta(ST.LOG_LIST, Object.assign({ token: D.logToken }, opts)); }
      catch (e) {
        if (!D.logToken || !CW.isDeltaRestart || !CW.isDeltaRestart(e)) throw e;
        D.logToken = null; D.logItems = null;
        d = await CW.listDelta(ST.LOG_LIST, opts);
      }
      if (d == null) { D.logItems = null; return false; }
      D.logItems = ST.mergeDelta(D.logToken ? D.logItems : [], d.items);
      D.logToken = d.next;
      return true;
    } catch (e) {
      console.warn(tag + " the log could not be read for the day's count:", (e && e.message) || e);
      return false;
    }
  };
  /** What the tablet counted of this person's taps today, or null while the
      log is being read for this opening of the sheet, or could not be read.

      ONLY A LOG READ SINCE THE SHEET OPENED COUNTS (review, 2026-09-28): the
      copy from an earlier opening misses every tap made since, which made
      Counted short and asked a false "you counted N and typed M". */
  D.counted = function (day) {
    const ds = D.sheet();
    if (!ds || !ds.counted || D.logState !== "ok" || D.logItems == null) return null;
    const mine = (typeof cfg.pendingLog === "function" ? cfg.pendingLog() : [])
      .map((f, i) => ({ id: "owed-" + i, fields: f }));
    const rows = ST.logRows(D.logItems.concat(mine), def.name);
    return ST.dayCounted(rows, who(), day || D.sheetDay(), stage());
  };

  /* ---- the draft, and WHICH DAY IT IS ABOUT ---------------------------------
     A DRAFT BELONGS TO THE DAY IT WAS STARTED (review, 2026-09-21), not to
     whatever day it happens to be when Save is tapped: typing at 23:55 and
     tapping Save at 00:01 must not file the evening's work under tomorrow, and
     the same tick must not throw away a draft. So the record carries its own
     `day`, the form shows THAT day, Save files it under it, and it survives
     until the end of the FOLLOWING day. It also carries `who`: a draft is one
     person's writing and the tablet is passed around. The target is stamped
     on at the same moment for the same reason.                             */
  function keep(d) {
    if (!d || d.stage !== stage()) return null;
    if (String(d.who || "") !== who()) return null;
    const day = ST.dayKey(String(d.day || ""));
    if (!day) return null;
    const age = Math.round((Date.parse(today() + "T12:00:00Z") - Date.parse(day + "T12:00:00Z")) / 86400000);
    if (!isFinite(age) || age < 0 || age > 1) return null;
    return { day: day, stage: d.stage, who: String(d.who || ""),
             counts: d.counts || {}, note: String(d.note || ""),
             target: d.target == null ? null : Number(d.target) };
  }
  function loadDraft() {
    let d = null;
    try { d = JSON.parse(localStorage.getItem(cfg.keys.draft) || "null"); } catch (e) { d = null; }
    return keep(d);
  }
  D.draftNow = function () {
    if (!keep(D.draft))
      D.draft = loadDraft() || { day: today(), stage: stage(), who: who(), counts: {}, note: "", target: D.target };
    return D.draft;
  };
  D.saveDraft = () => { try { localStorage.setItem(cfg.keys.draft, JSON.stringify(D.draft)); } catch (e) {} };
  D.clearDraft = () => { D.draft = null; try { localStorage.removeItem(cfg.keys.draft); } catch (e) {} };
  /** The day the sheet on screen is about: the draft's own. */
  D.sheetDay = () => D.draftNow().day;
  /** This person's sheet for that day, if it is already on the list ... */
  D.saved = day => D.rows.find(r => r.day === day && r.who === who()) || null;
  /** ... or still owed by this tablet. */
  D.owed = day => D.q[ST.dayTitle(def.name, stage(), day, who())] || null;
  /** What is typed now, added up. */
  D.draftTotal = function () {
    const ds = D.sheet();
    if (!ds) return 0;
    const d = D.draftNow();
    return ds.counts.reduce((n, c) => n + (ST.dayCount(d.counts[c[0]]) || 0), 0);
  };
  /** Is everything typed a whole number within the cap? */
  D.draftValid = function () {
    const ds = D.sheet();
    if (!ds) return false;
    const d = D.draftNow();
    return ds.counts.every(c => ST.dayCount(d.counts[c[0]]) != null);
  };
  /** Has anybody actually written anything? AN UNTOUCHED FORM IS NOT NOUGHTS
      (review, 2026-09-21); a note alone ("machine down all day") is a real
      entry. */
  D.draftTyped = function () {
    const ds = D.sheet();
    if (!ds) return false;
    const d = D.draftNow();
    return ds.counts.some(c => String(d.counts[c[0]] == null ? "" : d.counts[c[0]]).trim() !== "") ||
           String(d.note || "").trim() !== "";
  };
  /** still counting this person's taps: Save waits for the count */
  const counting = () => !!(D.sheet() && D.sheet().counted && D.logState === "reading");
  D.draftOk = () => D.draftValid() && D.draftTyped() && !counting();
  /** "This week: N of T": this person's saved sheets in the week of the day
      this sheet is about, plus what is typed now, against the target. */
  D.weekWords = function (extra) {
    const wk = ST.isoWeek(D.sheetDay()).key;
    const target = D.draftNow().target == null ? D.target : D.draftNow().target;
    const done = ST.dayWeekTotal(D.rows, who(), wk) + (extra || 0);
    const unit = (D.sheet() || {}).unit || "sheets";
    return { done: done, target: target,
             words: target == null ? done + " " + unit + " · no target set" : done + " of " + target + " " + unit,
             pct: target > 0 ? Math.max(0, Math.min(100, Math.round(done * 100 / target))) : 0 };
  };

  /** Save the sheet: one row queued, sent by the page's own queue. */
  D.save = function () {
    const ds = D.sheet();
    if (!ds || !who()) return false;
    const d = D.draftNow();
    if (D.saved(d.day) || D.owed(d.day)) return false;              // one per person per day
    if (!D.draftTyped()) { D.bad = "Fill in a number or write a line first."; cfg.render(); return false; }
    /* a stage that counts: the count must be fresh, or known to be unavailable */
    if (ds.counted && D.logState !== "ok" && D.logState !== "failed") {
      if (D.logState !== "reading") D.open();
      D.bad = "Still counting your taps — a moment."; cfg.render(); return false;
    }
    const counted = ds.counted ? D.counted(d.day) : null;
    const fields = ST.dayFields(def, stage(), {
      day: d.day, who: who(), counts: d.counts, note: d.note,
      weekTarget: d.target == null ? D.target : d.target, counted: counted, at: new Date().toISOString() });
    if (!fields) {
      D.bad = "Whole numbers only, please — no minus signs, no decimals, nothing over " + ST.DAY_COUNT_MAX + ".";
      cfg.render(); return false;
    }
    /* ONLY A LIST THAT IS NOT THERE refuses the save (review, 2026-09-21): a
       list that could not be READ is the workshop wifi, and the sheet queues. */
    if (D.missing) { D.bad = D.why || ST.DAY_MISSING_FLOOR; cfg.render(); return false; }
    /* the tablet's count against what was typed: said once, in the same
       question, and never a refusal - the person knows their own day. No taps
       at all is not a disagreement (owner: the count is a help, not a check) */
    const typed = ds.counts.reduce((n, c) => n + (fields[c[0]] || 0), 0);
    const differs = counted != null && counted > 0 && typed !== counted;
    const ask = (differs ? "You counted " + counted + " taps today and typed " + typed + ". Save anyway? " : "") +
      "Save the sheet for " + d.day + "? It cannot be changed from the tablet afterwards.";
    if (typeof confirm === "function" && !confirm(ask)) return false;
    D.bad = "";
    D.q[fields.Title] = { fields: fields, err: 0, refused: 0 };
    D.persist();
    D.clearDraft();
    cfg.touch();
    cfg.render();
    cfg.flush();
    return true;
  };
  /** Send the owed sheet. A refusal is very often the list's own unique rule
      catching a second tablet, or a replay of a row that landed and whose
      answer was lost - both mean ALREADY SAVED, so the list is read back
      before anything is called a failure. A row SharePoint keeps refusing
      (a 4xx) stops saying "waiting to send" after STU_DAY_REFUSE_MAX and says
      to tell the office; it STAYS QUEUED either way. */
  /** Is a row with this Title on the list already, whatever stage it is? */
  async function landedAlready(f, opts) {
    try {
      const items = await CW.listItems(ST.DAY_LIST, opts);
      const want = String(f.Title || "").trim().toUpperCase();
      return !!items && items.some(it => String(((it && it.fields) || {}).Title || "").trim().toUpperCase() === want);
    } catch (e) { return false; }
  }
  D.flush = async function () {
    /* not gated on this page's stage having a sheet: every queued row was
       rebuilt through its own stage's definition (cleanQ) */
    if (!siteId() || !D.owing()) return;
    const keys = Object.keys(D.q);
    for (let i = 0; i < keys.length; i++) {
      const e = D.q[keys[i]];
      if (!e) continue;
      /* the queued row's OWN stage, not the page's: a tablet switched from
         cutting to hotmelting may still owe a cutting sheet (review, 2026-09-28) */
      const rowOpts = { siteId: siteId(), fields: ST.dayFieldsFor(def, e.fields.Stage) };
      try {
        await CW.listAdd(ST.DAY_LIST, e.fields, rowOpts);
        delete D.q[keys[i]];
        await D.read();
      } catch (err) {
        if (await landedAlready(e.fields, rowOpts)) {
          delete D.q[keys[i]]; D.persist(); await D.read(); continue;
        }
        e.err = 1;
        /* only a refusal counts: a dropped connection is not the list saying no */
        const refused = !!(CW.isMissing && CW.isMissing(err)) || !!(CW.isRefused && CW.isRefused(err)) ||
                        /->\s*4\d\d\b/.test((err && err.message) || "");
        if (refused) e.refused = (Number(e.refused) || 0) + 1;
        console.warn(tag + " the day sheet is not saved yet" +
                     (e.refused >= STU_DAY_REFUSE_MAX ? " and has been refused " + e.refused + " times" : "") +
                     ":", (err && err.message) || err);
      }
      D.persist();
    }
  };
  /** What the card says about an owed sheet. */
  D.owedWords = function (e) {
    if (e && Number(e.refused) >= STU_DAY_REFUSE_MAX) return "could not be saved — tell the office";
    if (e && e.err) return "waiting to send — it will go when the tablet is back on the wifi";
    return "waiting to send…";
  };

  /* ---- drawn ----
     One column, full width, portrait, nothing to scroll sideways and nothing
     tapped under 44 px - the tablet rule of 2026-09-17. It takes the board's
     place rather than opening over it, and Back is the only way out.       */
  function lineHtml(r, cs) {
    return '<div class="dayline"><span class="daylday">' + esc(r.day) + '</span>' +
      '<span class="dayldow">' + esc(r.weekday.slice(0, 3)) + '</span>' +
      cs.map(c => '<span class="daylnum tab">' + esc(ST.dayCountShort(c)) + ' ' +
        (r.counts[c[0]] || 0) + '</span>').join("") +
      '<span class="dayltot tab">' + r.total + '</span></div>';
  }
  function historyHtml(cs) {
    const mine = D.rows.filter(r => r.who === who()).slice(0, 7);
    if (!mine.length) return "";
    return '<div class="dayhist"><div class="kick">Your last ' + mine.length + ' day' +
      (mine.length === 1 ? "" : "s") + '</div>' + mine.map(r => lineHtml(r, cs)).join("") + '</div>';
  }
  /** The tablet's own count, for a stage that keeps one. */
  function countedHtml(day) {
    const ds = D.sheet();
    if (!ds || !ds.counted) return "";
    const n = D.counted(day);
    const unit = ds.unit || "units";
    return '<div class="daycount">' + (n == null
      ? (D.logState === "failed"
          ? "Your taps could not be counted just now — the sheet saves without the tablet’s count."
          : "Counting your taps today…")
      : "You " + esc(ds.verb || "recorded") + " " +
        '<strong class="tab">' + n + '</strong> ' + esc(unit) + " " +
        (day === today() ? "today" : "on " + esc(ST.isoWeek(day).weekday))) + '</div>';
  }
  D.html = function () {
    const ds = D.sheet();
    if (!ds || !who()) return '<div class="msg">Nothing to fill in here.</div>';
    const cs = ds.counts, unit = ds.unit || "sheets";
    const day = D.sheetDay(), w = ST.isoWeek(day);
    const carried = day !== today();
    const head = '<div class="picker daysheet">' +
      '<div class="pickh">End of day</div>' +
      '<div class="picksub">' + esc(who() + " · " + w.weekday + " " + day + " · " + cfg.stageWords()) + '</div>' +
      (carried ? '<div class="daycarry">This is the sheet you started on ' + esc(w.weekday) +
         ', and it saves under that day.</div>' : "");
    const back = '<button class="pcancel" data-dayback="1">Back to the board</button></div>';
    const saved = D.saved(day), owed = D.owed(day);
    if (saved || owed) {
      const row = saved || ST.dayRows([{ id: "", fields: owed.fields }], cs)[0];
      /* the LOCAL clock (ST.stClock), not a slice of the ISO stamp */
      const when = saved ? "saved " + ST.stClock(saved.savedAt) + " — " + ST.DAY_SAVED_WORDS : D.owedWords(owed);
      return head +
        '<div class="dayread">' +
          cs.map(c => '<div class="dayrow"><span class="dayl">' + esc(c[1]) + '</span>' +
            '<span class="dayv tab">' + (row.counts[c[0]] || 0) + '</span></div>').join("") +
          (cs.length > 1 ? '<div class="dayrow"><span class="dayl">Total</span>' +
            '<span class="dayv tab">' + row.total + '</span></div>' : "") +
          (row.note ? '<div class="daynote">' + esc(row.note) + '</div>' : "") +
          '<div class="daysaid' + (owed ? " owed" : "") +
            (owed && Number(owed.refused) >= STU_DAY_REFUSE_MAX ? " bad" : "") + '">' + esc(when) + '</div>' +
        '</div>' + historyHtml(cs) + back;
    }
    const d = D.draftNow();
    const wk = D.weekWords(D.draftTotal());
    return head +
      (D.ok === false ? '<div class="cphint">' + esc(D.why) + '</div>' : "") +
      countedHtml(day) +
      '<div class="dayform">' +
        cs.map(c => '<label class="dayrow"><span class="dayl">' + esc(c[1]) + '</span>' +
          '<input class="daybox tab" type="text" inputmode="numeric" pattern="[0-9]*" ' +
            'data-daycount="' + esc(c[0]) + '" value="' + esc(String(d.counts[c[0]] == null ? "" : d.counts[c[0]])) +
            '" aria-label="' + esc(c[1]) + '"></label>').join("") +
        '<label class="dayrow daynoterow"><span class="dayl">Anything that got in the way</span>' +
          '<textarea class="daytext" data-daynote="1" rows="3" maxlength="' + ST.DAY_NOTE_MAX +
          '" placeholder="Machine down, waiting on glass, helped on another bench…">\n' +
          esc(d.note) + '</textarea></label>' +
        '<div class="daytot">' + esc(carried ? w.weekday : "Today") +
          ': <strong class="tab" id="daytoday">' + D.draftTotal() + '</strong> ' + esc(unit) + '</div>' +
        /* the week against its target, only where the stage has one */
        (hasTarget()
          ? '<div class="dayweek">This week: <strong class="tab" id="dayweeknum">' + esc(wk.words) + '</strong></div>' +
            '<div class="daybar"><span id="daybarfill" style="width:' + wk.pct + '%"></span></div>'
          : "") +
        (D.bad ? '<div class="daybad">' + esc(D.bad) + '</div>' : "") +
        '<button class="daysave" id="daysave" data-daysave="1"' +
          (D.draftOk() ? "" : ' disabled aria-disabled="true"') + '>Save the sheet for ' +
          esc(carried ? w.weekday : "today") + '</button>' +
      '</div>' + historyHtml(cs) + back;
  };
  /** The sheet's own clicks and keystrokes. The boxes are NOT redrawn as they
      are typed into - only the totals and the bar - because a card rebuilt on
      every keystroke is a caret lost on every keystroke. */
  D.wire = function (host) {
    if (!host || !host.querySelectorAll) return;
    const live = () => {
      const t = D.draftTotal(), wk = D.weekWords(t);
      const a = q$("#daytoday"); if (a) a.textContent = String(t);
      const b = q$("#dayweeknum"); if (b) b.textContent = wk.words;
      const c = q$("#daybarfill"); if (c) c.style.width = wk.pct + "%";
      const s = q$("#daysave"); if (s) s.disabled = !D.draftOk();
    };
    host.querySelectorAll("[data-daycount]").forEach(el => el.oninput = () => {
      cfg.touch();
      D.draftNow().counts[el.dataset.daycount] = el.value;
      D.saveDraft(); live();
    });
    host.querySelectorAll("[data-daynote]").forEach(el => el.oninput = () => {
      cfg.touch();
      D.draftNow().note = String(el.value || "").slice(0, ST.DAY_NOTE_MAX);
      D.saveDraft();
    });
    host.querySelectorAll("[data-daysave]").forEach(el => el.onclick = () => {
      if (el.disabled) return;
      D.save();
    });
    host.querySelectorAll("[data-dayback]").forEach(el => el.onclick = () => {
      D.shown = false; D.bad = ""; cfg.touch(); cfg.render();
    });
  };
  /** The header button: open the sheet and ask the lists once. For a stage
      that counts its own taps the log is asked every time, so the number is
      today's and not this morning's. */
  D.open = function () {
    if (!D.sheet() || !who()) return;
    const counts = !!D.sheet().counted;
    D.shown = true; D.bad = "";
    if (counts) D.logState = "reading";                 // the count on screen waits for this read
    cfg.touch(); cfg.render();
    const jobs = [];
    if (D.ok !== true) jobs.push(D.read());
    if (counts) jobs.push(D.readLog().then(ok => { D.logState = ok ? "ok" : "failed"; },
                                           () => { D.logState = "failed"; }));
    const redraw = () => { if (D.shown) cfg.render(); };
    if (jobs.length) Promise.all(jobs).then(redraw, redraw);
  };
  /** A change of person: the sheet closes and the half-typed draft is let go
      from memory (it is still in storage under whose it is). */
  D.reset = function () { D.shown = false; D.bad = ""; D.draft = null; };
  return D;
}

/* ---- another station's list, READ ONLY --------------------------------------
   2026-10-02 (docs/specs/2026-10-02-stations-see-each-other.md, section 5):
   one owner per fact, and everybody who needs the fact reads the owner's list.
   This is that read, written once - it began as the glass page's look at
   `Fabrication station` (2026-10-01) and is the same code for every page.

     const R = STU.stuListReader({ site: "own" | "floor", list: "Glass station",
                                   fields: [...], tag: "[glazing]" });
     if (await R.read()) render();        // in the page's tickOnce, and once,
                                          // NOT awaited, before its first paint
     R.ready     false until the first read has answered: say "checking"
     R.items     the list's rows [{ id, fields }], or null - NOT AVAILABLE
                 (no such site or list, 403, 404). Never read null as "nothing
                 done" or "ready".
     R.changed   did the last read change anything (it is what read() answers)
     R.missing   items are null because the list is positively not there (as
                 against refused, or not answered yet)
     R.asOf      when the last read that was ANSWERED went out (0: none yet) -
                 what is held is at least that fresh

   It holds the list by CW.listDelta on a token of its own, one read in the air
   at a time, and whether a reply is merged or replaces the list is decided by
   the token the request WENT OUT with. A 410 costs a fresh enumeration on the
   next turn. Only 403 and 404 take the rows away (a 404 also forgets that
   channel's cached site); anything else - offline, throttled, a bad gateway -
   keeps the last good read. After a failure it does not ask again for
   `retryMs` (a minute unless the page says otherwise). A list that refuses
   delta outright is read the plain way for five minutes, as the boards do.

   IT HAS NO WRITE PATH, and it keeps nothing outside the object it answers:
   its site comes from CW.stationSite(cfg.site) and is never the page's own
   SITEID, its token is never the page's TOKEN, and it cannot take a board
   away. Two readers on one page share nothing.                             */
const STU_READ_RETRY_MS = 60000;
const STU_DELTA_OFF_MS = 300000;
function stuListReader(cfg) {
  const tag = cfg.tag || "[station]";
  const retry = cfg.retryMs == null ? STU_READ_RETRY_MS : cfg.retryMs;
  const R = { items: null, ready: false, changed: false, missing: false, asOf: 0,
              busy: false, site: null, token: null, offAt: 0, deltaOffAt: 0 };
  const is = (k, e) => !!(CW[k] && CW[k](e));
  /* the list will not serve a delta at all - not a stale token (410), not a
     refusal or a missing list, and not a throttle (graph.js no longer hands a
     429 back as a restart; the test here stays as a second line of defence) */
  const noDelta = e => is("isDeltaRestart", e) && !is("isDeltaResync", e) && !is("isMissing", e) &&
    !is("isRefused", e) && !/->\s*429\b/.test((e && e.message) || "");
  R.read = async function () {
    if (R.busy) return false;
    if (R.offAt && Date.now() - R.offAt < retry) return false;
    R.busy = true;
    const began = Date.now();
    const had = R.items != null;
    let moved = false;
    try {
      const site = await CW.stationSite(cfg.site);
      if (site !== R.site) { R.site = site; R.token = null; R.deltaOffAt = 0; }   // a token only means anything in its own site
      const opts = { siteId: site, fields: cfg.fields };
      /* merge or replace is decided by the token the request WENT OUT with */
      const tok = R.token;
      let plain = !!(R.deltaOffAt && Date.now() - R.deltaOffAt < STU_DELTA_OFF_MS);
      let d = null;
      if (site && !plain) {
        try { d = await CW.listDelta(cfg.list, Object.assign({ token: tok || undefined }, opts)); }
        catch (e) {
          if (!noDelta(e)) throw e;
          R.deltaOffAt = Date.now(); plain = true;
          console.warn(tag + " “" + cfg.list + "” refused a delta; reading it the plain way for five minutes");
        }
      }
      if (site && plain) {
        const all = await CW.listItems(cfg.list, opts);
        d = all && { items: all, next: null };
      }
      if (d == null) {                     // no such site for this account, or no such list in it
        R.items = null; R.token = null; R.missing = !!site; R.offAt = Date.now();
      } else {
        if (tok && !plain) {
          moved = d.items.length > 0;
          R.items = ST.mergeDelta(R.items || [], d.items);
        } else {
          const was = JSON.stringify(R.items);
          R.items = d.items.filter(x => !x.removed).map(x => ({ id: x.id, fields: x.fields }));
          moved = JSON.stringify(R.items) !== was;
        }
        if (plain) R.token = null; else if (d.next) R.token = d.next;
        R.missing = false; R.offAt = 0; R.asOf = began;
      }
    } catch (e) {
      console.warn(tag + " “" + cfg.list + "” could not be read:", (e && e.message) || e);
      R.token = null;
      /* a token gone stale is not a failure: enumerate again on the next turn.
         Anything else waits. ONLY "not there" (404) and "not for you" (403)
         take the rows away; every other failure keeps the last read. */
      if (!is("isDeltaResync", e)) {
        R.offAt = Date.now();
        const gone = is("isMissing", e);
        if (gone || is("isRefused", e)) { R.items = null; R.missing = gone; }
        /* this reader's own channel, never the page's board's */
        if (gone && CW.forgetStationSite) CW.forgetStationSite(false, cfg.site);
      }
    } finally {
      R.busy = false;
    }
    R.changed = !R.ready || had !== (R.items != null) || moved;
    R.ready = true;
    return R.changed;
  };
  return R;
}

/* ---- the clock every station page polls on ----------------------------------
   2026-10-02. One timer, beating once a second, and ST.tickDue says whether a
   turn is due: every ST.REFRESH_MS, or every ST.TICK_FAST_MS for a while after
   burst() - which a page calls on ITS OWN TAP AND NOTHING ELSE. A change seen
   in a poll, or in a read of another station's list, is redrawn and does not
   touch the rate: when a change seen started a burst, every tablet on an
   account burst whenever any one of them was tapped, and ten tablets measured
   1020 requests a minute. A TURN NEVER STARTS WHILE THE LAST ONE IS STILL
   RUNNING: a slow reply costs skipped turns, not a pile of them. No pause and
   no visibility check - a tablet polls all the time.

     const TICK = STU.stuTicker(tickOnce);     // tickOnce: the page's own turn
     TICK.start();                             // in start(), in setInterval's place
     TICK.burst();                             // on this tablet's own tap, only   */
function stuTicker(fn) {
  const T = { busy: false, lastAt: 0, until: 0, timer: null };
  T.burst = () => { T.until = Date.now() + ST.TICK_BURST_MS; };
  /** One turn now, unless one is in the air. Answers whether it ran. */
  T.turn = async function () {
    if (T.busy && Date.now() - T.lastAt < ST.TICK_STUCK_MS) return false;
    T.busy = true; T.lastAt = Date.now();
    const mine = T.lastAt;
    try { await fn(); }
    catch (e) { console.warn("[station] a turn of the clock failed:", (e && e.message) || e); }
    finally { if (T.lastAt === mine) T.busy = false; }
    return true;
  };
  T.start = function () {
    if (T.timer) clearInterval(T.timer);
    T.lastAt = Date.now();
    T.timer = setInterval(() => { if (ST.tickDue(Date.now(), T.lastAt, T.until)) T.turn(); }, ST.TICK_BEAT_MS);
  };
  return T;
}

const STU = {
  STU_THEME_KEY, STU_DAY_REFUSE_MAX, STU_READ_RETRY_MS,
  stuEsc, stuAgo, stuThemeNow, stuApplyTheme, stuGate, stuPickerHtml, stuWirePicker, stuDaySheet,
  stuListReader, stuTicker
};
if (typeof window !== "undefined") window.STU = STU;
else if (typeof globalThis !== "undefined") globalThis.STU = STU;
if (typeof module !== "undefined" && module.exports) module.exports = STU;
