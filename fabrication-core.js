/* Fabrication station - the pure logic of the fourth floor station
   (docs/specs/2026-09-25-fabrication-station.md, Part A).

   Fabrication is the floor step after welding, for windows and doors. One list
   row per job AND product group, three counters per row: frames, sashes and
   transoms fabricated, each out of the group's own F / S / T on `Production`.
   Everybody can SEE every line; a person can only MOVE the lines of a product
   group they are eligible for (their `Stages` in `Station people` name the
   groups).

   Nothing in this file touches a DOM, Graph or a workbook. It is the slice the
   feeder sends, the rows the two boards draw, the clamp on a tap, the colour
   rule, the fields one write may carry, and the pure half of the office's
   colour painter (fbCellWant): which colour a product cell should be. The
   painting itself lives in app.js, because the tablet never writes the sheet.

   Borrowed, not copied: ST.floorOnly, ST.stripContact, ST.sectionInProduction,
   ST.stationPeople, ST.tabSearch, ST.feedPlan and ST.sliceHash take this
   station's definition (FAB, below). Every top-level name here starts FB_ or
   fb, because classic scripts share one global scope (HISTORY B29) and app.js
   already owns FAB_* and fab* for its floating button.                       */

/* ---- names ---------------------------------------------------------------- */
const FB_LIST = "Fabrication station";
const FB_NAME = "Fabrication";      // the `Station` word in the three shared lists
const FB_SITE = "floor";            // `Floor stations`, and nothing else
const FB_STAGE = "fab";             // the station report's one stage

/* ---- what is counted: all three parts (welding has no T; this station has) */
const FB_PARTS = ["frames", "sashes", "transoms"];
const FB_PART_LABEL = { frames: "Frames", sashes: "Sashes", transoms: "Transoms" };
const FB_PART_SUB = { frames: "f", sashes: "s", transoms: "t" };
const FB_TOTAL_FIELD = { frames: "Frames", sashes: "Sashes", transoms: "Transoms" };
const FB_DONE_FIELD = { frames: "FramesDone", sashes: "SashesDone", transoms: "TransomsDone" };
const FB_BY_FIELD = { frames: "FramesBy", sashes: "SashesBy", transoms: "TransomsBy" };
const FB_AT_FIELD = { frames: "FramesAt", sashes: "SashesAt", transoms: "TransomsAt" };

/* ---- door glazing: a fourth part, on PVC DOOR and PVC SMART only (owner,
   2026-10-01). NOT one of FB_PARTS: it is its own count. It is in no
   fabrication total, colour, tab, assignment or sheet paint; it has a role
   (`Stages` must name it: "PVC DOOR:glazing") and no assignment. Its total is
   the number of the job's DOORS DONE cells whose code belongs to the group. */
const FB_GLAZE = "glazing";
const FB_TAP_PARTS = FB_PARTS.concat([FB_GLAZE]);        // every counter a tap may move
FB_PART_LABEL[FB_GLAZE] = "Door glazing";
FB_TOTAL_FIELD[FB_GLAZE] = "GlazeTotal";
FB_DONE_FIELD[FB_GLAZE] = "GlazeDone";
FB_BY_FIELD[FB_GLAZE] = "GlazeBy";
FB_AT_FIELD[FB_GLAZE] = "GlazeAt";
const FB_GLAZE_GROUPS = ["PVC DOOR", "PVC SMART"];
/* door codes that are never door glazing here (owner: "… but not CD") */
const FB_GLAZE_NEVER = ["CD", "SFCD", "BF", "ACSD", "ACSS"];

/* ---- the columns ----------------------------------------------------------
   The feeder writes Title, the job facts and FedAt/FedBy. The floor (and the
   office's own board) writes FB_FLOOR_FIELDS and nothing else. `Urgent` is the
   office's alone and is Part B's: it is in FB_FIELDS so a read carries it, and
   in neither write list. No seed: there is no office record of fabrication. */
const FB_FEEDER_FIELDS = ["Job", "Group", "GroupSeq", "Customer", "Comment", "Seq", "Doors",
                          "Frames", "Sashes", "Transoms", "Section", "Active", "OnSheet",
                          "GlazeTotal"];
/* the feeder column added 2026-10-01. A list that has not got it yet must not
   have its whole feed refused: the office strips it from its writes and says
   so (app.js, fabrFeedWrite). The `Glass` column added the same day is no
   longer read or written by anybody (2026-10-02): it stays on the list. */
const FB_NEW_FEEDER_FIELDS = ["GlazeTotal"];
const FB_FLOOR_FIELDS = ["FramesDone", "SashesDone", "TransomsDone",
                         "FramesBy", "FramesAt", "SashesBy", "SashesAt", "TransomsBy", "TransomsAt",
                         "GlazeDone", "GlazeBy", "GlazeAt",
                         "DoneBy", "DoneAt"];
const FB_COUNTER_FIELDS = ["FramesDone", "SashesDone", "TransomsDone", "GlazeDone"];
const FB_OFFICE_FIELDS = ["Urgent"];
const FB_SEED_FIELDS = [];
/* `Modified` is SharePoint's own column: the server's clock for the row's last
   write. Read only - it is in no write list - and it is what the gold rule
   compares instead of a tablet's clock (review C-1). */
const FB_FIELDS = ["Title"].concat(FB_FEEDER_FIELDS, ["FedAt", "FedBy"], FB_FLOOR_FIELDS, FB_OFFICE_FIELDS, ["Modified"]);
const FB_FEEDER_WRITES = ["Title"].concat(FB_FEEDER_FIELDS, ["FedAt", "FedBy"]);
/* a row whose job or group has left the sheet: both flags off, never deleted
   (Active alone means "In production", so it cannot also mean "gone" - B31) */
const FB_GONE = { Active: "No", OnSheet: "No" };

/* ---- which product groups (owner, 2026-09-25) -----------------------------
   An ALLOW-list, matched through fbKey, using the sheet's own headers. The
   owner's "Arch & AnglesTH W" is two headers on the sheet, ARCH ANGLES and
   TH W, and "Composite PVC door" is the header COMPOSITE (sashes only: a CD's
   sash; its frame is on PVC DOOR). Owner's answer 2026-09-25: add all three,
   as the sheet names them. */
const FB_WINDOW_GROUPS = ["CASEMENT WINDOWS", "POLARIS 85MM CASEMENT", "4000 CASEMENT", "7000 CASEMENT",
  "POLARIS 85 TILT & TURN", "4000 TILT & TURN", "7000 TILT & TURN", "PVC FRENCH WDS",
  "ARCH ANGLES", "TH W", "ALU CLAD WINDOWS", "ALUCLAD TILT & TURN"];
const FB_DOOR_GROUPS = ["SIDELIGHTS", "SUPER DOOR", "BIFOLD", "PVC SMART", "COMPOSITE", "PVC DOOR"];

const FB_CUSTOMER_MAX = 70;
const FB_COMMENT_MAX = 140;
const FB_DOORS_MAX = 60;

/* ---- the two colours this station paints (rule 6, owner 2026-09-25) -------
   Neither occurs on the Production sheet today, and neither is yellow or gold,
   so nothing that reads the sheet mistakes fabrication for the office. */
const FB_PROCESS_HEX = "#D9D2E9";   // light lavender: started on that line
const FB_DONE_HEX = "#B4A7D6";      // purple: that line fully fabricated
const FB_WHITE_HEX = "#FFFFFF";

/* ---- small helpers -------------------------------------------------------- */
const fbTxt = v => String(v == null ? "" : v);
/** Upper case, every run of non-alphanumerics one space: the parser drops the
    "&" out of "TILT & TURN", so nothing may depend on it. */
const fbKey = v => fbTxt(v).trim().toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
function fbNum(v, dflt) {
  if (v == null || fbTxt(v).trim() === "") return dflt;
  const n = Number(v);
  return isFinite(n) ? n : dflt;
}
const fbInt = (v, dflt) => Math.round(fbNum(v, dflt == null ? 0 : dflt));
const fbClamp = (n, total) => Math.max(0, Math.min(total, Math.round(fbNum(n, 0))));
function fbST() {
  return (typeof ST !== "undefined" && ST) || (typeof window !== "undefined" && window.ST) || null;
}
/** Rule 3's strip. Fails CLOSED: with no strip to hand, no free text at all. */
function fbStrip(text, max) {
  const S = fbST();
  if (!S || typeof S.stripContact !== "function") return "";
  return S.stripContact(text, fbInt(max, FB_COMMENT_MAX));
}
const fbSectionLive = s => { const S = fbST(); return S ? !!S.sectionInProduction(s) : false; };

const FB_GROUP_KEYS = FB_WINDOW_GROUPS.concat(FB_DOOR_GROUPS).map(fbKey);
const FB_DOOR_KEYS = FB_DOOR_GROUPS.map(fbKey);
const fbAllowed = g => { const k = fbKey(g); return !!k && FB_GROUP_KEYS.indexOf(k) >= 0; };
const fbIsDoorGroup = g => FB_DOOR_KEYS.indexOf(fbKey(g)) >= 0;
const fbTitle = (job, group) => fbKey(job) + "|" + fbKey(group);

/* ---- what the feeder sends ------------------------------------------------ */
/** The job's DOORS DONE codes as labels: "2 CD, 1 DD". Labels only - nothing is
    computed from them (the sheet's F/S/T are already the right counts). */
function fbDoorLabels(j) {
  const n = {}, order = [];
  ((j && j.doors) || []).forEach(d => {
    const c = fbTxt(d && d.code).trim().toUpperCase();
    if (!c) return;
    if (!n[c]) { n[c] = 0; order.push(c); }
    n[c]++;
  });
  return fbStrip(order.map(c => n[c] + " " + c).join(", "), FB_DOORS_MAX);
}
function fbCommentOf(j) {
  const hit = ((j && j.notes) || []).find(x => x && x.k === "comment");
  return fbStrip(hit ? hit.t : "");
}
/** Which group a DOORS DONE code's glazing belongs to: SS is PVC SMART; CD,
    SFCD, BF and the AC codes are never glazed here; every other non-empty code
    (PVC, DD, SD, ...) is PVC DOOR. "" = not counted. */
function fbDoorGlazeGroup(code) {
  const c = fbKey(code);
  if (!c || FB_GLAZE_NEVER.indexOf(c) >= 0) return "";
  return c === "SS" ? "PVC SMART" : "PVC DOOR";
}
/** { "PVC DOOR": n, "PVC SMART": m } for one job, off j.doors (Production alone). */
function fbGlazeTotals(j) {
  const out = {};
  ((j && j.doors) || []).forEach(d => {
    const g = fbDoorGlazeGroup(d && d.code);
    if (g) out[g] = (out[g] || 0) + 1;
  });
  return out;
}

/* ---- the job's glass, as the fabricators are told it (2026-10-01) -----------
   One word per job, worked out from the job's own `Glass station` row by
   whoever is looking: since 2026-10-02 every screen reads that list itself
   (one owner per fact) and nobody copies the word anywhere (fbGlassOf).
     ""                         the job has no glass
     "wait" / "off"             the glass list has not answered yet / cannot be
                                read (fbGlassOf only): never ready
     "none"                     it has glass and nothing is recorded
     "part:c/t:h/t[:tuff n/m]"  started, not complete
     "done"                     cut and hotmelt complete, and tuff when the job
                                has tuff - or the office says the glass is done */
function fbGlassStatus(g, hasGlass) {
  if (!hasGlass) return "";
  if (!g) return "none";
  const t = Math.max(0, fbInt(g.total, 0)), tt = Math.max(0, fbInt(g.tuffTotal, 0));
  if (!(t > 0) && !(tt > 0)) return g.officeDone ? "done" : "none";
  /* THE TUFF RULE HOLDS FOR THE OFFICE'S "DONE" TOO (review A-1): OfficeDone
     speaks for DG and TG only and has never known about tuff, so a job that
     still owes tuff is not done whoever says the glass is. The office's done
     stands in for cutting and hotmelting, and for nothing else. */
  const c = g.officeDone ? t : fbClamp(g.cut, t), h = g.officeDone ? t : fbClamp(g.hotmelt, t);
  const tf = fbClamp(g.tuff, tt);
  if (c >= t && h >= t && tf >= tt) return "done";
  if (!(c > 0 || h > 0 || tf > 0)) return "none";
  return "part:" + c + "/" + t + ":" + h + "/" + t + (tt > 0 ? ":tuff " + tf + "/" + tt : "");
}
/** job -> its glass word, from the `Glass station` rows alone (a read-only
    reader's `items` and `ready`). No row for the job = no glass on the floor's
    list = "". A list not answered yet is "wait" and one that cannot be read
    (items null: 403, 404) is "off" - neither is ever "none" or "done". */
function fbGlassOf(items, ready) {
  const S = fbST();
  if (!ready) return () => "wait";
  if (!items || !S) return () => "off";
  const recs = S.jobRecords(items);
  return job => {
    const g = recs[S.jobKey(job)];
    return g ? fbGlassStatus(g, g.total > 0 || g.tuffTotal > 0) : "";
  };
}
/** The chip a glass word draws: { kind: "done" | "part" | "none" | "wait" | "off" | "", words }. */
function fbGlassChip(text) {
  const s = fbTxt(text).trim();
  if (!s) return { kind: "", words: "" };
  if (s === "wait") return { kind: "wait", words: "Glass: checking…" };
  if (s === "off") return { kind: "off", words: "Glass: not available" };
  if (s === "done") return { kind: "done", words: "Glass ✓ done" };
  if (s === "none") return { kind: "none", words: "Glass: not started" };
  const m = /^part:(\d+\/\d+):(\d+\/\d+)(?::tuff (\d+\/\d+))?$/.exec(s);
  if (!m) return { kind: "", words: "" };
  return { kind: "part", words: "Glass: cut " + m[1] + " · hotmelt " + m[2] + (m[3] ? " · tuff " + m[3] : "") };
}
/** Glass is ready and this job is not finished: the ones "Glass ready first"
    lifts. With nothing fabricated at all it also gets the "start this job" line. */
const fbGlassReady = c => !!c && fbTxt(c.glass) === "done" && !!c.active && !c.sheetDone && !c.finished;
const fbGlassStart = c => fbGlassReady(c) && !(c.done > 0);
/** Urgent first, then (when asked) glass-ready, then the rest - each keeping
    the board's own order. */
function fbGlassFirst(cards, on) {
  const u = [], g = [], rest = [];
  (cards || []).forEach(c => (fbCardUrgent(c) ? u : on && fbGlassReady(c) ? g : rest).push(c));
  return u.concat(g, rest);
}

/** One row per job and allowed product group with F, S or T > 0, off
    `Production` ALONE (j.prodsMain; never j.prods - HISTORY B20). Every section
    on the sheet: Active says In production, OnSheet says still on the sheet. */
function fbSlice(jobs, blockNames) {
  const names = blockNames || (jobs && jobs.blockNames) || [];
  const out = [], emitted = {};
  (jobs || []).forEach(j => {
    if (!j || !j.id || j.cat === "past") return;
    const job = fbKey(j.id);
    if (!job) return;
    const section = fbTxt(names[j.blk] || "").trim();
    const base = { job: job, customer: fbStrip(j.cust, FB_CUSTOMER_MAX), comment: fbCommentOf(j),
                   seq: fbNum(j.seq, 99999), section: section,
                   active: fbSectionLive(section), onSheet: true };
    const doors = fbDoorLabels(j);
    const glaze = fbGlazeTotals(j);
    (j.prodsMain || []).forEach((p, i) => {
      if (!p || !p.n || !fbAllowed(p.n)) return;
      const group = fbKey(p.n);
      const f = Math.max(0, fbInt(p.f, 0)), s = Math.max(0, fbInt(p.s, 0)), t = Math.max(0, fbInt(p.t, 0));
      if (!(f > 0 || s > 0 || t > 0)) return;
      const title = fbTitle(job, group);
      if (emitted[title]) {
        if (typeof console !== "undefined" && console.warn)
          console.warn("[fabrication] " + title + " appears twice on the sheet: the second is not fed.");
        return;
      }
      emitted[title] = 1;
      out.push(Object.assign({ title: title, group: group, groupSeq: i,
                               doors: fbIsDoorGroup(group) ? doors : "",
                               frames: f, sashes: s, transoms: t,
                               glazeTotal: glaze[group] || 0 }, base));
    });
  });
  out.sort(fbRowOrder);
  return out;
}
function fbRowOrder(a, b) {
  return (fbNum(a.seq, 99999) - fbNum(b.seq, 99999)) ||
         (fbNum(a.groupSeq, 999) - fbNum(b.groupSeq, 999)) ||
         (a.title < b.title ? -1 : a.title > b.title ? 1 : 0);
}
function fbFeederFields(r) {
  return { Job: r.job, Group: r.group, GroupSeq: fbInt(r.groupSeq, 0),
           Customer: fbTxt(r.customer), Comment: fbTxt(r.comment), Seq: fbNum(r.seq, 99999),
           Doors: fbTxt(r.doors),
           Frames: Math.max(0, fbInt(r.frames, 0)), Sashes: Math.max(0, fbInt(r.sashes, 0)),
           Transoms: Math.max(0, fbInt(r.transoms, 0)),
           Section: fbTxt(r.section), Active: r.active ? "Yes" : "No", OnSheet: r.onSheet ? "Yes" : "No",
           GlazeTotal: Math.max(0, fbInt(r.glazeTotal, 0)) };
}
const fbSeedFields = () => ({});
const fbHashRow = r => [r.title, r.job, r.group, r.groupSeq, r.customer, r.comment, r.seq, r.doors,
                        r.frames, r.sashes, r.transoms, r.section, !!r.active, !!r.onSheet,
                        fbInt(r.glazeTotal, 0)];

/* ---- what the screens draw ------------------------------------------------ */
const fbYes = v => fbTxt(v).trim().toLowerCase() === "yes";
const fbActive = f => fbYes(f && f.Active);
const fbOnSheet = f => fbYes(f && f.OnSheet) || fbActive(f);
/** "" nothing fabricated · "lavender" started · "purple" every one done. */
function fbColour(done, total) {
  const t = Math.max(0, fbInt(total, 0));
  if (!(t > 0)) return "";
  const d = fbClamp(done, t);
  return d >= t ? "purple" : d > 0 ? "lavender" : "";
}
function fbRollUp(colours) {
  const c = (colours || []).filter(x => x != null);
  if (!c.length) return "";
  if (c.every(x => x === "sheet")) return "sheet";
  if (c.every(x => x === "purple" || x === "sheet")) return "purple";
  return c.some(x => x) ? "lavender" : "";
}
/* ---- finished on the sheet (owner, 2026-09-28) -------------------------------
   "ready to fit, customer won't take but ready, and collect & supply should be
   green as they are finished in the main excel sheet". Exactly those three
   sections - not "Can sell as second hand", not In production. A job there is
   SHOWN finished: every line full, colour "sheet" (green), not tappable. It is
   display only: the list's counters are untouched (kept on `raw`), nothing is
   written, and the colour painter skips the job. */
const FB_SHEET_DONE = [/^READY TO FIT/, /^READY CUSTOMER WON/, /^COLLECT SUPPLY/];
const fbSheetDone = section => { const k = fbKey(section); return !!k && FB_SHEET_DONE.some(r => r.test(k)); };
/** One list row as the boards read it. Display clamp only. */
function fbRecord(it) {
  const f = (it && it.fields) || {};
  const tp = fbTxt(f.Title).split("|");
  const g = { id: fbTxt(it && it.id), title: fbTxt(f.Title),
              job: fbKey(f.Job) || fbKey(tp[0]), group: fbKey(f.Group) || fbKey(tp[1]),
              groupSeq: fbNum(f.GroupSeq, 999), customer: fbTxt(f.Customer), comment: fbTxt(f.Comment),
              doors: fbTxt(f.Doors), seq: fbNum(f.Seq, 99999), section: fbTxt(f.Section).trim(),
              active: fbActive(f), onSheet: fbOnSheet(f), fedAt: fbTxt(f.FedAt),
              doneAt: fbTxt(f.DoneAt), doneBy: fbTxt(f.DoneBy), urgent: fbTxt(f.Urgent),
              modified: fbTxt(f.Modified),
              by: {}, at: {}, lines: [], extra: [], raw: {} };
  g.sheetDone = fbSheetDone(g.section);
  /* door glazing: its own line, drawn under the group's others, in NONE of the
     fabrication totals, colours or tabs below (`extra`, never `lines`) */
  {
    const k = FB_GLAZE;
    const t = FB_GLAZE_GROUPS.indexOf(g.group) >= 0 ? Math.max(0, fbInt(f[FB_TOTAL_FIELD[k]], 0)) : 0;
    const real = fbClamp(f[FB_DONE_FIELD[k]], t);
    const d = g.sheetDone ? t : real;
    g.raw[k] = real; g[k] = d; g[k + "Total"] = t;
    g.by[k] = fbTxt(f[FB_BY_FIELD[k]]); g.at[k] = fbTxt(f[FB_AT_FIELD[k]]);
    if (t > 0) g.extra.push({ part: k, label: FB_PART_LABEL[k], done: d, total: t, by: g.by[k], at: g.at[k],
                              colour: g.sheetDone ? "sheet" : fbColour(d, t) });
  }
  let done = 0, total = 0;
  FB_PARTS.forEach(k => {
    const t = Math.max(0, fbInt(f[FB_TOTAL_FIELD[k]], 0));
    const real = fbClamp(f[FB_DONE_FIELD[k]], t);
    const d = g.sheetDone ? t : real;               // shown full; the list is not touched
    g.raw[k] = real;
    g[k] = d; g[k + "Total"] = t;
    g.by[k] = fbTxt(f[FB_BY_FIELD[k]]); g.at[k] = fbTxt(f[FB_AT_FIELD[k]]);
    if (t > 0) {
      g.lines.push({ part: k, label: FB_PART_LABEL[k], done: d, total: t,
                     by: g.by[k], at: g.at[k], colour: g.sheetDone ? "sheet" : fbColour(d, t) });
      done += d; total += t;
    }
  });
  g.done = done; g.total = total; g.left = Math.max(0, total - done);
  g.colour = fbRollUp(g.lines.map(l => l.colour));
  g.finished = total > 0 && done >= total;
  return g;
}
/** One card per job, groups in the sheet's order. De-duplicated by Title, the
    oldest id winning, as every list here does. `glassOf(job)` (fbGlassOf) is
    the job's glass word; left out, no card has one. */
function fbCards(items, keep, glassOf) {
  const best = {};
  (items || []).forEach(it => {
    if (!it) return;
    const f = it.fields || {};
    if (keep && !keep(f)) return;
    const t = fbKey(f.Title) || (fbKey(f.Job) + "|" + fbKey(f.Group));
    if (!t || t === "|") return;
    const prev = best[t];
    if (prev && (Number(prev.id) || 0) <= (Number(it.id) || 0)) return;
    best[t] = it;
  });
  const byJob = {};
  Object.keys(best).map(t => fbRecord(best[t])).forEach(r => {
    if (!r.job) return;
    const c = byJob[r.job] || (byJob[r.job] = { job: r.job, customer: "", comment: "", doors: "",
      seq: r.seq, section: "", active: false, doneAt: "", doneBy: "", glass: "", groups: [], done: 0, total: 0 });
    ["customer", "comment", "section", "doors"].forEach(k => { if (!c[k] && r[k]) c[k] = r[k]; });
    if (r.seq < c.seq) c.seq = r.seq;
    if (r.active) c.active = true;
    if (r.doneAt && (!c.doneAt || Date.parse(r.doneAt) > Date.parse(c.doneAt))) {
      c.doneAt = r.doneAt; c.doneBy = r.doneBy;
    }
    c.groups.push(r);
    c.done += r.done; c.total += r.total;
  });
  const out = Object.keys(byJob).map(j => {
    const c = byJob[j];
    c.groups.sort((a, b) => (a.groupSeq - b.groupSeq) || (a.group < b.group ? -1 : a.group > b.group ? 1 : 0));
    c.colour = fbRollUp(c.groups.map(g => g.colour));
    c.sheetDone = c.groups.some(g => g.sheetDone);
    c.finished = c.sheetDone || (c.total > 0 && c.done >= c.total);
    c.left = Math.max(0, c.total - c.done);
    if (typeof glassOf === "function") { try { c.glass = fbTxt(glassOf(c.job)); } catch (e) { c.glass = ""; } }
    /* a Door glazing line carries the job's glass word: its chip sits beside it */
    c.groups.forEach(g => g.extra.forEach(l => { l.glass = c.glass; }));
    /* decided on the WHOLE job, before any view narrows the card: glass ready
       and nothing fabricated at all */
    c.glassStart = fbGlassStart(c);
    return c;
  });
  out.sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
  return out;
}
/** The office's board: every job still on the sheet, every section. */
const fbOfficeBoard = (items, glassOf) => fbCards(items, fbOnSheet, glassOf);
/** One job's card for the drawer, or null. */
function fbJobCard(items, job, glassOf) {
  const want = fbKey(job);
  if (!want) return null;
  return fbCards(items, f => fbOnSheet(f) &&
    (fbKey(f.Job) || fbKey(fbTxt(f.Title).split("|")[0])) === want, glassOf)[0] || null;
}
/** The tablet's two tabs, in one pass: On floor = In production and not
    finished; Finished = every other on-sheet card. */
function fbTabs(items, glassOf) {
  const out = { floor: [], finished: [] };
  fbOfficeBoard(items, glassOf).forEach(c => out[c.active && !c.finished ? "floor" : "finished"].push(c));
  return out;
}
/** A job number, a customer or a group, matched anywhere. */
function fbFilter(cards, q) {
  const want = fbTxt(q).trim().toLowerCase();
  if (!want) return (cards || []).slice();
  return (cards || []).filter(c =>
    (c.job + " " + c.customer + " " + c.groups.map(g => g.group).join(" ")).toLowerCase().indexOf(want) >= 0);
}
/** Which tab a search shows, and how many matches the other one holds. */
function fbSearchTab(tabs, q, current) {
  const S = fbST();
  const r = S.tabSearch(tabs, current, q, fbFilter);
  return { tab: r.tab, other: r.more };
}

/* ---- who may move what (per group AND part, owner 2026-09-28) ----------------
   `Stages` in `Station people` is a comma list of entries, typed by hand:
     GROUP               every part of that group   ("PVC DOOR")
     GROUP:parts         only those parts, joined by +   ("PVC SMART:sashes",
                         "PVC DOOR:frames+transoms")
     ALL / ALL:parts     every fed group (the 255-character column fills up)
   Case and spaces do not matter; the same group twice is the union; an unknown
   group or part word is ignored. The ":" is split off BEFORE the group is put
   through fbKey, which would otherwise turn it into a space. */
function fbParseStages(text) {
  const out = {};
  let prev = "";                     // the group the last entry named
  fbTxt(text).split(/[,;]+/).forEach(entry => {
    /* "+" is the joiner, but "CASEMENT WINDOWS:frames,sashes" is what a hand
       types (review N2): an entry made only of part words continues the
       previous entry's group */
    const words = entry.split(/[+\s]+/).map(fbLow).filter(Boolean);
    if (prev && entry.indexOf(":") < 0 && words.length && words.every(w => FB_TAP_PARTS.indexOf(w) >= 0))
      entry = prev + ":" + entry;
    const i = entry.indexOf(":");
    const gk = fbKey(i >= 0 ? entry.slice(0, i) : entry);
    if (!gk) return;
    prev = entry.slice(0, i >= 0 ? i : entry.length);
    /* a bare group is frames + sashes + transoms and NOT door glazing: glazing
       is a role and has to be named ("PVC DOOR:glazing", "ALL:glazing") */
    const parts = i < 0 ? FB_PARTS.slice()
      : entry.slice(i + 1).split(/[+\s]+/).map(fbLow).filter(p => FB_TAP_PARTS.indexOf(p) >= 0);
    if (!parts.length) return;
    const groups = gk === "ALL" ? FB_GROUP_KEYS : FB_GROUP_KEYS.indexOf(gk) >= 0 ? [gk] : [];
    groups.forEach(g => {
      /* ... and it only exists on the two groups that are glazed here */
      const mine = parts.filter(p => p !== FB_GLAZE || FB_GLAZE_GROUPS.indexOf(g) >= 0);
      if (!mine.length) return;
      const k = g.toLowerCase(), m = out[k] || (out[k] = {});
      mine.forEach(p => { m[p] = true; });
    });
  });
  return out;
}
/** The people of this station, each with `stages` (the group keys they have
    any part of, lower case) and `parts` ({ group: { frames, sashes, transoms } }).
    ST.stationPeople still does the station / Active / PIN / sort work. */
function fbPeople(items) {
  const S = fbST();
  if (!S) return [];
  const parsed = {};
  const norm = (items || []).map(it => {
    if (!it || !it.fields) return it;
    const p = fbParseStages(it.fields.Stages);
    parsed[fbTxt(it.id)] = p;
    return { id: it.id, fields: Object.assign({}, it.fields, { Stages: Object.keys(p).join(",") }) };
  });
  return S.stationPeople(norm, FB_NAME, FB_GROUP_KEYS.map(k => k.toLowerCase()))
    .map(p => Object.assign(p, { parts: parsed[p.id] || {} }));
}
/** May this person move this group's lines - or, given a part, that part of
    it? A person object without `parts` (built by hand) holds every part of
    each group in `stages`, today's meaning. */
function fbEligible(person, group, part) {
  if (!person || !fbKey(group)) return false;
  const g = fbKey(group).toLowerCase();
  if (!person.parts) return (person.stages || []).indexOf(g) >= 0 && (!part || FB_PARTS.indexOf(fbLow(part)) >= 0);
  const m = person.parts[g];
  if (!m) return false;
  return part ? !!m[fbLow(part)] : FB_TAP_PARTS.some(p => m[p]);
}

/* ---- the tablet's view filter (owner, 2026-10-01) -----------------------------
   "Everything | My work | Assigned to me". DISPLAY ONLY: it decides which
   lines are drawn, never which may be tapped (fbCanTap is the gate).
     all      - the cards as they are;
     mine     - only the parts the person is eligible for (fbEligible, so the
                Stages syntax incl. GROUP:parts and ALL:parts);
     assigned - only the parts the person holds an Assigned row on (Qty > 0;
                a Requested row does not count).
   A group left with no line is dropped, a card left with no group is dropped.
   Kept groups and cards are copies with their counts and colours re-added from
   what is still drawn; the records the tap gate reads are not touched. */
const FB_VIEWS = ["all", "mine", "assigned"];
const fbViewOf = v => (FB_VIEWS.indexOf(v) >= 0 ? v : "all");
function fbViewFilter(cards, person, mode, idx) {
  const m = fbViewOf(mode);
  if (m === "all") return (cards || []).slice();
  const keepPart = (c, g, part) => m === "mine" ? fbEligible(person, g.group, part)
    : fbMine(idx, person, c.job, g.group, part) > 0;
  const out = [];
  (cards || []).forEach(c => {
    const groups = [];
    (c.groups || []).forEach(g => {
      const lines = (g.lines || []).filter(l => keepPart(c, g, l.part));
      /* door glazing is a role with no assignment: My work shows it to
         somebody who has the role, Assigned to me never does */
      const extra = m === "mine" ? (g.extra || []).filter(l => fbEligible(person, g.group, l.part)) : [];
      if (!lines.length && !extra.length) return;
      const done = lines.reduce((n, l) => n + l.done, 0), total = lines.reduce((n, l) => n + l.total, 0);
      groups.push(Object.assign({}, g, { lines: lines, extra: extra, done: done, total: total,
        left: Math.max(0, total - done), colour: fbRollUp(lines.map(l => l.colour)) }));
    });
    if (!groups.length) return;
    const done = groups.reduce((n, g) => n + g.done, 0), total = groups.reduce((n, g) => n + g.total, 0);
    out.push(Object.assign({}, c, { groups: groups, done: done, total: total, left: Math.max(0, total - done),
      colour: fbRollUp(groups.map(g => g.colour)) }));
  });
  return out;
}
/** The two tabs under a view (review M2, 2026-10-01). Under My work and
    Assigned to me a card is placed by the lines it SHOWS: every shown line
    done -> Finished; any shown line not done, on an active In production job
    -> On floor. A finished-on-sheet (green) card stays in Finished. Everything
    is the tabs as they are. Board order, then urgent first, as fbTabs' callers. */
function fbViewTabs(tabs, person, mode, idx) {
  const m = fbViewOf(mode);
  const t = tabs || { floor: [], finished: [] };
  if (m === "all") return { floor: (t.floor || []).slice(), finished: (t.finished || []).slice() };
  const cards = fbViewFilter((t.floor || []).concat(t.finished || []), person, m, idx);
  cards.sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
  const out = { floor: [], finished: [] };
  cards.forEach(c => {
    const shownDone = c.groups.every(g => g.lines.concat(g.extra || []).every(l => l.done >= l.total));
    out[c.active && !c.sheetDone && !shownDone ? "floor" : "finished"].push(c);
  });
  return { floor: fbUrgentFirst(out.floor), finished: fbUrgentFirst(out.finished) };
}
/** The remembered choice, per person: { "<person>": "all" | "mine" | "assigned" }. */
const fbViewFor = (store, name) => fbViewOf((store || {})[fbTxt(name)]);

/* ---- the writes one tap makes ------------------------------------------------ */
function fbApplyTap(row, part, delta) {
  const k = fbTxt(part).trim().toLowerCase();
  if (FB_TAP_PARTS.indexOf(k) < 0) return null;
  const total = Math.max(0, fbInt(row && row[k + "Total"], 0));
  const now = fbClamp(row && row[k], total);
  if (delta === "all") return total;
  if (delta === "none") return 0;
  const d = Number(delta);
  return isFinite(d) ? fbClamp(now + d, total) : now;
}
/** The counter PATCH, from the floor or the office: that part's count, its
    By/At and the last-touch pair. Nothing else can get in. */
function fbTapFields(part, value, who, at) {
  const k = fbTxt(part).trim().toLowerCase();
  if (FB_TAP_PARTS.indexOf(k) < 0) return null;
  const when = fbTxt(at) || new Date().toISOString(), name = fbTxt(who);
  const out = {};
  out[FB_DONE_FIELD[k]] = Math.max(0, fbInt(value, 0));
  out[FB_BY_FIELD[k]] = name; out[FB_AT_FIELD[k]] = when;
  out.DoneBy = name; out.DoneAt = when;
  return out;
}
const fbOfficeFields = fbTapFields;
function fbFloorOnly(fields) { const S = fbST(); return S ? S.floorOnly(fields, FAB) : {}; }
/** One `Station log` line (tablet only): GlassType carries the group. */
function fbLogEntry(e) {
  return { job: fbKey(e && e.job), station: FB_NAME, type: fbKey(e && e.group),
           stage: fbTxt(e && e.part).trim().toLowerCase(),
           from: fbInt(e && e.from, 0), to: fbInt(e && e.to, 0), who: fbTxt(e && e.who), at: fbTxt(e && e.at) };
}
const fbLogWords = (job, group, part) =>
  "Fabrication: " + fbKey(job) + " " + fbKey(group) + " " + fbTxt(part).trim().toLowerCase();
/** A queued tap meets somebody else's write: the other tablets' three answers
    (rise: re-base; fall stamped after the tap: drop; anything else: keep). */
function fbRebase(e, fields) {
  const f = fields || {};
  const k = fbTxt(e && e.part).trim().toLowerCase();
  if (FB_TAP_PARTS.indexOf(k) < 0) return { action: "keep" };
  const now = Number(f[FB_DONE_FIELD[k]]), was = Number(e && e.from);
  if (!isFinite(now) || !isFinite(was)) return { action: "keep" };
  if (now < was) {
    const rowAt = Date.parse(fbTxt(f.DoneAt)), tapAt = Date.parse(fbTxt(e && e.at));
    if (!isFinite(rowAt) || !isFinite(tapAt) || rowAt <= tapAt) return { action: "keep" };
    return { action: "drop" };
  }
  if (now <= was) return { action: "keep" };
  const total = Math.max(0, fbInt(f[FB_TOTAL_FIELD[k]], 0));
  return { action: "rebase", value: fbClamp(Math.round(now + (Number(e.value) - was)), total), from: Math.round(now) };
}
function fbCardSig(c) {
  return JSON.stringify([c.job, c.customer, c.comment, c.doors, c.section, c.seq, c.finished, c.colour, c.glass,
    c.groups.map(g => [g.id, g.group, g.colour, g.doors, g.urgent, FB_TAP_PARTS.map(k => [g[k], g[k + "Total"], g.by[k], g.at[k]])])]);
}

/* ---- the colour painter's rule (rule 6) -------------------------------------
   What one product cell's fill says, as this station reads it: "" white or no
   fill, "process" its lavender, "done" its purple, "other" anything else -
   yellow, gold, the Cut green, red - which it never paints over. */
function fbCellWord(hex) {
  let h = fbTxt(hex).trim().toUpperCase().replace(/^#/, "");
  if (h.length === 8) h = h.slice(-6);
  if (!h || h === "FFFFFF") return "";
  if (h === FB_PROCESS_HEX.slice(1)) return "process";
  if (h === FB_DONE_HEX.slice(1)) return "done";
  return "other";
}
const FB_WORD_HEX = { process: FB_PROCESS_HEX, done: FB_DONE_HEX, "": FB_WHITE_HEX };
/** What one cell should be painted, or null for "leave it":
      - the office's record says anything (process, done, cut...): leave it -
        fabrication never lowers the office;
      - the cell is a colour this station does not own: leave it;
      - done 0: white, but ONLY over one of this station's own two colours;
      - otherwise lavender (started) or purple (all done), unless already so. */
function fbCellWant(done, total, record, cell) {
  if (fbTxt(record).trim()) return null;
  if (cell === "other") return null;
  const t = Math.max(0, fbInt(total, 0)), d = fbClamp(done, t);
  const want = t > 0 && d >= t ? "done" : d > 0 ? "process" : "";
  if (!want) return cell === "process" || cell === "done" ? "" : null;
  return want === cell ? null : want;
}

/* ---- gold on the DOORS DONE cells when door glazing is complete (section C,
   owner 2026-10-01: "when door glazing is done make door done/golden in excel")
   The pure half: WHICH of a job's door items should be raised to done. The
   office does the raising, through the checkpoint record (app.js fabrGoldRun).
     - only a FULL count (done >= total > 0) raises anything; a partial count
       and a count that has dropped raise and clear nothing;
     - only the doors whose code is glazed in this group (never CD, SFCD, BF,
       the AC codes);
     - never a door already done, never one with a write in the air, never a
       cell carrying a colour the checkpoints do not own;
     - never over the office: a row the office (or a hand in Excel) wrote AFTER
       the glazing was recorded stands - an office clear made after the
       glazing finished is not gilded again.
   WHO SPOKE LAST IS DECIDED ON THE SERVER'S CLOCK (review C-1), never on the
   tablet's against the office browser's: `fabModified` is SharePoint's own
   `Modified` of the fabrication row as the glazing was recorded, and each
   state's `modified` is the record row's. A tablet whose clock runs fast or
   slow changes nothing. Where either stamp is missing it FAILS SAFE: an
   office / excel row is only ever raised from "process" - never from a clear.
   `stateOf(slot)` -> { status, office, modified (ms), pending, foreign }. */
function fbGoldPlan(o) {
  const t = Math.max(0, fbInt(o && o.total, 0));
  if (!(t > 0) || fbInt(o.done, 0) < t) return [];
  const fab = fbNum(o.fabModified, 0);
  return ((o.doors) || []).filter(d => d && fbDoorGlazeGroup(d.code) === fbKey(o.group)).filter(d => {
    const s = (typeof o.stateOf === "function" && o.stateOf(d.slot)) || {};
    if (s.pending || s.foreign || s.status === "done") return false;
    if (s.office) {
      const rec = fbNum(s.modified, 0);
      if (fab > 0 && rec > 0) { if (rec >= fab) return false; }
      else if (s.status !== "process") return false;
    }
    return true;
  }).map(d => d.slot);
}

/** The station report's Jobs sheet: one row per job and group. */
function fbReportJobs(data) {
  const S = fbST();
  const cap = (S && S.REPORT_CUSTOMER_MAX) || 60;
  const rows = [], jobs = [];
  /* the list's own counts (`raw`), never the "finished on the sheet" display */
  const rawDone = g => FB_PARTS.reduce((n, k) => n + (g[k + "Total"] > 0 ? (g.raw || g)[k] : 0), 0);
  ((data && data.board) || []).forEach(c => {
    jobs.push({ job: c.job, done: c.groups.reduce((n, g) => n + rawDone(g), 0), total: c.total });
    c.groups.forEach(g => { const r = g.raw || g, d = rawDone(g); rows.push([c.job, fbStrip(c.customer, cap),
      c.section, g.group, g.doors, r.frames, g.framesTotal, r.sashes, g.sashesTotal, r.transoms, g.transomsTotal,
      g.total, d, Math.max(0, g.total - d), g.doneBy, g.doneAt, g.total > 0 && d >= g.total ? "Yes" : "No",
      /* door glazing: its own count, in none of the totals to its left */
      g.glazingTotal > 0 ? r.glazing : "", g.glazingTotal > 0 ? g.glazingTotal : ""]); });
  });
  return { columns: ["Job", "Customer", "Section", "Group", "Doors", "Frames done", "Frames",
                     "Sashes done", "Sashes", "Transoms done", "Transoms", "Total", "Done", "Left",
                     "Last moved by", "When", "Complete", "Door glazing done", "Door glazing"],
           rows: rows, jobs: jobs };
}

/* ==== Part B: assignments, approval, urgent, notifications ===================
   `Fabrication assignments` (Floor stations): one row per piece of work given
   to a person - Title `JOB|GROUP|PART|<random 6>`, Job, Group, Part, Person,
   Qty, Status, RequestedBy/At, DecidedBy/At. Status is Requested (the tablet's
   Take), Assigned (the office assigns, or approves a request), Refused or
   Removed. Rows are never deleted.

   Who writes what, and nothing else:
     - the TABLET only ever creates a Requested row for the person signed in
       (fbRequestFields). It never PATCHes a row: not Status, not Qty, not Person;
     - the OFFICE creates Assigned rows and PATCHes Status/Qty/DecidedBy/At
       (fbAssignFields, fbApproveFields, fbDecideFields). */
const FB_ASSIGN_LIST = "Fabrication assignments";
const FB_ASSIGN_FIELDS = ["Title", "Job", "Group", "Part", "Person", "Qty", "Status",
                          "RequestedBy", "RequestedAt", "DecidedBy", "DecidedAt"];
const FB_STATUS = { requested: "Requested", assigned: "Assigned", refused: "Refused", removed: "Removed" };
const fbLow = v => fbTxt(v).trim().toLowerCase();
const fbPartKey = p => (FB_PARTS.indexOf(fbLow(p)) >= 0 ? fbLow(p) : "");
/** The key of one part of one group of one job. */
const fbLineKey = (job, group, part) => fbKey(job) + "|" + fbKey(group) + "|" + fbPartKey(part);
function fbRand6() {
  let s = "";
  for (let i = 0; i < 6; i++) s += "abcdefghjkmnpqrstuvwxyz23456789".charAt(Math.floor(Math.random() * 31));
  return s;
}
const fbAssignTitle = (job, group, part, rnd) => fbLineKey(job, group, part) + "|" + (rnd || fbRand6());

/** The list's rows, as the screens read them. A row with no job, group, part
    or person is not an assignment. */
function fbAssignRows(items) {
  const out = [];
  (items || []).forEach(it => {
    const f = (it && it.fields) || {};
    const tp = fbTxt(f.Title).split("|");
    const r = { id: fbTxt(it && it.id), job: fbKey(f.Job || tp[0]), group: fbKey(f.Group || tp[1]),
                part: fbPartKey(f.Part || tp[2]), person: fbTxt(f.Person).trim(),
                qty: Math.max(0, fbInt(f.Qty, 0)), status: fbLow(f.Status),
                requestedBy: fbTxt(f.RequestedBy), requestedAt: fbTxt(f.RequestedAt),
                decidedBy: fbTxt(f.DecidedBy), decidedAt: fbTxt(f.DecidedAt) };
    if (!r.job || !r.group || !r.part || !r.person) return;
    r.key = fbLineKey(r.job, r.group, r.part);
    out.push(r);
  });
  out.sort((a, b) => (Number(a.id) || 0) - (Number(b.id) || 0));
  return out;
}
/** Rows by line: { key: { assigned: [], requested: [] } }. Refused and Removed
    rows are history and are in neither. */
function fbAssignIndex(rows) {
  const out = {};
  (rows || []).forEach(r => {
    const b = out[r.key] || (out[r.key] = { assigned: [], requested: [] });
    if (r.status === "assigned") b.assigned.push(r);
    else if (r.status === "requested") b.requested.push(r);
  });
  return out;
}
const fbLineOf = (idx, job, group, part) => (idx || {})[fbLineKey(job, group, part)] || { assigned: [], requested: [] };
const fbAssignedSum = (idx, job, group, part, exceptId) =>
  fbLineOf(idx, job, group, part).assigned.filter(r => r.id !== fbTxt(exceptId)).reduce((n, r) => n + r.qty, 0);
/** How many of that part a person holds (Assigned), 0 for none. */
const fbMine = (idx, person, job, group, part) => !person ? 0 :
  fbLineOf(idx, job, group, part).assigned.filter(r => r.person === person.name).reduce((n, r) => n + r.qty, 0);
const fbRequested = (idx, person, job, group, part) => !!person &&
  fbLineOf(idx, job, group, part).requested.some(r => r.person === person.name);
/** A split may not give away more than the part's total on the sheet. `qty` is
    what is being added (or confirmed on approval, `exceptId` = that row). */
function fbSplitCheck(idx, job, group, part, total, qty, exceptId) {
  const t = Math.max(0, fbInt(total, 0)), q = fbInt(qty, 0);
  const have = fbAssignedSum(idx, job, group, part, exceptId);
  const free = Math.max(0, t - have);
  if (!(q > 0)) return { ok: false, free: free, msg: "The quantity has to be at least 1." };
  if (have + q > t) return { ok: false, free: free,
    msg: "Only " + free + " of " + t + " " + fbPartKey(part) + " are unassigned, so " + q + " cannot be given." };
  return { ok: true, free: free, msg: "" };
}
/** THE TABLET GATE (Part B). `state` is what is known about the assignments
    list, and the gate FAILS CLOSED (review P1, 2026-09-25):
      true  - read at least once: eligible AND holding an Assigned row;
      false - positively missing (the list is not there): Part A's gate,
              eligible may tap;
      null  - not known yet (never read successfully): nothing may be tapped. */
function fbCanTap(person, group, part, idx, state, job) {
  if (!fbEligible(person, group, part)) return false;
  /* door glazing is a role and has no assignment (owner, 2026-10-01): the
     person who has it may tap, whatever the assignments list says */
  if (fbLow(part) === FB_GLAZE) return true;
  if (state === false) return true;
  if (state !== true) return false;
  return fbMine(idx, person, job, group, part) > 0;
}
/** One counter per PART, not per person (review P2): with assignments on, All
    and None - which set the whole line - are allowed only to somebody holding
    all of it. − and + are always within the clamp. */
function fbActAllowed(person, idx, state, job, group, part, total, act) {
  if (act !== "all" && act !== "none") return true;
  if (fbLow(part) === FB_GLAZE) return true;           // no assignment on door glazing: All / None allowed
  if (state !== true) return true;                     // Part A, or locked anyway
  return fbMine(idx, person, job, group, part) >= Math.max(0, fbInt(total, 0));
}

/* ---- the bodies (the only shapes either side can send) ---- */
/** The tablet's one write to this list: a Requested row for the person
    signed in. */
function fbRequestFields(job, group, part, person, qty, at) {
  const who = fbTxt(person).trim(), when = fbTxt(at) || new Date().toISOString();
  if (!fbPartKey(part) || !who || !fbKey(job) || !fbKey(group)) return null;
  return { Title: fbAssignTitle(job, group, part), Job: fbKey(job), Group: fbKey(group), Part: fbPartKey(part),
           Person: who, Qty: Math.max(0, fbInt(qty, 0)), Status: FB_STATUS.requested,
           RequestedBy: who, RequestedAt: when };
}
/** The office's new assignment. */
function fbAssignFields(job, group, part, person, qty, who, at) {
  const when = fbTxt(at) || new Date().toISOString();
  return { Title: fbAssignTitle(job, group, part), Job: fbKey(job), Group: fbKey(group), Part: fbPartKey(part),
           Person: fbTxt(person).trim(), Qty: Math.max(0, fbInt(qty, 0)), Status: FB_STATUS.assigned,
           DecidedBy: fbTxt(who), DecidedAt: when };
}
/** The office approving a request, with the quantity it confirms. */
const fbApproveFields = (qty, who, at) => ({ Status: FB_STATUS.assigned, Qty: Math.max(0, fbInt(qty, 0)),
  DecidedBy: fbTxt(who), DecidedAt: fbTxt(at) || new Date().toISOString() });
/** Put an approval back (review P3: two office screens gave the same line
    away at once): the row is a request again, with the quantity it asked for. */
const fbUnapproveFields = (qty, who, at) => ({ Status: FB_STATUS.requested, Qty: Math.max(0, fbInt(qty, 0)),
  DecidedBy: fbTxt(who), DecidedAt: fbTxt(at) || new Date().toISOString() });
/** Refuse a request, or remove an assignment. Nothing else changes. */
const fbDecideFields = (status, who, at) => ({ Status: status === "refused" ? FB_STATUS.refused : FB_STATUS.removed,
  DecidedBy: fbTxt(who), DecidedAt: fbTxt(at) || new Date().toISOString() });

/* ---- urgent: `Urgent` on a Fabrication station row, office only ------------
   A comma list of `job`, `group`, `frames`, `sashes`, `transoms`. `job` is the
   job-level flag and is written on every row of the job. */
const FB_URGENT_WORDS = ["job", "group"].concat(FB_PARTS);
function fbUrgentOf(text) {
  const out = {};
  fbTxt(text).split(/[,;]+/).map(fbLow).forEach(w => { if (FB_URGENT_WORDS.indexOf(w) >= 0) out[w] = true; });
  return out;
}
function fbUrgentToggle(text, word, on) {
  const u = fbUrgentOf(text);
  if (on) u[word] = true; else delete u[word];
  return FB_URGENT_WORDS.filter(w => u[w]).join(",");
}
/** Is anything on this card urgent? And what, per group. */
function fbCardUrgent(c) {
  let any = false;
  (c.groups || []).forEach(g => {
    g.urgentOf = fbUrgentOf(g.urgent);
    if (Object.keys(g.urgentOf).length) any = true;
  });
  return any;
}
/** Urgent cards first, each side keeping the board's own order. */
function fbUrgentFirst(cards) {
  const u = [], rest = [];
  (cards || []).forEach(c => (fbCardUrgent(c) ? u : rest).push(c));
  return u.concat(rest);
}

/* ---- notifications on the tablet --------------------------------------------
   One key per thing worth telling the person signed in: each of their
   assignment rows in its current state, and each urgent flag on a line they
   hold. A key they have not seen is a notice. The keys change when the thing
   changes, so a changed Qty or a new urgent flag is a new notice. */
function fbNotices(rows, cards, person, idx) {
  if (!person) return [];
  const out = [];
  (rows || []).forEach(r => {
    if (r.person !== person.name) return;
    if (r.status === "requested") return;                    // their own Take: nothing to tell them
    const words = r.status === "assigned" ? "assigned to you: " + r.qty + " " + r.part
      : r.status === "refused" ? "request refused: " + r.part : "assignment removed: " + r.part;
    out.push({ key: "a|" + r.id + "|" + r.status + "|" + r.qty, job: r.job,
               text: r.job + " " + r.group + " — " + words });
  });
  (cards || []).forEach(c => (c.groups || []).forEach(g => {
    const u = fbUrgentOf(g.urgent);
    if (!Object.keys(u).length) return;
    const held = FB_PARTS.some(p => fbMine(idx, person, c.job, g.group, p) > 0);
    if (!held) return;
    out.push({ key: "u|" + c.job + "|" + g.group + "|" + Object.keys(u).sort().join(","), job: c.job,
               text: c.job + " " + g.group + " — urgent" });
  }));
  return out;
}
/** Seen state for one person: { key: 1 }. The first time a person is seen at
    all, everything current is taken as seen (a baseline, not a storm). */
function fbSeenFor(store, name, notices) {
  const s = store || {};
  if (!s[name]) { s[name] = {}; (notices || []).forEach(n => { s[name][n.key] = 1; }); }
  return s[name];
}
const fbUnseen = (notices, seen) => (notices || []).filter(n => !(seen || {})[n.key]);
/** Drop every seen key that is no longer a current notice (its row changed
    state, or the job left the board), so the store does not grow for ever
    (review P7). Answers whether anything was dropped. */
function fbPruneSeen(seen, notices) {
  const live = {};
  (notices || []).forEach(n => { live[n.key] = 1; });
  let changed = false;
  Object.keys(seen || {}).forEach(k => { if (!live[k]) { delete seen[k]; changed = true; } });
  return changed;
}

/* ---- the station definition (docs/STATIONS.md, "Adding a station") -------- */
const FAB = {
  key: "fabrication", name: FB_NAME, list: FB_LIST, site: FB_SITE, assignList: FB_ASSIGN_LIST,
  stages: FB_GROUP_KEYS.map(k => k.toLowerCase()),
  stageLabel: () => "Fabrication",
  reportStages: [FB_STAGE],
  reportLogStages: () => FB_PARTS.slice(),
  reportGroupLabel: "Group",
  reportJobs: fbReportJobs,
  fields: FB_FIELDS, feederFields: FB_FEEDER_FIELDS, floorFields: FB_FLOOR_FIELDS,
  counterFields: FB_COUNTER_FIELDS, seedFields: FB_SEED_FIELDS, feederWrites: FB_FEEDER_WRITES,
  goneFields: FB_GONE,
  feederOf: fbFeederFields, seedOf: fbSeedFields, hashOf: fbHashRow, rowOrder: fbRowOrder
};

const FABC = {
  FAB, FB_LIST, FB_NAME, FB_SITE, FB_STAGE, FB_PARTS, FB_PART_LABEL, FB_PART_SUB,
  FB_TOTAL_FIELD, FB_DONE_FIELD, FB_BY_FIELD, FB_AT_FIELD,
  FB_FIELDS, FB_FEEDER_FIELDS, FB_FLOOR_FIELDS, FB_COUNTER_FIELDS, FB_OFFICE_FIELDS,
  FB_SEED_FIELDS, FB_FEEDER_WRITES, FB_GONE, FB_WINDOW_GROUPS, FB_DOOR_GROUPS, FB_GROUP_KEYS,
  FB_PROCESS_HEX, FB_DONE_HEX, FB_WHITE_HEX, FB_WORD_HEX,
  fbKey, fbStrip, fbAllowed, fbIsDoorGroup, fbTitle, fbDoorLabels, fbCommentOf,
  fbSlice, fbRowOrder, fbFeederFields, fbSeedFields, fbHashRow,
  fbActive, fbOnSheet, fbColour, fbRollUp, fbRecord, fbCards, fbOfficeBoard, fbJobCard,
  fbTabs, fbFilter, fbSearchTab, fbPeople, fbEligible, fbParseStages, fbSheetDone,
  FB_VIEWS, fbViewOf, fbViewFilter, fbViewTabs, fbViewFor,
  FB_GLAZE, FB_TAP_PARTS, FB_GLAZE_GROUPS, FB_GLAZE_NEVER, FB_NEW_FEEDER_FIELDS,
  fbDoorGlazeGroup, fbGlazeTotals, fbGlassStatus, fbGlassOf, fbGlassChip, fbGlassReady, fbGlassStart, fbGlassFirst,
  fbGoldPlan,
  fbApplyTap, fbTapFields, fbOfficeFields, fbFloorOnly, fbLogEntry, fbLogWords, fbRebase, fbCardSig,
  fbCellWord, fbCellWant, fbReportJobs,
  FB_ASSIGN_LIST, FB_ASSIGN_FIELDS, FB_STATUS, FB_URGENT_WORDS, fbLineKey, fbAssignTitle,
  fbAssignRows, fbAssignIndex, fbLineOf, fbAssignedSum, fbMine, fbRequested, fbSplitCheck, fbCanTap,
  fbActAllowed, fbRequestFields, fbAssignFields, fbApproveFields, fbUnapproveFields, fbDecideFields,
  fbPruneSeen,
  fbUrgentOf, fbUrgentToggle, fbCardUrgent, fbUrgentFirst, fbNotices, fbSeenFor, fbUnseen
};
if (typeof window !== "undefined") window.FABC = FABC;
else if (typeof globalThis !== "undefined") globalThis.FABC = FABC;
if (typeof module !== "undefined" && module.exports) module.exports = FABC;
