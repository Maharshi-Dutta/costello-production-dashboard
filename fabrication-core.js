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

/* ---- the columns ----------------------------------------------------------
   The feeder writes Title, the job facts and FedAt/FedBy. The floor (and the
   office's own board) writes FB_FLOOR_FIELDS and nothing else. `Urgent` is the
   office's alone and is Part B's: it is in FB_FIELDS so a read carries it, and
   in neither write list. No seed: there is no office record of fabrication. */
const FB_FEEDER_FIELDS = ["Job", "Group", "GroupSeq", "Customer", "Comment", "Seq", "Doors",
                          "Frames", "Sashes", "Transoms", "Section", "Active", "OnSheet"];
const FB_FLOOR_FIELDS = ["FramesDone", "SashesDone", "TransomsDone",
                         "FramesBy", "FramesAt", "SashesBy", "SashesAt", "TransomsBy", "TransomsAt",
                         "DoneBy", "DoneAt"];
const FB_COUNTER_FIELDS = ["FramesDone", "SashesDone", "TransomsDone"];
const FB_OFFICE_FIELDS = ["Urgent"];
const FB_SEED_FIELDS = [];
const FB_FIELDS = ["Title"].concat(FB_FEEDER_FIELDS, ["FedAt", "FedBy"], FB_FLOOR_FIELDS, FB_OFFICE_FIELDS);
const FB_FEEDER_WRITES = ["Title"].concat(FB_FEEDER_FIELDS, ["FedAt", "FedBy"]);
/* a row whose job or group has left the sheet: both flags off, never deleted
   (Active alone means "In production", so it cannot also mean "gone" - B31) */
const FB_GONE = { Active: "No", OnSheet: "No" };

/* ---- which product groups (owner, 2026-09-25) -----------------------------
   An ALLOW-list, matched through fbKey. The owner's names, as given. Two of
   them match no header of the 2026-09 workbook copy and are kept as written
   until the owner says which headers they mean - see the brief's report:
   "ARCH & ANGLESTH W" (the sheet has "ARCH ANGLES" and "TH W") and
   "COMPOSITE PVC DOOR" (the sheet has "COMPOSITE"). An unmatched name feeds
   nothing; it is one edit here once confirmed. */
const FB_WINDOW_GROUPS = ["CASEMENT WINDOWS", "POLARIS 85MM CASEMENT", "4000 CASEMENT", "7000 CASEMENT",
  "POLARIS 85 TILT & TURN", "4000 TILT & TURN", "7000 TILT & TURN", "PVC FRENCH WDS",
  "ARCH & ANGLESTH W", "ALU CLAD WINDOWS", "ALUCLAD TILT & TURN"];
const FB_DOOR_GROUPS = ["SIDELIGHTS", "SUPER DOOR", "BIFOLD", "PVC SMART", "COMPOSITE PVC DOOR", "PVC DOOR"];

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
                               frames: f, sashes: s, transoms: t }, base));
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
           Section: fbTxt(r.section), Active: r.active ? "Yes" : "No", OnSheet: r.onSheet ? "Yes" : "No" };
}
const fbSeedFields = () => ({});
const fbHashRow = r => [r.title, r.job, r.group, r.groupSeq, r.customer, r.comment, r.seq, r.doors,
                        r.frames, r.sashes, r.transoms, r.section, !!r.active, !!r.onSheet];

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
  if (c.every(x => x === "purple")) return "purple";
  return c.some(x => x) ? "lavender" : "";
}
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
              by: {}, at: {}, lines: [] };
  let done = 0, total = 0;
  FB_PARTS.forEach(k => {
    const t = Math.max(0, fbInt(f[FB_TOTAL_FIELD[k]], 0));
    const d = fbClamp(f[FB_DONE_FIELD[k]], t);
    g[k] = d; g[k + "Total"] = t;
    g.by[k] = fbTxt(f[FB_BY_FIELD[k]]); g.at[k] = fbTxt(f[FB_AT_FIELD[k]]);
    if (t > 0) {
      g.lines.push({ part: k, label: FB_PART_LABEL[k], done: d, total: t,
                     by: g.by[k], at: g.at[k], colour: fbColour(d, t) });
      done += d; total += t;
    }
  });
  g.done = done; g.total = total; g.left = Math.max(0, total - done);
  g.colour = fbRollUp(g.lines.map(l => l.colour));
  g.finished = total > 0 && done >= total;
  return g;
}
/** One card per job, groups in the sheet's order. De-duplicated by Title, the
    oldest id winning, as every list here does. */
function fbCards(items, keep) {
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
      seq: r.seq, section: "", active: false, doneAt: "", doneBy: "", groups: [], done: 0, total: 0 });
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
    c.finished = c.total > 0 && c.done >= c.total;
    c.left = Math.max(0, c.total - c.done);
    return c;
  });
  out.sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
  return out;
}
/** The office's board: every job still on the sheet, every section. */
const fbOfficeBoard = items => fbCards(items, fbOnSheet);
/** One job's card for the drawer, or null. */
function fbJobCard(items, job) {
  const want = fbKey(job);
  if (!want) return null;
  return fbCards(items, f => fbOnSheet(f) &&
    (fbKey(f.Job) || fbKey(fbTxt(f.Title).split("|")[0])) === want)[0] || null;
}
/** The tablet's two tabs, in one pass: On floor = In production and not
    finished; Finished = every other on-sheet card. */
function fbTabs(items) {
  const out = { floor: [], finished: [] };
  fbOfficeBoard(items).forEach(c => out[c.active && !c.finished ? "floor" : "finished"].push(c));
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

/* ---- who may move what ------------------------------------------------------
   `Stages` in `Station people` is a comma list of product group names, typed by
   hand. Each is put through fbKey before ST.stationPeople reads it, so
   "Casement  windows" and "CASEMENT WINDOWS" are the same group. */
function fbPeople(items) {
  const S = fbST();
  if (!S) return [];
  const norm = (items || []).map(it => {
    if (!it || !it.fields) return it;
    const st = fbTxt(it.fields.Stages).split(/[,;]+/).map(s => fbKey(s).toLowerCase()).filter(Boolean);
    return { id: it.id, fields: Object.assign({}, it.fields, { Stages: st.join(",") }) };
  });
  return S.stationPeople(norm, FB_NAME, FB_GROUP_KEYS.map(k => k.toLowerCase()));
}
/** May this person move this group's lines? */
const fbEligible = (person, group) =>
  !!person && !!fbKey(group) && (person.stages || []).indexOf(fbKey(group).toLowerCase()) >= 0;

/* ---- the writes one tap makes ------------------------------------------------ */
function fbApplyTap(row, part, delta) {
  const k = fbTxt(part).trim().toLowerCase();
  if (FB_PARTS.indexOf(k) < 0) return null;
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
  if (FB_PARTS.indexOf(k) < 0) return null;
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
  if (FB_PARTS.indexOf(k) < 0) return { action: "keep" };
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
  return JSON.stringify([c.job, c.customer, c.comment, c.doors, c.section, c.seq, c.finished, c.colour,
    c.groups.map(g => [g.id, g.group, g.colour, g.doors, FB_PARTS.map(k => [g[k], g[k + "Total"], g.by[k], g.at[k]])])]);
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

/** The station report's Jobs sheet: one row per job and group. */
function fbReportJobs(data) {
  const S = fbST();
  const cap = (S && S.REPORT_CUSTOMER_MAX) || 60;
  const rows = [], jobs = [];
  ((data && data.board) || []).forEach(c => {
    jobs.push({ job: c.job, done: c.done, total: c.total });
    c.groups.forEach(g => rows.push([c.job, fbStrip(c.customer, cap), c.section, g.group, g.doors,
      g.frames, g.framesTotal, g.sashes, g.sashesTotal, g.transoms, g.transomsTotal,
      g.total, g.done, g.left, g.doneBy, g.doneAt, g.finished ? "Yes" : "No"]));
  });
  return { columns: ["Job", "Customer", "Section", "Group", "Doors", "Frames done", "Frames",
                     "Sashes done", "Sashes", "Transoms done", "Transoms", "Total", "Done", "Left",
                     "Last moved by", "When", "Complete"], rows: rows, jobs: jobs };
}

/* ---- the station definition (docs/STATIONS.md, "Adding a station") -------- */
const FAB = {
  key: "fabrication", name: FB_NAME, list: FB_LIST, site: FB_SITE,
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
  fbTabs, fbFilter, fbSearchTab, fbPeople, fbEligible,
  fbApplyTap, fbTapFields, fbOfficeFields, fbFloorOnly, fbLogEntry, fbLogWords, fbRebase, fbCardSig,
  fbCellWord, fbCellWant, fbReportJobs
};
if (typeof window !== "undefined") window.FABC = FABC;
else if (typeof globalThis !== "undefined") globalThis.FABC = FABC;
if (typeof module !== "undefined" && module.exports) module.exports = FABC;
