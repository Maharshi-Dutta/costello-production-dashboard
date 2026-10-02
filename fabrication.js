/* The Fabrication station page - the fabricators' own screen
   (docs/specs/2026-09-25-fabrication-station.md, Part A).

   What this page can do, in full: read four SharePoint lists in the `Floor
   stations` site ("Fabrication station", "Station people", "Station log",
   "Station comments"), PATCH one FramesDone / SashesDone / TransomsDone counter
   of a Fabrication station row with its By/At and the last-touch pair, POST
   one Station log line per counter write that landed, and POST one note per
   word somebody types for the office. Nothing else: no workbook, no delete, no
   job fact, no export.

   Built on welding.js's shape (the queue, the delta poll, the patched board);
   what differs is three parts per group and the ELIGIBILITY gate - everybody
   sees every line, and a person can move only the lines of a product group
   named in their `Stages`. Lines they cannot move are drawn greyed, and a tap
   on one says "not your line" rather than doing nothing.                   */

const $ = s => document.querySelector(s);
const esc = STU.stuEsc;
const F = FABC;

/* this page's own keys: never the glass or welding page's */
const PERSON_KEY = "cw_fabperson";
const QUEUE_KEY = "cw_fabq";
const LOGQ_KEY = "cw_fablogq";
const TAB_KEY = "cw_fabtab";             // "floor" | "finished"
const RETRY_MS = 5000;
const PEOPLE_MS = 600000;
const HINT_MS = 4000;                    // how long "not your line" stays on a card

let ITEMS = [], TOKEN = null, PEOPLE = [];
let READY = false, PEOPLE_READ = false;
let PROBLEM = "", SOFT = "", LASTREAD = 0;
let QUERY = "", TAB = "floor", PRESEARCH = null, TYPED = false;
let PERSON = null, LAST_TAP = 0;
let PINFOR = null, PINTYPED = "", PINBAD = false;
let HINT = {};                           // job -> { group, at }: a tap on somebody else's line
let retryT = null, buildT = null, peopleT = null;
let flushing = false;
const DELTA_OFF_MS = 300000;
let DELTA_OFF = 0;
const deltaOff = () => !!(DELTA_OFF && Date.now() - DELTA_OFF < DELTA_OFF_MS);

try { if (localStorage.getItem(TAB_KEY) === "finished") TAB = "finished"; } catch (e) {}
const saveTab = () => { try { localStorage.setItem(TAB_KEY, TAB); } catch (e) {} };

/* ---- who is on the station ----
   No idle lock (owner, 2026-10-02): the name stays, across reloads and new
   builds, until somebody presses Switch or the person leaves `Station people`.
   Eligibility follows whoever was picked last. */
function loadPerson() {
  let raw = null;
  try { raw = JSON.parse(localStorage.getItem(PERSON_KEY) || "null"); } catch (e) { raw = null; }
  if (!raw || !raw.name) return;
  const hit = PEOPLE.find(p => p.name === raw.name);
  if (!hit) return;
  PERSON = hit; LAST_TAP = Number(raw.at) || Date.now();
  loadView();
}
function savePerson() {
  try {
    if (PERSON) localStorage.setItem(PERSON_KEY, JSON.stringify({ name: PERSON.name, at: LAST_TAP }));
    else localStorage.removeItem(PERSON_KEY);
  } catch (e) {}
}
function touch() { LAST_TAP = Date.now(); savePerson(); }
function clearSearch() {
  QUERY = ""; TYPED = false;
  if (PRESEARCH != null) { TAB = PRESEARCH; PRESEARCH = null; try { window.scrollTo(0, 0); } catch (e) {} }
  const sb = $("#search");
  if (sb) sb.value = "";
}
function pickPerson(p) {
  PERSON = p; PINFOR = null; PINTYPED = ""; PINBAD = false; HINT = {};
  loadView(); clearSearch(); touch(); render();
}
function switchPerson() {
  PERSON = null; PINFOR = null; PINTYPED = ""; PINBAD = false; HINT = {};
  clearSearch(); savePerson(); render();
}
const who = () => (PERSON ? PERSON.name : "");
const mayDo = group => F.fbEligible(PERSON, group);

/* ---- the queues: one entry per row AND part, kept in localStorage and read
   back through the core's whitelist ---- */
let QUEUE = {}, LOGQ = {};
const qKey = (id, part) => String(id) + "|" + String(part);
function cleanQueue(raw) {
  const out = {};
  Object.keys(raw || {}).forEach(k => {
    const e = raw[k];
    if (!e || !e.id || F.FB_TAP_PARTS.indexOf(String(e.part)) < 0) return;
    const v = Number(e.value);
    if (!isFinite(v)) return;
    const from = Number(e.from);
    out[qKey(e.id, e.part)] = { id: String(e.id), part: String(e.part), value: Math.round(v),
                                who: String(e.who || ""), at: String(e.at || ""),
                                job: String(e.job || ""), group: String(e.group || ""),
                                title: String(e.title || ""), site: String(e.site || ""),
                                from: isFinite(from) ? Math.round(from) : 0, err: 0 };
  });
  return out;
}
function cleanLogQ(raw) {
  const out = {};
  Object.keys(raw || {}).forEach(k => {
    const e = raw[k];
    if (!e || !e.fields || !e.fields.Title) return;
    out[String(k)] = { key: String(k), err: 0, fields: ST.logFields({
      job: e.fields.Title, station: F.FB_NAME, type: e.fields.GlassType,
      stage: e.fields.Stage, from: e.fields.From, to: e.fields.To,
      who: e.fields.Who, at: e.fields.At }) };
  });
  return out;
}
try { QUEUE = cleanQueue(JSON.parse(localStorage.getItem(QUEUE_KEY) || "{}")); } catch (e) { QUEUE = {}; }
try { LOGQ = cleanLogQ(JSON.parse(localStorage.getItem(LOGQ_KEY) || "{}")); } catch (e) { LOGQ = {}; }
function saveQueue() {
  try { localStorage.setItem(QUEUE_KEY, JSON.stringify(QUEUE)); } catch (e) {}
  try { localStorage.setItem(LOGQ_KEY, JSON.stringify(LOGQ)); } catch (e) {}
}
function queueTap(rec, part, value) {
  const k = qKey(rec.id, part);
  const had = QUEUE[k];
  QUEUE[k] = { id: String(rec.id), part: part, value: value, who: who(),
               at: new Date().toISOString(), job: rec.job, group: rec.group,
               title: rec.title || (rec.job + "|" + rec.group), site: SITEID || "",
               from: had ? had.from : listValue(rec.id, part), err: 0 };
  saveQueue();
}
function queuedFor(id) {
  const out = {};
  Object.keys(QUEUE).forEach(k => {
    const e = QUEUE[k];
    if (String(e.id) === String(id)) out[F.FB_DONE_FIELD[e.part]] = e.value;
  });
  return Object.keys(out).length ? out : null;
}
const owedFor = id => Object.keys(QUEUE).some(k => String(QUEUE[k].id) === String(id));
const badFor = id => Object.keys(QUEUE).some(k => String(QUEUE[k].id) === String(id) && QUEUE[k].err);
const owedForCard = c => c.groups.some(g => owedFor(g.id));
const badForCard = c => c.groups.some(g => badFor(g.id));
let LOST = {};
const lostFor = job => LOST[String(job).trim().toUpperCase()] || null;
function listValue(id, part) {
  const it = ITEMS.find(x => String(x.id) === String(id));
  if (!it) return 0;
  const v = Number((it.fields || {})[F.FB_DONE_FIELD[part]]);
  return isFinite(v) ? Math.round(v) : 0;
}
/** A queued tap that will not be sent, kept on its card until the next tap:
    "moved" (the office changed the row after it) or "assign" (the line is no
    longer the tapper's). Never dropped in silence. */
function loseTap(e, why) {
  const job = String(e.job).trim().toUpperCase();
  const note = LOST[job] || { job: job, parts: [] };
  note.parts = note.parts.filter(p => !(p.part === e.part && p.group === e.group))
                 .concat([{ part: e.part, group: e.group, value: e.value, why: why, who: e.who }]);
  LOST[job] = note;
  console.warn("[fabrication] dropping a queued tap for " + e.job + " " + e.group + " " + e.part + ": " +
               (why === "assign" ? "that line is no longer assigned to " + e.who : "the row was moved after it was tapped"));
}
function rebaseQueue() {
  let moved = false;
  Object.keys(QUEUE).forEach(k => {
    const e = QUEUE[k];
    const it = ITEMS.find(x => String(x.id) === String(e.id));
    if (!it) return;
    const r = F.fbRebase(e, it.fields || {});
    if (r.action === "drop") {
      loseTap(e, "moved");
      delete QUEUE[k];
      moved = true;
      return;
    }
    if (r.action !== "rebase") return;
    e.value = r.value; e.from = r.from;
    moved = true;
  });
  if (moved) saveQueue();
}
function armRetry() {
  if (retryT) return;
  retryT = setTimeout(() => { retryT = null; flushQueue(); }, RETRY_MS);
}
/** The row a queued tap is about in the list as it is now: its id, or - if the
    lists moved site - the oldest row with its Title; null drops it. */
function currentId(e) {
  if (!e.site || e.site === SITEID) return e.id;
  const want = String(e.title || "").trim().toUpperCase();
  let best = null;
  ITEMS.forEach(it => {
    const t = String((it.fields || {}).Title == null ? "" : it.fields.Title).trim().toUpperCase();
    if (t === want && (!best || Number(it.id) < Number(best.id))) best = it;
  });
  return best ? String(best.id) : null;
}
/** Send everything owed: the counter first, the log line only once it landed. */
async function flushQueue() {
  if (flushing) return;
  if (!Object.keys(QUEUE).length && !Object.keys(LOGQ).length) return;
  if (!READY || !SITEID) { armRetry(); return; }
  flushing = true;
  const tried = {};
  let wrote = false;
  try {
    for (;;) {
      const k = Object.keys(QUEUE).find(x => !tried[x]);
      if (!k) break;
      tried[k] = 1;
      const e = QUEUE[k];
      if (!e) continue;
      /* captured before the await: e IS QUEUE[k], so a tap landing mid-write
         must be compared against what this pass actually sent (welding B19) */
      const want = e.value, when = e.at, whose = e.who, wantFrom = e.from;
      /* THE GATE AGAIN, AT THE MOMENT OF SENDING (review P5). A tap queued
         while the person held the line is not theirs to send once the office
         has taken the assignment away. Unknown gate: hold it for now. */
      if (ASSIGN_OK === null && e.part !== F.FB_GLAZE) continue;   // door glazing has no assignment to wait for
      /* a job that has moved on to a finished section on the sheet is shown
         finished and takes no taps, queued or not */
      /* decided from the list as it is now, not from the drawn records, which
         are only rebuilt while somebody is signed in (review N1) */
      const live = ITEMS.find(x => String(x.id) === String(e.id));
      if ((recordById(e.id) || {}).sheetDone || (live && F.fbSheetDone((live.fields || {}).Section))) {
        loseTap(e, "sheet");
        delete QUEUE[k]; saveQueue();
        continue;
      }
      const tapper = PEOPLE.find(p => p.name === whose) || null;
      if (!F.fbCanTap(tapper, e.group, e.part, IDX, ASSIGN_OK, e.job)) {
        loseTap(e, "assign");
        delete QUEUE[k]; saveQueue();
        continue;
      }
      const body = F.fbFloorOnly(F.fbTapFields(e.part, want, whose, when));
      const id = currentId(e);
      if (!id) { delete QUEUE[k]; saveQueue(); continue; }
      e.id = id; e.site = SITEID;
      try {
        await CW.listPatch(F.FB_LIST, id, body, listOpts());
        wrote = true;
        queueLog({ id: id, part: e.part, job: e.job, group: e.group,
                   value: want, at: when, who: whose, from: wantFrom });
        const now = QUEUE[k];
        if (now && now.value === want && now.at === when) delete QUEUE[k];
        else if (now) { now.err = 0; delete tried[k]; }
        const it = ITEMS.find(x => String(x.id) === String(id));
        if (it) (it.fields = it.fields || {})[F.FB_DONE_FIELD[e.part]] = want;
      } catch (err) {
        if (QUEUE[k]) QUEUE[k].err = 1;
        if (CW.isMissing && CW.isMissing(err) && CW.forgetStationSite) CW.forgetStationSite(false, F.FAB.site);
        console.warn("[fabrication] counter write failed for item " + id + ":", (err && err.message) || err);
      }
      saveQueue();
    }
    await flushLog();
  } finally {
    flushing = false;
  }
  render();
  if (wrote) await pollList();
  if (Object.keys(QUEUE).length || Object.keys(LOGQ).length) armRetry();
}
function queueLog(e) {
  const from = Number(e.from) || 0;
  if (from === e.value) return;
  const key = e.id + "|" + e.part + "|" + e.at;
  LOGQ[key] = { key: key, err: 0, fields: ST.logFields(F.fbLogEntry({
    job: e.job, group: e.group, part: e.part, from: from, to: e.value, who: e.who, at: e.at })) };
  saveQueue();
}
async function flushLog() {
  if (!SITEID) return;
  const keys = Object.keys(LOGQ);
  for (let i = 0; i < keys.length; i++) {
    const e = LOGQ[keys[i]];
    if (!e) continue;
    try { await CW.listAdd(ST.LOG_LIST, e.fields, logOpts()); delete LOGQ[keys[i]]; }
    catch (err) { e.err = 1; console.warn("[fabrication] log line not written yet:", (err && err.message) || err); }
    saveQueue();
  }
}

/* ---- the lists ---- */
let SITEID = null;
const listOpts = () => ({ siteId: SITEID, fields: F.FB_FIELDS });
const peopleOpts = () => ({ siteId: SITEID, fields: ST.PEOPLE_FIELDS });
const logOpts = () => ({ siteId: SITEID, fields: ST.LOG_FIELDS });
const commentOpts = () => ({ siteId: SITEID, fields: ST.COMMENT_FIELDS });
const NOTES = ST.stationComments({
  station: F.FB_NAME,
  listItems: (name, o) => CW.listItems(name, o),
  listAdd: (name, fields, o) => CW.listAdd(name, fields, o),
  opts: commentOpts
});

function trouble(e) {
  const m = (e && e.message) || String(e);
  console.warn("[fabrication] could not read SharePoint:", m);
  if (CW.isMissing && CW.isMissing(e)) {
    SITEID = null;
    if (CW.forgetStationSite) CW.forgetStationSite(false, F.FAB.site);
  }
  if (/interaction_required|login_required/.test(m)) { PROBLEM = "reauth"; SOFT = ""; READY = false; }
  else if (/permission needed/.test(m)) { PROBLEM = "consent"; SOFT = ""; READY = false; }
  else if (CW.isMissing && CW.isMissing(e)) { PROBLEM = "site"; SOFT = ""; READY = false; }
  else SOFT = "cannot reach SharePoint — retrying";
}
async function resolveSite() {
  if (!SITEID) SITEID = await CW.stationSite(F.FAB.site);
  return SITEID;
}
async function readPeople() {
  try {
    if (!(await resolveSite())) { PROBLEM = "site"; SOFT = ""; READY = false; render(); return false; }
    const items = await CW.listItems(ST.PEOPLE_LIST, peopleOpts());
    if (items == null) { PROBLEM = "people"; SOFT = ""; READY = false; render(); return false; }
    PEOPLE = F.fbPeople(items);
    PEOPLE_READ = true;
    if (PROBLEM === "people") PROBLEM = "";
    if (PERSON) {
      PERSON = PEOPLE.find(p => p.name === PERSON.name) || null;
      if (!PERSON) savePerson();
    } else loadPerson();
    render();
    return true;
  } catch (e) { trouble(e); render(); return false; }
}
async function readList() {
  try {
    if (!(await resolveSite())) { PROBLEM = "site"; SOFT = ""; READY = false; render(); return false; }
    if (siteMoved()) { TOKEN = null; DELTA_OFF = 0; }
    let items = null;
    if (deltaOff()) {
      items = await CW.listItems(F.FB_LIST, listOpts());
      if (items == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
      TOKEN = null;
    } else {
      try {
        const d = await CW.listDelta(F.FB_LIST, listOpts());
        if (d == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
        items = d.items.filter(x => !x.removed).map(x => ({ id: x.id, fields: x.fields }));
        TOKEN = d.next; DELTA_OFF = 0;
      } catch (e) {
        if (!CW.isDeltaRestart || !CW.isDeltaRestart(e)) throw e;
        if (!(CW.isDeltaResync && CW.isDeltaResync(e))) DELTA_OFF = Date.now();
        items = await CW.listItems(F.FB_LIST, listOpts());
        if (items == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
        TOKEN = null;
      }
    }
    ITEMS = items; READY = true; PROBLEM = ""; SOFT = ""; LASTREAD = Date.now();
    rebaseQueue(); render(); flushQueue();
    return true;
  } catch (e) { trouble(e); render(); return false; }
}
let SITE_GEN = 0;
function siteMoved() {
  const m = CW.stationSiteMoves ? CW.stationSiteMoves(F.FAB.site) : 0;
  if (m === SITE_GEN) return false;
  SITE_GEN = m;
  return true;
}
async function pollList() {
  try { SITEID = (await CW.stationSite(F.FAB.site)) || SITEID; } catch (e) {}
  if (siteMoved()) { TOKEN = null; DELTA_OFF = 0; }
  if (!TOKEN) return readList();
  try {
    const d = await CW.listDelta(F.FB_LIST, { siteId: SITEID, fields: F.FB_FIELDS, token: TOKEN });
    if (d == null) return readList();
    ITEMS = ST.mergeDelta(ITEMS, d.items);
    if (d.next) TOKEN = d.next;
    READY = true; SOFT = ""; LASTREAD = Date.now();
    if (d.items.length) TICK.burst();     // something moved: look again sooner for a while
    rebaseQueue(); render(); flushQueue();
    return true;
  } catch (e) {
    if (CW.isDeltaRestart && CW.isDeltaRestart(e)) {
      if (!(CW.isDeltaResync && CW.isDeltaResync(e))) DELTA_OFF = Date.now();
      TOKEN = null;
      return readList();
    }
    trouble(e); render(); return false;
  }
}

/* ---- taps ---- */
function tap(id, part, delta) {
  if (!PERSON) return;
  const rec = recordById(id);
  if (!rec) return;
  touch();
  /* THE GATE. The button is drawn greyed, and this is the rule: a line of a
     group this person does not do - or, once assignments exist, a part they do
     not hold an Assigned row on - is never queued. It says why. */
  const why = rec.sheetDone ? "sheet" : !mayDo(rec.group) ? "group"
    : !F.fbEligible(PERSON, rec.group, part) ? "part"
    : part !== F.FB_GLAZE && ASSIGN_OK === null ? "checking"
    : !mayTap(rec, part) ? "assign" : !mayAct(rec, part, delta) ? "whole" : "";
  if (why) {
    HINT[rec.job] = { group: rec.group, at: Date.now(), why: why };
    render();
    setTimeout(render, HINT_MS + 50);
    return;
  }
  const value = F.fbApplyTap(rec, part, delta);
  if (value == null || value === rec[part]) return;
  delete LOST[String(rec.job).trim().toUpperCase()];
  queueTap(rec, part, value);
  TICK.burst();
  render();
  flushQueue();
}
function itemsNow() {
  return ITEMS.map(it => {
    const q = queuedFor(String(it.id));
    return q ? { id: it.id, fields: Object.assign({}, it.fields, q) } : it;
  });
}
let RECS_BY_ID = {};
function boardNow() {
  const tabs = F.fbTabs(itemsNow());
  /* urgent first on every worker's list (Part B) */
  tabs.floor = F.fbUrgentFirst(tabs.floor);
  tabs.finished = F.fbUrgentFirst(tabs.finished);
  RECS_BY_ID = {};
  tabs.floor.concat(tabs.finished).forEach(c => c.groups.forEach(g => { RECS_BY_ID[String(g.id)] = g; }));
  /* the view filter (2026-10-01) narrows what is DRAWN, after the records the
     tap gate reads were taken from the whole board above */
  /* ... and places each card by the lines it shows (review M2) */
  const view = F.fbViewTabs(tabs, PERSON, VIEW, IDX);
  /* urgent stays first; then, when asked, the jobs whose glass is ready */
  view.floor = F.fbGlassFirst(view.floor, GLASSFIRST);
  view.finished = F.fbGlassFirst(view.finished, GLASSFIRST);
  return { tabs: view, all: tabs };
}

/* ---- the view: Everything | My work | Assigned to me (2026-10-01) ----------
   Display only. Remembered per person in cw_fabview; Everything by default. */
const VIEW_KEY = "cw_fabview";
let VIEW = "all";
/* "Glass ready first" (2026-10-01): remembered per TABLET, not per person */
const GLASSFIRST_KEY = "cw_fabglassfirst";
let GLASSFIRST = false;
try { GLASSFIRST = localStorage.getItem(GLASSFIRST_KEY) === "1"; } catch (e) {}
function setGlassFirst(on) {
  GLASSFIRST = !!on;
  try { localStorage.setItem(GLASSFIRST_KEY, GLASSFIRST ? "1" : "0"); } catch (e) {}
}
function viewStore() {
  try { return JSON.parse(localStorage.getItem(VIEW_KEY) || "{}") || {}; } catch (e) { return {}; }
}
function loadView() { VIEW = PERSON ? F.fbViewFor(viewStore(), PERSON.name) : "all"; }
function setView(v) {
  VIEW = F.fbViewOf(v);
  if (!PERSON) return;
  try { const s = viewStore(); s[PERSON.name] = VIEW; localStorage.setItem(VIEW_KEY, JSON.stringify(s)); } catch (e) {}
}

/* ---- Part B: the assignments list ------------------------------------------
   Read every tick with the board - since 2026-10-02 by delta, through the
   shared read-only reader (STU.stuListReader), not the whole list each time.
   ASSIGN_OK: null not read yet, true read, false the list is not there - and
   then this page behaves exactly as Part A (eligible = may tap) and says so in
   the footer. ONLY a list that is positively not there is `false`: a refusal,
   a blip, or a first read that has not answered leaves the last state, which
   for a page that has read nothing yet is null - the gate stays shut
   (HISTORY B32). The page's ONLY write to this list is a Requested row for
   the person signed in (Take), and it does not go through the reader. */
let ASSIGN_ROWS = [], IDX = {}, ASSIGN_OK = null;
const assignOpts = () => ({ siteId: SITEID, fields: F.FB_ASSIGN_FIELDS });
/* retried on the next turn after a failure: this list gates the taps */
const ASG = STU.stuListReader({ site: F.FAB.site, list: F.FB_ASSIGN_LIST, fields: F.FB_ASSIGN_FIELDS,
                                retryMs: ST.REFRESH_MS - 500, tag: "[fabrication]" });
/** Answers whether the assignments moved. */
async function readAssign() {
  const moved = await ASG.read();
  if (ASG.items) {
    if (moved || ASSIGN_OK !== true) { ASSIGN_ROWS = F.fbAssignRows(ASG.items); IDX = F.fbAssignIndex(ASSIGN_ROWS); }
    ASSIGN_OK = true;
  } else if (ASG.missing) { ASSIGN_OK = false; ASSIGN_ROWS = []; IDX = {}; }
  /* anything else: keep the last read, the gate does not flap on a blip */
  return moved;
}
const listOn = () => ASSIGN_OK === true;
/* three states, failing closed: see fbCanTap (review P1) */
const mayTap = (rec, part) => !rec.sheetDone && F.fbCanTap(PERSON, rec.group, part, IDX, ASSIGN_OK, rec.job);
const mayAct = (rec, part, act) =>
  F.fbActAllowed(PERSON, IDX, ASSIGN_OK, rec.job, rec.group, part, rec[part + "Total"], act);
/* TAKING[k]: 1 while the request is being sent; then the time it was sent, and
   for TAKE_HOLD_MS a second Take on that line is refused - a delta read can
   run a turn behind the write, and the row just added must not be asked for
   twice in the gap before it shows as "requested". */
const TAKE_HOLD_MS = 15000;
let TAKING = {};
async function take(id, part) {
  const rec = recordById(id);
  if (!rec || !PERSON || !listOn() || rec.sheetDone || !F.fbEligible(PERSON, rec.group, part)) return;
  touch();
  const k = rec.job + "|" + rec.group + "|" + part;
  if (TAKING[k] > 1 && Date.now() - TAKING[k] >= TAKE_HOLD_MS) delete TAKING[k];
  if (TAKING[k] || F.fbRequested(IDX, PERSON, rec.job, rec.group, part)) return;
  const free = Math.max(0, rec[part + "Total"] - F.fbAssignedSum(IDX, rec.job, rec.group, part));
  const body = F.fbRequestFields(rec.job, rec.group, part, PERSON.name, free, new Date().toISOString());
  if (!body || !(free > 0)) return;
  TAKING[k] = 1; render();
  try {
    await CW.listAdd(F.FB_ASSIGN_LIST, body, assignOpts());
    TAKING[k] = Date.now(); TICK.burst();
    await readAssign();
  } catch (e) {
    delete TAKING[k];
    SOFT = "that request could not be sent — try again"; console.warn("[fabrication] take failed:", (e && e.message) || e);
  }
  render();
}

/* ---- Part B: notices (banner, beep, badge) ---------------------------------- */
const SEEN_KEY = "cw_fabseen";
let SEEN = {};
try { SEEN = JSON.parse(localStorage.getItem(SEEN_KEY) || "{}") || {}; } catch (e) { SEEN = {}; }
const saveSeen = () => { try { localStorage.setItem(SEEN_KEY, JSON.stringify(SEEN)); } catch (e) {} };
let UNSEEN = [], BEEPED = {}, AUDIO = null;
function beep() {
  /* a short tone, no audio file. A browser may refuse sound until somebody has
     touched the page; that is not an error worth anything but silence. */
  try {
    const AC = window.AudioContext || window.webkitAudioContext;
    if (!AC) return;
    AUDIO = AUDIO || new AC();
    if (AUDIO.state === "suspended" && AUDIO.resume) AUDIO.resume().catch(() => {});
    const o = AUDIO.createOscillator(), g = AUDIO.createGain();
    o.frequency.value = 880; g.gain.value = 0.15;
    o.connect(g); g.connect(AUDIO.destination);
    o.start(); o.stop(AUDIO.currentTime + 0.18);
  } catch (e) {}
}
function noticesNow(cards) {
  if (!PERSON || !listOn()) { UNSEEN = []; return; }
  const all = F.fbNotices(ASSIGN_ROWS, cards, PERSON, IDX);
  const fresh = !SEEN[PERSON.name];
  const seen = F.fbSeenFor(SEEN, PERSON.name, all);
  /* saved only when it changed, and pruned to what is still current (P7) */
  if (F.fbPruneSeen(seen, all) || fresh) saveSeen();
  UNSEEN = F.fbUnseen(all, seen);
  if (UNSEEN.some(n => !BEEPED[n.key])) beep();
  UNSEEN.forEach(n => { BEEPED[n.key] = 1; });
}
function markSeen(job) {
  if (!PERSON || !UNSEEN.length) return;
  const seen = SEEN[PERSON.name] || (SEEN[PERSON.name] = {});
  let moved = false;
  UNSEEN.forEach(n => { if (!job || n.job === job) { seen[n.key] = 1; moved = true; } });
  if (moved) { saveSeen(); render(); }
}
const badgeFor = job => UNSEEN.filter(n => n.job === job).length;
const recordById = id => RECS_BY_ID[String(id)] || null;
const hintFor = job => {
  const h = HINT[job];
  return h && Date.now() - h.at < HINT_MS ? h : null;
};

/* ---- drawing ---- */
function barHtml(done, total) {
  const pct = total > 0 ? Math.max(0, Math.min(100, Math.round((done / total) * 100))) : 0;
  return '<span class="wbar" aria-hidden="true"><span class="wbarfill" style="width:' + pct + '%"></span></span>';
}
/** One part of one group. A line the person may not move is drawn greyed, NOT
    disabled: a disabled button swallows the tap, and the tap is what says
    "not your line". */
const URG = '<span class="urg" title="urgent" aria-label="urgent">!</span>';
/** What the label says under a line's name, and whether it offers Take. */
function lineState(rec, part) {
  if (rec.sheetDone) return { can: false, words: "finished on sheet" };
  if (!mayDo(rec.group)) return { can: false, words: "not your line" };
  if (!F.fbEligible(PERSON, rec.group, part)) return { can: false, words: "not your part" };
  if (part === F.FB_GLAZE) return { can: true, words: "" };          // a role, no assignment, no Take
  if (ASSIGN_OK === false) return { can: true, words: "" };           // no list: Part A
  if (ASSIGN_OK !== true) return { can: false, words: "checking assignments…" };
  const mine = F.fbMine(IDX, PERSON, rec.job, rec.group, part);
  if (mine > 0) return { can: true, words: "yours: " + mine };
  if (F.fbRequested(IDX, PERSON, rec.job, rec.group, part)) return { can: false, words: "requested — waiting" };
  const free = rec[part + "Total"] - F.fbAssignedSum(IDX, rec.job, rec.group, part);
  return { can: false, words: free > 0 ? "" : "assigned to others", take: free > 0 };
}
function stepHtml(rec, line) {
  const st = lineState(rec, line.part), mine = st.can;
  const lock = mine ? "" : ' aria-disabled="true"';
  const b = (t, act, cls, dead) => '<button class="' + cls + (dead ? " dead" : "") + '" data-id="' + esc(rec.id) +
    '" data-part="' + esc(line.part) + '" data-act="' + esc(act) + '"' +
    (dead && !lock ? ' aria-disabled="true"' : lock) + '>' + t + '</button>';
  const full = line.done >= line.total;
  /* All / None set the whole line: only for somebody holding all of it (P2) */
  const wholeOk = mayAct(rec, line.part, "all");
  const urgent = (rec.urgentOf || {})[line.part];
  return '<div class="step' + (mine ? "" : " locked") + ' c-' + (line.colour || "none") + '">' +
    '<span class="stepl">' + (urgent ? URG : "") + esc(line.label) +
      (mine ? '<span class="stepleft tab">' + esc(Math.max(0, line.total - line.done) + " left") + '</span>' : "") +
      (st.words ? '<span class="nomine' + (mine ? " yours" : "") + '">' + esc(st.words) + '</span>' : "") +
      (st.take ? '<button class="take" data-take="' + esc(rec.id) + '" data-part="' + esc(line.part) + '">Take</button>' : "") +
    '</span>' +
    '<span class="stepmid"><span class="stepn tab' + (full ? " full" : "") + '">' + line.done + ' / ' +
      line.total + '</span>' + barHtml(line.done, line.total) + '</span>' +
    '<span class="stepc">' + b("&minus;", "-1", "sbtn") + b("+", "1", "sbtn") +
      b(full ? "None" : "All", full ? "none" : "all", "sall", !wholeOk) + '</span>' +
  '</div>';
}
function groupHtml(rec) {
  rec.urgentOf = F.fbUrgentOf(rec.urgent);
  return '<div class="wgrp c-' + (rec.colour || "none") + '">' +
    '<div class="wghead"><span class="wgname cond">' + (rec.urgentOf.group ? URG : "") + esc(rec.group) + '</span>' +
      (rec.doors ? '<span class="wgdoors tab">' + esc(rec.doors) + '</span>' : "") +
      '<span class="wgcount tab">' + rec.done + ' / ' + rec.total + '</span></div>' +
    /* door glazing, when the group has it, is the last line: its own count */
    '<div class="steps">' + rec.lines.concat(rec.extra || []).map(l => stepHtml(rec, l)).join("") + '</div>' +
  '</div>';
}
function cardInner(c) {
  const owed = owedForCard(c), bad = badForCard(c), lost = lostFor(c.job), hint = hintFor(c.job);
  const jobUrgent = c.groups.some(g => F.fbUrgentOf(g.urgent).job);
  const badge = badgeFor(c.job);
  const glass = F.fbGlassChip(c.glass);
  return '<div class="chead">' +
      (jobUrgent ? URG : "") +
      '<span class="cond job">' + esc(c.job) + '</span>' +
      (badge ? '<span class="nbadge" title="new for you">' + badge + '</span>' : "") +
      '<span class="cust">' + esc(c.customer || "—") + '</span>' +
      (c.section && !c.active ? '<span class="csec">' + esc(c.section) + '</span>' : "") +
      (c.sheetDone ? '<span class="csheet">finished on sheet</span>' : "") +
      /* the job's glass, as the office fed it (read only; nothing when blank) */
      (glass.kind ? '<span class="gchip g-' + glass.kind + '">' + esc(glass.words) + '</span>' : "") +
    '</div>' +
    (c.comment ? '<div class="cfacts"><span class="ccmt">“' + esc(c.comment) + '”</span></div>' : "") +
    (c.glassStart ? '<div class="gstart">Glass is ready and nothing is fabricated yet — start this job</div>' : "") +
    '<div class="wgroups">' + c.groups.map(groupHtml).join("") + '</div>' +
    (hint ? '<div class="hint">' + (hint.why === "assign"
      ? "Not assigned to you yet — tap Take on the line, and the office approves it."
      : hint.why === "checking" ? "Checking assignments… try again in a moment."
      : hint.why === "part" ? "Not your part — you do other parts of " + esc(hint.group) + ", not this one."
      : hint.why === "sheet" ? "Finished on the sheet — this job has moved on, nothing to record here."
      : hint.why === "whole" ? "All/None only when the whole line is yours — use − and +."
      : "Not your line — " + esc(hint.group) + " is not one of your product groups.") + '</div>' : "") +
    (lost ? lost.parts.map(p => '<div class="unsaved">' + esc(p.group + " " + p.part + " " + p.value) +
        (p.why === "assign" ? " was not saved — that line is no longer assigned to " + esc(p.who || "you")
         : p.why === "sheet" ? " was not saved — the job is finished on the sheet"
                             : " was not saved — the office changed this job after that tap") + '</div>').join("") : "") +
    (bad ? '<div class="unsaved">not saved yet — retrying</div>'
         : owed ? '<div class="saving">saving…</div>' : "") +
    NOTES.html(c.job);
}

/* ---- the picker ---- */
function pickerHtml() {
  return STU.stuPickerHtml({
    people: PEOPLE, pinFor: PINFOR, pinTyped: PINTYPED, pinBad: PINBAD,
    labelOf: k => String(k).toUpperCase(),
    empty: "Nobody is set up for the Fabrication station yet. Ask the office to add " +
           "you to the “Station people” list."
  });
}
function wirePicker(host) {
  STU.stuWirePicker(host, {
    onPerson: name => {
      const p = PEOPLE.find(x => x.name === name);
      if (!p) return;
      if (ST.pinOk(p, "")) { pickPerson(p); return; }
      PINFOR = p; PINTYPED = ""; PINBAD = false; render();
    },
    onKey: k => {
      if (k === "cancel") { PINFOR = null; PINTYPED = ""; PINBAD = false; render(); return; }
      if (k === "back") { PINTYPED = PINTYPED.slice(0, -1); PINBAD = false; render(); return; }
      if (k === "ok") { submitPin(); return; }
      if (PINTYPED.length < 8) PINTYPED += k;
      PINBAD = false;
      if (PINFOR && PINTYPED.length >= String(PINFOR.pin).length) submitPin();
      else render();
    }
  });
}
function submitPin() {
  if (!PINFOR) return;
  if (ST.pinOk(PINFOR, PINTYPED)) { pickPerson(PINFOR); return; }
  PINTYPED = ""; PINBAD = true; render();
}

/* ---- the board, drawn once and then patched (welding.js's shape) ---- */
let NODES = {}, LIST = null, LIST_TAB = "", BOARD_PREV = null, QSIG = {}, PSIG = "", WIRED = false;
function qState(c) {
  const lost = lostFor(c.job), hint = hintFor(c.job);
  return (owedForCard(c) ? "o" : "") + (badForCard(c) ? "b" : "") +
         (lost ? "!" + lost.parts.map(p => p.group + p.part + p.value).join(",") : "") +
         (hint ? "h" + hint.group + hint.why : "") + "/" + NOTES.sig(c.job) +
         /* the assignment state of every line, and the notice badge: neither is
            in the Fabrication station list, so boardDiff cannot see them */
         "/" + c.groups.map(g => F.FB_TAP_PARTS.map(p => JSON.stringify(lineState(g, p))).join("")).join(";") +
         "/" + badgeFor(c.job) + (Object.keys(TAKING).length ? "t" : "");
}
/* the view is in here too: a switch redraws every card, since the card
   signature does not know which lines the view drew */
const pState = () => (PERSON ? PERSON.name + "|" + (PERSON.stages || []).join(",") + "|" + VIEW +
  (VIEW === "assigned" ? "|" + JSON.stringify(IDX) : "") : "");
function makeCard(c) {
  const el = document.createElement("div");
  el.className = "card c-" + (c.colour || "none") + (c.finished ? " done" : "");
  if (el.dataset) el.dataset.job = c.job;
  el.setAttribute("data-job", c.job);
  el.innerHTML = cardInner(c);
  return el;
}
function dressCard(el, c) {
  const act = document.activeElement;
  const at = act && act.dataset && act.dataset.cmbox === c.job ? act.selectionStart : null;
  el.className = "card c-" + (c.colour || "none") + (c.finished ? " done" : "");
  el.innerHTML = cardInner(c);
  if (at == null || !el.querySelector) return;
  const box = el.querySelector('[data-cmbox="' + c.job + '"]');
  if (!box) return;
  const to = Math.min(Number(at) || 0, String(box.value || "").length);
  try { box.focus(); if (box.setSelectionRange) box.setSelectionRange(to, to); } catch (e) {}
}
function paintBoard(host, board) {
  if (!LIST || LIST_TAB !== TAB) {
    host.innerHTML = "";
    LIST = document.createElement("div"); LIST.className = "grp";
    LIST_TAB = TAB;
    host.appendChild(LIST);
    NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
  }
  const diff = ST.boardDiff(BOARD_PREV, board, F.fbCardSig);
  const byJob = {};
  board.forEach(c => { byJob[c.job] = c; });
  diff.removed.forEach(j => {
    const el = NODES[j];
    if (el && el.remove) el.remove();
    delete NODES[j]; delete QSIG[j];
  });
  const changed = diff.changed.slice();
  const psig = pState(), pmoved = PSIG !== psig;
  PSIG = psig;
  board.forEach(c => {
    const sig = qState(c);
    if ((pmoved || QSIG[c.job] !== sig) && changed.indexOf(c.job) < 0 && diff.added.indexOf(c.job) < 0)
      changed.push(c.job);
    QSIG[c.job] = sig;
  });
  changed.forEach(j => { if (NODES[j]) dressCard(NODES[j], byJob[j]); });
  diff.added.forEach(j => { NODES[j] = makeCard(byJob[j]); });
  if (diff.order || diff.added.length || diff.removed.length) {
    const act = document.activeElement;
    const box = act && act.dataset && act.dataset.cmbox ? act : null;
    const at = box ? box.selectionStart : null;
    board.forEach(c => { if (NODES[c.job]) LIST.appendChild(NODES[c.job]); });
    if (box && document.activeElement !== box) {
      try { box.focus(); if (box.setSelectionRange && at != null) box.setSelectionRange(at, at); } catch (e) {}
    }
  }
  BOARD_PREV = board;
}

function render() {
  const host = $("#board");
  if (!host) return;
  const hdr = $("#whois");
  if (hdr) hdr.textContent = PERSON ? PERSON.name + ((PERSON.stages || []).length
    ? " · " + PERSON.stages.length + " group" + (PERSON.stages.length === 1 ? "" : "s") : " · no groups") : "";
  const sw = $("#switchbtn");
  if (sw) { sw.hidden = !PERSON; sw.style.display = PERSON ? "" : "none"; }
  const boarding = !PROBLEM && PEOPLE_READ && !!PERSON && READY;
  const sb = $("#search");
  if (sb) { sb.hidden = !boarding; sb.style.display = boarding ? "" : "none"; }
  const now = boarding ? boardNow() : null;
  /* the notices, before any card is drawn: the badge is part of the card */
  /* ... over the WHOLE board: notifications are never filtered by the view */
  if (now) noticesNow(now.all.floor.concat(now.all.finished)); else UNSEEN = [];
  const vw = $("#wview");
  if (vw) { vw.hidden = !boarding; vw.style.display = boarding ? "" : "none"; }
  const gf = $("#glassfirst");
  if (gf) {
    gf.hidden = !boarding; gf.style.display = boarding ? "" : "none";
    gf.className = GLASSFIRST ? "on" : "";
    gf.setAttribute("aria-pressed", GLASSFIRST ? "true" : "false");
  }
  [["#vall", "all"], ["#vmine", "mine"], ["#vasg", "assigned"]].forEach(([s, v]) => {
    const b = $(s);
    if (!b) return;
    b.className = VIEW === v ? "on" : "";
    b.setAttribute("aria-pressed", VIEW === v ? "true" : "false");
  });
  const ban = $("#banner");
  if (ban) {
    const on = !!UNSEEN.length;
    ban.hidden = !on; ban.style.display = on ? "" : "none";
    ban.textContent = on ? (UNSEEN.length > 1 ? UNSEEN.length + " new: " : "New: ") + UNSEEN[0].text +
      (UNSEEN.length > 1 ? " …" : "") + " — tap to dismiss" : "";
  }
  const afoot = $("#afoot");
  if (afoot) {
    const off = boarding && ASSIGN_OK === false;
    afoot.hidden = !off; afoot.style.display = off ? "" : "none";
    afoot.textContent = off ? "Assignments are not set up yet (the “Fabrication assignments” list is " +
      "not in the floor’s site), so every line of your own product groups can be tapped." : "";
  }
  let more = 0;
  if (now && QUERY.trim()) {
    if (TYPED) {
      const s = F.fbSearchTab(now.tabs, QUERY, TAB);
      if (s.tab !== TAB) { TAB = s.tab; try { window.scrollTo(0, 0); } catch (e) {} }
    }
    more = F.fbFilter(now.tabs[TAB === "floor" ? "finished" : "floor"], QUERY).length;
  }
  if (now) TYPED = false;
  const mb = $("#more");
  if (mb) {
    mb.hidden = !more; mb.style.display = more ? "" : "none";
    mb.textContent = more ? more + " more " + (TAB === "floor" ? "in Finished" : "on floor") + " ›" : "";
  }
  const tabs = $("#wtabs");
  if (tabs) { tabs.hidden = !boarding; tabs.style.display = boarding ? "" : "none"; }
  [["#tabfloor", "floor", "On floor"], ["#tabfin", "finished", "Finished"]].forEach(([s, t, label]) => {
    const b = $(s);
    if (!b) return;
    b.className = TAB === t ? "on" : "";
    b.setAttribute("aria-selected", TAB === t ? "true" : "false");
    b.textContent = label + (now ? " · " + now.tabs[t].length : "");
  });
  const upd = $("#upd");
  if (upd) upd.textContent = LASTREAD ? "updated " + STU.stuAgo(LASTREAD) : "";
  const soft = $("#soft");
  if (soft) { soft.textContent = SOFT; soft.hidden = !SOFT; soft.style.display = SOFT ? "" : "none"; }

  if (!boarding) {
    LIST = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    if (!PROBLEM && PEOPLE_READ && !PERSON) { host.innerHTML = pickerHtml(); wirePicker(host); return; }
    host.innerHTML = '<div class="msg">' + esc(words()) + againHtml() + '</div>';
    wireAgain();
    return;
  }
  const board = F.fbFilter(now.tabs[TAB], QUERY);
  if (!board.length) {
    LIST = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    /* Assigned to me with the list not read (or not there): say so, rather
       than an empty board that looks like "nothing assigned" */
    const notLoaded = VIEW === "assigned" && ASSIGN_OK !== true;
    /* decided from the CURRENT tab (review M1): the whole board's tab empty is
       the plain wording; the person's work in the other tab is offered as a
       tap; otherwise there is nothing of theirs at all */
    const other = TAB === "floor" ? "finished" : "floor";
    const otherN = now.tabs[other].length;
    host.innerHTML = '<div class="msg">' +
      (notLoaded ? (ASSIGN_OK === false ? "Assignments are not set up yet — switch to Everything or My work."
                                        : "Assignments not loaded yet — try again in a moment.")
       : QUERY.trim() ? "No job under " + (TAB === "floor" ? "On floor" : "Finished") + " matches “" + esc(QUERY) + "”."
       : VIEW === "all" || !now.all[TAB].length
         ? (TAB === "floor" ? "Nothing on the floor right now." : "No finished jobs yet.")
       : otherN ? '<button class="again" id="gotab">Nothing for you here — ' + otherN + " for you " +
           (other === "finished" ? "in Finished" : "on the floor") + " ›</button>"
       : "Nothing for you here — switch to Everything to see all jobs.") + '</div>';
    const gt = host.querySelector ? host.querySelector("#gotab") : null;
    if (gt) gt.onclick = () => switchTab(other);
    return;
  }
  paintBoard(host, board);
  wireBoard(host);
}
/** Go to a tab, as the tab buttons do (a search's tab is not remembered). */
function switchTab(t) {
  if (TAB !== t) {
    TAB = t;
    if (PRESEARCH == null) saveTab();
    try { window.scrollTo(0, 0); } catch (e) {}
  }
  touch(); render();
}
function words() {
  return PROBLEM === "reauth" ? "The sign-in has expired. Tap Sign out, then Sign in again."
    : PROBLEM === "consent" ? "Ask the office to grant the SharePoint permission."
    : PROBLEM === "people" ? "The “Station people” list is not in the floor’s site yet. Ask the office to add it."
    : PROBLEM === "list" ? "The “Fabrication station” list is not in the floor’s site yet. Ask the office to add it — " +
                           "nothing in the Excel file is involved."
    : PROBLEM === "site" ? "The floor’s SharePoint site is not there yet, or this account cannot see it. Ask the office."
    : "Reading the board…";
}
const againHtml = () => PROBLEM === "reauth"
  ? '<div><button class="again" id="reauth">Sign in again</button></div>'
  : '<div><button class="again" id="again">Try again</button></div>';
function wireAgain() {
  const a = $("#again");
  if (a) a.onclick = () => {
    PROBLEM = ""; SOFT = ""; SITEID = null; TOKEN = null;
    if (CW.forgetStationSite) CW.forgetStationSite(true, F.FAB.site);
    render(); readPeople(); readList();
  };
  const r = $("#reauth");
  if (r) r.onclick = async () => {
    try { await CW.signIn(CW.LIST_SCOPES); PROBLEM = ""; SITEID = null; TOKEN = null; await start(); }
    catch (e) { console.warn("[fabrication] sign-in again failed:", (e && e.message) || e); }
  };
}
function wireBoard(host) {
  if (WIRED || !host.addEventListener) return;
  WIRED = true;
  host.addEventListener("click", ev => {
    touch();
    /* a tap anywhere on a card clears that card's notice badge */
    const card = ev.target && ev.target.closest ? ev.target.closest(".card") : null;
    const cj = card && card.getAttribute ? card.getAttribute("data-job") : "";
    if (cj && badgeFor(cj)) markSeen(cj);
    let el = ev.target;
    for (let i = 0; el && i < 5; i++) {
      const d = el.dataset || {};
      if (d.take) {
        ev.stopPropagation();
        take(d.take, d.part);
        return;
      }
      if (d.act) {
        ev.stopPropagation();
        tap(d.id, d.part, d.act === "all" || d.act === "none" ? d.act : Number(d.act));
        return;
      }
      if (d.cmt) {
        ev.stopPropagation();
        NOTES.toggle(d.cmt);
        render();
        if (NOTES.isOpen(d.cmt) && NOTES.state.ok !== true) NOTES.read().then(() => render(), () => {});
        return;
      }
      if (d.cmsend) {
        ev.stopPropagation();
        if (el.disabled) return;
        sendNote(d.cmsend);
        return;
      }
      el = el.parentElement;
    }
  });
  host.addEventListener("input", ev => {
    const d = (ev.target && ev.target.dataset) || {};
    if (!d.cmbox) return;
    touch();
    NOTES.setDraft(d.cmbox, ev.target.value);
  });
}
async function sendNote(job) {
  if (!PERSON || !NOTES.draftOf(job).trim()) return;
  const going = NOTES.send(job, who());
  render();
  await going;
  render();
}

/* ---- staying current ---- */
const BUILD_MS = 120000;
let BUILD_NOW = null;
async function checkBuild() {
  try {
    const r = await fetch("version.json?t=" + Date.now(), { cache: "no-store" });
    if (!r.ok) return;
    const latest = (await r.json()).build;
    if (!latest) return;
    if (!BUILD_NOW) { BUILD_NOW = latest; return; }
    if (latest !== BUILD_NOW && !Object.keys(QUEUE).length && !Object.keys(LOGQ).length) location.reload(true);
  } catch (e) {}
}
/** One turn of the clock (STU.stuTicker: every five seconds, every two for a
    while after a tap or a change, never two turns at once). */
async function tickOnce() {
  render();
  if (!PEOPLE_READ) await readPeople();
  if (await readAssign()) TICK.burst();
  await pollList();
  if (await NOTES.poll()) render();
}
const TICK = STU.stuTicker(tickOnce);

/* ---- the page ---- */
async function start() {
  STU.stuGate(false);
  STU.stuApplyTheme(STU.stuThemeNow());
  const tb = $("#themebtn");
  if (tb) tb.onclick = () => STU.stuApplyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  $("#outbtn").onclick = () => { if (confirm("Sign out of the Fabrication station?")) CW.signOut(); };
  const sw = $("#switchbtn");
  if (sw) sw.onclick = () => switchPerson();
  const sb = $("#search");
  if (sb) sb.oninput = () => {
    QUERY = sb.value || "";
    TYPED = true;
    if (QUERY.trim()) { if (PRESEARCH == null) PRESEARCH = TAB; }
    else clearSearch();
    touch(); render();
  };
  const goTab = switchTab;
  [["#tabfloor", "floor"], ["#tabfin", "finished"]].forEach(([s, t]) => {
    const b = $(s);
    if (b) b.onclick = () => goTab(t);
  });
  const mb = $("#more");
  if (mb) mb.onclick = () => goTab(TAB === "floor" ? "finished" : "floor");
  const ban = $("#banner");
  if (ban) ban.onclick = () => { touch(); markSeen(""); };
  const gfb = $("#glassfirst");
  if (gfb) gfb.onclick = () => { setGlassFirst(!GLASSFIRST); touch(); render(); };
  [["#vall", "all"], ["#vmine", "mine"], ["#vasg", "assigned"]].forEach(([s, v]) => {
    const b = $(s);
    if (b) b.onclick = () => {
      if (VIEW !== v) { setView(v); try { window.scrollTo(0, 0); } catch (e) {} }
      touch(); render();
    };
  });
  render();
  await readPeople();
  await readAssign();
  await readList();
  await NOTES.read();
  await flushQueue();
  TICK.start();
  if (peopleT) clearInterval(peopleT);
  peopleT = setInterval(readPeople, PEOPLE_MS);
  if (buildT) clearInterval(buildT);
  checkBuild(); buildT = setInterval(checkBuild, BUILD_MS);
}

(async function boot() {
  STU.stuApplyTheme(STU.stuThemeNow());
  /* an expired sign-in renews itself when it can - never while a tap is owed */
  if (CW.stationRenew) CW.stationRenew(() => !Object.keys(QUEUE).length && !Object.keys(LOGQ).length);
  $("#signinbtn").onclick = async () => {
    try { await CW.signIn(CW.LIST_SCOPES); await start(); }
    catch (e) { STU.stuGate(true, "Sign-in failed:\n" + ((e && e.message) || e)); }
  };
  try {
    const acct = await CW.initAuth();
    if (acct) await start(); else STU.stuGate(true);
  } catch (e) {
    STU.stuGate(true, "Startup problem:\n" + ((e && e.message) || e));
  }
})();
