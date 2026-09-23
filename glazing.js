/* The Glazing station page - the glazer's own screen.

   What this page can do, in full: read four SharePoint lists in the `Floor
   stations` site ("Glazing station", "Station people", "Station log", "Station
   comments"), PATCH the Glazed counter of a Glazing station row - with its By
   and At and the last-touch pair beside it - POST one line to the Station log
   for each counter write that succeeded, and POST one note per word somebody
   types for the office. That is all.

   There is no workbook here - no exceljs, no parser.js, no download, no Excel
   API path anywhere in this file - no delete, no editing of a job's facts, no
   export, and no link back to the master dashboard.

   ONE NUMBER PER JOB: units glazed, out of the job's WINDOWS - doors are not
   glazed at this station (owner, 2026-09-23), so they are fed as a fact and
   counted nowhere, and a job of nothing but doors never reaches this page.
   There are no stages here and no product groups, which is why this page is shorter
   than welding's rather than a copy of it: a card is a job, a stepper and a
   note box.

   The three station pages share station-core.js (the list mechanics, the
   people, the log lines, the note channel, the delta merge and the board diff),
   glazing-core.js's `GLAZE` definition for this station's own columns and
   rules, and station-ui.js for the theme, the gate and the picker. What is here
   is this page: its queue, its board and its drawing.                        */

const $ = s => document.querySelector(s);
const esc = STU.stuEsc;
const GZ = GLZC;

/* This page's own storage keys. Deliberately not the other pages': the same
   tablet can have more than one open, and a glazer's queue must never be read
   as a glass or a welding tap. The theme is the one thing they share - it
   belongs to the device (STU_THEME_KEY). */
const PERSON_KEY = "cw_glzperson";
const QUEUE_KEY = "cw_glzq";
const LOGQ_KEY = "cw_glzlogq";
const RETRY_MS = 5000;
const PEOPLE_MS = 600000;                // the people list is re-read every ten minutes

let ITEMS = [];                 // the Glazing station list, as last read
let TOKEN = null;               // the delta token: what to ask for next time
let PEOPLE = [];
let READY = false;              // the first read of the board has answered
let PEOPLE_READ = false;
let PROBLEM = "";               // "site" | "list" | "people" | "consent" | "reauth"
let SOFT = "";                  // a passing failure: the last board stays, a line above it
let LASTREAD = 0;
let QUERY = "";
let FINOPEN = false;            // the Finished group is expanded
let PERSON = null;
let LAST_TAP = 0;
let PINFOR = null, PINTYPED = "", PINBAD = false;
let refreshT = null, retryT = null, buildT = null, peopleT = null, lockT = null;
let flushing = false;
const DELTA_OFF_MS = 300000;
let DELTA_OFF = 0;
const deltaOff = () => !!(DELTA_OFF && Date.now() - DELTA_OFF < DELTA_OFF_MS);

/* ---- who is on the station -------------------------------------------------
   The name is not a login. It decides nothing about permission - the station
   account's own access does that - and everything about whose name goes into
   the log. Ten minutes without a tap and the picker comes back, because the
   tablet is passed around. */
function loadPerson() {
  let raw = null;
  try { raw = JSON.parse(localStorage.getItem(PERSON_KEY) || "null"); } catch (e) { raw = null; }
  if (!raw || !raw.name) return;
  if (ST.personExpired(raw.at, Date.now(), ST.PERSON_LOCK_MS)) return;
  const hit = PEOPLE.find(p => p.name === raw.name);
  if (!hit) return;
  PERSON = hit; LAST_TAP = Number(raw.at) || Date.now();
}
function savePerson() {
  try {
    if (PERSON) localStorage.setItem(PERSON_KEY, JSON.stringify({ name: PERSON.name, at: LAST_TAP }));
    else localStorage.removeItem(PERSON_KEY);
  } catch (e) {}
}
function touch() { LAST_TAP = Date.now(); savePerson(); }
function pickPerson(p) { PERSON = p; PINFOR = null; PINTYPED = ""; PINBAD = false; touch(); render(); }
function switchPerson() {
  PERSON = null; PINFOR = null; PINTYPED = ""; PINBAD = false;
  savePerson(); render();
}
function lockIfIdle() {
  if (!PERSON) return;
  if (!ST.personExpired(LAST_TAP, Date.now(), ST.PERSON_LOCK_MS)) return;
  PERSON = null; savePerson(); render();
}
function who() { return PERSON ? PERSON.name : ""; }
/* One stage here, so anybody signed in holds it. canStage is still asked - a
   Stages column typed by hand can say anything, and a person holding nothing
   may move nothing, on this station exactly as on the other two. */
const mayGlaze = () => ST.canStage(PERSON, GZ.GLZ_STAGE);

/* ---- the queues -------------------------------------------------------------
   One entry per list row - which, with one counter per job, is also exactly
   what one log line is about. A later tap on the same row overwrites the number
   already waiting: the floor's latest count is what the office should see, not
   a replay of every button press.

   Kept in localStorage, so a tablet closed mid-write still owes it - and read
   back through the definition's own whitelist, because anyone holding the
   tablet can edit that storage and no edit of it may ever put a job fact on the
   wire. */
let QUEUE = {};
let LOGQ = {};

function cleanQueue(raw) {
  const out = {};
  Object.keys(raw || {}).forEach(k => {
    const e = raw[k];
    if (!e || !e.id) return;
    const v = Number(e.value);
    if (!isFinite(v)) return;
    const from = Number(e.from);
    out[String(e.id)] = { id: String(e.id), value: Math.round(v),
                          who: String(e.who || ""), at: String(e.at || ""),
                          job: String(e.job || ""), title: String(e.title || ""),
                          site: String(e.site || ""),
                          from: isFinite(from) ? Math.round(from) : 0, err: 0 };
  });
  return out;
}
function cleanLogQ(raw) {
  const out = {};
  Object.keys(raw || {}).forEach(k => {
    const e = raw[k];
    if (!e || !e.fields || !e.fields.Title) return;
    /* rebuilt through ST.logFields, so whatever was in storage comes out as the
       eight columns of the log list and nothing else */
    out[String(k)] = { key: String(k), err: 0, fields: ST.logFields(GZ.glzLogEntry({
      job: e.fields.Title, from: e.fields.From, to: e.fields.To,
      who: e.fields.Who, at: e.fields.At })) };
  });
  return out;
}
try { QUEUE = cleanQueue(JSON.parse(localStorage.getItem(QUEUE_KEY) || "{}")); } catch (e) { QUEUE = {}; }
try { LOGQ = cleanLogQ(JSON.parse(localStorage.getItem(LOGQ_KEY) || "{}")); } catch (e) { LOGQ = {}; }
function saveQueue() {
  try { localStorage.setItem(QUEUE_KEY, JSON.stringify(QUEUE)); } catch (e) {}
  try { localStorage.setItem(LOGQ_KEY, JSON.stringify(LOGQ)); } catch (e) {}
}

/* Who and when belong to the moment of the tap, not to the moment the write
   happens to leave: a tap made by the morning shift and sent after a reload
   three hours later was still theirs, at the time they made it.

   `from` belongs to the moment this row first went into the queue: it is the
   number the office could last see. Re-deriving it at flush time would read
   whatever the list says by then - and a write whose response was lost, or a
   poll that has already merged this very change, would make From equal To and
   swallow the log line entirely. A run of merged taps keeps the From of the
   first of them, which is what makes one line say 5 to 8.

   The site is stamped on with the item id, because a SharePoint item id only
   means anything in the list it came from; the Title is what finds the row
   again if the lists ever move. */
function queueTap(rec, value) {
  const k = String(rec.id);
  const had = QUEUE[k];
  QUEUE[k] = { id: k, value: value, who: who(), at: new Date().toISOString(),
               job: rec.job, title: rec.title || rec.job, site: SITEID || "",
               from: had ? had.from : listValue(rec.id), err: 0 };
  saveQueue();
}
/** The counter this tablet still owes for a row, so the screen shows the tapped
    number rather than the number the list last answered with. */
const queuedFor = id => (QUEUE[String(id)] ? { Glazed: QUEUE[String(id)].value } : null);
const owedFor = id => !!QUEUE[String(id)];
const badFor = id => !!(QUEUE[String(id)] && QUEUE[String(id)].err);
/** A tap that was dropped because somebody said something later, kept until the
    card is redrawn from a list that agrees - never dropped in silence. */
let LOST = {};
const lostFor = job => LOST[String(job).trim().toUpperCase()] || null;

/** What the list last said this counter was - the From of the log line. The
    queue is deliberately not consulted: From means "the number the office could
    see before this write", and that is the list's number. */
function listValue(id) {
  const it = ITEMS.find(x => String(x.id) === String(id));
  if (!it) return 0;
  const v = Number((it.fields || {}).Glazed);
  return isFinite(v) ? Math.round(v) : 0;
}

/** Re-base or drop every queued tap against the list as it now stands. The
    three answers, and why each is the right one, are in glazing-core's
    glzRebase - this is the part that has a queue and a screen to tell. */
function rebaseQueue() {
  let moved = false;
  Object.keys(QUEUE).forEach(k => {
    const e = QUEUE[k];
    const it = ITEMS.find(x => String(x.id) === String(e.id));
    if (!it) return;
    const r = GZ.glzRebase(e, it.fields || {});
    if (r.action === "drop") {
      LOST[String(e.job).trim().toUpperCase()] = { job: e.job, value: e.value };
      console.warn("[glazing] dropping a queued tap for " + e.job + ": " + e.value +
        " was tapped at " + e.at + ", and the row was moved after that (" +
        ((it.fields || {}).DoneAt || "") + ") - tap it again if it is still right");
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

/** The item a queued tap is about, in the list as it is now. Normally that is
    simply the id it was queued with. After a move it is not: the id belongs to
    the old site's list, so the row is found again by its Title, which is unique
    and is the one thing about it that does not change. Answering null means the
    row is not in the new list at all - the entry is dropped rather than written
    somewhere at random. */
function currentId(e) {
  if (!e.site || e.site === SITEID) return e.id;
  const want = String(e.title || "").trim().toUpperCase();
  let best = null;
  ITEMS.forEach(it => {
    const t = String((it.fields || {}).Title == null ? "" : it.fields.Title).trim().toUpperCase();
    if (t !== want) return;
    if (!best || Number(it.id) < Number(best.id)) best = it;      // the oldest, as everywhere else
  });
  return best ? String(best.id) : null;
}

/** Send everything owed, draining rather than snapshotting: a tap that lands
    while a write is in flight goes out in the same pass instead of waiting for
    a timer that nothing had armed.

    Order matters, and it is the whole point of the log: the counter first, and
    the log line only once that counter really landed. A log write that fails is
    owed like any other and never holds a counter up. */
async function flushQueue() {
  if (flushing) return;
  if (!Object.keys(QUEUE).length && !Object.keys(LOGQ).length) return;
  /* no site resolved means no list to write to. Coming back in five seconds is
     the whole of the answer: guessing one would send the tablet at the
     workbook, which is the one thing it must never do. */
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
      /* WHAT WENT, CAPTURED BEFORE THE AWAIT AND NOT READ BACK OFF THE ENTRY.
         `e` IS `QUEUE[k]` - the same object - so comparing e.value against
         QUEUE[k].value after the write compares a thing with itself and always
         agrees, and a tap or a re-base that landed while the PATCH was in the
         air would be deleted as though it had been sent. (welding.js learned
         this; docs/HISTORY.md B19.) */
      const want = e.value, when = e.at, whose = e.who, wantFrom = e.from;
      const body = GZ.glzFloorOnly(GZ.glzTapFields(want, whose, when));
      const id = currentId(e);
      if (!id) {
        console.warn("[glazing] dropping a queued tap for " + e.job +
                     ": that row is not in the list this tablet is now reading");
        delete QUEUE[k]; saveQueue(); continue;
      }
      e.id = id; e.site = SITEID;
      try {
        await CW.listPatch(GZ.GLZ_LIST, id, body, listOpts());
        wrote = true;
        /* the counter is in the list now, so the log line describing it may be
           owed - and only now: a line about a write that never happened would
           be a lie in a list nothing ever deletes from. */
        queueLog({ id: id, job: e.job, value: want, at: when, who: whose, from: wantFrom });
        const now = QUEUE[k];
        if (now && now.value === want && now.at === when) delete QUEUE[k];
        else if (now) { now.err = 0; delete tried[k]; }     // moved mid-write: run it again
        /* keep the local copy in step, so a second tap's From is right before
           the next poll has been anywhere near SharePoint */
        const it = ITEMS.find(x => String(x.id) === String(id));
        if (it) (it.fields = it.fields || {}).Glazed = want;
      } catch (err) {
        if (QUEUE[k]) QUEUE[k].err = 1;
        /* only a 404 is a reason to doubt the cached site: a refusal or a bad
           gateway says nothing about where the list is */
        if (CW.isMissing && CW.isMissing(err) && CW.forgetStationSite)
          CW.forgetStationSite(false, GZ.GLAZE.site);
        console.warn("[glazing] counter write failed for item " + id + ":", (err && err.message) || err);
      }
      saveQueue();
    }
    await flushLog();
  } finally {
    flushing = false;
  }
  render();
  if (wrote) await pollList();                 // the list agrees now: drop back to it
  if (Object.keys(QUEUE).length || Object.keys(LOGQ).length) armRetry();
}

/** One log line per sent write. The queue has already merged a run of quick
    taps into a single counter write, so this is one line saying 5 to 8 rather
    than three lines counting up to it. */
function queueLog(e) {
  const from = Number(e.from) || 0;
  if (from === e.value) return;                 // nothing actually moved
  const key = e.id + "|" + e.at;
  LOGQ[key] = { key: key, err: 0, fields: ST.logFields(GZ.glzLogEntry({
    job: e.job, from: from, to: e.value, who: e.who, at: e.at })) };
  saveQueue();
}
async function flushLog() {
  if (!SITEID) return;                          // nowhere to write it yet: it stays owed
  const keys = Object.keys(LOGQ);
  for (let i = 0; i < keys.length; i++) {
    const e = LOGQ[keys[i]];
    if (!e) continue;
    try {
      await CW.listAdd(ST.LOG_LIST, e.fields, logOpts());
      delete LOGQ[keys[i]];
    } catch (err) {
      e.err = 1;
      console.warn("[glazing] log line not written yet:", (err && err.message) || err);
    }
    saveQueue();
  }
}

/* ---- the lists ---- */
let SITEID = null;
const listOpts = () => ({ siteId: SITEID, fields: GZ.GLZ_FIELDS });
const peopleOpts = () => ({ siteId: SITEID, fields: ST.PEOPLE_FIELDS });
const logOpts = () => ({ siteId: SITEID, fields: ST.LOG_FIELDS });
const commentOpts = () => ({ siteId: SITEID, fields: ST.COMMENT_FIELDS });

/* ---- a word for the office -------------------------------------------------
   The note channel, shipped 2026-09-15 for the glass tablet and reached here by
   changing ONE word. Every line of it - the composer, the thread, the row and
   the two list calls - is in station-core.js. */
const NOTES = ST.stationComments({
  station: GZ.GLZ_NAME,
  listItems: (name, o) => CW.listItems(name, o),
  listAdd: (name, fields, o) => CW.listAdd(name, fields, o),
  opts: commentOpts
});

/** Everything that can go wrong with a read, decided in one place. Only
    SharePoint actually saying "there is no such site" or "there is no such
    list" takes the board away; anything else - the wifi in a workshop, a bad
    gateway, a refused token - leaves the last board on screen with one small
    line above it. */
function trouble(e) {
  const m = (e && e.message) || String(e);
  console.warn("[glazing] could not read SharePoint:", m);
  if (CW.isMissing && CW.isMissing(e)) {
    SITEID = null;
    if (CW.forgetStationSite) CW.forgetStationSite(false, GZ.GLAZE.site);
  }
  if (/interaction_required|login_required/.test(m)) { PROBLEM = "reauth"; SOFT = ""; READY = false; }
  else if (/permission needed/.test(m)) { PROBLEM = "consent"; SOFT = ""; READY = false; }
  else if (CW.isMissing && CW.isMissing(e)) { PROBLEM = "site"; SOFT = ""; READY = false; }
  else SOFT = "cannot reach SharePoint — retrying";
}

/** The site this station's lists are in: `Floor stations` and nothing else
    (GZ.GLAZE.site === "floor"). A missing one is not an error and not a fallback
    - it is the quiet explained state, the same one a missing list gets. */
async function resolveSite() {
  if (!SITEID) SITEID = await CW.stationSite(GZ.GLAZE.site);
  return SITEID;
}

async function readPeople() {
  try {
    if (!(await resolveSite())) { PROBLEM = "site"; SOFT = ""; READY = false; render(); return false; }
    const items = await CW.listItems(ST.PEOPLE_LIST, peopleOpts());
    if (items == null) { PROBLEM = "people"; SOFT = ""; READY = false; render(); return false; }
    PEOPLE = ST.stationPeople(items, GZ.GLZ_NAME, GZ.GLAZE.stages);
    PEOPLE_READ = true;
    if (PROBLEM === "people") PROBLEM = "";
    if (PERSON) {
      const hit = PEOPLE.find(p => p.name === PERSON.name);
      PERSON = hit || null;
      if (!PERSON) savePerson();
    } else {
      loadPerson();
    }
    render();
    return true;
  } catch (e) {
    trouble(e);
    render();
    return false;
  }
}

/** The whole list, and a fresh delta token with it. The initial delta call
    answers every item AND the token for the next poll, so this costs one
    enumeration rather than two. */
async function readList() {
  try {
    if (!(await resolveSite())) { PROBLEM = "site"; SOFT = ""; READY = false; render(); return false; }
    if (siteMoved()) { TOKEN = null; DELTA_OFF = 0; }
    let items = null;
    if (deltaOff()) {
      items = await CW.listItems(GZ.GLZ_LIST, listOpts());
      if (items == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
      TOKEN = null;
    } else {
      try {
        const d = await CW.listDelta(GZ.GLZ_LIST, listOpts());
        if (d == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
        items = d.items.filter(x => !x.removed).map(x => ({ id: x.id, fields: x.fields }));
        TOKEN = d.next;
        DELTA_OFF = 0;
      } catch (e) {
        if (!CW.isDeltaRestart || !CW.isDeltaRestart(e)) throw e;
        if (!(CW.isDeltaResync && CW.isDeltaResync(e))) {
          DELTA_OFF = Date.now();
          console.warn("[glazing] the board list refused a delta; polling the plain way for five minutes");
        }
        items = await CW.listItems(GZ.GLZ_LIST, listOpts());
        if (items == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
        TOKEN = null;
      }
    }
    ITEMS = items; READY = true; PROBLEM = ""; SOFT = ""; LASTREAD = Date.now();
    rebaseQueue();                      // the list may have moved under a waiting tap
    render();
    flushQueue();
    return true;
  } catch (e) {
    trouble(e);
    render();
    return false;
  }
}

/** Have this station's lists moved to another site since the last pass? A delta
    token only means anything in the site it was issued in. */
let SITE_GEN = 0;
function siteMoved() {
  const m = CW.stationSiteMoves ? CW.stationSiteMoves(GZ.GLAZE.site) : 0;
  if (m === SITE_GEN) return false;
  SITE_GEN = m;
  return true;
}
async function pollList() {
  try { SITEID = (await CW.stationSite(GZ.GLAZE.site)) || SITEID; } catch (e) { /* keep the last one */ }
  if (siteMoved()) { TOKEN = null; DELTA_OFF = 0; }
  if (!TOKEN) return readList();
  try {
    const d = await CW.listDelta(GZ.GLZ_LIST, { siteId: SITEID, fields: GZ.GLZ_FIELDS, token: TOKEN });
    if (d == null) return readList();
    ITEMS = ST.mergeDelta(ITEMS, d.items);
    if (d.next) TOKEN = d.next;
    READY = true; SOFT = ""; LASTREAD = Date.now();
    rebaseQueue();
    render();
    flushQueue();
    return true;
  } catch (e) {
    if (CW.isDeltaRestart && CW.isDeltaRestart(e)) {
      if (!(CW.isDeltaResync && CW.isDeltaResync(e))) DELTA_OFF = Date.now();
      TOKEN = null;
      return readList();
    }
    trouble(e);
    render();
    return false;
  }
}

/* ---- taps ---- */
function tap(id, delta) {
  /* belt and braces: the stepper is drawn disabled for somebody who holds no
     stage, and tap() refuses it anyway - a disabled button is a drawing, and
     this is the rule */
  if (!PERSON || !mayGlaze()) return;
  const rec = recordById(id);
  if (!rec) return;
  /* somebody is working the screen, whether or not the number could move */
  touch();
  const value = GZ.glzApplyTap(rec, delta);
  if (value == null || value === rec.glazed) return;      // already at the clamp
  /* the apology for a dropped tap goes the moment they tap again: they have
     been told, and they are now saying it a second time */
  delete LOST[String(rec.job).trim().toUpperCase()];
  /* queued first, drawn second: boardNow() lays the queue over the list, so the
     new number is on screen before the write has left the tablet */
  queueTap(rec, value);
  render();
  flushQueue();
}

/** The list with anything this tablet still owes laid over the top of it - so a
    tapped number never flickers back to the list's older one. */
function itemsNow() {
  return ITEMS.map(it => {
    const q = queuedFor(String(it.id));
    if (!q) return it;
    return { id: it.id, fields: Object.assign({}, it.fields, q) };
  });
}
let RECS_BY_ID = {};
function boardNow() {
  const board = GZ.glzBoard(itemsNow());
  RECS_BY_ID = {};
  board.forEach(c => { RECS_BY_ID[String(c.id)] = c; });
  return board;
}
const recordById = id => RECS_BY_ID[String(id)] || null;

/* ---- drawing ---- */
/** A progress bar as a block, because a workshop screen is read at arm's length
    and a 3px rule is not. Width, not characters, so a 49-unit job and a 1-unit
    job draw the same shape. */
function barHtml(done, total) {
  const pct = total > 0 ? Math.max(0, Math.min(100, Math.round((done / total) * 100))) : 0;
  return '<span class="wbar" aria-hidden="true"><span class="wbarfill" style="width:' + pct + '%"></span></span>';
}

/** The one stepper: the word, the count, the bar and the three buttons that
    move it - on the card itself, where a hand can reach them. Nothing to open,
    nothing to scroll. Every target is at least 56 px, because it is tapped with
    a work glove on. */
function stepHtml(c) {
  const mine = mayGlaze();
  const off = mine ? "" : ' disabled aria-disabled="true"';
  const b = (t, act, cls) => '<button class="' + cls + '" data-id="' + esc(c.id) +
    '" data-act="' + esc(act) + '"' + off + '>' + t + '</button>';
  const full = c.glazed >= c.total;
  return '<div class="step' + (mine ? "" : " locked") + ' c-' + (c.colour || "none") + '">' +
    '<span class="stepl">Glazed' +
      (mine ? '<span class="stepleft tab">' + esc(GZ.glzLeftWords(c.left)) + '</span>'
            : ' <span class="nomine">not yours</span>') + '</span>' +
    '<span class="stepmid">' +
      '<span class="stepn tab' + (full ? " full" : "") + '">' + c.glazed + ' / ' + c.total + '</span>' +
      barHtml(c.glazed, c.total) + '</span>' +
    '<span class="stepc">' + b("&minus;", "-1", "sbtn") + b("+", "1", "sbtn") +
      b(full ? "None" : "All", full ? "none" : "all", "sall") + '</span>' +
    '</div>';
}

/** Everything inside one card. The card element itself is kept between draws
    (see paintBoard), so only this string is ever rebuilt. */
function cardInner(c) {
  const owed = owedFor(c.id), bad = badFor(c.id);
  const lost = lostFor(c.job);
  /* no quantity line under the head: since 2026-09-23 the head already says
     "6 windows" and the quantity said the same six a second time */
  return '<div class="chead">' +
      '<span class="cond job">' + esc(c.job) + '</span>' +
      '<span class="cust">' + esc(c.customer || "—") + '</span>' +
      '<span class="cunits tab">' + esc(GZ.glzUnitWords(c.total)) + '</span>' +
    '</div>' +
    (c.comment
      ? '<div class="cfacts">' +
        '<span class="ccmt">“' + esc(c.comment) + '”</span></div>'
      : "") +
    stepHtml(c) +
    (lost ? '<div class="unsaved">' + esc(String(lost.value)) +
        ' was not saved — the office changed this job after that tap</div>' : "") +
    (bad ? '<div class="unsaved">not saved yet — retrying</div>'
         : owed ? '<div class="saving">saving…</div>' : "") +
    /* under everything the card counts: a word for the office about this job,
       and this station's own earlier words about it */
    NOTES.html(c.job);
}

/* ---- the picker ---- */
function pickerHtml() {
  return STU.stuPickerHtml({
    people: PEOPLE, pinFor: PINFOR, pinTyped: PINTYPED, pinBad: PINBAD,
    labelOf: () => "glazing",
    empty: "Nobody is set up for the Glazing station yet. Ask the office to add " +
           "you to the “Station people” list."
  });
}
function wirePicker(host) {
  STU.stuWirePicker(host, {
    onPerson: name => {
      const p = PEOPLE.find(x => x.name === name);
      if (!p) return;
      if (ST.pinOk(p, "")) { pickPerson(p); return; }       // no PIN column: straight in
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

/* ---- the board, drawn once and then patched ---------------------------------
   Six redraws a minute would throw away the scroll position and the tap a
   finger is already on its way to, so the cards are kept as nodes keyed by job
   and only the ones ST.boardDiff names are touched. */
let NODES = {};
let DOING = null, FINHEAD = null, FIN = null;
let BOARD_PREV = null;
let QSIG = {};                  // job -> what this tablet owed on it when last drawn
let PSIG = "";                  // who the cards were drawn for
let WIRED = false;

/* The "saving..." and "not saved yet" lines come from this tablet's own queue,
   which is not in the list and so is invisible to boardDiff. A row that has just
   failed to save therefore has to be named here, or the red line never appears
   on a board whose numbers did not move. */
function qState(c) {
  const lost = lostFor(c.job);
  return (owedFor(c.id) ? "o" : "") + (badFor(c.id) ? "b" : "") +
         (lost ? "!" + lost.value : "") +
         /* the note channel is not in the list either. The DRAFT is deliberately
            not in the signature - see dressCard. */
         "/" + NOTES.sig(c.job);
}
function pState() { return PERSON ? PERSON.name + "|" + (PERSON.stages || []).join(",") : ""; }

function makeCard(c) {
  const el = document.createElement("div");
  el.className = "card c-" + (c.colour || "none") + (c.finished ? " done" : "");
  if (el.dataset) el.dataset.job = c.job;
  el.setAttribute("data-job", c.job);
  el.innerHTML = cardInner(c);
  return el;
}
/** Redraw one card - and put the caret back if it was in this card's note box.
    Only the CARET is restored, never the text: the box is drawn from the draft,
    and the draft is what the input listener has been keeping current. */
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
  if (!DOING) {
    host.innerHTML = "";
    DOING = document.createElement("div"); DOING.className = "grp";
    FINHEAD = document.createElement("div"); FINHEAD.className = "grouphead";
    if (FINHEAD.addEventListener)
      FINHEAD.addEventListener("click", () => { FINOPEN = !FINOPEN; render(); });
    FIN = document.createElement("div"); FIN.className = "grp fin";
    host.appendChild(DOING); host.appendChild(FINHEAD); host.appendChild(FIN);
    NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
  }
  const diff = ST.boardDiff(BOARD_PREV, board, GZ.glzCardSig);
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
    /* a move BLURS whatever is inside it, so the caret is put back the way
       dressCard does after a redraw */
    const act = document.activeElement;
    const box = act && act.dataset && act.dataset.cmbox ? act : null;
    const at = box ? box.selectionStart : null;
    board.forEach(c => {
      const el = NODES[c.job];
      if (!el) return;
      (c.finished ? FIN : DOING).appendChild(el);
    });
    if (box && document.activeElement !== box) {
      try { box.focus(); if (box.setSelectionRange && at != null) box.setSelectionRange(at, at); }
      catch (e) {}
    }
  }
  const done = board.filter(c => c.finished).length;
  FINHEAD.textContent = done ? "Finished · " + done + (FINOPEN ? " ▾" : " ▸") : "";
  FINHEAD.hidden = !done;
  FIN.hidden = !done || !FINOPEN;
  BOARD_PREV = board;
}

/** Everything that is not a card: the header, the messages, the picker. */
function render() {
  const host = $("#board");
  if (!host) return;
  const hdr = $("#whois");
  if (hdr) hdr.textContent = PERSON ? PERSON.name + (mayGlaze() ? " · glazing" : " · no stages") : "";
  const sw = $("#switchbtn");
  if (sw) { sw.hidden = !PERSON; sw.style.display = PERSON ? "" : "none"; }

  const boarding = !PROBLEM && PEOPLE_READ && !!PERSON && READY;
  /* the search box belongs to the board: there is nothing to narrow on the
     picker, and a box over an error message only looks broken */
  const sb = $("#search");
  if (sb) { sb.hidden = !boarding; sb.style.display = boarding ? "" : "none"; }

  /* the board as it stands, before the box has narrowed it: the number in the
     header is read off this and is deliberately NOT narrowed - somebody looking
     a job number up must not make the day's work read shorter than it is */
  const live = boarding ? boardNow() : null;
  const gt = $("#gtotal");
  const on = boarding && mayGlaze();
  if (gt) {
    gt.hidden = !on;
    gt.style.display = on ? "" : "none";
    gt.textContent = on ? GZ.glzLeftWords(GZ.glzLeft(live)) : "";
  }
  const upd = $("#upd");
  if (upd) upd.textContent = LASTREAD ? "updated " + STU.stuAgo(LASTREAD) : "";
  const soft = $("#soft");
  if (soft) { soft.textContent = SOFT; soft.hidden = !SOFT; soft.style.display = SOFT ? "" : "none"; }

  /* Who are you? comes before the board and after any real problem with it */
  if (!boarding) {
    DOING = null; FIN = null; FINHEAD = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    if (!PROBLEM && PEOPLE_READ && !PERSON) { host.innerHTML = pickerHtml(); wirePicker(host); return; }
    host.innerHTML = '<div class="msg">' + esc(words()) + againHtml() + '</div>';
    wireAgain();
    return;
  }

  const board = GZ.glzFilter(live, QUERY);
  if (!board.length) {
    DOING = null; FIN = null; FINHEAD = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    host.innerHTML = '<div class="msg">' +
      (QUERY ? "No job on the board matches “" + esc(QUERY) + "”."
             : "Nothing on the board yet.") + '</div>';
    return;
  }
  paintBoard(host, board);
  wireBoard(host);
}

function words() {
  return PROBLEM === "reauth"
    ? "The sign-in has expired. Tap Sign out, then Sign in again."
    : PROBLEM === "consent"
    ? "Ask the office to grant the SharePoint permission."
    : PROBLEM === "people"
    ? "The “Station people” list is not in the floor’s site yet. Ask the office to add it."
    : PROBLEM === "list"
    ? "The “Glazing station” list is not in the floor’s site yet. Ask the office to add it — " +
      "nothing in the Excel file is involved."
    : PROBLEM === "site"
    ? "The floor’s SharePoint site is not there yet, or this account cannot see it. Ask the office."
    : "Reading the board…";
}
const againHtml = () => PROBLEM === "reauth"
  ? '<div><button class="again" id="reauth">Sign in again</button></div>'
  : '<div><button class="again" id="again">Try again</button></div>';
function wireAgain() {
  const a = $("#again");
  if (a) a.onclick = () => {
    PROBLEM = ""; SOFT = ""; SITEID = null; TOKEN = null;
    if (CW.forgetStationSite) CW.forgetStationSite(true, GZ.GLAZE.site);   // a person asked: look now
    render(); readPeople(); readList();
  };
  const r = $("#reauth");
  if (r) r.onclick = async () => {
    try { await CW.signIn(CW.LIST_SCOPES); PROBLEM = ""; SITEID = null; TOKEN = null; await start(); }
    catch (e) { console.warn("[glazing] sign-in again failed:", (e && e.message) || e); }
  };
}

/** One delegated listener for the whole board, wired once. Cards come and go
    under it; the listener does not. */
function wireBoard(host) {
  if (WIRED || !host.addEventListener) return;
  WIRED = true;
  host.addEventListener("click", ev => {
    /* anything at all on the board is a hand on the tablet, including a button
       that will not move: the ten-minute lock must not fire under a working
       hand */
    touch();
    let el = ev.target;
    for (let i = 0; el && i < 5; i++) {
      const d = el.dataset || {};
      if (d.act) {
        ev.stopPropagation();
        if (el.disabled) return;
        tap(d.id, d.act === "all" || d.act === "none" ? d.act : Number(d.act));
        return;
      }
      if (d.cmt) {
        ev.stopPropagation();
        NOTES.toggle(d.cmt);
        render();
        if (NOTES.isOpen(d.cmt) && NOTES.state.ok !== true)
          NOTES.read().then(() => render(), () => {});
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
  /* what is typed is remembered but NOT redrawn: a card rebuilt on every
     keystroke is a caret lost on every keystroke */
  host.addEventListener("input", ev => {
    const d = (ev.target && ev.target.dataset) || {};
    if (!d.cmbox) return;
    touch();
    NOTES.setDraft(d.cmbox, ev.target.value);
  });
}

async function sendNote(job) {
  if (!PERSON) return;
  if (!NOTES.draftOf(job).trim()) return;
  const going = NOTES.send(job, who());
  render();
  await going;
  render();
}

/* ---- staying current -------------------------------------------------------
   This page is opened once and left on a tablet for weeks, which is exactly
   where a browser serving a script from cache does the most damage. */
const BUILD_MS = 120000;
let BUILD_NOW = null;
async function checkBuild() {
  try {
    const r = await fetch("version.json?t=" + Date.now(), { cache: "no-store" });
    if (!r.ok) return;
    const latest = (await r.json()).build;
    if (!latest) return;
    if (!BUILD_NOW) { BUILD_NOW = latest; return; }          // the first look is the baseline
    if (latest !== BUILD_NOW && !Object.keys(QUEUE).length && !Object.keys(LOGQ).length)
      location.reload(true);
  } catch (e) { /* offline or blocked - the board is what matters */ }
}

/** One turn of the ten-second clock. A named function rather than a closure
    inside setInterval, so a test can take exactly one turn of it. */
async function tickOnce() {
  render();
  if (!PEOPLE_READ) await readPeople();
  await pollList();
  if (await NOTES.poll()) render();
}

/* ---- the page ---- */
async function start() {
  STU.stuGate(false);
  STU.stuApplyTheme(STU.stuThemeNow());
  const tb = $("#themebtn");
  if (tb) tb.onclick = () =>
    STU.stuApplyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  $("#outbtn").onclick = () => { if (confirm("Sign out of the Glazing station?")) CW.signOut(); };
  const sw = $("#switchbtn");
  if (sw) sw.onclick = () => switchPerson();
  /* the box is in the header, outside #board, so using it never rebuilds the
     node the caret is in */
  const sb = $("#search");
  if (sb) sb.oninput = () => { QUERY = sb.value || ""; touch(); render(); };
  render();
  await readPeople();
  await readList();
  /* what has already been said about today's jobs, so a second shift does not
     retype the first shift's note */
  await NOTES.read();
  await flushQueue();                            // taps owed from a previous visit
  if (refreshT) clearInterval(refreshT);
  refreshT = setInterval(tickOnce, ST.REFRESH_MS);
  if (peopleT) clearInterval(peopleT);
  peopleT = setInterval(readPeople, PEOPLE_MS);
  if (lockT) clearInterval(lockT);
  lockT = setInterval(lockIfIdle, 15000);
  if (buildT) clearInterval(buildT);
  checkBuild(); buildT = setInterval(checkBuild, BUILD_MS);
}

(async function boot() {
  STU.stuApplyTheme(STU.stuThemeNow());
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
