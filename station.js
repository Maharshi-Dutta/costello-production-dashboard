/* The Glass station page - the floor's own screen.

   What this page can do, in full: read three SharePoint lists ("Glass
   station", "Station people", "Station log"), PATCH the Cut / Hotmelt /
   Glazed counters of a Glass station row - with that stage's By and At and the
   last-touch pair beside them - and POST one line to the Station log for each
   counter write that succeeded. That is all. There is no workbook here - no
   exceljs, no parser.js, no download, no Excel API path anywhere in this file
   - no delete, no editing of a job's facts, no comments, no export, and no
   link back to the master dashboard.

   It is used on a shared tablet, signed in once with a shared station account
   and then passed between three people for weeks. So: the person picks their
   name (and their PIN) before they can move anything, the name locks itself
   after ten quiet minutes, the stages somebody does not hold are greyed rather
   than hidden, every tap shows on screen at once and is owed rather than
   awaited, the owed writes survive a reload in localStorage, and nothing on
   this page can open a modal or a consent popup while somebody is holding a
   sheet of glass.                                                           */

const $ = s => document.querySelector(s);
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

const PERSON_KEY = "cw_person";          // who is on the station right now, and when they last tapped
const QUEUE_KEY = "cw_stationq";         // counter writes this tablet owes
const LOGQ_KEY = "cw_stationlogq";       // log lines this tablet owes
const THEME_KEY = "cw_stationtheme";     // this device's own light/dark choice
const RETRY_MS = 5000;
const PEOPLE_MS = 600000;                // the people list is re-read every ten minutes

let ITEMS = [];                 // the Glass station list, as last read
let TOKEN = null;               // the delta token: what to ask for next time
let PEOPLE = [];                // the people who may record something here
let READY = false;              // the first read of the board has answered
let PEOPLE_READ = false;        // ... and the first read of the people list
let PROBLEM = "";               // "site" | "list" | "people" | "consent" | "reauth" - each takes the board away
let SOFT = "";                  // a passing failure: the last board stays, with this line above it
let LASTREAD = 0;
let QUERY = "";                 // what is in the search box, if anything
let FINOPEN = false;            // the Finished group is expanded
let PERSON = null;              // the person who picked their name
let LAST_TAP = 0;               // when they last touched anything, for the lock
let PINFOR = null;              // the person whose PIN is being asked for
let PINTYPED = "";
let PINBAD = false;
let refreshT = null, retryT = null, buildT = null, peopleT = null, lockT = null;
let flushing = false;
/* A list that will not serve a delta at all must not be asked for one six
   times a minute: it is marked off for five minutes and polled the plain way
   until then. A 410 "your token is too old" is NOT that - delta is working
   there, the token is simply stale - so it only costs a fresh enumeration. */
const DELTA_OFF_MS = 300000;
let DELTA_OFF = 0;
const deltaOff = () => !!(DELTA_OFF && Date.now() - DELTA_OFF < DELTA_OFF_MS);

/* ---- the device's own theme ----
   The owner asked for dark by default on the tablet - a workshop screen at
   arm's length - with a switch for whoever prefers the other one. It is a
   property of the device, not of the person, so it lives in localStorage and
   nothing on SharePoint knows about it. */
function themeNow() {
  try { return localStorage.getItem(THEME_KEY) === "light" ? "light" : "dark"; } catch (e) { return "dark"; }
}
function applyTheme(t) {
  try { localStorage.setItem(THEME_KEY, t); } catch (e) {}
  const r = document.documentElement;
  if (r && r.dataset) r.dataset.theme = t;
  const b = $("#themebtn");
  if (b) b.textContent = t === "dark" ? "Light" : "Dark";
}

/* ---- who is on the station --------------------------------------------------
   The name is not a login: it says which stages the steppers will move and it
   is what goes into the log. It is kept with the time of the last tap, because
   that pair is what the ten-minute lock reads. */
function loadPerson() {
  let saved = null;
  try { saved = JSON.parse(localStorage.getItem(PERSON_KEY) || "null"); } catch (e) { saved = null; }
  if (!saved || !saved.name) return;
  if (ST.personExpired(saved.at, Date.now(), ST.PERSON_LOCK_MS)) return;   // gone stale while the tablet slept
  /* a stamp from the future - a clock change, an edited localStorage - would
     otherwise keep the name unlocked for as long as it is ahead by */
  LAST_TAP = Math.min(Number(saved.at) || 0, Date.now());
  /* the stages come from the list, never from storage: an edit to
     localStorage must not be able to hand somebody a stage they do not hold */
  const hit = PEOPLE.find(p => p.name === saved.name);
  PERSON = hit || null;
}
function savePerson() {
  try {
    if (PERSON) localStorage.setItem(PERSON_KEY, JSON.stringify({ name: PERSON.name, at: LAST_TAP }));
    else localStorage.removeItem(PERSON_KEY);
  } catch (e) {}
}
/** Every tap pushes the lock back. */
function touch() { LAST_TAP = Date.now(); savePerson(); }
function pickPerson(p) { PERSON = p; PINFOR = null; PINTYPED = ""; PINBAD = false; touch(); render(); }
function switchPerson() {
  PERSON = null; PINFOR = null; PINTYPED = ""; PINBAD = false; LAST_TAP = 0;
  /* the box belongs to whoever was just holding the tablet: the next person
     must not be handed a board narrowed to somebody else's search */
  QUERY = "";
  const sb = $("#search");
  if (sb) sb.value = "";
  savePerson(); render();
}
/** The lock. Checked on a timer as well as on every draw, because a tablet
    left alone is exactly the case it exists for. */
function lockIfIdle() {
  if (!PERSON) return;
  if (!ST.personExpired(LAST_TAP, Date.now(), ST.PERSON_LOCK_MS)) return;
  switchPerson();
}
/** What goes in CutBy / HotmeltBy / DoneBy and in the log's Who. */
function who() { return PERSON ? PERSON.name : ""; }

/* ---- the queues -------------------------------------------------------------
   One counter entry per list item AND stage - which is also exactly what one
   log line is about. A later tap on the same row and stage overwrites the
   number in the entry already waiting: the floor's latest count is what the
   office should see, not a replay of every button press.

   Kept in localStorage, so a tablet closed mid-write still owes it - and read
   back through station-core's own whitelist, because anyone holding the tablet
   can edit that storage and no edit of it may ever put a job fact on the wire. */
let QUEUE = {};
let LOGQ = {};
const qKey = (id, stage) => String(id) + "|" + String(stage);

function cleanQueue(raw) {
  const out = {};
  Object.keys(raw || {}).forEach(k => {
    const e = raw[k];
    if (!e || !e.id || !ST.STAGE_FIELD[e.stage]) return;
    const v = Number(e.value);
    if (!isFinite(v)) return;
    const from = Number(e.from);
    out[qKey(e.id, e.stage)] = { id: String(e.id), stage: String(e.stage), value: Math.round(v),
                                 who: String(e.who || ""), at: String(e.at || ""),
                                 job: String(e.job || ""), type: String(e.type || ""),
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
    /* rebuilt through logFields, so whatever was in storage comes out as the
       eight columns of the log list and nothing else */
    out[String(k)] = { key: String(k), fields: ST.logFields({
      job: e.fields.Title, station: e.fields.Station, type: e.fields.GlassType,
      stage: e.fields.Stage, from: e.fields.From, to: e.fields.To,
      who: e.fields.Who, at: e.fields.At }), err: 0 };
  });
  return out;
}
try { QUEUE = cleanQueue(JSON.parse(localStorage.getItem(QUEUE_KEY) || "{}")); } catch (e) { QUEUE = {}; }
try { LOGQ = cleanLogQ(JSON.parse(localStorage.getItem(LOGQ_KEY) || "{}")); } catch (e) { LOGQ = {}; }
function saveQueue() {
  try { localStorage.setItem(QUEUE_KEY, JSON.stringify(QUEUE)); } catch (e) {}
  try { localStorage.setItem(LOGQ_KEY, JSON.stringify(LOGQ)); } catch (e) {}
}

/* ---- the office's lock ------------------------------------------------------
   The office can mark a job's glass finished, and that job then goes read-only
   here: the steppers grey out and tap() refuses them, the same way a stage
   somebody does not hold is already refused. Only the office can take it off
   again. The tablet learns it from the OfficeDone column and from nothing else
   - it still cannot see the workbook.

   The awkward case is a tap already sitting in the queue when the lock arrives:
   the wifi was out, or the write is still in the air. Sending it would put the
   floor's number over a job the office has just called finished, and on a
   locked row nobody on the floor could correct it afterwards. The office acted
   later, so by the same last-writer-wins rule the office keeps this one and the
   tap is dropped - but never quietly. It is moved here, said on the card in
   red, and written to the console; it clears when the office unlocks the job. */
const BLOCKED_KEY = "cw_stationblocked";
let BLOCKED = {};
try { BLOCKED = JSON.parse(localStorage.getItem(BLOCKED_KEY) || "{}"); } catch (e) { BLOCKED = {}; }
const saveBlocked = () => { try { localStorage.setItem(BLOCKED_KEY, JSON.stringify(BLOCKED)); } catch (e) {} };
const lockedItem = it => String(((it && it.fields) || {}).OfficeDone == null
  ? "" : it.fields.OfficeDone).trim().toLowerCase() === "yes";
const blockedFor = job => BLOCKED[String(job).trim().toUpperCase()] || null;

/** Take out of the queue anything the office has locked under it, and clear the
    notice for anything the office has unlocked. Called after every read of the
    list and again before every flush, because the lock can arrive in either. */
function dropBlocked() {
  let changed = false;
  const lock = {};                       // item id -> is it locked
  const byJob = {};                      // job number -> is it locked
  ITEMS.forEach(it => {
    const on = lockedItem(it);
    lock[String(it.id)] = on;
    const job = ST.jobKey((it.fields || {}).Job || (it.fields || {}).Title);
    if (job) byJob[job] = on;
  });
  Object.keys(QUEUE).forEach(k => {
    const e = QUEUE[k];
    if (!e || !lock[String(e.id)]) return;
    const job = ST.jobKey(e.job);
    const note = BLOCKED[job] || { job: job, stages: [], at: "" };
    note.stages = note.stages.filter(s => s.stage !== e.stage)
                    .concat([{ stage: e.stage, value: e.value }]);
    note.at = e.at;
    BLOCKED[job] = note;
    delete QUEUE[k];
    changed = true;
    console.warn("[station] " + job + " was marked complete by the office while a tap was " +
                 "still owed on it: " + ST.stageLabel(e.stage) + " " + e.value + " was not saved");
  });
  Object.keys(BLOCKED).forEach(job => {
    if (byJob[job] === false) { delete BLOCKED[job]; changed = true; }
  });
  if (changed) { saveQueue(); saveBlocked(); }
  return changed;
}

/* Who and when belong to the moment of the tap, not to the moment the write
   happens to leave: a tap made by the morning shift and sent after a reload
   three hours later was still theirs, at the time they made it. */
function queueTap(row, job, stage, value) {
  const k = qKey(row.id, stage);
  const had = QUEUE[k];
  /* From belongs to the moment this row and stage first went into the queue:
     it is the number the office could last see. Re-deriving it at flush time
     would read whatever the list says by then - and a write whose response was
     lost, or a poll that has already merged this very change, would make From
     equal To and swallow the log line entirely. A run of merged taps keeps the
     From of the first of them, which is what makes one line say 5 to 8. */
  /* the site is stamped on with the item id, because a SharePoint item id
     only means anything in the list it came from. If the floor's lists move
     between the tap and the write, this is what stops the write landing on
     whatever row happens to carry that number in the new list. */
  QUEUE[k] = { id: String(row.id), stage: stage, value: value, who: who(),
               at: new Date().toISOString(), job: job, type: ST.GLASS_TYPE, site: SITEID || "",
               from: had ? had.from : listValue(row.id, stage), err: 0 };
  saveQueue();
}
/** The counters this tablet is still owing for a row, so the screen shows the
    tapped number rather than the number the list last answered with. */
function queuedFor(id) {
  const out = {};
  Object.keys(QUEUE).forEach(k => {
    const e = QUEUE[k];
    if (String(e.id) !== String(id)) return;
    out[ST.STAGE_FIELD[e.stage]] = e.value;
  });
  return Object.keys(out).length ? out : null;
}
const owedFor = id => Object.keys(QUEUE).some(k => String(QUEUE[k].id) === String(id));
const badFor = id => Object.keys(QUEUE).some(k => String(QUEUE[k].id) === String(id) && QUEUE[k].err);

/** The item this queued tap is about, in the list as it is now. Normally that
    is simply the id it was queued with. After a move it is not: the id belongs
    to the old site's list, so the row is found again by its Title, which is
    unique and is the one thing about it that does not change. Answering null
    means the row is not in the new list at all - the entry is dropped rather
    than written somewhere at random. */
function currentId(e) {
  if (!e.site || e.site === SITEID) return e.id;
  const want = String(e.job).trim().toUpperCase();
  let best = null;
  ITEMS.forEach(it => {
    const t = String((it.fields || {}).Title == null ? "" : it.fields.Title).trim().toUpperCase();
    if (t !== want) return;
    if (!best || Number(it.id) < Number(best.id)) best = it;     // the oldest, as everywhere else
  });
  return best ? String(best.id) : null;
}

/** A tap is owed as a NUMBER, and a number means something only against what
    the list said when the tap was made. The office can raise a counter under
    a waiting tap - the feeder seeds a row from the office's own record, and a
    tablet that was out of wifi all morning can be holding a `+1` made against
    nought while the list has moved to twelve. Sending 1 would put the job
    back. So a queued entry whose row has risen past its `from` is re-based to
    the same movement against the new number: +1 on twelve becomes thirteen,
    clamped to the total, and its `from` moves with it so the log line still
    says what actually changed.

    Only a rise re-bases. A counter that has gone DOWN in the list is somebody
    correcting it on another tablet, and this tap is the newer statement of
    what is on the bench. */
function rebaseQueue() {
  let moved = false;
  Object.keys(QUEUE).forEach(k => {
    const e = QUEUE[k];
    const it = ITEMS.find(x => String(x.id) === String(e.id));
    if (!it) return;
    const f = it.fields || {};
    const now = Number(f[ST.STAGE_FIELD[e.stage]]);
    const was = Number(e.from);
    if (!isFinite(now) || !isFinite(was) || now <= was) return;
    const total = Math.max(0, Math.round(Number(f.Total) || 0));
    e.value = Math.max(0, Math.min(total, Math.round(now + (Number(e.value) - was))));
    e.from = Math.round(now);
    moved = true;
  });
  if (moved) saveQueue();
}

/** Come back to the queue in a moment. Armed whenever anything is still owed,
    so nothing can be left sitting there with no one due to pick it up. */
function armRetry() {
  if (retryT) return;
  retryT = setTimeout(() => { retryT = null; flushQueue(); }, RETRY_MS);
}

/** What the list last said this counter was - the From of the log line. The
    queue is deliberately not consulted: From means "the number the office
    could see before this write", and that is the list's number. */
function listValue(id, stage) {
  const it = ITEMS.find(x => String(x.id) === String(id));
  if (!it) return 0;
  const v = Number((it.fields || {})[ST.STAGE_FIELD[stage]]);
  return isFinite(v) ? Math.round(v) : 0;
}

/** Send everything owed, draining rather than snapshotting: a tap that lands
    while a write is in flight goes out in the same pass instead of waiting for
    a timer that nothing had armed. Each row+stage is tried once per pass; one
    that changed under a successful write is put back in the running.

    Order matters, and it is the whole point of the log: the counter first, and
    the log line only once that counter really landed. A log write that fails
    is owed like any other and never holds a counter up.                     */
async function flushQueue() {
  if (flushing) return;
  if (!Object.keys(QUEUE).length && !Object.keys(LOGQ).length) return;
  /* no site resolved means no list to write to. Coming back in five seconds is
     the whole of the answer: guessing one would send the tablet at the
     workbook, which is the one thing it must never do. */
  if (!READY || !SITEID) { armRetry(); return; }
  /* again here, not only after a read: the lock can have arrived in the poll
     that ran while this queue was waiting for its five seconds */
  dropBlocked();
  if (!Object.keys(QUEUE).length && !Object.keys(LOGQ).length) { render(); return; }
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
      const body = ST.floorOnly(ST.tapFields(e.stage, e.value, e.who, e.at));
      const id = currentId(e);
      if (!id) {
        /* the row this tap was about is not in the list any more - the lists
           moved and it was not copied across. Better lost than written onto
           somebody else's job. */
        console.warn("[station] dropping a queued tap for " + e.job + " " + e.type +
                     ": that row is not in the list this tablet is now reading");
        delete QUEUE[k]; saveQueue(); continue;
      }
      e.id = id; e.site = SITEID;
      try {
        await CW.listPatch(ST.STATION_LIST, id, body, listOpts());
        wrote = true;
        /* the counter is in the list now, so the log line describing it may be
           owed - and only now: a line about a write that never happened would
           be a lie in a list nothing ever deletes from */
        queueLog(e);
        const now = QUEUE[k];
        if (now && now.value === e.value && now.at === e.at) delete QUEUE[k];
        else if (now) { now.err = 0; delete tried[k]; }   // tapped again mid-write: run it again
        /* keep the local copy in step, so a second tap's From is right before
           the next poll has been anywhere near SharePoint */
        const it = ITEMS.find(x => String(x.id) === String(id));
        if (it) (it.fields = it.fields || {})[ST.STAGE_FIELD[e.stage]] = e.value;
      } catch (err) {
        if (QUEUE[k]) QUEUE[k].err = 1;
        /* only a 404 is a reason to doubt the cached site: a refusal or a bad
           gateway says nothing about where the list is, and dropping the site
           over one would restart the ten-minute cadence for nothing */
        if (CW.isMissing && CW.isMissing(err) && CW.forgetStationSite) CW.forgetStationSite();
        console.warn("[station] counter write failed for item " + id + ":", (err && err.message) || err);
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
  if (from === e.value) return;                 // nothing actually moved: nothing to record
  const key = e.id + "|" + e.stage + "|" + e.at;
  LOGQ[key] = { key: key, err: 0, fields: ST.logFields({
    job: e.job, type: e.type, stage: e.stage, from: from, to: e.value, who: e.who, at: e.at }) };
  saveQueue();
}
/** Send the owed log lines. A failure here is kept and retried with everything
    else; it never blocks a counter and never shows anybody a window. */
async function flushLog() {
  if (!SITEID) return;                          // nowhere to write it yet: it stays owed
  const keys = Object.keys(LOGQ);
  for (let i = 0; i < keys.length; i++) {
    const e = LOGQ[keys[i]];
    if (!e) continue;
    try {
      await CW.listAdd(ST.LOG_LIST, e.fields, listOpts());
      delete LOGQ[keys[i]];
    } catch (err) {
      e.err = 1;
      console.warn("[station] log line not written yet:", (err && err.message) || err);
    }
    saveQueue();
  }
}

/* ---- the lists ---- */
let SITEID = null;
const listOpts = () => ({ siteId: SITEID, fields: ST.STATION_FIELDS });
const peopleOpts = () => ({ siteId: SITEID, fields: ST.PEOPLE_FIELDS });

/** Everything that can go wrong with a read, decided in one place. Only
    SharePoint actually saying "there is no such site" or "there is no such
    list" takes the board away; anything else - the wifi in a workshop, a bad
    gateway, a refused token - leaves the last board on screen with one small
    line above it, because a floor screen that goes blank and says "ask the
    office to make a list" when the list is fine is worse than a board that is
    a minute out of date. */
function trouble(e) {
  const m = (e && e.message) || String(e);
  console.warn("[station] could not read SharePoint:", m);
  /* Only "there is no such thing here" is a reason to doubt the cached site.
     A refusal, a bad gateway or a dead workshop wifi says nothing about where
     the lists are, and dropping the site over one would restart the whole
     re-check cadence - and, on a 403, could flip a page that is happily on the
     real site over to the fallback. */
  if (CW.isMissing && CW.isMissing(e)) {
    SITEID = null;
    if (CW.forgetStationSite) CW.forgetStationSite();
  }
  if (/interaction_required|login_required/.test(m)) { PROBLEM = "reauth"; SOFT = ""; READY = false; }
  else if (/permission needed/.test(m)) { PROBLEM = "consent"; SOFT = ""; READY = false; }
  else if (CW.isMissing && CW.isMissing(e)) { PROBLEM = "site"; SOFT = ""; READY = false; }
  else SOFT = "cannot reach SharePoint — retrying";     // READY and ITEMS stay exactly as they were
}

/** The people who may record something here. Read at start and every ten
    minutes: somebody added to the list in SharePoint appears on the tablet
    without anyone touching it. */
async function readPeople() {
  try {
    if (!SITEID) SITEID = await CW.stationSite();
    if (!SITEID) { PROBLEM = "site"; SOFT = ""; READY = false; render(); return false; }
    const items = await CW.listItems(ST.PEOPLE_LIST, peopleOpts());
    if (items == null) { PROBLEM = "people"; SOFT = ""; READY = false; render(); return false; }
    PEOPLE = ST.stationPeople(items, ST.STATION_NAME);
    PEOPLE_READ = true;
    if (PROBLEM === "people") PROBLEM = "";
    /* the stages a chosen person holds are re-read with them: a change in
       SharePoint reaches the tablet without anybody signing out */
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
    enumeration rather than two. A tenant or a list where delta is refused
    falls back to the plain read and simply polls that way. */
async function readList() {
  try {
    if (!SITEID) SITEID = await CW.stationSite();
    if (!SITEID) { PROBLEM = "site"; SOFT = ""; READY = false; render(); return false; }
    if (siteMoved()) { TOKEN = null; DELTA_OFF = 0; }
    let items = null;
    if (deltaOff()) {
      items = await CW.listItems(ST.STATION_LIST, listOpts());
      if (items == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
      TOKEN = null;
    } else {
      try {
        const d = await CW.listDelta(ST.STATION_LIST, listOpts());
        if (d == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
        items = d.items.filter(x => !x.removed).map(x => ({ id: x.id, fields: x.fields }));
        TOKEN = d.next;
        DELTA_OFF = 0;                               // it served one: not a refusing list
      } catch (e) {
        if (!CW.isDeltaRestart || !CW.isDeltaRestart(e)) throw e;
        if (!(CW.isDeltaResync && CW.isDeltaResync(e))) {
          DELTA_OFF = Date.now();
          console.warn("[station] the board list refused a delta; polling the plain way for five minutes");
        }
        items = await CW.listItems(ST.STATION_LIST, listOpts());
        if (items == null) { PROBLEM = "list"; SOFT = ""; READY = false; render(); return false; }
        TOKEN = null;                                // no delta this time: poll the long way
      }
    }
    ITEMS = items; READY = true; PROBLEM = ""; SOFT = ""; LASTREAD = Date.now();
    rebaseQueue();                      // the list may have moved under a waiting tap
    dropBlocked();                      // ... and the office may have locked one of them
    render();
    flushQueue();                       // anything still owed goes now, not in five seconds
    return true;
  } catch (e) {
    trouble(e);
    render();
    return false;
  }
}

/** The ten-second poll: only what moved. A token that has gone stale (Graph
    answers 410 with a resync code) is not a failure - it means read the list
    and start again, which is exactly what readList() does. */
/** Have the floor's lists moved to another site since the last pass? A delta
    token only means anything in the site it was issued in. */
let SITE_GEN = 0;
function siteMoved() {
  const m = CW.stationSiteMoves ? CW.stationSiteMoves() : 0;
  if (m === SITE_GEN) return false;
  SITE_GEN = m;
  return true;
}
async function pollList() {
  /* resolved every pass, which is also what drives the ten-minute look for the
     site the lists are meant to end up in; it is a cached value in between */
  try { SITEID = (await CW.stationSite()) || SITEID; } catch (e) { /* keep the last one */ }
  if (siteMoved()) { TOKEN = null; DELTA_OFF = 0; }
  if (!TOKEN) return readList();
  try {
    const d = await CW.listDelta(ST.STATION_LIST, { siteId: SITEID, fields: ST.STATION_FIELDS, token: TOKEN });
    if (d == null) return readList();
    ITEMS = ST.mergeDelta(ITEMS, d.items);
    if (d.next) TOKEN = d.next;
    READY = true; SOFT = ""; LASTREAD = Date.now();
    rebaseQueue();
    dropBlocked();
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
function tap(id, stage, delta) {
  /* belt and braces: the stepper is drawn disabled for a stage this person
     does not hold, and tap() refuses it anyway - a disabled button is a
     drawing, and this is the rule */
  if (!PERSON || !ST.canStage(PERSON, stage)) return;
  const row = boardNow().find(g => g.id === id);
  if (!row) return;
  /* the office has marked this job's glass finished. Only the office can undo
     that, so there is nothing the floor can do here but see it. Same belt and
     braces as the stage check above: the buttons are drawn disabled too. */
  if (row.officeDone) return;
  const job = row.job;
  /* somebody is working the screen, whether or not the number could move:
     a stepper already at the total is still a hand on the tablet */
  touch();
  const value = ST.applyTap(row, stage, delta);
  if (value == null || value === row[ST.STAGE_ROW[stage]]) return;   // already at the clamp
  /* queued first, drawn second: boardNow() lays the queue over the list, so
     the new number is on screen before the write has left the tablet */
  queueTap(row, job, stage, value);
  render();
  flushQueue();
}

/** The board, with anything this tablet still owes laid over the top of it -
    so a tapped number never flickers back to the list's older one. */
function boardNow() {
  const over = ITEMS.map(it => {
    const q = queuedFor(String(it.id));
    if (!q) return it;
    return { id: it.id, fields: Object.assign({}, it.fields, q) };
  });
  return ST.jobBoard(over);
}

/* ---- drawing ---- */
const agoWords = at => {
  const s = Math.max(0, Math.round((Date.now() - at) / 1000));
  if (s < 10) return "just now";
  if (s < 90) return s + " s ago";
  if (s < 5400) return Math.round(s / 60) + " min ago";
  return Math.round(s / 3600) + " h ago";
};

/* One stage of one job: the name, the count, and the three buttons that move
   it - on the card itself, where a hand can reach them. There is nothing to
   open and nothing to scroll: the owner watched somebody hunt for a stepper
   inside an expanded card with a sheet of glass in their other hand.

   A stage this person does not hold is drawn with its number and its buttons
   disabled, not hidden: the floor can see how far the job has got without
   being able to move somebody else's part of it. */
function stepHtml(g, stage, label) {
  const v = g[ST.STAGE_ROW[stage]];
  /* each stage against its own quantity: cutting, hotmelting and glazing count
     the job's glasses, tuff counts the sheet's own TUFF number */
  const total = Math.max(0, Math.round(Number(g[ST.STAGE_TOTAL_ROW[stage]]) || 0));
  /* the office's lock is not about who is holding the tablet, so it greys every
     stepper on the card rather than the ones a person does not hold */
  const mine = ST.canStage(PERSON, stage) && !g.officeDone;
  const off = mine ? "" : ' disabled aria-disabled="true"';
  const b = (t, act, cls, dead) => '<button class="' + cls + '" data-id="' + esc(g.id) + '" data-stage="' + stage +
    '" data-act="' + act + '"' + (dead ? ' disabled aria-disabled="true"' : off) + '>' + t + '</button>';
  const full = total > 0 && v >= total;
  /* a row with no glasses on it has nothing to finish: All would set it to
     nought and read None a moment later, which says the opposite of the truth */
  const none = !(total > 0);
  return '<div class="step' + (mine ? "" : " locked") + '">' +
    '<span class="stepl">' + esc(label) +
      (g.officeDone ? "" : mine ? "" : ' <span class="nomine">not yours</span>') + '</span>' +
    '<span class="stepc">' + b("&minus;", "-1", "sbtn") +
      '<span class="stepn tab' + (full ? " full" : "") + '">' + v + '</span>' + b("+", "1", "sbtn") +
      b(none || !full ? "All" : "None", full ? "none" : "all", "sall", none) + '</span>' +
    '</div>';
}

/** Everything inside one card. The card element itself is kept between draws
    (see paintBoard), so only this string is ever rebuilt.

    The headline number is what THIS person has left on the job, not how big
    the job is: two people at the same tablet see two different numbers on the
    same card. It is drawn from PERSON, which is not part of the board, so
    paintBoard has to notice PERSON moving on its own - see pState(). */
function cardInner(g) {
  const owed = owedFor(g.id), bad = badFor(g.id);
  const lost = blockedFor(g.job);
  /* Tuff only on the jobs that have any: a "Tuff 0" stepper on every card is a
     fourth row of nothing on a screen somebody reads with a sheet of glass in
     their other hand. */
  const stages = ST.ALL_STAGES.filter(s => s[0] !== ST.TUFF_STAGE || g.tuffTotal > 0);
  return '<div class="chead">' +
      '<span class="cond job">' + esc(g.job) + '</span>' +
      '<span class="cust">' + esc(g.customer || "—") + '</span>' +
      '<span class="cnum tab">' + esc(ST.leftWords(ST.jobLeftFor(g, PERSON && PERSON.stages))) + '</span>' +
    '</div>' +
    '<div class="steps">' + stages.map(s => stepHtml(g, s[0], s[1])).join("") + '</div>' +
    (g.officeDone ? '<div class="officedone">the office has marked this job finished</div>' : "") +
    (lost ? '<div class="unsaved">' + esc(lost.stages.map(s =>
        ST.stageLabel(s.stage) + " " + s.value).join(", ")) +
        ' was not saved — the office marked this job finished first</div>' : "") +
    (bad ? '<div class="unsaved">not saved yet — retrying</div>'
         : owed ? '<div class="saving">saving…</div>' : "");
}

/* ---- the picker -------------------------------------------------------------
   Full screen, one big button per person, and a numeric pad for anybody whose
   PIN column is filled in. A wrong PIN shakes and says "Try again": there is
   no lockout, because locking a shared tablet out of the only screen the floor
   has would stop the work rather than protect it.                           */
function pickerHtml() {
  if (PINFOR) {
    const dots = PINTYPED.replace(/./g, "•");
    const key = t => '<button class="pk" data-pin="' + esc(t) + '">' + esc(t) + '</button>';
    return '<div class="picker">' +
      '<div class="pickh">' + esc(PINFOR.name) + '</div>' +
      '<div class="picksub">Enter your PIN</div>' +
      '<div class="pindots' + (PINBAD ? " shake" : "") + '">' + esc(dots || "····") + '</div>' +
      (PINBAD ? '<div class="pinbad">Try again</div>' : "") +
      '<div class="pinpad">' + ["1", "2", "3", "4", "5", "6", "7", "8", "9"].map(key).join("") +
        '<button class="pk" data-pin="back">←</button>' + key("0") +
        '<button class="pk go" data-pin="ok">OK</button></div>' +
      '<button class="pcancel" data-pin="cancel">Back to the names</button>' +
      '</div>';
  }
  if (!PEOPLE.length)
    return '<div class="picker"><div class="pickh">Who are you?</div>' +
      '<div class="msg">Nobody is set up for the Glass station yet. Ask the office to add ' +
      'you to the “Station people” list.</div></div>';
  return '<div class="picker">' +
    '<div class="pickh">Who are you?</div>' +
    '<div class="pgrid">' + PEOPLE.map(p =>
      '<button class="pbtn" data-person="' + esc(p.name) + '">' +
        '<span class="pname">' + esc(p.name) + '</span>' +
        '<span class="pstages">' + esc(p.stages.length
          ? p.stages.map(ST.stageLabel).join(" · ") : "no stages yet") + '</span>' +
      '</button>').join("") + '</div></div>';
}

function wirePicker(host) {
  host.querySelectorAll("[data-person]").forEach(el => el.onclick = () => {
    const p = PEOPLE.find(x => x.name === el.dataset.person);
    if (!p) return;
    if (ST.pinOk(p, "")) { pickPerson(p); return; }      // no PIN column: straight in
    PINFOR = p; PINTYPED = ""; PINBAD = false; render();
  });
  host.querySelectorAll("[data-pin]").forEach(el => el.onclick = () => {
    const k = el.dataset.pin;
    if (k === "cancel") { PINFOR = null; PINTYPED = ""; PINBAD = false; render(); return; }
    if (k === "back") { PINTYPED = PINTYPED.slice(0, -1); PINBAD = false; render(); return; }
    if (k === "ok") { submitPin(); return; }
    if (PINTYPED.length < 8) PINTYPED += k;
    PINBAD = false;
    /* six digits is the longest PIN the list allows, so a full one submits
       itself rather than making somebody find the OK button with a glove on */
    if (PINFOR && PINTYPED.length >= String(PINFOR.pin).length) submitPin();
    else render();
  });
}
function submitPin() {
  if (!PINFOR) return;
  if (ST.pinOk(PINFOR, PINTYPED)) { pickPerson(PINFOR); return; }
  PINTYPED = ""; PINBAD = true; render();
}

/* ---- the board, drawn once and then patched ---------------------------------
   Six redraws a minute would throw away the scroll position and the tap a
   finger is already on its way to, so the cards are kept as nodes
   keyed by job and only the ones boardDiff names are touched. The first paint
   builds the two groups; nothing after it ever sets the whole board's
   innerHTML again.                                                          */
let NODES = {};                 // job -> the card element
let DOING = null, FINHEAD = null, FIN = null;
let BOARD_PREV = null;          // the board these nodes were drawn from
let QSIG = {};                  // job -> what this tablet owed on it when it was last drawn
let PSIG = "";                  // who the cards were drawn for, and with which stages
let WIRED = false;

/* The "saving..." and "not saved yet" lines come from this tablet's own queue,
   which is not in the list and so is invisible to boardDiff. A row that has
   just failed to save therefore has to be named here, or the red line never
   appears on a board whose numbers did not move. */
function qState(g) {
  const lost = blockedFor(g.job);
  return (owedFor(g.id) ? "o" : "") + (badFor(g.id) ? "b" : "") +
         (lost ? "!" + lost.stages.map(s => s.stage + s.value).join(",") : "");
}

/* The person is not in the list either, and every card is drawn from them
   twice over: the headline number is their own remaining work, and a stage
   they do not hold is greyed. Switching person goes through the picker, which
   lets every node go anyway - but readPeople() re-reads the stages of whoever
   is signed in on the ten-second clock, so an edit to their Stages column in
   SharePoint changes what every card should say with no picker in between.
   Without this the cards would keep the numbers of the stages they used to
   hold until something else happened to move them. */
function pState() {
  return PERSON ? PERSON.name + "|" + (PERSON.stages || []).join(",") : "";
}

function makeCard(g) {
  const el = document.createElement("div");
  el.className = "card" + (g.finished ? " done" : "");
  if (el.dataset) el.dataset.job = g.job;
  el.setAttribute("data-job", g.job);
  el.innerHTML = cardInner(g);
  return el;
}
function dressCard(el, g) {
  el.className = "card" + (g.finished ? " done" : "");
  el.innerHTML = cardInner(g);
}
function paintBoard(host, board) {
  if (!DOING) {
    host.innerHTML = "";
    DOING = document.createElement("div"); DOING.className = "grp";
    FINHEAD = document.createElement("div"); FINHEAD.className = "grouphead";
    /* wired here rather than in wireBoard: this node is rebuilt whenever a
       message has taken the board's place, and it must come back live */
    if (FINHEAD.addEventListener)
      FINHEAD.addEventListener("click", () => { FINOPEN = !FINOPEN; render(); });
    FIN = document.createElement("div"); FIN.className = "grp fin";
    host.appendChild(DOING); host.appendChild(FINHEAD); host.appendChild(FIN);
    NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
  }
  const diff = ST.boardDiff(BOARD_PREV, board);
  const byJob = {};
  board.forEach(g => { byJob[g.job] = g; });

  diff.removed.forEach(j => {
    const el = NODES[j];
    if (el && el.remove) el.remove();
    delete NODES[j];
    delete QSIG[j];
  });
  const changed = diff.changed.slice();
  /* the person changing is every card changing, because every card is drawn
     from them - see pState() */
  const psig = pState(), pmoved = PSIG !== psig;
  PSIG = psig;
  board.forEach(g => {
    const sig = qState(g);
    if ((pmoved || QSIG[g.job] !== sig) && changed.indexOf(g.job) < 0 && diff.added.indexOf(g.job) < 0)
      changed.push(g.job);
    QSIG[g.job] = sig;
  });
  changed.forEach(j => { if (NODES[j]) dressCard(NODES[j], byJob[j]); });
  diff.added.forEach(j => { NODES[j] = makeCard(byJob[j]); });

  /* the order, and which group a card sits in, only get touched when something
     actually moved - appendChild on a node already in place is a move */
  if (diff.order || diff.added.length || diff.removed.length) {
    board.forEach(g => {
      const el = NODES[g.job];
      if (!el) return;
      (g.finished ? FIN : DOING).appendChild(el);
    });
  }
  const done = board.filter(g => g.finished).length;
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
  if (hdr) hdr.textContent = PERSON
    ? PERSON.name + (PERSON.stages.length ? " · " + PERSON.stages.map(ST.stageLabel).join(", ") : " · no stages")
    : "";
  const sw = $("#switchbtn");
  if (sw) { sw.hidden = !PERSON; sw.style.display = PERSON ? "" : "none"; }
  /* the search box belongs to the board: there is nothing to search on the
     picker, and a box over an error message only looks broken */
  const sb = $("#search");
  const boarding = !PROBLEM && PEOPLE_READ && !!PERSON && READY;
  if (sb) { sb.hidden = !boarding; sb.style.display = boarding ? "" : "none"; }
  /* the board as it stands, before the box has narrowed it: the number beside
     the box is read off this, and the cards below off the filtered copy */
  const live = boarding ? boardNow() : null;
  /* everything the person signed in has left to do, added up over the whole
     board: the same number as the cards, so it counts down with them. It is
     next to the box that narrows the cards and deliberately not narrowed by
     it - a job number typed in must not make somebody's day look shorter. It
     comes and goes with the search box for the same reason: a total floating
     over "ask the office for the permission" is a number about nothing. */
  const gt = $("#gtotal");
  if (gt) {
    gt.hidden = !boarding;
    gt.style.display = boarding ? "" : "none";
    gt.textContent = boarding ? ST.leftWords(ST.boardLeftFor(live, PERSON.stages)) : "";
  }
  const upd = $("#upd");
  if (upd) upd.textContent = LASTREAD ? "updated " + agoWords(LASTREAD) : "";
  /* the passing-failure line is part of the page, not of the board, so it is
     right whichever of the three things below is on screen */
  const soft = $("#soft");
  if (soft) { soft.textContent = SOFT; soft.hidden = !SOFT; soft.style.display = SOFT ? "" : "none"; }

  /* Who are you? comes before the board and after any real problem with it:
     there is no point asking a name on a screen that cannot reach the list,
     and no point drawing steppers before anybody has said whose they are. */
  if (PROBLEM || !PEOPLE_READ || !PERSON || !READY) {
    /* a message or the picker takes the board's place, so the card nodes are
       let go: the next good read paints them fresh */
    DOING = null; FIN = null; FINHEAD = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    if (!PROBLEM && PEOPLE_READ && !PERSON) { host.innerHTML = pickerHtml(); wirePicker(host); return; }
    host.innerHTML = '<div class="msg">' + esc(words()) + againHtml() + '</div>';
    wireAgain();
    return;
  }

  /* the search box narrows the board and never becomes it: an empty box is
     every card, and a box nothing matches says so rather than looking broken */
  const board = ST.boardFilter(live, QUERY);
  if (!board.length) {
    DOING = null; FIN = null; FINHEAD = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    host.innerHTML = '<div class="msg">' +
      (QUERY ? "No job on the board matches “" + esc(QUERY) + "”." : "Nothing on the board yet.") +
      '</div>';
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
    ? "The “Glass station” list is not in the floor’s site yet. Ask the office to add it."
    : PROBLEM === "site"
    ? "The floor’s SharePoint site is not there yet, or this account cannot see it. Ask the office."
    : "Reading the board…";
}
/* Try again is offered whenever the board cannot be shown - not only when the
   problem has a name. A tablet stuck on "Reading the board..." because one
   read failed before anything was set has to have something to tap. */
const againHtml = () => PROBLEM === "reauth"
  ? '<div><button class="again" id="reauth">Sign in again</button></div>'
  : '<div><button class="again" id="again">Try again</button></div>';
function wireAgain() {
  const a = $("#again");
  /* "Try again" forgets the cached site id as well: the commonest reason the
     board will not come back is an id that no longer means anything */
  if (a) a.onclick = () => {
    PROBLEM = ""; SOFT = ""; SITEID = null; TOKEN = null;
    if (CW.forgetStationSite) CW.forgetStationSite(true);   // a person asked: look now
    render(); readPeople(); readList();
  };
  const r = $("#reauth");
  if (r) r.onclick = async () => {
    try { await CW.signIn(CW.LIST_SCOPES); PROBLEM = ""; SITEID = null; TOKEN = null; await start(); }
    catch (e) { console.warn("[station] sign-in again failed:", (e && e.message) || e); }
  };
}

/** One delegated listener for the whole board, wired once. Cards come and go
    under it; the listener does not, so nothing has to be re-wired on a card
    that has just been redrawn. */
function wireBoard(host) {
  if (WIRED || !host.addEventListener) return;
  WIRED = true;
  host.addEventListener("click", ev => {
    /* anything at all on the board is a hand on the tablet, including a button
       that will not move because it is somebody else's stage or already at the
       total: the ten-minute lock must not fire under a working hand */
    touch();
    let el = ev.target;
    for (let i = 0; el && i < 5; i++) {
      const d = el.dataset || {};
      if (d.act) {
        ev.stopPropagation();
        if (el.disabled) return;
        tap(d.id, d.stage, d.act === "all" || d.act === "none" ? d.act : Number(d.act));
        return;
      }
      el = el.parentElement;
    }
  });
}

/* ---- staying current -------------------------------------------------------
   This page is opened once and left on a tablet for weeks, which is exactly
   where a browser serving a script from cache does the most damage. The build
   id is stamped into the script URLs by build.py, so a changed version.json
   means the scripts on this tablet are old. It reloads itself - but only when
   it owes nothing, because a reload with a queued tap in the balance would put
   the write off until whoever is on the station next taps something. */
const BUILD_MS = 120000;
let BUILD_NOW = null;
async function checkBuild() {
  try {
    const r = await fetch("version.json?t=" + Date.now(), { cache: "no-store" });
    if (!r.ok) return;
    const latest = (await r.json()).build;
    if (!latest) return;
    if (!BUILD_NOW) { BUILD_NOW = latest; return; }             // the first look is the baseline
    if (latest !== BUILD_NOW && !Object.keys(QUEUE).length && !Object.keys(LOGQ).length)
      location.reload(true);
  } catch (e) { /* offline or blocked - the board is what matters */ }
}

/** One turn of the ten-second clock. A named function rather than a closure
    inside setInterval, so a test can take exactly one turn of it. */
async function tickOnce() {
  render();
  /* a people list that has not answered yet is retried on the same clock: one
     failed read at start-up must not leave the tablet stuck at the picker's
     door until somebody thinks to reload it */
  if (!PEOPLE_READ) await readPeople();
  await pollList();
}

/* ---- the page ---- */
function showGate(on, err) {
  const g = $("#gate");
  g.hidden = !on; g.style.display = on ? "flex" : "none";
  $("#top").hidden = on; $("#top").style.display = on ? "none" : "flex";
  $("#main").hidden = on; $("#main").style.display = on ? "none" : "block";
  const e = $("#gateerr");
  e.style.display = err ? "block" : "none";
  e.textContent = err || "";
}

async function start() {
  showGate(false);
  applyTheme(themeNow());
  const tb = $("#themebtn");
  if (tb) tb.onclick = () => applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  $("#outbtn").onclick = () => { if (confirm("Sign out of the Glass station?")) CW.signOut(); };
  const sw = $("#switchbtn");
  if (sw) sw.onclick = () => switchPerson();
  /* the box is in the header, outside #board, so typing in it never rebuilds
     the node the caret is in - only the cards under it are redrawn */
  const sb = $("#search");
  if (sb) sb.oninput = () => { QUERY = sb.value || ""; touch(); render(); };
  render();
  await readPeople();
  await readList();
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
  applyTheme(themeNow());
  $("#signinbtn").onclick = async () => {
    try { await CW.signIn(CW.LIST_SCOPES); await start(); }
    catch (e) { showGate(true, "Sign-in failed:\n" + ((e && e.message) || e)); }
  };
  try {
    const acct = await CW.initAuth();
    if (acct) await start(); else showGate(true);
  } catch (e) {
    showGate(true, "Startup problem:\n" + ((e && e.message) || e));
  }
})();
