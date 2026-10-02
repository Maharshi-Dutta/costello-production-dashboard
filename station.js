/* The Glass station page - the floor's own screen.

   ONE PAGE, TWO TABLETS since 2026-09-21 (docs/specs/2026-09-21-glass-split-
   no-glazing.md). Cutting and hotmelting each have a tablet of their own and
   this file serves both: the stage comes from `?stage=cut` / `?stage=hotmelt`,
   else from the device's own `cw_stationstage`, else from a two-button chooser
   shown before the person picker. Everything the page draws - the header, who
   may sign in, which steppers a card has, what "N left" counts and which cards
   sink to the bottom - is about THIS page's stage and nothing else. Glazing
   left the station in the same ship and is not a stage here any more.

   What this page can do, in full: read three SharePoint lists ("Glass
   station", "Station people", "Station log"), PATCH the Cut / Hotmelt / Tuff
   counters of a Glass station row - with that stage's By and At and the
   last-touch pair beside them - and POST one line to the Station log for each
   counter write that succeeded. That is all. There is no workbook here - no
   exceljs, no parser.js, no download, no Excel API path anywhere in this file
   - no delete, no editing of a job's facts, no comments, no export, and no
   link back to the master dashboard.

   It is used on a shared tablet, signed in once with a shared station account
   and then passed between three people for weeks. So: the person picks their
   name (and their PIN) before they can move anything, the name stays until
   somebody presses Switch (since 2026-10-02 there is no idle lock: every tap
   is logged under whoever was picked last), the stages somebody does not hold are greyed rather
   than hidden, every tap shows on screen at once and is owed rather than
   awaited, the owed writes survive a reload in localStorage, and nothing on
   this page can open a modal or a consent popup while somebody is holding a
   sheet of glass.                                                           */

const $ = s => document.querySelector(s);
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

const PERSON_KEY = "cw_person";          // who is on the station right now, and when they last tapped
const CUTFIRST_KEY = "cw_glasscutfirst"; // the hotmelting tablet's "Cut first" switch, this device's own
const STAGE_KEY = "cw_stationstage";     // which stage THIS tablet is - the device's own, not the person's
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
let WAITONLY = false;           // the "N waiting on glass" capsule's filter; never remembered
/* the two tabs, On floor and Finished (owner, 2026-09-24; replaced the
   collapsed Finished group). Remembered on the device. A search may move the
   board to the other tab; PRETAB is the tab it was on before the typing began,
   and clearing the box goes back to it. */
const TAB_KEY = "cw_glasstab";
let TAB = "floor";
let PRETAB = null;
try { if (localStorage.getItem(TAB_KEY) === "finished") TAB = "finished"; } catch (e) {}
/* the waiting filter (2026-10-01) shows only the jobs fabrication has finished
   and this person has not. They are all On floor, so turning it on from
   Finished goes there; WAITTAB is the tab to go back to when it is turned off,
   and it is what gets saved meanwhile - the jump itself is never remembered.
   A tab the person taps while the filter is on is their own choice and stays. */
let WAITTAB = null;
const saveTab = () => { try { localStorage.setItem(TAB_KEY, WAITTAB || TAB); } catch (e) {} };
function setWait(on) {
  if (!!on === WAITONLY) return;
  WAITONLY = !!on;
  if (WAITONLY) { if (TAB !== "floor") { WAITTAB = TAB; TAB = "floor"; } }
  else { if (WAITTAB) TAB = WAITTAB; WAITTAB = null; }
}
/* "Cut first" (2026-10-02), on the hotmelting tablet only: the jobs whose
   cutting is ahead of their hotmelting come first on On floor. Off unless
   somebody turns it on; remembered on the device, like the tab. */
let CUTFIRST = false;
try { CUTFIRST = localStorage.getItem(CUTFIRST_KEY) === "1"; } catch (e) {}
let PERSON = null;              // the person who picked their name
let LAST_TAP = 0;               // when they last touched anything
let PINFOR = null;              // the person whose PIN is being asked for
let PINTYPED = "";
let PINBAD = false;
let retryT = null, buildT = null, peopleT = null;
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

/* ---- which tablet this is ---------------------------------------------------
   Three physical tablets, one page. The stage is a property of the DEVICE, like
   the theme, so it lives in localStorage - but the URL wins, so a tablet can be
   set up once by opening `glass.html?stage=hotmelt` and never asked again. A
   device that has been told neither is asked, once, on screen.

   A word that is not one of this station's own stages is not a stage, and
   ONLY A VALID URL VALUE WINS (review finding R4, 2026-09-21): a typo in the
   URL - `?stage=hotmlet` on a kiosk bookmark - falls back to the stage the
   device already knows it is, and only a tablet that has never been told
   anything is asked. Taking the typo as "told nothing at all" would have put a
   working tablet in front of the chooser and signed its person out, which is
   the worst of the three answers. */
let PAGE_STAGE = "";
/** Is this word one of this station's stages? */
const knownStage = k => ST.STAGE_KEYS.indexOf(String(k == null ? "" : k).trim().toLowerCase()) >= 0
  ? String(k).trim().toLowerCase() : "";
function setStage(k) {
  PAGE_STAGE = knownStage(k);
  if (PAGE_STAGE) { try { localStorage.setItem(STAGE_KEY, PAGE_STAGE); } catch (e) {} }
}
function stageFromUrl() {
  try {
    const m = /[?&]stage=([^&#]*)/.exec(String((typeof location !== "undefined" && location.search) || ""));
    return m ? decodeURIComponent(m[1]) : "";
  } catch (e) { return ""; }
}
function stageFromDevice() {
  try { return localStorage.getItem(STAGE_KEY) || ""; } catch (e) { return ""; }
}
/* the URL only wins when it says something this station understands */
setStage(knownStage(stageFromUrl()) || stageFromDevice());
/** Does THIS page count tuff? The cutter taps it, and only the cutter. */
const tuffHere = () => PAGE_STAGE === "cut";
/** The stages of one person that this page draws: its own, and tuff on the
    cutting page. A person's other stages belong to another tablet and are not
    shown here even when their row holds them. */
function myStages(p) {
  return ST.heldStages((p && p.stages) || [])
           .filter(k => k === PAGE_STAGE || (tuffHere() && k === ST.TUFF_STAGE));
}
/** Who may sign in on this tablet: the Glass people who hold something it
    draws. Somebody who only hotmelts is not offered the cutting tablet. */
const pagePeople = () => PEOPLE.filter(p => myStages(p).length > 0);

/* ---- who is on the station --------------------------------------------------
   The name is not a login: it says which stages the steppers will move and it
   is what goes into the log. THERE IS NO IDLE LOCK (owner, 2026-10-02): the
   name stays, across reloads and new builds, until somebody presses Switch,
   the tablet is made the other stage, or the person leaves `Station people`.
   What that costs is said in the brief: a tap is logged under whoever was
   picked last. */
function loadPerson() {
  let saved = null;
  try { saved = JSON.parse(localStorage.getItem(PERSON_KEY) || "null"); } catch (e) { saved = null; }
  if (!saved || !saved.name) return;
  LAST_TAP = Math.min(Number(saved.at) || 0, Date.now());
  /* the stages come from the list, never from storage: an edit to
     localStorage must not be able to hand somebody a stage they do not hold -
     and a name that holds nothing THIS page draws is not signed in here */
  const hit = pagePeople().find(p => p.name === saved.name);
  PERSON = hit || null;
}
function savePerson() {
  try {
    if (PERSON) localStorage.setItem(PERSON_KEY, JSON.stringify({ name: PERSON.name, at: LAST_TAP }));
    else localStorage.removeItem(PERSON_KEY);
  } catch (e) {}
}
/** Somebody is working the screen: the stamp kept beside the name. */
function touch() { LAST_TAP = Date.now(); savePerson(); }
/** The box belongs to whoever was just holding the tablet: the next person must
    not be handed a board narrowed - or a tab moved - by somebody else's search.
    Called on a switch and on a pick. */
function clearSearch() {
  QUERY = "";
  if (PRETAB) { TAB = PRETAB; saveTab(); }
  PRETAB = null;
  setWait(false);
  const sb = $("#search");
  if (sb) sb.value = "";
}
function pickPerson(p) {
  PERSON = p; PINFOR = null; PINTYPED = ""; PINBAD = false; clearSearch(); touch(); render();
}
function switchPerson() {
  PERSON = null; PINFOR = null; PINTYPED = ""; PINBAD = false; LAST_TAP = 0;
  clearSearch();
  /* AND NEITHER IS THE DAY SHEET (review, 2026-09-21). It was left open across
     a change of person, so the next name to be picked was shown the last
     person's half-typed numbers, on a form whose Save would have filed them
     under the new name. The typing is not lost - it is in storage under whose
     it is, and comes back when they pick their name again. */
  DAY.reset();
  savePerson(); render();
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
  DAY.persist();
}
/** Does this tablet owe anything at all? What keeps the retry timer armed. */
const owingAnything = () => !!(Object.keys(QUEUE).length || Object.keys(LOGQ).length ||
                               DAY.owing());
/** ... and what a RELOAD would actually put off (review, 2026-09-21).
    Deliberately NOT the day sheets. Every one of these three queues is in
    localStorage and rebuilt on the next load, so none of them is lost by a
    reload - but a day sheet the list keeps refusing would otherwise pin the
    tablet on an old build for as long as it sat there, which is exactly the
    state a new build is most likely to be the fix for. The counter and log
    queues keep the guard they have always had: they are minutes of somebody's
    tapping, and a reload delays them until the next tap. */
const owingWrites = () => !!(Object.keys(QUEUE).length || Object.keys(LOGQ).length);

/* ---- the office's lock ------------------------------------------------------
   The office can mark a job's GLASS finished, and that job's glass stages then
   go read-only here: the steppers grey out and tap() refuses them, the same way
   a stage somebody does not hold is already refused. Only the office can take
   it off again. The tablet learns it from the OfficeDone column and from
   nothing else - it still cannot see the workbook.

   TUFF IS OUTSIDE IT (review finding R2, 2026-09-21). `OfficeDone` is
   ST.officeComplete of the job's DG and TG - it says the office has ticked the
   GLASS off, and it has never said anything about tuff, which is a different
   department counting a different quantity. That distinction did no harm while
   the lock landed only after glazing, by which time the tuff was long counted;
   since 2026-09-21 the lock lands the moment hotmelting finishes, so a locked
   job is routinely one the cutter is still counting tuff on. So the tuff
   stepper stays live under the lock, tap() lets it through, and dropBlocked()
   never takes a queued tuff tap away.

   The awkward case is a GLASS tap already sitting in the queue when the lock arrives:
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
    /* the lock is about the job's GLASS, so a tuff tap is not its business:
       dropping one would delete work the cutter did while the office was
       ticking the glass off (R2) */
    if (e.stage === ST.TUFF_STAGE) return;
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

    Only a rise re-bases. A counter that has gone DOWN in the list was moved by
    somebody else - the office clearing the job, or a correction on another
    tablet - and the queue holds an ABSOLUTE number, so sending it flat would
    put the whole count back rather than lay this tap on top of it. Which of
    the two statements is the true one is decided the way everything else about
    this row is: whoever moved it LAST wins, on the stamps both sides write.
    Every writer sets DoneAt when it moves a row - the floor on every tap, and
    an office clear precisely so that this comparison stays honest - so a row
    stamped AFTER this tap was made is the later word and the tap is dropped.
    It is not stranded the way a tap that meets the office's LOCK is: nothing
    is stopping the floor from tapping again, because a clear UNLOCKS the job.
    A tie, or a stamp that will not parse, leaves the tap alone: the floor's own
    statement is what this tablet is for. */
function rebaseQueue() {
  let moved = false;
  Object.keys(QUEUE).forEach(k => {
    const e = QUEUE[k];
    const it = ITEMS.find(x => String(x.id) === String(e.id));
    if (!it) return;
    const f = it.fields || {};
    const now = Number(f[ST.STAGE_FIELD[e.stage]]);
    const was = Number(e.from);
    if (!isFinite(now) || !isFinite(was)) return;
    if (now < was) {
      const rowAt = Date.parse(String(f.DoneAt == null ? "" : f.DoneAt));
      const tapAt = Date.parse(String(e.at || ""));
      if (!isFinite(rowAt) || !isFinite(tapAt) || rowAt <= tapAt) return;
      console.warn("[station] dropping a queued tap for " + e.job + ": " +
        ST.stageLabel(e.stage) + " " + e.value + " was tapped at " + e.at +
        ", and the row was moved after that (" + f.DoneAt + ") - tap it again if it is still right");
      delete QUEUE[k];
      moved = true;
      return;
    }
    if (now <= was) return;
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
  if (!owingAnything()) return;
  /* no site resolved means no list to write to. Coming back in five seconds is
     the whole of the answer: guessing one would send the tablet at the
     workbook, which is the one thing it must never do. */
  if (!READY || !SITEID) { armRetry(); return; }
  /* again here, not only after a read: the lock can have arrived in the poll
     that ran while this queue was waiting for its five seconds */
  dropBlocked();
  if (!owingAnything()) { render(); return; }
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
        if (CW.isMissing && CW.isMissing(err) && CW.forgetStationSite) CW.forgetStationSite(false, ST.GLASS.site);
        console.warn("[station] counter write failed for item " + id + ":", (err && err.message) || err);
      }
      saveQueue();
    }
    await flushLog();
    /* last, and on its own clock of failures: a day sheet that will not send
       must never hold up a counter, and a counter that will not send must
       never lose somebody's day sheet */
    await DAY.flush();
  } finally {
    flushing = false;
  }
  render();
  if (wrote) await pollList();                 // the list agrees now: drop back to it
  if (owingAnything()) armRetry();
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
const commentOpts = () => ({ siteId: SITEID, fields: ST.COMMENT_FIELDS });

/* ---- a word for the office -------------------------------------------------
   The note channel, shipped 2026-09-15 (docs/specs/2026-09-15-station-comments.md).
   Every line of it - the composer, the thread, the row and the two list calls -
   is in station-core.js, and THIS IS THE ONE LINE A SECOND STATION PAGE WOULD
   CHANGE: `station` is the only station-specific input the channel takes. A
   cutting.html passes "Cutting" here and has the feature, with no new list, no
   new column and nothing to add to the office's drawer.

   It writes `Station comments` the same way the log lines are written - one
   POST from the tablet, straight to the list, no feeder, no upsert, no delete -
   and like every other thing on this page it never goes near the workbook. */
const NOTES = ST.stationComments({
  station: ST.STATION_NAME,
  listItems: (name, o) => CW.listItems(name, o),
  listAdd: (name, fields, o) => CW.listAdd(name, fields, o),
  opts: commentOpts
});

/* ---- the end-of-day sheet ---------------------------------------------------
   Shipped 2026-09-21 (docs/specs/2026-09-21-day-sheets-and-station-reports.md).
   The cutter's paper sheet, on the tablet: their name, the day, the counts the
   STATION DEFINITION names for this stage, a line about anything that got in
   the way, and this week measured against the office's target.

   IT IS SWITCHED ON BY THE DEFINITION, not by this page knowing what cutting
   is: ST.daySheetOf(ST.GLASS, PAGE_STAGE) answers with the counts or with
   nothing, and a stage with nothing gets no button, reads neither list and
   sends no request for either. That is the whole of "built so another page can
   switch it on later".

   Once saved it cannot be changed from here (owner's decision 2): the row is
   append-only from the tablet - one POST, no PATCH, no DELETE, ever - and a
   mistake is the office's to correct. The two lists are read-only apart from
   that one POST, and `Station targets` is never written here at all.

   The sheet itself - draft, save, queue, drawing - moved to station-ui.js on
   2026-09-28 (STU.stuDaySheet) so the welding page has it without a copy.
   Since then the HOTMELTING tablet has one too: DG and TG units, typed, beside
   the tablet's own count of the person's hotmelt taps today (`Counted`), and
   no weekly target (docs/specs/2026-09-28-day-sheets-welding-hotmelt-floor-log.md).
   The two storage keys are unchanged, so a sheet owed across the upgrade is
   still owed after it. */
const DAY = STU.stuDaySheet({
  def: ST.GLASS,
  stage: () => PAGE_STAGE,
  who: () => who(),
  siteId: () => SITEID,
  touch: () => touch(),
  render: () => render(),
  flush: () => flushQueue(),
  keys: { q: "cw_stationdayq", draft: "cw_daysheetdraft" },
  stageWords: () => stageWords(),
  /* this tablet's own lines not yet on the list are still today's taps */
  pendingLog: () => Object.keys(LOGQ).map(k => LOGQ[k].fields),
  tag: "[station]"
});


/* ---- what fabrication has done on each job (2026-10-01) ---------------------
   docs/specs/2026-10-01-glass-sees-fabrication.md. The glass tablet READS the
   `Fabrication station` list, and that is all it does with it: one delta read
   on the board's own clock, no write of any kind.

   IT IS KEPT APART FROM THE GLASS LISTS. That list is in the `Floor stations`
   site and the glass lists are pinned to the workbook's own (ST.GLASS.site).
   Since 2026-10-02 the read itself is the shared reader every station page
   uses (STU.stuListReader in station-ui.js): it asks CW.stationSite("floor") -
   graph.js's other pinned channel, with its own cached id, miss clock and move
   counter - and keeps what it answers in GFAB, never in SITEID. Its token is
   GFAB.token, never TOKEN. Nothing here calls trouble(), sets PROBLEM or SOFT,
   or forgets the glass site: a fabrication list that cannot be read costs the
   cards one muted line and the glass page nothing else.

   A read that fails is not tried again for a minute. A page loaded without
   fabrication-core.js has no such feature and asks nothing. */
const GFAB_FIELDS = ["Title", "Job", "Group", "GroupSeq", "Seq", "Section", "Active", "OnSheet",
                     "Frames", "Sashes", "Transoms", "FramesDone", "SashesDone", "TransomsDone"];
const GFAB = typeof FABC === "undefined" ? null
  : STU.stuListReader({ site: FABC.FB_SITE, list: FABC.FB_LIST, fields: GFAB_FIELDS });
let GFABMAP = {};               // job -> ST.fabOfJob's answer, from the last read
let GFABSIG = "";               // ... and what the cards were last told, to say when it moved
/** Read what moved. Answers true when what the cards draw has changed. Never
    throws, and never touches anything the glass lists use. */
async function readFab() {
  if (!GFAB || !(await GFAB.read())) return false;
  const map = {};
  if (GFAB.items) FABC.fbOfficeBoard(GFAB.items).forEach(c => { map[c.job] = ST.fabOfJob(c.groups); });
  const sig = JSON.stringify([!!GFAB.items, map]);
  const moved = sig !== GFABSIG;
  GFABMAP = map; GFABSIG = sig;
  return moved;
}
/** One job's fabrication: null (no such feature on this page, or the first
    read has not answered yet - no line at all), { state: "off" } (the list
    cannot be read), or ST.fabOfJob's answer. */
function fabOf(job) {
  if (!GFAB || !GFAB.ready) return null;
  if (!GFAB.items) return { state: "off" };
  return GFABMAP[FABC.fbKey(job)] || { state: "none" };
}
/** Fabrication done, and the person signed in still has glass to do on it. */
const waitingOn = g => !!PERSON && ST.glassWaiting(g, fabOf(g.job), myStages(PERSON));
/** The search box and the waiting filter, together. */
function shownCards(cards, q) {
  const b = ST.boardFilter(cards, q);
  return WAITONLY ? b.filter(waitingOn) : b;
}
/** The one line under a card's counters. */
function fabLineHtml(g) {
  const f = fabOf(g.job);
  if (!f || f.state === "none") return "";
  if (f.state === "off") return '<div class="fabl">Fabrication: not available</div>';
  const words = f.state === "done" ? "done" : f.state === "progress" ? "in progress" : "not started";
  const parts = f.state === "notstarted" ? [] : FABC.FB_PARTS.filter(k => f.parts[k])
    .map(k => FABC.FB_PART_LABEL[k] + " " + f.parts[k].done + "/" + f.parts[k].total);
  return '<div class="fabl f-' + (f.state === "done" && f.sheet ? "sheet" : f.state) + '">' +
    esc(["Fabrication: " + words].concat(parts).join(" · ")) + '</div>';
}

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
    if (CW.forgetStationSite) CW.forgetStationSite(false, ST.GLASS.site);
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
    if (!SITEID) SITEID = await CW.stationSite(ST.GLASS.site);
    if (!SITEID) { PROBLEM = "site"; SOFT = ""; READY = false; render(); return false; }
    const items = await CW.listItems(ST.PEOPLE_LIST, peopleOpts());
    if (items == null) { PROBLEM = "people"; SOFT = ""; READY = false; render(); return false; }
    PEOPLE = ST.stationPeople(items, ST.STATION_NAME);
    PEOPLE_READ = true;
    if (PROBLEM === "people") PROBLEM = "";
    /* the stages a chosen person holds are re-read with them: a change in
       SharePoint reaches the tablet without anybody signing out */
    if (PERSON) {
      const hit = pagePeople().find(p => p.name === PERSON.name);
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
    if (!SITEID) SITEID = await CW.stationSite(ST.GLASS.site);
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

/** The poll, on TICK's clock: only what moved. A token that has gone stale (Graph
    answers 410 with a resync code) is not a failure - it means read the list
    and start again, which is exactly what readList() does. */
/** Have the floor's lists moved to another site since the last pass? A delta
    token only means anything in the site it was issued in. */
let SITE_GEN = 0;
function siteMoved() {
  const m = CW.stationSiteMoves ? CW.stationSiteMoves(ST.GLASS.site) : 0;
  if (m === SITE_GEN) return false;
  SITE_GEN = m;
  return true;
}
/* ONE POLL IN THE AIR AT A TIME: the clock's turn and the look a landed write
   takes can ask at the same moment, and two deltas on one token would each
   merge and each hand back a token. The second asker is told false and the
   poll already running answers for both. */
let polling = false;
async function pollList() {
  if (polling) return false;
  polling = true;
  try { return await pollListNow(); } finally { polling = false; }
}
async function pollListNow() {
  /* resolved every pass, which is also what drives the ten-minute look for the
     site the lists are meant to end up in; it is a cached value in between */
  try { SITEID = (await CW.stationSite(ST.GLASS.site)) || SITEID; } catch (e) { /* keep the last one */ }
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
     drawing, and this is the rule. Since 2026-09-21 a stage this PAGE does not
     draw is refused for the same reason: the other tablet's stage is not this
     one's to move, however the click got here. */
  if (!PERSON || !ST.canStage(PERSON, stage)) return;
  if (stage !== PAGE_STAGE && !(tuffHere() && stage === ST.TUFF_STAGE)) return;
  /* either tab: a Finished card is tapped exactly as it always was, unless the
     office has locked it (owner's decision 2, 2026-09-24) */
  const tabs = tabsNow();
  const row = tabs.floor.concat(tabs.finished).find(g => g.id === id);
  if (!row) return;
  /* the office has marked this job's GLASS finished. Only the office can undo
     that, so there is nothing the floor can do here but see it. Same belt and
     braces as the stage check above: the buttons are drawn disabled too. Tuff
     is not glass and is not locked with it (R2). */
  if (row.officeDone && stage !== ST.TUFF_STAGE) return;
  const job = row.job;
  /* somebody is working the screen, whether or not the number could move:
     a stepper already at the total is still a hand on the tablet */
  touch();
  const value = ST.applyTap(row, stage, delta);
  if (value == null || value === row[ST.STAGE_ROW[stage]]) return;   // already at the clamp
  /* queued first, drawn second: boardNow() lays the queue over the list, so
     the new number is on screen before the write has left the tablet */
  queueTap(row, job, stage, value);
  TICK.burst();
  render();
  flushQueue();
}

/** The board, with anything this tablet still owes laid over the top of it -
    so a tapped number never flickers back to the list's older one.

    `finished` is re-read for THIS PAGE (2026-09-21): a card is done here when
    this page's stage is complete, so the cutter's finished jobs sink out of the
    cutter's way whether or not hotmelting has started. The record's own
    `finished` - both stages, and the tuff - is the office's answer and is what
    the office's board and the job row keep reading. Everything below (the
    card's class, the Finished group, boardDiff's "the order moved") follows
    this one field, so this is the whole of the change.

    AND THE CUTTER'S TUFF COUNTS HERE (owner, after the demo on 2026-09-21).
    Tuff is the cutting bench's own work, on the cutting tablet's own card, so a
    job with tuff still to count has not left the cutter's way however much
    glass is cut. The hotmelting page is untouched by it: tuff is not that
    bench's work and never appears on it. */
/* Since 2026-09-24 the tabs are the primary reading (tabsNow): On floor /
   Finished (ST.glassTabs), whose cards include the jobs still on the sheet in
   a later section. boardNow() is the In production cards (Active) among them,
   which the header counts. */
function tabsNow() {
  const over = ITEMS.map(it => {
    const q = queuedFor(String(it.id));
    if (!q) return it;
    return { id: it.id, fields: Object.assign({}, it.fields, q) };
  });
  return ST.glassTabs(over, g =>
    ST.stageComplete(g, PAGE_STAGE) && !(tuffHere() && ST.tuffOwed(g)));
}
const activeOf = tabs => tabs.floor.concat(tabs.finished).filter(g => g.active)
  .sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
function boardNow() { return activeOf(tabsNow()); }

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
  /* each stage against its own quantity: cutting and hotmelting count the
     job's glasses, tuff counts the sheet's own TUFF number */
  const total = Math.max(0, Math.round(Number(g[ST.STAGE_TOTAL_ROW[stage]]) || 0));
  /* the office's lock is not about who is holding the tablet, so it greys the
     glass steppers on the card rather than the ones a person does not hold -
     but it says nothing about TUFF, which stays live under it (R2) */
  const holds = ST.canStage(PERSON, stage);
  const locked = g.officeDone && stage !== ST.TUFF_STAGE;
  const mine = holds && !locked;
  const off = mine ? "" : ' disabled aria-disabled="true"';
  const b = (t, act, cls, dead) => '<button class="' + cls + '" data-id="' + esc(g.id) + '" data-stage="' + stage +
    '" data-act="' + act + '"' + (dead ? ' disabled aria-disabled="true"' : off) + '>' + t + '</button>';
  const full = total > 0 && v >= total;
  /* a row with no glasses on it has nothing to finish: All would set it to
     nought and read None a moment later, which says the opposite of the truth */
  const none = !(total > 0);
  return '<div class="step' + (mine ? "" : " locked") + '">' +
    '<span class="stepl">' + esc(label) +
      /* what is still to do on THIS stage of THIS job, under the name of the
         stage and beside the buttons that move it. It is the job's number, not
         a personal tally: two people who both cut read the same one here. It
         is drawn only for a stage this person holds - a number against
         somebody else's stage is not theirs to work off - and it is drawn even
         when the office has locked the card, because the counters underneath
         it have not changed and half the row is greyed already. It rides in
         space the row already had, so a card with four stages on it is no
         taller than it was with one number in the headline. */
      (holds ? '<span class="stepleft tab">' + esc(ST.leftWords(ST.stageLeft(g, stage))) + '</span>'
             : locked ? "" : ' <span class="nomine">not yours</span>') + '</span>' +
    '<span class="stepc">' + b("&minus;", "-1", "sbtn") +
      '<span class="stepn tab' + (full ? " full" : "") + '">' + v + '</span>' + b("+", "1", "sbtn") +
      b(none || !full ? "All" : "None", full ? "none" : "all", "sall", none) + '</span>' +
    '</div>';
}

/** The hotmelting tablet's look at the cutting bench (2026-10-02): how much of
    this job is cut, and how much of its tuff is done where it has any. Greyed
    like a stage that is not yours, with no button in it - there is nothing
    here to tap, and tap() refuses another tablet's stage whatever is drawn.
    Both counts are on the row this page already reads: no new read. */
function seenHtml(g) {
  if (PAGE_STAGE !== "hotmelt") return "";
  return '<div class="step locked seen"><span class="stepl tab">' +
    esc("Cut " + g.cut + "/" + g.total + (g.tuffTotal > 0 ? " · Tuff " + g.tuff + "/" + g.tuffTotal : "")) +
    '</span></div>';
}

/** Everything inside one card. The card element itself is kept between draws
    (see paintBoard), so only this string is ever rebuilt.

    The headline says how big the job is - the glasses it holds, and the tuff
    beside them where there is any - and it is the same for everybody. One
    number that tried to say how much of it was left for the person holding the
    tablet was shipped on 2026-09-09 and rejected by the owner on sight: it
    added up the stages they held, so 49 glasses and 10 tuff read 157. What is
    left is now one number per stage, drawn by stepHtml on the row of the stage
    it counts, where the buttons that move it are.

    The rows are still drawn from PERSON, which is not part of the board, so
    paintBoard has to notice PERSON moving on its own - see pState(). */
function cardInner(g) {
  const owed = owedFor(g.id), bad = badFor(g.id);
  const lost = blockedFor(g.job);
  /* THIS PAGE'S STAGE, and tuff beside it on the cutting page - only on the
     jobs that have tuff, and only for somebody who holds it. A second row of
     nothing on a screen somebody reads with a sheet of glass in their other
     hand is a row too many, and the other tablet's stage is not this tablet's
     business at all. */
  const stages = ST.ALL_STAGES.filter(s => s[0] === PAGE_STAGE ||
    (s[0] === ST.TUFF_STAGE && tuffHere() && g.tuffTotal > 0 && ST.canStage(PERSON, ST.TUFF_STAGE)));
  return '<div class="chead">' +
      '<span class="cond job">' + esc(g.job) + '</span>' +
      '<span class="cust">' + esc(g.customer || "—") + '</span>' +
      /* a job off In production, on the Finished tab: say where it is */
      (!g.active && g.section ? '<span class="csec">' + esc(g.section) + '</span>' : "") +
      /* fabrication has finished this job and this person's glass has not */
      (waitingOn(g) ? '<span class="fwait">FABRICATION DONE — GLASS WAITING</span>' : "") +
      '<span class="cnum tab">' + esc(ST.glassWords(g.total) +
          (g.tuffTotal > 0 ? " · " + g.tuffTotal + " tuff" : "")) + '</span>' +
    '</div>' +
    '<div class="steps">' + stages.map(s => stepHtml(g, s[0], s[1])).join("") + seenHtml(g) + '</div>' +
    fabLineHtml(g) +
    /* "glass", not "job": since 2026-09-21 the tuff stepper beside it is still
       live under this line, and a word that said otherwise would be wrong */
    (g.officeDone ? '<div class="officedone">the office has marked this job’s glass finished</div>' : "") +
    (lost ? '<div class="unsaved">' + esc(lost.stages.map(s =>
        ST.stageLabel(s.stage) + " " + s.value).join(", ")) +
        ' was not saved — the office marked this job finished first</div>' : "") +
    (bad ? '<div class="unsaved">not saved yet — retrying</div>'
         : owed ? '<div class="saving">saving…</div>' : "") +
    /* under everything the card counts: a word for the office about this job,
       and this station's own earlier words about it. Closed it is one button,
       so a card with nothing to say about it is the height it always was. */
    NOTES.html(g.job);
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
  const mine = pagePeople();
  if (!mine.length)
    return '<div class="picker"><div class="pickh">Who are you?</div>' +
      '<div class="msg">Nobody is set up for ' + esc(stageWords().toLowerCase()) +
      ' at the Glass station yet. Ask the office to add ' +
      'you to the “Station people” list.</div></div>';
  return '<div class="picker">' +
    '<div class="pickh">Who are you?</div>' +
    '<div class="pgrid">' + mine.map(p =>
      '<button class="pbtn" data-person="' + esc(p.name) + '">' +
        '<span class="pname">' + esc(p.name) + '</span>' +
        '<span class="pstages">' + esc(myStages(p).map(ST.stageLabel).join(" · ")) + '</span>' +
      '</button>').join("") + '</div></div>';
}

/* ---- which tablet is this? --------------------------------------------------
   Shown once, on a device that has been told neither by its URL nor by its own
   storage. Two buttons, because there are two glass tablets; the answer is
   remembered and the page carries on to the person picker. */
const stageWords = () => ST.stageLabel(PAGE_STAGE);
function chooserHtml() {
  return '<div class="picker">' +
    '<div class="pickh">Which tablet is this?</div>' +
    '<div class="picksub">Remembered on this device. The office can change it any time.</div>' +
    '<div class="pgrid">' + ST.STAGES.map(s =>
      '<button class="pbtn" data-stagepick="' + esc(s[0]) + '">' +
        '<span class="pname">' + esc(s[1]) + '</span></button>').join("") +
    '</div></div>';
}
function wireChooser(host) {
  host.querySelectorAll("[data-stagepick]").forEach(el => el.onclick = () => {
    setStage(el.dataset.stagepick);
    if (!PAGE_STAGE) return;
    /* the people this page may show have just changed, so whoever was signed in
       under the old answer is not signed in under this one */
    switchPerson();
    render();
  });
}
/** The header's way back to the chooser: this tablet becomes the other stage.
    It signs the person out, because the name on screen was chosen for the
    stage that is leaving. */
function askStage() {
  const other = ST.STAGE_KEYS.find(k => k !== PAGE_STAGE) || "";
  if (!other) return;
  if (typeof confirm === "function" &&
      !confirm("Make this the " + ST.stageLabel(other) + " tablet? It will sign you out.")) return;
  setStage(other);
  switchPerson();
  render();
}

function wirePicker(host) {
  host.querySelectorAll("[data-person]").forEach(el => el.onclick = () => {
    const p = pagePeople().find(x => x.name === el.dataset.person);
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
let LIST = null;                // the one column of cards, for the tab it was drawn for
let LIST_TAB = "";
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
         (lost ? "!" + lost.stages.map(s => s.stage + s.value).join(",") : "") +
         /* the note channel is not in the list either: opening a composer, a
            note arriving from the other shift and a send that failed all change
            what this card draws and none of them move a counter. The DRAFT is
            deliberately not in the signature - see dressCard. */
         "/" + NOTES.sig(g.job) +
         /* ... and neither is the fabrication line: it comes from another list,
            so it is in the card's signature here, and a fabrication tap redraws
            only the cards it changes */
         "/" + JSON.stringify(fabOf(g.job));
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
/** Redraw one card - and put the caret back if it was in this card's note box.
    A card is rebuilt whole (innerHTML), which is fine for numbers and buttons
    and would be very much not fine for a half-typed sentence: the ten-second
    poll would take the caret out of it six times a minute.

    Only the CARET is restored, never the text: the box is drawn from the draft,
    and the draft is what the input listener has been keeping current. That is
    also why a sent note does not come back - send() empties the draft, the box
    is redrawn empty, and the caret lands at nought in it. */
function dressCard(el, g) {
  const act = document.activeElement;
  const at = act && act.dataset && act.dataset.cmbox === g.job ? act.selectionStart : null;
  el.className = "card" + (g.finished ? " done" : "");
  el.innerHTML = cardInner(g);
  if (at == null || !el.querySelector) return;
  const box = el.querySelector('[data-cmbox="' + g.job + '"]');
  if (!box) return;
  const to = Math.min(Number(at) || 0, String(box.value || "").length);
  try { box.focus(); if (box.setSelectionRange) box.setSelectionRange(to, to); } catch (e) {}
}
function paintBoard(host, board) {
  /* a tab switch starts the column again: its cards are another set */
  if (!LIST || LIST_TAB !== TAB) {
    host.innerHTML = "";
    LIST = document.createElement("div"); LIST.className = "grp";
    LIST_TAB = TAB;
    host.appendChild(LIST);
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
    /* ... and a move BLURS whatever is inside it. Another job finishing while
       somebody is half way through a note would otherwise drop the caret and
       close the tablet's keyboard mid-sentence. The typing itself was never at
       risk (it is in the draft, and the node is moved rather than rebuilt), but
       having to find the box again with a sheet of glass in the other hand is
       the sort of thing that stops people using it. So the caret is put back
       where it was, the same way dressCard does after a redraw. */
    const act = document.activeElement;
    const box = act && act.dataset && act.dataset.cmbox ? act : null;
    const at = box ? box.selectionStart : null;
    board.forEach(g => {
      const el = NODES[g.job];
      if (!el) return;
      LIST.appendChild(el);
    });
    if (box && document.activeElement !== box) {
      try {
        box.focus();
        if (box.setSelectionRange && at != null) box.setSelectionRange(at, at);
      } catch (e) {}
    }
  }
  BOARD_PREV = board;
}

/** Everything that is not a card: the header, the messages, the picker. */
function render() {
  const host = $("#board");
  if (!host) return;
  /* which tablet this is, in the header, so nobody has to guess which of the
     two is in front of them */
  const br = $("#brand");
  if (br) br.textContent = PAGE_STAGE ? "GLASS · " + stageWords().toUpperCase() : "GLASS STATION";
  const sg = $("#stagebtn");
  if (sg) {
    sg.hidden = !PAGE_STAGE;
    sg.style.display = PAGE_STAGE ? "" : "none";
    sg.textContent = PAGE_STAGE ? stageWords() + " ▾" : "";
  }
  /* the name, and only the name, is the text: it is what every tap is logged
     under, and it must not be what an ellipsis eats. The stages they hold here
     ride in data-stages, which the stylesheet shows after it where there is room. */
  const hdr = $("#whois");
  if (hdr) {
    hdr.textContent = PERSON ? PERSON.name : "";
    hdr.dataset.stages = PERSON ? myStages(PERSON).map(ST.stageLabel).join(", ") || "no stages" : "";
  }
  const sw = $("#switchbtn");
  if (sw) { sw.hidden = !PERSON; sw.style.display = PERSON ? "" : "none"; }
  /* the end-of-day button: only on a stage whose definition has a day sheet,
     and only once somebody has said who they are */
  const dy = $("#daybtn");
  if (dy) {
    const on = !!DAY.sheet() && !!PERSON && !PROBLEM;
    dy.hidden = !on;
    dy.style.display = on ? "" : "none";
  }
  /* the search box belongs to the board: there is nothing to search on the
     picker, and a box over an error message only looks broken */
  const sb = $("#search");
  const boarding = !!PAGE_STAGE && !PROBLEM && PEOPLE_READ && !!PERSON && READY && !DAY.shown;
  if (sb) { sb.hidden = !boarding; sb.style.display = boarding ? "" : "none"; }
  /* the board as it stands, before the tab and the box have narrowed it: the
     number beside the box is read off the In production cards, and the cards
     below off the chosen tab, filtered */
  const now = boarding ? { tabs: tabsNow() } : null;
  const live = now ? activeOf(now.tabs) : null;
  /* how many jobs are waiting on this person's glass (the capsule, below) -
     counted here, before the tabs are drawn, because the filter letting go
     can put the board back on the tab it came from */
  const waitN = now ? now.tabs.floor.concat(now.tabs.finished).filter(waitingOn).length : 0;
  if (now && !waitN && GFAB && GFAB.items) setWait(false);
  /* "Cut first": the hotmelting tablet's own switch, with the board */
  const cf = $("#cutfirst");
  if (cf) {
    const on = boarding && PAGE_STAGE === "hotmelt";
    cf.hidden = !on; cf.style.display = on ? "" : "none";
    cf.className = CUTFIRST ? "on" : "";
    cf.setAttribute("aria-pressed", CUTFIRST ? "true" : "false");
    cf.textContent = "Cut first: " + (CUTFIRST ? "on" : "off");
  }
  const tabsEl = $("#gtabs");
  if (tabsEl) { tabsEl.hidden = !boarding; tabsEl.style.display = boarding ? "" : "none"; }
  [["#tabfloor", "floor", "On floor"], ["#tabfin", "finished", "Finished"]].forEach(([s, t, label]) => {
    const b = $(s);
    if (!b) return;
    b.className = TAB === t ? "on" : "";
    b.setAttribute("aria-selected", TAB === t ? "true" : "false");
    b.textContent = label + (now ? " · " + now.tabs[t].length : "");
  });
  /* both tabs match the search: a tappable line says how many more are on the
     other one (owner's decision 3, 2026-09-24) */
  /* the jobs fabrication has finished and this person has not: one capsule,
     only when there are any, and a tap on it shows only those. Counted over
     the whole board, like the number beside it. With none left the filter
     lets go - but only on a GOOD fabrication read that really says none: a
     read that cannot be made just now must not drop it, so the capsule stays
     up while the filter is on and there is always something to tap. */
  const gw = $("#gwait");
  if (gw) {
    const on = !!now && (waitN > 0 || WAITONLY);
    gw.hidden = !on; gw.style.display = on ? "" : "none";
    gw.className = WAITONLY ? "on" : "";
    gw.setAttribute("aria-pressed", WAITONLY ? "true" : "false");
    /* the number; the words after it are the stylesheet's, as many as fit */
    gw.textContent = on ? String(waitN) : "";
    gw.setAttribute("aria-label", waitN + " waiting on glass");
  }
  const pick = now ? ST.tabSearch(now.tabs, TAB, QUERY, shownCards) : null;
  const more = $("#more");
  if (more) {
    const on = !!(pick && pick.more);
    more.hidden = !on;
    more.style.display = on ? "" : "none";
    more.textContent = on ? pick.more + " more " + (pick.other === "finished" ? "in Finished" : "on floor") : "";
  }
  /* what is left for the person signed in, one number per stage they hold,
     added down the whole board: the very same numbers the cards show, said the
     same way, so the header and the cards always agree and every tap moves
     both. It is next to the box that narrows the cards and deliberately not
     narrowed by it - a job number typed in must not make somebody's day look
     shorter. It comes and goes with the search box for the same reason: a
     total floating over "ask the office for the permission" is a number about
     nothing. And somebody holding no stage at all gets no pill: they have
     nothing to work off, and a number here would be about somebody else's
     work. */
  const gt = $("#gtotal");
  if (gt) {
    /* this page's stage and no other, and tuff stays out of it as it always
       has: the number beside the box is the work this tablet is for */
    const lefts = boarding
      ? ST.boardLefts(live, myStages(PERSON).filter(k => k !== ST.TUFF_STAGE)) : [];
    const on = boarding && lefts.length > 0;
    gt.hidden = !on;
    gt.style.display = on ? "" : "none";
    gt.textContent = on
      ? lefts.map(e => e.label + " " + ST.leftWords(e.left)).join(" · ") : "";
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
  if (!PAGE_STAGE || PROBLEM || !PEOPLE_READ || !PERSON || !READY) {
    /* a message or the picker takes the board's place, so the card nodes are
       let go: the next good read paints them fresh */
    LIST = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    /* which tablet this is comes before everything, including a SharePoint
       problem: the answer is about the device and needs no list to give */
    if (!PAGE_STAGE) { host.innerHTML = chooserHtml(); wireChooser(host); return; }
    if (!PROBLEM && PEOPLE_READ && !PERSON) { host.innerHTML = pickerHtml(); wirePicker(host); return; }
    host.innerHTML = '<div class="msg">' + esc(words()) + againHtml() + '</div>';
    wireAgain();
    return;
  }

  /* the end-of-day sheet takes the board's place, the way the picker does: one
     thing on screen at a time, and the card nodes are let go while it is up */
  if (DAY.shown && DAY.sheet()) {
    LIST = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    host.innerHTML = DAY.html();
    DAY.wire(host);
    return;
  }
  DAY.shown = false;                     // a stage with no sheet has nothing to show

  /* the search box narrows the board and never becomes it: an empty box is
     every card, and a box nothing matches says so rather than looking broken */
  /* ... and "Cut first" only re-orders what is left: On floor, hotmelting page */
  const board = ST.cutFirst(shownCards(now.tabs[TAB], QUERY),
    CUTFIRST && PAGE_STAGE === "hotmelt" && TAB === "floor");
  if (!board.length) {
    LIST = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    host.innerHTML = '<div class="msg">' +
      (WAITONLY ? "No job waiting on glass under " + (TAB === "floor" ? "On floor" : "Finished") +
                  (QUERY ? " matches “" + esc(QUERY) + "”." : ".")
       : QUERY ? "No job under " + (TAB === "floor" ? "On floor" : "Finished") +
               " matches “" + esc(QUERY) + "”."
       : TAB === "floor" ? "Nothing on the floor right now."
       : "No finished jobs yet.") + '</div>';
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
    if (CW.forgetStationSite) CW.forgetStationSite(true, ST.GLASS.site);   // a person asked: look now
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
       total */
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
      /* the note channel: open or close the composer, or send what is in it.
         Neither can reach a counter, a queue or anything the workbook knows
         about - they only ever add one row to `Station comments`. */
      if (d.cmt) {
        ev.stopPropagation();
        NOTES.toggle(d.cmt);
        render();
        /* opened on a channel that has not answered yet - never looked, or the
           site was not resolved when the page started: ask now rather than show
           an empty thread until somebody reloads the tablet */
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
     keystroke is a caret lost on every keystroke. dressCard puts the caret back
     when something else redraws the card mid-sentence. */
  host.addEventListener("input", ev => {
    const d = (ev.target && ev.target.dataset) || {};
    if (!d.cmbox) return;
    touch();
    NOTES.setDraft(d.cmbox, ev.target.value);
  });
}

/** Send the note typed on one job's card. One row appended to `Station
    comments`, from the person signed in, for this station. Nothing else is
    written anywhere, and nothing is ever edited or removed. */
async function sendNote(job) {
  if (!PERSON) return;                       // the board is not drawn without one
  if (!NOTES.draftOf(job).trim()) return;
  /* send() marks the job as sending before it awaits anything, so this render
     is the one that greys the button and says "Sending…" */
  const going = NOTES.send(job, who());
  render();
  await going;
  render();
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
    if (latest !== BUILD_NOW && !owingWrites()) location.reload(true);
  } catch (e) { /* offline or blocked - the board is what matters */ }
}

/** One turn of the clock (STU.stuTicker: every five seconds, every two for a
    while after this tablet's own tap, never two turns at once). A named
    function, so a test can take exactly one turn of it. */
async function tickOnce() {
  render();
  /* a people list that has not answered yet is retried on the same clock: one
     failed read at start-up must not leave the tablet stuck at the picker's
     door until somebody thinks to reload it */
  if (!PEOPLE_READ) await readPeople();
  await pollList();
  /* fabrication's list, read only, on the same clock; it cannot fail loudly */
  if (await readFab()) render();
  /* the note channel keeps its own counsel: it only asks the list anything
     while somebody has a composer open, and never more often than
     ST.COMMENT_POLL_MS. A tablet nobody is writing on sends no request for it
     at all. A failure there is quiet by construction and cannot touch the
     board. */
  if (await NOTES.poll()) render();
}
const TICK = STU.stuTicker(tickOnce);

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

/** One keystroke in the search box (see start()). */
function onSearch(value) {
  const was = QUERY.trim();
  QUERY = String(value || "");
  const q = QUERY.trim();
  const from = TAB;
  if (q && !was && PRETAB == null) PRETAB = TAB;
  if (!q) { if (PRETAB) TAB = PRETAB; PRETAB = null; saveTab(); }
  else TAB = ST.tabSearch(tabsNow(), TAB, QUERY, shownCards).tab;
  if (TAB !== from) { try { window.scrollTo(0, 0); } catch (e) {} }
  touch(); render();
}

async function start() {
  showGate(false);
  applyTheme(themeNow());
  const tb = $("#themebtn");
  if (tb) tb.onclick = () => applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  $("#outbtn").onclick = () => { if (confirm("Sign out of the Glass station?")) CW.signOut(); };
  const sw = $("#switchbtn");
  if (sw) sw.onclick = () => switchPerson();
  const sgb = $("#stagebtn");
  if (sgb) sgb.onclick = () => askStage();
  const dyb = $("#daybtn");
  if (dyb) dyb.onclick = () => DAY.open();
  /* the box is in the header, outside #board, so typing in it never rebuilds
     the node the caret is in - only the cards under it are redrawn */
  const sb = $("#search");
  /* SEARCH ACROSS BOTH TABS (owner's decision 3, 2026-09-24): the tab this
     board was on when the typing began is remembered; while there is a query,
     a tab with no match gives way to the other tab if that one has one; an
     emptied box goes back to the remembered tab. Decided on typing only, so a
     tab the person taps mid-search stays where they put it. */
  if (sb) sb.oninput = () => onSearch(sb.value);
  [["#tabfloor", "floor"], ["#tabfin", "finished"]].forEach(([s, t]) => {
    const b = $(s);
    if (b) b.onclick = () => {
      if (TAB !== t) { TAB = t; WAITTAB = null; if (PRETAB == null) saveTab(); try { window.scrollTo(0, 0); } catch (e) {} }
      touch(); render();
    };
  });
  const mb = $("#more");
  if (mb) mb.onclick = () => {
    /* the person's own choice of tab, as a tap on the tab itself is */
    TAB = TAB === "floor" ? "finished" : "floor"; WAITTAB = null;
    try { window.scrollTo(0, 0); } catch (e) {}
    touch(); render();
  };
  const cfb = $("#cutfirst");
  if (cfb) cfb.onclick = () => {
    CUTFIRST = !CUTFIRST;
    try { localStorage.setItem(CUTFIRST_KEY, CUTFIRST ? "1" : "0"); } catch (e) {}
    touch(); render();
  };
  const gwb = $("#gwait");
  if (gwb) gwb.onclick = () => {
    setWait(!WAITONLY);
    try { window.scrollTo(0, 0); } catch (e) {}
    touch(); render();
  };
  render();
  await readPeople();
  /* NOT awaited: a slow floor site must never hold the glass board back. Until
     it answers the cards carry no fabrication line at all. */
  readFab().then(ch => { if (ch) render(); });
  await readList();
  /* what has already been said about today's jobs, so a second shift does not
     retype the first shift's note. One read, quiet, and unable to fail loudly:
     a missing list is a line inside the composer, never a board taken away. */
  await NOTES.read();
  /* the day sheets (and the target, where the stage has one), once - and ONLY
     on a stage whose definition has a sheet */
  if (DAY.sheet()) await DAY.read();
  await flushQueue();                            // taps owed from a previous visit
  TICK.start();
  if (peopleT) clearInterval(peopleT);
  peopleT = setInterval(() => {
    readPeople();
    /* the office's target and anybody else's sheets ride the ten-minute clock:
       a day sheet moves once a day and a target less often than that */
    if (DAY.sheet()) DAY.read().then(() => { if (DAY.shown) render(); }, () => {});
  }, PEOPLE_MS);
  if (buildT) clearInterval(buildT);
  checkBuild(); buildT = setInterval(checkBuild, BUILD_MS);
}

(async function boot() {
  applyTheme(themeNow());
  /* an expired sign-in is renewed without anybody tapping, when it can be -
     and never by leaving the page while a tap is still owed */
  if (CW.stationRenew) CW.stationRenew(() => !owingWrites());
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
