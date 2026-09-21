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
  /* AND NEITHER IS THE DAY SHEET (review, 2026-09-21). It was left open across
     a change of person, so the next name to be picked was shown the last
     person's half-typed numbers, on a form whose Save would have filed them
     under the new name. The typing is not lost - it is in storage under whose
     it is, and comes back when they pick their name again. */
  DAYOPEN = false; DAYBAD = ""; DAYDRAFT = null;
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
  try { localStorage.setItem(DAYQ_KEY, JSON.stringify(DAYQ)); } catch (e) {}
}
/** Does this tablet owe anything at all? What keeps the retry timer armed. */
const owingAnything = () => !!(Object.keys(QUEUE).length || Object.keys(LOGQ).length ||
                               Object.keys(DAYQ).length);
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
    await flushDay();
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
   that one POST, and `Station targets` is never written here at all.       */
const DAYQ_KEY = "cw_stationdayq";       // a saved sheet this tablet still owes
const DRAFT_KEY = "cw_daysheetdraft";    // ... and one typed but not yet saved
let DAYOPEN = false;             // the sheet is on screen, in the board's place
let DAYROWS = [];                // this stage's saved sheets, as last read
let DAY_OK = null;               // null not looked · false missing or unreachable · true read it
let DAY_WHY = "";
let TARGET = null;               // the weekly target in force, or null
let DAYDRAFT = null;             // { day, stage, who, counts, note, target } typed but not saved
let DAYBAD = "";                 // what is wrong with what is typed, in words
/* "there is no such list" and "I could not reach the list" are two different
   answers and only the first of them is a reason to refuse a save (review,
   2026-09-21). DAY_OK is false for both; this says which. */
let DAY_MISSING = false;
let DAYQ = {};
try { DAYQ = cleanDayQ(JSON.parse(localStorage.getItem(DAYQ_KEY) || "{}")); } catch (e) { DAYQ = {}; }

/** Whatever is in storage, rebuilt through the row builder - so an edited
    localStorage can put no column on the wire that a save could not. */
function cleanDayQ(raw) {
  const out = {};
  Object.keys(raw || {}).forEach(k => {
    const e = raw[k], f = (e && e.fields) || null;
    if (!f || !f.Title) return;
    const ds = ST.daySheetOf(ST.GLASS, f.Stage);
    if (!ds) return;
    const counts = {};
    ds.counts.forEach(c => { counts[c[0]] = f[c[0]]; });
    const rebuilt = ST.dayFields(ST.GLASS, f.Stage, { day: f.Day, who: f.Who, counts: counts,
                                                      note: f.Note, weekTarget: f.WeekTarget, at: f.SavedAt });
    /* how often SharePoint has refused it survives the reload with it: a row
       that has been refused three times must not read as brand new every time
       somebody restarts the tablet */
    if (rebuilt) out[rebuilt.Title] = { fields: rebuilt, err: 0,
                                        refused: Math.max(0, Math.round(Number(e.refused) || 0)) };
  });
  return out;
}

/** This page's day sheet, or null - and the answer to every "does this tablet
    have one" question below. */
const daySheet = () => (typeof ST.daySheetOf === "function" ? ST.daySheetOf(ST.GLASS, PAGE_STAGE) : null);
const dayOpts = () => ({ siteId: SITEID, fields: ST.dayFieldsFor(ST.GLASS, PAGE_STAGE) });
const targetOpts = () => ({ siteId: SITEID, fields: ST.TARGET_FIELDS });
const dayToday = () => ST.dayKey(new Date());

/** Both lists, quietly. A missing one is a state, never an error: the board
    must not be takeable away by a sheet nobody has opened. */
async function readDay() {
  if (!daySheet() || !SITEID) return false;
  let got = false;
  try {
    const items = await CW.listItems(ST.DAY_LIST, dayOpts());
    if (items == null) {
      DAY_OK = false; DAY_MISSING = true; DAY_WHY = ST.DAY_MISSING_FLOOR; return false;
    }
    DAYROWS = ST.dayRows(items, daySheet().counts,
                         { station: ST.STATION_NAME, stage: PAGE_STAGE });
    DAY_OK = true; DAY_MISSING = false; DAY_WHY = ""; got = true;
  } catch (e) {
    /* a throw is the wifi, a bad gateway, a refused token - never "there is no
       such list", which is the null above. DAY_MISSING is deliberately left
       alone: a save must still be queued through a bad afternoon. */
    if (DAY_OK !== true) { DAY_OK = false; DAY_WHY = ST.DAY_UNREACHABLE; }
    console.warn("[station] the day sheets could not be read:", (e && e.message) || e);
    return false;
  }
  /* the target is a nicety - the sheet saves perfectly well without one, and
     "no target set" is a thing it is allowed to say */
  try {
    const t = await CW.listItems(ST.TARGET_LIST, targetOpts());
    const hit = t == null ? null : ST.targetOf(t, ST.STATION_NAME, PAGE_STAGE);
    TARGET = hit ? hit.target : null;
  } catch (e) { /* keep the last one */ }
  return got;
}

/* ---- the draft, and WHICH DAY IT IS ABOUT -----------------------------------
   A DRAFT BELONGS TO THE DAY IT WAS STARTED (review, 2026-09-21), not to
   whatever day it happens to be when Save is tapped. The first build keyed it
   on "today", which broke in both directions on the one shift it matters on:
   typing at 23:55 and tapping Save at 00:01 filed the evening's work under
   tomorrow - the wrong day, the wrong ISO week, possibly the wrong target, and
   on a row the cutter cannot correct - and the same tick silently threw away a
   draft that had been typed and not yet saved.

   So the record carries its own `day`, the form's header shows THAT day, Save
   files it under it, and it survives until the end of the FOLLOWING day before
   being dropped. It also carries `who`: a draft is one person's writing, and
   the tablet is passed around - without that, whoever picked their name next
   would be handed somebody else's numbers to save under their own.

   The target is stamped on at the same moment for the same reason: the sheet
   is about that day, so it is measured against the target that was in force
   then rather than whatever the office has changed it to since.             */
function draftKeep(d) {
  if (!d || d.stage !== PAGE_STAGE) return null;
  if (String(d.who || "") !== who()) return null;        // one person's writing
  const day = ST.dayKey(String(d.day || ""));
  if (!day) return null;
  /* today, or yesterday: a night shift's sheet is still there in the morning,
     and the day after that it is gone rather than sitting there for a week */
  const age = Math.round((Date.parse(dayToday() + "T12:00:00Z") -
                          Date.parse(day + "T12:00:00Z")) / 86400000);
  if (!isFinite(age) || age < 0 || age > 1) return null;
  return { day: day, stage: d.stage, who: String(d.who || ""),
           counts: d.counts || {}, note: String(d.note || ""),
           target: d.target == null ? null : Number(d.target) };
}
function loadDraft() {
  let d = null;
  try { d = JSON.parse(localStorage.getItem(DRAFT_KEY) || "null"); } catch (e) { d = null; }
  return draftKeep(d);
}
function draftNow() {
  if (!draftKeep(DAYDRAFT))
    DAYDRAFT = loadDraft() || { day: dayToday(), stage: PAGE_STAGE, who: who(),
                                counts: {}, note: "", target: TARGET };
  return DAYDRAFT;
}
function saveDraft() { try { localStorage.setItem(DRAFT_KEY, JSON.stringify(DAYDRAFT)); } catch (e) {} }
function clearDraft() { DAYDRAFT = null; try { localStorage.removeItem(DRAFT_KEY); } catch (e) {} }
/** The day the sheet on screen is about: the draft's own, which is normally
    today and is yesterday for a draft carried over midnight. */
const sheetDay = () => draftNow().day;

/** This person's sheet for that day, if it is already on the list. (`owedFor`
    up in the counter queue is a different question about a different queue,
    which is why these two say "daySheet" out loud.) */
function daySheetSaved(day) {
  const nm = who();
  return DAYROWS.find(r => r.day === day && r.who === nm) || null;
}
/** ... or still owed by this tablet, which reads the same to whoever typed it
    apart from the word: "waiting to send" rather than "saved". */
function daySheetOwed(day) {
  return DAYQ[ST.dayTitle(ST.STATION_NAME, PAGE_STAGE, day, who())] || null;
}
/** What is typed now, added up - the "Today: N" under the boxes. */
function draftTotal() {
  const ds = daySheet();
  if (!ds) return 0;
  const d = draftNow();
  let n = 0;
  ds.counts.forEach(c => { n += ST.dayCount(d.counts[c[0]]) || 0; });
  return n;
}
/** Is everything typed a whole number within the cap? */
function draftValid() {
  const ds = daySheet();
  if (!ds) return false;
  const d = draftNow();
  return ds.counts.every(c => ST.dayCount(d.counts[c[0]]) != null);
}
/** Has anybody actually written anything? AN UNTOUCHED FORM IS NOT FOUR
    NOUGHTS (review, 2026-09-21): Save was live the moment the sheet opened, so
    one stray tap filed a day of nothing on a row nobody on the floor can
    correct. A day with no sheets cut but a note ("machine down all day") is a
    real entry and saves perfectly well. */
function draftTyped() {
  const ds = daySheet();
  if (!ds) return false;
  const d = draftNow();
  return ds.counts.some(c => String(d.counts[c[0]] == null ? "" : d.counts[c[0]]).trim() !== "") ||
         String(d.note || "").trim() !== "";
}
/** May Save be tapped at all? */
const draftOk = () => draftValid() && draftTyped();

/** "This week: N of T": this person's saved sheets in the week of the day this
    sheet is about, plus what is typed now, against the target in force. */
function weekWords(extra) {
  const wk = ST.isoWeek(sheetDay()).key;
  const target = draftNow().target == null ? TARGET : draftNow().target;
  const done = ST.dayWeekTotal(DAYROWS, who(), wk) + (extra || 0);
  const unit = (daySheet() || {}).unit || "sheets";
  return { done: done, target: target,
           words: target == null ? done + " " + unit + " · no target set"
                                 : done + " of " + target + " " + unit,
           pct: target > 0 ? Math.max(0, Math.min(100, Math.round(done * 100 / target))) : 0 };
}

/** Save the sheet: one row queued, sent by the page's own queue. */
function saveDay() {
  const ds = daySheet();
  if (!ds || !PERSON) return false;
  const d = draftNow();
  if (daySheetSaved(d.day) || daySheetOwed(d.day)) return false;   // one per person per day
  if (!draftTyped()) { DAYBAD = "Fill in a number or write a line first."; render(); return false; }
  const fields = ST.dayFields(ST.GLASS, PAGE_STAGE, {
    day: d.day, who: who(), counts: d.counts, note: d.note,
    weekTarget: d.target == null ? TARGET : d.target, at: new Date().toISOString() });
  if (!fields) {
    DAYBAD = "Whole numbers only, please — no minus signs, no decimals, nothing over " +
             ST.DAY_COUNT_MAX + ".";
    render(); return false;
  }
  /* ONLY A LIST THAT IS NOT THERE refuses the save (review, 2026-09-21). A list
     that could not be READ is the workshop wifi, and a sheet typed out at the
     end of a shift must be queued through that exactly like any other offline
     save - refusing it was the one way this feature could lose somebody's day. */
  if (DAY_MISSING) { DAYBAD = DAY_WHY || ST.DAY_MISSING_FLOOR; render(); return false; }
  if (typeof confirm === "function" &&
      !confirm("Save the sheet for " + d.day + "? It cannot be changed from the tablet afterwards.")) return false;
  DAYBAD = "";
  DAYQ[fields.Title] = { fields: fields, err: 0, refused: 0 };
  saveQueue();
  clearDraft();
  touch();
  render();
  flushQueue();
  return true;
}
/** Send the owed sheet. A refusal is very often the list's own unique rule
    catching a second tablet, or this tablet replaying a row that landed and
    whose answer was lost on the workshop wifi. Both of those mean ALREADY
    SAVED, so the list is read back before anything is called a failure - and
    then the row that is there is what the person is shown.

    A row SharePoint keeps refusing (a 4xx: the column was renamed, the account
    lost its write, the list was locked) is not the wifi, and after
    DAY_REFUSE_MAX of them the card stops saying "waiting to send" - which
    invites somebody to stand there waiting for it - and says to tell the
    office. It STAYS QUEUED either way: the tablet cannot show it, so throwing
    it away would be the only way the day was really lost. */
const DAY_REFUSE_MAX = 3;
async function flushDay() {
  if (!SITEID || !daySheet()) return;
  const keys = Object.keys(DAYQ);
  for (let i = 0; i < keys.length; i++) {
    const e = DAYQ[keys[i]];
    if (!e) continue;
    try {
      await CW.listAdd(ST.DAY_LIST, e.fields, dayOpts());
      delete DAYQ[keys[i]];
      await readDay();
    } catch (err) {
      const landed = (await readDay()) && DAYROWS.some(r => r.title === e.fields.Title);
      if (landed) { delete DAYQ[keys[i]]; saveQueue(); continue; }
      e.err = 1;
      /* only a refusal counts. A dropped connection is not the list saying no,
         and a tablet carried out of range for an hour must not end the day
         telling somebody to go and see the office about nothing. */
      const refused = !!(CW.isMissing && CW.isMissing(err)) || !!(CW.isRefused && CW.isRefused(err)) ||
                      /->\s*4\d\d\b/.test((err && err.message) || "");
      if (refused) e.refused = (Number(e.refused) || 0) + 1;
      console.warn("[station] the day sheet is not saved yet" +
                   (e.refused >= DAY_REFUSE_MAX ? " and has been refused " + e.refused + " times" : "") +
                   ":", (err && err.message) || err);
    }
    saveQueue();
  }
}
/** What the card says about an owed sheet. */
function owedWords(e) {
  if (e && Number(e.refused) >= DAY_REFUSE_MAX)
    return "could not be saved — tell the office";
  if (e && e.err) return "waiting to send — it will go when the tablet is back on the wifi";
  return "waiting to send…";
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

/** The ten-second poll: only what moved. A token that has gone stale (Graph
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
async function pollList() {
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
  const row = boardNow().find(g => g.id === id);
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
function boardNow() {
  const over = ITEMS.map(it => {
    const q = queuedFor(String(it.id));
    if (!q) return it;
    return { id: it.id, fields: Object.assign({}, it.fields, q) };
  });
  const board = ST.jobBoard(over);
  board.forEach(g => {
    g.finished = ST.stageComplete(g, PAGE_STAGE) && !(tuffHere() && ST.tuffOwed(g));
  });
  return board;
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
      '<span class="cnum tab">' + esc(ST.glassWords(g.total) +
          (g.tuffTotal > 0 ? " · " + g.tuffTotal + " tuff" : "")) + '</span>' +
    '</div>' +
    '<div class="steps">' + stages.map(s => stepHtml(g, s[0], s[1])).join("") + '</div>' +
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

/* ---- the end-of-day sheet, drawn ---------------------------------------------
   One column, full width, portrait, nothing to scroll sideways and nothing
   tapped under 44 px - the tablet rule of 2026-09-17. It takes the board's
   place rather than opening over it, so there is one thing on screen at a time
   and Back is the only way out.                                             */
/** One saved sheet, read-only: the day, the counts and the total. */
function dayLineHtml(r, counts) {
  return '<div class="dayline"><span class="daylday">' + esc(r.day) + '</span>' +
    '<span class="dayldow">' + esc(r.weekday.slice(0, 3)) + '</span>' +
    /* the definition's own short label, not the first word of the long one -
       which read "Other 2" for the obscure count */
    counts.map(c => '<span class="daylnum tab">' + esc(ST.dayCountShort(c)) + ' ' +
      (r.counts[c[0]] || 0) + '</span>').join("") +
    '<span class="dayltot tab">' + r.total + '</span></div>';
}
function daySheetHtml() {
  const ds = daySheet();
  if (!ds || !PERSON) return '<div class="msg">Nothing to fill in here.</div>';
  const counts = ds.counts, unit = ds.unit || "sheets";
  /* THE DAY THIS SHEET IS ABOUT, which is the draft's own and is yesterday for
     one carried over midnight - not "today", which is what filed an evening's
     work under the wrong date */
  const day = sheetDay(), w = ST.isoWeek(day);
  const carried = day !== dayToday();
  const head = '<div class="picker daysheet">' +
    '<div class="pickh">End of day</div>' +
    '<div class="picksub">' + esc(PERSON.name + " · " + w.weekday + " " + day +
      " · " + stageWords()) + '</div>' +
    (carried ? '<div class="daycarry">This is the sheet you started on ' + esc(w.weekday) +
       ', and it saves under that day.</div>' : "");
  const back = '<button class="pcancel" data-dayback="1">Back to the board</button></div>';
  const saved = daySheetSaved(day), owed = daySheetOwed(day);

  /* already done: the sheet as it was saved, and the one sentence that says
     what to do about a mistake */
  if (saved || owed) {
    const row = saved || ST.dayRows([{ id: "", fields: owed.fields }], counts)[0];
    /* the LOCAL clock (ST.stClock), not a slice of the ISO stamp: a sheet
       saved at 12:28 read "11:28" here all summer */
    const when = saved ? "saved " + ST.stClock(saved.savedAt) + " — " + ST.DAY_SAVED_WORDS
                       : owedWords(owed);
    return head +
      '<div class="dayread">' +
        counts.map(c => '<div class="dayrow"><span class="dayl">' + esc(c[1]) + '</span>' +
          '<span class="dayv tab">' + (row.counts[c[0]] || 0) + '</span></div>').join("") +
        '<div class="dayrow"><span class="dayl">Total</span>' +
          '<span class="dayv tab">' + row.total + '</span></div>' +
        (row.note ? '<div class="daynote">' + esc(row.note) + '</div>' : "") +
        '<div class="daysaid' + (owed ? " owed" : "") +
          (owed && Number(owed.refused) >= DAY_REFUSE_MAX ? " bad" : "") + '">' + esc(when) + '</div>' +
      '</div>' + dayHistoryHtml(counts) + back;
  }

  const d = draftNow();
  const wk = weekWords(draftTotal());
  return head +
    (DAY_OK === false ? '<div class="cphint">' + esc(DAY_WHY) + '</div>' : "") +
    '<div class="dayform">' +
      counts.map(c => '<label class="dayrow"><span class="dayl">' + esc(c[1]) + '</span>' +
        '<input class="daybox tab" type="text" inputmode="numeric" pattern="[0-9]*" ' +
          'data-daycount="' + esc(c[0]) + '" value="' + esc(String(d.counts[c[0]] == null ? "" : d.counts[c[0]])) +
          '" aria-label="' + esc(c[1]) + '"></label>').join("") +
      '<label class="dayrow daynoterow"><span class="dayl">Anything that got in the way</span>' +
        '<textarea class="daytext" data-daynote="1" rows="3" maxlength="' + ST.DAY_NOTE_MAX +
        '" placeholder="Machine down, waiting on glass, helped on another bench…">\n' +
        esc(d.note) + '</textarea></label>' +
      '<div class="daytot">' + esc(carried ? w.weekday : "Today") +
        ': <strong class="tab" id="daytoday">' + draftTotal() + '</strong> ' + esc(unit) + '</div>' +
      '<div class="dayweek">This week: <strong class="tab" id="dayweeknum">' + esc(wk.words) + '</strong></div>' +
      '<div class="daybar"><span id="daybarfill" style="width:' + wk.pct + '%"></span></div>' +
      (DAYBAD ? '<div class="daybad">' + esc(DAYBAD) + '</div>' : "") +
      '<button class="daysave" id="daysave" data-daysave="1"' +
        (draftOk() ? "" : ' disabled aria-disabled="true"') + '>Save the sheet for ' +
        esc(carried ? w.weekday : "today") + '</button>' +
    '</div>' + dayHistoryHtml(counts) + back;
}
/** This person's last seven saved days, read-only, one line each. */
function dayHistoryHtml(counts) {
  const mine = DAYROWS.filter(r => r.who === who()).slice(0, 7);
  if (!mine.length) return "";
  return '<div class="dayhist"><div class="kick">Your last ' + mine.length + ' day' +
    (mine.length === 1 ? "" : "s") + '</div>' +
    mine.map(r => dayLineHtml(r, counts)).join("") + '</div>';
}
/** The sheet's own clicks and keystrokes. The boxes are NOT redrawn as they
    are typed into - only the two totals and the bar are - because a card
    rebuilt on every keystroke is a caret lost on every keystroke. */
function wireDaySheet(host) {
  if (!host || !host.querySelectorAll) return;
  const live = () => {
    const t = draftTotal(), wk = weekWords(t);
    const a = $("#daytoday"); if (a) a.textContent = String(t);
    const b = $("#dayweeknum"); if (b) b.textContent = wk.words;
    const c = $("#daybarfill"); if (c) c.style.width = wk.pct + "%";
    const s = $("#daysave"); if (s) s.disabled = !draftOk();
  };
  host.querySelectorAll("[data-daycount]").forEach(el => el.oninput = () => {
    touch();
    draftNow().counts[el.dataset.daycount] = el.value;
    saveDraft(); live();
  });
  host.querySelectorAll("[data-daynote]").forEach(el => el.oninput = () => {
    touch();
    draftNow().note = String(el.value || "").slice(0, ST.DAY_NOTE_MAX);
    saveDraft();
  });
  host.querySelectorAll("[data-daysave]").forEach(el => el.onclick = () => {
    if (el.disabled) return;
    saveDay();
  });
  host.querySelectorAll("[data-dayback]").forEach(el => el.onclick = () => {
    DAYOPEN = false; DAYBAD = ""; touch(); render();
  });
}
/** The header button. It opens the sheet and asks the two lists once, so the
    first thing somebody sees is their week rather than a blank target. */
function openDaySheet() {
  if (!daySheet() || !PERSON) return;
  DAYOPEN = true; DAYBAD = ""; touch(); render();
  if (DAY_OK !== true) readDay().then(() => { if (DAYOPEN) render(); }, () => {});
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
         (lost ? "!" + lost.stages.map(s => s.stage + s.value).join(",") : "") +
         /* the note channel is not in the list either: opening a composer, a
            note arriving from the other shift and a send that failed all change
            what this card draws and none of them move a counter. The DRAFT is
            deliberately not in the signature - see dressCard. */
         "/" + NOTES.sig(g.job);
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
      (g.finished ? FIN : DOING).appendChild(el);
    });
    if (box && document.activeElement !== box) {
      try {
        box.focus();
        if (box.setSelectionRange && at != null) box.setSelectionRange(at, at);
      } catch (e) {}
    }
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
  const hdr = $("#whois");
  if (hdr) hdr.textContent = PERSON
    ? PERSON.name + (myStages(PERSON).length
        ? " · " + myStages(PERSON).map(ST.stageLabel).join(", ") : " · no stages")
    : "";
  const sw = $("#switchbtn");
  if (sw) { sw.hidden = !PERSON; sw.style.display = PERSON ? "" : "none"; }
  /* the end-of-day button: only on a stage whose definition has a day sheet,
     and only once somebody has said who they are */
  const dy = $("#daybtn");
  if (dy) {
    const on = !!daySheet() && !!PERSON && !PROBLEM;
    dy.hidden = !on;
    dy.style.display = on ? "" : "none";
  }
  /* the search box belongs to the board: there is nothing to search on the
     picker, and a box over an error message only looks broken */
  const sb = $("#search");
  const boarding = !!PAGE_STAGE && !PROBLEM && PEOPLE_READ && !!PERSON && READY && !DAYOPEN;
  if (sb) { sb.hidden = !boarding; sb.style.display = boarding ? "" : "none"; }
  /* the board as it stands, before the box has narrowed it: the number beside
     the box is read off this, and the cards below off the filtered copy */
  const live = boarding ? boardNow() : null;
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
    DOING = null; FIN = null; FINHEAD = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
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
  if (DAYOPEN && daySheet()) {
    DOING = null; FIN = null; FINHEAD = null; NODES = {}; BOARD_PREV = null; QSIG = {}; PSIG = "";
    host.innerHTML = daySheetHtml();
    wireDaySheet(host);
    return;
  }
  DAYOPEN = false;                       // a stage with no sheet has nothing to show

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

/** One turn of the ten-second clock. A named function rather than a closure
    inside setInterval, so a test can take exactly one turn of it. */
async function tickOnce() {
  render();
  /* a people list that has not answered yet is retried on the same clock: one
     failed read at start-up must not leave the tablet stuck at the picker's
     door until somebody thinks to reload it */
  if (!PEOPLE_READ) await readPeople();
  await pollList();
  /* the note channel keeps its own counsel: it only asks the list anything
     while somebody has a composer open, and never more often than
     ST.COMMENT_POLL_MS. A tablet nobody is writing on sends no request for it
     at all. A failure there is quiet by construction and cannot touch the
     board. */
  if (await NOTES.poll()) render();
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
  const sgb = $("#stagebtn");
  if (sgb) sgb.onclick = () => askStage();
  const dyb = $("#daybtn");
  if (dyb) dyb.onclick = () => openDaySheet();
  /* the box is in the header, outside #board, so typing in it never rebuilds
     the node the caret is in - only the cards under it are redrawn */
  const sb = $("#search");
  if (sb) sb.oninput = () => { QUERY = sb.value || ""; touch(); render(); };
  render();
  await readPeople();
  await readList();
  /* what has already been said about today's jobs, so a second shift does not
     retype the first shift's note. One read, quiet, and unable to fail loudly:
     a missing list is a line inside the composer, never a board taken away. */
  await NOTES.read();
  /* the day sheets and the target, once - and ONLY on a stage whose definition
     has a sheet, so the hotmelting tablet never asks for either list */
  if (daySheet()) await readDay();
  await flushQueue();                            // taps owed from a previous visit
  if (refreshT) clearInterval(refreshT);
  refreshT = setInterval(tickOnce, ST.REFRESH_MS);
  if (peopleT) clearInterval(peopleT);
  peopleT = setInterval(() => {
    readPeople();
    /* the office's target and anybody else's sheets ride the ten-minute clock:
       a day sheet moves once a day and a target less often than that */
    if (daySheet()) readDay().then(() => { if (DAYOPEN) render(); }, () => {});
  }, PEOPLE_MS);
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
