/* Glass station - the pure logic of the first floor station.

   The floor works on a separate page (glass.html) signed in with a shared
   station account that has NO access to the production workbook at all. It
   sees three SharePoint lists in its own site: "Glass station" (the work),
   "Station people" (who may record what) and "Station log" (what was
   recorded, by whom, when). The master dashboard is the only thing that puts
   job facts into the first of those; the floor is the only thing that puts
   counters into it and lines into the log. Neither side ever writes the
   other's columns.

   Nothing in this file touches a DOM, Graph, or a workbook: it is the slice
   the feeder should send, the plan of writes that gets it there, the grouping
   the two boards draw, the clamp on a tap, who is allowed to make it, what the
   log line says, and what changed between two boards. That is what
   test_station.js exercises directly.

   Loaded in the browser by index.html (before app.js) and by glass.html, and
   in Node by the tests, the same way checkpoints.js is.                     */

/* The lists the floor uses, and the site they live in. The site is a small,
   separate SharePoint site so the station account can be a member of it
   WITHOUT being a member of the site that holds the workbook. */
const STATION_LIST = "Glass station";
const PEOPLE_LIST = "Station people";
const LOG_LIST = "Station log";
const STATION_SITE = "FloorStations";
/* Which station these people and log lines belong to. One day there will be a
   second station in the same site, reading the same two lists with a different
   word here - that is the whole reason the Station column exists. */
const STATION_NAME = "Glass";

/* The three things the floor records, in the order they are shown. The keys
   are the board's own; the four maps below say which list columns they mean.
   Every stage applies to every glass unit: there is no unit that skips one. */
const STAGES = [["cut", "Cutting"], ["hotmelt", "Hotmelting"], ["glazed", "Glazing"]];
const STAGE_KEYS = STAGES.map(s => s[0]);
const STAGE_FIELD = { cut: "Cut", hotmelt: "Hotmelt", glazed: "Glazed" };
const STAGE_ROW = { cut: "cut", hotmelt: "hotmelt", glazed: "glazed" };
const STAGE_BY = { cut: "CutBy", hotmelt: "HotmeltBy", glazed: "GlazedBy" };
const STAGE_AT = { cut: "CutAt", hotmelt: "HotmeltAt", glazed: "GlazedAt" };
const stageLabel = k => (STAGES.find(s => s[0] === k) || [k, k])[1];

/* Every column of the list, for the field selection of a read. The feeder
   writes Title and the six job facts plus FedAt/FedBy; the floor writes the
   eleven columns of FLOOR_FIELDS and nothing else, ever. No column here
   carries a phone number, an eircode, a county, a price, a comment or a
   product: the floor is told a job number, a customer name and a number of
   glasses, and nothing else - not even which kinds of glass they are. */
const FEEDER_FIELDS = ["Job", "Customer", "GlassType", "Total", "Seq", "Active"];
/* The floor is told a number of glasses, never a kind of glass. The column
   stays on the list (nothing creates or deletes columns) and every row carries
   the same literal in it, so nothing can read a type back off the list even by
   accident. DG and TG are the two kinds that make up that number: TUFF / NOT
   TUFF describe those same units and are not added, and ARCH, ASTRAGAL, FANCY
   and EXTRA are not glasses the floor cuts, hotmelts and glazes. */
const GLASS_TYPE = "GLASS";
const TOTAL_TYPES = ["DG", "TG"];
/* The floor's whitelist: the three counters, a By/At pair for each of them,
   and the last-touch pair. A PATCH from the tablet is filtered to this list on
   the way into the queue and again on the way out of it. */
const FLOOR_FIELDS = ["Cut", "Hotmelt", "Glazed",
                      "CutBy", "CutAt", "HotmeltBy", "HotmeltAt", "GlazedBy", "GlazedAt",
                      "DoneBy", "DoneAt"];
const STATION_FIELDS = ["Title"].concat(FEEDER_FIELDS, ["FedAt", "FedBy"], FLOOR_FIELDS);
/* The three counters, and the ONE exception to "the feeder never writes the
   floor's columns": a row the feeder is creating, or a row the floor has never
   tapped (DoneAt empty), starts from what the office has already ticked off in
   the workbook - otherwise a job the office finished last week arrives on the
   tablet reading nothing done. The moment the floor touches a row, DoneAt is
   set and the feeder never writes a counter on it again. The By/At pairs and
   the last-touch pair are NEVER the feeder's, seeded or not: they say who did
   it, and the office did not. */
const SEED_FIELDS = ["Cut", "Hotmelt", "Glazed"];
const FEEDER_WRITES = ["Title"].concat(FEEDER_FIELDS, ["FedAt", "FedBy"], SEED_FIELDS);
const PEOPLE_FIELDS = ["Title", "Station", "Stages", "PIN", "Active"];
/* The office needs the names and the stages for the log window's filters, and
   has no business reading anybody's PIN: it asks for four of the five. */
const PEOPLE_FIELDS_OFFICE = ["Title", "Station", "Stages", "Active"];
const LOG_FIELDS = ["Title", "Station", "GlassType", "Stage", "From", "To", "Who", "At"];
const CUSTOMER_MAX = 70;

/* One tablet is passed between three people, so the name that is chosen has to
   stop meaning anything by itself. Ten minutes without a tap and the picker
   comes back. One constant, easy to change after the demo. */
const PERSON_LOCK_MS = 600000;
/* How often each screen looks for someone else's changes. Ten seconds is the
   number the owner asked for: a tap on the floor is on the office screen
   before anybody has finished saying what they just did. */
const REFRESH_MS = 10000;
/* How far back the office reads the log to begin with. The list is never
   deleted from, so after a year it is thousands of lines of last spring: the
   window and the drawer both only ever look at recent work. */
const LOG_DAYS = 90;
/** The ISO stamp LOG_DAYS ago - the front of the log worth reading. */
function logSince(days, now) {
  const d = stNum(days, LOG_DAYS);
  const t = stNum(now, Date.now());
  return new Date(t - d * 86400000).toISOString();
}

/* ---- small helpers -------------------------------------------------------- */
const stTxt = v => String(v == null ? "" : v);
const stKey = v => stTxt(v).trim().toUpperCase();
function stNum(v, dflt) {
  if (v == null || stTxt(v).trim() === "") return dflt;
  const n = Number(v);
  return isFinite(n) ? n : dflt;
}
const stClamp = (n, total) => Math.max(0, Math.min(total, Math.round(stNum(n, 0))));
/** SharePoint item ids are numbers in a string. The oldest is the smallest. */
function stItemAge(a, b) {
  const na = Number(a && a.id), nb = Number(b && b.id);
  if (isFinite(na) && isFinite(nb)) return na - nb;
  const sa = stTxt(a && a.id), sb = stTxt(b && b.id);
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/** "In production" for the list's Active column. The section names come from
    the sheet's own divider rows (parser.js), so this matches on the prefix and
    ignores case: Ready to fit, Collect & supply only and everything else are
    not in production, and neither is a job that has left the sheet (cat past). */
function inProduction(j, blockNames) {
  if (!j || j.cat === "past") return false;
  const name = stTxt((blockNames || [])[j.blk]);
  return /^\s*in production/i.test(name);
}

/* Rows sort by the office's own order first, so the floor sees the office's
   order, and by Title after it, so one slice always makes one plan. */
function stRowOrder(a, b) {
  const d = stNum(a.seq, 99999) - stNum(b.seq, 99999);
  if (d) return d;
  return a.title < b.title ? -1 : a.title > b.title ? 1 : 0;
}

/** How many glasses a job has, as the floor counts them: DG plus TG, and
    nothing else. The sheet's other glass columns describe those same units
    (TUFF / NOT TUFF) or are not glass the floor works on (ARCH, ASTRAGAL,
    FANCY, EXTRA), so adding any of them would send the floor to cut sheets
    that do not exist. */
function glassTotal(j) {
  const g = (j && j.glass) || {};
  let n = 0;
  Object.keys(g).forEach(k => {
    if (TOTAL_TYPES.indexOf(stKey(k)) < 0) return;
    n += Math.round(stNum(g[k], 0));
  });
  return Math.max(0, n);
}

/* ---- what the office has already ticked off --------------------------------
   The office ticks glass off job by job in the workbook's own colours long
   before a tablet appears on the floor, and a job that was finished in the
   office last week must not arrive on the tablet reading nothing done. That
   record is the glass checkpoints, and this turns it into the three counters
   the floor's row starts at.

   `counts` is plain data - one entry per glass item of the job, `{type, total,
   status, done}` where status is the sheet's own word ("done", "process",
   "cut" or nothing) and done is the office's count where it has one. It is
   built by whoever has checkpoints.js to hand (app.js), so nothing in this
   file has to know about a job object or a global.                          */
function officeSeed(row, counts) {
  const total = Math.max(0, Math.round(stNum(row && row.total, 0)));
  const none = { cut: 0, hotmelt: 0, glazed: 0 };
  if (!(total > 0)) return none;
  const mine = (counts || []).filter(c => c && TOTAL_TYPES.indexOf(stKey(c.type)) >= 0);
  if (!mine.length) return none;
  const st = c => stTxt(c.status).trim().toLowerCase();
  const all = n => ({ cut: n, hotmelt: n, glazed: n });
  /* gold on every one of them: the office says this job's glass is done, so
     the floor's row is done too */
  if (mine.every(c => st(c) === "done")) return all(total);
  /* part way there: the office's own count, a gold item counting its whole
     quantity. All three stages start there - the office does not record which
     stage its count belongs to, and starting them level is the reading that
     asks the floor for the least re-typing. */
  if (mine.some(c => st(c) === "done" || st(c) === "process")) {
    let done = 0;
    mine.forEach(c => {
      const t = Math.max(0, Math.round(stNum(c.total, 0)));
      done += st(c) === "done" ? t : Math.max(0, Math.round(stNum(c.done, 0)));
    });
    return all(stClamp(done, total));
  }
  /* the sheet's Cut green and nothing else: cut, and only cut */
  if (mine.some(c => st(c) === "cut")) return { cut: total, hotmelt: 0, glazed: 0 };
  return none;
}

/* ---- what the feeder should be sending ------------------------------------
   ONE row per job, for jobs that are IN PRODUCTION and have glass. The floor
   is told a job number, a customer name and a number of glasses; the kinds of
   glass are the office's business and stay in the office. Nothing else: the
   floor's list is today's work, not a copy of the workbook, so a job that has
   been delivered or was never on the floor has no business being sent at all.
   A job that leaves production is not removed from the slice's world -
   feedPlan marks the item it already made as Active = No - so the list settles
   at the floor's jobs plus a short tail of finished ones. Every row here
   therefore carries active:true; the field stays because it is what becomes
   the Active column, and because feedPlan reads it.

   `countsOf` is optional and is how the office's own record reaches the seed:
   a function from the job to its glass counts, or a map of them by job number.
   Without it every row seeds at nothing, which is what a caller with no
   checkpoints to hand should send.                                          */
function glassSlice(jobs, blockNames, countsOf) {
  const names = blockNames || (jobs && jobs.blockNames) || [];
  const look = j => {
    if (typeof countsOf === "function") return countsOf(j) || [];
    if (countsOf && typeof countsOf === "object") return countsOf[stKey(j.id)] || [];
    return [];
  };
  const out = [];
  (jobs || []).forEach(j => {
    if (!j || !j.id) return;
    const job = stKey(j.id);
    if (!job) return;
    const active = inProduction(j, names);
    if (!active) return;                    // not on the floor: not the floor's business
    const total = glassTotal(j);
    if (!(total > 0)) return;               // no glass, nothing for the glass station to do
    const row = { title: job, job: job, customer: stTxt(j.cust).trim().slice(0, CUSTOMER_MAX),
                  total: total, seq: stNum(j.seq, 99999), active: active };
    row.seed = officeSeed(row, look(j));
    out.push(row);
  });
  out.sort(stRowOrder);
  return out;
}

/** The six job facts of one slice row, in list shape. */
function feederFields(row) {
  return { Job: row.job, Customer: row.customer, GlassType: GLASS_TYPE,
           Total: row.total, Seq: row.seq, Active: row.active ? "Yes" : "No" };
}
/** The three counters of a slice row's seed, in list shape. */
function seedFields(row) {
  const s = (row && row.seed) || {};
  const out = {};
  STAGE_KEYS.forEach(k => { out[STAGE_FIELD[k]] = stClamp(s[k], stNum(row && row.total, 0)); });
  return out;
}
/** Has the floor ever tapped this row? DoneAt is set by the first tap and by
    nothing else, so an empty one means the row is still the office's to seed. */
const untouched = fields => stTxt(fields && fields.DoneAt).trim() === "";
/** Is the value already in the list the same as the one the feeder wants?
    Numbers are compared as numbers (SharePoint may answer 6 or "6"); a missing
    or blank cell is never "the same as" anything, so it gets written. */
function sameField(have, want) {
  if (typeof want === "number") {
    if (have == null || stTxt(have).trim() === "") return false;
    return Number(have) === want;
  }
  return stTxt(have) === stTxt(want);
}

/* ---- the plan --------------------------------------------------------------
   HARD RULE, and the reason this is a separate pure function whose output the
   tests read: the plan never contains a delete, and the only floor columns it
   can ever name are the three counters - on a row it is creating, or on a row
   whose DoneAt is still empty. A By, an At or the last-touch pair is never in
   a plan at all, and a row the floor has tapped keeps every number on it. A
   job that leaves production is marked Active = No; its counters, and whoever
   recorded them, stay where they are.

   Duplicates: two items carrying one Title should not happen (Title is unique
   on the list and the feeder is its only writer) but if they do, the oldest is
   the one the feeder maintains and the boards read - exactly the rule the
   phases list already uses - and the others are left alone rather than
   deleted. The one exception is going inactive, which every item of that Title
   is told, so a stray duplicate cannot keep a finished job on the floor's
   board.                                                                     */
function feedPlan(slice, items, opts) {
  const at = (opts && opts.at) || new Date().toISOString();
  const by = stTxt(opts && opts.by);
  const rows = (slice || []).slice().sort(stRowOrder);

  const byTitle = {};
  (items || []).forEach(it => {
    if (!it) return;
    const t = stKey((it.fields || {}).Title);
    if (!t) return;
    (byTitle[t] = byTitle[t] || []).push(it);
  });
  Object.keys(byTitle).forEach(t => byTitle[t].sort(stItemAge));

  const adds = [], todo = [], seen = {};
  let unchanged = 0;

  rows.forEach(r => {
    const t = stKey(r.title);
    seen[t] = 1;
    const want = feederFields(r);
    const mine = byTitle[t];
    if (!mine || !mine.length) {
      /* a row this feeder is creating starts where the office's own record
         already is - that is the whole of the seeding rule on a new row */
      adds.push(Object.assign({ Title: r.title }, want, seedFields(r), { FedAt: at, FedBy: by }));
      return;
    }
    const have = mine[0].fields || {};
    const diff = {};
    FEEDER_FIELDS.forEach(k => { if (!sameField(have[k], want[k])) diff[k] = want[k]; });
    /* and a row the floor has never tapped is still the office's to say. Once
       DoneAt is set - the first tap writes it - this is skipped for ever, so
       nothing the office does can walk over the floor's own count. */
    if (untouched(have)) {
      const seed = seedFields(r);
      /* a blank counter IS nought here, unlike a blank job fact: a row nobody
         has ever written a number on does not need three zeros put in it */
      SEED_FIELDS.forEach(k => { if (stNum(have[k], 0) !== seed[k]) diff[k] = seed[k]; });
    }
    if (!Object.keys(diff).length) { unchanged++; return; }
    /* FedAt/FedBy say when this item was last brought up to date and by whose
       dashboard - the FedAt column's whole meaning - so they ride along with
       every change, and only with a change. */
    diff.FedAt = at; diff.FedBy = by;
    todo.push({ seq: r.seq, title: r.title, id: stTxt(mine[0].id), fields: diff });
  });

  /* everything the sheet no longer has: marked inactive, never removed */
  Object.keys(byTitle).forEach(t => {
    if (seen[t]) return;
    byTitle[t].forEach(it => {
      const have = it.fields || {};
      if (stTxt(have.Active).trim().toLowerCase() === "no") { unchanged++; return; }
      todo.push({ seq: stNum(have.Seq, 99999), title: t, id: stTxt(it.id),
                  fields: { Active: "No", FedAt: at, FedBy: by } });
    });
  });

  todo.sort(stRowOrder);
  return { adds: adds, patches: todo.map(x => ({ id: x.id, fields: x.fields })), unchanged: unchanged };
}

/** A short, stable fingerprint of the slice. The feeder skips a run when it
    matches the last one and that run is under ten minutes old, so a dashboard
    left open on a quiet afternoon does not re-read the list every refresh.
    JSON.stringify does the separating, so a customer name carrying a bar or a
    comma cannot make two different slices hash to the same string. The seed is
    in it as well: a job the office has just ticked off has to reach the floor
    on the next load, not in ten minutes' time. */
function sliceHash(slice) {
  const rows = (slice || []).slice().sort(stRowOrder);
  let s = "";
  rows.forEach(r => {
    const seed = r.seed || {};
    s += JSON.stringify([r.title, r.job, r.customer, r.total, r.seq, !!r.active,
                         seed.cut || 0, seed.hotmelt || 0, seed.glazed || 0]) + "\n";
  });
  let h = 0x811c9dc5;                                     // FNV-1a, 32 bit
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return rows.length + "-" + h.toString(16);
}

/* ---- what the screens draw --------------------------------------------------
   One card per job, because there is one list row per job. `keep` decides
   which rows are in: the two boards take the active ones, the office drawer
   takes one job's row whether it is still on the floor or not. Counters are
   clamped for display: if the office shortens a job from 8 glasses to 4 after
   the floor has cut 6, the card reads 4 of 4 rather than 150%. The clamp is
   display only - nothing here writes a corrected number back anywhere.

   Each stage also carries the who and the when of the last change to it, which
   is what the office reads out as "Cutting 8 of 8 - Person A - Tue 14:02".

   A job is finished, and goes gold on both screens, when all three counters
   have reached the total and there is a total to reach: a row with no glasses
   on it is not something anybody finished.                                  */
const isActive = f => stTxt(f && f.Active).trim().toLowerCase() === "yes";

/** The job a list row is about, whichever iteration wrote it: v3 writes the
    job number as the Title, v2 wrote `JOB|TYPE`. */
const rowJob = f => stKey(f && f.Job) || stKey(f && f.Title).split("|")[0];
/** Is this row the v3 one - the row whose Title IS the job number? */
const titleIsJob = f => !!stKey(f && f.Title) && stKey(f && f.Title) === rowJob(f);

function buildJobs(items, keep) {
  const best = {};
  (items || []).forEach(it => {
    if (!it) return;
    const f = it.fields || {};
    if (!keep(f)) return;
    const job = rowJob(f);
    if (!job) return;
    /* One card per JOB, not per row. The live list still holds the v2 rows -
       one per job AND glass type, Title `JOB|TYPE` - until they are cleared,
       and a job must not appear twice on the board or be called finished
       because some old ARCH row happens to be full. The row whose Title is
       the job number is the one this iteration maintains, so it wins; between
       two of the same kind the oldest id wins, as everywhere else. */
    const prev = best[job];
    if (prev) {
      const pf = prev.fields || {};
      if (titleIsJob(pf) !== titleIsJob(f)) { if (titleIsJob(pf)) return; }
      else if (stItemAge(prev, it) <= 0) return;
    }
    best[job] = it;
  });

  const out = Object.keys(best).sort().map(t => {
    const it = best[t], f = it.fields || {};
    const total = Math.max(0, Math.round(stNum(f.Total, 0)));
    const g = { id: stTxt(it.id), job: stKey(f.Job) || t, customer: stTxt(f.Customer),
                seq: stNum(f.Seq, 99999), fedAt: stTxt(f.FedAt), active: isActive(f),
                total: total, by: {}, at: {}, bars: {} };
    STAGE_KEYS.forEach(k => {
      g[STAGE_ROW[k]] = stClamp(f[STAGE_FIELD[k]], total);
      g.by[k] = stTxt(f[STAGE_BY[k]]);
      g.at[k] = stTxt(f[STAGE_AT[k]]);
      g.bars[k] = { done: g[STAGE_ROW[k]], total: total, by: g.by[k], at: g.at[k] };
    });
    g.finished = total > 0 && STAGE_KEYS.every(k => g[STAGE_ROW[k]] >= total);
    return g;
  });
  out.sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
  return out;
}

/** "12 glasses", "1 glass" - the one number the floor is given about a job,
    said the same way on the tablet and on the office's board. */
const glassWords = n => Math.max(0, Math.round(stNum(n, 0))) +
  (Math.round(stNum(n, 0)) === 1 ? " glass" : " glasses");

/** "28 left" - the same number after it has started counting down, which has
    to say so. The tablet's number is how much work the person signed in has
    left, not the size of the job and not even a count of glasses (somebody
    who holds two stages has two jobs of work on each glass), so a bare "28"
    beside a customer's name on a workshop screen would be read as the size of
    the job it used to mean. "left" is true of both; "glasses" would be false
    for the two-stage person. The office keeps glassWords: there, the number
    really is how big the job is. */
const leftWords = n => Math.max(0, Math.round(stNum(n, 0))) + " left";

/** How much of one job is left for ONE person: every stage they hold, counted
    separately and added up.

    A person who cuts AND hotmelts a 14-glass job has 28 things left to do on
    it, not 14, and every single tap of either stage takes one off. The other
    reading - the smallest of their counters, so the number means glasses
    rather than jobs of work - was tried and rejected by the owner: it lets
    somebody cut six and watch the big number on the wall sit still, because
    hotmelt had not caught up. A number that does not move while somebody
    works is worse than a number that counts something slightly abstract.

    Somebody who holds no stage at all has the whole job left: they cannot
    move any of it, and saying nought would read as "done". stationPeople has
    already dropped any word in the Stages column that is not a stage, so a
    person whose column was all nonsense arrives here holding nothing and gets
    that same answer.

    Two people therefore see two different numbers for the same job, and
    neither of them is the office's `Glass 12/24`. That is what was asked for -
    the tablet answers "how much have I left", not "how big is this job".

    Note what this is NOT: it does not decide gold. A job goes gold, and folds
    into Finished, when it is complete for everybody (buildJobs' `finished`),
    never when it reaches nought for whoever is holding the tablet.          */
function jobLeftFor(g, stages) {
  const total = Math.max(0, Math.round(stNum(g && g.total, 0)));
  const seen = {};
  let held = 0, left = 0;
  (stages || []).forEach(s => {
    const k = stTxt(s).trim().toLowerCase();
    /* a Stages column typed by hand can say anything; a word that is not one
       of the three stages is not a stage this person holds, and must not be
       the one place on the tablet that throws */
    if (STAGE_KEYS.indexOf(k) < 0) return;
    /* "cut, cut" is one stage held, not two. stationPeople already drops the
       repeat, so this only matters if something else ever calls in - but a
       number on a wall that doubles because a column was typed twice is not a
       thing to leave to another function's care */
    if (seen[k]) return;
    seen[k] = true;
    held++;
    /* the clamp is what keeps this in 0...total per stage, so a row saying
       -9 or 99 or nothing at all cannot make the wall read a negative */
    left += total - stClamp(g && g[STAGE_ROW[k]], total);
  });
  return held ? left : total;
}

/** The same sum over the whole board: everything the person signed in still
    has to do. It is given the whole board and never the searched one -
    somebody looking a job number up must not make their own work read smaller
    than it is, which is the one way this number could tell a lie. */
function boardLeftFor(board, stages) {
  return (board || []).reduce((sum, g) => sum + (g ? jobLeftFor(g, stages) : 0), 0);
}

/** The cards somebody typing in the tablet's search box is looking for: a job
    number or a customer name, matched anywhere in either. An empty box is
    every card - the box narrows the board, it never becomes the board. */
function boardFilter(board, q) {
  const want = stTxt(q).trim().toLowerCase();
  if (!want) return (board || []).slice();
  return (board || []).filter(g =>
    (g.job + " " + g.customer).toLowerCase().indexOf(want) >= 0);
}

/** The floor's board: the jobs that are on the floor now. */
function jobBoard(items) { return buildJobs(items, isActive); }

/** One job's record whatever its Active is, or null. The office drawer uses
    this rather than the board: a job that has been delivered is off the
    floor's screen, but what the floor recorded on it is still worth reading,
    and "not fed to the floor yet" would be a lie about it. `active` says which
    of the two it is. */
function jobRecord(items, job) {
  const want = stKey(job);
  if (!want) return null;
  /* rowJob reads a v2 `JOB|TYPE` Title as well as a v3 one, and buildJobs
     prefers the v3 row, so a leftover old row can neither become the drawer's
     record nor hide the real one */
  const mine = buildJobs(items, f => rowJob(f) === want);
  return mine[0] || null;
}

/** The same records, all of them at once, as a map from job number to record.
    jobRecord() walks the whole list to answer about one job, which is right
    for a drawer asking once and wrong for the office's job list, which asks
    about every row it draws. One pass, the same de-duplication, the same
    records - the caller keeps the map for as long as the list it was built
    from is the same array. Reads only; it writes nothing anywhere. */
function jobRecords(items) {
  const out = {};
  buildJobs(items, () => true).forEach(g => { out[g.job] = g; });
  return out;
}

/** The new value of one counter after a tap. delta is a number, "all" or
    "none". The only rule is the clamp: the floor may record glazing before it
    records hotmelt, and nothing here stops them - the counters are a record of
    what happened, not a workflow to be enforced. */
function applyTap(row, stage, delta) {
  const field = STAGE_ROW[stage];
  if (!field) return null;
  const total = Math.max(0, Math.round(stNum(row && row.total, 0)));
  const now = stClamp(row && row[field], total);
  if (delta === "all") return total;
  if (delta === "none") return 0;
  const d = Number(delta);
  if (!isFinite(d)) return now;
  return stClamp(now + d, total);
}

/* ---- who is allowed to record what ------------------------------------------
   One shared account, three people. The name that is picked is not a login: it
   decides which stages the steppers will move, and it is what goes into the
   log. The PIN is compared here, on the tablet, against a column the station
   account can read - it is a deterrent on a device that is passed around a
   workshop, not a secret, and STATIONS.md says so in those words.           */
function stationPeople(items, station) {
  const want = stTxt(station || STATION_NAME).trim().toLowerCase();
  const out = [];
  (items || []).forEach(it => {
    if (!it) return;
    const f = it.fields || {};
    const name = stTxt(f.Title).trim();
    if (!name) return;
    if (stTxt(f.Station).trim().toLowerCase() !== want) return;
    if (stTxt(f.Active).trim().toLowerCase() !== "yes") return;
    /* "hotmelt, cut" and "Hotmelt,Cut" are the same two stages: the owner
       types this column by hand, so it is read forgivingly and filtered down
       to stages that actually exist */
    const seen = {};
    const stages = stTxt(f.Stages).split(/[,;/]+/).map(s => s.trim().toLowerCase())
                     .filter(s => STAGE_KEYS.indexOf(s) >= 0)
                     .filter(s => (seen[s] ? false : (seen[s] = true)));
    out.push({ id: stTxt(it.id), name: name, stages: stages,
               pin: stTxt(f.PIN).trim(), station: stTxt(f.Station).trim() });
  });
  out.sort((a, b) => (a.name.toLowerCase() < b.name.toLowerCase() ? -1
                    : a.name.toLowerCase() > b.name.toLowerCase() ? 1 : 0));
  return out;
}
/** May this person move this stage? A person holding no stages at all may move
    nothing: every stepper is greyed and tap() refuses every one of them. */
function canStage(person, stage) {
  if (!person || !stage) return false;
  return (person.stages || []).indexOf(String(stage).toLowerCase()) >= 0;
}
/** An empty PIN column means this person is not asked for one. Anything else
    must match exactly - the digits typed on the pad, compared as text. */
function pinOk(person, entered) {
  const want = stTxt(person && person.pin).trim();
  if (!want) return true;
  return want === stTxt(entered).trim();
}
/** Has the chosen name gone stale? The tablet is passed around, so a name that
    has not tapped anything for PERSON_LOCK_MS stops meaning anybody. */
function personExpired(lastTapAt, now, lockMs) {
  const t = stNum(lastTapAt, 0);
  if (!t) return true;
  return (stNum(now, Date.now()) - t) >= stNum(lockMs, PERSON_LOCK_MS);
}

/* ---- the writes one tap makes ----------------------------------------------
   Two of them, in order: the counter, and then - only if that succeeded - the
   log line. The counter PATCH carries that stage's number, that stage's By and
   At, and the last-touch pair. Nothing else can get into it: floorOnly() is
   the same filter the tablet's queue runs on the way in and on the way out, so
   a job fact typed into localStorage by somebody holding the tablet still
   cannot reach the list.                                                    */
function floorOnly(fields) {
  const out = {};
  FLOOR_FIELDS.forEach(k => {
    if (!fields || !(k in fields)) return;
    const v = fields[k];
    if (STAGE_KEYS.some(s => STAGE_FIELD[s] === k)) {
      if (typeof v === "number" && isFinite(v)) out[k] = v;   // a counter is a number
      return;
    }
    if (typeof v === "string" && v) out[k] = v;               // a By or an At is text
  });
  return out;
}
/** The counter PATCH for one tap: that stage's number, that stage's By/At, and
    the last-touch pair. Never anything else. */
function tapFields(stage, value, who, at) {
  if (!STAGE_FIELD[stage]) return null;
  const when = stTxt(at) || new Date().toISOString();
  const name = stTxt(who);
  const out = {};
  out[STAGE_FIELD[stage]] = Math.round(stNum(value, 0));
  out[STAGE_BY[stage]] = name;
  out[STAGE_AT[stage]] = when;
  out.DoneBy = name;
  out.DoneAt = when;
  return out;
}
/** One Station log line. From is the value last known to be in the list, To is
    the value sent - so a run of quick taps that the queue merged into one
    write is one line saying 5 to 8, not three lines saying 5, 6, 7, 8.
    GlassType is the same literal every row of the board carries: the column is
    still on the list, and there are no glass types on the floor to put in it. */
function logFields(e) {
  return { Title: stKey(e && e.job), Station: stTxt((e && e.station) || STATION_NAME),
           GlassType: stKey((e && e.type) || GLASS_TYPE), Stage: stTxt(e && e.stage).toLowerCase(),
           From: Math.round(stNum(e && e.from, 0)), To: Math.round(stNum(e && e.to, 0)),
           Who: stTxt(e && e.who), At: stTxt(e && e.at) || new Date().toISOString() };
}

/* ---- reading the log back --------------------------------------------------
   The office reads this list and never writes it; nothing anywhere deletes
   from it. Rows come back newest first, because that is the only order anybody
   ever wants to read a log in.                                              */
function logRows(items, station) {
  const want = stTxt(station || STATION_NAME).trim().toLowerCase();
  const out = [];
  (items || []).forEach(it => {
    if (!it) return;
    const f = it.fields || {};
    const job = stKey(f.Title);
    if (!job) return;
    const st = stTxt(f.Station).trim();
    if (st && st.toLowerCase() !== want) return;      // another station's line
    out.push({ id: stTxt(it.id), job: job, station: st || STATION_NAME,
               type: stKey(f.GlassType), stage: stTxt(f.Stage).trim().toLowerCase(),
               from: Math.round(stNum(f.From, 0)), to: Math.round(stNum(f.To, 0)),
               who: stTxt(f.Who), at: stTxt(f.At) });
  });
  /* ISO stamps sort as text; the item id breaks a tie, because two taps inside
     one second are perfectly possible on a tablet */
  out.sort((a, b) => (a.at < b.at ? 1 : a.at > b.at ? -1 : -stItemAge(a, b)));
  return out;
}
/** The log window's four filters. A missing filter matches everything; `day`
    is a YYYY-MM-DD and matches the front of the ISO stamp.

    `job` matches as a substring, because the log window's box is somebody
    typing "530" to find a job - but with `exact: true` it must be the whole
    job number, which is what the drawer needs: R530 is not R5303, and a
    drawer that quietly listed another job's work would be a lie about it. */
function logFilter(rows, f) {
  const who = stTxt(f && f.who).trim().toLowerCase();
  const stage = stTxt(f && f.stage).trim().toLowerCase();
  const job = stKey(f && f.job);
  const day = stTxt(f && f.day).trim();
  const exact = !!(f && f.exact);
  return (rows || []).filter(r =>
    (!who || r.who.toLowerCase() === who) &&
    (!stage || r.stage === stage) &&
    (!job || (exact ? r.job === job : r.job.indexOf(job) >= 0)) &&
    (!day || r.at.slice(0, 10) === day));
}
/** How many lines, and how many units, per person and per stage - the count
    that sits at the top of the log window for whatever is filtered. */
function logCounts(rows) {
  const grab = keyOf => {
    const at = {};
    (rows || []).forEach(r => {
      const k = keyOf(r);
      if (!k) return;
      const c = at[k] || (at[k] = { key: k, lines: 0, units: 0 });
      c.lines++;
      const d = r.to - r.from;
      if (d > 0) c.units += d;
    });
    return Object.keys(at).map(k => at[k])
             .sort((a, b) => (b.units - a.units) || (b.lines - a.lines) ||
                             (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
  };
  return { people: grab(r => r.who), stages: grab(r => r.stage) };
}
/** The newest line for one job, for the board card's one-line "last: ...". */
function logLast(rows, job) {
  const want = stKey(job);
  return (rows || []).find(r => r.job === want) || null;
}

/* ---- keeping two boards in step --------------------------------------------
   Ten-second polling means the board is rebuilt six times a minute. Throwing
   the whole thing away each time loses an open card, the scroll position, and
   the tap somebody's finger is already on its way to. boardDiff says exactly
   which cards moved, so only those nodes are touched.

   The signature is everything a card actually draws. Anything not in it -
   another job's counters, a FedAt that did not change the words on screen -
   cannot make a card redraw.                                                */
function cardSig(g) {
  return JSON.stringify([g.id, g.job, g.customer, g.seq, g.total, g.finished,
    STAGE_KEYS.map(k => [g[STAGE_ROW[k]], g.by[k], g.at[k]])]);
}
function boardDiff(prev, next) {
  const was = {}, now = {};
  (prev || []).forEach(g => { was[g.job] = cardSig(g); });
  (next || []).forEach(g => { now[g.job] = cardSig(g); });
  const added = [], changed = [], removed = [];
  Object.keys(now).forEach(j => {
    if (!(j in was)) added.push(j);
    else if (was[j] !== now[j]) changed.push(j);
  });
  Object.keys(was).forEach(j => { if (!(j in now)) removed.push(j); });
  /* the finished group is a different place on the screen, so a card becoming
     finished is a move as well as a change */
  const order = list => list.map(g => g.job + (g.finished ? "!" : "")).join(",");
  return { added: added.sort(), changed: changed.sort(), removed: removed.sort(),
           order: order(prev || []) !== order(next || []) };
}

/* ---- delta ------------------------------------------------------------------
   Graph's delta feed answers only what moved since the last token. The same
   item can appear twice in one feed (it changed twice) and the LAST occurrence
   is the true one; a deleted item arrives carrying "@removed". Merging is pure
   and therefore testable, which is the point of it living here rather than in
   station.js: the fiddly part of a ten-second poll is not the request.      */
function mergeDelta(items, changes) {
  const out = (items || []).filter(Boolean).map(it => ({ id: stTxt(it.id), fields: it.fields || {} }));
  const at = {};
  out.forEach((it, i) => { at[it.id] = i; });
  (changes || []).forEach(c => {
    if (!c || c.id == null) return;
    const id = stTxt(c.id);
    if (c.removed) {
      if (id in at) { out[at[id]] = null; delete at[id]; }
      return;
    }
    /* Graph can answer a change with no fields bag at all (a column the
       select did not ask for moved, say). Overwriting a row with {} would
       blank a job off the board over nothing, so an empty change keeps what
       is already there and only a real bag replaces it. */
    const fields = c.fields && Object.keys(c.fields).length ? c.fields : null;
    if (id in at) {
      if (fields) out[at[id]] = { id: id, fields: fields };
      return;                                         // the last occurrence wins
    }
    at[id] = out.length;
    out.push({ id: id, fields: fields || {} });
  });
  return out.filter(Boolean);
}

const ST = {
  STATION_LIST, PEOPLE_LIST, LOG_LIST, STATION_SITE, STATION_NAME,
  STAGES, STAGE_KEYS, STAGE_FIELD, STAGE_ROW, STAGE_BY, STAGE_AT, stageLabel,
  STATION_FIELDS, FEEDER_FIELDS, FLOOR_FIELDS, PEOPLE_FIELDS, LOG_FIELDS,
  SEED_FIELDS, FEEDER_WRITES, GLASS_TYPE, TOTAL_TYPES,
  PEOPLE_FIELDS_OFFICE, CUSTOMER_MAX, PERSON_LOCK_MS, REFRESH_MS, LOG_DAYS, logSince,
  inProduction, glassTotal, officeSeed, glassSlice, feederFields, seedFields, feedPlan, sliceHash,
  jobBoard, jobRecord, jobRecords, jobKey: stKey, boardFilter, glassWords, leftWords,
  jobLeftFor, boardLeftFor, applyTap, boardDiff, mergeDelta,
  stationPeople, canStage, pinOk, personExpired,
  floorOnly, tapFields, logFields, logRows, logFilter, logCounts, logLast
};
if (typeof window !== "undefined") window.ST = ST;
else if (typeof globalThis !== "undefined") globalThis.ST = ST;
if (typeof module !== "undefined" && module.exports) module.exports = ST;
