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

/* The two things the floor records about a glass unit, in the order they are
   shown. The keys are the board's own; the four maps below say which list
   columns they mean. Every stage applies to every glass unit: there is no unit
   that skips one.

   GLAZING LEFT THIS STATION on 2026-09-21 (docs/specs/2026-09-21-glass-split-
   no-glazing.md). The owner's correction: glazing is the last step of the whole
   job, not a glass step, so it stops deciding whether glass is done and gets a
   station of its own later. The `Glazed`, `GlazedBy` and `GlazedAt` columns
   stay on the list with whatever is in them - nothing rewrites, resets or
   deletes them - and nothing reads or writes them any more. */
const STAGES = [["cut", "Cutting"], ["hotmelt", "Hotmelting"]];
const STAGE_KEYS = STAGES.map(s => s[0]);
/* Display only, for history. `Station log` lines written before 2026-09-21
   carry Stage = "glazed", and a log that suddenly reads "glazed" in lower case
   where every other line reads "Cutting" would look broken. This map is read by
   stageLabel and by nothing else: a word in here is not a stage, cannot be
   tapped, cannot be seeded and is not in any whitelist. */
const LEGACY_STAGE_LABELS = { glazed: "Glazing" };
/* Tuff is a third thing the same person counts, added 2026-09-10: a different
   department, the same hands, and its own quantity off the sheet's TUFF column
   rather than the DG + TG one. It is NOT one of the two above, and the
   difference matters in three places - it is not part of the job's glass total,
   it is not one of the counters the feeder may seed, and the office's lock is
   never about it. So the two keep their own list (STAGES / STAGE_KEYS) and all
   three have their own beside it.

   IT DOES DECIDE WHETHER A JOB IS SHOWN AS FINISHED, since the owner's word
   after the demo on 2026-09-21 ("if job has tuff and is not done dont move
   it"). That clause of the 2026-09-10 decision is reversed and no other; see
   tuffOwed(). */
const TUFF_STAGE = "tuff";
const ALL_STAGES = STAGES.concat([[TUFF_STAGE, "Tuff"]]);
const ALL_STAGE_KEYS = ALL_STAGES.map(s => s[0]);
const STAGE_FIELD = { cut: "Cut", hotmelt: "Hotmelt", tuff: "Tuff" };
const STAGE_ROW = { cut: "cut", hotmelt: "hotmelt", tuff: "tuff" };
const STAGE_BY = { cut: "CutBy", hotmelt: "HotmeltBy", tuff: "TuffBy" };
const STAGE_AT = { cut: "CutAt", hotmelt: "HotmeltAt", tuff: "TuffAt" };
/* Which total a stage counts against, on a board record. Two of them count
   glasses; tuff counts tuff units, and clamping it to the glass total would
   read 8 of 8 on a job with eleven tuff and eight glasses. */
const STAGE_TOTAL_ROW = { cut: "total", hotmelt: "total", tuff: "tuffTotal" };
/* a stage's word - falling back to the legacy map, so an old log line still
   reads "Glazing" rather than "glazed" */
const stageLabel = k => (ALL_STAGES.find(s => s[0] === k) || [k, LEGACY_STAGE_LABELS[k] || k])[1];

/* Every column of the list, for the field selection of a read. The feeder
   writes Title and the six job facts plus FedAt/FedBy; the floor writes the
   eleven columns of FLOOR_FIELDS and nothing else, ever. No column here
   carries a phone number, an eircode, a county, a price, a comment or a
   product: the floor is told a job number, a customer name and a number of
   glasses, and nothing else - not even which kinds of glass they are. */
/* TuffTotal and OfficeDone joined these on 2026-09-10. TuffTotal is a job fact
   like Total - the sheet's own TUFF quantity, which the tuff counter needs
   something to count towards. OfficeDone is the office saying "this job's glass
   is finished here", which greys the tablet's steppers; it is the office's own
   column and only the office ever writes it, so it belongs in this list and not
   in the floor's. */
const FEEDER_FIELDS = ["Job", "Customer", "GlassType", "Total", "TuffTotal", "Seq", "Active", "OfficeDone"];
/* The floor is told a number of glasses, never a kind of glass. The column
   stays on the list (nothing creates or deletes columns) and every row carries
   the same literal in it, so nothing can read a type back off the list even by
   accident. DG and TG are the two kinds that make up that number: TUFF / NOT
   TUFF describe those same units and are not added, and ARCH, ASTRAGAL, FANCY
   and EXTRA are not glasses the floor cuts and hotmelts. */
const GLASS_TYPE = "GLASS";
const TOTAL_TYPES = ["DG", "TG"];
/* The floor's whitelist: the three counters, a By/At pair for each of them,
   and the last-touch pair. A PATCH from the tablet is filtered to this list on
   the way into the queue and again on the way out of it.

   Glazed/GlazedBy/GlazedAt left it on 2026-09-21, which is what makes the
   removal safe rather than merely tidy: a `glazed` tap sitting in an old
   tablet's cw_stationq cannot be written, because the filter it runs through on
   the way out of the queue no longer has a column to put it in. */
const FLOOR_FIELDS = ["Cut", "Hotmelt", "Tuff",
                      "CutBy", "CutAt", "HotmeltBy", "HotmeltAt",
                      "TuffBy", "TuffAt", "DoneBy", "DoneAt"];
const STATION_FIELDS = ["Title"].concat(FEEDER_FIELDS, ["FedAt", "FedBy"], FLOOR_FIELDS);
/* The two counters, and the ONE exception to "the feeder never writes the
   floor's columns": a row the feeder is creating, or a row the floor has never
   tapped (DoneAt empty), starts from what the office has already ticked off in
   the workbook - otherwise a job the office finished last week arrives on the
   tablet reading nothing done. The moment the floor touches a row, DoneAt is
   set and the feeder never writes a counter on it again. The By/At pairs and
   the last-touch pair are NEVER the feeder's, seeded or not: they say who did
   it, and the office did not.

   Tuff is deliberately NOT in this list. The seeding exception was given for
   the glass counters, in the v3 spec's "Seeding" section and nowhere else
   (CLAUDE.md rule 3); nobody has widened it, so the tuff counter starts at
   nought on every row and only the floor ever moves it. `Glazed` was in here
   until 2026-09-21 and is not any more: the feeder no longer seeds it and
   leaves whatever is in it. */
const SEED_FIELDS = ["Cut", "Hotmelt"];
const FEEDER_WRITES = ["Title"].concat(FEEDER_FIELDS, ["FedAt", "FedBy"], SEED_FIELDS);

/* ---- a STATION DEFINITION ---------------------------------------------------
   New on 2026-09-16, with the welding station (docs/specs/2026-09-16-welding-
   station.md). Three functions in this file used to know the glass list's own
   column names: feedPlan, sliceHash and floorOnly. They now take a definition
   instead, and GLASS is the definition of the station they used to be about -
   so every existing call site and every one of the 246 glass checks reads
   exactly as it did, and the welding station hands in WELDC.WELD instead.

   Nothing glass-specific was added to this file to make that work, and nothing
   welding-specific may be either: a definition is data, it lives beside its own
   station's rules (welding-core.js), and the next station is one more of them.

   What a definition has to carry:
     key / name / list  what it is called, and the list it feeds
     site               "own" = the workbook's own site, resolved by path;
                        "floor" = the `Floor stations` site, resolved by path.
                        Each is the ONLY site that channel can ever answer.
                        Owner's decision 1 of 2026-09-16. CW.stationSite(which)
                        is what honours it, and the block comment there is the
                        whole of what each word means.
     fields             every column, for the $select of a read
     feederFields       the job facts the feeder owns
     floorFields        the columns the tablet may write, and nothing else
     counterFields      which of those are numbers rather than stamps
     seedFields         the counters the feeder may seed on an untouched row
     feederOf(row)      one slice row's job facts, in list shape
     seedOf(row)        one slice row's seed, in list shape
     hashOf(row)        everything about a row worth re-feeding for            */
const GLASS = {
  key: "glass",
  name: STATION_NAME,
  list: STATION_LIST,
  /* THE PIN (owner's decision 1, 2026-09-16). The Floor stations site did not
     exist when this station shipped, so its three lists were made in the
     workbook's own site. Creating the site for the welding station would
     otherwise have moved this page to a site with no "Glass station" list in
     it at its next ten-minute re-check.

     "own" means THE WORKBOOK'S OWN SITE, resolved by path, with the Floor
     stations site never looked up on that channel at all - not on an empty
     cache, not after a forget, not on Try again. An earlier version of this
     meant "whatever site was cached, else look the old way", which a single
     404 on a list call was enough to defeat; see the graph.js block comment.

     THIS ONE WORD IS THE WHOLE OF THE GLASS MOVE. On the day the three glass
     lists are copied into `Floor stations`, it becomes "floor" and nothing
     else in the app changes. */
  site: "own",
  stages: ALL_STAGE_KEYS,
  stageLabel: stageLabel,
  /* WHICH STAGES GET AN END-OF-DAY SHEET, and what is counted on it
     (2026-09-21). Cutting only for now, by the owner's decision 3; hotmelting
     or welding get one by gaining a line here and nothing else. The keys are
     the list's own column names; the words are what the tablet asks for. */
  /* [column, the question the tablet asks, the heading a table puts over it].
     The short one is there because the office's window has one row per sheet
     and no room for "Other obscure sheets cut" over a column of numbers - and
     without it the four counts read as "12 4 0 2" with nothing to say which is
     which (found in a screenshot, 2026-09-21). dayCountShort falls back to the
     long one, so a station that gives only two is still drawn. */
  daySheets: {
    cut: { counts: [["Clear", "Clear glass sheets cut", "Clear"],
                    ["KGlass", "K-glass sheets cut", "K-glass"],
                    ["Satin", "Satin sheets cut", "Satin"],
                    ["Obscure", "Other obscure sheets cut", "Obscure"]],
           unit: "sheets" }
  },
  /* the stages a report may be run for: the two real ones. Tuff is a counter
     the cutter also moves, not a station somebody reports on. */
  reportStages: STAGE_KEYS,
  /* WHICH `Station log` STAGES FEED A REPORT STAGE (review, 2026-09-21). Here
     they are the same word - a cutting line is logged `Stage = cut` - but they
     are not the same QUESTION, and welding proves it: its tablet logs one line
     per PART (`frames`, `sashes`), so a report of the welding stage that
     matched on "weld" found nothing at all and shipped an empty Days and no
     Activity sheet. Asking the definition is what keeps stationReport free of
     both answers. */
  reportLogStages: k => [stTxt(k).trim().toLowerCase()],
  reportJobs: (data, stage) => glassReportJobs(data, stage),
  fields: STATION_FIELDS,
  feederFields: FEEDER_FIELDS,
  floorFields: FLOOR_FIELDS,
  counterFields: ALL_STAGE_KEYS.map(k => STAGE_FIELD[k]),
  seedFields: SEED_FIELDS,
  feederWrites: FEEDER_WRITES,
  feederOf: row => feederFields(row),
  seedOf: row => seedFields(row),
  hashOf: r => {
    const seed = r.seed || {};
    return [r.title, r.job, r.customer, r.total, r.tuffTotal || 0, r.seq,
            !!r.active, !!r.officeDone,
            seed.cut || 0, seed.hotmelt || 0];
  }
};
/** The definition to use when a caller named none: the glass station, which is
    what every one of these functions was about before there was a second. */
const stDef = d => d || GLASS;

/* The SECOND exception, and the last one: an office CLEAR puts this job's
   counters back to nought (owner, 2026-09-10,
   docs/specs/2026-09-10-office-clears-the-floor.md).

   Why it had to exist. Un-ticking a job's glass in the office cleared the
   workbook correctly every time, but nothing cleared the floor's counters, so
   the tablet stayed gold for ever and the only way back was somebody tapping
   minus forty-nine times. The office, the workbook and the floor now say the
   same thing after a clear, which is the whole of the feature.

   It is ONE named case and nothing more. The office still never writes
   `Station people`, never writes `Station log`, never writes a per-stage By or
   At, and never writes a counter for any other reason - not to correct one,
   not to "tidy up", not from a reconciliation pass. The proof that it cannot
   is in the shape of officeClearFields below: it takes a name and a time, so
   there is no argument it could carry a counter or a per-stage stamp IN. Every
   counter it writes is a literal nought, written here, once. */
const OFFICE_CLEAR_FIELDS = ["Cut", "Hotmelt", "Tuff", "DoneBy", "DoneAt", "OfficeDone"];
/** The whole body of an office clear: three noughts, the last-touch pair, and
    the unlock. `Glazed` was one of the noughts until 2026-09-21; a clear now
    leaves whatever is in it, like everything else on this build (rule 8).

    DoneBy/DoneAt are written, and not writing them would be the bug: the
    last-writer-wins rule reads floorStamp off those fields (spec §3), so
    counters dropped to nought under a stale stamp would look like old news and
    the tablet's "last touch" line would name whoever tapped last rather than
    whoever actually moved the row. The per-stage By/At pairs are NOT written -
    they say who did that stage's WORK, and nobody did.

    OfficeDone = "No" joined them on 2026-09-10, after the owner watched a card
    go correctly black with its counters at nought and then sit GREYED, saying
    the office had marked the job finished, for over a minute. The lock was
    only ever released by the feeder's next run, and the feeder derives it from
    what the master currently shows - so while the master had not caught up it
    kept writing the lock straight back on. Carrying the unlock in the clear's
    own PATCH frees the card on the tablet's next ten-second poll instead.

    It widens nothing. OfficeDone is a FEEDER field - the office's own column,
    which the office already writes and the tablet never can (floorOnly drops
    it). The counters are the only floor columns on this path, and they are
    still the same ones, still only ever nought. */
function officeClearFields(who, at) {
  const out = {};
  ALL_STAGE_KEYS.forEach(k => { out[STAGE_FIELD[k]] = 0; });
  out.DoneBy = stTxt(who);
  out.DoneAt = stTxt(at) || new Date().toISOString();
  out.OfficeDone = "No";
  return out;
}
/** Is there anything on this row for an office clear to clear? Two things have
    to be true, and both of them are about not writing for nothing:

    the floor must really have tapped it (DoneAt), because on a row they never
    touched the counters are the office's own seed echoed back and the FEEDER
    already puts them right on its next run - writing DoneAt on such a row
    would be worse than useless, because it would mark the row "touched" for
    ever and switch its seeding off;

    and at least one counter must be above nought, because writing four noughts
    over four noughts tells nobody anything. */
function floorWorkToClear(g) {
  if (!g || !stTxt(g.doneAt).trim()) return false;
  return ALL_STAGE_KEYS.some(k => Math.round(stNum(g[STAGE_ROW[k]], 0)) > 0);
}
/* the word each counter is said with when the office is asked to destroy it */
const CLEAR_WORD = { cut: "cut", hotmelt: "hotmelted", tuff: "tuff" };
/** What is on this row, in the words a person reads: "49 cut, 49 hotmelted,
    12 tuff". "" when there is nothing on it. ONE source for those words, so
    the question asked before a clear and the apology made when one could not
    be written cannot describe the same row differently. */
function clearWords(g) {
  if (!floorWorkToClear(g)) return "";
  const parts = [];
  ALL_STAGE_KEYS.forEach(k => {
    const n = Math.round(stNum(g[STAGE_ROW[k]], 0));
    if (n > 0) parts.push(n + " " + CLEAR_WORD[k]);
  });
  return parts.join(", ");
}
/** What the office is asked before a clear destroys work the floor has really
    recorded - naming what will go, because it can be a whole afternoon of it,
    from one click, on a job somebody is standing at. "" when there is nothing
    to lose, and then nothing is asked at all. */
function clearWarning(g) {
  const words = clearWords(g);
  if (!words) return "";
  return "The floor has recorded " + words + " on this job. Clearing the " +
         "glass here will set all of those back to zero. Clear it anyway?";
}

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
  return sectionInProduction((blockNames || [])[j.blk]);
}
/** The same question asked of a SECTION NAME rather than of a job - which is
    what a station list row carries, because the feeder copies the name onto it.
    One regex for both, so a tablet's filter and the feeder's slice can never
    disagree about what "In production" is. */
const sectionInProduction = name => /^\s*in production/i.test(stTxt(name));

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

/** How many tuff units a job has, off the sheet's own TUFF column. It is a
    count of the same glasses seen another way, so it is never added to the
    number above: a job can be 8 glasses and 11 tuff at the same time, and the
    tablet shows the two side by side rather than as one 19. */
function tuffTotal(j) {
  const g = (j && j.glass) || {};
  let n = 0;
  Object.keys(g).forEach(k => { if (stKey(k) === "TUFF") n += Math.round(stNum(g[k], 0)); });
  return Math.max(0, n);
}

/* ---- what the office has already ticked off --------------------------------
   The office ticks glass off job by job in the workbook's own colours long
   before a tablet appears on the floor, and a job that was finished in the
   office last week must not arrive on the tablet reading nothing done. That
   record is the glass checkpoints, and this turns it into the two counters
   the floor's row starts at.

   `counts` is plain data - one entry per glass item of the job, `{type, total,
   status, done}` where status is the sheet's own word ("done", "process",
   "cut" or nothing) and done is the office's count where it has one. It is
   built by whoever has checkpoints.js to hand (app.js), so nothing in this
   file has to know about a job object or a global.                          */
function officeSeed(row, counts) {
  const total = Math.max(0, Math.round(stNum(row && row.total, 0)));
  const none = { cut: 0, hotmelt: 0 };
  if (!(total > 0)) return none;
  const mine = (counts || []).filter(c => c && TOTAL_TYPES.indexOf(stKey(c.type)) >= 0);
  if (!mine.length) return none;
  const st = c => stTxt(c.status).trim().toLowerCase();
  const all = n => ({ cut: n, hotmelt: n });
  /* gold on every one of them: the office says this job's glass is done, so
     the floor's row is done too */
  if (mine.every(c => st(c) === "done")) return all(total);
  /* Yellow, on every one of them (owner, 2026-09-10, re-read 2026-09-21). An
     office yellow is read back as exactly what yellow now means, so the row
     round-trips to yellow again instead of being flattened to blank.

     It had to be fixed. Before this, a yellow item with no stored count seeded
     at nought (cpStored is empty for one), the floor's row read "nothing
     done", and the first tap anywhere on that job made the colour writer paint
     the office's own yellow out - a mark destroyed by a feature that had
     learned nothing from doing it. It was found on the one yellow glass cell
     in the owner's whole workbook.

     HOTMELT STAYS AT NOUGHT since 2026-09-21, and that is the point of the
     branch rather than an omission. Yellow now means ONE of the two stages is
     complete; cutting comes first, so cutting is the one. Seeding both would
     make the cell read gold and assert finished work that has not happened.

     ALL of them, not some. The floor's row carries one combined DG + TG number
     and there is nowhere to put a per-type split (spec §8), so this branch can
     only be taken when the office is saying the whole of the job's glass is
     yellow. A job with DG yellow and TG blank falls through to the count
     below: seeding cut to the total there would tell the FLOOR that glass
     which still needs cutting is cut, and a wrong instruction on the workshop
     screen is worse than a lost colour on a report. */
  if (mine.every(c => st(c) === "done" || st(c) === "process"))
    return { cut: total, hotmelt: 0 };
  /* part way there: the office's own count, a gold item counting its whole
     quantity. Both stages start there - the office does not record which
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
  if (mine.some(c => st(c) === "cut")) return { cut: total, hotmelt: 0 };
  return none;
}

/** Has the office finished this job's glass in its own record? That is what
    locks the job on the tablet (the OfficeDone column): the steppers grey out
    and tap() refuses them, and only the office can undo it by un-ticking.

    The question is asked of the same glass the floor is given - DG and TG - and
    not of ARCH, ASTRAGAL, FANCY or EXTRA, which the office ticks by hand and
    which describe work the floor never sees. A job with no DG and no TG is not
    "complete": it was never on the floor's board to lock. */
function officeComplete(counts) {
  const mine = (counts || []).filter(c => c && TOTAL_TYPES.indexOf(stKey(c.type)) >= 0 &&
                                          Math.round(stNum(c.total, 0)) > 0);
  if (!mine.length) return false;
  return mine.every(c => stTxt(c.status).trim().toLowerCase() === "done");
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
    const counts = look(j);
    const row = { title: job, job: job, customer: stTxt(j.cust).trim().slice(0, CUSTOMER_MAX),
                  total: total, tuffTotal: tuffTotal(j), seq: stNum(j.seq, 99999), active: active,
                  officeDone: officeComplete(counts) };
    row.seed = officeSeed(row, counts);
    out.push(row);
  });
  out.sort(stRowOrder);
  return out;
}

/** The job facts of one slice row, in list shape. */
function feederFields(row) {
  return { Job: row.job, Customer: row.customer, GlassType: GLASS_TYPE,
           Total: row.total, TuffTotal: Math.max(0, Math.round(stNum(row.tuffTotal, 0))),
           Seq: row.seq, Active: row.active ? "Yes" : "No",
           OfficeDone: row.officeDone ? "Yes" : "No" };
}
/** The two counters of a slice row's seed, in list shape. */
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
   can ever name are the two counters - on a row it is creating, or on a row
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
  /* which station's columns this plan is about. Omitted = the glass station,
     which is what this function was about before there was a second one, so
     every existing call site is unchanged. */
  const def = stDef(opts && opts.def);
  const order = def.rowOrder || stRowOrder;
  const rows = (slice || []).slice().sort(order);

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
    const want = def.feederOf(r);
    const mine = byTitle[t];
    if (!mine || !mine.length) {
      /* a row this feeder is creating starts where the office's own record
         already is - that is the whole of the seeding rule on a new row */
      adds.push(Object.assign({ Title: r.title }, want, def.seedOf(r), { FedAt: at, FedBy: by }));
      return;
    }
    const have = mine[0].fields || {};
    const diff = {};
    def.feederFields.forEach(k => { if (!sameField(have[k], want[k])) diff[k] = want[k]; });
    /* and a row the floor has never tapped is still the office's to say. Once
       DoneAt is set - the first tap writes it - this is skipped for ever, so
       nothing the office does can walk over the floor's own count. */
    if (untouched(have)) {
      const seed = def.seedOf(r);
      /* A SEED IS ONLY EVER RAISED (review finding R1, 2026-09-21), and the
         rehearsal on a saved copy of the real list is why. The office's yellow
         used to seed `Cut = Hotmelt = total`; since the colour rule changed it
         seeds `Cut = total, Hotmelt = 0` for the same record, and this branch
         wrote every difference in BOTH directions - so the first load after the
         ship would have PATCHed `Hotmelt` from the total back to nought on
         seven rows and handed the hotmelting tablet finished work as work to
         do. The office never "un-did" anything on those rows; only the reading
         of its mark changed, and a re-reading is not a reason to take a number
         away from somebody.

         The ONE case that may lower is a seed that is noughts all through: that
         is a real office un-tick (the record now says nothing of this job is
         done), and clearing the row is the whole point of re-seeding it. */
      const clearing = def.seedFields.every(k => !(seed[k] > 0));
      /* a blank counter IS nought here, unlike a blank job fact: a row nobody
         has ever written a number on does not need zeros put in it */
      def.seedFields.forEach(k => {
        const now = stNum(have[k], 0);
        if (now === seed[k]) return;
        if (seed[k] < now && !clearing) return;
        diff[k] = seed[k];
      });
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
function sliceHash(slice, def) {
  const d = stDef(def);
  const rows = (slice || []).slice().sort(d.rowOrder || stRowOrder);
  let s = "";
  rows.forEach(r => { s += JSON.stringify(d.hashOf(r)) + "\n"; });
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

   A job is finished, and goes gold on the office's board, when both glass
   counters have reached the total, there is a total to reach, and the job's
   tuff (if it has any) is counted too: a row with no glasses on it is not
   something anybody finished. Each TABLET has its own
   answer to "done here" since 2026-09-21 - the cutter's finished jobs leave
   the cutter's way before hotmelting has started - and that is the page's own
   reading of this record, not a second field on it.                         */
const isActive = f => stTxt(f && f.Active).trim().toLowerCase() === "yes";

/** Does this job still owe tuff? Owner, after the demo on 2026-09-21: "why are
    moving the jobs to complete when tuff is left? if job has tuff and is not
    done dont move it, if done then move."

    DISPLAY ONLY, and the boundary matters because 2026-09-10 put tuff outside
    the finished rule on purpose and this reverses exactly one clause of that.
    It decides whether a card is drawn gold and sinks to the bottom, on the
    office's board and on the cutting tablet, and NOTHING else: not the colour
    the Production sheet is painted (`glassColours` reads the counters direct),
    not the office's lock (`officeComplete` asks the office's own DG and TG
    checkpoints and has never known about tuff), not the feeder's `Active`,
    seed or hash, and not the header's "N left", which is this page's stage
    alone. A job with no tuff on it owes none and is unaffected. */
function tuffOwed(g) {
  return Math.max(0, Math.round(stNum(g && g.tuffTotal, 0))) > 0 && !stageComplete(g, TUFF_STAGE);
}

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
                total: total, tuffTotal: Math.max(0, Math.round(stNum(f.TuffTotal, 0))),
                /* the office's own "this job's glass is finished here". It greys
                   the tablet's steppers; only the office can take it off again */
                officeDone: stTxt(f.OfficeDone).trim().toLowerCase() === "yes",
                /* the first tap writes this and nothing else ever does, so it is
                   the one honest answer to "has the floor touched this row" */
                doneAt: stTxt(f.DoneAt),
                /* who made that last touch. Read only - the office writes this
                   field on a clear and the tablet on every tap, and since
                   2026-09-11 the colour writer needs a name to put in the
                   record's `Who` when it carries the floor's work up. */
                doneBy: stTxt(f.DoneBy),
                by: {}, at: {}, bars: {} };
    ALL_STAGE_KEYS.forEach(k => {
      const t2 = Math.max(0, Math.round(stNum(g[STAGE_TOTAL_ROW[k]], 0)));
      g[STAGE_ROW[k]] = stClamp(f[STAGE_FIELD[k]], t2);
      g.by[k] = stTxt(f[STAGE_BY[k]]);
      g.at[k] = stTxt(f[STAGE_AT[k]]);
      g.bars[k] = { done: g[STAGE_ROW[k]], total: t2, by: g.by[k], at: g.at[k] };
    });
    /* The two glass stages, AND the tuff the job is carrying (owner, after the
       demo, 2026-09-21: "why are moving the jobs to complete when tuff is
       left? if job has tuff and is not done dont move it, if done then move").
       This REVERSES the 2026-09-10 decision that tuff was outside the finished
       rule - and only that one. Tuff is still outside the job's glass total,
       still outside the office's lock, and still outside the header's "N left":
       what changed is only whether a card is SHOWN as done and sinks to the
       bottom. A job with no tuff on it finishes on the glass stages, exactly as
       before. */
    g.finished = total > 0 && STAGE_KEYS.every(k => g[STAGE_ROW[k]] >= total) && !tuffOwed(g);
    return g;
  });
  out.sort((a, b) => (a.seq - b.seq) || (a.job < b.job ? -1 : a.job > b.job ? 1 : 0));
  return out;
}

/** "12 glasses", "1 glass" - the one number the floor is given about a job,
    said the same way on the tablet and on the office's board. */
const glassWords = n => Math.max(0, Math.round(stNum(n, 0))) +
  (Math.round(stNum(n, 0)) === 1 ? " glass" : " glasses");

/** "20 left" - a number that has started counting down, which has to say so.
    Every number the tablet counts down now belongs to one stage of one job and
    is drawn on that stage's own row; the card's headline says how big the job
    is, in glasses, and does not move. "left" is the word that tells the two
    apart at arm's length. The office keeps glassWords everywhere: there, the
    number really is only ever the size of the job. */
const leftWords = n => Math.max(0, Math.round(stNum(n, 0))) + " left";

/** How much of ONE stage of ONE job is still to do: that stage's own quantity
    less its own counter.

    This is the only "what is left" arithmetic on the tablet, and it is
    deliberately small. The number is the JOB's, not a person's: two people who
    both cut see the same "20 left" on the same card, because 20 is genuinely
    what is left to cut on it - the owner's rule of 2026-09-10: if one person
    finishes 20 cuts out of 40, 20 are left for the others. Nothing here knows
    who is holding the tablet, so nothing here can make that untrue.

    Each stage counts against its OWN quantity (STAGE_TOTAL_ROW): cutting,
    hotmelting against the job's DG + TG, tuff against the sheet's
    own TUFF number. Tuff is never folded into the glass number - a job can be
    8 glasses and 11 tuff at the same time, and 19 is a count of nothing.

    The clamp is what keeps it in 0...total, so a row saying -9, or 99, or
    nothing at all, cannot put a negative or a NaN on a workshop wall. A word
    that is not one of the three stages is not a stage and gets 0 rather than a
    throw: the Stages column is typed by hand.

    Note what this is NOT. It does not decide gold: a job goes gold, and folds
    into Finished, when it is complete for everybody (jobBoard's `finished`).
    It does not decide a colour either: glassColours reads the raw counters
    through stageComplete and must keep doing so.                            */
function stageLeft(g, stage) {
  const k = stTxt(stage).trim().toLowerCase();
  if (ALL_STAGE_KEYS.indexOf(k) < 0) return 0;
  const t = Math.max(0, Math.round(stNum(g && g[STAGE_TOTAL_ROW[k]], 0)));
  return t - stClamp(g && g[STAGE_ROW[k]], t);
}

/** The stages a person actually holds, in the order the card draws them.
    "cut, cut" is one stage held, not two, and a word that is not a stage is
    not one they hold. stationPeople already drops both, but a number on a wall
    that doubles because a column was typed twice is not a thing to leave to
    another function's care. The order is the board's own, never the order the
    column happened to be typed in, so the header reads down in the same order
    as the steppers on every card. */
function heldStages(stages) {
  const seen = {};
  (stages || []).forEach(s => {
    const k = stTxt(s).trim().toLowerCase();
    if (ALL_STAGE_KEYS.indexOf(k) >= 0) seen[k] = true;
  });
  return ALL_STAGE_KEYS.filter(k => seen[k]);
}

/** One number per stage the person holds, on one job: `{stage, label, left}`
    in the order the steppers are drawn.

    Never one number across stages. That was tried, shipped and rejected by the
    owner on sight: a job of 49 glasses with 10 tuff read 157, which is
    49 x 3 + 10 and a count of nothing at all. Each entry here is a real number
    of real things, drawn on the row of the stage that moves it.

    A person holding nothing gets an empty list. They can tap nothing, and a
    number here would be a claim about work that is not theirs; the card still
    says how big the job is, which is what they are there to read. */
function jobLefts(g, stages) {
  return heldStages(stages).map(k => ({ stage: k, label: stageLabel(k), left: stageLeft(g, k) }));
}

/** The same numbers added down the whole board - one per stage held, never one
    across stages - so the header and the cards say the same thing and both
    count down on the same tap.

    It is given the whole board and never the searched one: somebody looking a
    job number up must not make the day's work read smaller than it is. */
function boardLefts(board, stages) {
  return heldStages(stages).map(k => ({
    stage: k, label: stageLabel(k),
    left: (board || []).reduce((n, g) => n + (g ? stageLeft(g, k) : 0), 0) }));
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
    "none". The only rule is the clamp: the floor may record hotmelting before
    it records cutting, and nothing here stops them - the counters are a record
    of what happened, not a workflow to be enforced. */
function applyTap(row, stage, delta) {
  const field = STAGE_ROW[stage];
  if (!field) return null;
  /* against that stage's own quantity: "All" on tuff means all eleven tuff, not
     all eight glasses */
  const total = Math.max(0, Math.round(stNum(row && row[STAGE_TOTAL_ROW[stage]], 0)));
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
function stationPeople(items, station, stageKeys) {
  const want = stTxt(station || STATION_NAME).trim().toLowerCase();
  /* which words in the Stages column are stages at all. Omitted = the glass
     station's three, which is what this function reads; the welding
     page passes its own single "weld". */
  const known = (stageKeys && stageKeys.length) ? stageKeys : ALL_STAGE_KEYS;
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
                     .filter(s => known.indexOf(s) >= 0)
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
function floorOnly(fields, def) {
  const d = stDef(def);
  const counters = d.counterFields || [];
  const out = {};
  d.floorFields.forEach(k => {
    if (!fields || !(k in fields)) return;
    const v = fields[k];
    if (counters.indexOf(k) >= 0) {
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

/* ---- a word from the floor to the office ------------------------------------
   Shipped 2026-09-15. Spec: docs/specs/2026-09-15-station-comments.md.

   The floor can move a counter and it can finish a job; until now it could not
   say anything. A shortage, a wrong measurement on the sheet, a pallet that
   arrived broken - none of that is a tick, and the only channel for it was
   somebody walking to the office. So: one note, typed on the card of the job it
   is about, read by the office in that job's drawer.

   THE WHOLE POINT OF THIS BLOCK IS THAT IT IS NOT THE GLASS STATION'S. More
   station tablets are coming (cutting, fabrication, glazing, windows, doors)
   and every one of them will want the same channel. So the composer, the
   thread, the row it writes and the two list calls all live here, behind
   stationComments(), whose ONLY station-specific input is cfg.station: a
   future cutting.html changes that one word and has the feature.

   Three rules, from the spec and from CLAUDE.md:

     · APPEND-ONLY. One row per note, never upserted by Title, never patched,
       never deleted - by either side. A mistake gets a follow-up note. There is
       no edit path and no delete path here to find;
     · the tablet writes this list DIRECTLY, exactly as it writes Station log,
       and touches no workbook of any kind;
     · the tablet sees only its OWN station's notes on a job. The office drawer
       is where every station's notes about a job come together.              */
const COMMENT_LIST = "Station comments";
/* Title is a row id (JOB|unix-ms), not a key: this list has no unique rule on
   it and nothing ever reads a row back by it. Text is the note itself, the one
   free-text field on any floor list. */
const COMMENT_FIELDS = ["Title", "Job", "Station", "Who", "Text", "At"];
/* A note is a sentence or two about a job, not a document. The cap is on the
   composer and again on the row builder, so a long paste cannot arrive through
   a replayed draft either. */
const COMMENT_MAX = 2000;
/* How often a tablet with a composer open asks the list what else has been
   said. Only with one open: a closed composer has nothing to show and a read
   six times a minute of a list nobody is looking at is a poll for its own sake. */
const COMMENT_POLL_MS = 20000;

/* No list is created by code (CLAUDE.md rule 3). A missing list says which one,
   plainly, on whichever screen is asking, and writes nothing anywhere. */
const COMMENT_MISSING_FLOOR = "The “Station comments” list is not in the floor’s site yet, " +
  "so notes cannot be left here. Ask the office to add it — nothing in the Excel file is involved.";
const COMMENT_MISSING_OFFICE = "The “Station comments” list is not in the floor’s site yet, " +
  "so the floor cannot leave notes. Ask the manager to add it — nothing in the Excel file is involved.";
const COMMENT_UNREACHABLE = "The floor’s notes could not be read just now — retrying.";
const COMMENT_CHECKING = "checking…";
const COMMENT_EMPTY_FLOOR = "No notes on this job yet.";
const COMMENT_EMPTY_OFFICE = "No notes from the floor on this job yet.";
/* A note that would not send is kept in the box and said so: the floor has
   typed something they meant somebody to read, and dropping it silently is the
   one thing this feature must not do. */
const COMMENT_UNSENT = "not sent — tap Send again";
/* ... and the other way round: the note DID go, but there is still writing in
   the box because somebody carried on typing while it was in the air. Without
   a word here the box looks exactly like a note that failed to send, and the
   obvious thing to do about it - tap Send again - would post the first note
   twice. */
const COMMENT_KEPT = "sent — what is in the box is a new note";

/** The row id: the job and the instant, which is unique enough for a list
    nothing ever looks a row up in. */
function commentTitle(job, at) {
  const t = stNum(at, NaN);
  const ms = isFinite(t) ? t : Date.parse(stTxt(at));
  return stKey(job) + "|" + (isFinite(ms) ? Math.round(ms) : Date.now());
}
/** The one row the composer sends. Six columns, all text, none of them a
    counter, a colour, a cell or anything the workbook has ever heard of. */
function commentFields(e) {
  const at = stTxt(e && e.at) || new Date().toISOString();
  return { Title: commentTitle(e && e.job, at),
           Job: stKey(e && e.job),
           Station: stTxt((e && e.station) || STATION_NAME).trim(),
           Who: stTxt(e && e.who),
           Text: stTxt(e && e.text).trim().slice(0, COMMENT_MAX),
           At: at };
}

/** Two `At` stamps, compared by the MOMENT they name rather than as text.
    -1 / 0 / 1, like any comparator.

    ISO stamps mostly sort as text, and that is what this used to lean on - but
    not always: "…10:00:00.100Z" is a tenth of a second AFTER "…10:00:00Z" and
    sorts BEFORE it as a string, because "." is below "Z". The tablet writes
    milliseconds (`new Date().toISOString()`); a row typed into SharePoint by
    hand often has none, and the office's "have I seen this note" stamp is
    whichever of the two it last saw - so one undated-looking second could hide
    every note that followed it. A stamp neither side can parse as a date (a
    hand-typed "yesterday") falls back to the text comparison it always had,
    which is the best that can be said about it. */
function atCmp(a, b) {
  const ta = Date.parse(stTxt(a)), tb = Date.parse(stTxt(b));
  if (isFinite(ta) && isFinite(tb)) return ta < tb ? -1 : ta > tb ? 1 : 0;
  const sa = stTxt(a), sb = stTxt(b);
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/** The list, read into rows - OLDEST FIRST on both screens, because this reads
    as a conversation and not as a log.

    `f.job` narrows to one job; `f.station` narrows to one station, which is
    what the tablet passes and what the office drawer deliberately does not. A
    row whose Station column is empty is shown to whoever asks, the same rule
    logRows has always used: it can only have come off a tablet, and hiding it
    from every screen would lose a note somebody typed. */
function commentRows(items, f) {
  const job = stKey(f && f.job);
  const station = stTxt(f && f.station).trim().toLowerCase();
  const out = [];
  (items || []).forEach(it => {
    if (!it) return;
    const fl = it.fields || {};
    /* the Job column, with the Title's own prefix as the fallback - a row typed
       into SharePoint by hand can easily have one and not the other */
    const j = stKey(fl.Job) || stKey(stTxt(fl.Title).split("|")[0]);
    const text = stTxt(fl.Text).trim();
    if (!j || !text) return;                       // an empty note is not a note
    const st = stTxt(fl.Station).trim();
    if (job && j !== job) return;
    if (station && st && st.toLowerCase() !== station) return;
    out.push({ id: stTxt(it.id), job: j, station: st, who: stTxt(fl.Who).trim(),
               text: text.slice(0, COMMENT_MAX), at: stTxt(fl.At) });
  });
  /* by the moment each stamp names (atCmp, not text - a note with milliseconds
     on it sorts after one without); the item id breaks a tie, because two notes
     inside one second are possible and the older id is the older note */
  out.sort((a, b) => atCmp(a.at, b.at) || stItemAge(a, b));
  return out;
}

/** "just now" / "12 min ago" / "3 h ago" / "2 days ago", from an ISO stamp. The
    tablet's own agoWords takes a millisecond number; this one is pure, takes
    what the list carries, and can be tested without a clock. */
function commentAgo(at, now) {
  const t = Date.parse(stTxt(at));
  if (!isFinite(t)) return "";
  const s = Math.max(0, Math.round((stNum(now, Date.now()) - t) / 1000));
  if (s < 45) return "just now";
  if (s < 5400) return Math.round(s / 60) + " min ago";
  if (s < 172800) return Math.round(s / 3600) + " h ago";
  return Math.round(s / 86400) + " days ago";
}
/** The Changes-panel line for a note the office has not seen before. The
    sentence is the owner's, from the spec, and is built here so both the page
    and its test read the same words. */
function commentChangeWords(r) {
  return "New floor note on " + stKey(r && r.job) +
         ", from " + (stTxt(r && r.station).trim() || "a floor station") +
         ", " + (stTxt(r && r.who).trim() || "somebody");
}

/* This file draws no DOM - it never has - but it does build the strings both
   pages set as innerHTML, exactly as logFields builds the row both pages read.
   So it needs its own escape rather than borrowing a page's. */
const stEsc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

/* ---- the channel itself -----------------------------------------------------
   One object per station page, made once at start-up. It holds the list, what
   is typed but not sent, which composers are open, and the two calls that read
   and append - and it takes the station's name, the list transport and a clock,
   so nothing in here knows what a Graph is and the tests need no network.

   Everything a page has to do with it:

       const NOTES = ST.stationComments({ station: ST.STATION_NAME,
                                          listItems: CW.listItems,
                                          listAdd: CW.listAdd,
                                          opts: listOpts });
       ... NOTES.read()  once at start
       ... NOTES.poll()  on the page's own clock
       ... NOTES.html(job)  inside the card
       ... NOTES.toggle / setDraft / send  from the card's clicks

   `station` is the ONLY station-specific input. That is the feature.        */
function stationComments(cfg) {
  cfg = cfg || {};
  const station = stTxt(cfg.station || STATION_NAME).trim() || STATION_NAME;
  const list = stTxt(cfg.list || COMMENT_LIST);
  const opts = () => (typeof cfg.opts === "function" ? cfg.opts() : (cfg.opts || {}));
  const clock = () => stNum(typeof cfg.now === "function" ? cfg.now() : cfg.now, Date.now());

  /* ok: null not looked yet · false missing, refused or unreachable · true read
     it. `why` is which of those, in words, for the card. */
  const S = { items: null, ok: null, why: "", readAt: 0,
              open: {}, draft: {}, sending: {}, failed: {}, kept: {} };

  const key = job => stKey(job);
  /** This station's notes on one job, oldest first. */
  const rows = job => commentRows(S.items || [], { job: job, station: station });
  const isOpen = job => !!S.open[key(job)];
  const draftOf = job => stTxt(S.draft[key(job)]);

  function toggle(job) {
    const k = key(job);
    if (S.open[k]) delete S.open[k]; else S.open[k] = 1;
    return !!S.open[k];
  }
  function setDraft(job, text) {
    const k = key(job);
    S.draft[k] = stTxt(text).slice(0, COMMENT_MAX);
    /* the "that was sent, this is new" note has been read by then: the next
       keystroke is somebody carrying on, and a hint that outstays the moment it
       explains is just another thing on the card to read */
    if (S.kept[k]) delete S.kept[k];
  }
  const anyOpen = () => Object.keys(S.open).length > 0;

  /** The whole list, quietly. A missing list is a state, not an error; anything
      else leaves whatever was last read exactly where it is. Never throws:
      a note channel must not be able to take a working board away. */
  async function read() {
    if (typeof cfg.listItems !== "function") return false;
    try {
      const items = await cfg.listItems(list, opts());
      if (items == null) {
        S.ok = false; S.why = stTxt(cfg.missing) || COMMENT_MISSING_FLOOR;
        S.readAt = clock();
        return false;
      }
      S.items = items; S.ok = true; S.why = ""; S.readAt = clock();
      return true;
    } catch (e) {
      S.readAt = clock();
      if (S.ok !== true) { S.ok = false; S.why = COMMENT_UNREACHABLE; }
      return false;
    }
  }
  /** The page's clock, answered only when somebody has a composer open and the
      last read is old enough to be worth replacing. */
  async function poll() {
    if (!anyOpen()) return false;
    if (S.readAt && clock() - S.readAt < COMMENT_POLL_MS) return false;
    return await read();
  }

  /** Append one note. One POST, no read, no upsert, no id looked up: this list
      is a log and every row is its own record.

      The row goes into the copy in hand at once so the thread shows it without
      waiting for a read, exactly as a tap shows before its write lands. A
      refused write puts the text back in the box and says so. */
  async function send(job, who) {
    const k = key(job);
    /* exactly what is in the box, kept beside the trimmed copy that is sent.
       The box stays editable while the POST is in the air - only the Send
       button is disabled - and on workshop wifi that is a second or two in
       which somebody can start the next note. */
    const raw = draftOf(job);
    const text = raw.trim();
    if (!text) return null;
    if (S.ok === false) return null;               // nowhere to write it
    if (S.sending[k]) return null;                 // one at a time per job
    if (typeof cfg.listAdd !== "function") return null;
    const fields = commentFields({ job: job, station: station, who: who,
                                   text: text, at: new Date(clock()).toISOString() });
    S.sending[k] = 1; delete S.failed[k]; delete S.kept[k];
    try {
      const made = await cfg.listAdd(list, fields, opts());
      const id = made && made.id != null ? String(made.id) : "local:" + fields.Title;
      S.items = (S.items || []).concat([{ id: id, fields: fields }]);
      /* EMPTY THE BOX ONLY IF IT STILL HOLDS WHAT WENT (review finding 3).
         Clearing it unconditionally throws away whatever has been typed since
         the tap, silently, with no message and no way back - the one thing this
         feature must never do with something somebody meant to be read. If it
         has moved on, the note that landed is in the thread above and the new
         typing is left exactly where it is - and the box SAYS SO, or writing
         left behind on purpose looks exactly like a note that failed to send. */
      if (draftOf(job) === raw) S.draft[k] = "";
      else S.kept[k] = COMMENT_KEPT;
      return fields;
    } catch (e) {
      S.failed[k] = COMMENT_UNSENT;                // the typing stays in the box
      return null;
    } finally {
      delete S.sending[k];
    }
  }

  /** One thread line: who, how long ago, and what they said. */
  function lineHtml(r, now) {
    return '<div class="cmrow">' +
      '<span class="cmwho">' + stEsc(r.who || "—") + '</span>' +
      '<span class="cmwhen">' + stEsc(commentAgo(r.at, now)) + '</span>' +
      '<div class="cmtext">' + stEsc(r.text) + '</div></div>';
  }

  /** The composer and the thread, for one card. Closed it is one button with a
      count on it; open it is this station's notes on the job, oldest first,
      and a box to add to them. */
  function html(job) {
    const k = key(job);
    const now = clock();
    const mine = rows(job);
    const open = isOpen(job);
    const n = mine.length;
    const head = '<button class="cmtog" data-cmt="' + stEsc(job) + '" aria-expanded="' + (open ? "true" : "false") + '">' +
      '<svg class="cmicon" width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">' +
      '<path fill="currentColor" d="M2 2h12v9H6.6L3 13.8V11H2z"/></svg>' +
      (n ? "Notes (" + n + ")" : "Note") + '</button>';
    if (!open) return '<div class="cmt">' + head + '</div>';
    const body =
      S.ok === null ? '<div class="cmhint">' + stEsc(COMMENT_CHECKING) + '</div>'
      : S.ok === false ? '<div class="cmhint">' + stEsc(S.why || COMMENT_MISSING_FLOOR) + '</div>'
      : (n ? '<div class="cmthread">' + mine.map(r => lineHtml(r, now)).join("") + '</div>'
           : '<div class="cmhint">' + stEsc(COMMENT_EMPTY_FLOOR) + '</div>') +
        /* the "\n" straight after the tag is not decoration: HTML throws away a
           newline immediately after <textarea>, so without one of its own a
           draft that starts with a blank line loses it on the next redraw */
        '<textarea class="cmbox" data-cmbox="' + stEsc(job) + '" rows="2" maxlength="' + COMMENT_MAX + '" ' +
          'placeholder="A word for the office about this job…">\n' + stEsc(draftOf(job)) + '</textarea>' +
        '<div class="cmfoot">' +
          '<span class="cmnote' + (S.failed[k] ? " bad" : "") + '">' +
            stEsc(S.failed[k] || S.kept[k] ||
                  "Sent to the office. Nothing in the Excel file is involved.") + '</span>' +
          '<button class="cmsend" data-cmsend="' + stEsc(job) + '"' + (S.sending[k] ? " disabled" : "") + '>' +
          (S.sending[k] ? "Sending…" : "Send") + '</button></div>';
    return '<div class="cmt open">' + head + '<div class="cmbody">' + body + '</div></div>';
  }

  /** What this card is showing of the channel, for the board's repaint diff.
      The DRAFT is deliberately not in it: a card must not be rebuilt under the
      fingers typing into it. */
  function sig(job) {
    const k = key(job);
    const mine = rows(job);
    return (isOpen(job) ? "o" : "") + mine.length + ":" +
           (mine.length ? mine[mine.length - 1].id : "") +
           (S.sending[k] ? "s" : "") + (S.failed[k] ? "!" : "") + (S.kept[k] ? "=" : "") +
           "|" + (S.ok === null ? "?" : S.ok ? "y" : "n");
  }

  return { station: station, list: list, state: S,
           rows: rows, isOpen: isOpen, draftOf: draftOf, anyOpen: anyOpen,
           toggle: toggle, setDraft: setDraft,
           read: read, poll: poll, send: send, html: html, sig: sig };
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
  return JSON.stringify([g.id, g.job, g.customer, g.seq, g.total, g.tuffTotal,
    g.finished, g.officeDone,
    ALL_STAGE_KEYS.map(k => [g[STAGE_ROW[k]], g.by[k], g.at[k]])]);
}
function boardDiff(prev, next, sigOf) {
  /* the signature of a card is the one station-specific thing here. Omitted =
     the glass card's, which is what it has always been; the welding tablet
     passes WELDC.weldCardSig. Everything else - what added, changed, removed
     and "the order moved" mean - is the same question on any board. */
  const sig = typeof sigOf === "function" ? sigOf : cardSig;
  const was = {}, now = {};
  (prev || []).forEach(g => { was[g.job] = sig(g); });
  (next || []).forEach(g => { now[g.job] = sig(g); });
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

/* ---- the colour the Production sheet should be showing ----------------------
   New on 2026-09-10, and the first time anything the floor does reaches the
   master sheet. The tablet still cannot: it has no access to the workbook at
   all. This is the office's own reading of the floor's row, and the office
   dashboard is what paints the cells - the same shape as the feeder, in
   reverse.

   Four columns, and only these four. ARCH, ASTRAGAL, FANCY and EXTRA are
   ticked by hand in the office and are never written by this feature, in
   either direction, under any state of the floor's row (owner, 2026-09-10).

   THE RULE CHANGED ON 2026-09-21 (docs/specs/2026-09-21-glass-split-no-
   glazing.md), because glazing left this station and could no longer be the
   gate for gold. It is now a count of how many of the two glass stages are
   complete:

     DG, TG, NOT TUFF   blank   neither cutting nor hotmelting complete
                        yellow  exactly one of them complete
                        gold    both complete
     TUFF               gold once the tuff count is complete, blank once it has
                        been tapped and is not complete, and NO OPINION AT ALL
                        while nobody has ever tapped it

   It walks back down as well as up - a job tapped back to nought goes yellow
   and then blank again - because the owner asked for it to be reversible: "do
   not make it one way because it would be hard to reverse it when it done by
   mistake".

   "Complete" means the counter has reached that stage's own quantity and there
   is a quantity to reach: a row with no glasses on it has finished nothing.

   TUFF'S THIRD ANSWER IS `null`, AND IT IS NOT THE SAME AS BLANK (review
   finding R5, 2026-09-21). `Tuff` is the one counter the feeder may never seed,
   so on every row the cutter has not tapped it reads nought - which said
   nothing at all while glazing painted TUFF gold, and would now read as "the
   floor says no tuff is done" and wipe a gold TUFF cell the office ticked by
   hand. A `null` is carried all the way to the writer, which plans neither a
   paint nor a clear for that column. The moment somebody taps tuff, `TuffAt`
   is written and the column has an opinion for ever after - including "blank",
   when it is tapped back down to nought on purpose.

   Nothing here reads `Glazed` any more. A row still carrying a full `Glazed`
   from before this build plans exactly what its Cut and Hotmelt say, which is
   the whole of rule 8: the old data is left alone and has no influence.     */
const COLOUR_TYPES = ["dg", "tg", "tuff", "not tuff"];
/** Is one stage of a board record at its own total? */
function stageComplete(g, k) {
  const t = Math.max(0, Math.round(stNum(g && g[STAGE_TOTAL_ROW[k]], 0)));
  return t > 0 && stClamp(g && g[STAGE_ROW[k]], t) >= t;
}
/** Has the floor ever tapped the TUFF counter on this row? It is the one
    counter the feeder may never seed, so the stamp is the only honest answer -
    a nought with no stamp beside it is nobody's statement about anything. */
const tuffSpoken = g => stTxt(g && g.at && g.at[TUFF_STAGE]).trim() !== "";
/** What each of the four columns should be showing: "gold", "yellow", ""
    (no colour at all) or, for TUFF alone, `null` - no opinion, so the writer
    neither paints nor clears it. Pure: it reads a board record and nothing
    else. */
function glassColours(g) {
  /* how many of the two, not which: cutting before hotmelting and hotmelting
     before cutting are the same half-done job to the sheet */
  const done = STAGE_KEYS.reduce((n, k) => n + (stageComplete(g, k) ? 1 : 0), 0);
  const glass = done >= STAGE_KEYS.length ? "gold" : done > 0 ? "yellow" : "";
  const tuff = stageComplete(g, TUFF_STAGE) ? "gold" : (tuffSpoken(g) ? "" : null);
  const out = {};
  COLOUR_TYPES.forEach(t => { out[t] = t === "tuff" ? tuff : glass; });
  return out;
}
/** When the floor last moved anything on this job, as the ISO stamp it wrote -
    "" when they never have. Every tap sets DoneAt, so that is normally the
    answer; the per-stage stamps are taken as well because a row written by
    some future hand that forgot one must not read as "never touched".

    The newest is taken by PARSED TIME, and a stamp that will not parse is
    passed over rather than compared. Taking a text maximum, which ISO stamps
    ordinarily allow, was wrong for a reason worth remembering: every column
    read here is editable by hand in SharePoint, and one junk character in one
    of them would win the text comparison ("zzz" beats every real stamp),
    answer a time nothing can read, and switch the whole feature off for that
    job in silence. Failing closed was the right direction; failing closed
    because somebody's keyboard slipped was not.

    This is one half of the last-writer-wins comparison in §3 of the spec. The
    other half - when the office last said something about the same job's glass
    - is the dashboard's own record (Dashboard Progress and Dashboard Log) and
    is worked out in app.js, which is the only side that can see it. */
function floorStamp(g) {
  let best = "", bestAt = 0;
  const look = v => {
    const s = stTxt(v).trim();
    if (!s) return;
    /* the floor writes new Date().toISOString() and nothing else, so Date.parse
       is the whole of what this needs; anything it cannot read is not a time */
    const t = Date.parse(s);
    if (!isFinite(t) || t <= bestAt) return;
    best = s; bestAt = t;
  };
  look(g && g.doneAt);
  ALL_STAGE_KEYS.forEach(k => look(g && g.at && g.at[k]));
  return best;
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

/* ---- rule 3's strip, in one place ------------------------------------------
   MOVED HERE FROM welding-core.js on 2026-09-21 (docs/specs/2026-09-21-day-
   sheets-and-station-reports.md). It was the welding COMMENT column's guard and
   nothing outside that file could reach it; a station report exports free text
   somebody typed on a tablet - a day sheet's note, a floor note - and rule 3
   says no phone number and no eircode leaves the app in an export, ever. So the
   shapes are written down once, here, and weldStripDigits is now a call to this
   with the comment's own cap. Welding's behaviour is unchanged by construction.

   FOUR SHAPES, and the first of them is the one that matters:

     · a PHONE NUMBER AS PEOPLE TYPE ONE - six or more digits with spaces,
       brackets, dots or dashes between them, and an optional leading "+":
       `086 123 4567`, `+353 86 123 4567`, `087-123-4567`, `(086) 123 4567`.
       A phone number in a workshop note is almost never ten digits in a row;
     · digit groups joined by `/`, `,`, `:` or `_`, when the whole match holds
       nine or more digits - so a date (`12/09/2026`) and a small size
       (`spacer 4/20/4`) are left alone;
     · any run of six or more digits;
     · anything eircode-shaped (a letter, two digits or a digit and W, then four
       alphanumerics).

   WHAT SURVIVES, and is tested so it stays surviving: a job number (`R5303`),
   the sheet's own short dates (`12.09`, `12/09/2026`) and small quantities.
   WHAT IS OVER-STRIPPED, said rather than discovered: a full ISO date typed
   into free text, and a dash-joined run of sizes (`cill 150-2100`). That is the
   right way round to be wrong - the owner's rule is that no phone number leaves
   the app, and a lost size in a note costs nobody anything.                  */
const STRIP_PHONE_RE = /(?:\+?\d[\s().‐-―-]{0,2}){6,}\d?/g;
const STRIP_SEP_RE = /\d+(?:[\/,:_]\d+)+/g;
const STRIP_DIGITS_RE = /\d{6,}/g;
const STRIP_EIR_RE = /\b[A-Za-z]\d(?:\d|[Ww])\s?[A-Za-z0-9]{4}\b/g;
const STRIP_MAX = 140;
/** Free text with anything phone- or eircode-shaped replaced by an ellipsis,
    whitespace tidied and the result capped. Blank in, blank out. Pure. */
function stripContact(text, max) {
  const cap = Math.max(1, Math.round(stNum(max, STRIP_MAX)));
  const t = stTxt(text)
    .replace(STRIP_PHONE_RE, "…")
    .replace(STRIP_SEP_RE, m => ((m.match(/\d/g) || []).length >= 9 ? "…" : m))
    .replace(STRIP_DIGITS_RE, "…")
    .replace(STRIP_EIR_RE, "…");
  return t.replace(/\s+/g, " ").trim().slice(0, cap);
}

/* ---- the end-of-day sheet ---------------------------------------------------
   Shipped 2026-09-21 (docs/specs/2026-09-21-day-sheets-and-station-reports.md).
   The cutter fills in a paper sheet every day: their name, the day, four counts
   and a line about anything that got in the way. This is that sheet, on the
   tablet, against a weekly target the office sets.

   IT IS NOT THE CUTTING STATION'S, in exactly the way the note channel is not
   the glass station's: which counts a sheet has - and whether a stage has one at
   all - comes from the station DEFINITION (`def.daySheets[stage]`), so
   hotmelting or welding get one by gaining a line of definition and nothing
   else. A stage with no entry draws no button and reads neither list.

   Two lists, in whichever site the station's lists live:
     `Station day sheets`  one row per person per station-stage per day,
                           APPEND-ONLY from the tablet (one POST, no PATCH, no
                           DELETE). The office may correct the counts and the
                           note, and never deletes;
     `Station targets`     one row per station-stage, written by the office
                           only, read-only on every tablet.
   No list is created by code: a missing one is a quiet explained state.     */
const DAY_LIST = "Station day sheets";
const TARGET_LIST = "Station targets";
/* every column but the counts, which are the definition's own */
const DAY_BASE_FIELDS = ["Title", "Station", "Stage", "Day", "Who",
                         "Note", "WeekTarget", "SavedAt", "EditedBy", "EditedAt"];
const TARGET_FIELDS = ["Title", "WeeklyTarget", "SetBy", "SetAt"];
/* a line about the day, not a document */
const DAY_NOTE_MAX = 500;
const DAY_MISSING_FLOOR = "The “Station day sheets” list is not in the floor’s site yet, so the " +
  "day sheet cannot be saved here. Ask the office to add it — nothing in the Excel file is involved.";
const DAY_MISSING_OFFICE = "The “Station day sheets” list is not in the floor’s site yet, so there " +
  "are no day sheets to show. Ask the manager to add it — nothing in the Excel file is involved.";
const TARGET_MISSING_OFFICE = "The “Station targets” list is not in the floor’s site yet, so a " +
  "weekly target cannot be set. Ask the manager to add it — nothing in the Excel file is involved.";
const DAY_UNREACHABLE = "The day sheets could not be read just now — retrying.";
const DAY_SAVED_WORDS = "ask the office to correct a mistake";

/** The heading a table puts over one count column, falling back to the
    question the tablet asks when a definition gives no short one. */
const dayCountShort = c => stTxt((c || [])[2] || (c || [])[1] || (c || [])[0]);
/** A stamp as the LOCAL clock reads it, "14:02" - and "" for anything that is
    not a time. One implementation for both screens: the office's stWhen builds
    its own line on top of this, the tablet says "saved 14:02" with it, and
    neither of them slices the ISO string, which is the UTC clock and is an
    hour out all summer. */
function stClock(iso) {
  const d = new Date(stTxt(iso));
  if (isNaN(d.getTime())) return "";
  const p = n => (n < 10 ? "0" : "") + n;
  return p(d.getHours()) + ":" + p(d.getMinutes());
}

/** This stage's day sheet, or null - the whole of "which pages have one". */
function daySheetOf(def, stage) {
  const ds = (stDef(def).daySheets || {})[stTxt(stage).trim().toLowerCase()];
  return (ds && ds.counts && ds.counts.length) ? ds : null;
}
/** The stages of a station that have one, in the definition's own order. */
function daySheetStages(def) {
  const d = stDef(def);
  return (d.stages || []).filter(k => !!daySheetOf(d, k));
}
/** The columns of a read: the fixed ones plus this stage's own counts. */
function dayFieldsFor(def, stage) {
  const ds = daySheetOf(def, stage);
  return DAY_BASE_FIELDS.concat(ds ? ds.counts.map(c => c[0]) : []);
}
/** A LOCAL calendar day as YYYY-MM-DD. Local, deliberately: a sheet saved at
    23:30 belongs to the day the person worked, not to UTC's idea of it. A
    string that is already a day is taken as one rather than re-parsed, which is
    the same trap the other way round. */
function dayKey(d) {
  if (typeof d === "string") {
    const m = /^(\d{4}-\d{2}-\d{2})/.exec(d.trim());
    if (m) return m[1];
  }
  const x = d == null ? new Date() : (d instanceof Date ? d : new Date(d));
  if (!x || isNaN(x.getTime())) return "";
  const p = n => (n < 10 ? "0" : "") + n;
  return x.getFullYear() + "-" + p(x.getMonth() + 1) + "-" + p(x.getDate());
}
const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
/** Which ISO week a day is in: Monday to Sunday, `2026-W39`, with the Monday
    and the Sunday beside it. The arithmetic is done at noon UTC so no timezone
    can move a day across a week boundary, and the year comes from the week's
    THURSDAY - which is what makes 2026-12-31 and 2027-01-01 the same week. */
function isoWeek(day) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dayKey(day));
  if (!m) return { key: "", monday: "", sunday: "", weekday: "", dow: -1 };
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], 12));
  const dow = (d.getUTCDay() + 6) % 7;                    // Monday = 0
  const mon = new Date(d.getTime() - dow * 86400000);
  const thu = new Date(mon.getTime() + 3 * 86400000);
  /* NOON at both ends: with midnight on one side and noon on the other, the
     half day rounds a January Thursday into week 2 */
  const jan1 = Date.UTC(thu.getUTCFullYear(), 0, 1, 12);
  const wk = Math.floor(Math.round((thu.getTime() - jan1) / 86400000) / 7) + 1;
  const iso = x => x.toISOString().slice(0, 10);
  return { key: thu.getUTCFullYear() + "-W" + (wk < 10 ? "0" + wk : String(wk)),
           monday: iso(mon), sunday: iso(new Date(mon.getTime() + 6 * 86400000)),
           weekday: DAY_NAMES[dow], dow: dow };
}
/** The Title, which is the key: a second save for the same person, stage and
    day is refused by SharePoint as well as by the page. */
function dayTitle(station, stage, day, who) {
  return [stTxt(station).trim(), stTxt(stage).trim().toLowerCase(),
          dayKey(day), stTxt(who).trim()].join("|");
}
/** One count as it may be saved: a whole number, nought or more. A blank is
    nought (the paper sheet leaves them blank); anything else - a decimal, a
    minus, a word - is null, and the form refuses the save rather than rounding
    somebody's afternoon into a number they did not write.

    CAPPED AT DAY_COUNT_MAX (review, 2026-09-21). Nobody cuts ten thousand
    sheets in a day, and a finger held on a number pad is the likeliest way one
    would arrive - on a row the cutter cannot correct afterwards, into a week
    total, and into the office's report. Above the cap it is refused at the form
    like any other thing that is not a count. */
const DAY_COUNT_MAX = 9999;
function dayCount(v) {
  const s = stTxt(v).trim();
  if (s === "") return 0;
  if (!/^\d+$/.test(s)) return null;
  const n = Number(s);
  if (!isFinite(n) || n > DAY_COUNT_MAX) return null;
  return n;
}
/** The one row a save POSTs, or null when a count is not a whole number. Every
    column it can carry is here: there is no path that could add another. */
function dayFields(def, stage, e) {
  const ds = daySheetOf(def, stage);
  if (!ds) return null;
  const d = stDef(def);
  e = e || {};
  const day = dayKey(e.day || new Date());
  const who = stTxt(e.who).trim();
  if (!day || !who) return null;
  const out = { Title: dayTitle(d.name, stage, day, who), Station: d.name,
                Stage: stTxt(stage).trim().toLowerCase(), Day: day, Who: who };
  const counts = e.counts || {};
  let bad = false;
  ds.counts.forEach(c => {
    const n = dayCount(counts[c[0]]);
    if (n == null) bad = true; else out[c[0]] = n;
  });
  if (bad) return null;
  out.Note = stTxt(e.note).trim().slice(0, DAY_NOTE_MAX);
  /* the target in force when this was saved, so an old week stays true after
     the office changes it. No target set = the column is not written at all */
  const t = stNum(e.weekTarget, null);
  if (t != null && isFinite(t)) out.WeekTarget = Math.max(0, Math.round(t));
  out.SavedAt = stTxt(e.at) || new Date().toISOString();
  return out;
}
/** What the OFFICE may correct, and the proof it can correct nothing else: a
    builder that takes counts, a note, a name and a time. null when a count is
    not a whole number. */
function dayOfficeFields(counts, e) {
  e = e || {};
  const out = {};
  let bad = false;
  (counts || []).forEach(c => {
    const n = dayCount((e.counts || {})[c[0]]);
    if (n == null) bad = true; else out[c[0]] = n;
  });
  if (bad) return null;
  out.Note = stTxt(e.note).trim().slice(0, DAY_NOTE_MAX);
  out.EditedBy = stTxt(e.who);
  out.EditedAt = stTxt(e.at) || new Date().toISOString();
  return out;
}
/** The list read into rows and filtered. `counts` is the definition's own list
    of [key, label]; `f` may narrow by station, stage, who, from/to day, weekday
    (0 = Monday) and ISO week. Newest day first.

    ONE ROW PER TITLE, OLDEST ITEM ID WINNING (review, 2026-09-21). Title is the
    key and the owner turns enforce-unique-values on, but the code must not
    depend on that having been done: without this, a list with the rule left off
    - or two rows typed in by hand - doubles that person's day in the office
    window, in the week subtotal, in the board's line, in the tablet's own "this
    week" and in the report, all silently. It is the same rule every other list
    here already uses (buildJobs, weldRecords, targetOf), and it runs before the
    filters so a duplicate carrying a hand-edited Day cannot hide the real row. */
function dayRows(items, counts, f) {
  const cols = counts || [];
  const want = f || {};
  const station = stTxt(want.station).trim().toLowerCase();
  const stage = stTxt(want.stage).trim().toLowerCase();
  const who = stTxt(want.who).trim().toLowerCase();
  const from = dayKey(want.from || ""), to = dayKey(want.to || "");
  const week = stTxt(want.week).trim();
  const dow = want.weekday == null || want.weekday === "" ? -1 : Number(want.weekday);
  const best = {}, order = [];
  (items || []).forEach(it => {
    if (!it) return;
    /* a row with no Title cannot have come off a tablet and is nobody's
       duplicate: it keeps its own identity rather than colliding on "" */
    const k = stTxt((it.fields || {}).Title).trim().toUpperCase() || ("#" + stTxt(it.id));
    if (!best[k]) { best[k] = it; order.push(k); return; }
    if (stItemAge(it, best[k]) < 0) best[k] = it;
  });
  const out = [];
  order.forEach(k => {
    const it = best[k];
    const fl = it.fields || {};
    const day = dayKey(stTxt(fl.Day));
    if (!day) return;
    const st = stTxt(fl.Station).trim(), sg = stTxt(fl.Stage).trim().toLowerCase();
    if (station && st.toLowerCase() !== station) return;
    if (stage && sg !== stage) return;
    const nm = stTxt(fl.Who).trim();
    if (who && nm.toLowerCase() !== who) return;
    if (from && day < from) return;
    if (to && day > to) return;
    const w = isoWeek(day);
    if (week && w.key !== week) return;
    if (dow >= 0 && w.dow !== dow) return;
    const c = {};
    let total = 0;
    cols.forEach(k => {
      const n = Math.max(0, Math.round(stNum(fl[k[0]], 0)));
      c[k[0]] = n; total += n;
    });
    const wt = stNum(fl.WeekTarget, null);
    out.push({ id: stTxt(it.id), title: stTxt(fl.Title), station: st, stage: sg,
               day: day, week: w.key, monday: w.monday, weekday: w.weekday, dow: w.dow,
               who: nm, counts: c, total: total, note: stTxt(fl.Note),
               weekTarget: wt == null || !isFinite(wt) ? null : Math.max(0, Math.round(wt)),
               savedAt: stTxt(fl.SavedAt), editedBy: stTxt(fl.EditedBy), editedAt: stTxt(fl.EditedAt) });
  });
  out.sort((a, b) => (a.day < b.day ? 1 : a.day > b.day ? -1 :
                      (a.who < b.who ? -1 : a.who > b.who ? 1 : 0)));
  return out;
}
/** One person's sheets this ISO week, added up - the "This week: N of T" on the
    tablet and the week subtotals in the office. */
function dayWeekTotal(rows, who, week) {
  const nm = stTxt(who).trim().toLowerCase(), wk = stTxt(week).trim();
  return (rows || []).reduce((n, r) =>
    n + ((!nm || r.who.toLowerCase() === nm) && (!wk || r.week === wk) ? r.total : 0), 0);
}
/** The rows grouped into ISO weeks, newest week first, each with its subtotal,
    the target that was in force and the difference.

    A week's target is the `WeekTarget` on its most recently saved sheet - so
    changing the target today cannot rewrite what last week was measured
    against. A week with no stored target falls back to the live one, which is
    what the current week needs before anybody has saved a sheet in it. */
function dayWeeks(rows, counts, liveTarget) {
  const cols = counts || [];
  const by = {}, order = [];
  (rows || []).forEach(r => {
    let g = by[r.week];
    if (!g) {
      g = by[r.week] = { week: r.week, monday: r.monday, rows: [], counts: {}, total: 0,
                         target: null, diff: null, targetAt: "" };
      cols.forEach(c => { g.counts[c[0]] = 0; });
      order.push(r.week);
    }
    g.rows.push(r);
    cols.forEach(c => { g.counts[c[0]] += r.counts[c[0]] || 0; });
    g.total += r.total;
    if (r.weekTarget != null && (!g.targetAt || atCmp(r.savedAt, g.targetAt) >= 0)) {
      g.target = r.weekTarget; g.targetAt = r.savedAt;
    }
  });
  const live = stNum(liveTarget, null);
  order.sort((a, b) => (a < b ? 1 : a > b ? -1 : 0));
  return order.map(k => {
    const g = by[k];
    if (g.target == null && live != null && isFinite(live)) g.target = Math.max(0, Math.round(live));
    g.diff = g.target == null ? null : g.total - g.target;
    return g;
  });
}

/* ---- the weekly target ---- */
const targetTitle = (station, stage) =>
  stTxt(station).trim() + "|" + stTxt(stage).trim().toLowerCase();
/** The three columns an office save writes beside the Title. */
function targetFields(n, who, at) {
  return { WeeklyTarget: Math.max(0, Math.round(stNum(n, 0))),
           SetBy: stTxt(who), SetAt: stTxt(at) || new Date().toISOString() };
}
/** The target in force for one station-stage, or null. */
function targetOf(items, station, stage) {
  const want = targetTitle(station, stage).toLowerCase();
  let hit = null;
  (items || []).forEach(it => {
    if (!it) return;
    const fl = it.fields || {};
    if (stTxt(fl.Title).trim().toLowerCase() !== want) return;
    if (!hit || stItemAge(it, hit) < 0) hit = it;      // the oldest, as everywhere else
  });
  if (!hit) return null;
  const fl = hit.fields || {};
  const n = stNum(fl.WeeklyTarget, null);
  if (n == null || !isFinite(n)) return null;
  return { id: stTxt(hit.id), target: Math.max(0, Math.round(n)),
           by: stTxt(fl.SetBy), at: stTxt(fl.SetAt) };
}

/* ---- what the glass station says about itself in a report -------------------
   The report is assembled by one pure function (export.js, stationReport) with
   no station-specific branch in it: everything a station has to say about
   itself comes off its own definition. This is the glass station's half of
   that, and welding-core.js has the matching one. A third station writes one
   more of these and appears in the report by being defined.                  */
/** One row per job on the board, at this stage, in the report's own order.

    THE CUSTOMER GOES THROUGH THE STRIP (rule 3, review 2026-09-21). It is free
    text off the sheet like any other: the office types a site contact into a
    customer name often enough that welding's feeder has stripped it since the
    day that station shipped, and a report that carried it as typed would put a
    phone number in an exported file - which is the one thing rule 3 forbids
    outright. 60 characters, the width the column is read at. */
const REPORT_CUSTOMER_MAX = 60;
function glassReportJobs(data, stage) {
  const k = stTxt(stage).trim().toLowerCase();
  const board = (data && data.board) || [];
  const rows = [], jobs = [];
  board.forEach(g => {
    const total = Math.max(0, Math.round(stNum(g[STAGE_TOTAL_ROW[k]], 0)));
    const done = stClamp(g[STAGE_ROW[k]], total);
    jobs.push({ job: g.job, done: done, total: total });
    rows.push([g.job, stripContact(g.customer, REPORT_CUSTOMER_MAX),
               g.officeDone ? "the office says complete" : g.active ? "on the floor" : "off the board",
               total, done, Math.max(0, total - done),
               g.by && g.by[k] ? g.by[k] : g.doneBy, (g.at && g.at[k]) || g.doneAt,
               total > 0 && done >= total ? "Yes" : "No"]);
  });
  return { columns: ["Job", "Customer", "Status", "Total", "Done at this stage", "Left",
                     "Last moved by", "When", "Complete"],
           rows: rows, jobs: jobs };
}

const ST = {
  GLASS, stDef,
  STATION_LIST, PEOPLE_LIST, LOG_LIST, STATION_SITE, STATION_NAME,
  STAGES, STAGE_KEYS, STAGE_FIELD, STAGE_ROW, STAGE_BY, STAGE_AT, stageLabel,
  LEGACY_STAGE_LABELS,
  ALL_STAGES, ALL_STAGE_KEYS, STAGE_TOTAL_ROW, TUFF_STAGE,
  COLOUR_TYPES, stageComplete, glassColours, tuffSpoken, tuffOwed, floorStamp,
  STATION_FIELDS, FEEDER_FIELDS, FLOOR_FIELDS, PEOPLE_FIELDS, LOG_FIELDS,
  SEED_FIELDS, FEEDER_WRITES, GLASS_TYPE, TOTAL_TYPES,
  OFFICE_CLEAR_FIELDS, officeClearFields, floorWorkToClear, clearWords, clearWarning,
  PEOPLE_FIELDS_OFFICE, CUSTOMER_MAX, PERSON_LOCK_MS, REFRESH_MS, LOG_DAYS, logSince,
  COMMENT_LIST, COMMENT_FIELDS, COMMENT_MAX, COMMENT_POLL_MS,
  COMMENT_MISSING_FLOOR, COMMENT_MISSING_OFFICE, COMMENT_UNREACHABLE, COMMENT_CHECKING,
  COMMENT_EMPTY_FLOOR, COMMENT_EMPTY_OFFICE, COMMENT_UNSENT, COMMENT_KEPT,
  commentTitle, commentFields, commentRows, commentAgo, commentChangeWords, stationComments, atCmp,
  stripContact, STRIP_MAX,
  DAY_LIST, TARGET_LIST, DAY_BASE_FIELDS, TARGET_FIELDS, DAY_NOTE_MAX,
  DAY_MISSING_FLOOR, DAY_MISSING_OFFICE, TARGET_MISSING_OFFICE, DAY_UNREACHABLE, DAY_SAVED_WORDS,
  daySheetOf, daySheetStages, dayFieldsFor, dayCountShort, dayKey, stClock, isoWeek, DAY_NAMES,
  dayTitle, dayCount, DAY_COUNT_MAX, dayFields, dayOfficeFields, dayRows, dayWeekTotal, dayWeeks,
  targetTitle, targetFields, targetOf, glassReportJobs, REPORT_CUSTOMER_MAX,
  inProduction, sectionInProduction, glassTotal, tuffTotal, officeSeed, officeComplete,
  glassSlice, feederFields, seedFields, feedPlan, sliceHash,
  jobBoard, jobRecord, jobRecords, jobKey: stKey, boardFilter, glassWords, leftWords,
  stageLeft, heldStages, jobLefts, boardLefts, applyTap, boardDiff, mergeDelta,
  stationPeople, canStage, pinOk, personExpired,
  floorOnly, tapFields, logFields, logRows, logFilter, logCounts, logLast
};
if (typeof window !== "undefined") window.ST = ST;
else if (typeof globalThis !== "undefined") globalThis.ST = ST;
if (typeof module !== "undefined" && module.exports) module.exports = ST;
