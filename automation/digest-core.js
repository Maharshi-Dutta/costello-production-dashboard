/* digest-core.js - the pure logic behind the "job alerts" email digest.

   No I/O, no ExcelScript, no Node globals: every function takes plain arrays
   (the values matrices of the four sheets) and returns plain data. The same
   code runs in two places:

     * node automation/test_digest.js   - requires this file directly
     * automation/alerts-digest.ts      - an Office Script, which cannot import
                                          anything, so the block between the
                                          CORE markers below is copied into it.

   Keep them in sync by regenerating, never by hand-editing the .ts:

       node automation/build-script.js

   The type comments are annotations for the generated TypeScript:
   build-script.js rewrites  x /*:string* /  into  x: string , and Node ignores
   them. Add one wherever the Office Scripts compiler could not infer a type
   (function parameters, and empty arrays that are filled later).

   HARD RULE: no email address, person's name or domain appears in this file.
   The admin address (which defines the only allowed recipient domain) and the
   dashboard URL are read at run time from the Dashboard Config sheet.        */

/* --- BEGIN SHARED CORE (copied into alerts-digest.ts by build-script.js) --- */

/* Job numbers look like C5255 / AB1234. Same rule as parser.js. */
const JOB_RE = /^[A-Z]{1,2}\d{3,5}$/;

const DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** One cell of a values matrix as text ("" for blank). */
function cell(row /*:Cell[]*/, c /*:number*/) {
  if (!row) return "";
  const v = row[c];
  return v == null ? "" : String(v);
}

/** Header text reduced to lower-case words, exactly as parser.js does it, so
    "COMMENT " and "Comment" and "comment:" all compare equal. */
function norm(v /*:string*/) {
  return String(v == null ? "" : v).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

/** Text -> HTML. Also used for attribute values, so quotes are escaped too. */
function escapeHtml(v /*:string*/) {
  return String(v == null ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function newBlock(idx /*:number*/, name /*:string*/, divider /*:number*/) {
  const jobs /*:JobRef[]*/ = [];
  return { idx: idx, name: name, divider: divider, jobs: jobs, last: divider };
}

/** Sections of the Production sheet, ported from blocksFromValues in
    parser.js - the divider rule must stay identical or the dashboard and the
    mailer would disagree about which jobs are finished.
    A divider is a row whose text matches AND whose job-number cell (column C)
    is not a job number, so a job whose comment says "not sent to floor" can
    never be mistaken for a section break. The first "not sent to floor" opens
    "Ready to fit", the second (and any after it) "In production".            */
function blocksFromValues(matrix /*:Cell[][]*/) {
  const names = ["Can sell as second hand"];
  const blocks /*:Block[]*/ = [newBlock(0, names[0], 0)];
  let cur /*:Block*/ = blocks[0];
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i] || [], r = i + 1;
    const j = cell(row, 2).trim().toUpperCase();
    const isJob = JOB_RE.test(j);
    let line = "";
    for (let c = 0; c < 11; c++) line += " " + cell(row, c);
    line = norm(line);
    let hit = "";
    if (!isJob) {
      if (line.indexOf("can sell as second hand") >= 0) hit = "";           // labels the block above
      else if (line.indexOf("customers won t take") >= 0 || line.indexOf("customers wont take") >= 0)
        hit = "Ready, customer won't take";
      else if (line.indexOf("collect or supply only") >= 0) hit = "Collect & supply only";
      else if (line.indexOf("not sent to floor") >= 0)
        /* display names only - the sheet's own divider text is never changed */
        hit = names.indexOf("Ready to fit") < 0 ? "Ready to fit" : "In production";
    }
    if (hit) { names.push(hit); cur = newBlock(names.length - 1, hit, r); blocks.push(cur); continue; }
    if (isJob) { cur.jobs.push({ row: r, id: j }); cur.last = r; }
  }
  return { names: names, blocks: blocks };
}

/** Job number -> its 1-based row, for the jobs sitting in "In production".
    Everything else is finished (or has left the sheet) and is never emailed. */
function productionJobRows(matrix /*:Cell[][]*/) {
  const rows = Object.create(null);      // no prototype: a job id can never collide with "constructor"
  const B = blocksFromValues(matrix);
  for (let i = 0; i < B.blocks.length; i++) {
    const b = B.blocks[i];
    if (b.name !== "In production") continue;
    for (let k = 0; k < b.jobs.length; k++) rows[b.jobs[k].id] = b.jobs[k].row;
  }
  return rows;
}

/** The (group, sub) header row pair, ported from headerRows in parser.js. */
function headerRowsFromValues(matrix /*:Cell[][]*/) {
  for (let r = 0; r < 6; r++) {
    const row = matrix[r] || [];
    const maxC = Math.min(row.length, 40);
    for (let c = 0; c < maxC; c++) {
      const n = norm(cell(row, c));
      if (n.indexOf("dates on contract") >= 0 || n === "dates") return [r, r + 1];
    }
  }
  return [1, 2];                         // rows 2 and 3, the usual layout
}

/** The three comment-bearing columns, found by header text and never by
    letter (they move). Same matching as mapSheet in parser.js: an exact
    "comment" for the COMMENT column (first one wins), a contains-match for
    "brendan" and for "notes of specials" (the last one wins, as in parser.js).
    -1 means the column is not on this sheet.                                 */
function commentColumns(matrix /*:Cell[][]*/) {
  const hr = headerRowsFromValues(matrix);
  const g = matrix[hr[0]] || [];
  const cols /*:Cols*/ = { comment: -1, brendan: -1, specials: -1 };
  for (let c = 0; c < g.length; c++) {
    const lab = norm(cell(g, c));
    if (!lab) continue;
    if (lab === "comment" && cols.comment < 0) cols.comment = c;
    if (lab.indexOf("brendan") >= 0) cols.brendan = c;
    if (lab.indexOf("notes of specials") >= 0) cols.specials = c;
  }
  return cols;
}

/** Comments written through the dashboard: Dashboard Log rows whose column D
    is "Comment" (A=When, B=Who, C=Job, D=What changed, E=From, F=To). Oldest
    first, because the log is only ever appended to. The header row cannot be
    mistaken for data - its column D reads "What changed".                    */
function logComments(matrix /*:Cell[][]*/) {
  const byJob = Object.create(null);
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const job = cell(row, 2).trim().toUpperCase();
    if (!job || norm(cell(row, 3)) !== "comment") continue;
    const text = cell(row, 5).trim();
    if (!text) continue;
    if (!byJob[job]) byJob[job] = [];
    byJob[job].push(text);
  }
  return byJob;
}

/** Dashboard Alerts rows (Job | Email | Added by | When). Rows without an "@"
    - the header row, and rows blanked by a removal - simply fall out, and a
    duplicated (job, email) pair is kept only once.                           */
function readSubscriptions(matrix /*:Cell[][]*/) {
  const out /*:Sub[]*/ = [];
  const seen = Object.create(null);
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const job = cell(row, 0).trim().toUpperCase();
    const email = cell(row, 1).trim().toLowerCase();
    if (!job || email.indexOf("@") < 1) continue;
    const key = job + "|" + email;
    if (seen[key]) continue;
    seen[key] = 1;
    out.push({ job: job, email: email, who: cell(row, 2).trim(), when: cell(row, 3).trim() });
  }
  return out;
}

/** Dashboard Config is a Key | Value sheet. Keys are matched loosely (case
    and punctuation ignored) so "Dashboard URL" and "dashboardUrl" both work.
    Missing keys come back as "". First value wins.                           */
function readConfig(matrix /*:Cell[][]*/) {
  const conf /*:Conf*/ = { admin: "", dashboardUrl: "" };
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const key = norm(cell(row, 0)).replace(/ /g, "");
    const value = cell(row, 1).trim();
    if (!key || !value) continue;
    if (key === "admin" && !conf.admin) conf.admin = value;
    if (key === "dashboardurl" && !conf.dashboardUrl) conf.dashboardUrl = value;
  }
  return conf;
}

/** Every comment on one job: the Production sheet's COMMENT / Notes from
    Brendan / Notes of specials cells, then the dashboard comments from the
    log (newest last). Blank ones are dropped.                                */
function commentsForJob(prodMatrix /*:Cell[][]*/, cols /*:Cols*/, row /*:number*/,
                        logByJob /*:any*/, jobId /*:string*/) {
  const out /*:string[]*/ = [];
  const r = prodMatrix[row - 1] || [];
  const from = [cols.comment, cols.brendan, cols.specials];
  for (let i = 0; i < from.length; i++) {
    if (from[i] < 0) continue;
    const t = cell(r, from[i]).trim();
    if (t) out.push(t);
  }
  const logged = logByJob[jobId];
  if (logged) for (let i = 0; i < logged.length; i++) if (logged[i]) out.push(String(logged[i]));
  return out;
}

/** "Fri 5 Sep" - the run date, in the flow's own time zone. */
function formatDate(d /*:Date*/) {
  return DAY_NAMES[d.getDay()] + " " + d.getDate() + " " + MONTH_NAMES[d.getMonth()];
}

function plural(n /*:number*/) { return n === 1 ? " job" : " jobs"; }

function subjectFor(count /*:number*/, d /*:Date*/) {
  return "Job alerts: " + count + plural(count) + " - " + formatDate(d);
}

/** Only http(s) links are put in the mail; anything else in the config cell is
    ignored rather than turned into a link. */
function safeUrl(url /*:string*/) {
  const u = String(url == null ? "" : url).trim();
  return (u.indexOf("http://") === 0 || u.indexOf("https://") === 0) ? u : "";
}

/** The mail body: an intro line, then per job a heading and a bullet list of
    its comments (nothing else per job), then the dashboard link. */
function htmlFor(jobs /*:JobDigest[]*/, dashboardUrl /*:string*/) {
  const n = jobs.length;
  const url = safeUrl(dashboardUrl);
  let h = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#17171A">';
  h += "<p>You are subscribed to " + n + plural(n) + " still in production. " +
       "Comments are listed oldest first.</p>";
  for (let i = 0; i < jobs.length; i++) {
    const job = jobs[i];
    h += '<h3 style="margin:18px 0 4px 0;font-size:15px">' + escapeHtml(job.id) + "</h3>";
    h += '<ul style="margin:0;padding-left:20px">';
    if (job.comments.length) {
      for (let k = 0; k < job.comments.length; k++)
        h += "<li>" + escapeHtml(job.comments[k]) + "</li>";
    } else {
      h += "<li><em>(no comments)</em></li>";
    }
    h += "</ul>";
  }
  h += '<p style="margin-top:22px;font-size:12.5px;color:#6B6B70">';
  h += url ? '<a href="' + escapeHtml(url) + '">Open the dashboard</a>'
           : "Open the dashboard for the full picture.";
  h += "</p></div>";
  return h;
}

/** The whole job: four values matrices in, one digest per address out.
    - only addresses whose domain equals the admin's domain are considered
      (the dashboard refuses any other, so they should never exist);
    - only jobs still in the "In production" section are listed;
    - an address with nothing left to say gets no entry at all.
    Bad rows are skipped, never thrown on. With no usable admin address there
    is no domain rule, so nothing is sent.                                    */
function buildDigests(configValues /*:Cell[][]*/, alertsValues /*:Cell[][]*/,
                      productionValues /*:Cell[][]*/, logValues /*:Cell[][]*/,
                      now /*:Date*/) {
  const digests /*:Digest[]*/ = [];
  const conf = readConfig(configValues || []);
  const admin = conf.admin.toLowerCase();
  const at = admin.indexOf("@");
  if (at < 1) return digests;
  const domain = admin.slice(at + 1);
  if (!domain) return digests;

  const prod = productionValues || [];
  const rows = productionJobRows(prod);
  const cols = commentColumns(prod);
  const logByJob = logComments(logValues || []);
  const subs = readSubscriptions(alertsValues || []);

  const byEmail = Object.create(null);
  const emails /*:string[]*/ = [];
  for (let i = 0; i < subs.length; i++) {
    const s = subs[i];
    if (s.email.slice(s.email.indexOf("@") + 1) !== domain) continue;  // never mail outside the admin's domain
    const row = rows[s.job];
    if (row === undefined) continue;                                   // finished, or gone from the sheet
    if (!byEmail[s.email]) { byEmail[s.email] = []; emails.push(s.email); }
    byEmail[s.email].push({ id: s.job, row: row });
  }
  emails.sort();

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    const picked = byEmail[email];
    picked.sort(function (a /*:any*/, b /*:any*/) { return a.row - b.row; });   // sheet order
    const list /*:JobDigest[]*/ = [];
    for (let k = 0; k < picked.length; k++)
      list.push({ id: picked[k].id,
                  comments: commentsForJob(prod, cols, picked[k].row, logByJob, picked[k].id) });
    if (!list.length) continue;
    digests.push({ email: email, subject: subjectFor(list.length, now),
                   html: htmlFor(list, conf.dashboardUrl) });
  }
  return digests;
}

/* --- END SHARED CORE --- */

/* Node only. Office Scripts never sees this: build-script.js copies only the
   block between the CORE markers. */
if (typeof module !== "undefined") module.exports = {
  JOB_RE: JOB_RE, cell: cell, norm: norm, escapeHtml: escapeHtml,
  blocksFromValues: blocksFromValues, productionJobRows: productionJobRows,
  headerRowsFromValues: headerRowsFromValues, commentColumns: commentColumns,
  logComments: logComments, readSubscriptions: readSubscriptions,
  readConfig: readConfig, commentsForJob: commentsForJob,
  formatDate: formatDate, subjectFor: subjectFor, safeUrl: safeUrl,
  htmlFor: htmlFor, buildDigests: buildDigests
};
