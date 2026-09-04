/* Sign-in + Microsoft Graph access for the Costello production dashboard.
   Reads the live workbook by downloading it (full fidelity, including fills);
   writes surgically through the Excel API so the file is never rewritten. */

const CLIENT_ID = "a989939b-17f3-4c9c-adb7-4d8338f4878a";
const TENANT_ID = "cb4cfc4c-96f4-44c0-b37b-a467826f86d6";
const SCOPES = ["Files.ReadWrite.All", "User.Read"];
const SITE_PATH = "costellowindowsie.sharepoint.com:/sites/ProductionProgress";
const FILE_MATCH = "production work in progress";
const G = "https://graph.microsoft.com/v1.0";

let msalApp = null;
function app() {
  if (msalApp) return msalApp;
  if (typeof msal === "undefined") throw new Error("The Microsoft sign-in library did not load.");
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/" + TENANT_ID,
      redirectUri: window.location.origin   // no path: must match the Entra entry exactly
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
  });
  return msalApp;
}

let account = null;
let sessionId = null;          // workbook session: keeps the file warm server-side

async function initAuth() {
  await app().initialize();
  const res = await app().handleRedirectPromise();
  if (res && res.account) account = res.account;
  if (!account) {
    const all = app().getAllAccounts();
    if (all.length) account = all[0];
  }
  return account;
}

async function signIn() {
  const res = await app().loginPopup({ scopes: SCOPES, prompt: "select_account" });
  account = res.account;
  return account;
}

function signOut() {
  return app().logoutPopup({ account: account });
}

let tokenOverride = null;     // rehearsal harness only: a token without the sign-in library
async function token() {
  if (tokenOverride) return tokenOverride();
  if (!account) throw new Error("not signed in");
  try {
    const r = await app().acquireTokenSilent({ scopes: SCOPES, account: account });
    return r.accessToken;
  } catch (e) {
    const r = await app().acquireTokenPopup({ scopes: SCOPES, account: account });
    return r.accessToken;
  }
}

async function headers(extra) {
  const h = { Authorization: "Bearer " + (await token()) };
  if (sessionId) h["workbook-session-id"] = sessionId;
  return Object.assign(h, extra || {});
}

/* Retries: 5xx while Excel loads a big workbook, and - importantly - the
   InvalidSession 400 that Graph returns once a workbook session goes idle.
   That one arrives as a 400, so it needs handling separately from server
   errors: drop the stale session, open a fresh one, and try again. */
let openingSession = false;

async function call(method, path, body, asBuffer) {
  let lastStatus = 0, lastText = "", sessionRetried = false;
  for (let a = 0; a < 5; a++) {
    const init = { method, headers: await headers(body ? { "Content-Type": "application/json" } : null) };
    if (body) init.body = JSON.stringify(body);
    const r = await fetch(G + path, init);
    if (r.ok) {
      if (asBuffer) return await r.arrayBuffer();
      if (r.status === 204) return {};
      const t = await r.text();
      return t ? JSON.parse(t) : {};
    }
    lastStatus = r.status;
    lastText = await r.text();

    if (!sessionRetried && !openingSession &&
        (lastText.indexOf("InvalidSession") >= 0 || lastText.indexOf("invalidSessionReCreatable") >= 0)) {
      sessionRetried = true;
      sessionId = null;
      try { await openSession(); } catch (e) { sessionId = null; }  // stateless still works
      continue;
    }
    if ([429, 500, 503, 504].indexOf(r.status) >= 0) {
      await new Promise(s => setTimeout(s, 2000 + a * 2500));
      continue;
    }
    break;
  }
  throw new Error(method + " " + path.split("/workbook")[1] + " -> " + lastStatus + " " + lastText.slice(0, 200));
}

/* ---- the workbook ---- */
let fileRef = null;

async function findFile() {
  if (fileRef) return fileRef;
  const cached = localStorage.getItem("cw_fileref");
  if (cached) { fileRef = JSON.parse(cached); return fileRef; }
  const site = await call("GET", "/sites/" + SITE_PATH);
  const kids = await call("GET", "/sites/" + site.id + "/drive/root/children");
  const hit = (kids.value || []).find(i =>
    i.name.toLowerCase().endsWith(".xlsx") && i.name.toLowerCase().indexOf(FILE_MATCH) >= 0);
  if (!hit) throw new Error("Workbook not found in the ProductionProgress library.");
  fileRef = {
    siteId: site.id, itemId: hit.id, name: hit.name,
    base: "/sites/" + site.id + "/drive/items/" + hit.id + "/workbook",
    content: "/sites/" + site.id + "/drive/items/" + hit.id + "/content",
    meta: "/sites/" + site.id + "/drive/items/" + hit.id
  };
  localStorage.setItem("cw_fileref", JSON.stringify(fileRef));
  return fileRef;
}

async function openSession() {
  const f = await findFile();
  openingSession = true;
  try {
    const s = await call("POST", f.base + "/createSession", { persistChanges: true });
    sessionId = s.id;
  } catch (e) {
    sessionId = null;      // a session is an optimisation; Graph works without one
  } finally {
    openingSession = false;
  }
  return sessionId;
}

async function lastModified() {
  const f = await findFile();
  const m = await call("GET", f.meta + "?$select=lastModifiedDateTime,lastModifiedBy");
  return {
    at: m.lastModifiedDateTime,
    by: (m.lastModifiedBy && m.lastModifiedBy.user && m.lastModifiedBy.user.displayName) || ""
  };
}

async function downloadWorkbook() {
  const f = await findFile();
  const buf = await call("GET", f.content, null, true);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

/* ---- version history ------------------------------------------------------
   SharePoint keeps every save as a version. Listing and downloading them is
   read-only. restoreVersion is the ONE deliberate whole-file write in this
   app: it makes the chosen version current, and SharePoint keeps the state it
   replaced as a new version, so a rollback is itself rollable.               */
async function listVersions(top) {
  const f = await findFile();
  const r = await call("GET", f.meta + "/versions?$top=" + (top || 60));
  return (r.value || []).map(v => ({
    id: v.id, at: v.lastModifiedDateTime, size: v.size || 0,
    by: (v.lastModifiedBy && v.lastModifiedBy.user && v.lastModifiedBy.user.displayName) || "unknown"
  }));
}
async function downloadVersion(versionId) {
  const f = await findFile();
  const buf = await call("GET", f.meta + "/versions/" + versionId + "/content", null, true);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}
async function restoreVersion(versionId) {
  const f = await findFile();
  sessionId = null;                         // the file is about to be replaced under any open session
  return call("POST", f.meta + "/versions/" + versionId + "/restoreVersion", {});
}

/* ---- one writer at a time per dashboard sheet -----------------------------
   Every upsert below is "read usedRange, work out the row, write it". Two of
   those running at once read the same rowCount and write the same row, so one
   change silently replaces the other. Chaining them per sheet costs nothing
   (they are a few hundred milliseconds each) and makes that impossible.     */
const chains = {};
function serialised(sheet, fn) {
  const prev = chains[sheet] || Promise.resolve();
  const next = prev.then(fn, fn);                 // an earlier failure must not stall the queue
  chains[sheet] = next.catch(() => {});
  return next;
}
/* The ensure*Sheet functions below remember the promise, not a flag: a second
   caller arriving while the first is still creating the sheet waits for it
   instead of adding the sheet a second time. A failure clears the memo so the
   next caller retries. */

/* ---- the audit log -------------------------------------------------------
   HARD RULE: every function below addresses LOG_SHEET and nothing else. The
   sheet name is a constant, never a parameter, so no call site can point this
   at Production or any other sheet.                                          */
const LOG_SHEET = "Dashboard Log";
const LOG_HEADERS = [["When", "Who", "Job", "What changed", "From", "To"]];
let logReady = null;

function ensureLogSheet() {
  if (logReady) return logReady;
  logReady = makeLogSheet();
  logReady.catch(() => { logReady = null; });
  return logReady;
}
async function makeLogSheet() {
  const f = await findFile();
  const ws = await call("GET", f.base + "/worksheets");
  const exists = (ws.value || []).some(w => w.name === LOG_SHEET);
  if (!exists) {
    await call("POST", f.base + "/worksheets/add", { name: LOG_SHEET });
    await call("PATCH", f.base + "/worksheets('" + LOG_SHEET + "')/range(address='A1:F1')", { values: LOG_HEADERS });
    await call("PATCH", f.base + "/worksheets('" + LOG_SHEET + "')/range(address='A1:F1')/format/font", { bold: true });
    await call("PATCH", f.base + "/worksheets('" + LOG_SHEET + "')/range(address='A1:F1')/format/fill", { color: "#17171A" });
    await call("PATCH", f.base + "/worksheets('" + LOG_SHEET + "')/range(address='A1:F1')/format/font", { color: "#FFFFFF" });
    /* Excel turns a date-like string into a serial number; give the column a
       date format so it reads properly and still sorts as a real date. */
    const fmt = []; for (let i = 0; i < 1999; i++) fmt.push(["@"]);   // text: no locale guessing
    await call("PATCH", f.base + "/worksheets('" + LOG_SHEET + "')/range(address='A2:A2000')", { numberFormat: fmt });
    const widths = { A: 130, B: 220, C: 70, D: 230, E: 150, F: 190 };
    for (const col in widths)
      await call("PATCH", f.base + "/worksheets('" + LOG_SHEET + "')/range(address='" + col + ":" + col + "')/format", { columnWidth: widths[col] });
  }
  return true;
}

/** Append one row to the log sheet. Never writes anywhere else. */
async function appendLog(who, job, what, from, to) {
  try {
    await ensureLogSheet();
    const f = await findFile();
    /* queued: two log lines a moment apart would otherwise both take the same row */
    return await serialised(LOG_SHEET, async () => {
      const used = await call("GET", f.base + "/worksheets('" + LOG_SHEET + "')/usedRange?$select=rowCount");
      const next = (used.rowCount || 1) + 1;
      const when = new Date();
      const pad = n => (n < 10 ? "0" : "") + n;
      /* ISO order, written as text. "03/09/2026" is read as 9 March by a US-locale
         Excel and 3 September by an Irish one; this is the same everywhere, and
         still sorts correctly because ISO sorts lexicographically. */
      const stampStr = when.getFullYear() + "-" + pad(when.getMonth() + 1) + "-" + pad(when.getDate()) +
                       " " + pad(when.getHours()) + ":" + pad(when.getMinutes());
      await call("PATCH", f.base + "/worksheets('" + LOG_SHEET + "')/range(address='A" + next + ":F" + next + "')",
        { values: [[stampStr, who || "unknown", job, what, String(from == null ? "" : from), String(to == null ? "" : to)]],
          numberFormat: [["@", "@", "@", "@", "@", "@"]] });
      return true;
    });
  } catch (e) {
    console.warn("log append failed:", e.message);
    return false;   // never let a logging failure block the real edit
  }
}

/* ---- saved views ---------------------------------------------------------
   Stores only decisions - "someone put job X in group Y" - never job data.
   Everything about a job is always read live from Production, so this cannot
   go stale. Addresses VIEWS_SHEET by constant, like the log.               */
const VIEWS_SHEET = "Dashboard Views";
const VIEWS_HEADERS = [["View", "Job", "Group", "Order", "Set by", "When"]];
let viewsReady = null;

function ensureViewsSheet() {
  if (viewsReady) return viewsReady;
  viewsReady = makeViewsSheet();
  viewsReady.catch(() => { viewsReady = null; });
  return viewsReady;
}
async function makeViewsSheet() {
  const f = await findFile();
  const ws = await call("GET", f.base + "/worksheets");
  if (!(ws.value || []).some(w => w.name === VIEWS_SHEET)) {
    await call("POST", f.base + "/worksheets/add", { name: VIEWS_SHEET });
    const S = f.base + "/worksheets('" + VIEWS_SHEET + "')";
    await call("PATCH", S + "/range(address='A1:F1')", { values: VIEWS_HEADERS });
    await call("PATCH", S + "/range(address='A1:F1')/format/font", { bold: true, color: "#FFFFFF" });
    await call("PATCH", S + "/range(address='A1:F1')/format/fill", { color: "#17171A" });
    const widths = { A: 170, B: 80, C: 90, D: 60, E: 220, F: 130 };
    for (const c in widths)
      await call("PATCH", S + "/range(address='" + c + ":" + c + "')/format", { columnWidth: widths[c] });
  }
  return true;
}

/** Put one job in one group of one view. Updates the existing line if there is
    one, otherwise appends. Only ever writes to VIEWS_SHEET. */
async function saveAssignment(view, job, group, order, who) {
  await ensureViewsSheet();
  const f = await findFile();
  const S = f.base + "/worksheets('" + VIEWS_SHEET + "')";
  const used = await call("GET", S + "/usedRange?$select=values,rowCount");
  const rows = used.values || [];
  let target = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() === view && String(rows[i][1] || "").trim().toUpperCase() === job.toUpperCase()) {
      target = i + 1; break;
    }
  }
  if (!target) target = (used.rowCount || 1) + 1;
  const d = new Date(), p = n => (n < 10 ? "0" : "") + n;
  const when = d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + " " + p(d.getHours()) + ":" + p(d.getMinutes());
  await call("PATCH", S + "/range(address='A" + target + ":F" + target + "')", {
    values: [[view, job, group === null || group === undefined ? "" : String(group), String(order == null ? "" : order), who || "", when]],
    numberFormat: [["@", "@", "@", "@", "@", "@"]]
  });
  return target;
}

/** Remove a job from a view (blank its line). */
async function clearAssignment(view, job) {
  await ensureViewsSheet();
  const f = await findFile();
  const S = f.base + "/worksheets('" + VIEWS_SHEET + "')";
  const used = await call("GET", S + "/usedRange?$select=values");
  const rows = used.values || [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() === view && String(rows[i][1] || "").trim().toUpperCase() === job.toUpperCase()) {
      await call("PATCH", S + "/range(address='A" + (i + 1) + ":F" + (i + 1) + "')", { values: [["", "", "", "", "", ""]] });
      return true;
    }
  }
  return false;
}

/* ---- checkpoint progress -------------------------------------------------
   Excel keeps only the colour of a cell, so "6 of 10" has to be written down
   somewhere: here, one upserted row per (Job, Item). Everything below
   addresses PROGRESS_SHEET by constant, like the log and the views. */
const PROGRESS_SHEET = "Dashboard Progress";
const PROGRESS_HEADERS = [["Job", "Item", "Done", "Total", "Who", "When"]];
const TEXT6 = [["@", "@", "@", "@", "@", "@"]];
let progressReady = null;

/* same ISO-order text stamp as the log, for the same locale reason */
function nowStamp() {
  const d = new Date(), p = n => (n < 10 ? "0" : "") + n;
  return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + " " + p(d.getHours()) + ":" + p(d.getMinutes());
}

function ensureProgressSheet() {
  if (progressReady) return progressReady;
  progressReady = makeProgressSheet();
  progressReady.catch(() => { progressReady = null; });
  return progressReady;
}
async function makeProgressSheet() {
  const f = await findFile();
  const ws = await call("GET", f.base + "/worksheets");
  if (!(ws.value || []).some(w => w.name === PROGRESS_SHEET)) {
    await call("POST", f.base + "/worksheets/add", { name: PROGRESS_SHEET });
    const S = f.base + "/worksheets('" + PROGRESS_SHEET + "')";
    await call("PATCH", S + "/range(address='A1:F1')", { values: PROGRESS_HEADERS });
    await call("PATCH", S + "/range(address='A1:F1')/format/font", { bold: true, color: "#FFFFFF" });
    await call("PATCH", S + "/range(address='A1:F1')/format/fill", { color: "#17171A" });
    /* text throughout: a count is not a sum and "2026-09-04 10:42" is not a
       date Excel should re-interpret in whatever locale it happens to run in */
    const fmt = []; for (let i = 0; i < 1999; i++) fmt.push(["@", "@", "@", "@", "@", "@"]);
    await call("PATCH", S + "/range(address='A2:F2000')", { numberFormat: fmt });
    const widths = { A: 70, B: 210, C: 60, D: 60, E: 220, F: 130 };
    for (const c in widths)
      await call("PATCH", S + "/range(address='" + c + ":" + c + "')/format", { columnWidth: widths[c] });
  }
  return true;
}

/** Store one item's count. Updates the existing (Job, Item) line, else appends. */
async function saveProgress(job, item, done, total, who) {
  await ensureProgressSheet();
  const f = await findFile();
  const S = f.base + "/worksheets('" + PROGRESS_SHEET + "')";
  return serialised(PROGRESS_SHEET, async () => {
    const used = await call("GET", S + "/usedRange?$select=values,rowCount");
    const rows = used.values || [];
    let target = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim().toUpperCase() === String(job).toUpperCase() &&
          String(rows[i][1] || "").trim() === String(item)) { target = i + 1; break; }
    }
    if (!target) target = (used.rowCount || 1) + 1;
    await call("PATCH", S + "/range(address='A" + target + ":F" + target + "')", {
      values: [[job, item, String(done), String(total), who || "", nowStamp()]],
      numberFormat: TEXT6
    });
    return target;
  });
}

/** Store several items of one job: one read of the sheet, then all the rows in
    a single batch - a whole group of taps costs two requests, not two per item. */
async function saveProgressMany(job, items, who) {
  await ensureProgressSheet();
  const f = await findFile();
  const S = f.base + "/worksheets('" + PROGRESS_SHEET + "')";
  return serialised(PROGRESS_SHEET, async () => {
    const used = await call("GET", S + "/usedRange?$select=values,rowCount");
    const rows = used.values || [], at = {};
    for (let i = 1; i < rows.length; i++)
      if (String(rows[i][0] || "").trim().toUpperCase() === String(job).toUpperCase())
        at[String(rows[i][1] || "").trim()] = i + 1;
    let next = (used.rowCount || 1) + 1;
    const when = nowStamp();
    const reqs = (items || []).map(x => {
      const r = at[x.item] || next++;
      return { method: "PATCH", url: S + "/range(address='A" + r + ":F" + r + "')",
               body: { values: [[job, x.item, String(x.done), String(x.total), who || "", when]], numberFormat: TEXT6 } };
    });
    if (reqs.length) await batchWrite(reqs);
    return reqs.length;
  });
}

/* ---- writes: always addressed by cell, never by rewriting the file ---- */
const A1 = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };

async function setFill(sheet, address, color) {
  const f = await findFile();
  return call("PATCH", f.base + "/worksheets('" + sheet + "')/range(address='" + address + "')/format/fill", { color });
}
async function clearFill(sheet, address) {
  const f = await findFile();
  return call("POST", f.base + "/worksheets('" + sheet + "')/range(address='" + address + "')/format/fill/clear", {});
}
async function setValues(sheet, address, values) {
  const f = await findFile();
  return call("PATCH", f.base + "/worksheets('" + sheet + "')/range(address='" + address + "')", { values });
}
async function readColumn(sheet, address) {
  const f = await findFile();
  const r = await call("GET", f.base + "/worksheets('" + sheet + "')/range(address='" + address + "')");
  return r.values;
}

/** Re-find a job's row immediately before writing: rows move in this sheet, and a
    cached row number would eventually write onto somebody else's job. */
async function rowForJob(sheet, jobId) {
  const vals = await readColumn(sheet, "C1:C400");
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] == null ? "" : vals[i][0]).trim().toUpperCase();
    if (v === jobId.toUpperCase()) return i + 1;
  }
  throw new Error("Job " + jobId + " is no longer on the " + sheet + " sheet - it may have been moved or removed.");
}


/* ---- moving a job between sections of the Production sheet ----------------
   The only structural write in this app. A move is: insert a blank row straight
   after the target section's last job, copy the job's row into it (values,
   number formats, fills, fonts, borders, alignment, height), confirm the copy
   is there, then delete the original. Every step finds rows by job number at
   that moment - nothing is written by a remembered row number. If anything
   fails before the delete, the inserted row is removed again, so the sheet is
   left exactly as it was.                                                   */
const PROD_SHEET = "Production";
const ROW_COLS = 90;                                   // A..CL, the sheet's used width
const lastCol = A1(ROW_COLS);

async function liveBlocks() {
  const f = await findFile();
  const r = await call("GET", f.base + "/worksheets('" + PROD_SHEET + "')/range(address='A1:K600')?$select=values");
  return blocksFromValues(r.values || []);            // parser.js - the same rule as the download
}

/** Run many small Graph requests as JSON batches (20 per batch), a few batches
    in flight at a time - Excel queues per session and refuses a flood with 429
    OperationQueueFull. Returns bodies in input order; throws on any failure. */
async function batchGet(urls, limit) {
  const out = new Array(urls.length), starts = [];
  for (let i = 0; i < urls.length; i += 20) starts.push(i);
  let next = 0;
  const worker = async () => {
    while (next < starts.length) {
      const s = starts[next++], part = urls.slice(s, s + 20);
      const res = await batchRun(part.map((u, n) => ({ id: String(n + 1), method: "GET", url: u })));
      res.forEach((body, n) => { out[s + n] = body; });
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit || 3, starts.length) }, worker));
  return out;
}
/** Sequential batches of writes (they share one workbook session). */
async function batchWrite(reqs) {
  for (let i = 0; i < reqs.length; i += 20) {
    const part = reqs.slice(i, i + 20).map((r, n) => Object.assign({ id: String(n + 1) }, r));
    await batchRun(part);
  }
}
const sleep = ms => new Promise(s => setTimeout(s, ms));
async function batchRun(reqs) {
  let sessionRetried = false;
  for (let attempt = 0; attempt < 6; attempt++) {
    const h = {}; if (sessionId) h["workbook-session-id"] = sessionId;
    const body = { requests: reqs.map(r => Object.assign({}, r,
      { headers: Object.assign({}, h, r.body ? { "Content-Type": "application/json" } : {}) })) };
    const r = await call("POST", "/$batch", body);
    const byId = {}; (r.responses || []).forEach(x => { byId[x.id] = x; });
    const bad = reqs.map(q => byId[q.id]).find(x => !x || x.status >= 400);
    if (!bad) return reqs.map(q => (byId[q.id] || {}).body || {});
    const txt = JSON.stringify((bad && bad.body) || {});
    if (!sessionRetried && /InvalidSession|invalidSessionReCreatable/.test(txt)) {
      sessionRetried = true; sessionId = null; await openSession(); continue;
    }
    if (attempt < 5 && (bad.status === 429 || bad.status >= 500 || /OperationQueueFull|tooManyRequests/.test(txt))) {
      await sleep(1000 + attempt * 1500);          // the queue drains in a second or two
      continue;                                    // re-running finished items is harmless: same value again
    }
    throw new Error("batch item failed: " + (bad ? bad.status : "missing") + " " + txt.slice(0, 200));
  }
}

/** Everything about one row that must travel with it. Values come with their
    number formats in one read. Fills and fonts carry the job's status and the
    office's colour notes, so they are read live - but per group of cells that
    the last download says share one fill / one font, each group in a single
    request. A group that is no longer uniform live is read cell by cell.
    Either way the result is what is in the sheet at this moment.           */
const normFont = fo => ({ bold: !!fo.bold, italic: !!fo.italic, size: fo.size || 11, name: fo.name || "Calibri",
                          color: fo.color || "#000000", underline: fo.underline || "None" });
async function captureRow(row, tmpl) {
  const f = await findFile(), S = f.base + "/worksheets('" + PROD_SHEET + "')";
  const addr = "A" + row + ":" + lastCol + row;
  const R = (a, b) => S + "/range(address='" + A1(a) + row + ":" + A1(b) + row + "')";
  const cells = await call("GET", S + "/range(address='" + addr + "')?$select=values,valueTypes,numberFormat,formulas");
  const T = (tmpl && tmpl.cells && tmpl.cells.length >= ROW_COLS) ? tmpl.cells : null;
  const flruns = runs(ROW_COLS, c => T ? (T[c - 1].flk || "") : "c" + c, c => null);
  const fruns = runs(ROW_COLS, c => T ? (T[c - 1].fk || "") : "c" + c, c => null);
  const urls = [S + "/range(address='" + addr + "')/format?$select=rowHeight"];
  flruns.forEach(r => urls.push(R(r.from, r.to) + "/format/fill"));
  fruns.forEach(r => urls.push(R(r.from, r.to) + "/format/font"));
  const got = await batchGet(urls, 3);
  const height = (got[0] || {}).rowHeight || null;
  const fills = new Array(ROW_COLS), fonts = new Array(ROW_COLS), fbFill = [], fbFont = [];
  flruns.forEach((r, i) => {
    const fi = got[1 + i] || {};
    if (fi.color === null || fi.color === undefined) { for (let c = r.from; c <= r.to; c++) fbFill.push(c); }
    else for (let c = r.from; c <= r.to; c++) fills[c - 1] = fi.color;     // "" = no fill
  });
  fruns.forEach((r, i) => {
    const fo = got[1 + flruns.length + i] || {};
    const uniform = ["bold", "italic", "size", "name", "color"].every(k => fo[k] !== null && fo[k] !== undefined);
    if (uniform) { for (let c = r.from; c <= r.to; c++) fonts[c - 1] = normFont(fo); }
    else for (let c = r.from; c <= r.to; c++) fbFont.push(c);
  });
  if (fbFill.length || fbFont.length) {
    const u2 = fbFill.map(c => R(c, c) + "/format/fill").concat(fbFont.map(c => R(c, c) + "/format/font"));
    const g2 = await batchGet(u2, 3);
    fbFill.forEach((c, i) => { fills[c - 1] = (g2[i] || {}).color || ""; });
    fbFont.forEach((c, i) => { fonts[c - 1] = normFont(g2[fbFill.length + i] || {}); });
  }
  return { values: cells.values[0], types: cells.valueTypes[0], numberFormat: cells.numberFormat[0],
           formulas: cells.formulas[0], fills, fonts, height,
           reads: urls.length + fbFill.length + fbFont.length, fallback: fbFill.length + fbFont.length };
}

/* Excel parses what it is given the way it parses typing: "07/04" becomes a
   date, "9434" a number, "TRUE" a boolean. A text cell that looks like one of
   those is written with a leading apostrophe - Excel's own way of saying
   "this is text" - so it comes back exactly as it was. */
function guardText(v) {
  const s = String(v);
  if (/^[=+\-@]/.test(s) || /^\d/.test(s) || /^(true|false)$/i.test(s)) return "'" + s;
  return s;
}
function cellOut(cap, c) {
  const f = cap.formulas[c], v = cap.values[c], t = cap.types[c];
  if (typeof f === "string" && f.charAt(0) === "=") return f;
  if (t === "Empty" || v === "" || v == null) return "";
  if (t === "String") return guardText(v);
  return v;                                            // Double, Boolean, Error - as they are
}
/** Consecutive cells with the same key -> [{from, to, key, val}] (1-based columns). */
function runs(n, keyOf, valOf) {
  const out = [];
  for (let c = 1; c <= n; c++) {
    const k = keyOf(c);
    if (out.length && out[out.length - 1].key === k) out[out.length - 1].to = c;
    else out.push({ from: c, to: c, key: k, val: valOf(c) });
  }
  return out;
}
const BORDER_STYLE = { thin: ["Continuous", "Thin"], medium: ["Continuous", "Medium"], thick: ["Continuous", "Thick"],
  hair: ["Continuous", "Hairline"], dashed: ["Dash", "Thin"], dotted: ["Dot", "Thin"], double: ["Double", "Thick"],
  mediumDashed: ["Dash", "Medium"], dashDot: ["DashDot", "Thin"], mediumDashDot: ["DashDot", "Medium"],
  dashDotDot: ["DashDotDot", "Thin"], mediumDashDotDot: ["DashDotDot", "Medium"], slantDashDot: ["SlantDashDot", "Medium"] };

/** Re-assert the bottom edge of the job row just above a row we inserted or
    deleted. Writing an edge on a neighbour makes Excel move the shared edge
    onto the written row, and a later deletion of the neighbour takes it away;
    giving the edge back to the row that stays keeps the line drawn. Only ever
    writes edges the row's own template already has - never "no border".   */
async function restoreBottomEdge(row, tmpl) {
  if (!row || row < 1 || !tmpl || !tmpl.cells || tmpl.cells.length < ROW_COLS) return 0;
  const f = await findFile(), S = f.base + "/worksheets('" + PROD_SHEET + "')";
  const R = (a, b) => S + "/range(address='" + A1(a) + row + ":" + A1(b) + row + "')";
  const T = tmpl.cells, reqs = [];
  runs(ROW_COLS, c => T[c - 1].bottom ? T[c - 1].bottom.style + "/" + T[c - 1].bottom.color : "", c => T[c - 1].bottom).forEach(r => {
    if (!r.val) return;
    const st = BORDER_STYLE[r.val.style] || ["Continuous", "Thin"];
    reqs.push({ method: "PATCH", url: R(r.from, r.to) + "/format/borders/EdgeBottom", body: { style: st[0], weight: st[1], color: r.val.color || "#000000" } });
  });
  if (reqs.length) await batchWrite(reqs);
  return reqs.length;
}

async function writeRow(row, cap, tmpl) {
  const f = await findFile(), S = f.base + "/worksheets('" + PROD_SHEET + "')";
  const R = (from, to) => S + "/range(address='" + A1(from) + row + ":" + A1(to) + row + "')";
  const out = []; for (let c = 0; c < ROW_COLS; c++) out.push(cellOut(cap, c));
  await call("PATCH", R(1, ROW_COLS), { formulas: [out], numberFormat: [cap.numberFormat] });
  const reqs = [];
  runs(ROW_COLS, c => cap.fills[c - 1], c => cap.fills[c - 1]).forEach(r =>
    reqs.push(r.val ? { method: "PATCH", url: R(r.from, r.to) + "/format/fill", body: { color: r.val } }
                    : { method: "POST", url: R(r.from, r.to) + "/format/fill/clear", body: {} }));
  runs(ROW_COLS, c => JSON.stringify(cap.fonts[c - 1]), c => cap.fonts[c - 1]).forEach(r =>
    reqs.push({ method: "PATCH", url: R(r.from, r.to) + "/format/font", body: r.val }));
  /* the API reports heights rounded to whole pixels (22.15 -> 21.75); the
     downloaded file has the exact value, so prefer it when we have it */
  const height = (tmpl && tmpl.height) || cap.height;
  if (height) reqs.push({ method: "PATCH", url: R(1, ROW_COLS) + "/format", body: { rowHeight: height } });
  if (tmpl && tmpl.cells && tmpl.cells.length >= ROW_COLS) {
    const T = tmpl.cells;
    runs(ROW_COLS, c => (T[c - 1].h || "") + "|" + (T[c - 1].v || "") + "|" + T[c - 1].wrap, c => T[c - 1]).forEach(r => {
      const b = { wrapText: !!r.val.wrap };
      if (r.val.h) b.horizontalAlignment = r.val.h.charAt(0).toUpperCase() + r.val.h.slice(1);
      if (r.val.v) b.verticalAlignment = r.val.v.charAt(0).toUpperCase() + r.val.v.slice(1);
      reqs.push({ method: "PATCH", url: R(r.from, r.to) + "/format", body: b });
    });
    /* Borders. Excel keeps ONE definition per shared edge: writing any edge
       on a row makes it re-normalise that row and strip the shared edges off
       its neighbours. So: the landing row gets its bottom edge and verticals
       from the template (an inserted row does not reliably inherit them),
       never its top - and the row above it has its own bottom re-asserted
       afterwards (restoreBottomEdge), so the line between them belongs to
       the row that stays put if the moved row is later moved on again.     */
    const sideKey = s => s ? s.style + "/" + s.color : "";
    runs(ROW_COLS, c => [sideKey(T[c - 1].bottom), sideKey(T[c - 1].left), sideKey(T[c - 1].right)].join("|"), c => T[c - 1]).forEach(r => {
      const edge = (name, s) => {
        const st = s ? (BORDER_STYLE[s.style] || ["Continuous", "Thin"]) : null;
        reqs.push({ method: "PATCH", url: R(r.from, r.to) + "/format/borders/" + name,
                    body: st ? { style: st[0], weight: st[1], color: s.color || "#000000" } : { style: "None" } });
      };
      edge("EdgeBottom", r.val.bottom);
      edge("EdgeLeft", r.val.left);
      edge("EdgeRight", r.val.right);
      if (r.to > r.from) edge("InsideVertical", r.val.right || r.val.left);
    });
  }
  await batchWrite(reqs);
  return reqs.length;
}

/** Where a job is right now on the Production sheet, and every section. */
async function locateJob(jobId) {
  const B = await liveBlocks();
  let hit = null;
  B.blocks.forEach(b => b.jobs.forEach(j => { if (j.id === jobId) hit = { row: j.row, block: b }; }));
  return { blocks: B, hit };
}

/** Move one job into a section of the Production sheet (by section index, see
    blocksFromValues). tmplFor(jobId) returns templateForJob() from the last
    download (or null) - used for the moved row and for the row it lands under.
    onStep(text) reports progress. Resolves {moved, from, to, fromRow, row}.  */
async function moveJobRow(jobId, targetIdx, tmplFor, onStep) {
  const tmpl = typeof tmplFor === "function" ? tmplFor(jobId) : (tmplFor || null);
  jobId = String(jobId).trim().toUpperCase();
  const step = t => { if (onStep) try { onStep(t); } catch (e) {} };
  const f = await findFile(), S = f.base + "/worksheets('" + PROD_SHEET + "')";
  const rowOf = n => S + "/range(address='" + n + ":" + n + "')";
  const L = await locateJob(jobId);
  const tb = L.blocks.blocks[targetIdx];
  if (!tb) throw new Error("There is no section " + targetIdx + " on the Production sheet.");
  if (!L.hit) throw new Error("Job " + jobId + " is not on the Production sheet.");
  const src = L.hit.row, sb = L.hit.block;
  if (sb.idx === targetIdx) return { moved: false, from: sb.name, to: tb.name, fromRow: src, row: src };
  if (!tb.last) throw new Error("The '" + tb.name + "' section has no divider or jobs in Excel right now, so there is nowhere safe to put " + jobId + ". Place one job there in Excel first.");
  const tgt = tb.last + 1;                             // straight after the section's last job
  step("reading " + jobId + " (row " + src + ")");
  const cap = await captureRow(src, tmpl);
  step("making room in " + tb.name);
  await call("POST", rowOf(tgt) + "/insert", { shift: "Down" });
  const srcNow = src + (src >= tgt ? 1 : 0);
  try {
    step("writing the copy at row " + tgt);
    await writeRow(tgt, cap, tmpl);
    const aboveId = tb.jobs.length ? tb.jobs[tb.jobs.length - 1].id : null;
    if (aboveId && typeof tmplFor === "function") await restoreBottomEdge(tgt - 1, tmplFor(aboveId));
    const col = await readColumn(PROD_SHEET, "C1:C600");
    const at = [];
    col.forEach((v, i) => { if (String(v[0] == null ? "" : v[0]).trim().toUpperCase() === jobId) at.push(i + 1); });
    if (at.length !== 2 || at.indexOf(tgt) < 0 || at.indexOf(srcNow) < 0)
      throw new Error("The sheet changed under " + jobId + " while it was being moved (now at rows " + at.join(", ") + ").");
  } catch (e) {
    /* Put the sheet back. Rows may have shifted meanwhile, so never trust a
       remembered number: re-read the sections and identify our copy as the
       occurrence of the job inside the target section (the original is in
       another section). If nothing landed, the inserted row is still blank
       and sits where the target row went - shifted exactly like the original. */
    try {
      const B2 = await liveBlocks(), occ = [];
      B2.blocks.forEach(b => b.jobs.forEach(j => { if (j.id === jobId) occ.push({ row: j.row, idx: b.idx }); }));
      const ours = occ.filter(o => o.idx === targetIdx), orig = occ.filter(o => o.idx !== targetIdx);
      let victim = null;
      if (ours.length === 1 && orig.length === 1) victim = ours[0].row;
      else if (occ.length === 1) {
        const cand = tgt + (occ[0].row - srcNow);
        const probe = await call("GET", S + "/range(address='A" + cand + ":K" + cand + "')?$select=values");
        if (((probe.values && probe.values[0]) || []).every(v => v === "" || v == null)) victim = cand;
      }
      if (victim == null) throw new Error("could not tell which row is the copy");
      await call("POST", rowOf(victim) + "/delete", { shift: "Up" });
    } catch (e2) {
      throw new Error(e.message + " Could not tidy up (" + e2.message + ") - check the Production sheet for a blank or duplicate row near row " + tgt + ".");
    }
    throw e;
  }
  step("removing the old row " + srcNow);
  await call("POST", rowOf(srcNow) + "/delete", { shift: "Up" });
  /* the job row that used to sit above the old row gets its bottom edge back too */
  const prev = sb.jobs.filter(j => j.row < src).pop();
  if (prev && prev.row === src - 1 && typeof tmplFor === "function") await restoreBottomEdge(srcNow - 1, tmplFor(prev.id));
  return { moved: true, from: sb.name, to: tb.name, fromRow: src, row: tgt - (srcNow < tgt ? 1 : 0) };
}

window.CW = {
  initAuth, signIn, signOut, token, findFile, openSession, lastModified,
  downloadWorkbook, setFill, clearFill, setValues, rowForJob, A1,
  ensureLogSheet, appendLog, LOG_SHEET,
  ensureViewsSheet, saveAssignment, clearAssignment, VIEWS_SHEET,
  ensureProgressSheet, saveProgress, saveProgressMany, PROGRESS_SHEET, batchWrite,
  listVersions, downloadVersion, restoreVersion,
  liveBlocks, locateJob, moveJobRow, captureRow, batchGet,
  _setToken(fn) { tokenOverride = fn; }, _setFile(ref) { fileRef = ref; }, _setSession(id) { sessionId = id; },
  get account() { return account; }
};
