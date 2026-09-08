/* Sign-in + Microsoft Graph access for the Costello production dashboard.
   Reads the live workbook by downloading it (full fidelity, including fills);
   writes surgically through the Excel API so the file is never rewritten. */

const CLIENT_ID = "a989939b-17f3-4c9c-adb7-4d8338f4878a";
const TENANT_ID = "cb4cfc4c-96f4-44c0-b37b-a467826f86d6";
/* Sign-in and every workbook call use SCOPES. The SharePoint list that holds
   the hand-set phases needs Sites.ReadWrite.All as well - asked for ONLY on
   list requests, so a tenant that has not consented to it yet still signs in
   and works; the phase list simply reports that it needs permission. */
const SCOPES = ["Files.ReadWrite.All", "User.Read"];
const LIST_SCOPES = ["Files.ReadWrite.All", "Sites.ReadWrite.All", "User.Read"];
const SITE_PATH = "costellowindowsie.sharepoint.com:/sites/ProductionProgress";
/* The separate site the floor stations live in, on the same hostname. It is
   named here rather than taken from station-core.js because graph.js is loaded
   on its own by the phases tests, and because this is the URL segment of the
   site - station-core.js' STATION_SITE is the same word for the same site and
   the two must match. */
const STATION_SITE_NAME = "FloorStations";
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

/* The master dashboard signs in with SCOPES, so a tenant that has not
   consented to Sites.ReadWrite.All still gets a working dashboard. The station
   page has nothing BUT lists, so it passes LIST_SCOPES and asks for the list
   permission at the door - there is no consent flow on a shared tablet. */
async function signIn(scopes) {
  const res = await app().loginPopup({ scopes: scopes || SCOPES, prompt: "select_account" });
  account = res.account;
  return account;
}

function signOut() {
  return app().logoutPopup({ account: account });
}

let tokenOverride = null;     // rehearsal harness only: a token without the sign-in library
/* quiet = never open a consent popup (background reads); a popup is only ever
   opened for the sign-in scopes, or for the list scopes from a user's click. */
async function token(scopes, quiet) {
  if (tokenOverride) return tokenOverride(scopes || SCOPES, !!quiet);
  if (!account) throw new Error("not signed in");
  scopes = scopes || SCOPES;
  try {
    const r = await app().acquireTokenSilent({ scopes: scopes, account: account });
    return r.accessToken;
  } catch (e) {
    if (quiet) throw new Error("permission needed: " + ((e && e.errorCode) || (e && e.message) || "consent"));
    const r = await app().acquireTokenPopup({ scopes: scopes, account: account });
    return r.accessToken;
  }
}
/** Ask (once, with a popup if needed) for the list permission - from a click only. */
async function listConsent() { return token(LIST_SCOPES, false); }
/** Have we got the list permission already? A quiet attempt, so it can be
    asked in the background - before a feed run, say - without a popup ever
    appearing in front of someone who did not click anything. */
async function hasListConsent() {
  try { await token(LIST_SCOPES, true); return true; } catch (e) { return false; }
}

/* Which requests need Sites.ReadWrite.All. This decides on the SHAPE of the
   path and never on a stray character: every A1 address in the app has a colon
   in it (range(address='A1:K600')), so "has a colon" would quietly send every
   fill, row move, log line and progress row asking for the list permission -
   which fails outright on a tenant that has not granted it, and loses the
   interactive re-auth popup on one that has.

     · /workbook or /drive/ - the file itself: SCOPES, and not quiet, so an
       expired token still opens the popup that gets the work saved;
     · /lists - a SharePoint list: LIST_SCOPES, quietly;
     · /sites/{host}:/{path} - looking a site up by its path: LIST_SCOPES,
       quietly, EXCEPT the workbook's own site, which findFile() resolves at
       sign-in and which must keep working for a tenant that has never
       consented to the list permission.

   The one awkward case is the interim arrangement in which the floor's lists
   live in the workbook's own site: stationSite() then looks that site up too,
   and that lookup is a list call - quiet, and never allowed to throw a consent
   window at a tablet on the floor. It is the same path findFile() uses, so it
   is marked with a $select that findFile() never sends, and the rule below
   reads the mark. Nothing else in the app sends it.                        */
const SITE_AS_LIST = "?$select=id,displayName";
function needsListScope(path) {
  if (path == null) return false;
  /* a list path is a list call whatever else is in it. This is tested first so
     that a site whose id or name happens to carry "/drive" in it - and any
     future list path that does - cannot be read as a workbook call and sent
     with the wrong scopes, non-quietly, at a tablet on the floor. */
  if (path.indexOf("/lists") >= 0) return true;
  if (path.indexOf("/workbook") >= 0 || path.indexOf("/drive/") >= 0) return false;
  if (!/^\/sites\/[^/]+:\//.test(path)) return false;
  return path.indexOf(SITE_PATH) < 0 || path.indexOf(SITE_AS_LIST) >= 0;
}
/** The scopes one request is made with - the single place that decides, so a
    test can ask the same question the request asks. */
function scopeFor(path) { return needsListScope(path) ? LIST_SCOPES : SCOPES; }

/* The workbook session header belongs to workbook requests only: sending a
   stale one at /sites/{id}/lists would earn an InvalidSession 400 on a call
   that has nothing to do with the workbook. */
async function headers(extra, path) {
  const isList = needsListScope(path);
  const h = { Authorization: "Bearer " + (await token(scopeFor(path), isList)) };
  if (sessionId && (path == null || path.indexOf("/workbook") >= 0)) h["workbook-session-id"] = sessionId;
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
    const init = { method, headers: await headers(body ? { "Content-Type": "application/json" } : null, path) };
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
  /* list paths have no "/workbook" in them, so fall back to the whole path */
  throw new Error(method + " " + (path.split("/workbook")[1] || path) + " -> " + lastStatus + " " + lastText.slice(0, 200));
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

/* ---- job alerts ----------------------------------------------------------
   Who gets emailed about a job. One row per (Job, Email), the address stored
   lower-cased, removal = blanking the row like clearAssignment. Everything
   here addresses ALERTS_SHEET by constant, like the log, the views and the
   progress sheet, so no call site can point a write at Production. Nothing is
   read from or written to the Dashboard Config sheet here: the dashboard only
   ever reads that one, out of the downloaded workbook.                      */
const ALERTS_SHEET = "Dashboard Alerts";
const ALERTS_HEADERS = [["Job", "Email", "Added by", "When"]];
const TEXT4 = [["@", "@", "@", "@"]];
let alertsReady = null;

function ensureAlertsSheet() {
  if (alertsReady) return alertsReady;
  alertsReady = makeAlertsSheet();
  alertsReady.catch(() => { alertsReady = null; });
  return alertsReady;
}
async function makeAlertsSheet() {
  const f = await findFile();
  const ws = await call("GET", f.base + "/worksheets");
  if (!(ws.value || []).some(w => w.name === ALERTS_SHEET)) {
    await call("POST", f.base + "/worksheets/add", { name: ALERTS_SHEET });
    const S = f.base + "/worksheets('" + ALERTS_SHEET + "')";
    await call("PATCH", S + "/range(address='A1:D1')", { values: ALERTS_HEADERS });
    await call("PATCH", S + "/range(address='A1:D1')/format/font", { bold: true, color: "#FFFFFF" });
    await call("PATCH", S + "/range(address='A1:D1')/format/fill", { color: "#17171A" });
    /* text throughout, like the Progress sheet: an address is not a formula and
       "2026-09-04 10:42" is not a date Excel should re-read in its own locale */
    const fmt = []; for (let i = 0; i < 1999; i++) fmt.push(["@", "@", "@", "@"]);
    await call("PATCH", S + "/range(address='A2:D2000')", { numberFormat: fmt });
    const widths = { A: 70, B: 260, C: 220, D: 130 };
    for (const c in widths)
      await call("PATCH", S + "/range(address='" + c + ":" + c + "')/format", { columnWidth: widths[c] });
  }
  return true;
}

/** Subscribe one address to one job. Upsert: an existing (Job, Email) line is
    written again rather than added a second time, and a line an earlier
    removal blanked is used before the sheet is made any longer. */
async function addAlert(job, email, who) {
  await ensureAlertsSheet();
  const f = await findFile();
  const S = f.base + "/worksheets('" + ALERTS_SHEET + "')";
  const j = String(job).trim().toUpperCase(), e = String(email).trim().toLowerCase();
  return serialised(ALERTS_SHEET, async () => {
    const used = await call("GET", S + "/usedRange?$select=values,rowCount");
    const rows = used.values || [];
    let target = 0, spare = 0;
    for (let i = 1; i < rows.length; i++) {
      const rj = String(rows[i][0] == null ? "" : rows[i][0]).trim().toUpperCase();
      const re = String(rows[i][1] == null ? "" : rows[i][1]).trim().toLowerCase();
      if (rj === j && re === e) { target = i + 1; break; }
      if (!spare && !rj && !re) spare = i + 1;
    }
    if (!target) target = spare || (used.rowCount || 1) + 1;
    await call("PATCH", S + "/range(address='A" + target + ":D" + target + "')", {
      values: [[j, e, who || "", nowStamp()]], numberFormat: TEXT4
    });
    return target;
  });
}

/** Unsubscribe: blank the line (A:D), the way a view assignment is cleared.
    Blanks every matching line, in case one was typed twice by hand in Excel. */
async function removeAlert(job, email) {
  await ensureAlertsSheet();
  const f = await findFile();
  const S = f.base + "/worksheets('" + ALERTS_SHEET + "')";
  const j = String(job).trim().toUpperCase(), e = String(email).trim().toLowerCase();
  return serialised(ALERTS_SHEET, async () => {
    const used = await call("GET", S + "/usedRange?$select=values");
    const rows = used.values || [];
    let n = 0;
    for (let i = 1; i < rows.length; i++) {
      const rj = String(rows[i][0] == null ? "" : rows[i][0]).trim().toUpperCase();
      const re = String(rows[i][1] == null ? "" : rows[i][1]).trim().toLowerCase();
      if (rj !== j || re !== e) continue;
      await call("PATCH", S + "/range(address='A" + (i + 1) + ":D" + (i + 1) + "')", { values: [["", "", "", ""]] });
      n++;
    }
    return n;
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

/* ---- SharePoint lists ------------------------------------------------------
   The hand-set phases are shared through a SharePoint list in the same site,
   NOT through the workbook: nothing in this block addresses a worksheet, a
   range, a cell format or the file at all. Every path is
   /sites/{siteId}/lists/... and the only thing taken from the workbook is the
   site id that findFile() already worked out.

   The list itself is made by hand in SharePoint. If it is not there, listId()
   answers null and listItems() answers null with it, so the dashboard can say
   so plainly instead of trying to create anything.                          */
const LISTIDS_KEY = "cw_listids";
const listIdMemo = {};                 // cache key -> id, for this page load

/* Every function below takes an optional opts = { siteId, fields }. Without
   it they behave exactly as they always have: the workbook's own site, and the
   phases list's five columns. With it they address another site - the floor
   stations site, which the station account can see and the workbook's site it
   cannot - and ask for another set of columns. That is the whole of the
   generalisation: no call site that omits opts changes behaviour. */
const PHASE_SELECT = "Title,Phase,PhaseName,SetBy,SetAt";
/** The site a list call is about: the workbook's, unless one was named. A
    named site is never looked up through findFile(), because the account
    reading it may have no access to the workbook at all. */
async function listSiteId(opts) {
  if (opts && opts.siteId) return opts.siteId;
  /* A caller that named a site and passed nothing has not resolved one yet.
     Falling through to findFile() here is how the station page ends up asking
     for /drive/root/children - the one request it must never make - so an
     explicit empty siteId is an error rather than a default. */
  if (opts && "siteId" in opts) throw new Error("no site resolved for this list yet");
  const f = await findFile();
  return f.siteId;
}
/** Two sites can hold two lists with the same display name, so the cache is
    keyed by both - and the default site keeps the bare display name it has
    always used, so cw_listids does not change shape for anyone. */
const listKey = (displayName, siteId) => (siteId ? siteId + "|" : "") + displayName;

function listIdCache() {
  try { return JSON.parse(localStorage.getItem(LISTIDS_KEY) || "{}"); } catch (e) { return {}; }
}
function rememberListId(name, id) {
  const all = listIdCache();
  if (all[name] === id) return;
  all[name] = id;
  try { localStorage.setItem(LISTIDS_KEY, JSON.stringify(all)); } catch (e) {}
}

/** The id of a list, found once by its display name and then remembered - in
    memory for this page and in localStorage for the next one. A list that is
    not there is never cached: the moment the manager creates it, the next call
    finds it without anyone clearing anything. */
async function listId(displayName, opts) {
  const siteId = await listSiteId(opts);
  const key = listKey(displayName, opts && opts.siteId ? siteId : null);
  if (listIdMemo[key]) return listIdMemo[key];
  const cached = listIdCache()[key];
  if (cached) { listIdMemo[key] = cached; return cached; }
  const r = await call("GET", "/sites/" + siteId + "/lists?$select=id,displayName");
  const want = String(displayName).toLowerCase();
  const hit = (r.value || []).find(l => String(l.displayName || "").toLowerCase() === want);
  if (!hit) return null;
  listIdMemo[key] = hit.id;
  rememberListId(key, hit.id);
  return hit.id;
}

/** The next page of a collection arrives as an absolute URL. It is normally
    the same base call() puts on the front, but an absolute origin must be cut
    off either way rather than handed to call() whole - and so must the version
    segment call() is about to add back. */
function graphPath(url) {
  let s = String(url || "");
  if (s.indexOf(G) === 0) return s.slice(G.length);
  const m = /^https?:\/\/[^/]+(\/.*)$/.exec(s);
  if (m) s = m[1];
  return s.replace(/^\/(v1\.0|beta)(?=\/)/, "");
}

const LIST_PAGE_CAP = 50;                   // 50 pages of 999: a runaway guard
/** Every item of a list, following @odata.nextLink to the end.
    null (not []) means the list does not exist. */
async function listItems(displayName, opts) {
  const id = await listId(displayName, opts);
  if (!id) return null;
  const siteId = await listSiteId(opts);
  const select = opts && opts.fields && opts.fields.length ? opts.fields.join(",") : PHASE_SELECT;
  let path = "/sites/" + siteId + "/lists/" + id +
             "/items?expand=fields(select=" + select + ")&$top=999";
  const out = [];
  let page = 0;
  for (; page < LIST_PAGE_CAP && path; page++) {
    const r = await call("GET", path);
    (r.value || []).forEach(it => out.push({ id: String(it.id), fields: it.fields || {} }));
    path = r["@odata.nextLink"] ? graphPath(r["@odata.nextLink"]) : null;
  }
  /* the cap is a guard against a runaway loop, not a page size - if it is ever
     reached, the caller is holding an incomplete list and should know */
  if (path) console.warn("[graph] “" + displayName + "” has more than " +
    (LIST_PAGE_CAP * 999) + " items: only the first " + out.length + " were read.");
  return out;
}

/* ---- what moved since last time -------------------------------------------
   The floor taps a counter and the office is meant to see it inside ten
   seconds. Re-reading the whole list six times a minute on two screens would
   be six hundred rows a minute of the same unchanged text, so the polls use
   Graph's delta feed instead: the first call enumerates the list and hands
   back a deltaLink, and every call after it passes that token and gets only
   the items that have changed since.

   Three things about the feed the callers have to know, all of them handled
   here rather than in the pages:

     · a deleted item arrives carrying "@removed" instead of a fields bag;
     · the same item can appear more than once in one feed, and the LAST
       occurrence is the true one (ST.mergeDelta applies that rule);
     · a token that is too old, or a list whose server state has moved on,
       is answered 410 Gone with resyncChangesApplyDifferences or
       resyncChangesUploadDifferences - which means "start again", not "this
       failed".

   So every 4xx from the delta endpoint - 410 included - is re-thrown as one
   recognisable error, and the caller answers it the one way that is always
   right: read the whole list with listItems() once, then start a fresh delta
   enumeration with no token. Anything else (offline, a 5xx that outlived
   call()'s retries) is thrown untouched, so a passing failure stays a passing
   failure and does not cost a full read.                                    */
const DELTA_FAIL = "delta must be restarted";
const DELTA_TOP = 500;
/** Is this the error that means "throw the token away and read the lot"? */
function isDeltaRestart(e) {
  const m = (e && e.message) || String(e || "");
  return m.indexOf(DELTA_FAIL) >= 0;
}
/** ... and was it the 410 Gone that says "your token is too old, enumerate
    again"? That one means delta itself is perfectly well - anything else from
    the delta endpoint means this list will not serve one, and a caller that
    cannot tell them apart asks a refusing list six times a minute for ever. */
function isDeltaResync(e) {
  if (!isDeltaRestart(e)) return false;
  const m = (e && e.message) || String(e || "");
  return /->\s*410\b/.test(m) || m.indexOf("resyncChanges") >= 0;
}
/** One delta pass. Without opts.token it enumerates the list from scratch;
    with it, only what has changed since that token was issued.
    Answers { items:[{id, fields, removed}], next } - `next` is the deltaLink
    to hand back next time. null (not {}) means the list does not exist. */
async function listDelta(displayName, opts) {
  let path;
  if (opts && opts.token) {
    path = graphPath(opts.token);
    /* Graph normally carries the $expand through the deltaLink, but it is not
       obliged to, and a token path without it answers items with no fields at
       all - which reads on the other end as "every row went blank". Put it
       back if it is not there. */
    if (path.indexOf("expand=fields") < 0) {
      const sel = opts.fields && opts.fields.length ? opts.fields.join(",") : PHASE_SELECT;
      path += (path.indexOf("?") >= 0 ? "&" : "?") + "expand=fields(select=" + sel + ")";
    }
  } else {
    const id = await listId(displayName, opts);
    if (!id) return null;
    const siteId = await listSiteId(opts);
    const select = opts && opts.fields && opts.fields.length ? opts.fields.join(",") : PHASE_SELECT;
    path = "/sites/" + siteId + "/lists/" + id +
           "/items/delta?expand=fields(select=" + select + ")&$top=" + DELTA_TOP;
  }
  const items = [];
  let next = null;
  for (let page = 0; page < 50 && path; page++) {
    let r;
    try {
      r = await call("GET", path);
    } catch (e) {
      const m = (e && e.message) || String(e || "");
      /* 4xx, and the 410 Gone that carries a resync code, both mean the same
         thing to a caller: this token is no use, read the list instead */
      if (/->\s*4\d\d\b/.test(m) || m.indexOf("resyncChanges") >= 0)
        throw new Error(DELTA_FAIL + ": " + m);
      throw e;
    }
    (r.value || []).forEach(it => {
      items.push({ id: String(it.id), fields: it.fields || {}, removed: !!it["@removed"] });
    });
    if (r["@odata.nextLink"]) { path = graphPath(r["@odata.nextLink"]); continue; }
    next = r["@odata.deltaLink"] || null;
    path = null;
  }
  return { items: items, next: next };
}

/* ---- one item per Title, even with two browsers writing at once -----------
   serialised() keeps one tab in order; it cannot see the tab on the next desk.
   Two people setting the same job at the same moment both read "nothing there"
   and both POST. The list is created with "enforce unique values" on Title, so
   the second POST is refused - and a refusal there is not an error, it means
   the item exists now. Either way the answer is the same: read the items for
   that Title back, keep the OLDEST id, delete any others, and PATCH the
   survivor with what this caller meant to write. That is done after a POST
   succeeds as well, so a list without the unique rule still ends up with one
   item per job rather than two.                                             */
const itemIdOrder = (a, b) => {
  const na = Number(a.id), nb = Number(b.id);
  if (isFinite(na) && isFinite(nb)) return na - nb;
  return String(a.id) < String(b.id) ? -1 : String(a.id) > String(b.id) ? 1 : 0;
};
/** Every item of a list carrying one Title, oldest first. */
async function listItemsFor(displayName, title, opts) {
  const all = (await listItems(displayName, opts)) || [];
  const key = String(title).trim().toUpperCase();
  return all.filter(x => String(x.fields.Title == null ? "" : x.fields.Title).trim().toUpperCase() === key)
            .sort(itemIdOrder);
}

async function listUpsert(displayName, title, fields, opts) {
  await listConsent();                    // a click: the one place the list permission may be asked for

  const id = await listId(displayName, opts);
  if (!id) throw new Error("The \u201c" + displayName + "\u201d list is not in SharePoint.");
  const siteId = await listSiteId(opts);
  const base = "/sites/" + siteId + "/lists/" + id;
  const key = String(title).trim().toUpperCase();
  const body = Object.assign({ Title: key }, fields || {});
  /* keep the oldest, remove the rest, and write the intended fields onto it */
  async function settle(mine) {
    const keep = mine[0];
    for (let i = 1; i < mine.length; i++) await call("DELETE", base + "/items/" + mine[i].id);
    await call("PATCH", base + "/items/" + keep.id + "/fields", body);
    return { id: keep.id, created: false, deduped: mine.length - 1 };
  }
  return serialised("list:" + displayName, async () => {
    const mine = await listItemsFor(displayName, key, opts);
    if (mine.length) return await settle(mine);
    let made;
    try {
      made = await call("POST", base + "/items", { fields: body });
    } catch (e) {
      /* refused: either the unique-values rule caught a second browser, or
         something else went wrong. Only the first of those leaves an item
         behind, so look before deciding it was a real failure. */
      const now = await listItemsFor(displayName, key, opts);
      if (!now.length) throw e;
      return await settle(now);
    }
    const after = await listItemsFor(displayName, key, opts);
    if (after.length > 1) return await settle(after);      // both browsers got through
    return { id: made && made.id != null ? String(made.id) : (after[0] && after[0].id) || null,
             created: true, deduped: 0 };
  });
}

/** Remove the items for one Title - every one of them, so a duplicate left by
    two browsers writing at once cannot survive a clear.
    false = there was nothing to remove. */
async function listDelete(displayName, title) {
  await listConsent();                    // a click: the one place the list permission may be asked for

  const id = await listId(displayName);
  if (!id) return false;
  const f = await findFile();
  const base = "/sites/" + f.siteId + "/lists/" + id;
  const key = String(title).trim().toUpperCase();
  return serialised("list:" + displayName, async () => {
    const mine = await listItemsFor(displayName, key);
    if (!mine.length) return false;
    for (let i = 0; i < mine.length; i++) await call("DELETE", base + "/items/" + mine[i].id);
    return true;
  });
}

/* ---- plain writes, for the station feeder and the floor's counters ---------
   listUpsert above reads before it writes, because two people can set the same
   phase at the same second. These two do not: the feeder is the only writer of
   a job's facts and the floor is the only writer of its counters, and both
   already know the item id they mean, so a POST or a PATCH is the whole of it.
   Neither ever opens a consent popup - headers() asks for the list scopes
   quietly, so a background feed on a tenant that has not granted them fails
   with "permission needed" instead of throwing a window at somebody.       */
async function listAdd(displayName, fields, opts) {
  const id = await listId(displayName, opts);
  if (!id) throw new Error("The “" + displayName + "” list is not in SharePoint.");
  const siteId = await listSiteId(opts);
  const made = await call("POST", "/sites/" + siteId + "/lists/" + id + "/items", { fields: fields || {} });
  return { id: made && made.id != null ? String(made.id) : null };
}
/** One item of a list, by its id, with the columns opts asks for. null = there
    is no such item (or no such list) any more.

    This exists for decisions that must be made on what the list says NOW
    rather than on what a read a moment ago said: the station feeder plans its
    writes from one read and then sends up to sixty of them, and a tap from the
    floor can land in between. One item is one small GET. */
async function listItem(displayName, itemId, opts) {
  const id = await listId(displayName, opts);
  if (!id) return null;
  const siteId = await listSiteId(opts);
  const select = opts && opts.fields && opts.fields.length ? opts.fields.join(",") : PHASE_SELECT;
  try {
    const r = await call("GET", "/sites/" + siteId + "/lists/" + id + "/items/" + itemId +
                                "?expand=fields(select=" + select + ")");
    return { id: String(r.id), fields: r.fields || {} };
  } catch (e) {
    if (isMissing(e)) return null;      // deleted between the plan and the write
    throw e;
  }
}
async function listPatch(displayName, itemId, fields, opts) {
  const id = await listId(displayName, opts);
  if (!id) throw new Error("The “" + displayName + "” list is not in SharePoint.");
  const siteId = await listSiteId(opts);
  return call("PATCH", "/sites/" + siteId + "/lists/" + id + "/items/" + itemId + "/fields", fields || {});
}

/* ---- where the floor's lists live -----------------------------------------
   The intended home is a separate SharePoint site on the same hostname as the
   workbook's, holding one list per station: the station account is a member of
   THAT site and of nothing else, which is what keeps the workbook out of its
   reach.

   Creating a site needs a Global Administrator, which the owner has not got
   yet, so there is an interim arrangement: the three lists live in the
   workbook's own site and the station account is a member of it. That trades
   the isolation away - in this arrangement the station account CAN open the
   workbook - and the owner accepted that for now. See docs/STATIONS.md.

   So this resolves the Floor stations site if it is there and falls back to
   the workbook's own site if it is not, and it keeps looking: the day the site
   appears and the rows are copied across, both pages move to it on their own
   with no code change and nothing for anybody to clear. The fallback is
   resolved by a plain site lookup, never through findFile() - nothing in this
   path may touch /drive/ or /workbook, because the whole point of the station
   account is that it has no business there.                                 */
const STATION_SITE_KEY = "cw_stationsite";
const STATION_MISS_MS = 60000;        // nothing resolved at all, or a refusal: look again in a minute
const STATION_RECHECK_MS = 600000;    // running on the fallback: look for the real site every ten minutes
let stationSiteId = null;             // the site the lists are being read from
let stationSiteOwn = false;           // ... and whether that is the fallback (the workbook's own site)
let stationSiteMissAt = 0;            // when the Floor stations site was last looked for and not found
let stationSoftMiss = false;          // that miss was a refusal, not a 404: worth trying again soon
let stationSiteGen = 0;               // bumped when the lists move from one site to another
let stationSiteRead = false;          // has localStorage been consulted this page?

/** What is remembered between page loads: the id, and which of the two sites
    it is. An older build stored a bare id string; that is read as the real
    site, which is what it was. */
function loadStationSite() {
  if (stationSiteRead) return;
  stationSiteRead = true;
  let raw = null;
  try { raw = localStorage.getItem(STATION_SITE_KEY); } catch (e) { return; }
  if (!raw) return;
  if (raw.charAt(0) === "{") {
    try {
      const o = JSON.parse(raw);
      if (o && o.id) { stationSiteId = String(o.id); stationSiteOwn = !!o.own; }
    } catch (e) {}
    return;
  }
  stationSiteId = raw; stationSiteOwn = false;
}
function rememberStationSite(id, own) {
  if (stationSiteId && stationSiteId !== id) {
    /* the lists have moved. A list id is only meaningful in the site it was
       found in, so the ones cached against the old site go with it. */
    forgetListIdsFor(stationSiteId);
    stationSiteGen++;
  }
  stationSiteId = id; stationSiteOwn = !!own;
  stationSiteRead = true;
  try { localStorage.setItem(STATION_SITE_KEY, JSON.stringify({ id: id, own: !!own })); } catch (e) {}
}
/** Forget the list ids found in one site, in memory and in localStorage. */
function forgetListIdsFor(siteId) {
  const pre = siteId + "|";
  Object.keys(listIdMemo).forEach(k => { if (k.indexOf(pre) === 0) delete listIdMemo[k]; });
  const all = listIdCache();
  let hit = false;
  Object.keys(all).forEach(k => { if (k.indexOf(pre) === 0) { delete all[k]; hit = true; } });
  if (hit) try { localStorage.setItem(LISTIDS_KEY, JSON.stringify(all)); } catch (e) {}
}

/** The site the floor's lists are in, or null if neither can be resolved.
    Anything that is not an answer - offline, a bad gateway - is thrown, so the
    pages can tell "there is nowhere to read this from" apart from "I cannot
    see it just now" and leave the board they are already showing alone. */
async function stationSite() {
  loadStationSite();
  /* the real site, already found: there is nothing left to look for */
  if (stationSiteId && !stationSiteOwn) return stationSiteId;

  /* on the fallback (or with nothing at all yet), the Floor stations site is
     looked for again from time to time - every ten minutes while the fallback
     is working, and every minute when nothing resolved or the lookup was
     refused, because a refusal is usually a permission somebody is about to
     grant rather than a site that will never exist */
  const hold = stationSiteId && !stationSoftMiss ? STATION_RECHECK_MS : STATION_MISS_MS;
  if (stationSiteMissAt && Date.now() - stationSiteMissAt < hold) return stationSiteId;

  const host = SITE_PATH.split(":")[0];
  let real = null, missed = false, soft = false;
  try {
    const site = await call("GET", "/sites/" + host + ":/sites/" + STATION_SITE_NAME);
    real = (site && site.id) || null;
    missed = !real;
  } catch (e) {
    /* not there (404) and cannot see it (403) both mean "not today". Only the
       first is a settled fact; a refusal is worth asking about again soon. */
    if (!isMissing(e) && !isRefused(e)) throw e;
    missed = true; soft = isRefused(e) && !isMissing(e);
  }
  if (real) {
    stationSiteMissAt = 0; stationSoftMiss = false;
    rememberStationSite(real, false);
    return stationSiteId;
  }
  if (missed) { stationSiteMissAt = Date.now(); stationSoftMiss = soft; }
  if (stationSiteId) return stationSiteId;              // carry on with the fallback

  /* nothing to fall back to yet: resolve the workbook's own site, by path.
     SITE_AS_LIST marks it as the station's lookup rather than findFile()'s, so
     it is asked for quietly with the list scopes and can never put a consent
     window in front of somebody holding a sheet of glass. */
  let own = null;
  try {
    const site = await call("GET", "/sites/" + SITE_PATH + SITE_AS_LIST);
    own = (site && site.id) || null;
  } catch (e) {
    if (!isMissing(e) && !isRefused(e)) throw e;
    return null;
  }
  if (!own) return null;
  rememberStationSite(own, true);
  return stationSiteId;
}
/** Graph's two ways of saying "there is no such thing here". */
function isMissing(e) {
  const m = (e && e.message) || String(e || "");
  return /->\s*404\b/.test(m) || m.indexOf("itemNotFound") >= 0;
}
/** ... and its way of saying "there may well be, but not for you". */
function isRefused(e) {
  const m = (e && e.message) || String(e || "");
  return /->\s*403\b/.test(m) || m.indexOf("accessDenied") >= 0;
}
/** Forget the cached site, so the next call resolves it again. Called when a
    call against the cached id says the thing is not there: the id may be from
    another tenant, another browser profile, or a site since rebuilt.

    The move counter goes up unconditionally, and this is the whole point of
    it. Without that, a page that forgot the site and then resolved a DIFFERENT
    one would tell nobody: rememberStationSite() sees no previous id to compare
    against, so it stays quiet, and both screens carry on polling a delta token
    issued in the old site while writing into the new one. Forgetting a site is
    a move whether or not anything is known about where to next.

    The re-check cadence is NOT reset. `lookAgain` is for a person tapping "Try
    again", which is a reason to look right now; a failed call is not.        */
function forgetStationSite(lookAgain) {
  if (stationSiteId) forgetListIdsFor(stationSiteId);
  stationSiteId = null; stationSiteOwn = false; stationSiteRead = true;
  stationSiteGen++;
  if (lookAgain) { stationSiteMissAt = 0; stationSoftMiss = false; }
  try { localStorage.removeItem(STATION_SITE_KEY); } catch (e) {}
}
/** Which site the lists are being read from now, for a caller holding state
    that only means anything in one site - a delta token, say. It changes when
    the lists move, and never otherwise. */
function stationSiteMoves() { return stationSiteGen; }


window.CW = {
  initAuth, signIn, signOut, token, findFile, openSession, lastModified,
  downloadWorkbook, setFill, clearFill, setValues, rowForJob, A1,
  ensureLogSheet, appendLog, LOG_SHEET,
  ensureViewsSheet, saveAssignment, clearAssignment, VIEWS_SHEET,
  ensureProgressSheet, saveProgress, saveProgressMany, PROGRESS_SHEET, batchWrite,
  ensureAlertsSheet, addAlert, removeAlert, ALERTS_SHEET,
  listVersions, downloadVersion, restoreVersion,
  listId, listItems, listItem, listItemsFor, listUpsert, listDelete, listAdd, listPatch,
  listDelta, isDeltaRestart, isDeltaResync,
  listConsent, hasListConsent, LIST_SCOPES, SCOPES,
  stationSite, forgetStationSite, isMissing, isRefused, stationSiteMoves, STATION_SITE_NAME,
  liveBlocks, locateJob, moveJobRow, captureRow, batchGet,
  _setToken(fn) { tokenOverride = fn; }, _setFile(ref) { fileRef = ref; }, _setSession(id) { sessionId = id; },
  /* tests only: forget which dashboard sheets have been seen, so the creation
     branch of the ensure*Sheet functions can be exercised again */
  _resetSheetMemo() { logReady = null; viewsReady = null; progressReady = null; alertsReady = null; },
  /* tests only: forget the list ids found so far */
  _resetListIds() { Object.keys(listIdMemo).forEach(k => delete listIdMemo[k]); },
  /* tests only: which scopes a path is asked for - the same call headers()
     makes, so the answer cannot drift from what really goes out */
  _scopeFor: scopeFor,
  /* tests only: pin (or forget, with null) the site the lists are read from */
  _setStationSite(id, own) {
    stationSiteId = id || null; stationSiteOwn = !!own;
    stationSiteMissAt = 0; stationSoftMiss = false;
    /* clearing it puts the page back to before it had ever looked, so a test
       can put something in localStorage and watch it be read the way a fresh
       page load would read it */
    stationSiteRead = !!id;
    try { if (id) localStorage.setItem(STATION_SITE_KEY, JSON.stringify({ id: id, own: !!own }));
          else localStorage.removeItem(STATION_SITE_KEY); } catch (e) {}
  },
  /* tests only: which site is in use, and whether it is the fallback */
  _stationSiteInfo() { return { id: stationSiteId, own: stationSiteOwn, gen: stationSiteGen }; },
  /* tests only: pretend the last look for the Floor stations site was then, so
     a ten-minute re-check can be reached without waiting ten minutes */
  _stationSiteLookedAt(at) { stationSiteMissAt = at; },
  get account() { return account; }
};
