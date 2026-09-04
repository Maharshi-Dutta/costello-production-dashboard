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

async function token() {
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

/* ---- the audit log -------------------------------------------------------
   HARD RULE: every function below addresses LOG_SHEET and nothing else. The
   sheet name is a constant, never a parameter, so no call site can point this
   at Production or any other sheet.                                          */
const LOG_SHEET = "Dashboard Log";
const LOG_HEADERS = [["When", "Who", "Job", "What changed", "From", "To"]];
let logReady = false;

async function ensureLogSheet() {
  if (logReady) return true;
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
  logReady = true;
  return true;
}

/** Append one row to the log sheet. Never writes anywhere else. */
async function appendLog(who, job, what, from, to) {
  try {
    await ensureLogSheet();
    const f = await findFile();
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
let viewsReady = false;

async function ensureViewsSheet() {
  if (viewsReady) return true;
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
  viewsReady = true;
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

window.CW = {
  initAuth, signIn, signOut, token, findFile, openSession, lastModified,
  downloadWorkbook, setFill, clearFill, setValues, rowForJob, A1,
  ensureLogSheet, appendLog, LOG_SHEET,
  ensureViewsSheet, saveAssignment, clearAssignment, VIEWS_SHEET,
  listVersions, downloadVersion, restoreVersion,
  get account() { return account; }
};
