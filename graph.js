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

/* Graph occasionally 504s while Excel loads a large workbook - retry those. */
async function call(method, path, body, asBuffer) {
  let last;
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
    last = r;
    if ([429, 500, 503, 504].indexOf(r.status) >= 0) {
      await new Promise(s => setTimeout(s, 2000 + a * 2500));
      continue;
    }
    break;
  }
  throw new Error(method + " " + path + " -> " + last.status + " " + (await last.text()).slice(0, 300));
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
  try {
    const s = await call("POST", f.base + "/createSession", { persistChanges: true });
    sessionId = s.id;
  } catch (e) { sessionId = null; }   // sessions are an optimisation, not a requirement
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
  get account() { return account; }
};
