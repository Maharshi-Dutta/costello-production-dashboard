/* The floor tablet's SHELL - the parts of a station page that have nothing to
   do with which station it is.

   New on 2026-09-16 with the welding station. Every floor page needs the same
   five things and none of them is about glass or welding: the device's own
   light/dark choice, "how long ago was that" in words, the sign-in gate, the
   "Who are you?" picker with its PIN pad, and one HTML escape. Written once
   here so the third station (PA Lam) is a definition and a renderer and not a
   third copy of a page.

   NOTHING IN HERE TOUCHES GRAPH, A LIST OR A WORKBOOK. It is strings, a
   localStorage key for a theme, and two wiring helpers. Every identifier is
   prefixed `stu`/`STU` so that a page which loads this file beside another
   station's script cannot collide with it - browser scripts share one global
   lexical scope, and a collision there is a page that does not start.

   A NOTE ON glass.html: it loads this file but station.js does not yet call
   into it. station.js runs inside its own vm context in test_station.js, whose
   246 checks the welding spec says must pass unchanged, and that harness does
   not load this file - so moving the glass page onto this shell is a change to
   that suite, and belongs with the next piece of glass work rather than with
   this one. The file is on the page so that move is a code change and not an
   HTML one. See docs/STATIONS.md, "Adding a station".                       */

const STU_THEME_KEY = "cw_stationtheme";      // this device's choice, not this person's

const stuEsc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

/** "just now" / "40 s ago" / "12 min ago" / "3 h ago", from a millisecond
    stamp. Short on purpose: it sits in a header read at arm's length. */
function stuAgo(at) {
  const t = Number(at);
  if (!isFinite(t) || !t) return "";
  const s = Math.max(0, Math.round((Date.now() - t) / 1000));
  if (s < 10) return "just now";
  if (s < 90) return s + " s ago";
  if (s < 5400) return Math.round(s / 60) + " min ago";
  return Math.round(s / 3600) + " h ago";
}

/* ---- the device's own theme -------------------------------------------------
   The owner asked for dark by default on a tablet - a workshop screen at arm's
   length - with a switch for whoever prefers the other one. It belongs to the
   device and not to the person, so it lives in localStorage and survives
   somebody switching person, signing out or the page reloading itself. */
function stuThemeNow() {
  try { return localStorage.getItem(STU_THEME_KEY) === "light" ? "light" : "dark"; }
  catch (e) { return "dark"; }
}
/** Set the theme, remember it, and put the right word on the button. */
function stuApplyTheme(t, btn) {
  const dark = t !== "light";
  try { document.documentElement.dataset.theme = dark ? "dark" : "light"; } catch (e) {}
  try { localStorage.setItem(STU_THEME_KEY, dark ? "dark" : "light"); } catch (e) {}
  const b = btn || document.querySelector("#themebtn");
  if (b) b.textContent = dark ? "Light" : "Dark";
  return dark ? "dark" : "light";
}

/* ---- the sign-in gate -------------------------------------------------------
   Three nodes every station page has: #gate (signed out), #top (the header)
   and #main (the board). One of the first two is on screen and never both. */
function stuGate(on, err) {
  const q = s => document.querySelector(s);
  const set = (sel, hide, disp) => {
    const el = q(sel);
    if (!el) return;
    el.hidden = hide;
    if (el.style) el.style.display = hide ? "none" : disp;
  };
  const g = q("#gate");
  if (g) { g.hidden = !on; if (g.style) g.style.display = on ? "flex" : "none"; }
  set("#top", on, "flex");
  set("#main", on, "block");
  const e = q("#gateerr");
  if (e) {
    if (e.style) e.style.display = err ? "block" : "none";
    e.textContent = err || "";
  }
}

/* ---- "Who are you?" ---------------------------------------------------------
   Full screen, one big button per person, and a numeric pad for anybody whose
   PIN column is filled in. A wrong PIN shakes and says "Try again": there is no
   lockout, because locking a shared tablet out of the only screen the floor has
   would stop the work rather than protect it.

   Station-neutral in two places, and they are the two that used to be glass's:
   `empty` is the sentence shown when nobody is set up for THIS station, and
   `labelOf` turns a stage key into the word this station calls it. A station
   with one stage passes a labelOf that answers one word.                    */
function stuPickerHtml(o) {
  o = o || {};
  const people = o.people || [];
  const labelOf = typeof o.labelOf === "function" ? o.labelOf : (k => String(k));
  if (o.pinFor) {
    const dots = String(o.pinTyped || "").replace(/./g, "•");
    const key = t => '<button class="pk" data-pin="' + stuEsc(t) + '">' + stuEsc(t) + '</button>';
    return '<div class="picker">' +
      '<div class="pickh">' + stuEsc(o.pinFor.name) + '</div>' +
      '<div class="picksub">Enter your PIN</div>' +
      '<div class="pindots' + (o.pinBad ? " shake" : "") + '">' + stuEsc(dots || "····") + '</div>' +
      (o.pinBad ? '<div class="pinbad">Try again</div>' : "") +
      '<div class="pinpad">' + ["1", "2", "3", "4", "5", "6", "7", "8", "9"].map(key).join("") +
        '<button class="pk" data-pin="back">←</button>' + key("0") +
        '<button class="pk go" data-pin="ok">OK</button></div>' +
      '<button class="pcancel" data-pin="cancel">Back to the names</button>' +
      '</div>';
  }
  if (!people.length)
    return '<div class="picker"><div class="pickh">Who are you?</div>' +
      '<div class="msg">' + stuEsc(o.empty || "Nobody is set up for this station yet. " +
        "Ask the office to add you to the “Station people” list.") + '</div></div>';
  return '<div class="picker">' +
    '<div class="pickh">Who are you?</div>' +
    '<div class="pgrid">' + people.map(p =>
      '<button class="pbtn" data-person="' + stuEsc(p.name) + '">' +
        '<span class="pname">' + stuEsc(p.name) + '</span>' +
        '<span class="pstages">' + stuEsc((p.stages || []).length
          ? p.stages.map(labelOf).join(" · ") : "no stages yet") + '</span>' +
      '</button>').join("") + '</div></div>';
}

/** Wire the picker that stuPickerHtml just drew. `onPerson` gets the name that
    was tapped; `onKey` gets one of the pad's keys ("0".."9", "back", "ok",
    "cancel"). Neither knows anything about a list. */
function stuWirePicker(host, o) {
  if (!host || !host.querySelectorAll) return;
  o = o || {};
  host.querySelectorAll("[data-person]").forEach(el => {
    el.onclick = () => { if (typeof o.onPerson === "function") o.onPerson(el.dataset.person); };
  });
  host.querySelectorAll("[data-pin]").forEach(el => {
    el.onclick = () => { if (typeof o.onKey === "function") o.onKey(el.dataset.pin); };
  });
}

const STU = {
  STU_THEME_KEY,
  stuEsc, stuAgo, stuThemeNow, stuApplyTheme, stuGate, stuPickerHtml, stuWirePicker
};
if (typeof window !== "undefined") window.STU = STU;
else if (typeof globalThis !== "undefined") globalThis.STU = STU;
if (typeof module !== "undefined" && module.exports) module.exports = STU;
