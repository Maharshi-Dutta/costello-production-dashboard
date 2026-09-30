/* The Sales page - its rules and its writes, with no DOM.

   docs/specs/2026-09-30-sales-page.md. The sales team's page (sales.html) is
   the office dashboard with customer columns and a small set of writes the
   office page does not have: the six customer cells, the row's text colour,
   a delivery date (a SharePoint list, never the sheet), delete with restore,
   and requests to the office. Every write here refuses unless the page is the
   Sales page (window.CW_PAGE === "sales"), before it asks anything of Graph.

   Loaded by index.html (the office reads the same two lists: requests and
   delivery dates) and by sales.html, before app.js, and by test_sales.js in
   Node. Everything lives inside SALESC: classic scripts share one global
   scope (HISTORY B29), so nothing else is declared at the top level.       */
const SALESC = (function () {
  const PEOPLE = "Sales people", JOBS = "Sales jobs", BACKUPS = "Sales job backups", REQUESTS = "Sales requests";
  const PEOPLE_FIELDS = ["Title", "Active"];
  const JOBS_FIELDS = ["Title", "DeliveryDate", "SetBy", "SetAt"];
  const BACKUP_FIELDS = ["Title", "Job", "Kind", "Section", "Field", "From", "To", "Who", "At",
                         "Restored", "RestoredBy", "RestoredAt"];
  const REQUEST_FIELDS = ["Title", "Job", "Kind", "Text", "From", "At", "Reply", "ReplyBy", "ReplyAt", "ReplySeen"];
  const REQUEST_KINDS = [["status", "Status"], ["move", "Please move"], ["other", "Other"]];

  const isSales = () => typeof window !== "undefined" && !!window && window.CW_PAGE === "sales";
  function guard() { if (!isSales()) throw new Error("Only the Sales page may do this."); }
  const str = v => String(v == null ? "" : v);
  const yes = v => /^(yes|true|1)$/i.test(str(v).trim());

  /* ---- the row's text colour: one colour, so urgent and booked exclude ---- */
  const COLOURS = { urgent: "#FF0000", booked: "#00B050", "": "#000000" };
  const COLOUR_WORD = { urgent: "Urgent", booked: "Booked", "": "Black" };
  const flagWord = f => (f === "urgent" || f === "booked") ? f : "";
  /** Pressing one toggle: on if it was off (the other one goes), off if it was on. */
  const nextFlag = (cur, which) => flagWord(cur) === which ? "" : which;

  /* ---- moving: only once the office has marked the job ready ------------- */
  const NO_MOVE = ["In production", "Can sell as second hand"];
  const moveAllowed = j => !!(j && j.done);
  /** [{idx, name}] - every section but the two above and the one it is in. */
  const moveTargets = (names, cur) => (names || []).map((name, idx) => ({ idx, name }))
    .filter(t => NO_MOVE.indexOf(t.name) < 0 && t.idx !== cur);

  /* ---- the six customer cells, located by header through m.ident ---------- */
  const FIELDS = [
    { key: "cust", label: "Customer", of: j => j.cust },
    { key: "phone", label: "Phone no", of: j => j.ph, text: true },
    { key: "area", label: "Area", of: j => j.area },
    { key: "eir", label: "Eircode", of: j => j.eir, text: true },
    { key: "off", label: "Office no", of: j => j.off },
    { key: "colour", label: "Windows colour", of: j => j.colour }
  ];
  const fieldOf = key => FIELDS.find(f => f.key === key) || null;

  /* ---- the backup of a row, as JSON that fits a multi-line text column ---- */
  const ROW_MAX = 60000;
  const ARRS = { values: "", types: "Empty", numberFormat: "General", formulas: "", fills: "", fonts: null };
  function slimCap(cap) {
    const out = { height: cap.height || null };
    Object.keys(ARRS).forEach(k => {
      const o = {};
      (cap[k] || []).forEach((v, i) => { if (v != null && v !== "") o[i] = v; });
      out[k] = o;
    });
    return out;
  }
  function fullCap(s, n) {
    const cap = { height: s.height || null };
    Object.keys(ARRS).forEach(k => {
      const a = [];
      for (let i = 0; i < (n || 90); i++) a.push(s[k] && s[k][i] != null ? s[k][i] : ARRS[k]);
      cap[k] = a;
    });
    return cap;
  }
  /** The JSON stored in Row. Throws when it would not fit - the delete is refused. */
  function rowJson(cap, sectionName) {
    const s = JSON.stringify({ cap: slimCap(cap), sectionName: str(sectionName) });
    if (s.length > ROW_MAX) throw new Error("This row is too large to back up (" + s.length +
      " characters, the limit is " + ROW_MAX + "), so it was not deleted.");
    return s;
  }
  function parseRow(json) {
    const o = JSON.parse(str(json));
    return { cap: fullCap(o.cap || {}), sectionName: str(o.sectionName) };
  }

  /* ---- requests and replies ---------------------------------------------- */
  const reqF = r => (r && r.fields) || r || {};
  const unanswered = reqs => (reqs || []).filter(r => !str(reqF(r).Reply).trim());
  const unreadReplies = reqs => (reqs || []).filter(r => str(reqF(r).Reply).trim() && !yes(reqF(r).ReplySeen));
  const openFor = (reqs, job) => unanswered(reqs).some(r => str(reqF(r).Job).toUpperCase() === str(job).toUpperCase());
  const newestFirst = reqs => (reqs || []).slice().sort((a, b) => str(reqF(b).At).localeCompare(str(reqF(a).At)));
  /** The office's window: unanswered first, then newest first within each. */
  const officeOrder = reqs => {
    const n = newestFirst(reqs), u = unanswered(n);
    return u.concat(n.filter(r => u.indexOf(r) < 0));
  };

  /* ---- delivery dates ---------------------------------------------------- */
  const DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  /** "2026-10-07" -> "Wed 07 Oct"; anything else back as it was. */
  function dayWords(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(str(iso));
    if (!m) return str(iso);
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return DAYS[d.getUTCDay()] + " " + m[3] + " " + MONTHS[+m[2] - 1];
  }
  /** Sales jobs items -> { JOB: { date, by, at } }, empty dates left out. */
  function deliveryMap(items) {
    const out = {};
    (items || []).forEach(it => {
      const f = reqF(it), job = str(f.Title).trim().toUpperCase(), date = str(f.DeliveryDate).trim();
      if (job && date) out[job] = { date, by: str(f.SetBy), at: str(f.SetAt), id: it.id };
    });
    return out;
  }

  /* ---- the writes, one at a time per page -------------------------------- */
  let chain = Promise.resolve();
  function serial(fn) {
    const run = chain.then(fn, fn);
    chain = run.then(() => {}, () => {});
    return run;
  }
  function need(ctx) {
    guard();
    if (!ctx || !ctx.CW) throw new Error("no Graph layer");
    if (!str(ctx.who).trim()) throw new Error("Pick your name first.");
  }
  const now = ctx => (ctx.now ? ctx.now() : new Date().toISOString());
  const title = (job, at) => str(job).toUpperCase() + "|" + at;

  /** Add a backup item and read it back; throws unless Row came back the same. */
  async function backup(ctx, fields) {
    const made = await ctx.CW.listAdd(BACKUPS, fields);
    if (!made || !made.id) throw new Error("The backup was not saved, so nothing was changed.");
    if (fields.Row != null) {
      const back = await ctx.CW.listItem(BACKUPS, made.id, { fields: ["Row"] });
      if (!back || str(back.fields && back.fields.Row) !== fields.Row)
        throw new Error("The backup could not be read back, so nothing was changed.");
    }
    return made;
  }

  /** One customer cell (2b). Written, logged, then recorded in the backups. */
  function saveCustomer(ctx, j, key, value) {
    need(ctx);
    const F = fieldOf(key);
    if (!F) throw new Error("not a customer field: " + key);
    const col = ctx.identCol(key);
    if (!col) throw new Error("The Production sheet has no '" + F.label + "' column.");
    const from = str(F.of(j)), to = str(value).trim();
    return serial(async () => {
      const r = await ctx.CW.salesLocate(j.id);
      await ctx.CW.salesSetCell(j.id, r, col, to, !!F.text);
      ctx.note(j.id, F.label, from, to);
      const at = now(ctx);
      try {
        await ctx.CW.listAdd(BACKUPS, { Title: title(j.id, at), Job: j.id, Kind: "edit", Field: F.label,
                                        From: from, To: to, Who: ctx.who, At: at, Restored: "No" });
      } catch (e) { /* the cell is written and logged; the record is the extra */ }
      return { row: r, from, to };
    });
  }

  /** The row's text colour (2a): backed up and read back BEFORE the paint. */
  function setColour(ctx, j, want, sectionName) {
    need(ctx);
    want = flagWord(want);
    const from = COLOUR_WORD[flagWord(j.flag)] || str(j.flag), to = COLOUR_WORD[want];
    return serial(async () => {
      const r = await ctx.CW.salesLocate(j.id);
      const cap = await ctx.CW.captureRow(r, ctx.tmplFor ? ctx.tmplFor(j.id) : null);
      const at = now(ctx);
      await backup(ctx, { Title: title(j.id, at), Job: j.id, Kind: "colour", Section: str(sectionName),
                          From: from, To: to, Row: rowJson(cap, sectionName), Who: ctx.who, At: at, Restored: "No" });
      await ctx.CW.salesSetFont(j.id, r, COLOURS[want]);
      ctx.note(j.id, "Text colour", from, to);
      return { row: r, from, to };
    });
  }

  /** Delete the row (2c), only after its backup has been read back whole. */
  function deleteJob(ctx, j, sectionName) {
    need(ctx);
    return serial(async () => {
      const r = await ctx.CW.salesLocate(j.id);
      const cap = await ctx.CW.captureRow(r, ctx.tmplFor ? ctx.tmplFor(j.id) : null);
      const json = rowJson(cap, sectionName);            // too big: refused before anything is saved
      const at = now(ctx);
      const made = await backup(ctx, { Title: title(j.id, at), Job: j.id, Kind: "delete", Section: str(sectionName),
                                       Row: json, Who: ctx.who, At: at, Restored: "No" });
      const d = await ctx.CW.salesDeleteRow(j.id, ctx.tmplFor);   // found and checked again in there
      ctx.note(j.id, "Deleted", str(sectionName), "");
      return { itemId: made.id, row: d.row };
    });
  }

  /** Put a deleted row back at the bottom of the section it came from. */
  function restoreJob(ctx, item) {
    need(ctx);
    const f = reqF(item);
    if (str(f.Kind) !== "delete") throw new Error("Only a delete can be restored.");
    if (yes(f.Restored)) throw new Error("That job has already been restored.");
    const job = str(f.Job).toUpperCase();
    return serial(async () => {
      /* Row is read fresh: the window's listing leaves the big column out */
      const full = await ctx.CW.listItem(BACKUPS, item.id, { fields: ["Row"] });
      if (!full || !str(full.fields && full.fields.Row)) throw new Error("The backup of " + job + " could not be read.");
      const saved = parseRow(full.fields.Row);
      const section = saved.sectionName || str(f.Section);
      const put = await ctx.CW.salesInsertRow(job, saved.cap, section, ctx.tmplFor);
      const at = now(ctx);
      await ctx.CW.listPatch(BACKUPS, item.id, { Restored: "Yes", RestoredBy: ctx.who, RestoredAt: at });
      await ctx.CW.listAdd(BACKUPS, { Title: title(job, at), Job: job, Kind: "restore", Section: section,
                                      Who: ctx.who, At: at, Restored: "No" });
      ctx.note(job, "Restored", "", section);
      return put;
    });
  }

  /** The delivery date: a list row per job, never the sheet. "" clears it. */
  function setDelivery(ctx, j, date, from) {
    need(ctx);
    const to = str(date).trim();
    if (to && !/^\d{4}-\d{2}-\d{2}$/.test(to)) throw new Error("A delivery date is YYYY-MM-DD.");
    return serial(async () => {
      await ctx.CW.listUpsert(JOBS, j.id, { DeliveryDate: to, SetBy: ctx.who, SetAt: now(ctx) });
      ctx.note(j.id, "Delivery date", str(from), to);
      return to;
    });
  }

  /** A question to the office about one job. */
  function sendRequest(ctx, job, kind, text) {
    need(ctx);
    const k = REQUEST_KINDS.some(p => p[0] === kind) ? kind : "other", t = str(text).trim().slice(0, 1000);
    if (!t) throw new Error("Write the request first.");
    return serial(async () => {
      const at = now(ctx);
      const made = await ctx.CW.listAdd(REQUESTS, { Title: title(job, at), Job: str(job).toUpperCase(), Kind: k,
                                                     Text: t, From: ctx.who, At: at, ReplySeen: "No" });
      ctx.note(str(job).toUpperCase(), "Sales request", "", t);
      return made;
    });
  }

  /** The replies on screen that have not been marked seen: ReplySeen = Yes. */
  function markSeen(ctx, reqs) {
    guard();
    const todo = unreadReplies(reqs);
    return serial(async () => {
      for (let i = 0; i < todo.length; i++) await ctx.CW.listPatch(REQUESTS, todo[i].id, { ReplySeen: "Yes" });
      return todo.length;
    });
  }

  return {
    PEOPLE, JOBS, BACKUPS, REQUESTS, PEOPLE_FIELDS, JOBS_FIELDS, BACKUP_FIELDS, REQUEST_FIELDS, REQUEST_KINDS,
    COLOURS, COLOUR_WORD, NO_MOVE, FIELDS, ROW_MAX,
    isSales, flagWord, nextFlag, moveAllowed, moveTargets, fieldOf, yes,
    slimCap, fullCap, rowJson, parseRow,
    unanswered, unreadReplies, openFor, newestFirst, officeOrder, dayWords, deliveryMap,
    serial, saveCustomer, setColour, deleteJob, restoreJob, setDelivery, sendRequest, markSeen
  };
})();
if (typeof window !== "undefined") window.SALESC = SALESC;
if (typeof module !== "undefined" && module.exports) module.exports = SALESC;
