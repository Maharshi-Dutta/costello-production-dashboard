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
  /* 2026-10-01 (reports, complaints, photos): five more columns, added by the
     session's script. Until all five are there nothing reads or writes them. */
  const REQUEST_EXTRA = ["Customer", "Status", "ResolvedBy", "ResolvedAt", "Photos"];
  const REQUEST_KINDS = [["status", "Status"], ["move", "Please move"], ["report", "Report"],
                         ["complaint", "Customer complaint"], ["other", "Other"]];
  const FORM_KINDS = ["status", "move", "report", "complaint"];     // what the send form offers
  const NO_JOB_OK = ["report", "complaint"];                          // these may name a customer instead

  const isSales = () => typeof window !== "undefined" && !!window && window.CW_PAGE === "sales";
  function guard() { if (!isSales()) throw new Error("Only the Sales page may do this."); }
  const str = v => String(v == null ? "" : v);
  const yes = v => /^(yes|true|1)$/i.test(str(v).trim());

  /* ---- the row's text colour: one colour, so urgent and booked exclude ---- */
  const COLOURS = { urgent: "#FF0000", booked: "#00B050", "": "#000000" };
  const COLOUR_WORD = { urgent: "Urgent", booked: "Booked", trade: "Trade order", hold: "On hold", "": "Black" };
  const INK_WORD = { trade: "pink", hold: "blue" };
  const flagWord = f => (f === "urgent" || f === "booked") ? f : "";
  /** What the row's text says now, as logged: the real word, Trade order and On hold included. */
  const fromWord = f => COLOUR_WORD[f] || (f ? String(f) : "Black");
  /** Pressing one toggle: on if it was off (the other one goes), off if it was on. */
  const nextFlag = (cur, which) => flagWord(cur) === which ? "" : which;
  /** A Trade order / On hold row is somebody else's colour: ask before replacing it. */
  function replaceQuestion(cur, want) {
    if (!INK_WORD[cur]) return "";
    return "This row's text is " + INK_WORD[cur] + " (" + COLOUR_WORD[cur] + "). Replace with " +
      (want === "urgent" ? "red (Urgent)" : want === "booked" ? "green (Booked)" : "black") + "?";
  }
  /** "0871234567" -> "•••567": phone and eircode are never logged whole. */
  const mask = v => { const s = String(v == null ? "" : v).trim(); return s ? "•••" + s.slice(-3) : ""; };

  /* ---- moving: only once the office has marked the job ready ------------- */
  const NO_MOVE = ["In production", "Can sell as second hand"];
  const moveAllowed = j => !!(j && j.done);
  /** [{idx, name}] - every section but the two above and the one it is in. */
  const moveTargets = (names, cur) => (names || []).map((name, idx) => ({ idx, name }))
    .filter(t => NO_MOVE.indexOf(t.name) < 0 && t.idx !== cur);

  /* ---- the six customer cells, located by header through m.ident ---------- */
  /* hdr: the header the parser located the column by (parser.js IDENT);
     secret: logged masked (rule 4), whole only in the backups list */
  const FIELDS = [
    { key: "cust", label: "Customer", hdr: "customer", of: j => j.cust },
    { key: "phone", label: "Phone no", hdr: "phone no", of: j => j.ph, secret: true },
    { key: "area", label: "Area", hdr: "area", of: j => j.area },
    { key: "eir", label: "Eircode", hdr: "eircode", of: j => j.eir, secret: true },
    { key: "off", label: "Office no", hdr: "office no", of: j => j.off },
    { key: "colour", label: "Windows colour", hdr: "windows colour", of: j => j.colour }
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
  const extrasIn = cols => !!cols && REQUEST_EXTRA.every(c => cols.indexOf(c) >= 0);
  const isComplaint = r => str(reqF(r).Kind) === "complaint";
  /** Open unless the list says Resolved: a complaint whose Status is empty (sent
      before the columns existed, or while the column check had not finished)
      is open - and can be resolved, since the columns now exist (review 1). */
  const openComplaint = r => isComplaint(r) && str(reqF(r).Status) !== "Resolved";
  /** The office's bell: unanswered messages plus open complaints, each once. */
  const bellCount = reqs => (reqs || []).filter(r => !str(reqF(r).Reply).trim() || openComplaint(r)).length;
  /** The office's window: open complaints first, then unanswered, then the
      rest - newest first within each. */
  const officeOrder = reqs => {
    const n = newestFirst(reqs), c = n.filter(openComplaint), u = unanswered(n).filter(r => c.indexOf(r) < 0);
    return c.concat(u, n.filter(r => c.indexOf(r) < 0 && u.indexOf(r) < 0));
  };
  const kindWord = k => (REQUEST_KINDS.find(p => p[0] === k) || ["", str(k) || "Other"])[1];
  const photoCount = r => Number(reqF(r).Photos) || 0;

  /* ---- photos: shrunk in the browser, then one folder per message --------- */
  const PHOTO_MAX = 8, PHOTO_SIDE = 1600, PHOTO_QUALITY = 0.82, PHOTO_BYTES = 40 * 1024 * 1024;
  /** The size a photo is drawn at: longest side PHOTO_SIDE, aspect kept, never upscaled. */
  function fitSize(w, h, max) {
    max = max || PHOTO_SIDE;
    const k = Math.min(1, max / Math.max(w, h, 1));
    return { w: Math.max(1, Math.round(w * k)), h: Math.max(1, Math.round(h * k)) };
  }
  /** Why these files cannot be added to `already` photos ("" = they can). */
  function photoRefusal(files, already) {
    files = files || [];
    if ((already || 0) + files.length > PHOTO_MAX) return "At most " + PHOTO_MAX + " photos per message.";
    const bad = files.find(f => !/^image\//i.test(str(f && f.type)));
    if (bad) return "“" + str(bad.name || "That file") + "” is not a photo.";
    /* refused before the browser tries to decode it (review 8) */
    const huge = files.find(f => Number(f && f.size) > PHOTO_BYTES);
    return huge ? "“" + str(huge.name || "That file") + "” is over 40 MB, so it was not added." : "";
  }
  /** The next free file number in a folder listing: one past the highest n.jpg,
      so a re-send never lands on (and with replace, never overwrites) a photo. */
  const nextPhotoNo = list => (list || []).reduce((m, p) => Math.max(m, parseInt(str(p && p.name), 10) || 0), 0);
  /** The message's folder in the library: its Title, with | and : made safe. */
  const photoFolder = t => str(t).replace(/[|:"*<>?\/\\#%]/g, "-");
  const pad2 = n => (n < 10 ? "0" : "") + n;
  function whenWords(iso) {
    const d = new Date(str(iso));
    return isNaN(d) ? str(iso) : DAYS[d.getDay()] + " " + pad2(d.getDate()) + " " + MONTHS[d.getMonth()] + " " + pad2(d.getHours()) + ":" + pad2(d.getMinutes());
  }

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
  /* Every Sales write - and the Sales page's move - runs through here, one at
     a time. `busy()` is true from the moment one is asked for until the last
     has finished, and `onBusy` hears each change, so the page can hold every
     write control on every drawer and window while anything is in flight. */
  let chain = Promise.resolve(), inFlight = 0;
  const api = { onBusy: null };
  const busyTell = () => { if (typeof api.onBusy === "function") try { api.onBusy(inFlight > 0); } catch (e) {} };
  function serial(fn) {
    inFlight++; if (inFlight === 1) busyTell();
    const run = chain.then(fn, fn);
    chain = run.then(() => {}, () => {});
    const done = () => { inFlight--; if (!inFlight) busyTell(); };
    run.then(done, done);
    return run;
  }
  const busy = () => inFlight > 0;
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
    const col = ctx.identCol(key), hdrRow = ctx.hdrRow ? ctx.hdrRow() : 0;
    if (!col) throw new Error("The Production sheet has no '" + F.label + "' column.");
    const from = str(F.of(j)), to = str(value).trim();
    return serial(async () => {
      const L = await ctx.CW.salesLocate(j.id);
      await ctx.CW.salesSetCell(j.id, L.row, col, to, { row: hdrRow, label: F.hdr });
      /* the Dashboard Log never carries a whole phone number or eircode */
      ctx.note(j.id, F.label, F.secret ? mask(from) : from, F.secret ? mask(to) : to);
      const at = now(ctx);
      let warn = "";
      try {
        await ctx.CW.listAdd(BACKUPS, { Title: title(j.id, at), Job: j.id, Kind: "edit", Field: F.label,
                                        From: from, To: to, Who: ctx.who, At: at, Restored: "No" });
      } catch (e) { warn = "Written and logged, but the Sales job backups record failed: " + ((e && e.message) || e); }
      return { row: L.row, from, to, warn };
    });
  }

  /** The row's text colour (2a): backed up and read back BEFORE the paint. */
  function setColour(ctx, j, want) {
    need(ctx);
    want = flagWord(want);
    const from = fromWord(j.flag), to = COLOUR_WORD[want];
    return serial(async () => {
      const L = await ctx.CW.salesLocate(j.id);          // the section is the live one, not the download's
      const cap = await ctx.CW.captureRow(L.row, ctx.tmplFor ? ctx.tmplFor(j.id) : null);
      const at = now(ctx);
      await backup(ctx, { Title: title(j.id, at), Job: j.id, Kind: "colour", Section: L.section,
                          From: from, To: to, Row: rowJson(cap, L.section), Who: ctx.who, At: at, Restored: "No" });
      await ctx.CW.salesSetFont(j.id, L.row, COLOURS[want]);   // refused if the row has moved since
      ctx.note(j.id, "Text colour", from, to);
      return { row: L.row, from, to };
    });
  }

  /** Delete the row (2c), only after its backup has been read back whole -
      and only the row that was backed up: refused if it has moved since. */
  function deleteJob(ctx, j) {
    need(ctx);
    return serial(async () => {
      const L = await ctx.CW.salesLocate(j.id);
      const cap = await ctx.CW.captureRow(L.row, ctx.tmplFor ? ctx.tmplFor(j.id) : null);
      const json = rowJson(cap, L.section);              // too big: refused before anything is saved
      const at = now(ctx);
      const made = await backup(ctx, { Title: title(j.id, at), Job: j.id, Kind: "delete", Section: L.section,
                                       Row: json, Who: ctx.who, At: at, Restored: "No" });
      const d = await ctx.CW.salesDeleteRow(j.id, L.row, ctx.tmplFor,
        () => ctx.note(j.id, "Deleted", L.section, ""));   // logged straight after the delete itself
      return { itemId: made.id, row: d.row, section: L.section, edgeError: d.edgeError || "" };
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
      const put = await ctx.CW.salesInsertRow(job, saved.cap, section, ctx.tmplFor,
        () => ctx.note(job, "Restored", "", section));    // logged the moment the row is confirmed back
      const at = now(ctx);
      /* the row is back whatever happens next; a list failure is reported, and
         a second restore is refused anyway because the job is on the sheet */
      try {
        await ctx.CW.listPatch(BACKUPS, item.id, { Restored: "Yes", RestoredBy: ctx.who, RestoredAt: at });
        await ctx.CW.listAdd(BACKUPS, { Title: title(job, at), Job: job, Kind: "restore", Section: section,
                                        Who: ctx.who, At: at, Restored: "No" });
      } catch (e) { put.warn = "Restored and logged, but the backups list was not updated: " + ((e && e.message) || e); }
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

  const errText = e => (e && e.message) || String(e);
  /** Has the list got the five new columns? ctx.extras is a flag, or (on the
      page) a function that runs the column check first if it has not finished
      or had failed, so a complaint is never sent without its Status (review 1). */
  const extrasOf = async ctx => typeof ctx.extras === "function" ? !!(await ctx.extras()) : !!ctx.extras;
  /** Upload photos into one message's folder, one by one, numbered from `start`
      + 1, then PATCH Photos to what the folder really holds (its listing; had +
      landed only if the listing fails). A failed photo never undoes the message. */
  async function putPhotos(ctx, id, folder, blobs, start, had) {
    if (!(await ctx.CW.salesPhotoDrive()))
      return { landed: 0, failed: blobs.length, photos: had, note: "The “Sales photos” library is not in SharePoint yet, so no photo was sent." };
    let landed = 0;
    for (let i = 0; i < blobs.length; i++) {
      try { await ctx.CW.salesPhotoPut(folder, (start + i + 1) + ".jpg", blobs[i]); landed++; }
      catch (e) { if (typeof console !== "undefined") console.warn("[sales] photo: " + errText(e)); }
    }
    const failed = blobs.length - landed;
    let note = failed ? failed + " of " + blobs.length + " photos failed — send them again from the message" : "";
    let photos = had + landed;
    try { photos = (await ctx.CW.salesPhotoList(folder)).length; }          // review 3: the folder is the truth
    catch (e) { if (typeof console !== "undefined") console.warn("[sales] photo listing: " + errText(e)); }
    if (photos !== had) {
      try { await ctx.CW.listPatch(REQUESTS, id, { Photos: photos }); }
      catch (e) { note = (note ? note + ". " : "") + "The photos are saved, but their count could not be: " + errText(e); }
    }
    return { landed, failed, photos, note };
  }

  /** A message to the office: Status / Move about a job, or a Report /
      Customer complaint about a job or a customer. Free text plus photos
      (Blobs, already shrunk). The item is added first; the photos follow; a
      photo that fails leaves the message sent. One Dashboard Log line: the
      kind, the job and the photo count - never the text. */
  function sendMessage(ctx, m) {
    need(ctx);
    m = m || {};
    const kind = REQUEST_KINDS.some(p => p[0] === m.kind) ? m.kind : "other";
    const job = str(m.job).trim().toUpperCase().slice(0, 40), cust = str(m.customer).trim().slice(0, 255);
    const text = str(m.text).trim().slice(0, 4000), photos = (m.photos || []).filter(Boolean);
    if (!text) throw new Error("Write the message first.");
    if (!job && NO_JOB_OK.indexOf(kind) < 0) throw new Error("A " + kindWord(kind).toLowerCase() + " message needs a job number.");
    /* review 9: a Status or Move is about a job on the sheet */
    if (NO_JOB_OK.indexOf(kind) < 0 && ctx.hasJob && !ctx.hasJob(job))
      throw new Error(job + " is not on the Production sheet. Check the number, or send it as a Report.");
    if (photos.length > PHOTO_MAX) throw new Error("At most " + PHOTO_MAX + " photos per message.");
    return serial(async () => {
      const extras = await extrasOf(ctx);
      const at = now(ctx);
      const fields = { Title: title(job || "NO JOB", at), Job: job, Kind: kind, Text: text, From: ctx.who, At: at, ReplySeen: "No" };
      if (extras) { fields.Customer = cust; fields.Photos = 0; if (kind === "complaint") fields.Status = "Open"; }
      else if (cust) fields.Text = "Customer: " + cust + "\n" + text;     // no Customer column yet: keep it in the text
      const made = await ctx.CW.listAdd(REQUESTS, fields);
      const id = made && made.id;
      let r = { landed: 0, failed: 0, note: "" };
      if (photos.length) {
        r = extras ? await putPhotos(ctx, id, photoFolder(fields.Title), photos, 0, 0)
          : { landed: 0, failed: photos.length, note: "Sent without the photos: the “Sales requests” list has no Photos column yet." };
        if (extras) fields.Photos = r.photos;
      }
      ctx.note(job || "—", "Sales message: " + kindWord(kind), "", r.landed + (r.landed === 1 ? " photo" : " photos"));
      return Object.assign({ id, fields }, r);
    });
  }
  /** The drawer's old shape: a question about one job. */
  const sendRequest = (ctx, job, kind, text) => sendMessage(ctx, { job, kind, text });

  /** More photos on a message already sent (a failed photo, or one forgotten). */
  function addPhotos(ctx, item, blobs) {
    need(ctx);
    blobs = (blobs || []).filter(Boolean);
    if (!blobs.length) throw new Error("Pick a photo first.");
    return serial(async () => {
      if (!(await extrasOf(ctx))) throw new Error("The “Sales requests” list has no Photos column yet, so photos cannot be sent.");
      const cur = await ctx.CW.listItem(REQUESTS, item.id, { fields: REQUEST_FIELDS.concat(REQUEST_EXTRA) });
      if (!cur) throw new Error("That message is no longer in the list.");
      const folder = photoFolder(cur.fields.Title);
      /* the folder says what is there and which numbers are taken (review 3) */
      let list = null;
      try { list = await ctx.CW.salesPhotoList(folder); } catch (e) {}
      const had = list ? list.length : photoCount(cur), start = list ? nextPhotoNo(list) : had;
      if (had + blobs.length > PHOTO_MAX) throw new Error("At most " + PHOTO_MAX + " photos per message (" + had + " there already).");
      const r = await putPhotos(ctx, item.id, folder, blobs, start, photoCount(cur));
      ctx.note(str(cur.fields.Job) || "—", "Sales photos added", had, r.photos);
      return r;
    });
  }

  /** Mark a complaint Resolved, or Reopen it - from either page. Reads the item
      first: if it already says what was asked, nothing is written and the
      error names who and when. Writes Status, ResolvedBy, ResolvedAt only. */
  async function setComplaint(CW, id, want, who, at) {
    want = want === "Resolved" ? "Resolved" : "Open";
    const cur = await CW.listItem(REQUESTS, id, { fields: REQUEST_FIELDS.concat(REQUEST_EXTRA) });
    if (!cur) throw new Error("That message is no longer in the list.");
    const f = cur.fields || {};
    if (str(f.Kind) !== "complaint") throw new Error("Only a complaint can be resolved.");
    const is = str(f.Status) === "Resolved" ? "Resolved" : "Open";
    if (is === want) {
      const by = str(f.ResolvedBy).split("@")[0];
      const e = new Error(want === "Resolved" ? "Already resolved by " + (by || "someone") + (f.ResolvedAt ? ", " + whenWords(f.ResolvedAt) : "") + " - nothing was written."
        : "It is already open" + (by ? " (reopened by " + by + (f.ResolvedAt ? ", " + whenWords(f.ResolvedAt) : "") + ")" : "") + " - nothing was written.");
      e.current = f; throw e;
    }
    const fields = { Status: want, ResolvedBy: str(who), ResolvedAt: at || new Date().toISOString() };
    /* review 4: only if nobody has changed the item since that read */
    const etag = cur.etag || f["@odata.etag"];
    try { await CW.listPatch(REQUESTS, id, fields, etag ? { ifMatch: etag } : undefined); }
    catch (e) {
      if (!/->\s*412\b/.test(errText(e))) throw e;
      const now = await CW.listItem(REQUESTS, id, { fields: REQUEST_FIELDS.concat(REQUEST_EXTRA) });
      const nf = (now && now.fields) || f, nby = str(nf.ResolvedBy).split("@")[0];
      const e2 = new Error("Someone got there first - it now says " + (str(nf.Status) || "Open") +
        (nby ? " (" + nby + (nf.ResolvedAt ? ", " + whenWords(nf.ResolvedAt) : "") + ")" : "") + ". Nothing was written.");
      e2.current = nf; throw e2;
    }
    return Object.assign({}, f, fields);
  }
  /** The Sales page's side of it: a name, the queue, a log line. */
  function resolve(ctx, id, want) {
    need(ctx);
    return serial(async () => {
      const f = await setComplaint(ctx.CW, id, want, ctx.who, now(ctx));
      ctx.note(str(f.Job) || "—", "Complaint", f.Status === "Resolved" ? "Open" : "Resolved", f.Status);
      return f;
    });
  }

  /** The replies on screen that have not been marked seen: ReplySeen = Yes. */
  /* ids already queued or written this page load: a reply is marked once (review 2) */
  const seenQueued = {};
  function markSeen(ctx, reqs) {
    need(ctx);
    const todo = unreadReplies(reqs).filter(r => !seenQueued[r.id]);
    todo.forEach(r => { seenQueued[r.id] = 1; });
    return serial(async () => {
      for (let i = 0; i < todo.length; i++) {
        try { await ctx.CW.listPatch(REQUESTS, todo[i].id, { ReplySeen: "Yes" }); }
        catch (e) { for (let k = i; k < todo.length; k++) delete seenQueued[todo[k].id]; throw e; }   // may be asked again
      }
      return todo.length;
    });
  }

  /** The office's reply: Reply, ReplyBy, ReplyAt only - and only if nobody has
      replied since this screen read the list (read the item first). Answers
      the item's fields as they are after the write. */
  async function officeReply(CW, id, text, who, at) {
    if (isSales()) throw new Error("The office replies from the office page.");
    const t = str(text).trim().slice(0, 1000);
    if (!t) throw new Error("Write the reply first.");
    const cur = await CW.listItem(REQUESTS, id, { fields: REQUEST_FIELDS });
    if (!cur) throw new Error("That request is no longer in the list.");
    if (str(cur.fields.Reply).trim()) {
      const e = new Error("Already answered by " + (str(cur.fields.ReplyBy).split("@")[0] || "someone") + " - nothing was written.");
      e.current = cur.fields; throw e;
    }
    const fields = { Reply: t, ReplyBy: str(who), ReplyAt: at || new Date().toISOString() };
    await CW.listPatch(REQUESTS, id, fields);
    return Object.assign({}, cur.fields, fields);
  }

  /** The Sales page's move (decision 5), in the same queue as every other
      Sales write. `move(ids, idx)` is app.js's moveJobsInSheet. */
  function moveJob(ctx, j, idx, names, move) {
    need(ctx);
    const name = (names || [])[idx];
    if (!moveAllowed(j)) throw new Error("The office has not marked " + j.id + " ready yet.");
    if (!name || NO_MOVE.indexOf(name) >= 0) throw new Error(j.id + " cannot be moved to " + (name || "that section") + " from here.");
    return serial(() => move([j.id], idx));
  }

  /** Several jobs at once - a drag onto a section, or "move to" on a selection
      (amendment B). The same gate and targets as the drawer's Move and the
      same queue; a job the office has not marked ready is refused and nothing
      moves for it. Answers { moved, refused: [ids] }. */
  const REFUSED_NOTE = "the office has not marked this job ready - send a request";
  function moveMany(ctx, jobs, idx, names, move) {
    need(ctx);
    const name = (names || [])[idx];
    if (!name || NO_MOVE.indexOf(name) >= 0) throw new Error("Jobs cannot be moved to " + (name || "that section") + " from here.");
    const list = (jobs || []).filter(Boolean);
    const ready = list.filter(moveAllowed), refused = list.filter(j => !moveAllowed(j)).map(j => j.id);
    if (!ready.length) return Promise.resolve({ moved: 0, refused });
    return serial(() => move(ready.map(j => j.id), idx)).then(n => ({ moved: Number(n) || 0, refused }));
  }

  return Object.assign(api, {
    REFUSED_NOTE, moveMany,
    PEOPLE, JOBS, BACKUPS, REQUESTS, PEOPLE_FIELDS, JOBS_FIELDS, BACKUP_FIELDS, REQUEST_FIELDS, REQUEST_KINDS,
    REQUEST_EXTRA, FORM_KINDS, NO_JOB_OK, PHOTO_MAX, PHOTO_SIDE, PHOTO_QUALITY, PHOTO_BYTES, nextPhotoNo,
    extrasIn, isComplaint, openComplaint, bellCount, kindWord, photoCount, fitSize, photoRefusal, photoFolder, whenWords,
    sendMessage, addPhotos, setComplaint, resolve,
    COLOURS, COLOUR_WORD, NO_MOVE, FIELDS, ROW_MAX,
    isSales, flagWord, fromWord, nextFlag, replaceQuestion, mask, moveAllowed, moveTargets, fieldOf, yes,
    slimCap, fullCap, rowJson, parseRow,
    unanswered, unreadReplies, openFor, newestFirst, officeOrder, dayWords, deliveryMap,
    serial, busy, saveCustomer, setColour, deleteJob, restoreJob, setDelivery, sendRequest, markSeen, moveJob, officeReply
  });
})();
if (typeof window !== "undefined") window.SALESC = SALESC;
if (typeof module !== "undefined" && module.exports) module.exports = SALESC;
