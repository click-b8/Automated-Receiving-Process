/* Receiving Dashboard — ACME Medical receiving automation (plain JS, no build step).
 * Reads Project Map/_scripts/.dashboard_data.json (written by refresh_dashboard.py),
 * the vault's own files, and (optionally) one monday.com board. Read-only everywhere:
 * the only thing it ever starts is the read-only collector script. */
const { Plugin, ItemView, Notice, PluginSettingTab, Setting, requestUrl, setIcon, TFile } = require("obsidian");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const VIEW = "receiving-dashboard-view";
const DEFAULTS = {
  pythonCmd: "python",
  poRoot: "",
  refreshMinutes: 5,
  mondayEnabled: false,
  mondayToken: "",
  mondayBoardId: "",
  autoDhr: true,
};

/* ---------------- pure helpers (unit-tested) ---------------- */
function parseQuery(q) {
  const out = { cmds: [], terms: [], not: [] };
  (q || "").trim().split(/\s+/).filter(Boolean).forEach((t) => {
    if (t.startsWith("@")) out.cmds.push(t.slice(1).toLowerCase());
    else if (t.startsWith("!") && t.length > 1) out.not.push(t.slice(1).toLowerCase());
    else out.terms.push(t.toLowerCase());
  });
  return out;
}
function textMatch(hay, q) {
  const h = hay.toLowerCase();
  return q.terms.every((t) => h.includes(t)) && !q.not.some((t) => h.includes(t));
}
function withinDays(ts, days) { return ts && Date.now() / 1000 - ts <= days * 86400; }
function dateCmdOk(ts, cmds) {
  if (cmds.includes("today") && !withinDays(ts, 1)) return false;
  if (cmds.includes("week") && !withinDays(ts, 7)) return false;
  if (cmds.includes("month") && !withinDays(ts, 31)) return false;
  return true;
}
const PO_CMDS = { filed: (p) => p.status === "filed", partial: (p) => p.status === "partial",
  progress: (p) => p.status === "progress", waiting: (p) => p.prf_waiting.length > 0,
  nomatrix: (p) => p.parts.length > 0 && !p.in_matrix, inmatrix: (p) => p.in_matrix,
  signed: (p) => p.status === "filed",
  edge: (p) => p.how.some((h) => h.includes("edge") || h.startsWith("process_")),
  receive: (p) => p.how.includes("receive.py") };
function filterPOs(pos, q) {
  const Q = parseQuery(q);
  return pos.filter((p) => {
    for (const c of Q.cmds) { if (PO_CMDS[c] && !PO_CMDS[c](p)) return false; }
    if (!dateCmdOk(p.last, Q.cmds)) return false;
    const hay = [p.po, p.vendor, p.status, p.how.join(" "), p.parts.map((x) => x.part).join(" "), p.prf_waiting.join(" ")].join(" ");
    return textMatch(hay, Q);
  });
}
const RESULT_CMDS = { proposed: "Proposed", pass: "Pass", fail: "Fail", notrun: "Not run", retest: "Re-test required", hold: "On hold" };
function filterMatrix(rows, q) {
  const Q = parseQuery(q);
  return rows.filter((r) => {
    for (const c of Q.cmds) {
      if (RESULT_CMDS[c] && r.result !== RESULT_CMDS[c]) return false;
      if (c === "must" && !r.must) return false;
    }
    return textMatch([r.id, r.scenario, r.result, r.ref].join(" "), Q);
  });
}
const AREAS = [
  ["Instruction", (p) => p === "CLAUDE.md" || p.startsWith("Project Map/01 Governance/")],
  ["Workflow", (p) => p.startsWith("Project Map/05 QE Workflow/") || p.endsWith("QE Receiving Workflow.canvas")],
  ["Code notes", (p) => p.startsWith("Project Map/02 Code/")],
  ["Validation", (p) => p.startsWith("Project Map/03 Validation/") || p.startsWith("Validation/")],
  ["History", (p) => p.startsWith("Project Map/04 History/")],
  ["Run manifests", (p) => p.startsWith("logs/")],
  ["Backups", (p) => p.startsWith("_backup/")],
  ["Edge-case scripts", (p) => !p.includes("/") && /^(process|prefill)_.*\.py$|^(find_stadiums|inspect_annots)\.py$/.test(p)],
  ["Project Map", (p) => p.startsWith("Project Map/")],
  ["Tool files", () => true],
];
function areaOf(p) { for (const [n, f] of AREAS) if (f(p)) return n; return "Tool files"; }
const HIDDEN = [/^\.obsidian\//, /^gs10\.07\.1\//, /^__pycache__\//, /^_backup\/Project Map\.pre-/, /^_backup\/canvas-graph\.pre-/, /^Project Map\/_scripts\/\./];
function hiddenPath(p) { return HIDDEN.some((r) => r.test(p)); }
const TYPE_CMDS = { notes: ["md"], canvas: ["canvas"], pdf: ["pdf"], xlsx: ["xlsx", "xls"], docx: ["docx", "doc"],
  py: ["py"], json: ["json"], images: ["png", "jpg", "jpeg", "svg", "gif"] };
const AREA_CMDS = { wi: "Instruction", workflow: "Workflow", code: "Code notes", validation: "Validation", history: "History",
  manifests: "Run manifests", backups: "Backups", scripts: "Edge-case scripts", tools: "Tool files" };
function filterFiles(files, q, area) {
  const Q = parseQuery(q);
  const types = Q.cmds.filter((c) => TYPE_CMDS[c]).flatMap((c) => TYPE_CMDS[c]);
  const areas = Q.cmds.filter((c) => AREA_CMDS[c]).map((c) => AREA_CMDS[c]);
  return files.filter((f) => {
    if (area && area !== "All" && f.area !== area) return false;
    if (types.length && !types.includes(f.ext)) return false;
    if (areas.length && !areas.includes(f.area)) return false;
    if (!dateCmdOk(f.mtime / 1000, Q.cmds)) return false;
    return textMatch(f.path, Q);
  });
}
function fmtDate(ts) { if (!ts) return ""; const d = new Date(ts * 1000); return d.toISOString().slice(0, 10); }
function fmtSize(b) { if (b == null) return ""; if (b < 1024) return b + " B"; if (b < 1048576) return (b / 1024).toFixed(0) + " KB"; return (b / 1048576).toFixed(1) + " MB"; }
function pct(n, d) { return d ? Math.round((100 * n) / d) : 0; }
function poFromText(s) { const m = /(\d{2})-(\d{4})/.exec(s || ""); return m ? `PO${m[1]}-${m[2]}` : ""; }

/* ---------------- DHR rule (pure, unit-tested) ----------------
 * A PO is ready for DHR when the latest Box scan shows every compiled part in Completed Receiving(s),
 * no PRF is waiting to be compiled, and exactly one monday item in "Product Release" matches it. */
const DHR_FROM_STATUSES = ["Working on it", "Waiting for Signature", "Ready to start", ""];
function dhrCandidates(data, items, nowMs) {
  const move = [], skipped = [];
  if (!data || !data.scan_ok) return { move, skipped, stale: true };
  if (nowMs - Date.parse(data.generated) > 30 * 60000) return { move, skipped, stale: true };
  for (const p of data.pos) {
    if (p.status !== "filed" || !p.parts.length || p.prf_waiting.length) continue;
    const m = items.filter((i) => i.po === p.po);
    if (!m.length) continue;
    const why = (r) => skipped.push({ po: p.po, item: m[0].name, reason: r });
    if (m.some((i) => i.status === "DHR" || i.group === "DHR" || /put away/i.test(i.group))) continue; // already at or past DHR
    if (m.length > 1) { why(`${m.length} monday items match this PO (shipments?) — set DHR by hand`); continue; }
    const it = m[0];
    if (/\bship/i.test(it.name)) { why("item is for one shipment of the PO — set DHR by hand"); continue; }
    if (it.group !== "Product Release") { why(`item is in "${it.group}", not Product Release`); continue; }
    if (!DHR_FROM_STATUSES.includes(it.status)) { why(`status is "${it.status}"`); continue; }
    const lastFiled = Math.max(...p.parts.map((x) => x.filed || 0)) * 1000;
    if (it.created && lastFiled < Date.parse(it.created)) { why("monday item was created after the last package was signed off"); continue; }
    move.push({ po: p.po, item: it, parts: p.parts.map((x) => x.part) });
  }
  return { move, skipped, stale: false };
}

/* ---------------- delay tracking (pure, unit-tested) ---------------- */
const DELAY_RULES = [
  ["Signature logistics", /docusign|too large|sign via|file size/i],
  ["Priority / expedite", /priorit|expedite|order today|rush/i],
  ["Missing documents", /missing|w\/?o(ut)? packing|no packing|emailing for coc|coc request|request(ing)? (the )?coc|receiving list|no coc/i],
  ["Quantity discrepancy", /\bqty\b|quantity/i],
  ["Material / finish discrepancy", /material|mat\/fin|finish|surface|roughness|passivat|anodiz/i],
  ["Dimensional out of spec", /out of spec|\boos\b|dimension|disposition/i],
  ["Testing failure / root cause", /test|broke|root cause|fail/i],
  ["Vendor document error", /traveler|vendor (form|doc)|post-steril|wrong date/i],
  ["PRF correction", /section \d|\bprf\b|lot number|correction|incomplete/i],
  ["Inspection method question", /\baql\b|inspection question|sampling|method/i],
];
function categorizeIssue(text) {
  const t = (text || "").trim(); if (!t) return "";
  for (const [label, re] of DELAY_RULES) if (re.test(t)) return label;
  return "Other";
}
const HOLD_STATUSES = ["Stuck", "Waiting for Signature"];
function delayModel(items, events, issues, nowMs) {
  // events: [{itemId, t (ms), to, from}]; issues: [{itemId, t, issue}]
  const byItem = {}; events.forEach((e) => (byItem[e.itemId] = byItem[e.itemId] || []).push(e));
  Object.values(byItem).forEach((l) => l.sort((a, b) => a.t - b.t));
  const issuesBy = {}; issues.forEach((i) => (issuesBy[i.itemId] = issuesBy[i.itemId] || []).push(i));
  const itemById = {}; items.forEach((i) => (itemById[String(i.id)] = i));
  const episodes = [];
  for (const [id, evs] of Object.entries(byItem)) {
    const it = itemById[id]; if (!it) continue;
    evs.forEach((e, k) => {
      if (!HOLD_STATUSES.includes(e.to)) return;
      const next = evs.slice(k + 1).find((x) => x.to !== e.to);
      const end = next ? next.t : nowMs;
      const iss = (issuesBy[id] || []).filter((x) => x.t >= e.t - 3600e3 && x.t <= end).map((x) => x.issue);
      let reasons = (it.delayReason || "").split(",").map((x) => x.trim()).filter(Boolean);
      let source = "monday";
      if (e.to === "Stuck" && iss.length) { const inf = [...new Set(iss.map(categorizeIssue).filter(Boolean))]; if (!reasons.length || next) { reasons = inf; source = "inferred"; } }
      if (!reasons.length) { reasons = [e.to === "Stuck" ? "Not recorded" : "Waiting for signature"]; source = "none"; }
      episodes.push({ itemId: id, po: it.po, name: it.name, status: e.to, start: e.t, end, open: !next, days: (end - e.t) / 864e5,
        reasons, source, issues: iss, waitingOn: it.waitingOn || "" });
    });
  }
  const held = items.filter((i) => HOLD_STATUSES.includes(i.status)).map((i) => {
    const ep = episodes.filter((x) => x.itemId === String(i.id) && x.open && x.status === i.status).pop();
    const lastIssue = ((issuesBy[String(i.id)] || []).slice(-1)[0] || {}).issue || i.text || "";
    const recorded = (i.delayReason || "").trim();
    return { item: i, days: ep ? ep.days : null, reason: recorded || categorizeIssue(lastIssue) || "", reasonSource: recorded ? "monday" : (lastIssue ? "inferred" : "none"),
      issue: lastIssue, blank: i.status === "Stuck" ? !recorded : !(i.waitingOn || "").trim() };
  }).sort((a, b) => (b.days || 0) - (a.days || 0));
  const causes = {};
  episodes.filter((e) => e.status === "Stuck").forEach((e) => e.reasons.forEach((r) => {
    const c = (causes[r] = causes[r] || { reason: r, count: 0, days: 0, pos: new Set() });
    c.count += 1; c.days += e.days; if (e.po) c.pos.add(e.po);
  }));
  const waits = {};
  episodes.filter((e) => e.status === "Waiting for Signature").forEach((e) => {
    const who = e.waitingOn || "Not recorded"; const w = (waits[who] = waits[who] || { who, count: 0, days: 0 }); w.count += 1; w.days += e.days; });
  return { episodes, held, causes: Object.values(causes).sort((a, b) => b.count - a.count || b.days - a.days),
    waits: Object.values(waits).sort((a, b) => b.days - a.days) };
}
function parseStatusEvents(logs) {
  const out = [];
  (logs || []).forEach((l) => { try { const d = JSON.parse(l.data);
    out.push({ itemId: String(d.pulse_id), t: Number(l.created_at) / 1e4, to: d.value && d.value.label ? d.value.label.text : "",
      from: d.previous_value && d.previous_value.label ? d.previous_value.label.text : "" }); } catch (e) {} });
  return out;
}
function parseStuckIssues(updates) {
  const out = [];
  (updates || []).forEach((u) => { const m = /STUCK\s*-[^\n]*\n+\s*Issue:\s*([^\n]*)/i.exec(u.text_body || "");
    if (m) out.push({ itemId: String(u.item_id), t: Date.parse(u.created_at), issue: m[1].replace(/\s*-\s*Auto\s*$/i, "").trim() }); });
  return out;
}

/* ---------------- plugin ---------------- */
class ReceivingDashboardPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULTS, await this.loadData());
    this.data = null; this.dataMtime = 0; this.running = false; this.lastError = ""; this.monday = null;
    this.registerView(VIEW, (leaf) => new DashboardView(leaf, this));
    this.addRibbonIcon("gauge", "Receiving Dashboard", () => this.activate());
    this.addCommand({ id: "open", name: "Open Receiving Dashboard", callback: () => this.activate() });
    this.addCommand({ id: "refresh", name: "Refresh dashboard data now", callback: () => this.runCollector(true) });
    this.addSettingTab(new DashboardSettings(this.app, this));
    this.app.workspace.onLayoutReady(() => { this.loadDataFile(true); this.startTimers(); this.runCollector(false); this.fetchMonday(); });
  }
  onunload() { this.stopTimers(); }
  base() { return this.app.vault.adapter.getBasePath(); }
  dataFile() { return path.join(this.base(), "Project Map", "_scripts", ".dashboard_data.json"); }
  script() { return path.join(this.base(), "Project Map", "_scripts", "refresh_dashboard.py"); }
  startTimers() {
    this.stopTimers();
    const mins = Math.max(1, Number(this.settings.refreshMinutes) || 5);
    this.t1 = window.setInterval(() => { this.runCollector(false); this.fetchMonday(); }, mins * 60000);
    this.t2 = window.setInterval(() => this.loadDataFile(false), 20000);
  }
  stopTimers() { if (this.t1) window.clearInterval(this.t1); if (this.t2) window.clearInterval(this.t2); }
  views() { return this.app.workspace.getLeavesOfType(VIEW).map((l) => l.view); }
  redraw() { this.views().forEach((v) => v.render()); }
  runCollector(manual) {
    if (this.running) return;
    this.running = true; this.redraw();
    const args = [this.script()];
    if (this.settings.poRoot) args.push("--po-root", this.settings.poRoot);
    execFile(this.settings.pythonCmd || "python", args, { cwd: this.base(), timeout: 15 * 60000, windowsHide: true },
      (err, stdout, stderr) => {
        this.running = false;
        if (err) {
          this.lastError = ((stderr || "") + " " + err.message).trim().slice(-400);
          if (manual) new Notice("Dashboard refresh failed — see the red banner on the dashboard.");
        } else {
          this.lastError = "";
          if (manual) new Notice("Dashboard refreshed");
        }
        this.loadDataFile(true);
      });
  }
  loadDataFile(force) {
    fs.stat(this.dataFile(), (err, st) => {
      if (err) { if (force) this.redraw(); return; }
      if (!force && st.mtimeMs === this.dataMtime) return;
      fs.readFile(this.dataFile(), "utf8", (e2, txt) => {
        if (e2) return;
        try { this.data = JSON.parse(txt); this.dataMtime = st.mtimeMs; } catch (e) { return; }
        this.redraw();
        this.autoDhr();
      });
    });
  }
  async fetchMonday() {
    const s = this.settings;
    if (!s.mondayEnabled || !s.mondayToken || !s.mondayBoardId) { this.monday = null; return; }
    const query = `query { me { name } boards(ids: [${Number(s.mondayBoardId)}]) { name url groups { id title }
      items_page(limit: 500) { items { id name url created_at updated_at group { title }
      column_values { id text column { title } } } }
      activity_logs(limit: 1000, column_ids: ["status"]) { data created_at }
      updates(limit: 300) { item_id created_at text_body } } }`;
    try {
      const r = await requestUrl({ url: "https://api.monday.com/v2", method: "POST", throw: false,
        headers: { "Content-Type": "application/json", Authorization: s.mondayToken, "API-Version": "2024-10" },
        body: JSON.stringify({ query }) });
      const b = r.json && r.json.data && r.json.data.boards && r.json.data.boards[0];
      if (!b) throw new Error((r.json && JSON.stringify(r.json.errors || r.json.error_message)) || "HTTP " + r.status);
      const items = b.items_page.items.map((it) => {
        const cols = {}; it.column_values.forEach((c) => { cols[c.column.title] = c.text || ""; if (c.column.title === "Status") this.statusColId = c.id; });
        return { id: it.id, name: it.name, url: it.url, group: it.group.title, status: cols.Status || "",
          date: cols.Date || "", person: cols.Person || "", notes: cols.Notes || "", updated: it.updated_at, created: it.created_at, po: poFromText(it.name),
          waitingOn: cols["Waiting on"] || "", delayReason: cols["Delay reason"] || "", text: cols.Text || "" };
      });
      const me = r.json.data.me ? r.json.data.me.name : "";
      this.monday = { name: b.name, url: b.url, items, error: "", fetched: new Date(), account: me,
        events: parseStatusEvents(b.activity_logs), issues: parseStuckIssues(b.updates) };
    } catch (e) { this.monday = { error: String(e.message || e), items: [] }; }
    this.redraw();
    this.autoDhr();
  }
  async mondayCall(query, variables) {
    const r = await requestUrl({ url: "https://api.monday.com/v2", method: "POST", throw: false,
      headers: { "Content-Type": "application/json", Authorization: this.settings.mondayToken, "API-Version": "2024-10" },
      body: JSON.stringify({ query, variables }) });
    if (!r.json || r.json.errors || r.json.error_message) throw new Error(JSON.stringify((r.json && (r.json.errors || r.json.error_message)) || r.status));
    return r.json.data;
  }
  dhrLogFile() { return path.join(this.base(), "Project Map", "_scripts", ".monday_dhr_log.jsonl"); }
  readDhrLog() { try { return fs.readFileSync(this.dhrLogFile(), "utf8").trim().split("\n").filter(Boolean).map((l) => JSON.parse(l)).reverse(); } catch (e) { return []; } }
  async autoDhr() {
    const s = this.settings, M = this.monday;
    if (!s.autoDhr || !s.mondayEnabled || !s.mondayToken || !M || M.error || this.dhrBusy) return;
    this.dhrDone = this.dhrDone || new Set();
    const res = dhrCandidates(this.data, M.items, Date.now());
    this.dhrState = res;
    const todo = res.move.filter((m) => !this.dhrDone.has(m.item.id));
    if (!todo.length) { this.redraw(); return; }
    this.dhrBusy = true;
    for (const m of todo) {
      const entry = { time: new Date().toISOString(), po: m.po, item: m.item.name, itemId: m.item.id, from: m.item.status, parts: m.parts, account: M.account };
      try {
        await this.mondayCall(`mutation ($b: ID!, $i: ID!, $c: String!, $v: String) { change_simple_column_value(board_id: $b, item_id: $i, column_id: $c, value: $v) { id } }`,
          { b: String(s.mondayBoardId), i: String(m.item.id), c: this.statusColId || "status", v: "DHR" });
        await this.mondayCall(`mutation ($i: ID!, $body: String!) { create_update(item_id: $i, body: $body) { id } }`,
          { i: String(m.item.id), body: `Status set to DHR automatically by the Receiving Dashboard. All ${m.parts.length} compiled part(s) for ${m.po} are signed off in Completed Receivings in Box: ${m.parts.join(", ")}. Box scan ${this.data.generated.replace("T", " ")}.` });
        entry.result = "DHR set";
        new Notice(`${m.po}: monday status set to DHR`);
      } catch (e) { entry.result = "FAILED: " + String(e.message || e).slice(0, 200); new Notice(`${m.po}: couldn't set DHR in monday`); }
      this.dhrDone.add(m.item.id);
      try { fs.appendFileSync(this.dhrLogFile(), JSON.stringify(entry) + "\n"); } catch (e) {}
    }
    this.dhrBusy = false;
    this.fetchMonday();
  }
  async activate() {
    let leaf = this.app.workspace.getLeavesOfType(VIEW)[0];
    if (!leaf) { leaf = this.app.workspace.getLeaf("tab"); await leaf.setViewState({ type: VIEW, active: true }); }
    this.app.workspace.revealLeaf(leaf);
  }
  /* opening things */
  openVault(p) {
    const f = this.app.vault.getAbstractFileByPath(p);
    if (!(f instanceof TFile)) { new Notice("Not found in vault: " + p); return; }
    if (["md", "canvas", "pdf", "png", "jpg", "jpeg", "svg", "gif"].includes(f.extension)) this.app.workspace.getLeaf("tab").openFile(f);
    else this.openExternal(path.join(this.base(), p));
  }
  openExternal(abs) { const { shell } = require("electron"); shell.openPath(abs).then((e) => { if (e) new Notice("Could not open: " + e); }); }
  reveal(abs) { const { shell } = require("electron"); shell.showItemInFolder(abs); }
  async saveSettings() { await this.saveData(this.settings); this.startTimers(); this.fetchMonday(); }
}

/* ---------------- view ---------------- */
class DashboardView extends ItemView {
  constructor(leaf, plugin) {
    super(leaf); this.plugin = plugin;
    this.tab = "overview"; this.q = { pos: "", matrix: "", files: "", monday: "" };
    this.sort = {}; this.expanded = new Set(); this.fileArea = "All"; this.fileLimit = 200;
  }
  getViewType() { return VIEW; }
  getDisplayText() { return "Receiving Dashboard"; }
  getIcon() { return "gauge"; }
  async onOpen() { this.render(); }
  go(tab, query) {
    this.tab = tab;
    if (query != null) {
      this.q[tab === "files" ? "files" : tab] = query;
      const k = /^@(\w+)$/.test(query || "") ? query.slice(1) : null;   // keep the matching tile highlighted
      if (tab === "pos") this.posTile = ["filed", "partial", "progress", "waiting", "nomatrix"].includes(k) ? k : null;
      if (tab === "matrix") this.matrixTile = ["pass", "fail", "notrun", "hold", "retest", "proposed", "must"].includes(k) ? k : null;
    }
    this.render();
  }

  render() {
    const root = this.containerEl.children[1]; root.empty(); root.addClass("rd-root");
    const P = this.plugin, D = P.data;
    const head = root.createDiv("rd-head");
    const t = head.createDiv("rd-title"); setIcon(t.createSpan("rd-logo"), "gauge"); t.createSpan({ text: "Receiving Dashboard" });
    const meta = head.createDiv("rd-meta");
    if (D) {
      meta.createSpan({ cls: "rd-pill " + (D.scan_ok ? "ok" : "bad"), text: D.scan_ok ? "Box connected" : "Box not reachable" });
      meta.createSpan({ cls: "rd-muted", text: "Updated " + D.generated.replace("T", " ") });
    }
    if (P.monday) meta.createSpan({ cls: "rd-pill " + (P.monday.error ? "bad" : "ok"), text: P.monday.error ? "monday error" : "monday connected" });
    const btn = meta.createEl("button", { cls: "rd-btn", text: P.running ? "Refreshing…" : "Refresh" });
    setIcon(btn.createSpan({ cls: "rd-btn-ico" }), "refresh-cw"); btn.disabled = P.running;
    btn.onclick = () => P.runCollector(true);
    if (P.lastError) root.createDiv({ cls: "rd-banner bad", text: "Refresh failed: " + P.lastError });
    if (D && !D.scan_ok) root.createDiv({ cls: "rd-banner warn", text: "The Box PO folder couldn't be read (" + D.roots.join(", ") + "). Check Box Drive is running, or set the PO folder in Settings → Receiving Dashboard." });

    const tabs = root.createDiv("rd-tabs");
    const list = [["overview", "Overview", "layout-dashboard"], ["pos", "POs & parts", "package"], ["workflow", "Workflow", "route"],
      ["matrix", "Test matrix", "check-square"], ["files", "Project files", "folder-tree"]];
    if (P.settings.mondayEnabled) { list.push(["delays", "Delays", "hourglass"]); list.push(["monday", "monday", "calendar-check"]); }
    list.forEach(([id, label, ico]) => {
      const b = tabs.createDiv({ cls: "rd-tab" + (this.tab === id ? " active" : "") }); setIcon(b.createSpan(), ico); b.createSpan({ text: label });
      b.onclick = () => this.go(id);
    });
    const body = root.createDiv("rd-body");
    if (!D && !["files", "delays", "monday"].includes(this.tab)) {
      body.createDiv({ cls: "rd-empty", text: P.running ? "Collecting data for the first time — this can take a few minutes while Box Drive downloads the filed packages once." : "No data yet. Click Refresh." });
      return;
    }
    ({ overview: () => this.overview(body, D), pos: () => this.posTab(body, D), workflow: () => this.workflow(body, D),
       matrix: () => this.matrixTab(body, D), files: () => this.filesTab(body), monday: () => this.mondayTab(body), delays: () => this.delaysTab(body, D) })[this.tab]();
  }

  /* ---- overview ---- */
  overview(el, D) {
    const S = D.summary;
    const g = el.createDiv("rd-kpis");
    const kpi = (label, value, sub, onclick, tone, ratio) => {
      const k = g.createDiv("rd-kpi " + (tone || "")); k.createDiv({ cls: "rd-kpi-label", text: label });
      k.createDiv({ cls: "rd-kpi-value", text: String(value) }); if (sub) k.createDiv({ cls: "rd-kpi-sub", text: sub });
      if (ratio != null) { const bar = k.createDiv("rd-meter"); bar.createDiv({ cls: "rd-meter-fill" }).style.width = Math.min(100, ratio) + "%"; }
      if (onclick) { k.addClass("click"); k.onclick = onclick; }
    };
    kpi("POs processed", S.po_processed, `of ${S.po_folders} PO folders`, () => this.go("pos", ""), "accent", pct(S.po_processed, S.po_folders));
    kpi("Parts processed", S.parts, "compiled packages", () => this.go("pos", ""));
    kpi("POs fully signed off", S.po_filed, `${pct(S.po_filed, S.po_processed)}% · ${S.po_partial} partly signed off`, () => { this.posTile = "filed"; this.go("pos", "@filed"); }, "good", pct(S.po_filed, S.po_processed));
    kpi("Parts signed off", S.parts_filed, `${pct(S.parts_filed, S.parts)}% of ${S.parts} · in Completed Receiving`, () => this.go("pos", "@filed"), "good", pct(S.parts_filed, S.parts));
    kpi("Awaiting signature", S.parts_awaiting, "parts in In-Progress Receiving", () => this.go("pos", "@progress"), S.parts_awaiting ? "warn" : "");
    kpi("PRFs not compiled", S.po_prf_waiting, "POs with a PRF but no package", () => this.go("pos", "@waiting"), S.po_prf_waiting ? "warn" : "");
    kpi("Test matrix complete", pct(S.tc_pass, S.tc_total) + "%", `${S.tc_pass}/${S.tc_total} passed`, () => this.go("matrix", ""), "accent", pct(S.tc_pass, S.tc_total));
    kpi("Must-Pass cases", `${S.must_pass}/${S.must_total}`, `${pct(S.must_pass, S.must_total)}% passed`, () => this.go("matrix", "@must"), "", pct(S.must_pass, S.must_total));
    kpi("POs not in matrix", S.not_in_matrix, "processed but unmapped", () => this.go("pos", "@nomatrix"), S.not_in_matrix ? "warn" : "");
    kpi("WI Open Items", `${S.oi_open}/${S.oi_total}`, "still open in CLAUDE.md", () => this.plugin.openVault("Project Map/03 Validation/Open Items (Master List).md"), "bad");

    const charts = el.createDiv("rd-charts");
    this.card(charts, "PO status", (c) => donut(c, [
      ["Signed off", S.po_filed, "var(--color-green)", "filed"], ["Partly signed off", S.po_partial, "var(--color-yellow)", "partial"],
      ["In progress", S.po_progress, "var(--color-orange)", "progress"]], S.po_processed, "POs",
      (k) => { this.posTile = k; this.go("pos", "@" + k); }));
    const tb = S.tc_by || {};
    this.card(charts, "Test case results", (c) => donut(c, [
      ["Pass", tb["Pass"] || 0, "var(--color-green)", "pass"], ["Fail", tb["Fail"] || 0, "var(--color-red)", "fail"],
      ["Proposed", tb["Proposed"] || 0, "var(--color-blue)", "proposed"], ["Re-test", tb["Re-test required"] || 0, "var(--color-purple)", "retest"], ["On hold", tb["On hold"] || 0, "var(--color-yellow)", "hold"],
      ["Not run", tb["Not run"] || 0, "var(--background-modifier-border)", "notrun"]], S.tc_total, "cases",
      (k) => { this.matrixTile = k; this.go("matrix", "@" + k); }));
    this.card(charts, "Weekly throughput (16 weeks)", (c) => bars(c, D.weekly), "wide");

    const lists = el.createDiv("rd-two");
    this.card(lists, "Needs attention", (c) => {
      const rows = [];
      D.pos.forEach((p) => p.parts.filter((x) => x.status === "progress").forEach((x) => rows.push([p, x.part, "Awaiting signature / filing", x.file])));
      D.pos.forEach((p) => p.prf_waiting.forEach((w) => rows.push([p, w, "PRF ready, not compiled", ""])));
      if (!rows.length) c.createDiv({ cls: "rd-muted", text: "Nothing waiting." });
      rows.slice(0, 25).forEach(([p, part, what, file]) => {
        const r = c.createDiv("rd-row");
        r.createSpan({ cls: "rd-mono", text: p.po }); r.createSpan({ text: `${p.vendor} · ${part}` }); r.createSpan({ cls: "rd-chip warn", text: what });
        const a = r.createSpan("rd-actions");
        if (file) this.iconBtn(a, "file-text", "Open package", () => this.plugin.openExternal(file));
        this.iconBtn(a, "folder-open", "Open PO folder", () => this.plugin.openExternal(p.path));
      });
    });
    this.card(lists, "Recent activity", (c) => {
      D.recent.slice(0, 15).forEach((x) => {
        const p = D.pos.find((y) => y.po === x.po);
        const r = c.createDiv("rd-row");
        r.createSpan({ cls: "rd-mono rd-muted", text: fmtDate(x.mtime) }); r.createSpan({ cls: "rd-mono", text: x.po });
        r.createSpan({ text: x.part }); this.statusChip(r, x.status);
        const part = p && p.parts.find((y) => y.part === x.part);
        if (part && part.file) { const a = r.createSpan("rd-actions"); this.iconBtn(a, "file-text", "Open package", () => this.plugin.openExternal(part.file)); }
      });
    });

    this.card(el, "Jump to", (c) => {
      const links = c.createDiv("rd-links");
      const L = (label, ico, fn) => { const b = links.createDiv("rd-link"); setIcon(b.createSpan(), ico); b.createSpan({ text: label }); b.onclick = fn; };
      L("WI-7.4.3-01 (CLAUDE.md)", "book-open", () => this.plugin.openVault("CLAUDE.md"));
      L("Project Index", "list", () => this.plugin.openVault("Project Map/00 Index.md"));
      L("Project Map board", "layout-grid", () => this.plugin.openVault("Project Map/Project Map.canvas"));
      L("QE workflow board", "route", () => this.plugin.openVault("Project Map/QE Receiving Workflow.canvas"));
      L("Test Case Matrix (Word)", "file-check", () => this.plugin.openVault("Validation/QF_7.4.3-06_Attachment_A_Test_Case_Matrix.docx"));
      L("Validation Protocol (Word)", "file-check", () => this.plugin.openVault("Validation/QF_7.4.3-06_Validation_Protocol.docx"));
      L("Open Items", "alert-triangle", () => this.plugin.openVault("Project Map/03 Validation/Open Items (Master List).md"));
      L("Reasoning Reference", "message-square", () => this.plugin.openVault("Project Map/01 Governance/Comment & Discrepancy Reasoning Reference.md"));
      if (D.roots && D.roots[0]) L("Box PO year folder", "folder-open", () => this.plugin.openExternal(D.roots[0]));
      L("Tools folder", "folder-cog", () => this.plugin.openExternal(this.plugin.base()));
    }, "wide");
  }

  /* ---- POs ---- */
  posTab(el, D) {
    const pos = D.pos;
    const S = D.summary, PF = this.posTile || null;
    const tileQ = { filed: "@filed", partial: "@partial", progress: "@progress", waiting: "@waiting", nomatrix: "@nomatrix" };
    this.tiles(el, [
      { key: "filed", label: "Signed off", value: S.po_filed, sub: "all parts in Completed", tone: "good" },
      { key: "partial", label: "Partly signed off", value: S.po_partial, tone: S.po_partial ? "warn" : "" },
      { key: "progress", label: "In progress", value: S.po_progress, sub: "nothing signed off yet", tone: S.po_progress ? "warn" : "" },
      { key: "waiting", label: "PRFs not compiled", value: S.po_prf_waiting, tone: S.po_prf_waiting ? "warn" : "" },
      { key: "nomatrix", label: "Not in test matrix", value: S.not_in_matrix },
    ], PF, (k) => { this.posTile = k; this.q.pos = k ? tileQ[k] : ""; this.render(); }, true);
    this.search(el, "pos", "Search POs, vendors, parts…   @filed @partial @progress @waiting @nomatrix @edge @receive @week @month  !exclude");
    const mon = this.plugin.monday && !this.plugin.monday.error ? this.plugin.monday.items : null;
    const rows = filterPOs(pos, this.q.pos).map((p) => {
      const m = mon && mon.find((i) => i.po === p.po);
      return Object.assign({}, p, { _parts: `${p.filed}/${p.parts.length}`, _mon: m ? m.status || m.group : "" });
    });
    const cols = [["po", "PO"], ["vendor", "Vendor"], ["status", "Status"], ["_parts", "Signed off / parts"],
      ["in_matrix", "In matrix"], ["how", "Ran via"]];
    if (mon) cols.push(["_mon", "monday"]);
    cols.push(["last", "Last activity"], ["", ""]);
    el.createDiv({ cls: "rd-muted rd-count", text: `${rows.length} of ${pos.length} POs` });
    this.table(el, "pos", cols, rows, (tr, p) => {
      tr.createEl("td", { cls: "rd-mono", text: p.po }); tr.createEl("td", { text: p.vendor });
      this.statusChip(tr.createEl("td"), p.status); tr.createEl("td", { cls: "rd-num", text: p._parts });
      tr.createEl("td", { text: p.parts.length ? (p.in_matrix ? "✓" : "—") : "" });
      tr.createEl("td", { cls: "rd-small", text: p.how.join(", ") });
      if (mon) tr.createEl("td", { cls: "rd-small", text: p._mon });
      tr.createEl("td", { cls: "rd-mono rd-muted", text: fmtDate(p.last) });
      const a = tr.createEl("td", { cls: "rd-actions" });
      this.iconBtn(a, "folder-open", "Open PO folder", (e) => { e.stopPropagation(); this.plugin.openExternal(p.path); });
      tr.addClass("click");
      tr.onclick = () => { this.expanded.has(p.po) ? this.expanded.delete(p.po) : this.expanded.add(p.po); this.render(); };
      if (this.expanded.has(p.po)) {
        const ex = tr.parentElement.createEl("tr", { cls: "rd-expand" }).createEl("td", { attr: { colspan: String(cols.length) } });
        p.parts.forEach((x) => {
          const r = ex.createDiv("rd-row");
          r.createSpan({ cls: "rd-mono", text: x.part }); this.statusChip(r, x.status);
          r.createSpan({ cls: "rd-muted rd-small", text: x.filed ? "signed off " + fmtDate(x.filed) : "compiled " + fmtDate(x.compiled) });
          const ac = r.createSpan("rd-actions");
          if (x.file) { this.textBtn(ac, "file-text", "Package", () => this.plugin.openExternal(x.file)); this.textBtn(ac, "search", "Show", () => this.plugin.reveal(x.file)); }
          if (x.prf) this.textBtn(ac, "sheet", "PRF", () => this.plugin.openExternal(x.prf));
        });
        p.prf_waiting.forEach((w) => { const r = ex.createDiv("rd-row"); r.createSpan({ cls: "rd-mono", text: w }); r.createSpan({ cls: "rd-chip warn", text: "PRF ready, not compiled" }); });
        if (p.prf_files.length) {
          const r = ex.createDiv("rd-row rd-small"); r.createSpan({ cls: "rd-muted", text: "Inspection Forms:" });
          p.prf_files.forEach((f) => this.textBtn(r, f.toLowerCase().endsWith(".pdf") ? "file-text" : "sheet", path.basename(f), () => this.plugin.openExternal(f)));
        }
        const m = mon && mon.find((i) => i.po === p.po);
        if (m) { const r = ex.createDiv("rd-row rd-small"); r.createSpan({ cls: "rd-muted", text: "monday:" }); r.createSpan({ text: `${m.name} · ${m.group} · ${m.status}` }); this.textBtn(r, "external-link", "Open", () => window.open(m.url)); }
      }
    }, (p, k) => (k === "how" ? p.how.join(",") : k === "in_matrix" ? (p.in_matrix ? 1 : 0) : p[k]));
  }

  /* ---- workflow ---- */
  workflow(el, D) {
    const S = D.summary;
    const counts = { W04: [S.po_prf_waiting, "POs with PRFs not compiled", "@waiting"], W06: [S.parts, "parts compiled", ""],
      W07: [S.parts_awaiting, "awaiting signature", "@progress"], W08: [S.parts_awaiting, "awaiting signature", "@progress"],
      W09: [S.parts_filed, "parts signed off", "@filed"] };
    const files = this.app.vault.getMarkdownFiles().filter((f) => f.path.startsWith("Project Map/05 QE Workflow/W")).sort((a, b) => a.path.localeCompare(b.path));
    const grid = el.createDiv("rd-steps");
    files.forEach((f, i) => {
      const id = f.basename.slice(0, 3); const title = f.basename.slice(6);
      const card = grid.createDiv("rd-step");
      card.createDiv({ cls: "rd-step-n", text: id });
      card.createDiv({ cls: "rd-step-t", text: title });
      const c = counts[id];
      if (c) { const b = card.createDiv({ cls: "rd-step-count click", text: `${c[0]} ${c[1]}` }); b.onclick = (e) => { e.stopPropagation(); this.go("pos", c[2]); }; }
      const cache = this.app.metadataCache.getFileCache(f);
      const links = (cache && cache.links ? cache.links : []).map((l) => l.link).filter((l) => /^Step |Reasoning|Edge-Case|File Access|Folder Structure|Manual vs/.test(l));
      const chips = card.createDiv("rd-step-chips");
      links.forEach((l) => { const ch = chips.createSpan({ cls: "rd-chip code click", text: l.replace(/ - .*$/, "").replace("Comment & Discrepancy ", "").replace("Manual vs Claude-Assisted Workflow", "Manual vs Claude") }); ch.title = l;
        ch.onclick = (e) => { e.stopPropagation(); const t = this.app.metadataCache.getFirstLinkpathDest(l, f.path); if (t) this.plugin.openVault(t.path); }; });
      if (!links.length) chips.createSpan({ cls: "rd-muted rd-small", text: "physical / manual" });
      card.addClass("click"); card.onclick = () => this.plugin.openVault(f.path);
      if (i < files.length - 1) grid.createDiv({ cls: "rd-step-arrow", text: "→" });
    });
    el.createDiv({ cls: "rd-muted rd-small", text: "Click a step to open its note (editable). Purple chips are the automation steps embedded in it; counts jump to the matching POs." });
  }

  /* ---- matrix ---- */
  matrixTab(el, D) {
    const bar = el.createDiv("rd-toolbar");
    this.textBtn(bar, "file-check", "Open Test Case Matrix (Word)", () => this.plugin.openVault("Validation/QF_7.4.3-06_Attachment_A_Test_Case_Matrix.docx"));
    this.textBtn(bar, "file-text", "Matrix notes", () => this.plugin.openVault("Project Map/03 Validation/Attachment A - Test Case Matrix.md"));
    if (D.matrix_error) el.createDiv({ cls: "rd-banner warn", text: D.matrix_error });
    const by = {}; D.matrix.forEach((r) => (by[r.result] = (by[r.result] || 0) + 1));
    const mustN = D.matrix.filter((r) => r.must).length, MF = this.matrixTile || null;
    const mq = { pass: "@pass", fail: "@fail", notrun: "@notrun", hold: "@hold", retest: "@retest", proposed: "@proposed", must: "@must" };
    const mdefs = [{ key: "pass", label: "Pass", value: by["Pass"] || 0, tone: "good" }, { key: "fail", label: "Fail", value: by["Fail"] || 0, tone: (by["Fail"] ? "bad" : "") },
      { key: "notrun", label: "Not run", value: by["Not run"] || 0 }, { key: "hold", label: "On hold", value: by["On hold"] || 0, tone: by["On hold"] ? "warn" : "" }];
    if (by["Re-test required"]) mdefs.push({ key: "retest", label: "Re-test", value: by["Re-test required"], tone: "warn" });
    if (by["Proposed"]) mdefs.push({ key: "proposed", label: "Proposed", value: by["Proposed"], tone: "accent" });
    mdefs.push({ key: "must", label: "Must-Pass cases", value: mustN, tone: "accent" });
    this.tiles(el, mdefs, MF, (k) => { this.matrixTile = k; this.q.matrix = k ? mq[k] : ""; this.render(); }, true);
    this.search(el, "matrix", "Search test cases…   @pass @proposed @fail @notrun @retest @hold @must");
    const rows = filterMatrix(D.matrix, this.q.matrix);
    el.createDiv({ cls: "rd-muted rd-count", text: `${rows.length} of ${D.matrix.length} test cases` });
    this.table(el, "matrix", [["id", "TC"], ["scenario", "Scenario"], ["must", "Must-pass"], ["result", "Result"], ["ref", "PO ref"]], rows, (tr, r) => {
      tr.createEl("td", { cls: "rd-mono", text: r.id }); tr.createEl("td", { text: r.scenario });
      tr.createEl("td", { text: r.must ? "●" : "" }); this.statusChip(tr.createEl("td"), r.result);
      const td = tr.createEl("td", { cls: "rd-small" });
      (r.ref.match(/\d{2}-\d{4}/g) || []).forEach((n) => { const po = "PO" + n; const ch = td.createSpan({ cls: "rd-chip code click", text: po }); ch.onclick = () => this.go("pos", po); });
    });
  }

  /* ---- files navigator ---- */
  filesTab(el) {
    const all = this.app.vault.getFiles().filter((f) => !hiddenPath(f.path)).map((f) => ({ file: f, path: f.path, name: f.name,
      ext: f.extension.toLowerCase(), area: areaOf(f.path), size: f.stat.size, mtime: f.stat.mtime }));
    const areas = ["All", ...AREAS.map((a) => a[0])];
    const chips = el.createDiv("rd-areas");
    areas.forEach((a) => {
      const n = a === "All" ? all.length : all.filter((f) => f.area === a).length;
      const c = chips.createSpan({ cls: "rd-area" + (this.fileArea === a ? " active" : ""), text: `${a} ${n}` });
      c.onclick = () => { this.fileArea = a; this.fileLimit = 200; this.render(); };
    });
    this.search(el, "files", "Search every project file…   @notes @canvas @pdf @xlsx @docx @py @json @images  @wi @workflow @code @validation @manifests @backups @scripts  @week @month");
    const rows = filterFiles(all, this.q.files, this.fileArea);
    el.createDiv({ cls: "rd-muted rd-count", text: `${rows.length} files` + (rows.length > this.fileLimit ? ` · showing ${this.fileLimit}` : "") });
    this.table(el, "files", [["name", "Name"], ["area", "Area"], ["ext", "Type"], ["size", "Size"], ["mtime", "Modified"], ["", ""]], rows, (tr, f) => {
      const n = tr.createEl("td"); n.createDiv({ text: f.name }); n.createDiv({ cls: "rd-muted rd-small", text: f.path.includes("/") ? path.dirname(f.path) : "" });
      tr.createEl("td", { cls: "rd-small", text: f.area }); tr.createEl("td", { cls: "rd-chip-cell" }).createSpan({ cls: "rd-chip ext-" + f.ext, text: f.ext });
      tr.createEl("td", { cls: "rd-num rd-muted", text: fmtSize(f.size) }); tr.createEl("td", { cls: "rd-mono rd-muted", text: fmtDate(f.mtime / 1000) });
      const a = tr.createEl("td", { cls: "rd-actions" });
      this.iconBtn(a, "external-link", "Open with default app", (e) => { e.stopPropagation(); this.plugin.openExternal(path.join(this.plugin.base(), f.path)); });
      this.iconBtn(a, "search", "Show in folder", (e) => { e.stopPropagation(); this.plugin.reveal(path.join(this.plugin.base(), f.path)); });
      tr.addClass("click"); tr.onclick = () => this.plugin.openVault(f.path);
    }, null, this.fileLimit, () => { this.fileLimit += 300; this.render(); }, ["mtime", -1]);
  }

  /* ---- delays ---- */
  playbook() {
    const f = this.app.vault.getAbstractFileByPath("Project Map/04 History/Delay Playbook.md");
    const out = {}; if (!f) return out;
    try {
      const txt = require("fs").readFileSync(path.join(this.plugin.base(), f.path), "utf8");
      txt.split(/\n## /).slice(1).forEach((sec) => { const [h, ...rest] = sec.split("\n"); out[h.trim()] = rest.join("\n").trim().split(/\n\s*\n/)[0]; });
    } catch (e) {}
    return out;
  }
  delaysTab(el, D) {
    const M = this.plugin.monday;
    if (!M) { el.createDiv({ cls: "rd-empty", text: "Add a monday API token in Settings → Receiving Dashboard to track delays." }); return; }
    if (M.error) { el.createDiv({ cls: "rd-banner bad", text: "monday: " + M.error }); return; }
    const model = delayModel(M.items, M.events || [], M.issues || [], Date.now());
    const pb = this.playbook();
    const stuck = model.held.filter((h) => h.item.status === "Stuck"), wfs = model.held.filter((h) => h.item.status === "Waiting for Signature");
    const blanks = model.held.filter((h) => h.blank);
    const avg = (l) => { const d = l.map((h) => h.days).filter((x) => x != null); return d.length ? (d.reduce((a, b) => a + b, 0) / d.length).toFixed(1) : "—"; };
    const top = model.causes.find((c) => !["Not recorded", "Other"].includes(c.reason));
    const F = this.delayFilter || null;
    const pick = (v) => { this.delayFilter = v; this.render(); };
    this.tiles(el, [
      { key: "stuck", label: "Stuck now", value: stuck.length, sub: `avg ${avg(stuck)} days so far`, tone: stuck.length ? "bad" : "" },
      { key: "wfs", label: "Waiting for signature", value: wfs.length, sub: `avg ${avg(wfs)} days so far`, tone: wfs.length ? "warn" : "" },
      { key: "blank", label: "Missing reason / who", value: blanks.length, sub: "fill in on the monday item", tone: blanks.length ? "warn" : "good" },
      { key: top ? "cause:" + top.reason : "cause:", label: "Top cause", value: top ? top.count : 0, sub: top ? top.reason : "no history yet", tone: "accent" },
    ], F, pick);
    const causeSel = F && F.startsWith("cause:") ? F.slice(6) : null;
    const heldShown = model.held.filter((h) => !F ? true : F === "stuck" ? h.item.status === "Stuck" : F === "wfs" ? h.item.status === "Waiting for Signature"
      : F === "blank" ? h.blank : causeSel ? (h.reason || "").split(",").map((x) => x.trim()).includes(causeSel) : true);
    if (F) this.filterBar(el, F === "stuck" ? "Stuck now" : F === "wfs" ? "Waiting for signature" : F === "blank" ? "Missing reason / who" : "Cause: " + causeSel, () => pick(null));
    this.card(el, F ? `Held right now — ${heldShown.length} of ${model.held.length}` : "Held right now", (c) => {
      if (!heldShown.length) { c.createDiv({ cls: "rd-muted", text: F ? "No PO currently held matches this filter." : "Nothing is Stuck or Waiting for Signature." }); }
      const box = (po) => (D && D.pos ? D.pos.find((p) => p.po === po) : null);
      heldShown.forEach((h) => {
        const r = c.createDiv("rd-row");
        r.createSpan({ cls: "rd-mono", text: h.item.po || "—" }); r.createSpan({ text: h.item.name.replace(/^PO:?\s*/i, "") });
        r.createSpan({ cls: "rd-chip " + (h.item.status === "Stuck" ? "bad" : "purple"), text: h.item.status });
        r.createSpan({ cls: "rd-chip", text: "waiting on: " + (h.item.waitingOn || "?") });
        r.createSpan({ cls: "rd-chip " + (h.reasonSource === "monday" ? "code" : "warn"), text: (h.reason || "no reason") + (h.reasonSource === "inferred" ? " (inferred)" : "") });
        if (h.days != null) r.createSpan({ cls: "rd-mono rd-muted", text: h.days.toFixed(1) + " d" });
        const b = h.item.po && box(h.item.po); if (b) r.createSpan({ cls: "rd-small rd-muted", text: `Box: ${b.filed}/${b.parts.length} signed off` });
        const a = r.createSpan("rd-actions"); this.iconBtn(a, "external-link", "Open in monday", () => window.open(h.item.url));
        if (b) this.iconBtn(a, "folder-open", "Open PO folder", () => this.plugin.openExternal(b.path));
        if (h.issue) r.createDiv({ cls: "rd-small rd-muted", text: "Issue: " + h.issue }).style.flexBasis = "100%";
        const tip = pb[h.reason]; if (tip && h.item.status === "Stuck") r.createDiv({ cls: "rd-small", text: "Fix: " + tip }).style.flexBasis = "100%";
      });
    }, "wide");

    if (causeSel) this.card(el, `Every hold for "${causeSel}"`, (c) => {
      const eps = model.episodes.filter((e) => e.status === "Stuck" && e.reasons.includes(causeSel)).sort((a, b) => b.start - a.start);
      if (!eps.length) { c.createDiv({ cls: "rd-muted", text: "No past holds recorded for this reason." }); return; }
      eps.forEach((e) => { const r = c.createDiv("rd-row");
        r.createSpan({ cls: "rd-mono rd-muted", text: new Date(e.start).toISOString().slice(0, 10) }); r.createSpan({ cls: "rd-mono", text: e.po || "—" });
        r.createSpan({ text: e.name.replace(/^PO:?\s*/i, "") }); r.createSpan({ cls: "rd-mono rd-muted", text: e.days.toFixed(1) + " d" + (e.open ? " (still held)" : "") });
        if (e.issues.length) r.createDiv({ cls: "rd-small rd-muted", text: "Issue: " + e.issues.join(" · ") }).style.flexBasis = "100%"; });
    }, "wide");
    const two = el.createDiv("rd-two");
    this.card(two, "What causes delays (Stuck history) — click a reason", (c) => {
      if (!model.causes.length) { c.createDiv({ cls: "rd-muted", text: "No Stuck history in the last 1,000 status changes." }); return; }
      const max = Math.max(...model.causes.map((x) => x.count));
      model.causes.forEach((x) => {
        const r = c.createDiv("rd-cause click" + (causeSel === x.reason ? " active" : ""));
        r.onclick = () => pick(causeSel === x.reason ? null : "cause:" + x.reason);
        const head = r.createDiv("rd-row"); head.createSpan({ cls: "rd-strong", text: x.reason });
        head.createSpan({ cls: "rd-muted rd-small", text: `${x.count}× · ${(x.days / x.count).toFixed(1)} d avg · ${[...x.pos].join(", ")}` });
        r.createDiv("rd-meter").createDiv({ cls: "rd-meter-fill" }).style.width = (100 * x.count / max) + "%";
        if (pb[x.reason]) r.createDiv({ cls: "rd-small rd-muted", text: pb[x.reason] });
      });
      c.createDiv({ cls: "rd-small rd-muted", text: "Reasons come from the monday Delay reason column; older holds are inferred from the STUCK \"Issue:\" emails. Edit the fixes in Delay Playbook." });
    });
    this.card(two, "Waiting for signature — by person", (c) => {
      if (!model.waits.length) { c.createDiv({ cls: "rd-muted", text: "No signature waits recorded yet." }); return; }
      model.waits.forEach((w) => { const r = c.createDiv("rd-row"); r.createSpan({ cls: "rd-strong", text: w.who });
        r.createSpan({ cls: "rd-muted rd-small", text: `${w.count} wait(s) · ${(w.days / w.count).toFixed(1)} d avg · ${w.days.toFixed(1)} d total` }); });
      c.createDiv({ cls: "rd-small rd-muted", text: "\"Not recorded\" = Waiting on wasn't set. Fill it in on the item to see who signatures wait on." });
    });
    const t = el.createDiv("rd-toolbar");
    this.textBtn(t, "file-text", "Edit Delay Playbook", () => this.plugin.openVault("Project Map/04 History/Delay Playbook.md"));
    this.textBtn(t, "external-link", "Open monday board", () => window.open(M.url));
  }

  /* ---- monday ---- */
  mondayTab(el) {
    const M = this.plugin.monday;
    if (!M) { el.createDiv({ cls: "rd-empty", text: "Add a monday API token and board ID in Settings → Receiving Dashboard." }); return; }
    if (M.error) { el.createDiv({ cls: "rd-banner bad", text: "monday: " + M.error }); return; }
    const auto = this.card(el, "DHR automation", (c) => {
      const st = this.plugin.dhrState, s = this.plugin.settings;
      const head = c.createDiv("rd-row");
      head.createSpan({ cls: "rd-chip " + (s.autoDhr ? "good" : ""), text: s.autoDhr ? "On — fully automatic" : "Off" });
      head.createSpan({ cls: "rd-muted rd-small", text: `Sets Status = DHR when every compiled part of a PO is in Completed Receivings (your monday automation then moves the group). Posts as: ${M.account || "?"}` });
      if (st && st.stale) c.createDiv({ cls: "rd-muted rd-small", text: "Waiting for a fresh Box scan (needs one from the last 30 minutes)." });
      if (st && st.skipped.length) {
        c.createDiv({ cls: "rd-small rd-strong", text: "Signed off in Box but not changed automatically:" });
        st.skipped.forEach((k) => { const r = c.createDiv("rd-row rd-small"); r.createSpan({ cls: "rd-mono", text: k.po }); r.createSpan({ text: k.item }); r.createSpan({ cls: "rd-chip warn", text: k.reason }); });
      }
      const log = this.plugin.readDhrLog().slice(0, 15);
      c.createDiv({ cls: "rd-small rd-strong", text: log.length ? "Recent changes:" : "No changes made yet." });
      log.forEach((e) => { const r = c.createDiv("rd-row rd-small"); r.createSpan({ cls: "rd-mono rd-muted", text: e.time.slice(0, 16).replace("T", " ") });
        r.createSpan({ cls: "rd-mono", text: e.po }); r.createSpan({ text: `${e.from || "—"} → DHR` }); r.createSpan({ cls: "rd-chip " + (e.result === "DHR set" ? "good" : "bad"), text: e.result }); });
    });
    const bar = el.createDiv("rd-toolbar"); bar.createSpan({ cls: "rd-strong", text: M.name });
    this.textBtn(bar, "external-link", "Open board", () => window.open(M.url));
    const groups = {}; M.items.forEach((i) => { groups[i.group] = (groups[i.group] || 0) + 1; });
    const GF = this.mondayGroup || null;
    this.tiles(el, Object.entries(groups).map(([k, v]) => ({ key: k, label: k, value: v, tone: k === "DHR" ? "good" : k === "MRB" ? "bad" : "" })),
      GF, (k) => { this.mondayGroup = k; this.render(); }, true);
    if (GF) this.filterBar(el, "Group: " + GF, () => { this.mondayGroup = null; this.render(); });
    this.search(el, "monday", "Search monday items…");
    const Q = parseQuery(this.q.monday);
    const rows = M.items.filter((i) => (!GF || i.group === GF) && textMatch([i.name, i.group, i.status, i.person, i.notes].join(" "), Q));
    const known = new Set(((this.plugin.data && this.plugin.data.pos) || []).map((p) => p.po));
    this.table(el, "monday", [["name", "Item"], ["group", "Group"], ["status", "Status"], ["person", "Person"], ["date", "Date"], ["po", "PO"], ["", ""]], rows, (tr, i) => {
      tr.createEl("td", { text: i.name }); tr.createEl("td", { cls: "rd-small", text: i.group }); tr.createEl("td").createSpan({ cls: "rd-chip", text: i.status || "—" });
      tr.createEl("td", { cls: "rd-small", text: i.person }); tr.createEl("td", { cls: "rd-mono rd-muted", text: i.date });
      const td = tr.createEl("td"); if (i.po) { const ch = td.createSpan({ cls: "rd-chip code" + (known.has(i.po) ? " click" : ""), text: i.po }); if (known.has(i.po)) ch.onclick = () => this.go("pos", i.po); }
      const a = tr.createEl("td", { cls: "rd-actions" }); this.iconBtn(a, "external-link", "Open in monday", () => window.open(i.url));
    });
  }

  /* ---- widgets ---- */
  tiles(el, defs, active, onPick, small) {
    // defs: [{key,label,value,sub,tone}] — clicking a tile filters; clicking the active tile clears.
    const g = el.createDiv("rd-kpis" + (small ? " small" : ""));
    defs.forEach((d) => {
      const k = g.createDiv("rd-kpi click " + (d.tone || "") + (active === d.key ? " active" : ""));
      k.createDiv({ cls: "rd-kpi-label", text: d.label }); k.createDiv({ cls: "rd-kpi-value", text: String(d.value) });
      if (d.sub) k.createDiv({ cls: "rd-kpi-sub", text: d.sub });
      k.title = active === d.key ? "Click to show everything again" : "Click to show only these";
      k.onclick = () => onPick(active === d.key ? null : d.key);
    });
    return g;
  }
  filterBar(el, text, onClear) {
    const b = el.createDiv("rd-filterbar"); b.createSpan({ text: "Showing: " }); b.createSpan({ cls: "rd-strong", text });
    const x = b.createSpan({ cls: "rd-tbtn", text: "Show all" }); x.onclick = onClear;
  }
  card(parent, title, fill, cls) { const c = parent.createDiv("rd-card " + (cls || "")); c.createDiv({ cls: "rd-card-title", text: title }); fill(c.createDiv("rd-card-body")); return c; }
  statusChip(el, s) {
    const map = { filed: ["Signed off", "good"], partial: ["Partly signed off", "warn"], progress: ["In progress", "warn"], waiting: ["PRF waiting", "warn"],
      archived: ["Archived copy", ""], Pass: ["Pass", "good"], Fail: ["Fail", "bad"], "Re-test required": ["Re-test", "purple"], Proposed: ["Proposed", "purple"], "On hold": ["On hold", "warn"], "Not run": ["Not run", ""] };
    const [t, c] = map[s] || [s, ""];
    el.createSpan({ cls: "rd-chip " + c, text: t });
  }
  iconBtn(el, ico, title, fn) { const b = el.createSpan({ cls: "rd-ibtn" }); b.title = title; setIcon(b, ico); b.onclick = fn; return b; }
  textBtn(el, ico, label, fn) { const b = el.createSpan({ cls: "rd-tbtn" }); setIcon(b.createSpan(), ico); b.createSpan({ text: label }); b.onclick = (e) => { e.stopPropagation(); fn(e); }; return b; }
  search(el, key, ph) {
    const w = el.createDiv("rd-search"); setIcon(w.createSpan("rd-search-ico"), "search");
    const inp = w.createEl("input", { type: "text", placeholder: ph }); inp.value = this.q[key] || "";
    let t; inp.oninput = () => { window.clearTimeout(t); t = window.setTimeout(() => { this.q[key] = inp.value;
      if (key === "pos") this.posTile = null; if (key === "matrix") this.matrixTile = null; this.render(); this.refocus(key); }, 200); };
    inp.dataset.key = key;
    if (this._focus === key) { window.setTimeout(() => { inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length); }, 0); this._focus = null; }
  }
  refocus(key) { this._focus = key; const inp = this.containerEl.querySelector(`input[data-key="${key}"]`); if (inp) { inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length); } }
  table(el, key, cols, rows, drawRow, getter, limit, more, defSort) {
    const st = this.sort[key] || (defSort ? { col: defSort[0], dir: defSort[1] } : null);
    if (st) {
      const g = getter || ((r, k) => r[k]);
      rows = rows.slice().sort((a, b) => { const x = g(a, st.col), y = g(b, st.col); return (x > y ? 1 : x < y ? -1 : 0) * st.dir; });
    }
    const wrap = el.createDiv("rd-table-wrap"); const t = wrap.createEl("table", { cls: "rd-table" });
    const hr = t.createEl("thead").createEl("tr");
    cols.forEach(([k, label]) => {
      const th = hr.createEl("th", { text: label + (st && st.col === k ? (st.dir > 0 ? " ▲" : " ▼") : "") });
      if (k) th.onclick = () => { const cur = this.sort[key] || st; this.sort[key] = { col: k, dir: cur && cur.col === k ? -cur.dir : 1 }; this.render(); };
    });
    const tb = t.createEl("tbody");
    (limit ? rows.slice(0, limit) : rows).forEach((r) => drawRow(tb.createEl("tr"), r));
    if (limit && rows.length > limit && more) { const m = wrap.createDiv({ cls: "rd-more click", text: `Show more (${rows.length - limit} hidden)` }); m.onclick = more; }
    if (!rows.length) wrap.createDiv({ cls: "rd-empty", text: "Nothing matches." });
  }
}

/* ---------------- SVG charts ---------------- */
const NS = "http://www.w3.org/2000/svg";
function svgEl(tag, attrs, parent) { const e = document.createElementNS(NS, tag); Object.entries(attrs).forEach(([k, v]) => e.setAttribute(k, v)); if (parent) parent.appendChild(e); return e; }
function donut(el, segs, total, unit, onPick) {
  const wrap = el.createDiv("rd-donut");
  const svg = svgEl("svg", { viewBox: "0 0 120 120", width: "140", height: "140" }); wrap.appendChild(svg);
  const R = 46, C = 2 * Math.PI * R; let off = 0;
  svgEl("circle", { cx: 60, cy: 60, r: R, fill: "none", "stroke-width": 14, style: "stroke: var(--background-modifier-border)" }, svg);
  const sum = segs.reduce((a, s) => a + s[1], 0) || 1;
  segs.forEach(([label, v, color]) => {
    if (!v) return; const len = (v / sum) * C;
    const c = svgEl("circle", { cx: 60, cy: 60, r: R, fill: "none", "stroke-width": 14, "stroke-dasharray": `${len} ${C - len}`,
      "stroke-dashoffset": String(-off), transform: "rotate(-90 60 60)", style: `stroke:${color}` }, svg);
    svgEl("title", {}, c).textContent = `${label}: ${v}`; off += len;
  });
  const tx = svgEl("text", { x: 60, y: 58, "text-anchor": "middle", class: "rd-donut-num" }, svg); tx.textContent = String(total);
  const tu = svgEl("text", { x: 60, y: 76, "text-anchor": "middle", class: "rd-donut-unit" }, svg); tu.textContent = unit;
  const leg = wrap.createDiv("rd-legend");
  segs.forEach(([label, v, color, key]) => { const r = leg.createDiv("rd-leg"); r.createSpan({ cls: "rd-dot" }).style.background = color; r.createSpan({ text: label }); r.createSpan({ cls: "rd-num", text: String(v) });
    if (onPick && key) { r.addClass("click"); r.title = "Show only these"; r.onclick = () => onPick(key); } });
}
function bars(el, weeks) {
  const W = 560, H = 150, pad = 22, n = weeks.length || 1, bw = (W - pad * 2) / n;
  const max = Math.max(1, ...weeks.map((w) => Math.max(w.compiled, w.filed)));
  const svg = svgEl("svg", { viewBox: `0 0 ${W} ${H + 24}`, class: "rd-bars", preserveAspectRatio: "none" }); el.appendChild(svg);
  weeks.forEach((w, i) => {
    const x = pad + i * bw;
    [["compiled", "var(--interactive-accent)", 0], ["filed", "var(--color-green)", 1]].forEach(([k, col, j]) => {
      const h = (w[k] / max) * H; const r = svgEl("rect", { x: x + 3 + j * (bw / 2 - 3), y: H - h, width: Math.max(2, bw / 2 - 4), height: h, rx: 2, style: `fill:${col}` }, svg);
      svgEl("title", {}, r).textContent = `Week of ${w.week}: ${w[k]} ${k}`;
    });
    if (i % 2 === 0) { const t = svgEl("text", { x: x + bw / 2, y: H + 16, "text-anchor": "middle", class: "rd-axis" }, svg); t.textContent = w.week.slice(5); }
  });
  const leg = el.createDiv("rd-legend row");
  [["Compiled", "var(--interactive-accent)"], ["Filed", "var(--color-green)"]].forEach(([l, c]) => { const r = leg.createDiv("rd-leg"); r.createSpan({ cls: "rd-dot" }).style.background = c; r.createSpan({ text: l }); });
}

/* ---------------- settings ---------------- */
class DashboardSettings extends PluginSettingTab {
  constructor(app, plugin) { super(app, plugin); this.plugin = plugin; }
  display() {
    const el = this.containerEl; el.empty(); const s = this.plugin.settings;
    const save = async () => { await this.plugin.saveSettings(); };
    new Setting(el).setName("Python command").setDesc("Command that runs Python on this computer (python, py, or a full path).")
      .addText((t) => t.setValue(s.pythonCmd).onChange(async (v) => { s.pythonCmd = v.trim() || "python"; await save(); }));
    new Setting(el).setName("PO year folder").setDesc("Leave blank for the default 2026 folder in Box Drive. Read-only: nothing is ever written there.")
      .addText((t) => t.setPlaceholder("C:\\Users\\...\\Purchase Orders\\2026").setValue(s.poRoot).onChange(async (v) => { s.poRoot = v.trim(); await save(); }));
    new Setting(el).setName("Refresh every (minutes)").setDesc("How often the dashboard re-scans Box while Obsidian is open. It also picks up any refresh the scheduled task makes.")
      .addText((t) => t.setValue(String(s.refreshMinutes)).onChange(async (v) => { s.refreshMinutes = Math.max(1, parseInt(v) || 5); await save(); }));
    el.createEl("h3", { text: "monday.com (optional)" });
    new Setting(el).setName("Show monday tab").addToggle((t) => t.setValue(s.mondayEnabled).onChange(async (v) => { s.mondayEnabled = v; await save(); this.plugin.redraw(); }));
    new Setting(el).setName("API token").setDesc("monday → your avatar → Developers → My access tokens. Stored in this vault's plugin settings on this computer only; use a read-only account token if you can.")
      .addText((t) => { t.inputEl.type = "password"; t.setValue(s.mondayToken).onChange(async (v) => { s.mondayToken = v.trim(); await save(); }); });
    new Setting(el).setName("Set DHR automatically").setDesc("When every compiled part of a PO is in Completed Receivings in Box and exactly one Product Release item matches it, set that item's Status to DHR and post an update explaining why. Changes are made with the API token above, so they show in monday under that token's owner.")
      .addToggle((t) => t.setValue(s.autoDhr).onChange(async (v) => { s.autoDhr = v; await save(); }));
    new Setting(el).setName("Board ID").setDesc("Default is the \"New inspection\" board in the Inspection workspace. Item names containing a PO number (e.g. 26-1004) are matched to PO folders.")
      .addText((t) => t.setValue(s.mondayBoardId).onChange(async (v) => { s.mondayBoardId = v.trim(); await save(); }));
  }
}

module.exports = ReceivingDashboardPlugin;
module.exports.helpers = { categorizeIssue, delayModel, parseStatusEvents, parseStuckIssues, dhrCandidates, parseQuery, filterPOs, filterMatrix, filterFiles, areaOf, hiddenPath, poFromText, pct };
