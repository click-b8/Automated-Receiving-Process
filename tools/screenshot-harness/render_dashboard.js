// Renders the REAL Receiving Dashboard plugin (main.js + styles.css) inside a mock Obsidian
// environment fed with synthetic demo data, then screenshots each tab.
// node render_dashboard.js <tools-folder> <out-dir>
const fs = require("fs"), path = require("path");
const { chromium } = require("playwright");
const TOOLS = path.resolve(process.argv[2]); const OUT = path.resolve(process.argv[3]);
const PLUG = path.join(TOOLS, ".obsidian/plugins/receiving-dashboard");
const DEMO = path.join(__dirname, "demo");
fs.mkdirSync(OUT, { recursive: true });

const ICONS = ["alert-triangle","book-open","external-link","file-check","file-text","folder-cog","folder-open","gauge","layout-grid",
  "list","message-square","refresh-cw","route","search","sheet","layout-dashboard","package","check-square","folder-tree","hourglass",
  "calendar-check","files","git-fork","layout-template","settings","help-circle","file-plus","folder","network","terminal-square","x","plus","chevron-down","more-horizontal"];
const icons = {};
for (const n of ICONS) { const p = path.join(__dirname, "node_modules/lucide-static/icons", n + ".svg"); if (fs.existsSync(p)) icons[n] = fs.readFileSync(p, "utf8").replace(/<!--[\s\S]*?-->/g, ""); }

// vault file list from the public tools-folder (what the Files tab shows)
const files = [];
(function walk(d) { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name);
  if (e.isDirectory()) walk(p); else { const rel = path.relative(TOOLS, p).split(path.sep).join("/"); const st = fs.statSync(p);
    files.push({ path: rel, size: st.size, mtime: Date.parse("2026-10-" + String(1 + (rel.length % 19)).padStart(2, "0") + "T12:00:00"),
      content: rel.endsWith(".md") ? fs.readFileSync(p, "utf8") : "" }); } } })(TOOLS);
// a few representative non-note files so the navigator looks like a working tools folder (names are synthetic)
[["logs/PO26-1214_SR90-1131_20261014-1512_manifest.json", 5400], ["logs/PO26-1219_SR90-1152_20261016-0931_manifest.json", 5120],
 ["Validation/QF_7.4.3-06_Validation_Protocol.docx", 61200], ["Validation/QF_7.4.3-06_Attachment_A_Test_Case_Matrix.docx", 88400],
 ["Validation/QF_7.4.3-06_Attachment_B_Instruction_Guide.docx", 40100], ["parts_cache.json", 8653], ["templates/PRF template CONTROLS.xlsx", 74210],
 ["_backup/PO26-1203_pre-rerun_2026-09-24/PO26-1203 Receiving Part 1 - SR90-1104.pdf", 2210400]].forEach(([p, s], i) =>
  files.push({ path: p, size: s, mtime: Date.parse("2026-10-1" + (i % 9) + "T09:00:00"), content: "" }));

const env = { icons, files, data: fs.readFileSync(path.join(DEMO, "dashboard_data.json"), "utf8"),
  monday: JSON.parse(fs.readFileSync(path.join(DEMO, "monday_response.json"), "utf8")),
  dhrlog: fs.readFileSync(path.join(DEMO, "dhr_log.jsonl"), "utf8") };

const THEME = `
body{margin:0;background:#1e1e1e;color:#dadada;font-family:Inter,-apple-system,"Segoe UI",Roboto,sans-serif;font-size:15px;
 --background-primary:#1e1e1e;--background-primary-alt:#242424;--background-secondary:#262626;--background-secondary-alt:#2a2a2a;
 --background-modifier-border:#363636;--background-modifier-border-hover:#2f2f2f;--background-modifier-hover:rgba(255,255,255,.075);
 --text-normal:#dadada;--text-muted:#a3a3a3;--text-faint:#666;--text-on-accent:#fff;--interactive-accent:#8a5cf5;
 --color-red:#fb464c;--color-red-rgb:251,70,76;--color-orange:#e9973f;--color-orange-rgb:233,151,63;--color-yellow:#e0de71;--color-yellow-rgb:224,222,113;
 --color-green:#44cf6e;--color-green-rgb:68,207,110;--color-cyan:#53dfdd;--color-cyan-rgb:83,223,221;--color-blue:#4f9dff;--color-blue-rgb:79,157,255;
 --color-purple:#a882ff;--color-purple-rgb:168,130,255;--font-monospace:"JetBrains Mono",Menlo,Consolas,monospace}
button{background:#363636;color:#dadada;border:0;border-radius:6px;padding:6px 12px;font:inherit;font-size:.9em}
input{background:#1e1e1e;color:#dadada;border:1px solid #363636;border-radius:6px;padding:6px 10px;font:inherit}
.win{display:flex;width:1440px;min-height:900px;background:#1e1e1e}
.ribbon{width:44px;background:#262626;border-right:1px solid #363636;display:flex;flex-direction:column;align-items:center;padding-top:12px;gap:14px;color:#a3a3a3}
.ribbon svg{width:18px;height:18px}.ribbon .on{color:#8a5cf5}
.main{flex:1;display:flex;flex-direction:column;min-width:0}
.tabbar{height:38px;background:#262626;border-bottom:1px solid #363636;display:flex;align-items:flex-end;padding-left:8px;gap:4px}
.tabh{display:flex;align-items:center;gap:8px;padding:7px 14px;font-size:13px;color:#a3a3a3;border-radius:8px 8px 0 0}
.tabh.on{background:#1e1e1e;color:#dadada}.tabh svg{width:14px;height:14px}
.view-content{overflow:visible}
svg{vertical-align:middle}`;

const HARNESS = `<!doctype html><html><head><meta charset="utf-8"><style>${THEME}</style><style>${fs.readFileSync(path.join(PLUG, "styles.css"), "utf8")}</style></head>
<body><div class="win"><div class="ribbon"></div><div class="main"><div class="tabbar"></div><div id="host"></div></div></div>
<script>
const ENV = ${JSON.stringify(env)};
// ---- Obsidian DOM helpers ----
const P = HTMLElement.prototype;
function applyOpts(el, o){ if(o==null) return; if(typeof o==="string"){ o.split(" ").filter(Boolean).forEach(c=>el.classList.add(c)); return; }
  if(o.cls) (Array.isArray(o.cls)?o.cls:o.cls.split(" ")).filter(Boolean).forEach(c=>el.classList.add(c));
  if(o.text!=null) el.textContent=o.text; if(o.attr) for(const k in o.attr) el.setAttribute(k,o.attr[k]);
  for(const k of ["type","placeholder","value","title","href"]) if(o[k]!=null) el[k]=o[k]; }
P.createEl=function(tag,o,cb){const e=document.createElement(tag);applyOpts(e,o);this.appendChild(e);if(cb)cb(e);return e;};
P.createDiv=function(o,cb){return this.createEl("div",o,cb);}; P.createSpan=function(o,cb){return this.createEl("span",o,cb);};
P.empty=function(){this.innerHTML="";}; P.addClass=function(...c){this.classList.add(...c.flatMap(x=>x.split(" ")).filter(Boolean));};
P.removeClass=function(...c){this.classList.remove(...c);}; P.toggleClass=function(c,v){this.classList.toggle(c,v);}; P.setText=function(t){this.textContent=t;};
function setIcon(el,name){ el.innerHTML = ENV.icons[name] || ENV.icons["file-text"] || ""; }
// ---- mock vault ----
class TFile{ constructor(f){ this.path=f.path; this.name=f.path.split("/").pop(); const i=this.name.lastIndexOf(".");
  this.extension=i>0?this.name.slice(i+1):""; this.basename=i>0?this.name.slice(0,i):this.name; this.stat={size:f.size,mtime:f.mtime,ctime:f.mtime}; this._c=f.content; } }
const TF = ENV.files.map(f=>new TFile(f));
const app = { vault: { adapter:{ getBasePath:()=>"C:\\\\Users\\\\<you>\\\\Desktop\\\\Receiving Task Automation Tools" },
    getFiles:()=>TF, getMarkdownFiles:()=>TF.filter(f=>f.extension==="md"), getAbstractFileByPath:(p)=>TF.find(f=>f.path===p)||null },
  metadataCache:{ getFileCache:(f)=>({ links:[...(f._c||"").matchAll(/\\[\\[([^\\]|#]+)(?:#[^\\]|]*)?(?:\\|[^\\]]*)?\\]\\]/g)].map(m=>({link:m[1]})) }),
    getFirstLinkpathDest:(l)=>TF.find(f=>f.basename===l)||null },
  workspace:{ onLayoutReady:(cb)=>setTimeout(cb,0), getLeavesOfType:()=>window.__leaves||[], getLeaf:()=>({}), revealLeaf(){} } };
// ---- mock modules ----
class Plugin{ constructor(app,m){this.app=app;this.manifest=m;} async loadData(){return {pythonCmd:"python",refreshMinutes:5,mondayEnabled:true,mondayToken:"demo",mondayBoardId:"1234567890",autoDhr:true};}
  async saveData(){} registerView(){} addRibbonIcon(){} addCommand(){} addSettingTab(){} }
class ItemView{ constructor(leaf){ this.leaf=leaf; this.app=leaf.app; this.containerEl=document.createElement("div"); this.containerEl.createDiv("view-header"); this.containerEl.createDiv("view-content"); } }
class PluginSettingTab{ constructor(app,p){this.app=app;this.plugin=p;} } class Setting{ constructor(){ return new Proxy(this,{get:(t,k)=>()=>t}); } }
class Notice{ constructor(){} }
async function requestUrl(){ return { status:200, json: ENV.monday }; }
const fsMock = { stat:(p,cb)=>cb(null,{mtimeMs:1}), readFile:(p,enc,cb)=>cb(null,ENV.data),
  readFileSync:(p)=>{ if(/dhr_log/.test(p)) return ENV.dhrlog; const rel=p.replace(/\\\\/g,"/").split("Receiving Task Automation Tools/")[1]; const f=TF.find(x=>x.path===rel); if(f) return f._c; throw new Error("nf"); },
  appendFileSync:()=>{} };
const pathMock = { join:(...a)=>a.join("\\\\"), dirname:(p)=>p.split("/").slice(0,-1).join("/") };
const modules = { obsidian:{Plugin,ItemView,Notice,PluginSettingTab,Setting,requestUrl,setIcon,TFile}, fs:fsMock, path:pathMock,
  child_process:{ execFile:(c,a,o,cb)=>setTimeout(()=>cb(null,"",""),0) }, electron:{ shell:{ openPath:async()=>"", showItemInFolder(){} } } };
function require(n){ return modules[n]; }
const module = { exports:{} }; const exports = module.exports;
</script>
<script>${fs.readFileSync(path.join(PLUG, "main.js"), "utf8")}</script>
<script>
(async()=>{
  const PluginClass = module.exports.default || module.exports;
  const plugin = new PluginClass(app, {id:"receiving-dashboard"}); await plugin.onload();
  const ViewClass = Object.getPrototypeOf(Object.getPrototypeOf(plugin)).constructor; // not used
  const leaf = { app };
  // find DashboardView via registerView capture
})();
</script></body></html>`;

// main.js defines DashboardView as a module-level class; expose it by appending an export line
const mainSrc = "(function(require, module, exports){\n" + fs.readFileSync(path.join(PLUG, "main.js"), "utf8") + "\n;window.__DV = DashboardView;\n})(require, module, exports);";
const html = HARNESS.replace(fs.readFileSync(path.join(PLUG, "main.js"), "utf8"), () => mainSrc)
  .replace(/\(async\(\)=>\{[\s\S]*?\}\)\(\);\n<\/script><\/body>/, `(async()=>{
  const PluginClass = module.exports.default || module.exports;
  const plugin = new PluginClass(app, {id:"receiving-dashboard"}); await plugin.onload();
  const view = new window.__DV({ app }, plugin); window.__leaves=[{view}];
  document.getElementById("host").appendChild(view.containerEl);
  const rib=document.querySelector(".ribbon"); ["files","search","layout-grid","git-fork","gauge","terminal-square"].forEach((n,i)=>{const s=document.createElement("span"); s.innerHTML=ENV.icons[n]||""; if(n==="gauge") s.className="on"; rib.appendChild(s);});
  const tb=document.querySelector(".tabbar"); [["00 Index","file-text"],["Receiving Dashboard","gauge"],["QE Receiving Workflow","layout-grid"]].forEach(([t,i])=>{const d=document.createElement("div"); d.className="tabh"+(t==="Receiving Dashboard"?" on":""); d.innerHTML=(ENV.icons[i]||"")+"<span>"+t+"</span>"; tb.appendChild(d);});
  await new Promise(r=>setTimeout(r,300)); await plugin.fetchMonday(); await new Promise(r=>setTimeout(r,300));
  window.__view=view; window.__ready=true;
})();
</script></body>`);
fs.writeFileSync(path.join(__dirname, "harness.html"), html);

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.5 });
  await page.clock.setFixedTime(new Date("2026-10-20T10:45:00"));
  page.on("pageerror", (e) => console.error("PAGEERR", e.message));
  await page.goto("file://" + path.join(__dirname, "harness.html"));
  await page.waitForFunction(() => window.__ready === true, null, { timeout: 15000 });
  const shots = [["overview", null, "dashboard-overview"], ["pos", null, "dashboard-pos"], ["workflow", null, "dashboard-workflow"],
    ["matrix", null, "dashboard-test-matrix"], ["files", null, "dashboard-files"], ["delays", null, "dashboard-delays"], ["monday", null, "dashboard-monday"]];
  for (const [tab, q, name] of shots) {
    await page.evaluate(([t, q]) => { window.__view.go(t, q); }, [tab, q]);
    if (tab === "pos") await page.evaluate(() => { const v = window.__view; const p = v.plugin.data.pos.find(x => x.parts.length > 2 && x.status !== "filed") || v.plugin.data.pos[50]; v.expanded.add(p.po); v.sort.pos = { col: "last", dir: -1 }; v.render(); });
    await page.waitForTimeout(250);
    const el = await page.$(".win");
    const bb = await el.boundingBox();
    const maxH = ["dashboard-overview", "dashboard-delays", "dashboard-workflow"].includes(name) ? bb.height : Math.min(bb.height, 1120);
    await page.setViewportSize({ width: 1440, height: Math.ceil(bb.height) });
    await page.screenshot({ path: path.join(OUT, name + ".png"), clip: { x: 0, y: 0, width: 1440, height: maxH } });
    console.log("saved", name);
  }
  await browser.close();
})();
