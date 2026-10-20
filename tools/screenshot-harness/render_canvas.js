// Renders an Obsidian .canvas file (from the public tools-folder) to PNG in an Obsidian-like dark style.
// node render_canvas.js <tools-folder> <canvas-rel-path> <out.png> [scale]
const fs = require("fs"), path = require("path");
const { chromium } = require("playwright");
const { marked } = require("marked");
const [TOOLS, REL, OUTP, SC] = [path.resolve(process.argv[2]), process.argv[3], path.resolve(process.argv[4]), Number(process.argv[5] || 0.6)];
const cv = JSON.parse(fs.readFileSync(path.join(TOOLS, REL), "utf8"));
const COL = { 1: "#fb464c", 2: "#e9973f", 3: "#e0de71", 4: "#44cf6e", 5: "#53dfdd", 6: "#a882ff" };
const col = (c) => (c ? (COL[c] || c) : "#5a5a5a");
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
const wl = (md) => md.replace(/\[\[([^\]|]+)\|([^\]]+)\]\]/g, '<span class="wl">$2</span>').replace(/\[\[([^\]]+)\]\]/g, '<span class="wl">$1</span>');
let minX = 1e9, minY = 1e9, maxX = -1e9, maxY = -1e9;
cv.nodes.forEach((n) => { minX = Math.min(minX, n.x); minY = Math.min(minY, n.y); maxX = Math.max(maxX, n.x + n.width); maxY = Math.max(maxY, n.y + n.height); });
const PAD = 80; const W = maxX - minX + PAD * 2, H = maxY - minY + PAD * 2;
const X = (x) => x - minX + PAD, Y = (y) => y - minY + PAD;
const byId = Object.fromEntries(cv.nodes.map((n) => [n.id, n]));
function anchor(n, side) { const x = X(n.x), y = Y(n.y);
  return { top: [x + n.width / 2, y], bottom: [x + n.width / 2, y + n.height], left: [x, y + n.height / 2], right: [x + n.width, y + n.height / 2] }[side || "right"]; }
const dir = { top: [0, -1], bottom: [0, 1], left: [-1, 0], right: [1, 0] };
let edges = "";
cv.edges.forEach((e) => { const a = byId[e.fromNode], b = byId[e.toNode]; if (!a || !b) return;
  const fs_ = e.fromSide || "right", ts_ = e.toSide || "left"; const p = anchor(a, fs_), q = anchor(b, ts_);
  const d = Math.max(60, Math.hypot(q[0] - p[0], q[1] - p[1]) / 3);
  const c1 = [p[0] + dir[fs_][0] * d, p[1] + dir[fs_][1] * d], c2 = [q[0] + dir[ts_][0] * d, q[1] + dir[ts_][1] * d];
  const c = col(e.color) === "#5a5a5a" ? "#8a8a8a" : col(e.color);
  edges += `<path d="M${p} C${c1} ${c2} ${q}" stroke="${c}" stroke-width="3" fill="none" marker-end="url(#ar${e.color || 0})"/>`;
  if (e.label) { const mx = (p[0] + 3 * c1[0] + 3 * c2[0] + q[0]) / 8, my = (p[1] + 3 * c1[1] + 3 * c2[1] + q[1]) / 8;
    edges += `<foreignObject x="${mx - 110}" y="${my - 16}" width="220" height="32"><div class="elabel">${esc(e.label)}</div></foreignObject>`; } });
const markers = [0, 1, 2, 3, 4, 5, 6].map((k) => `<marker id="ar${k}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="${k ? COL[k] : "#8a8a8a"}"/></marker>`).join("");
let nodes = "";
cv.nodes.filter((n) => n.type === "group").forEach((n) => { nodes += `<div class="group" style="left:${X(n.x)}px;top:${Y(n.y)}px;width:${n.width}px;height:${n.height}px;--c:${col(n.color)}"><div class="glabel">${esc(n.label || "")}</div></div>`; });
cv.nodes.filter((n) => n.type !== "group").forEach((n) => {
  let inner = "";
  if (n.type === "text") inner = `<div class="md">${marked.parse(wl(n.text || ""))}</div>`;
  else if (n.type === "file") { const f = n.file; const name = f.split("/").pop(); const isMd = f.endsWith(".md");
    let prev = ""; if (isMd && fs.existsSync(path.join(TOOLS, f))) { prev = fs.readFileSync(path.join(TOOLS, f), "utf8").replace(/^---[\s\S]*?---\s*/, "").split("\n").filter((l) => !/^#\s/.test(l)).join("\n").slice(0, 600); }
    inner = `<div class="ftitle">${esc(isMd ? name.slice(0, -3) : name)}</div>` + (prev ? `<div class="md small">${marked.parse(wl(prev))}</div>` : `<div class="fext">${esc(f)}</div>`); }
  nodes += `<div class="node ${n.type}" style="left:${X(n.x)}px;top:${Y(n.y)}px;width:${n.width}px;height:${n.height}px;--c:${col(n.color)}">${inner}</div>`; });
const html = `<!doctype html><html><head><meta charset="utf-8"><style>
body{margin:0;background:#1e1e1e;font-family:Inter,-apple-system,"Segoe UI",sans-serif;color:#dadada}
#c{position:relative;width:${W}px;height:${H}px;background:#1e1e1e;background-image:radial-gradient(#3a3a3a 1.4px,transparent 1.4px);background-size:22px 22px;transform-origin:0 0;transform:scale(${SC})}
svg{position:absolute;left:0;top:0;overflow:visible}
.group{position:absolute;border:3px solid var(--c);border-radius:14px;background:color-mix(in srgb,var(--c) 7%,transparent)}
.glabel{position:absolute;top:-46px;left:0;font-size:24px;font-weight:600;color:#dadada;background:color-mix(in srgb,var(--c) 30%,#1e1e1e);padding:4px 12px;border-radius:8px;white-space:nowrap}
.node{position:absolute;box-sizing:border-box;border:3px solid var(--c);border-radius:10px;background:#262626;overflow:hidden;padding:10px 14px;font-size:15px;line-height:1.35}
.node[style*="#5a5a5a"]{border-color:#4a4a4a}
.node.text{background:color-mix(in srgb,var(--c) 10%,#232323)}
.md h1{font-size:26px;margin:.1em 0 .3em}.md h2{font-size:20px;margin:.1em 0 .3em}.md h3{font-size:17px;margin:.1em 0 .3em}
.md p,.md ul{margin:.2em 0}.md ul{padding-left:1.2em}.md code{background:#333;border-radius:4px;padding:0 4px;font-size:.9em}
.md.small{font-size:12px;color:#a3a3a3;margin-top:4px}.md.small h2,.md.small h3{font-size:13px}
.ftitle{font-weight:700;font-size:17px}.fext{font-size:12px;color:#8a8a8a;margin-top:4px;font-family:Menlo,Consolas,monospace}
.wl{color:#a882ff}.elabel{display:inline-block;max-width:220px;background:#1e1e1e;color:#cfcfcf;font-size:14px;padding:3px 8px;border-radius:6px;border:1px solid #444;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
</style></head><body><div id="c"><svg width="${W}" height="${H}"><defs>${markers}</defs>${edges}</svg>${nodes}</div></body></html>`;
fs.writeFileSync(path.join(__dirname, "canvas.html"), html);
(async () => { const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: Math.ceil(W * SC), height: Math.ceil(H * SC) }, deviceScaleFactor: 1.5 });
  await p.goto("file://" + path.join(__dirname, "canvas.html")); await p.waitForTimeout(300);
  await p.screenshot({ path: OUTP, clip: { x: 0, y: 0, width: Math.ceil(W * SC), height: Math.ceil(H * SC) } }); await b.close(); console.log("saved", OUTP, Math.ceil(W * SC), Math.ceil(H * SC)); })();
