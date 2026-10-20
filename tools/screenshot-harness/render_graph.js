// Obsidian-style Graph view of the vault. Builds nodes/edges from [[wikilinks]] in the notes under <vault>,
// lays them out with d3-force, colors them with the vault's graph.json color groups, renders to PNG.
// Node labels go through scrub.py's mapping first (passed in as a JSON label map) so nothing private is drawn.
// node render_graph.js <vault-dir> <out.png> [labelmap.json] [focus-note-basename]
const fs = require("fs"), path = require("path");
const d3 = require("d3-force");
const { chromium } = require("playwright");
const VAULT = path.resolve(process.argv[2]), OUT = path.resolve(process.argv[3]);
const LABELS = process.argv[4] && fs.existsSync(process.argv[4]) ? JSON.parse(fs.readFileSync(process.argv[4], "utf8")) : {};
const FOCUS = process.argv[5] || null;
const md = [];
(function walk(d) { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name);
  if (e.name.startsWith(".") || e.name === "_scripts") continue; if (e.isDirectory()) walk(p); else if (/\.(md|canvas)$/.test(e.name)) md.push(path.relative(VAULT, p).split(path.sep).join("/")); } })(VAULT);
const id = (p) => p.replace(/\.md$/, "");
const base = (p) => p.split("/").pop().replace(/\.md$/, "");
const byBase = {}; md.forEach((p) => (byBase[base(p)] = p));
const nodes = new Map(), links = [];
const add = (k, kind, p) => { if (!nodes.has(k)) nodes.set(k, { id: k, kind, path: p || k, deg: 0 }); return nodes.get(k); };
md.forEach((p) => add(id(p), "note", p));
md.filter((p) => p.endsWith(".md")).forEach((p) => {
  const t = fs.readFileSync(path.join(VAULT, p), "utf8");
  for (const m of t.matchAll(/\[\[([^\]|#]+)(?:#[^\]|]*)?(?:\|[^\]]*)?\]\]/g)) {
    let tgt = m[1].trim(); let key;
    if (byBase[tgt.replace(/\.md$/, "")]) key = id(byBase[tgt.replace(/\.md$/, "")]);
    else if (byBase[tgt]) key = id(byBase[tgt]);
    else { key = tgt; add(key, /\.(md)$/.test(tgt) || !/\.\w{2,5}$/.test(tgt) ? "note" : "file", tgt); if (!/\.\w{2,5}$/.test(tgt)) continue; }
    if (key === id(p)) continue; links.push({ source: id(p), target: key });
  } });
// drop unresolved note links (hideUnresolved: true)
let N = [...nodes.values()].filter((n) => n.kind === "file" || md.includes(n.path) || md.includes(n.path + ".md"));
let keep = new Set(N.map((n) => n.id));
let L = links.filter((l) => keep.has(l.source) && keep.has(l.target));
if (FOCUS) { const f = N.find((n) => n.path.split("/").pop().replace(/\.md$/, "") === FOCUS).id;
  const nb = new Set([f]); L.forEach((l) => { if (l.source === f) nb.add(l.target); if (l.target === f) nb.add(l.source); });
  N = N.filter((n) => nb.has(n.id)); keep = nb; L = L.filter((l) => nb.has(l.source) && nb.has(l.target)); }
L.forEach((l) => { nodes.get(l.source).deg++; nodes.get(l.target).deg++; });
function color(n) { const p = n.path;
  if (n.kind === "file") { if (p.startsWith("logs/")) return "#20a39e"; if (p.startsWith("_backup/")) return "#6b7a8f"; if (/^(process|prefill)_|^(find_|inspect_)/.test(p)) return "#e05252";
    if (p.startsWith("Validation/")) return "#2e7d32"; return "#9e9e9e"; }
  if (p.includes("05 QE Workflow")) return "#e8833a"; if (p.includes("01 Governance") || /CLAUDE/.test(p)) return "#8e6cef";
  if (p.includes("02 Code")) return "#2bb3c0"; if (p.includes("03 Validation")) return "#4caf50"; if (p.includes("04 History")) return "#d4b106";
  if (p.includes("06 Files")) return "#bdbdbd"; return "#dadada"; }
let sim = d3.forceSimulation(N).force("link", d3.forceLink(L).id((d) => d.id).distance((l) => (l.target.kind === "file" ? 26 : (FOCUS ? 200 : 130))).strength(0.7))
  .force("charge", d3.forceManyBody().strength((d) => (d.kind === "file" ? -18 : (FOCUS ? -900 : -520)))).force("center", d3.forceCenter(0, 0))
  .force("x", d3.forceX(0).strength(0.04)).force("y", d3.forceY(0).strength(0.05)).force("collide", d3.forceCollide((d) => 4 + Math.sqrt(d.deg) * 2.2)).stop();
for (let i = 0; i < 600; i++) sim.tick();
let x0 = 1e9, y0 = 1e9, x1 = -1e9, y1 = -1e9; N.forEach((n) => { x0 = Math.min(x0, n.x); y0 = Math.min(y0, n.y); x1 = Math.max(x1, n.x); y1 = Math.max(y1, n.y); });
const W = 1800, H = 1250, pad = 90, s = Math.min((W - 2 * pad) / (x1 - x0), (H - 2 * pad) / (y1 - y0));
const P = (n) => [pad + (n.x - x0) * s + ((W - 2 * pad) - (x1 - x0) * s) / 2, pad + (n.y - y0) * s + ((H - 2 * pad) - (y1 - y0) * s) / 2];
const R = (n) => (n.kind === "file" ? 3.2 : 4 + Math.sqrt(n.deg) * 2.1);
const label = (n) => { const b = n.path.split("/").pop().replace(/\.md$/, ""); return LABELS[b] || b; };
const focusSet = false ? new Set([FOCUS, ...L.filter((l) => l.source.id.endsWith(FOCUS) || l.target.id.endsWith(FOCUS)).flatMap((l) => [l.source.id, l.target.id])]) : null;
const dim = (n) => focusSet && ![...focusSet].some((f) => n.id === f || n.id.endsWith("/" + f) || n.id === f);
let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="100%" height="100%" fill="#1e1e1e"/>`;
L.forEach((l) => { const [a, b] = [P(l.source), P(l.target)]; const on = focusSet && !dim(l.source) && !dim(l.target);
  svg += `<line x1="${a[0]}" y1="${a[1]}" x2="${b[0]}" y2="${b[1]}" stroke="${on ? "#8e6cef" : "#4a4a4a"}" stroke-opacity="${focusSet && !on ? 0.25 : 0.75}" stroke-width="${on ? 1.6 : 1}"/>`; });
N.forEach((n) => { const [x, y] = P(n); svg += `<circle cx="${x}" cy="${y}" r="${R(n)}" fill="${color(n)}" fill-opacity="${dim(n) ? 0.25 : 1}"/>`; });
N.filter((n) => n.kind === "note" && (FOCUS || /_Hub|00 Index|QE Receiving Workflow|WI-7.4.3-01 Overview|receive.py|Files - /.test(n.path))).forEach((n) => { const [x, y] = P(n);
  svg += `<text x="${x}" y="${y + R(n) + 14}" text-anchor="middle" font-family="Inter,Segoe UI,sans-serif" font-size="${FOCUS ? 15 : (n.deg > 15 ? 14 : 12.5)}" fill="#dadada" fill-opacity="${dim(n) ? 0.3 : 0.92}">${label(n).replace(/&/g, "&amp;").replace(/</g, "&lt;")}</text>`; });
// legend
const leg = [["Instruction / WI", "#8e6cef"], ["QE workflow", "#e8833a"], ["Code", "#2bb3c0"], ["Validation", "#4caf50"], ["History", "#d4b106"], ["File hubs", "#bdbdbd"],
  ["Run manifests", "#20a39e"], ["Backups", "#6b7a8f"], ["Edge-case scripts", "#e05252"], ["Validation docs", "#2e7d32"]];
leg.forEach(([t, c], i) => { svg += `<circle cx="${30}" cy="${30 + i * 22}" r="6" fill="${c}"/><text x="44" y="${35 + i * 22}" font-family="Inter,Segoe UI,sans-serif" font-size="13" fill="#bbb">${t}</text>`; });
svg += `</svg>`;
fs.writeFileSync(OUT.replace(/\.png$/, ".svg"), svg);
(async () => { const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: W, height: H }, deviceScaleFactor: 1.2 });
  await p.setContent(`<html><body style="margin:0">${svg}</body></html>`); await p.screenshot({ path: OUT }); await b.close();
  fs.unlinkSync(OUT.replace(/\.png$/, ".svg")); console.log("graph", N.length, "nodes", L.length, "edges ->", OUT); })();
