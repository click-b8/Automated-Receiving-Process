# Screenshot harness

Produces every image in `docs/images/` from the **real** plugin code and vault files, fed with **synthetic** data. No real record is involved.

```bash
cd tools/screenshot-harness
npm install                      # playwright, d3-force, marked, lucide-static
npx playwright install chromium  # if you don't already have a browser for Playwright
python demo_data.py              # writes ./demo/*.json (invented POs, vendors, parts, monday board)

node render_dashboard.js ../../tools-folder ../../docs/images
node render_canvas.js ../../tools-folder "Project Map/Project Map.canvas"        ../../docs/images/canvas-project-map.png 0.42
node render_canvas.js ../../tools-folder "Project Map/QE Receiving Workflow.canvas" ../../docs/images/canvas-qe-workflow.png 0.55
node render_graph.js  ../../tools-folder ../../docs/images/graph-view.png
node render_graph.js  ../../tools-folder ../../docs/images/graph-local-workflow.png "" "QE Receiving Workflow"
```

| Script | How it works |
|---|---|
| `render_dashboard.js` | Loads `main.js` + `styles.css` into a page with a small mock of the Obsidian API (`Plugin`, `ItemView`, DOM helpers, `requestUrl`, a vault built from the tools folder), Obsidian's default dark theme variables and Lucide icons. Instantiates the real plugin and view, switches tabs and screenshots each one. Time is frozen with Playwright's clock |
| `render_canvas.js` | Reads a `.canvas` JSON file and draws groups, text and file cards (with note previews) and labelled bezier edges in Obsidian's canvas colors |
| `render_graph.js` | Builds nodes and edges from `[[wikilinks]]`, lays them out with d3-force, colors by the vault's graph groups, labels hubs only (or every node in local mode) |

The published graph images were rendered from the internal vault's link structure after the same scrubbing used for this repo, so they show the full file clusters. Run against `tools-folder` you'll get the smaller public vault.
