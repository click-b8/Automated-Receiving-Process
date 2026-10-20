"""
refresh_dashboard.py — collects the data behind the Receiving Dashboard plugin
(writes Project Map\\_scripts\\.dashboard_data.json; --md also writes the old note).

    python "Project Map\\_scripts\\refresh_dashboard.py"
    python "Project Map\\_scripts\\refresh_dashboard.py" --po-root "<other year folder>"

READ-ONLY on everything outside "Project Map\\": it lists the Box PO folders,
reads the signed packages in Completed Receiving(s) once each (to count
digital signatures; cached by size + modified time so Box Drive only
downloads a file the first time), reads logs\\ manifests, the edge-case
scripts' names, the Test Case Matrix .docx and CLAUDE.md. The only files it
writes are the dashboard note and its cache, both inside "Project Map\\".
Not part of the validated tool (receive.py / prf_helpers.py are untouched).
Standard library only.
"""
import argparse, datetime as dt, json, os, re, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

VAULT = Path(__file__).resolve().parents[2]
PM = VAULT / "Project Map"
OUT = PM / "07 Dashboard" / "Receiving Dashboard.md"
CACHE = PM / "_scripts" / ".dashboard_cache.json"
DEFAULT_ROOTS = [r"C:\Users\<you>\Box\Shared\Purchase Orders\2026"]
MATRIX = VAULT / "Validation" / "QF_7.4.3-06_Attachment_A_Test_Case_Matrix.docx"

PO_DIR = re.compile(r"^(\d{2})-(\d{4})\s*-?\s*(.*)$")
PKG = re.compile(r"^(?:PO\s*)?(\d{2})-(\d{4})\s*-?\s*Receiving\s+Part\s*(\d+)\s*-\s*(.+?)\.pdf$", re.I)
PRF = re.compile(r"PRF", re.I)

def po_key(yy, nnnn): return f"PO{yy}-{nnnn}"

def is_completed_dir(name):  # real folders are "Completed Receivings"; WI says "Completed Receiving"
    return name.lower().startswith("completed receiving")

def is_inprogress_dir(name): return name.lower().startswith("in-progress")

def safe_scandir(p):
    try:
        return list(os.scandir(p))
    except OSError:
        return []

def count_signatures(path, cache):
    """Number of digital signatures = '/ByteRange' entries. Cached by size+mtime."""
    st = os.stat(path)
    key = str(path)
    hit = cache.get(key)
    if hit and hit["size"] == st.st_size and hit["mtime"] == int(st.st_mtime):
        return hit["sigs"]
    try:
        with open(path, "rb") as f:
            sigs = f.read().count(b"/ByteRange")
    except OSError:
        sigs = -1
    cache[key] = {"size": st.st_size, "mtime": int(st.st_mtime), "sigs": sigs}
    return sigs

# ------------------------------------------------------------------ Box scan
def add_file(po, part, fp, comp, archived):
    """Record one package PDF. A part is signed off once its package is in Completed Receiving(s)."""
    rec = po["parts"].setdefault(part, {"inprog": None, "completed": None, "archived_only": True, "mtime": 0,
                                        "path_completed": "", "path_inprog": "", "path_any": ""})
    mt = os.stat(fp).st_mtime
    rec["mtime"] = max(rec["mtime"], mt)
    rec["path_any"] = rec["path_any"] or fp
    if archived:
        return
    rec["archived_only"] = False
    if comp:
        rec["completed"] = max(rec["completed"] or 0, mt)
        rec["path_completed"] = fp
    else:
        rec["inprog"] = max(rec["inprog"] or 0, mt)
        rec["path_inprog"] = fp

def scan_pos(roots, cache):
    pos = {}
    for root in roots:
        for d in safe_scandir(root):
            if not d.is_dir():
                continue
            m = PO_DIR.match(d.name)
            if not m:
                continue
            key = po_key(m.group(1), m.group(2))
            po = pos.setdefault(key, {"vendor": m.group(3).strip(" -") or "?", "folder": d.name,
                                      "root": root, "parts": {}, "prfs": set()})
            for sub in safe_scandir(d.path):
                if not sub.is_dir():
                    continue
                comp, prog = is_completed_dir(sub.name), is_inprogress_dir(sub.name)
                if sub.name.lower() == "inspection forms":
                    for f in safe_scandir(sub.path):
                        if f.is_file() and PRF.search(f.name) and f.name.lower().endswith((".xlsx", ".pdf")):
                            if f.name.lower().endswith(".xlsx"):
                                po["prfs"].add(f.name)
                            po.setdefault("prf_files", []).append(f.path)
                if not (comp or prog):
                    continue
                for dirpath, dirnames, filenames in os.walk(sub.path):
                    archived = "archive" in Path(dirpath).relative_to(sub.path).as_posix().lower()
                    for fn in filenames:
                        if not fn.lower().endswith(".pdf"):
                            continue
                        fp = os.path.join(dirpath, fn)
                        pm = PKG.match(fn)
                        if not pm:
                            # Non-standard name (e.g. "Tissue- PO26-1014 Receiving.pdf"): anything filed in
                            # Completed Receiving(s) is a signed-off package; resolve its part below.
                            if comp and not archived:
                                po.setdefault("_loose", []).append(fp)
                            continue
                        part = re.split(r"\s+(?:rev|lot)\b", pm.group(4).strip(), flags=re.I)[0].strip()
                        if archived:      # archived copies carry suffixes ("... No signatures")
                            part = part.split()[0]
                        add_file(po, part, fp, comp, archived)
            # attribute loose filed PDFs to a part: by name, else the PO's only part, else the file itself
            for fp in po.pop("_loose", []):
                stem = Path(fp).stem
                known = [k for k in po["parts"] if k.lower() in stem.lower()]
                if not known and len(po["parts"]) == 1:
                    known = list(po["parts"])
                part = known[0] if known else re.sub(r"(?i)\bPO\s*\d{2}-\d{4}\b|receiving", "", stem).strip(" -_") or stem
                add_file(po, part, fp, True, False)
    return pos

# ------------------------------------------------------------------ local sources
def scan_manifests():
    out = {}
    for f in (VAULT / "logs").glob("*_manifest.json"):
        m = re.match(r"(PO\d{2}-\d{4})_", f.name)
        if not m:
            continue
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            data = {}
        via = data.get("script") or "receive.py"
        out.setdefault(m.group(1), set()).add(via)
    return out

def scan_edge_scripts():
    out = {}
    for f in VAULT.glob("*.py"):
        m = re.match(r"(?:process|prefill)_(?:PO)?(\d{2})-(\d{4})", f.name)
        if m:
            out.setdefault(po_key(m.group(1), m.group(2)), []).append(f.name)
    return out

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
def cell_text(tc):
    """Text of a table cell, ignoring struck-through runs (obsolete wording)."""
    paras = []
    for p in tc.iter(W + "p"):
        s = ""
        for r in p.iter(W + "r"):
            rpr = r.find(W + "rPr")
            if rpr is not None and rpr.find(W + "strike") is not None and \
               rpr.find(W + "strike").get(W + "val", "true") not in ("0", "false"):
                continue
            s += "".join(t.text or "" for t in r.iter(W + "t"))
        paras.append(s)
    return "\n".join(x for x in paras if x.strip()).strip()

def classify_result(txt, ref):
    t = txt.lower()
    if not t:
        return "Not run"
    if "re-test" in t or "retest" in t:
        return "Re-test required"
    if t.startswith("not yet executable") or t.startswith("not yet run") or t.startswith("not run"):
        return "Not run"
    if t.startswith("proposed"):
        return "Proposed"
    if t.startswith("hold") or "on hold" in t:
        return "On hold"
    if "->" in t and "pass" in t.split("->")[-1]:
        return "Pass"
    if t.startswith("fail"):
        return "Fail"
    if t.startswith("pass"):
        return "Pass"
    return "Other"

def scan_matrix():
    rows, refs = [], set()
    try:
        with zipfile.ZipFile(MATRIX) as z:
            root = ET.fromstring(z.read("word/document.xml"))
    except Exception as e:
        return rows, refs, f"could not read matrix: {e}"
    for tbl in root.iter(W + "tbl"):
        for tr in tbl.iter(W + "tr"):
            cells = [cell_text(tc) for tc in tr.iter(W + "tc")]
            if len(cells) < 7 or not re.match(r"TC-\d+", cells[0]):
                continue
            ref = cells[5]
            for yy, n in re.findall(r"\b(\d{2})-(\d{4})\b", ref):
                refs.add(po_key(yy, n))
            rows.append({"id": cells[0].split()[0], "scenario": cells[1].split("\n")[0],
                         "must": cells[4].lower().startswith("yes"),
                         "result": classify_result(cells[6], ref), "ref": ref.replace("\n", " ")})
        break
    return rows, refs, ""

def scan_open_items():
    try:
        s = (VAULT / "CLAUDE.md").read_text(encoding="utf-8")
    except OSError:
        return None
    sec = s.split("## OPEN ITEMS", 1)[-1].split("\n## ", 1)[0]
    items = re.findall(r"^(\d+)\.\s+(.*)$", sec, re.M)
    closed = [n for n, t in items if t.lstrip("*").startswith("~~")]
    return {"total": len(items), "struck": len(closed), "open": len(items) - len(closed)}

# ------------------------------------------------------------------ render
def bar(n, total, width=20):
    if not total:
        return ""
    k = round(width * n / total)
    return "█" * k + "░" * (width - k)

def esc(x): return str(x).replace("|", "/")

def pct(n, d): return f"{(100*n/d):.0f}%" if d else "—"

def fmt_date(ts): return dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d") if ts else ""

def render(pos, manifests, edges, matrix, refs, merr, oi, roots, scan_ok):
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    proc = {k: v for k, v in pos.items() if v["parts"]}
    parts = [(k, p, r) for k, v in proc.items() for p, r in v["parts"].items()]
    n_parts = len(parts)
    comp_parts = [x for x in parts if x[2]["completed"]]
    signed_parts = [x for x in comp_parts if x[2]["sigs"] > 0]
    await_parts = [x for x in parts if not x[2]["completed"] and x[2]["inprog"]]
    po_complete = [k for k, v in proc.items() if all(r["completed"] for r in v["parts"].values())]
    po_partial = [k for k, v in proc.items() if k not in po_complete and any(r["completed"] for r in v["parts"].values())]
    po_open = [k for k in proc if k not in po_complete and k not in po_partial]
    # PRFs waiting: PO has PRFs but fewer packages than PRFs
    waiting = []
    for k, v in pos.items():
        prf_parts = {re.sub(r"(?i)^prf\s*|\s*prf$|\s+rev\b.*$", "", Path(f).stem).strip() for f in v["prfs"]}
        missing = [p for p in prf_parts if p and not any(p.lower() in q.lower() for q in v["parts"])]
        if missing:
            waiting.append((k, v["vendor"], sorted(missing)))
    not_in_matrix = sorted(k for k in proc if k not in refs)
    tot = len(matrix)
    by = {}
    for r in matrix:
        by[r["result"]] = by.get(r["result"], 0) + 1
    must = [r for r in matrix if r["must"]]
    must_pass = sum(1 for r in must if r["result"] == "Pass")
    passed = by.get("Pass", 0)

    L = ["---", "tags: [dashboard, generated]", "---", "",
         "# Receiving Automation Dashboard", "",
         f"*Refreshed **{now}** by `Project Map/_scripts/refresh_dashboard.py`. "
         "Reads the Box PO folders, `logs\\`, the Test Case Matrix and CLAUDE.md; writes nothing outside Project Map. "
         "Edits here are overwritten on the next refresh.*", ""]
    if not scan_ok:
        L += ["> [!warning] Box PO folders not reachable",
              f"> None of these paths could be read: {', '.join(roots)}. PO counts below are empty. "
              "Check that Box Drive is running, or pass `--po-root`.", ""]
    L += ["## At a glance", "",
          "| | Count | |", "|---|---:|---|",
          f"| PO folders in scope | {len(pos)} | {', '.join(Path(r).name for r in roots)} |",
          f"| **POs processed** (≥1 compiled package) | **{len(proc)}** | {pct(len(proc), len(pos))} of PO folders |",
          f"| **Parts processed** (compiled packages) | **{n_parts}** | |",
          f"| POs fully filed to Completed Receiving | {len(po_complete)} | {bar(len(po_complete), len(proc))} {pct(len(po_complete), len(proc))} |",
          f"| POs partly filed | {len(po_partial)} | |",
          f"| POs still in progress | {len(po_open)} | |",
          f"| Parts in Completed Receiving | {len(comp_parts)} | {bar(len(comp_parts), n_parts)} {pct(len(comp_parts), n_parts)} |",
          f"| …with a digital signature detected | {len(signed_parts)} | {pct(len(signed_parts), len(comp_parts))} of filed parts |",
          f"| Parts awaiting signature / filing | {len(await_parts)} | in In-Progress Receiving only |",
          f"| POs with PRFs not yet compiled | {len(waiting)} | see below |",
          f"| **Test Case Matrix complete (Pass)** | **{passed}/{tot}** | {bar(passed, tot)} {pct(passed, tot)} |",
          f"| Must-Pass-100% cases passed | {must_pass}/{len(must)} | {bar(must_pass, len(must))} {pct(must_pass, len(must))} |",
          f"| Processed POs not in the Test Case Matrix | {len(not_in_matrix)} | {pct(len(not_in_matrix), len(proc))} of processed |"]
    if oi:
        L.append(f"| WI Open Items still open | {oi['open']}/{oi['total']} | {oi['struck']} struck (closed or partly closed) |")
    L += ["", "## PO status", "", "```mermaid", "pie showData",
          f'    "Fully filed" : {len(po_complete)}', f'    "Partly filed" : {len(po_partial)}',
          f'    "In progress" : {len(po_open)}', "```", ""]
    L += ["## Test Case Matrix", "", "```mermaid", "pie showData"]
    L += [f'    "{k}" : {v}' for k, v in sorted(by.items(), key=lambda x: -x[1])] + ["```", ""]
    if merr:
        L += [f"> [!warning] {merr}", ""]
    L += ["| TC | Scenario | Must-pass | Result | PO ref |", "|---|---|:-:|---|---|"]
    for r in matrix:
        icon = {"Pass": "✅", "Fail": "❌", "Re-test required": "🔁", "On hold": "⏸", "Not run": "⬜"}.get(r["result"], "•")
        L.append(f"| {r['id']} | {esc(r['scenario'][:70])} | {'●' if r['must'] else ''} | {icon} {r['result']} | {esc(r['ref'][:40])} |")
    L += ["", "## Processed POs not in the Test Case Matrix", "",
          "Candidates to map onto an existing test case, or evidence of a scenario the matrix doesn't cover yet.", "",
          "| PO | Vendor | Parts | Filed | How it ran |", "|---|---|---:|---:|---|"]
    for k in not_in_matrix:
        v = proc[k]
        how = []
        if k in manifests: how.append("manifest: " + ", ".join(sorted(manifests[k])))
        if k in edges: how.append("edge-case script")
        L.append(f"| {k} | {esc(v['vendor'])} | {len(v['parts'])} | {sum(1 for r in v['parts'].values() if r['completed'])} | {'; '.join(how) or '—'} |")
    L += ["", "## Awaiting signature or filing", "", "| PO | Vendor | Part | Compiled |", "|---|---|---|---|"]
    for k, p, r in sorted(await_parts, key=lambda x: x[2]["inprog"] or 0, reverse=True):
        L.append(f"| {k} | {esc(proc[k]['vendor'])} | {esc(p)} | {fmt_date(r['inprog'])} |")
    L += ["", "## PRFs not yet compiled", "", "| PO | Vendor | PRF part(s) without a package |", "|---|---|---|"]
    for k, vend, miss in sorted(waiting):
        L.append(f"| {k} | {esc(vend)} | {esc(', '.join(miss))} |")
    L += ["", "## Recent activity (last 15 packages)", "", "| Date | PO | Part | Status |", "|---|---|---|---|"]
    for k, p, r in sorted(parts, key=lambda x: x[2]["mtime"], reverse=True)[:15]:
        st = ("Filed, " + (f"{r['sigs']} signature(s)" if r["sigs"] > 0 else "no signature detected")) if r["completed"] \
             else ("In progress" if r["inprog"] else "Archived copy only")
        L.append(f"| {fmt_date(r['mtime'])} | {k} | {esc(p)} | {st} |")
    L += ["", "## How these numbers are counted", "",
          "- **Processed PO / part:** a `PO##-#### Receiving Part N - <part>.pdf` package exists in the PO's `In-Progress Receiving` or `Completed Receiving(s)` folder (archived copies count as processed only).",
          "- **Filed:** the package is in `Completed Receiving(s)` (top level, not Archive). **Signature detected:** the PDF contains a digital-signature dictionary (`/ByteRange`); the count is the number of signatures applied.",
          "- **Test Case Matrix:** struck-through text is ignored, so revised rows count by their current result. Complete = Pass.",
          "- **Not in the matrix:** the PO number doesn't appear in any row's PO Ref column.",
          "- POs from before the automation (no compiled packages) are counted in \"PO folders in scope\" only.",
          "", "Part of [[00 Index]] · [[QE Receiving Workflow]] · [[Attachment A - Test Case Matrix]]", ""]
    return "\n".join(L)

DATA = PM / "_scripts" / ".dashboard_data.json"

def prf_part(name):
    return re.sub(r"(?i)^prf\s*|\s*prf$|\s+rev\b.*$", "", Path(name).stem).strip()

def collect(roots, cache):
    scan_ok = any(os.path.isdir(r) for r in roots)
    pos = scan_pos(roots, cache)
    manifests, edges = scan_manifests(), scan_edge_scripts()
    matrix, refs, merr = scan_matrix()
    oi = scan_open_items() or {"total": 0, "struck": 0, "open": 0}
    out_pos, recent, events = [], [], []
    for k, v in sorted(pos.items()):
        parts = []
        for name, r in v["parts"].items():
            st = "filed" if r["completed"] else ("progress" if r["inprog"] else "archived")
            parts.append({"part": name, "status": st,
                          "compiled": r["inprog"] or r["completed"] or r["mtime"],
                          "filed": r["completed"], "mtime": r["mtime"],
                          "file": r["path_completed"] or r["path_inprog"] or r["path_any"],
                          "prf": next((f for f in v.get("prf_files", []) if name.lower() in Path(f).name.lower()
                                       and f.lower().endswith(".xlsx")), "")})
            recent.append({"po": k, "vendor": v["vendor"], "part": name, "status": st, "mtime": r["mtime"]})
            events.append((r["inprog"] or r["completed"] or r["mtime"], r["completed"]))
        waiting = sorted(p for p in {prf_part(f) for f in v["prfs"]}
                         if p and not any(p.lower() in q.lower() for q in v["parts"]))
        if not parts and not waiting:
            continue
        if parts:
            done = [x for x in parts if x["status"] == "filed"]
            status = "filed" if len(done) == len(parts) else ("partial" if done else "progress")
        else:
            status = "waiting"
        how = sorted(manifests.get(k, set()))
        if k in edges:
            how.append("edge-case script")
        out_pos.append({"po": k, "vendor": v["vendor"], "folder": v["folder"],
                        "path": os.path.join(v.get("root", roots[0]), v["folder"]),
                        "status": status, "parts": parts, "prf_waiting": waiting,
                        "prf_files": sorted(v.get("prf_files", [])),
                        "filed": sum(1 for x in parts if x["status"] == "filed"),
                        "in_matrix": k in refs, "how": how,
                        "last": max([x["mtime"] for x in parts] or [0])})
    proc = [p for p in out_pos if p["parts"]]
    allparts = [x for p in proc for x in p["parts"]]
    # weekly throughput, last 16 ISO weeks
    today = dt.date.today()
    weeks = []
    for w in range(15, -1, -1):
        start = today - dt.timedelta(days=today.weekday() + 7 * w)
        weeks.append({"week": start.isoformat(), "compiled": 0, "filed": 0})
    def wk(ts):
        d = dt.date.fromtimestamp(ts)
        return (d - dt.timedelta(days=d.weekday())).isoformat()
    idx = {w["week"]: w for w in weeks}
    for comp, filed in events:
        if comp and wk(comp) in idx: idx[wk(comp)]["compiled"] += 1
        if filed and wk(filed) in idx: idx[wk(filed)]["filed"] += 1
    by = {}
    for r in matrix:
        by[r["result"]] = by.get(r["result"], 0) + 1
    must = [r for r in matrix if r["must"]]
    summary = {
        "po_folders": len(pos), "po_processed": len(proc), "parts": len(allparts),
        "po_filed": sum(1 for p in proc if p["status"] == "filed"),
        "po_partial": sum(1 for p in proc if p["status"] == "partial"),
        "po_progress": sum(1 for p in proc if p["status"] == "progress"),
        "parts_filed": sum(1 for x in allparts if x["status"] == "filed"),
        "parts_awaiting": sum(1 for x in allparts if x["status"] == "progress"),
        "po_prf_waiting": sum(1 for p in out_pos if p["prf_waiting"]),
        "tc_total": len(matrix), "tc_pass": by.get("Pass", 0), "tc_by": by,
        "must_total": len(must), "must_pass": sum(1 for r in must if r["result"] == "Pass"),
        "not_in_matrix": sum(1 for p in proc if not p["in_matrix"]),
        "oi_open": oi["open"], "oi_total": oi["total"],
    }
    return {"generated": dt.datetime.now().isoformat(timespec="seconds"), "scan_ok": scan_ok,
            "roots": roots, "matrix_error": merr, "summary": summary, "pos": out_pos,
            "matrix": matrix, "weekly": weeks,
            "recent": sorted(recent, key=lambda x: x["mtime"], reverse=True)[:40]}, (pos, manifests, edges, matrix, refs, merr, oi)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--po-root", action="append", help="PO year folder (repeatable)")
    ap.add_argument("--md", action="store_true", help="also write the old markdown dashboard note")
    a = ap.parse_args()
    roots = a.po_root or DEFAULT_ROOTS
    cache = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}
    data, raw = collect(roots, cache)
    tmp = DATA.with_suffix(".tmp")
    tmp.write_text(json.dumps(data), encoding="utf-8")
    os.replace(tmp, DATA)                     # atomic: the UI never reads a half-written file
    CACHE.write_text(json.dumps(cache), encoding="utf-8")
    if a.md:
        pos, manifests, edges, matrix, refs, merr, oi = raw
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_text(render(pos, manifests, edges, matrix, refs, merr, oi, roots, data["scan_ok"]), encoding="utf-8")
    s = data["summary"]
    print(f"Dashboard data refreshed: {s['po_folders']} PO folders, {s['po_processed']} processed, "
          f"{s['parts']} parts, {s['tc_total']} test cases -> {DATA}")

if __name__ == "__main__":
    sys.exit(main())
