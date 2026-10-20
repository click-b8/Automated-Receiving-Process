"""Synthetic demo data for the Receiving Dashboard screenshots.
Every PO, vendor, part, lot and person here is invented."""
import json, random, datetime as dt
from pathlib import Path

random.seed(7)
OUT = Path(__file__).parent / "demo"
OUT.mkdir(exist_ok=True)
NOW = dt.datetime(2026, 10, 20, 10, 30)
VENDORS = ["Northwind Machining", "Contoso Sterile Pack", "Fabrikam Biologics", "Litware Anodizing",
           "Tailspin Polymers", "Adatum Instruments", "Proseware Sterilization", "Wingtip Swiss Turn"]
ROOT = r"C:\Users\<you>\Box\Shared\Purchase Orders\2026"

import calendar
def ts(d): return calendar.timegm(d.timetuple()) + d.microsecond / 1e6

pos, recent = [], []
pn = 1100
for i in range(64):
    po = f"PO26-{1201 + i}"
    vendor = random.choice(VENDORS)
    n = random.choice([1, 1, 1, 2, 2, 3, 4])
    age = (63 - i) * 1.6 + random.random()
    base = NOW - dt.timedelta(days=age)
    r = random.random()
    if age > 40: status = "filed" if r > 0.06 else "progress"
    else: status = "filed" if r < 0.42 else ("progress" if r < 0.74 else ("partial" if r < 0.87 else "waiting"))
    if status == "partial" and n == 1: n = 2
    parts = []
    for k in range(n):
        pn += random.randint(1, 9)
        part = f"SR90-{pn:04d}"
        comp = ts(base + dt.timedelta(hours=2 * k))
        if status == "filed" or (status == "partial" and k == 0):
            pst, filed = "filed", min(comp + random.randint(1, 5) * 86400, ts(NOW) - 3600)
        elif status == "waiting":
            continue
        else:
            pst, filed = "progress", None
        parts.append({"part": part, "status": pst, "compiled": comp, "filed": filed, "mtime": filed or comp,
                      "file": rf"{ROOT}\{po[2:]} - {vendor}\{'Completed Receivings' if pst == 'filed' else 'In-Progress Receiving'}\{po} Receiving Part {k+1} - {part}.pdf",
                      "prf": ""})
    waiting = [f"PRF SR90-{pn + 50:04d} Rev A.xlsx"] if status == "waiting" else []
    how = ["receive.py"] if random.random() < 0.8 else ["edge-case script"]
    pos.append({"po": po, "vendor": vendor, "folder": f"{po[2:]} - {vendor}", "path": rf"{ROOT}\{po[2:]} - {vendor}",
                "status": status, "parts": parts, "filed": sum(1 for p in parts if p["status"] == "filed"),
                "prf_waiting": waiting, "prf_files": [], "how": how if parts else [],
                "in_matrix": random.random() < 0.35, "last": max([p["mtime"] for p in parts] or [ts(base)])})
    for p in parts:
        recent.append({"po": po, "vendor": vendor, "part": p["part"], "status": p["status"], "mtime": p["mtime"]})
recent.sort(key=lambda x: -x["mtime"])

scen = [
    ("Happy path — single part, single batch", True), ("Multi-part PO, single batch", True),
    ("Multi-batch / partial receiving (different receiving dates)", True),
    ("Same part + revision across two batches — lot-suffix disambiguation", True),
    ("Contract-manufacturer off-the-shelf part (traveler present)", True),
    ("Contract-manufacturer off-the-shelf part, no traveler provided", False),
    ("Confidential biologic batch records — Director review before QM", True),
    ("Drawing revision mismatch (ordered vs. received)", True), ("No engineering drawing on file", True),
    ("DVT-accepted part (no dimensional inspection)", True),
    ("Manufacturer inspection report shows an out-of-spec dimension", True), ("Missing required COC", True),
    ("Multiple materials/finishes requiring multiple COCs", True), ("File-access boundary — within one PO", True),
    ("PRF Section 1 mismatch (engineer pre-fill error)", True), ("Run manifest content and placement", False),
    ("Sibling-PO isolation", True), ("PO#/year mismatch", True), ("Native-checkbox integrity after a plain run", True),
    ("Multi-part PO, one line held for missing documents", True), ("Quantity short of ordered, within tolerance", True),
    ("Two PO revisions on file with conflicting qty", True), ("Sterilization-only service vendor, no traveler", False),
    ("Disposition via APDL lookup", True), ("Stadium/oval critical-dimension callouts", True),
    ("Re-run against a hand-corrected PRF (no-clobber guard)", True),
    ("Section 5 text drafted per the Reasoning Reference", True), ("Compiled before packing slip exists", True),
    ("Scanned packing slip — signature read by vision", True), ("COC bundled inside packing-slip PDF", True),
    ("Drawing exists only inside the PO PDF", True), ("Vendor sales-order acknowledgement in PO folder", True),
    ("Explicitly unsterilized lot from an always-sterile vendor", True), ("Sterile part with no sterilization COC", True),
    ("Vendor paperwork naming a different customer or PO", True), ("COC material differs from drawing callout", True),
    ("Large multi-part PO (10+ parts) signed off part by part", True), ("Rework shipment: vendor lot ≠ engraved lot", True),
]
results = ["Pass"] * 30 + ["Not run"] * 5 + ["On hold", "Proposed", "Re-test required"]
random.shuffle(results)
matrix = [{"id": f"TC-{i+1:02d}", "scenario": s, "must": m, "result": results[i],
           "ref": f"{pos[i % len(pos)]['po'][2:]}" if results[i] == "Pass" else ""} for i, (s, m) in enumerate(scen)]

weekly = []
for w in range(16):
    d = (NOW - dt.timedelta(weeks=15 - w))
    d -= dt.timedelta(days=d.weekday())
    c = random.randint(4, 22); weekly.append({"week": d.strftime("%Y-%m-%d"), "compiled": c, "filed": max(0, c - random.randint(-3, 6))})
weekly[-1]["filed"] = 2

parts_all = [p for x in pos for p in x["parts"]]
S = {"po_folders": 152, "po_processed": sum(1 for p in pos if p["parts"]),
     "po_filed": sum(1 for p in pos if p["status"] == "filed"), "po_partial": sum(1 for p in pos if p["status"] == "partial"),
     "po_progress": sum(1 for p in pos if p["status"] == "progress"), "po_prf_waiting": sum(1 for p in pos if p["prf_waiting"]),
     "parts": len(parts_all), "parts_filed": sum(1 for p in parts_all if p["status"] == "filed"),
     "parts_awaiting": sum(1 for p in parts_all if p["status"] == "progress"),
     "not_in_matrix": sum(1 for p in pos if p["parts"] and not p["in_matrix"]),
     "tc_total": len(matrix), "tc_pass": sum(1 for r in matrix if r["result"] == "Pass"),
     "must_total": sum(1 for r in matrix if r["must"]), "must_pass": sum(1 for r in matrix if r["must"] and r["result"] == "Pass"),
     "tc_by": {}, "oi_open": 14, "oi_total": 31}
for r in matrix: S["tc_by"][r["result"]] = S["tc_by"].get(r["result"], 0) + 1
data = {"generated": NOW.strftime("%Y-%m-%dT%H:%M:%S"), "scan_ok": True, "roots": [ROOT], "summary": S,
        "pos": pos, "recent": recent[:40], "weekly": weekly, "matrix": matrix, "matrix_error": ""}
(OUT / "dashboard_data.json").write_text(json.dumps(data, indent=1))

# ---- monday board (shape of the GraphQL response the plugin requests) ----
people = ["QE Engineer", "Ops Receiving", "Quality Manager", "Engineering Reviewer"]
groups = ["New inspection", "Inspection", "Product Release", "DHR", "Put away"]
items, logs, updates = [], [], []
iid = 50000000001
reasons = ["Missing documents", "Dimensional out of spec", "Material / finish discrepancy", "Quantity discrepancy",
           "Signature logistics", "Vendor document error", "PRF correction", "Testing failure / root cause"]
issue_text = {"Missing documents": "Emailing vendor for COC, no packing slip in box",
              "Dimensional out of spec": "Dim 3 out of spec, disposition request sent to Engineering",
              "Material / finish discrepancy": "COC finish callout differs from drawing (passivation)",
              "Quantity discrepancy": "Qty short 4 pcs vs PO, waiting on vendor confirmation",
              "Signature logistics": "Package too large to email, send via Box link",
              "Vendor document error": "Traveler has wrong manufacture date",
              "PRF correction": "Section 1 lot number correction",
              "Testing failure / root cause": "Bench test broke at 2 of 10, root cause opened"}
def mk_cv(status, person, date, waiting="", reason="", text=""):
    vals = {"Status": status, "Person": person, "Date": date, "Notes": "", "Waiting on": waiting, "Delay reason": reason, "Text": text}
    return [{"id": k.lower().replace(" ", "_"), "text": v, "column": {"title": k}} for k, v in vals.items()]
for i, p in enumerate(pos[-34:]):
    iid += 1
    g = random.choice(groups[:3]) if p["status"] != "filed" else random.choice(["Product Release", "DHR", "Put away", "DHR"])
    st = {"New inspection": "Ready to start", "Inspection": random.choice(["Working on it", "Stuck"]),
          "Product Release": random.choice(["Waiting for Signature", "Working on it", "Stuck"]),
          "DHR": "DHR", "Put away": "Done"}[g]
    created = NOW - dt.timedelta(days=random.randint(3, 40))
    reason = random.choice(reasons) if st == "Stuck" and random.random() < 0.7 else ""
    waiting = random.choice(["Quality Manager", "Engineering Reviewer", "Director of Operations", ""]) if st == "Waiting for Signature" else ""
    items.append({"id": str(iid), "name": f"PO: {p['po'][2:]} {p['vendor'].split()[0]} - ({random.choice(['Screws', 'Trays', 'Cannulas', 'Syringes', 'Inserters', 'Caps'])})",
                  "url": f"https://example.monday.com/boards/1234567890/pulses/{iid}", "created_at": created.isoformat() + "Z",
                  "updated_at": NOW.isoformat() + "Z", "group": {"title": g},
                  "column_values": mk_cv(st, random.choice(people), (created + dt.timedelta(days=3)).strftime("%Y-%m-%d"), waiting, reason)})
    # status history: a couple of past Stuck episodes per item, then current status
    t = created
    hist = []
    for _ in range(random.randint(0, 2)):
        t += dt.timedelta(days=random.uniform(0.5, 3)); hist.append(("Stuck", t))
        t += dt.timedelta(days=random.uniform(0.5, 6)); hist.append(("Working on it", t))
    if st in ("Stuck", "Waiting for Signature"):
        t += dt.timedelta(days=random.uniform(0.3, 2)); hist.append((st, t))
    prev = "Ready to start"
    for to, when in hist:
        logs.append({"data": json.dumps({"pulse_id": int(iid), "value": {"label": {"text": to}}, "previous_value": {"label": {"text": prev}}}),
                     "created_at": str(int(ts(when) * 1e7))})
        if to == "Stuck":
            r_ = reason or random.choice(reasons)
            updates.append({"item_id": str(iid), "created_at": (when + dt.timedelta(minutes=5)).isoformat() + "Z",
                            "text_body": f"STUCK - {p['po']}\n\nIssue: {issue_text[r_]} - Auto"})
        prev = to
monday = {"data": {"me": {"name": "Demo User"}, "boards": [{"name": "Receiving – New Inspection (demo)",
          "url": "https://example.monday.com/boards/1234567890", "groups": [{"id": g, "title": g} for g in groups],
          "items_page": {"items": items}, "activity_logs": logs, "updates": updates}]}}
(OUT / "monday_response.json").write_text(json.dumps(monday, indent=1))
dhr = [{"time": (NOW - dt.timedelta(hours=h)).isoformat() + "Z", "po": p["po"], "item": f"PO: {p['po'][2:]} {p['vendor'].split()[0]}",
        "itemId": "0", "from": "Waiting for Signature", "parts": [x["part"] for x in p["parts"]], "account": "Demo User", "result": "DHR set"}
       for h, p in zip([30, 52, 75], [x for x in pos if x["status"] == "filed"][-3:])]
(OUT / "dhr_log.jsonl").write_text("\n".join(json.dumps(x) for x in dhr) + "\n")
print(json.dumps(S, indent=0)[:400])
