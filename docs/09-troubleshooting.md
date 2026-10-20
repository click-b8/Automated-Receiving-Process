# 09 · Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `UnicodeEncodeError` on start (old builds) | Windows cp1252 console | Current `main()` forces UTF-8. Otherwise `set PYTHONUTF8=1` |
| `[ERROR] Required folder missing: COCs/` | Run from the wrong folder, or the subfolder is named differently | `cd` into the PO folder. Rename to the WI's names, or add a fallback in the `F_*` constants |
| `PO document not found` | PO PDF lacks the PO number in its filename | Rename to `26-1085 - Vendor.pdf`. Vendor sales-order acknowledgements are intentionally ignored |
| Run stops: *PO revision conflict* | Two PO copies disagree on rev or qty | Decide which governs, then move the other to `Archive\` |
| Part shows **Pending** | No PRF found | Name it `PRF <part#> Rev <X>.xlsx` in `Inspection Forms\` |
| Part **HELD** | Multi-part PO, this part has no documents of its own | Get the vendor paperwork |
| Part **BLOCKED**: Section 1 mismatch | PRF Section 1 ≠ sources | Fix whichever is wrong (PRF or source). The tool never picks one |
| `UserContentError` | The package was edited after the last run | Keep your edits, or re-run with `--force` on explicit instruction |
| `ChecklistCheckboxError` | A checkbox's linked cell couldn't be resolved, or the read-back didn't match | Don't patch around it. The PRF template may differ, so check the cell map |
| `LogoMissingError` | PRF and fallback template have no letterhead image | Put the controlled template in `templates\` |
| Checkboxes in Excel aren't clickable / logo vanished after saving | A PRF saved by an old build, then re-saved in Excel | Re-create from the template and re-run. Check other PRFs from the same period |
| Receiving-list highlight covers the wrong row | Shared lot numbers / sterilization round-trip rows | The tool retries by part number. Verify visually and report the layout |
| `LibreOffice not found` | Not installed / not on PATH | Install it. The tool also checks the default install paths |
| Compiled PDF can't be signed | Ghostscript not found, so vendor encryption survived the merge | Install Ghostscript, or drop a `gs10.x\` folder next to the tool |
| Dashboard: *Box not reachable* | Box Drive not running, or a different year folder | Start Box Drive, or set **PO year folder** in settings |
| Dashboard: first refresh takes minutes | Box Drive downloading every filed package once to count signatures | Expected. Later scans use the size+mtime cache |
| Dashboard: *monday error* | Bad token / board ID / API version | Re-paste the token. Check the board ID from its URL |
| DHR not set automatically | See the *not changed automatically* list on the monday tab | Usually: several items match, wrong group, Stuck status, or Box scan older than 30 min |
| Graph view is a hairball | Hubs not rebuilt after new runs | `python "Project Map\_scripts\refresh_file_map.py"` |
