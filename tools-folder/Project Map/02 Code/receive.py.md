---
tags: [code, script]
---

# receive.py

Main automation script. **~3,600 lines** (was ~2,490 as of 2026-07-22 — grew
substantially in the 2026-09-11 batch). Version marker `SCRIPT_VERSION =
"1.0"` at the top still does not track [[WI-7.4.3-01 Overview|WI revisions]]
— still worth noting as a documentation gap. Imports from
[[prf_helpers.py]] — never reimplements its functions inline. **Lives at
`Desktop\Receiving Task Automation Tools\receive.py`** — see
[[Folder Structure]] (corrected 2026-09-11; this was previously described as
living under a nonexistent `_Automation_Master\tools\`).

## Model tiering

```python
MODEL = {
    "classify":  "claude-haiku-4-5-20251001",
    "extract":   "claude-haiku-4-5-20251001",
    "checklist": "claude-sonnet-4-6",
    "narrative": "claude-sonnet-4-6",
    "complex":   "claude-sonnet-4-6",
}
```
[[Step 7 - Section 2 Determination (Fixed Rules)|Section 2 determinations]] are **never** made by either model — pure Python functions only.

## Function map, grouped by WI step

| Group | Functions | Step |
|---|---|---|
| Cache | `load_cache`, `save_cache`, `cache_key` | [[Step 5 - Parse Drawing with Cache]] |
| Claude API | `call_claude`, `_get_client` | (shared) |
| Text extraction | `extract_text`, `extract_page_text` | [[Step 4 - Extract Data from Documents]] |
| Folder resolution | `find_drawings_dir`, `find_receiving_list`, `find_prf`, `find_po_document`, **`find_po_documents_all`** (new 2026-09-11) | [[Step 1 - Verify Folders and Identify Parts]], [[Step 2 - Match & Classify Documents]] |
| Page splitting/classification | `classify_page`, `split_pdf_by_page_type`, `_classify_scanned_combined_via_vision`, `classify_coc`, `_vision_classify_page` (2026-09-08) | [[Step 2 - Match & Classify Documents]] |
| Per-part document scoping | **`_filter_classified_for_part`** (new 2026-09-11) | [[Step 2 - Match & Classify Documents]] |
| Completeness check | `check_document_completeness` | [[Step 6.5 - Vendor Document Completeness Check]] |
| Data extraction | `extract_po_data`, `extract_packing_data`, `extract_receiving_data`, `extract_coc_data`, `_parse_json`, `_parse_json_dict` | [[Step 4 - Extract Data from Documents]] |
| Drawing parse | `parse_drawing` | [[Step 5 - Parse Drawing with Cache]] |
| Cross-reference | `cross_reference` (now includes a 10% shortage-tolerance branch, 2026-09-11) | [[Step 6 - Cross-Reference Source Documents]] |
| Vision key-field check | `_vision_extract_key_fields`, `check_key_field_consistency` | [[Step 6.6 - Vision-Based Key-Field Cross-Check]] |
| Invoice stripping | `strip_invoice_pages` (2026-07-2x) | Step 3 (packing slip filtering) |
| Section 2 logic | `_part_is_sterile`, `determine_section2` (2.4 default + disposition rewritten 2026-09-11) | [[Step 7 - Section 2 Determination (Fixed Rules)]] |
| Drawing callouts | `_find_callout_coords`, `_find_stadium_oval_coords` (new 2026-09-11), `annotate_drawing_callouts` | [[Step 3 - Compile the Receiving Package]] |
| Checkbox XML plumbing | `_col_to_letter`, `_resolve_sheet_xml_path`, `_parse_cell_ref`, `_parse_merge_ranges`, `_snap_to_merge_anchor`, `_extract_vml_caption`, `_read_vml_checkbox_states`, `_anchor_cell` | [[Step 8 - Fill PRF Sections 2, 5, 6]] |
| Calibration check | `check_calibration_date`, `_extract_date` (2026-07-2x, cell fixed 2026-09-08) | Step 3 cross-check |
| PRF fill/export | `fill_prf`, `extract_po_pages_only` | [[Step 8 - Fill PRF Sections 2, 5, 6]], [[Step 3 - Compile the Receiving Package]] |
| PO revision conflict | new gate inside `main()`, using `find_po_documents_all` (2026-09-11) | [[Step 6 - Cross-Reference Source Documents]] |
| Partial-PO hold | new block inside `process_part()`, returns `"held"` status (2026-09-11, corrected same day) | [[Step 2 - Match & Classify Documents]] |
| Orchestration | `process_part` (new `all_part_numbers` param, `LogoMissingError` handling, `"held"` status), `main` (UTF-8 console fix, PO-revision gate) | all steps |

## Constants of note

- `BIOLOGIC_VENDORS = {VendorB, VendorA, VendorC}`, `VENDORD_VENDORS = {VendorD}`, `TRAVELER_VENDORS`, `BIOLOGIC_KEYWORDS` — see [[Step 7 - Section 2 Determination (Fixed Rules)]]
- `F_SHIPPING`, `F_COCS`, `F_DRAWINGS`/`F_DRAWINGS2`, `F_FORMS`, `F_OUTPUT` — folder-name constants, see [[Folder Structure]]
- `CACHE_FILE`, `TMP_DIR` — resolved from `TOOLS_DIR` (imported from [[prf_helpers.py]])
- **`APDL_FILE`** (new 2026-09-11) — `TOOLS_DIR / "qf_7.2-01_apdl_catalog.json"`, the 434-entry Approved Product Destination List catalog the disposition rule now looks up against, replacing the retired "label" keyword check. See [[Open Items (Master List)#Item 13]].

## `import re as _re` — a real bug that shipped once

The merge-aware checkbox helpers (`_parse_merge_ranges` etc.) use `_re` at module scope, but `receive.py` originally only had function-local `import re as _re` statements — causing a live `NameError` the first time the merge-aware fix ran for real. Fixed by adding `import re as _re` to the top-of-file imports. See [[Changelog Timeline]] and the process lesson it produced: **verify fixes against the real, unmodified file, not just an isolated test copy** (this bug was specifically hidden by an isolated verification script that imported `re` differently than the real file did — the SAME lesson recurred in a new form on 2026-09-11, see [[Changelog Timeline#Recurring lesson across this timeline]]).

## `main()` — 2026-09-11 additions

- Forces UTF-8 on stdout/stderr at the very top, before any `print()` can
  trip a default Windows cp1252 console (`UnicodeEncodeError` on the
  script's own ✓/⏳/⚠ output — found 2026-09-08, actually fixed in code
  2026-09-11).
- New PO-revision-conflict gate right after `find_po_document()`: uses
  `find_po_documents_all()` to find every PO-document candidate, compares
  revisions and per-part-number ordered quantities across all of them, and
  `sys.exit(1)`s the ENTIRE run (not one part) if they disagree — see
  [[Open Items (Master List)#Item 19]] for the deliberate decision to keep
  this stricter than the real bespoke script's corroborate-and-flag
  behavior.
- Computes `all_part_numbers` (every ready part's number on this PO) and
  passes it into every `process_part()` call, enabling the partial-PO hold
  and per-part document scoping.

## Depends on

- [[prf_helpers.py]] (hard dependency, explicit named import — see the
  corrected `REQUIRED DEPENDENCY` block at the top of [[CLAUDE.md]], fixed
  2026-09-11 to list the real current import set instead of the
  never-existed `LOGS_DIR`/`write_manifest`)
- `anthropic`, `pdfplumber`, `openpyxl`, `fitz` (PyMuPDF) — see [[Tools and Dependencies]]

## Used by

- [[Manual vs Claude-Assisted Workflow]] (CLI entry point: `python receive.py PO26-1085`, now run from `Desktop\Receiving Task Automation Tools\` directly — see [[Folder Structure]])
- `receive.bat` (thin Windows wrapper)
- Bypassed entirely by the [[Edge-Case Scripts]] for specific real POs whose circumstances don't fit the general path — now ~14 scripts, not 2 (2026-09-11 scope decision: fix recurring root causes here, keep genuinely one-off circumstances bespoke — see [[Open Items (Master List)#Item EC — Undocumented Edge-Case Scripts|Item EC]])
