---
tags: [code, script]
---

# prf_helpers.py

Shared utility functions imported by [[receive.py]] (and the [[Edge-Case Scripts]]). **~1,805 lines** (was ~580 as of 2026-07-22 — more than tripled by the 2026-09-11 checkbox-integrity/no-clobber/logo-hardening/text-clipping additions). Controlled alongside receive.py; **lives at `Desktop\Receiving Task Automation Tools\prf_helpers.py`**, read/write, not the read-only controlled copy the vault previously described — see [[Folder Structure]].

## Function map

| Function | Purpose |
|---|---|
| `_cb_font`, `check_box`, `uncheck_box` | Draw ☑/☐ glyphs into a cell — the rendering-only checkbox fix, see [[Step 8 - Fill PRF Sections 2, 5, 6]] |
| `ChecklistCheckboxError` (class) | Raised when a checkbox's true state can't be verified — blocks the run rather than guessing |
| `verify_checkbox_fidelity` | Pre-save/post-save gate confirming what was written matches what was intended |
| `save_preserving_drawings` | Saves the workbook without losing embedded drawings/images openpyxl would otherwise strip |
| **`restore_form_controls`, `set_native_checkboxes`, `verify_native_checkboxes`, `verify_xlsx_package_integrity`** (all new 2026-09-11) | Native VML/ctrlProp checkbox integrity — `save_preserving_drawings()` used to carry these zip parts forward without updating the sheet XML's own references to them, so Excel silently pruned them (and the ACME Medical letterhead logo) the next time a human opened+saved the file. These four keep the native, clickable Excel checkboxes intact through a save, not just the rendered ☑/☐ glyphs. See [[Open Items (Master List)#Item 19]]'s sibling fix and [[Changelog Timeline#2026-09-11]]. |
| **`UserContentError`, `assert_no_user_edits`** (new 2026-09-11) | No-clobber guard — refuses to silently overwrite a PRF a human has since hand-edited; `receive.py` gained a `--force` flag to override deliberately. |
| **`LogoMissingError`** (new 2026-09-11) | Raised by `export_and_stamp_pdf()` when neither the PRF's own letterhead image nor the fallback template has a usable logo. `receive.py`'s `process_part()` catches this and BLOCKS just that one part rather than crashing the whole PO run. |
| `_load_font`, `export_and_stamp_pdf` | xlsx → PDF export via LibreOffice headless. `FALLBACK_LOGO_TEMPLATE` constant corrected 2026-09-11 to the real file confirmed on disk (`templates\PRF template CONTROLS.xlsx`) — the constant previously named a guessed filename that doesn't exist. |
| `_find_ghostscript`, `_gs_redistill`, `compile_pdf` | PDF merge + Ghostscript re-distill for a fully editable, unlocked compiled PDF — see [[Step 3 - Compile the Receiving Package]] |
| `_is_lot_subrow`, `_is_received_row`, `_has_receiving_date`, `_bold_line_flags`, `highlight_receiving_line` | Full-row receiving-list highlight — content + gap boundary test + bold-line hard boundary (2026-09-08), see [[Step 2 - Match & Classify Documents]] |
| `_is_sterile_output_row`, `_parse_recv_date` (2026-09-11, first-pass fix) | Sterilization round-trip keyword/date disambiguation — kept as a secondary fallback after the real fix below; see [[Open Items (Master List)#Item 24]] |
| **`highlight_receiving_line`'s new `part_number` parameter** (2026-09-11, the real fix) | When a lot-number search yields 2+ row candidates and NONE qualifies as received (the actual PO26-1009 shared-lot-number failure), retries the search using the part number instead — verified against the real `process_PO26-1009.py`'s working logic and real row coordinates. |
| **`TextClippingError`, `check_row_text_clipping`** (new 2026-09-11) | General-purpose geometric text-clipping detector (rendered text bottom vs. a cell's own row boundary). Written but NOT wired to a real Section 4 cell — see [[Open Items (Master List)#Item 22]]. |
| `normalize_checkbox_marks` | Related checkbox-mark normalization utility |

## ~~Not defined here (documented gap)~~ — obsoleted 2026-10-05

~~`write_manifest()` and `LOGS_DIR` — [[WI-7.4.3-01 Overview]] describes both as existing, importable from this file. **Neither exists** — reconfirmed 2026-09-08, which is also what surfaced that the whole `_Automation\logs\` destination they'd write to doesn't exist either. See [[Step 9 - Run Manifest (Desktop Dev Aid)]].~~

*Both have existed since 2026-09-17: `LOGS_DIR = TOOLS_DIR / "logs"` (the Desktop tools folder), and `write_manifest(po_number, part_number, wi_revision, data)` writes `<PO#>_<part#>_<YYYYMMDD-HHMM>_manifest.json` there. Location ratified 2026-10-05. See [[Step 9 - Run Manifest (Desktop Dev Aid)]].*

## Checkbox cell resolution (the hardest part of this file)

`_read_vml_checkbox_states()` resolves a checkbox's answer cell via, in order: `<x:Row>`/`<x:Column>` → `<x:FmlaLink>` → `<x:Anchor>` → then snaps to a containing **merged cell's** top-left address if applicable (`_parse_cell_ref`, `_parse_merge_ranges`, `_snap_to_merge_anchor`). Reads every VML file linked to a sheet, not just the first. A cell resolvable by none of these methods correctly raises `ChecklistCheckboxError` — this is by design, not a gap to patch with a fourth guessing fallback. See [[Open Items (Master List)#Item 7]] and [[Changelog Timeline]] for the real-file history behind this.

## Depends on

`openpyxl`, `fitz` (PyMuPDF), Ghostscript binary (external, see [[Tools and Dependencies]]), LibreOffice (external, via `soffice` headless call)

## Depended on by

- [[receive.py]] (explicit named import, listed in [[WI-7.4.3-01 Overview]] as a required dependency — the import list itself was corrected 2026-09-11, see [[CLAUDE.md]]'s `REQUIRED DEPENDENCY` block)
- [[Edge-Case Scripts]] (most of the ~14 per-PO scripts reuse these helpers directly, several use `highlight_row()`-style logic of their own that was the actual source for the 2026-09-11 `part_number` fix above)
