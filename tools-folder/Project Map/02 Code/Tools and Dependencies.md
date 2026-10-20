---
tags: [code, dependencies]
---

# Tools and Dependencies

Full detail lives in `requirements.txt` at the vault root; summarized here.

## Python packages (`pip install -r requirements.txt`)

| Package | Used for |
|---|---|
| `anthropic>=0.40.0` | Claude API calls — [[Step 4 - Extract Data from Documents|extraction]], [[Step 5 - Parse Drawing with Cache|drawing parsing]], [[Step 6.5 - Vendor Document Completeness Check|completeness check]], [[Step 6.6 - Vision-Based Key-Field Cross-Check|vision cross-check]] |
| `pdfplumber>=0.11.0` | PDF text extraction — primary, zero-token path, tried before any model call |
| `openpyxl>=3.1.0` | Excel PRF read/write |
| `pymupdf>=1.24.0` (imports as `fitz`) | PDF merge, annotation, highlighting, page rendering for vision calls |

Python 3.10+ required.

## System dependencies (not pip)

- **LibreOffice** — required for PRF xlsx → PDF export (`export_and_stamp_pdf()` in [[prf_helpers.py]]). Installer present at vault root: `LibreOffice_26.2.3_Win_x86-64.msi`.
- **Ghostscript** — required for fully editable compiled PDFs (strips security flags/encryption/AcroForm locks — `compile_pdf()` in [[prf_helpers.py]]). Local install present at vault root: `gs10.07.1/`.

## Fallback letterhead template — corrected 2026-09-11

`FALLBACK_LOGO_TEMPLATE` in [[prf_helpers.py]] (used by `export_and_stamp_pdf()` when a PRF's own letterhead image is missing, raising [[Open Items (Master List)#Item 22|`LogoMissingError`]] if even the fallback has none) previously named a **guessed** filename that doesn't exist on disk. Confirmed via direct device listing: the only file actually present in `templates\` is `PRF template CONTROLS.xlsx` — the constant now points at that real file. Not yet confirmed that this file's own embedded image actually IS the ACME Medical logo (only the filename was confirmed, not the file's contents) — worth a direct check before relying on it in production.

## Anthropic API key

Required in environment as `ANTHROPIC_API_KEY` for every Haiku/Sonnet call the script makes. Setup instructions in [[Manual vs Claude-Assisted Workflow]].

## PATH setup (optional, one-time)

`requirements.txt` documents adding the tools folder to PATH so `receive.py` can be invoked from any PO folder without a full path — see the file directly for the PowerShell/bash snippets.

## Non-controlled artifacts (do not treat as source)

- `parts_cache.json` — auto-created cache, keyed by part#+revision, see [[Step 5 - Parse Drawing with Cache]]. Do not hand-edit.
- `qf_7.2-01_apdl_catalog.json` — new 2026-09-11, the 434-entry APDL catalog the disposition rule now looks up against. See [[Open Items (Master List)#Item 13]].
- `__pycache__/` — Python bytecode cache, ignorable.

## Used by

- [[receive.py]], [[prf_helpers.py]], [[Edge-Case Scripts]] all draw on this same dependency set
