---
tags: [governance, sandbox]
---

# Folder Structure — Flat Tools Folder + Per-PO Purchase Order Tree

Part of [[WI-7.4.3-01 Overview]]. **REWRITTEN 2026-09-11** — see
[[File Access Rules]] for why the previous year-rooted `_Automation_Master`
diagram was fictional. This page now matches [[CLAUDE.md]]'s corrected
FILE ACCESS RULES section AND a direct `device_list_dir` listing of the real
Desktop folder taken 2026-09-11.

```
Desktop\Receiving Task Automation Tools\        <- the tool's real home, flat
|-- receive.py                                     READ/WRITE (not read-only)
|-- prf_helpers.py                                 READ/WRITE
|-- CLAUDE.md                                      READ/WRITE (this WI, = the vault root doc)
|-- receive.bat                                    thin Windows wrapper
|-- requirements.txt
|-- parts_cache.json                               drawing-parse cache (auto-created)
|-- qf_7.2-01_apdl_catalog.json                    APDL catalog (added 2026-09-11, see Open Item #13)
|-- __pycache__\                                   Python bytecode cache, ignorable
|-- templates\                                     fallback letterhead template(s) — see [[Tools and Dependencies]]
|-- gs10.07.1\                                     local Ghostscript install
|-- LibreOffice_26.2.3_Win_x86-64.msi               LibreOffice installer
|-- Validation\                                     QF 7.4.3-06 package (Protocol, Attachment A/B .docx)
|-- Project Map\                                    <- this vault
|-- process_PO26-1016.py, process_PO26-1016_part23.py, ...   see [[Edge-Case Scripts]] —
|                                                       now ~14 per-PO scripts, not 2
|-- inspect_annots.py                               diagnostic, not production

...\Purchase Orders\                 <- separate tree, the Box PO structure
|-- <year>\                                          e.g. 2026\
|     |-- <PO folder>\                               e.g. "26-1005 - SterilizerCo\"
|     |     |-- Shipping Docs\        (PO doc, packing slip, receiving list)
|     |     |-- COCs\                 (external mfr COCs + travelers)
|     |     |-- Drawings\             (primary; fallback: Drawings & Stepfiles\)
|     |     |-- Inspection Forms\     (pre-filled PRF xlsx + pdf)
|     |     |-- In-Progress Receiving\  (compiled PDFs, auto-created)
|     |     |-- Archive\, Completed Receiving\, Invoices\, Quote\  (IGNORE on PO runs; 2026-10-05: Completed Receiving\ is the read-only source for the Reasoning Reference harvest)
|     |-- Template\, Unstored COC's\, .claude\        (also present, IGNORE for a normal run)
```

**There is no `_Automation_Master\`, no `_Automation\tools\`, no
`_Automation\logs\`, and no year-rooted copy of the tools.** One flat
Desktop folder serves every year and every PO. See
[[Open Items (Master List)#Item 8]] for the traceability-manifest
consequence of this ~~(no manifest is ever written — the old diagram claimed
one would land in `_Automation\logs\`, which doesn't exist).~~
*2026-10-05: manifests have been written since 2026-09-17, to
`Desktop\Receiving Task Automation Tools\logs\`, which is now the decided
location — see [[Step 9 - Run Manifest (Desktop Dev Aid)]].*

RECEIVING LIST search order: PO folder root first, then `Shipping Docs\`.
DRAWINGS: `Drawings\` first, fall back to `Drawings & Stepfiles\`.

## Output per part

```
<PO folder>\In-Progress Receiving\PO##-#### Receiving Part X - Part#.pdf
<PO folder>\Inspection Forms\PRF Part# Rev X.xlsx   (updated in place)
<PO folder>\Inspection Forms\PRF Part# Rev X.pdf    (exported by automation)
```

~~No manifest file is written anywhere — see~~
~~[[Step 9 - Run Manifest (Desktop Dev Aid)]].~~

*2026-10-05: one manifest per part per run is written OUTSIDE the PO folder, to
`Desktop\Receiving Task Automation Tools\logs\<PO#>_<part#>[_Lot<lot>]_<YYYYMMDD-HHMM>_manifest.json`
— a development aid, not part of the output package. See [[Step 9 - Run Manifest (Desktop Dev Aid)]].*

## See also

- [[File Access Rules]] — the corrected access boundary this diagram now matches
- [[Open Items (Master List)#Item 2]] — whether a real `_Automation_Master`
  controlled deployment should still be built, or whether the WI should
  simply be revised to describe this flat-folder practice permanently
