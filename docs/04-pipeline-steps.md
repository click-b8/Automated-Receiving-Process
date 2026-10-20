# 04 · Pipeline steps (what `receive.py` actually does)

The numbering follows the work instruction (WI-7.4.3-01 = `tools-folder/CLAUDE.md`). Each step has a matching note in the vault under `Project Map/01 Governance/`.

| # | Step | Kind | Can block? |
|---|---|---|---|
| 1 | Verify folders, identify parts from the PO | deterministic | yes: missing `Inspection Forms\` / `COCs\`, no PO document |
| 2 | Match and classify documents per part | rules first, Haiku on ambiguity | holds a part (`HELD`) on a partial PO |
| 2.5 | Validate PRF Section 1 against sources | deterministic | **yes, per part** |
| 3 | Compile the receiving package | deterministic | no |
| 4 | Extract data from documents | pdfplumber → Haiku | no |
| 5 | Parse the drawing (cached) | cache → Sonnet | no |
| 6 | Cross-reference sources against each other | deterministic | **PO revision conflict halts the whole PO** |
| 6.5 | Vendor document completeness | vision (Haiku), non-blocking | no |
| 6.6 | Vision key-field cross-check | vision, soft flag only | no |
| 7 | Section 2 determination | **fixed Python rules** | no (unresolved → `FLAG`, boxes left empty) |
| 8 | Fill PRF Sections 2, 5, 6 + export | deterministic | yes, per part: checkbox verification, logo, no-clobber |
| 9 | Run manifest | deterministic | no |

---

## Step 1 · Verify folders and identify parts

- `cwd` must be the PO folder. `Inspection Forms\` and `COCs\` must exist.
- `find_po_document()` locates the PO PDF. It skips files that are clearly not a PO copy, such as vendor sales-order acknowledgements and quotes.
- `extract_po_data()` returns PO#, date, revision, vendor and every line item (part#, description, rev, qty). Quote pages are ignored.
- A part is **ready** when its PRF exists (`find_prf()`). Otherwise it's **pending** and listed at the end.

## Step 2 · Match and classify

| Document | Primary | Fallback |
|---|---|---|
| Drawing | `Drawings\` | `Drawings & Stepfiles\`, or drawing pages inside the PO PDF |
| COCs | `COCs\` (every PDF classified) | COC pages bundled in the packing-slip PDF |
| Receiving list | PO root | `Shipping Docs\` |
| Packing slip | `Shipping Docs\` | PO root |
| PRF | `Inspection Forms\` | none |

**Classification is rules first, at zero tokens.** Sterilizer names and "irradiation/gamma/e-beam" → `ster_coc`. "traveler/job traveler/router" → `traveler`. "certificate of conformance/compliance" → `mfg_coc`. Company-generated packing/PO docs → excluded. Ambiguous or scanned pages escalate to Claude Haiku. Combined PDFs are split page by page (`split_pdf_by_page_type()`). A scanned multi-cert bundle is rendered and classified by vision (`_classify_scanned_combined_via_vision()`).

**Vendor override.** Some machine shops name their inspection report "traveler". For any vendor outside `TRAVELER_VENDORS` that document is a *manufacturer inspection report*, so it drives 2.3, not 2.4.

**Per-part scoping.** `_filter_classified_for_part()` keeps a multi-part PO from sharing one part's COCs with another.

**Partial-PO hold.** On a multi-part PO, a part with no documents of its own and no packing slip naming it, while its siblings do have documents, returns `HELD` rather than being compiled with the wrong paperwork.

**Always excluded.** Company-generated COCs and packing slips, quote pages, commercial/proforma invoice pages (`strip_invoice_pages()`), confidential biologic batch records/travelers, and duplicate pages.

## Step 2.5 · Validate PRF Section 1 (blocking gate)

The tool reads Section 1 and **never writes to it**. It reads `M3` product, `M4` part#, `M5` rev, `E3` PO#, `E4` supplier, `E6` qty received and `M6` lot, and compares each against the same fact extracted independently from the sources. On any mismatch:

- **halt this part** (no package, no PRF write); other parts continue
- print the mismatch **above every other flag**, as PRF value vs. source value
- never guess which is right, and never "correct" Section 1

This is the most audit-relevant gate in the tool. It stops a quality record from being drafted on top of a PRF that disagrees with the receiving evidence.

## Step 3 · Compile the package

Output: `In-Progress Receiving\PO26-xxxx Receiving Part N - <part#>.pdf`

**Order (SOP 7.4.3 / Proof Receiving Checklist):**

1. PO (relevant pages only, quote pages stripped)
2. Packing slip(s) (vendor's own, invoice pages stripped)
3. Receiving list, with the **whole** multi-line row for this lot highlighted (`highlight_receiving_line()`)
4. Drawing(s). On a revision mismatch: the **ordered** rev first, then the **received** rev. Numbered callout bubbles go only on the received-rev drawing.
5. PRF (filled, exported, letterhead stamped)
6. Manufacturer COC(s)
7. Sterilization COC
8. Traveler (non-confidential only)
9. Manufacturer inspection report / other vendor QC documents

`compile_pdf()` merges with PyMuPDF, then **re-distills with Ghostscript**. Vendor PDFs often arrive encrypted or form-locked, and a locked package can't be signed in Adobe.

**Callout bubbles.** `_find_callout_coords()` finds critical-dimension callouts drawn as circled digits, parenthesized numbers, or (via vector geometry) hand-drawn stadium/oval shapes. The oval style is drawn in red so it can't be confused with a drawing's own BOM balloons.

## Step 4 · Extract data

`pdfplumber` first. Escalate to Haiku only when the text layer is empty or incomplete.

| Function | From | Fields |
|---|---|---|
| `extract_po_data()` | PO | PO#, date, rev, vendor, line items |
| `extract_packing_data()` | packing slip | vendor, part#, description, qty shipped, lot, **Ops signature present** (vision on scans) |
| `extract_receiving_data()` | receiving list | part#, qty to receive / received, date, location, lot, rev, mfg/exp/ster dates |
| `extract_coc_data()` | mfg / ster COC, mfg report | vendor part#, lot, qty, dates, material/finish, dose results, out-of-spec dims |

## Step 5 · Parse the drawing (with cache)

Key: `"<part#>_Rev<rev>"` in `parts_cache.json` (tools folder, never a PO folder). On a miss, Sonnet reads the title block and every numbered inspection callout (nominal, limits, gaging method). Drawings are controlled and rarely change, so after the first run this is a free lookup.

## Step 6 · Cross-reference

`cross_reference()` compares **sources with each other**:

- part# (PO / packing / receiving / drawing / COC)
- revision ordered vs. received; on a mismatch both drawings are compiled
- lot (packing / receiving / COC)
- quantity: PO vs. received vs. shipped. A **shortage of 10% or less** is a soft `QTY SHORTAGE (WITHIN 10% TOLERANCE)`. Overage or a larger shortage → `QTY MISMATCH`.
- mfg date ≠ sterilization date (same day = CRITICAL)
- COC vendor part# vs. the drawing's supplier part#
- PRF Section 1 vs. every source (soft `MISMATCH`, in addition to the 2.5 gate)
- calibration date: Section 3 next-calibration date vs. inspection date (soft)

**PO revision-conflict halt.** If two PO copies in the folder disagree on revision or a line quantity, the **entire run stops** before any part is processed. Quality chose this over auto-resolving from corroborating documents.

## Step 6.5 · Vendor document completeness (vision)

Every page about to be compiled is rendered and Haiku is asked one narrow question: *is any section that requires an entry completely blank, with no data and no N/A?* For example, an empty post-sterilization QC block. This is non-blocking. If the model call fails, the result degrades to "no issues found".

## Step 6.6 · Vision key-field cross-check

For scanned documents where text extraction returned nothing, page 1 is rendered and description, vendor, lot and qty are read by vision, then compared across documents. Findings carry a `VISION CROSS-CHECK:` prefix. They are **shown, but never by themselves draft Section 5 or set 6.1 = Yes**, which was an explicit Quality decision.

## Step 7 · Section 2 determination (fixed rules)

`determine_section2()` is plain Python. Every input is a fact from Steps 1–6. Every output is one of the allowed answers or `FLAG`, and `FLAG` leaves both boxes empty.

| Item | Question | Rule (simplified) |
|---|---|---|
| **2.1** | Receiving list and packing slip signed | `Completed` if the packing slip has the Ops signature and date **and** a receiving list is present, else `FLAG` |
| **2.2** | Mfg certs match drawing notes | `N/A` if drawing material **and** finish are N/A. `Yes` if a cert covers each material and each finish. `No` + Section 5 note if any is missing or differs |
| **2.3** | Mfg inspection report, dims in spec | `N/A` if the drawing has no numbered callouts. `Yes` if the report is present and all pass. `No` if any are out of spec. `FLAG` if the report is missing while callouts exist |
| **2.4** | Contract-manufacturer travelers reviewed | `Yes` for confidential biologic product from a listed contract manufacturer (Director review). For genuine traveler vendors: `Yes` if travelers are present and complete. Everyone else, including service vendors that assemble nothing: `N/A` + explicit flag, never a bare "No" |
| **2.5** | Sterilization COC reviewed | `Yes` if the ster COC is present and dose parameters passed. `N/A` if not sterile (explicit "unsterilized" wording is checked **before** vendor heuristics). `No`/flag if required and missing. Section 1.3 QDA = Yes also forces the check |
| **2.6** | Confidential docs / Director review | `Yes` **only** for biologic product from the listed contract manufacturers, so the Director of Operations signs before the QM. Otherwise `No` |
| **Disposition** | Raw vs. Finished Goods | Look up the part# in the Approved Product Destination List. Not listed → **Raw Goods**. Listed → `FLAG` (both unchecked), except a defined vendor profile known to ship finished, labeled goods → **Finished Goods** |

The design rule behind all of this: **a missing fact produces a flag, never a default answer.** Several rules were rewritten after real runs, for example retiring a `"label" in description` keyword heuristic for the catalog lookup. Each rewrite carries a test case.

## Step 8 · Fill PRF Sections 2, 5, 6

**Write boundary:** only these cells, every run.

<a id="prf-cell-map"></a>

```
Sheet "Sections 1-2"
  2.1 → I20 (Completed)
  2.2 → I21 Yes | L21 No | O21 N/A
  2.3 → I22 Yes | L22 No | O22 N/A
  2.4 → I23 Yes | L23 No | O23 N/A
  2.5 → I24 Yes | L24 No | O24 N/A
  2.6 → I25 Yes | L25 No           (no N/A)
  H29 QM initials, I32/O32 Director sig/date → re-asserted "N/A" when 2.4 / 2.6 ≠ Yes

Sheet "Sections 4-6"
  Section 5: L12 = N/A checkbox | B14 = discrepancy text (plain text, discrepancy_flags only)
  6.1:       K30 Yes | N30 No
  6.2 N/A:   B31
  Disposition: J36 Raw Goods | N36 Finished Goods (both unchecked on FLAG)

Never written: Section 1 (E3–E6, M3–M6, 1.2, 1.3), Section 3, Section 4
```

**Native Excel checkboxes.** This was the hardest bug in the project. openpyxl rewrites the sheet XML without references to the VML/ctrlProps parts. The parts stay in the zip **orphaned**, and the next time a person opens and saves the file in Excel, Excel quietly deletes them, along with the letterhead logo. The fix:

- `save_preserving_drawings()` carries the drawing, VML and ctrlProps parts forward
- `restore_form_controls()` re-links them in the sheet XML
- `set_native_checkboxes()` / `verify_native_checkboxes()` set and **read back** the real control state
- `verify_xlsx_package_integrity()` checks that no part is orphaned
- ☑/☐ glyphs are written only to a throwaway copy used for the PDF export, because LibreOffice can't render VML controls

**No-clobber guard.** `assert_no_user_edits()` refuses to regenerate a package a person has already edited. `--force` overrides it, on explicit instruction only.

**Export.** `export_and_stamp_pdf()` runs LibreOffice headless, then stamps the letterhead, title and document-number block on every page. If no logo can be found anywhere, `LogoMissingError` blocks that part.

**Section 5 language.** Plain text only. It covers real discrepancies only, never process notices. It contains **no NCR or disposition recommendation**, because that's Engineering's call under SOP 8.3. It follows the wording of signed precedents (the *Comment & Discrepancy Reasoning Reference*, built only from fully signed packages).

## Step 9 · Run manifest (development aid)

`write_manifest()` → `<tools>\logs\<PO#>_<part#>[_Lot<lot>]_<YYYYMMDD-HHMM>_manifest.json`. It records which file fed each fact, the Section 1/2/5/6 values, the package page order, exclusions and every flag. It is **not** a quality record and is kept out of Box on purpose. The signed package is the record. See [`examples/example-run-manifest.json`](../examples/example-run-manifest.json).
