---
tags: [governance, step]
---

# Step 2 — Match and Classify Documents for Each Part

Part of [[WI-7.4.3-01 Overview]].

## What gets matched, and where

| Document | Primary location | Fallback |
|---|---|---|
| Drawing | `Drawings\` | `Drawings & Stepfiles\` |
| COCs | `COCs\` (every PDF classified) | — |
| Receiving list | PO root | `Shipping Docs\`, then `Receiving List\` |
| Packing slip | `Shipping Docs\` | PO root |
| Pre-filled PRF | `Inspection Forms\` | — (read-only, never modified) |

## COC classification rules (rule-based first, zero tokens)

- "sterilizerco" / "irradiation" / "gamma" / "e-beam" → `ster_coc`
- "traveler" / "job traveler" / "router" → `traveler`
- "certificate of conformance/compliance" → `mfg_coc`
- ACME Medical name + "packing"/"purchase order" → `sgt_doc_exclude`

**Vendor override:** VendorF and VendorG issue manufacturer inspection reports labeled "traveler" in the filename that are NOT real travelers — classified as `mfg_inspection_report` regardless of the filename keyword. See `TRAVELER_VENDORS` in [[receive.py]].

Ambiguous files escalate to `claude-haiku-4-5-20251001`. Combined PDFs (>2 pages) are split page-by-page via `split_pdf_by_page_type()`.

**Excluded always:** ACME Medical-generated docs, commercial/proforma invoices (a bundled-invoice packing slip variant fixed 2026-07-2x via `strip_invoice_pages()`), confidential Biologic travelers (VendorB/VendorA/VendorC Biologic only), and — since 2026-09-08 — an already-compiled-elsewhere COC accidentally bundled inside a scanned packing slip's own text-empty pages (`_vision_classify_page()`).

## Per-part document scoping on multi-part POs — NEW 2026-09-11

`_filter_classified_for_part()` (new in [[receive.py]]) closes a
previously-unknown traceability bug: COC/traveler/mfg-report classification
had never been scoped to a specific part on a multi-part PO at all — every
part shared the same PO-wide `classified` dict, filename-substring matched
against the part's own number using the same convention already used for
drawing and PRF matching.

## Partial-PO hold — NEW 2026-09-11, corrected same day ([[Open Items (Master List)#Item 20]])

A ready part on a MULTI-part PO now returns a new `"held"` status
(distinct from `"error"`) when it has zero COC-type documents attributable
to it AND no packing-slip/shipping-ticket-like file names it specifically —
while sibling parts on the same PO DO have documents. Codifies the
PO26-1008 finding (CAT-XX-65/Lot A084: no shipping ticket, no COC, while two
siblings had everything). **Original version also required the receiving
list to have no line for the part (a triple absence) — found wrong the
same day**: the real precedent's receiving list DID have a line for this
lot already (595/596 logged received), so that third condition would have
made the tool miss the very case it exists for. Now a double absence only.
Deliberately keyed off "siblings on this PO DO have documents," so it can't
misfire on the far more common single-part-PO case where documents
routinely aren't part-number-named at all — which also means it can
under-fire the same way if a PO's own naming convention never names any
document with a part number at all.

## Receiving-list highlighting

`highlight_receiving_line()` (in [[prf_helpers.py]]) must be called **after** receiving-list data extraction, using the **lot number** as the search term (falls back to part number only if lot extraction failed) — order matters, this was a real bug (see [[Changelog Timeline]]). It highlights the full multi-line row using both a vertical-gap test and a content-based row-boundary test (`_is_lot_subrow()`), and — since 2026-09-08 — clusters lines into visual ROWS by vertical-centre proximity with bold-line as a hard boundary (rebuilt after a THIRD real-world bleed recurrence, see [[Open Items (Master List)#Item 5]]). Since 2026-09-11, it also takes an optional `part_number` to re-anchor the search when a shared lot number across a sterilization round-trip's two rows leaves no row qualifying as "received" at all — see [[Open Items (Master List)#Item 24]].

## Feeds into

- [[Step 3 - Compile the Receiving Package]] (uses the classified documents directly)
- [[Step 4 - Extract Data from Documents]] (extracts structured fields from what's classified here)
- [[Step 6.5 - Vendor Document Completeness Check]] (runs on every document that will be compiled)
