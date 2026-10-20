---
tags: [governance, step, vision-judgment-call, soft-flag]
---

# Step 6.6 — Vision-Based Key-Field Cross-Check

Part of [[WI-7.4.3-01 Overview]]. Added 2026-07-22, most recent addition to the WI. Separate again from [[Step 6 - Cross-Reference Source Documents]] and [[Step 6.5 - Vendor Document Completeness Check]].

## Root cause it closes — "DISC-1"

`cross_reference()` ([[Step 6 - Cross-Reference Source Documents]]) can only compare fields [[Step 4 - Extract Data from Documents|Step 4's text extraction]] actually produced. On PO26-1003/SR90-1004, the PO's page 1 had 6 extractable characters and the packing list had 0 — both essentially blank text layers — so extraction returned nothing for description/vendor/lot/qty, and cross_reference() silently found no issue even though the PO said "DOWEL PIN" and the packing list said "CAPTIVE PIN." A human caught it only on manual review. See [[Real-World Runs (Case Studies)#PO26-1003 — CatalogSupplier]].

## What `check_key_field_consistency()` does ([[receive.py]])

For the PO, packing slip, receiving list, and manufacturer COC: if a document's already-text-extracted description, vendor, lot#, and qty are **all** empty, renders page 1 and asks a vision model (`_vision_extract_key_fields()`) to read those four fields directly off the image. Documents that already produced something from text extraction are **not** sent to vision (cost control). Then compares description/vendor/lot#/qty across whichever documents have a value — qty normalizes to digits-only first so "525" and "525 EA" don't false-flag.

## Soft flag — by explicit Quality direction

Findings get a `VISION CROSS-CHECK:` prefix — always shown in the run summary but **excluded from `discrepancy_flags`**, so they can never by themselves draft a [[Step 8 - Fill PRF Sections 2, 5, 6|Section 5 narrative]] or set Section 6.1 to "Yes." Unlike the other two flag prefixes given this same exclusion (`RECEIVING LIST:`, `CALLOUT PLACEMENT:` — pure tooling limitations), a `VISION CROSS-CHECK:` finding **can** describe something genuinely wrong with the part. This was a deliberate, explicit choice — not a default — and is flagged as an open question in [[Open Items (Master List)#Item 12]] (DISC-1 itself was arguably NCR-worthy per SOP 8.3).

## Status

Not yet confirmed against a live vision call or a real scanned document end-to-end. See [[Open Items (Master List)#Item 12]]. Documented in [[QF 7.4.3-06 Validation Protocol]] Rev 05.
