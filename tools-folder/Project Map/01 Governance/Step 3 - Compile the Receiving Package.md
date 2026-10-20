---
tags: [governance, step]
---

# Step 3 — Compile the Receiving Package

Part of [[WI-7.4.3-01 Overview]]. Output: `PO##-####\In-Progress Receiving\PO##-##### Receiving Part X - Product Part#.pdf`

## Document order (per SOP 7.4.3 Rev L / the Quality Manager's Proof Receiving Checklist)

1. PO (quote pages stripped)
2. Packing Slip (vendor-only, never ACME Medical-generated; commercial/proforma invoice pages bundled into a scanned packing slip are now stripped too — `strip_invoice_pages()`, 2026-07-2x)
3. Receiving List (highlighted copy, QF 7.4.3-01)
4. Drawing(s) — rev mismatch: ordered rev first, received rev second; bubble callouts only on the received-revision drawing. **Note (2026-09-11, Attachment A):** one real run compiled the REVERSE order (received rev first) on explicit one-off engineer instruction — flagged as a possible WI-amendment question for Quality, not treated as a Fail. See [[Attachment A - Test Case Matrix]].
5. PRF (QF 7.4.3-02 — after [[Step 8 - Fill PRF Sections 2, 5, 6|Sections 2/5/6]] are filled and exported)
6. Manufacturer COC(s)
7. Sterilization COC (if present)
8. Traveler / Job Traveler (if present, non-confidential)
9. Manufacturer inspection report / other QC documents

## Excluded

ACME Medical-generated COCs/packing slips/POs, quote pages, confidential Biologic travelers, commercial/proforma invoices, duplicate pages (including, since 2026-09-08, a COC bundled inside a scanned packing slip's own pages that would otherwise be compiled twice).

## Drawing callout bubbling — three conventions now, not two

`_find_callout_coords()` (in [[receive.py]]) recognizes a critical-dimension
callout three ways: a pre-existing circled Unicode digit, a parenthesized
number, and — new 2026-09-11 — a hand-drawn stadium/oval shape around a
plain digit (`_find_stadium_oval_coords()`, via `page.get_drawings()`
vector-path geometry). The third convention is drawn in **red**, distinct
from the black used for the other two, so an inspection bubble can never be
confused with a drawing's own black BOM-item or revision-change balloon.
**Not yet verified against a real drawing** — and, per
[[Open Items (Master List)#Item 23]], this is an independently-designed
technique, not a port of the real bespoke script's own (very different,
per-part-hardcoded) method. See [[Edge-Case Scripts]].

## Implemented by

`compile_pdf()` in [[prf_helpers.py]] — PyMuPDF merge + Ghostscript re-distill (strips security flags/encryption/AcroForm locks). See [[Tools and Dependencies]] for the Ghostscript requirement.

## Known deviation

The `mfg_report` (manufacturer inspection report) document was extracted from and used for [[Step 7 - Section 2 Determination (Fixed Rules)|2.3]] since early on, but was silently **not** compiled into the package for a long time — position 9 was documented but never wired in. Fixed 2026-07-20; see [[Changelog Timeline]].
