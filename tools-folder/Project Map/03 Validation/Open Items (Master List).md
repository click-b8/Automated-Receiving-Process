---
tags: [validation, open-items]
---

# Open Items — Master List

Every gap that blocks the WI moving from *Development* to *Rev A*. In the real vault each item has the evidence, the decision owner and the date it was confirmed; here only the titles are kept so you can see the kind of issues a tool like this accumulates during validation.

- Item 1 — ~~Section 2.4 Edge Case (confidential Biologic)~~ — FIXED 2026-07-27
- Item 2 — `_Automation_Master` Canonical Location — OPEN, bigger than before
- Item 3 — Yearly Refresh Procedure — OPEN, same fate as Item 2
- Item 4 — Section 1.3 QDA → 2.4/2.5 Cross-Check Rule — OPEN
- Item 5 — ~~Receiving-List Highlight Bleed~~ — FIXED 2026-07-20, recurred and re-fixed twice more
- Item 6 — Vendor Document Completeness Check — OPEN
- Item 7 — Checkbox Cell-Resolution Fallbacks & PRF Filename Convention — OPEN
- Item 8 — Traceability Manifest — ~~STILL NOT IMPLEMENTED, BLOCKS REV A~~ IMPLEMENTED 2026-09-17, LOCATION DECIDED 2026-10-05, two sub-gaps open
- Item 9 — ~~Section 6 Checkbox Cell-Mapping Bug~~ — FIXED 2026-07-22, then CONFIRMED 2026-07-27
- Item 9b — ~~Section 6 label-deletion bug~~ — FIXED 2026-07-27
- Item 10 — 2.2 Material/Finish Annotation — STILL DISABLED, cell unknown
- Item 11 — ~~Stray/Unverified Checkbox at G2~~ — RESOLVED 2026-07-27
- Item 12 — Vision-Based Key-Field Cross-Check — OPEN
- Item 13 — ~~Disposition Rule~~ — CODE FIX IMPLEMENTED 2026-09-11, CORRECTED 2026-09-14, ratification still needed
- Item 14 — Section 2.1 False Negative on a Signed Packing Slip — OPEN, not investigated
- Item 15 — Section 5 Narrative Contains Literal Markdown — OPEN, cosmetic but on a signed record
- Item 16 — J8 (Section 1.2 AQL Checkbox) Fix — NEEDS VERIFICATION
- Item 17 — Calibration-Date Soft Check — cell reference now fixed, second concern still open
- Item 18 — Section 2.4 Default (non-Biologic, non-traveler-vendor, VendorD) — code changed 2026-09-11, not yet SOP-ratified
- Item 19 — PO Revision-Conflict Halt — implemented 2026-09-11, deliberately stricter than the proven bespoke-script behavior
- Item 20 — Partial-PO Hold — implemented 2026-09-11, corrected same day
- Item 21 — Quantity-Shortage 10% Tolerance — Quality-confirmed, not yet an SOP citation
- Item 22 — Section 4 Text-Clipping Guard — function written, NOT wired to a real cell
- Item 23 — Stadium/Oval "Note 4" Drawing-Callout Support — implemented via a DIFFERENT technique than the bespoke script, unverified
- Item 24 — `highlight_receiving_line()` Sterilization Round-Trip Disambiguation — corrected same day
- Item 31 — Section 5 drafting doesn't use the Reasoning Reference in `receive.py` yet — NEW 2026-10-05, OPEN
- Item EC — Undocumented Edge-Case Scripts (proposed by this vault, still not in CLAUDE.md)

Related: [[QF 7.4.3-06 Validation Protocol]], [[Attachment A - Test Case Matrix]], [[_Hub - Validation]].
