---
tags: [validation, controlled-document]
---

# QF 7.4.3-06 — Validation Protocol (Receiving Automation Tool)

Lives in `Validation\QF_7.4.3-06_Validation_Protocol.docx` at the vault root. **Currently Rev 05 (2026-07-22).** Defines the approach, acceptance criteria, and sign-off requirements for validating [[receive.py|the Tool]] before rollout, per 21 CFR 820.70(i) and ISO 13485:2016 cl. 4.1.6.

## Structure

- **1.0 Purpose / 2.0 Scope** — what's in/out of scope (dock-level incoming inspection, the engineer's Section 1/3/4 pre-fill, and Box's own file-sync behavior are explicitly OUT of scope)
- **9.0 Open Items** — see [[Open Items (Master List)]] for the full cross-referenced list (this protocol's own internal numbering is 9.1–9.11, which does not map 1:1 onto CLAUDE.md's Open Items 1–12 — see that note for the mapping and gaps)
- **10.0 Acceptance Criteria by Function** — one row per function (Section 1 validation, file-access boundary, document classification, cross-reference checks, [[Step 7 - Section 2 Determination (Fixed Rules)|Section 2 logic]], [[Step 8 - Fill PRF Sections 2, 5, 6|write boundary]], [[Step 9 - Run Manifest (Desktop Dev Aid)|manifest]], checkbox rendering/resolution, [[Step 6.5 - Vendor Document Completeness Check|completeness check]], [[Step 6.6 - Vision-Based Key-Field Cross-Check|vision cross-check]], etc.)
- **11.0 Pass/Fail Criteria & Deviation Handling**
- **12.0 Test Strategy** — real PO folders under the actual Box directory, not synthetic fixtures, except where a scenario must be deliberately staged (e.g. TC-13 Section 1 mismatch, TC-15/16 boundary tests)
- **13.0 Revalidation Triggers** — 13.1 (full matrix rerun) vs. 13.2 (partial regression only)
- **14.0 Sign-off** — Author/Validation Lead (<qe-engineer>), Engineering Reviewer, Operations Reviewer (Director of Operations), Quality/Final Approver (Quality Manager)

## Revision history highlights

- Rev 03 → Rev 04: documented the checkbox-reading fixes, [[Step 6.5 - Vendor Document Completeness Check|Step 6.5]], and `find_prf()` suffix fallback; added Open Items 9.7–9.10.
- Rev 04 → Rev 05: added Open Item 9.11 and a new acceptance-criteria row for [[Step 6.6 - Vision-Based Key-Field Cross-Check]].

## Target rollout

Friday, July 24, 2026 — contingent on closure of Section 9.0 open items and completion of all Must-Pass-100% test cases in [[Attachment A - Test Case Matrix]].

## Related documents

- [[Attachment A - Test Case Matrix]] (QF 7.4.3-06 Attachment A)
- [[Attachment B - Instruction Guide]] (QF 7.4.3-06 Attachment B)
- [[WI-7.4.3-01 Overview]] (the instruction this protocol validates)
