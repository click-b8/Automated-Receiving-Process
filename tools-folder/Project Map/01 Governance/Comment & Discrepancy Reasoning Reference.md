---
tags: [governance, section-5, language, new-2026-10-05]
---

# Comment & Discrepancy Reasoning Reference

Part of [[WI-7.4.3-01 Overview]]. **New 2026-10-05.** The authoritative text,
including the precedent table and the full harvest procedure, is the section
of the same name in [[CLAUDE.md]]. This page summarizes it and links it into
the vault.

## Why it exists

Comments and discrepancy notes were being rewritten after compile or after
Engineering Reviewer/Quality Manager feedback. Instead of a manifest-driven learning loop or new
folders in Box, QE and Engineering decided 2026-10-05 that the tool learns from
**what actually got signed**.

## Sources — signed material only

1. [[Attachment A - Test Case Matrix]] — edge-case POs with a Pass result.
2. Each PO's `Completed Receiving\` subfolder — the final Adobe-signed
   package. Signed Section 4/5/6 wording is the reference; a difference from
   the draft (manifest in [[Step 9 - Run Manifest (Desktop Dev Aid)|Desktop logs]], or the In-Progress
   version) is the reviewer's correction.

Never a source: unsigned In-Progress packages, draft PRFs, chat transcripts,
or verbal comments that aren't on a signed record.

## Scope

- Applies to **Section 5** (B14) and run-summary flag wording.
- **Section 4 unchanged:** engineer-owned, read-only. Signed Section 4 text
  is used only as vocabulary — never written, pre-filled or suggested.
- Never changes a [[Step 7 - Section 2 Determination (Fixed Rules)|Step 7]]
  rule, and never recommends an NCR or disposition (SOP 8.3).

## Section 5 drafting rules

1. Plain text, no Markdown ([[Open Items (Master List)|Item 15]]).
2. Only `discrepancy_flags` — never process notices (see
   [[Step 8 - Fill PRF Sections 2, 5, 6]]).
3. No NCR or disposition language.
4. Follow a matching precedent's signed wording.
5. *Provisional until the first harvest:* facts with sources, one
   discrepancy per sentence, names as printed on the documents.

## Precedents seeded 2026-10-05 (none yet checked against a signed record)

| Ref | PO | Situation → decision |
|---|---|---|
| P-01 | PO26-1008 (VendorD) | Qty short <1%, pre-acknowledged → Section 5 N/A, 6.1 No |
| P-02 | PO26-1005/0431 (SterilizerCo) | Sterilization-only vendor → 2.4 N/A |
| P-03 | PO26-1003 (CatalogSupplier) | DOWEL PIN vs CAPTIVE PIN → real discrepancy, SOP 8.3 |
| P-04 | PO26-1003 | Highlight failure → tooling notice, not a discrepancy |
| P-05 | PO26-1007 (VendorE) | Blank label stock → Raw Goods |
| P-06 | PO26-1008 (VendorD) | Anodized + etched screws → Finished Goods, 2.4 N/A — **conflicts with its own signed PRF** |
| P-07 | PO26-1004 (VendorB) | UnSterilized Biologic QDA → 2.5 N/A, 2.4/2.6 Yes, Director of Operations first |

## Harvest procedure (summary)

Batches only; QE's explicit permission naming the POs
([[File Access Rules]]); `Completed Receiving\` read-only; text extraction
with no model calls; every reviewer edit sorted as **wording** (precedent),
**added information** (extraction gap → Open Item), or **correction**
(defect → Open Items and test cases, never a precedent). Changes follow the
WI revision process.

## Validation and open work

- TC-27 (Section 5 language) and TC-28 (harvest boundary) in
  [[Attachment A - Test Case Matrix]].
- [[Open Items (Master List)#Item 31|Item 31]]: `receive.py`'s prompt and
  manifest don't use this yet.
