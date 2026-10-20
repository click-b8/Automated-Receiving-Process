---
tags: [governance, step, vision-judgment-call, non-blocking]
---

# Step 6.5 — Vendor Document Completeness Check

Part of [[WI-7.4.3-01 Overview]]. Added 2026-07-20. Separate from [[Step 6 - Cross-Reference Source Documents]] (sources vs. each other) — this checks a **single vendor document's own internal completeness**.

## Root cause it closes

A VendorA "QC Final Inspection" page (PO26-1002) had its "2nd Pack QC Inspection Post-Sterilization" section left completely blank — no data, no Pass/Fail, no N/A box checked — and nothing flagged it, because automation only ever extracted structured fields from `mfg_coc`/`ster_coc` via fixed JSON schemas; `mfg_report`/`other` documents got zero content review. See [[Real-World Runs (Case Studies)#PO26-1002 — VendorA]].

## What `check_document_completeness()` does ([[receive.py]])

Renders **every page** of every document about to be compiled (mfg_coc, ster_coc, traveler, mfg_report, other) and asks `claude-haiku-4-5-20251001` to find any section requiring an entry (data field / checkbox / Pass-Fail-N/A selector) left completely blank — no data AND no N/A marked. A section with data, or with its own N/A checked, is fine.

## Blocking behavior

**Does NOT block the run** on a failed review (API error, unparseable response) — degrades to "no issues found." This is a supplementary due-diligence check, not a [[Step 7 - Section 2 Determination (Fixed Rules)|Section 2 input]].

## Ordering bug (fixed)

This check originally ran **after** [[Step 8 - Fill PRF Sections 2, 5, 6|fill_prf()]] had already saved the PRF — meaning its flags could never reach Section 5/6.1, only the console summary. Fixed 2026-07-22 by reordering `process_part()`. See [[Changelog Timeline]].

## Status

Not yet confirmed as authoritative — see [[Open Items (Master List)#Item 6]]. A vision-model judgment call, not a deterministic Step 7 rule.
