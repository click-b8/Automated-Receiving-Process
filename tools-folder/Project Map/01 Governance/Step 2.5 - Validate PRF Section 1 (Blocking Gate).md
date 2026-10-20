---
tags: [governance, step, blocking-gate]
---

# Step 2.5 — Validate the Engineer's Pre-Filled Section 1 (read-only, blocking)

Part of [[WI-7.4.3-01 Overview]]. Runs for every ready part, **before** compilation and before any [[Step 7 - Section 2 Determination (Fixed Rules)|Section 2]] logic.

## What it checks

Reads Section 1 fields (never modifies them): `Product Name (M3)`, `Product P/N (M4)`, `Revision (M5)`, `PO# (E3)`, `Supplier (E4)`, `Quantity Received (E6)`, `Lot/Batch Number (M6)`. Compares each against the same fact independently extracted in [[Step 1 - Verify Folders and Identify Parts]] / [[Step 4 - Extract Data from Documents]] from the PO, receiving list, packing slip, COC, and drawing.

## On any mismatch

- **HALT** — do not compile, do not fill or export a PRF, for this part, this run only. Other ready parts are unaffected.
- Record the mismatch in the summary **above every other flag category**, in an exact format naming the PRF value vs. the source-document value.
- **Never guess which value is correct. Never write a corrected value into Section 1** — that field stays engineer-owned; this step is detection only.

## Why this exists

This is the single most FDA-relevant gate in the whole tool: it stops the automation from ever compiling a package or writing quality-record content off of a PRF whose Section 1 disagrees with the actual receiving evidence. It is a hard, deterministic halt — no vision judgment call, no soft flag.

## Related

- [[QF 7.4.3-06 Validation Protocol]] Section 9.3 — "resolved by design, verify" (Test Case TC-13 is Must-Pass-100%; this item is VP-only, not mirrored in [[Open Items (Master List)|CLAUDE.md's own Open Items list]])
- [[Attachment A - Test Case Matrix]] TC-13
