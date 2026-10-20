---
tags: [governance, step]
---

# Step 1 — Verify Folders and Identify All Parts

Part of [[WI-7.4.3-01 Overview]].

## What happens

1. Confirm the PO#'s encoded year matches the year folder Claude Code was opened in (see [[File Access Rules]]) — stop and flag if not.
2. Confirm `PO##-####\Inspection Forms\` and `PO##-####\COCs\` exist — stop if missing.
3. Open the PO document (check PO folder root first, then `Shipping Docs\`). Extract: PO number, PO date, PO revision, vendor name, and every line item (part#, description, revision, qty ordered). Quote pages are ignored — only pages with a "Purchase Order" header count.
4. Scan `Inspection Forms\` for files matching `PRF*{part_number}*.xlsx`. Parts with a PRF present are "ready"; parts without one are "pending" and reported at the end, not blocked forever.

## Implemented by

- `extract_po_data()` and `find_prf()` in [[receive.py]]
- `find_prf()` has a PRF-as-suffix fallback (e.g. `SR90-1004 PRF.xlsx`) added after a real part was silently treated as pending — see [[Open Items (Master List)#Item 7]] and [[Real-World Runs (Case Studies)#PO26-1003 — CatalogSupplier]]

## Feeds into

- [[Step 2 - Match & Classify Documents]] (needs the parts list and PO data)
- [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)]] (compares against these same extracted facts)
