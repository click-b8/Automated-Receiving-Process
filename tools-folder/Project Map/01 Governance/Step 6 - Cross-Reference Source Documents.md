---
tags: [governance, step]
---

# Step 6 — Cross-Reference Source Documents Against Each Other

Part of [[WI-7.4.3-01 Overview]]. Distinct from [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)]], which checks the PRF against the sources — this step checks the sources (PO, packing slip, receiving list, COC, drawing) against **one another**.

## What `cross_reference()` checks ([[receive.py]])

- Part# consistency (PO / Packing / Receiving / Drawing)
- Revision consistency (Ordered vs Received — flags both if different, both drawings get compiled per [[Step 3 - Compile the Receiving Package]])
- Lot# consistency (Packing / Receiving / COC)
- Qty: PO vs Received, Shipped vs Received — **since 2026-09-11, a SHORTAGE (received < expected) of 10% or less is a soft, informational flag (`QTY SHORTAGE (WITHIN 10% TOLERANCE)`) rather than a hard `QTY MISMATCH`** — see below. An overage, or a shortage beyond 10%, is unchanged.
- Mfg date ≠ Ster date (must differ — same-day is flagged CRITICAL)
- COC vendor part# vs. drawing supplier part#
- PRF Section 1 vs. every other source (added 2026-07-27 — closes this WI's original never-implemented "Step 2.5" gap; soft flag, does not block, not yet verified against a deliberately-mismatched fixture)

## 10% quantity-shortage tolerance — NEW 2026-09-11 ([[Open Items (Master List)#Item 21]])

Codifies a real, repeated finding on PO26-1008 (997/1,000, 795/800, 595/596
— all short by under 1%, all pre-acknowledged by the customer, engineer's
ruling was Section 5 = N/A / 6.1 = No every time). Quality-confirmed
threshold, **not yet tied to a cited SOP provision**. Does NOT itself
record the customer/engineering acknowledgment anywhere on the PRF — if
that belongs on the signed record, it's still the engineer's own Section 4
entry; automation never writes there.

## PO revision-conflict halt — NEW 2026-09-11, a Step-6-level rule ([[Open Items (Master List)#Item 19]])

New `find_po_documents_all()` + a gate in `main()` (runs before ANY part on
the PO is processed, since this is a property of the PO itself, not one
line item): when more than one PO document is found in scope and they
disagree on revision or a line item's ordered quantity, the **entire run
for this PO halts** — not just one part — naming exactly which copies
disagree and on what. Codifies the PO26-1005 finding (Rev B/385 ea vs. Rev
C/400 ea, file timestamps inconclusive). **Deliberately stricter than the
real, proven bespoke-script behavior**: `process_PO26-1005.py` actually
auto-resolves using corroborating documents (receiving list + PRF Section
1.1) and just flags the conflict — QE was asked to choose between
matching that behavior or keeping an unconditional halt, and chose the
halt. Never yet exercised against a real conflicting pair through this
governed path itself.

## Known historical gap — closed by Step 6.6

`cross_reference()` has **never** compared part description or vendor name across documents — only Part#/Rev/Lot#/Qty/dates/vendor-part#. On a fully scanned document pair with nothing extracted, this silently finds "no issue" even when the documents actually disagree. This is exactly what let "DISC-1" through on PO26-1003 (PO said "DOWEL PIN", packing slip said "CAPTIVE PIN") — see [[Real-World Runs (Case Studies)#PO26-1003 — CatalogSupplier]] and [[Step 6.6 - Vision-Based Key-Field Cross-Check]], which was added specifically to close this gap for scanned documents.

## Feeds into

- Flags list → filtered into `discrepancy_flags` → [[Step 8 - Fill PRF Sections 2, 5, 6]] (Section 5 narrative / Section 6.1)
