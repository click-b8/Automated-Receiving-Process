---
tags: [governance, step]
---

# Step 4 — Extract Data from Documents

Part of [[WI-7.4.3-01 Overview]]. `pdfplumber` first (zero tokens), escalate to `claude-haiku-4-5-20251001` only when text extraction is incomplete.

## Extraction functions ([[receive.py]])

| Function | Document | Key fields |
|---|---|---|
| `extract_po_data()` | PO | po_number, po_date, po_revision, vendor, line_items[] |
| `extract_packing_data()` | Packing slip | vendor, part_number, description, qty_shipped, lot_number, ops_signature_present |
| `extract_receiving_data()` | Receiving list | part_number, description, qty_received, lot_number, exp/mfg/ster dates |
| `extract_coc_data()` | Mfg COC / Ster COC / Mfg report | vendor_part_number, lot_number, qty, dates, material/finish, out_of_spec_dimensions |

## Known weak point

When a document is a scanned image with little or no text layer (e.g. CatalogSupplier's PO/packing list on PO26-1003 — 6 and 0 extractable characters respectively), these functions return empty dicts, and [[Step 6 - Cross-Reference Source Documents]] has nothing to compare. This is exactly the gap [[Step 6.6 - Vision-Based Key-Field Cross-Check]] was built to partially close. See [[Real-World Runs (Case Studies)#PO26-1003 — CatalogSupplier]].

## Feeds into

- [[Step 5 - Parse Drawing with Cache]] (drawing extraction is a separate, cached path)
- [[Step 6 - Cross-Reference Source Documents]]
- [[Step 7 - Section 2 Determination (Fixed Rules)]] (every input fact Step 7's functions consume comes from here or Step 5)
