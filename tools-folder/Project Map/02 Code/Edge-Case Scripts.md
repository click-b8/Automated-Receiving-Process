---
tags: [code, edge-case]
---

# Edge-Case Scripts

Per-PO scripts (`process_PO26-xxxx.py`, `prefill_PO26-xxxx.py`) used when a real PO doesn't fit the general path of [[receive.py]] — e.g. a PRF that is already signed and must not be modified, an unusual multi-shipment split, or a vendor document layout the tool can't parse yet.

## Rules for using them

- They live in the tools folder, never in a PO folder, and reuse [[prf_helpers.py]] instead of copying logic.
- They are **outside the validated path**. Quality decided to *fix root causes* in `receive.py` rather than add a review procedure for these scripts: anything that recurs gets folded back into the general tool, and the script stays only for the genuinely one-off part.
- The dashboard's *POs & parts* tab shows which POs ran through an edge-case script (`@edge`).

The real scripts contain vendor- and part-specific data and are not published in this repo.

Related: [[_Hub - Code]], [[Open Items (Master List)]].
