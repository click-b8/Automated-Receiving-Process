---
tags: [workflow, qe-workflow]
---

# Staged for Inspection (Monday)

Part of [[QE Receiving Workflow]] · [[W03 - Received into Quarantine in ERP|← W03]] · [[W05 - Physical Inspection, Dimensions and Testing|W05 →]]

**Who:** QE / Engineering  
**Where:** Quarantine

## Physical work
- Parts wait in Quarantine for the inspection day.
- Start the PRF (QF 7.4.3-02) from the template and fill Section 1: part info, quantity inspected (N/A, 100% or AQL), QDA.
- Save it to `Inspection Forms\` as `PRF <part#> Rev X.xlsx` (add `Lot X` when the same part and rev appear twice).

## Records
- QF 7.4.3-02 PRF, Section 1

## Automation embedded in this step
- [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)]]
- [[Edge-Case Scripts]]

## Watch-outs
- Section 1 is engineer-owned. Automation only cross-checks it against the source documents later (Step 2.5). The `prefill_PO26-*.py` scripts in the tools folder are edge-case helpers, outside the governed path.
