---
tags: [workflow, qe-workflow]
---

# Compile with the Automation

Part of [[QE Receiving Workflow]] · [[W05 - Physical Inspection, Dimensions and Testing|← W05]] · [[W07 - Engineer Review and Sign|W07 →]]

**Who:** QE runs it (manually or in a Claude Code session)  
**Where:** Desktop tools folder → PO folder

## Physical work
- Run `python receive.py PO26-xxxx`.
- Read the summary: Section 1 mismatches first, then each part's flags.

## Records
- Compiled package in `In-Progress Receiving\`
- PRF xlsx/pdf with Sections 2, 5, 6 filled
- Manifest in Desktop `logs\` (development aid only)

## Automation embedded in this step
- [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)]]
- [[Step 3 - Compile the Receiving Package]]
- [[Step 6 - Cross-Reference Source Documents]]
- [[Step 6.5 - Vendor Document Completeness Check]]
- [[Step 6.6 - Vision-Based Key-Field Cross-Check]]
- [[Step 7 - Section 2 Determination (Fixed Rules)]]
- [[Step 8 - Fill PRF Sections 2, 5, 6]]
- [[Step 9 - Run Manifest (Desktop Dev Aid)]]
- [[Comment & Discrepancy Reasoning Reference]]
- [[Manual vs Claude-Assisted Workflow]]

## Watch-outs
- Automation drafts. It never approves. Flags are the tool declining to guess; resolve them by hand from the source documents.
