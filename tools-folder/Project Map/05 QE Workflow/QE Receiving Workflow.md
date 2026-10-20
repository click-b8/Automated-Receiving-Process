---
tags: [workflow, moc]
---

# QE Receiving Workflow — from pallet to Completed Receiving

The receiving process as the quality engineer lives it. Each physical step lists the automation that runs inside it. Grounded in SOP 7.4.3 Rev L Table 1 (Operations → Engineering → Quality → Operations); items marked *confirm* are assumptions to check.

Visual board: [[QE Receiving Workflow.canvas]]. For a graph of just this workflow, open this note's **local graph** (depth 1), or filter the global graph with `path:"Project Map/05 QE Workflow"`.

## Sequence

1. [[W01 - Parts Arrive at the Receiving Dock]]
2. [[W02 - Paperwork Gathered and Filed to the PO Folder]]
3. [[W03 - Received into Quarantine in ERP]]
4. [[W04 - Staged for Inspection (Monday)]]
5. [[W05 - Physical Inspection, Dimensions and Testing]]
6. [[W06 - Compile with the Automation]]
7. [[W07 - Engineer Review and Sign]]
8. [[W08 - Routed for Signatures by Email]]
9. [[W09 - Filed to Completed Receiving]]
10. [[W10 - Released to Inventory or Returned]]

## Who does what (SOP 7.4.3 Table 1)

| SOP step | Who | Workflow steps |
|---|---|---|
| 1 Receive into Quarantine (QF 7.4.3-01) | Operations | W01, W03 |
| 2 Store documentation in PO folder | Operations | W02 |
| 3 Complete QF 7.4.3-02 | Engineering | W04–W07 |
| 4 Review and accept QF 7.4.3-02 | Quality (Director of Operations / Engineering Reviewer / Quality Manager as required) | W08 |
| 5 Move to inventory or return to vendor | Operations | W10 |

W09 (filing to Completed Receiving) is the record step that sits between SOP steps 4 and 5.

## Related
- [[WI-7.4.3-01 Overview]] — the controlled instruction behind the automation
- [[00 Index]]
