# Outlook integration

Email is still how receiving work moves **between people**: vendors, signers and the production board. The validated tool deliberately **does not send email**. Outlook sits around it in four places.

```
              vendor COC / packing-slip requests            STUCK notices
 Vendors  ◄──────────────────────────────────────  QE  ──────────────────────►  monday item updates
                                                    │                                   │
                                                    │ compiled package (PDF)            ▼
                                                    ▼                           Delays tab (dashboard)
                       Director of Ops → Engineering Reviewer → Quality Manager
                                                    │
                                                    ▼ signed PDF filed to Completed Receivings (Box)
```

## 1 · Vendor document requests

Missing documents are the most common delay. When `receive.py` flags a missing packing slip, receiving list, COC, sterilization certificate or first-article/inspection report, the **Delay Playbook** says to treat the flag as *"request now"*, not *"note in Section 5"*:

> Ask the vendor for the COC, inspection report / FAI and signed packing slip when the PO ships, and don't compile until Shipping Docs and COCs are complete.

The vendor's reply attachments are saved into the PO's `COCs\` or `Shipping Docs\` in Box, and the tool is re-run. A COC dated **after** receipt is flagged by the tool. The reply email is filed with the package as evidence.

## 2 · Signature routing

The compiled package is emailed from Outlook to the signers in the order the PRF requires:

| Condition (from Step 7 / Step 8) | Who signs first |
|---|---|
| 2.6 = Yes (confidential biologic batch records) | **Director of Operations**, before the QM |
| 2.4 = Yes | Quality Manager initials (H29) |
| Section 6 has discrepancies | **Engineering Reviewer**, before the QM |
| always | **Quality Manager**, final review and acceptance |

Packages over the mail size limit go as a **Box link**, or by DocuSign with the PRF attached separately. That rule is the *Signature logistics* fix in the Delay Playbook. While waiting, the monday item sits in *Waiting for Signature* with **Waiting on** set to the signer, which feeds the dashboard's *waiting by person* chart.

## 3 · STUCK notices → monday

When a PO stalls, a short email-style notice goes to the PO's monday item and is stored there as an update:

```
STUCK - PO26-0085
Issue: Qty short 4 pcs vs PO, waiting on vendor confirmation - Auto
```

The dashboard parses these (`parseStuckIssues()`) to infer a delay reason for holds where nobody filled the *Delay reason* column. Keep the `Issue:` line short and specific, because it's what the classifier reads.

## 4 · Claude's Microsoft 365 connector

In Claude (claude.ai / Cowork) sessions with the Microsoft 365 connector enabled, Claude can search and read Outlook mail and calendar on the user's behalf. In this project it's used for:

- **finding vendor replies**: "find the COC the vendor sent for lot 2604002 last week" (`outlook_email_search`)
- **chasing signatures**: checking whether a package went out and to whom
- **triage**: pulling recent threads about a held PO before deciding the next step

These are **assistive, read-only lookups during development and triage**. Nothing in the validated path depends on them, and Claude doesn't send email on its own.

## Why not automate the email?

Routing changes who is accountable for a signed quality record. Sending it automatically would bring email into the validated scope (recipient logic, attachment integrity, delivery evidence). The tool does the part that's safe to automate: it **tells you who has to sign and in what order**. People send it.
