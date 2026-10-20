# 03 · Using `receive.py`

## Before you run it

The tool processes **only parts whose PRF already exists**. A part with no PRF is listed as *Pending* and skipped. Per part:

1. **Documents are in the PO folder** (Box):

   ```
   26-1085 - VendorA\
   ├── 26-1085 - VendorA.pdf          PO document (root, or in Shipping Docs\)
   ├── Shipping Docs\                 packing slip(s), receiving list
   ├── COCs\                          manufacturer COCs, ster COCs, travelers, inspection reports
   ├── Drawings\                      one PDF per part, named by part number   (fallback: "Drawings & Stepfiles\")
   └── Inspection Forms\              PRF <part#> Rev <X>.xlsx  ← the engineer's pre-filled PRF
   ```

2. **The PRF Section 1 is filled** (PO#, vendor, inspection date, qty, product, part#, rev, lot, plus 1.2 quantity inspected and 1.3 QDA), along with Section 3 measurements and Section 4 comments if applicable. These sections are yours, and the tool only **reads** them.
3. **The PRF is named** `PRF SR90-1202 Rev A.xlsx`. If the same part and rev arrive twice on one PO, add `Lot <lot>`. Legacy `PRF-SR90-1202_Rev_A.xlsx` and `SR90-1202 Rev A PRF.xlsx` are also recognized.

## Run

```powershell
cd "C:\Users\<you>\Box\Shared\Purchase Orders\2026\26-1085 - VendorA"
python "C:\Users\<you>\Desktop\Receiving Task Automation Tools\receive.py" PO26-1085

# just some parts
python ...\receive.py PO26-1085 --parts SR90-1202,SR90-1203

# re-run over a package a human has already edited (discards those edits — only on explicit instruction)
python ...\receive.py PO26-1085 --force
```

The working directory **is** the PO folder. The tool never reads sibling POs or parent folders.

## What you get

```
26-1085 - VendorA\
├── In-Progress Receiving\
│   └── PO26-1085 Receiving Part 1 - SR90-1202.pdf     ← compiled package, ready to review & sign
└── Inspection Forms\
    ├── PRF SR90-1202 Rev A.xlsx                        ← Sections 2/5/6 filled, native checkboxes intact
    └── PRF SR90-1202 Rev A.pdf                         ← exported, letterhead stamped

<tools folder>\logs\PO26-1085_SR90-1202_20261020-1012_manifest.json   ← dev aid, not a quality record
```

And a console summary:

```
====================================================
 PO26-1085 — Complete | Vendor: VendorA
====================================================

 Part SR90-1202  Rev A  Lot 9904002
  In-Progress Receiving/PO26-1085 Receiving Part 1 - SR90-1202.pdf  (14 pages)
  Inspection Forms/PRF SR90-1202 Rev A.xlsx  updated
  Inspection Forms/PRF SR90-1202 Rev A.pdf  exported
  logs/PO26-1085_SR90-1202_20261020-1012_manifest.json  written
  2.1 Completed | 2.2 N/A | 2.3 N/A | 2.4 N/A | 2.5 Yes | 2.6 No
  Section 3: N/A
  Disposition: Raw Goods
  Flags: None
  PDF order: 1.PO | 2.PackingSlip | 3.ReceivingList | 4.Drawing Rev A | 5.PRF | 6.MfgCOC | 7.SterCOC

 Next steps:
  - Review suggested Sections 2/5/6 — edit freely before signing
  - Resolve any flags before routing
  - Engineer sends compiled PDF for signatures (not automated)
====================================================
```

## Reading the result

| Status / output | Meaning | What to do |
|---|---|---|
| `Flags: None` | Every cross-check passed | Review, sign, route |
| `⚠ FLAGS (n)` with `MISMATCH …` | Two sources disagree (part#, rev, lot, qty, dates, vendor part#, Section 1) | Resolve from the source documents. Section 5 is drafted, and 6.1 = Yes when there's a real discrepancy |
| `FLAG` on a Section 2 line | The rule couldn't decide, so **both boxes are left empty** | Decide and tick the box yourself |
| `QTY SHORTAGE (WITHIN 10% TOLERANCE)` | Informational. A shortage of 10% or less is treated as a soft flag | Record any acknowledgement in Section 4 (yours) |
| `HELD: …` | Multi-part PO, and this part has no documents of its own while its siblings do | Get the vendor paperwork. This isn't a tool error |
| `ERROR` / `BLOCKED` | Structural problem: Section 1 mismatch gate, missing logo, checkbox verification failure, PRF already hand-edited | Read the message. Bring structural problems to whoever maintains the tool |
| Run halts before any part | **PO revision conflict**: two PO copies disagree on rev or qty | Decide which PO governs, remove or rename the other, and re-run |

## After the run: the engineer's part

- Open the compiled package. Check the receiving-list highlight covers the whole row for your lot, and that callout bubbles sit on the **received** revision drawing.
- Resolve every flag. Edit the Section 5 draft freely, because the signed text is yours.
- Confirm Section 3 (your measurements) and Section 4 (your comments) are complete. The tool never writes to them.
- Sign in Adobe and route by email (see [Outlook](integrations/outlook.md)):
  - **2.6 = Yes** (confidential biologic batch records): the **Director of Operations** signs first
  - **2.4 = Yes**: Quality Manager initials
  - **Section 6 discrepancies**: the **Engineering Reviewer** signs before the Quality Manager
  - **Quality Manager** signs last
- When fully signed, file it in the PO's `Completed Receivings\`.

## Manual vs. Claude-assisted runs

| Situation | Do this |
|---|---|
| Routine PO, PRFs filled | Run it yourself. No AI session needed |
| Completes with flags | Still just you. The flags are the tool declining to guess |
| Exception, `ChecklistCheckboxError`, `UserContentError`, a whole part blocked | Bring it to a Claude Code / Cowork session. Something structural needs diagnosis |
| New vendor or unfamiliar document layout | Run once. Escalate only if the output looks **wrong**, not merely flagged |
| You want the logic changed | Always change it through the WI and code together, then add a test case |
| A PO genuinely doesn't fit | A per-PO edge-case script (outside the validated path). Fold anything recurring back into `receive.py` |

## Common gotchas

- **Windows console crash on ✓/⚠ characters.** Fixed: `main()` forces UTF-8 on stdout. On very old setups, `set PYTHONUTF8=1`.
- **"PO document not found".** The PO PDF must contain the PO number in its name. Vendor *sales-order acknowledgements* are deliberately not treated as a second PO.
- **Excel checkboxes stopped being clickable.** Only an old PRF saved with plain openpyxl does this. The current save path restores form-control wiring. See [04 · Pipeline → Step 8](04-pipeline-steps.md#step-8--fill-prf-sections-2-5-6).
- More in [09 · Troubleshooting](09-troubleshooting.md).
