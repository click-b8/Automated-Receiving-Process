---
tags: [workflow, qe-workflow]
---

# Physical Inspection, Dimensions and Testing

Part of [[QE Receiving Workflow]] · [[W04 - Staged for Inspection (Monday)|← W04]] · [[W06 - Compile with the Automation|W06 →]]

**Who:** QE / Engineering (SOP Table 1, step 3)  
**Where:** Quarantine inspection area, well lit (SOP §6.5)

## Physical work
- Sample per the AQL, or inspect 100%.
- Measure the numbered critical dimensions on the drawing with calibrated tools.
- Bench / DVT testing where required. For sterile packaging, check for foreign objects with the TAPPI chart (over 0.25 mm² is rejected, SOP §6.6).
- Record the measurements in Section 3. Record comments, DVT reference and calibrated-tool info in Section 4. Save the PRF (xlsx + pdf).

## Records
- PRF Section 3 (measurements)
- PRF Section 4 (engineer comments)

## Automation embedded in this step
- [[Step 5 - Parse Drawing with Cache]]

## Watch-outs
- A rejection opens an NCR (SOP §6.7, SOP 8.3). Section 4 is never written by automation. Claude sessions can set up the Section 3 table from the drawing (`fill_section3_table()`); the run checks the calibration date against the inspection date.
