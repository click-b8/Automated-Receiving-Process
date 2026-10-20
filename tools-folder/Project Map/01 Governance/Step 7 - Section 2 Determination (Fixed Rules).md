---
tags: [governance, step, deterministic]
---

# Step 7 — Section 2 Determination — FIXED RULES ONLY

Part of [[WI-7.4.3-01 Overview]]. **The core design principle of the entire project**: these are deterministic Python functions (`determine_section2()` and its sub-functions in [[receive.py]]), never model judgment calls. Every input is a fact already extracted in Steps 1/4/5/6. If a required fact is missing, the function returns a flag — it never guesses.

## The six checkbox functions

| # | Function | Resolves to |
|---|---|---|
| 2.1 | `determine_2_1()` | Completed / FLAG — needs packing-slip ops sig+date AND receiving list present |
| 2.2 | `determine_2_2()` | N/A / Yes / No — material+finish cert coverage |
| 2.3 | `determine_2_3()` | N/A / Yes / No / FLAG — dimensional callouts pass/fail |
| 2.4 | `determine_2_4()` | Yes / N/A (see below) — traveler review; Biologic+vendor and VendorD special-cased |
| 2.5 | `determine_2_5()` | N/A / Yes / No — sterilization COC; also driven by Section 1.3 QDA (see below) |
| 2.6 | `determine_2_6()` | Yes / No — confidential Biologic docs, Director of Operations reviews before Quality Manager if Yes |

`"FLAG"` means: leave **both** checkboxes unchecked, add a Section 5/summary explanation. Never pick a default answer to fill a checkbox when the function can't resolve one.

## 2.4 default — CHANGED 2026-09-11, reconfirmed 2026-09-14 ([[Open Items (Master List)#Item 18]])

Both the generic catch-all AND the VendorD branch used to resolve a missing
traveler to a bare **"No"** — a settled negative finding, not a
documentation gap. Confirmed with Quality after two real, repeated engineer
corrections (PO26-1005/PO26-1009 SterilizerCo; PO26-1008 VendorD): a vendor that
assembles nothing and supplies no document actually labeled a traveler now
resolves to **N/A + explicit flag** instead — same shape everywhere else
this codebase uses FLAG, never a bare negative that could be mistaken for
"confirmed not applicable." Confirmed Biologic vendors (VendorB/VendorA/VendorC)
are unaffected — still `Yes` regardless of traveler availability
(2026-07-27 fix, [[Open Items (Master List)#Item 1]]). **Not yet formally
SOP-ratified** — see the Open Item for what still needs confirming before
Rev A.
**Reconfirmed 2026-09-14 (QE, direct instruction):** VendorD never provides
travelers for any lot — this branch's "no traveler" case is not an
exception, it is the only case VendorD will ever hit — so 2.4 for VendorD must
always be N/A, never "No." `receive.py`'s `is_vendord` branch already did this;
the actual gap was the bespoke `process_PO26-1008.py` script bypassing
`determine_2_4()` and hardcoding "No" — fixed to "N/A" that same day. See
[[Real-World Runs (Case Studies)]].

## Disposition rule — REPLACED 2026-09-11, CORRECTED 2026-09-14 ([[Open Items (Master List)#Item 13]])

Long history, now resolved by a completely different mechanism than either
of the two heuristics that came before it:

1. Originally: an over-broad default (anything not Biologic/VendorD/sterilized
   → Finished Goods) — misclassified an ordinary packaging component.
2. **2026-07-27**: replaced with a `"label" in description` keyword check
   (direct QE correction, PO26-1006/VendorK).
3. **2026-09-08**: that keyword check confirmed **unsafe in both
   directions** on PO26-1007/VendorE — blank, unprinted label stock is Raw
   Goods (the real test is *printed* product labeling, not the word
   "label" itself); this part only got the right answer by two accidents.
4. **2026-09-11**: keyword check retired entirely. `determine_section2()`
   now looks the part number up directly against a QF 7.2-01 Approved
   Product Destination List (APDL) catalog
   (`qf_7.2-01_apdl_catalog.json`, 434 entries): **not on the APDL →
   Raw Goods**; **on the APDL → FLAG** (both checkboxes left unchecked,
   naming the APDL's own description) — being catalog-listed only proves
   this catalog number *can* be a finished item, not that *this lot* is one
   rather than an earlier-stage Raw Goods receipt of the same part number,
   justified at the time by "the ProductS facet screw appearing at more than one
   stage across PO26-1008/PO26-1001."
5. **2026-09-14 correction (QE, direct instruction):** that justification
   was wrong — PO26-1008's Raw Goods answer was a 2026-09-10 engineer
   misjudgment, not a genuine multi-stage receipt. VendorD anodizes AND
   laser-etches these screws before shipping them back; what ACME Medical
   receives from VendorD already IS the finished, shelf-ready catalog item.
   `determine_section2()` now special-cases **on the APDL + VendorD vendor →
   Finished Goods** (the ProductS facet screw family and any future VendorD
   anodized/laser-etched screw matching that profile); on the APDL + any
   other vendor, the FLAG behavior from step 4 is unchanged. This does not
   touch the real, signed PO26-1008 PRF itself — code/documentation only.
   One consequence worth tracking: the FLAG branch now has **no** confirmed
   real-world positive example left (every other real confirmation resolved
   via "not on the APDL," not an actual hit) — **still needs Engineering Reviewer/Quality Manager
   to ratify the FLAG-on-APDL-hit behavior for non-VendorD vendors**, and to
   confirm the catalog JSON was transcribed correctly from the live
   document.

## Section 1.3 QDA → 2.4/2.5 cross-check (Open Item #4, 2026-07-20)

A Section 1.3 QDA of **YES** now feeds `determine_section2()`, raising a flag on 2.4/2.5 if a traveler/sterilization COC is missing, rather than letting the vendor-only heuristic silently resolve to N/A. Added from a verbal engineering instruction — **not yet tied to an approved SOP/WI revision**. See [[Open Items (Master List)#Item 4]].

## Biologic / vendor special cases

`BIOLOGIC_VENDORS = {VendorB, VendorA, VendorC}`, matched on product line keywords OR "biologic" in description. `TRAVELER_VENDORS` (genuine contract-manufacturer travelers) vs. everyone else (VendorF/VendorG "traveler"-named docs are really `mfg_inspection_report`, see [[Step 2 - Match & Classify Documents]]).

## Feeds into

- [[Step 8 - Fill PRF Sections 2, 5, 6]] — writes these six determinations, plus disposition, to fixed cells (cell map itself corrected 2026-09-11 — see that page)
