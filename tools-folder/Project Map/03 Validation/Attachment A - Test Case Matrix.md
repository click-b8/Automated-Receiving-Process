---
tags: [validation, controlled-document, living-document]
---

# QF 7.4.3-06 Attachment A — Test Case Matrix

Lives in `Validation\QF_7.4.3-06_Attachment_A_Test_Case_Matrix.docx` at the vault root. **A living document** — rows are added/dated/filled progressively as real POs are processed, not pre-populated with synthetic results (per [[QF 7.4.3-06 Validation Protocol]] Section 12.0 Test Strategy).

## 2026-09-11 update — no longer untouched

Previously left deliberately untouched by this vault ("actively maintained
elsewhere"). Updated this pass:

- **TC-05** filled in: PO Ref `VendorB | 26-1015`, Result `Pass`, Tester
  `QE`, Date `9/9/2026`.
- **TC-19 through TC-26** added — new rows for: native-checkbox integrity
  after a plain unmodified run; multi-part PO partial hold; accepted
  quantity-shortage tolerance; PO-revision conflict with no other document
  resolving it; non-assembling service vendor with no traveler (2.4
  default); disposition via APDL lookup rather than a keyword; stadium/oval
  drawing callouts; and re-running against a PRF the engineer has already
  hand-corrected and saved (the native-checkbox/logo permanent-loss
  scenario).
- **Three note rows appended** (not edits to existing TCs, per the
  document's own "insert additional rows as needed" convention):
  - **Note (TC-05)** — this run's drawing order was compiled
    received-rev-first then ordered-rev-first, the REVERSE of Step 3's
    documented order, per the engineer's explicit one-off instruction.
    Flagged as a possible WI amendment question, not a Fail.
  - **Note (TC-17)** and **Note (TC-18)** — both test an
    `_Automation\tools\`/`_Automation_Master\tools\` deployment structure
    that the 2026-09-08 finding (and every session since) confirms doesn't
    exist — see [[File Access Rules]], [[Folder Structure]]. Recommends
    Quality decide whether to retire these test cases, rewrite them to
    describe the real Desktop-folder structure, or keep them pending until
    the master/yearly-copy deployment is actually built.

## 2026-10-14 (later) — tester confirmation and two Fails closed

- QE confirmed every Proposed result as **Pass**, plus TC-34, TC-36 and TC-41. The rules behind TC-34 and TC-36 are still open as Open Items #27 and #25.
- **TC-25** (stadium/oval callouts): original Fail kept; **Pass on rerun** with 26-1023 (8 stadiums found and numbered, package signed off), 26-1036 and 26-1044. These ran through `find_stadiums.py` in the per-PO scripts. `receive.py`'s own oval detection is still unverified on a VendorD drawing (Open Item #23).
- **TC-26** (rerun against a hand-edited PRF): original Fail kept; **Pass on rerun** with 26-1023 (three compiles against the engineer's saved PRF, checkbox and package guards all ran, package signed off), via `process_PO26-1023.py`.
- Matrix now: **36 of 44 Pass**, 34 of 40 Must-Pass. Still open: TC-09, TC-15–18, TC-27, TC-28 (not run) and TC-23 (on hold).

## 2026-10-14 update — edge cases from processed POs

Pre-change copy: `_backup\QF_7.4.3-06_Attachment_A_Test_Case_Matrix.pre-2026-10-14.docx`. Nothing deleted.

**Every result added on 10/14 is marked "Proposed"** (a session's reading of the run manifests, per-PO script findings and the WI changelog). Each one becomes a real Pass or Fail only after the tester checks it against the signed package in Completed Receivings and initials and dates it. The [[_Hub - Validation|dashboard]] counts these separately as *Proposed*, not Pass.

- **TC-07** (no ACME drawing) → VendorE 26-1007, proposed Pass; confirm 2.2 and 2.3 = N/A on the signed PRF.
- **TC-14** (manifest) → proposed re-test evidence from VendorB 26-1004: two manifests in Desktop `logs\`, no overwrite.
- **New TC-29 to TC-42:**

| TC | Edge case | PO(s) | Status |
|---|---|---|---|
| TC-29 | Engineer finds out-of-spec dims; no vendor FAI | 26-1030, 26-1027 | Proposed |
| TC-30 | Compiled before packing slip / receiving list exists | 26-1027, 26-1030, 26-1032 → 26-1033 | Proposed |
| TC-31 | Partial shipment across lots, balance open, no lots on slip | 26-1023, 26-1027 | Proposed |
| TC-32 | Scanned packing slip, signature read by vision | 26-1004 | Proposed |
| TC-33 | COC bundled in packing-slip PDF | 26-1004, 26-1007 | Proposed |
| TC-34 | Drawing only inside the PO PDF | 26-1004, 26-1045 | On hold (Open Item #27) |
| TC-35 | Vendor sales-order acknowledgement ≠ PO copy | 26-1004 | Proposed |
| TC-36 | UnSterilized lot from always-sterile vendor | 26-1004 | On hold (Open Item #25) |
| TC-37 | Sterile part with no ster cert | 26-1032, 26-1045 (+ 26-1033 question) | Proposed |
| TC-38 | Biologic when the PO page is unreadable | 26-1004, 26-1045 | Proposed |
| TC-39 | Vendor paperwork for a different customer / part no. | 26-1032, 26-1033 | Proposed |
| TC-40 | COC material ≠ drawing; COC dated after receipt | 26-1046 | Proposed |
| TC-41 | Large multi-part PO (17 parts) | 26-1010 | Candidate, not run |
| TC-42 | Rework shipment, vendor lot ≠ engraved lot | 26-1030 | Proposed |

**Worth Quality's attention:** on 26-1033 the signed PRF has 2.5 = N/A while the drawing label still reads STERILE EO (see TC-37).

## 2026-10-05 update — manifest decision and Reasoning Reference

All changes made by strike-through, never deletion; italic text beside each
struck passage is the replacement. Pre-change copy:
`_backup\QF_7.4.3-06_Attachment_A_Test_Case_Matrix.pre-2026-10-05.docx`.

- **Intro:** a dated revision note added under the opening paragraph.
- **TC-01 / TC-02:** manifest clauses struck and moved to TC-14. Their
  7/27–7/28 Passes stand for every other criterion; the manifest clause
  could not have been met then (no manifest before 2026-09-17).
- **TC-14:** "shared placement" and the `_Automation\logs\` expected output
  struck; replaced with Desktop `logs\` placement and the real Step 9 fields.
  Must-Pass changed Yes → No (development aid, not a release criterion).
  Result "Pass" struck → **re-test required**. The 7/28 Pass could not have
  been valid; CLAUDE.md's 2026-09-17 changelog says TC-14 was corrected that
  day, but the correction never reached the .docx — recorded now.
- **TC-15:** evidence source "manifest/log" struck → console output and
  session transcript (the manifest never recorded file access).
- **TC-27 (new):** Section 5 text drafted per the
  [[Comment & Discrepancy Reasoning Reference]] — blocked on
  [[Open Items (Master List)#Item 31|Item 31]] and Item 15.
- **TC-28 (new):** harvest access boundary and read-only behavior.
- **Four note rows appended:** Note (TC-14), Note (TC-01/TC-02),
  Note (TC-15), Note (TC-27/TC-28), each explaining why text was obsoleted.

## Known Must-Pass-100% cases referenced elsewhere in this vault

- ~~**TC-01, TC-02, TC-14** — currently cannot Pass as written; each requires manifest existence/content that [[Step 9 - Run Manifest (Desktop Dev Aid)|doesn't exist yet]]~~ *2026-10-05: manifest clauses struck from TC-01/TC-02 and moved to TC-14; TC-14 revised to Desktop placement, no longer Must-Pass, re-test required*
- **TC-27, TC-28** — new 2026-10-05, both Must-Pass-100% (Section 5 language; harvest access boundary)
- **TC-13** — [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)|Section 1 halt-and-flag behavior]]
- **TC-04a** — the [[Open Items (Master List)#Item 1|Section 2.4 edge case]] (original confidential-Biologic question, now resolved 2026-07-27)
- **TC-15/TC-16** — boundary tests, deliberately staged with real historical documents rather than occurring naturally
- **TC-17, TC-18** — structural-premise mismatch, see the note above
- **TC-19, TC-26** — new Must-Pass-100% cases from the 2026-09-11 batch (native checkbox integrity, PO-revision conflict)

## Related

- [[QF 7.4.3-06 Validation Protocol]] (parent document)
- [[Real-World Runs (Case Studies)]] (the actual PO runs that map onto rows in this matrix)
- [[Open Items (Master List)]] (every new TC row above traces to a numbered open item)
