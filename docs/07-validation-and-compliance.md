# 07 · Validation and compliance

> Not legal or regulatory advice. This is how the project approaches validation. Your Quality and Regulatory functions decide what's sufficient for your QMS.

## Regulatory basis

Software used as part of production or the quality system has to be validated for its intended use:

- **FDA QMSR (21 CFR Part 820, effective 2 Feb 2026)** incorporates **ISO 13485:2016** by reference. The software-validation expectation is now **ISO 13485 §4.1.6** (QMS software) and **§7.5.6** (process validation), read through the QMSR. The project began under the old §820.70(i) *automated processes* wording, and **documents that still cite 820.70(i) should be updated to the QMSR / ISO 13485 references.** An auditor will notice outdated citations.
- **FDA Computer Software Assurance (CSA)** guidance supports a **risk-based** approach: assurance effort matched to the risk of the software's role, with credit for unscripted and exploratory testing where risk is lower.
- **21 CFR Part 11** applies to the electronic records and signatures. The signed release record is the Adobe-signed package in `Completed Receivings`. The signature controls belong to that system, not to this tool.

## Intended use (as written for validation)

> Assist the receiving QE by compiling vendor and receiving documents into one package per part, cross-checking them, and pre-filling PRF QF 7.4.3-02 Sections 2, 5 and 6 as a **draft for human review**. The tool does not approve, release or disposition product.

That framing sets the risk. A wrong Section 2 answer would be a wrong entry on a release record, so it is **mitigated by design**: deterministic rules, flag-don't-guess, mandatory human review and unchanged signatures. Validation focuses on those controls.

## The validation package (QF 7.4.3-06)

| Document | Role |
|---|---|
| **Validation Protocol** | Scope, intended use, risk, acceptance criteria, test strategy (*living*: test cases are added as real POs exercise new paths) |
| **Attachment A: Test Case Matrix** | One row per scenario: **TC-ID, scenario, expected result, must-pass, PO reference, result, tester, date** |
| **Attachment B: Instruction Guide** | Plain-language quick start for colleagues |
| **WI-7.4.3-01 (`CLAUDE.md`)** | The controlled spec the code implements. Stays *Revision 02 (Development)* until the Open Items needed for Rev A are closed |

### Test strategy

- **Real POs, not synthetic fixtures.** Each test case is tied to a real PO run, and its PO reference is recorded. The dashboard shows which processed POs aren't yet mapped to any test case (*POs not in matrix*).
- **Fix → test case.** Every defect is fixed in code **and** gets a test case, and the WI and vault are updated together.
- **Must-pass cases** cover the controls that protect the record: the Section 1 gate, file-access boundary, sibling-PO isolation, native-checkbox integrity, the no-clobber guard, the PO-revision halt, confidential-record routing and missing-COC detection.
- **Real-file smoke test.** A fix verified on an isolated copy also has to pass one run against the real, unmodified file before it's called done. This rule exists because two shipped bugs were hidden by isolated test copies.

### Example scenarios (from Attachment A)

| Area | Scenarios |
|---|---|
| Happy paths | single part; multi-part PO; multi-batch / partial receiving; same part + rev across two lots |
| Document problems | missing required COC; multiple materials/finishes needing multiple COCs; COC bundled inside the packing slip; scanned packing slip with the signature read by vision; COC dated after receipt; vendor paperwork naming another customer or PO |
| Drawing | revision mismatch (ordered vs. received); no drawing on file; drawing only inside the PO PDF; stadium/oval callouts |
| Rules | confidential biologic records → Director review first; sterilization-only service vendor; explicitly unsterilized lot from an always-sterile vendor; disposition by catalog lookup |
| Integrity | native checkboxes after a plain run; re-run against a hand-corrected PRF; sibling-PO isolation; PO#/year mismatch |
| Process | quantity short within tolerance; two PO revisions disagreeing; multi-part PO with one line held; large 10+ part PO signed off part by part |

## Things an auditor is likely to ask, and the honest answers

| Question | Where the project stands |
|---|---|
| *Is the tool validated?* | Not yet. WI is Rev 02 (Development), and Open Items block Rev A. Until then every output is reviewed as a draft and signed by people under SOP 7.4.3, exactly as before. |
| *How do you control the version that produced a record?* | `SCRIPT_VERSION` doesn't yet track WI revisions (open item). The run manifest records WI revision, timestamp, inputs and outputs per part, but it's a **development aid kept outside Box**, not a quality record. **Before Rev A: put `receive.py`/`prf_helpers.py` under document control (hash or tag per release) and print it in the package or PRF footer.** |
| *What about the per-PO scripts?* | Edge-case scripts run **outside** the validated path. Quality's decision was to fix recurring root causes in `receive.py` rather than create a review procedure. Each use should be visible (the dashboard tags those POs), and the count should trend down. |
| *Can the AI change a quality decision?* | No. Section 2 and the disposition are deterministic code. Model output is data, vision findings are soft, and the WI file-access rule limits Claude Code sessions to one PO folder. |
| *Who can change the rules?* | Rule changes go through the WI (controlled), code, a test case and the vault. Verbal instructions that became code are marked "not yet SOP-ratified" in the Open Items list until Quality ratifies them. |
| *Are records ever silently overwritten?* | The no-clobber guard refuses to regenerate a package a person has edited. `--force` exists for explicit operator instruction only. |
| *Data integrity of the PRF file itself?* | Native controls and the logo are verified after every save (`verify_native_checkboxes`, `verify_xlsx_package_integrity`). The 2026 checkbox incident is documented with root cause, fix and affected-record review. |

## Known gaps worth closing before Rev A

- Update regulatory citations from §820.70(i) to QMSR / ISO 13485 §4.1.6.
- Version-stamp the tool into its outputs.
- Record the **extraction method** (pdfplumber / model / cache) per fact in the manifest.
- Ratify the verbally-instructed rules (10% shortage tolerance, 2.4 N/A + flag default, QDA → 2.4/2.5 cross-check, disposition FLAG on a catalog hit).
- Decide whether to keep or retire test cases written for a deployment structure that was never built.
- Confirm the approved-destination catalog JSON was transcribed correctly from the controlled list, and define how it's refreshed when that list is revised.
