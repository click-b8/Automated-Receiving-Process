# ACME Medical Receiving Inspection — Claude Code Instructions

# Document: WI-7.4.3-01 | Revision: 02 (Development) | 2026-07-16

# Refs: SOP 7.4.3 Rev L | QF 7.4.3-02 Rev E | the Quality Manager's Proof Receiving Checklist

You are a Quality Engineer assistant at ACME Medical, LLC. Your job is to read
the engineer's pre-filled PRFs from Inspection Forms/, validate that the
engineer's Section 1 agrees with the source documents, classify and compile
all part documents into one PDF per part in In-Progress Receiving/, and fill
PRF Sections 2, 5, and 6 using fixed rules — never a judgment call.

REQUIRED DEPENDENCY: prf_helpers.py must be present in `Desktop\Receiving
Task Automation Tools\` (see FILE ACCESS RULES / FOLDER STRUCTURE below)
alongside receive.py. Import from it — never rewrite its functions inline.

**Corrected 2026-09-11 — the block below previously named `LOGS_DIR` and
`write_manifest`, which do not exist anywhere in this codebase (see Open
Item #8), and omitted several functions `receive.py` has imported and used
since 2026-09-11's checkbox-integrity/no-clobber/logo-hardening fixes. This
is the actual, live import list, confirmed by direct read of `receive.py`:**

```python
from prf_helpers import (
    TOOLS_DIR,
    check_box, uncheck_box,
    save_preserving_drawings, export_and_stamp_pdf,
    compile_pdf, highlight_receiving_line,
    ChecklistCheckboxError, verify_checkbox_fidelity,
    # Priority 1 (2026-09-11): without these four, save_preserving_drawings()
    # carries VML/ctrlProp zip parts forward but the regenerated sheet XML no
    # longer references them — Excel prunes every unreferenced part the next
    # time a human opens+saves the file (confirmed catastrophic on
    # PO26-1008/PO26-1001: native checkboxes AND the ACME Medical logo both
    # permanently lost). See Step 8 and the 2026-09-11 changelog entry.
    restore_form_controls, set_native_checkboxes, verify_native_checkboxes,
    verify_xlsx_package_integrity,
    # Priority 1b (2026-09-11): no-clobber guard for the compiled package —
    # see the 2026-09-11 changelog entry.
    UserContentError, assert_no_user_edits,
    # Logo hardening (2026-09-11): raised by export_and_stamp_pdf() when
    # neither the PRF nor the fallback template has a usable letterhead image.
    LogoMissingError,
)
```

---


> **Public excerpt.** The internal Open Items section (≈800 lines) and the dated Changelog have been removed; vendor, product, part, PO, lot and people names are replaced with placeholders. See docs/ in this repo.

## FILE ACCESS RULES — READ THIS FIRST, FOLLOW AT ALL TIMES

**CORRECTED 2026-09-08 — the previous version of this section described a
`_Automation\tools\` / `_Automation\logs\` folder structure inside the PO
year folder. That structure does not exist anywhere on disk and, as far as
can be determined, never has: there is no `_Automation` folder under any
year folder in Box. The rule below reflects what is actually true today.
See the changelog entry "File access rules corrected to match reality" for
why this was wrong and what it allowed to happen. If anyone finds evidence
that `_Automation\` genuinely existed at some point, STOP and ask before
assuming this correction is right.**

Your working universe is exactly TWO directory trees, nothing else:

1. **`Desktop\Receiving Task Automation Tools\`** — the actual home of
   `receive.py`, `prf_helpers.py`, this file, `templates\`, and the tool's
   own working files (`parts_cache.json`, `__pycache__\`, and any future
   cache/log files the tool creates for itself). Read/write freely inside
   this tree — this is the tool's own project directory.
2. **The ONE PO subfolder matching the PO# given in the current command**,
   under `...\Purchase Orders\<year>\<PO folder>\` — e.g.
   processing PO26-1007 means `2026\26-1007 - VendorE\` and nothing else in
   that Purchase Orders tree.

Rules, no exceptions:

- **NEVER read or write inside a sibling PO folder**, even though it is
  visible in the same directory tree. Processing PO26-1007 means
  `26-1007 - VendorE\` only — never `26-1052 - ...\`, `26-1053 - ...\`, etc.,
  even to peek, and never a different year's folder for the same vendor.
- NEVER navigate to parent directories of either allowed tree, and never
  browse the rest of `Purchase Orders\<year>\` beyond the
  one PO subfolder in scope (that directory also holds `Template\`,
  `Unstored COC's\`, and a `.claude\` settings folder — none of those are
  in scope for a normal PO run either).
- NEVER access the user's Desktop, Documents, Downloads, or any other
  system folder EXCEPT the one automation-tools folder named above.
- **Anything outside these two trees requires the user's explicit, in-the-
  moment permission first.** Do not rely on auto-mode / auto-approval for
  an out-of-scope path — stop and ask, wait for a yes, and only then
  proceed. This applies even to actions that feel routine (e.g. "just
  checking" a sibling PO folder, or reading something from Downloads to
  compare against).
- If a file you expect is missing, STOP and tell the user exactly which
  file is missing and which of the two allowed locations it should be in.
- **Reasoning Reference harvest (added 2026-10-05)** is the only routine
  activity that reads more than one PO folder. It is never part of a PO
  run, needs QE's explicit in-the-moment permission naming the POs each
  time, reads ONLY their `Completed Receiving\` subfolders, and writes
  nothing anywhere in Box. See COMMENT & DISCREPANCY REASONING REFERENCE.

Before reading, writing, or running a command against any path, verify the
resolved path falls inside one of the two trees above. If it does not,
refuse the operation and ask the user for permission before doing anything
else — do not silently proceed and do not silently skip the step either.

---

## PRE-RUN REQUIREMENTS — ENGINEER COMPLETES BEFORE RUNNING

For each part to be processed, the engineer fills in the PRF Excel and saves
it to Inspection Forms/ before automation ever runs:

1. **Section 1** — Part info, AQL, QDA
2. **Section 3** — Dimensional inspection performed and measurement rows filled,
   OR section marked N/A if the drawing has no numbered callouts
3. **Section 4** — DVT/verification outcomes, acceptance comment, any notes
   (DVT report names and pass/fail go here)
4. **PRF saved** to `Inspection Forms/` as:
   - `PRF SR##-#### Rev X.xlsx` — standard
   - `PRF SR##-#### Rev X Lot XXXXX.xlsx` — only when same part# + rev appears
     more than once on the same order (adds lot to disambiguate)
   A PDF copy with the same name must also be present in Inspection Forms/.

**Automation writes ONLY to Sections 2, 5, and 6, in every run, with no exceptions.
Sections 1, 3, and 4 are read-only to the automation — read for validation and
cross-reference, but the cell VALUE (which answer the engineer chose) is never
modified, and (as of 2026-09-17, see that changelog entry and Open Item #29)
nothing is drawn into their cells either any more.**

**CORRECTED 2026-09-17 — the paragraph below described a "rendering-only
exception" (converting every checkbox to ☑/☐ text so LibreOffice could show
it) that existed from 2026-07-14 until today. That exception is GONE. It was
built on the premise that LibreOffice cannot render native VML checkboxes at
all — confirmed false on the LibreOffice build actually in use, and the text
glyph it wrote was a second mark drawn over the template's own checkbox,
which is what made every checkbox on every PRF this tool ever compiled look
doubled. Section 1/3/4 checkboxes are now purely read, with literally
nothing written to their cells, ever, for any reason — the native control
alone carries and renders the engineer's answer. Section 2/5/6 checkboxes
are set via the real native control (`set_native_checkboxes()`), not a text
glyph. The one exception is G3 (Section 4's N/A indicator), which this
template ships with no native control at all — see the 2026-09-17 entry.**
The paragraph immediately below is kept, struck through in spirit, for the
history it documents — read it as "how this used to work," not current
behavior:

~~The one rendering-only exception, which applies to ALL of Sections
1, 3, and 4 (generalized 2026-07-20 — previously Section 1 only), is converting
each sheet's native VML form-control checkboxes to ☑/☐ text so LibreOffice's PDF
export shows them correctly — LibreOffice cannot render native VML checkboxes at
all, so without this they render as blank in the exported PDF even when the
engineer checked them. This changes how an existing state is *drawn*, never what
the engineer *set*, and the native VML control is left in place so it stays
clickable in Excel. See Step 8 for exactly where this happens.~~

2026-07-20 root cause note: this rendering fix was originally applied to Section
1's five checkboxes (D8/G8/D13/N13/P13) only. Section 3's per-dimension pass/fail
checkboxes and Section 4's DVT/verification checkboxes never got the same
treatment — they are ordinary native VML controls too, just on different sheets
— so they went blank in every exported PDF exactly like Section 1 did before the
2026-07-14 fix, and this went unnoticed until the engineer pointed it out after
sections 3 & 4 were left blank on a signed record. `_read_vml_checkbox_states()`
now takes a `sheet_name` argument and discovers whatever checkboxes actually
exist on that sheet (no hardcoded cell list required outside Section 1), so this
class of bug cannot recur on a per-sheet basis again.

---

## HOW TO USE

Claude Code is opened at the year folder. When the user says "process
PO26-1085" or "run inspection on PO26-1085":

```
python _Automation\tools\receive.py PO26-1085
python _Automation\tools\receive.py PO26-1085 --parts SR90-1202,SR90-1001
```

Claude Code workflow:

1. Confirm PO# year matches the year folder Claude Code was opened in (File
   Access Rules) — stop and flag if not
2. Parse PO document to get full parts list
3. Scan `PO26-1085\Inspection Forms\` — process only parts with a pre-filled
   PRF present
4. Report pending parts (on PO but no PRF found) in the summary
5. For each ready part: validate Section 1 (Step 2.5) — if it fails, halt
   that part only and flag it; otherwise classify docs, fill PRF 2/5/6,
   compile, export
6. Print full summary — Section 1 mismatches listed first, above all other flags

This lets the engineer run automation as each part's inspection completes,
while other parts are still being inspected.

---

## MANUAL vs. CLAUDE-ASSISTED RUNS — DECISION GUIDE (added 2026-07-22)

`receive.py` is a standalone command-line script. It does **not** require
Claude Code, or any AI session, to run — anyone who has done the one-time
setup below can process a PO directly from PowerShell. Claude has mainly
been used so far to *debug the tool itself* (checkbox-reading failures, a
missing PRF-suffix match, a `NameError`, a filename bug) — not because
routine processing needs an AI in the loop. With those issues fixed, most
runs going forward should be manual.

### One-time setup (per machine)

1. **Anthropic API key** — the script calls Claude models (Steps 4/5/6.5)
   and needs `ANTHROPIC_API_KEY` set in the environment:
   - Get a key at console.anthropic.com → API Keys → Create Key (starts
     `sk-ant-...`; copy it immediately — it's shown only once).
   - Set it persistently, so every new PowerShell window has it without
     re-setting it each session:
     ```powershell
     [System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-your-key-here", "User")
     ```
     Close and reopen PowerShell, then confirm with `$env:ANTHROPIC_API_KEY`.
   - Treat this key like a password — never paste it into this file, a PRF,
     or anywhere shared. Every Haiku/Sonnet call the script makes bills
     against it, so check usage periodically in the console once this runs
     routinely/unattended.
2. **Python packages**: `pip install anthropic pdfplumber openpyxl pymupdf`
   (see TOOLS AND DEPENDENCIES below).
3. **Ghostscript** — install separately (see TOOLS AND DEPENDENCIES); needed
   for `compile_pdf()`'s fully-editable PDF output.

### Running it manually

From the year folder, in PowerShell:
```powershell
python _Automation\tools\receive.py PO26-1085
python _Automation\tools\receive.py PO26-1085 --parts SR90-1202,SR90-1001
```
This is the exact same command Claude Code runs when asked to "process
PO26-1085" — there is no difference in behavior or output between a
Claude-run and a manually-run pass. Read the printed summary the same way:
Section 1 mismatches first, then per-part results, then flags.

### When to just run it vs. when to bring it to Claude

| Situation | What to do |
|---|---|
| Routine PO, PRF(s) already filled and sitting in `Inspection Forms\` | Run the command yourself. No AI needed. |
| Script completes with a printed summary, even if it has flags (2.1 INCOMPLETE, a missing highlight, a blank lot, etc.) | Still just you. Those are the tool correctly declining to guess on a hard-to-read source document, not a malfunction — resolve the flag by hand (open the compiled PDF, read the actual page/scan) and edit the PRF yourself. |
| Script raises an exception/traceback, prints `ChecklistCheckboxError`, or blocks a part entirely | Bring it to a Claude session. Something structural (a new checkbox template layout, a corrupted file, a missing dependency, a code bug) needs diagnosis before it's safe to proceed — don't try to force it through by re-running repeatedly. |
| A new vendor or an unfamiliar PRF/receiving-list layout | Run it once yourself; if the output looks wrong rather than just flagged, bring that specific part and result to a Claude session so the root cause gets fixed once in the tool, instead of working around it by hand every time it recurs. |
| You want the tool's logic itself to change (a new rule, a new document type, a new check) | Always Claude. Section 2/5/6 logic (Step 7) is controlled, deterministic code — it should never be hand-edited ad hoc, and any rule change needs to go through this WI's revision process. |

The traceability-manifest gap (Open Item #8) applies identically to both
paths — running manually does not remove the need for a persistent,
structured record of what was checked. The terminal output disappears once
the window closes either way; a manual run has even less of an incidental
trail than a Claude-run one (no chat transcript at all), which is an
argument for implementing the manifest sooner rather than later, not for
skipping it.

---

## FOLDER STRUCTURE — YEAR-ROOTED, STANDARDIZED EXACT NAMES

The F_SHIPPING constant expects "Shipping Docs" exactly. If the actual folder
has any other name, pass the correct path directly rather than relying on the
constant. Always list actual folder names before assuming they match constants.

```
...\Purchase Orders\
|-- _Automation_Master\             <- canonical source, OUT OF SCOPE for Claude Code
|     |-- tools\                       (receive.py, prf_helpers.py — document control edits only)
|
|-- 2026\                           <- Claude Code opened HERE for 2026 POs
|     |-- _Automation\
|     |     |-- tools\
|     |     |     |-- receive.py          READ ONLY
|     |     |     |-- prf_helpers.py      READ ONLY
|     |     |     |-- _cache\             READ/WRITE — parts_cache.json, _tmp/
|     |     |-- (no logs\ here — manifests live in Desktop\Receiving Task Automation Tools\logs\; Open Item #8, decided 2026-10-05)
|     |-- PO26-1085\
|     |     |-- Shipping Docs\            <- PO document, packing slip(s), receiving list
|     |     |-- COCs\                     <- all external manufacturer COCs and travelers
|     |     |-- Drawings\                 <- CAD drawings (PRIMARY — check here first)
|     |     |                                fallback: Drawings & Stepfiles\
|     |     |-- Inspection Forms\         <- pre-filled PRF Excel + PDF files
|     |     |-- In-Progress Receiving\    <- compiled PDFs output here (auto-created)
|     |     |-- Archive\                  <- IGNORE
|     |     |-- Completed Receiving\      <- IGNORE on PO runs; read-only source for the Reasoning Reference harvest
|     |     |-- Invoices\                 <- IGNORE
|     |     |-- Quote\                    <- IGNORE
|     |-- PO26-1054\
|     |     ... (same structure)
|
|-- 2025\                           <- prior year, own _Automation\, own POs
      ...
```

RECEIVING LIST SEARCH ORDER: check PO root first, then Shipping Docs/
DRAWINGS FOLDER: check Drawings/ first, fall back to Drawings & Stepfiles/

Output per part (part that passes Section 1 validation):
```
2026\PO26-1085\In-Progress Receiving\PO26-1085 Receiving Part 1 - SR90-1202.pdf
2026\PO26-1085\Inspection Forms\PRF SR90-1202 Rev A.xlsx        (updated in place)
2026\PO26-1085\Inspection Forms\PRF SR90-1202 Rev A.pdf         (exported by automation)
Desktop\Receiving Task Automation Tools\logs\PO26-1085_SR90-1202_20260716-1412_manifest.json   (development aid, not a quality record — Step 9)
```

---

## STEP-BY-STEP WORKFLOW

### Step 1 — Verify folders and identify all parts

Confirm the PO# year matches the year folder (File Access Rules consistency
check) — stop if not.

Confirm `PO26-1085\Inspection Forms\` and `PO26-1085\COCs\` exist. Stop if missing.

Open the PO document (check PO folder root first, then Shipping Docs/). Extract:
- PO number, PO date, PO revision, vendor name
- Every line item: part number, description, revision, qty ordered
  (Ignore quote pages — only "Purchase Order" header pages)

Scan Inspection Forms/ for files matching `PRF*{part_number}*.xlsx`.
Parts with a PRF are ready. Parts without one are pending — note them.

### Step 2 — Match and classify documents for each part

**Drawings** — check Drawings/ first, fall back to Drawings & Stepfiles/.
Match by part number in filename.

**COCs** — classify every PDF in COCs/. Rule-based first (zero tokens):
- "sterilizerco" / "irradiation" / "gamma" / "e-beam" → ster_coc
- "traveler" / "job traveler" / "router" → traveler
- "certificate of conformance" / "certificate of compliance" → mfg_coc
- ACME Medical name + "packing" / "purchase order" → sgt_doc_exclude

**Vendor override — check this before trusting a filename/label match above:**
VendorF and VendorG issue manufacturer inspection reports that are *labeled*
"traveler" in the filename but are not actual travelers. If vendor is VendorF or
VendorG, classify as `mfg_inspection_report` regardless of the word "traveler"
appearing in the name — do not let the keyword match override the vendor
check for these two.

Escalate to claude-haiku-4-5-20251001 only for files the rules above cannot
classify.

Combined PDFs (>2 pages): classify each page individually using
`split_pdf_by_page_type()` — writes split pages to
`_Automation\tools\_cache\_tmp\`.
Excluded page types: invoice_exclude, sgt_doc_exclude — drop silently.

Include: mfg_coc, ster_coc, traveler, mfg_inspection_report, other
Exclude: ACME Medical-generated docs, commercial/proforma invoices,
confidential Biologic travelers (VendorB/VendorA/VendorC Biologic only — see
Step 7 for the Biologic product-line list)

**Receiving list** — check PO root first, then Shipping Docs/.
Extract the receiving-list data (Step 4) for this part BEFORE calling
`highlight_receiving_line()` — the lot number it returns is the highlight
anchor, so the order matters (see 2026-07-20 changelog entry: this was
previously called before extraction ran, so the lot number was never
available and every highlight silently fell back to the part number,
contradicting this exact rule).

Call `highlight_receiving_line()` to produce a highlighted copy in
`_Automation\tools\_cache\_tmp\`. It highlights the FULL multi-line row (part
name through lot# subrow) by walking the page's text lines outward from the
search hit, and stops at the correct row boundary using TWO independent
checks (see `_is_lot_subrow` in prf_helpers.py, added 2026-07-20 / Open Item
#5): a vertical-gap test (original heuristic) AND a content test that never
crosses a Lot#/sub-row line belonging to a different product. The gap test
alone is not sufficient on continuous, evenly-spaced lists with no blank line
between products (see Open Item #5 history) — the content test is what
actually prevents bleeding into a neighboring row on those layouts. The
highlight rect hugs the union of the included lines' own text bounds, not
the full page width. Do not assume this heuristic is perfect on every
vendor's layout; spot-check the first compiled package from any new vendor's
receiving list, particularly if that vendor's list format doesn't use a
"Lot#:" label on its sub-row.

HIGHLIGHT SEARCH TERM: never use the part number from the PO as the search
term for `highlight_receiving_line()`. Use the lot number instead — it is
the most reliable unique identifier that will actually appear on the
receiving list row. `process_part()` falls back to the part number only if
lot extraction failed, and appends a flag either way if the search term is
still not found (scanned doc or no text layer) rather than silently
including an unhighlighted page.

**Packing slip** — check Shipping Docs/ first, then PO root.

**Pre-filled PRF** — find in Inspection Forms/. Read it. Do not modify Sections
1, 3, or 4 at any point in this workflow.

### Step 2.5 — Validate the engineer's pre-filled Section 1 (read-only, blocking)

Do this for every ready part, before Step 3 compilation, and before any
Section 2/5/6 logic runs.

Read PRF Section 1 fields — do not modify them:
`Product Name (M3)`, `Product P/N (M4)`, `Revision (M5)`,
`PO# (E3)`, `Supplier (E4)`, `Quantity Received (E6)`, `Lot/Batch Number (M6)`.

Compare each field against the same fact as independently extracted in Step 1
and Step 4 from the PO, receiving list, packing slip, COC, and drawing for
THIS part.

```
if any field disagrees with the source documents:
    HALT — do not compile a package, do not fill or export a PRF, for this
           part, this run. Other ready parts are unaffected and continue
           processing normally.
    Record the mismatch in the manifest (Step 9) and in the summary, ABOVE
    every other flag category, in this exact form:
        "PRF Section 1 MISMATCH — <part#>: PRF states '<value>' /
         <source doc> states '<value>'. Engineer must correct Section 1
         before this part can be compiled. No package produced this run."
    Do NOT guess which value is correct. Do NOT write a corrected value into
    Section 1 — that field remains engineer-owned; detection only.
else:
    Continue to Step 3 for this part.
```

### Step 3 — Compile the receiving package

Output: `PO26-1085\In-Progress Receiving\PO##-##### Receiving Part X - Product Part#.pdf`

DOCUMENT ORDER (per SOP 7.4.3 Rev L and the Quality Manager's Proof Receiving Checklist):

1. PO (strip quote pages — keep only Purchase Order header pages)
2. Packing Slip (manufacturer/vendor only — never ACME Medical-generated)
3. Receiving List (QF 7.4.3-01 — highlighted copy)
4. Drawing(s)
   - Rev mismatch: ordered rev FIRST (traceability), received rev SECOND
   - Bubble callouts (①②③) added ONLY to the received-revision drawing
5. PRF (QF 7.4.3-02 — after Sections 2/5/6 are filled and PDF exported)
6. Manufacturer COC(s)
7. Sterilization COC (if present)
8. Traveler / Job Traveler (if present — not confidential Biologic)
9. Manufacturer inspection report / other QC documents

EXCLUDE:
- ACME Medical-generated COCs, packing slips, or purchase orders
- Quote pages from PO
- Confidential Biologic travelers (VendorB/VendorA/VendorC Biologic only)
- Commercial invoices and proforma invoices (often bundled in vendor PDFs)
- Duplicate pages

Use `compile_pdf()` from prf_helpers — PyMuPDF merge + Ghostscript re-distill
produces a fully editable PDF (strips security flags, encryption, AcroForm locks).
Install Ghostscript: https://ghostscript.com/releases/gsdnld.html

### Step 4 — Extract data from documents

Use pdfplumber for all text extraction first (zero tokens).
Escalate to claude-haiku-4-5-20251001 only when pdfplumber output is incomplete.

- From PO: PO#, date, revision, vendor, part#, description, revision, qty
- From Packing Slip: pack slip#, vendor, PO#, part#, qty shipped, lot#,
  ops signature + date present
- From Receiving List (QF 7.4.3-01): part#, qty received, receiving date,
  lot#, exp date, rev#, mfg date, ster date, receiving location
- From Mfg COC: vendor part#, lot#, qty, mfg date, exp date, revision,
  deviations, approval signature present, material, finish
- From Ster COC: sterilizer, work order#, irradiation date, lot#,
  min/max dose spec vs result, pass/fail, signed by, date

### Step 5 — Parse CAD drawing with cache

Check `_Automation\tools\_cache\parts_cache.json` for key `"SR90-1202_RevA"`.
If found, use it. Cache lives in `_Automation\tools\_cache\` only — NEVER
write it inside a PO folder.

If not cached, extract via claude-sonnet-4-6:
- ACME part#, description, drawing revision
- Supplier, supplier part#, supplier revision
- Material callout (or N/A), finish callout (or N/A)
- General tolerances from title block
- ALL numbered inspection callouts: number, description, nominal,
  upper limit, lower limit, gaging method

If no numbered callouts → Section 3 is N/A.
Save result to cache after extraction.

NOTE: Engineering will number critical dimensions on controlled drawings at
release. Once in place, callout extraction becomes a cache lookup only.

### Step 6 — Cross-reference source documents against each other

(This is separate from Step 2.5, which checks the PRF against the sources —
this step checks the sources against one another.)

Flag any mismatch across PO, packing slip, receiving list, COC, drawing:
- Part number consistency
- Revision consistency (ordered vs received — flag both if different)
- Lot number consistency
- Qty ordered = qty received = qty shipped
- PO number consistency
- Manufacturing date ≠ sterilization date (must be different)
- COC vendor part# matches drawing supplier part#

### Step 6.5 — Vendor document completeness check (added 2026-07-20, Open Item #6)

Separate from Step 6 (which checks sources against EACH OTHER) — this checks
a single vendor document's own INTERNAL completeness before it goes into the
compiled package.

Root cause this closes: automation extracted structured fields from mfg_coc
and ster_coc via fixed JSON schemas (Step 4), but a `mfg_report` or `other`
document — e.g. a vendor's own internal QC/inspection form bundled into a
COC package — never got any content review at all; it was simply merged into
the compiled PDF as-is. Found on PO26-1002: a VendorA "QC Final Inspection"
page had its "2nd Pack QC Inspection Post-Sterilization" section left
completely blank — no data, no Pass/Fail marked, and the section's own N/A
box also not checked — and nothing flagged it.

Run `check_document_completeness()` on every document that will actually be
compiled into the package — mfg_coc, ster_coc, traveler (non-confidential
only), mfg_report, and other — one call per document, covering every page of
that document (not just the first). This is a vision call regardless of
whether the page has extractable text, because a page can have SOME text
(typed field labels) while its checkboxes or hand-marks are only visible in
the image, exactly like the PO26-1002 case.

The model looks for a section that appears to require an entry — a data
field, a checkbox, a Pass/Fail/N/A selector — left completely blank: no data
AND no checkbox marked, not even that section's own N/A option. A section
with data filled in, or with its own N/A checked, is fine. Returns one flag
per issue found, folded into the same `flags` list as Step 6's cross-checks.

This does NOT block the run on failure (API error, unparseable response) —
unlike the Section 1 checkbox fidelity gate, this is a supplementary
due-diligence check, not a Section 2 input, so a failed review degrades to
"no issues found" rather than fabricating a flag or halting the part.

**Not yet confirmed as authoritative — see Open Item #6.** This is a vision-
model judgment call, not a deterministic rule like Step 7 below, and can have
false positives/negatives depending on document layout. Treat its flags as
"review this before releasing," not as a substitute for Engineering Reviewer/the Quality Manager's own
review of the compiled package.

### Step 6.6 — Vision-based key-field cross-check (added 2026-07-22, Open Item #12)

Separate again from Step 6 and Step 6.5 — this exists because Step 6's
`cross_reference()` can only compare fields that Step 4's text extraction
actually produced. On a scanned document with little or no text layer (found
on PO26-1003/CatalogSupplier: the PO's page 1 had 6 extractable characters, the
packing list had 0), text extraction returns empty for every field, so
`cross_reference()`'s mismatch check has nothing to compare and silently
reports no issue — even when a real difference exists on the page. This is
exactly how "DISC-1" slipped through unflagged on PO26-1003: the PO line
item read "DOWEL PIN," the packing list read "CAPTIVE PIN," and nothing
caught it until a human read the compiled package by eye.

Run `check_key_field_consistency()` after Step 6's `cross_reference()`. For
each of the PO, packing slip, receiving list, and manufacturer COC: if that
document's already text-extracted description, vendor, lot#, and qty are
ALL empty, render its first page and ask a vision model
(`_vision_extract_key_fields()`) to read those four fields directly off the
image. A document that already produced something from text extraction is
NOT sent to vision — this keeps the added API cost limited to genuinely
scanned/blank-text documents, not every document on every run. Then compare
description, vendor, lot#, and qty across whichever documents ended up with
a value for that field, the same shape of check as `cross_reference()` uses.
Qty comparison normalizes to digits-only first, so "525" and "525 EA" do not
falsely flag.

**This is a SOFT flag, by explicit Quality direction — not a hard
discrepancy.** Findings are prefixed `VISION CROSS-CHECK:` and always appear
in the run summary, but `process_part()` excludes this prefix from
`discrepancy_flags`, so a finding here can never by itself draft a Section 5
narrative or set Section 6.1 to "Yes" — the same treatment as the
`RECEIVING LIST:` and `CALLOUT PLACEMENT:` process-notice prefixes, even
though (unlike those two) a `VISION CROSS-CHECK:` finding can describe
something genuinely wrong with the part, not just a tooling limitation. QE
reviews and dispositions it manually, per SOP 8.3, same as DISC-1 was
handled by hand on PO26-1003.

Like Step 6.5, a failed vision call does not block the run — it degrades to
"nothing to report" for that document rather than fabricating a flag.

**Not yet confirmed as authoritative — see Open Item #12.** This is a
vision-model judgment call over a single rendered page, not a deterministic
Step 7 rule, and a description comparison in particular can flag benign
vendor-catalog naming differences alongside genuine mismatches. Treat its
flags as "QE should look at this," not as a substitute for the human review
that originally caught DISC-1.

### Step 7 — Section 2 determination — FIXED RULES ONLY

These are deterministic functions, not model judgment calls, exactly like the
Step 2 COC classification rules above. Every input is a fact already
extracted in Steps 1, 4, 5, or 6. If a required input fact is missing or
unextractable, the function returns a flag — it never guesses.

```python
BIOLOGIC_VENDORS = {"VendorB", "VendorA", "VendorC"}
BIOLOGIC_PRODUCT_LINES = {"productv", "putty", "ProductX", "productz"}
# any product description matching a biologic material counts as Biologic,
# even if it isn't one of the four names above — match on "biologic" too.

def is_biologic(product_description, product_line):
    if product_line in BIOLOGIC_PRODUCT_LINES:
        return True
    return "biologic" in product_description.lower()

def determine_2_1(receiving_list_present, packing_slip_has_ops_sig_and_date):
    if packing_slip_has_ops_sig_and_date and receiving_list_present:
        return "Completed"
    return "FLAG"  # missing signature/date or receiving list — do not guess Completed

def determine_2_2(drawing_material, drawing_finish, coc_material_present,
                   coc_finish_present, coc_is_general_only):
    if drawing_material == "N/A" and drawing_finish == "N/A":
        return "N/A"
    if coc_is_general_only:
        return "Yes"  # general COC covers both, only when it's the sole doc given
    if coc_material_present and coc_finish_present:
        return "Yes"
    return "No"  # name the specific missing cert(s) in Section 5

def determine_2_3(has_numbered_callouts, section_3_all_pass, section_3_discrepancy_present):
    if not has_numbered_callouts:
        return "N/A"
    if section_3_discrepancy_present:
        return "No"
    if section_3_all_pass:
        return "Yes"
    return "FLAG"  # Section 3 present but pass/fail not determinable from engineer's entries

def determine_2_4(vendor, is_confidential_biologic, is_traveler_vendor, is_vendord,
                   has_traveler, section1_qda=None):
    # CORRECTED 2026-09-11 (WI-7.4.3-01 gap analysis 2026-09-11, Priority 2 —
    # real data from PO26-1005/PO26-1009/PO26-1008). Both the generic
    # catch-all AND the VendorD branch previously resolved a missing traveler to
    # a bare "No", which reads as a settled negative finding rather than a
    # documentation gap needing engineer follow-up. Confirmed with Quality:
    # for a vendor that assembles nothing and supplies no document actually
    # labeled a traveler, 2.4 should default to N/A (checkbox pair left
    # unchecked) with an explicit flag — never a silent or bare "No" — so it
    # can never be mistaken for "confirmed not applicable."
    if is_confidential_biologic:
        return "Yes"  # confidential Biologic — never request, never include in package
    if is_traveler_vendor:
        return "Yes" if has_traveler else "N/A"  # FLAG naming the missing traveler
    if is_vendord:
        # VendorD is expected to supply a traveler for every lot (see FILE
        # ACCESS RULES era note / Open Item below) — missing one is a
        # documentation gap on THIS lot, not a closed "not applicable".
        return "Yes" if has_traveler else "N/A"  # FLAG naming the missing traveler
    return "N/A"  # VendorF/VendorG/other metal-machining vendors — no contract-mfr traveler
    # Separately (existing Open Item #4, unchanged by this fix): if Section
    # 1.3 QDA = YES and no traveler was found, a flag is raised regardless of
    # which branch above resolved 2.4, because a QDA'd lot is expected to
    # have a traveler even for a vendor this function would otherwise treat
    # as N/A with no comment.

def determine_2_5(is_sterilized, ster_coc_present, dose_params_passed, section1_qda=None):
    # 2026-07-20 addition (Open Item #4): a Section 1.3 QDA of YES also forces
    # this off of a silent N/A, even if the vendor/description heuristic alone
    # would not have flagged the part as sterile.
    qda_yes = str(section1_qda or "").strip().upper() == "YES"
    if not is_sterilized and not qda_yes:
        return "N/A"
    if ster_coc_present and dose_params_passed:
        return "Yes"
    return "No"  # FLAG: missing ster COC on a part expected to have one

def determine_2_6(vendor, is_biologic_part):
    if is_biologic_part and vendor in BIOLOGIC_VENDORS:
        return "Yes"  # FLAG PROMINENTLY: Director of Operations reviews and signs BEFORE Quality Manager
    return "No"
```

Any function above that returns `"FLAG"` means: write the corresponding
checkbox cell to neither Yes/No/N/A — leave that specific checkbox pair fully
unchecked (both ☐), and add a line to Section 5 and the summary explaining
what fact is missing and why the determination could not be made. Never pick
a default answer to fill a checkbox when the deterministic function cannot
resolve one.

Verification testing (DVT / first article — entered by engineer in Section 4):
Automation reads Section 4 to check for a DVT reference, never writes to it.
DVT accepted does NOT count for AQL — engineer sets Section 1.2; automation
never changes it.

Special cases (all still resolved by the functions above, listed here for
quick reference only — do not hand-apply these instead of running the code):
- VendorA/VendorB/VendorC off-the-shelf, non-Biologic: 2.4 per `determine_2_4`,
  2.6 = No. Section 4 (engineer's own text) typically reads: "These parts are
  off-the-shelf from [Vendor] therefore sections 2.X not applicable. No
  dimensional inspection required." — automation does not write this; it is
  the engineer's Section 4 entry, read only for the Step 2.5 cross-check.
- Biologic from VendorB/VendorA/VendorC: 2.6 = Yes — flag prominently; Director of Operations signs
  BEFORE Quality Manager.
- Rev mismatch: include both drawings (Step 3); Section 5 describes the
  discrepancy.
- Multiple materials/finishes on drawing: one COC required per material AND
  per finish, unless the sole document given is a general COC (see
  `determine_2_2`).
- No ACME drawing on file: Section 3 = N/A, 2.2 = N/A, 2.3 = N/A.
- Section 1.3 QDA = YES (Open Item #4, added 2026-07-20): read-only from the
  engineer's own Section 1 checkboxes, cross-checked against 2.4 and 2.5 — a
  QDA'd lot is expected to have both a sterilization COC and (usually) a
  traveler. If either is missing when QDA = YES, `determine_section2` raises
  a flag (`2.4_flag` / `2.5_flag`) rather than letting the vendor-only
  heuristic quietly resolve to N/A. This is new logic, not yet tied to an
  approved SOP revision — see Open Item #4.

### Step 8 — Fill PRF Sections 2, 5, and 6 only

Load the engineer's pre-filled PRF with `openpyxl.load_workbook()`.

**CORRECTED 2026-09-17 — do NOT convert Section 1's checkboxes to ☑/☐ text.**
That conversion is what caused Open Item #29 (every checkbox on every
exported PRF rendering as a doubled mark) — the native VML control renders
correctly on its own on the LibreOffice build actually in use, so writing a
text glyph over it drew a second, redundant mark on top of a control that
was already visible. Section 1 (and Section 3 and Section 4) checkboxes are
now read ONLY, for the Step 2.5 cross-check and Section 4's DVT-reference
check — nothing is written to their cells, not even a "just redraw what's
already there" rendering pass. The native control the engineer already set
is the only thing that renders, and it renders correctly.

**NEVER write a value to Sections 1, 3, or 4, for any reason, under any
circumstance.** Run information is written to the Desktop tools folder's
`logs\` instead (Step 9), never into the signed record itself.

**Section 4 stays exactly as it is: the engineer's own comment field.**
Automation and Claude sessions read it (Step 2.5 cross-check, DVT
reference, and as vocabulary for the COMMENT & DISCREPANCY REASONING
REFERENCE) but never write to it, pre-fill it, suggest text into it, or
"tidy" it. Reconfirmed 2026-10-05.

Write ONLY to Sections 2, 5, and 6, using the return values from Step 7's
functions.

CHECKBOX MAP — Sections 1-2 sheet:
```
2.1  → check_box(ws, 'I20', with_label='Completed')       [only if not FLAG]
2.2  → I21=Yes | L21=No | O21=N/A
2.3  → I22=Yes | L22=No | O22=N/A
2.4  → I23=Yes | L23=No | O23=N/A
2.5  → I24=Yes | L24=No | O24=N/A
2.6  → I25=Yes | L25=No   (no N/A for 2.6)
```
Label cells K21-K25, N21-N25, Q21-Q24 — NEVER touch (no value, no rendering pass).
Section 1 cells (E3,E4,E5,E6,M3,M4,M5,M6) — NEVER touch (read-only in Step 2.5).

CONDITIONAL SIGNATURE/INITIALS FIELDS — Sections 1-2 sheet (added 2026-09-08,
explicit QE instruction, PO26-1007/VendorE — see CHANGELOG):
```
H29        "If Item 2.4 is YES, Quality Manager Review and Approval
            Initials:" (label B29:G29) — write "N/A" whenever 2.4 != "Yes".
            When 2.4 == "Yes", clear any stray "N/A" and leave blank for
            the Quality Manager to actually initial.
I32 / O32  "If Item 2.6 is YES, Director of Operations Review and Approval
            Signature:" (label B32:H32) plus its Date field (label N32) —
            write "N/A" into BOTH cells whenever 2.6 != "Yes" (i.e. "No").
            When 2.6 == "Yes" (confidential Biologic only), clear any
            stray "N/A" in either cell and leave both blank for Director of Operations to sign
            and date.
```
Same rule as the Yes/No/N/A checkboxes above: this only ever writes "N/A" or
clears a stray "N/A" back to blank — it never invents a signature, initials,
or a date, and it never changes what 2.4/2.6 themselves resolved to.

"NEVER touch" below means the automation never changes WHICH answer is
recorded. It does NOT mean the cell is untouched by bytes — every native VML
checkbox on Sections 1-2, Section 3, and Sections 4-6 gets the same
rendering-only ☑/☐ text conversion (see PRE-RUN REQUIREMENTS above and
`_read_vml_checkbox_states()` in receive.py), which only redraws whatever
state the engineer already set so it survives LibreOffice's PDF export.

Section 1.2 checkboxes (D8,G8) — value NEVER touched; rendering pass applies.
Section 1.3 checkboxes (D13,N13,P13) — value NEVER touched; rendering pass applies.

Section 3 sheet — DO NOT TOUCH any answer (engineer-filled). Generalized
2026-07-20: every native checkbox discovered on this sheet gets the same
rendering-only pass as Section 1 (see above) BEFORE the automation's own W2
(Section 3 N/A indicator) write — W2 remains fully automation-owned and is
set explicitly, overriding whatever the rendering pass drew for that one cell.

Section 4 — DO NOT TOUCH any answer, full stop. Read-only, for the Step 2.5
cross-check and for confirming a DVT reference exists. Generalized 2026-07-20:
Section 4 lives on the "Sections 4-6" sheet, so every native checkbox
discovered there (including any Section 4 DVT/verification checkboxes) gets
the same rendering-only pass as Section 1 BEFORE the automation's own Section
4/5/6 writes (G3, L12, I30/M30, I36/M36, B31) — those remain fully
automation-owned and are set explicitly afterward, overriding the rendering
pass for those specific cells only. No Section 4 ANSWER is ever written by
automation.

Sections 4-6 sheet — automation writes only these cells:
```
Section 5: L12 = N/A checkbox | B14 = discrepancy text (sonnet-drafted, factual only)
Section 6.1: K30=Yes | N30=No
Section 6.2 N/A: ws['B31'].value = '6.2\n☐ N/A'   (when no discrepancies)
Disposition: J36=Raw Goods | N36=Finished Goods (or both left unchecked — see
             Open Item #13: disposition is now an APDL catalog lookup, not a
             keyword match — a hit FLAGS both boxes unchecked for engineer
             confirmation rather than auto-checking Finished Goods)
```
**Corrected 2026-09-11 — this cell map previously read `I30=Yes | M30=No` and
`I36=Raw Goods | M36=Finished Goods`. Those are the OLD, pre-2026-07-22
cells, already documented as wrong and fixed in code by Open Item #9 (found
via PO26-1003/SR90-1004's manual QE correction) more than six weeks before
this correction — `fill_prf()` has written K30/N30/J36/N36 since then. This
section's own text was simply never updated to match, so the controlled
document contradicted its own Open Items list, its own Changelog, and the
actual code the whole time. Caught while syncing the Project Map vault to
current status, not from a new code change — a reminder to re-check this
section's literal cell references any time a Step 8 write-cell fix lands,
not just the Open Items list.**

Save with `save_preserving_drawings(wb, prf_path, prf_path)`.
Export PDF with `export_and_stamp_pdf(xlsx_path, forms_dir, prf_path)`.

### Step 9 — Manifest, output, and cleanup

Write a run manifest for this part to
`Desktop\Receiving Task Automation Tools\logs\<PO#>_<part#>[_Lot<lot>]_<YYYYMMDD-HHMM>_manifest.json`
(`LOGS_DIR` in `prf_helpers.py`; flat; one file per part per run, never
overwritten; **never inside a PO folder and never anywhere in Box** —
decided 2026-10-05, Open Item #8) containing:
- Timestamp, WI-7.4.3-01 revision that ran, prf_helpers.py version
- Every extracted fact used in Steps 1/4/5/6 and which method produced it
  (pdfplumber vs. haiku vs. sonnet vs. cache hit)
- The Step 2.5 Section 1 validation result (pass, or the specific mismatch)
- Every Step 6 cross-reference check performed and its result
- The Step 7 return value for each of 2.1–2.6 and the inputs that produced it
- The Section 5 text automation drafted and the Section 6 answers it set
  (`receive.py` does not record the drafted Section 5 text yet — Open Item #31)

**What this manifest is and is not (decided 2026-10-05).** It is a
development aid for QE and for the Claude session maintaining this WI.
It records what automation drafted, so a later comparison against the
signed record shows what reviewers changed (see COMMENT & DISCREPANCY
REASONING REFERENCE). It is **not** a quality record: it is not part of the
receiving package, not shared, and not relied on for release. The signed
PRF and compiled package in the PO's Completed Receiving subfolder are the
only record of what was checked and decided, so anything a reviewer or
auditor must see belongs on the PRF itself. The filename alone (PO#, part#,
optional lot suffix, timestamp) must identify the run without opening it.

Confirm both Excel and PDF in the summary for every part that passed Step 2.5:
```
Inspection Forms/PRF SR90-1202 Rev A.xlsx  updated ✓
Inspection Forms/PRF SR90-1202 Rev A.pdf   exported ✓
logs/PO26-1085_SR90-1202_20260716-1412_manifest.json  written ✓   (Desktop tools folder)
```

Clean `_Automation\tools\_cache\_tmp\` after each run. Do not clean
`logs\` in the Desktop tools folder — manifests are kept while they are
useful for WI improvement; losing them affects no quality record.

---

## COMMENT & DISCREPANCY REASONING REFERENCE (added 2026-10-05)

**Purpose.** Make the Section 5 text and run-summary flags that automation
drafts read the way Engineering Reviewer, Quality Manager and Director of Operations actually sign them, so less gets
rewritten after compile. Decided 2026-10-05 (QE, with Engineering) in
place of a manifest-driven learning loop or any new folder in Box.

**Sources — only these two, both already signed:**
1. **Validation Test Case Matrix** (QF 7.4.3-06 Attachment A) — the
   edge-case POs recorded there with a Pass result, for the language and
   reasoning used on each scenario.
2. **Each PO's `Completed Receiving\` subfolder** — the final,
   Adobe-signed compiled package. The Section 4 comment, Section 5 text and
   Section 6 answers as signed are the reference wording. Where they differ
   from what automation drafted (manifest, or the In-Progress Receiving
   version), the signed version is the reviewer's correction.

Never a source: an unsigned In-Progress package, a draft PRF, a chat
transcript, or a verbal comment that isn't on a signed record. If it isn't
on a signed PRF or a signed test case, it doesn't go in.

**Scope — what this reference may and may not change.**
- Applies to Section 5 discrepancy text (B14) and the wording of
  run-summary flags, whether drafted by `receive.py` or by a Claude Code
  session (see Open Item #31 for the `receive.py` gap).
- **Section 4 is unchanged: engineer-owned, read-only to automation**
  (Step 8). Signed Section 4 text is collected here only as vocabulary, so
  Section 5 describes things in the terms Engineering already uses. It is
  never text that automation writes into, suggests for, or pre-fills in
  Section 4.
- Never changes a Step 7 determination (2.1–2.6, disposition). Those are
  fixed rules. If a precedent seems to contradict a Step 7 rule, that is a
  rule question for Engineering Reviewer/Quality Manager through this WI's revision process, not
  something to settle by wording.
- Never recommends an NCR or a disposition (SOP 8.3): Engineering decides.

**Section 5 drafting rules.**
1. Plain text only — no Markdown characters (`**`, `##`, list markers). It
   lands on a signed record (Open Item #15).
2. Draft only from `discrepancy_flags`. Process notices (`RECEIVING LIST:`,
   `CALLOUT PLACEMENT:`, `VISION CROSS-CHECK:`, `DRAWING SOURCE:`,
   `QTY SHORTAGE (WITHIN 10% TOLERANCE)`) never become Section 5 text
   (changelog 2026-07-22; Open Items #12, #21, #27).
3. No NCR recommendation or disposition language (SOP 8.3).
4. Where a precedent below covers the same situation, follow its signed
   wording and reasoning.
5. *Provisional until the first harvest replaces it with signed language:*
   state each discrepancy as facts with their sources, naming the document
   and the value each one shows (e.g. "PO line item describes DOWEL PIN;
   packing list describes CAPTIVE PIN."), one discrepancy per sentence,
   using part, lot and document names as printed.

**Decision precedents.** Seeded 2026-10-05 from rulings already recorded
in this file. None has yet been checked against the signed record; the
"Signed text" column is filled at the first harvest. Where a ruling and
the signed PRF disagree, the row says so. Do not quote a row as signed
language until that column is filled.

| Ref | PO / Part | Situation | Decision and reasoning | Source | Signed text / check |
|---|---|---|---|---|---|
| P-01 | PO26-1008 / CAT-XX-55, -60, -65 (VendorD) | Qty received short <1% (997/1,000, 795/800, 595/596), shortage pre-acknowledged | Section 5 = N/A, 6.1 = No. Acknowledgment belongs in Section 4, written by the engineer | Open Item #21; TC-21 | Pending harvest |
| P-02 | PO26-1005, PO26-1009 (SterilizerCo) | Sterilization-only service vendor, no document labeled a traveler | 2.4 corrected from Yes to N/A: a vendor that assembles nothing has no traveler to review | Open Item #18; TC-23 | Pending harvest |
| P-03 | PO26-1003 / SR90-1004 (CatalogSupplier) | PO says DOWEL PIN, packing list says CAPTIVE PIN (DISC-1) | Genuine description mismatch. QE dispositioned per SOP 8.3 | Open Item #12 | Pending harvest |
| P-04 | PO26-1003 / SR90-1004 | Receiving-list highlight failed on a scanned list | Tooling notice, not a discrepancy: QE cleared the Section 5 narrative and reset 6.1 | Changelog 2026-07-22, item 3 | Pending harvest |
| P-05 | PO26-1007 / PN-0001 (VendorE) | Blank (unprinted) ALARA label stock | Raw Goods. Only printed product labeling makes a label Finished Goods | Changelog 2026-09-08; Open Item #13 | Pending harvest |
| P-06 | PO26-1008 / ProductS facet screws (VendorD) | Anodized and laser-etched by VendorD; VendorD never supplies travelers | Finished Goods; 2.4 = N/A | Open Items #13, #18; changelog 2026-09-14 | **Conflict:** ruling post-dates the signed PRF, which shows Raw Goods and 2.4 = No. Record both at harvest; do not treat the signed PRF as the precedent |
| P-07 | PO26-1004 / SR90-1002 (VendorB) | UnSterilized Putty (Biologic) received into Quarantine for QDA | 2.5 = N/A (no ster COC can exist yet); 2.4 = Yes, 2.6 = Yes, Director of Operations signs before Quality Manager | Open Items #25, #30 | Pending harvest; #25 still awaits Quality confirmation |

**Adding entries — harvest procedure.**
1. **Who and when:** QE, or a Claude session at QE's direction, in
   batches (only POs completed since the last harvest). Never during a PO
   run.
2. **Access:** needs QE's explicit permission naming the POs each time
   (FILE ACCESS RULES). Read only the `Completed Receiving\` subfolders;
   never write, rename or move anything in any PO folder.
3. **Method:** text extraction (pdfplumber) of the signed PRF pages. Adobe
   digital signatures leave the text layer intact, so no model calls are
   needed to extract. Compare against the manifest in `logs\`, or the
   In-Progress Receiving version where no manifest exists.
4. **Classify every difference:**
   - *Wording only* (same meaning) → add or update a precedent row with the
     verbatim signed text.
   - *Added information* (the reviewer added something automation didn't
     state) → note what automation should have extracted; open an Open
     Item if it's a real gap.
   - *Correction of a fact or answer* → a defect: handle it through Open
     Items, the changelog and the Test Case Matrix. Never record it as a
     language precedent.
5. **Record:** one row per distinct situation, with PO, part, signer roles
   and harvest date. When a newer signed record supersedes a row, update
   the row and note the earlier one in the changelog. Every harvest gets a
   changelog entry.
6. **Control:** changes to this section are changes to WI-7.4.3-01 and
   follow its revision process like any other rule change, including
   regression against the Test Case Matrix before Rev A.

---

## OUTPUT SUMMARY FORMAT

```
====================================================
PO26-1085 — Complete | Vendor: VendorA | WI-7.4.3-01 Rev 02
====================================================

*** SECTION 1 MISMATCHES — RESOLVE BEFORE ANYTHING ELSE ***
  SR90-1001: PRF states Product P/N 'SR90-1202' / PO states 'SR90-1001'.
             Engineer must correct Section 1. No package produced this run.

Vendor: VendorA | Parts on order: 3 | PRFs ready: 2 | Section 1 mismatch: 1

Part 1: SR90-1202  Rev A  Lot 9900007
  In-Progress Receiving/PO26-1085 Receiving Part 1 - SR90-1202.pdf  (N pages)
  Inspection Forms/PRF SR90-1202 Rev A.xlsx  updated
  Inspection Forms/PRF SR90-1202 Rev A.pdf   exported
  logs/PO26-1085_SR90-1202_20260716-1412_manifest.json  written (Desktop tools folder)
  2.1 Completed | 2.2 N/A | 2.3 N/A | 2.4 N/A | 2.5 Yes | 2.6 No
  Section 3: N/A (no inspection callouts)
  Disposition: Raw Goods
  Flags: None
  PDF order: 1.PO 2.PackingSlip 3.ReceivingList 4.Drawing 5.PRF 6.MfgCOC 7.SterCOC
  Excluded: [list with reason]

Pending (no PRF in Inspection Forms/): SR90-1049 Rev A

====================================================
Next steps:
- Fix Section 1 mismatches above and re-run for those parts only
- Review suggested Sections 2/5/6 — edit freely before signing
- Resolve flags before routing for signature
- 2.4 YES → Quality Manager initials required
- 2.6 YES → Director of Operations signs FIRST, then route to Quality Manager
- Section 6 discrepancies → Engineering Reviewer signs before Quality Manager
- Engineer sends compiled PDF for signatures (not automated)
====================================================
```

---

## WHAT THE ENGINEER DOES AFTER

- Fix any Section 1 mismatches first — nothing else runs for that part until resolved
- Review automation-suggested content — edit as needed
- Confirm Section 2 checklist answers
- Review Section 5 discrepancy draft — Engineering decides on NCR (SOP 8.3)
- Sign and route:
  - 2.4 YES → Quality Manager initials
  - 2.6 YES → Director of Operations signs FIRST
  - Section 6 discrepancies → Engineering Reviewer signs before Quality Manager
  - Final: Quality Manager signature on Section 6 (Quality Manager)

---

## TOOLS AND DEPENDENCIES

Canonical master (document-control edits happen here ONLY — never at runtime):
```
...\Purchase Orders\_Automation_Master\tools\
    receive.py
    prf_helpers.py
```

Working copy used by Claude Code, refreshed from the master once per year
(see Open Items #3):
```
...\Purchase Orders\2026\_Automation\
    tools\
        receive.py         — main automation script (READ ONLY at runtime)
        prf_helpers.py      — shared utility functions (READ ONLY at runtime;
                               never rewrite inline — import from here)
        _cache\
            parts_cache.json — auto-created on first run of the year
            _tmp\            — split-page and highlight scratch files
```
Manifests are NOT kept in this tree — see Step 9:
`Desktop\Receiving Task Automation Tools\logs\` (decided 2026-10-05).

Ghostscript — install separately for fully editable compiled PDFs.

Python packages:
```
pip install anthropic pdfplumber openpyxl pymupdf
```

Model tiering:
```
claude-haiku-4-5-20251001  → COC/page classification, field extraction
claude-sonnet-4-6          → cross-reference reasoning, Section 5 narrative,
                              drawing callout parsing, complex vendor docs
```
Section 2 determinations are never made by either model — they are pure
functions (Step 7) run on facts the models or pdfplumber extracted.

Every file path is relative to the year folder Claude Code was opened in.
Reject any path that resolves outside it, and reject any path that resolves
into a sibling PO folder other than the one named in the current command.

---

