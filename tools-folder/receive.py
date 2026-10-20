#!/usr/bin/env python3
"""
receive.py — ACME Medical Receiving Inspection Automation
Document: WI-7.4.3-01 | Version: 1.0 | 2026-06-02

PRE-RUN (engineer completes first for each part):
  Fill Sections 1, 3, 4 of the PRF Excel → save to Inspection Forms/
  Naming: PRF SR##-#### Rev X.xlsx  (add Lot only if same part+rev appears twice)

USAGE:
  python receive.py PO26-1085                              # all parts with PRFs ready
  python receive.py PO26-1085 --parts SR90-1202,SR90-1001  # specific parts only
"""

SCRIPT_VERSION = "1.0"

# For the Step 9 manifest's "WI-7.4.3-01 revision that ran" field. Keep this
# in sync with CLAUDE.md's own footer line by hand — added 2026-09-17
# alongside write_manifest(); there is no single source of truth shared
# between the two files today, which is worth fixing properly at some point
# rather than trusting two hand-maintained copies to never drift.
WI_REVISION = "Revision 02 (Development)"

import argparse, json, os, re as _re, shutil, sys
from datetime import date
from pathlib import Path

import fitz
import pdfplumber
import anthropic
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from prf_helpers import (
    TOOLS_DIR,
    check_box, uncheck_box,
    save_preserving_drawings, export_and_stamp_pdf,
    compile_pdf, highlight_receiving_line,
    ChecklistCheckboxError, verify_checkbox_fidelity,
    # 2026-09-11 addition (WI-7.4.3-01 gap analysis, Priority 1): these four
    # functions have existed in prf_helpers.py since at least 2026-08-12 but
    # were never imported or called anywhere in receive.py — confirmed by
    # direct grep, not inferred. Without restore_form_controls(), openpyxl's
    # regenerated sheet XML never re-references the VML/ctrlProps parts
    # save_preserving_drawings() carries into the zip, so they ride along as
    # orphaned/unreferenced parts; the first time the delivered PRF is opened
    # and re-saved in Excel, Excel prunes every unreferenced part it finds —
    # confirmed on PO26-1001 (CAT-XX-60): both the native checkboxes AND the
    # ACME Medical letterhead logo (xl/media/image1.png) were permanently gone
    # after one ordinary Excel open+save, not merely "orphaned." See
    # fill_prf() and CLAUDE.md's corresponding changelog entry.
    restore_form_controls, set_native_checkboxes, verify_native_checkboxes,
    verify_xlsx_package_integrity,
    # 2026-09-17 addition: actually closes Open Item #11 (G2) — removes the
    # leftover native checkbox shape itself, not just stray cell text. See
    # fill_prf() call site and this function's own docstring.
    strip_orphan_checkbox,
    # 2026-09-17 addition: Step 9 traceability manifest, resolves Open Item
    # #8. See process_part()'s call site and this function's own docstring
    # for the known extraction-method-tracking gap it does NOT resolve.
    LOGS_DIR, write_manifest,
    # No-clobber guard (Priority 1b): exists in prf_helpers.py, never called.
    # Wired into process_part() below, guarding the compiled package (the
    # artifact a human actually signs/annotates) before it is regenerated.
    UserContentError, assert_no_user_edits,
    # 2026-09-11 addition (Task #18 — logo hardening): export_and_stamp_pdf()
    # can now raise LogoMissingError if neither the PRF being exported nor
    # the fixed fallback template has a usable letterhead image in xl/media/.
    # Caught around the export call in process_part() below so a missing
    # logo halts only that one part (BLOCKED), consistent with every other
    # part-level failure mode in this file, instead of crashing the whole run.
    LogoMissingError,
    # 2026-09-17 (QE, direct report: the marks "are X's instead of the
    # natural checks"). normalize_checkbox_marks() has existed in
    # prf_helpers.py since 2026-07 and WI Step 3's CAUTION names it as a
    # required step, but nothing ever called it. Wired into process_part()
    # below — see that call site.
    normalize_checkbox_marks,
)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
CACHE_FILE = TOOLS_DIR / "parts_cache.json"
TMP_DIR    = TOOLS_DIR / "_tmp"

# QF 7.2-01 Rev L (Approved Product Destination List) catalog-number lookup —
# added 2026-09-11 to resolve Open Item #13 (CLAUDE.md). Extracted once from
# the controlled QF 7.2-01 document; refresh this file whenever QF 7.2-01 is
# revised (see CLAUDE.md Open Item #13 changelog entry for the refresh note).
APDL_FILE = TOOLS_DIR / "qf_7.2-01_apdl_catalog.json"

_apdl_cache = None

def load_apdl() -> dict:
    """
    {CATALOG_NUMBER (upper, stripped): {"description", "udi", "shelf_life"}}
    from QF 7.2-01. Returns {} (never raises) if the file is missing or
    unparseable — determine_section2() treats a missing/empty APDL as "cannot
    confirm this part is a distributable catalog item," which is the SAFE
    direction for the disposition rule (see determine_section2 disposition
    comment): it degrades toward auto-Raw-Goods, never toward silently
    calling something Finished Goods that the APDL was actually there to
    catch. A missing APDL file also appends a flag so this degradation is
    never silent — see determine_section2.
    """
    global _apdl_cache
    if _apdl_cache is not None:
        return _apdl_cache
    try:
        with open(APDL_FILE, encoding="utf-8") as f:
            _apdl_cache = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"  [WARN] Could not load QF 7.2-01 APDL catalog ({APDL_FILE.name}): {e}. "
              f"Disposition will default to Raw Goods for every part this run — "
              f"see CLAUDE.md Open Item #13.")
        _apdl_cache = {}
    return _apdl_cache

# Folder names — primary and fallback where noted
F_SHIPPING  = "Shipping Docs"
F_COCS      = "COCs"
F_DRAWINGS  = "Drawings"            # primary
F_DRAWINGS2 = "Drawings & Stepfiles" # fallback
F_FORMS     = "Inspection Forms"
F_OUTPUT    = "In-Progress Receiving"

# Model tiering
MODEL = {
    "classify":  "claude-haiku-4-5-20251001",
    "extract":   "claude-haiku-4-5-20251001",
    "checklist": "claude-sonnet-4-6",
    "narrative": "claude-sonnet-4-6",
    "complex":   "claude-sonnet-4-6",
}

# Vendor rules (case-insensitive comparison)
# ONLY Biologic products from these vendors have confidential travelers (2.6=Yes, Director of Operations reviews)
BIOLOGIC_VENDORS = {"vendora", "vendorb", "vendorc"}
VENDORD_VENDORS       = {"vendord"}   # travelers required

# TRAVELER VENDOR RULE (authoritative — see WI-7.4.3-01 Step 7 §2.4):
# Genuine contract-manufacturer "Travelers" come ONLY from these vendors. A
# "traveler"-named document from any other manufacturer (VendorF, VendorG, and other
# metal / machining shops) is really that manufacturer's INSPECTION report — it
# drives 2.3, NOT 2.4, and is compiled as a manufacturer inspection report, never
# in the traveler slot. For those vendors 2.4 = N/A.
TRAVELER_VENDORS  = {"sterilizerco", "vendora", "vendorb", "vendorc"}

# Biologic product keywords — any match = confidential traveler regardless of vendor
BIOLOGIC_KEYWORDS = {
    "biologic", "biologic", "productx", "productx",
    "productz", "productw", "productv", "putty",
    "biologic", "cancellous", "demineralized",
}

def is_biologic(description: str) -> bool:
    """Return True if the product description indicates a biologic material."""
    d = description.lower()
    return any(kw in d for kw in BIOLOGIC_KEYWORDS)

_client = None

def _get_client():
    """Instantiate the Anthropic client on first use, not at import time.
    Lets compile-only tooling import this module without an API key present."""
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


# ─────────────────────────────────────────────────────────────────────────────
# CACHE
# ─────────────────────────────────────────────────────────────────────────────

def load_cache() -> dict:
    try:
        return json.loads(CACHE_FILE.read_text()) if CACHE_FILE.exists() else {}
    except Exception:
        return {}

def save_cache(cache: dict):
    CACHE_FILE.write_text(json.dumps(cache, indent=2))

def cache_key(part_number: str, revision: str) -> str:
    return f"{part_number}_Rev{revision.upper()}"


# ─────────────────────────────────────────────────────────────────────────────
# CLAUDE API
# ─────────────────────────────────────────────────────────────────────────────

def call_claude(prompt: str, task: str = "checklist") -> str:
    model = MODEL.get(task, MODEL["checklist"])
    resp = _get_client().messages.create(
        model=model, max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )
    return resp.content[0].text.strip()


# ─────────────────────────────────────────────────────────────────────────────
# TEXT EXTRACTION — pdfplumber first, AI only on escalation
# ─────────────────────────────────────────────────────────────────────────────

def extract_text(pdf_path: Path, max_chars: int = 6000) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)[:max_chars]
    except Exception as e:
        print(f"  [WARN] pdfplumber failed on {pdf_path.name}: {e}")
        return ""


def extract_page_text(pdf_path: Path, page_idx: int) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return pdf.pages[page_idx].extract_text() or ""
    except Exception:
        return ""


# ─────────────────────────────────────────────────────────────────────────────
# FOLDER RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

def find_drawings_dir(po_folder: Path) -> Path | None:
    """Drawings/ first (primary), fall back to Drawings & Stepfiles/."""
    for name in [F_DRAWINGS, F_DRAWINGS2]:
        d = po_folder / name
        if d.is_dir():
            return d
    return None

def find_receiving_list(po_folder: Path) -> Path | None:
    """
    Check PO root first, then Shipping Docs/ (requires 'receiving' in name),
    then dedicated Receiving List/ subfolder (any PDF there IS the list).
    """
    # Root — must have "receiving" in filename
    for f in sorted(po_folder.glob("*.pdf")):
        if "receiving" in f.name.lower():
            return f
    # Shipping Docs/ — must have "receiving" in filename
    sd = po_folder / F_SHIPPING
    if sd.is_dir():
        for f in sorted(sd.glob("*.pdf")):
            if "receiving" in f.name.lower():
                return f
    # Dedicated Receiving List/ subfolder — any PDF here is the receiving list
    rl = po_folder / "Receiving List"
    if rl.is_dir():
        pdfs = sorted(rl.glob("*.pdf"))
        if pdfs:
            return pdfs[0]
    return None

def find_prf(forms_dir: Path, part_number: str) -> Path | None:
    """
    Find engineer's pre-filled PRF.
    Standard naming: PRF SR##-#### Rev X.xlsx  (spaces, per convention).
    Also accepts legacy PRF-SR##-####_Rev_X.xlsx.

    PRF-as-suffix fallback (added 2026-07-21): engineers have twice now saved
    the file as "SR##-#### Rev X PRF.xlsx" — PRF at the END, not the start
    (found on PO26-1002/VendorA, then again on PO26-1003/CatalogSupplier). The
    prefix patterns above never match that, so find_prf() returned None and
    the part was silently treated as "no PRF found / pending" instead of
    being processed — even though a correctly pre-filled PRF existed right
    there in Inspection Forms/. Fallback: any .xlsx in forms_dir whose name
    contains BOTH "PRF" and the part number, case-insensitively, anywhere in
    the filename. Tried only after the stricter prefix patterns, so a
    same-named non-PRF file is very unlikely to be picked up by mistake.
    """
    for pattern in [f"PRF {part_number}*.xlsx",
                    f"PRF*{part_number}*.xlsx",
                    f"PRF-{part_number}*.xlsx"]:
        matches = sorted(forms_dir.glob(pattern))
        if matches:
            return matches[0]
    pn_upper = part_number.upper()
    suffix_matches = sorted(
        f for f in forms_dir.glob("*.xlsx")
        if pn_upper in f.stem.upper() and "PRF" in f.stem.upper()
    )
    return suffix_matches[0] if suffix_matches else None

# Filename markers for a document that legitimately carries the PO number in
# its name but is NOT a copy of ACME Medical's own Purchase Order. Added
# 2026-09-17 (found on PO26-1004/VendorB): the folder held
# "PO P26-1004 Sales Order Acknowledgement.pdf" — VendorB's own acknowledgement of
# the order, a vendor document — alongside the real PO
# ("26-1004 - VendorB.pdf"). Both name the PO number, so:
#   - find_po_documents_all() counted the acknowledgement as a second copy of
#     the governing PO, compared its line Rev "1" against the PO's own
#     Rev "A", and HALTED the entire run on a fabricated
#     "PO REVISION CONFLICT" (the two documents are not two copies of the
#     same record — a vendor's Rev 1 and ACME Medical's Rev A are different
#     numbering schemes for the same item, exactly as printed on the COC:
#     "SR90-1002 / Rev 1(Customer Rev A)").
#   - find_po_document() picked the acknowledgement over the real PO whenever
#     it was called with the PO number parsed from the document ("P26-1004",
#     which the real PO's own filename does not contain) instead of the CLI
#     argument ("PO26-1004", which it does) — see the po_doc parameter on
#     process_part() for the other half of that fix.
# "receiving"/"packing" were already excluded in find_po_documents_all() and
# in find_po_document()'s last-resort fallback only; this makes one shared
# list and applies it to every branch of both functions.
_NOT_PO_COPY_MARKERS = (
    "receiving", "packing", "pack slip", "packing slip",
    "acknowledgement", "acknowledgment", "sales order", "order confirmation",
    "quote", "quotation", "invoice",
)


def _is_not_po_copy(name: str) -> bool:
    """True if this filename marks a document that is not a copy of the PO."""
    n = name.lower()
    return any(m in n for m in _NOT_PO_COPY_MARKERS)


def find_po_document(po_folder: Path, po_number: str) -> Path | None:
    """Find PO PDF in root or Shipping Docs/."""
    for f in po_folder.glob("*.pdf"):
        if _is_not_po_copy(f.name):
            continue
        if po_number.replace("PO", "").strip("-") in f.name or "PO" in f.name.upper():
            return f
    sd = po_folder / F_SHIPPING
    if sd.is_dir():
        for f in sd.glob("*.pdf"):
            if _is_not_po_copy(f.name):
                continue
            if po_number.replace("PO", "").strip("-") in f.name:
                return f
    # Fallback: any PDF in root that isn't a receiving list, packing slip, or
    # other non-PO document that merely names the PO number.
    candidates = [f for f in po_folder.glob("*.pdf") if not _is_not_po_copy(f.name)]
    return candidates[0] if candidates else None


def find_po_documents_all(po_folder: Path, po_number: str) -> list:
    """
    Return EVERY PDF that looks like a copy of the governing PO — not just
    the first one, unlike find_po_document() above. Added 2026-09-11
    (Priority 3 item #3, WI-7.4.3-01 gap analysis 2026-09-11, PO26-1005/
    SterilizerCo): find_po_document() silently returns its first match and
    is never checked for a second, conflicting copy. PO26-1005 had two real
    PO copies on file that disagreed (Rev B/385 ea vs. Rev C/400 ea), with
    file modified/"render" timestamps that did NOT reliably establish which
    was newer — main() uses this list to detect that condition and halt
    the whole run for this PO rather than silently compiling against
    whichever copy find_po_document() happened to pick first.

    Same search surface and exclusions as find_po_document() (PO root, then
    Shipping Docs/; never a file named like a receiving list, packing slip,
    or vendor sales-order acknowledgement — see _NOT_PO_COPY_MARKERS), just
    not short-circuited on the first hit.
    """
    seen = []
    def _scan(d: Path):
        if not d.is_dir():
            return
        for f in sorted(d.glob("*.pdf")):
            if _is_not_po_copy(f.name):
                continue
            if (po_number.replace("PO", "").strip("-") in f.name) or ("PO" in f.name.upper()):
                if f not in seen:
                    seen.append(f)
    _scan(po_folder)
    _scan(po_folder / F_SHIPPING)
    return seen


# ─────────────────────────────────────────────────────────────────────────────
# PAGE SPLITTING — classify pages in combined PDFs
# ─────────────────────────────────────────────────────────────────────────────

PAGE_LABELS = ["packing_slip", "mfg_coc", "mfg_report", "ster_coc", "traveler",
               "invoice_exclude", "sgt_doc_exclude", "other"]

def classify_page(text: str, filename: str) -> str:
    """Rule-based page classification; Haiku only for ambiguous pages."""
    t = text.lower()
    fn = filename.lower()

    if "packing list" in t or "packing slip" in t or "shipping confirmation" in t:
        return "packing_slip"
    if ("commercial invoice" in t or "proforma invoice" in t or
            ("invoice" in t and "commercial" in t)):
        return "invoice_exclude"
    if "sterilization" in t and ("certificate" in t or "lot" in t):
        return "ster_coc"
    if "irradiation" in t or "gamma" in t or "e-beam" in t:
        return "ster_coc"
    if "certificate of conformance" in t or "certificate of compliance" in t:
        return "mfg_coc"
    # Manufacturer inspection / dimensional report — distinct from COC
    if (("inspection report" in t or "dimensional report" in t or "dimension report" in t
         or "first article" in t or "first piece inspection" in t
         or ("inspection" in t and "dimensions" in t and ("in spec" in t or "pass" in t or "result" in t)))
            and "certificate" not in t):
        return "mfg_report"
    if "job traveler" in t or "work traveler" in t or "router" in t:
        return "traveler"
    # ACME Medical-generated docs (exclude) — PO copies and drawings inside vendor packages
    if "acme medical" in t and ("packing" in t or "purchase order" in t or "component" in t):
        return "sgt_doc_exclude"
    # Drawing pages embedded in vendor combined docs — exclude (already added from Drawings/)
    if (("drawn by" in t or "checked by" in t or "drawing #" in t or "drawing number" in t)
            and ("revision" in t or " rev " in t or "scale" in t)):
        return "sgt_doc_exclude"
    if not t.strip():
        return "other"

    # Escalate to Haiku for ambiguous pages
    prompt = (
        "Classify this page from a medical device receiving package.\n"
        "Reply with exactly one label:\n"
        "  packing_slip  — vendor packing list or shipping document\n"
        "  mfg_coc       — manufacturer certificate of conformance\n"
        "  mfg_report    — manufacturer dimensional inspection report or first article report\n"
        "  ster_coc      — sterilization certificate\n"
        "  traveler      — job traveler or work order\n"
        "  invoice_exclude — commercial or proforma invoice (exclude)\n"
        "  sgt_doc_exclude — ACME Medical-generated document (exclude)\n"
        "  other         — anything else\n\n"
        f"Filename: {filename}\nText:\n{text[:1500]}"
    )
    result = call_claude(prompt, "classify").lower().strip().split()[0]
    return result if result in PAGE_LABELS else "other"


def split_pdf_by_page_type(pdf_path: Path, tmp_dir: Path) -> dict:
    """
    Classify each page and split into typed PDFs in tmp_dir.
    Returns dict: {label: [Path, ...]} for included types.
    Excluded types (invoice_exclude, sgt_doc_exclude) are dropped silently.

    If the PDF is scanned (all pages have empty text), the whole document is
    classified by filename via classify_coc rather than per-page splitting.
    """
    tmp_dir.mkdir(exist_ok=True)
    doc = fitz.open(str(pdf_path))
    page_count = len(doc)

    # Check if document is scanned (all pages empty text)
    all_texts = [extract_page_text(pdf_path, i) for i in range(min(page_count, 5))]
    if not any(t.strip() for t in all_texts):
        # Scanned document — classify whole file by filename first.
        # If that returns "other" (no filename keywords matched), escalate to
        # vision so a combined package like "ACME Medical COC-WO Lot#..." is not
        # silently dropped, leaving has_ster_coc and has_traveler False.
        doc.close()
        whole_label = classify_coc(pdf_path)
        if whole_label.endswith("_exclude"):
            return {}
        if whole_label == "other":
            whole_label = _classify_scanned_combined_via_vision(pdf_path)
        # whole_label may now be a pipe-joined multi-type string e.g. "mfg_coc|ster_coc|traveler"
        # Emit one copy of the file per detected type so classified dict is fully populated.
        labels = [lbl for lbl in whole_label.split("|") if lbl and not lbl.endswith("_exclude")]
        if not labels:
            return {}
        import shutil as _sh
        result = {}
        for lbl in labels:
            out_path = tmp_dir / f"{pdf_path.stem}_{lbl}.pdf"
            _sh.copy(str(pdf_path), str(out_path))
            result.setdefault(lbl, []).append(out_path)
        return result

    buckets: dict = {}

    for i in range(page_count):
        text = extract_page_text(pdf_path, i)
        label = classify_page(text, pdf_path.name)
        if label.endswith("_exclude"):
            continue
        if label not in buckets:
            buckets[label] = []
        buckets[label].append(i)

    results = {}
    for label, pages in buckets.items():
        out_path = tmp_dir / f"{pdf_path.stem}_{label}.pdf"
        out_doc = fitz.open()
        for p in pages:
            out_doc.insert_pdf(doc, from_page=p, to_page=p)
        out_doc.save(str(out_path))
        out_doc.close()
        results.setdefault(label, []).append(out_path)

    doc.close()
    return results


# ─────────────────────────────────────────────────────────────────────────────
# DOCUMENT CLASSIFICATION (whole-document)
# ─────────────────────────────────────────────────────────────────────────────


def _classify_scanned_combined_via_vision(pdf_path: Path) -> str:
    """
    Vision fallback for scanned PDFs that classify_coc() couldn't identify by filename.
    Renders up to the first 3 pages at 120 DPI and asks claude-haiku-4-5-20251001
    which document types are present in the package.

    Returns a pipe-joined string of detected labels, e.g. "mfg_coc|ster_coc|traveler".
    Falls back to "other" if the API call fails.

    Used only by split_pdf_by_page_type() after classify_coc() returns "other"
    for a scanned document. Not called for text-based PDFs.
    """
    import base64, json as _json

    doc = fitz.open(str(pdf_path))
    n_pages = len(doc)
    pages_to_check = min(3, n_pages)

    content_parts = [
        {
            "type": "text",
            "text": (
                "You are reviewing pages from a medical device supplier document package. "
                "Identify every document type present anywhere across these pages. "
                "Reply with ONLY a JSON array of strings chosen from this exact set: "
                '["mfg_coc", "mfg_report", "ster_coc", "traveler", "other"]. '
                "Definitions:\n"
                "  mfg_coc    = Manufacturer Certificate of Conformance or Compliance — "
                "shows part/lot number and is signed by the manufacturer.\n"
                "  mfg_report = Manufacturer dimensional inspection report or first article "
                "report — shows measured dimensions with pass/fail results against a spec.\n"
                "  ster_coc   = Sterilisation Certificate of Processing (e.g. from SterilizerCo) — "
                "shows irradiation/EO dose, work order, pass/fail.\n"
                "  traveler   = Job Traveler or Shop Traveler — internal manufacturing routing "
                "document with process steps, work order numbers, sign-off lines.\n"
                "  other      = Anything else (pick list, label reconciliation, etc.).\n"
                "A single PDF may contain multiple types across its pages. "
                "Include every type you can identify. Output ONLY the JSON array, no explanation."
            ),
        }
    ]

    for i in range(pages_to_check):
        try:
            page = doc[i]
            mat = fitz.Matrix(120 / 72, 120 / 72)   # 120 DPI — enough for classification
            pix = page.get_pixmap(matrix=mat)
            img_b64 = base64.standard_b64encode(pix.tobytes("jpeg")).decode("ascii")
            content_parts.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64},
            })
        except Exception:
            pass
    doc.close()

    try:
        response = _get_client().messages.create(
            model=MODEL["classify"],
            max_tokens=128,
            messages=[{"role": "user", "content": content_parts}],
        )
        raw = response.content[0].text.strip()
        types = _json.loads(raw)
        valid = {"mfg_coc", "mfg_report", "ster_coc", "traveler", "other"}
        types = [t for t in types if t in valid]
        if types:
            return "|".join(types)
    except Exception as e:
        print(f"  [WARN] Vision classification failed for {pdf_path.name}: {e}")

    return "other"


# ─────────────────────────────────────────────────────────────────────────────
# VENDOR DOCUMENT COMPLETENESS CHECK — added 2026-07-20
# ─────────────────────────────────────────────────────────────────────────────
# Root cause this fixes: automation extracted structured fields from mfg_coc
# and ster_coc via fixed JSON schemas, but mfg_report and other (e.g. a
# vendor's own internal QC/inspection form bundled into a COC package) never
# got ANY content review — they were just merged into the compiled package
# as-is. Found on PO26-1002: a VendorA "QC Final Inspection" page had its
# "2nd Pack QC Inspection Post-Sterilization" section left completely blank —
# no data entered, no Pass/Fail marked, and the section's own N/A box also
# not checked — and nothing flagged it. This is new logic derived directly
# from that finding; see CLAUDE.md Open Item #6 before treating its output as
# authoritative — it is a vision-model judgment call, not a deterministic
# rule like Step 7's Section 2 functions, and can have false positives/
# negatives depending on document layout.

def check_document_completeness(pdf_path: Path, label: str = "") -> list:
    """
    Vision-based completeness check for a single vendor document that is
    about to be included in the compiled receiving package. Renders every
    page and asks claude-haiku-4-5-20251001 to identify any section that
    appears to require an entry (a data field, a checkbox, a Pass/Fail/N/A
    selector) but has been left completely blank — no data written AND no
    checkbox marked, not even that section's own N/A option.

    Runs on every page of the document (not just the first), per WI-7.4.3-01
    Step 6.5 — this is a vision call regardless of whether the page has an
    extractable text layer, because a page can have SOME text (typed labels)
    while its checkboxes/hand-marks are only visible, not extractable (this
    is exactly the PO26-1002 case: typed field labels, hand-drawn checkmarks).

    Returns a list of human-readable flag strings (empty list if the document
    appears complete). A failure to review (API error, unparseable response)
    does NOT block the run — this is a supplementary due-diligence check, not
    a Section 2 input — it is logged as a WARN and treated as "no issues
    found" rather than guessed at, since raising a fabricated flag would be
    worse than missing one here (unlike the Section 1 checkbox fidelity gate,
    which blocks precisely because THAT input feeds a required, signed
    decision).
    """
    import base64, json as _json

    try:
        doc = fitz.open(str(pdf_path))
        n_pages = len(doc)

        content_parts = [
            {
                "type": "text",
                "text": (
                    "You are reviewing every page of a vendor/manufacturer quality "
                    "document (COC, traveler, inspection report, or QC record) that "
                    "is about to be included in a medical device receiving package. "
                    "For EACH page, check whether any section that appears to require "
                    "an entry — a data field (a blank line or box meant to hold a "
                    "value), a checkbox, or a Pass/Fail/N/A selector — has been left "
                    "completely blank: no data written in, AND no checkbox marked "
                    "(including that section's own N/A option, if it has one). A "
                    "section is FINE if it has data filled in, OR if its own N/A box "
                    "is checked. A section is a PROBLEM only if it has NEITHER a "
                    "completed entry NOR N/A marked. Ignore signature lines and dates "
                    "unless the entire record is unsigned. Ignore sections that are "
                    "clearly not part of this document's applicable scope (e.g. a "
                    "second sheet that is entirely unrelated).\n"
                    "Reply with ONLY a JSON array (no explanation, no markdown "
                    "fences). One object per issue found; an empty array [] means "
                    "the document appears complete:\n"
                    '[{"page": 1, "section": "<section name/label as printed on the '
                    'page>", "issue": "<one sentence describing what is blank>"}]'
                ),
            }
        ]
        for i in range(n_pages):
            try:
                page = doc[i]
                mat = fitz.Matrix(150 / 72, 150 / 72)  # 150 DPI — enough to read checkboxes/handwriting
                pix = page.get_pixmap(matrix=mat)
                img_b64 = base64.standard_b64encode(pix.tobytes("jpeg")).decode("ascii")
                content_parts.append({
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64},
                })
            except Exception:
                pass
        doc.close()

        response = _get_client().messages.create(
            model=MODEL["classify"],
            max_tokens=600,
            messages=[{"role": "user", "content": content_parts}],
        )
        raw = response.content[0].text.strip()
        clean = raw.lstrip("```json").lstrip("```").rstrip("```").strip()
        issues = _json.loads(clean)
        if not isinstance(issues, list):
            return []

        flags = []
        doc_ref = f" ({label})" if label else f" ({pdf_path.name})"
        for item in issues:
            if not isinstance(item, dict):
                continue
            page_no = item.get("page", "?")
            section = item.get("section", "(unnamed section)")
            issue = item.get("issue", "left blank")
            flags.append(
                f"INCOMPLETE VENDOR DOCUMENT{doc_ref}: page {page_no} — "
                f"'{section}' — {issue}. Confirm with vendor/QA before releasing."
            )
        return flags
    except Exception as e:
        print(f"  [WARN] Completeness check failed for {pdf_path.name}: {e}")
        return []


def classify_coc(pdf_path: Path) -> str:
    """
    Classify a whole COC-folder document. Rule-based first, then filename,
    then Haiku. For scanned PDFs (empty text), filename is the primary signal.
    """
    name = pdf_path.name.lower()
    text = extract_text(pdf_path, 2000).lower()
    text_empty = not text.strip()

    # Rule-based on text content (works for text-based PDFs)
    if not text_empty:
        if any(k in name or k in text for k in
               ["sterilizerco", "sterilization certificate", "irradiation", "gamma", "e-beam"]):
            return "ster_coc"
        if any(k in name or k in text for k in
               ["traveler", "job traveler", "work traveler", "router"]):
            return "traveler"
        if "certificate of conformance" in text or "certificate of compliance" in text:
            return "mfg_coc"
        if (("inspection report" in text or "dimensional report" in text
             or "first article" in text or "first piece inspection" in text
             or ("inspection" in text and "dimensions" in text
                 and ("in spec" in text or "pass" in text or "result" in text)))
                and "certificate" not in text):
            return "mfg_report"
        if "acme medical" in text and ("packing" in text or "purchase order" in text):
            return "sgt_doc_exclude"

    # Filename-based fallback (catches scanned PDFs and ambiguous text)
    if any(k in name for k in ["sterilizerco", "sterilization", "irradiation", "gamma"]):
        return "ster_coc"
    if any(k in name for k in ["traveler", "router", "work order", "job trav"]):
        return "traveler"
    if any(k in name for k in ["coc", "certificate", "cert of"]):
        return "mfg_coc"
    if any(k in name for k in ["inspection report", "insp report", "dim report",
                                "dimensional", "first article", "far "]):
        return "mfg_report"
    if any(k in name for k in ["packing", "pack slip", "shipping"]):
        return "packing_slip"

    if text_empty:
        return "other"   # scanned + no filename signal — include as other

    prompt = (
        "Classify this document from a medical device COC folder.\n"
        "Reply with exactly one: mfg_coc, mfg_report, ster_coc, traveler, sgt_doc_exclude, other\n"
        "  mfg_report = manufacturer dimensional inspection report or first article report\n"
        "  mfg_coc    = certificate of conformance or compliance (NOT an inspection report)\n\n"
        f"Filename: {pdf_path.name}\nText:\n{extract_text(pdf_path, 1500)}"
    )
    return call_claude(prompt, "classify").lower().strip().split()[0]


# ─────────────────────────────────────────────────────────────────────────────
# DATA EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def _parse_json(raw: str) -> dict:
    try:
        clean = raw.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
        return json.loads(clean)
    except Exception:
        return {}

def _parse_json_dict(raw: str) -> dict:
    """Like _parse_json but guarantees a dict. If the model returns a list,
    take the first dict element; otherwise return {}."""
    data = _parse_json(raw)
    if isinstance(data, dict):
        return data
    if isinstance(data, list):
        return next((x for x in data if isinstance(x, dict)), {})
    return {}

def extract_po_data(po_path: Path) -> dict:
    text = extract_text(po_path)
    prompt = (
        "Extract from this Purchase Order. Ignore quote pages — use only pages with "
        "'Purchase Order' header. Return JSON only:\n"
        '{"po_number":"","po_date":"","po_revision":"","vendor":"",'
        '"line_items":[{"part_number":"","description":"","revision":"","qty":0}]}\n\n'
        f"Text:\n{text}"
    )
    return _parse_json_dict(call_claude(prompt, "extract"))

_PACKING_SCHEMA = (
    '{"pack_slip_number":"","vendor":"","po_number":"","part_number":"",'
    '"description":"","qty_shipped":0,"lot_number":"","ops_signature_present":false,'
    '"ops_date":"","no_discrepancies_noted":false}'
)


def extract_packing_data(pdf_path: Path) -> dict:
    text = extract_text(pdf_path)
    # Vision fallback for a packing slip with no text layer — added
    # 2026-09-17, root-causing CLAUDE.md Open Item #14 ("Section 2.1 false
    # negative on a signed packing slip", found on PO26-1006 and reproduced
    # exactly on PO26-1004/VendorB). The cause is not signature detection per se:
    # these packing slips are flatbed SCANS with an empty text layer, so
    # pdfplumber hands this function nothing and EVERY field comes back
    # falsy — pack_slip_number, vendor, lot_number and
    # ops_signature_present alike. determine_2_1() then reads
    # "packing slip not found or unsigned" on a slip that is sitting right
    # there, physically signed and dated by Operations ("NO Discrepancies /
    # [Ops signature] / 8/14/26" on PO26-1004's). A signed record must not
    # carry a finding that the source document contradicts on sight.
    # Only the no-text-layer case takes this path, so a normal text-based
    # packing slip costs nothing extra.
    if not text.strip():
        print(f"  Packing slip: no text layer — reading {pdf_path.name} visually")
        data = _vision_extract_packing_fields(pdf_path)
        if data:
            return data
        print("  [WARN] Vision read of the packing slip returned nothing — "
              "2.1 will report INCOMPLETE. Verify the packing slip by hand.")
        return {}
    prompt = (
        "Extract from this packing slip. Return JSON only:\n"
        + _PACKING_SCHEMA + "\n\n"
        f"Text:\n{text}"
    )
    return _parse_json_dict(call_claude(prompt, "extract"))


def _vision_extract_packing_fields(pdf_path: Path) -> dict:
    """
    Read a scanned packing slip's fields off the rendered page image.
    Same plumbing and same never-raises contract as
    _vision_extract_key_fields() (see Open Item #12), but it returns the
    full packing-slip schema — including ops_signature_present/ops_date,
    which are the fields 2.1 actually turns on and which only exist as ink
    on the page for a handwritten sign-off.

    Only page 1 is read: the Operations sign-off is on the packing slip's own
    first page by convention, and any certificate bundled on a later page is
    handled separately by split_bundled_certificate_pages(). Returns {} on
    any API/parse failure so the caller degrades to "nothing read" and 2.1
    reports INCOMPLETE rather than a fabricated "Completed".
    """
    import base64, json as _json
    try:
        doc = fitz.open(str(pdf_path))
        if len(doc) == 0:
            doc.close()
            return {}
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72))
        img_b64 = base64.standard_b64encode(pix.tobytes("jpeg")).decode("ascii")
        doc.close()

        prompt_text = (
            "This is a scanned packing slip from a medical device supplier. It "
            "has no machine-extractable text, so read it visually. Reply with "
            "ONLY this JSON object (no explanation, no markdown fences), using "
            '"" (or 0 / false) for anything not present or not legible:\n'
            + _PACKING_SCHEMA + "\n"
            "pack_slip_number = the vendor's own packing slip / pack slip number.\n"
            "vendor = the company that issued and shipped this slip.\n"
            "po_number = the customer PO number printed on it.\n"
            "part_number / description / qty_shipped / lot_number = the line "
            "item as printed.\n"
            "ops_signature_present = true ONLY if a handwritten signature or "
            "initials appear on the page (a receiving sign-off). Do not count "
            "typed names, printed form labels, or the vendor's letterhead.\n"
            "ops_date = the handwritten date next to that signature, as written.\n"
            "no_discrepancies_noted = true if a handwritten note says there "
            "were no discrepancies, exceptions, shortages or damage."
        )
        response = _get_client().messages.create(
            model=MODEL["classify"],
            max_tokens=500,
            messages=[{"role": "user", "content": [
                {"type": "text", "text": prompt_text},
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                              "data": img_b64}},
            ]}],
        )
        raw = response.content[0].text.strip()
        clean = raw.lstrip("```json").lstrip("```").rstrip("```").strip()
        data = _json.loads(clean)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        print(f"  [WARN] Vision packing-slip read failed for {pdf_path.name}: {e}")
        return {}

def extract_receiving_data(pdf_path: Path, part_number: str) -> dict:
    text = extract_text(pdf_path)
    prompt = (
        f"From this receiving list, extract ONLY the row for part number {part_number}. "
        "Return a single JSON object (not an array). If the list has multiple rows, "
        f"return only the one whose part number matches {part_number}. JSON only:\n"
        '{"part_number":"","description":"","qty_to_receive":0,"qty_received":0,'
        '"receiving_date":"","receiving_location":"","lot_number":"","revision":"",'
        '"exp_date":"","mfg_date":"","ster_date":""}\n\n'
        f"Text:\n{text}"
    )
    data = _parse_json(call_claude(prompt, "extract"))
    # Defend against the model returning a list of rows despite instructions:
    # pick the row matching the requested part number, else the first row.
    if isinstance(data, list):
        pn_norm = part_number.strip().upper()
        match = next(
            (row for row in data
             if isinstance(row, dict)
             and str(row.get("part_number", "")).strip().upper() == pn_norm),
            None,
        )
        data = match or (data[0] if data and isinstance(data[0], dict) else {})
    return data if isinstance(data, dict) else {}

def extract_coc_data(pdf_path: Path, doc_type: str) -> dict:
    text = extract_text(pdf_path)
    if doc_type == "ster_coc":
        prompt = (
            "Extract from this sterilization COC. Return JSON only:\n"
            '{"sterilizer":"","work_order":"","irradiation_date":"","lot_number":"",'
            '"min_dose_spec":"","max_dose_spec":"","min_dose_result":"","max_dose_result":"",'
            '"pass_fail":"","signed_by":"","signed_date":""}\n\n'
            f"Text:\n{text}"
        )
    elif doc_type == "mfg_report":
        prompt = (
            "Extract from this manufacturer dimensional inspection report. Return JSON only.\n"
            "For out_of_spec_dimensions: list every dimension that FAILED — where the measured "
            "value is outside the tolerance band. A dimension shown in red font, flagged, starred, "
            "or with a measured value outside its upper/lower limits is out of spec. "
            "Return an empty list [] only if every single dimension passed.\n"
            '{"vendor_part_number":"","lot_number":"","revision":"",'
            '"all_dims_in_spec":true,'
            '"out_of_spec_dimensions":['
            '{"characteristic":"","nominal":"","upper_limit":"","lower_limit":"","measured":"","note":""}]}\n\n'
            f"Text:\n{text}"
        )
    else:
        prompt = (
            "Extract from this manufacturer COC. Return JSON only:\n"
            '{"vendor_part_number":"","lot_number":"","qty":0,"mfg_date":"","exp_date":"",'
            '"revision":"","deviations":"","approval_signature_present":false,'
            '"approval_date":"","material_callout":"","finish_callout":""}\n\n'
            f"Text:\n{text}"
        )
    return _parse_json_dict(call_claude(prompt, "extract"))


# ─────────────────────────────────────────────────────────────────────────────
# DRAWING PARSE WITH CACHE
# ─────────────────────────────────────────────────────────────────────────────

def parse_drawing(pdf_path: Path, part_number: str, revision: str, cache: dict) -> dict:
    """Parse drawing. Returns cached result if part+rev already known."""
    key = cache_key(part_number, revision)
    if key in cache:
        print(f"  [CACHE] {key}")
        return cache[key]

    text = extract_text(pdf_path)
    prompt = (
        "Parse this engineering drawing. Return JSON only:\n"
        '{"sgt_part_number":"","description":"","drawing_revision":"","supplier":"",'
        '"supplier_part_number":"","supplier_revision":"","material":"N/A","finish":"N/A",'
        '"sterility_status":"","general_tolerances":"","inspection_callouts":['
        '{"number":1,"description":"","nominal":"","upper_limit":"","lower_limit":"","gaging_method":""}]}\n'
        "If no numbered inspection callouts exist, return inspection_callouts as []\n"
        'sterility_status: copy the title block\'s sterility wording verbatim if it '
        'has one (e.g. "UNSTERILIZED", "STERILE", "STERILE R"); return "" if the '
        'drawing does not state one. Do not infer it.\n\n'
        f"Text:\n{text}"
    )
    data = _parse_json_dict(call_claude(prompt, "complex"))
    if not data:
        data = {"sgt_part_number": part_number, "drawing_revision": revision,
                "material": "N/A", "finish": "N/A", "inspection_callouts": []}

    cache[key] = data
    save_cache(cache)
    return data


# ─────────────────────────────────────────────────────────────────────────────
# CROSS-REFERENCE VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

import re as _re
from datetime import date, datetime

# Section 3's calibrated-tool info line ("Name, Tool #, Next Calibration
# Date") is a single freeform text field the engineer types, not a
# structured date cell — same reason mfg_date/ster_date below are compared
# as raw strings rather than real dates: nothing in this codebase has ever
# needed to parse a date out of free text before.
#
# VERIFIED 2026-09-08 against a real QF 7.4.3-02 Rev E "Section 3" sheet
# (PO26-1007/PN-0001) — resolves the CONFIRM below, which was right to doubt it.
# The sheet has two merged ranges on row 4: B4:L4 holds the static LABEL
# ("Calibrated Tool Information (Name, Tool #, Next Calibration Date):") and
# M4:V4 is the engineer's entry field. The old value, "B4", read the label,
# which is never blank and never "N/A" — so check_calibration_date() fired on
# EVERY part ever processed, emitting "calibrated tool field has text (...) but
# no date could be extracted". That flag is not in _informational_prefixes, so
# it landed in discrepancy_flags and by itself drafted a Section 5 discrepancy
# narrative and set Section 6.1 = "Yes" on a signed quality record. Found on
# PO26-1007/PN-0001, where Section 3 is legitimately N/A (no drawing, no
# dimensional inspection) and there were no discrepancies at all.
# Corrected to the value cell, M4, so a genuinely empty/N/A field is correctly
# silent — which is the documented intent of this check.
_S3_CALIBRATION_CELL = "M4"

_DATE_PATTERNS = [
    r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b",   # 3/15/2027, 03/15/27
    r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b",     # 2027-03-15
    r"\b(\d{1,2})-(\d{1,2})-(\d{4})\b",     # 03-15-2027
]

def _extract_date(text) -> date | None:
    """
    Best-effort date extraction from a free-text field or a native
    Excel date value. Returns None (never raises) if nothing date-like
    is found — callers must treat None as "not comparable," not as an error,
    since a genuinely N/A calibration field or an unusual date format are
    both expected and should never block or falsely flag a part.
    """
    if text is None:
        return None
    if isinstance(text, datetime):
        return text.date()
    if isinstance(text, date):
        return text
    s = str(text).strip()
    if not s or s.upper() == "N/A":
        return None
    for i, pattern in enumerate(_DATE_PATTERNS):
        m = _re.search(pattern, s)
        if not m:
            continue
        g = m.groups()
        try:
            if i == 1:  # Y-m-d
                return date(int(g[0]), int(g[1]), int(g[2]))
            # m/d/Y or m-d-Y — normalize a 2-digit year
            year = g[2] if len(g[2]) == 4 else ("20" + g[2] if int(g[2]) < 70 else "19" + g[2])
            return date(int(year), int(g[0]), int(g[1]))
        except (ValueError, IndexError):
            continue
    return None


def check_calibration_date(prf_path: Path) -> list:
    """
    Soft flag (never blocks) — added per QE direction: if the calibrated
    tool's Next Calibration Date (Section 3, top of sheet) is BEFORE the
    inspection date (Sections 1-2, E5), the tool was out of calibration at
    the time of inspection. Read-only on both fields; never writes anything.

    Deliberately conservative: if either date can't be parsed, or the
    calibration field is genuinely "N/A" (no calibrated tool used, per SOP
    7.4.3 6.5 / the Quality Manager's Checklist), this returns no flag at all rather
    than guessing. A separate informational flag fires only when the
    calibration field has real, non-N/A text but no date could be extracted
    from it — that case is worth a human glance, but is NOT the same finding
    as an actually-expired tool and is worded differently on purpose.
    """
    flags = []
    try:
        wb = load_workbook(prf_path, read_only=True, data_only=True)
        inspection_date_raw = wb["Sections 1-2"]["E5"].value
        calibration_raw = wb["Section 3"][_S3_CALIBRATION_CELL].value
        wb.close()
    except Exception as e:
        flags.append(
            f"CALIBRATION CHECK: could not read inspection date or calibration "
            f"field ({e}) — not compared this run."
        )
        return flags

    if calibration_raw is None or str(calibration_raw).strip().upper() in ("", "N/A"):
        return flags  # No calibrated tool used — nothing to check, per SOP 6.5

    inspection_date = _extract_date(inspection_date_raw)
    cal_date = _extract_date(calibration_raw)

    if inspection_date is None:
        flags.append(
            "CALIBRATION CHECK: could not parse the PRF inspection date "
            f"(Sections 1-2 E5 = '{inspection_date_raw}') — calibration date "
            "not compared this run."
        )
        return flags

    if cal_date is None:
        flags.append(
            f"CALIBRATION CHECK: calibrated tool field has text ('{calibration_raw}') "
            "but no date could be extracted from it — confirm the tool's "
            "calibration is current before signing."
        )
        return flags

    if cal_date < inspection_date:
        flags.append(
            f"CALIBRATION: tool's Next Calibration Date ({cal_date}) is BEFORE "
            f"the inspection date ({inspection_date}) — this tool may have been "
            f"out of calibration when this inspection was performed. Confirm "
            f"before signing."
        )
    return flags


def cross_reference(po_item: dict, packing: dict, receiving: dict,
                    mfg_coc: dict, drawing: dict, prf_section1: dict = None,
                    po_number_expected: str = None) -> list:
    """
    prf_section1: engineer's read-only PRF Section 1 fields, read the same
    way vendor is already read from E4 elsewhere in process_part() —
    {"product_name", "product_pn", "revision", "po_number", "lot_number",
     "qty_received"}. Added 2026-07-27, closing the original gap that
     motivated WI-7.4.3-01's draft "Step 2.5": nothing previously compared
     the PRF itself against the PO/receiving list/packing slip/drawing, so a
     wrong part name/number in a pre-filled PRF could reach a compiled
     package undetected. This folds the PRF in as one more source on the
     existing Part#/Rev/Lot# checks below, using the same non-empty-values
     -only comparison shape as every other field — not a new mechanism.
     NOT YET VERIFIED against a real mismatched fixture — see CLAUDE.md
     2026-07-27 changelog entry. Treated as a flag only: this can never
     block compilation, only add a MISMATCH to discrepancy_flags, same as
     any other finding here.
    po_number_expected: the actual PO# being processed (po_data.get(
     "po_number")), passed separately since po_item is per-line-item and
     does not itself carry the PO#.
    """
    flags = []
    prf_section1 = prf_section1 or {}

    def mismatch(field, sources):
        vals = {k: v for k, v in sources.items() if v and str(v).strip()}
        unique = set(str(v).strip().upper() for v in vals.values())
        if len(unique) > 1:
            flags.append(f"MISMATCH {field}: " +
                         " | ".join(f"{k}={v}" for k, v in vals.items()))

    mismatch("Part#", {
        "PO": po_item.get("part_number", ""),
        "Packing": packing.get("part_number", ""),
        "Receiving": receiving.get("part_number", ""),
        "Drawing": drawing.get("sgt_part_number", ""),
        "PRF Section 1": prf_section1.get("product_pn", ""),
    })
    mismatch("Rev", {
        "PO": po_item.get("revision", ""),
        "Receiving": receiving.get("revision", ""),
        "Drawing": drawing.get("drawing_revision", ""),
        "COC": mfg_coc.get("revision", ""),
        "PRF Section 1": prf_section1.get("revision", ""),
    })
    mismatch("Lot#", {
        "Packing": packing.get("lot_number", ""),
        "Receiving": receiving.get("lot_number", ""),
        "COC": mfg_coc.get("lot_number", ""),
        "PRF Section 1": prf_section1.get("lot_number", ""),
    })
    mismatch("Product Name", {
        "PO": po_item.get("description", ""),
        "PRF Section 1": prf_section1.get("product_name", ""),
    })
    mismatch("PO#", {
        "Order": po_number_expected or "",
        "PRF Section 1": prf_section1.get("po_number", ""),
    })

    po_qty   = po_item.get("qty", 0)
    rec_qty  = receiving.get("qty_received", 0)
    ship_qty = packing.get("qty_shipped", 0)
    prf_qty  = prf_section1.get("qty_received", 0)
    for a, b, label in [(po_qty, rec_qty, "PO vs Received"),
                        (ship_qty, rec_qty, "Shipped vs Received"),
                        (prf_qty, rec_qty, "PRF Section 1 vs Received")]:
        try:
            a_int, b_int = int(a), int(b)
            # Skip when source quantity is 0 — PO/packing slip/PRF parse failed
            if a_int == 0:
                continue
            if a_int != b_int:
                # 2026-09-11 fix (Task #12 / Priority 3 item #2, WI-7.4.3-01
                # gap analysis 2026-09-11, PO26-1008/VendorD: 997/1,000, 795/800,
                # 595/596 — all short by <1%, all pre-acknowledged by the
                # customer as acceptable; the engineer's ruling was Section 5
                # = N/A, 6.1 = No, i.e. NOT a discrepancy needing an NCR).
                # A SHORTAGE (received < expected) within 10% of the expected
                # quantity — Quality-confirmed threshold, not yet tied to an
                # approved SOP provision (see Open Item below) — is recorded
                # as a soft, informational note: it still appears in the run
                # summary so the engineer sees it, but process_part()
                # excludes it from discrepancy_flags (same treatment as the
                # RECEIVING LIST:/CALLOUT PLACEMENT:/VISION CROSS-CHECK:
                # prefixes), so it never by itself drafts a Section 5
                # narrative or sets Section 6.1 to "Yes". An OVERAGE
                # (received > expected), or a shortage beyond 10%, is still a
                # hard QTY MISMATCH exactly as before. This does NOT record
                # the customer/engineering acknowledgment itself on the PRF —
                # per the VendorD precedent, that belongs in Section 4 by the
                # engineer's own hand if it needs to be on the signed record;
                # automation still never writes to Section 4.
                if 0 < b_int < a_int and (a_int - b_int) / a_int <= 0.10:
                    pct = (a_int - b_int) / a_int
                    flags.append(
                        f"QTY SHORTAGE (WITHIN 10% TOLERANCE) {label}: {a} vs {b} "
                        f"received ({pct:.1%} short) — not treated as a discrepancy "
                        f"per the Quality-confirmed 10% threshold; confirm customer/"
                        f"engineering acknowledgment of the shortage is on file "
                        f"(Section 4, by hand, if it needs to be on the signed record)."
                    )
                else:
                    flags.append(f"QTY MISMATCH {label}: {a} != {b}")
        except (ValueError, TypeError):
            pass

    mfg  = receiving.get("mfg_date", "")
    ster = receiving.get("ster_date", "")
    if mfg and ster and mfg.strip() == ster.strip():
        flags.append("CRITICAL: Mfg date = Ster date — must be different")

    coc_vpn = mfg_coc.get("vendor_part_number", "")
    drw_spn = drawing.get("supplier_part_number", "")
    if coc_vpn and drw_spn and coc_vpn.strip().upper() != drw_spn.strip().upper():
        flags.append(f"Vendor part# mismatch: COC={coc_vpn} | Drawing supplier={drw_spn}")

    ord_rev = po_item.get("revision", "").strip().upper()
    rec_rev = receiving.get("revision", "").strip().upper()
    if ord_rev and rec_rev and ord_rev != rec_rev:
        flags.append(
            f"REV MISMATCH: Ordered Rev {ord_rev} / Received Rev {rec_rev} — "
            f"include BOTH drawings (ordered first, received second)"
        )
    return flags


# ─────────────────────────────────────────────────────────────────────────────
# VISION-BASED KEY-FIELD CROSS-CHECK — added 2026-07-22 (DISC-1 / Open Item #12)
# ─────────────────────────────────────────────────────────────────────────────
# Root cause this addresses: cross_reference() above only ever compares
# fields that pdfplumber/Haiku TEXT extraction actually produced. On a
# scanned document with little or no text layer (PO26-1003/CatalogSupplier: the PO's
# page 1 had 6 extractable characters, the packing list had 0), the matching
# extraction dict comes back empty for every field, so mismatch()'s own
# non-empty-values-only comparison has nothing to compare and silently finds
# no issue. This is exactly how DISC-1 slipped through: the PO line item read
# "DOWEL PIN" and the packing list read "CAPTIVE PIN," but text extraction
# returned nothing for either description field, so no comparison ever ran —
# a human caught it only on manual review of the compiled package.
#
# check_key_field_consistency() closes this gap for description, vendor,
# lot#, and qty specifically, by rendering page 1 of any document whose text
# extraction came back completely empty for all four fields and asking a
# vision model to read them directly off the image, then comparing across
# whatever documents have a value (text-extracted or vision-read) the same
# way cross_reference()'s mismatch() helper does. Documents that already got
# something from text extraction are NOT sent to vision — this keeps the
# added API cost limited to genuinely scanned/blank-text documents, not
# every document on every run.
#
# Per explicit direction from Quality (2026-07-22): this is a SOFT flag —
# unlike cross_reference()'s MISMATCH findings, a result here is
# informational only. It is surfaced in the run summary for human review,
# but process_part() deliberately excludes the "VISION CROSS-CHECK:" prefix
# from discrepancy_flags, so it can never by itself drive Section 5's
# narrative or flip Section 6.1 — the same treatment already given to
# "RECEIVING LIST:" and "CALLOUT PLACEMENT:" flags. A description/vendor/
# lot/qty read off a scanned image by a vision model, and a same-wording
# comparison across documents that may legitimately use different vendor
# phrasing for an identical part, is corroborating evidence for QE to weigh —
# not a deterministic Section 2 input like Step 7's fixed-rule functions.
#
# Like check_document_completeness() (Step 6.5), a failed vision call does
# NOT block the run — logged as a WARN and treated as "nothing to report"
# for that document, since a fabricated flag would be worse than a missed
# one here. Not yet confirmed against a range of real vendor documents —
# see CLAUDE.md Open Item #12.

def _vision_extract_key_fields(pdf_path: Path, label: str) -> dict:
    """
    Render page 1 only of pdf_path at 150 DPI and ask
    claude-haiku-4-5-20251001 to read part description, vendor name, lot
    number, and quantity directly off the image. Returns a dict with those
    four keys ("" if not present/legible). Never raises — an API or parse
    failure returns {} so the caller treats it as "nothing read," not a
    fabricated blank-field claim.
    """
    import base64, json as _json
    try:
        doc = fitz.open(str(pdf_path))
        if len(doc) == 0:
            doc.close()
            return {}
        page = doc[0]
        mat = fitz.Matrix(150 / 72, 150 / 72)
        pix = page.get_pixmap(matrix=mat)
        img_b64 = base64.standard_b64encode(pix.tobytes("jpeg")).decode("ascii")
        doc.close()

        prompt_text = (
            "This is a page from a medical device supplier document "
            f"({label}). It has little or no machine-extractable text, so "
            "read it visually. Reply with ONLY this JSON object (no "
            "explanation, no markdown fences), using \"\" for anything not "
            "present or not legible:\n"
            '{"description":"","vendor":"","lot_number":"","qty":""}\n'
            "description = the part/item name or description as printed.\n"
            "vendor = the company name that issued or shipped this document.\n"
            "lot_number = lot, batch, or heat number, if shown.\n"
            "qty = the quantity shown for this line item, as printed (text is fine)."
        )
        response = _get_client().messages.create(
            model=MODEL["classify"],
            max_tokens=300,
            messages=[{"role": "user", "content": [
                {"type": "text", "text": prompt_text},
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                              "data": img_b64}},
            ]}],
        )
        raw = response.content[0].text.strip()
        clean = raw.lstrip("```json").lstrip("```").rstrip("```").strip()
        data = _json.loads(clean)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        print(f"  [WARN] Vision key-field read failed for {pdf_path.name}: {e}")
        return {}


def _vision_classify_page(pdf_path: Path, page_idx: int) -> str:
    """
    Classify ONE page of a scanned document by rendering it and reading the
    image, for pages whose text layer is empty. Returns one of PAGE_LABELS,
    or "" if the call failed or the reply wasn't a recognized label — the
    caller must treat "" as "unknown, keep the page", never as an exclusion.

    Added 2026-09-08 (found on PO26-1007/VendorE, reported by QE on the first
    compiled package). classify_page() is text-only and returns "other" for a
    page with no extractable text, so a scanned vendor packing slip that has
    the manufacturer's COC bundled in as its second page had that COC page
    kept in the packing-slip slot AND compiled again from COCs/ in the COC
    slot — the same certificate twice in one package, once before the PRF and
    once after. Per WI-7.4.3-01 Step 3 the COC belongs only in position 6,
    after the PRF, and duplicate pages are an explicit exclusion.
    """
    import base64
    try:
        doc = fitz.open(str(pdf_path))
        if page_idx >= len(doc):
            doc.close()
            return ""
        pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72))
        img_b64 = base64.standard_b64encode(pix.tobytes("jpeg")).decode("ascii")
        doc.close()

        prompt_text = (
            "This is one page from a medical device receiving package. It has no "
            "machine-extractable text, so read it visually. Reply with EXACTLY one "
            "of these labels and nothing else:\n"
            "  packing_slip    — vendor packing list, shipping notification, or delivery note\n"
            "  mfg_coc         — manufacturer certificate of conformance/compliance\n"
            "  mfg_report      — manufacturer dimensional inspection or first article report\n"
            "  ster_coc        — sterilization / irradiation certificate\n"
            "  traveler        — job traveler or work order\n"
            "  invoice_exclude — commercial or proforma invoice\n"
            "  sgt_doc_exclude — a ACME Medical-generated document\n"
            "  other           — anything else"
        )
        response = _get_client().messages.create(
            model=MODEL["classify"],
            max_tokens=20,
            messages=[{"role": "user", "content": [
                {"type": "text", "text": prompt_text},
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                              "data": img_b64}},
            ]}],
        )
        label = response.content[0].text.strip().strip(".").lower()
        return label if label in PAGE_LABELS else ""
    except Exception as e:
        print(f"  [WARN] Vision page classification failed for "
              f"{pdf_path.name} p{page_idx + 1}: {e}")
        return ""


def check_key_field_consistency(doc_paths: dict, extracted: dict) -> list:
    """
    doc_paths: {label: Path|None} for the documents relevant to this part,
               e.g. {"PO": ..., "Packing": ..., "Receiving": ..., "COC": ...}
    extracted: matching {label: dict} of already text-extracted data for
               each document (po line item / packing_data / receiving_data /
               mfg_coc_data), each expected to expose whichever of
               description/vendor/lot_number/qty it normally carries.

    For each labeled document with a path: if the already-extracted
    description, vendor, lot_number, and qty are ALL empty (text extraction
    had nothing to give cross_reference() for any of them — the DISC-1
    pattern), render page 1 and ask a vision model to read them directly via
    _vision_extract_key_fields(). Otherwise use the text-extracted values
    as-is (no vision call spent on a document that already produced data).

    Then compares description, vendor, lot#, and qty across whichever
    documents ended up with a non-empty value for that field — same
    mismatch-detection shape as cross_reference() above.

    Returns a list of "VISION CROSS-CHECK: ..." flags — soft/informational
    only (see module comment above); process_part() must keep these out of
    discrepancy_flags. Empty list if nothing to report.
    """
    fields_by_doc: dict = {}
    for label, path in doc_paths.items():
        if not path:
            continue
        base = extracted.get(label, {}) or {}
        has_any = any(str(base.get(k, "")).strip() for k in
                      ("description", "vendor", "lot_number", "qty"))
        if has_any:
            fields_by_doc[label] = {
                "description": base.get("description", ""),
                "vendor": base.get("vendor", ""),
                "lot_number": base.get("lot_number", ""),
                "qty": base.get("qty", ""),
            }
        else:
            vision = _vision_extract_key_fields(path, label)
            if vision:
                fields_by_doc[label] = vision

    flags = []

    def _mismatch(field_key, display_name, normalize=None, absent_values=()):
        # absent_values (2026-09-17, PO26-1004/VendorB): values that mean "this
        # document produced nothing for this field", not "this document says
        # zero". A qty of 0 is what extract_po_data()/extract_coc_data()
        # return when the source page has no text layer to read — on
        # PO26-1004 that yielded "Qty differs across documents — PO=0 |
        # Packing=24 | Receiving=24 | COC=0", a finding about two failed
        # extractions rather than about the lot, which is exactly the kind of
        # noise that trains a reviewer to skim past this check's real hits.
        # cross_reference() already skips a 0 quantity for the same reason
        # ("Skip when source quantity is 0 — PO/packing slip/PRF parse
        # failed"); this brings the vision cross-check in line with it.
        vals = {label: str(d.get(field_key, "")).strip()
                for label, d in fields_by_doc.items()
                if str(d.get(field_key, "")).strip()
                and str(d.get(field_key, "")).strip() not in absent_values}
        if len(vals) < 2:
            return
        norm = normalize or (lambda v: v.upper())
        unique = set(norm(v) for v in vals.values())
        if len(unique) > 1:
            flags.append(
                f"VISION CROSS-CHECK: {display_name} differs across documents — " +
                " | ".join(f"{k}={v}" for k, v in vals.items()) +
                " — soft flag (vision-read/OCR-fallback); confirm against the "
                "drawing and COC before releasing. May be a benign vendor-"
                "catalog naming difference rather than a wrong part."
            )

    _mismatch("description", "Part description")
    _mismatch("vendor", "Vendor")
    _mismatch("lot_number", "Lot#")
    _mismatch("qty", "Qty", normalize=lambda v: (_re.sub(r"[^0-9]", "", v) or v.upper()),
              absent_values=("0", "0.0", "0.00"))
    return flags


# ─────────────────────────────────────────────────────────────────────────────
# QUALITY MANAGER'S CHECKLIST LOGIC — Section 2
# Per SOP 7.4.3 Rev L and the Quality Manager's Proof Receiving Checklist
# ─────────────────────────────────────────────────────────────────────────────

# Keywords that indicate a part is sterilized — drawing notes, description, or product name
STERILITY_KEYWORDS = {
    "sterile", "sterilize", "sterilized", "sterilization",
    "eo sterilized", "eo sterile", "ethylene oxide",
    "gamma", "gamma irradiated", "e-beam", "electron beam",
    "irradiated", "irradiation", "autoclave", "autoclaved",
    "single use", "single-use", "disposable",
}

# Vendors whose standard product lines are always sterile
# (ster COC missing = flag, not silent N/A)
ALWAYS_STERILE_VENDORS = {"vendora", "vendorb", "vendorc"}

# Explicit "this item is NOT sterile" language. Added 2026-09-17 (found on
# PO26-1004/VendorB). VendorB is in ALWAYS_STERILE_VENDORS, but it also supplies
# UNSTERILIZED bulk product — SR90-1002 "ProductX Putty Synthetic XF -
# 12CC UnSterilized" says so on its drawing title block, the PO line, the
# packing slip and the COC, and is received into Quarantine for QDA before
# any sterilization step exists to certify. The vendor-only heuristic
# resolved 2.5 to "No" + "MISSING: obtain ster COC before releasing" on a
# lot for which no sterilization COC can exist — a false finding on a signed
# quality record. This list is checked BEFORE the vendor heuristic; note that
# "unsterilized" CONTAINS "sterilized", so the positive STERILITY_KEYWORDS
# scan below matches it too and cannot be relied on to tell them apart.
UNSTERILE_KEYWORDS = {
    "unsterilized", "unsterilised", "unsterile",
    "un-sterilized", "un-sterilised", "un-sterile",
    "non-sterile", "nonsterile", "non sterile", "non-sterilized",
    "not sterile", "not sterilized",
}


def _part_is_sterile(vendor: str, product_description: str, drawing: dict,
                      extra_descriptions=()) -> bool:
    """
    Return True if the part is expected to be sterile based on:
      1. An explicit NOT-sterile callout anywhere in the part's own
         description/drawing text — this wins over everything below, OR
      2. Vendor is a known always-sterile supplier, OR
      3. Product description or drawing material/finish notes contain sterility keywords.
    Used to decide whether a missing ster COC is a flag (No) vs genuinely N/A.

    extra_descriptions: additional description strings extracted from THIS
    lot's own documents (packing slip, receiving list, drawing sterility
    status). Added 2026-09-17 — the PO line item's description is often the
    shortest form of the product name on file ("ProductX Putty 12cc"),
    while the vendor's own documents and the drawing title block carry the
    "UnSterilized" qualifier that actually settles this question.

    The not-sterile override deliberately does NOT fire when the same text
    also carries positive sterility language after the not-sterile phrases
    are removed — e.g. a sterilization round-trip whose unsterile input leg
    and sterile output leg are named on the same row (PO26-1005/PO26-1009).
    In that ambiguous case the part stays "sterile", so a missing ster COC
    still flags rather than silently resolving to N/A.
    """
    search_fields = [
        product_description,
        *[d for d in extra_descriptions if d],
        drawing.get("description", ""),
        drawing.get("sterility_status", ""),
        drawing.get("material", ""),
        drawing.get("finish", ""),
        drawing.get("general_tolerances", ""),
    ]
    combined = " ".join(str(f) for f in search_fields if f).lower()

    stripped = combined
    for kw in UNSTERILE_KEYWORDS:
        stripped = stripped.replace(kw, " ")
    said_unsterile = stripped != combined
    if said_unsterile and not any(kw in stripped for kw in STERILITY_KEYWORDS):
        return False

    vl = vendor.lower().strip()
    if any(v in vl for v in ALWAYS_STERILE_VENDORS):
        return True
    return any(kw in combined for kw in STERILITY_KEYWORDS)


def determine_section2(vendor: str, drawing: dict, classified_cocs: dict,
                       packing: dict, receiving: dict,
                       product_description: str = "",
                       mfg_report_data: dict = None,
                       section1_qda: str = None,
                       part_number: str = "") -> dict:
    """
    section1_qda: the engineer's Section 1.3 "QDA within window?" answer
    ("YES" / "NO" / "N/A" / None if unreadable), read read-only from the PRF's
    native checkboxes. Added 2026-07-20 per Open Item #4 (CLAUDE.md) — a QDA
    of YES means this lot went through Quality Data Analysis, which per
    engineering direction should never resolve to a missing sterilization COC
    (2.5) or missing traveler (2.4) without an explicit flag, even for a
    vendor/description combination that would otherwise let those checks
    resolve quietly. This is a new rule not yet tied to an approved SOP
    revision — see Open Item #4 before treating it as final.

    part_number: this line item's ACME Medical part/catalog number (from the PO
    line item — the most authoritative source for it, not re-derived from
    the drawing here). Added 2026-09-11 to drive the QF 7.2-01 Approved
    Product Destination List disposition lookup — see the "Disposition"
    block below and Open Item #13.
    """
    vl = vendor.lower().strip()
    is_vendord      = any(v in vl for v in VENDORD_VENDORS)

    # 2026-09-17 fix (found on PO26-1004/VendorB): the Biologic test used to read
    # the PO line item's description ALONE. On a PO whose own page has no text
    # layer, extract_po_data() returns no line items at all and main() falls
    # back to deriving the part from the PRF filename — which carries a part
    # number and revision but NO description. product_description then arrives
    # here as "", is_biologic("") is False, and a confidential Biologic lot
    # (ProductX Putty from VendorB) silently resolves 2.4 to N/A and 2.6 to
    # "No" — i.e. the Director-of-Operations review that 2.6 exists to trigger
    # would never have been flagged, and a confidential Biologic traveler
    # would have been eligible for inclusion in the compiled package. Now
    # tested against every independently-extracted description available for
    # this lot (receiving list, packing slip, drawing title block), not just
    # the PO's. Deliberately NOT read from the PRF's own Section 1: Step 2.5
    # validates the PRF against these sources, so feeding the PRF back in as a
    # source would make that check circular.
    _desc_sources = [
        product_description,
        receiving.get("description", ""),
        packing.get("description", ""),
        drawing.get("description", ""),
    ]
    _combined_description = " ".join(str(d) for d in _desc_sources if d)
    biologic   = is_biologic(_combined_description)
    is_confidential_biologic = biologic and any(v in vl for v in BIOLOGIC_VENDORS)

    has_mfg_coc    = "mfg_coc"     in classified_cocs.values()
    has_ster_coc   = "ster_coc"    in classified_cocs.values()
    has_traveler   = "traveler"    in classified_cocs.values()
    has_mfg_report = "mfg_report"  in classified_cocs.values()

    material = drawing.get("material", "N/A")
    finish   = drawing.get("finish",   "N/A")
    callouts = drawing.get("inspection_callouts", [])
    ops_sig  = packing.get("ops_signature_present", False)

    r = {}

    # 2.1 — Receiving list and packing slip signed (SOP 5.1.2 / QF 7.4.3-01)
    has_packing = packing.get("pack_slip_number","") or packing.get("vendor","") or ops_sig
    r["2.1"] = "Completed" if has_packing else "INCOMPLETE — packing slip not found or unsigned"
    if not has_packing:
        # 2026-09-17: an unresolved 2.1 now raises a flag like every other
        # Step 7 FLAG condition. Until now the "INCOMPLETE" string above went
        # only into the run summary's own 2.1 field — it never reached `flags`,
        # never reached Section 5, and (see fill_prf) the I20 "Completed" box
        # was checked anyway. WI-7.4.3-01 Step 7 is explicit that a FLAG must
        # be explained in Section 5 and the summary, and Step 8 that I20 is
        # checked "[only if not FLAG]".
        r["2.1_flag"] = (
            "MISSING: no packing slip could be found, or none with an "
            "Operations signature and date on it, for this lot — 2.1 could not "
            "be resolved and the 'Completed' box is left unchecked. Confirm the "
            "signed packing slip and receiving list by hand (SOP 7.4.3 / "
            "QF 7.4.3-01) before releasing."
        )

    # 2.2 — Mfg certs match drawing notes
    if material.upper() == "N/A" and finish.upper() == "N/A":
        r["2.2"] = "N/A"
    elif has_mfg_coc:
        r["2.2"] = "Yes"
    else:
        r["2.2"] = "No"

    # 2.3 — Manufacturer inspection report
    # No numbered callouts → N/A.
    # Callouts present but no report found → No + flag (audit finding per 820.80).
    # Report found → Yes only if ALL dims passed. If any are out of spec → No + flag.
    if not callouts:
        r["2.3"] = "N/A"
    elif not has_mfg_report:
        r["2.3"] = "No"
        r["2.3_flag"] = ("MISSING: Drawing has numbered inspection callouts but no "
                         "manufacturer inspection report was found in COCs/. "
                         "Obtain report before releasing.")
    else:
        report = mfg_report_data or {}
        all_pass = report.get("all_dims_in_spec", True)
        oos = report.get("out_of_spec_dimensions", [])
        # Treat as failed if either flag says so
        if not all_pass or oos:
            r["2.3"] = "No"
            oos_desc = "; ".join(
                f"{d.get('characteristic','')} nominal={d.get('nominal','')} "
                f"measured={d.get('measured','')} [{d.get('note','')}]"
                for d in oos if d.get("characteristic") or d.get("measured")
            ) or "see report"
            r["2.3_flag"] = (
                f"OUT OF SPEC: Manufacturer inspection report contains dimensions outside "
                f"tolerance — {oos_desc}. Engineering review required before releasing."
            )
        else:
            r["2.3"] = "Yes"

    # 2.4 — Contract manufacturer travelers
    # Genuine contract-manufacturer travelers come ONLY from SterilizerCo/VendorA/
    # VendorB/VendorC (TRAVELER_VENDORS). A "traveler"-named file from a metal / machining
    # manufacturer (VendorF, VendorG, etc.) is really that maker's inspection report — for
    # those vendors 2.4 = N/A and the file is handled as a manufacturer inspection
    # report (see process_part reclassification + WI-7.4.3-01 Step 7 §2.4).
    is_traveler_vendor = any(v in vl for v in TRAVELER_VENDORS)
    qda_yes = str(section1_qda or "").strip().upper() == "YES"
    if is_confidential_biologic:
        # Confirmed with Quality 2026-07-27 (resolves WI-7.4.3-01 Open Item #1):
        # per the Quality Manager's Proof Receiving Checklist, 2.4 is APPLICABLE for
        # confidential Biologic from VendorB/VendorA/VendorC — check YES regardless
        # of traveler availability. The traveler itself is still never
        # requested and never included in the compiled package (see
        # traveler_inc exclusion logic in process_part()).
        r["2.4"] = "Yes"   # confidential Biologic — never request, never include
    elif is_traveler_vendor:
        r["2.4"] = "Yes" if has_traveler else "N/A"
    elif is_vendord:
        # 2026-09-11 fix (WI-7.4.3-01 gap analysis, PO26-1008): this used to
        # resolve straight to "No" whenever no traveler was found, with no
        # flag — a bare "No" reads as a settled negative finding, but VendorD
        # genuinely is expected to supply a traveler ("VendorD: travelers
        # required" above), so a missing one here is a documentation gap on
        # THIS lot needing engineer follow-up, not a closed determination.
        # Aligned with the other traveler-vendor branch immediately above:
        # both now resolve to N/A (checkbox pair left unchecked, per the
        # file-wide FLAG convention) when no traveler is present, with an
        # explicit flag naming what's missing so it can never be mistaken
        # for "confirmed not applicable."
        #
        # Reconfirmed 2026-09-14 (QE, direct instruction): VendorD never
        # provides travelers at all, so this branch's "No traveler found"
        # case is not a rare exception — it is the ONLY case that will ever
        # occur for a VendorD-vendor part. This branch already resolves it to
        # N/A (never a bare "No"), which is exactly the intended behavior;
        # no code change was needed here. See CLAUDE.md Open Item #18 and
        # the 2026-09-14 changelog entry — the actual bug this instruction
        # was aimed at was in the bespoke `process_PO26-1008.py` script,
        # which hardcodes 2.4 = "No" instead of calling this function at
        # all; that script has been corrected separately.
        if has_traveler:
            r["2.4"] = "Yes"
        else:
            r["2.4"] = "N/A"
            r["2.4_flag"] = ("MISSING: VendorD is expected to supply a traveler for every "
                             "lot, but none was found in COCs/. Confirm whether a "
                             "traveler exists for this lot before releasing (see "
                             "CLAUDE.md Open Item #1).")
    else:
        # VendorF / VendorG / other metal-machining vendors — no contract-mfr traveler
        r["2.4"] = "N/A"
    # Open Item #4: QDA=YES on Section 1.3 means this lot went through Quality
    # Data Analysis — a traveler is expected to exist for that lot. Don't let
    # a non-confidential-Biologic, non-VendorD vendor resolve 2.4 to a quiet N/A
    # when QDA is YES and no traveler was actually found; flag it instead.
    if qda_yes and not is_confidential_biologic and not has_traveler and r["2.4"] in ("N/A", "No"):
        _qda_msg = ("Section 1.3 QDA is marked YES for this lot, but no traveler was "
                    "found in COCs/. Confirm whether a traveler should exist for this "
                    "lot before releasing (see CLAUDE.md Open Item #4).")
        # Append rather than overwrite: the VendorD branch above may already have
        # set its own 2.4_flag for the same missing-traveler fact (2026-09-11
        # fix) — both are real, independent reasons to double-check this
        # lot, and one should never silently replace the other.
        r["2.4_flag"] = (r["2.4_flag"] + " " + _qda_msg) if r.get("2.4_flag") else _qda_msg

    # 2.5 — Sterilization COC
    # Determine if part is expected to be sterile before deciding Yes / No / N/A.
    # Silently marking N/A for a missing ster COC on a sterile part is an audit finding.
    # extra_descriptions (2026-09-17): this lot's own vendor/receiving
    # descriptions, which carry the sterile/unsterilized qualifier far more
    # often than the PO's short line-item description does. See
    # _part_is_sterile() and the PO26-1004/VendorB entry in CLAUDE.md.
    part_is_sterile = _part_is_sterile(
        vendor, product_description, drawing,
        extra_descriptions=(packing.get("description", ""),
                            receiving.get("description", "")),
    )
    if has_ster_coc:
        r["2.5"] = "Yes"
    elif part_is_sterile or qda_yes:
        r["2.5"] = "No"
        reasons = []
        if part_is_sterile:
            reasons.append("part appears to be sterile (vendor or description)")
        if qda_yes:
            reasons.append("Section 1.3 QDA is marked YES for this lot (see Open Item #4)")
        r["2.5_flag"] = ("MISSING: " + " and ".join(reasons) + " but no sterilization COC "
                         "was found in COCs/. Obtain ster COC before releasing.")
    else:
        r["2.5"] = "N/A"

    # 2.6 — Confidential docs / Director review
    # YES only for Biologic from VendorB/VendorA/VendorC → Director of Operations reviews and signs BEFORE Quality Manager
    r["2.6"] = "Yes" if is_confidential_biologic else "No"
    r["is_biologic"] = is_confidential_biologic

    r["section3_na"] = len(callouts) == 0

    # Disposition — 2026-09-11 fix (resolves Open Item #13; supersedes the
    # 2026-07-27 "label" keyword fix). QE confirmed 2026-09-08 (PO26-1007/VendorE)
    # that the keyword proxy is unsafe in BOTH directions: it produces a
    # false NEGATIVE on a genuinely printed label whose description omits the
    # word "label," and — the live risk — a false POSITIVE on BLANK label
    # stock whose description happens to contain the word "label" (blank
    # stock is raw material ACME Medical prints later, not Finished Goods). The
    # checklist's own real criteria are reusable metal instruments with
    # laser-etched part number/lot/UDI, reusable metal trays, and anodized
    # implants — verified via the Approved Product Destination List
    # (QF 7.2-01) or a catalog number in the drawing's filename, NOT a
    # "label" text match.
    #
    # New rule, per the QF 7.2-01 cross-reference now available (see
    # load_apdl()): a part number's mere PRESENCE on the APDL does not by
    # itself settle disposition — the same catalog number can in principle
    # be received mid-process (e.g. returned from an outside process before
    # its finishing step) as Raw Goods, before it later becomes the
    # shelf-ready Finished Good the APDL entry describes. So:
    #   - NOT found on the APDL  -> auto-resolve Raw Goods. This is the safe
    #     default direction: nothing ACME Medical doesn't sell as a finished
    #     catalog item under this number could be Finished Goods.
    #   - Found on the APDL, VendorD vendor -> auto-resolve Finished Goods. See
    #     the 2026-09-14 correction note directly below — VendorD only ever
    #     ships ACME Medical's anodized, laser-etched screws already in their
    #     finished, shelf-ready state; there is no earlier-stage "raw"
    #     receipt of these catalog numbers from VendorD to be ambiguous about.
    #   - Found on the APDL, any other vendor -> do NOT auto-resolve to
    #     Finished Goods. Leave BOTH disposition checkboxes (J36/N36)
    #     unchecked and add a Section 5 / summary flag naming the APDL entry
    #     found, requiring the engineer to confirm whether THIS lot is
    #     actually the finished catalog item or an earlier-stage Raw Goods
    #     receipt of the same part number. This mirrors the file-wide FLAG
    #     convention (never guess when a deterministic function cannot
    #     resolve one) rather than picking a default in either direction.
    #
    # CORRECTION — 2026-09-14 (QE, direct instruction; supersedes the
    # 2026-09-11 entry above and its CLAUDE.md changelog counterpart). This
    # rule previously cited "PO26-1008/PO26-1001, VendorD ProductS facet screws" as
    # real-world evidence that an APDL-listed part number can legitimately
    # be received at an earlier, non-finished stage — i.e. it treated
    # PO26-1008's ProductS facet screws as a Raw Goods example. That was a
    # mistaken engineering call on the real PRF, not a confirmed multi-stage
    # scenario: VendorD anodizes and laser-etches these screws before shipping
    # them to ACME Medical, so what VendorD ships back IS the finished, laser-
    # etched catalog item — Finished Goods — not an intermediate Raw Goods
    # receipt. Corrected rule: for a VendorD-vendor part that is on the APDL
    # (i.e. any VendorD anodized screw with laser etching received — the ProductS
    # facet screw family CAT-XX-55/-60/-65/-70/-75 and any future part
    # matching that same profile), disposition auto-resolves to Finished
    # Goods, the same as this function already does for 2.4 (is_vendord is
    # already computed above). The APDL-hit-not-VendorD case (e.g. SR90-1003/
    # VendorK, PO26-1007/VendorE blank label stock, the SterilizerCo returns) is
    # UNCHANGED — none of those real confirmations actually exercised the
    # FLAG branch (all resolved to Raw Goods via "not found on the APDL"),
    # so this correction narrows the one case that did, without touching a
    # design nobody has yet found a counter-example for. This does NOT
    # touch the actual, already-signed PO26-1008 PRF/quality records
    # themselves — those are out of scope for this code fix.
    #
    # Looked up by the PO line item's own part number (authoritative — see
    # part_number param), falling back to the drawing's ACME part number if
    # the PO didn't supply one (some vendors' POs extract with no part#).
    _lookup_pn = (part_number or drawing.get("sgt_part_number", "") or "").strip().upper()
    _apdl = load_apdl()
    _apdl_hit = _apdl.get(_lookup_pn) if _lookup_pn else None
    if _apdl_hit is None:
        r["disposition"] = "Raw Goods"
    elif is_vendord:
        # VendorD anodized/laser-etched screw carve-out (2026-09-14 correction,
        # see comment block above) — VendorD only returns these already
        # finished, so an APDL hit here is not ambiguous the way it can be
        # for another vendor.
        r["disposition"] = "Finished Goods"
    else:
        r["disposition"] = "FLAG"
        r["disposition_flag"] = (
            f"DISPOSITION: Part# {_lookup_pn} appears on the QF 7.2-01 Approved "
            f"Product Destination List as \"{_apdl_hit.get('description','')}\" "
            f"(shelf life: {_apdl_hit.get('shelf_life','?')}). A catalog number's "
            f"presence on the APDL does not by itself mean THIS lot is the "
            f"finished, laser-etched/labeled catalog item — it may be an "
            f"earlier-stage Raw Goods receipt of the same part number (e.g. "
            f"returned from an outside process pre-laser-etch). Engineer must "
            f"confirm Raw Goods vs. Finished Goods by hand before releasing; "
            f"disposition checkboxes are left unchecked pending that call."
        )

    return r


# ─────────────────────────────────────────────────────────────────────────────
# DRAWING CALLOUT ANNOTATION
# ─────────────────────────────────────────────────────────────────────────────

def _find_callout_coords(page: fitz.Page, callout_number: int) -> tuple | None:
    """
    Search the drawing's text layer for a circled number matching callout_number.
    Looks for Unicode circled digits (① ② … ⑩), parenthesized forms ((1) (2) …),
    plain integers enclosed by a hand-drawn stadium/oval shape (see
    _find_stadium_oval_coords), and plain integers that appear isolated near
    dimension text.

    Returns (x, y, method) — the callout's bounding box center and which pass
    matched ("circled", "paren", or "stadium") — or None if not found. The
    caller uses the method to pick a bubble color (see annotate_drawing_
    callouts): a "stadium" match overlays an existing hand-drawn oval, which
    is visually distinct from the Unicode-circled/parenthesized convention
    and gets a red-bordered bubble instead of black so the two never look
    the same on a compiled package.
    """
    import unicodedata

    # Unicode circled digit codepoints: ① = U+2460, ② = U+2461 … ⑳ = U+2473
    circled = {}
    for cp in range(0x2460, 0x2474):   # ① through ⑳
        digit = cp - 0x2460 + 1
        circled[chr(cp)] = digit

    target = callout_number
    words = page.get_text("words")   # list of (x0, y0, x1, y1, word, ...)

    # Pass 1: circled Unicode digit exact match
    for w in words:
        x0, y0, x1, y1, word = w[0], w[1], w[2], w[3], w[4].strip()
        if word in circled and circled[word] == target:
            return ((x0 + x1) / 2, (y0 + y1) / 2, "circled")

    # Pass 2: parenthesized form  (1)  (10)  — allowing surrounding whitespace
    import re
    pat = re.compile(r'^\(?0*' + str(target) + r'\)?$')
    for w in words:
        x0, y0, x1, y1, word = w[0], w[1], w[2], w[3], w[4].strip()
        if pat.match(word):
            return ((x0 + x1) / 2, (y0 + y1) / 2, "paren")

    # Pass 3: plain digit enclosed by a hand-drawn stadium/oval shape — added
    # 2026-09-11 (WI-7.4.3-01 gap analysis, PO26-1008/VendorD ProductS facet screws).
    # This vendor's drawings mark critical dimensions with a plain (uncircled,
    # unparenthesized) digit enclosed by a drawn oval/stadium outline and a
    # "Note 4 — DENOTES CRITICAL DIMENSION(S) TO BE INSPECTED" legend, a
    # convention Pass 1/2 above cannot recognize at all (no Unicode circled
    # glyph, no parens — just an ordinary digit plus separate vector artwork).
    coords = _find_stadium_oval_coords(page, target)
    if coords is not None:
        return (coords[0], coords[1], "stadium")

    return None


def _find_stadium_oval_coords(page: fitz.Page, callout_number: int) -> tuple | None:
    """
    Find a plain (uncircled) digit whose text is enclosed by a hand-drawn
    oval/stadium vector shape — the "Note 4 — DENOTES CRITICAL DIMENSION(S)
    TO BE INSPECTED" callout convention. Returns (x, y) of the digit's
    bounding-box center, or None.

    Distinguishing a genuinely enclosed digit from an arbitrary digit that
    happens to sit near unrelated line art is the hard part here — this
    requires BOTH: (1) the digit's center falls inside a drawn path's
    bounding box, and (2) that path's box is only modestly larger than the
    digit's own box in both dimensions (an oval hand-drawn around a single
    digit, not a whole dimension line, a border, or a title-block rule).
    NOT yet verified against a real VendorD drawing in this sandbox (no
    PyMuPDF-renderable fixture available here) — see CLAUDE.md Open Item for
    this fix; treat as a first pass, and spot-check the first drawing this
    matches against by eye before trusting it unattended.
    """
    target = str(callout_number)
    words = page.get_text("words")
    candidates = [w for w in words if w[4].strip() == target]
    if not candidates:
        return None
    try:
        drawings = page.get_drawings()
    except Exception:
        return None
    if not drawings:
        return None
    for w in candidates:
        x0, y0, x1, y1 = w[0], w[1], w[2], w[3]
        dw, dh = max(x1 - x0, 0.1), max(y1 - y0, 0.1)
        wx, wy = (x0 + x1) / 2, (y0 + y1) / 2
        for d in drawings:
            rect = d.get("rect")
            if rect is None or rect.is_empty:
                continue
            rw, rh = rect.width, rect.height
            if not (rect.x0 - 1 <= wx <= rect.x1 + 1 and rect.y0 - 1 <= wy <= rect.y1 + 1):
                continue
            # Modestly larger than the digit alone in both directions —
            # excludes whole dimension lines, borders, and title-block rules,
            # which are much larger than a single-digit bounding box in at
            # least one dimension.
            if dw * 1.2 <= rw <= dw * 6 and dh * 1.2 <= rh <= dh * 6:
                return (wx, wy)
    return None


def annotate_drawing_callouts(pdf_path: Path, callouts: list, tmp_dir: Path) -> tuple:
    """
    Add numbered bubble callouts to a copy of the drawing.

    For each callout, searches the drawing's text layer for the existing circled
    number (① ② … or (1) (2) …) and overlays a clean white-filled bubble with a
    black border directly on top of the found position so it registers with the
    dimension it belongs to.

    If a callout's position cannot be found (scanned drawing or vector-only symbol),
    that bubble is skipped entirely and a flag is returned.

    Returns:
        (annotated_pdf_path, missed_callout_numbers)
        missed_callout_numbers is a list of ints; empty list = all placed successfully.
    """
    if not callouts:
        return pdf_path, []

    out = tmp_dir / f"{pdf_path.stem}_annotated.pdf"
    # Open the ORIGINAL (read-only source) and save to a NEW path. Opening a copy
    # and saving back to the same path raises "save to original must be incremental".
    doc = fitz.open(str(pdf_path))

    missed = []
    for i, _ in enumerate(callouts, start=1):
        # Search every page — circled numbers may be on sheet 2, detail views, etc.
        found_on = None
        for pg in doc:
            result = _find_callout_coords(pg, i)
            if result is not None:
                found_on = (pg, result)
                break
        if found_on is None:
            missed.append(i)
            continue

        pg, (cx, cy, method) = found_on
        r = 9

        # Bubble border+number color: RED for every inspection callout,
        # regardless of which convention matched. Changed 2026-09-17 (QE,
        # direct report — standing rule going forward, not a one-off): this
        # used to be black for the circled-Unicode/parenthesized conventions
        # and red only for a "stadium" oval match (added 2026-09-11,
        # specifically so a stadium-derived bubble wouldn't be confused with
        # a drawing's own black BOM-item/revision-change balloon). That same
        # confusion risk exists for the other two conventions too — nothing
        # stops a drawing from also carrying black BOM/revision balloons
        # alongside a circled or parenthesized critical-dimension callout —
        # so every inspection bubble this tool draws is red now, making the
        # rule uniform instead of dependent on which text convention the
        # drawing happened to use.
        border_color = (1, 0, 0)

        # White-filled circle with colored border — overlays existing symbol
        shape = pg.new_shape()
        shape.draw_circle((cx, cy), r)
        shape.finish(color=border_color, fill=(1, 1, 1), width=1.5)
        shape.commit()

        # Number text centered in bubble
        offset_x = 5 if i < 10 else 8
        pg.insert_text((cx - offset_x, cy + 4),
                       str(i), fontsize=9, color=border_color, fontname="Helv")

    doc.save(str(out), garbage=4, deflate=True)
    doc.close()
    return out, missed


# ─────────────────────────────────────────────────────────────────────────────
# PRF FILLING — Sections 2, 5, 6 ONLY (engineer owns 1, 3, 4)
# ─────────────────────────────────────────────────────────────────────────────

def _col_to_letter(col_1idx: int) -> str:
    """Convert 1-indexed column number to Excel column letter(s). e.g. 4 → D."""
    result = ""
    n = col_1idx
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _resolve_sheet_xml_path(z, sheet_name: str) -> str:
    """
    Resolve the internal xl/worksheets/sheetN.xml path for a visible sheet name
    (e.g. "Sections 1-2") by reading workbook.xml (name -> r:id) and
    workbook.xml.rels (r:id -> target), instead of assuming it is always
    sheet1.xml. Internal sheet numbering does not always match tab order —
    it can shift after repeated saves — so hardcoding "sheet1" was fragile
    and a likely contributor to silent checkbox read failures.
    Returns None if the sheet can't be located.
    """
    import re as _re
    names = set(z.namelist())
    if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
        return None
    wb_xml = z.read("xl/workbook.xml").decode("utf-8", errors="replace")
    rels_xml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="replace")

    # Attribute-order-independent: OOXML does not guarantee Id="" appears
    # before Target="" in a <Relationship> tag (2026-07-20 fix — a fixed-order
    # regex here silently failed to resolve ANY relationship for files where a
    # writer emitted Type/Target/Id order rather than Id/Type/Target, which
    # would have made Section 3/4 checkbox discovery silently find nothing —
    # exactly the "read failure treated as empty state" bug class this whole
    # checkbox-integrity system exists to eliminate).
    rid_to_target = {}
    for rel_m in _re.finditer(r'<Relationship\b([^>]*)/?>', rels_xml):
        attrs = rel_m.group(1)
        id_m = _re.search(r'\bId="([^"]+)"', attrs)
        target_m = _re.search(r'\bTarget="([^"]+)"', attrs)
        if id_m and target_m:
            rid_to_target[id_m.group(1)] = target_m.group(1)

    sheet_m = _re.search(
        r'<sheet\b[^>]*name="' + _re.escape(sheet_name) + r'"[^>]*r:id="([^"]+)"',
        wb_xml,
    )
    if not sheet_m:
        # name= and r:id= can appear in either order in the attribute list
        sheet_m = _re.search(
            r'<sheet\b[^>]*r:id="([^"]+)"[^>]*name="' + _re.escape(sheet_name) + r'"',
            wb_xml,
        )
        if sheet_m:
            rid = sheet_m.group(1)
        else:
            return None
    else:
        rid = sheet_m.group(1)

    target = rid_to_target.get(rid)
    if not target:
        return None
    if target.startswith("/xl/"):
        target = target[1:]
    elif not target.startswith("xl/"):
        target = "xl/" + target.lstrip("/")
    return target


def _parse_cell_ref(cell: str):
    """'D8' -> (4, 8) as (1-indexed col, row). Returns (None, None) if unparseable."""
    m = _re.match(r"^([A-Z]+)(\d+)$", cell or "")
    if not m:
        return None, None
    col_letters, row = m.group(1), int(m.group(2))
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col, row


def _parse_merge_ranges(sheet_xml_content: str):
    """
    Parse <mergeCell ref="D8:F12"/> entries from a worksheet's own XML.
    Returns a list of (col1, row1, col2, row2) tuples, 1-indexed inclusive.
    """
    ranges = []
    for m in _re.finditer(r'<mergeCell\s+ref="([A-Z]+\d+):([A-Z]+\d+)"', sheet_xml_content):
        col1, row1 = _parse_cell_ref(m.group(1))
        col2, row2 = _parse_cell_ref(m.group(2))
        if col1 is not None and col2 is not None:
            ranges.append((col1, row1, col2, row2))
    return ranges


def _snap_to_merge_anchor(cell: str, merge_ranges) -> str:
    """
    If `cell` falls inside a merged range, return that range's top-left cell
    instead (the only writable/addressable member of a merge). Added
    2026-07-21 (PO26-1003/CatalogSupplier, second fix same day): the QF 7.4.3-02
    Rev E master form's Section 1.2 checkboxes for "N/A" (D8) and "100%
    Inspection" (G8) are drawn two rows below their own label, at D10/G10 --
    but D8:F12 and G8:I12 are merged ranges, so D10/G10 sit INSIDE the
    same merged cell as D8/G8. The template is not defective (it works for
    every other receiving); the bug was in this file's own <x:Anchor>
    fallback (added earlier the same day), which resolved the checkbox's
    raw drawing position without checking whether that position falls
    inside a merged cell whose top-left is the actual answer address.
    Section 1.3's checkboxes happened to sit exactly on their merge's
    top-left already (D13:E15 etc.), which is why they were never affected
    and why this went unnoticed until Section 1.2 specifically was tested.
    This snap is deterministic, not a guess: merged ranges are an explicit,
    unambiguous part of the sheet's own structure (<mergeCell ref="...">),
    not inferred from proximity or heuristics — the gate's "never guess"
    rule is unaffected.
    """
    col, row = _parse_cell_ref(cell)
    if col is None:
        return cell
    for col1, row1, col2, row2 in merge_ranges:
        if col1 <= col <= col2 and row1 <= row <= row2:
            return _col_to_letter(col1) + str(row1)
    return cell


def _extract_vml_caption(shape_body: str) -> str:
    """
    Extract the human-readable caption Excel displays next to a form-control
    checkbox (its <v:textbox> content), stripped of markup. Returns "" if the
    shape has no caption text. Used so Section 3/4 checkboxes — whose exact
    cell layout and labels are NOT hardcoded anywhere in this file — can still
    be rendered with a meaningful label instead of a bare glyph.
    """
    import re as _re
    tb_m = _re.search(r"<v:textbox[^>]*>(.*?)</v:textbox>", shape_body, _re.DOTALL)
    if not tb_m:
        return ""
    text = _re.sub(r"<[^>]+>", " ", tb_m.group(1))
    text = _re.sub(r"&nbsp;|&#160;", " ", text)
    return _re.sub(r"\s+", " ", text).strip()


def _read_vml_checkbox_states(prf_path: Path, sheet_name: str, required_cells=None) -> dict:
    """
    Parse VML + ctrlProp XML files in the given sheet of the PRF to determine
    which native form-control checkboxes exist and what state they're in.
    Returns {cell_coord: {"checked": bool, "caption": str}},
    e.g. {"D8": {"checked": True, "caption": "N/A"}}.

    Generalized 2026-07-20 (previously hardcoded to the "Sections 1-2" sheet
    only). Root cause this generalization fixes: Section 1.2/1.3 were the
    ONLY checkboxes ever converted from invisible native VML controls to
    visible ☑/☐ text before PDF export — every other native checkbox in the
    PRF (Section 3's per-dimension pass/fail boxes, Section 4's DVT/
    verification boxes) was left as a LibreOffice-invisible VML control and
    rendered as a blank box in the exported PDF even when the engineer had
    checked it. This function makes no assumption about which cells hold
    checkboxes on a given sheet — it discovers every native checkbox shape
    actually present and reads its real state (and caption), so Step 8 can
    render ALL of them the same way Section 1 already was, on any sheet.

    required_cells: optional iterable of cell coords that MUST resolve, or
    this raises ChecklistCheckboxError (used only for Section 1's five
    decision-critical checkboxes, where an unresolved state must never be
    silently defaulted). When omitted (Section 3 / Section 4 usage), this
    returns whatever checkboxes are actually found — possibly {} — because
    the exact set varies by template revision and must never be hardcoded.

    Raises ChecklistCheckboxError on any STRUCTURAL read failure (sheet not
    found, or a shape's checked-state genuinely can't be parsed). Never
    silently returns a partial/empty result when required_cells is given —
    per the 2026-07-14 fix, a read failure must stop the run, not guess.
    When required_cells is None and the sheet simply has no VML/ctrlProp
    parts at all, that is treated as "zero checkboxes on this sheet" (return
    {}), not a structural failure.

    Handles three real-world ctrlProp checked-state formats:
      Format A  checked="1"         attribute on the root element
      Format B  val="1"             attribute on a <mc:checked> / <xdr:checked> child
      Format C  <x:Checked>1</x:Checked>  text element inside the VML shape itself

    Cell address is resolved, in order:
      1. <x:Row> / <x:Column>  (primary — explicit VML coordinates)
      2. <x:FmlaLink>          (fallback — linked-cell formula reference)
      3. <x:Anchor>            (fallback, added 2026-07-21 — see below)
    A shape lacking all three is skipped, same as before.

    2026-07-21 addition (PO26-1003/CatalogSupplier PRF): a template was found whose
    checkboxes have NEITHER Row/Column NOR FmlaLink — they are positioned
    purely via drawing anchor. <x:Anchor> is "fromCol, fromColOff, fromRow,
    fromRowOff, toCol, toColOff, toRow, toRowOff"; the checkbox's own cell is
    the from-column/from-row (0-indexed) — the cell the control visually sits
    on top of. Without this fallback, every checkbox on such a template
    resolves to nothing, `states` stays empty, and the Section 1 required-
    cells gate correctly (but unhelpfully) raises ChecklistCheckboxError for
    every one of D8/G8/D13/N13/P13, blocking the part even though the
    engineer's answers are recoverable. This fallback is best-effort, not a
    guarantee: if a template's checkboxes resolve via none of the three
    methods, the gate is CORRECT to keep blocking rather than guess — do not
    add a fourth fallback that guesses a cell from shape position/order alone.

    2026-07-21 addition: now reads EVERY VML drawing file associated with the
    sheet (previously only the first). If a sheet's checkboxes happen to be
    split across more than one VML part, the rest were previously silently
    never discovered — same "read failure treated as empty state" bug class
    as the earlier Id/Target attribute-order fix. Purely additive: sheets
    with a single VML file (the common case) behave exactly as before.

    2026-07-21 correction (same PRF, found within hours of the two additions
    above): the <x:Anchor> fallback initially still failed to resolve D8/G8
    on this same PRF (and on a genuinely blank copy of the QF 7.4.3-02 Rev E
    master form itself — confirming the QF form is NOT defective, contrary
    to an earlier hypothesis raised and then disproven during this same
    investigation). Root cause: D8 and G8 are merged cells (D8:F12, G8:I12)
    holding the "N/A" / "100% Inspection" labels, but their checkboxes are
    drawn two rows below the label, at D10/G10 — which sit INSIDE those same
    merged ranges, not on the merge's own top-left address. Section 1.3's
    checkboxes never hit this because their merges happen to start at the
    checkbox's own anchor row, so raw anchor == merge top-left already.
    Fix: after resolving a raw cell via any of the three methods above,
    `_snap_to_merge_anchor()` checks whether it falls inside a merged range
    (parsed from the sheet's own <mergeCell ref="..."> entries) and, if so,
    substitutes that range's top-left cell. This is deterministic, not a
    heuristic — merged ranges are explicit, unambiguous sheet structure, so
    this does not weaken the "never guess" rule the required-cells gate
    exists to enforce.
    """
    import zipfile as _zf, re as _re
    states = {}
    try:
        with _zf.ZipFile(str(prf_path), "r") as z:
            names = set(z.namelist())

            sheet_xml = _resolve_sheet_xml_path(z, sheet_name)
            if not sheet_xml:
                raise ChecklistCheckboxError(
                    f"Could not locate the '{sheet_name}' sheet inside {prf_path.name} "
                    f"via workbook.xml — internal workbook structure may have changed."
                )
            sheet_dir, sheet_file = sheet_xml.rsplit("/", 1)

            # Merged ranges on this sheet, for merge-aware cell resolution
            # (2026-07-21 fix — see _snap_to_merge_anchor docstring).
            merge_ranges = []
            if sheet_xml in names:
                sheet_content = z.read(sheet_xml).decode("utf-8", errors="replace")
                merge_ranges = _parse_merge_ranges(sheet_content)

            rels_path = f"{sheet_dir}/_rels/{sheet_file}.rels"
            if rels_path not in names:
                if required_cells:
                    raise ChecklistCheckboxError(
                        f"No _rels file found for the {sheet_name} sheet ({rels_path}) in "
                        f"{prf_path.name} — the native checkboxes may have been stripped."
                    )
                return {}  # no relationships at all for this sheet — no checkboxes to find
            rels_xml = z.read(rels_path).decode("utf-8", errors="replace")

            # Map r:id → normalized target path.
            # Attribute-order-independent (see matching fix + comment in
            # _resolve_sheet_xml_path above — same bug class, same fix here).
            rid_to_target = {}
            for rel_m in _re.finditer(r'<Relationship\b([^>]*)/?>', rels_xml):
                attrs = rel_m.group(1)
                id_m = _re.search(r'\bId="([^"]+)"', attrs)
                target_m = _re.search(r'\bTarget="([^"]+)"', attrs)
                if not (id_m and target_m):
                    continue
                rid, target = id_m.group(1), target_m.group(1)
                if target.startswith("../"):
                    target = "xl/" + target[3:]
                elif not target.startswith("xl/"):
                    target = "xl/worksheets/" + target
                rid_to_target[rid] = target

            # Read ctrlProp checked states: {path: is_checked}
            # Excel stores checked state in three different ways depending on version:
            #   Format A: checked="1"  as an attribute on the root ctrlPr element
            #   Format B: val="1"      as an attribute on a <*:checked> child element
            # Both are tried; either being truthy means the box is checked.
            ctrl_checked = {}
            for name in names:
                if "ctrlProp" in name and name.endswith(".xml"):
                    xml = z.read(name).decode("utf-8", errors="replace")
                    # Format A: checked="1" anywhere in the file
                    fmt_a = bool(_re.search(r'\bchecked="1"', xml))
                    # Format B: <*:checked ... val="1" ...> or <*:checked>1</*:checked>
                    fmt_b = bool(_re.search(r'<[^>:]*:?checked\b[^>]*\bval="1"', xml) or
                                 _re.search(r'<[^>:]*:?checked\b[^/]*/?>1</', xml))
                    ctrl_checked[name] = fmt_a or fmt_b

            # Build rid → is_checked for ctrlProp relationships
            rid_to_checked = {
                rid: ctrl_checked.get(target, False)
                for rid, target in rid_to_target.items()
                if "ctrlProp" in target
            }

            # Locate ALL VML drawing files for this sheet (2026-07-21: previously
            # only the first was read — see docstring note on why that's fixed).
            vml_paths = [v for v in rid_to_target.values() if v.endswith(".vml") and v in names]
            if not vml_paths:
                if required_cells:
                    raise ChecklistCheckboxError(
                        f"No VML drawing file found for the {sheet_name} sheet in "
                        f"{prf_path.name} — native checkboxes may have been stripped "
                        f"by a prior save."
                    )
                return {}  # sheet has no VML at all — zero checkboxes, not a failure

            for vml_path in vml_paths:
                vml = z.read(vml_path).decode("utf-8", errors="replace")

                # Parse each VML shape that is a Checkbox
                for shape_m in _re.finditer(r"<v:shape\b[^>]*>(.*?)</v:shape>", vml, _re.DOTALL):
                    body = shape_m.group(1)
                    if not _re.search(r"[Cc]heck[Bb]ox", body):
                        continue

                    # ── Resolve cell address ──────────────────────────────────
                    # Primary: <x:Row> / <x:Column> (0-indexed VML coordinates)
                    row_m = _re.search(r"<x:Row>(\d+)</x:Row>", body)
                    col_m = _re.search(r"<x:Column>(\d+)</x:Column>", body)
                    cell = None
                    if row_m and col_m:
                        row = int(row_m.group(1)) + 1   # VML is 0-indexed
                        col = int(col_m.group(1)) + 1
                        cell = _col_to_letter(col) + str(row)
                    else:
                        # Fallback 1: <x:FmlaLink> stores the linked cell address
                        link_m = _re.search(r"<x:FmlaLink>[^!<]*!?\$?([A-Z]+)\$?(\d+)</x:FmlaLink>",
                                            body, _re.IGNORECASE)
                        if link_m:
                            cell = link_m.group(1).upper() + link_m.group(2)
                        else:
                            # Fallback 2 (added 2026-07-21, see docstring): <x:Anchor>
                            # "fromCol, fromColOff, fromRow, fromRowOff, toCol,
                            # toColOff, toRow, toRowOff" — cell is from-col/from-row.
                            anchor_m = _re.search(r"<x:Anchor>\s*([\d,\s]+)</x:Anchor>", body)
                            if anchor_m:
                                parts = [p.strip() for p in anchor_m.group(1).split(",")]
                                if len(parts) >= 3:
                                    try:
                                        from_col = int(parts[0])
                                        from_row = int(parts[2])
                                        cell = _col_to_letter(from_col + 1) + str(from_row + 1)
                                    except ValueError:
                                        cell = None
                    if cell is None:
                        continue

                    # Merge-aware snap (2026-07-21 fix): if the resolved cell
                    # sits inside a merged range, the merge's top-left is the
                    # actual answer address, not the raw drawing position.
                    if merge_ranges:
                        cell = _snap_to_merge_anchor(cell, merge_ranges)

                    # ── Resolve checked state ─────────────────────────────────
                    # ctrlProp via RelID is authoritative (Formats A + B handled above).
                    # Fall back to Format C: <x:Checked>1</x:Checked> inside the VML shape.
                    rel_m = _re.search(r'<x:RelID\s[^>]*r:id="([^"]+)"', body)
                    if rel_m:
                        is_checked = rid_to_checked.get(rel_m.group(1), False)
                    else:
                        chk = _re.search(r"<x:Checked>(\d+)</x:Checked>", body)
                        is_checked = bool(chk and chk.group(1) == "1")

                    states[cell] = {"checked": is_checked, "caption": _extract_vml_caption(body)}

    except ChecklistCheckboxError:
        raise
    except Exception as e:
        # Any other unexpected parse failure is just as unsafe to silently swallow
        # as the cases we raise explicitly above — surface it the same way instead
        # of returning a partial/empty dict for the caller to default from.
        raise ChecklistCheckboxError(
            f"Unexpected error reading checkbox state from the {sheet_name} sheet of "
            f"{prf_path.name}: {e!r}"
        ) from e

    if required_cells:
        missing = [c for c in required_cells if c not in states]
        if missing:
            raise ChecklistCheckboxError(
                f"{sheet_name} checkbox state could not be resolved for {', '.join(missing)} "
                f"in {prf_path.name} — the VML shape(s) for these cells were not found or "
                f"did not match the expected format."
            )

    return states


# Section 1 checkbox cells and their display labels (per QF 7.4.3-02 Rev E layout)
#
# 2026-07-2x fix: J8 (1.2's third option — AQL sampling, as opposed to D8
# N/A or G8 100% Inspection) was never in this map, so unlike D8/G8 it was
# never explicitly read from its real VML state and re-rendered on save —
# it relied entirely on save_preserving_drawings() carrying the original
# VML/ctrlProp shape through untouched. QE reported this checkbox
# disappearing on repeated runs. Adding it here makes it behave exactly
# like D8/G8: read once, re-drawn every run from its actual state, same as
# every other Section 1 checkbox. NOT independently confirmed as the full
# root cause — if the real failure is something else (e.g. a merge-range
# interaction specific to J8, or something inside save_preserving_drawings
# itself), this will not catch it. Needs the same kind of direct check
# against a blank template's real XML that root-caused G2/B31/K30 before
# treating this as closed — see WI-7.4.3-01 Open Items.
_S1_CHECKBOXES = {
    "D8":  "N/A",              # 1.2 N/A (quantity inspected)
    "G8":  "100% Inspection",  # 1.2 100% Inspection
    "J8":  "AQL",              # 1.2 AQL (quantity inspected) — added 2026-07-2x
    "D13": "N/A",              # 1.3 QDA N/A
    "N13": "YES",              # 1.3 QDA YES
    "P13": "NO",               # 1.3 QDA NO
}


def _anchor_cell(ws, coord):
    """
    Return the writable cell for coord. If coord falls inside a merged range,
    return the range's top-left anchor cell (the only writable member).
    Otherwise return the cell itself.
    """
    from openpyxl.utils import range_boundaries, get_column_letter
    target = ws[coord]
    # Fast path: not a merged cell
    from openpyxl.cell.cell import MergedCell
    if not isinstance(target, MergedCell):
        return target
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            min_col, min_row, _, _ = range_boundaries(str(rng))
            return ws.cell(row=min_row, column=min_col)
    return target  # shouldn't happen, but fail safe


def fill_prf(prf_path: Path, section2: dict, discrepancy_flags: list, drawing_data: dict = None,
            s1_states: dict = None) -> Path:
    """
    Read engineer's pre-filled PRF. Fill Sections 2, 5, 6.

    discrepancy_flags: 2026-07-22 rename/behavior change (found on
    PO26-1003/SR90-1004). This must be the QUALITY-DISCREPANCY subset of a
    part's flags, not the full flags list process_part() reports in its
    summary — process-limitation notes (e.g. "RECEIVING LIST: search term
    not found", "CALLOUT PLACEMENT: could not locate callout") do not
    describe anything wrong with the received part, and a run whose only
    "flags" were of that kind was incorrectly getting a Section 5 narrative
    and a Section 6.1 "discrepancies present = Yes" — both had to be
    manually cleared to N/A/No by QE review. See process_part() for how
    this subset is built before being passed in here.

    STALE DOCSTRING CORRECTED 2026-09-17 — this used to say every native VML
    checkbox got a ☑/☐ text glyph written into its cell "so it renders
    correctly in LibreOffice" (LibreOffice was believed unable to render
    native VML controls at all). That premise is false on the LibreOffice
    build actually in use, confirmed by rendering a real export at 400 DPI:
    the native control renders on its own. The glyph write was never a
    harmless fallback — it was a second mark drawn over the template's own
    checkbox, on top of an already-visible control, which is exactly what
    made every checkbox on every exported PRF look doubled (a bold glyph
    sitting on/beside a thin box outline). See Open Item #29 and this
    changelog's 2026-09-17 entry for the full history, including the earlier
    (wrong) conclusion in this codebase that only G2 carried this risk.

    Per QE's direct instruction, that glyph write is REMOVED — the
    template's native form controls are the checkboxes of record for both
    engineer-owned and automation-owned answers:
      * Section 1/3/4 (engineer-owned): read only, never written. The
        engineer's own native control already renders their answer; nothing
        here re-draws it.
      * Section 2/5/6 (automation-owned): set via set_native_checkboxes()
        (existing since 2026-09-11) further below. The label text next to
        each checkbox (e.g. "Yes", "N/A", "Raw Goods") is preserved or
        written as plain text; only the leading glyph character is gone.
      * G3 is the one known exception: this template ships it with no native
        form control at all (confirmed via _control_map()), so a plain ☑/☐
        glyph is still written there — there is nothing else to carry its
        state. If a future template variant does have a native control at
        G3, this should be revisited so G3 stops being a special case.
    A stray glyph left by an OLDER run of this tool in an automation-owned
    cell (Section 2/5/6) is cleared on the next run over that same file. A
    glyph an older run left in a Section 1/3/4 cell is NOT touched, even to
    clean it up — those cells are permanently off-limits to automation,
    cosmetic leftover or not. See the 2026-09-17 changelog entry for what
    this means for PRFs already processed before this fix: their Section 1
    checkboxes will keep showing a doubled mark until someone removes the
    stray character from that cell by hand in Excel (safe — it only ever
    touches decorative cell text, never the real checkbox, which already
    carries the correct answer).

    s1_states: pre-read Section 1 VML states (from process_part's Section
    1 → Section 2 cross-check) — reused here to avoid parsing the PRF's VML
    twice. Reads fresh if not provided (e.g. when called standalone/in tests).
    """
    # Read Section 1 VML checkbox states BEFORE openpyxl loads the file
    # (openpyxl would strip VML context; reading raw zip preserves accuracy).
    #
    # 2026-07-14 fix: this used to swallow read failures and silently default
    # every unresolved checkbox to unchecked (☐) — indistinguishable from the
    # engineer having actually left it unchecked. That was the root cause of
    # Section 1.2/1.3 checkboxes coming out wrong or blank on repeated runs.
    # _read_vml_checkbox_states() now raises ChecklistCheckboxError instead of
    # ever returning a partial/empty result, so we stop here rather than guess.
    if s1_states is None:
        try:
            s1_states = _read_vml_checkbox_states(
                prf_path, "Sections 1-2", required_cells=list(_S1_CHECKBOXES.keys())
            )
        except ChecklistCheckboxError as e:
            raise RuntimeError(
                f"BLOCKED — {prf_path.name}: Section 1.2/1.3 checkbox state could not be "
                f"verified ({e}). This automation previously defaulted unreadable "
                f"checkboxes to unchecked, which produced incorrect quality records — "
                f"that behavior was removed on 2026-07-14 (see CLAUDE.md changelog). "
                f"Open the PRF in Excel, confirm the actual state of 1.2 (D8/G8) and "
                f"1.3 (D13/N13/P13), re-save from Excel, and re-run. Do not proceed by "
                f"guessing these values."
            ) from e
    vml_states = s1_states
    print(f"  Section 1 checkboxes read: {vml_states}")

    # Discover every OTHER native checkbox on Section 3 and Sections 4-6 (Section
    # 4's own checkboxes live on the "Sections 4-6" sheet, above where Section
    # 5/6 start). Unlike Section 1, no fixed cell list is required or assumed —
    # the whole point is to stop hardcoding which cells matter. A structural
    # read failure here still raises (corrupt file); a sheet with genuinely no
    # native checkboxes just returns {}.
    try:
        s3_checkbox_states = _read_vml_checkbox_states(prf_path, "Section 3")
    except ChecklistCheckboxError as e:
        raise RuntimeError(
            f"BLOCKED — {prf_path.name}: Section 3 checkbox state could not be read ({e}). "
            f"Open the PRF in Excel, confirm it isn't corrupted, re-save, and re-run."
        ) from e
    try:
        s46_checkbox_states = _read_vml_checkbox_states(prf_path, "Sections 4-6")
    except ChecklistCheckboxError as e:
        raise RuntimeError(
            f"BLOCKED — {prf_path.name}: Sections 4-6 checkbox state could not be read "
            f"({e}). Open the PRF in Excel, confirm it isn't corrupted, re-save, and re-run."
        ) from e
    if s3_checkbox_states:
        print(f"  Section 3 checkboxes discovered: {list(s3_checkbox_states)}")
    if s46_checkbox_states:
        print(f"  Sections 4-6 checkboxes discovered: {list(s46_checkbox_states)}")

    # Track every checkbox this run writes so we can verify fidelity right
    # before saving — catches write-side or zip-surgery regressions too.
    _expected_s12 = {}
    _expected_s3 = {}
    _expected_s46 = {}

    # Track ONLY the automation-DETERMINED checkboxes (Section 2/5/6 answers,
    # plus the two automation-owned N/A indicators W2/G3) separately from the
    # dicts above, which also include Section 1/3/4 cells this function only
    # re-renders (never determines). Used below to set the REAL native
    # form-control state to match — see the restore_form_controls block.
    _native_s12 = {}
    _native_s3  = {}
    _native_s46 = {}

    # Pristine copy of the PRF exactly as it exists on entry to this
    # function, BEFORE openpyxl touches it — added 2026-09-11 (WI-7.4.3-01
    # gap analysis, Priority 1; confirmed catastrophic on PO26-1001). This
    # file's own sheet XML still carries the intact <drawing>/<legacyDrawing>/
    # <controls> tail elements that openpyxl's save always drops; it is the
    # only valid "template" restore_form_controls() can splice those elements
    # back from after save_preserving_drawings() overwrites prf_path in
    # place below. See that call site for why this matters: without it, the
    # VML/ctrlProps parts save_preserving_drawings() carries into the zip are
    # never re-referenced by the sheet XML, and Excel silently PRUNES every
    # unreferenced part — checkboxes and the ACME Medical logo included — the
    # very next time a human opens and saves the delivered file.
    _pristine_copy = TMP_DIR / f"__pristine_{prf_path.name}"
    TMP_DIR.mkdir(exist_ok=True)
    shutil.copy2(str(prf_path), str(_pristine_copy))

    wb = load_workbook(prf_path, rich_text=True)
    s12 = wb["Sections 1-2"]
    s3  = wb["Section 3"]
    s46 = wb["Sections 4-6"]

    # ── Section 1 — READ ONLY, nothing written ───────────────────────────────
    # REMOVED 2026-09-17 (QE, direct instruction). This loop used to write a
    # ☑/☐ text glyph plus a re-derived label into each of Section 1's checkbox
    # cells, on the premise (stated in PRE-RUN REQUIREMENTS and this function's
    # own docstring) that "LibreOffice cannot render native VML checkboxes at
    # all, so without this they render as blank in the exported PDF."
    #
    # That premise is false on this LibreOffice build, established by rendering
    # the real export at 400 DPI: the native control DOES render. So the glyph
    # was never the only mark — it was a SECOND mark written over the
    # template's own checkbox, which is why every field on the exported PRF
    # showed a doubled box. It was also overwriting the template's own cell
    # contents to do it: D8 'N/A' -> '☑  N/A', G8 '         100% Inspection' ->
    # '☐  100% Inspection' (losing the template's own leading padding), K30
    # 'Yes' -> '☐  Yes', B31 '6.2\n         \n       N/A' -> '6.2\n☑  N/A',
    # and a bare '☑'/'☐' into Section 2's answer cells, which the template
    # ships EMPTY. Every one of those is a style change to a controlled form,
    # made on cells this WI elsewhere calls read-only.
    #
    # The template's native form controls are now the checkboxes of record,
    # for engineer-owned and automation-owned answers alike:
    #   * Section 1/3/4 (engineer-owned): read, never written — the engineer's
    #     own native control already carries and now renders their answer.
    #   * Section 2/5/6 (automation-owned): set via set_native_checkboxes()
    #     further below, which has existed since 2026-09-11 and is what makes
    #     dropping the glyph possible at all — before it, the real control was
    #     left at whatever the blank template shipped with and the glyph was
    #     the only thing carrying the answer.
    #   * normalize_checkbox_marks() then turns a checked control's X
    #     appearance into a clean vector checkmark on the exported PDF (see
    #     process_part(); WI Step 3's CAUTION requires checkmarks, never X's).
    # _expected_s12/_s3/_s46 stay populated below purely so the run still has
    # an explicit record of what it intended; the glyph-fidelity gate they fed
    # is replaced by verify_native_checkboxes() (see the gate block).

    # ── Section 3 — READ ONLY, nothing written ───────────────────────────────
    # REMOVED 2026-09-17 alongside the Section 1 loop above, same reason and
    # same evidence. The engineer's per-dimension pass/fail controls render
    # themselves; W2 (automation-owned) is set natively just below.
    for coord, info in s3_checkbox_states.items():
        _expected_s3[coord] = info["checked"]

    # ── Section 3 — restore N/A checkbox (W2) ────────────────────────────────
    if section2.get("section3_na"):
        v2 = _anchor_cell(s3, "V2")
        # Guard: never overwrite a formula cell. Per template, R2/U2 are formula-linked.
        # 2026-09-17: also never RESTYLE a label the template already carries.
        # V2 ships as "N/A" in this template, so the old unconditional write
        # re-set its font and alignment to this function's own idea of them on
        # every run — a silent style change to a controlled form, the same
        # class of overwrite as the glyph writes removed above. Now written
        # only when the cell is genuinely empty.
        v2_is_formula = isinstance(v2.value, str) and v2.value.startswith("=")
        if v2_is_formula:
            print(f"  [WARN] Section 3 N/A label cell resolved to formula cell "
                  f"{v2.coordinate} — skipped label write, set checkbox only.")
        elif not (v2.value and str(v2.value).strip()):
            v2.value = "N/A"
            v2.font  = Font(name="Calibri", size=11)
            v2.alignment = Alignment(horizontal="right", vertical="center")
        _expected_s3["W2"] = True
        _native_s3["W2"] = True
    else:
        # Per CLAUDE.md, W2 must be unchecked when Section 3 has dimensions
        # (engineer-filled). Set explicitly rather than left alone — the same
        # "assume the old value is still correct" gap that caused the Section 1
        # bug — but now on the native control, not a text glyph.
        _expected_s3["W2"] = False
        _native_s3["W2"] = False

    # ── Section 2 ────────────────────────────────────────────────────────────
    # 2.1 — checked only when determine_2_1() actually resolved to
    # "Completed". 2026-09-17 fix: this used to check I20 unconditionally,
    # even when determine_section2() had returned "INCOMPLETE — packing slip
    # not found or unsigned", so a run that could NOT establish a signed
    # packing slip still asserted "Completed" on the signed quality record —
    # the one thing WI-7.4.3-01 Step 7's FLAG convention exists to prevent,
    # and contrary to Step 8's own "[only if not FLAG]" note on this cell.
    _s21_completed = str(section2.get("2.1", "")).strip().lower() == "completed"
    _expected_s12['I20'] = _s21_completed
    _native_s12['I20'] = _s21_completed

    # 2.2 material/finish annotation — DISABLED 2026-07-22 (found on
    # PO26-1003/SR90-1004 via manual QE correction, recorded in that part's
    # manifest). This used to overwrite cell B21 with a rich-text paragraph
    # ("Manufacturing Certifications Match Drawing Notes... Material: ...
    # Finishing: ..."), but B21 is actually the row's own "2.2" item-number
    # label — the same class of cell as Section 2's documented K21-K25/
    # N21-N25/Q21-Q24 "NEVER touch" labels, just not previously identified as
    # one. Every run was clobbering that label with the material/finish text,
    # requiring QE to manually restore "2.2" afterward. Disabled outright
    # rather than guessing a replacement cell — repeating the "assume a
    # coordinate without confirming it against the real template" mistake
    # this WI exists to prevent. See CLAUDE.md Open Item #12: needs the real
    # QF 7.4.3-02 template inspected (by a session with Box file access) to
    # find where a material/finish annotation actually belongs, if anywhere,
    # before this is re-enabled.
    # if drawing_data:
    #     mat = drawing_data.get("material", "N/A") or "N/A"
    #     fin = drawing_data.get("finish",   "N/A") or "N/A"
    #     plain = InlineFont(rFont="Calibri", sz=10)
    #     under = InlineFont(rFont="Calibri", sz=10, u="single")
    #     label_cell = _anchor_cell(s12, "B21")
    #     label_cell.value = CellRichText(
    #         TextBlock(plain, "Manufacturing Certifications Match Drawing Notes "
    #                          "(list all that apply):\nMaterial: "),
    #         TextBlock(under, mat),
    #         TextBlock(plain, "\nFinishing: "),
    #         TextBlock(under, fin),
    #     )
    #     label_cell.alignment = Alignment(wrap_text=True, vertical="top")

    def set_row(yes, no, na, val):
        # Native control only (2026-09-17) — these three cells ship EMPTY in the
        # template; the visible "Yes"/"No"/"N/A" text lives in its own label
        # cell (K/N/Q), which this function has never touched and still doesn't.
        if val == "Yes":
            _expected_s12[yes] = True; _expected_s12[no] = False
            if na: _expected_s12[na] = False
        elif val == "No":
            _expected_s12[yes] = False; _expected_s12[no] = True
            if na: _expected_s12[na] = False
        else:
            _expected_s12[yes] = False; _expected_s12[no] = False
            if na: _expected_s12[na] = True
        # Mirror into _native_s12 — these are the automation-determined
        # Section 2 answers, the whole reason the native-checkbox wiring
        # below exists (see restore_form_controls block).
        for _c in (yes, no, na):
            if _c:
                _native_s12[_c] = _expected_s12[_c]

    set_row('I21','L21','O21', section2.get("2.2","N/A"))
    set_row('I22','L22','O22', section2.get("2.3","N/A"))
    set_row('I23','L23','O23', section2.get("2.4","N/A"))
    set_row('I24','L24','O24', section2.get("2.5","N/A"))

    if section2.get("2.6") == "Yes":
        _expected_s12['I25'] = True; _expected_s12['L25'] = False
    else:
        _expected_s12['I25'] = False; _expected_s12['L25'] = True
    _native_s12['I25'] = _expected_s12['I25']; _native_s12['L25'] = _expected_s12['L25']

    # ── Signature/initials fields conditional on 2.4 / 2.6 ──────────────────
    # Added 2026-09-08 per explicit QE instruction (PO26-1007/VendorE): the PRF
    # template has two conditional sign-off fields on this same "Sections
    # 1-2" sheet that this function never touched before, even though they
    # are governed entirely by 2.4/2.6 (values this function already
    # computes) and this function already owns writing to Section 2:
    #   H29        "If Item 2.4 is YES, Quality Manager Review and Approval
    #               Initials:" (merged label B29:G29, entry cell H29:I29).
    #               Ships BLANK in the template — no default N/A of its own.
    #   I32 / O32  "If Item 2.6 is YES, Director of Operations Review and
    #               Approval Signature:" (label B32:H32) and its Date field
    #               (label N32). The template DOES ship a placeholder "N/A"
    #               in both I32 and O32.
    # Leaving H29 untouched meant a 2.4 = No/N/A part shipped with a BLANK
    # Quality Manager initials field — indistinguishable from "reviewer
    # forgot to sign" rather than "not applicable, reviewed." And because
    # nothing ever re-asserted the I32/O32 default, any live PRF where that
    # placeholder had been cleared upstream (copy/paste from a prior part,
    # a template edit, etc.) would ship with a blank Director signature
    # field on a 2.6 = No part with no automated check to catch it.
    #
    # Fix: whenever the governing item is NOT "Yes", write "N/A" into its
    # dependent field(s), the same convention the template already uses for
    # I32/O32. Whenever the governing item IS "Yes" (2.6 = Yes only for
    # confidential Biologic per TRAVELER_VENDORS — see determine_section2),
    # clear out any stray leftover "N/A" instead, since that cell must then
    # carry a real signature/initials, not a leftover default — never
    # invents "Yes"/"No" here, only ever N/A vs. untouched-for-human-entry.
    qm_initials_cell = _anchor_cell(s12, "H29")
    if section2.get("2.4") == "Yes":
        if isinstance(qm_initials_cell.value, str) and qm_initials_cell.value.strip().upper() == "N/A":
            qm_initials_cell.value = None
    else:
        qm_initials_cell.value = "N/A"

    director_sig_cell  = _anchor_cell(s12, "I32")
    director_date_cell = _anchor_cell(s12, "O32")
    if section2.get("2.6") == "Yes":
        for _cell in (director_sig_cell, director_date_cell):
            if isinstance(_cell.value, str) and _cell.value.strip().upper() == "N/A":
                _cell.value = None
    else:
        director_sig_cell.value = "N/A"
        director_date_cell.value = "N/A"

    # ── Section 4 — render every discovered native checkbox read-only ────────
    # Same rendering-only pass as Section 3, applied to whatever checkboxes
    # exist on the Sections 4-6 sheet (this covers Section 4 — DVT/verification
    # checkboxes — since it shares a sheet with Sections 5/6). This draws the
    # engineer's own already-set state; it never changes an answer. The
    # explicit G3 write immediately below is automation-owned and runs AFTER
    # this loop, so it always wins for G3 specifically. Section 5/6 cells this
    # function writes further down (L12, I30/M30, I36/M36, B31, etc.) are all
    # written AFTER this loop too, so they are unaffected by it.
    #
    # 2026-07-27 fix (found on PO26-1006/SR90-1003, resolves Open Item #11):
    # G2 is a leftover template-construction artifact, not a real field — it
    # is not referenced anywhere in this WI's Section 4/5/6 spec and has no
    # associated label text of its own, so the generalized "render every
    # discovered checkbox" pass was drawing a bare, unlabeled ☑/☐ floating
    # above SECTION 4's own N/A checkbox on every run since 2026-07-20.
    # Confirmed directly by the QE (not inferred): this should never render
    # at all. Excluded from the rendering pass, and any stray glyph a past
    # run already wrote there is cleared below.
    #
    # 2026-09-17 correction: this exclusion + text-clear was NOT sufficient
    # on its own — it stopped this pass from writing a NEW glyph into G2,
    # but did nothing about the native VML checkbox shape already sitting
    # there, which renders in LibreOffice's PDF export on its own regardless
    # of cell text. Confirmed on a real compiled PRF (SR90-1005, PO26-1010)
    # that the empty box was still visible next to Section 4's N/A on every
    # PRF produced since the July fix. The actual shape removal now happens
    # in strip_orphan_checkbox() near the end of this function, after
    # restore_form_controls() — the exclusion/clear below is still needed
    # too (it stops this run's own loop from writing a glyph there before
    # that later removal happens).
    _S46_IGNORED_CHECKBOXES = {'G2'}
    g2_stray = s46['G2']
    if g2_stray.value not in (None, ''):
        g2_stray.value = None

    # 2026-09-17: rendering-only glyph writes removed here too (see the
    # Section 1 block). Section 4's own controls render themselves; Sections
    # 5/6's automation-owned answers are set natively below.
    for coord, info in s46_checkbox_states.items():
        if coord in _S46_IGNORED_CHECKBOXES:
            continue
        _expected_s46[coord] = info["checked"]

    # ── Section 4 — do not modify; engineer owns this section entirely ─────────
    # Ensure Section 4 N/A checkbox is unchecked if engineer wrote content
    existing_s4 = s46['B4'].value
    if existing_s4 and str(existing_s4).strip():
        _expected_s46['G3'] = False
        # 2026-09-17 (QE, direct report: "section 4's box next to N/A
        # disappeared"). Section 4's N/A is the ONE checkbox field in this
        # template with no native form control behind it — confirmed by
        # reading vmlDrawing4.vml's own anchors, which carry exactly six
        # shapes for this sheet: K30 (6.1 Yes), J36 (Raw Goods), L12
        # (Section 5 N/A), N30 (6.1 No), B31 (6.2 N/A) and N36 (Finished
        # Goods). Nothing is anchored at or near G3, and the 'G2' that
        # _read_vml_checkbox_states() reports is the anchorless phantom
        # already documented as a template artifact (Open Item #11).
        #
        # So the box that has always appeared next to Section 4's "N/A" was
        # drawn by THIS automation's own ☑/☐ glyph — it is the one field where
        # the glyph was never a duplicate of anything, and removing the glyph
        # overlay wholesale took it away with the real duplicates. Restored
        # here as an explicit, narrowly-scoped FALLBACK, and only when the
        # template genuinely has no control to use instead:
        #   native control present -> set the control, write no glyph
        #                             (every other checkbox on the form)
        #   no native control      -> write the glyph, since it is then the
        #                             only mark that can carry the answer
        # Keyed off the discovered control list rather than hardcoded to G3,
        # so the same fallback covers any other template field that turns out
        # to have no control behind it.
        if 'G3' in s46_checkbox_states:
            _native_s46['G3'] = False
        else:
            uncheck_box(s46, 'G3', with_label='N/A')
            print("  Section 4 N/A (G3): no native form control in this template — "
                  "wrote the ☐ glyph as the field's only mark (see CLAUDE.md "
                  "2026-09-17).")

    # ── Section 5 ─────────────────────────────────────────────────────────────
    if discrepancy_flags:
        disc_text = call_claude(
            "Write a concise Section 5 discrepancy description for a medical device "
            "receiving inspection PRF. Do NOT recommend NCR disposition — Engineering decides.\n"
            "Flags:\n" + "\n".join(f"- {f}" for f in discrepancy_flags),
            task="narrative"
        )
        _expected_s46['L12'] = False
        _native_s46['L12'] = False
        b14 = _anchor_cell(s46, 'B14')
        b14.value = disc_text
        b14.alignment = Alignment(wrap_text=True, vertical='top')
    else:
        _expected_s46['L12'] = True
        _native_s46['L12'] = True

    # ── Section 6 ─────────────────────────────────────────────────────────────
    # 2026-07-22 fix (found on PO26-1003/SR90-1004 via manual QE correction,
    # recorded in that part's manifest): this code previously wrote the 6.1
    # Yes/No checkbox to I30/M30 and the disposition checkbox to I36/M36,
    # which produced a spurious "doubled" glyph while leaving K30/N30/J36/N36
    # (the real answer checkboxes) blank. Corrected to write only to the real
    # cells — that part of the fix was right.
    #
    # 2026-07-27 follow-up fix (found on PO26-1006/SR90-1003, confirmed
    # directly against a real blank QF 7.4.3-02 template — resolves the
    # "NOT yet re-confirmed" caveat from Open Item #11): the 2026-07-22 fix
    # assumed K30/N30/J36/N36 were bare answer-only checkboxes with the
    # "Yes"/"No"/"Raw Goods"/"Finished Goods" text living in some other,
    # separate label cell (the way Section 2's I/K, L/N, O/Q pairs work).
    # That assumption was wrong: on the real template the label text is
    # baked into the SAME cell as the checkbox glyph (e.g. K30's pristine
    # value is "☐Yes", not just "☐"). check_box()/uncheck_box() called
    # without with_label overwrite the entire cell value with a bare glyph,
    # so every run silently deleted "Yes"/"No"/"Raw Goods"/"Finished Goods"
    # from the signed record — a second, different instance of the same
    # class of bug the 2026-07-22 fix was trying to close. Fixed by passing
    # with_label explicitly here, the same pattern already used everywhere
    # else a label must survive the checkbox rendering pass.
    # 2026-07-27 fix (found on PO26-1006/SR90-1003, resolves part of Open Item
    # #11's uncertainty): B31 (the "6.2 [ ] N/A" label+checkbox) is one of the
    # native checkboxes s46_checkbox_states discovers, so the generic
    # rendering pass above already drew SOME glyph into it — but only a bare
    # ☑/☐ with no caption, since this cell's VML shape has no caption text of
    # its own (the "6.2"/"N/A" text is baked into the cell value, the same
    # pattern as K30/N30/J36/N36 above). The old code only overrode that with
    # the real "6.2\nN/A" label in the no-discrepancy branch, so whenever
    # discrepancies WERE present (the common case — 6.2 doesn't apply as N/A
    # then either, since real disposition data belongs in the table below,
    # but the "6.2"/"N/A" text should still be visible, just unchecked) the
    # bare glyph from the generic pass was left showing with no label at all.
    # Now written explicitly in both branches so the label always survives.
    # 2026-09-17 fix (QE, direct report): B31 was hand-rolled outside
    # check_box()/uncheck_box() (it carries two lines — "6.2" and the N/A
    # checkbox — in one merged cell, so the shared helper's single-glyph
    # value format doesn't fit it), and its Alignment call hardcoded
    # vertical='top' with no horizontal set at all. That pinned "6.2"/"☐ N/A"
    # to the top-left corner of the merged field instead of the field's
    # center, the same drift-to-the-side/top symptom just fixed in
    # check_box()/uncheck_box() itself, just not reachable by that fix since
    # this cell never calls them. Centered on both axes now, matching every
    # other checkbox+label field in the template.
    # 2026-09-17: B31's hand-rolled value/font/alignment write is gone with the
    # rest of the glyph overlay. B31 has a native control of its own (it is in
    # s46_checkbox_states), and the template's own cell text — '6.2', a blank
    # line, then 'N/A' — is the label; the old code replaced that whole string
    # with '6.2\n☑  N/A' and re-set the font and alignment on every run. That
    # is what the 2026-07-27 and 2026-09-17 entries above were repeatedly
    # patching: both were fixing how a mark drawn over the original looked,
    # rather than not drawing it. The native control carries the answer now and
    # the template's label is left exactly as the engineer's file has it.
    if discrepancy_flags:
        _expected_s46['K30'] = True; _expected_s46['N30'] = False
        _expected_s46['B31'] = False
        _native_s46['K30'] = True; _native_s46['N30'] = False; _native_s46['B31'] = False
    else:
        _expected_s46['K30'] = False; _expected_s46['N30'] = True
        _expected_s46['B31'] = True
        _native_s46['K30'] = False; _native_s46['N30'] = True; _native_s46['B31'] = True

    # 2026-09-11: disposition can now resolve to "FLAG" (Open Item #13 fix —
    # the part# is on the QF 7.2-01 APDL, so it MIGHT be Finished Goods, but
    # automation cannot tell mid-process Raw Goods from the finished catalog
    # item by catalog number alone). Per the file-wide FLAG convention, leave
    # BOTH checkboxes unchecked rather than guess either answer — the
    # engineer resolves it by hand and checks the correct one themselves.
    _disposition = section2.get("disposition")
    if _disposition == "Finished Goods":
        _expected_s46['J36'] = False; _expected_s46['N36'] = True
    elif _disposition == "FLAG":
        _expected_s46['J36'] = False; _expected_s46['N36'] = False
    else:
        _expected_s46['J36'] = True; _expected_s46['N36'] = False
    _native_s46['J36'] = _expected_s46['J36']; _native_s46['N36'] = _expected_s46['N36']

    # ── Fidelity gate — added 2026-07-14 ───────────────────────────────────────
    # Confirm every checkbox this run intended to write actually landed correctly
    # BEFORE saving. Raises ChecklistCheckboxError (loud, unmissable) instead of
    # letting a wrong/blank checkbox reach In-Progress/ or Inspections/.
    # 2026-09-17: verify_checkbox_fidelity() checked that the ☑/☐ TEXT GLYPH
    # this function used to write had landed in the cell. There is no glyph to
    # verify any more — the native control is the answer of record, and
    # verify_native_checkboxes() (already called after the save below) is the
    # equivalent fail-loud gate for it. The helper itself is left in
    # prf_helpers.py untouched for any caller that still renders glyphs.
    #
    # Instead, clean up a stray glyph an EARLIER run of this tool may have left
    # in one of the automation-owned cells, so re-running over a
    # previously-processed PRF restores the template's own text instead of
    # leaving a doubled mark behind. Scoped to automation-owned cells only:
    # Section 1/3/4 cells are engineer-owned and are not touched even to tidy
    # them (a past run's glyph there is a pre-existing record to be corrected
    # by hand, not silently rewritten by automation).
    # Note: iterates the _native_* dicts only, so a cell written by the
    # no-native-control glyph fallback above (e.g. Section 4's G3) is never in
    # scope here — it is not a native cell, and its glyph is the field's only
    # mark, not leftover residue to clear.
    _stray = 0
    for _ws, _cells in ((s12, _native_s12), (s3, _native_s3), (s46, _native_s46)):
        for _coord in _cells:
            _c = _anchor_cell(_ws, _coord)
            if not isinstance(_c.value, str) or not any(g in _c.value for g in ("☑", "☐")):
                continue
            _cleaned = _c.value.replace("☑", "").replace("☐", "").strip()
            _c.value = _cleaned or None
            _stray += 1
    if _stray:
        print(f"  Cleared {_stray} leftover checkbox glyph(s) written into cell text "
              f"by an earlier run — the template's native checkboxes are the "
              f"checkboxes of record (see the 2026-09-17 CLAUDE.md entry).")

    save_preserving_drawings(wb, str(prf_path), str(prf_path))

    # ── Restore native form-control references — added 2026-09-11 ────────────
    # (WI-7.4.3-01 gap analysis, Priority 1; confirmed catastrophic on
    # PO26-1001/CAT-XX-60). save_preserving_drawings() carries the VML/
    # ctrlProps parts into the saved zip, but openpyxl's regenerated sheet
    # XML no longer references them (it drops the <drawing>/<legacyDrawing>/
    # <controls> elements entirely) — so those parts are present but
    # ORPHANED. Excel does not merely ignore an orphaned part: the next time
    # a human opens this file and saves it in Excel, Excel prunes every part
    # with no live reference — discarding the native checkboxes AND the
    # ACME Medical letterhead logo (xl/media/image1.png) permanently, not just
    # "until re-run." restore_form_controls() splices the pristine copy's
    # own trailing elements back in and forces the matching worksheet .rels
    # so every restored r:id resolves; it raises ChecklistCheckboxError
    # (loud, blocks the run) rather than silently continuing if the saved
    # file holds a worksheet relationship the pristine copy does not
    # (template-wins would silently discard it).
    try:
        restore_form_controls(str(prf_path), str(_pristine_copy))
        verify_xlsx_package_integrity(str(prf_path))
    except ChecklistCheckboxError as e:
        raise RuntimeError(
            f"BLOCKED — {prf_path.name}: could not restore native form-control "
            f"references after saving ({e}). Delivering this file as-is risks "
            f"Excel silently discarding checkboxes and/or the logo the next "
            f"time a human opens and saves it — do not deliver it. Open the "
            f"PRF, confirm it isn't corrupted, and re-run."
        ) from e

    # ── Remove the G2 leftover checkbox artifact — added 2026-09-17 ──────────
    # Actually closes Open Item #11. The 2026-07-27 fix (_S46_IGNORED_
    # CHECKBOXES below) only ever cleared stray TEXT a prior run had written
    # into G2's cell — it never touched the native checkbox shape itself,
    # which restore_form_controls() just spliced back in from the template
    # along with every other control. That shape renders on its own in
    # LibreOffice's PDF export regardless of cell text, confirmed directly
    # on a real, already-issued compiled PRF (SR90-1005, PO26-1010): a
    # small, empty, unlabeled checkbox sits next to Section 4's N/A on
    # every PRF this tool has ever produced. Must run AFTER
    # restore_form_controls() (which would otherwise put it right back) and
    # BEFORE the native-checkbox verification below. See
    # strip_orphan_checkbox()'s docstring for why both halves (VML shape +
    # <control> element) are removed together, not just one.
    try:
        strip_orphan_checkbox(str(prf_path), "Sections 4-6", "G2")
        verify_xlsx_package_integrity(str(prf_path))
    except ChecklistCheckboxError as e:
        raise RuntimeError(
            f"BLOCKED — {prf_path.name}: could not remove the G2 leftover "
            f"checkbox artifact ({e}). Delivering this file as-is risks a "
            f"half-removed checkbox breaking native-checkbox verification on "
            f"a future run — do not deliver it. Open the PRF, confirm it "
            f"isn't corrupted, and re-run."
        ) from e

    # ── Set REAL native checkbox state for automation-determined answers ─────
    # Everything above (check_box()/uncheck_box()) only ever drew a ☑/☐ TEXT
    # glyph into the cell, for LibreOffice's benefit — it never touched the
    # actual native form-control checkbox underneath, which for every
    # Section 2/5/6 cell this function determines has, until now, been left
    # at whatever state the blank template shipped with (normally
    # unchecked): the glyph looked right in the exported PDF while the real,
    # clickable Excel checkbox told a different story. set_native_checkboxes()
    # makes the real control agree with the glyph for exactly the
    # automation-owned cells tracked in _native_s12/_native_s3/_native_s46
    # above — never for the Section 1/3/4 cells this function only
    # re-renders (read-only; left exactly as the engineer's own native
    # control already had them). verify_native_checkboxes() is the same
    # fail-loud post-write gate as verify_checkbox_fidelity(), for the
    # native control instead of the text glyph.
    _native_states = {s: d for s, d in {
        "Sections 1-2": _native_s12, "Section 3": _native_s3, "Sections 4-6": _native_s46,
    }.items() if d}
    if _native_states:
        set_native_checkboxes(str(prf_path), _native_states)
        verify_native_checkboxes(str(prf_path), _native_states)

    # ── Post-save re-verification — added 2026-07-14 ───────────────────────────
    # save_preserving_drawings() does zip surgery (re-injecting VML/drawing parts
    # openpyxl strips on save). Re-open the file we just wrote from disk and
    # confirm the checkboxes still read back correctly — this is the gate that
    # catches a write-side or zip-merge regression, not just a read-side one.
    # 2026-09-17: re-reads the NATIVE control states back off the saved file,
    # which is what this gate was always for — catching a write-side or
    # zip-surgery regression, not just a read-side one — now that the native
    # control rather than a text glyph carries the answer. Same fail-loud
    # contract; verify_native_checkboxes() raises ChecklistCheckboxError.
    if _native_states:
        verify_native_checkboxes(str(prf_path), _native_states)

    return prf_path


# ─────────────────────────────────────────────────────────────────────────────
# PO PAGE STRIPPING
# ─────────────────────────────────────────────────────────────────────────────

def _po_page_kind(pdf_path: Path, page_idx: int) -> str:
    """
    Classify one page of a PO PDF as "drawing", "quote", or "po".

    Factored out of extract_po_pages_only() 2026-09-17 so the drawing pages
    it identifies can also be ROUTED to the drawing slot of the compiled
    package (extract_drawing_pages_from_po below) instead of only being
    stripped and thrown away. The indicator expressions are unchanged.
    """
    text = extract_page_text(pdf_path, page_idx).lower()
    is_drawing = (("drawn by" in text or "checked by" in text) and
                  ("revision" in text or " rev " in text)) or \
                 ("tolerance" in text and ("scale" in text or "drawn" in text)) or \
                 ("drawing" in text and "revision" in text and "sheet" in text)
    is_quote = ("quotation" in text) or \
               ("quote" in text and ("validity" in text or "estimate" in text))
    if is_drawing:
        return "drawing"
    if is_quote:
        return "quote"
    return "po"


def extract_drawing_pages_from_po(pdf_path: Path, tmp_dir: Path,
                                   part_number: str = "") -> tuple:
    """
    Return (Path | None, names_part: bool) for the drawing page(s) bundled
    into a PO PDF, extracted to their own file in tmp_dir.

    Added 2026-09-17 (found on PO26-1004/VendorB). extract_po_pages_only() has
    always known that "many POs have the engineering drawing attached as the
    last page(s)" and stripped those pages out of the PO slot — but nothing
    ever put them back anywhere, and process_part() looks for a drawing only
    in Drawings/ or Drawings & Stepfiles/. On PO26-1004 the only copy of
    SR90-1002 Rev A on file is page 2 of the PO PDF and Drawings &
    Stepfiles/ is empty, so the compiled package would have had NO drawing
    at all (WI-7.4.3-01 Step 3 position 4), and drawing_data would have
    fallen back to its material="N/A"/finish="N/A" default — which silently
    resolves 2.2 to N/A instead of running determine_2_2 against the real
    title block. Both are quality-record defects, not cosmetic.

    names_part is True when the returned pages' own text layer contains
    part_number — i.e. the drawing is positively identified as this part's.
    When no drawing page names the part (a scanned/vector-only title block,
    or a multi-part PO), every drawing page is still returned but
    names_part is False so the caller can flag it for human confirmation
    rather than silently compiling a possibly-wrong drawing.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
    if page_count <= 1:
        return None, False

    drawing_pages = [i for i in range(page_count)
                     if _po_page_kind(pdf_path, i) == "drawing"]
    if not drawing_pages:
        return None, False

    pn = (part_number or "").strip().upper()
    named = [i for i in drawing_pages
             if pn and pn in extract_page_text(pdf_path, i).upper()]
    use_pages, names_part = (named, True) if named else (drawing_pages, False)

    tmp_dir.mkdir(exist_ok=True)
    out_path = tmp_dir / f"po_bundled_drawing_{pn or 'unknown'}.pdf"
    doc = fitz.open(str(pdf_path))
    out_doc = fitz.open()
    for p in use_pages:
        out_doc.insert_pdf(doc, from_page=p, to_page=p)
    out_doc.save(str(out_path))
    out_doc.close()
    doc.close()
    return out_path, names_part


def extract_po_pages_only(pdf_path: Path, tmp_dir: Path) -> Path:
    """
    Strip drawing pages and quote pages appended to the PO PDF.
    Many POs have the engineering drawing attached as the last page(s).
    Returns path to cleaned PDF, or original if nothing stripped.

    The stripped drawing pages are not discarded — extract_drawing_pages_from_po()
    above re-extracts them for the drawing slot when no standalone drawing
    file exists (2026-09-17).
    """
    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
    if page_count <= 1:
        return pdf_path

    keep = [i for i in range(page_count) if _po_page_kind(pdf_path, i) == "po"]

    if len(keep) == page_count:
        return pdf_path  # Nothing to strip

    out_path = tmp_dir / f"po_stripped.pdf"
    doc = fitz.open(str(pdf_path))
    out_doc = fitz.open()
    for p in keep:
        out_doc.insert_pdf(doc, from_page=p, to_page=p)
    out_doc.save(str(out_path))
    out_doc.close()
    doc.close()
    print(f"  PO: stripped {page_count - len(keep)} non-PO page(s)")
    return out_path


def strip_invoice_pages(pdf_path: Path, tmp_dir: Path,
                        drop_labels: set = None) -> Path:
    """
    drop_labels: EXTRA page labels to strip beyond the always-excluded
    invoice_exclude / sgt_doc_exclude. Added 2026-09-08 (found on
    PO26-1007/VendorE, reported by QE): the caller passes the document types it is
    already compiling independently out of COCs/ (mfg_coc, ster_coc, traveler,
    mfg_report), so a copy of one of those bundled into the packing slip PDF
    is dropped from the packing-slip slot instead of appearing twice in the
    package — once before the PRF and once in its own documented position.
    Only types the caller confirms it is including separately are ever passed,
    so this can never drop the only copy of a certificate.

    Pages whose text layer is empty are classified visually via
    _vision_classify_page() before this filter is applied; a failed/unknown
    vision result keeps the page (see that function's docstring).

    Strip commercial/proforma invoice pages (and any ACME Medical-generated
    pages) bundled inside a packing slip PDF. Keeps everything else,
    including pages that don't classify cleanly as "packing_slip" — this is
    deliberately exclude-only, not an allow-list, so a legitimate rider or
    continuation page that classify_page() can't confidently label is never
    dropped. Mirrors extract_po_pages_only()'s shape.

    2026-07-2x fix: WI-7.4.3-01 Step 3 has always said commercial/proforma
    invoices "often come bundled with packing slips from international
    vendors" and must be stripped — but that stripping was only ever wired
    up for combined PDFs inside COCs/ (via split_pdf_by_page_type(), called
    when a COC file has >2 pages). The packing slip found in Shipping Docs/
    was taken wholesale into packing_pdfs_inc with no page-level filtering
    at all, so a bundled invoice rode straight into every compiled package
    unless someone caught it by eye. Returns path to cleaned PDF, or the
    original if nothing was stripped (single-page packing slips are the
    common case and pay only the page-count check below).
    """
    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
    if page_count <= 1:
        return pdf_path

    drop = {"invoice_exclude", "sgt_doc_exclude"} | set(drop_labels or ())

    keep = []
    dropped_labels = []
    for i in range(page_count):
        text = extract_page_text(pdf_path, i)
        if text.strip():
            label = classify_page(text, pdf_path.name)
        else:
            # No text layer — classify_page() can only return "other" here, so
            # look at the page image instead (2026-09-08; see
            # _vision_classify_page). "" means unknown → keep the page.
            label = _vision_classify_page(pdf_path, i)
            print(f"  Packing slip p{i + 1}: no text layer — vision classified as "
                  f"'{label or 'unknown (keeping)'}'")
        if label in drop:
            dropped_labels.append(label)
            continue
        keep.append(i)

    if len(keep) == page_count:
        return pdf_path  # Nothing to strip

    if not keep:
        # Everything on this "packing slip" classified as excluded — almost
        # certainly a misclassification, not a genuinely empty packing slip.
        # Fail safe by keeping the original rather than compiling nothing.
        print(f"  WARNING: strip_invoice_pages found no keepable pages in "
              f"{pdf_path.name} (all classified {set(dropped_labels)}) — "
              f"keeping original file unfiltered rather than dropping it "
              f"entirely. Verify this packing slip manually.")
        return pdf_path

    out_path = tmp_dir / f"packing_slip_stripped.pdf"
    doc = fitz.open(str(pdf_path))
    out_doc = fitz.open()
    for p in keep:
        out_doc.insert_pdf(doc, from_page=p, to_page=p)
    out_doc.save(str(out_path))
    out_doc.close()
    doc.close()
    print(f"  Packing slip: stripped {page_count - len(keep)} page(s) "
          f"({set(dropped_labels)}) from {pdf_path.name}")
    return out_path


# Document types that have their own documented position in the compiled
# package (WI-7.4.3-01 Step 3 positions 6-9) and their own role in the
# Step 7 Section 2 determination — as opposed to a page that simply belongs
# with the packing slip.
_CERTIFICATE_LABELS = ("mfg_coc", "ster_coc", "traveler", "mfg_report")


def split_bundled_certificate_pages(pdf_path: Path, tmp_dir: Path) -> dict:
    """
    Split certificate/QC pages out of a packing-slip (or other shipping) PDF.
    Returns {label: [Path, ...]} for _CERTIFICATE_LABELS pages only; every
    other page is left alone.

    Added 2026-09-17 (found on PO26-1004/VendorB). COC classification has only
    ever scanned COCs/ — so a certificate the vendor bundled into the same
    PDF as the packing slip was never classified at all. That is not a
    cosmetic filing issue: it means has_mfg_coc/has_ster_coc/has_traveler are
    False in determine_section2() (2.2 resolves "No"/"N/A" on a lot that DOES
    have a valid COC on file), extract_coc_data() never runs on it (its lot#/
    qty/DOM/revision are never cross-referenced in Step 6),
    check_document_completeness() never reviews it (Step 6.5), and the page
    rides along in the packing-slip slot rather than position 6-9.

    On PO26-1004 the VendorB Certificate of Conformance for lot LOT-0001 is page 2
    of "Packing List 26-1004.pdf" and COCs/ is empty. The complementary
    machinery already exists on the other side: process_part() passes the
    types it is compiling separately to strip_invoice_pages(drop_labels=...)
    (added 2026-09-08 for PO26-1007/VendorE, the same vendor habit seen from the
    other direction), so a page promoted here is automatically dropped from
    the packing-slip slot instead of appearing twice.

    Classification per page is classify_page() when the page has a text
    layer, falling back to _vision_classify_page() when it does not — the
    same pair strip_invoice_pages() already uses, which matters because
    these bundled scans routinely have no text layer at all (PO26-1004's is
    a two-page scan of exactly this shape). An unknown/failed vision result
    promotes nothing, leaving the page with the packing slip.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
    if page_count <= 1:
        return {}

    buckets: dict = {}
    for i in range(page_count):
        text = extract_page_text(pdf_path, i)
        if text.strip():
            label = classify_page(text, pdf_path.name)
        else:
            label = _vision_classify_page(pdf_path, i)
        if label in _CERTIFICATE_LABELS:
            buckets.setdefault(label, []).append(i)

    if not buckets:
        return {}

    tmp_dir.mkdir(exist_ok=True)
    results: dict = {}
    doc = fitz.open(str(pdf_path))
    for label, pages in buckets.items():
        out_path = tmp_dir / f"{pdf_path.stem}_bundled_{label}.pdf"
        out_doc = fitz.open()
        for p in pages:
            out_doc.insert_pdf(doc, from_page=p, to_page=p)
        out_doc.save(str(out_path))
        out_doc.close()
        results.setdefault(label, []).append(out_path)
        print(f"  Packing slip: promoted p{[p + 1 for p in pages]} to "
              f"'{label}' (bundled certificate — compiled in its own position)")
    doc.close()
    return results


def _filter_classified_for_part(classified: dict, pn: str, all_part_numbers) -> dict:
    """
    Scope a PO-wide classified-COC dict down to the documents relevant to
    ONE part. Added 2026-09-11 (groundwork for Priority 3 item #1 / partial-
    PO hold, WI-7.4.3-01 gap analysis 2026-09-11): classify_coc()/
    split_pdf_by_page_type() only ever label a document by TYPE (mfg_coc/
    ster_coc/traveler/mfg_report/other) — nothing previously scoped
    `classified` to a specific part number, so on a multi-part PO every
    ready part drew from the exact same PO-wide dict via
    `next((p for p, t in classified.items() if t == ...), None)`, meaning
    part 2 and part 3 could silently compile part 1's COC into their own
    quality record. This was never exercised through receive.py itself —
    real multi-part POs (e.g. PO26-1008/VendorD) have gone through bespoke
    per-PO scripts instead (see Priority 3 item #7) — but is a real,
    FDA-relevant traceability defect in the governed path regardless of
    whether it has yet produced a wrong compiled package.

    A document is "part-specific" if its filename contains ANY ready part
    number on this PO as a substring — it then belongs ONLY to the part(s)
    whose number it names. A document whose filename contains NONE of the
    PO's ready part numbers is "unscoped" (e.g. a vendor's general COC that
    doesn't name a part number at all — the common case on the single-part
    POs this tool has mostly processed so far) and stays available to every
    part, exactly as before. Same filename-substring convention already
    used for drawing matching (`drawings_dir.glob(f"*{pn}*.pdf")`) and PRF
    matching (`find_prf()`) — not a new mechanism, just applied here too.
    """
    pn_u = (pn or "").upper()
    other_parts = {str(p).upper() for p in (all_part_numbers or ()) if str(p).upper() != pn_u}
    out = {}
    for path, label in classified.items():
        stem_u = path.stem.upper()
        names_this_part = bool(pn_u) and pn_u in stem_u
        names_other_part = any(op in stem_u for op in other_parts if op)
        if names_this_part or not names_other_part:
            out[path] = label
    return out


# ─────────────────────────────────────────────────────────────────────────────
# PART PROCESSING
# ─────────────────────────────────────────────────────────────────────────────

def process_part(po_folder: Path, po_data: dict, part_item: dict,
                 part_idx: int, cache: dict, cli_po_number: str = "",
                 force: bool = False, all_part_numbers=frozenset(),
                 po_doc: Path = None) -> dict:
    """
    po_doc: the PO document main() already resolved and validated for this
    run (added 2026-09-17, found on PO26-1004/VendorB). This function used to
    re-resolve it itself via find_po_document(po_folder,
    po_data["po_number"]) — i.e. with the PO number as PARSED OUT of the
    document ("P26-1004"), not the CLI argument main() searched with
    ("PO26-1004"). Those are different search strings and they can select
    different files: on PO26-1004 the real PO ("26-1004 - VendorB.pdf") does not
    contain "P26-1004" in its filename while VendorB's sales-order
    acknowledgement does, so the compiled package would have carried the
    vendor's acknowledgement in position 1 as "the PO." Reusing main()'s
    already-resolved path removes the possibility of the two disagreeing at
    all. Falls back to the old lookup (CLI number first) when not supplied,
    so a caller that omits it is not newly broken.

    all_part_numbers: every ready part number on this PO run (added
    2026-09-11) — used only to scope `classified` COC documents to this
    part via _filter_classified_for_part(); an empty/omitted set degrades
    to the pre-2026-09-11 behavior (every classified document is treated
    as unscoped/available to every part), so a caller that doesn't pass it
    is not newly broken, just not newly protected either.
    Process a single part through the full receiving workflow.

    force: passed straight through to assert_no_user_edits() at the compiled-
    package write below (added 2026-09-11) — discards a human's edits to an
    already-delivered package on an EXPLICIT operator instruction (--force on
    the CLI) only. Never set True by default.
    """
    pn  = part_item.get("part_number", "")
    rev = part_item.get("revision", "")
    vendor = po_data.get("vendor", "")
    # Prefer the PO number parsed from the PO document itself; fall back to
    # the CLI-supplied PO# (2026-07-22 fix, found on PO26-1003/CatalogSupplier) when
    # the source PO is a scanned image with no extractable text and parsing
    # yields nothing. This is not a guess — the CLI argument is exactly which
    # PO the user invoked this run for — it just avoids the output filename
    # coming out blank/malformed (e.g. " Receiving Part 1 - SR90-1004.pdf")
    # on a part whose PO document has no usable text layer.
    po_num = po_data.get("po_number", "") or cli_po_number

    forms_dir   = po_folder / F_FORMS
    cocs_dir    = po_folder / F_COCS
    drawings_dir = find_drawings_dir(po_folder)
    output_dir  = po_folder / F_OUTPUT
    output_dir.mkdir(exist_ok=True)
    TMP_DIR.mkdir(exist_ok=True)

    print(f"\n  Part {part_idx}: {pn} Rev {rev}")
    flags: list = []

    # ── Pre-filled PRF ────────────────────────────────────────────────────────
    prf_path = find_prf(forms_dir, pn)
    if not prf_path:
        return {"status": "error", "message": f"PRF not found in {F_FORMS}/"}

    # ── PO document ──────────────────────────────────────────────────────────
    # Resolved once here, from main()'s already-validated path when supplied
    # (see the po_doc parameter's docstring above), and reused by the drawing
    # fallback, the vision cross-check, and doc_order below.
    raw_po = po_doc or find_po_document(
        po_folder, cli_po_number or po_data.get("po_number", ""))

    # ── Drawing ───────────────────────────────────────────────────────────────
    drawing_path = None
    if drawings_dir:
        matches = sorted(drawings_dir.glob(f"{pn}*.pdf")) or \
                  sorted(drawings_dir.glob(f"*{pn}*.pdf"))
        drawing_path = matches[0] if matches else None

    # Fallback: the drawing is bundled into the PO PDF itself (2026-09-17,
    # PO26-1004/VendorB) — extract_po_pages_only() strips those pages out of the
    # PO slot, so without this they were simply lost. See
    # extract_drawing_pages_from_po().
    if not drawing_path and raw_po:
        _bundled_drawing, _names_part = extract_drawing_pages_from_po(raw_po, TMP_DIR, pn)
        if _bundled_drawing:
            drawing_path = _bundled_drawing
            print(f"  Drawing: no file in {F_DRAWINGS}/ or {F_DRAWINGS2}/ — using "
                  f"drawing page(s) bundled in {raw_po.name}")
            flags.append(
                f"DRAWING SOURCE: no standalone drawing file was found in "
                f"{F_DRAWINGS}/ or {F_DRAWINGS2}/ — the drawing compiled into this "
                f"package was taken from the drawing page(s) bundled inside "
                f"{raw_po.name}"
                + ("." if _names_part else
                   f", and NONE of those pages' text layer names {pn}, so the "
                   f"automation could not positively confirm the drawing belongs to "
                   f"this part. Verify the compiled drawing is {pn} Rev {rev} before "
                   f"signing.")
            )

    drawing_data = parse_drawing(drawing_path, pn, rev, cache) if drawing_path else {
        "material": "N/A", "finish": "N/A", "inspection_callouts": [],
        "sgt_part_number": pn, "drawing_revision": rev
    }

    # If vendor unknown from PO (e.g. scanned PDF), read from engineer's PRF E4
    if vendor.lower() in ("unknown", ""):
        try:
            wb_check = load_workbook(prf_path, read_only=True, data_only=True)
            prf_vendor = wb_check["Sections 1-2"]["E4"].value
            wb_check.close()
            if prf_vendor:
                vendor = str(prf_vendor).strip()
                print(f"  Vendor from PRF: {vendor}")
        except Exception:
            pass

    # ── Receiving list ────────────────────────────────────────────────────────
    receiving_list = find_receiving_list(po_folder)
    # Extract receiving-list data BEFORE highlighting — the highlight anchor
    # needs the lot number, which only exists after this extraction runs.
    # (2026-07-20 fix: this used to happen in the opposite order, which is why
    # the highlight call below always fell back to searching on the part
    # number, contradicting the "use lot number" rule already documented in
    # WI-7.4.3-01 Step 2 — the lot number simply wasn't available yet.)
    receiving_data = extract_receiving_data(receiving_list, pn) if receiving_list else {}
    receiving_hl = None
    if receiving_list:
        hl_path = TMP_DIR / f"receiving_hl_{pn}.pdf"
        lot_for_hl = str(receiving_data.get("lot_number", "")).strip()
        search_term = lot_for_hl or pn
        # part_number=pn (2026-09-11): lets highlight_receiving_line() retry on
        # the part number when a lot-number search matches 2+ rows and NONE
        # qualifies as "received" (SterilizerCo sterilization round-trip case,
        # verified against PO26-1009's real receiving list — see that
        # function's docstring). Harmless when search_term already IS pn
        # (the function skips the retry when part_number == search_term) and
        # a no-op whenever the ordinary lot-number match already resolves.
        matched = highlight_receiving_line(str(receiving_list), search_term, str(hl_path),
                                            part_number=pn)
        if not matched:
            flags.append(
                f"RECEIVING LIST: search term '{search_term}' "
                f"({'lot number' if lot_for_hl else 'part number — lot extraction failed'}) "
                f"not found in the receiving list's text layer (scanned page or no text "
                f"layer) — the row could not be highlighted. Verify the correct row "
                f"manually before compiling."
            )
        receiving_hl = hl_path

    # ── Packing slip + COC classification from Shipping Docs/ + COCs/ ─────────
    shipping_dir = po_folder / F_SHIPPING
    packing_pdfs = list(shipping_dir.glob("*.pdf")) if shipping_dir.is_dir() else []
    packing_slip_path = next(
        (f for f in packing_pdfs if "packing" in f.name.lower() or "slip" in f.name.lower()),
        None
    )
    # Also check root for packing slip
    if not packing_slip_path:
        packing_slip_path = next(
            (f for f in po_folder.glob("*.pdf")
             if "packing" in f.name.lower() or "slip" in f.name.lower()), None
        )

    # Classify COC folder contents + handle combined PDFs
    classified: dict = {}   # {Path: label}
    if cocs_dir.is_dir():
        for coc_file in cocs_dir.glob("*.pdf"):
            with pdfplumber.open(coc_file) as pdf:
                page_count = len(pdf.pages)

            if page_count > 2:
                # Likely combined — split by page type
                split = split_pdf_by_page_type(coc_file, TMP_DIR)
                for label, paths in split.items():
                    for p in paths:
                        classified[p] = label
            else:
                label = classify_coc(coc_file)
                if not label.endswith("_exclude"):
                    classified[coc_file] = label

    # Certificates the vendor bundled into the packing-slip PDF itself
    # (2026-09-17, PO26-1004/VendorB) — see split_bundled_certificate_pages().
    # Added after the COCs/ pass so a standalone certificate in COCs/ always
    # takes precedence and a bundled duplicate of it is not added twice.
    if packing_slip_path:
        _bundled_types = {t for t in classified.values()}
        for _label, _paths in split_bundled_certificate_pages(
                packing_slip_path, TMP_DIR).items():
            if _label in _bundled_types:
                print(f"  Packing slip: bundled '{_label}' page(s) skipped — a "
                      f"standalone {_label} already exists in {F_COCS}/")
                continue
            for _p in _paths:
                classified[_p] = _label

    # Traveler-vendor rule (see TRAVELER_VENDORS / WI-7.4.3-01 Step 7 §2.4):
    # a "traveler"-named doc from a metal / machining manufacturer (VendorF, VendorG, …)
    # is really that maker's inspection report. Reclassify traveler→mfg_report for
    # any vendor that is NOT a genuine traveler vendor (and not VendorD), so it drives
    # 2.3 and compiles in the manufacturer-inspection-report slot, not the traveler
    # slot, and 2.4 resolves to N/A.
    _vl = vendor.lower().strip()
    if not (any(v in _vl for v in TRAVELER_VENDORS) or any(v in _vl for v in VENDORD_VENDORS)):
        for _p, _t in list(classified.items()):
            if _t == "traveler":
                classified[_p] = "mfg_report"

    # Scope classified COC documents to THIS part on a multi-part PO —
    # see _filter_classified_for_part() docstring (2026-09-11).
    classified = _filter_classified_for_part(classified, pn, all_part_numbers)

    # ── Partial-PO hold: no vendor documentation at all for this lot ─────────
    # (Priority 3 item #1, WI-7.4.3-01 gap analysis 2026-09-11, PO26-1008/VendorD
    # precedent: CAT-XX-65/Lot A084 had no shipping ticket and no COC anywhere
    # in the folder, while its two siblings on the same PO had everything and
    # processed normally. The engineer's instruction was to hold that one
    # line item — no package compiled, no PRF written — while the rest of
    # the PO continued.) This is a real, legitimate business state (the
    # vendor hasn't delivered paperwork yet), not a code/data defect, so it
    # gets a dedicated "held" status rather than the generic "error" bucket.
    #
    # CORRECTED 2026-09-11 (same day as the original implementation): this
    # originally required a TRIPLE absence — classified-COC absent, no
    # part-named packing slip, AND no receiving-list line for the part. That
    # third condition is WRONG and was removed after reading the real
    # process_PO26-1008.py directly (via the device bridge): in the actual
    # CAT-XX-65/A084 precedent, the receiving list DOES carry a line for this
    # lot (595 of 596 received into Quarantine — the engineer's own Section 1
    # already reflects it) — the script's own printed hold reason names only
    # "NO VendorD shipping ticket and NO VendorD Certificate of Compliance," never a
    # missing receiving-list line. Requiring receiving-list absence too would
    # have made this exact real case FAIL to hold (a false negative on the
    # very precedent this feature exists for) — a physical receipt can be
    # logged before the vendor's own shipping ticket/COC paperwork arrives,
    # and that paperwork gap is precisely the traceability problem 2.1/2.2
    # need to catch. Fires now on a DOUBLE absence only:
    #   1. classified is empty AFTER part-scoping above — no COC-type
    #      document (mfg_coc/ster_coc/traveler/mfg_report/other) attributable
    #      to this part, whether because none exists or none names this part
    #      while other parts' documents do (see _filter_classified_for_part);
    #      and
    #   2. no packing-slip/shipping-ticket-like file anywhere (root or
    #      Shipping Docs/) names this part number specifically.
    # Both must hold: a generic, PO-wide packing slip or COC that simply
    # doesn't name any part number stays "unscoped" (see above) and keeps
    # this part off the hold path, matching every single-part PO processed
    # to date. Only when OTHER parts on this PO clearly have named documents
    # and THIS one has none anywhere does this fire — regardless of whether
    # a receiving-list line exists for it.
    if all_part_numbers and len(all_part_numbers) > 1:
        _pn_named_packing = any(
            pn.upper() in f.stem.upper()
            for f in (list(shipping_dir.glob("*.pdf")) if shipping_dir.is_dir() else [])
            + list(po_folder.glob("*.pdf"))
        )
        _has_receiving_line = bool(receiving_data)
        if not classified and not _pn_named_packing:
            return {"status": "held", "message": (
                f"HELD — {pn} Rev {rev}: no shipping ticket and no COC-type document "
                f"could be found anywhere in {F_COCS}/, {F_SHIPPING}/, or the PO root "
                f"that names this part/lot, while other parts on this PO do have "
                f"documents on file"
                + (f" (a receiving-list line for this part WAS found — the physical "
                   f"receipt is logged, but the vendor's own paperwork is still "
                   f"missing)" if _has_receiving_line else
                   f" (no receiving-list line for this part was found either)")
                + f". Per the PO26-1008/VendorD precedent (WI-7.4.3-01 gap analysis "
                f"2026-09-11), this line item is held — no package compiled, no PRF "
                f"written — until vendor documentation arrives. Other ready parts on "
                f"this PO are unaffected and continue processing normally. If "
                f"documentation for this part DOES exist but isn't named with the "
                f"part number, rename it (or process this part manually) and re-run."
            )}

    # ── Data extraction ───────────────────────────────────────────────────────
    # (receiving_data was already extracted above, before highlighting — not repeated here)
    packing_data    = extract_packing_data(packing_slip_path) if packing_slip_path else {}
    mfg_coc_path    = next((p for p, t in classified.items() if t == "mfg_coc"), None)
    mfg_coc_data    = extract_coc_data(mfg_coc_path, "mfg_coc") if mfg_coc_path else {}
    mfg_report_path = next((p for p, t in classified.items() if t == "mfg_report"), None)
    mfg_report_data = extract_coc_data(mfg_report_path, "mfg_report") if mfg_report_path else {}

    # ── PRF Section 1 (read-only) for cross-check — added 2026-07-27 ─────────
    # Same read-only pattern as the vendor-from-E4 lookup above. Read once,
    # here, rather than at the cross_reference() call site, so a read failure
    # degrades to "no PRF comparison this run" (logged, not fatal) instead of
    # blocking the part — this check is a flag, not a gate, per Quality
    # direction (WI-7.4.3-01 2026-07-27 changelog).
    prf_section1 = {}
    try:
        wb_s1 = load_workbook(prf_path, read_only=True, data_only=True)
        s1 = wb_s1["Sections 1-2"]
        prf_section1 = {
            "product_name": s1["M3"].value,
            "product_pn":   s1["M4"].value,
            "revision":     s1["M5"].value,
            "po_number":    s1["E3"].value,
            "qty_received": s1["E6"].value,
            "lot_number":   s1["M6"].value,
        }
        wb_s1.close()
    except Exception as e:
        flags.append(
            f"PRF SECTION 1 READ: could not read Section 1 fields for cross-check "
            f"({e}) — Product Name/P/N/Rev/PO#/Lot#/Qty were NOT compared against "
            f"source documents this run. Verify Section 1 manually before signing."
        )

    # ── Cross-reference ───────────────────────────────────────────────────────
    flags.extend(cross_reference(part_item, packing_data, receiving_data,
                                 mfg_coc_data, drawing_data,
                                 prf_section1=prf_section1,
                                 po_number_expected=po_data.get("po_number", "")))

    # ── Calibration date soft check — added per QE direction ────────────────
    # Separate read from the Section 1 block above since it touches "Section 3"
    # too, not just "Sections 1-2" — kept as its own small workbook open/close
    # so a failure here can never affect the Section 1 cross-check above it,
    # or vice versa.
    flags.extend(check_calibration_date(prf_path))

    # ── Vision-based key-field cross-check (soft flag — DISC-1 / Open Item #12) ─
    # Catches description/vendor/lot/qty mismatches across documents whose text
    # layer was too sparse for cross_reference() above to compare anything at
    # all (see check_key_field_consistency() docstring). raw_po is resolved
    # here (rather than down at doc_order construction, where it used to be
    # computed) so this check has the actual PO PDF to fall back to vision on;
    # the later doc_order block reuses this same variable instead of calling
    # find_po_document() a second time.
    # 2026-09-17: raw_po is now resolved once at the top of this function
    # (from main()'s validated po_doc) — see the po_doc parameter docstring.
    _vision_doc_paths = {
        "PO": raw_po,
        "Packing": packing_slip_path,
        "Receiving": receiving_list,
        "COC": mfg_coc_path,
    }
    _vision_extracted = {
        "PO": {"description": part_item.get("description", ""),
               "vendor": vendor,
               "lot_number": "",
               "qty": part_item.get("qty", "")},
        "Packing": {"description": packing_data.get("description", ""),
                    "vendor": packing_data.get("vendor", ""),
                    "lot_number": packing_data.get("lot_number", ""),
                    "qty": packing_data.get("qty_shipped", "")},
        "Receiving": {"description": receiving_data.get("description", ""),
                      "vendor": "",
                      "lot_number": receiving_data.get("lot_number", ""),
                      "qty": receiving_data.get("qty_received", "")},
        "COC": {"description": "",
                "vendor": "",
                "lot_number": mfg_coc_data.get("lot_number", ""),
                "qty": mfg_coc_data.get("qty", "")},
    }
    flags.extend(check_key_field_consistency(_vision_doc_paths, _vision_extracted))

    # ── Section 1 → Section 2 cross-check ─────────────────────────────────────
    # Read Section 1's native checkboxes (read-only) so Section 2's determination
    # can factor in what the engineer already recorded, per Open Item #4 (CLAUDE.md):
    # e.g. QDA marked YES on Section 1.3 raises the bar for 2.4/2.5 rather than
    # letting a vendor-only heuristic silently resolve them.
    try:
        s1_states = _read_vml_checkbox_states(
            prf_path, "Sections 1-2", required_cells=list(_S1_CHECKBOXES.keys())
        )
    except ChecklistCheckboxError as e:
        return {"status": "error", "message": (
            f"BLOCKED — {prf_path.name}: Section 1 checkbox state could not be verified "
            f"({e}). Open the PRF in Excel, confirm 1.2 (D8/G8) and 1.3 (D13/N13/P13), "
            f"re-save, and re-run. Do not proceed by guessing these values."
        )}
    section1_qda = ("YES" if s1_states.get("N13", {}).get("checked") else
                    "N/A" if s1_states.get("D13", {}).get("checked") else
                    "NO"  if s1_states.get("P13", {}).get("checked") else None)

    # ── Section 2 logic ───────────────────────────────────────────────────────
    section2 = determine_section2(vendor, drawing_data, classified,
                                  packing_data, receiving_data,
                                  product_description=part_item.get("description",""),
                                  mfg_report_data=mfg_report_data,
                                  section1_qda=section1_qda,
                                  part_number=pn)

    # Promote any section2 inline flags into the main flags list
    for key in ("2.1_flag", "2.3_flag", "2.4_flag", "2.5_flag", "disposition_flag"):
        if key in section2:
            flags.append(section2.pop(key))

    # ── Annotate drawing callouts ─────────────────────────────────────────────
    callouts = drawing_data.get("inspection_callouts", [])
    final_drawing = drawing_path
    if drawing_path and callouts:
        final_drawing, missed_callouts = annotate_drawing_callouts(drawing_path, callouts, TMP_DIR)
        if missed_callouts:
            flags.append(
                f"CALLOUT PLACEMENT: Could not locate callout(s) {missed_callouts} in drawing "
                f"text layer (likely scanned or vector-only). Those bubble(s) were skipped — "
                f"add manually before signing."
            )

    # ── Document classification lists (needed by the completeness check next,
    # and reused below for doc_order) ─────────────────────────────────────────
    mfg_coc_inc   = [p for p, t in classified.items() if t == "mfg_coc"]
    ster_coc_inc  = [p for p, t in classified.items() if t == "ster_coc"]

    # Travelers: include for ALL parts EXCEPT Biologic from VendorB/VendorA/VendorC.
    # Non-biologic parts from VendorA/VendorB/VendorC — travelers ARE included.
    confidential_bg = section2.get("is_biologic", False)
    if confidential_bg:
        traveler_inc = []  # Biologic only — confidential, exclude entirely
    else:
        traveler_inc = [p for p, t in classified.items() if t == "traveler"]

    # Other QC docs (seal inspection, final inspection, pick lists, labels etc.)
    # Include for all vendors — sgt_doc_exclude filter handles ACME Medical-generated pages.
    # Drawings in combined docs are handled by extract_po_pages_only + sgt_doc_exclude.
    other_inc = [p for p, t in classified.items() if t == "other"]

    # Manufacturer inspection report — per WI-7.4.3-01 Step 3 DOCUMENT ORDER
    # (position 9, "Manufacturer inspection report / other QC documents") this
    # is supposed to be compiled into the package, but doc_order below never
    # included it — mfg_report_path was extracted (feeds 2.3 in Step 7) and
    # then silently dropped instead of being compiled in. Found 2026-07-20
    # while wiring in the completeness check below; see CLAUDE.md Open Item #7.
    mfg_report_inc = [mfg_report_path] if mfg_report_path else []

    # ── Vendor document completeness check (Open Item #6) ────────────────────
    # Run on every doc that will actually be COMPILED into the package (not on
    # anything excluded, e.g. a confidential Biologic traveler that never
    # made it into traveler_inc) — checks each one, page by page, for a
    # section left blank with no N/A marked. See check_document_completeness()
    # above for what this catches and why it doesn't block the run on failure.
    #
    # 2026-07-22 fix (found on PO26-1003/SR90-1004): this loop used to run
    # AFTER fill_prf() had already written and saved Sections 2/5/6 — meaning
    # any "INCOMPLETE VENDOR DOCUMENT" flag it found could never reach
    # Section 5 or Section 6.1; it only ever showed up in the console
    # summary. Moved before fill_prf() so these flags are actually considered
    # when Section 5/6.1 are populated, which is the entire point of Step 6.5.
    for _doc_path, _doc_label in (
        [(p, "Mfg COC") for p in mfg_coc_inc] +
        [(p, "Sterilization COC") for p in ster_coc_inc] +
        [(p, "Traveler") for p in traveler_inc] +
        [(p, "Mfg Inspection Report") for p in mfg_report_inc] +
        [(p, "Other/QC Document") for p in other_inc]
    ):
        flags.extend(check_document_completeness(_doc_path, _doc_label))

    # ── Discrepancy-relevant subset of flags (2026-07-22 fix) ─────────────────
    # Section 5's narrative and Section 6.1's "discrepancies present" answer
    # must only be driven by flags that describe something actually wrong
    # with the received part/documents — not by pure automation/process
    # limitations. "RECEIVING LIST:" (couldn't auto-highlight a scanned list)
    # and "CALLOUT PLACEMENT:" (couldn't place a drawing bubble) are notices
    # about what the TOOL couldn't do, not findings about the PART, and
    # should never by themselves flip Section 6.1 to "Yes" or draft a
    # discrepancy narrative — found on PO26-1003/SR90-1004, where a
    # highlight-failure notice alone had incorrectly triggered both, requiring
    # QE to manually clear the narrative and reset 6.1 to N/A/No. Every other
    # flag type (cross-reference mismatches, Section 2 FLAG conditions,
    # INCOMPLETE VENDOR DOCUMENT) still counts.
    #
    # "VISION CROSS-CHECK:" added 2026-07-22 alongside the two existing
    # prefixes, per explicit direction from Quality: a description/vendor/
    # lot/qty mismatch read via vision-fallback on a scanned document (Open
    # Item #12 / DISC-1) is a soft, informational flag for human review —
    # it must reach the run summary but must NOT by itself draft a Section 5
    # narrative or set Section 6.1 to "Yes." Unlike the other two prefixes,
    # this one CAN describe something genuinely wrong with the part (a real
    # description mismatch), not just a tooling limitation — it is excluded
    # from discrepancy_flags by deliberate choice, not because it is process-
    # only, so do not fold it back in without Quality re-confirming that.
    # "QTY SHORTAGE (WITHIN 10% TOLERANCE)" added 2026-09-11 (Task #12 /
    # Priority 3 item #2): like VISION CROSS-CHECK, this can describe a real
    # condition on the part (a genuine shortfall), not just a tooling
    # limitation — it is excluded here by deliberate Quality-confirmed
    # threshold decision (10%), not because it is process-only. A shortage
    # beyond 10%, or any overage, is still a plain "QTY MISMATCH" flag and
    # is NOT covered by this prefix, so it still drives Section 5/6.1 as before.
    # "DRAWING SOURCE:" added 2026-09-17 (PO26-1004/VendorB): a note that the
    # drawing compiled into the package came from the PO PDF's own bundled
    # drawing page rather than a standalone file in Drawings/ describes where
    # automation FOUND a document, not a defect in the received part — it
    # belongs in the summary for the engineer to eyeball, never in a Section 5
    # discrepancy narrative or Section 6.1.
    _informational_prefixes = ("RECEIVING LIST:", "CALLOUT PLACEMENT:", "VISION CROSS-CHECK:",
                                "QTY SHORTAGE (WITHIN 10% TOLERANCE)", "DRAWING SOURCE:")
    discrepancy_flags = [f for f in flags if not f.startswith(_informational_prefixes)]

    # ── Fill PRF Sections 2, 5, 6 ────────────────────────────────────────────
    fill_prf(prf_path, section2, discrepancy_flags, drawing_data, s1_states=s1_states)

    # ── Export PRF PDF ────────────────────────────────────────────────────────
    # 2026-09-11 (Task #18 — logo hardening): export_and_stamp_pdf() can now
    # raise LogoMissingError if neither this PRF nor the fixed fallback
    # template has a usable ACME Medical letterhead image in xl/media/ — see
    # prf_helpers.py's _find_logo()/FALLBACK_LOGO_TEMPLATE. Caught here so a
    # missing logo halts only this part (BLOCKED, same shape as every other
    # part-level failure above) instead of raising uncaught and crashing the
    # whole multi-part run.
    try:
        prf_pdf_path = export_and_stamp_pdf(str(prf_path), str(forms_dir), str(prf_path))
    except LogoMissingError as e:
        return {"status": "error", "message": (
            f"BLOCKED — {prf_path.name}: could not export a stamped PDF ({e}). "
            f"Confirm the fallback letterhead template path in prf_helpers.py "
            f"(FALLBACK_LOGO_TEMPLATE) is correct for this machine, or restore "
            f"the ACME Medical logo image in this PRF's own xl/media/, then re-run."
        )}

    # ── Checked boxes must render as CHECKMARKS, never X's ───────────────────
    # LibreOffice draws a CHECKED native form control as an X-in-a-box. With
    # the ☑/☐ text-glyph overlay removed (so each field shows one mark, the
    # template's own), that X is now the only thing a reviewer sees on a
    # populated box — and WI Step 3's CAUTION is explicit that a signed PRF
    # must show checkmarks, never X's. normalize_checkbox_marks() rewrites
    # each X as a clean vector checkmark inside the template's own box,
    # leaving empty boxes and any bitmap-checkmark variant alone.
    #
    # Verified on this PRF's real export at 400 DPI: 13 X-boxes in, 13 vector
    # checkmarks out, 15 empty boxes untouched. Also verified by rendering the
    # compiled output that compile_pdf()'s Ghostscript re-distill PRESERVES
    # these vector checkmarks — the global WI warns it may mangle them and
    # advises a PyMuPDF-only merge; not true of this Ghostscript build, so
    # compile_pdf() is left alone.
    try:
        _marks = normalize_checkbox_marks(prf_pdf_path)
        print(f"  PRF PDF: {_marks} checked box(es) rendered as checkmarks")
        if _marks == 0:
            flags.append(
                "CHECKBOX RENDER: no checked checkbox was found to normalize on the "
                "exported PRF PDF. If any Section 1/2/5/6 answer is checked, confirm "
                "by eye that it shows a checkmark (not an X, and not a blank box) "
                "before routing for signature."
            )
    except Exception as e:
        # Appearance-only post-process — never block the part over it, but say
        # so, because the delivered PDF then shows X's on populated boxes.
        flags.append(
            f"CHECKBOX RENDER: could not normalize the exported PRF PDF's checked "
            f"boxes to checkmarks ({e}) — populated boxes may print as X's. Confirm "
            f"the PRF pages by eye before routing for signature."
        )

    # ── Compile package (SOP 7.4.3 / the Quality Manager's Proof Checklist order) ─────────
    # 1.PO  2.Packing Slip  3.Receiving List  4.Drawing  5.PRF
    # 6.Mfg COC  7.Ster COC  8.Traveler (non-confidential only)
    # (raw_po was already resolved earlier, above the vision cross-check —
    # reused here rather than calling find_po_document() a second time.)
    po_pdfs = [extract_po_pages_only(raw_po, TMP_DIR) if raw_po else None]

    # Packing slip — Shipping Docs copy if found; else first classified only (no duplicates)
    # Both branches now pass through strip_invoice_pages() — see that function's
    # docstring for why this was previously unfiltered.
    #
    # 2026-09-08: also strip from the packing slip any page that duplicates a
    # document type this package is ALREADY compiling separately, in its own
    # documented position (WI-7.4.3-01 Step 3 positions 6-9). Found on
    # PO26-1007/VendorE: the signed packing slip is a 2-page scan whose second
    # page is the manufacturer's COC, and the same COC also sits in COCs/ —
    # so the certificate appeared twice, once right after the packing slip and
    # again after the PRF. Only types actually present in the include lists
    # below are passed, so a certificate that exists ONLY inside the packing
    # slip PDF is never dropped.
    _already_included = set()
    if mfg_coc_inc:    _already_included.add("mfg_coc")
    if ster_coc_inc:   _already_included.add("ster_coc")
    if traveler_inc:   _already_included.add("traveler")
    if mfg_report_inc: _already_included.add("mfg_report")

    if packing_slip_path:
        packing_pdfs_inc = [strip_invoice_pages(packing_slip_path, TMP_DIR,
                                                drop_labels=_already_included)]
    else:
        first_ps = next((p for p, t in classified.items() if t == "packing_slip"), None)
        packing_pdfs_inc = [strip_invoice_pages(first_ps, TMP_DIR,
                                                drop_labels=_already_included)] if first_ps else []

    receiving_inc = [receiving_hl or receiving_list] if (receiving_hl or receiving_list) else []
    drawing_inc   = [final_drawing] if final_drawing else []
    prf_inc       = [prf_pdf_path] if prf_pdf_path else []

    doc_order = (
        [str(p) for p in po_pdfs if p] +
        [str(p) for p in packing_pdfs_inc if p] +
        [str(p) for p in receiving_inc if p] +
        [str(p) for p in drawing_inc if p] +
        [str(p) for p in prf_inc if p] +
        [str(p) for p in mfg_coc_inc] +
        [str(p) for p in ster_coc_inc] +
        [str(p) for p in traveler_inc] +
        [str(p) for p in mfg_report_inc] +
        [str(p) for p in other_inc]
    )

    out_name = f"{po_num} Receiving Part {part_idx} - {pn}.pdf"
    out_path = output_dir / out_name

    # ── No-clobber guard (added 2026-09-11) ──────────────────────────────────
    # find_user_edits()/assert_no_user_edits() have existed in prf_helpers.py
    # since 2026-08-12 (PO26-1011) but were never called anywhere in
    # receive.py — confirmed by direct grep, not inferred. Re-running this
    # script for a part whose compiled package a human already
    # signed/annotated (the engineer types lot release info straight into
    # Section 4 of the COMPILED PDF, and Quality signs that same file — see
    # assert_no_user_edits()'s docstring) would silently overwrite that work
    # with a freshly regenerated file. Detect this BEFORE touching the real
    # path: regenerate to a throwaway probe file first, diff it against
    # whatever already exists at out_path, and only replace out_path if
    # nothing was added since — or the operator explicitly passed --force.
    if out_path.exists():
        _probe_path = TMP_DIR / f"__freshprobe_{out_name}"
        compile_pdf(doc_order, str(_probe_path))
        try:
            assert_no_user_edits(str(out_path), str(_probe_path), force=force)
        except UserContentError as e:
            return {"status": "error", "message": (
                f"BLOCKED — {out_name}: {e} Re-run with --force only on an "
                f"explicit operator instruction to discard those edits."
            )}
        shutil.move(str(_probe_path), str(out_path))
    else:
        compile_pdf(doc_order, str(out_path))
    pages = fitz.open(str(out_path)).page_count

    # ── Step 9 traceability manifest — added 2026-09-17, resolves Open Item #8 ─
    # write_manifest()/LOGS_DIR did not exist anywhere in this codebase before
    # today; no part processed by this tool has ever produced one (see
    # CLAUDE.md Open Item #8's full history, and the 2026-09-17 changelog
    # entry closing it). This covers the one outcome Step 9 is really about —
    # a part that reached a compiled package — not every early-return/BLOCKED
    # path above (PRF not found, checkbox-state error, Section 1 held, logo
    # missing, user-edit conflict): those are tool/process halts before a
    # real quality determination exists, and none of them produced a package
    # this manifest could describe. That's a deliberate v1 scope decision,
    # not an oversight — a held/blocked part leaving no persisted record of
    # WHY is a real, separate gap worth a follow-up, flagged in the
    # changelog rather than solved here.
    #
    # KNOWN GAP this does NOT resolve: Step 9 also wants each extracted fact
    # tagged with which method produced it (pdfplumber vs. haiku vs. sonnet
    # vs. cache hit). That instrumentation doesn't exist anywhere in this
    # script's extraction call sites — adding it would mean threading a
    # tracking parameter through Steps 1/4/5/6, a much larger change than
    # this session's scope. write_manifest() records this honestly as
    # "NOT INSTRUMENTED" rather than fabricating a value — see that
    # function's docstring.
    #
    # A manifest-write failure is logged as a flag, not a hard block: the
    # package above is already correctly compiled and the quality
    # determination is already correct regardless of whether this side
    # record succeeds — same posture as the "PRF SECTION 1 READ" soft-flag
    # failure earlier in this function.
    try:
        _manifest_path = write_manifest(
            po_number=po_num, part_number=pn, wi_revision=WI_REVISION,
            data={
                "vendor": vendor,
                "revision": rev,
                "lot_number": receiving_data.get("lot_number", ""),
                "section1_cross_check": {
                    "note": "WI-7.4.3-01's draft 'Step 2.5' is implemented today as "
                            "a soft flag inside cross_reference() (see that "
                            "function's docstring) — a Section 1 mismatch shows up "
                            "in cross_reference_findings below, not as a separate "
                            "pass/fail gate here.",
                    "section1_qda": section1_qda,
                    "prf_section1_fields": prf_section1,
                },
                "cross_reference_findings": [
                    f for f in flags
                    if f.startswith(("MISMATCH", "REV MISMATCH", "QTY MISMATCH"))
                ],
                "section2_results": section2,
                "all_flags": flags,
                "discrepancy_flags": discrepancy_flags,
                "disposition": section2.get("disposition"),
                "output_package": out_name,
                "prf_xlsx": prf_path.name,
                "prf_pdf": Path(prf_pdf_path).name if prf_pdf_path else "",
                "doc_order": [Path(p).name for p in doc_order],
            })
        manifest_name = _manifest_path.name
    except Exception as e:
        flags.append(
            f"MANIFEST: could not write the Step 9 traceability manifest for this "
            f"part ({e}). The compiled package and PRF above are unaffected and "
            f"correct — only this side record failed. Investigate before relying "
            f"on {LOGS_DIR} for this part's audit trail."
        )
        manifest_name = ""

    return {
        "status": "ok", "part_number": pn, "revision": rev,
        "lot": receiving_data.get("lot_number",""),
        "output": out_name, "pages": pages,
        "section2": section2, "flags": flags,
        "prf_xlsx": prf_path.name,
        "prf_pdf":  Path(prf_pdf_path).name if prf_pdf_path else "",
        "doc_order": [Path(p).name for p in doc_order],
        "manifest": manifest_name,
    }


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # 2026-09-11 fix (CLAUDE.md 2026-09-08 changelog entry, PO26-1007/VendorE):
    # a default Windows console (cp1252) crashes on this script's own ✓/⏳/⚠
    # console output with UnicodeEncodeError, before a single part is
    # processed — not a logic bug, just an un-set console encoding. Force
    # UTF-8 on stdout/stderr at the very top of main(), before any print()
    # call below can trip over it. reconfigure() is a no-op failure mode on
    # a stream that doesn't support it (never raises); PYTHONUTF8=1 or an
    # already-UTF-8 console makes this a harmless no-op either way.
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

    parser = argparse.ArgumentParser(description="ACME Medical Receiving Automation")
    parser.add_argument("po_number", help="PO number, e.g. PO26-1085")
    parser.add_argument("--parts", help="Comma-separated part numbers to process")
    parser.add_argument("--force", action="store_true",
                        help="Discard any human edits already made to a delivered "
                             "compiled package instead of halting that part (see "
                             "assert_no_user_edits in prf_helpers.py). Use only on "
                             "an explicit operator instruction to discard those edits.")
    args = parser.parse_args()

    po_number = args.po_number
    requested = {p.strip().upper() for p in args.parts.split(",")} if args.parts else None
    po_folder = Path.cwd()

    print(f"\n{'='*52}")
    print(f" receive.py v{SCRIPT_VERSION} — {po_number}")
    print(f"{'='*52}")

    # Validate required folders
    for req in [F_FORMS, F_COCS]:
        if not (po_folder / req).is_dir():
            print(f"[ERROR] Required folder missing: {req}/")
            sys.exit(1)

    cache = load_cache()

    # Parse PO
    po_doc = find_po_document(po_folder, po_number)
    if not po_doc:
        print(f"[ERROR] PO document not found for {po_number}")
        sys.exit(1)
    print(f" PO document: {po_doc.name}")

    # ── PO revision-conflict gate (Priority 3 item #3, 2026-09-11) ───────────
    # More than one PO document sitting in the folder is not, by itself, an
    # error (a re-issued PO can legitimately replace an earlier copy nobody
    # removed). It IS an error if two copies disagree on revision or on a
    # line item's ordered quantity, because nothing available here reliably
    # says which one supersedes the other — confirmed on PO26-1005, where
    # two real PO copies disagreed (Rev B/385 ea vs. Rev C/400 ea) and file
    # modified/"render" timestamps did NOT clearly establish which was
    # newer. Picking the alphabetically-higher revision letter is equally
    # unsafe (a mis-filed old copy could still be lettered to sort higher).
    # Same posture as Step 2.5's Section 1 mismatch gate: halt the ENTIRE
    # run for this PO — not just one part — and require the engineer/Quality
    # to resolve which copy governs before re-running. Runs before any
    # part-level processing, because this ambiguity is a property of the PO
    # itself, not of one line item.
    po_doc_candidates = find_po_documents_all(po_folder, po_number)
    if len(po_doc_candidates) > 1:
        _po_datas = [(p, extract_po_data(p)) for p in po_doc_candidates]
        _revisions = {str(d.get("po_revision", "")).strip().upper()
                      for _, d in _po_datas if str(d.get("po_revision", "")).strip()}
        _conflict_lines = []
        if len(_revisions) > 1:
            _conflict_lines.append(
                "PO REVISION CONFLICT: multiple PO copies found with different "
                "revisions — " + " | ".join(
                    f"{p.name}=Rev {d.get('po_revision','?') or '?'}" for p, d in _po_datas)
            )
        _qty_by_part = {}
        for _p, _d in _po_datas:
            for _item in _d.get("line_items", []):
                _pn2 = str(_item.get("part_number", "")).strip().upper()
                _q2  = _item.get("qty", "")
                if _pn2:
                    _qty_by_part.setdefault(_pn2, set()).add((str(_q2), _p.name))
        for _pn2, _vals in _qty_by_part.items():
            _qtys = {v[0] for v in _vals}
            if len(_qtys) > 1:
                _conflict_lines.append(
                    f"PO REVISION CONFLICT: {_pn2} ordered qty disagrees across PO "
                    f"copies — " + " | ".join(f"{fn}={q}" for q, fn in sorted(_vals))
                )
        if _conflict_lines:
            print(f"\n{'='*52}")
            print(f" {po_number} — HALTED: PO revision conflict, no part processed")
            print(f"{'='*52}")
            for line in _conflict_lines:
                print(f"  {line}")
            print(
                f"\n  Multiple copies of this PO's governing document were found "
                f"in {po_folder.name}/ and/or {F_SHIPPING}/, and they disagree on "
                f"revision and/or ordered quantity. Nothing here — including file "
                f"modified dates — reliably establishes which copy is authoritative "
                f"(see WI-7.4.3-01 gap analysis 2026-09-11, PO26-1005). No part on "
                f"this PO was processed. Engineer/Quality must confirm which PO "
                f"revision governs (remove or clearly archive the superseded copy), "
                f"then re-run."
            )
            sys.exit(1)
        # Multiple copies exist but agree on everything compared above — not
        # a conflict; proceed with the one find_po_document() already picked.

    po_data = extract_po_data(po_doc)
    all_parts = po_data.get("line_items", [])
    vendor = po_data.get("vendor", "Unknown")
    print(f" Vendor: {vendor} | Parts on order: {len(all_parts)}")

    forms_dir = po_folder / F_FORMS

    # Fallback: if PO parsing returned no parts (e.g. scanned PDF),
    # derive parts directly from PRF filenames in Inspection Forms/
    if not all_parts:
        print(f"  [WARN] PO parsing returned no parts — scanning {F_FORMS}/ for available PRFs")
        for prf_file in sorted(forms_dir.glob("PRF*.xlsx")):
            name = prf_file.stem  # e.g. "PRF SR90-1006 Rev C"
            if name.upper().startswith("PRF "):
                rest = name[4:].strip()
                rev_split = rest.upper().split(" REV ")
                pn  = rev_split[0].strip()
                rev = rev_split[1].split()[0].strip() if len(rev_split) > 1 else ""
                all_parts.append({
                    "part_number": pn, "revision": rev,
                    "description": "", "qty": 0
                })
                print(f"  Found: {prf_file.name}  ->  Part {pn}  Rev {rev}")
        if not all_parts:
            print(f"[ERROR] No parts found from PO or {F_FORMS}/. Check PO document.")
            sys.exit(1)

    # Scan Inspection Forms/ for available PRFs
    ready, pending = [], []
    for i, item in enumerate(all_parts, start=1):
        pn = item.get("part_number", "")
        if requested and pn.upper() not in requested:
            continue
        prf = find_prf(forms_dir, pn)
        if prf:
            ready.append((i, item))
            print(f"  ✓ {pn} Rev {item.get('revision','')} — PRF ready: {prf.name}")
        else:
            pending.append(pn)
            print(f"  ⏳ {pn} Rev {item.get('revision','')} — PRF not found, pending inspection")

    if not ready:
        print("\n[INFO] No PRFs ready to process.")
        sys.exit(0)

    print(f"\n Processing {len(ready)} of {len(all_parts)} parts...\n")

    # All ready part numbers this run — passed to process_part() so it can
    # scope multi-part-PO documents to the right part and detect a
    # genuinely-undocumented held lot (see _filter_classified_for_part()
    # and the partial-PO hold check, both 2026-09-11).
    all_part_numbers = {item.get("part_number", "").strip().upper() for _, item in ready}

    # Process each ready part
    results = []
    for part_idx, part_item in ready:
        result = process_part(po_folder, po_data, part_item, part_idx, cache,
                               cli_po_number=po_number, force=args.force,
                               all_part_numbers=all_part_numbers,
                               po_doc=po_doc)
        results.append(result)

    # Cleanup _tmp
    try:
        shutil.rmtree(TMP_DIR)
    except Exception:
        pass

    # Summary
    print(f"\n{'='*52}")
    print(f" {po_number} — Complete | Vendor: {vendor}")
    print(f"{'='*52}")

    for r in results:
        if r["status"] == "held":
            # Distinct from ERROR (2026-09-11, Priority 3 item #1): a held
            # part is a normal business state (vendor paperwork not on file
            # yet), not a tooling/data defect — labeling it ERROR would read
            # as something broken that needs a Claude session, when really
            # it just needs the vendor documents.
            print(f"\n  HELD: {r.get('message','')}")
            continue
        if r["status"] != "ok":
            print(f"\n  ERROR: {r.get('message','')}")
            continue
        s2 = r["section2"]
        print(f"\n Part {r.get('part_number')}  Rev {r.get('revision')}  Lot {r.get('lot')}")
        print(f"  {F_OUTPUT}/{r['output']}  ({r['pages']} pages)")
        print(f"  {F_FORMS}/{r['prf_xlsx']}  updated")
        print(f"  {F_FORMS}/{r['prf_pdf']}  exported")
        if r.get('manifest'):
            print(f"  {LOGS_DIR.name}/{r['manifest']}  written")
        else:
            print(f"  Step 9 manifest: NOT written this run — see MANIFEST flag above")
        s2_str = " | ".join(f"2.{k[-1]} {v}" for k, v in s2.items() if k.startswith("2."))
        print(f"  {s2_str}")
        print(f"  Section 3: {'N/A' if s2.get('section3_na') else 'Data present (engineer-filled)'}")
        print(f"  Disposition: {s2.get('disposition','?')}")
        if r["flags"]:
            print(f"  ⚠  FLAGS ({len(r['flags'])}):")
            for f in r["flags"]:
                print(f"     - {f}")
        else:
            print("  Flags: None")
        print(f"  PDF order: " + " | ".join(
            f"{i+1}.{n}" for i, n in enumerate(r["doc_order"])))

    if pending:
        print(f"\n Pending (no PRF yet): {', '.join(pending)}")

    print(f"\n Next steps:")
    print(f"  - Review suggested Sections 2/5/6 — edit freely before signing")
    print(f"  - Resolve any flags before routing")
    print(f"  - Engineer sends compiled PDF for signatures (not automated)")
    print(f"{'='*52}\n")


if __name__ == "__main__":
    main()

