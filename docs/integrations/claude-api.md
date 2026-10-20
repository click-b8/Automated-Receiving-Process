# Claude API integration

`receive.py` uses the Anthropic Python SDK (`anthropic>=0.40`) for the parts of receiving that need **reading**: messy PDFs, scans, page classification, drawing title blocks. It **never** uses a model to make a quality decision.

## Model tiering

```python
MODEL = {
    "classify":  "claude-haiku-4-5-20251001",   # page/document classification, vision reads
    "extract":   "claude-haiku-4-5-20251001",   # field extraction when pdfplumber comes up short
    "checklist": "claude-sonnet-4-6",           # (default for call_claude)
    "narrative": "claude-sonnet-4-6",           # Section 5 discrepancy draft
    "complex":   "claude-sonnet-4-6",           # drawing parse (title block + callouts)
}
```

| Where | Function | Model | Input | Output |
|---|---|---|---|---|
| Step 2 | `classify_page()`, `classify_coc()` (only if the rules can't decide) | Haiku | page text | label: `packing_slip`, `mfg_coc`, `ster_coc`, `traveler`, `mfg_report`, `exclude`… |
| Step 2 | `_classify_scanned_combined_via_vision()`, `_vision_classify_page()` | Haiku (vision) | rendered page PNGs | per-page labels, so a scanned multi-cert bundle can be split |
| Step 4 | `extract_po_data()`, `extract_packing_data()`, `extract_receiving_data()`, `extract_coc_data()` | Haiku | text, or a page image for scans | strict JSON (`_parse_json`) |
| Step 4 | `_vision_extract_packing_fields()` | Haiku (vision) | packing-slip image | Ops signature present? date? lot? |
| Step 5 | `parse_drawing()` | Sonnet | drawing text / pages | part#, rev, supplier, material, finish, numbered callouts. **Cached** by part+rev |
| Step 6.5 | `check_document_completeness()` | Haiku (vision) | every page to be compiled | sections left completely blank |
| Step 6.6 | `_vision_extract_key_fields()` | Haiku (vision) | page 1 of empty-text docs | description, vendor, lot, qty → soft cross-check |
| Step 8 | Section 5 draft | Sonnet | `discrepancy_flags` + facts | plain-text discrepancy statement, no NCR or disposition language |

## Guardrails

1. **pdfplumber first.** Text extraction costs no tokens. A model is called only when the text layer is empty or a required field is missing.
2. **Deterministic decisions.** Section 2 (2.1–2.6), the disposition, cross-reference comparisons, the quantity tolerance and the PO-revision halt are plain Python. The model's output is *data* fed into those rules.
3. **Structured output.** Prompts ask for JSON with fixed keys. `_parse_json` / `_parse_json_dict` reject anything else, and a failed parse becomes "not extracted", which leads to a flag, never a guess.
4. **Vision findings are soft.** Step 6.6 results carry a `VISION CROSS-CHECK:` prefix and can't on their own put text in Section 5 or set 6.1 = Yes.
5. **Fail safe.** API errors in the supplementary checks (6.5/6.6) degrade to "no finding". In extraction they produce an empty field, and the cross-check flags it.
6. **Lazy client.** `_get_client()` creates the client on first use, so tooling that only imports the module (tests, the dashboard) doesn't need a key.

## Cost and latency

- `parts_cache.json` makes the drawing parse, the most expensive call, a one-time cost per part revision.
- Most text-native POs use only a handful of Haiku calls. Scanned vendor paperwork is what drives vision calls.
- Watch usage in the Anthropic Console once runs are routine. Every call bills against `ANTHROPIC_API_KEY`.

## Data handling

Vendor certificates, packing slips and drawings are sent to the API. Before using this in production:

- confirm your organization's policy on sending controlled drawings and supplier documents to a third-party API, and whether your agreement includes zero data retention
- keep the key in the user environment (`setx` / `SetEnvironmentVariable`), never in code or the vault
- remember the run manifest records **which file** fed each fact, but not (yet) **which method** (pdfplumber vs. Haiku vs. cache). That's an open item in the WI, deliberately not faked

## Claude Code and Cowork (the development loop)

Separate from the runtime API calls, Claude is also the **development partner** for the tool:

- `CLAUDE.md` (WI-7.4.3-01) is Claude Code's project instruction file, with the same rules the code implements and the same file-access boundary
- in Cowork, Claude debugs real runs, writes the patch and the matching test-case row, updates the vault notes, keeps the canvases and graph in sync, and drafts validation documents
- the standing lesson from that loop, recorded in the vault, is that **any fix verified on an isolated copy also needs one smoke test against the real, unmodified file** before it's reported as done
