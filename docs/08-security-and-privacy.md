# 08 · Security and privacy

## What this public repo leaves out

This is a sanitized copy of an internal project. A deterministic scrubber was run over every file before publishing, and the results were checked with a denylist search and by reading the files.

| Category | Treatment |
|---|---|
| Company name | → `ACME Medical` / `ACME` |
| Vendors (contract manufacturers, sterilizers, machine shops, distributors) | → `VendorA`…`VendorL`, `SterilizerCo`, `CatalogSupplier` |
| Product names, catalog numbers | → `ProductX`, `ProductS`, `CAT-XX-nn`; "bone graft" → "biologic" |
| People | → roles: *QE*, *Director of Operations*, *Engineering Reviewer*, *Quality Manager* |
| PO, part and lot numbers | → consistent fakes (`PO26-1xxx`, `SR90-xxxx`, `99xxxxx`), so cross-references still line up |
| File paths, Box folder names | → `C:\Users\<you>\Box\Shared\Purchase Orders\…` |
| monday board / item IDs, API tokens | removed |
| ERP product name | → "ERP" |

**Not included at all:** run manifests, case studies, the per-PO edge-case scripts, backups, the PRF Excel template, the approved-destination catalog, validation Word documents, every vendor document, and the plugin's `data.json` (which holds the monday token).

**Kept on purpose:** internal QMS document numbers (SOP 7.4.3, QF 7.4.3-02, WI-7.4.3-01, QF 7.4.3-06), because the docs and code cross-reference them, and the anonymized open-item and test-case *titles*, because they show what validating a tool like this involves.

Screenshots use **synthetic data** produced by `tools/screenshot-harness/demo_data.py`.

## Secrets in the running system

| Secret | Where it lives | Guidance |
|---|---|---|
| `ANTHROPIC_API_KEY` | Windows user environment variable | Never in code, the vault or git. Rotate if exposed |
| monday API token | `.obsidian/plugins/receiving-dashboard/data.json` (plain text, local) | Use the least-privileged account. Keep the vault off shared drives. Excluded by `.gitignore` |
| Box access | Box Drive session of the logged-in user | The tool has no Box credentials of its own |

## Data flows to third parties

- **Anthropic API:** document text and page images from the one PO being processed. Check your organization's policy and data-retention terms. See [Claude API](integrations/claude-api.md#data-handling).
- **monday.com:** the plugin reads board data and, with auto-DHR on, writes a Status value and an update containing the PO number and part numbers.
- Nothing is sent anywhere else. The collector and the vault scripts use only the Python standard library and local files.

## Least privilege, by design

- `receive.py` refuses any path outside its working PO folder and its own tools folder.
- The dashboard and its collector are read-only on Box.
- `refresh_file_map.py` writes only inside `Project Map/`.
- Claude Code sessions inherit the same file-access rule from `CLAUDE.md`.
