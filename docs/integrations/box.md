# Box integration

Box is the **system of record** for purchase-order documents and signed release packages. The automation reaches it through **Box Drive** (the desktop sync client), which presents Box as ordinary Windows paths. There is no Box API key in the tool.

## PO folder convention

```
C:\Users\<you>\Box\Shared\Purchase Orders\
└── 2026\
    ├── 26-1085 - VendorA\                 "<yy-nnnn> - <Vendor>"  (the dashboard parses this)
    │   ├── 26-1085 - VendorA.pdf          PO document (quote pages are ignored)
    │   ├── Shipping Docs\                 packing slip(s), receiving list (QF 7.4.3-01)
    │   ├── COCs\                          manufacturer COCs, sterilization COCs, travelers, inspection reports
    │   ├── Drawings\                      <part#>_Rev_<X>.pdf  (fallback folder name: "Drawings & Stepfiles")
    │   ├── Inspection Forms\              PRF <part#> Rev <X>.xlsx (+ exported .pdf)
    │   ├── In-Progress Receiving\         ← receive.py writes compiled packages here
    │   ├── Completed Receivings\          ← fully signed packages; the release record
    │   └── Archive\ · Invoices\ · Quote\  ignored by the tool
    └── Template\                          ignored
```

The tool identifies subfolders **by content when names vary**. The WI says: list what's actually there, then decide which folder holds shipping documents, drawings and COCs.

## Who touches Box, and how

| Actor | Access | What |
|---|---|---|
| Operations | read/write | scan packing slip, receiving list, COCs, drawings |
| QE (Excel) | read/write | PRF in `Inspection Forms\` |
| `receive.py` | read/write **one PO folder** (cwd) | writes only `In-Progress Receiving\<package>.pdf` and `Inspection Forms\PRF … .xlsx/.pdf` |
| QE (Adobe) | read/write | signs, then moves the package to `Completed Receivings\` |
| `refresh_dashboard.py` | **read-only, every PO folder** | lists folders, reads filed PDFs to count signatures |
| Claude (Cowork / claude.ai) via the Box connector | per user permissions | search, read drawings and docs during development and debugging |

## Rules that came from audits of the process

- **No new folders in Box.** Run manifests, caches, backups and the vault stay in the local tools folder. Engineering agreed that the Box PO tree contains only the documents SOP 7.4.3 defines.
- **Signed off = in `Completed Receiving(s)`.** The dashboard and the monday DHR automation both rely on this one observable fact rather than on someone remembering to update a status.
- **Box Drive download cost.** The collector caches signature counts by file size and modified time, so each filed package is downloaded only once.
- **Sibling isolation.** The tool runs with cwd = one PO folder and refuses any resolved path outside it. That stops one PO's documents leaking into another's package, and is covered by its own test case.

## Using the Box connector with Claude

With the Box connector enabled in Claude (claude.ai / Cowork), a session can:

- find a drawing or COC by part number or lot across folders (`search_files_keyword`, `search_files_metadata`)
- read a document's text to check an extraction (`get_file_content`)
- list a PO folder to confirm what Operations filed before a run (`list_folder_content_by_folder_id`)

Use it for **investigation and development**. Production runs use Box Drive paths, so the validated tool has no network dependency on Box.
