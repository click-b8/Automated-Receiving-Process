# 01 · Architecture

## The problem in one picture

A receiving record has a lot of sources and one output:

```
PO ─┐
Packing slip ─┤                                  ┌─ PRF (QF 7.4.3-02), Sections 2/5/6 pre-filled
Receiving list (QF 7.4.3-01) ─┤   receive.py      │
Drawing (+ older rev) ─┼──►  classify ► extract ──┼─ Compiled package PDF, fixed order, one per part
Manufacturer COC(s) ─┤   cross-check ► rules      │
Sterilization COC ─┤                              └─ Run manifest (dev aid, local only)
Traveler / mfg report ─┘
```

Every source has to agree on **part number, revision, lot, quantity, PO number and dates**. Section 2 of the PRF records six yes/no/N-A answers derived from those documents (the Quality Manager's "Proof Receiving Checklist"). The compiled PDF has to follow the order in SOP 7.4.3.

## Components

| Component | Language | Runs where | Writes to |
|---|---|---|---|
| `receive.py` | Python 3.10+ | QE workstation, **cwd = one PO folder** | that PO folder's `Inspection Forms\` and `In-Progress Receiving\`; tools-folder `logs\`, `parts_cache.json`, `_tmp\` |
| `prf_helpers.py` | Python | imported by `receive.py` | (library: PDF merge/highlight, xlsx save with native checkboxes, PDF export) |
| `CLAUDE.md` = WI-7.4.3-01 | Markdown | read by Claude Code on launch; read by people | n/a |
| Obsidian vault `Project Map/` | Markdown + canvas | Obsidian desktop | notes only |
| `refresh_dashboard.py` | Python, stdlib only | Windows scheduled task every 15 min, or on demand | `Project Map/_scripts/.dashboard_data.json` and its cache |
| `refresh_file_map.py` | Python, stdlib only | same | `Project Map/` hub notes |
| Receiving Dashboard plugin | Plain JS (no build) | inside Obsidian | `.monday_dhr_log.jsonl`; monday.com (Status column + an update) when auto-DHR is on |

## Data flow, end to end

1. **Dock → Box.** Operations receives the pallet, signs the packing slip, fills the receiving list, receives the lot into Quarantine in the ERP, then scans everything into the PO folder in Box (`Shipping Docs\`, `COCs\`, `Drawings\`).
2. **Inspection.** The QE inspects, measures and tests, then fills PRF **Section 1** (part info, quantity inspected, QDA), **Section 3** (measurements) and **Section 4** (comments). The PRF is saved to `Inspection Forms\`.
3. **Compile.** From inside the PO folder the QE runs `python receive.py PO26-xxxx`. For each part the tool:
   - lists the folder and identifies parts from the PO (Step 1)
   - classifies every document and scopes it to the part (Step 2)
   - **blocks** if PRF Section 1 disagrees with the sources (Step 2.5)
   - extracts fields with pdfplumber first and Claude only when needed (Steps 4/5)
   - cross-checks sources against each other (Step 6), checks vendor documents are complete (6.5) and cross-checks key fields by vision (6.6)
   - decides Section 2 with **fixed rules** (Step 7)
   - writes Sections 2/5/6, keeps native Excel checkboxes intact and exports the PRF PDF (Step 8)
   - merges the package in SOP order and highlights the receiving-list row (Step 3)
   - writes a run manifest to the local `logs\` folder (Step 9)
4. **Review & sign.** The QE opens the package, resolves every flag, edits the Section 5 draft if needed and signs in Adobe.
5. **Route.** Outlook email to the signers. Order: Director of Operations first if 2.6 = Yes, Engineering Reviewer before the Quality Manager if Section 6 has discrepancies, Quality Manager last.
6. **File.** The fully signed package goes to the PO's `Completed Receivings\`. That move is what the dashboard treats as "signed off".
7. **Track.** The dashboard collector scans Box read-only, counts digital signatures in filed packages and joins the result with the monday.com board. When every compiled part of a PO is signed off it can set the monday item to **DHR**.

## Trust boundaries

```
┌──────────── QE workstation ────────────────────────────────────────────┐
│  tools folder (Desktop)            Box Drive sync (C:\Users\<you>\Box) │
│  ┌──────────────────────┐   cwd    ┌─────────────────────────────┐     │
│  │ receive.py           │ ───────► │ ONE PO folder (read/write)  │     │
│  │ refuses paths outside│          │ sibling POs: never touched  │     │
│  │ cwd (+ tools folder) │          └─────────────────────────────┘     │
│  └──────────┬───────────┘          ┌─────────────────────────────┐     │
│             │                      │ all PO folders (READ-ONLY)  │◄─── dashboard collector
│             │ HTTPS                └─────────────────────────────┘     │
└─────────────┼──────────────────────────────────────────────────────────┘
              ▼                                   ▲ HTTPS (token)
      Anthropic API (document text/images)    monday.com API (board read, Status write)
```

- **The file-access rule is part of the controlled WI.** The tool works only inside the folder it was launched in, plus its own tools folder. It never climbs to a parent folder or searches Desktop or Downloads. If a file it expects is missing, it stops and says which file and where it should go. See `CLAUDE.md` → *FILE ACCESS RULES*.
- **The dashboard is read-only on Box.** It only lists folders and reads PDFs to count `/ByteRange` signature dictionaries.
- **The model never makes a quality decision.** Claude classifies pages, reads fields from scans and flags inconsistencies. Section 2 answers and the disposition come from `determine_section2()` and plain `if` statements.

## Why a CLI and not a service?

- The validated scope stays small: one script, one folder, one run, one manifest.
- It works offline from monday and Outlook. Box Drive gives it ordinary file paths.
- Claude Code is optional. The same command runs from PowerShell with no AI session (see [Using receive.py](03-using-receive-py.md#manual-vs-claude-assisted-runs)).

## Why Obsidian?

The project grew to roughly 6,700 lines of code, a 3,500-line work instruction, a validation package and dozens of real-run lessons. Obsidian keeps all of it in plain Markdown **in the same folder as the tool**, so:

- the WI, the code notes and the validation status are cross-linked (`[[Step 7 - …]]`), and Graph view shows what depends on what
- canvases give two visual maps: the project map and the physical QE workflow
- a small plugin turns the folder into a live operations dashboard without a server

See [05 · Obsidian vault](05-obsidian-vault.md) and [06 · Receiving Dashboard](06-receiving-dashboard.md).
