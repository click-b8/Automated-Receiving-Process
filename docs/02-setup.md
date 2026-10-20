# 02 · Setup

Windows 10/11 is the reference platform: Box Drive, Adobe signing, and Excel for the native checkboxes. The Python parts also run on macOS and Linux.

## 1. The tools folder

Copy `tools-folder/` to the QE's machine. In production it lives at `Desktop\Receiving Task Automation Tools\`. That one folder is both:

- **the tool** (`receive.py`, `prf_helpers.py`, `CLAUDE.md`, caches, `logs\`)
- **the Obsidian vault** (`Project Map\`, `.obsidian\`)

Keep it **outside Box**. Run manifests and caches are development aids and are deliberately kept out of the quality-record tree.

## 2. Python and packages

```powershell
python --version                 # 3.10 or newer
pip install -r requirements.txt  # anthropic, pdfplumber, openpyxl, pymupdf
```

## 3. System dependencies

| Tool | Why | Install |
|---|---|---|
| **LibreOffice** | Headless export of the filled PRF `.xlsx` to PDF (`export_and_stamp_pdf()`) | libreoffice.org, or `winget install TheDocumentFoundation.LibreOffice` |
| **Ghostscript** | Re-distills the merged package so it's editable and signable (removes encryption, security flags and AcroForm locks from vendor PDFs) | ghostscript.com. The tool looks on PATH and in a local `gs10.x\` folder next to it |

## 4. Anthropic API key

```powershell
[System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
```

Treat it like a password. Every classification, extraction and vision call bills against it. See [integrations/claude-api.md](integrations/claude-api.md) for which models are used where and how to keep cost down.

## 5. Company files you must supply

These are controlled company documents and are **not** in this repo.

| File | Where | Used for |
|---|---|---|
| PRF template (QF 7.4.3-02) `.xlsx` with native checkboxes and the letterhead logo | each PO's `Inspection Forms\` (the QE starts each PRF from it); a fallback copy in `templates\` | Fill target. The fallback provides the logo if a PRF has lost it |
| Approved Product Destination List catalog `qf_7.2-01_apdl_catalog.json` | tools folder root | Disposition rule (Raw vs Finished Goods). If it's missing the tool warns and defaults every part to Raw Goods |
| Your Section 2 vendor rules | constants at the top of `receive.py` (`BIOLOGIC_VENDORS`, `TRAVELER_VENDORS`, `ALWAYS_STERILE_VENDORS`, `VENDORD_VENDORS`…) | Placeholders here (`vendora`, `vendorb`…). Replace them with your real vendor names, lowercase |

The PRF cell map (which cell is 2.3 "No", where Section 5 text goes, and so on) is hard-coded for QF 7.4.3-02 Rev E. If your form differs, see [04 · Pipeline → PRF cell map](04-pipeline-steps.md#prf-cell-map).

## 6. Box

Install **Box Drive** so PO folders appear as normal paths, for example:

```
C:\Users\<you>\Box\Shared\Purchase Orders\2026\26-1085 - VendorA\
```

See [integrations/box.md](integrations/box.md) for the required PO subfolders.

## 7. Optional: run from anywhere

Add the tools folder to `PATH`, or use the wrapper:

```bat
:: receive.bat
python "%~dp0receive.py" %*
```

## 8. Obsidian

1. Install Obsidian, then **Open folder as vault** → choose the tools folder.
2. Settings → Community plugins → turn off Restricted mode → enable **Receiving Dashboard**. It's already in `.obsidian/plugins/receiving-dashboard/`.
3. Click the **gauge** icon in the left ribbon, or run *Open Receiving Dashboard* from the command palette.
4. Settings → Receiving Dashboard:
   - **Python command**: usually `python`
   - **PO year folder**: leave blank for the default; it is read-only
   - **Refresh every N minutes**: how often the dashboard rescans while Obsidian is open
   - **monday.com**: see [integrations/monday.md](integrations/monday.md)
5. Optional background refresh, even when Obsidian is closed:

   ```bat
   "Project Map\_scripts\install_dashboard_schedule.bat"
   ```

   This creates a Windows scheduled task, *Receiving Dashboard Refresh*, that runs `refresh_all.py` every 15 minutes with `pythonw` (no window). Remove it with `schtasks /Delete /TN "Receiving Dashboard Refresh" /F`.

## 9. Optional: Claude Code / Cowork

`CLAUDE.md` is written as Claude Code project instructions. Open Claude Code in a **PO folder** that contains a copy of the WI, and ask "process PO26-1085". It follows the same steps and file-access rules and runs the same script. Engineers copy the canonical WI into each PO folder rather than editing it locally.

Day to day, the AI session is mostly used to **debug and improve the tool**, maintain the vault and draft validation documents. Routine POs are just `python receive.py`.
