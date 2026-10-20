---
tags: [governance, workflow]
---

# Manual vs. Claude-Assisted Runs — Decision Guide

Part of [[WI-7.4.3-01 Overview]]. Added 2026-07-22. [[receive.py]] is a standalone CLI script — it does **not** require Claude Code or any AI session to run. Claude has mainly been used to *debug the tool itself*, not because routine processing needs an AI in the loop.

**Commands corrected 2026-09-11** — see [[Folder Structure]]: the tool lives
directly at `Desktop\Receiving Task Automation Tools\`, not under a
year-rooted `_Automation\tools\`.

## One-time setup (per machine)

1. **Anthropic API key** — the script calls Claude models directly (Steps 4/5/6.5/6.6). Get one at console.anthropic.com → API Keys, set persistently:
   ```powershell
   [System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-your-key-here", "User")
   ```
   Treat like a password. Every Haiku/Sonnet call bills against it — check usage periodically once this runs routinely/unattended.
2. `pip install -r requirements.txt` (see [[Tools and Dependencies]])
3. Ghostscript installed separately (see [[Tools and Dependencies]])

## Running it manually (from inside the PO folder or with a full path, PowerShell)

```powershell
python "C:\Users\<you>\Desktop\Receiving Task Automation Tools\receive.py" PO26-0085
python "C:\Users\<you>\Desktop\Receiving Task Automation Tools\receive.py" PO26-0085 --parts SR25-1202,SR90-1001
```

Identical command and output to what a Claude Code session runs — no behavior difference. (Adding the tools folder to PATH, documented in `requirements.txt`, lets you drop the full path.)

## Windows console crash — fixed 2026-09-11

A default Windows console (cp1252) used to crash on this script's own ✓/⏳/⚠
console output with `UnicodeEncodeError`, before a single part was
processed — found 2026-09-08 on PO26-1007/VendorE, worked around that day with
`PYTHONUTF8=1`, then actually fixed in code 2026-09-11:
`main()` now force-reconfigures stdout/stderr to UTF-8 at the top, before
any print() can trip over it. A fresh machine no longer needs the
environment-variable workaround.

## When to just run it vs. bring it to Claude

| Situation | Action |
|---|---|
| Routine PO, PRF(s) already filled | Run it yourself, no AI needed |
| Completes with flags (2.1 INCOMPLETE, missing highlight, blank lot, `HELD`) | Still just you — the tool correctly declining to guess, resolve by hand |
| Raises an exception / `ChecklistCheckboxError` / `UserContentError` / blocks a part entirely | Bring to Claude — something structural needs diagnosis |
| New vendor / unfamiliar layout | Run once yourself; escalate only if output looks *wrong*, not just flagged |
| Want the tool's logic itself changed | Always Claude — [[Step 7 - Section 2 Determination (Fixed Rules)|Section 2 logic]] is controlled, deterministic code, never hand-edited ad hoc |
| A PO's circumstances don't fit the general path at all | See [[Edge-Case Scripts]] — bespoke per-PO scripts are now an accepted, deliberate pattern (2026-09-11 scope decision: "fix root causes only" — a bespoke script handles the truly one-off part of a real PO; only the general, recurring parts of what it worked around get folded into `receive.py`/`prf_helpers.py`) |

## Unresolved regardless of path

~~[[Step 9 - Run Manifest (Desktop Dev Aid)|The manifest gap]] applies identically to both paths — a manual run has even less of an incidental trail than a Claude-run one (no chat transcript at all).~~

*2026-10-05: both paths now write the same manifest to the Desktop `logs\` folder (since 2026-09-17). It is a development aid, not a quality record. Note: a manual run's Section 5 text does not yet follow the [[Comment & Discrepancy Reasoning Reference]] — see [[Open Items (Master List)#Item 31]].*
