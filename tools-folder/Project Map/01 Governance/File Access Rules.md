---
tags: [governance, sandbox]
---

# File Access Rules

Part of [[WI-7.4.3-01 Overview]].

**REWRITTEN 2026-09-11 to match [[CLAUDE.md]]'s own 2026-09-08 correction.**
Everything below this notice describes the REAL structure, confirmed
directly against the live Desktop folder via the device bridge
(2026-09-11). The version of this page before today described a
`_Automation\tools\` / `_Automation\logs\` year-folder structure inside Box.
**That structure does not exist and, as far as anyone has been able to
determine, never has.** See [[Open Items (Master List)#Item 2]] and
[[Open Items (Master List)#Item 8]] for how this was discovered, and
[[Changelog Timeline#2026-09-08]] for the correction itself.

## The real working universe — exactly TWO directory trees

1. **`Desktop\Receiving Task Automation Tools\`** — the actual home of
   `receive.py`, `prf_helpers.py`, [[CLAUDE.md]] itself, `templates\`, this
   vault (`Project Map\`), `Validation\`, every per-PO
   [[Edge-Case Scripts|edge-case script]], and the tool's own working files
   (`parts_cache.json`, `__pycache__\`). Read/write freely inside this tree
   — it's the tool's real project directory, not a read-only controlled
   copy.
2. **The ONE PO subfolder matching the PO# given in the current command**,
   under `...\Purchase Orders\<year>\<PO folder>\` — nothing
   else in that Purchase Orders tree.

## Hard rules, no exceptions

- Never read/write inside a sibling PO folder, even to peek, and never a
  different year's folder for the same vendor.
- Never browse the rest of `Purchase Orders\<year>\` beyond
  the one PO subfolder in scope (that directory also holds `Template\`,
  `Unstored COC's\`, and a `.claude\` settings folder — none in scope for a
  normal PO run).
- Never access Desktop/Documents/Downloads or any other system folder
  EXCEPT the one automation-tools folder named above.
- Anything outside these two trees requires the user's explicit,
  in-the-moment permission first — never rely on auto-mode to silently
  approve an out-of-scope path.
- If an expected file is missing, STOP and say exactly which file and which
  of the two allowed locations it belongs in.

## What changed from the old (fictional) model

| Old page said | Real practice |
|---|---|
| receive.py/prf_helpers.py live at `_Automation_Master\tools\`, READ ONLY to Claude Code | They live at `Desktop\Receiving Task Automation Tools\`, read/write |
| A `_cache\` subfolder under `_Automation\tools\` | `parts_cache.json` sits directly in the Desktop tools folder |
| A `_Automation\logs\` folder holds traceability manifests | ~~No manifest is written anywhere — see [[Step 9 - Run Manifest (Desktop Dev Aid)]]~~ *2026-10-05: manifests are written to `Desktop\Receiving Task Automation Tools\logs\` only (since 2026-09-17), never in Box — see [[Step 9 - Run Manifest (Desktop Dev Aid)]]* |
| Year-rooted `_Automation\` copy, refreshed yearly from `_Automation_Master\` | No such copy exists; one flat tools folder serves every year/PO |

## Reasoning Reference harvest — added 2026-10-05

The only routine activity that reads more than one PO folder. It is never
part of a PO run, needs QE's explicit in-the-moment permission naming the
POs each time, reads **only** their `Completed Receiving\` subfolders, and
writes nothing anywhere in Box. See
[[Comment & Discrepancy Reasoning Reference]]; validated by TC-28.

## Consistency check (unchanged)

Every PO# encodes its year (PO26-xxxx = 2026). Before processing, confirm
the PO# given matches the year folder being opened — stop and flag if not.

## Still-open question

Whether `Desktop\Receiving Task Automation Tools\` — living on one
engineer's own machine, not a shared/controlled Document Control location —
is itself an acceptable home for a controlled WI's tooling, or whether
Quality wants an actual `_Automation_Master`-style controlled deployment
built as originally specified. See [[Open Items (Master List)#Item 2]].

## See also

- [[Folder Structure]] for the full real layout (also rewritten today)
- [[Open Items (Master List)#Item 8]] — ~~the traceability-manifest gap this
  correction surfaced~~ *(2026-10-05: manifest built 2026-09-17 and its location decided; two sub-gaps remain)*
