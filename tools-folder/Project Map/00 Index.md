---
tags: [moc, index]
---

# Receiving Automation Project — Map of Content

Entry point for the whole project. See also [[Project Map.canvas]] for a single visual board with every node below and the dependency lines between them (canvas itself not re-laid-out in this pass — see the note at the end of this page).

This vault documents **ACME Medical's Receiving Inspection Automation project**: a controlled work instruction ([[WI-7.4.3-01 Overview|WI-7.4.3-01]], = the vault-root [[CLAUDE.md]] file), a Python automation tool ([[receive.py]] + [[prf_helpers.py]]) that reads engineer-completed PRFs, validates them against source documents, compiles per-part receiving packages, and fills fixed Section 2/5/6 checklist fields — and the formal validation package (QF 7.4.3-06) that governs its rollout.

**This vault was refreshed 2026-09-11 after sitting untouched since
2026-07-22 (VP Rev 05).** Everything below reflects current status; where a
page changed meaningfully in this pass, it says so at the top.

## Dashboard (added 2026-10-12)

The **Receiving Dashboard** is an Obsidian plugin, not a note: click the gauge icon in the left ribbon, or run "Open Receiving Dashboard" from the command palette. Tabs: Overview (live counts and charts), POs & parts (open packages, PRFs and PO folders), Workflow (W01–W10 with live counts), Test matrix, Project files (every file in this folder, searchable, opens for editing) and monday (optional). It re-scans the Box PO folder every 5 minutes while Obsidian is open; settings are under Settings → Receiving Dashboard. Data comes from `Project Map/_scripts/refresh_dashboard.py` (read-only on Box).

## Cluster hubs (added 2026-10-12)

One hub per colour in Graph view. Each hub links its own pages and links back here, so each colour clumps together and the clusters connect through this Index. Rebuilt by `Project Map/_scripts/refresh_file_map.py` — run it after a receiving run to pick up new notes and files.

- [[_Hub - Instruction (WI-7.4.3-01)]] — purple
- [[_Hub - Code]] — cyan
- [[_Hub - Validation]] — green
- [[_Hub - History]] — yellow
- [[QE Receiving Workflow]] — orange
- [[_Hub - Files]] — grey: manifests (teal), backups (slate), edge-case scripts (red), validation documents (dark green), tool files

## 01 · Governance — the controlled work instruction

- [[WI-7.4.3-01 Overview]] — top-level summary, start here
- [[Step 1 - Verify Folders and Identify Parts]]
- [[Step 2 - Match & Classify Documents]] — updated: per-part document scoping, partial-PO hold
- [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)]]
- [[Step 3 - Compile the Receiving Package]]
- [[Step 4 - Extract Data from Documents]]
- [[Step 5 - Parse Drawing with Cache]]
- [[Step 6 - Cross-Reference Source Documents]] — updated: 10% shortage tolerance, PO-revision-conflict halt
- [[Step 6.5 - Vendor Document Completeness Check]]
- [[Step 6.6 - Vision-Based Key-Field Cross-Check]]
- [[Step 7 - Section 2 Determination (Fixed Rules)]] — updated: 2.4 default, disposition rule rewritten 2026-09-11, disposition corrected again 2026-09-14
- [[Step 8 - Fill PRF Sections 2, 5, 6]] — updated: native-checkbox integrity, corrected cell map
- [[Step 9 - Run Manifest (Desktop Dev Aid)]] — ~~gap, blocks Rev A, still unresolved~~ *obsoleted 2026-10-05: implemented 2026-09-17; Desktop `logs\` only, development aid, not a quality record*
- [[Comment & Discrepancy Reasoning Reference]] — **new 2026-10-05**: how Section 5 language improves from signed precedent; Section 4 stays engineer-only
- [[File Access Rules]] — rewritten: real Desktop-folder structure, not the fictional year-rooted one
- [[Folder Structure]] — rewritten, same reason
- [[Manual vs Claude-Assisted Workflow]] — updated: real commands, UTF-8 fix, edge-case-script guidance

## 02 · Code — the actual scripts

- [[receive.py]] — main automation, ~3,600 lines now (was ~2,490), full function map updated
- [[prf_helpers.py]] — shared utilities, ~1,805 lines now (was ~580), full function map updated
- [[Edge-Case Scripts]] — ⚠️ now ~14 undocumented per-PO scripts (was 2) — audit observation, larger than before
- [[Tools and Dependencies]] — Python packages, Ghostscript, LibreOffice, API key

## 03 · Validation — the controlled validation package (QF 7.4.3-06)

- [[QF 7.4.3-06 Validation Protocol]] — Rev 05 (not re-verified against CLAUDE.md's now-24-item Open Items list in this pass — see [[Open Items (Master List)]]'s own caveat)
- [[Attachment A - Test Case Matrix]] — updated: TC-05 filled, TC-19–26 added, 3 new notes
- [[Attachment B - Instruction Guide]] — Rev 05, plain-language quick start (not touched this pass)
- [[Open Items (Master List)]] — now **24** tracked items (was 12 + 1 proposed) — fully resynced

## 05 · QE Workflow — the physical process (new 2026-10-05)

- [[QE Receiving Workflow]] — pallet → paperwork → ERP → staged → inspect / test → compile → review & sign → email for signatures → Completed Receiving → inventory, with the automation embedded in each step
- [[QE Receiving Workflow.canvas]] — the same thing as a visual board, one row per physical step
- Graph view: open the local graph of [[QE Receiving Workflow]], or filter the global graph with `path:"Project Map/05 QE Workflow"`. Graph colors: orange workflow, purple WI/CLAUDE.md, cyan code, green validation, yellow history.

## 04 · History — how we got here

- [[Changelog Timeline]] — now covers 2026-07-14 through 2026-09-11 (was stopping at 2026-07-22)
- [[Real-World Runs (Case Studies)]] — now 8 POs (was 3): the original PO26-1002/0356/0266, plus PO26-1006, PO26-1007, and the 2026-09-11 batch (PO26-1015, PO26-1008, PO26-1009/0430, PO26-1001)

## Quick answers

- **"Is this validated yet?"** No — [[WI-7.4.3-01 Overview]] is Revision 02 (Development), not Rev A. See [[Open Items (Master List)]] for what's blocking it (now 24 items, not 12).
- **"What's the single biggest open gap?"** ~~[[Step 9 - Run Manifest (Desktop Dev Aid)|The traceability manifest]] — documented as existing since Rev 01, still never actually implemented, confirmed again 2026-09-08.~~ *Obsoleted 2026-10-05: the manifest exists (2026-09-17) and its location is decided. Not re-ranked in this pass; see [[Open Items (Master List)]].*
- **"What runs outside the controlled path entirely?"** [[Edge-Case Scripts]] — now ~14 scripts, up from 2. Quality made a scope decision 2026-09-11 to fix root causes rather than adopt a review procedure — worth Quality's attention regardless.
- **"Can I just run this myself without Claude?"** Yes — see [[Manual vs Claude-Assisted Workflow]] (commands corrected 2026-09-11 to the real Desktop path).
- **"What's the biggest fix in this project's history?"** Native checkbox integrity + no-clobber guard (2026-09-11) — an ordinary Excel save on an orphaned-controls PRF was permanently destroying both the checkboxes and the ACME Medical logo with no error. See [[Step 8 - Fill PRF Sections 2, 5, 6]].
- **"Where did the fictional folder structure come from, and is it really gone?"** It was always aspirational, never real — confirmed 2026-09-08 and reconfirmed via direct device listing 2026-09-11. See [[File Access Rules]] and [[Folder Structure]] (both rewritten). [[Attachment A - Test Case Matrix|TC-17/TC-18]] still test the fictional structure and need a Quality decision.

## ~~Known limitation of this refresh pass~~ — obsoleted 2026-10-05

~~`Project Map.canvas` (the visual board) was **not** re-laid-out to add~~
~~nodes for new pages or the larger cross-reference set — it still reflects~~
~~the pre-2026-09-11 page set. The markdown pages above are the authoritative,~~
~~current source; treat the canvas as a stale visual index until it's~~
~~rebuilt.~~

*Obsoleted 2026-10-05: the canvas was re-laid-out into five clusters — the WI/CLAUDE.md instruction cluster (joined to the automation by two bridge edges), code, validation, history, and the new QE workflow. Every earlier node was kept. Six cross-cluster edges were taken off the board to keep it readable and are listed in its "Cross-cluster links" card. The markdown pages remain the authoritative source.*
