---
tags: [governance, top-level]
---

# WI-7.4.3-01 — Receiving Inspection Automation Work Instruction

The controlled instruction document that governs the entire project. Lives at the vault root as [[CLAUDE.md]] (same file — CLAUDE.md *is* WI-7.4.3-01; Claude Code loads it automatically as project instructions). Currently **Revision 02 (Development)** — not yet promoted to Rev A. Refs: SOP 7.4.3 Rev L, QF 7.4.3-02 Rev E (the PRF template), the Quality Manager's Proof Receiving Checklist.

## What it makes Claude/the tool do

> Read the engineer's pre-filled PRFs, validate Section 1 against source documents, classify and compile all part documents into one PDF per part, and fill PRF Sections 2, 5, and 6 using fixed rules — never a judgment call.

## The 9-step workflow (+ sub-steps)

1. [[Step 1 - Verify Folders and Identify Parts]]
2. [[Step 2 - Match & Classify Documents]] — now also handles per-part document scoping and the partial-PO hold
3. [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)]] — read-only, blocking
4. [[Step 3 - Compile the Receiving Package]]
5. [[Step 4 - Extract Data from Documents]]
6. [[Step 5 - Parse Drawing with Cache]]
7. [[Step 6 - Cross-Reference Source Documents]] — now also the 10% shortage tolerance and PO-revision-conflict halt
8. [[Step 6.5 - Vendor Document Completeness Check]] — vision judgment call, non-blocking
9. [[Step 6.6 - Vision-Based Key-Field Cross-Check]] — vision judgment call, soft flag
10. [[Step 7 - Section 2 Determination (Fixed Rules)]] — deterministic, never model judgment; 2.4 default and disposition rule both rewritten 2026-09-11, disposition corrected again 2026-09-14 (VendorD anodized/laser-etched screws → Finished Goods)
11. [[Step 8 - Fill PRF Sections 2, 5, 6]] — the only sheets/cells ever written; native-checkbox integrity added 2026-09-11
12. [[Step 9 - Run Manifest (Desktop Dev Aid)]] — ~~documented but does not exist in code, reconfirmed 2026-09-08~~ *2026-10-05: implemented 2026-09-17; Desktop `logs\` only; development aid, not a quality record*

## Supporting governance rules

- [[File Access Rules]] — **rewritten 2026-09-11**: the real two-tree sandbox (`Desktop\Receiving Task Automation Tools\` + the one active PO folder), not the fictional year-folder structure
- [[Folder Structure]] — **rewritten 2026-09-11**, same reason
- [[Manual vs Claude-Assisted Workflow]] — **updated 2026-09-11**: real commands, the UTF-8 console fix, edge-case-script guidance
- [[Comment & Discrepancy Reasoning Reference]] — **new 2026-10-05**: Section 5 / run-summary language drawn from signed precedent; Section 4 stays engineer-only

## Implemented by

- [[receive.py]] (main script, ~3,600 lines) importing from [[prf_helpers.py]] (~1,805 lines)
- Governed dependency requirement: prf_helpers.py must exist alongside receive.py at `Desktop\Receiving Task Automation Tools\` — receive.py imports from it, never reimplements its functions inline. CLAUDE.md's own `REQUIRED DEPENDENCY` import list was itself found stale and corrected 2026-09-11 (it named `LOGS_DIR`/`write_manifest`, which don't exist, and omitted several functions in active use).

## Tracked / validated by

- [[QF 7.4.3-06 Validation Protocol]] (Rev 05) — the controlled validation record for this Tool; its 9.1–9.11 numbering has not been re-verified against CLAUDE.md's current, much longer Open Items list
- [[Open Items (Master List)]] — **24 open items** as of 2026-09-11 (was 12 + 1 proposed), none yet resolved to Rev A
- [[Attachment A - Test Case Matrix]], [[Attachment B - Instruction Guide]]

## Key design principle (repeated throughout the WI)

Section 2 (1.2–2.6) determinations are made **exclusively by fixed Python functions** ([[Step 7 - Section 2 Determination (Fixed Rules)]]), never by model judgment. Anywhere a required fact is missing or unreadable, the correct behavior is to **flag and halt/leave blank — never guess**. This shows up repeatedly: [[Step 2.5 - Validate PRF Section 1 (Blocking Gate)]] halts on mismatch rather than guessing which value is right; the checkbox-reading gate (see [[prf_helpers.py]]) raises `ChecklistCheckboxError` rather than guessing a cell; [[Step 7 - Section 2 Determination (Fixed Rules)]] returns `"FLAG"` rather than defaulting an answer. The 2026-09-11 batch extended this same principle twice more: the PO-revision-conflict gate halts the whole run rather than guessing which PO copy governs (deliberately stricter than the real bespoke script's own corroborate-and-flag practice — see [[Open Items (Master List)#Item 19]]), and the new `"held"` part status leaves a part fully unprocessed rather than guessing at missing vendor paperwork.

## See also

- [[Changelog Timeline]] for how this document has evolved (now covers 2026-07-14 through 2026-09-11)
- [[Real-World Runs (Case Studies)]] for the actual PO runs that drove most of these rules (now 8 POs, not 3)
