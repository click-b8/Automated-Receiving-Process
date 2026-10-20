---
tags: [history, changelog]
---

# Changelog Timeline

> **Public template.** The real vault keeps a dated, entry-by-entry history of every fix (which PO surfaced it, what changed in code, which test case covers it). That detail is internal and has been removed here. Keep your own entries in this shape:
>
> `### YYYY-MM-DD — short title` → *found on* (PO / part, anonymized), *symptom*, *root cause*, *fix* (function names), *verified by* (test case ID), *docs updated* (WI section, matrix row).

## Milestones (high level)

| When | Milestone |
|---|---|
| 2026-05 | First version of the controlled instruction (WI) and the `receive.py` CLI |
| 2026-06 | Test Case Matrix started; first real POs compiled end-to-end |
| 2026-07 | Section 2 rules hardened (traveler / confidential-record logic), receiving-list highlight fixes |
| 2026-08 | Native Excel checkbox integrity rework (`restore_form_controls()`) |
| 2026-09 | Validation batch: partial-PO hold, PO-revision halt, APDL disposition lookup, run manifests |
| 2026-10 | Obsidian Receiving Dashboard, monday.com DHR automation and delay tracking |

Related: [[Real-World Runs (Case Studies)]], [[Open Items (Master List)]], [[_Hub - History]].

## Recurring lesson across this timeline

Nearly every multi-day bug chain in this history shares one shape: a fix is made, verified against an **isolated copy** or a **synthetic fixture**, reported as done — and then a real run against the actual file surfaces a gap the isolated copy couldn't have shown (different import scoping, a merged-cell geometry the fixture didn't reproduce, a cell that's actually a static label). The standing rule this produced: **any fix verified via an isolated copy also needs one direct smoke-test against the real, unmodified file before being reported as done.**

**2026-09-11 added a new variant of the same lesson.** Two of that day's
fixes (stadium/oval callouts, sterilization round-trip disambiguation) were
first written from a *description* of what a bespoke script did (the gap
analysis document), not from reading that script's actual source — and in
both cases the description turned out to not match the real mechanism
closely enough to trust. Both were caught and corrected the same day, only
after the real bespoke scripts were staged and read directly via the device
bridge. Extended rule: **when a fix is meant to codify what a bespoke
script already does, read that script's actual source before writing the
fix — a summary of its behavior is not a substitute**, the same way an
isolated test copy is not a substitute for the real file.
