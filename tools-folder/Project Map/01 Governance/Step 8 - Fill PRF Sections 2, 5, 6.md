---
tags: [governance, step, write-boundary]
---

# Step 8 — Fill PRF Sections 2, 5, and 6 Only

Part of [[WI-7.4.3-01 Overview]]. **The write boundary**: automation writes ONLY to Sections 2, 5, and 6, every run, no exceptions. Sections 1, 3, 4 are read-only — the cell *value* is never modified.

## The one rendering-only exception

Converting every sheet's native VML checkboxes to ☑/☐ **text** before PDF export (LibreOffice cannot render native VML form controls — without this, checked boxes render blank in the PDF). This redraws an existing state, never changes it, and applies to Sections 1-2, Section 3, and Sections 4-6 alike (generalized 2026-07-20 after Section 3/4 checkboxes were found rendering blank the same way Section 1 once did).

## Native checkbox integrity + no-clobber guard — NEW 2026-09-11 (the biggest fix in this WI's history)

Root cause: `save_preserving_drawings()` carried the underlying VML/ctrlProp
zip parts forward on save, but the regenerated sheet XML no longer
referenced them — Excel silently prunes every unreferenced zip part the
NEXT time a human opens and saves the file. Confirmed catastrophic on
PO26-1008/PO26-1001: **native checkboxes AND the ACME Medical letterhead logo
were both permanently lost**, with no error at the time it happened. Fixed
with four new [[prf_helpers.py]] functions — `restore_form_controls`,
`set_native_checkboxes`, `verify_native_checkboxes`,
`verify_xlsx_package_integrity` — that keep the real, clickable native
controls intact through a save, not just the rendered glyph. Paired with a
**no-clobber guard** (`UserContentError`/`assert_no_user_edits`, plus a new
`--force` CLI flag) so the tool never silently overwrites a PRF a human has
since hand-edited. See [[Changelog Timeline#2026-09-11]].

## LogoMissingError — part-level block, not a crash

`export_and_stamp_pdf()` now raises `LogoMissingError` when neither the
PRF's own letterhead image nor the fallback template
(`templates\PRF template CONTROLS.xlsx` — filename corrected 2026-09-11,
was previously a guessed name that doesn't exist) has a usable logo.
`process_part()` catches this and blocks just that one part with a clear
message, instead of crashing the whole PO run.

## Checkbox map (Sections 1-2 sheet)

```
2.1 → I20 (Completed, only if not FLAG)
2.2 → I21=Yes | L21=No | O21=N/A
2.3 → I22=Yes | L22=No | O22=N/A
2.4 → I23=Yes | L23=No | O23=N/A
2.5 → I24=Yes | L24=No | O24=N/A
2.6 → I25=Yes | L25=No (no N/A)
```

Plus, since 2026-09-08 (explicit QE instruction): H29 (Quality Manager
initials, active only if 2.4=Yes) and I32/O32 (Director signature+date,
active only if 2.6=Yes) are now actively re-asserted to "N/A" whenever
their governing item isn't "Yes" — previously these only ever held
whatever the blank template shipped with.

## Section 4 — engineer-owned, reconfirmed 2026-10-05

Section 4 is the engineer's own comment field. Automation and Claude
sessions read it (Step 2.5 cross-check, DVT reference, and as vocabulary for
the [[Comment & Discrepancy Reasoning Reference]]) but never write to it,
pre-fill it, suggest text into it, or tidy it.

## Section 5 language — added 2026-10-05

The B14 text must follow the [[Comment & Discrepancy Reasoning Reference]]:
plain text, `discrepancy_flags` only, facts with sources, no NCR language,
signed precedent where one fits. `receive.py`'s own prompt doesn't apply
this yet — [[Open Items (Master List)#Item 31]]; validated by TC-27.

## Sections 4-6 sheet — CORRECTED 2026-09-11 (was stale in this WI's own body text for six weeks)

```
Section 5: L12 = N/A checkbox | B14 = discrepancy text (sonnet-drafted)
Section 6.1: K30=Yes | N30=No
6.2 N/A: B31 = '6.2\n☑ N/A'
Disposition: J36=Raw Goods | N36=Finished Goods (or both unchecked — APDL hit, see Step 7)
```

**This page previously (and [[CLAUDE.md]]'s own STEP 8 section, until
today) showed `I30=Yes | M30=No` and `I36=Raw Goods | M36=Finished Goods`.**
Those are the OLD, wrong cells — already found and fixed in the real code
on 2026-07-22 ([[Open Items (Master List)#Item 9]]), then confirmed
against a real blank template on 2026-07-27 (which also caught a second bug:
K30/N30/J36/N36 carry their label text baked into the SAME cell as the
checkbox glyph, requiring `with_label=` on every call —
[[Open Items (Master List)#Item 9b]]). CLAUDE.md's own procedural text
simply never got updated to match its own Open Items list and Changelog —
caught and corrected 2026-09-11 while syncing this vault, not from a new
code change.

## `discrepancy_flags` — the filtered subset that drives Section 5/6.1

Section 5's narrative and Section 6.1 must only be driven by flags describing something wrong with the **part**, not the tool's own process limitations. `RECEIVING LIST:`, `CALLOUT PLACEMENT:`, `VISION CROSS-CHECK:`, and, since 2026-09-11, `QTY SHORTAGE (WITHIN 10% TOLERANCE)` are all excluded from `discrepancy_flags` by deliberate design.

## Disabled, not fixed

The 2.2 material/finish annotation (used to overwrite cell B21, which is actually the row's own "2.2" label) is **disabled outright**, not routed to a guessed replacement cell — see [[Open Items (Master List)#Item 10]].

## Save/export

`save_preserving_drawings(wb, prf_path, prf_path)` then `export_and_stamp_pdf()` — both in [[prf_helpers.py]], both hardened 2026-09-11.

## Feeds into

- [[Step 3 - Compile the Receiving Package]] (the exported PRF PDF slots in at position 5)
- [[Step 9 - Run Manifest (Desktop Dev Aid)]] ~~(documented as the next step; still does not actually run)~~ *(2026-10-05: runs since 2026-09-17; Desktop `logs\` only, development aid)*
