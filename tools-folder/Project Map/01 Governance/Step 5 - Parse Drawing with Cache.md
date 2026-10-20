---
tags: [governance, step]
---

# Step 5 — Parse CAD Drawing with Cache

Part of [[WI-7.4.3-01 Overview]].

## What happens

Check `_Automation\tools\_cache\parts_cache.json` for key `"{part#}_Rev{revision}"`. If found, use it (zero-token cache hit). If not, extract via `claude-sonnet-4-6` (`parse_drawing()` in [[receive.py]]):

- ACME part#, description, drawing revision
- Supplier, supplier part#, supplier revision
- Material callout (or N/A), finish callout (or N/A)
- General tolerances from title block
- All numbered inspection callouts: number, description, nominal, upper/lower limit, gaging method

If no numbered callouts exist → [[Step 7 - Section 2 Determination (Fixed Rules)|Section 3 is N/A]].

## Cache rule

The cache lives **only** in `_Automation\tools\_cache\` — never inside a PO folder (see [[File Access Rules]]). Once engineering numbers critical dimensions on controlled drawings at release, callout extraction becomes a cache lookup only, not a fresh model call.

## Feeds into

- [[Step 6 - Cross-Reference Source Documents]] (drawing revision/supplier-part# cross-checks)
- [[Step 7 - Section 2 Determination (Fixed Rules)]] (2.2 material/finish, 2.3 callout pass/fail)
- Drawing callout bubbling (`annotate_drawing_callouts()`) in [[Step 3 - Compile the Receiving Package]]
