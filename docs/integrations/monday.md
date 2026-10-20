# monday.com integration

monday.com is where Operations, Engineering and Quality track each PO through receiving. The dashboard plugin talks to the **monday GraphQL API v2** (`https://api.monday.com/v2`, `API-Version: 2024-10`) directly from Obsidian using a personal API token.

## Board shape the plugin expects

One board, for example *Receiving – New Inspection*. One **item per PO** (or per PO shipment), named so the PO number can be parsed: `PO: 26-0085 VendorA - (Screws)`.

| Groups (typical flow) | |
|---|---|
| New inspection → Inspection → **Product Release** → **DHR** → Put away (and MRB for holds) | |

| Column (title must match) | Type | Used for |
|---|---|---|
| **Status** | status | `Ready to start`, `Working on it`, `Stuck`, `Waiting for Signature`, `DHR`, `Done` |
| Person | people | shown in the table |
| Date | date | shown in the table |
| Notes / Text | text | free text. `Text` also feeds delay inference |
| **Waiting on** | text/dropdown | who a *Waiting for Signature* item is waiting on (Director of Operations, Engineering Reviewer, Quality Manager…) |
| **Delay reason** | dropdown | one or more of: Missing documents · Material / finish discrepancy · Dimensional out of spec · Quantity discrepancy · Testing failure / root cause · Vendor document error · PRF correction · Signature logistics · Inspection method question · Priority / expedite · Other |

The plugin reads, in **one GraphQL query**: the board's name, URL and groups; up to 500 items with every column value; the last 1,000 **Status** changes from `activity_logs`; and the last 300 **updates** (`text_body`).

## Read side: what you see

- **monday tab.** Group tiles and a searchable item table, with PO chips that link to the Box-side status.
- **POs & parts tab.** A *monday* column showing each PO's current status or group.
- **Delays tab.** Hold episodes rebuilt from the status history. See [06 · Dashboard → Delays](../06-receiving-dashboard.md#delays).

## Write side: automatic DHR

When a PO is complete in Box, the plugin moves its monday item to **DHR** (Device History Record) so nobody has to remember to.

**Rule (`dhrCandidates()`):** for every PO where the latest Box scan (no older than 30 minutes) shows

1. status **filed**: every compiled part's package is in `Completed Receivings`
2. **no PRF waiting** to be compiled
3. **exactly one** monday item matches the PO, and it
   - is in group **Product Release**
   - has status *Working on it*, *Waiting for Signature*, *Ready to start* or blank
   - isn't a per-shipment item (name contains "ship…")
   - was created **before** the last package was signed off

…then it runs two mutations:

```graphql
mutation ($b: ID!, $i: ID!, $c: String!, $v: String) {
  change_simple_column_value(board_id: $b, item_id: $i, column_id: $c, value: $v) { id }
}   # value = "DHR"

mutation ($i: ID!, $body: String!) {
  create_update(item_id: $i, body: $body) { id }
}   # "Status set to DHR automatically by the Receiving Dashboard. All 3 compiled part(s) for PO26-0085
    #  are signed off in Completed Receivings in Box: SR90-1101, SR90-1102, SR90-1103. Box scan 2026-10-20 10:30."
```

Your own monday automation (for example *when Status changes to DHR, move item to group DHR*) does the rest.

Anything that doesn't qualify is listed under **"Signed off in Box but not changed automatically"** with the reason: several items match, item is in another group, status is Stuck, item was created after sign-off… A person decides those.

Every change is appended to `Project Map/_scripts/.monday_dhr_log.jsonl`:

```json
{"time":"2026-10-19T04:30:00Z","po":"PO26-1262","item":"PO: 26-1262 Wingtip","itemId":"…","from":"Waiting for Signature","parts":["SR90-1752"],"account":"Demo User","result":"DHR set"}
```

## STUCK notices → delay analytics

When an item goes **Stuck**, the team's notice, sent by email through Outlook and recorded on the monday item as an update, has the shape:

```
STUCK - PO26-0085
Issue: Emailing vendor for COC, no packing slip in box - Auto
```

`parseStuckIssues()` pulls the `Issue:` line. `categorizeIssue()` maps it to a delay reason with keyword rules, so old holds without a *Delay reason* still group correctly, shown as "(inferred)". See [Outlook](outlook.md).

## Setup

1. monday → avatar → **Developers → My access tokens** → copy the token. Use an account with only the board access it needs. Changes appear under that account's name.
2. Obsidian → Settings → Receiving Dashboard → **Show monday tab**, paste the **API token** and the **Board ID** (the number in the board URL).
3. Leave **Set DHR automatically** off until you've watched a few cycles in the *not changed automatically* list.

> **Security:** Obsidian stores the token in plain text in `.obsidian/plugins/receiving-dashboard/data.json`. Keep that file off shared drives and out of git (this repo's `.gitignore` excludes it). Rotate the token if it was ever copied anywhere else.
