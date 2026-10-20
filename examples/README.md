# examples/

Everything here is **synthetic**. The vendors, parts, lots and people are invented.

| File | What it is |
|---|---|
| `example-run-manifest.json` | Shape of one Step 9 run manifest (`logs\<PO>_<part>_<stamp>_manifest.json`) |
| `demo-dashboard-data.json` | Shape of `Project Map/_scripts/.dashboard_data.json`, written by `refresh_dashboard.py` and read by the plugin |
| `demo-monday-response.json` | Shape of the single monday GraphQL response the plugin requests (items, activity logs, updates) |
| `demo-monday-dhr-log.jsonl` | Shape of `.monday_dhr_log.jsonl`, the audit trail of automatic DHR changes |

To try the dashboard without Box or monday, copy `demo-dashboard-data.json` to `tools-folder/Project Map/_scripts/.dashboard_data.json`, open the vault in Obsidian and enable the plugin. The collector will overwrite the file on its next run unless you point **PO year folder** at an empty folder.
