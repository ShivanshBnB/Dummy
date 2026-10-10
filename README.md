# Post-wall QC automation tracker

A workbook that maps the post-wall quality checks a site engineer performs (80 at present) to the
cheapest capture method that could automate each one, with columns to track the work.

| File | What it is |
|---|---|
| `qc_automation_tracker.xlsx` | The workbook. Open it in Excel or upload it to Google Drive to use it as a Sheet. |
| `build_workbook.py` | Rebuilds the workbook from the CSV. Edit the `DESIGN` and `METHODS` tables in it to change the proposals. |
| `data/qc_post_wall_source.csv` | The QC sheet export the workbook was built from (Sheet2, 55 rows). |
| `data/qc_post_wall_added_YYYY-MM-DD.csv` | Further checks, as headerless rows in the same column order. Passed to the script as `--extra`; rows whose first cell is not a number are skipped. |
| `data/qc_tracker_export_YYYY-MM-DD.csv` | The QC Tracker tab exported from the team's working copy. Passed to the script as `--overlay`, its SE, approach and tracking columns override the script's values. |
| `data/approach_split.csv` | ID, Approach group and Approach description only: a block to paste next to the free-text approach column in the working copy. |

## Tabs

- **Read me**: how to use the workbook, the proposed pipeline, what one viewpoint can and cannot measure, what changes for the site engineer, and a suggested order of work.
- **QC Tracker**: one row per check. Source columns (trimmed), the site engineer's notes (SE Comments, Accuracy Requirements SE, Feasibility with Gyro), the team's approach (group and description), proposed design (check family, primary method, automation level, confidence, how it works, ARCore relevance, fallback, inputs the prompt needs, target tolerance, shared detector, extra site time) and tracking columns (status, owner, sprint, validation accuracy and sample, last updated, notes).
- **Methods**: the eleven method codes, what each asks of the site engineer, extra seconds per use (editable estimates), expected accuracy, build effort and limits.
- **Dashboard**: counts by status, method, automation level, confidence, check family and stage. Formulas only.
- **Experiments**: a log of what was tried and decided.
- **Source data**: the exports as received, with the file each row came from.

## Rebuild

```bash
python3 -I build_workbook.py data/qc_post_wall_source.csv qc_automation_tracker.xlsx \
    --extra data/qc_post_wall_added_2026-10-10.csv \
    --overlay data/qc_tracker_export_2026-10-08.csv
```

To add checks, drop them in a headerless CSV in the source column order and pass it as another `--extra`, then add their design to the `DESIGN`, `ARCORE` and `APPROACH` tables in the script (the build stops and names any ID that is missing). To pick up new edits from the working copy, export its QC Tracker tab as CSV into `data/` and pass that file as `--overlay`. Blank cells in the export keep the script's value; the `Extra SE time (s)` column is always regenerated as a formula.

Requires `openpyxl`. After rebuilding, open the file once in Excel or LibreOffice so the formulas calculate.
