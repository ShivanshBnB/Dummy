# Post-wall QC automation tracker

A workbook that maps the 55 post-wall quality checks a site engineer performs to the
cheapest capture method that could automate each one, with columns to track the work.

| File | What it is |
|---|---|
| `qc_automation_tracker.xlsx` | The workbook. Open it in Excel or upload it to Google Drive to use it as a Sheet. |
| `build_workbook.py` | Rebuilds the workbook from the CSV. Edit the `DESIGN` and `METHODS` tables in it to change the proposals. |
| `data/qc_post_wall_source.csv` | The QC sheet export the workbook was built from (Sheet2, 55 rows). |

## Tabs

- **Read me**: how to use the workbook, the proposed pipeline, what one viewpoint can and cannot measure, what changes for the site engineer, and a suggested order of work.
- **QC Tracker**: one row per check. Source columns (trimmed), proposed design (check family, primary method, automation level, confidence, how it works, fallback, inputs the prompt needs, target tolerance, shared detector, extra site time) and tracking columns (status, owner, sprint, validation accuracy and sample, last updated, notes).
- **Methods**: the ten method codes, what each asks of the site engineer, extra seconds per use (editable estimates), expected accuracy, build effort and limits.
- **Dashboard**: counts by status, method, automation level, confidence, check family and stage. Formulas only.
- **Experiments**: a log of what was tried and decided.
- **Source data**: the export as received.

## Rebuild

```bash
python3 -I build_workbook.py data/qc_post_wall_source.csv qc_automation_tracker.xlsx
```

Requires `openpyxl`. After rebuilding, open the file once in Excel or LibreOffice so the formulas calculate.
