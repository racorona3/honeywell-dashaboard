# Honeywell Program Dashboard — Refresh Script

This repository contains a small Python script, `refresh_dashboard.py`, that reads two sheets from an Excel workbook, computes YTD revenue, margin and past-due analytics, and writes a self-contained `index.html` dashboard (uses Chart.js). The generated `index.html` can be published via GitHub Pages or uploaded manually.

## Requirements

- Python 3.8+
- pip packages:
  - `pandas`
  - `openpyxl`

Install dependencies:
`python -m pip install pandas openpyxl`

## Files

- `refresh_dashboard.py` — main script (reads Excel, computes metrics, writes `index.html`)
- `index.html` — generated dashboard (not committed by default; commit only if you want GitHub Pages to serve it)

## Quick start (manual)

1. Close the Excel workbook before running the script (Excel may lock the file).
2. Edit the CONFIG block at the top of `refresh_dashboard.py` if you need to change:
   - `EXCEL_FILE` — input Excel workbook path
   - `SAW_SHEET` — sheet name for current SAW rows
   - `HIST_SHEET` — sheet name for historical snapshots
   - `OUTPUT_FILE` — where to write `index.html`
3. Run:
`python "C:\path\to\refresh_dashboard.py"`
4. Upload or commit the generated `index.html` to the branch served by GitHub Pages (or otherwise host the file).

## What the script does

- Loads two Excel sheets:
  - Current SAW data (rows with `Status`, `PO/Bin`, `Extended Price`, `ActionBy - New`, `Customer Name`, etc.)
  - Historical snapshots with `Date`, `Revenue ($)`, `PGM (%)`, `Baseline Plan (Revenue)`, etc.
- Computes metrics: YTD revenue, program gross margin (PGM), past-due dollars/lines, backlog, top function and site callouts, plus trend series.
- Produces a single-file `index.html` with embedded Chart.js charts and visualizations.

## Configuration notes

- The script currently uses hard-coded file paths in the CONFIG block. For portability, consider converting to CLI args (see Suggested improvements).
- Confirm that your Excel columns match the names used in the script (case-sensitive): `Date`, `Revenue ($)`, `PGM (%)`, `Baseline Plan (Revenue)`, `Baseline Plan PGM`, `Past Due Dollars ($)`, `Past Due Lines`, `Extended Price`, `Status`, `PO/Bin`, `ActionBy - New`, `Customer Name`.

## Troubleshooting

- "Excel file not found" — verify the `EXCEL_FILE` path and that the file is closed.
- `KeyError` or missing column errors — open the Excel workbook and confirm exact column names and sheet names.
- Empty historical sheet — the script expects at least one historical row; if empty it may fail. Add at least one snapshot or add error handling.
- Write errors for `index.html` — check write permissions and that the output directory exists.
- Incorrect numeric formats in charts — validate whether revenue columns are in dollars (not thousands) and whether `PGM (%)` is fractional (0.25) or already a percent (25). Adjust code accordingly.

## Security & data handling

- Do not commit the source Excel workbook if it contains sensitive/proprietary data.
- The script embeds some text values into JavaScript. If names can contain arbitrary characters, consider using `json.dumps(...)` for safe JS embedding (prevents accidental injection or broken JS).

## Suggested improvements (optional)

- Convert the script to accept CLI arguments with `argparse` (input/output paths, `--dry-run`, `--debug`).
- Replace ad-hoc string escaping with `json.dumps(...)` when embedding Python lists/strings in the HTML/JS to avoid manual escaping bugs.
- Validate required columns on load and print a helpful message listing missing columns.
- Add a `--sample` mode that writes a small example Excel for testing.
- Add unit tests around the `compute()` function using pytest and small DataFrame fixtures.
- Add logging and an optional `--verbose`/`--debug` flag.

## Example automation (CI / scheduled)

Simple automation approach:
1. Run `python refresh_dashboard.py` on a scheduled runner or local machine.
2. Commit and push the updated `index.html` to the branch that GitHub Pages serves.
3. GitHub Pages will redeploy the updated static site.

## Contact / Next steps

If you want, I can:
- produce a patched `refresh_dashboard.py` that uses `argparse` and `json.dumps` for safer embedding,
- add basic column validation and helpful error messages,
- or generate a small sample Excel and pytest-based unit tests for `compute()`.

Tell me which improvements you want and I will generate the changes.

