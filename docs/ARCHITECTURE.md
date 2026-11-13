# Architecture

## Pipeline

```
CSV / JSON input ──▶ load_input() ──▶ qc_checks() ──▶ generate_report() ──▶ .xlsx
     (pandas)          (findings)        (openpyxl workbook)
```

1. **Load** (`load_input` in `src/generate_report.py`): pandas reads the
   CSV/JSON, normalises headers, coerces `Units` / `Unit Price` to numbers,
   and derives `Amount = Units × Unit Price`.
2. **QC** (`qc_checks`): scans the frame for missing values, exact duplicate
   rows, `Units <= 0`, negative amounts, and amounts above 100,000. Findings
   carry the input row numbers so problems can be traced back.
3. **Build** (`generate_report`): creates the workbook with openpyxl —
   `Report`, `Summary`, and `QC Checks` sheets — and saves it.

## Sheets

- **Report** — title rows, styled header, banded data rows, frozen panes,
  an auto-filter, a totals row using live `SUM` formulas, and two
  conditional-formatting rules on the Amount column.
- **Summary** — `groupby("Category")` totals, a grand-total row, and an
  openpyxl `BarChart` anchored below the table.
- **QC Checks** — one row per finding (or a single "All checks passed" row).

## Demo data

`--demo` (and `assets/sample_sales.csv`) use a 20-row fictional October 2026
sales dataset for a network-installation supplier — hardware and services
across Noida, Delhi, and Gurugram. No real company data is involved.

## Extending it

- Add a column: extend `REQUIRED_COLUMNS` / `ALL_COLUMNS` and the width map
  in `_add_report_sheet`.
- Add a QC rule: append a finding dict in `qc_checks` — it appears on the
  QC sheet automatically.
- Change the "unusually large" threshold (currently 100,000) in `qc_checks`.
