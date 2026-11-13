# Excel Report Generator

Turn raw CSV or JSON data into a formatted Excel workbook — styled tables, a
summary sheet with a chart, and a QC sheet that flags data-quality problems
*before* the report is built. The monthly report every operations team builds
by hand, in one command.

## Features

- Reads **CSV or JSON** and produces a styled `.xlsx` with three sheets:
  - **Report** — the full data table with a frozen header, filters, banded
    rows, a totals row (live `SUM` formulas), and conditional formatting
    (negative amounts in red, very large amounts in green).
  - **Summary** — totals grouped by category, plus a bar chart.
  - **QC Checks** — missing values, duplicate rows, out-of-range numbers,
    and negative amounts, with the input rows they came from.
- **CLI flags** for input file, output file, and report title.
- **`--demo` mode** generates a report from a built-in sample dataset, so
  you can try it with no input file at all.
- `Amount` is computed from `Units × Unit Price` when the input does not
  already have it.

## Tech Stack

- Python 3.10+
- [openpyxl](https://openpyxl.readthedocs.io/) — workbook creation, styling,
  conditional formatting, and charts
- [pandas](https://pandas.pydata.org/) — CSV/JSON loading and aggregation
- pytest — tests

## How to Run

```bash
pip install -r requirements.txt

# Demo report from the built-in sample dataset
python3 src/generate_report.py --demo

# Your own data (CSV or JSON)
python3 src/generate_report.py --input assets/sample_sales.csv --output report.xlsx

# Custom title and output path
python3 src/generate_report.py --input data/sales.json --output monthly.xlsx --title "October 2026 Sales"
```

See all options with:

```bash
python3 src/generate_report.py --help
```

### Input format

The input needs these columns (headers are case-insensitive, extra columns
are ignored):

| Column | Meaning |
| --- | --- |
| `Date` | Row date, e.g. `2026-10-01` |
| `Product` | Item or service name |
| `Category` | Grouping used by the Summary sheet |
| `Region` | Location, e.g. `Noida` |
| `Units` | Quantity sold |
| `Unit Price` | Price per unit |

A ready-made example is in `assets/sample_sales.csv`. A generated report is
in `assets/sample_report.xlsx`.

## Tests

```bash
python3 -m pytest tests/ -v
```

The suite generates a real workbook in a temporary directory and checks its
sheets, values, totals formula, chart, and QC findings.

## Screenshots

See [`screenshots/SCREENSHOTS.md`](screenshots/SCREENSHOTS.md) — open
`assets/sample_report.xlsx` in Excel / LibreOffice Calc to see the styled
Report, Summary, and QC Checks sheets.

## License

MIT — see [LICENSE](LICENSE).
