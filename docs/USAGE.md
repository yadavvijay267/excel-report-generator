# Usage

## Quick start

```bash
pip install -r requirements.txt
python3 src/generate_report.py --demo
```

This writes `report.xlsx` in the current directory.

## Options

| Flag | Description |
| --- | --- |
| `--input, -i` | Input `.csv` or `.json` file (required unless `--demo`) |
| `--output, -o` | Output `.xlsx` file (default: `report.xlsx`) |
| `--title, -t` | Title printed on the Report sheet (default: `Monthly Sales Report`) |
| `--demo` | Use the built-in sample dataset instead of an input file |
| `--help` | Show help and exit |

## Examples

```bash
# From the sample CSV shipped with the repo
python3 src/generate_report.py --input assets/sample_sales.csv --output report.xlsx

# From JSON — an array of objects with the same column names
python3 src/generate_report.py --input data/expenses.json --title "October Expenses" --output expenses.xlsx

# See the QC sheet catch problems: try it on a file with a blank cell or a
# negative quantity and open the QC Checks sheet in the output workbook.
```

## Reading the output

- **Report** — the data. Sort/filter with the header dropdowns; the totals
  row recalculates if you edit values.
- **Summary** — category totals and the bar chart.
- **QC Checks** — fix the flagged input rows and re-run to get a clean sheet.
