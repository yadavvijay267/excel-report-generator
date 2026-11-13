#!/usr/bin/env python3
"""Excel Report Generator.

Turn raw CSV or JSON sales/expense data into a formatted Excel workbook:

* ``Report``   - the full styled data table, with a totals row and
                 conditional formatting that highlights unusually large
                 amounts and negative values.
* ``Summary``  - totals grouped by category, plus a bar chart.
* ``QC Checks``- data-quality findings found before the report was built
                 (missing values, duplicate rows, out-of-range numbers).

Usage::

    python3 src/generate_report.py --demo
    python3 src/generate_report.py --input data/sales.csv --output report.xlsx
    python3 src/generate_report.py --input data/sales.json --title "Q3 Sales"

The demo mode writes a built-in monthly sales dataset, so the tool can be
tried with no input file at all.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import Counter
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------

NAVY = "1F3864"
CYAN = "2BB3C0"
LIGHT_GREY = "F2F2F2"
WHITE = "FFFFFF"
RED_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color=WHITE)
TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=11, color="595959")
TOTAL_FONT = Font(name="Calibri", size=11, bold=True, color=NAVY)
THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

# Columns the report understands. ``Amount`` is computed from
# ``Units * Unit Price`` when it is missing from the input.
REQUIRED_COLUMNS = ["Date", "Product", "Category", "Region", "Units", "Unit Price"]
ALL_COLUMNS = REQUIRED_COLUMNS + ["Amount"]

DEMO_ROWS = [
    ("2026-10-01", "Fiber Splice Kit", "Hardware", "Noida", 12, 1450.00),
    ("2026-10-02", "OTDR Rental (day)", "Services", "Delhi", 3, 5200.00),
    ("2026-10-03", "Cat6 Cable (305m)", "Hardware", "Noida", 8, 3850.00),
    ("2026-10-04", "Patch Panel 24-port", "Hardware", "Gurugram", 5, 2100.00),
    ("2026-10-05", "Site Survey", "Services", "Noida", 2, 7500.00),
    ("2026-10-06", "Fiber Splice Kit", "Hardware", "Delhi", 7, 1450.00),
    ("2026-10-07", "Network Audit", "Services", "Gurugram", 1, 12000.00),
    ("2026-10-08", "Cat6 Cable (305m)", "Hardware", "Noida", 4, 3850.00),
    ("2026-10-09", "Cable Tray (3m)", "Hardware", "Delhi", 15, 640.00),
    ("2026-10-10", "OTDR Rental (day)", "Services", "Noida", 2, 5200.00),
    ("2026-10-11", "Patch Panel 24-port", "Hardware", "Noida", 6, 2100.00),
    ("2026-10-12", "Fiber Splice Kit", "Hardware", "Gurugram", 9, 1450.00),
    ("2026-10-13", "Site Survey", "Services", "Delhi", 1, 7500.00),
    ("2026-10-14", "Cable Tray (3m)", "Hardware", "Noida", 20, 640.00),
    ("2026-10-15", "Network Audit", "Services", "Noida", 2, 12000.00),
    ("2026-10-16", "Cat6 Cable (305m)", "Hardware", "Gurugram", 6, 3850.00),
    ("2026-10-17", "Fiber Splice Kit", "Hardware", "Noida", 11, 1450.00),
    ("2026-10-18", "OTDR Rental (day)", "Services", "Delhi", 4, 5200.00),
    ("2026-10-19", "Patch Panel 24-port", "Hardware", "Delhi", 3, 2100.00),
    ("2026-10-20", "Site Survey", "Services", "Gurugram", 3, 7500.00),
]


def demo_dataframe() -> pd.DataFrame:
    """Return the built-in demo dataset as a DataFrame."""
    frame = pd.DataFrame(DEMO_ROWS, columns=REQUIRED_COLUMNS)
    frame["Amount"] = frame["Units"] * frame["Unit Price"]
    return frame


# ---------------------------------------------------------------------------
# Input loading
# ---------------------------------------------------------------------------

def load_input(path: Path) -> pd.DataFrame:
    """Load a CSV or JSON file into a DataFrame, normalising column names."""
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    if path.suffix.lower() == ".csv":
        frame = pd.read_csv(path)
    elif path.suffix.lower() == ".json":
        frame = pd.read_json(path)
    else:
        raise ValueError(f"Unsupported input format: {path.suffix} (use .csv or .json)")

    # Normalise headers: strip whitespace and title-case them.
    frame.columns = [str(c).strip().title() for c in frame.columns]
    # Accept common aliases for the price column.
    frame = frame.rename(columns={"Unitprice": "Unit Price", "Unit_Price": "Unit Price"})
    missing = [c for c in REQUIRED_COLUMNS if c not in frame.columns]
    if missing:
        raise ValueError(f"Input is missing required column(s): {', '.join(missing)}")
    frame = frame[REQUIRED_COLUMNS].copy()
    frame["Units"] = pd.to_numeric(frame["Units"], errors="coerce")
    frame["Unit Price"] = pd.to_numeric(frame["Unit Price"], errors="coerce")
    frame["Amount"] = frame["Units"] * frame["Unit Price"]
    return frame


# ---------------------------------------------------------------------------
# QC checks
# ---------------------------------------------------------------------------

def qc_checks(frame: pd.DataFrame) -> list[dict]:
    """Inspect the raw data and return a list of quality findings.

    Each finding is a dict with ``check``, ``detail`` and ``rows`` keys.
    An empty list means the data passed every check.
    """
    findings: list[dict] = []

    for column in frame.columns:
        missing = int(frame[column].isna().sum())
        if missing:
            rows = [str(i + 2) for i in frame.index[frame[column].isna()][:5]]
            findings.append({
                "check": "Missing value",
                "detail": f"{missing} missing value(s) in column '{column}'",
                "rows": ", ".join(rows),
            })

    duplicates = int(frame.duplicated().sum())
    if duplicates:
        findings.append({
            "check": "Duplicate row",
            "detail": f"{duplicates} exact duplicate row(s) in the dataset",
            "rows": "see Report sheet",
        })

    bad_units = frame[(frame["Units"].isna()) | (frame["Units"] <= 0)]
    if not bad_units.empty:
        findings.append({
            "check": "Out-of-range value",
            "detail": f"{len(bad_units)} row(s) with Units <= 0 or not a number",
            "rows": ", ".join(str(i + 2) for i in bad_units.index[:5]),
        })

    negative = frame[frame["Amount"] < 0]
    if not negative.empty:
        findings.append({
            "check": "Negative amount",
            "detail": f"{len(negative)} row(s) with a negative Amount",
            "rows": ", ".join(str(i + 2) for i in negative.index[:5]),
        })

    huge = frame[frame["Amount"] > 100_000]
    if not huge.empty:
        findings.append({
            "check": "Unusually large amount",
            "detail": f"{len(huge)} row(s) with Amount above 100,000 — please verify",
            "rows": ", ".join(str(i + 2) for i in huge.index[:5]),
        })

    if not findings:
        findings.append({
            "check": "All checks passed",
            "detail": "No missing values, duplicates, or out-of-range values found.",
            "rows": "—",
        })
    return findings


# ---------------------------------------------------------------------------
# Workbook building
# ---------------------------------------------------------------------------

def _style_header_row(ws, row: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def _add_report_sheet(wb: Workbook, frame: pd.DataFrame, title: str) -> None:
    ws = wb.active
    ws.title = "Report"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(ALL_COLUMNS))
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(ALL_COLUMNS))
    ws.cell(row=2, column=1,
           value=f"{len(frame)} rows  •  Generated by excel-report-generator").font = SUBTITLE_FONT

    header_row = 4
    for col, name in enumerate(ALL_COLUMNS, start=1):
        ws.cell(row=header_row, column=col, value=name)
    _style_header_row(ws, header_row, len(ALL_COLUMNS))

    for r, row in enumerate(frame.itertuples(index=False), start=header_row + 1):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = THIN_BORDER
            if c in (5, 6, 7):  # Units, Unit Price, Amount
                cell.number_format = "#,##0.00" if c > 5 else "#,##0"
                cell.alignment = Alignment(horizontal="right")
            if r % 2 == 0:
                cell.fill = PatternFill(start_color=LIGHT_GREY, end_color=LIGHT_GREY,
                                        fill_type="solid")

    total_row = header_row + 1 + len(frame)
    ws.cell(row=total_row, column=1, value="TOTAL").font = TOTAL_FONT
    for col in (5, 6, 7):
        letter = get_column_letter(col)
        cell = ws.cell(row=total_row, column=col)
        cell.value = f"=SUM({letter}{header_row + 1}:{letter}{total_row - 1})"
        cell.font = TOTAL_FONT
        cell.number_format = "#,##0.00" if col > 5 else "#,##0"
        cell.border = THIN_BORDER
    ws.cell(row=total_row, column=1).border = THIN_BORDER

    # Conditional formatting: flag negative amounts red, very large ones green.
    amount_letter = get_column_letter(7)
    data_range = f"{amount_letter}{header_row + 1}:{amount_letter}{total_row - 1}"
    ws.conditional_formatting.add(
        data_range, CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    ws.conditional_formatting.add(
        data_range, CellIsRule(operator="greaterThan", formula=["50000"], fill=GREEN_FILL))

    widths = {"A": 13, "B": 24, "C": 12, "D": 12, "E": 9, "F": 13, "G": 14}
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.auto_filter.ref = f"A{header_row}:G{total_row - 1}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _add_summary_sheet(wb: Workbook, frame: pd.DataFrame) -> None:
    ws = wb.create_sheet("Summary")
    ws.merge_cells("A1:C1")
    ws.cell(row=1, column=1, value="Summary by Category").font = TITLE_FONT
    for col, name in enumerate(["Category", "Rows", "Total Amount"], start=1):
        ws.cell(row=3, column=col, value=name)
    _style_header_row(ws, 3, 3)

    grouped = (frame.groupby("Category", as_index=False)
               .agg(Rows=("Amount", "size"), Total=("Amount", "sum"))
               .sort_values("Total", ascending=False))
    for r, (_, row) in enumerate(grouped.iterrows(), start=4):
        ws.cell(row=r, column=1, value=row["Category"]).border = THIN_BORDER
        ws.cell(row=r, column=2, value=int(row["Rows"])).border = THIN_BORDER
        amount = ws.cell(row=r, column=3, value=float(row["Total"]))
        amount.number_format = "#,##0.00"
        amount.border = THIN_BORDER

    total_row = 4 + len(grouped)
    ws.cell(row=total_row, column=1, value="TOTAL").font = TOTAL_FONT
    ws.cell(row=total_row, column=2, value=f"=SUM(B4:B{total_row - 1})").font = TOTAL_FONT
    total_cell = ws.cell(row=total_row, column=3,
                        value=f"=SUM(C4:C{total_row - 1})")
    total_cell.font = TOTAL_FONT
    total_cell.number_format = "#,##0.00"

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 18

    chart = BarChart()
    chart.type = "col"
    chart.title = "Total Amount by Category"
    chart.y_axis.title = "Amount"
    chart.x_axis.title = "Category"
    chart.style = 10
    chart.height = 7.5
    chart.width = 14
    data = Reference(ws, min_col=3, min_row=3, max_row=total_row - 1)
    cats = Reference(ws, min_col=1, min_row=4, max_row=total_row - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    ws.add_chart(chart, f"A{total_row + 2}")


def _add_qc_sheet(wb: Workbook, findings: list[dict]) -> None:
    ws = wb.create_sheet("QC Checks")
    ws.merge_cells("A1:C1")
    ws.cell(row=1, column=1, value="QC Checks — data quality before reporting").font = TITLE_FONT
    for col, name in enumerate(["Check", "Detail", "Row(s) in input"], start=1):
        ws.cell(row=3, column=col, value=name)
    _style_header_row(ws, 3, 3)
    for r, finding in enumerate(findings, start=4):
        ws.cell(row=r, column=1, value=finding["check"]).border = THIN_BORDER
        ws.cell(row=r, column=2, value=finding["detail"]).border = THIN_BORDER
        ws.cell(row=r, column=3, value=finding["rows"]).border = THIN_BORDER
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 58
    ws.column_dimensions["C"].width = 18


def generate_report(frame: pd.DataFrame, output: Path, title: str) -> dict:
    """Build the workbook from *frame* and write it to *output*.

    Returns a small stats dict (rows, total amount, QC finding count).
    """
    output.parent.mkdir(parents=True, exist_ok=True)
    findings = qc_checks(frame)
    wb = Workbook()
    _add_report_sheet(wb, frame, title)
    _add_summary_sheet(wb, frame)
    _add_qc_sheet(wb, findings)
    wb.save(output)
    return {
        "rows": len(frame),
        "total_amount": float(frame["Amount"].sum()),
        "qc_findings": len(findings),
        "output": str(output),
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="generate_report.py",
        description="Turn a CSV/JSON dataset into a formatted Excel report "
                    "with a summary chart and QC checks.")
    parser.add_argument("--input", "-i", type=Path, default=None,
                        help="Input .csv or .json file (required unless --demo)")
    parser.add_argument("--output", "-o", type=Path, default=Path("report.xlsx"),
                        help="Output .xlsx file (default: report.xlsx)")
    parser.add_argument("--title", "-t", default="Monthly Sales Report",
                        help="Report title shown on the Report sheet")
    parser.add_argument("--demo", action="store_true",
                        help="Generate the report from the built-in demo dataset")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.demo:
        frame = demo_dataframe()
    elif args.input is not None:
        try:
            frame = load_input(args.input)
        except (FileNotFoundError, ValueError) as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 2
    else:
        print("Error: provide --input <file> or use --demo.", file=sys.stderr)
        return 2

    if frame.empty:
        print("Error: input contains no data rows.", file=sys.stderr)
        return 2

    stats = generate_report(frame, args.output, args.title)
    print(f"Wrote {stats['output']} — {stats['rows']} rows, "
          f"total {stats['total_amount']:,.2f}, "
          f"{stats['qc_findings']} QC finding(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
