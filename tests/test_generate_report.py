"""Tests for the Excel Report Generator.

Run from the project root with::

    python3 -m pytest tests/ -v
"""

import sys
from pathlib import Path

import pytest
import pandas as pd
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import generate_report  # noqa: E402


@pytest.fixture()
def sample_frame():
    return generate_report.demo_dataframe()


def test_demo_dataframe_has_expected_shape(sample_frame):
    assert list(sample_frame.columns) == [
        "Date", "Product", "Category", "Region", "Units", "Unit Price", "Amount",
    ]
    assert len(sample_frame) == 20
    assert (sample_frame["Amount"] == sample_frame["Units"] * sample_frame["Unit Price"]).all()


def test_generate_report_creates_workbook(tmp_path, sample_frame):
    output = tmp_path / "report.xlsx"
    stats = generate_report.generate_report(sample_frame, output, "Test Report")
    assert output.exists()
    assert output.stat().st_size > 0
    assert stats["rows"] == 20
    assert stats["total_amount"] == pytest.approx(float(sample_frame["Amount"].sum()))

    wb = load_workbook(output)
    assert wb.sheetnames == ["Report", "Summary", "QC Checks"]

    report = wb["Report"]
    assert report["A1"].value == "Test Report"
    assert report["A4"].value == "Date"
    assert report["G4"].value == "Amount"
    # Totals row uses a live SUM formula over the Amount column.
    total_row = 4 + 1 + 20
    assert report.cell(row=total_row, column=1).value == "TOTAL"
    assert str(report.cell(row=total_row, column=7).value).startswith("=SUM(")

    summary = wb["Summary"]
    assert summary["A1"].value == "Summary by Category"
    categories = {summary.cell(row=r, column=1).value for r in range(4, 6)}
    assert categories == {"Hardware", "Services"}
    assert len(summary._charts) == 1  # one bar chart

    qc = wb["QC Checks"]
    assert qc["A1"].value.startswith("QC Checks")


def test_qc_checks_flag_bad_data():
    frame = generate_report.demo_dataframe()
    frame.loc[0, "Units"] = -3  # out-of-range + negative amount
    frame["Amount"] = frame["Units"] * frame["Unit Price"]  # recompute like load_input does
    frame.loc[1, "Product"] = None  # missing value
    frame = pd.concat([frame, frame.iloc[[2]]], ignore_index=True)  # duplicate row
    findings = generate_report.qc_checks(frame)
    kinds = {f["check"] for f in findings}
    assert "Missing value" in kinds
    assert "Duplicate row" in kinds
    assert "Out-of-range value" in kinds
    assert "Negative amount" in kinds


def test_qc_checks_pass_on_clean_data(sample_frame):
    findings = generate_report.qc_checks(sample_frame)
    assert len(findings) == 1
    assert findings[0]["check"] == "All checks passed"


def test_load_input_csv(tmp_path):
    src = Path(__file__).resolve().parents[1] / "assets" / "sample_sales.csv"
    frame = generate_report.load_input(src)
    assert len(frame) == 20
    assert "Amount" in frame.columns


def test_load_input_missing_file(tmp_path):
    with pytest.raises(FileNotFoundError):
        generate_report.load_input(tmp_path / "nope.csv")


def test_cli_help_exits_zero(capsys):
    with pytest.raises(SystemExit) as exc:
        generate_report.build_parser().parse_args(["--help"])
    assert exc.value.code == 0


def test_cli_demo_writes_xlsx(tmp_path):
    output = tmp_path / "demo.xlsx"
    assert generate_report.main(["--demo", "--output", str(output)]) == 0
    assert output.exists()


def test_cli_without_input_or_demo_fails():
    assert generate_report.main([]) == 2
