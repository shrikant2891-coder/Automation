#!/usr/bin/env python3
"""Generate a sample .xlsm report in samples/."""

from __future__ import annotations

from pathlib import Path

from reports.xlsm_report import create_report, save_report

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "samples" / "sample_report.xlsm"


def main() -> Path:
    workbook = create_report(
        rows=[
            ("RPT-001", "Invoice sync", "Passed", 12),
            ("RPT-002", "Vendor export", "Passed", 8),
            ("RPT-003", "Payroll reconcile", "Failed", 3),
        ],
        headers=("Report ID", "Job", "Status", "Rows"),
        sheet_name="Automation Report",
    )
    # Mark as macro-enabled workbook container when saved as .xlsm
    workbook.is_template = False
    path = save_report(workbook, OUTPUT)
    print(f"Wrote {path}")
    return path


if __name__ == "__main__":
    main()
