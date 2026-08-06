"""Tests for .xlsm report support."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from reports.xlsm_report import (
    SUPPORTED_EXTENSIONS,
    create_report,
    is_supported_report,
    load_report,
    read_sheet_rows,
    save_report,
    write_report_from_dicts,
)


def test_supported_extensions_include_xlsm() -> None:
    assert ".xlsm" in SUPPORTED_EXTENSIONS
    assert is_supported_report("report.xlsm")
    assert is_supported_report("report.XLSX")
    assert not is_supported_report("report.csv")


def test_create_and_save_xlsm(tmp_path: Path) -> None:
    path = tmp_path / "automation_report.xlsm"
    workbook = create_report(
        rows=[("A1", 1), ("A2", 2)],
        headers=("Name", "Value"),
        sheet_name="Results",
    )
    saved = save_report(workbook, path)

    assert saved.exists()
    assert saved.suffix == ".xlsm"

    rows = read_sheet_rows(saved, sheet_name="Results", data_only=True)
    assert rows[0] == ("Name", "Value")
    assert rows[1] == ("A1", 1)
    assert rows[2] == ("A2", 2)


def test_load_rejects_unsupported_extension(tmp_path: Path) -> None:
    path = tmp_path / "notes.txt"
    path.write_text("not excel", encoding="utf-8")
    with pytest.raises(ValueError, match="Unsupported report type"):
        load_report(path)


def test_write_report_from_dicts_xlsm(tmp_path: Path) -> None:
    path = tmp_path / "dict_report.xlsm"
    write_report_from_dicts(
        path,
        [
            {"id": "R1", "status": "ok"},
            {"id": "R2", "status": "error"},
        ],
        sheet_name="Summary",
    )
    rows = read_sheet_rows(path, sheet_name="Summary")
    assert rows[0] == ("id", "status")
    assert rows[1] == ("R1", "ok")
    assert rows[2] == ("R2", "error")


def test_roundtrip_preserves_sheet_name(tmp_path: Path) -> None:
    path = tmp_path / "named.xlsm"
    workbook = create_report(headers=("Col",), sheet_name="Jobs")
    save_report(workbook, path)
    loaded = load_report(path)
    assert loaded.sheetnames == ["Jobs"]


def test_xlsm_uses_macro_enabled_content_type(tmp_path: Path) -> None:
    path = tmp_path / "macro_report.xlsm"
    save_report(create_report(headers=("A",)), path)
    content_types = ZipFile(path).read("[Content_Types].xml").decode("utf-8")
    assert "sheet.macroEnabled.main+xml" in content_types
    assert "spreadsheetml.sheet.main+xml" not in content_types
