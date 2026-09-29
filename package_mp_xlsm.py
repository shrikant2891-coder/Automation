#!/usr/bin/env python3
"""Package MP Summary into macro-enabled MP Summary.xlsm using RetailX VBA shell."""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

VBA_SKIP_PREFIX = "Attribute VB_Name"
VBA_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
SHEET_MAIN_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
MACRO_MAIN_CT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
RETAILX_SHELL = Path("RetailX.xlsm")


def load_vba_source(bas_path: Path) -> str:
    lines = []
    for line in bas_path.read_text(encoding="utf-8").splitlines():
        if line.startswith(VBA_SKIP_PREFIX):
            continue
        lines.append(line)
    return "\r\n".join(lines)


def build_workbook(xlsx_path: Path) -> None:
    subprocess.run(
        [sys.executable, "build_mp_uploader.py", "--input", str(xlsx_path), "--output", str(xlsx_path)],
        check=True,
    )


def inject_mp_module(xlsm_path: Path, bas_path: Path, shell_path: Path) -> None:
    """Use the working RetailX.xlsm VBA project shell and swap in MP macro."""
    shutil.copy2(shell_path, xlsm_path)
    vba = load_vba_source(bas_path)
    with ExcelFile(str(xlsm_path)) as wb:
        project = wb.vba_project()
        for name in list(wb.module_names()):
            if name == "GenerateUploader":
                project.delete_module(name)
        if "GenerateMPUploader" in wb.module_names():
            project.delete_module("GenerateMPUploader")
        project.add_module("GenerateMPUploader", vba, kind=VBAModuleKind.standard)
        wb.save()


def patch_content_types(xml: str) -> str:
    if 'Extension="bin"' not in xml:
        xml = xml.replace(
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>',
            1,
        )
    return xml.replace(SHEET_MAIN_CT, MACRO_MAIN_CT)


def next_rel_id(xml: str) -> str:
    ids = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', xml)]
    return f"rId{max(ids, default=0) + 1}"


def patch_workbook_rels(xml: str) -> str:
    if VBA_REL_TYPE in xml:
        return xml
    rid = next_rel_id(xml)
    insert = f'<Relationship Id="{rid}" Type="{VBA_REL_TYPE}" Target="vbaProject.bin"/>'
    return xml.replace("</Relationships>", insert + "</Relationships>")


def merge_workbook_data(xlsx_path: Path, xlsm_shell_path: Path, xlsm_path: Path) -> None:
    """Start from populated xlsx and inject RetailX-compatible vbaProject.bin."""
    tmp = xlsm_path.with_suffix(".tmp.xlsm")
    skip_from_xlsx = {"[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/vbaProject.bin"}

    with zipfile.ZipFile(xlsm_shell_path, "r") as zmacro:
        vba_bin = zmacro.read("xl/vbaProject.bin")

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        base_ct = patch_content_types(zin.read("[Content_Types].xml").decode("utf-8")).encode("utf-8")
        base_rels = patch_workbook_rels(zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")).encode("utf-8")

    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in skip_from_xlsx:
                continue
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("[Content_Types].xml", base_ct)
        zout.writestr("xl/_rels/workbook.xml.rels", base_rels)
        zout.writestr("xl/vbaProject.bin", vba_bin)

    shutil.move(tmp, xlsm_path)


def verify_xlsm(path: Path) -> None:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    if not names:
        raise RuntimeError(f"{path} has no sheets after packaging")

    with ExcelFile(str(path)) as xlsm:
        modules = xlsm.module_names()
        if "GenerateMPUploader" not in modules:
            raise RuntimeError(f"GenerateMPUploader module missing: {modules}")

    print(f"Verified sheets: {names}")
    print(f"Verified VBA modules: {modules}")


def package(
    xlsx_path: Path,
    xlsm_path: Path,
    bas_path: Path,
    shell_path: Path,
    skip_build: bool = False,
) -> None:
    if not shell_path.exists():
        raise SystemExit(f"RetailX shell not found: {shell_path}")
    if not skip_build:
        build_workbook(xlsx_path)
    shell_tmp = xlsm_path.with_suffix(".shell.xlsm")
    inject_mp_module(shell_tmp, bas_path, shell_path)
    merge_workbook_data(xlsx_path, shell_tmp, xlsm_path)
    shell_tmp.unlink(missing_ok=True)
    verify_xlsm(xlsm_path)
    print(f"Created macro-enabled workbook: {xlsm_path}")


def main():
    ap = argparse.ArgumentParser(description="Create MP Summary.xlsm with RetailX-compatible VBA shell")
    ap.add_argument("--input", default="MP Summary.xlsx")
    ap.add_argument("--output", default="MP Summary.xlsm")
    ap.add_argument("--bas", default="GenerateMPUploader.bas")
    ap.add_argument("--shell", default=str(RETAILX_SHELL))
    ap.add_argument("--skip-build", action="store_true")
    args = ap.parse_args()
    package(
        Path(args.input),
        Path(args.output),
        Path(args.bas),
        Path(args.shell),
        args.skip_build,
    )


if __name__ == "__main__":
    main()
