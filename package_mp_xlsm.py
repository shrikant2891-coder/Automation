#!/usr/bin/env python3
"""Package MP Summary.xlsx (with data sheets) into macro-enabled MP Summary.xlsm."""

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


def create_vba_project_bytes(bas_path: Path) -> bytes:
  tmp = Path("/tmp/mp_macro_shell.xlsm")
  vba = load_vba_source(bas_path)
  with ExcelFile.create_new(str(tmp)) as wb:
      wb.vba_project().add_module("GenerateMPUploader", vba, kind=VBAModuleKind.standard)
      wb.save()
  with zipfile.ZipFile(tmp, "r") as z:
      return z.read("xl/vbaProject.bin")


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
    insert = (
        f'<Relationship Id="{rid}" Type="{VBA_REL_TYPE}" Target="vbaProject.bin"/>'
    )
    return xml.replace("</Relationships>", insert + "</Relationships>")


def inject_macro_into_xlsx(xlsx_path: Path, xlsm_path: Path, vba_bin: bytes) -> None:
    """Start from populated xlsx and add VBA with correct sheet relationships."""
    tmp = xlsm_path.with_suffix(".tmp.xlsm")

    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = patch_content_types(data.decode("utf-8")).encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = patch_workbook_rels(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("xl/vbaProject.bin", vba_bin)

    shutil.move(tmp, xlsm_path)


def verify_xlsm(path: Path) -> None:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    if not names:
        raise RuntimeError(f"{path} has no sheets after packaging")
    print(f"Verified sheets: {names}")


def package(xlsx_path: Path, xlsm_path: Path, bas_path: Path, skip_build: bool = False) -> None:
    if not skip_build:
        build_workbook(xlsx_path)
    vba_bin = create_vba_project_bytes(bas_path)
    inject_macro_into_xlsx(xlsx_path, xlsm_path, vba_bin)
    verify_xlsm(xlsm_path)
    print(f"Created macro-enabled workbook: {xlsm_path}")


def main():
    ap = argparse.ArgumentParser(description="Create MP Summary.xlsm with embedded VBA macro")
    ap.add_argument("--input", default="MP Summary.xlsx")
    ap.add_argument("--output", default="MP Summary.xlsm")
    ap.add_argument("--bas", default="GenerateMPUploader.bas")
    ap.add_argument("--skip-build", action="store_true")
    args = ap.parse_args()
    package(Path(args.input), Path(args.output), Path(args.bas), args.skip_build)


if __name__ == "__main__":
    main()
