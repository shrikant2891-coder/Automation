#!/usr/bin/env python3
"""Package MP Summary.xlsx (with data sheets) into macro-enabled MP Summary.xlsm."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

VBA_SKIP_PREFIX = "Attribute VB_Name"


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


def create_macro_shell(xlsm_path: Path, bas_path: Path) -> None:
    vba = load_vba_source(bas_path)
    with ExcelFile.create_new(str(xlsm_path)) as wb:
        wb.vba_project().add_module("GenerateMPUploader", vba, kind=VBAModuleKind.standard)
        wb.save()


def merge_workbook_data(xlsx_path: Path, xlsm_path: Path) -> None:
    """Replace workbook data in xlsm with content from populated xlsx, keeping VBA."""
    tmp = xlsm_path.with_suffix(".tmp.xlsm")
    skip_from_xlsx = {"[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/vbaProject.bin"}

    with zipfile.ZipFile(xlsm_path, "r") as zmacro:
        vba_bin = zmacro.read("xl/vbaProject.bin")
        macro_ct = zmacro.read("[Content_Types].xml")
        macro_rels = zmacro.read("xl/_rels/workbook.xml.rels")

    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in skip_from_xlsx:
                continue
            zout.writestr(item, zin.read(item.filename))

        zout.writestr("[Content_Types].xml", macro_ct)
        zout.writestr("xl/_rels/workbook.xml.rels", macro_rels)
        zout.writestr("xl/vbaProject.bin", vba_bin)

    shutil.move(tmp, xlsm_path)


def package(xlsx_path: Path, xlsm_path: Path, bas_path: Path, skip_build: bool = False) -> None:
    if not skip_build:
        build_workbook(xlsx_path)
    create_macro_shell(xlsm_path, bas_path)
    merge_workbook_data(xlsx_path, xlsm_path)
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
