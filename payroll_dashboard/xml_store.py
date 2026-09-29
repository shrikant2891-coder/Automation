"""Read/write payroll.xml (single source of truth for the dashboard)."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_PATH = Path(__file__).resolve().parent.parent / "payroll-dashboard" / "config" / "payroll.xml"


def _path(path: Path | None) -> Path:
    return path or DEFAULT_PATH


def load(path: Path | None = None) -> ET.ElementTree:
    p = _path(path)
    if not p.is_file():
        raise FileNotFoundError(f"Payroll XML not found: {p}")
    return ET.parse(p)


def save(tree: ET.ElementTree, path: Path | None = None) -> None:
    p = _path(path)
    root = tree.getroot()
    root.set("updated", datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))
    ET.indent(tree, space="  ")
    p.parent.mkdir(parents=True, exist_ok=True)
    tree.write(p, encoding="unicode", xml_declaration=True)


def get_setting(root: ET.Element, name: str, default: str = "") -> str:
    for el in root.findall("./settings/setting"):
        if el.get("name") == name:
            return el.get("value", default)
    return default


def set_setting(root: ET.Element, name: str, value: str) -> None:
    settings = root.find("settings")
    if settings is None:
        settings = ET.SubElement(root, "settings")
    for el in settings.findall("setting"):
        if el.get("name") == name:
            el.set("value", value)
            return
    ET.SubElement(settings, "setting", name=name, value=value)


def find_employee(root: ET.Element, code: str) -> ET.Element | None:
    for emp in root.findall("./employees/employee"):
        if emp.get("code") == code:
            return emp
    return None


def list_months(root: ET.Element) -> list[dict[str, Any]]:
    out = []
    for m in root.findall("./months/month"):
        out.append(
            {
                "index": int(m.get("index", 0)),
                "label": m.get("label", ""),
                "days": int(m.get("days", 0)),
            }
        )
    return sorted(out, key=lambda x: x["index"])


def employee_to_dict(emp: ET.Element) -> dict[str, Any]:
    code = emp.get("code", "")
    sal: dict[str, str] = {}
    for el in emp.findall("./salary"):
        sal.update({k: str(v) for k, v in el.attrib.items()})
    master = {el.tag: (el.text or "") for el in emp.findall("./master/*")}
    attendance = []
    for m in emp.findall("./attendance/month"):
        attendance.append(
            {
                "index": int(m.get("index", 0)),
                "label": m.get("label", ""),
                "days": float(m.get("days", 0) or 0),
            }
        )
    return {
        "code": code,
        "salary": sal,
        "master": master,
        "attendance": sorted(attendance, key=lambda x: x["index"]),
    }


def root_to_dict(root: ET.Element) -> dict[str, Any]:
    settings = {
        el.get("name"): el.get("value")
        for el in root.findall("./settings/setting")
    }
    employees = [
        employee_to_dict(emp) for emp in root.findall("./employees/employee")
    ]
    increments = []
    for inc in root.findall("./increments/increment"):
        increments.append(
            {
                "emp_code": inc.get("emp_code", ""),
                "effective_month": inc.get("effective_month", ""),
                "new_basic": inc.get("new_basic", ""),
                "old_basic": inc.get("old_basic", ""),
                "arrear_months": inc.get("arrear_months", "3"),
                "notes": inc.get("notes", ""),
            }
        )
    return {
        "fy": root.get("fy", ""),
        "version": root.get("version", ""),
        "updated": root.get("updated", ""),
        "settings": settings,
        "months": list_months(root),
        "employees": employees,
        "increments": increments,
    }
