"""Parse plain-text commands and apply them to payroll.xml."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import Callable

from .xml_store import find_employee, load, save, set_setting


class CommandError(Exception):
    pass


def _require(parts: list[str], n: int, usage: str) -> None:
    if len(parts) < n:
        raise CommandError(f"Usage: {usage}")


def apply_command(cmd: str, path=None) -> str:
    cmd = " ".join(cmd.strip().split())
    if not cmd:
        raise CommandError("Empty command")

    tree = load(path)
    root = tree.getroot()
    parts = cmd.split()
    verb = parts[0].lower()

    handlers: dict[str, Callable] = {
        "set": lambda: _cmd_set(root, parts[1:]),
        "add": lambda: _cmd_add(root, parts[1:]),
        "show": lambda: _cmd_show(root, parts[1:]),
        "help": lambda: _cmd_help(),
    }
    if verb not in handlers:
        raise CommandError(f"Unknown verb '{verb}'. Try: help")

    if verb == "help":
        return handlers[verb]()

    msg = handlers[verb]()
    save(tree, path)
    return msg


def _cmd_help() -> str:
    return (
        "Commands:\n"
        "  set setting <name> <value>\n"
        "  set employee <code> basic <amount>\n"
        "  set employee <code> pf_status Yes|No\n"
        "  set employee <code> status Active|Inactive\n"
        "  set attendance <code> <month_label> <days>\n"
        "  add increment <code> <effective_month> <new_basic> <old_basic> [arrear_months]\n"
        "  show settings | show employee <code>\n"
    )


def _cmd_show(root: ET.Element, parts: list[str]) -> str:
    _require(parts, 1, "show settings | show employee <code>")
    if parts[0] == "settings":
        lines = []
        for el in root.findall("./settings/setting"):
            lines.append(f"{el.get('name')}={el.get('value')}")
        return "\n".join(lines) or "(no settings)"
    if parts[0] == "employee":
        _require(parts, 2, "show employee <code>")
        emp = find_employee(root, parts[1])
        if not emp:
            raise CommandError(f"Employee not found: {parts[1]}")
        from xml.etree.ElementTree import tostring

        return ET.tostring(emp, encoding="unicode")
    raise CommandError("show settings | show employee <code>")


def _cmd_set(root: ET.Element, parts: list[str]) -> str:
    _require(parts, 2, "set setting|employee|attendance ...")
    target = parts[0].lower()

    if target == "setting":
        _require(parts, 3, "set setting <name> <value>")
        set_setting(root, parts[1], parts[2])
        return f"Updated setting {parts[1]}={parts[2]}"

    if target == "employee":
        _require(parts, 4, "set employee <code> <field> <value>")
        code, field, value = parts[1], parts[2].lower(), parts[3]
        emp = find_employee(root, code)
        if not emp:
            raise CommandError(f"Employee not found: {code}")
        if field in ("basic", "pf_status", "status"):
            fields: dict[str, str] = {"pf_status": "No", "status": "Active"}
            for el in emp.findall("salary"):
                fields.update({k: str(v) for k, v in el.attrib.items()})
            for el in list(emp.findall("salary")):
                emp.remove(el)
            fields[field] = value
            ET.SubElement(
                emp,
                "salary",
                {k: v for k, v in fields.items() if v != ""},
            )
            return f"Updated employee {code} {field}={value}"
        raise CommandError("Fields: basic, pf_status, status")

    if target == "attendance":
        _require(parts, 4, "set attendance <code> <month_label> <days>")
        code, month_label, days = parts[1], parts[2], parts[3]
        emp = find_employee(root, code)
        if not emp:
            raise CommandError(f"Employee not found: {code}")
        att = emp.find("attendance")
        if att is None:
            att = ET.SubElement(emp, "attendance")
        m_idx = None
        for m in root.findall("./months/month"):
            if m.get("label") == month_label:
                m_idx = m.get("index")
                break
        if not m_idx:
            raise CommandError(f"Unknown month: {month_label}")
        found = None
        for el in att.findall("month"):
            if el.get("label") == month_label:
                found = el
                break
        if found is None:
            found = ET.SubElement(
                att, "month", index=m_idx, label=month_label, days=str(days)
            )
        else:
            found.set("days", str(days))
        return f"Attendance {code} {month_label}={days}"

    raise CommandError("set setting|employee|attendance")


def _cmd_add(root: ET.Element, parts: list[str]) -> str:
    _require(parts, 1, "add increment ...")
    if parts[0] != "increment":
        raise CommandError("Only: add increment ...")
    _require(parts, 6, "add increment <code> <month> <new_basic> <old_basic> [months]")
    code, eff, new_b, old_b = parts[1], parts[2], parts[3], parts[4]
    arrear = parts[5] if len(parts) > 5 else "3"
    incs = root.find("increments")
    if incs is None:
        incs = ET.SubElement(root, "increments")
    ET.SubElement(
        incs,
        "increment",
        emp_code=code,
        effective_month=eff,
        new_basic=new_b,
        old_basic=old_b,
        arrear_months=arrear,
        notes="Added via command",
    )
    return f"Added increment for {code} effective {eff}"
