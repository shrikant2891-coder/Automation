"""Payroll calculation engine (mirrors .xlsm logic)."""

from __future__ import annotations

import math
from typing import Any

from .xml_store import list_months, root_to_dict


def _f(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def month_index(months: list[dict], label: str) -> int:
    for m in months:
        if m["label"] == label:
            return m["index"]
    raise ValueError(f"Unknown month label: {label}")


def month_by_index(months: list[dict], idx: int) -> dict:
    for m in months:
        if m["index"] == idx:
            return m
    raise ValueError(f"Unknown month index: {idx}")


def settings_float(settings: dict, key: str, default: float) -> float:
    return _f(settings.get(key), default)


def pf_cap_for_month(settings: dict, month_idx: int) -> float:
    legacy = settings_float(settings, "pf_cap_legacy", 1800)
    new = settings_float(settings, "pf_cap_new", 3000)
    eff = int(settings_float(settings, "pf_effective_month_index", 6))
    if month_idx >= eff:
        return new
    return legacy


def salary_components(
    basic: float, pf_status: str, month_idx: int, settings: dict
) -> dict[str, float]:
    hra = basic / 2
    pf_yes = str(pf_status).strip().lower() == "yes"
    cap = pf_cap_for_month(settings, month_idx) if pf_yes else 0.0
    if pf_yes:
        special = hra - cap
    else:
        special = hra
    gross = basic + hra + special
    return {
        "basic": basic,
        "hra": hra,
        "special_allowance": special,
        "pf_employee": cap,
        "gross": gross,
    }


def annual_taxable_projection(emp: dict, months: list[dict], settings: dict) -> float:
    """Approximate annual salary (Salary sheet column K style)."""
    basic = _f(emp["salary"].get("basic"))
    pf = emp["salary"].get("pf_status", "No")
    total = 0.0
    att_map = {a["index"]: a["days"] for a in emp["attendance"]}
    for m in months:
        days_in_month = m["days"]
        att = att_map.get(m["index"], 0)
        if att <= 0:
            continue
        comp = salary_components(basic, pf, m["index"], settings)
        total += comp["gross"] / days_in_month * att
    return total


def income_tax_annual(taxable_annual: float) -> float:
    """Port of Income Tax sheet formula (simplified slab engine)."""
    x = taxable_annual - 75000
    if x <= 1_200_000:
        tax = 0.0
    else:
        def slab(upper: float, rate: float, prev: float) -> float:
            return max(0, min(x, upper) - prev) * rate

        tax = (
            slab(400_000, 0.0, 0)
            + slab(800_000, 0.05, 400_000)
            + slab(1_200_000, 0.10, 800_000)
            + slab(1_600_000, 0.15, 1_200_000)
            + slab(2_000_000, 0.20, 1_600_000)
            + slab(2_400_000, 0.25, 2_000_000)
            + max(0, x - 2_400_000) * 0.30
        )
        tax *= 1.04
        alt = (x - 1_200_000) * 1.04
        tax = min(tax, alt)
    return float(math.ceil(max(0, tax)))


def monthly_tds(annual_tax: float) -> float:
    return round(annual_tax / 12, 0)


def pf_arrears_jul_onwards(
    emp: dict, selected_month_idx: int, settings: dict
) -> float:
    if selected_month_idx < 4:
        return 0.0
    pf_status = emp["salary"].get("pf_status", "No")
    if str(pf_status).strip().lower() != "yes":
        return 0.0
    att_map = {a["index"]: a["days"] for a in emp["attendance"]}
    total = 0.0
    for m_idx in range(4, selected_month_idx + 1):
        if att_map.get(m_idx, 0) > 0:
            total += pf_cap_for_month(settings, m_idx)
    return total


def increment_arrear(emp_code: str, month_idx: int, increments: list[dict]) -> float:
    for inc in increments:
        if inc["emp_code"] != emp_code:
            continue
        # legacy split: partial recognition before month 7
        new_b = _f(inc["new_basic"])
        old_b = _f(inc["old_basic"])
        diff = new_b - old_b
        months = max(1, int(_f(inc.get("arrear_months"), 3)))
        if month_idx > 6:
            return diff
        if month_idx < 7:
            return diff / months * min(months, max(0, month_idx - 3))
    return 0.0


def build_slip(
    data: dict[str, Any], emp_code: str, month_label: str
) -> dict[str, Any]:
    months = data["months"]
    settings = data["settings"]
    m_idx = month_index(months, month_label)
    m = month_by_index(months, m_idx)
    emp = next((e for e in data["employees"] if e["code"] == emp_code), None)
    if not emp:
        raise ValueError(f"Employee not found: {emp_code}")

    att_map = {a["index"]: a["days"] for a in emp["attendance"]}
    days_att = att_map.get(m_idx, 0)
    comp = salary_components(
        _f(emp["salary"].get("basic")),
        emp["salary"].get("pf_status", "No"),
        m_idx,
        settings,
    )
    days_in_month = m["days"]
    prorate = (days_att / days_in_month) if days_in_month and days_att else 0

    basic_p = round(comp["basic"] * prorate, 0)
    hra_p = round(comp["hra"] * prorate, 0)
    special_p = round(comp["special_allowance"] * prorate, 0)
    gross_p = basic_p + hra_p + special_p

    pf_current = comp["pf_employee"] if m_idx >= 4 and days_att > 0 else 0
    pf_arrear = pf_arrears_jul_onwards(emp, m_idx, settings) - pf_current
    if pf_arrear < 0:
        pf_arrear = 0

    annual_taxable = annual_taxable_projection(emp, months, settings)
    annual_tax = income_tax_annual(annual_taxable)
    tds = monthly_tds(annual_tax) if days_att > 0 else 0

    inc_arr = increment_arrear(emp_code, m_idx, data["increments"])

    total_deductions = pf_current + pf_arrear + tds
    net = gross_p + inc_arr - total_deductions

    return {
        "emp_code": emp_code,
        "emp_name": emp["master"].get("name", ""),
        "month": month_label,
        "month_index": m_idx,
        "days_in_month": days_in_month,
        "days_attended": days_att,
        "skip_pdf": days_att <= 0,
        "earnings": {
            "basic": basic_p,
            "hra": hra_p,
            "special_allowance": special_p,
            "increment_arrear": round(inc_arr, 0),
            "gross": gross_p,
        },
        "deductions": {
            "pf_employee": pf_current,
            "pf_arrear": round(pf_arrear, 0),
            "tds": tds,
            "total": round(total_deductions, 0),
        },
        "net_pay": round(net, 0),
        "components_full_month": comp,
        "tax": {"annual_taxable": round(annual_taxable, 0), "annual_tax": annual_tax},
    }


def compute_dashboard(data: dict[str, Any]) -> dict[str, Any]:
    """Enrich all employees with computed salary row for selected default view."""
    settings = data["settings"]
    rows = []
    for emp in data["employees"]:
        basic = _f(emp["salary"].get("basic"))
        pf = emp["salary"].get("pf_status", "No")
        pre = salary_components(basic, pf, 5, settings)
        post = salary_components(basic, pf, 6, settings)
        rows.append(
            {
                "code": emp["code"],
                "name": emp["master"].get("name", ""),
                "status": emp["salary"].get("status", ""),
                "pf_status": pf,
                "basic": basic,
                "pre_sep_gross": pre["gross"],
                "post_sep_gross": post["gross"],
                "pf_pre": pre["pf_employee"],
                "pf_post": post["pf_employee"],
            }
        )
    return {"salary_rows": rows}


def from_xml_root(root) -> dict[str, Any]:
    data = root_to_dict(root)
    data["computed"] = compute_dashboard(data)
    return data
