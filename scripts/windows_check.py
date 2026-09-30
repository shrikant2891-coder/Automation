#!/usr/bin/env python3
"""Quick environment check for dashboard (run: python scripts/windows_check.py)."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
errors = []


def ok(msg: str) -> None:
    print(f"  OK  {msg}")


def bad(msg: str) -> None:
    print(f"  FAIL  {msg}")
    errors.append(msg)


print("Python:", sys.version.split()[0], sys.executable)
print("Project root:", ROOT)
print()

for name in (
    "requirements-dashboard.txt",
    "payroll_dashboard/server.py",
    "payroll-dashboard/config/payroll.xml",
    "start-dashboard.bat",
):
    p = ROOT / name
    if p.is_file():
        ok(name)
    else:
        bad(f"Missing: {name}")

print()
try:
    import fastapi  # noqa: F401
    import uvicorn  # noqa: F401

    ok("fastapi and uvicorn installed")
except ImportError as e:
    bad(f"Packages not installed: {e}. Run: python -m pip install -r requirements-dashboard.txt")

print()
if errors:
    print("Fix the FAIL items above.")
    sys.exit(1)
print("All checks passed. Run start-dashboard.bat or:")
print("  python -m uvicorn payroll_dashboard.server:app --host 127.0.0.1 --port 8080")
