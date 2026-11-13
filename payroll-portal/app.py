"""Payroll portal — admin and employee self-service (mirrors Three Rocks advisory flow)."""

from __future__ import annotations

import os
import secrets
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date
from functools import wraps
from pathlib import Path

from flask import (
    Flask,
    flash,
    g,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash

BASE = Path(__file__).resolve().parent
DB_PATH = BASE / "data" / "payroll.db"

app = Flask(__name__)
app.secret_key = os.environ.get("PAYROLL_SECRET", "dev-payroll-change-in-production")


def _hash_password(password: str) -> str:
    return generate_password_hash(password, method="pbkdf2:sha256")


def csrf_token() -> str:
    if "_csrf" not in session:
        session["_csrf"] = secrets.token_hex(16)
    return session["_csrf"]


def verify_csrf() -> bool:
    return request.form.get("_c", "") == session.get("_csrf", "")


@contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS admins (
                id INTEGER PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY,
                emp_code TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                email TEXT,
                department TEXT,
                designation TEXT,
                doj TEXT,
                basic REAL DEFAULT 0,
                hra REAL DEFAULT 0,
                special_allowance REAL DEFAULT 0,
                pf_employee_pct REAL DEFAULT 12,
                password_hash TEXT NOT NULL,
                active INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS pay_runs (
                id INTEGER PRIMARY KEY,
                month_year TEXT NOT NULL,
                status TEXT DEFAULT 'draft',
                created_at TEXT DEFAULT (datetime('now')),
                UNIQUE(month_year)
            );
            CREATE TABLE IF NOT EXISTS payslips (
                id INTEGER PRIMARY KEY,
                pay_run_id INTEGER NOT NULL,
                employee_id INTEGER NOT NULL,
                days_worked REAL DEFAULT 30,
                basic REAL,
                hra REAL,
                special_allowance REAL,
                gross REAL,
                pf_employee REAL,
                tds REAL,
                other_deductions REAL DEFAULT 0,
                net_pay REAL,
                FOREIGN KEY (pay_run_id) REFERENCES pay_runs(id),
                FOREIGN KEY (employee_id) REFERENCES employees(id),
                UNIQUE(pay_run_id, employee_id)
            );
            """
        )
        admin = conn.execute(
            "SELECT id FROM admins WHERE username = ?", ("admin",)
        ).fetchone()
        if not admin:
            conn.execute(
                "INSERT INTO admins (username, password_hash) VALUES (?, ?)",
                ("admin", _hash_password("admin123")),
            )
        if conn.execute("SELECT COUNT(*) c FROM employees").fetchone()["c"] == 0:
            seed = [
                ("E001", "Shrikant Patel", "shrikant@example.com", "Analytics", "Data Scientist", "2024-01-15", 45000, 18000, 12000),
                ("E002", "Priya Sharma", "priya@example.com", "Finance", "Accountant", "2023-06-01", 38000, 15200, 8000),
                ("E003", "Rahul Mehta", "rahul@example.com", "Operations", "HR Executive", "2022-11-20", 32000, 12800, 6000),
            ]
            default_pw = _hash_password("employee123")
            for row in seed:
                conn.execute(
                    """INSERT INTO employees
                    (emp_code, name, email, department, designation, doj, basic, hra, special_allowance, password_hash)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (*row, default_pw),
                )


@app.before_request
def load_globals():
    g.csrf = csrf_token()


@app.context_processor
def inject_helpers():
    return {"csrf_token": csrf_token}


def admin_required(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Please log in as an administrator.", "bad")
            return redirect(url_for("index"))
        return f(*args, **kwargs)

    return wrapped


def employee_required(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if session.get("role") != "employee":
            flash("Please log in as an employee.", "bad")
            return redirect(url_for("employee_login"))
        return f(*args, **kwargs)

    return wrapped


def calc_payslip(emp: sqlite3.Row, days: float = 30.0) -> dict:
    factor = min(max(days, 0), 30) / 30.0
    basic = round(emp["basic"] * factor, 2)
    hra = round(emp["hra"] * factor, 2)
    special = round(emp["special_allowance"] * factor, 2)
    gross = round(basic + hra + special, 2)
    pf = round(basic * (emp["pf_employee_pct"] / 100.0), 2)
    tds = round(max(gross - pf - 50000 / 12, 0) * 0.05, 2)
    net = round(gross - pf - tds, 2)
    return {
        "basic": basic,
        "hra": hra,
        "special_allowance": special,
        "gross": gross,
        "pf_employee": pf,
        "tds": tds,
        "other_deductions": 0.0,
        "net_pay": net,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    if session.get("role") == "admin":
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        if not verify_csrf():
            flash("Invalid session. Try again.", "bad")
            return redirect(url_for("index"))
        username = request.form.get("u", "").strip()
        password = request.form.get("p", "")
        with get_db() as conn:
            row = conn.execute(
                "SELECT * FROM admins WHERE username = ?", (username,)
            ).fetchone()
        if row and check_password_hash(row["password_hash"], password):
            session.clear()
            session["role"] = "admin"
            session["admin_id"] = row["id"]
            session["admin_name"] = row["username"]
            session["_csrf"] = secrets.token_hex(16)
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "bad")
    return render_template("login.html", title="Payroll Login")


@app.route("/me", methods=["GET", "POST"])
def employee_login():
    if session.get("role") == "employee":
        return redirect(url_for("employee_home"))
    if request.method == "POST":
        if not verify_csrf():
            flash("Invalid session. Try again.", "bad")
            return redirect(url_for("employee_login"))
        emp_code = request.form.get("u", "").strip().upper()
        password = request.form.get("p", "")
        with get_db() as conn:
            row = conn.execute(
                "SELECT * FROM employees WHERE emp_code = ? AND active = 1", (emp_code,)
            ).fetchone()
        if row and check_password_hash(row["password_hash"], password):
            session.clear()
            session["role"] = "employee"
            session["employee_id"] = row["id"]
            session["employee_name"] = row["name"]
            session["_csrf"] = secrets.token_hex(16)
            return redirect(url_for("employee_home"))
        flash("Invalid employee ID or password.", "bad")
    return render_template("employee_login.html", title="Employee Login")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "ok")
    return redirect(url_for("index"))


@app.route("/dashboard")
@admin_required
def dashboard():
    with get_db() as conn:
        emp_count = conn.execute(
            "SELECT COUNT(*) c FROM employees WHERE active = 1"
        ).fetchone()["c"]
        last_run = conn.execute(
            "SELECT * FROM pay_runs ORDER BY id DESC LIMIT 1"
        ).fetchone()
        total_net = 0.0
        if last_run:
            total_net = conn.execute(
                "SELECT COALESCE(SUM(net_pay), 0) s FROM payslips WHERE pay_run_id = ?",
                (last_run["id"],),
            ).fetchone()["s"]
        recent = conn.execute(
            """SELECT pr.month_year, COUNT(ps.id) slips, COALESCE(SUM(ps.net_pay), 0) total
            FROM pay_runs pr LEFT JOIN payslips ps ON ps.pay_run_id = pr.id
            GROUP BY pr.id ORDER BY pr.id DESC LIMIT 6"""
        ).fetchall()
    return render_template(
        "admin_dashboard.html",
        emp_count=emp_count,
        last_run=last_run,
        total_net=total_net,
        chart_data=recent,
    )


@app.route("/employees", methods=["GET", "POST"])
@admin_required
def employees():
    if request.method == "POST" and request.form.get("action") == "add":
        if not verify_csrf():
            flash("Invalid session.", "bad")
            return redirect(url_for("employees"))
        code = request.form.get("emp_code", "").strip().upper()
        name = request.form.get("name", "").strip()
        if not code or not name:
            flash("Employee ID and name are required.", "bad")
        else:
            try:
                with get_db() as conn:
                    conn.execute(
                        """INSERT INTO employees
                        (emp_code, name, email, department, designation, doj, basic, hra, special_allowance, password_hash)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (
                            code,
                            name,
                            request.form.get("email", "").strip(),
                            request.form.get("department", "").strip(),
                            request.form.get("designation", "").strip(),
                            request.form.get("doj", "") or None,
                            float(request.form.get("basic", 0) or 0),
                            float(request.form.get("hra", 0) or 0),
                            float(request.form.get("special_allowance", 0) or 0),
                            _hash_password(request.form.get("password", "employee123") or "employee123"),
                        ),
                    )
                flash(f"Employee {code} added.", "ok")
            except sqlite3.IntegrityError:
                flash(f"Employee ID {code} already exists.", "bad")
        return redirect(url_for("employees"))
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM employees WHERE active = 1 ORDER BY emp_code"
        ).fetchall()
    return render_template("employees.html", employees=rows)


@app.route("/payrun", methods=["GET", "POST"])
@admin_required
def payrun():
    if request.method == "POST":
        if not verify_csrf():
            flash("Invalid session.", "bad")
            return redirect(url_for("payrun"))
        month_year = request.form.get("month_year", "").strip()
        if not month_year:
            flash("Select a pay month.", "bad")
            return redirect(url_for("payrun"))
        with get_db() as conn:
            try:
                cur = conn.execute(
                    "INSERT INTO pay_runs (month_year, status) VALUES (?, 'finalized')",
                    (month_year,),
                )
                pay_run_id = cur.lastrowid
            except sqlite3.IntegrityError:
                flash(f"Pay run for {month_year} already exists.", "warn")
                return redirect(url_for("payrun"))
            emps = conn.execute(
                "SELECT * FROM employees WHERE active = 1"
            ).fetchall()
            for emp in emps:
                slip = calc_payslip(emp)
                conn.execute(
                    """INSERT INTO payslips
                    (pay_run_id, employee_id, basic, hra, special_allowance, gross,
                     pf_employee, tds, other_deductions, net_pay)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        pay_run_id,
                        emp["id"],
                        slip["basic"],
                        slip["hra"],
                        slip["special_allowance"],
                        slip["gross"],
                        slip["pf_employee"],
                        slip["tds"],
                        slip["other_deductions"],
                        slip["net_pay"],
                    ),
                )
        flash(f"Pay run for {month_year} finalized for {len(emps)} employees.", "ok")
        return redirect(url_for("payslips", month=month_year))
    today = date.today()
    default_month = today.strftime("%Y-%m")
    with get_db() as conn:
        runs = conn.execute(
            "SELECT * FROM pay_runs ORDER BY id DESC"
        ).fetchall()
    return render_template("payrun.html", default_month=default_month, runs=runs)


@app.route("/payslips")
@admin_required
def payslips():
    month = request.args.get("month", "")
    with get_db() as conn:
        runs = conn.execute("SELECT * FROM pay_runs ORDER BY id DESC").fetchall()
        selected = None
        slips = []
        if month:
            selected = conn.execute(
                "SELECT * FROM pay_runs WHERE month_year = ?", (month,)
            ).fetchone()
        elif runs:
            selected = runs[0]
            month = selected["month_year"]
        if selected:
            slips = conn.execute(
                """SELECT ps.*, e.emp_code, e.name, e.department
                FROM payslips ps
                JOIN employees e ON e.id = ps.employee_id
                WHERE ps.pay_run_id = ?
                ORDER BY e.emp_code""",
                (selected["id"],),
            ).fetchall()
    return render_template(
        "payslips.html", runs=runs, month=month, slips=slips, selected=selected
    )


@app.route("/employee")
@employee_required
def employee_home():
    eid = session["employee_id"]
    with get_db() as conn:
        emp = conn.execute("SELECT * FROM employees WHERE id = ?", (eid,)).fetchone()
        slips = conn.execute(
            """SELECT ps.*, pr.month_year FROM payslips ps
            JOIN pay_runs pr ON pr.id = ps.pay_run_id
            WHERE ps.employee_id = ? ORDER BY pr.id DESC LIMIT 12""",
            (eid,),
        ).fetchall()
    latest_net = slips[0]["net_pay"] if slips else 0
    return render_template(
        "employee_dashboard.html",
        employee=emp,
        slips=slips,
        latest_net=latest_net,
    )


@app.route("/employee/payslip/<int:slip_id>")
@employee_required
def employee_payslip(slip_id: int):
    eid = session["employee_id"]
    with get_db() as conn:
        slip = conn.execute(
            """SELECT ps.*, pr.month_year, e.name, e.emp_code, e.department, e.designation
            FROM payslips ps
            JOIN pay_runs pr ON pr.id = ps.pay_run_id
            JOIN employees e ON e.id = ps.employee_id
            WHERE ps.id = ? AND ps.employee_id = ?""",
            (slip_id, eid),
        ).fetchone()
    if not slip:
        flash("Payslip not found.", "bad")
        return redirect(url_for("employee_home"))
    return render_template("payslip_detail.html", slip=slip)


if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
