# Payroll Portal

Web payroll application modeled on the [Three Rocks Advisory payroll login](https://payroll.threerocksadvisory.com/index.php): separate **admin** and **employee** login, dashboard, employee master, monthly pay runs, payslip register, and employee self-service.

## Quick start

```bash
cd payroll-portal
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open **http://localhost:8080**

### Demo credentials

| Role     | Login        | Password      |
|----------|--------------|---------------|
| Admin    | `admin`      | `admin123`    |
| Employee | `E001`       | `employee123` |

## Features

- **Admin**: dashboard with KPIs, employee list (add employees), finalize monthly pay run (Basic + HRA + allowance, PF, TDS, net), payslip grid.
- **Employee** (`/me`): view payslip history and open printable slip detail.
- Styling aligned with the reference site (`static/assets/app.css`).

Data is stored in `data/payroll.db` (SQLite). Set `PAYROLL_SECRET` in production.
