# XML Payroll Dashboard

Web dashboard driven by **`payroll-dashboard/config/payroll.xml`**. The same file powers the UI and CLI—commands you run update XML and the dashboard refreshes.

## Start the dashboard

```bash
cd /workspace
pip install -r requirements-dashboard.txt
uvicorn payroll_dashboard.server:app --host 0.0.0.0 --port 8080
```

Open **http://localhost:8080**

## Command line (same commands as the dashboard bar)

```bash
python -m payroll_dashboard.cli help

python -m payroll_dashboard.cli "set setting pf_cap_new 3000"
python -m payroll_dashboard.cli "set employee INF0001 basic 32000"
python -m payroll_dashboard.cli "set attendance INF0001 Sep'26 28"
python -m payroll_dashboard.cli "add increment INF0099 Oct'26 40000 35000 3"
```

## Re-import from Excel (.xlsm)

After you change the workbook:

```bash
python3 scripts/export_xlsm_to_xml.py
```

Then refresh the dashboard (or restart the server).

## Dashboard sections (maps to Excel sheets)

| Tab | Excel sheet |
|-----|-------------|
| Generator | Generator + Salary Slip preview |
| Settings | Payroll Settings |
| Salary | Salary sheet (computed pre/post Sep gross & PF) |
| Master Data | Master Data |
| Attendance | Salary sheet attendance columns |
| Increments | Incremental sheet |
| Income Tax | Income Tax (annual estimate) |
| XML | Raw `payroll.xml` |

## API

- `GET /api/payroll` — full data + computed salary rows  
- `GET /api/slip?emp=INF0001&month=Sep'26` — payslip JSON  
- `POST /api/command` — body `{"command": "set setting pf_cap_new 3000"}`  
- `GET /api/xml` — raw XML  

PDF export is not in the web app; use Excel macros for bulk PDF, or print the slip preview from the browser.
