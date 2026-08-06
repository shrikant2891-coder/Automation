# Automation

Python helpers for Excel automation reports, including **`.xlsm`** (macro-enabled) workbooks.

## Supported report formats

| Extension | Description |
|-----------|-------------|
| `.xlsm` | Excel macro-enabled workbook (**supported**) |
| `.xlsx` | Excel workbook |
| `.xltm` | Macro-enabled template |
| `.xltx` | Excel template |

Sheet data can be created, read, and written for all of the above. When you load an existing `.xlsm` / `.xltm` file, VBA macros are **preserved on save**. This library does not execute macros (use Excel / COM / xlwings for that).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```python
from reports import create_report, save_report, read_sheet_rows, is_supported_report

assert is_supported_report("out/daily.xlsm")

wb = create_report(
    headers=("Job", "Status"),
    rows=[("sync", "ok"), ("export", "failed")],
    sheet_name="Report",
)
save_report(wb, "out/daily.xlsm")

print(read_sheet_rows("out/daily.xlsm"))
```

Generate the checked-in sample:

```bash
python scripts/generate_sample_xlsm.py
```

Sample output: [`samples/sample_report.xlsm`](samples/sample_report.xlsm)

## Tests

```bash
pytest
```
