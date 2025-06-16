# Area Sum Utility

This repository contains sample data (`0615.xlsx`) and small scripts for computing the total of column F. The new `calc_area.py` script is minimal:

```bash
python calc_area.py
```

First install the requirements:

```bash
pip install -r requirements.txt
```

It reads `0615.xlsx`, converts the sixth column (column F) to numbers and prints the sum.

## Task report

`task_report.py` compares tasks listed in
`ZS-沪乍杭-线路任务单一览表-补定测.xlsx` with those already returned in
`对应表格.xlsx`.
The ZS workbook contains many worksheets named with the two-digit task order. The script scans all sheets whose names begin with digits and extracts codes like "03-01" from the first columns.
It then compares them against the returned codes in `对应表格.xlsx` and prints the identifiers of tasks that still need to be returned. Place both spreadsheets in the repository root before running the script.

Run it after installing the requirements. You can override the filenames with `--zs` and `--returned`:

```bash
python task_report.py --zs ZS-沪乍杭-线路任务单一览表-补定测.xlsx --returned 对应表格.xlsx
```
