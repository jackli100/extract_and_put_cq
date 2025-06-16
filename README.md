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
It prints the identifiers of tasks that still need to be returned. Both spreadsheets are read without headers. Place these two spreadsheets in the repository root before running the script.

Run it after installing the requirements:

```bash
python task_report.py
```
