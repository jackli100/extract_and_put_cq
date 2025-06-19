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
It then compares them against the returned codes in `对应表格.xlsx` and prints a tidy summary for each form. Cells in the returned table may contain several task codes separated by spaces such as "0101 0102 0105"; the script recognises each of them individually. The code also classifies tasks containing the keywords "抄平"/"超平" as **道路抄平** and those with "核补" as **核补地形** (unless already categorised as 道路抄平) and outputs a count of how many of each category have been returned. Place both spreadsheets in the repository root before running the script. You can also write an Excel report summarising the remaining tasks using the `-o/--output` option and a PNG bar chart with `-c/--chart`. The report adds a `Category` sheet with totals and a `CategoryDetail` sheet listing each classified task and whether it has been returned. The script also writes the returned flag back into the ZS workbook in memory and saves the updated workbook as `<original>_processed.xlsx`, leaving the source file untouched.

Run it after installing the requirements. You can override the filenames with `--zs` and `--returned`:

```bash
python task_report.py --zs ZS-沪乍杭-线路任务单一览表-补定测.xlsx --returned 对应表格.xlsx
python task_report.py -o report.xlsx -c progress.png
```

## CAD Utilities

Three small scripts help with DWG/DXF files:

- `collect_dwg.py SRC DEST` copies all `.dwg` files under `SRC` to `DEST` skipping duplicate file names.
- `convert_dwg_to_dxf.py SRC DEST` converts DWG files in `SRC` to DXF files written to `DEST` using the ODA File Converter.
- `merge_dxf.py OUTPUT FILES...` merges the contents of multiple DXF files into one file `OUTPUT`.

The conversion and merge scripts require `ezdxf` and the ODA File Converter to be installed.
