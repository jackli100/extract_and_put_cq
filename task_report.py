# -*- coding: utf-8 -*-
"""Generate a report of tasks that still need to be returned.

The script reads two Excel files:
 - `ZS-沪乍杭-线路任务单一览表-补定测.xlsx`: column A lists every task that must be returned. Each
   value in column A should contain a form number plus a task index,
   e.g. "0101" meaning form 01 task 01.
 - `对应表格.xlsx`: the first column is a file name, the second column
   contains returned task identifiers. The returned task identifier has
   the same four character format where the first two digits are the
   form number and the last two digits are the task index. If the last
   two characters are ``AL`` it means all tasks for that form have been
   returned. ``0000`` means the mapping is unknown and is ignored.

The script compares both files and prints the tasks that are still
unreturned.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from typing import Iterable, List

try:
    import pandas as pd
    import matplotlib.pyplot as plt
except Exception:
    sys.stderr.write("pandas and matplotlib are required to run this script\n")
    raise


DEFAULT_ZS_FILE = "ZS-沪乍杭-线路任务单一览表-补定测.xlsx"
DEFAULT_RETURNED_FILE = "对应表格.xlsx"

# Match a task code consisting of two digits for the form number followed by two
# digits or ``AL`` for the task index. The negative look-behind/ahead ensure the
# code is not part of a longer sequence of digits.
TASK_CODE_RE = re.compile(r"(?<!\d)(\d{2})(\d{2}|AL)(?!\d)", re.IGNORECASE)


def _parse_task_code(value: object):
    """Parse a task code like ``0101`` or ``03AL``.

    Returns ``(form_no, index)`` or ``None`` if the value cannot be
    parsed or represents an unknown mapping (``0000``).
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper() == "0000":
        return None
    m = TASK_CODE_RE.search(text)
    if not m:
        return None
    form_no, index = m.groups()
    index = index.upper()
    return form_no, index


TASK_TEXT_RE = re.compile(r"(\d+)[-－](\d+)")


def _load_required_tasks(filename: str = DEFAULT_ZS_FILE):
    try:
        xls = pd.ExcelFile(filename)
    except FileNotFoundError:
        sys.stderr.write(f"File '{filename}' not found\n")
        return set()
    except Exception as e:
        sys.stderr.write(f"Failed to read '{filename}': {e}\n")
        return set()

    tasks = set()
    for sheet in xls.sheet_names:
        if not re.match(r"\d+", sheet):
            continue  # skip summary or other sheets
        df = xls.parse(sheet, header=None)
        for row in df.itertuples(index=False):
            for cell in row:
                if isinstance(cell, str):
                    m = TASK_TEXT_RE.search(cell)
                    if m:
                        form_no, index = m.groups()
                        tasks.add((form_no.zfill(2), index.zfill(2)))
                        break
    return tasks


def _load_returned_tasks(filename: str = DEFAULT_RETURNED_FILE):
    try:
        df = pd.read_excel(filename, header=None)
    except FileNotFoundError:
        sys.stderr.write(f"File '{filename}' not found\n")
        return {}
    except Exception as e:
        sys.stderr.write(f"Failed to read '{filename}': {e}\n")
        return {}

    returned = defaultdict(set)  # form_no -> set of indices or {"AL"}
    for value in df.iloc[:, 1].dropna().tolist():
        text = str(value)
        for form_no, index in TASK_CODE_RE.findall(text.upper()):
            # ignore unknown mapping "0000"
            if form_no == "00" and index == "00":
                continue
            returned[form_no].add(index.upper())
    return returned


def compute_unreturned(required_tasks, returned_tasks):
    remaining = []
    for form_no, index in sorted(required_tasks):
        returned_indices = returned_tasks.get(form_no)
        if not returned_indices:
            remaining.append((form_no, index))
            continue
        if "AL" in returned_indices:
            continue  # all tasks for this form already returned
        if index not in returned_indices:
            remaining.append((form_no, index))
    return remaining


def compute_summary(required_tasks, returned_tasks):
    required_map: defaultdict[str, set[str]] = defaultdict(set)
    for form_no, index in required_tasks:
        required_map[form_no].add(index)

    summary = []
    remaining = []
    for form_no in sorted(required_map):
        required_indices = required_map[form_no]
        returned_indices = returned_tasks.get(form_no, set())
        if "AL" in returned_indices:
            returned_count = len(required_indices)
            missing_indices: List[str] = []
        else:
            missing_indices = sorted(required_indices - returned_indices)
            returned_count = len(required_indices) - len(missing_indices)
        ratio = (
            returned_count / len(required_indices)
            if required_indices
            else 0
        )
        summary.append(
            {
                "Form": form_no,
                "Required": len(required_indices),
                "Returned": returned_count,
                "Missing": " ".join(missing_indices),
                "Ratio": ratio,
            }
        )
        for index in missing_indices:
            remaining.append((form_no, index))
    return remaining, summary


def save_report(
    remaining: Iterable[tuple[str, str]],
    summary: List[dict],
    output_file: str,
) -> None:
    """Save a visual report of remaining tasks to ``output_file``.

    The report contains a sheet listing each remaining task and a summary
    sheet with counts per form number and return ratios.
    """
    if not remaining:
        df = pd.DataFrame(columns=["Form", "TaskIndex"])
    else:
        df = pd.DataFrame(remaining, columns=["Form", "TaskIndex"])
    summary_df = pd.DataFrame(summary)
    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Remaining")
            summary_df.to_excel(writer, index=False, sheet_name="Summary")
    except Exception as e:
        sys.stderr.write(f"Failed to write report '{output_file}': {e}\n")


def save_bar_chart(summary: List[dict], chart_file: str) -> None:
    """Save a bar chart of return ratios for each form."""
    if not summary:
        return
    forms = [item["Form"] for item in summary]
    ratios = [item["Ratio"] for item in summary]
    plt.figure(figsize=(max(6, len(forms) * 0.6), 4))
    plt.bar(forms, ratios, color="skyblue")
    plt.ylim(0, 1)
    plt.xlabel("Form")
    plt.ylabel("Returned / Required")
    plt.title("Return Ratio by Form")
    plt.tight_layout()
    try:
        plt.savefig(chart_file)
    except Exception as e:
        sys.stderr.write(f"Failed to save chart '{chart_file}': {e}\n")
    finally:
        plt.close()


def main(argv=None):
    parser = argparse.ArgumentParser(description="Report unreturned tasks")
    parser.add_argument(
        "--zs",
        default=DEFAULT_ZS_FILE,
        help="Excel file listing required tasks (default: %(default)s)",
    )
    parser.add_argument(
        "--returned",
        default=DEFAULT_RETURNED_FILE,
        help="Excel file listing already returned tasks (default: %(default)s)",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Write an Excel report to this file",
    )
    parser.add_argument(
        "-c",
        "--chart",
        help="Save a bar chart of return ratios to this PNG file",
    )
    args = parser.parse_args(argv)

    required = _load_required_tasks(args.zs)
    returned = _load_returned_tasks(args.returned)
    if not required:
        sys.stderr.write("No required tasks loaded or file missing.\n")
    if not returned:
        sys.stderr.write("No returned task information loaded or file missing.\n")

    remaining, summary = compute_summary(required, returned)
    if args.output:
        save_report(remaining, summary, args.output)
        print(f"Report written to {args.output}")
    if args.chart:
        save_bar_chart(summary, args.chart)
        print(f"Chart saved to {args.chart}")

    for idx, item in enumerate(summary, 1):
        if not item["Missing"]:
            print(f"{idx}: \u5168\u9f50")  # 全齐
        else:
            print(f"{idx}: \u7f3a{item['Missing']}")  # 缺...


if __name__ == "__main__":
    main()
