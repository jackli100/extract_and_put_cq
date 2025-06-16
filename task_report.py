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

import re
import sys
from collections import defaultdict

try:
    import pandas as pd
except Exception:
    sys.stderr.write("pandas is required to run this script\n")
    raise


TASK_CODE_RE = re.compile(r"(\d{2})(\d{2}|AL)$", re.IGNORECASE)


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


def _load_required_tasks(filename: str = "ZS-沪乍杭-线路任务单一览表-补定测.xlsx"):
    try:
        df = pd.read_excel(filename, header=None)
    except FileNotFoundError:
        sys.stderr.write(f"File '{filename}' not found\n")
        return set()
    except Exception as e:
        sys.stderr.write(f"Failed to read '{filename}': {e}\n")
        return set()

    tasks = set()
    for value in df.iloc[:, 0].tolist():
        parsed = _parse_task_code(value)
        if parsed:
            tasks.add(parsed)
    return tasks


def _load_returned_tasks(filename: str = "对应表格.xlsx"):
    try:
        df = pd.read_excel(filename, header=None)
    except FileNotFoundError:
        sys.stderr.write(f"File '{filename}' not found\n")
        return {}
    except Exception as e:
        sys.stderr.write(f"Failed to read '{filename}': {e}\n")
        return {}

    returned = defaultdict(set)  # form_no -> set of indices or {"AL"}
    for value in df.iloc[:, 1].tolist():
        parsed = _parse_task_code(value)
        if parsed:
            form_no, index = parsed
            returned[form_no].add(index)
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


def main():
    required = _load_required_tasks()
    returned = _load_returned_tasks()
    if not required:
        sys.stderr.write("No required tasks loaded or file missing.\n")
    if not returned:
        sys.stderr.write("No returned task information loaded or file missing.\n")

    remaining = compute_unreturned(required, returned)
    if remaining:
        print("Tasks still to be returned:")
        for form_no, index in remaining:
            print(f"{form_no}{index}")
    else:
        print("All tasks have been returned.")


if __name__ == "__main__":
    main()
