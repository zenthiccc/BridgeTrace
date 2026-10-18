"""Generate a fully synthetic leave export + calendar so the demo runs end to end.

Usage:
    python scripts/make_sample_data.py

Writes:
    data/raw/leave_export_20260701_080000.xlsx   (fake employees, one per scenario)
    samples/sample_calendar.xlsx                 (from samples/sample_calendar.md)
"""
from __future__ import annotations

import sys
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import config  # noqa: E402
from convert_calendar import convert  # noqa: E402

HEADERS = [
    "Department", "Employee ID", "Employee Name", "Internal Ref", "Leave Code", "End Date",
    "Start Date", "Last Modified", "Start Time", "End Time", "Leave Days", "Leave Hours",
]

# (employee id, name, [(code, start, end, start_time)]) - one scenario per fake employee.
SCENARIOS = [
    ("E001", "Employee A", [  # SL -> AL -> SL (AM)          : included
        ("SL", date(2026, 1, 13), None, time(9)),
        ("AL", date(2026, 1, 14), None, time(9)),
        ("SL", date(2026, 1, 15), None, time(9)),
    ]),
    ("E002", "Employee B", [  # SL -> weekend -> SL          : included
        ("SL", date(2026, 3, 6), None, None),
        ("SLH", date(2026, 3, 9), None, time(8, 30)),
    ]),
    ("E003", "Employee C", [  # SL -> holiday -> SL          : included
        ("SL", date(2026, 6, 16), None, None),
        ("SLX", date(2026, 6, 18), date(2026, 6, 19), None),
    ]),
    ("E004", "Employee D", [  # AL -> SL                     : excluded
        ("AL", date(2026, 2, 2), None, None),
        ("SL", date(2026, 2, 3), None, None),
    ]),
    ("E005", "Employee E", [  # SL -> AL -> SL (PM)          : excluded
        ("SL", date(2026, 4, 7), None, None),
        ("AL", date(2026, 4, 8), None, None),
        ("SLH", date(2026, 4, 9), None, time(13, 30)),
    ]),
    ("E006", "Employee F", [  # SL -> substitute workday -> SL : excluded
        ("SL", date(2026, 6, 19), None, None),
        ("SL", date(2026, 6, 22), None, None),
    ]),
]


def build_rows():
    for emp_id, name, leaves in SCENARIOS:
        for code, start, end, start_time in leaves:
            end = end or start
            days = (end - start).days + 1
            yield [
                "DEPT-A", emp_id, name, "REF-0000", code,
                datetime.combine(end, time()), datetime.combine(start, time()),
                datetime.combine(start, time()), start_time or time(0), time(0),
                days, days * 8,
            ]


def main() -> int:
    raw_path = config.DATA_RAW / "leave_export_20260701_080000.xlsx"
    raw_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for row in build_rows():
        ws.append(row)
    wb.save(raw_path)
    print(f"Wrote synthetic export -> {raw_path}")

    count = convert(config.CALENDAR_MD, config.CALENDAR_FILE)
    print(f"Wrote calendar ({count} rows) -> {config.CALENDAR_FILE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
