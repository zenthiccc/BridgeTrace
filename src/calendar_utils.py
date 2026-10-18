"""Working-calendar helpers.

Reads a calendar xlsx (columns `Date` and `Status`) to decide whether a date is a
working day. Substitute working days override normal weekend status and count as
*worked* days. A markdown parser is also kept so the xlsx can be regenerated from a
human-readable source (see convert_calendar.py).
"""
from __future__ import annotations

import re
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

import config

# Matches rows like: | 2026-01-01 | Thursday | Public holiday | New Year's Day |
_ROW_RE = re.compile(
    r"\|\s*(\d{4})-(\d{2})-(\d{2})\s*\|\s*[A-Za-z]+\s*\|\s*([^|]+?)\s*\|"
)


def _classify(status: str, non_working: set[date], substitute_working: set[date], d: date) -> None:
    status_l = status.strip().lower()
    if "substitute" in status_l:
        substitute_working.add(d)
    elif "holiday" in status_l:  # public holiday / holiday / company holiday
        non_working.add(d)


class WorkCalendar:
    def __init__(self, non_working: set[date], substitute_working: set[date]):
        self.non_working = non_working
        self.substitute_working = substitute_working

    @classmethod
    def from_xlsx(cls, path: str | Path) -> "WorkCalendar":
        df = pd.read_excel(path)
        cols = {str(c).strip().lower(): c for c in df.columns}
        date_col = cols.get(config.CAL_COL_DATE.lower())
        status_col = cols.get(config.CAL_COL_STATUS.lower())
        if date_col is None or status_col is None:
            raise ValueError(
                f"Calendar file must have '{config.CAL_COL_DATE}' and "
                f"'{config.CAL_COL_STATUS}' columns. Found: {list(df.columns)}"
            )

        non_working: set[date] = set()
        substitute_working: set[date] = set()
        for _, row in df.iterrows():
            dval, sval = row[date_col], row[status_col]
            if pd.isna(dval) or pd.isna(sval):
                continue
            _classify(str(sval), non_working, substitute_working, pd.Timestamp(dval).date())
        return cls(non_working, substitute_working)

    @classmethod
    def from_markdown(cls, path: str | Path) -> "WorkCalendar":
        text = Path(path).read_text(encoding="utf-8")
        non_working: set[date] = set()
        substitute_working: set[date] = set()
        for year, month, day, status in _ROW_RE.findall(text):
            _classify(status, non_working, substitute_working, date(int(year), int(month), int(day)))
        return cls(non_working, substitute_working)

    @property
    def year(self) -> int | None:
        """The calendar year, inferred from its dates (None if the calendar is empty)."""
        dates = self.non_working | self.substitute_working
        return min(dates).year if dates else None

    def is_working_day(self, d: date) -> bool:
        """True if `d` is a normal working day (Mon-Fri or a substitute working day)."""
        if d in self.substitute_working:
            return True
        if d in self.non_working:
            return False
        return d.weekday() < 5

    def has_working_day_between(self, start_exclusive: date, end_exclusive: date) -> bool:
        """True if any working day lies strictly between the two dates."""
        d = start_exclusive + timedelta(days=1)
        while d < end_exclusive:
            if self.is_working_day(d):
                return True
            d += timedelta(days=1)
        return False
