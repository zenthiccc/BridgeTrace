"""Load and normalize the raw leave data from an Excel export."""
from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path
from typing import Optional

import pandas as pd

import config
from models import Record


def norm_code(value) -> str:
    """Normalize a code/id cell to a plain string (e.g. 1001.0 -> '1001')."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_date(value) -> Optional[date]:
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=False)
    return None if pd.isna(parsed) else parsed.date()


def parse_time(value) -> Optional[time]:
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time()
    # Excel stores times as a fraction of a day.
    if isinstance(value, (int, float)):
        total_seconds = int(round(float(value) * 24 * 3600)) % (24 * 3600)
        return time(total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60)
    parsed = pd.to_datetime(str(value), errors="coerce")
    return None if pd.isna(parsed) else parsed.time()


def load_records(source: str | Path | pd.DataFrame) -> list[Record]:
    """Return normalized Record objects from the raw workbook (a path or its DataFrame)."""
    df = source if isinstance(source, pd.DataFrame) else pd.read_excel(source)

    missing = [c for c in config.REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Raw file is missing required columns: {missing}. "
            f"Found columns: {list(df.columns)}"
        )

    records: list[Record] = []
    for _, row in df.iterrows():
        code = norm_code(row[config.COL_LEAVE_CODE])
        start_date = parse_date(row[config.COL_START_DATE])
        end_date = parse_date(row[config.COL_END_DATE])

        if not code or start_date is None:
            continue
        if end_date is None:
            end_date = start_date

        records.append(
            Record(
                emp_id=norm_code(row[config.COL_EMP_ID]),
                emp_name=str(row[config.COL_EMP_NAME]).strip(),
                leave_code=code,
                start_date=start_date,
                end_date=end_date,
                start_time=parse_time(row[config.COL_START_TIME]),
                end_time=parse_time(row[config.COL_END_TIME]),
                is_target=code in config.TARGET_LEAVE_CODES,
            )
        )

    return records
