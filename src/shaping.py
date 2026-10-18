"""An Excel Power Query (see scripts/shaping_query.m) reproduced with pandas.

Keeps the configured columns in order, sets their types, appends blank tracking
columns, keeps only target-leave rows, and sorts newest change first.
"""
from __future__ import annotations

import pandas as pd

import config
from load_data import norm_code, parse_date, parse_time


def _convert(series: pd.Series, kind: str) -> pd.Series:
    """Coerce a column like Power Query's "Changed Type" step; bad values become blank."""
    if kind == "int":
        return pd.to_numeric(series, errors="coerce").round().astype("Int64")
    if kind == "number":
        return pd.to_numeric(series, errors="coerce")
    if kind == "date":
        return pd.to_datetime(series.map(parse_date))
    if kind == "time":
        return series.map(parse_time)
    return series.map(lambda v: norm_code(v) or None)


def transform(raw: pd.DataFrame) -> pd.DataFrame:
    """Apply the shaping steps to the raw export and return the resulting table."""
    missing = [c for c in config.SHAPE_COLUMNS if c not in raw.columns]
    if missing:
        raise ValueError(f"Raw file is missing columns used by the shaping step: {missing}")

    df = raw[list(config.SHAPE_COLUMNS)].copy()
    for column, kind in config.SHAPE_COLUMNS.items():
        df[column] = _convert(df[column], kind)
    for column in config.SHAPE_ADDED_COLUMNS:
        df[column] = None

    df = df[df[config.COL_LEAVE_CODE].isin(config.TARGET_LEAVE_CODES)]
    return df.sort_values(
        config.SHAPE_SORT_COLUMN, ascending=False, kind="stable", na_position="last"
    ).reset_index(drop=True)
