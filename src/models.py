"""Data model for leave records and detected sequences."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Optional


@dataclass
class Record:
    """One leave row from the raw export, normalized."""

    emp_id: str
    emp_name: str
    leave_code: str
    start_date: date
    end_date: date
    start_time: Optional[time]
    end_time: Optional[time]
    is_target: bool


@dataclass
class Sequence:
    """A qualifying continuous target-leave sequence, ready for output."""

    emp_id: str
    emp_name: str
    start_date: date
    end_date: date
    consecutive_days: int
    tokens: list[str]  # interleaved leave codes and bridge labels (weekend/holiday)
