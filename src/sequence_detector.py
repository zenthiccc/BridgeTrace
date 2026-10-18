"""Core logic: detect continuous target-leave sequences.

A qualifying sequence:
  * begins with a target leave (TL) record,
  * ends with a TL record,
  * every day in between is bridged by other leave / weekend / holiday
    (no ordinary worked day, and no substitute working day, in the gaps), and
  * contains at least one bridge (a pure consecutive TL->TL run is excluded).

AM/PM rule: when a TL resumes *after a bridge*, it stays continuous only if it
starts before AM_PM_CUTOFF. Starting at/after the cutoff means the employee
worked that morning, which breaks continuity.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import timedelta
from typing import Optional

import config
from calendar_utils import WorkCalendar
from models import Record, Sequence


def _gap_is_bridged(prev: Record, cur: Record, cal: WorkCalendar) -> tuple[bool, bool]:
    """Return (bridged, gap_is_bridge).

    bridged        - True if no worked day separates the two records.
    gap_is_bridge  - True if a non-empty non-working gap (weekend/holiday) lies between them.
    """
    if cur.start_date <= prev.end_date:
        return True, False
    if cal.has_working_day_between(prev.end_date, cur.start_date):
        return False, False
    return True, (cur.start_date - prev.end_date).days > 1


def _resumes_in_pm(rec: Record) -> bool:
    if rec.start_time is None:
        return False  # no time recorded -> treat as a full day
    return rec.start_time >= config.AM_PM_CUTOFF


def _gap_bridge_tokens(prev: Record, cur: Record, cal: WorkCalendar) -> list[str]:
    """Labels for the non-working days bridging two records, consecutive duplicates collapsed."""
    labels: list[str] = []
    d = prev.end_date + timedelta(days=1)
    while d < cur.start_date:
        if d in cal.non_working:
            label = "holiday"
        elif d.weekday() >= 5:
            label = "weekend"
        else:
            label = None
        if label and (not labels or labels[-1] != label):
            labels.append(label)
        d += timedelta(days=1)
    return labels


def _finalize(chain: list[Record], cal: WorkCalendar) -> Optional[Sequence]:
    """Trim to end on a TL record and emit a Sequence if it qualifies."""
    last_target = next((i for i in range(len(chain) - 1, -1, -1) if chain[i].is_target), None)
    if last_target is None:
        return None
    chain = chain[: last_target + 1]
    if len(chain) < 2:
        return None

    has_bridge = any(
        not chain[k].is_target or _gap_is_bridged(chain[k - 1], chain[k], cal)[1]
        for k in range(1, len(chain))
    )
    if not has_bridge:
        return None

    first, last = chain[0], chain[-1]
    tokens: list[str] = [first.leave_code]
    for k in range(1, len(chain)):
        tokens.extend(_gap_bridge_tokens(chain[k - 1], chain[k], cal))
        tokens.append(chain[k].leave_code)

    return Sequence(
        emp_id=first.emp_id,
        emp_name=first.emp_name,
        start_date=first.start_date,
        end_date=last.end_date,
        consecutive_days=(last.end_date - first.start_date).days + 1,
        tokens=tokens,
    )


def detect_for_employee(records: list[Record], cal: WorkCalendar) -> list[Sequence]:
    """Detect qualifying sequences for a single employee's records."""
    records = sorted(records, key=lambda r: (r.start_date, r.start_time or config.AM_PM_CUTOFF.min))
    sequences: list[Sequence] = []

    i, n = 0, len(records)
    while i < n:
        if not records[i].is_target:
            i += 1
            continue

        chain = [records[i]]
        j = i + 1
        while j < n:
            prev, cur = chain[-1], records[j]
            bridged, gap_is_bridge = _gap_is_bridged(prev, cur, cal)
            if not bridged:
                break
            crosses_bridge = (not prev.is_target) or gap_is_bridge
            if cur.is_target and crosses_bridge and _resumes_in_pm(cur):
                break
            chain.append(cur)
            j += 1

        seq = _finalize(chain, cal)
        if seq is not None:
            sequences.append(seq)
        i = j if j > i else i + 1

    return sequences


def detect_all(records: list[Record], cal: WorkCalendar) -> list[Sequence]:
    """Detect qualifying sequences across all employees."""
    by_emp: dict[str, list[Record]] = defaultdict(list)
    for rec in records:
        by_emp[rec.emp_id].append(rec)

    results: list[Sequence] = []
    for emp_records in by_emp.values():
        results.extend(detect_for_employee(emp_records, cal))
    results.sort(key=lambda s: (s.emp_id, s.start_date))
    return results
