"""Validate the detector against the include/exclude patterns in the README."""
from datetime import date, time

import config
from calendar_utils import WorkCalendar
from models import Record
from sequence_detector import detect_for_employee

CAL = WorkCalendar.from_markdown(config.CALENDAR_MD)

TL = "SL"  # target leave (sick leave)
AL = "AL"  # annual leave (bridge)


def rec(code, start, end=None, start_time=None):
    return Record(
        emp_id="E1",
        emp_name="Test",
        leave_code=code,
        start_date=start,
        end_date=end or start,
        start_time=start_time,
        end_time=None,
        is_target=code in config.TARGET_LEAVE_CODES,
    )


def detect(records):
    return detect_for_employee(records, CAL)


# --- INCLUDED patterns -----------------------------------------------------

def test_tl_al_tl_am_included():
    records = [
        rec(TL, date(2026, 1, 13)),
        rec(AL, date(2026, 1, 14)),
        rec(TL, date(2026, 1, 15), start_time=time(9, 0)),
    ]
    seqs = detect(records)
    assert len(seqs) == 1
    assert seqs[0].tokens == [TL, AL, TL]


def test_tl_weekend_tl_included():
    seqs = detect([rec(TL, date(2026, 1, 9)), rec(TL, date(2026, 1, 12))])  # Fri, Mon
    assert len(seqs) == 1
    assert seqs[0].consecutive_days == 4
    assert seqs[0].tokens == [TL, "weekend", TL]


def test_tl_holiday_tl_included():
    seqs = detect([rec(TL, date(2026, 6, 16)), rec(TL, date(2026, 6, 18))])  # Tue, holiday Wed, Thu
    assert len(seqs) == 1
    assert seqs[0].tokens == [TL, "holiday", TL]


# --- EXCLUDED patterns -----------------------------------------------------

def test_al_tl_excluded():
    assert detect([rec(AL, date(2026, 1, 6)), rec(TL, date(2026, 1, 7))]) == []


def test_tl_al_tl_pm_excluded():
    records = [
        rec(TL, date(2026, 1, 13)),
        rec(AL, date(2026, 1, 14)),
        rec(TL, date(2026, 1, 15), start_time=time(14, 0)),
    ]
    assert detect(records) == []


def test_pure_consecutive_target_excluded():
    assert detect([rec(TL, date(2026, 1, 13)), rec(TL, date(2026, 1, 14))]) == []


def test_worked_day_gap_breaks_chain():
    records = [
        rec(TL, date(2026, 1, 12)),
        rec(AL, date(2026, 1, 13)),
        rec(TL, date(2026, 1, 15)),  # Wed 14 is a normal working day
    ]
    assert detect(records) == []


def test_substitute_working_day_breaks_chain():
    # Fri 19 Jun, Sat 20 Jun is a substitute working day, Mon 22 Jun
    assert detect([rec(TL, date(2026, 6, 19)), rec(TL, date(2026, 6, 22))]) == []


# --- Calendar sanity -------------------------------------------------------

def test_calendar_classification():
    assert CAL.is_working_day(date(2026, 6, 20)) is True    # substitute working Saturday
    assert CAL.is_working_day(date(2026, 6, 17)) is False   # company holiday
    assert CAL.is_working_day(date(2026, 1, 10)) is False   # ordinary Saturday
    assert CAL.year == 2026
