"""Configuration: leave codes, AM/PM cutoff, column headers, and file paths.

All values here are generic placeholders. Adapt them to your own HR system export.
"""
from __future__ import annotations

from datetime import time
from pathlib import Path

# --- Project paths ---------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
DATA_RAW = ROOT / "data" / "raw"
DATA_OUTPUT = ROOT / "data" / "output"
# Human-readable calendar source and the generated xlsx the pipeline reads.
CALENDAR_MD = ROOT / "samples" / "sample_calendar.md"
CALENDAR_FILE = ROOT / "samples" / "sample_calendar.xlsx"

# Calendar xlsx column headers (the two the pipeline relies on).
CAL_COL_DATE = "Date"
CAL_COL_STATUS = "Status"

# --- Business rules --------------------------------------------------------
# Leave codes that count as the *target* leave (sick leave). Everything else is
# "other leave" and can bridge two target-leave periods.
TARGET_LEAVE_CODES = {"SL", "SLH", "SLX"}

# A target leave resuming after a bridge stays continuous only if it starts
# before this time (AM). Start time >= cutoff (PM) breaks continuity.
AM_PM_CUTOFF = time(12, 0)

# --- Raw export column headers ---------------------------------------------
COL_EMP_ID = "Employee ID"
COL_EMP_NAME = "Employee Name"
COL_LEAVE_CODE = "Leave Code"
COL_START_DATE = "Start Date"
COL_END_DATE = "End Date"
COL_START_TIME = "Start Time"
COL_END_TIME = "End Time"

REQUIRED_COLUMNS = [
    COL_EMP_ID,
    COL_EMP_NAME,
    COL_LEAVE_CODE,
    COL_START_DATE,
    COL_END_DATE,
    COL_START_TIME,
    COL_END_TIME,
]

# --- Output column headers (continuous-leave sheet) --------------------------
OUT_EMP_ID = "Employee ID"
OUT_EMP_NAME = "Employee Name"
OUT_START_DATE = "Start Date"
OUT_END_DATE = "End Date"
OUT_CONSECUTIVE_DAYS = "Number of consecutive days"
OUT_LEAVE_CHAIN = "Leave types in the sequence"

# --- Report workbook -------------------------------------------------------
# Named from the raw export's date: leave_export_20260803.xlsx ->
# "Continuous Leave Report 2026-08-03.xlsx".
REPORT_FILENAME = "Continuous Leave Report {:%Y-%m-%d}.xlsx"
SHEET_FILTERED = "Target Leaves"         # shaped/filtered export (first sheet)
SHEET_CONTINUOUS = "Continuous Leave"    # detected sequences (second sheet)
FILTERED_TABLE = "TargetLeaves"          # Excel table name on the first sheet

# --- Shaping step (replicates an Excel Power Query) -------------------------
# Columns kept, in output order, with the type the query's "Changed Type" sets.
SHAPE_COLUMNS = {
    "Department": "text",
    "Employee ID": "text",
    "Employee Name": "text",
    "Leave Code": "text",
    "Start Date": "date",
    "End Date": "date",
    "Last Modified": "date",
    "Start Time": "time",
    "End Time": "time",
    "Leave Days": "number",
    "Leave Hours": "number",
}
# Blank columns appended to be filled in manually.
SHAPE_ADDED_COLUMNS = ["Supporting Documents", "Reviewer Notes"]
# Most recently changed records first.
SHAPE_SORT_COLUMN = "Last Modified"

# --- Leave code -> display text --------------------------------------------
LEAVE_TEXT = {
    "SL": "Sick Leave",
    "SLH": "Sick Leave (Half Day)",
    "SLX": "Extended Sick Leave",
    "AL": "Annual Leave",
    "PL": "Personal Leave",
    "FL": "Family Leave",
    "BL": "Bereavement Leave",
    "OL": "Other Leave",
}
