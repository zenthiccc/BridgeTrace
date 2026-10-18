# BridgeTrace

*Continuous-leave sequence detection engine.*

A Python desktop tool that turns a raw **leave export** (Excel) into a
two-sheet review workbook:

1. **Target Leaves** — the export shaped and filtered the way it was previously done
   by hand with an Excel Power Query (kept columns, typed values, tracking columns,
   newest changes first).
2. **Continuous Leave** — only the cases where a **target leave type** (e.g. sick leave)
   is used **continuously** with other leave types, weekends, or public holidays
   bridging two target-leave periods.

The business question behind sheet 2: *"Is another leave type (or a non-working day)
being used in between, consecutively, with sick leave?"*

> This is a generic reference implementation. All leave codes, column names, calendar
> entries, employees, and sample data are fictional placeholders.

See [ARCHITECTURE.md](ARCHITECTURE.md) for the design, data flow, and core algorithm.

---

## 1. What counts as "continuous"

A sequence qualifies when it **starts** with a target leave (TL), **ends** with a TL,
and every day in between is covered by a *bridge* (another leave type, a weekend, or a
holiday) — with no ordinary worked day breaking the chain.

| Pattern | Result | Reason |
| ------- | ------ | ------ |
| `TL → AL → TL` | Included | Annual leave bridges two target leaves |
| `TL → holiday → TL` | Included | Holiday bridges two target leaves |
| `TL → weekend → TL` | Included | Weekend bridges two target leaves |
| `TL → AL → TL (AM)` | Included | Target leave resumes in the morning → never returned to work |
| `AL → TL` | Excluded | Does not start with a target leave |
| `TL → AL → TL (PM)` | Excluded | Employee worked the morning before resuming → continuity broken |
| `TL → TL` (adjacent weekdays) | Excluded | No bridge — just one long absence |
| `TL → substitute working day → TL` | Excluded | A weekend made into a working day counts as worked |

**AM/PM rule:** when a target leave resumes after a bridge, its start time decides —
before 12:00 = AM (continuous), 12:00 or later = PM (break).

---

## 2. Inputs

**Leave export (`.xlsx`)** — one row per leave record. Required columns (configurable in
[src/config.py](src/config.py)):

`Employee ID`, `Employee Name`, `Leave Code`, `Start Date`, `End Date`, `Start Time`,
`End Time` — plus the columns kept on the first sheet (`Department`, `Last Modified`,
`Leave Days`, `Leave Hours`).

**Working calendar (`.xlsx`)** — columns `Date` and `Status`:

| Status | Working day? | Acts as bridge? |
| ------ | ------------ | --------------- |
| `Public holiday` / `Holiday` / `Company holiday` | No | Yes |
| Unlisted Sat/Sun | No | Yes |
| `Substitute working day` | **Yes** | **No** |
| Unlisted Mon–Fri | Yes | No |

The calendar can be authored as markdown ([samples/sample_calendar.md](samples/sample_calendar.md))
and converted with `python src/convert_calendar.py`.

---

## 3. Output

One workbook per export, named from the `YYYYMMDD` in the export's file name
(e.g. `leave_export_20260701_080000.xlsx` → `Continuous Leave Report 2026-07-01.xlsx`).

**Sheet `Continuous Leave`:**

| Column | Description |
| ------ | ----------- |
| `Employee ID` / `Employee Name` | Who |
| `Start Date` / `End Date` | Span of the continuous sequence |
| `Number of consecutive days` | Calendar days in the span |
| `Leave types in the sequence` | e.g. `Sick Leave -> Annual Leave -> Weekend -> Sick Leave` |

---

## 4. Quick start (synthetic demo)

```powershell
pip install -r requirements.txt
python scripts/make_sample_data.py      # fake export + calendar
python src/main.py                       # CLI: newest export in data/raw -> data/output
python src/gui.py                        # or the desktop GUI
pytest                                   # unit + end-to-end tests
```

### Standalone Windows executable (no Python install required)

```powershell
pip install -r requirements.txt -r requirements-build.txt
pyinstaller build_exe.spec --noconfirm --clean   # -> dist/BridgeTrace.exe
```

---

## 5. Project structure

```
bridgetrace/
├─ README.md
├─ ARCHITECTURE.md
├─ build_exe.spec               # PyInstaller one-file, windowed build
├─ requirements.txt / requirements-build.txt
├─ samples/
│  └─ sample_calendar.md        # fictional working calendar (markdown source)
├─ scripts/
│  ├─ make_sample_data.py       # synthetic export + calendar for demos
│  └─ shaping_query.m           # the original Power Query (generic form)
├─ data/raw, data/output        # inputs / generated reports (git-ignored)
├─ src/
│  ├─ config.py                 # codes, headers, AM/PM cutoff, report names, paths
│  ├─ models.py                 # Record & Sequence dataclasses
│  ├─ load_data.py              # read & normalize raw rows
│  ├─ shaping.py                # Power Query steps in pandas -> sheet 1
│  ├─ calendar_utils.py         # WorkCalendar: working-day rules
│  ├─ convert_calendar.py       # calendar markdown -> xlsx
│  ├─ sequence_detector.py      # continuous-leave detection + AM/PM rule
│  ├─ main.py                   # run_pipeline(), report writer, CLI
│  └─ gui.py                    # customtkinter GUI
└─ tests/
   ├─ test_sequences.py         # include/exclude patterns
   └─ test_report.py            # shaping, naming, end-to-end workbook
```

## 6. Tech stack

Python 3.10+, pandas, openpyxl, customtkinter, pytest, PyInstaller.
