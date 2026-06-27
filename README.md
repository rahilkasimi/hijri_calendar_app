# Fatimid (Misri) Hijri Calendar — Python Edition

A standalone desktop application that reproduces the **Fatimid / Misri Hijri
calendar** workbook (`CALENDAR (Updated).xlsx`) as pure Python.

It shows a full Hijri month grid with the matching Gregorian dates, the five
daily prayer times (Fajr, Sunrise, Noon/Zawal, Maghrib) for a selectable city,
and the religious events/occasions for each day — exactly like the original
Excel sheet.

![App preview](preview.png)

---

## Why this exists

The calendar logic used to live entirely inside Excel formulas spread across a
hidden `Calculations` worksheet, a `Time database`, an `events` sheet and a
`Fatimi Calendar` layout sheet. This project lifts **all** of that logic into
documented, testable Python — **with no external calendar/conversion library**.
The Excel workbook is treated as the **source of truth**, and every table and
formula was copied verbatim. See **[`LOGIC_MAPPING.md`](LOGIC_MAPPING.md)** for
the formula-by-formula mapping.

---

## Quick start

```bash
# 1. (optional) create a virtual environment
python3 -m venv .venv && source .venv/bin/activate

# 2. install the one runtime dependency
pip install -r requirements.txt

# 3. on Linux only, make sure Tk is available
#    sudo apt-get install python3-tk

# 4. run the app
python3 app.py
```

> **This is a desktop GUI application, not a web/localhost server.** It opens a
> native window via Tkinter/FreeSimpleGUI and therefore needs a graphical
> display. On a headless machine (no monitor) you must run it through a desktop
> session, VNC, or an X server such as Xvfb.

---

## Using the app

- **City dropdown** — pick the city; all prayer times update from the embedded
  time database.
- **‹ / ›** — previous / next month. **‹‹ / ››** — previous / next year.
- **Today** — jump to the Hijri month containing today's Gregorian date.
- **Go to** — type a Hijri year + month and jump straight to it.
- Each day cell shows the **Hijri day number** (event days are highlighted in
  the workbook's green/yellow style) and a **blue chip** with the Gregorian day.
- The right-hand panel lists the month's **events** and the per-day count,
  mirroring the workbook's "No. of Events" cell.

---

## Project layout

| File | Purpose |
|---|---|
| `calendar_core.py` | Pure-Python Gregorian ↔ Hijri conversion engine. All lookup tables are copied from `Calculations`; every function carries the Excel cell references it reproduces. |
| `calendar_builder.py` | Builds a full month model (grid, dual dates, prayer times, events) from the core engine + the JSON data tables. |
| `app.py` | The FreeSimpleGUI desktop UI, laid out to mirror the Excel screenshot. |
| `data/cities.json` | City → time-database column mapping (`Fatimi Calendar`!IH:IK). |
| `data/timedb.json` | Sunrise / Noon / Sunset per Gregorian month-day per city (hidden `Time database`). |
| `data/events.json` | Religious events keyed by Hijri month-day (hidden `events` sheet). |
| `LOGIC_MAPPING.md` | Detailed Excel-formula → Python mapping. |
| `requirements.txt` | Runtime dependency (FreeSimpleGUI). |

---

## How the calculation works (summary)

Both calendars are reduced to a single **serial ordinal day number** and matched
through that integer:

```
Gregorian fields  ->  ordinal  ->  Hijri fields      (and the reverse)
```

- **Fatimi tables** (`HIJRI_*` in `calendar_core.py`) come from
  `Calculations!A4:L34`: 30-year cycles of 10631 days, the 354/355-day
  *Basita*/*Kabisa* year pattern, and 30/29-day months.
- **Gregorian tables** (`GREG_*`) come from `Calculations!N4:Y66`: centuries
  (with the 1582 reform baked in), 4-year blocks of 1461 days, and a
  leap-aware month grid.
- Decoding an ordinal uses Excel's `MATCH(value, column, 1)` — the largest
  table value ≤ the lookup — replicated as `_match_le()` with the workbook's
  "step back one row when the remainder is 0" boundary correction.

### Prayer-time rules
```
SR (Sunrise) = Time database value
FJ (Fajr)    = Sunrise − 1:15      ('Fatimi Calendar'!IA11)
NN (Noon)    = Time database value
ZW (Zawal)   = Noon + 0:00         ('Fatimi Calendar'!IA13  ->  ZW == NN)
MG (Maghrib) = Time database value (Sunset)
```

---

## Validation against the original Excel

The engine was checked cell-by-cell against the workbook's 12 cached month
sheets (1443 AH):

| Check | Result |
|---|---|
| Gregorian ↔ Hijri day conversions | **294 / 294 exact** |
| Weekday placement in the grid | **exact** |
| Prayer-time cells | **1449 / 1470 exact** |
| Reference anchor (H 1448-01-01) | **Monday 15 June 2026** — matches the screenshot |

The 21 differing prayer cells are all **Sunday Noon/Zawal** values. They stem
from a copy-paste artefact in the *original* workbook (an absolute `$G$8`
reference) and differ by ≤4 minutes; the Python engine performs the correct
per-day lookup. This is documented in detail in `LOGIC_MAPPING.md` (§7).

---

## Notes

- No `hijri-converter`, `ummalqura`, `convertdate`, or any other calendar
  library is used — only Python's standard library plus the embedded tables.
- The data JSON files were extracted once from the workbook; the app does not
  need Excel or `openpyxl` at runtime.
