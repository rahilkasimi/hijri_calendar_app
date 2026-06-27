"""
calendar_builder.py
===================
Assembles a complete Fatimid month view (the thing the GUI draws) out of the
pure-Python engine in :mod:`calendar_core` plus the three static data tables
extracted verbatim from the workbook:

    data/timedb.json  -> the hidden "Time database" sheet (Sunrise/Noon/Sunset
                          per Gregorian month-day for every supported city).
    data/cities.json  -> the "Fatimi Calendar"!IH:IK table mapping a city name
                          to its Sunrise/Noon/Sunset column offset.
    data/events.json  -> the hidden "events" sheet keyed by Hijri month-day.

PRAYER-TIME RULES (replicated from month sheet '6' + Fatimi Calendar cells):
    SR (Sunrise) = Time database value for the city / Gregorian day.
    FJ (Fajr)    = SR  -  'Fatimi Calendar'!IA11   (= 1 h 15 m before sunrise).
    NN (Noon)    = Time database "Noon" value for the city / Gregorian day.
    ZW (Zawal)   = NN  +  'Fatimi Calendar'!IA13   (= +0 min  ->  ZW == NN).
    MG (Maghrib) = Time database "Sunset" value for the city / Gregorian day.

EVENT RULES (replicated from the 'events' sheet + month-sheet SUMIF logic):
    Events are keyed by (Hijri month, Hijri day).  Rows whose occasion text is
    "No Event Recorded" are skipped.  Each day can carry several events
    (the sheet's "C" sub-index); the GUI shows the per-day count + a master list.
"""

import json
import os
import datetime

import calendar_core as cc

_DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

with open(os.path.join(_DATA, "timedb.json"), encoding="utf-8") as _f:
    TIMEDB = json.load(_f)
with open(os.path.join(_DATA, "cities.json"), encoding="utf-8") as _f:
    CITIES = json.load(_f)
with open(os.path.join(_DATA, "events.json"), encoding="utf-8") as _f:
    EVENTS = json.load(_f)

# Offsets taken from 'Fatimi Calendar'!IA11 (Fajr) and IA13 (Zawal).
FAJR_BEFORE_SUNRISE_MIN = 75      # IA11 = 01:15
ZAWAL_AFTER_NOON_MIN = 0          # IA13 = 00:00

CITY_NAMES = list(CITIES.keys())
WEEKDAYS = ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"]


def _shift(hhmm, minutes):
    """Add/subtract whole minutes to a 'HH:MM' string, wrapping at 24 h."""
    if not hhmm:
        return ""
    h, m = map(int, hhmm.split(":"))
    total = (h * 60 + m + minutes) % (24 * 60)
    return f"{total // 60:02d}:{total % 60:02d}"


def prayer_times(city, greg_year, greg_month, greg_day):
    """Return the 5 displayed times {FJ,SR,NN,ZW,MG} for a Gregorian date.

    Mirrors the month-sheet VLOOKUP into 'Time database' (keyed by month-day,
    the workbook always uses the 2014 reference row set) for the chosen city.
    """
    offs = CITIES.get(city)
    row = TIMEDB.get(f"{greg_month}-{greg_day}")
    if not offs or not row:
        return {k: "" for k in ("FJ", "SR", "NN", "ZW", "MG")}
    sr = row.get(str(offs["sunrise"]), "")
    nn = row.get(str(offs["noon"]), "")
    mg = row.get(str(offs["sunset"]), "")
    return {
        "FJ": _shift(sr, -FAJR_BEFORE_SUNRISE_MIN),   # SR - 1:15
        "SR": sr,
        "NN": nn,
        "ZW": _shift(nn, ZAWAL_AFTER_NOON_MIN),       # NN + 0
        "MG": mg,
    }


def day_events(hijri_month, hijri_day):
    """List of event dicts for a Hijri (month, day) - empty if none recorded."""
    return EVENTS.get(f"{hijri_month}-{hijri_day}", [])


class DayCell:
    """One day in the month grid - holds both calendars, prayer times, events."""

    def __init__(self, hy, hm, hd, city):
        self.hijri_year = hy
        self.hijri_month = hm           # 1..12
        self.hijri_day = hd             # 1..30
        gy, gm, gd = cc.hijri_to_gregorian(hy, hm, hd)
        self.greg_year = gy
        self.greg_month = gm            # 1..12
        self.greg_day = gd
        self.weekday = cc.weekday_of_gregorian(gy, gm, gd)   # 0=Sun
        self.times = prayer_times(city, gy, gm, gd)
        self.events = day_events(hm, hd)

    @property
    def event_count(self):
        return len(self.events)

    @property
    def greg_month_name(self):
        return cc.GREG_MONTH_NAMES[self.greg_month - 1]


def build_month(hijri_year, hijri_month, city="Qatar"):
    """Build the full data model for one Fatimid month.

    Returns a dict with:
        title           "Muharram 1448"
        year_type       "Kabisa" / "Basita"
        days            list[DayCell]  (1 .. month length)
        weeks           6x7 grid of DayCell|None laid out Sun..Sat
        greg_span       e.g. "June - July 2026"
        events          flat list of (code, day, occasion) for the month
    """
    month_len = cc.hijri_month_length(hijri_year, hijri_month)
    days = [DayCell(hijri_year, hijri_month, d, city) for d in range(1, month_len + 1)]

    # ---- lay the days into a Sun..Sat grid -------------------------------
    weeks = []
    week = [None] * 7
    for cell in days:
        week[cell.weekday] = cell
        if cell.weekday == 6:           # Saturday -> close the row
            weeks.append(week)
            week = [None] * 7
    if any(w is not None for w in week):
        weeks.append(week)

    # ---- Gregorian span string -------------------------------------------
    first, last = days[0], days[-1]
    if first.greg_month == last.greg_month:
        span = f"{first.greg_month_name} {first.greg_year}"
    elif first.greg_year == last.greg_year:
        span = f"{first.greg_month_name} - {last.greg_month_name} {last.greg_year}"
    else:
        span = (f"{first.greg_month_name} {first.greg_year} - "
                f"{last.greg_month_name} {last.greg_year}")

    # ---- month event list (sorted by Hijri day then sub-index) -----------
    month_events = []
    for d in range(1, month_len + 1):
        for ev in day_events(hijri_month, d):
            month_events.append({
                "code": ev.get("code") or "",
                "day": d,
                "occasion": ev.get("occasion") or "",
            })

    return {
        "hijri_year": hijri_year,
        "hijri_month": hijri_month,
        "title": f"{cc.HIJRI_MONTH_NAMES[hijri_month - 1]} {hijri_year}",
        "year_type": "کبیسَہ" if cc.is_kabisa(hijri_year) else "Basita",
        "month_len": month_len,
        "days": days,
        "weeks": weeks,
        "greg_span": span,
        "events": month_events,
        "city": city,
    }


if __name__ == "__main__":
    m = build_month(1448, 1, "Qatar")
    print(m["title"], "|", m["year_type"], "|", m["greg_span"], "| days:", m["month_len"])
    for cell in m["days"][:6]:
        print(f"  H{cell.hijri_day:>2}  G{cell.greg_day:>2} {cell.greg_month_name[:3]}  "
              f"{WEEKDAYS[cell.weekday][:3]}  FJ {cell.times['FJ']}  SR {cell.times['SR']}  "
              f"MG {cell.times['MG']}  events={cell.event_count}")
    print("Total month events:", len(m["events"]))
