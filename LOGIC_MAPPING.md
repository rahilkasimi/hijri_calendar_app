# Logic Mapping — Excel → Python

This document maps **every** calculation in `CALENDAR (Updated).xlsx` to its
pure-Python equivalent. The workbook is the **source of truth**; the Python
code reproduces its lookup tables and formula chain exactly, with **no external
calendar/conversion library**.

The heavy lifting happens on the hidden **`Calculations`** worksheet. The
visible month sheets (`1`…`12`), `Fatimi Calendar`, `Date Converter`,
`Time database` and `events` only *reference* that engine.

---

## 1. The core idea — a shared "serial ordinal day"

Both calendars are expressed as a single integer: the **ordinal day number**
counted from a common origin. Converting between calendars therefore means:

```
source calendar fields  ->  ordinal integer  ->  target calendar fields
```

Two parallel ordinal systems are built on the `Calculations` sheet and aligned
by the epoch constant **8231** (`Calculations!C5`), so one integer can be read
in either calendar.

| Concept | Excel location | Python (`calendar_core.py`) |
|---|---|---|
| Fatimi ordinal tables | `Calculations!A4:L34` | `HIJRI_CYCLE_*`, `HIJRI_YEAR_*`, `HIJRI_MONTH_*` |
| Gregorian ordinal tables | `Calculations!N4:Y66` | `GREG_CENTURY_*`, `GREG_4YR_*`, `GREG_MONTH_*` |
| Greg → Hijri formula block | `Calculations!AH6:AH33` | `gregorian_to_hijri()` |
| Hijri → Greg formula block | `Calculations!AG37:AH70` | `hijri_to_gregorian()` |

---

## 2. Fatimi (Hijri) tables — `Calculations!A4:L34`

| Excel column | Meaning | Python symbol |
|---|---|---|
| `B5:B65` | cycle start year `(cycle-1)*30` | `HIJRI_CYCLE_YEARS` |
| `C5:C65` | ordinal day at start of cycle, `8231 + 10631*(cycle-1)` | `HIJRI_CYCLE_ORD` |
| `F5:F34` | length of each year in a 30-year cycle (354 *Basita* / 355 *Kabisa*) | `HIJRI_YEAR_LEN` |
| `G5:G34` | cumulative days from cycle start to start of that year | `HIJRI_YEAR_CUM` |
| `K5:K16` | month length (30 / 29 alternating, last month leap-aware) | `HIJRI_MONTH_LEN` |
| `L5:L16` | cumulative days from year start to start of month | `HIJRI_MONTH_CUM` |

**Leap (Kabisa) years**: cycle positions `2, 5, 8, 10, 13, 16, 19, 21, 24, 27,
29, 30` are 355 days. This pattern is encoded directly in `HIJRI_YEAR_LEN`
(copied from `F5:F34`), so 1448 AH is correctly detected as *Kabisa* — exactly
as the workbook shows.

Each 30-year cycle spans **10631 days** (`C(n+1) − C(n)`), matching the
classic tabular Islamic calendar.

---

## 3. Gregorian tables — `Calculations!N4:Y66`

| Excel column | Meaning | Python symbol |
|---|---|---|
| `O/P` (`N4:P…`) | century base year → ordinal at century start | `GREG_CENTURY_YEAR`, `GREG_CENTURY_ORD` |
| `S/T` (`R4:T28`) | 4-year block → cumulative days (1461 = 365×3 + 366) | `GREG_4YR_*` |
| `W/X/Y` (`V18:Y66`) | 48-month table → day-of-block offset + year-position | `GREG_MONTH_*` |
| `W4:AB16` | month grid: day-of-year offset to the 1st of each month for year-positions 0/1/2/3 | `GREG_MONTH_GRID` |

The **1582 Gregorian reform** is handled implicitly by the hard-coded century
ordinals (e.g. the jump around `358665 / 365240`). No special-case code is
needed — the table values already absorb the 10 skipped days, exactly as the
workbook stores them.

---

## 4. The `MATCH(…, 1)` decode step

Throughout the workbook, decoding an ordinal back into calendar fields uses:

```excel
MATCH(value, ascending_column, 1)   ' position of the LARGEST value <= `value`
```

plus the workbook's recurring boundary correction: *"if the remainder is 0,
step back one row."* Both are replicated by the helper:

```python
def _match_le(value, table):
    """Excel MATCH(value, table, 1): index of largest table[i] <= value."""
```

and the explicit "remainder == 0 → previous row" logic inside
`hijri_to_ordinal()` / `ordinal_to_hijri()` (mirrors the `AG41` boundary cell).

> **Bug we had to reproduce faithfully:** an off-by-one originally appeared
> because Excel's cumulative columns are 1-based against a 0-based Python list.
> Fixed by indexing `HIJRI_YEAR_CUM[year_in_cycle-1]` /
> `HIJRI_YEAR_LEN[year_in_cycle-1]` with the same boundary rule the sheet uses.

---

## 5. Gregorian → Hijri — `Calculations!AH6:AH33`

```
1. ordinal = GREG_CENTURY_ORD[match] + GREG_4YR_T[match]
             + GREG_MONTH_GRID[year-pos][month] + day        (AH6:AH20)
2. subtract Fatimi epoch and decode with TABLE 1:
     cycle  = _match_le(ordinal, HIJRI_CYCLE_ORD)            (AH22:AH26)
     year   = _match_le(rem, HIJRI_YEAR_CUM) within cycle    (AH27:AH30)
     month  = _match_le(rem, HIJRI_MONTH_CUM)                (AH31:AH33)
     day    = remainder + 1
```
→ `calendar_core.gregorian_to_hijri(y, m, d)`

## 6. Hijri → Gregorian — `Calculations!AG37:AH70`

The mirror image: build the ordinal from the Fatimi tables, then decode with
the Gregorian century / 4-year / month-grid tables.
→ `calendar_core.hijri_to_gregorian(y, m, d)`

The **day-of-week** is `ordinal mod 7` aligned to the workbook's weekday cell,
which is why H 1448-01-01 lands on **Monday 15 June 2026** — identical to the
reference screenshot.

---

## 7. Prayer times — month sheets + `Fatimi Calendar`

Source values come from the hidden **`Time database`** sheet (Sunrise / Noon /
Sunset per Gregorian month-day for each city). The `Fatimi Calendar`!`IH:IK`
table maps a **city name → column offset** into that database.

| Prayer | Excel rule | Python (`calendar_builder.py`) |
|---|---|---|
| **SR** Sunrise | `Time database` lookup | `db["sunrise"]` |
| **FJ** Fajr | `SR − 'Fatimi Calendar'!IA11` (= 01:15 before sunrise) | `sunrise − timedelta(1h15m)` |
| **NN** Noon | `Time database` lookup | `db["noon"]` |
| **ZW** Zawal | `NN + 'Fatimi Calendar'!IA13` (= +00:00 → `ZW == NN`) | `noon + timedelta(0)` |
| **MG** Maghrib | `Time database` lookup (Sunset) | `db["sunset"]` |

The Time DB is keyed by Gregorian **month-day** (the stored year 2014 is a
constant placeholder and is ignored).

### The documented Sunday-noon quirk
In the original workbook, the **Noon/Zawal** cells for **Sundays** use an
absolute reference `$G$8` (a copy-paste artefact) instead of the row-relative
day, which makes those 21 cells differ by ≤4 minutes from the correct per-day
value. The Python engine performs the **correct per-day lookup**, so it differs
from Excel in exactly those 21 Sunday Noon/Zawal cells. Validation:

```
Day conversions : 294 / 294 cells exact   (all 12 cached month sheets, 1443 AH)
Prayer times    : 1449 / 1470 cells exact (21 diffs = the Sunday $G$8 quirk)
```

---

## 8. Events — `events` sheet + month-sheet `SUMIF`

Events are keyed by **(Hijri month, Hijri day)**. Rows whose occasion text is
`"No Event Recorded"` are skipped. A single day may carry several events (the
sheet's `C` sub-index); the GUI shows a per-day **count** plus a master list,
mirroring the month sheet's `SUMIF`/`COUNTIF` "No. of Events" cell.
→ `calendar_builder.build_month()` attaches events to each day.
