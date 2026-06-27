"""
calendar_core.py
================
Pure-Python re-implementation of the Hijri (Fatimid / Misri) calendar engine that
lives inside ``CALENDAR (Updated).xlsx``.

NOTHING in this module uses an external calendar/conversion library.  Every value
is produced by faithfully re-creating the *exact* lookup tables and the
VLOOKUP / MATCH formula chain found on the hidden ``Calculations`` worksheet of
the workbook (the "source of truth").

--------------------------------------------------------------------------------
EXCEL  ->  PYTHON  LOGIC  MAP  (see LOGIC_MAPPING.md for the full narrative)
--------------------------------------------------------------------------------
The workbook converts a date to a *serial ordinal day number* and back.  Two
parallel ordinal systems are built so a single integer can be matched between
the Gregorian and the Fatimid calendars:

TABLE 1 - FATIMI  (Calculations!A4:L34)
    A/B/C : 30-year cycles.  B = (cycle-1)*30 years,  C = ordinal day at the
            start of that cycle.  Each cycle is 10631 days  (C(n+1)-C(n)).
            C5 = 8231  is the epoch offset that aligns the two ordinal systems.
    E/F/G : the 30 years inside one cycle.  F = length of that Hijri year
            (354 normal "Basita" / 355 leap "Kabisa"),  G = cumulative days
            from the start of the cycle to the start of that year.
    I/K/L : the 12 Hijri months.  K = month length (30/29 alternating),
            L = cumulative days from the start of the year to start of month.

TABLE 2 - GREGORIAN (Calculations!N4:AB16 / R4:T28 / V18:Y66)
    N/O/P : centuries.  O = century base year (600,700,...,1582,1600,...),
            P = ordinal day at the start of that century (handles the 1582
            Gregorian reform via the hard-coded jumps 358665 / 365240 ...).
    R/S/T : 4-year blocks.  S = multiple of 4 years, T = cumulative days
            (1461 = 365*3 + 366 for every 4-year block).
    V/W/X/Y : 48 months spanning a 4-year block.  X = cumulative day-of-block
            at the start of the month,  Y = which of the 4 years (0..3).
    W4:AB16 : "month grid" - day-of-year offset to the 1st of each month for
            year-position 0/1/2/3 inside a 4-year block (leap-year aware).

Gregorian -> Hijri   : build ordinal from Gregorian (a+b+c+d) then decode with
                       TABLE 1.   (Calculations!AH6:AH33)
Hijri -> Gregorian   : build ordinal from Hijri  then decode with TABLE 2.
                       (Calculations!AG37:AH70)

The decode step everywhere uses Excel ``MATCH(value, column, 1)`` = position of
the largest table value <= ``value`` (ascending data) plus the workbook's
"if the remainder is 0 step back one row" boundary correction.
"""

# ===========================================================================
#  EXACT lookup tables copied from the recalculated workbook (data_only read).
#  These are the SAME numbers Excel stores in the Calculations sheet.
# ===========================================================================

# ---- TABLE 1 : FATIMI -----------------------------------------------------
# 30-year cycles.  Index 0 -> cycle 1.
HIJRI_CYCLE_YEARS = [c * 30 for c in range(0, 61)]          # B5:B65  (0,30,60,...)
HIJRI_CYCLE_ORD = [8231 + c * 10631 for c in range(0, 61)]  # C5:C65  (epoch 8231, +10631/cycle)

# Length of each of the 30 Hijri years inside a cycle (F5:F34).
HIJRI_YEAR_LEN = [354, 355, 354, 354, 355, 354, 354, 355, 354, 355,
                  354, 354, 355, 354, 354, 355, 354, 354, 355, 354,
                  355, 354, 354, 355, 354, 354, 355, 354, 355, 355]
# Cumulative days from start of cycle to start of year e  (G5:G34).
HIJRI_YEAR_CUM = []
_acc = 0
for _d in HIJRI_YEAR_LEN:
    HIJRI_YEAR_CUM.append(_acc)
    _acc += _d
# -> [0, 354, 709, 1063, ... , 10277]

# Month lengths (K5:K16) and cumulative day offsets (L5:L16).
HIJRI_MONTH_LEN = [30, 29, 30, 29, 30, 29, 30, 29, 30, 29, 30, 29]
HIJRI_MONTH_CUM = []
_acc = 0
for _d in HIJRI_MONTH_LEN:
    HIJRI_MONTH_CUM.append(_acc)
    _acc += _d
# -> [0,30,59,89,118,148,177,207,236,266,295,325]

HIJRI_MONTH_NAMES = [
    "Muharram", "Safar", "Rabi ul Awwal", "Rabi ul Akhir",
    "Jamad il Awwal", "Jamad il Akhir", "Rajab", "Shaban",
    "Ramadan", "Shawwal", "Zi al qa'ad", "Zi al Hajj",
]

# ---- TABLE 2 : GREGORIAN --------------------------------------------------
# Centuries  (O5:O24 base years, P5:P24 ordinal day at century start).
GREG_CENT_YEAR = [600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500,
                  1582, 1600, 1700, 1800, 1900, 2000, 2100, 2200, 2300, 2400]
GREG_CENT_ORD = [0, 36525, 73050, 109575, 146100, 182625, 219150, 255675,
                 292200, 328725, 358665, 365240, 401764, 438288, 474812,
                 511337, 547861, 584385, 620909, 657433]

# 4-year blocks  (S5:S28 years multiple of 4, T5:T28 ordinal).
GREG_4YR_YEAR = [4 * i for i in range(1, 25)]            # 4,8,...,96
GREG_4YR_ORD = [1461 * i for i in range(1, 25)]          # 1461,2922,...

# 48-month block table (V19:Y66) - month index, cumulative day-of-block, year-offset.
GREG_MONTH_CUM = [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335,      # yr0
                  366, 397, 425, 456, 486, 517, 547, 578, 609, 639, 670, 700,  # yr1
                  731, 762, 790, 821, 851, 882, 912, 943, 974, 1004, 1035, 1065,  # yr2
                  1096, 1127, 1155, 1186, 1216, 1247, 1277, 1308, 1339, 1369, 1400, 1430]  # yr3
GREG_MONTH_YROFF = ([0] * 12) + ([1] * 12) + ([2] * 12) + ([3] * 12)
GREG_MONTH_NAME = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"] * 4

GREG_MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"]

# "Month grid"  W4:AB16  ->  day-of-year offset to 1st of month for the four
# year-positions (0,1,2,3) inside a 4-year block.  Index by month 1..12.
#   column order: [yroff0, yroff1, yroff2, yroff3]
MONTH_GRID = {
    1:  [0,   366,  731,  1096],
    2:  [31,  397,  762,  1127],
    3:  [60,  425,  790,  1155],
    4:  [91,  456,  821,  1186],
    5:  [121, 486,  851,  1216],
    6:  [152, 517,  882,  1247],
    7:  [182, 547,  912,  1277],
    8:  [213, 578,  943,  1308],
    9:  [244, 609,  974,  1339],
    10: [274, 639,  1004, 1369],
    11: [305, 670,  1035, 1400],
    12: [335, 700,  1065, 1430],
}


# ===========================================================================
#  Excel function helpers
# ===========================================================================
def _match_le(value, table):
    """Replicate Excel MATCH(value, table, 1) for ascending data.

    Returns the 1-based position of the LARGEST element that is <= value.
    If value is smaller than table[0], Excel returns #N/A; we return 0 so the
    callers (which wrap the lookup in IFERROR) behave identically.
    """
    pos = 0
    for i, v in enumerate(table):
        if v <= value:
            pos = i + 1          # 1-based, like Excel
        else:
            break
    return pos


# ===========================================================================
#  GREGORIAN  ->  HIJRI        (Calculations!AH6 : AH33)
# ===========================================================================
def gregorian_to_ordinal(gy, gm, gd):
    """Serial ordinal day number for a Gregorian date (Calculations!AH11).

    AH11 = a + b + c + d  where
        a (AH6)  = century ordinal           VLOOKUP(year-year%100, O:P)
        b (AH8)  = 4-year-block ordinal       from largest S <= (year%100)
        c (AH9)  = day-of-year offset to 1st  MONTH_GRID[month][rem_years]
        d (AH10) = day of month
    """
    last2 = gy % 100                      # AG7 = RIGHT(year,2)
    cent_base = gy - last2                # AG6
    # a : century ordinal (exact VLOOKUP in O:P)
    a = GREG_CENT_ORD[GREG_CENT_YEAR.index(cent_base)]
    # b : whole 4-year blocks inside the century
    pos = _match_le(last2, GREG_4YR_YEAR)         # MATCH(AG7, S, 1)
    if pos == 0:
        block_years = 0                            # AG8 (IFERROR -> 0)
        b = 0                                      # AH8
    else:
        block_years = GREG_4YR_YEAR[pos - 1]       # AG8
        b = GREG_4YR_ORD[pos - 1]                  # AH8
    rem_years = last2 - block_years                # AG9  (0..3)
    # c : day-of-year offset to the 1st of this month for that year-position
    c = MONTH_GRID[gm][rem_years]                  # AH9 = VLOOKUP(month, W:AB, rem+3)
    d = gd                                          # AH10
    return a + b + c + d                            # AH11


def gregorian_to_hijri(gy, gm, gd):
    """Convert a Gregorian (year, month 1-12, day) to Fatimid (year, month, day).

    Decodes the ordinal with TABLE 1 exactly like Calculations!AH13:AH33.
    Returns ``(hijri_year, hijri_month_1_12, hijri_day)``.
    """
    ordinal = gregorian_to_ordinal(gy, gm, gd)        # AH11

    # ---- find the 30-year cycle -----------------------------------------
    p = _match_le(ordinal, HIJRI_CYCLE_ORD)           # AH13
    diff = ordinal - HIJRI_CYCLE_ORD[p - 1]           # AH15
    cyc_pos = (p - 1) if diff == 0 else p             # AH16  (boundary fix)
    cyc_ord = HIJRI_CYCLE_ORD[cyc_pos - 1]            # AH17
    days_in_cycle = ordinal - cyc_ord                 # AH18
    base_years = HIJRI_CYCLE_YEARS[cyc_pos - 1]       # AH19 = (cycle-1)*30

    # ---- find the year inside the cycle ---------------------------------
    p = _match_le(days_in_cycle, HIJRI_YEAR_CUM)      # AH21
    diff = days_in_cycle - HIJRI_YEAR_CUM[p - 1]      # AH23
    yr_pos = (p - 1) if diff == 0 else p              # AH24
    yr_cum = HIJRI_YEAR_CUM[yr_pos - 1]               # AH25
    days_in_year = days_in_cycle - yr_cum             # AH26

    # ---- find the month inside the year ---------------------------------
    p = _match_le(days_in_year, HIJRI_MONTH_CUM)      # AH28
    diff = days_in_year - HIJRI_MONTH_CUM[p - 1]      # AH30
    mon_pos = (p - 1) if diff == 0 else p             # AH31
    mon_cum = HIJRI_MONTH_CUM[mon_pos - 1]            # AH32
    day = days_in_year - mon_cum                      # AH33

    hijri_year = base_years + yr_pos                  # D5 = AH19 + AH24
    return hijri_year, mon_pos, day


# ===========================================================================
#  HIJRI  ->  GREGORIAN        (Calculations!AG37 : AH70)
# ===========================================================================
def hijri_to_ordinal(hy, hm, hd):
    """Serial ordinal day number for a Hijri date (Calculations!AH46).

    AH46 = C[cycle] + G[year-in-cycle] + L[month] + day
    """
    cyc_p = _match_le(hy, HIJRI_CYCLE_YEARS)          # AG39 = MATCH(year, B, 1)
    rem0 = hy - HIJRI_CYCLE_YEARS[cyc_p - 1]          # AG40
    # AG41 : if the year sits exactly on a cycle boundary, step back one cycle
    cyc_pos = (cyc_p - 1) if rem0 == 0 else cyc_p     # AG41
    cyc_ord = HIJRI_CYCLE_ORD[cyc_pos - 1]            # AH41 = C[cycle]
    # AG43 = year-in-cycle (1..30).  AH43 = VLOOKUP(AG43, E:G, 3) = days BEFORE
    # that year -> HIJRI_YEAR_CUM is 0-based for year 1, so use (n-1).
    year_in_cycle = hy - HIJRI_CYCLE_YEARS[cyc_pos - 1]   # AG43
    year_cum = HIJRI_YEAR_CUM[year_in_cycle - 1]      # AH43
    month_cum = HIJRI_MONTH_CUM[hm - 1]               # AH44 = VLOOKUP(month, J:L, 3)
    return cyc_ord + year_cum + month_cum + hd        # AH46


def hijri_to_gregorian(hy, hm, hd):
    """Convert a Fatimid (year, month 1-12, day) to Gregorian (year, month, day).

    Decodes the ordinal with TABLE 2 exactly like Calculations!AH48:AH70.
    Returns ``(greg_year, greg_month_1_12, greg_day)``.
    """
    ordinal = hijri_to_ordinal(hy, hm, hd)            # AH46

    # ---- century ---------------------------------------------------------
    p = _match_le(ordinal, GREG_CENT_ORD)             # AH48
    diff = ordinal - GREG_CENT_ORD[p - 1]             # AH50
    cent_pos = (p - 1) if diff == 0 else p            # AH51
    cent_ord = GREG_CENT_ORD[cent_pos - 1]            # AH52
    days_in_cent = ordinal - cent_ord                 # AH53
    cent_base = GREG_CENT_YEAR[cent_pos - 1]          # AH54

    # ---- 4-year blocks ---------------------------------------------------
    p = _match_le(days_in_cent, GREG_4YR_ORD)         # AH56
    if p == 0:
        block_years = 0
        days_after_blocks = days_in_cent              # AH61 (when no full block)
    else:
        diff = days_in_cent - GREG_4YR_ORD[p - 1]     # AH58
        blk_pos = (p - 1) if diff == 0 else p         # AH59
        block_years = GREG_4YR_YEAR[blk_pos - 1]      # AH62
        days_after_blocks = days_in_cent - GREG_4YR_ORD[blk_pos - 1]  # AH61

    # ---- month inside the (<=4 year) remainder block ---------------------
    p = _match_le(days_after_blocks, GREG_MONTH_CUM)  # AH64
    diff = days_after_blocks - GREG_MONTH_CUM[p - 1]  # AH66
    mon_pos = (p - 1) if diff == 0 else p             # AH67
    mon_cum = GREG_MONTH_CUM[mon_pos - 1]             # AH69
    day = days_after_blocks - mon_cum                 # AH70

    greg_month = ((mon_pos - 1) % 12) + 1             # month name index
    year_off = GREG_MONTH_YROFF[mon_pos - 1]          # AH68
    greg_year = cent_base + block_years + year_off    # D12 = AH54+AH62+AH68
    return greg_year, greg_month, day


# ===========================================================================
#  Convenience helpers
# ===========================================================================
def hijri_year_length(hy):
    """354 (Basita) or 355 (Kabisa) - length of a Fatimid year (uses TABLE 1).

    Uses the same cycle/year-in-cycle resolution as :func:`hijri_to_ordinal`
    (with the AG41 boundary correction) so the leap-year pattern in
    HIJRI_YEAR_LEN (1-based for year 1) is indexed correctly.
    """
    cyc_p = _match_le(hy, HIJRI_CYCLE_YEARS)
    rem0 = hy - HIJRI_CYCLE_YEARS[cyc_p - 1]
    cyc_pos = (cyc_p - 1) if rem0 == 0 else cyc_p
    year_in_cycle = hy - HIJRI_CYCLE_YEARS[cyc_pos - 1]    # 1..30
    return HIJRI_YEAR_LEN[year_in_cycle - 1]


def is_kabisa(hy):
    """True if the Fatimid year is a leap (Kabisa / 355-day) year."""
    return hijri_year_length(hy) == 355


def hijri_month_length(hy, hm):
    """Length of a Fatimid month.  Months alternate 30/29; in a Kabisa year
    the 12th month (Zi al Hajj) gains one day -> 30 instead of 29."""
    base = HIJRI_MONTH_LEN[hm - 1]
    if hm == 12 and is_kabisa(hy):
        return 30
    return base


def _gregorian_abs_day(gy, gm, gd):
    """Proleptic Gregorian Julian-Day-Number (used ONLY for the weekday of the
    grid, NOT for the Hijri conversion which is 100% table-driven)."""
    a = (14 - gm) // 12
    y = gy + 4800 - a
    m = gm + 12 * a - 3
    jdn = gd + (153 * m + 2) // 5 + 365 * y + y // 4 - y // 100 + y // 400 - 32045
    return jdn


def weekday_of_gregorian(gy, gm, gd):
    """Return weekday index 0=Sunday .. 6=Saturday for the calendar grid."""
    jdn = _gregorian_abs_day(gy, gm, gd)
    # JDN 0 was a Monday; (jdn+1)%7 -> 0=Sunday
    return (jdn + 1) % 7


if __name__ == "__main__":
    # Quick self-test printed when run directly.
    for hy, hm, hd in [(1443, 6, 1), (1448, 1, 1), (1445, 1, 1)]:
        gy, gm, gd = hijri_to_gregorian(hy, hm, hd)
        back = gregorian_to_hijri(gy, gm, gd)
        wd = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][weekday_of_gregorian(gy, gm, gd)]
        print(f"H {hy}-{hm}-{hd} -> G {gy}-{gm:02d}-{gd:02d} ({wd})  round-trip={back}")
