"""
app_streamlit.py  -  Fatimid (Misri) Hijri Calendar  -  Streamlit front-end
====
A Streamlit web application that reproduces the look & feel of the
``CALENDAR (Updated).xlsx`` calendar sheet:

  * dual Hijri / Gregorian dates in every day cell,
  * Greg month name shown below the Greg day number (e.g. "JUN"),
  * the five daily prayer times FJ / SR / NN / ZW / MG displayed to the
    RIGHT of the Gregorian date block (matching the Excel layout),
  * colour-coded days (event days get the dark-green + yellow treatment,
    Fridays/weekends are tinted, today is highlighted),
  * a per-day "No. of Events" indicator,
  * the event listing table at the bottom,
  * month / year navigation and a city selector.

All calendar maths comes from calendar_core and calendar_builder.
No external calendar libraries.
"""

import datetime
import streamlit as st

import calendar_core as cc
import calendar_builder as cb

# ── Colour palette ──────────────────────────────────────────────────────────
DARK_GREEN  = "#3c5d2e"
MED_GREEN   = "#7d9b54"
LIGHT_GREEN = "#c5d8a4"
PALE        = "#eef3e4"
BLUE_HDR    = "#4f81bd"
LIGHT_BLUE  = "#dce6f2"
RED_DATE    = "#c00000"
YELLOW      = "#f2ef00"
WHITE       = "#ffffff"
TODAY_BG    = "#ffd966"

PRAYER_ORDER = ["FJ", "SR", "NN", "ZW", "MG"]


# ── Helpers ──────────────────────────────────────────────────────────────────
def convert_gregorian_to_hijri(gy, gm, gd):
    """Convert a Gregorian date to Hijri using the project's existing core logic."""
    return cc.gregorian_to_hijri(gy, gm, gd)


def hijri_today():
    t = datetime.date.today()
    return convert_gregorian_to_hijri(t.year, t.month, t.day)


def _init_state():
    if "hy" not in st.session_state:
        hy, hm, _ = hijri_today()
        st.session_state.hy   = hy
        st.session_state.hm   = hm
        st.session_state.city = "Qatar"


# ── Navigation callbacks ─────────────────────────────────────────────────────
def _prev_month():
    st.session_state.hm -= 1
    if st.session_state.hm < 1:
        st.session_state.hm = 12
        st.session_state.hy -= 1

def _next_month():
    st.session_state.hm += 1
    if st.session_state.hm > 12:
        st.session_state.hm = 1
        st.session_state.hy += 1

def _prev_year():
    st.session_state.hy -= 1

def _next_year():
    st.session_state.hy += 1

def _go_today():
    hy, hm, _ = hijri_today()
    st.session_state.hy = hy
    st.session_state.hm = hm


# ── Day-cell HTML renderer ───────────────────────────────────────────────────
def _cell_html(cell, today_hijri):
    """Return an HTML string for one day cell, matching the Excel layout."""
    if cell is None:
        return f'<div style="background:{WHITE};min-height:90px;border:1px solid #ddd;"></div>'

    has_events = cell.event_count > 0
    is_today   = (cell.hijri_year, cell.hijri_month, cell.hijri_day) == today_hijri
    is_weekend = cell.weekday in (5, 6)

    num_bg   = DARK_GREEN  if has_events else LIGHT_GREEN
    num_fg   = YELLOW      if has_events else "#1b1b1b"
    body_bg  = TODAY_BG    if is_today   else (LIGHT_BLUE if is_weekend else PALE)
    greg_fg  = RED_DATE    if cell.weekday in (0, 6) else "#16407a"

    # ── top row: [Hijri day] [Greg day / Greg month] [Prayer times] ──────────
    # Hijri day box
    hijri_box = (
        f'<div style="background:{num_bg};color:{num_fg};'
        f'font-family:Calibri;font-size:2.2rem;font-weight:bold;text-align:center;'
        f'padding:2px 7px;min-width:44px;line-height:1.1;">'
        f'{cell.hijri_day}</div>'
    )

    # Greg date chip: day number + month abbreviation stacked
    greg_month_abbr = cell.greg_month_name[:3].upper()
    greg_box = (
        f'<div style="background:{LIGHT_BLUE};color:{greg_fg};'
        f'font-family:Calibri;'
        f'text-align:center;padding:1px 5px;min-width:38px;">'
        f'<div style="font-size:26px;font-weight:bold;line-height:1.2;">{cell.greg_day}</div>'
        f'<div style="font-size:13px;font-weight:bold;line-height:1.1;">{greg_month_abbr}</div>'
        f'</div>'
    )

    # Prayer times column (to the right of greg box)
    prayer_html = ""
    for lab in PRAYER_ORDER:
        t = cell.times.get(lab, "")
        prayer_html += (
            f'<div style="display:flex;gap:2px;line-height:1.2;">'
            f'<span style="color:#274b8c;font-weight:bold;font-size:11px;min-width:18px;">{lab}</span>'
            f'<span style="color:#222;font-size:12px;">{t}</span>'
            f'</div>'
        )
    prayer_box = (
        f'<div style="background:{body_bg};padding:1px 4px 1px 6px;flex:1;font-family:Calibri;text-align:center;">'
        f'{prayer_html}'
        f'</div>'
    )

    top_row = (
        f'<div style="display:flex;align-items:center;justify-content:center;gap:0;font-family:Calibri;">'
        f'{hijri_box}{greg_box}{prayer_box}'
        f'</div>'
    )

    # ── event count row ───────────────────────────────────────────────────────
    event_row = ""
    if has_events:
        event_row = (
            f'<div style="background:{body_bg};color:#7a0000;'
            f'font-size:10px;font-style:italic;padding:1px 3px;">'
            f'No. of Events &nbsp;{cell.event_count}</div>'
        )

    return (
        f'<div style="background:{body_bg};border:1px solid #ccc;'
        f'min-height:80px;overflow:hidden;text-align:center;">'
        f'{top_row}{event_row}'
        f'</div>'
    )


# ── Main render ──────────────────────────────────────────────────────────────
def render():
    _init_state()
    hy   = st.session_state.hy
    hm   = st.session_state.hm
    city = st.session_state.city

    model       = cb.build_month(hy, hm, city)
    today_hijri = hijri_today()

    # ── Global CSS ────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
      .block-container { padding-top: 3rem !important; }
      div[data-testid="column"] > div { padding: 0 2px !important; }
      .stButton > button {backgrouNd-color: """+PALE+"""; color: """+DARK_GREEN+"""; border: none; border-radius: 30px;
        padding: 2px 8px; font-size: 12px; height: 30px;
      }
    </style>
    """, unsafe_allow_html=True)

    # ── Page title ───────────────────────────────────────────────────────────
    st.markdown(
        '<div style="text-align:center;margin-bottom:8px;font-family:Calibri;">'
        f'<div style="background:{DARK_GREEN};color:{LIGHT_GREEN};font-size:30px;font-weight:bold;">'
        'سلیمانی (فاطمي) حجری کلندر</div>'
        '</div>',
        unsafe_allow_html=True
    )

    # ── Header band ───────────────────────────────────────────────────────────
    hijri_month_name = cc.HIJRI_MONTH_NAMES[hm - 1]
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:0;margin-bottom:4px;font-family:Calibri;">'
        f'<div style="background:{DARK_GREEN};color:#fff;font-size:20px;font-weight:bold;'
        f'padding:6px 18px;flex:2;text-align:center;">{hijri_month_name}</div>'
        f'<div style="background:{MED_GREEN};color:#fff;font-size:20px;font-weight:bold;'
        f'padding:6px 12px;text-align:center;">{hy}</div>'
        f'<div style="background:{LIGHT_GREEN};color:{DARK_GREEN};font-size:20px;font-weight:bold;'
        f'padding:6px 12px;text-align:center;">{model["year_type"]}</div>'
        f'<div style="background:{BLUE_HDR};color:#fff;font-size:20px;font-weight:bold;'
        f'padding:6px 16px;flex:1;text-align:center;">{city}</div>'
        f'<div style="background:{WHITE};color:#16407a;font-size:20px;font-weight:bold;'
        f'padding:6px 16px;flex:2;text-align:center;border:1px solid #ccc;">'
        f'{model["greg_span"]}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── Navigation row ────────────────────────────────────────────────────────
    nav_cols = st.columns([0.9, 0.9, 0.9, 0.9, 0.9, 1, 1, 1, 1.0, 0.9, 1, 1.3, 0.9], gap="xxsmall", vertical_alignment="center", width="stretch")
    nav_cols[0].button("<< Year",  on_click=_prev_year,  key="prevy", use_container_width=True)
    nav_cols[1].button("< Month",  on_click=_prev_month, key="prevm", use_container_width=True)
    nav_cols[2].button("Today",    on_click=_go_today,   key="today", use_container_width=True)
    nav_cols[3].button("Month >",  on_click=_next_month, key="nextm", use_container_width=True)
    nav_cols[4].button("Year >>",  on_click=_next_year,  key="nexty", use_container_width=True)

    city_sel = nav_cols[5].selectbox(
        "City", cb.CITY_NAMES,
        index=cb.CITY_NAMES.index(city) if city in cb.CITY_NAMES else 0,
        key="city_sel", label_visibility="collapsed",
    )
    if city_sel != city:
        st.session_state.city = city_sel
        st.rerun()

    go_year = nav_cols[6].number_input(
        "Hijri Year", value=hy, min_value=1, max_value=9999,
        step=1, key="go_year", label_visibility="collapsed",
    )
    go_month = nav_cols[7].selectbox(
        "Hijri Month", cc.HIJRI_MONTH_NAMES,
        index=hm - 1, key="go_month", label_visibility="collapsed", help="Select the Hijri month to jump to"
    )
    if nav_cols[8].button("Go", key="go_btn",width=80, help="Jump to the specified Hijri month and year"):
        st.session_state.hy = int(go_year)
        st.session_state.hm = cc.HIJRI_MONTH_NAMES.index(go_month) + 1
        st.rerun()
    
    nav_cols[9].caption(" Gregorian → Hijri",width="stretch")
    conv_date = nav_cols[10].date_input(
        "Gregorian Date",
        value=datetime.date.today(),
        key="conv_date",
        label_visibility="collapsed",min_value=datetime.date(622, 7, 15),help="Select a Gregorian date to convert to Hijri"       
    )

    conv_hy, conv_hm, conv_hd = convert_gregorian_to_hijri(conv_date.year, conv_date.month, conv_date.day)
    
    next_date = conv_date + datetime.timedelta(days=1)
    next_hy, next_hm, next_hd = convert_gregorian_to_hijri(next_date.year, next_date.month, next_date.day)
    
    
    conv_hm_name = cc.HIJRI_MONTH_NAMES[conv_hm - 1]
    next_hm_name = cc.HIJRI_MONTH_NAMES[next_hm - 1]

    nav_cols[11].markdown(
            f'<div style="font-family:Arial; font-size:13px;color:#274b8c;">'
            f'Hijri date: <strong>{conv_hd:02d} {conv_hm_name} {conv_hy:04d}</strong><br />'
            f'After Maghrib: <strong>{next_hd:02d} {next_hm_name} {next_hy:04d}</strong></div>',
            unsafe_allow_html=True,
        )
    # _ = nav_cols[12].empty()

    st.markdown("<hr style='margin:4px 0;'>", unsafe_allow_html=True)

    # ── Weekday header ────────────────────────────────────────────────────────
    hdr_cols = st.columns(7,gap="small")
    for i, day_name in enumerate(cb.WEEKDAYS):
        hdr_cols[i].markdown(
            f'<div style="background:{MED_GREEN};color:#fff;text-align:center;'
            f'font-family:Calibri;font-weight:bold;font-size:14px;padding:4px 0;">{day_name}</div>',
            unsafe_allow_html=True,
        )

    # ── Calendar grid ─────────────────────────────────────────────────────────
    for week in model["weeks"]:
        week_cols = st.columns(7,gap="small",vertical_alignment="center")
        for col_idx, cell in enumerate(week):
            week_cols[col_idx].markdown(
                _cell_html(cell, today_hijri),
                unsafe_allow_html=True,
            )

    # ── Events table ──────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:{DARK_GREEN};color:#fff;text-align:center;'
        f'font-weight:bold;font-size:14px;padding:5px;margin-top:8px;">Events</div>',
        unsafe_allow_html=True,
    )

    events = model["events"]
    if events:
        # Split into two halves side by side (like the Excel layout)
        mid = (len(events) + 1) // 2
        left_events  = events[:mid]
        right_events = events[mid:]

        def _events_table_html(evts):
            rows_html = ""
            for i, e in enumerate(evts):
                bg = "#e4ecd5" if i % 2 == 0 else "#f3f6ec"
                rows_html += (
                    f'<tr style="background:{bg};">'
                    f'<td style="padding:2px 6px;font-size:11px;white-space:nowrap;">{e["code"]}</td>'
                    f'<td style="padding:2px 6px;font-size:11px;text-align:center;">{e["day"]}</td>'
                    f'<td style="padding:2px 6px;font-size:11px;">{e["occasion"]}</td>'
                    f'</tr>'
                )
            return (
                f'<table style="width:100%;border-collapse:collapse;">'
                f'<thead><tr style="background:{DARK_GREEN};color:#fff;">'
                f'<th style="padding:3px 6px;font-size:9px;text-align:left;">Evnt no</th>'
                f'<th style="padding:3px 6px;font-size:9px;text-align:center;">Date</th>'
                f'<th style="padding:3px 6px;font-size:9px;text-align:left;">Events</th>'
                f'</tr></thead><tbody>{rows_html}</tbody></table>'
            )

        left_col, right_col = st.columns(2)
        left_col.markdown(_events_table_html(left_events),   unsafe_allow_html=True)
        right_col.markdown(_events_table_html(right_events), unsafe_allow_html=True)
    else:
        st.markdown(
            '<p style="font-style:italic;color:#888;font-size:11px;">'
            'No events recorded for this month.</p>',
            unsafe_allow_html=True,
        )

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown(
        '<p style="font-size:12px;color:#aaa;margin-top:10px;">'
        'Fatimid (Misri) Hijri Calendar — logic migrated verbatim from '
        'CALENDAR (Updated).xlsx (Calendar and Conversion Logic Provided by Muneer Tailor,Qatar)- Application Developed by Rahil Kasimi</p>',
        unsafe_allow_html=True,
    )


# ── Entry point ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Fatimi Hijri Calendar",
    layout="wide",
    initial_sidebar_state="collapsed",
)

render()
