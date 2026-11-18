import io
import math
import random

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ═════════════════════════════════════════════
#  PAGE CONFIG
# ═════════════════════════════════════════════

# Dark theme applied from inside the app, so no .streamlit/config.toml is needed.
for _key, _val in {
    "theme.base": "dark",
    "theme.primaryColor": "#3ECF8E",
    "theme.backgroundColor": "#0F0F0F",
    "theme.secondaryBackgroundColor": "#1C1C1C",
    "theme.textColor": "#EDEDED",
}.items():
    try:
        st._config.set_option(_key, _val)
    except Exception:
        pass

st.set_page_config(
    page_title="Attendance Sheet Generator",
    page_icon="📋",
    layout="wide",
)

# ═════════════════════════════════════════════
#  DESIGN TOKENS + CSS
# ═════════════════════════════════════════════

BG, SURFACE, BORDER_C = "#0F0F0F", "#1C1C1C", "#2E2E2E"
TEXT, MUTED = "#EDEDED", "#A0A0A0"
GREEN, GREEN_HOVER = "#3ECF8E", "#2EB87A"
RED, AMBER = "#F04438", "#F5A524"

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

html, body, [class*="css"], .stApp {{ font-family: 'Inter', sans-serif; }}
.stApp {{ background: {BG}; color: {TEXT}; }}
[data-testid="stHeader"] {{ background: transparent; }}
#MainMenu, footer {{ visibility: hidden; }}
.block-container {{ padding-top: 2.2rem; padding-bottom: 4rem; max-width: 1240px; }}

/* Hero */
.hero {{ margin-bottom: 1.8rem; }}
.hero-badge {{
  display: inline-flex; align-items: center; gap: 8px;
  padding: 5px 12px; border-radius: 999px;
  background: rgba(62,207,142,.10); border: 1px solid rgba(62,207,142,.28);
  color: {GREEN}; font-size: 12px; font-weight: 600; letter-spacing: .04em;
  text-transform: uppercase;
}}
.hero h1 {{
  font-size: 2.4rem; font-weight: 700; letter-spacing: -0.03em;
  margin: .9rem 0 .3rem 0; padding: 0; color: {TEXT};
}}
.hero h1 span {{ color: {GREEN}; }}
.hero p {{ color: {MUTED}; font-size: 1rem; margin: 0; max-width: 620px; }}

/* Cards (bordered containers) */
[data-testid="stVerticalBlockBorderWrapper"] {{
  background: {SURFACE}; border: 1px solid {BORDER_C} !important;
  border-radius: 14px; padding: .4rem .5rem;
}}
.sec-title {{ display: flex; align-items: center; gap: 10px; margin-bottom: .15rem; }}
.sec-title span {{ font-size: 1.05rem; font-weight: 600; color: {TEXT}; }}
.sec-sub {{ color: {MUTED}; font-size: .85rem; margin-bottom: .9rem; }}
.batch-tag {{
  font-family: 'JetBrains Mono', monospace; font-size: 12px; font-weight: 600;
  color: {GREEN}; margin: .9rem 0 .2rem 0; text-transform: uppercase; letter-spacing: .06em;
}}

/* Inputs */
label, [data-testid="stWidgetLabel"] p {{ color: {MUTED} !important; font-size: .85rem !important; font-weight: 500; }}
[data-baseweb="input"], [data-baseweb="base-input"] {{ background: {BG} !important; border-radius: 10px !important; }}
[data-baseweb="input"] {{ border: 1px solid {BORDER_C} !important; }}
[data-baseweb="input"]:focus-within {{ border-color: {GREEN} !important; box-shadow: 0 0 0 3px rgba(62,207,142,.15); }}
.stNumberInput input {{ font-family: 'JetBrains Mono', monospace; color: {TEXT}; }}
.stNumberInput button {{ background: {SURFACE} !important; color: {MUTED} !important; border-color: {BORDER_C} !important; }}
.stNumberInput button:hover {{ color: {GREEN} !important; }}

/* Buttons */
.stButton > button, .stDownloadButton > button {{
  width: 100%; border-radius: 10px; font-weight: 600; font-size: .95rem;
  padding: .7rem 1rem; transition: all .15s ease; font-family: 'Inter', sans-serif;
}}
.stButton > button[kind="primary"], [data-testid="stBaseButton-primary"],
.stDownloadButton > button {{
  background: {GREEN}; color: #06140D; border: 1px solid {GREEN};
}}
.stButton > button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover,
.stDownloadButton > button:hover {{
  background: {GREEN_HOVER}; border-color: {GREEN_HOVER}; color: #06140D;
  box-shadow: 0 6px 22px rgba(62,207,142,.25);
}}
.stDownloadButton > button p, .stButton > button[kind="primary"] p {{ color: #06140D !important; font-weight: 600; }}

/* Metric cards */
.metrics {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 1rem; }}
.metric {{
  background: {SURFACE}; border: 1px solid {BORDER_C}; border-radius: 14px; padding: 16px 18px;
}}
.metric .m-top {{ display: flex; align-items: center; justify-content: space-between; }}
.metric .m-label {{ color: {MUTED}; font-size: 12px; font-weight: 500; text-transform: uppercase; letter-spacing: .05em; }}
.metric .m-value {{ font-family: 'JetBrains Mono', monospace; font-size: 1.7rem; font-weight: 600; color: {TEXT}; margin-top: 8px; }}

/* Distribution */
.dist {{ margin: .2rem 0 1.1rem 0; }}
.dist-head {{ display: flex; justify-content: space-between; font-size: .85rem; margin-bottom: 8px; }}
.dist-head b {{ color: {TEXT}; font-weight: 600; }}
.dist-head span {{ color: {MUTED}; font-family: 'JetBrains Mono', monospace; font-size: 12px; }}
.bar {{ display: flex; height: 10px; border-radius: 999px; overflow: hidden; background: {BG}; gap: 2px; }}
.bar i {{ display: block; height: 100%; }}
.legend {{ display: flex; gap: 18px; margin-top: 8px; flex-wrap: wrap; }}
.legend div {{ display: flex; align-items: center; gap: 6px; color: {MUTED}; font-size: 12px; }}
.legend em {{ width: 8px; height: 8px; border-radius: 50%; display: inline-block; }}
.legend strong {{ color: {TEXT}; font-family: 'JetBrains Mono', monospace; font-weight: 500; }}

/* Empty state */
.empty {{
  border: 1px dashed {BORDER_C}; border-radius: 14px; padding: 4.5rem 2rem; text-align: center;
  background: repeating-linear-gradient(135deg, transparent 0 14px, rgba(255,255,255,.012) 14px 28px);
}}
.empty .e-icon {{
  width: 52px; height: 52px; border-radius: 14px; margin: 0 auto 1rem auto;
  background: rgba(62,207,142,.10); border: 1px solid rgba(62,207,142,.25);
  display: flex; align-items: center; justify-content: center;
}}
.empty h3 {{ margin: 0 0 .3rem 0; font-size: 1.1rem; font-weight: 600; color: {TEXT}; }}
.empty p {{ margin: 0; color: {MUTED}; font-size: .9rem; }}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {BORDER_C}; }}
.stTabs [data-baseweb="tab"] {{ color: {MUTED}; font-weight: 500; padding: 8px 14px; border-radius: 8px 8px 0 0; }}
.stTabs [aria-selected="true"] {{ color: {GREEN} !important; }}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {GREEN} !important; }}

/* Dataframe */
[data-testid="stDataFrame"] {{ border: 1px solid {BORDER_C}; border-radius: 12px; overflow: hidden; }}

/* Section label above preview */
.pv-title {{ font-size: .8rem; color: {MUTED}; text-transform: uppercase; letter-spacing: .06em; font-weight: 600; margin: 1.1rem 0 .5rem 0; }}

@media (max-width: 900px) {{ .metrics {{ grid-template-columns: repeat(2, 1fr); }} }}
</style>
""",
    unsafe_allow_html=True,
)

# ═════════════════════════════════════════════
#  LUCIDE ICONS (inline SVG)
# ═════════════════════════════════════════════

ICONS = {
    "clipboard": '<rect width="8" height="4" x="8" y="2" rx="1" ry="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><path d="m9 14 2 2 4-4"/>',
    "calendar": '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
    "flask": '<path d="M10 2v7.527a2 2 0 0 1-.211.896L4.72 20.55a1 1 0 0 0 .9 1.45h12.76a1 1 0 0 0 .9-1.45l-5.069-10.127A2 2 0 0 1 14 9.527V2"/><path d="M8.5 2h7"/><path d="M7 16h10"/>',
    "users": '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
    "percent": '<line x1="19" x2="5" y1="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/>',
    "layers": '<path d="m12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83Z"/><path d="m22 17.65-9.17 4.16a2 2 0 0 1-1.66 0L2 17.65"/><path d="m22 12.65-9.17 4.16a2 2 0 0 1-1.66 0L2 12.65"/>',
}


def icon(name, size=18, color=GREEN):
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" '
        f'fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        f"{ICONS[name]}</svg>"
    )


def section(icon_name, title, subtitle):
    st.markdown(
        f'<div class="sec-title">{icon(icon_name, 20)}<span>{title}</span></div>'
        f'<div class="sec-sub">{subtitle}</div>',
        unsafe_allow_html=True,
    )


# ═════════════════════════════════════════════
#  ATTENDANCE LOGIC
# ═════════════════════════════════════════════

def random_distribution():
    low_pct = round(random.uniform(20, 25), 1)
    mid_pct = round(random.uniform(25, 40), 1)
    return low_pct, mid_pct


def assign_target_percents(num_rows, low_pct, mid_pct):
    n_low = int(num_rows * low_pct / 100)
    n_mid = int(num_rows * mid_pct / 100)
    n_high = num_rows - n_low - n_mid

    def spread(lo, hi, count):
        if count == 0:
            return []
        lo, hi = int(lo), int(hi)
        if count <= (hi - lo + 1):
            return random.sample(range(lo, hi + 1), count)
        step = (hi - lo) / (count + 1)
        return [round(lo + (i + 1) * step, 1) for i in range(count)]

    pcts = spread(75, 80, n_low) + spread(80, 90, n_mid) + spread(90, 100, n_high)
    random.shuffle(pcts)
    return pcts


def generate_attendance_data(names, num_days, low_pct, mid_pct, start_index=1):
    all_pcts = assign_target_percents(len(names), low_pct, mid_pct)
    rows = []
    for i, name in enumerate(names):
        target_pct = all_pcts[i]
        present_days = min(
            max(int(math.ceil(0.75 * num_days)), round((target_pct / 100) * num_days)),
            num_days,
        )
        attendance = ["P"] * num_days
        absent_needed = num_days - present_days
        if absent_needed > 0:
            absent_placed = 0
            attempts = 0
            while absent_placed < absent_needed and attempts < 1000:
                attempts += 1
                burst = random.randint(1, min(3, absent_needed - absent_placed))
                start = random.randint(0, num_days - burst)
                if all(attendance[start + k] == "P" for k in range(burst)):
                    for k in range(burst):
                        attendance[start + k] = "A"
                    absent_placed += burst
            p_indices = [j for j in range(num_days) if attendance[j] == "P"]
            still_needed = absent_needed - absent_placed
            if still_needed > 0 and len(p_indices) >= still_needed:
                for j in random.sample(p_indices, still_needed):
                    attendance[j] = "A"
        actual_pct = round((attendance.count("P") / num_days) * 100, 1)
        rows.append([start_index + i, name] + attendance + [f"{int(actual_pct)}%"])
    return rows


def band_counts(rows):
    """Return (low, mid, high) student counts from generated rows."""
    low = mid = high = 0
    for r in rows:
        pct = int(r[-1].replace("%", ""))
        if pct < 80:
            low += 1
        elif pct < 90:
            mid += 1
        else:
            high += 1
    return low, mid, high


def avg_attendance(rows):
    if not rows:
        return 0.0
    return sum(int(r[-1].replace("%", "")) for r in rows) / len(rows)


def make_dataframe(rows, num_days):
    cols = ["S.No", "Student Name"] + [f"D{i}" for i in range(1, num_days + 1)] + ["Attendance %"]
    return pd.DataFrame(rows, columns=cols)


# ═════════════════════════════════════════════
#  EXCEL EXPORT
# ═════════════════════════════════════════════

HEADER_FILL = PatternFill("solid", start_color="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
TITLE_FILL = PatternFill("solid", start_color="2E75B6")
TITLE_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
BATCH_FILL = PatternFill("solid", start_color="D6E4F0")
BATCH_FONT = Font(bold=True, color="1F4E79", name="Arial", size=10)
P_FILL = PatternFill("solid", start_color="C6EFCE")
P_FONT = Font(color="276221", name="Arial", size=9)
A_FILL = PatternFill("solid", start_color="FFC7CE")
A_FONT = Font(color="9C0006", name="Arial", size=9)
PCT_HIGH_FILL = PatternFill("solid", start_color="C6EFCE")
PCT_MID_FILL = PatternFill("solid", start_color="FFEB9C")
PCT_LOW_FILL = PatternFill("solid", start_color="FFC7CE")
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
thin = Side(style="thin", color="BFBFBF")
XL_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_header(cell):
    cell.fill, cell.font, cell.alignment, cell.border = HEADER_FILL, HEADER_FONT, CENTER, XL_BORDER


def write_attendance_block(ws, start_row, rows, num_days, label=None):
    if label:
        c = ws.cell(start_row, 1, label)
        c.font, c.fill, c.alignment = BATCH_FONT, BATCH_FILL, LEFT
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=2 + num_days + 1)
        start_row += 1
    headers = ["S.No", "Student Name"] + [f"Day {i}" for i in range(1, num_days + 1)] + ["Attendance %"]
    for col, h in enumerate(headers, 1):
        style_header(ws.cell(start_row, col, h))
    start_row += 1
    for row_data in rows:
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(start_row, col, val)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = CENTER
            cell.border = XL_BORDER
            if val == "P":
                cell.fill, cell.font = P_FILL, P_FONT
            elif val == "A":
                cell.fill, cell.font = A_FILL, A_FONT
            elif isinstance(val, str) and val.endswith("%"):
                pct = int(val.replace("%", ""))
                cell.fill = PCT_HIGH_FILL if pct >= 92 else (PCT_MID_FILL if pct >= 82 else PCT_LOW_FILL)
                cell.font = Font(bold=True, name="Arial", size=9)
            elif col == 2:
                cell.alignment = LEFT
        start_row += 1
    return start_row + 1


def set_col_widths(ws, num_days):
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 22
    for i in range(3, 3 + num_days):
        ws.column_dimensions[get_column_letter(i)].width = 6
    ws.column_dimensions[get_column_letter(3 + num_days)].width = 14


def build_workbook_bytes(lecture, labs):
    wb = Workbook()
    ws_t = wb.active
    ws_t.title = "Lecture_Attendance"
    ws_t.freeze_panes = "C3"
    ld = lecture["days"]
    ws_t.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2 + ld + 1)
    t = ws_t.cell(1, 1, "LECTURE ATTENDANCE REGISTER")
    t.fill, t.font, t.alignment = TITLE_FILL, TITLE_FONT, CENTER
    ws_t.row_dimensions[1].height = 22
    next_row = write_attendance_block(ws_t, 2, lecture["rows"], ld)
    set_col_widths(ws_t, ld)

    # Distribution summary
    next_row += 1

    def band_rows(rows, n):
        lo, mi, hi = band_counts(rows)
        pc = lambda x: f"{round(x / n * 100, 1)}%"
        return [
            ("  low  (75–80%)", pc(lo), f"→ {lo} students"),
            ("  mid  (80–90%)", pc(mi), f"→ {mi} students"),
            ("  high (90–100%)", pc(hi), f"→ {hi} students"),
        ]

    n_lec = len(lecture["rows"])
    summary_rows = [("📊 DISTRIBUTION SUMMARY", None, None), ("Lectures", None, None)]
    summary_rows += band_rows(lecture["rows"], n_lec)
    for bi, b in enumerate(labs, 1):
        summary_rows.append((f"Lab Batch {bi} ({len(b['rows'])} students)", None, None))
        summary_rows += band_rows(b["rows"], len(b["rows"]))

    band_fills = {
        "low": PatternFill("solid", start_color="FFC7CE"),
        "mid": PatternFill("solid", start_color="FFEB9C"),
        "high": PatternFill("solid", start_color="C6EFCE"),
    }
    for label, pct, count in summary_rows:
        row = next_row
        c1 = ws_t.cell(row, 1, label)
        c1.font = Font(name="Arial", size=9, bold=("SUMMARY" in label or label == "Lectures" or "Batch" in label))
        c1.alignment = LEFT
        if pct:
            c2, c3 = ws_t.cell(row, 2, pct), ws_t.cell(row, 3, count)
            c2.font, c3.font = Font(name="Arial", size=9, bold=True), Font(name="Arial", size=9)
            c2.alignment, c3.alignment = CENTER, LEFT
            band = "low" if "75" in label else ("mid" if "80" in label else "high")
            c1.fill = c2.fill = c3.fill = band_fills[band]
        elif "SUMMARY" in label:
            c1.fill = PatternFill("solid", start_color="2E75B6")
            c1.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            ws_t.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
            c1.alignment = CENTER
        else:
            c1.fill = PatternFill("solid", start_color="D6E4F0")
            ws_t.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        for col in range(1, 4):
            ws_t.cell(row, col).border = XL_BORDER
        next_row += 1

    if labs:
        ws_l = wb.create_sheet("Lab_Attendance")
        ws_l.freeze_panes = "C2"
        max_days = max(b["days"] for b in labs)
        ws_l.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2 + max_days + 1)
        t = ws_l.cell(1, 1, "LAB ATTENDANCE REGISTER")
        t.fill, t.font, t.alignment = TITLE_FILL, TITLE_FONT, CENTER
        ws_l.row_dimensions[1].height = 22
        cur_row = 2
        for i, b in enumerate(labs, 1):
            cur_row = write_attendance_block(ws_l, cur_row, b["rows"], b["days"], label=f"  Batch {i}")
        set_col_widths(ws_l, max_days)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_result(n_students, lecture_days, lab_cfg):
    names = [f"Student_{i}" for i in range(1, n_students + 1)]
    lo, mi = random_distribution()
    lecture = {"days": lecture_days, "rows": generate_attendance_data(names, lecture_days, lo, mi, 1)}

    labs, s_idx = [], 1
    for cfg in lab_cfg:
        b_names = [f"Student_{i}" for i in range(1, cfg["students"] + 1)]
        bl, bm = random_distribution()
        rows = generate_attendance_data(b_names, cfg["days"], bl, bm, s_idx)
        labs.append({"days": cfg["days"], "rows": rows})
        s_idx += cfg["students"]

    return {"lecture": lecture, "labs": labs, "xlsx": build_workbook_bytes(lecture, labs)}


# ═════════════════════════════════════════════
#  UI HELPERS
# ═════════════════════════════════════════════

def metric_card(label, value, icon_name):
    return (
        f'<div class="metric"><div class="m-top"><span class="m-label">{label}</span>{icon(icon_name, 16)}</div>'
        f'<div class="m-value">{value}</div></div>'
    )


def dist_block(title, rows):
    lo, mi, hi = band_counts(rows)
    n = max(len(rows), 1)
    w = lambda x: f"{x / n * 100:.2f}%"
    return (
        f'<div class="dist"><div class="dist-head"><b>{title}</b><span>{len(rows)} students · avg {avg_attendance(rows):.0f}%</span></div>'
        f'<div class="bar"><i style="width:{w(lo)};background:{RED}"></i>'
        f'<i style="width:{w(mi)};background:{AMBER}"></i>'
        f'<i style="width:{w(hi)};background:{GREEN}"></i></div>'
        f'<div class="legend">'
        f'<div><em style="background:{RED}"></em>75–80% <strong>{lo}</strong></div>'
        f'<div><em style="background:{AMBER}"></em>80–90% <strong>{mi}</strong></div>'
        f'<div><em style="background:{GREEN}"></em>90–100% <strong>{hi}</strong></div>'
        f"</div></div>"
    )


def style_preview(df):
    def cell_style(v):
        if v == "P":
            return "background-color:#12351F;color:#3ECF8E;text-align:center;font-weight:600"
        if v == "A":
            return "background-color:#3B1517;color:#F87171;text-align:center;font-weight:600"
        return ""

    def pct_style(v):
        p = int(str(v).replace("%", ""))
        color = GREEN if p >= 92 else (AMBER if p >= 82 else RED)
        return f"color:{color};font-weight:600;text-align:center"

    styler = df.style
    mapper = styler.map if hasattr(styler, "map") else styler.applymap
    day_cols = [c for c in df.columns if c.startswith("D")]
    styler = mapper(cell_style, subset=day_cols)
    mapper = styler.map if hasattr(styler, "map") else styler.applymap
    return mapper(pct_style, subset=["Attendance %"])


# ═════════════════════════════════════════════
#  HERO
# ═════════════════════════════════════════════

st.markdown(
    f'<div class="hero"><div class="hero-badge">{icon("clipboard", 14)} Attendance toolkit</div>'
    f'<h1>Attendance Sheet <span>Generator</span></h1>'
    f"<p>Create realistic lecture and lab attendance registers in seconds. "
    f"Set the numbers, generate, preview, and download a formatted Excel file.</p></div>",
    unsafe_allow_html=True,
)

# ═════════════════════════════════════════════
#  LAYOUT
# ═════════════════════════════════════════════

left, right = st.columns([1, 1.65], gap="large")

with left:
    with st.container(border=True):
        section("calendar", "Lectures", "Theory register for the whole class")
        n_students = st.number_input("Total number of students", min_value=1, max_value=500, value=60, step=1)
        lecture_days = st.number_input("Number of lecture days", min_value=1, max_value=120, value=30, step=1)

    with st.container(border=True):
        section("flask", "Lab", "Separate register for every lab batch")
        include_lab = st.toggle("Include lab attendance", value=False)
        lab_cfg = []
        if include_lab:
            n_batches = st.number_input("Number of lab batches", min_value=1, max_value=12, value=2, step=1)
            for i in range(int(n_batches)):
                st.markdown(f'<div class="batch-tag">Batch {i + 1}</div>', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                b_students = c1.number_input("Students", min_value=1, max_value=500, value=20, step=1, key=f"bs_{i}")
                b_days = c2.number_input("Lab days", min_value=1, max_value=120, value=10, step=1, key=f"bd_{i}")
                lab_cfg.append({"students": int(b_students), "days": int(b_days)})

    generate = st.button("Generate attendance sheet", type="primary")

if generate:
    with st.spinner("Generating attendance…"):
        st.session_state["result"] = build_result(int(n_students), int(lecture_days), lab_cfg)

with right:
    result = st.session_state.get("result")

    if not result:
        st.markdown(
            f'<div class="empty"><div class="e-icon">{icon("layers", 24)}</div>'
            f"<h3>Nothing generated yet</h3>"
            f"<p>Fill in the details on the left and hit <b>Generate</b> to see a live preview here.</p></div>",
            unsafe_allow_html=True,
        )
    else:
        lecture, labs = result["lecture"], result["labs"]
        lec_rows = lecture["rows"]

        st.markdown(
            '<div class="metrics">'
            + metric_card("Students", len(lec_rows), "users")
            + metric_card("Lecture days", lecture["days"], "calendar")
            + metric_card("Avg attendance", f"{avg_attendance(lec_rows):.0f}%", "percent")
            + metric_card("Lab batches", len(labs), "flask")
            + "</div>",
            unsafe_allow_html=True,
        )

        st.download_button(
            "Download Excel file",
            data=result["xlsx"],
            file_name="attendance_sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown('<div class="pv-title">Distribution</div>', unsafe_allow_html=True)
        with st.container(border=True):
            html = dist_block("Lectures", lec_rows)
            for i, b in enumerate(labs, 1):
                html += dist_block(f"Lab · Batch {i}", b["rows"])
            st.markdown(html, unsafe_allow_html=True)

        st.markdown('<div class="pv-title">Preview</div>', unsafe_allow_html=True)
        tabs = st.tabs(["Lectures"] + [f"Lab · Batch {i}" for i in range(1, len(labs) + 1)])
        with tabs[0]:
            st.dataframe(
                style_preview(make_dataframe(lec_rows, lecture["days"])),
                hide_index=True,
                height=420,
            )
        for tab, b in zip(tabs[1:], labs):
            with tab:
                st.dataframe(
                    style_preview(make_dataframe(b["rows"], b["days"])),
                    hide_index=True,
                    height=420,
                )
