"""
QA Work-Hours Analytics Dashboard
----------------------------------
A single-page Streamlit dashboard for analyzing QA team work-hour trackers.

Upload the QA tracker workbook -> get instant KPI summary + charts
(Donut, Waffle, Range, Bar/Column) -> filter Daily / Monthly / Yearly ->
export to Excel or PDF.

Run:  streamlit run app.py
"""

import io
import re
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore")

# ============================================================================
# PAGE CONFIG & GLOBAL STYLE
# ============================================================================
st.set_page_config(
    page_title="QA Work Hours Dashboard",
    page_icon="\U0001F4CA",
    layout="wide",
    initial_sidebar_state="expanded",
)

PRIMARY = "#5B5FEF"
PRIMARY_DARK = "#3E42C9"
BILLABLE_COLOR = "#22C55E"
NONBILL_COLOR = "#F59E0B"
NOTWORKED_COLOR = "#EF4444"
BG_CARD = "#FFFFFF"
TEXT_MUTED = "#6B7280"

CUSTOM_CSS = f"""
<style>
    .stApp {{
        background: linear-gradient(180deg, #F7F8FC 0%, #EEF0FA 100%);
    }}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    /* Keep the header bar itself functional (it contains the sidebar
       expand/collapse arrow) but make it visually blend into the page. */
    header[data-testid="stHeader"] {{
        background: transparent;
        height: 3rem;
    }}

    .app-title {{
        font-size: 2.1rem;
        font-weight: 800;
        color: #1F2340;
        margin-bottom: 0px;
    }}
    .app-subtitle {{
        color: {TEXT_MUTED};
        font-size: 0.95rem;
        margin-top: -6px;
        margin-bottom: 1.2rem;
    }}

    .kpi-card {{
        background: {BG_CARD};
        border-radius: 16px;
        padding: 18px 20px 14px 20px;
        box-shadow: 0 4px 18px rgba(31,35,64,0.06);
        border: 1px solid #ECEDF7;
        text-align: left;
        height: 118px;
    }}
    .kpi-label {{
        font-size: 0.78rem;
        font-weight: 600;
        color: {TEXT_MUTED};
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 6px;
    }}
    .kpi-value {{
        font-size: 1.9rem;
        font-weight: 800;
        color: #1F2340;
        line-height: 1.1;
    }}
    .kpi-sub {{
        font-size: 0.78rem;
        color: {TEXT_MUTED};
        margin-top: 4px;
    }}
    .section-header {{
        font-size: 1.25rem;
        font-weight: 700;
        color: #1F2340;
        margin-top: 1.6rem;
        margin-bottom: 0.4rem;
        border-left: 5px solid {PRIMARY};
        padding-left: 10px;
    }}
    .qa-card {{
        background: {BG_CARD};
        border-radius: 16px;
        padding: 16px 18px;
        box-shadow: 0 4px 18px rgba(31,35,64,0.06);
        border: 1px solid #ECEDF7;
        margin-bottom: 14px;
    }}
    .qa-name {{
        font-size: 1.05rem;
        font-weight: 800;
        color: #1F2340;
    }}
    .badge {{
        display:inline-block;
        padding: 3px 10px;
        border-radius: 999px;
        font-size: 0.72rem;
        font-weight: 700;
        margin-left: 8px;
    }}
    .badge-good {{ background:#DCFCE7; color:#16A34A; }}
    .badge-warn {{ background:#FEF3C7; color:#D97706; }}
    .badge-bad  {{ background:#FEE2E2; color:#DC2626; }}

    div[data-testid="stFileUploader"] {{
        background: {BG_CARD};
        border-radius: 16px;
        padding: 14px;
        border: 1.5px dashed #C7CAEF;
    }}

    /* Give the page some breathing room so charts don't touch the bottom edge */
    .block-container {{
        padding-bottom: 4rem !important;
        padding-top: 1rem !important;
    }}

    /* Hide Plotly's floating toolbar (camera/zoom/pan) to prevent accidental clicks */
    .modebar {{
        display: none !important;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Disable the floating Plotly toolbar (camera/zoom/pan) everywhere so users
# can't accidentally trigger it while scrolling/clicking on charts.
PLOTLY_CONFIG = {"displayModeBar": False, "displaylogo": False, "scrollZoom": False}

# ============================================================================
# DATA LOADING & CLEANING
# ============================================================================

REQUIRED_CANON = ["QA Name", "Date", "Day", "Month",
                   "Billable Hours", "Non-Billable Hours",
                   "Hours Not Worked", "Total Hours"]

# Map of messy possible header text -> canonical column name
HEADER_ALIASES = {
    "enter qa name here": "QA Name",
    "qa name": "QA Name",
    "date": "Date",
    "day": "Day",
    "month": "Month",
    "billable hours": "Billable Hours",
    "non-billable hours": "Non-Billable Hours",
    "non billable hours": "Non-Billable Hours",
    "hours not worked": "Hours Not Worked",
    "total hours": "Total Hours",
    "project name": "Project Name",
    "comment": "Comment",
    "comments": "Comment",
}

SKIP_SHEETS = {"master data", "config", "instructions"}
WEEKEND_TOKENS = {"sat", "sun", "saturday", "sunday"}


def _norm(s):
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _find_header_row(raw: pd.DataFrame, max_scan: int = 5):
    """Scan first few rows to find the row that looks like column headers."""
    for i in range(min(max_scan, len(raw))):
        row_vals = [_norm(v) for v in raw.iloc[i].tolist()]
        hits = sum(1 for v in row_vals if v in HEADER_ALIASES)
        if hits >= 3:
            return i
    return 0


def _to_number(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if s == "" or _norm(s) in WEEKEND_TOKENS:
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


def parse_workbook(file) -> pd.DataFrame:
    """Read every per-QA sheet in the workbook and return one tidy dataframe."""
    xls = pd.ExcelFile(file)
    frames = []

    for sheet in xls.sheet_names:
        if _norm(sheet) in SKIP_SHEETS:
            continue

        raw = xls.parse(sheet, header=None)
        if raw.empty or raw.shape[0] < 2:
            continue

        header_row = _find_header_row(raw)
        header_vals = raw.iloc[header_row].tolist()
        col_map = {}
        for idx, v in enumerate(header_vals):
            key = _norm(v)
            if key in HEADER_ALIASES:
                col_map[idx] = HEADER_ALIASES[key]

        if "Billable Hours" not in col_map.values():
            continue  # not a data sheet we recognize

        body = raw.iloc[header_row + 1:].copy()
        body = body.rename(columns=col_map)
        keep_cols = [c for c in body.columns if isinstance(c, str) and c in HEADER_ALIASES.values()]
        body = body[keep_cols]

        # Drop columns that got duplicated (keep first occurrence)
        body = body.loc[:, ~body.columns.duplicated()]

        # Fallback QA name = sheet name if column missing/blank
        if "QA Name" not in body.columns:
            body["QA Name"] = sheet
        body["QA Name"] = body["QA Name"].fillna(sheet)
        body.loc[body["QA Name"].astype(str).str.strip() == "", "QA Name"] = sheet

        # Parse date
        if "Date" in body.columns:
            body["Date"] = pd.to_datetime(body["Date"], errors="coerce")
        else:
            body["Date"] = pd.NaT

        # Drop rows with no date (blank separator rows / "Total" rows / notes)
        body = body[body["Date"].notna()]
        if body.empty:
            continue

        # Numeric columns
        for c in ["Billable Hours", "Non-Billable Hours", "Hours Not Worked", "Total Hours"]:
            if c in body.columns:
                body[c] = body[c].apply(_to_number)
            else:
                body[c] = np.nan

        # Drop weekend / off rows where all hour fields are NaN (Sat/Sun placeholders)
        hour_cols = ["Billable Hours", "Non-Billable Hours", "Hours Not Worked", "Total Hours"]
        body = body[~body[hour_cols].isna().all(axis=1)]
        if body.empty:
            continue

        body[hour_cols] = body[hour_cols].fillna(0.0)

        # Recompute Total Hours if inconsistent or missing/zero but components exist
        computed_total = body["Billable Hours"] + body["Non-Billable Hours"] + body["Hours Not Worked"]
        body["Total Hours"] = np.where(
            (body["Total Hours"] <= 0) | (body["Total Hours"].isna()),
            computed_total,
            body["Total Hours"],
        )

        # Day / Month derived from Date if not present or blank
        body["Day"] = body["Date"].dt.day_name().str.slice(0, 3)
        body["Month"] = body["Date"].dt.strftime("%b")
        body["Year"] = body["Date"].dt.year

        body["QA Name"] = body["QA Name"].astype(str).str.strip()
        # Normalize inconsistent name variants (e.g. "Saujanya.Gouda" vs "Saujanya Gouda")
        body["QA Name"] = body["QA Name"].str.replace(r"[._]+", " ", regex=True)
        body["QA Name"] = body["QA Name"].str.replace(r"\s+", " ", regex=True).str.strip()
        body["QA Name"] = body["QA Name"].str.title()

        final = body[["QA Name", "Date", "Day", "Month", "Year",
                      "Billable Hours", "Non-Billable Hours",
                      "Hours Not Worked", "Total Hours"]].copy()
        final = final[final["QA Name"] != ""]
        frames.append(final)

    if not frames:
        return pd.DataFrame(columns=["QA Name", "Date", "Day", "Month", "Year",
                                      "Billable Hours", "Non-Billable Hours",
                                      "Hours Not Worked", "Total Hours"])

    out = pd.concat(frames, ignore_index=True)
    out = out.drop_duplicates(subset=["QA Name", "Date"], keep="first")
    out = out.sort_values(["QA Name", "Date"]).reset_index(drop=True)
    return out


# ============================================================================
# CHART BUILDERS
# ============================================================================

def donut_chart(billable, nonbill, notworked, title="Overall Hours Split"):
    labels = ["Billable", "Non-Billable", "Not Worked"]
    values = [billable, nonbill, notworked]
    colors = [BILLABLE_COLOR, NONBILL_COLOR, NOTWORKED_COLOR]
    fig = go.Figure(data=[go.Pie(
        labels=labels, values=values, hole=0.62,
        marker=dict(colors=colors, line=dict(color="#FFFFFF", width=3)),
        textinfo="percent", textfont=dict(size=13, color="white", family="Arial Black"),
        hovertemplate="%{label}: %{value:.1f} hrs (%{percent})<extra></extra>",
        sort=False,
    )])
    total = billable + nonbill + notworked
    util = (billable / total * 100) if total else 0
    fig.update_layout(
        title=dict(text=title, font=dict(size=16, color="#1F2340"), x=0.0, xanchor="left"),
        annotations=[dict(text=f"<b>{util:.0f}%</b><br><span style='font-size:11px;color:#6B7280'>Utilization</span>",
                           x=0.5, y=0.5, showarrow=False, font=dict(size=22, color="#1F2340"))],
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5),
        margin=dict(t=50, b=60, l=10, r=10),
        height=360,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def bar_chart_by_qa(df_period, title="Billable vs Non-Billable vs Not Worked (per QA)"):
    grp = df_period.groupby("QA Name")[["Billable Hours", "Non-Billable Hours", "Hours Not Worked"]].sum().reset_index()
    grp = grp.sort_values("Billable Hours", ascending=True)

    fig = go.Figure()
    fig.add_trace(go.Bar(y=grp["QA Name"], x=grp["Billable Hours"], name="Billable",
                          orientation="h", marker_color=BILLABLE_COLOR,
                          hovertemplate="Billable: %{x:.1f} hrs<extra></extra>"))
    fig.add_trace(go.Bar(y=grp["QA Name"], x=grp["Non-Billable Hours"], name="Non-Billable",
                          orientation="h", marker_color=NONBILL_COLOR,
                          hovertemplate="Non-Billable: %{x:.1f} hrs<extra></extra>"))
    fig.add_trace(go.Bar(y=grp["QA Name"], x=grp["Hours Not Worked"], name="Not Worked",
                          orientation="h", marker_color=NOTWORKED_COLOR,
                          hovertemplate="Not Worked: %{x:.1f} hrs<extra></extra>"))
    fig.update_layout(
        barmode="stack",
        title=dict(text=title, font=dict(size=16, color="#1F2340"), x=0.0, xanchor="left", y=0.98, yanchor="top"),
        xaxis=dict(title="Hours", gridcolor="#EEF0FA"),
        yaxis=dict(title=""),
        legend=dict(orientation="h", yanchor="bottom", y=1.14, xanchor="center", x=0.5),
        margin=dict(t=95, b=40, l=10, r=10),
        height=max(360, 42 * len(grp) + 140),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def trend_column_chart(df_period, granularity, title="Total Hours Trend"):
    d = df_period.copy()
    if granularity == "Daily":
        d["bucket"] = d["Date"].dt.strftime("%d %b")
        order_key = d["Date"]
    elif granularity == "Monthly":
        d["bucket"] = d["Date"].dt.strftime("%b %Y")
        order_key = d["Date"]
    else:  # Yearly
        d["bucket"] = d["Date"].dt.year.astype(str)
        order_key = d["Date"]

    grp = d.groupby("bucket").agg(
        Billable=("Billable Hours", "sum"),
        NonBillable=("Non-Billable Hours", "sum"),
        NotWorked=("Hours Not Worked", "sum"),
        order=("Date", "min"),
    ).reset_index().sort_values("order")

    fig = go.Figure()
    fig.add_trace(go.Bar(x=grp["bucket"], y=grp["Billable"], name="Billable", marker_color=BILLABLE_COLOR))
    fig.add_trace(go.Bar(x=grp["bucket"], y=grp["NonBillable"], name="Non-Billable", marker_color=NONBILL_COLOR))
    fig.add_trace(go.Bar(x=grp["bucket"], y=grp["NotWorked"], name="Not Worked", marker_color=NOTWORKED_COLOR))
    fig.update_layout(
        barmode="stack",
        title=dict(text=title, font=dict(size=16, color="#1F2340"), x=0.0, xanchor="left", y=0.98, yanchor="top"),
        xaxis=dict(title="", gridcolor="#EEF0FA"),
        yaxis=dict(title="Hours", gridcolor="#EEF0FA"),
        legend=dict(orientation="h", yanchor="bottom", y=1.16, xanchor="center", x=0.5),
        margin=dict(t=95, b=40, l=10, r=10),
        height=400,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def qa_mini_donut(row, qa_name):
    b, n, w = row["Billable Hours"], row["Non-Billable Hours"], row["Hours Not Worked"]
    fig = go.Figure(data=[go.Pie(
        labels=["Billable", "Non-Billable", "Not Worked"], values=[b, n, w], hole=0.65,
        marker=dict(colors=[BILLABLE_COLOR, NONBILL_COLOR, NOTWORKED_COLOR], line=dict(color="white", width=2)),
        textinfo="none",
        hovertemplate="%{label}: %{value:.1f} hrs<extra></extra>",
        sort=False,
    )])
    total = b + n + w
    util = (b / total * 100) if total else 0
    fig.update_layout(
        annotations=[dict(text=f"<b>{util:.0f}%</b>", x=0.5, y=0.5, showarrow=False,
                           font=dict(size=16, color="#1F2340"))],
        showlegend=False, margin=dict(t=0, b=0, l=0, r=0), height=140,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


# ============================================================================
# EXPORT HELPERS
# ============================================================================

def _fig_to_png_bytes(fig, width=900, height=500, scale=2):
    """Render a Plotly figure to PNG bytes (requires kaleido)."""
    try:
        return fig.to_image(format="png", width=width, height=height, scale=scale)
    except Exception:
        return None


def to_excel_bytes(summary_df, detail_df, kpis, period_label, chart_figs, qa_mini_figs):
    """Build an Excel workbook that mirrors the on-screen dashboard:
    a Dashboard sheet with the KPI numbers + every chart as an embedded image,
    plus the raw Summary and Detail Data sheets for further analysis.
    Expects PRE-RENDERED (label, png_bytes) pairs for speed."""
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    wb = Workbook()

    # ---- Dashboard sheet ----
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    header_fill = PatternFill(start_color="5B5FEF", end_color="5B5FEF", fill_type="solid")
    title_font = Font(size=16, bold=True, color="1F2340")
    kpi_label_font = Font(size=10, bold=True, color="6B7280")
    kpi_val_font = Font(size=18, bold=True, color="1F2340")

    ws["B2"] = "QA Work Hours Analytics Dashboard"
    ws["B2"].font = title_font
    ws["B3"] = f"Period: {period_label}   |   Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}"
    ws["B3"].font = Font(size=10, color="6B7280")

    kpi_cells = [
        ("QA TEAM SIZE", f'{kpis["team_size"]}'),
        ("BILLABLE HOURS", f'{kpis["billable"]:.1f}'),
        ("NON-BILLABLE HOURS", f'{kpis["nonbill"]:.1f}'),
        ("UTILIZATION", f'{kpis["utilization"]:.1f}%'),
        ("TOTAL HOURS", f'{kpis["total"]:.1f}'),
    ]
    col_start = 2  # column B
    for i, (label, val) in enumerate(kpi_cells):
        col = get_column_letter(col_start + i * 2)
        ws[f"{col}5"] = label
        ws[f"{col}5"].font = kpi_label_font
        ws[f"{col}6"] = val
        ws[f"{col}6"].font = kpi_val_font

    row_cursor = 9
    for title, fig in chart_figs:
        ws[f"B{row_cursor}"] = title
        ws[f"B{row_cursor}"].font = Font(size=12, bold=True, color="1F2340")
        row_cursor += 1
        png = _fig_to_png_bytes(fig, width=900, height=500)
        if png:
            img = XLImage(io.BytesIO(png))
            img.width, img.height = 560, 311
            ws.add_image(img, f"B{row_cursor}")
            row_cursor += 20
        else:
            row_cursor += 2

    # Per-QA mini donuts, 4 per row
    ws[f"B{row_cursor}"] = "Individual QA Breakdown"
    ws[f"B{row_cursor}"].font = Font(size=13, bold=True, color="1F2340")
    row_cursor += 2
    col_positions = ["B", "F", "J", "N"]
    start_row = row_cursor
    for idx, (qa_name, fig) in enumerate(qa_mini_figs):
        col = col_positions[idx % 4]
        r = start_row + (idx // 4) * 14
        ws[f"{col}{r}"] = qa_name
        ws[f"{col}{r}"].font = Font(size=11, bold=True, color="1F2340")
        png = _fig_to_png_bytes(fig, width=350, height=350)
        if png:
            img = XLImage(io.BytesIO(png))
            img.width, img.height = 220, 220
            ws.add_image(img, f"{col}{r + 1}")

    ws.column_dimensions["A"].width = 2

    # ---- Data sheets ----
    ws_summary = wb.create_sheet("QA Summary")
    ws_summary.append(list(summary_df.columns))
    for cell in ws_summary[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    for row in summary_df.round(2).itertuples(index=False):
        ws_summary.append(list(row))
    for col_cells in ws_summary.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws_summary.column_dimensions[col_cells[0].column_letter].width = max(12, length + 2)

    ws_detail = wb.create_sheet("Detail Data")
    ws_detail.append(list(detail_df.columns))
    for cell in ws_detail[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    for row in detail_df.itertuples(index=False):
        row = list(row)
        # format Date nicely
        ws_detail.append(row)
    for col_cells in ws_detail.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws_detail.column_dimensions[col_cells[0].column_letter].width = max(12, min(length + 2, 30))

    wb.save(buf)
    buf.seek(0)
    return buf


def to_pdf_bytes(summary_df, kpis, period_label, chart_figs, qa_mini_figs):
    """Build a PDF that visually mirrors the dashboard: KPI cards, then every
    team chart as an image, then a grid of per-QA mini donut cards, then the
    summary table. Expects PRE-RENDERED (label, png_bytes) pairs for speed."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                     Spacer, Image as RLImage, PageBreak)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                             topMargin=16 * mm, bottomMargin=14 * mm,
                             leftMargin=14 * mm, rightMargin=14 * mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title2", parent=styles["Title"], fontSize=18,
                                  textColor=rl_colors.HexColor("#1F2340"))
    sub_style = ParagraphStyle("Sub", parent=styles["Normal"], fontSize=10,
                                textColor=rl_colors.HexColor("#6B7280"))
    h2_style = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=13,
                               textColor=rl_colors.HexColor("#1F2340"), spaceBefore=6, spaceAfter=6)

    elements = [
        Paragraph("QA Work Hours Dashboard Report", title_style),
        Paragraph(f"Period: {period_label} &nbsp;|&nbsp; Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}", sub_style),
        Spacer(1, 10),
    ]

    kpi_data = [["QA Team Size", "Billable Hours", "Non-Billable Hours", "Utilization %", "Total Hours"],
                [str(kpis["team_size"]), f'{kpis["billable"]:.1f}', f'{kpis["nonbill"]:.1f}',
                 f'{kpis["utilization"]:.1f}%', f'{kpis["total"]:.1f}']]
    kpi_table = Table(kpi_data, colWidths=[150] * 5)
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#5B5FEF")),
        ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BACKGROUND", (0, 1), (-1, 1), rl_colors.HexColor("#F7F8FC")),
        ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.HexColor("#ECEDF7")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(kpi_table)
    elements.append(Spacer(1, 14))

    # Team charts, two per row (already rendered upstream)
    chart_imgs = []
    for title, fig in chart_figs:
        png = _fig_to_png_bytes(fig, width=800, height=460)
        if png:
            chart_imgs.append((title, png))

    elements.append(Paragraph("Team Overview", h2_style))
    row_imgs = []
    for i, (title, png) in enumerate(chart_imgs):
        img = RLImage(io.BytesIO(png), width=370, height=213)
        cell = [Paragraph(title, ParagraphStyle("ct", fontSize=9, textColor=rl_colors.HexColor("#1F2340"))), img]
        row_imgs.append(cell)
        if len(row_imgs) == 2:
            t = Table([row_imgs], colWidths=[390, 390])
            t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
            elements.append(t)
            elements.append(Spacer(1, 8))
            row_imgs = []
    if row_imgs:
        t = Table([row_imgs], colWidths=[390] * len(row_imgs))
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        elements.append(t)

    elements.append(PageBreak())

    # Per-QA mini donuts grid
    elements.append(Paragraph("Individual QA Breakdown", h2_style))
    mini_row = []
    for i, (qa_name, fig) in enumerate(qa_mini_figs):
        png = _fig_to_png_bytes(fig, width=260, height=260)
        img = RLImage(io.BytesIO(png), width=110, height=110) if png else Paragraph("", styles["Normal"])
        cell = [Paragraph(f"<b>{qa_name}</b>", ParagraphStyle("qn", fontSize=9,
                           textColor=rl_colors.HexColor("#1F2340"), alignment=1)), img]
        mini_row.append(cell)
        if len(mini_row) == 5:
            t = Table([mini_row], colWidths=[156] * 5)
            t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
            elements.append(t)
            elements.append(Spacer(1, 6))
            mini_row = []
    if mini_row:
        t = Table([mini_row], colWidths=[156] * len(mini_row))
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
        elements.append(t)

    elements.append(PageBreak())
    elements.append(Paragraph("Per-QA Summary Table", h2_style))
    elements.append(Spacer(1, 4))

    table_data = [list(summary_df.columns)] + summary_df.round(1).astype(str).values.tolist()
    n_cols = len(summary_df.columns)
    col_width = 780 / n_cols
    tbl = Table(table_data, colWidths=[col_width] * n_cols, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#1F2340")),
        ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#ECEDF7")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F7F8FC")]),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elements.append(tbl)

    doc.build(elements)
    buf.seek(0)
    return buf


# ============================================================================
# APP HEADER
# ============================================================================

st.markdown('<div class="app-title">\U0001F4CA QA Work Hours Analytics Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">Upload the QA hours tracker workbook to get an instant, management-ready view of team utilization.</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("Upload QA Work Hours Excel file (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("\U0001F446 Upload your QA hours tracker workbook to begin. Each QA's own sheet will be detected and combined automatically.")
    st.stop()

with st.spinner("Reading and analyzing the workbook..."):
    data = parse_workbook(uploaded)

if data.empty:
    st.error("Couldn't find any recognizable QA hour records in this workbook. Please check the sheet format.")
    st.stop()

# ============================================================================
# SIDEBAR FILTERS
# ============================================================================
st.sidebar.header("\U0001F50D Filters")

view_mode = st.sidebar.radio("View by", ["Daily", "Monthly", "Yearly"], index=1)

years_available = sorted(data["Year"].dropna().unique().astype(int).tolist())
sel_year = st.sidebar.selectbox("Year", years_available, index=len(years_available) - 1)

df_y = data[data["Year"] == sel_year]

if view_mode == "Monthly":
    months_available = [m for m in MONTH_ORDER if m in df_y["Month"].unique()]
    sel_month = st.sidebar.selectbox("Month", months_available, index=len(months_available) - 1 if months_available else 0)
    df_period = df_y[df_y["Month"] == sel_month]
    period_label = f"{sel_month} {sel_year}"
elif view_mode == "Daily":
    dates_available = sorted(df_y["Date"].dt.date.unique().tolist())
    if dates_available:
        sel_date = st.sidebar.selectbox("Date", dates_available, index=len(dates_available) - 1)
        df_period = df_y[df_y["Date"].dt.date == sel_date]
        period_label = sel_date.strftime("%d %b %Y")
    else:
        df_period = df_y
        period_label = f"{sel_year}"
else:  # Yearly
    df_period = df_y
    period_label = f"{sel_year}"

qa_list = sorted(df_period["QA Name"].unique().tolist())
sel_qas = st.sidebar.multiselect("QA Team Members", qa_list, default=qa_list)
df_period = df_period[df_period["QA Name"].isin(sel_qas)]

if df_period.empty:
    st.warning("No data for the selected filters.")
    st.stop()

# ============================================================================
# KPI SUMMARY
# ============================================================================
team_size = df_period["QA Name"].nunique()
billable_total = df_period["Billable Hours"].sum()
nonbill_total = df_period["Non-Billable Hours"].sum()
notworked_total = df_period["Hours Not Worked"].sum()
total_hours = df_period["Total Hours"].sum()
utilization = (billable_total / total_hours * 100) if total_hours > 0 else 0

kpis = dict(team_size=team_size, billable=billable_total, nonbill=nonbill_total,
            utilization=utilization, total=total_hours)

st.markdown(f'<div class="section-header">Summary &nbsp;·&nbsp; {period_label}</div>', unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)
kpi_cells = [
    (k1, "QA Team Size", f"{team_size}", "active members"),
    (k2, "Billable Hours", f"{billable_total:,.1f}", "hrs logged"),
    (k3, "Non-Billable Hours", f"{nonbill_total:,.1f}", "hrs logged"),
    (k4, "Utilization", f"{utilization:.1f}%", "billable / total"),
    (k5, "Total Hours", f"{total_hours:,.1f}", "hrs tracked"),
]
for col, label, val, sub in kpi_cells:
    with col:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{val}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

# ============================================================================
# CHARTS - TEAM LEVEL
# ============================================================================
st.markdown('<div class="section-header">Team Overview</div>', unsafe_allow_html=True)

# Build every team chart once, so the same figure objects power both the
# on-screen dashboard and the Excel/PDF exports (what you see is what you get).
fig_donut = donut_chart(billable_total, nonbill_total, notworked_total)
fig_bar = bar_chart_by_qa(df_period)
fig_trend = trend_column_chart(df_period, view_mode)

# Donut centered in the middle column of a 3-column row
d1, d2, d3 = st.columns([1, 2, 1])
with d2:
    st.plotly_chart(fig_donut, use_container_width=True, config=PLOTLY_CONFIG)

# Bar and Trend charts go full width, stacked
st.plotly_chart(fig_bar, use_container_width=True, config=PLOTLY_CONFIG)
st.plotly_chart(fig_trend, use_container_width=True, config=PLOTLY_CONFIG)

team_chart_figs = [
    ("Overall Hours Split (Donut)", fig_donut),
    ("Billable vs Non-Billable vs Not Worked (Bar)", fig_bar),
    ("Total Hours Trend", fig_trend),
]

# ============================================================================
# PER-QA CARDS
# ============================================================================
st.markdown('<div class="section-header">Individual QA Breakdown</div>', unsafe_allow_html=True)

qa_summary = df_period.groupby("QA Name").agg(
    **{
        "Billable Hours": ("Billable Hours", "sum"),
        "Non-Billable Hours": ("Non-Billable Hours", "sum"),
        "Hours Not Worked": ("Hours Not Worked", "sum"),
        "Total Hours": ("Total Hours", "sum"),
        "Days Logged": ("Date", "nunique"),
    }
).reset_index()
qa_summary["Utilization %"] = np.where(
    qa_summary["Total Hours"] > 0,
    qa_summary["Billable Hours"] / qa_summary["Total Hours"] * 100,
    0,
)
qa_summary = qa_summary.sort_values("Utilization %", ascending=False)

n_cols = 3
qa_rows = [qa_summary.iloc[i:i + n_cols] for i in range(0, len(qa_summary), n_cols)]

qa_mini_figs = []  # collected for export reuse
for chunk in qa_rows:
    cols = st.columns(n_cols)
    for col, (_, row) in zip(cols, chunk.iterrows()):
        with col:
            util = row["Utilization %"]
            if util >= 70:
                badge = '<span class="badge badge-good">Good</span>'
            elif util >= 40:
                badge = '<span class="badge badge-warn">Watch</span>'
            else:
                badge = '<span class="badge badge-bad">Low</span>'
            st.markdown(f"""
            <div class="qa-card">
                <div class="qa-name">{row['QA Name']} {badge}</div>
                <div class="kpi-sub">Days logged: {int(row['Days Logged'])} &nbsp;|&nbsp; Total: {row['Total Hours']:.1f} hrs</div>
            </div>
            """, unsafe_allow_html=True)
            mini_fig = qa_mini_donut(row, row["QA Name"])
            st.plotly_chart(mini_fig, use_container_width=True, config=PLOTLY_CONFIG,
                             key=f"mini_{row['QA Name']}_{period_label}")
            qa_mini_figs.append((row["QA Name"], mini_fig))

# ============================================================================
# DETAIL TABLE
# ============================================================================
with st.expander("\U0001F4CB View Detailed QA Summary Table"):
    st.dataframe(
        qa_summary[["QA Name", "Billable Hours", "Non-Billable Hours", "Hours Not Worked",
                     "Total Hours", "Utilization %", "Days Logged"]].round(1),
        use_container_width=True, hide_index=True,
    )

with st.expander("\U0001F5C2\uFE0F View Raw Daily Records (Selected Period)"):
    st.dataframe(
        df_period[["QA Name", "Date", "Day", "Month", "Billable Hours",
                    "Non-Billable Hours", "Hours Not Worked", "Total Hours"]]
        .sort_values(["QA Name", "Date"]).round(1),
        use_container_width=True, hide_index=True,
    )

# ============================================================================
# EXPORTS
# ============================================================================
st.markdown('<div class="section-header">Export</div>', unsafe_allow_html=True)
e1, e2 = st.columns(2)

export_summary = qa_summary[["QA Name", "Billable Hours", "Non-Billable Hours", "Hours Not Worked",
                              "Total Hours", "Utilization %", "Days Logged"]].round(2)
export_detail = df_period[["QA Name", "Date", "Day", "Month", "Billable Hours",
                            "Non-Billable Hours", "Hours Not Worked", "Total Hours"]].sort_values(["QA Name", "Date"])

with e1:
    try:
        with st.spinner("Building Excel report with charts..."):
            excel_bytes = to_excel_bytes(export_summary, export_detail, kpis, period_label,
                                          team_chart_figs, qa_mini_figs)
        st.download_button(
            "\U0001F4E5 Export to Excel",
            data=excel_bytes,
            file_name=f"QA_Dashboard_{period_label.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"Excel export failed: {e}")
with e2:
    try:
        with st.spinner("Building PDF report with charts..."):
            pdf_bytes = to_pdf_bytes(export_summary, kpis, period_label, team_chart_figs, qa_mini_figs)
        st.download_button(
            "\U0001F4C4 Export to PDF",
            data=pdf_bytes,
            file_name=f"QA_Dashboard_{period_label.replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"PDF export unavailable: {e}")

st.caption("Built for QA Management \u00b7 Streamlit Dashboard \u00b7 Ready for desktop packaging (Windows/Mac)")