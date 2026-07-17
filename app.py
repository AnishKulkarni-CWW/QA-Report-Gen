"""
QA Work-Hours Analytics Dashboard
----------------------------------
A single-page Streamlit dashboard for analyzing QA team work-hour trackers.

Upload the QA tracker workbook -> instant KPI summary + charts (line trend,
daily intensity grid, QA comparison bars, hours mix, searchable daily log) ->
filter by QA / Year / Month / Week / Day (multi-select + calendar) -> click
"Prepare Export" once to build Excel/PDF snapshots of exactly what's on screen.

Run:  streamlit run app.py
"""

import io
import re
import warnings
from datetime import datetime, date

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ============================================================================
# PAGE CONFIG & GLOBAL STYLE
# ============================================================================
st.set_page_config(
    page_title="QA Hours Dashboard",
    page_icon="\U0001F4CA",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---- Palette (matches the reference screenshot: dark navy header, teal/slate accents) ----
INK = "#0F1729"           # dark navy header background
INK_TEXT = "#E7EAF3"
PAGE_BG = "#F4F5F8"
CARD_BG = "#FFFFFF"
BORDER = "#E7E9F0"
TEXT_MAIN = "#1F2333"
TEXT_MUTED = "#7A7F91"

BILLABLE_COLOR = "#2F9E8F"     # teal
NONBILL_COLOR = "#E0A72E"      # amber
NOTWORKED_COLOR = "#C1543D"    # brick red
ACCENT = "#2F9E8F"

QA_PALETTE = ["#2F9E8F", "#D97A46", "#6E7FC9", "#C1543D",
              "#8E9257", "#DCB13A", "#4C7A96", "#9C6FA6"]

CUSTOM_CSS = f"""
<style>
    .stApp {{
        background: {PAGE_BG};
    }}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header[data-testid="stHeader"] {{
        background: transparent;
        height: 3rem;
    }}
    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 3rem !important;
        max-width: 1400px;
    }}

    /* Top banner mimicking the reference design */
    .top-banner {{
        background: {INK};
        border-radius: 16px;
        padding: 18px 28px;
        margin-bottom: 20px;
        color: {INK_TEXT};
    }}
    .top-banner .eyebrow {{
        font-size: 0.7rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: #8B93AD;
        font-weight: 600;
        margin-bottom: 2px;
    }}
    .top-banner .title {{
        font-size: 1.5rem;
        font-weight: 800;
        color: #FFFFFF;
    }}

    .kpi-card {{
        background: {CARD_BG};
        border-radius: 14px;
        padding: 16px 18px 12px 18px;
        box-shadow: 0 1px 3px rgba(15,23,41,0.06);
        border: 1px solid {BORDER};
        height: 108px;
    }}
    .kpi-label {{
        font-size: 0.72rem;
        font-weight: 700;
        color: {TEXT_MUTED};
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 6px;
    }}
    .kpi-value {{
        font-size: 1.8rem;
        font-weight: 800;
        color: {TEXT_MAIN};
        line-height: 1.1;
    }}
    .kpi-value .unit {{
        font-size: 0.95rem;
        font-weight: 600;
        color: {TEXT_MUTED};
        margin-left: 3px;
    }}
    .kpi-sub {{
        font-size: 0.74rem;
        color: {TEXT_MUTED};
        margin-top: 4px;
    }}

    .panel {{
        background: {CARD_BG};
        border-radius: 16px;
        padding: 20px 22px;
        box-shadow: 0 1px 3px rgba(15,23,41,0.06);
        border: 1px solid {BORDER};
        margin-bottom: 20px;
    }}
    .panel-title {{
        font-size: 1.05rem;
        font-weight: 800;
        color: {TEXT_MAIN};
    }}
    .panel-sub {{
        font-size: 0.78rem;
        color: {TEXT_MUTED};
    }}

    .qa-card {{
        background: {CARD_BG};
        border-radius: 14px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(15,23,41,0.06);
        border: 1px solid {BORDER};
        margin-bottom: 12px;
    }}
    .qa-name {{
        font-size: 1.0rem;
        font-weight: 800;
        color: {TEXT_MAIN};
    }}

    div[data-testid="stFileUploader"] {{
        background: {CARD_BG};
        border-radius: 14px;
        padding: 12px;
        border: 1.5px dashed #C9CEDD;
    }}

    .modebar {{ display: none !important; }}

    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background: {CARD_BG};
        border-right: 1px solid {BORDER};
    }}

    .stButton > button {{
        border-radius: 10px;
        font-weight: 700;
    }}
    div[data-testid="stDownloadButton"] > button {{
        border-radius: 10px;
        font-weight: 700;
        background: {INK};
        color: white;
        border: none;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Disable Plotly's floating toolbar everywhere (prevents accidental clicks).
PLOTLY_CONFIG = {"displayModeBar": False, "displaylogo": False, "scrollZoom": False}

TODAY = pd.Timestamp(datetime.now().date())

# ============================================================================
# DATA LOADING & CLEANING
# ============================================================================

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
    """Read every per-QA sheet in the workbook and return one tidy dataframe.
    Rows dated after today are dropped, even if present in the source file."""
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
            continue

        body = raw.iloc[header_row + 1:].copy()
        body = body.rename(columns=col_map)
        keep_cols = [c for c in body.columns if isinstance(c, str) and c in HEADER_ALIASES.values()]
        body = body[keep_cols]
        body = body.loc[:, ~body.columns.duplicated()]

        if "QA Name" not in body.columns:
            body["QA Name"] = sheet
        body["QA Name"] = body["QA Name"].fillna(sheet)
        body.loc[body["QA Name"].astype(str).str.strip() == "", "QA Name"] = sheet

        if "Date" in body.columns:
            body["Date"] = pd.to_datetime(body["Date"], errors="coerce")
        else:
            body["Date"] = pd.NaT

        body = body[body["Date"].notna()]
        if body.empty:
            continue

        for c in ["Billable Hours", "Non-Billable Hours", "Hours Not Worked", "Total Hours"]:
            if c in body.columns:
                body[c] = body[c].apply(_to_number)
            else:
                body[c] = np.nan

        hour_cols = ["Billable Hours", "Non-Billable Hours", "Hours Not Worked", "Total Hours"]
        body = body[~body[hour_cols].isna().all(axis=1)]
        if body.empty:
            continue

        body[hour_cols] = body[hour_cols].fillna(0.0)

        computed_total = body["Billable Hours"] + body["Non-Billable Hours"] + body["Hours Not Worked"]
        body["Total Hours"] = np.where(
            (body["Total Hours"] <= 0) | (body["Total Hours"].isna()),
            computed_total,
            body["Total Hours"],
        )

        body["Day"] = body["Date"].dt.day_name().str.slice(0, 3)
        body["Month"] = body["Date"].dt.strftime("%b")
        body["Year"] = body["Date"].dt.year

        if "Comment" in body.columns:
            body["Comment"] = body["Comment"].fillna("").astype(str).str.strip()
        else:
            body["Comment"] = ""

        body["QA Name"] = body["QA Name"].astype(str).str.strip()
        body["QA Name"] = body["QA Name"].str.replace(r"[._]+", " ", regex=True)
        body["QA Name"] = body["QA Name"].str.replace(r"\s+", " ", regex=True).str.strip()
        body["QA Name"] = body["QA Name"].str.title()

        final = body[["QA Name", "Date", "Day", "Month", "Year",
                      "Billable Hours", "Non-Billable Hours",
                      "Hours Not Worked", "Total Hours", "Comment"]].copy()
        final = final[final["QA Name"] != ""]
        frames.append(final)

    if not frames:
        return pd.DataFrame(columns=["QA Name", "Date", "Day", "Month", "Year",
                                      "Billable Hours", "Non-Billable Hours",
                                      "Hours Not Worked", "Total Hours", "Comment"])

    out = pd.concat(frames, ignore_index=True)
    out = out.drop_duplicates(subset=["QA Name", "Date"], keep="first")

    # Hard rule: never show data beyond today, even if the workbook has it.
    out = out[out["Date"] <= TODAY]

    out = out.sort_values(["QA Name", "Date"]).reset_index(drop=True)
    return out


# ============================================================================
# CHART BUILDERS
# ============================================================================

def qa_color_map(qa_names):
    return {name: QA_PALETTE[i % len(QA_PALETTE)] for i, name in enumerate(sorted(qa_names))}


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
        title=dict(text=title, font=dict(size=15, color=TEXT_MAIN), x=0.0, xanchor="left"),
        annotations=[dict(text=f"<b>{util:.0f}%</b><br><span style='font-size:11px;color:#7A7F91'>Utilization</span>",
                           x=0.5, y=0.5, showarrow=False, font=dict(size=20, color=TEXT_MAIN))],
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.08, xanchor="center", x=0.5),
        margin=dict(t=50, b=55, l=10, r=10),
        height=360,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def bar_chart_by_qa(df_period, title="QA Comparison — Total Hours"):
    grp = df_period.groupby("QA Name")["Total Hours"].sum().reset_index()
    grp = grp.sort_values("Total Hours", ascending=True)
    colors = qa_color_map(grp["QA Name"].unique())

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=grp["QA Name"], x=grp["Total Hours"], orientation="h",
        marker_color=[colors[n] for n in grp["QA Name"]],
        text=[f"{v:,.1f}" for v in grp["Total Hours"]],
        textposition="outside",
        hovertemplate="%{y}: %{x:.1f} hrs<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color=TEXT_MAIN), x=0.0, xanchor="left"),
        xaxis=dict(title="Hours", gridcolor="#F0F1F5"),
        yaxis=dict(title=""),
        showlegend=False,
        margin=dict(t=50, b=20, l=10, r=40),
        height=max(320, 42 * len(grp) + 90),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def hours_mix_chart(df_period, title="Hours Mix — Billable / Non-Billable / Not Worked"):
    grp = df_period.groupby("QA Name")[["Billable Hours", "Non-Billable Hours", "Hours Not Worked"]].sum().reset_index()
    grp = grp.sort_values("Billable Hours", ascending=False)

    fig = go.Figure()
    fig.add_trace(go.Bar(x=grp["QA Name"], y=grp["Billable Hours"], name="Billable", marker_color=BILLABLE_COLOR,
                          hovertemplate="Billable: %{y:.1f} hrs<extra></extra>"))
    fig.add_trace(go.Bar(x=grp["QA Name"], y=grp["Non-Billable Hours"], name="Non-Billable", marker_color=NONBILL_COLOR,
                          hovertemplate="Non-Billable: %{y:.1f} hrs<extra></extra>"))
    fig.add_trace(go.Bar(x=grp["QA Name"], y=grp["Hours Not Worked"], name="Not Worked", marker_color=NOTWORKED_COLOR,
                          hovertemplate="Not Worked: %{y:.1f} hrs<extra></extra>"))
    fig.update_layout(
        barmode="stack",
        title=dict(text=title, font=dict(size=15, color=TEXT_MAIN), x=0.0, xanchor="left"),
        xaxis=dict(title="", gridcolor="#F0F1F5"),
        yaxis=dict(title="Hours", gridcolor="#F0F1F5"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
        margin=dict(t=50, b=10, l=10, r=10),
        height=420,
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
                           font=dict(size=16, color=TEXT_MAIN))],
        showlegend=False, margin=dict(t=0, b=0, l=0, r=0), height=140,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


# ============================================================================
# EXPORT HELPERS  (only ever called after the user clicks "Prepare Export")
# ============================================================================

def _fig_to_png_bytes(fig, width=900, height=500, scale=2):
    try:
        return fig.to_image(format="png", width=width, height=height, scale=scale)
    except Exception:
        return None


def to_excel_bytes(summary_df, detail_df, kpis, period_label, chart_figs, qa_mini_figs):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    wb = Workbook()

    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    header_fill = PatternFill(start_color="0F1729", end_color="0F1729", fill_type="solid")
    title_font = Font(size=16, bold=True, color="0F1729")
    kpi_label_font = Font(size=10, bold=True, color="7A7F91")
    kpi_val_font = Font(size=18, bold=True, color="0F1729")

    ws["B2"] = "QA Work Hours Dashboard"
    ws["B2"].font = title_font
    ws["B3"] = f"Period: {period_label}   |   Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}"
    ws["B3"].font = Font(size=10, color="7A7F91")

    kpi_cells = [
        ("QA TEAM SIZE", f'{kpis["team_size"]}'),
        ("BILLABLE HOURS", f'{kpis["billable"]:.1f}'),
        ("NON-BILLABLE HOURS", f'{kpis["nonbill"]:.1f}'),
        ("UTILIZATION", f'{kpis["utilization"]:.1f}%'),
        ("TOTAL HOURS", f'{kpis["total"]:.1f}'),
    ]
    for i, (label, val) in enumerate(kpi_cells):
        col = get_column_letter(2 + i * 2)
        ws[f"{col}5"] = label
        ws[f"{col}5"].font = kpi_label_font
        ws[f"{col}6"] = val
        ws[f"{col}6"].font = kpi_val_font

    row_cursor = 9
    for title, fig in chart_figs:
        ws[f"B{row_cursor}"] = title
        ws[f"B{row_cursor}"].font = Font(size=12, bold=True, color="0F1729")
        row_cursor += 1
        png = _fig_to_png_bytes(fig, width=900, height=500)
        if png:
            img = XLImage(io.BytesIO(png))
            img.width, img.height = 560, 311
            ws.add_image(img, f"B{row_cursor}")
            row_cursor += 20
        else:
            row_cursor += 2

    ws[f"B{row_cursor}"] = "Individual QA Breakdown"
    ws[f"B{row_cursor}"].font = Font(size=13, bold=True, color="0F1729")
    row_cursor += 2
    col_positions = ["B", "F", "J", "N"]
    start_row = row_cursor
    for idx, (qa_name, fig) in enumerate(qa_mini_figs):
        col = col_positions[idx % 4]
        r = start_row + (idx // 4) * 14
        ws[f"{col}{r}"] = qa_name
        ws[f"{col}{r}"].font = Font(size=11, bold=True, color="0F1729")
        png = _fig_to_png_bytes(fig, width=350, height=350)
        if png:
            img = XLImage(io.BytesIO(png))
            img.width, img.height = 220, 220
            ws.add_image(img, f"{col}{r + 1}")

    ws.column_dimensions["A"].width = 2

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
        ws_detail.append(list(row))
    for col_cells in ws_detail.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws_detail.column_dimensions[col_cells[0].column_letter].width = max(12, min(length + 2, 30))

    wb.save(buf)
    buf.seek(0)
    return buf


def to_pdf_bytes(summary_df, kpis, period_label, chart_figs, qa_mini_figs):
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
                                  textColor=rl_colors.HexColor("#0F1729"))
    sub_style = ParagraphStyle("Sub", parent=styles["Normal"], fontSize=10,
                                textColor=rl_colors.HexColor("#7A7F91"))
    h2_style = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=13,
                               textColor=rl_colors.HexColor("#0F1729"), spaceBefore=6, spaceAfter=6)

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
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0F1729")),
        ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BACKGROUND", (0, 1), (-1, 1), rl_colors.HexColor("#F4F5F8")),
        ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.HexColor("#E7E9F0")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(kpi_table)
    elements.append(Spacer(1, 14))

    chart_imgs = []
    for title, fig in chart_figs:
        png = _fig_to_png_bytes(fig, width=800, height=460)
        if png:
            chart_imgs.append((title, png))

    elements.append(Paragraph("Team Overview", h2_style))
    row_imgs = []
    for i, (title, png) in enumerate(chart_imgs):
        img = RLImage(io.BytesIO(png), width=370, height=213)
        cell = [Paragraph(title, ParagraphStyle("ct", fontSize=9, textColor=rl_colors.HexColor("#0F1729"))), img]
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
    elements.append(Paragraph("Individual QA Breakdown", h2_style))
    mini_row = []
    for i, (qa_name, fig) in enumerate(qa_mini_figs):
        png = _fig_to_png_bytes(fig, width=260, height=260)
        img = RLImage(io.BytesIO(png), width=110, height=110) if png else Paragraph("", styles["Normal"])
        cell = [Paragraph(f"<b>{qa_name}</b>", ParagraphStyle("qn", fontSize=9,
                           textColor=rl_colors.HexColor("#0F1729"), alignment=1)), img]
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
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0F1729")),
        ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#E7E9F0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F4F5F8")]),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elements.append(tbl)

    doc.build(elements)
    buf.seek(0)
    return buf


# ============================================================================
# APP HEADER / UPLOAD
# ============================================================================

st.markdown("""
<div class="top-banner">
    <div class="eyebrow">QA MANAGEMENT SYSTEM &middot; V2</div>
    <div class="title">QA Hours Dashboard</div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload QA Work Hours Excel file (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("\U0001F446 Upload your QA hours tracker workbook to begin. Each QA's own sheet is detected and combined automatically. Only data up to today's date will ever be shown.")
    st.stop()

with st.spinner("Reading and analyzing the workbook..."):
    data = parse_workbook(uploaded)

if data.empty:
    st.error("Couldn't find any recognizable QA hour records in this workbook (or all rows were in the future). Please check the sheet format.")
    st.stop()

# ============================================================================
# SIDEBAR FILTERS
# ============================================================================
st.sidebar.header("\U0001F50D Filters")

view_mode = st.sidebar.radio("View by", ["Daily", "Weekly", "Monthly", "Yearly"], index=2)

# ---- QA selection: a real closed dropdown (multiselect), with quick All/None ----
all_qas = sorted(data["QA Name"].unique().tolist())

if "qa_multiselect" not in st.session_state:
    st.session_state.qa_multiselect = all_qas.copy()
# keep state in sync if a new file introduces different QAs
st.session_state.qa_multiselect = [qa for qa in st.session_state.qa_multiselect if qa in all_qas]

cA, cB = st.sidebar.columns(2)
if cA.button("Select All QAs", use_container_width=True):
    st.session_state.qa_multiselect = all_qas.copy()
if cB.button("Clear QAs", use_container_width=True):
    st.session_state.qa_multiselect = []

sel_qas = st.sidebar.multiselect("QA Team Members", all_qas, key="qa_multiselect")

st.sidebar.markdown("---")

# ---- Time period selection depending on view mode ----
years_available = sorted(data["Year"].dropna().unique().astype(int).tolist())

if view_mode == "Yearly":
    sel_years = st.sidebar.multiselect("Years", years_available, default=years_available)
    df_period = data[data["Year"].isin(sel_years)]
    period_label = ", ".join(str(y) for y in sel_years) if sel_years else "No years selected"

elif view_mode == "Monthly":
    sel_years = st.sidebar.multiselect("Years", years_available, default=years_available)
    months_in_scope = [m for m in MONTH_ORDER if m in data[data["Year"].isin(sel_years)]["Month"].unique()]
    sel_months = st.sidebar.multiselect("Months", months_in_scope, default=months_in_scope)
    df_period = data[data["Year"].isin(sel_years) & data["Month"].isin(sel_months)]
    period_label = f"{', '.join(sel_months) if sel_months else 'No months'} ({', '.join(str(y) for y in sel_years)})"

elif view_mode == "Weekly":
    sel_years = st.sidebar.multiselect("Years", years_available, default=years_available)
    df_y = data[data["Year"].isin(sel_years)]
    df_y = df_y.assign(_week=df_y["Date"].dt.to_period("W"))
    week_options = sorted(df_y["_week"].unique())
    week_labels = {w: f"{w.start_time.strftime('%d %b')} – {w.end_time.strftime('%d %b %Y')}" for w in week_options}
    sel_weeks = st.sidebar.multiselect(
        "Weeks", week_options, default=week_options,
        format_func=lambda w: week_labels.get(w, str(w)),
    )
    df_period = df_y[df_y["_week"].isin(sel_weeks)].drop(columns="_week")
    period_label = f"{len(sel_weeks)} week(s) selected" if sel_weeks else "No weeks selected"

else:  # Daily -> calendar-style multi-date picker
    min_d, max_d = data["Date"].min().date(), min(data["Date"].max().date(), TODAY.date())
    sel_dates = st.sidebar.date_input(
        "Select date(s)", value=(min_d, max_d), min_value=min_d, max_value=max_d,
    )
    if isinstance(sel_dates, (tuple, list)) and len(sel_dates) == 2:
        d_start, d_end = sel_dates
        df_period = data[(data["Date"].dt.date >= d_start) & (data["Date"].dt.date <= d_end)]
        period_label = f"{d_start.strftime('%d %b %Y')} – {d_end.strftime('%d %b %Y')}"
    elif isinstance(sel_dates, (tuple, list)) and len(sel_dates) == 1:
        d_only = sel_dates[0]
        df_period = data[data["Date"].dt.date == d_only]
        period_label = d_only.strftime("%d %b %Y")
    else:
        df_period = data[data["Date"].dt.date == sel_dates]
        period_label = sel_dates.strftime("%d %b %Y")

df_period = df_period[df_period["QA Name"].isin(sel_qas)]

if df_period.empty:
    st.warning("No data for the selected filters. Adjust the QA / time filters in the sidebar.")
    st.stop()

qa_colors = qa_color_map(all_qas)

# ============================================================================
# KPI SUMMARY
# ============================================================================
# "QA Team Size" reflects how many QAs are CHECKED in the filter — not just
# how many happen to have logged rows within the current date/month/year
# window (a QA with zero hours in a narrow window should still count as
# selected, matching what the sidebar shows).
team_size = len(sel_qas)
billable_total = df_period["Billable Hours"].sum()
nonbill_total = df_period["Non-Billable Hours"].sum()
notworked_total = df_period["Hours Not Worked"].sum()
total_hours = df_period["Total Hours"].sum()
utilization = (billable_total / total_hours * 100) if total_hours > 0 else 0
days_logged = df_period["Date"].nunique()
avg_hours_per_day = (total_hours / (df_period.groupby("QA Name")["Date"].nunique().sum())) if len(df_period) else 0

kpis = dict(team_size=team_size, billable=billable_total, nonbill=nonbill_total,
            utilization=utilization, total=total_hours)

st.markdown(f'<div class="panel-title" style="margin-bottom:12px;">Summary &nbsp;·&nbsp; {period_label}</div>', unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)
kpi_cells = [
    (k1, "QA TEAM SIZE", f"{team_size}", "active members"),
    (k2, "TOTAL HOURS", f"{total_hours:,.1f}", f"across {days_logged} days"),
    (k3, "BILLABLE SHARE", f"{utilization:.1f}%", f"{billable_total:,.0f} of {total_hours:,.0f} hrs"),
    (k4, "AVG HOURS / DAY", f"{avg_hours_per_day:.2f}", "per active QA-day"),
    (k5, "NON-BILLABLE HOURS", f"{nonbill_total:,.1f}", "hrs logged"),
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

fig_donut = donut_chart(billable_total, nonbill_total, notworked_total)
fig_bar = bar_chart_by_qa(df_period)
fig_mix = hours_mix_chart(df_period)

with st.container(border=True):
    st.plotly_chart(fig_bar, use_container_width=True, config=PLOTLY_CONFIG)

with st.container(border=True):
    st.plotly_chart(fig_mix, use_container_width=True, config=PLOTLY_CONFIG)

# Donut, centered
d1, d2, d3 = st.columns([1, 2, 1])
with d2:
    with st.container(border=True):
        st.plotly_chart(fig_donut, use_container_width=True, config=PLOTLY_CONFIG)

team_chart_figs = [
    ("QA Comparison — Total Hours", fig_bar),
    ("Hours Mix", fig_mix),
    ("Overall Hours Split (Donut)", fig_donut),
]

# ============================================================================
# PER-QA CARDS
# ============================================================================
st.markdown('<div class="panel-title">Individual QA Breakdown</div>', unsafe_allow_html=True)
st.write("")

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

qa_mini_figs = []
for chunk in qa_rows:
    cols = st.columns(n_cols)
    for col, (_, row) in zip(cols, chunk.iterrows()):
        with col:
            st.markdown(f"""
            <div class="qa-card">
                <div class="qa-name">{row['QA Name']}</div>
                <div class="kpi-sub">Days logged: {int(row['Days Logged'])} &nbsp;|&nbsp; Total: {row['Total Hours']:.1f} hrs</div>
            </div>
            """, unsafe_allow_html=True)
            mini_fig = qa_mini_donut(row, row["QA Name"])
            st.plotly_chart(mini_fig, use_container_width=True, config=PLOTLY_CONFIG,
                             key=f"mini_{row['QA Name']}_{period_label}")
            qa_mini_figs.append((row["QA Name"], mini_fig))

# ============================================================================
# DAILY LOG TABLE
# ============================================================================
st.markdown('<div class="panel-title" style="margin-top:1.2rem;">Daily Log</div>', unsafe_allow_html=True)
st.caption("Every cleaned daily record for the QAs and period selected in the sidebar filters · sortable")

log_df = df_period.copy()

log_df_display = log_df[["Date", "Day", "QA Name", "Billable Hours", "Non-Billable Hours",
                          "Hours Not Worked", "Total Hours", "Comment"]].sort_values("Date", ascending=False)
log_df_display["Date"] = log_df_display["Date"].dt.strftime("%Y-%m-%d")

st.dataframe(log_df_display.round(2), use_container_width=True, hide_index=True, height=380)
st.caption(f"{len(log_df_display):,} rows")

with st.expander("\U0001F4CB View Per-QA Summary Table"):
    st.dataframe(
        qa_summary[["QA Name", "Billable Hours", "Non-Billable Hours", "Hours Not Worked",
                     "Total Hours", "Utilization %", "Days Logged"]].round(1),
        use_container_width=True, hide_index=True,
    )

# ============================================================================
# EXPORT — explicit "Prepare Export" gate.
# Nothing below is computed until the button is pressed, and once the export
# bytes exist they are cached in session_state so simply changing an unrelated
# filter afterwards does NOT silently regenerate/re-run the export.
# ============================================================================
st.markdown('<div class="panel-title" style="margin-top:1.2rem;">Export</div>', unsafe_allow_html=True)
st.caption("Exports always reflect exactly what is on screen right now. Nothing is generated until you click Prepare Export.")

export_summary = qa_summary[["QA Name", "Billable Hours", "Non-Billable Hours", "Hours Not Worked",
                              "Total Hours", "Utilization %", "Days Logged"]].round(2)
export_detail = df_period[["QA Name", "Date", "Day", "Month", "Billable Hours",
                            "Non-Billable Hours", "Hours Not Worked", "Total Hours", "Comment"]].sort_values(["QA Name", "Date"])

prepare_clicked = st.button("\u2699\ufe0f Prepare Exports", type="primary", use_container_width=False)

if prepare_clicked:
    with st.spinner("Building Excel and PDF reports..."):
        st.session_state["export_excel_bytes"] = to_excel_bytes(
            export_summary, export_detail, kpis, period_label, team_chart_figs, qa_mini_figs)
        st.session_state["export_pdf_bytes"] = to_pdf_bytes(
            export_summary, kpis, period_label, team_chart_figs, qa_mini_figs)
        st.session_state["export_period_label"] = period_label
    st.success("Exports ready below \u2014 changing filters now will NOT regenerate them until you click Prepare Exports again.")

if "export_excel_bytes" in st.session_state:
    e1, e2 = st.columns(2)
    fname_period = st.session_state.get("export_period_label", period_label).replace(" ", "_").replace(",", "")
    with e1:
        st.download_button(
            "\U0001F4E5 Download Excel",
            data=st.session_state["export_excel_bytes"],
            file_name=f"QA_Dashboard_{fname_period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with e2:
        st.download_button(
            "\U0001F4C4 Download PDF",
            data=st.session_state["export_pdf_bytes"],
            file_name=f"QA_Dashboard_{fname_period}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

st.caption("Built for QA Management \u00b7 Streamlit Dashboard \u00b7 Ready for desktop packaging (Windows/Mac)")