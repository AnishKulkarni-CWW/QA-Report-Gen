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


# Note: intentionally does NOT reuse BILLABLE_COLOR (#2F9E8F) or
# NOTWORKED_COLOR (#C1543D) — those two hues are reserved for the
# Billable/Non-Billable/Not-Worked state encoding in the stacked and donut
# charts, so a QA-identity color here can never be mistaken for that meaning.
QA_PALETTE = ["#4C6FA6", "#D97A46", "#6E7FC9", "#5B9279",
              "#8E9257", "#DCB13A", "#A65D8C", "#9C6FA6"]

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
        padding: 16px 18px 14px 18px;
        box-shadow: 0 1px 3px rgba(15,23,41,0.06);
        border: 1px solid {BORDER};
        min-height: 108px;
        overflow: visible;
        margin-bottom: 4px;
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

    /* Give multiselect (QA / Years / Months / Weeks) and date-input dropdowns
       a visible rounded border so they read as proper closed input fields. */
    div[data-baseweb="select"] > div {{
        border: 1.5px solid {BORDER} !important;
        border-radius: 10px !important;
        box-shadow: none !important;
    }}
    div[data-baseweb="select"] > div:focus-within {{
        border-color: {ACCENT} !important;
    }}
    div[data-testid="stDateInput"] > div {{
        border: 1.5px solid {BORDER} !important;
        border-radius: 10px !important;
    }}
    div[data-testid="stDateInput"] input {{
        border: none !important;
    }}

    /* Breathing room after the KPI summary row, before the next panel/card */
    .kpi-row-spacer {{
        height: 20px;
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


def _teal_scale(t):
    """Map t in [0, 1] to a hex color on a light-to-dark teal ramp (same
    family as BILLABLE_COLOR). Used for value-encoded bars where color
    represents magnitude, not category identity."""
    t = max(0.0, min(1.0, t))
    light = (0xD3, 0xEE, 0xEA)   # pale teal
    dark = (0x17, 0x5E, 0x54)    # deep teal
    rgb = tuple(int(light[i] + (dark[i] - light[i]) * t) for i in range(3))
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def donut_chart(billable, nonbill, notworked, title="Team Utilization Split"):
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
    # Canonical sort used across the whole app: Total Hours, descending.
    # Kept as ascending=True here on purpose — Plotly draws horizontal bars
    # bottom-to-top, so an ascending sort on the underlying data renders the
    # highest total at the TOP of the chart, i.e. visually descending.
    grp = df_period.groupby("QA Name")["Total Hours"].sum().reset_index()
    grp = grp.sort_values("Total Hours", ascending=True)

    # Single-hue value scale (teal, matching BILLABLE_COLOR's family) instead
    # of the categorical QA_PALETTE. Color here doesn't encode QA identity —
    # it encodes magnitude — so a continuous scale reads more honestly than a
    # rainbow of unrelated per-person colors. This also keeps the
    # teal/amber/red trio reserved strictly for Billable/Non-Billable/Not-Worked
    # in the stacked and donut charts, with no accidental reuse here.
    max_hours = grp["Total Hours"].max() if len(grp) else 0
    bar_colors = [
        _teal_scale(v / max_hours if max_hours else 0) for v in grp["Total Hours"]
    ]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=grp["QA Name"], x=grp["Total Hours"], orientation="h",
        marker_color=bar_colors,
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
    # Canonical sort: Total Hours, descending — same rule used by the bar
    # chart, the breakdown cards, and the summary table below.
    grp = df_period.groupby("QA Name")[
        ["Billable Hours", "Non-Billable Hours", "Hours Not Worked"]
    ].sum().reset_index()
    grp["Total Hours"] = grp["Billable Hours"] + grp["Non-Billable Hours"] + grp["Hours Not Worked"]
    grp = grp.sort_values("Total Hours", ascending=False)

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
        # Show each slice's own share as an on-chart label (matching the big
        # donut's textinfo="percent") so the split is readable at a glance,
        # not only on hover.
        textinfo="percent",
        textposition="inside",
        textfont=dict(size=10, color="white", family="Arial Black"),
        hovertemplate="%{label}: %{value:.1f} hrs (%{percent})<extra></extra>",
        sort=False,
    )])
    total = b + n + w
    util = (b / total * 100) if total else 0
    fig.update_layout(
        annotations=[dict(text=f"<b>{util:.0f}%</b><br><span style='font-size:8px;color:#7A7F91'>Utilization</span>",
                           x=0.5, y=0.5, showarrow=False,
                           font=dict(size=15, color=TEXT_MAIN))],
        showlegend=False, margin=dict(t=0, b=0, l=0, r=0), height=150,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig


# ============================================================================
# EXPORT HELPERS  (only ever called after the user clicks "Prepare Export")
# ============================================================================

def _safe_filename_fragment(text, max_len=60):
    """Strip characters that are unsafe in Windows/Mac filenames (plus
    parentheses, which are legal but look messy in a filename) and collapse
    whitespace/commas into underscores, so export filenames never break when
    downloaded on either OS. Truncates long fragments so the full filename
    stays reasonable."""
    text = re.sub(r'[\\/:*?"<>|()]', "", text)
    text = re.sub(r"[\s,]+", "_", text.strip())
    text = re.sub(r"_+", "_", text)
    text = text.strip("_")
    return text[:max_len] if len(text) > max_len else text


def _build_export_filename_base(period_label, sel_qas, all_qas):
    """Build the shared filename base (without extension) for an export,
    reflecting both the time period AND the QA selection at the moment
    "Prepare Exports" was clicked -- so a file filtered to 2 of 8 QAs is
    identifiable from its filename alone, not just its contents."""
    period_part = _safe_filename_fragment(period_label)

    if not sel_qas or set(sel_qas) == set(all_qas):
        # No QA-specific suffix needed: either everyone is included (the
        # default, most common case) or nothing is (period_label alone,
        # e.g. "No months selected", already says enough).
        qa_part = ""
    elif len(sel_qas) <= 3:
        # A handful of named QAs: spell them out, it's still a short, readable filename.
        qa_part = "_" + "_".join(_safe_filename_fragment(q, max_len=20) for q in sorted(sel_qas))
    else:
        # Many QAs but not all: naming them all would make the filename
        # unwieldy, so summarize by count instead.
        qa_part = f"_{len(sel_qas)}QAs"

    base = f"QA_Dashboard_{period_part}{qa_part}"
    return base if base else "QA_Dashboard_Export"


KALEIDO_PROBE_TIMEOUT_SECONDS = 8   # one-time-per-session viability check (see _check_kaleido_available)
KALEIDO_TIMEOUT_SECONDS = 10        # per real figure render; called up to ~22 times in one export cycle,
                                     # so this is a multiplicative cost, not a fixed one -- kept tight but
                                     # still generous given the render resolutions were also cut ~5x below.


def _run_with_timeout(fn, timeout_seconds, *args, **kwargs):
    """Run fn(*args, **kwargs) in a worker thread and enforce a hard wall-clock
    timeout. Needed because a broken kaleido install (most commonly: it can't
    find/launch a Chrome/Chromium runtime) doesn't always raise an exception —
    it can hang the underlying subprocess call indefinitely instead. A plain
    try/except never catches a hang, so this is the only reliable guard.
    Uses a thread (not signal.alarm) so it also works on Windows, since this
    app is meant to be packaged into a Windows .exe / Mac .dmg later.

    IMPORTANT: uses shutdown(wait=False) — if we let the ThreadPoolExecutor's
    context manager do its default wait=True shutdown, it blocks until the
    hung worker thread actually finishes, defeating the entire point of the
    timeout (the caller would still wait the full hang duration). The
    abandoned worker thread is left to die in the background (or run out
    the process's lifetime); Python threads don't get force-killed, but the
    caller gets control back at the deadline either way, which is what
    actually matters here."""
    from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError

    pool = ThreadPoolExecutor(max_workers=1)
    future = pool.submit(fn, *args, **kwargs)
    try:
        result = future.result(timeout=timeout_seconds)
        pool.shutdown(wait=False)
        return result
    except FutureTimeoutError:
        pool.shutdown(wait=False)
        raise TimeoutError(
            f"Chart image rendering did not finish within {timeout_seconds} seconds "
            "and was abandoned. This almost always means kaleido's headless "
            "Chrome/Chromium subprocess is hanging or failing to launch in this "
            "environment, not that your computer is slow."
        )
    except Exception:
        pool.shutdown(wait=False)
        raise


def _check_kaleido_available():
    """Kaleido (>=1.3.0, see requirements.txt) is required to turn Plotly
    figures into static images for the Excel/PDF exports. A successful
    `import kaleido` does NOT guarantee fig.to_image() actually works —
    kaleido v1 needs a real Chrome/Chromium installation on the machine (it
    no longer bundles one, unlike the old 0.2.1 that's no longer used here),
    and if that's missing it raises a clear "Chrome not installed"
    RuntimeError rather than hanging. This still runs a real, tiny,
    TIME-LIMITED end-to-end render so a broken install is always caught
    quickly instead of silently producing blank chart pages.

    Returns (ok: bool, message: str | None, chrome_missing: bool). When
    chrome_missing is True, the caller can offer a one-click "Install
    Chrome" button (see the Export section) rather than this function
    silently starting a ~100MB+ download on its own — a health check
    shouldn't have a surprise multi-second network side effect."""
    try:
        import kaleido  # noqa: F401
    except ImportError:
        return False, (
            "The 'kaleido' package is not installed, so charts cannot be "
            "rendered into the Excel/PDF exports. Install it with:\n\n"
            "    pip install kaleido==1.3.0\n\n"
            "then restart the app and click Prepare Exports again."
        ), False

    def _probe():
        probe_fig = go.Figure(data=[go.Bar(x=[1], y=[1])])
        png_bytes = probe_fig.to_image(format="png", width=100, height=100, scale=1)
        if not png_bytes or len(png_bytes) < 200 or not png_bytes.startswith(b"\x89PNG"):
            raise RuntimeError("kaleido produced invalid/empty image data")
        return png_bytes

    try:
        _run_with_timeout(_probe, KALEIDO_PROBE_TIMEOUT_SECONDS)
        return True, None, False
    except Exception as e:
        chrome_missing = "chrome" in str(e).lower()
        if chrome_missing:
            return False, (
                "Kaleido needs a real Chrome installation on this machine to "
                "render charts \u2014 it no longer bundles one (unlike the older "
                "kaleido 0.2.1). Click **Install Chrome for Exports** below to "
                "fetch it automatically (~100MB, one-time, takes a minute), or "
                "install Chrome yourself from google.com/chrome. Either way, "
                "restart the Streamlit app afterwards and click Prepare "
                "Exports again."
            ), True
        return False, (
            f"Kaleido failed a quick test export ({e}). Chrome appears to be "
            "installed, so this is likely a version mismatch \u2014 make sure "
            "`kaleido` and `plotly` are both current:\n\n"
            "    pip install --upgrade kaleido plotly\n\n"
            "(a kaleido older than 1.3.0 paired with a newer plotly is a "
            "known source of export errors), then fully restart the "
            "Streamlit app (not just refresh the browser) and click Prepare "
            "Exports again."
        ), False



def _fig_to_png_bytes(fig, width=900, height=500, scale=2):
    """Render a Plotly figure to PNG bytes, with a hard timeout. Raises on
    failure, on a timeout, OR on a suspiciously tiny/invalid result (instead
    of silently returning None, hanging forever, or trusting corrupt bytes)
    so export problems are always visible, never baked into a broken-looking
    PDF/Excel file with empty chart slots or an unresponsive UI."""
    def _render():
        return fig.to_image(format="png", width=width, height=height, scale=scale)

    png_bytes = _run_with_timeout(_render, KALEIDO_TIMEOUT_SECONDS)
    if not png_bytes or len(png_bytes) < 500 or not png_bytes.startswith(b"\x89PNG"):
        raise RuntimeError(
            "Chart image rendering returned invalid/empty PNG data — "
            "kaleido may be installed but not working correctly in this "
            "environment (common cause: missing Chrome/Chromium runtime "
            "that kaleido depends on)."
        )
    return png_bytes


def to_excel_bytes(summary_df, detail_df, kpis, period_label, chart_titles, chart_pngs, qa_names, mini_pngs, include_images=True):
    """chart_pngs and mini_pngs are dicts of already-rendered PNG bytes,
    keyed by chart title / QA name respectively -- rendered once, shared with
    to_pdf_bytes, rather than re-rendered here from scratch. Pass an empty
    dict (not None) when include_images is False."""
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
    note_font = Font(size=9, italic=True, color="C1543D")

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
    images_failed = False
    if include_images:
        for title in chart_titles:
            ws[f"B{row_cursor}"] = title
            ws[f"B{row_cursor}"].font = Font(size=12, bold=True, color="0F1729")
            row_cursor += 1
            png = chart_pngs.get(title)
            if png is not None:
                img = XLImage(io.BytesIO(png))
                img.width, img.height = 560, 311
                ws.add_image(img, f"B{row_cursor}")
                row_cursor += 20
            else:
                images_failed = True
                ws[f"B{row_cursor}"] = "(chart image unavailable — see chart data in the tables below)"
                ws[f"B{row_cursor}"].font = note_font
                row_cursor += 2

        ws[f"B{row_cursor}"] = "Individual QA Breakdown"
        ws[f"B{row_cursor}"].font = Font(size=13, bold=True, color="0F1729")
        row_cursor += 2
        col_positions = ["B", "F", "J", "N"]
        start_row = row_cursor
        for idx, qa_name in enumerate(qa_names):
            col = col_positions[idx % 4]
            r = start_row + (idx // 4) * 14
            ws[f"{col}{r}"] = qa_name
            ws[f"{col}{r}"].font = Font(size=11, bold=True, color="0F1729")
            png = mini_pngs.get(qa_name)
            if png is not None:
                img = XLImage(io.BytesIO(png))
                img.width, img.height = 220, 220
                ws.add_image(img, f"{col}{r + 1}")
            else:
                images_failed = True
    else:
        ws[f"B{row_cursor}"] = "Chart images were skipped for this export (kaleido unavailable). All figures are in the tables below."
        ws[f"B{row_cursor}"].font = note_font
        row_cursor += 2

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
    return buf, images_failed


def to_pdf_bytes(summary_df, detail_df, kpis, period_label, chart_titles, chart_pngs, qa_names, mini_pngs, include_images=True):
    """chart_pngs and mini_pngs are the SAME already-rendered PNG dicts passed
    to to_excel_bytes -- rendering happens exactly once per figure, shared
    between both export formats, rather than once per format."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                     Spacer, Image as RLImage, PageBreak, KeepTogether)
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

    images_failed = False
    chart_imgs = []
    if include_images:
        for title in chart_titles:
            png = chart_pngs.get(title)
            if png is not None:
                chart_imgs.append((title, png))
            else:
                images_failed = True

    elements.append(Paragraph("Team Overview", h2_style))

    if chart_imgs:
        # Normal path: charts rendered fine, show them as images.
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
    else:
        # Fallback path: no images available for any reason (kaleido disabled,
        # unavailable, or every render failed). Never leave a near-empty page —
        # show the same information as real numbers instead of an image.
        elements.append(Paragraph(
            "Chart images could not be rendered in this environment (kaleido/Chrome "
            "unavailable). The figures below are the same data the on-screen charts "
            "show, presented as tables instead.",
            sub_style))
        elements.append(Spacer(1, 8))

        elements.append(Paragraph("QA Comparison &mdash; Total Hours", ParagraphStyle(
            "h3a", fontSize=11, textColor=rl_colors.HexColor("#0F1729"), spaceBefore=4, spaceAfter=4)))
        comp_df = summary_df[["QA Name", "Total Hours"]].sort_values("Total Hours", ascending=False)
        comp_data = [["QA Name", "Total Hours"]] + comp_df.round(1).astype(str).values.tolist()
        comp_tbl = Table(comp_data, colWidths=[250, 150])
        comp_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0F1729")),
            ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#E7E9F0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F4F5F8")]),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(comp_tbl)
        elements.append(Spacer(1, 14))

        elements.append(Paragraph("Hours Mix &mdash; Billable / Non-Billable / Not Worked", ParagraphStyle(
            "h3b", fontSize=11, textColor=rl_colors.HexColor("#0F1729"), spaceBefore=4, spaceAfter=4)))
        mix_df = summary_df[["QA Name", "Billable Hours", "Non-Billable Hours", "Hours Not Worked"]]
        mix_data = [["QA Name", "Billable", "Non-Billable", "Not Worked"]] + mix_df.round(1).astype(str).values.tolist()
        mix_tbl = Table(mix_data, colWidths=[220, 120, 120, 120])
        mix_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0F1729")),
            ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#E7E9F0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F4F5F8")]),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(mix_tbl)
        elements.append(Spacer(1, 14))

        elements.append(Paragraph("Team Utilization Split", ParagraphStyle(
            "h3c", fontSize=11, textColor=rl_colors.HexColor("#0F1729"), spaceBefore=4, spaceAfter=4)))
        elements.append(Paragraph(
            f"Billable: {kpis['billable']:.1f} hrs &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Non-Billable: {kpis['nonbill']:.1f} hrs &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Utilization: {kpis['utilization']:.1f}%",
            sub_style))

    # Only force a fresh page here when there are real chart images that need
    # the room -- confirmed by direct visual inspection that forcing a page
    # break in the image-less fallback path left 60-85% of the page blank,
    # since the fallback content (a couple of small tables) is far shorter
    # than an embedded chart image. Letting ReportLab's natural flow pack
    # sections together in the fallback case avoids that wasted space.
    if include_images:
        elements.append(PageBreak())
    elements.append(Paragraph("Individual QA Breakdown", h2_style))

    mini_row = []
    any_mini_image = False
    for i, qa_name in enumerate(qa_names):
        img = None
        if include_images:
            png = mini_pngs.get(qa_name)
            if png is not None:
                img = RLImage(io.BytesIO(png), width=110, height=110)
                any_mini_image = True
            else:
                images_failed = True
        if img is None:
            img = Paragraph("", styles["Normal"])
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

    if not any_mini_image:
        # No per-QA donut images either — replace the mostly-empty name grid
        # with the real per-QA utilization numbers instead.
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(
            "Donut images could not be rendered in this environment. Per-QA "
            "utilization (the number the donuts show) is below:",
            sub_style))
        elements.append(Spacer(1, 6))
        util_df = summary_df[["QA Name", "Billable Hours", "Non-Billable Hours",
                               "Hours Not Worked", "Total Hours", "Utilization %"]]
        util_data = [["QA Name", "Billable", "Non-Billable", "Not Worked", "Total", "Utilization %"]] + \
                    util_df.round(1).astype(str).values.tolist()
        util_tbl = Table(util_data, colWidths=[170, 100, 110, 100, 90, 110])
        util_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0F1729")),
            ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#E7E9F0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F4F5F8")]),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(util_tbl)

    # Same reasoning as above: only force a page break here if the section
    # just rendered actually used real images (any_mini_image is more precise
    # than the general include_images flag, since individual mini-donut
    # renders can fail even when include_images is True overall).
    if any_mini_image:
        elements.append(PageBreak())

    per_qa_table_data = [list(summary_df.columns)] + summary_df.round(1).astype(str).values.tolist()
    n_cols = len(summary_df.columns)
    col_width = 780 / n_cols
    per_qa_tbl = Table(per_qa_table_data, colWidths=[col_width] * n_cols, repeatRows=1)
    per_qa_tbl.setStyle(TableStyle([
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
    elements.append(KeepTogether([
        Paragraph("Per-QA Summary Table", h2_style),
        Spacer(1, 4),
        per_qa_tbl,
    ]))

    # ---- Daily Log (mirrors the on-screen Daily Log table) ----
    # Same reasoning as the two page breaks above, corrected after actually
    # rendering and inspecting the output: in the no-images fallback case the
    # Per-QA Summary Table is short enough to leave real room on its page, so
    # forcing a break here left that page mostly blank too. Conditional on
    # include_images lets Daily Log flow naturally in the fallback case.
    if include_images:
        elements.append(PageBreak())
    elements.append(Paragraph("Daily Log", h2_style))
    elements.append(Paragraph(
        f"{len(detail_df):,} rows &middot; sorted by date, most recent first",
        sub_style))
    elements.append(Spacer(1, 4))

    log_for_pdf = detail_df.copy().sort_values("Date", ascending=False)
    log_for_pdf["Date"] = pd.to_datetime(log_for_pdf["Date"]).dt.strftime("%Y-%m-%d")
    log_cols = ["Date", "Day", "QA Name", "Billable Hours", "Non-Billable Hours",
                "Hours Not Worked", "Total Hours", "Comment"]
    log_for_pdf = log_for_pdf[log_cols]

    # Cap rows rendered directly in the PDF table for file-size/perf sanity;
    # the full, uncapped data is always in the companion Excel export.
    MAX_PDF_LOG_ROWS = 500
    truncated = len(log_for_pdf) > MAX_PDF_LOG_ROWS
    log_for_pdf_show = log_for_pdf.head(MAX_PDF_LOG_ROWS)

    log_table_data = [log_cols] + log_for_pdf_show.round(2).astype(str).values.tolist()
    log_col_widths = [70, 40, 90, 65, 75, 70, 65, 300]
    log_tbl = Table(log_table_data, colWidths=log_col_widths, repeatRows=1)
    log_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0F1729")),
        ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#E7E9F0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F4F5F8")]),
        ("ALIGN", (0, 0), (6, -1), "CENTER"),
        ("ALIGN", (7, 0), (7, -1), "LEFT"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    elements.append(log_tbl)
    if truncated:
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(
            f"Showing the most recent {MAX_PDF_LOG_ROWS:,} of {len(log_for_pdf):,} rows. "
            "Download the Excel export for the complete daily log.",
            sub_style))

    doc.build(elements)
    buf.seek(0)
    return buf, images_failed


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

# "Clear All" (button lives at the very bottom of the sidebar, see below)
# has to take effect here, BEFORE any filter widget below is instantiated —
# Streamlit widgets read their bound session_state key at creation time, so
# clearing state after the widgets already rendered this run would only be
# visible on the NEXT interaction, not immediately. The button itself sets
# this flag and calls st.rerun(); this block is what actually acts on it.
if st.session_state.get("_clear_all_filters"):
    st.session_state["_clear_all_filters"] = False
    st.session_state["qa_multiselect"] = []
    # Empty each, matching exactly what that filter's own individual Clear
    # button does — Clear All is just "press every Clear button at once".
    for _k in ("years_multiselect", "months_multiselect", "weeks_multiselect"):
        st.session_state[_k] = []

view_mode = st.sidebar.radio("View by", ["Daily", "Weekly", "Monthly", "Yearly"], index=2)

# ---- QA selection: a real closed dropdown (multiselect), with quick All/None ----
all_qas = sorted(data["QA Name"].unique().tolist())

if "qa_multiselect" not in st.session_state:
    st.session_state.qa_multiselect = all_qas.copy()
# keep state in sync if a new file introduces different QAs
st.session_state.qa_multiselect = [qa for qa in st.session_state.qa_multiselect if qa in all_qas]

if st.sidebar.button("Clear QAs", use_container_width=True):
    st.session_state.qa_multiselect = []

sel_qas = st.sidebar.multiselect("QA Team Members", all_qas, key="qa_multiselect")

st.sidebar.markdown("---")

# ---- Time period selection depending on view mode ----
years_available = sorted(data["Year"].dropna().unique().astype(int).tolist())

def _sidebar_multiselect_with_clear(label, options, state_key, clear_label, format_func=None):
    """A multiselect that defaults to "all selected" the first time it's seen,
    persists its selection in session_state under state_key (same pattern as
    the QA Team Members widget above), and renders its own small "Clear <x>"
    button directly above it. Matches the existing, working Clear QAs
    pattern exactly: the button mutates st.session_state[state_key] BEFORE
    the multiselect widget below it is instantiated in this same script run,
    so the widget picks up the cleared value immediately with no st.rerun()
    needed — unlike passing default= each run, which would snap back to
    "all selected" since nothing would ever persist a cleared state.
    """
    if state_key not in st.session_state:
        st.session_state[state_key] = list(options)
    # keep state in sync if the available options changed (e.g. new file, or
    # Years selection narrowing which Months/Weeks exist)
    st.session_state[state_key] = [v for v in st.session_state[state_key] if v in options]

    if st.sidebar.button(clear_label, use_container_width=True, key=f"{state_key}_clear_btn"):
        st.session_state[state_key] = []

    kwargs = {"key": state_key}
    if format_func is not None:
        kwargs["format_func"] = format_func
    return st.sidebar.multiselect(label, options, **kwargs)


if view_mode == "Yearly":
    sel_years = _sidebar_multiselect_with_clear("Years", years_available, "years_multiselect", "Clear Years")
    df_period = data[data["Year"].isin(sel_years)]
    period_label = ", ".join(str(y) for y in sel_years) if sel_years else "No years selected"

elif view_mode == "Monthly":
    sel_years = _sidebar_multiselect_with_clear("Years", years_available, "years_multiselect", "Clear Years")
    months_in_scope = [m for m in MONTH_ORDER if m in data[data["Year"].isin(sel_years)]["Month"].unique()]
    sel_months = _sidebar_multiselect_with_clear("Months", months_in_scope, "months_multiselect", "Clear Months")
    df_period = data[data["Year"].isin(sel_years) & data["Month"].isin(sel_months)]
    # Dynamic label: "Feb 2026" for a single month/year, "Feb, Mar (2026)"
    # for several months in one year, "Feb 2025, Mar 2026" if years differ —
    # never a fixed placeholder string.
    if not sel_months or not sel_years:
        period_label = "No months selected"
    elif len(sel_years) == 1:
        year_txt = str(sel_years[0])
        if len(sel_months) == 1:
            period_label = f"{sel_months[0]} {year_txt}"
        else:
            period_label = f"{', '.join(sel_months)} ({year_txt})"
    else:
        period_label = f"{', '.join(sel_months)} ({', '.join(str(y) for y in sel_years)})"

elif view_mode == "Weekly":
    sel_years = _sidebar_multiselect_with_clear("Years", years_available, "years_multiselect", "Clear Years")
    df_y = data[data["Year"].isin(sel_years)]
    df_y = df_y.assign(_week=df_y["Date"].dt.to_period("W"))
    week_options = sorted(df_y["_week"].unique())
    week_labels = {w: f"{w.start_time.strftime('%d %b')} – {w.end_time.strftime('%d %b %Y')}" for w in week_options}
    sel_weeks = _sidebar_multiselect_with_clear(
        "Weeks", week_options, "weeks_multiselect", "Clear Weeks",
        format_func=lambda w: week_labels.get(w, str(w)),
    )
    df_period = df_y[df_y["_week"].isin(sel_weeks)].drop(columns="_week")
    if not sel_weeks:
        period_label = "No weeks selected"
    elif len(sel_weeks) == 1:
        period_label = week_labels.get(sel_weeks[0], str(sel_weeks[0]))
    else:
        period_label = f"{len(sel_weeks)} weeks selected"

else:  # Daily -> calendar-style multi-date picker
    min_d, max_d = data["Date"].min().date(), min(data["Date"].max().date(), TODAY.date())
    sel_dates = st.sidebar.date_input(
        "Select date(s)", value=(min_d, max_d), min_value=min_d, max_value=max_d,
    )
    if isinstance(sel_dates, (tuple, list)) and len(sel_dates) == 2:
        d_start, d_end = sel_dates
        df_period = data[(data["Date"].dt.date >= d_start) & (data["Date"].dt.date <= d_end)]
        if d_start == d_end:
            period_label = d_start.strftime("%d %b %Y")
        else:
            period_label = f"{d_start.strftime('%d %b %Y')} – {d_end.strftime('%d %b %Y')}"
    elif isinstance(sel_dates, (tuple, list)) and len(sel_dates) == 1:
        d_only = sel_dates[0]
        df_period = data[data["Date"].dt.date == d_only]
        period_label = d_only.strftime("%d %b %Y")
    else:
        df_period = data[data["Date"].dt.date == sel_dates]
        period_label = sel_dates.strftime("%d %b %Y")

st.sidebar.markdown("---")
if st.sidebar.button("\U0001F9F9 Clear All Filters", use_container_width=True):
    st.session_state["_clear_all_filters"] = True
    st.rerun()

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

st.markdown(f'<div class="panel-title" style="margin-bottom:12px;">QA &mdash; {period_label}</div>', unsafe_allow_html=True)

k1, k2, k3, k4, k5, k6 = st.columns(6)
kpi_cells = [
    (k1, "QA TEAM SIZE", f"{team_size}", "active members"),
    (k2, "TOTAL HOURS", f"{total_hours:,.1f}", f"across {days_logged} days"),
    (k3, "BILLABLE SHARE", f"{utilization:.1f}%", f"{billable_total:,.0f} of {total_hours:,.0f} hrs"),
    (k4, "AVG HOURS / DAY", f"{avg_hours_per_day:.2f}", "per active QA-day"),
    (k5, "NON-BILLABLE HOURS", f"{nonbill_total:,.1f}", "hrs logged"),
    (k6, "HOURS NOT WORKED", f"{notworked_total:,.1f}", "hrs logged"),
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

st.markdown('<div class="kpi-row-spacer"></div>', unsafe_allow_html=True)

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
    ("Team Utilization Split (Donut)", fig_donut),
]

# ============================================================================
# PER-QA CARDS
# ============================================================================
st.markdown(
    '<div class="panel-title">Individual QA Breakdown</div>'
    '<div class="panel-sub" style="margin-top:2px;">'
    'The number in the center of each donut is that QA\u2019s <b>Billable Hours \u00f7 Total Hours</b> '
    '(billable share of their own hours) for the period selected above \u2014 not a share of the whole team.'
    '</div>',
    unsafe_allow_html=True,
)
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
# Canonical sort applied everywhere in the app: Total Hours, descending.
# qa_summary feeds the breakdown cards below, the "View Per-QA Summary
# Table" expander, and both the Excel and PDF exports — sorting it here
# once keeps all of those consistent with the bar chart and stacked chart.
qa_summary = qa_summary.sort_values("Total Hours", ascending=False)

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



st.markdown('<div class="panel-title" style="margin-top:1.2rem;">Daily Log</div>', unsafe_allow_html=True)
st.caption(
    "Every cleaned daily record for the QAs and period selected in the sidebar filters · sortable. "
    "All rows are shown in this one scrollable table (no separate pages) — scroll inside the "
    "table itself, vertically for more rows and horizontally for the Comment column, to see everything."
)

log_df = df_period.copy()

log_df_display = log_df[["Date", "Day", "QA Name", "Billable Hours", "Non-Billable Hours",
                          "Hours Not Worked", "Total Hours", "Comment"]].sort_values("Date", ascending=False)
log_df_display["Date"] = log_df_display["Date"].dt.strftime("%Y-%m-%d")

st.dataframe(
    log_df_display.round(2),
    use_container_width=True,
    hide_index=True,
    height=380,
    column_config={
        "Date": st.column_config.TextColumn("Date", width=95),
        "Day": st.column_config.TextColumn("Day", width=60),
        "QA Name": st.column_config.TextColumn("QA Name", width=150),
        "Billable Hours": st.column_config.NumberColumn("Billable Hours", width=100, format="%.2f"),
        "Non-Billable Hours": st.column_config.NumberColumn("Non-Billable Hours", width=110, format="%.2f"),
        "Hours Not Worked": st.column_config.NumberColumn("Hours Not Worked", width=110, format="%.2f"),
        "Total Hours": st.column_config.NumberColumn("Total Hours", width=100, format="%.2f"),
        # Comment's explicit width, combined with the explicit widths above,
        # sums to well past any realistic rendered table width even with
        # use_container_width=True -- that's what makes the grid's built-in
        # horizontal scroll actually engage (confirmed directly: scrolling
        # the grid horizontally does bring later columns and more Comment
        # text into view). The scrollbar itself has no visible indicator in
        # Streamlit's grid (a real, confirmed platform limitation, not a bug
        # in this app).
        "Comment": st.column_config.TextColumn(
            "Comment", width=500,
            help="Long comments are truncated here — scroll the table sideways, or drag this column's edge to widen it, to read one in full.",
        ),
    },
)
st.caption(f"{len(log_df_display):,} rows total \u00b7 all shown above, scroll to see the rest \u00b7 the table also scrolls sideways for the Comment column, though the scrollbar itself isn't always visible")

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
    st.session_state["_run_export_build"] = True

if st.session_state.get("_run_export_build"):
    st.session_state["_run_export_build"] = False

    if "_kaleido_check_cache" in st.session_state:
        # Already probed once this session -- reuse that result instead of
        # spending another up-to-KALEIDO_TIMEOUT_SECONDS on a question we
        # already have the answer to. This is invalidated (see the Chrome
        # install button below) whenever there's a real reason the answer
        # might have changed.
        kaleido_ok, kaleido_msg, chrome_missing = st.session_state["_kaleido_check_cache"]
    else:
        with st.spinner(f"Checking chart rendering (up to {KALEIDO_PROBE_TIMEOUT_SECONDS}s, once per session)..."):
            kaleido_ok, kaleido_msg, chrome_missing = _check_kaleido_available()
        st.session_state["_kaleido_check_cache"] = (kaleido_ok, kaleido_msg, chrome_missing)

    if not kaleido_ok:
        st.warning(
            f"{kaleido_msg}\n\n"
            "Proceeding to build the Excel and PDF **without chart images** — "
            "all KPI numbers, the per-QA summary table, and the full Daily Log "
            "will still be complete."
        )
        if chrome_missing:
            if st.button("\U0001F310 Install Chrome for Exports (one-time, ~100MB)"):
                try:
                    import plotly.io as pio
                    with st.spinner("Downloading Chrome for chart rendering \u2014 this can take a minute..."):
                        _run_with_timeout(pio.get_chrome, 120)
                    st.success("Chrome installed. Building your exports now...")
                    st.session_state.pop("_kaleido_check_cache", None)
                    st.session_state["_run_export_build"] = True
                    st.rerun()
                except Exception as install_err:
                    st.error(
                        f"Automatic Chrome install failed: {install_err}\n\n"
                        "Please install Chrome manually from google.com/chrome, "
                        "then restart the Streamlit app and click Prepare Exports again."
                    )

    chart_pngs = {}
    mini_pngs = {}
    if kaleido_ok:
        with st.spinner("Rendering charts (once, shared by both Excel and PDF)..."):
            # Render at the LARGER of the two destinations' needs (Excel's
            # 560x311 / 220x220 vs PDF's 555x320 / 165x165) -- the same PNG
            # then gets displayed slightly smaller in the PDF, which is safe
            # downscaling with no visible quality loss, and means each figure
            # is rendered exactly once instead of once per export format.
            for title, fig in team_chart_figs:
                try:
                    chart_pngs[title] = _fig_to_png_bytes(fig, width=560, height=311)
                except Exception:
                    pass  # left out of the dict; both builders already handle a missing entry
            for qa_name, fig in qa_mini_figs:
                try:
                    mini_pngs[qa_name] = _fig_to_png_bytes(fig, width=220, height=220)
                except Exception:
                    pass

    try:
        with st.spinner("Building Excel and PDF reports..."):
            chart_titles = [title for title, _ in team_chart_figs]
            qa_names_for_export = [qa_name for qa_name, _ in qa_mini_figs]
            excel_buf, excel_images_failed = to_excel_bytes(
                export_summary, export_detail, kpis, period_label,
                chart_titles, chart_pngs, qa_names_for_export, mini_pngs,
                include_images=kaleido_ok)
            pdf_buf, pdf_images_failed = to_pdf_bytes(
                export_summary, export_detail, kpis, period_label,
                chart_titles, chart_pngs, qa_names_for_export, mini_pngs,
                include_images=kaleido_ok)
            st.session_state["export_excel_bytes"] = excel_buf
            st.session_state["export_pdf_bytes"] = pdf_buf
            st.session_state["export_period_label"] = period_label
            st.session_state["export_filename_base"] = _build_export_filename_base(period_label, sel_qas, all_qas)
            st.session_state["export_generated_at"] = datetime.now().strftime("%d %b %Y, %H:%M:%S")

        if kaleido_ok and not excel_images_failed and not pdf_images_failed:
            st.success("Exports ready below, with all charts included \u2014 changing filters now will NOT regenerate them until you click Prepare Exports again.")
        else:
            st.info("Exports ready below \u2014 data-complete, but some chart images could not be rendered (see warning above). Changing filters now will NOT regenerate them until you click Prepare Exports again.")
    except Exception as e:
        st.error(f"Export generation failed: {e}")
        st.session_state.pop("export_excel_bytes", None)
        st.session_state.pop("export_pdf_bytes", None)

if "export_excel_bytes" in st.session_state:
    gen_at = st.session_state.get("export_generated_at", "unknown time")
    st.caption(f"\u2705 These exports were generated at **{gen_at}**. If that's not recent, click Prepare Exports again after fixing any errors above.")
    e1, e2 = st.columns(2)
    fname_base = st.session_state.get(
        "export_filename_base",
        _build_export_filename_base(st.session_state.get("export_period_label", period_label), sel_qas, all_qas),
    )
    with e1:
        st.download_button(
            "\U0001F4E5 Download Excel",
            data=st.session_state["export_excel_bytes"],
            file_name=f"{fname_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with e2:
        st.download_button(
            "\U0001F4C4 Download PDF",
            data=st.session_state["export_pdf_bytes"],
            file_name=f"{fname_base}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

st.caption("Built for QA Management \u00b7 Streamlit Dashboard \u00b7 Ready for desktop packaging (Windows/Mac)")