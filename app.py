"""
Metals Risk Premia — Interactive Dashboard
============================================
Streamlit dashboard for exploring LME & CME metals data.

Data files required (place in same directory):
  1. Metals Cash and 3M.xlsx
  2. Metals Futures Curve.csv  (actually .xlsx based on screenshots)

Deploy: Push to GitHub → connect to streamlit.io/cloud
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from scipy import stats
from datetime import timedelta
import warnings
warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════

st.set_page_config(
    page_title="Metals Risk Premia Dashboard",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ═══════════════════════════════════════════════
# STYLING
# ═══════════════════════════════════════════════

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');

    /* Global */
    .stApp { font-family: 'DM Sans', sans-serif; }
    
    /* Hide default streamlit elements */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    
    /* Metric cards */
    .metric-card {
        background: linear-gradient(135deg, #1A1F2E 0%, #0F1724 100%);
        border: 1px solid #2D3748;
        border-radius: 12px;
        padding: 16px 20px;
        margin: 4px 0;
    }
    .metric-card h4 {
        color: #94A3B8;
        font-size: 0.75rem;
        font-weight: 500;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        margin: 0 0 4px 0;
    }
    .metric-card .value {
        color: #E2E8F0;
        font-family: 'JetBrains Mono', monospace;
        font-size: 1.5rem;
        font-weight: 700;
        margin: 0;
    }
    .metric-card .delta-pos { color: #34D399; font-size: 0.85rem; }
    .metric-card .delta-neg { color: #F87171; font-size: 0.85rem; }
    
    /* Section headers */
    .section-header {
        font-family: 'DM Sans', sans-serif;
        color: #E2E8F0;
        font-size: 1.1rem;
        font-weight: 700;
        border-bottom: 2px solid #3B82F6;
        padding-bottom: 8px;
        margin: 24px 0 16px 0;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background-color: #0F1724;
        padding: 4px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 20px;
        font-weight: 500;
    }
    
    /* Backwardation / Contango badges */
    .badge-backwardation {
        background: rgba(52, 211, 153, 0.15);
        color: #34D399;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    .badge-contango {
        background: rgba(248, 113, 113, 0.15);
        color: #F87171;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    
    /* Title */
    .main-title {
        font-family: 'DM Sans', sans-serif;
        font-weight: 700;
        font-size: 1.8rem;
        color: #E2E8F0;
        margin-bottom: 0;
    }
    .main-subtitle {
        color: #64748B;
        font-size: 0.9rem;
        margin-top: 0;
    }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# CHART THEME
# ═══════════════════════════════════════════════

CHART_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(14,17,23,0.8)",
    font=dict(family="DM Sans, sans-serif", color="#94A3B8"),
    xaxis=dict(gridcolor="rgba(45,55,72,0.5)", zerolinecolor="rgba(45,55,72,0.5)"),
    yaxis=dict(gridcolor="rgba(45,55,72,0.5)", zerolinecolor="rgba(45,55,72,0.5)"),
    margin=dict(l=60, r=30, t=50, b=50),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
    hoverlabel=dict(bgcolor="#1A1F2E", font_size=12, font_family="JetBrains Mono"),
)

COLORS = {
    "primary": "#3B82F6",
    "secondary": "#8B5CF6",
    "accent": "#06B6D4",
    "green": "#34D399",
    "red": "#F87171",
    "amber": "#FBBF24",
    "orange": "#FB923C",
    "pink": "#F472B6",
    "slate": "#64748B",
}

METAL_COLORS = {
    "Copper": "#FB923C",
    "Aluminium": "#94A3B8",
    "Zinc": "#3B82F6",
    "Nickel": "#8B5CF6",
    "Lead": "#64748B",
    "Tin": "#06B6D4",
    "Gold": "#FBBF24",
    "Silver": "#CBD5E1",
    "Platinum": "#A78BFA",
    "Palladium": "#F472B6",
}


# ═══════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════

@st.cache_data(ttl=3600)
def load_cash_3m_data(file):
    """Load Metals Cash and 3M.xlsx — one sheet per metal."""
    xls = pd.ExcelFile(file)
    data = {}

    # LME metals sheets
    lme_sheets = [s for s in xls.sheet_names if "LME" in s]
    for sheet in lme_sheets:
        df = pd.read_excel(xls, sheet_name=sheet, header=[0, 1])
        # Flatten multi-level columns
        df.columns = [
            f"{c[0].strip()}_{c[1].strip()}" if "Unnamed" not in str(c[0]) else c[1].strip()
            for c in df.columns
        ]
        # Find date column
        date_col = [c for c in df.columns if "date" in c.lower() or "Date" in c]
        if date_col:
            df = df.rename(columns={date_col[0]: "Date"})
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"])
            df = df.set_index("Date").sort_index()
        metal_name = sheet.replace("LME ", "").strip()
        data[metal_name] = df

    # CME Cash Prices sheet
    cme_sheets = [s for s in xls.sheet_names if "CME" in s or "Cash" in s]
    for sheet in cme_sheets:
        df = pd.read_excel(xls, sheet_name=sheet)
        date_col = [c for c in df.columns if "date" in c.lower() or "Date" in c]
        if date_col:
            df = df.rename(columns={date_col[0]: "Date"})
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"])
            df = df.set_index("Date").sort_index()
        data["CME_Cash"] = df

    return data


@st.cache_data(ttl=3600)
def load_futures_curve_data(file):
    """
    Load Metals Futures Curve — one sheet per metal with F1-F27.
    Handles: .xlsx, .xls, .csv (with encoding fallbacks), and
    xlsx files incorrectly saved with .csv extension.
    """
    fname = file.name if hasattr(file, 'name') else str(file)
    data = {}

    # ── Try reading as Excel first (even if extension is .csv) ──
    try:
        xls = pd.ExcelFile(file)
        return _parse_curve_excel(xls)
    except Exception:
        pass

    # ── If Excel fails, try CSV with multiple encodings ──
    if fname.lower().endswith('.csv'):
        for encoding in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16']:
            try:
                file.seek(0)  # Reset file pointer
                df = pd.read_csv(file, encoding=encoding)
                if not df.empty:
                    data["Sheet1"] = _parse_single_curve_df(df)
                    return data
            except Exception:
                continue

        # Last resort: read as bytes, decode, then parse
        try:
            file.seek(0)
            raw = file.read()
            # Try to detect if it's actually xlsx bytes
            if raw[:4] == b'PK\x03\x04':  # ZIP/XLSX magic bytes
                import io
                file.seek(0)
                xls = pd.ExcelFile(io.BytesIO(raw))
                return _parse_curve_excel(xls)
            # Otherwise try as text
            for enc in ['utf-8', 'latin-1', 'cp1252']:
                try:
                    text = raw.decode(enc)
                    import io
                    df = pd.read_csv(io.StringIO(text))
                    if not df.empty:
                        data["Sheet1"] = _parse_single_curve_df(df)
                        return data
                except Exception:
                    continue
        except Exception:
            pass

    st.error(f"Could not read '{fname}'. Try saving it as .xlsx from Excel and re-uploading.")
    return data


def _parse_curve_excel(xls):
    """Parse an Excel file with one sheet per metal, multi-row headers."""
    data = {}

    for sheet in xls.sheet_names:
        try:
            # First pass: read raw to detect header structure
            df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=5)

            # Detect header rows by looking for "Date", "F1", "Price" etc.
            header_rows = []
            for i in range(min(4, len(df_raw))):
                row_vals = [str(v).strip().lower() for v in df_raw.iloc[i].values if pd.notna(v)]
                if any(kw in " ".join(row_vals) for kw in ["date", "f1", "f2", "price", "volume"]):
                    header_rows.append(i)

            # Read with detected headers
            if len(header_rows) >= 2:
                df = pd.read_excel(xls, sheet_name=sheet, header=header_rows)
            elif len(header_rows) == 1:
                df = pd.read_excel(xls, sheet_name=sheet, header=header_rows[0])
            else:
                # Default: try first 3 rows as header (matches your screenshot: Row1=blank, Row2=F1/F2, Row3=Price/Vol/OI)
                df = pd.read_excel(xls, sheet_name=sheet, header=[0, 1, 2])

        except Exception:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=[0, 1])
            except Exception:
                df = pd.read_excel(xls, sheet_name=sheet)

        data[sheet] = _parse_single_curve_df(df)

    return data


def _parse_single_curve_df(df):
    """Parse a single dataframe with futures curve data into standardized format."""
    # Flatten multi-level columns
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for col_tuple in df.columns:
            parts = [str(p).strip() for p in col_tuple
                     if pd.notna(p) and "Unnamed" not in str(p) and str(p).strip()]
            new_cols.append("_".join(parts) if parts else str(col_tuple))
        df.columns = new_cols

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    # Find and set date index
    date_col = [c for c in df.columns if "date" in c.lower()]
    if date_col:
        df = df.rename(columns={date_col[0]: "Date"})
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
        df = df.set_index("Date").sort_index()

    # Extract price columns for each future month
    # Match patterns: "F1_Price", "F1 Price", "F1_price", or just columns containing "F1" and "price"
    prices = {}
    for col in df.columns:
        col_lower = col.lower().replace(" ", "_")
        for i in range(1, 28):
            patterns = [
                f"f{i}_price", f"f{i}_Price",
                f"F{i}_Price", f"F{i}_price",
            ]
            # Also match "F1" at start with "price" somewhere
            if any(p.lower() in col_lower for p in patterns):
                prices[f"F{i}"] = pd.to_numeric(df[col], errors="coerce")
                break
            # Fallback: column starts with F{i} and contains "price"
            elif col_lower.startswith(f"f{i}") and "price" in col_lower:
                prices[f"F{i}"] = pd.to_numeric(df[col], errors="coerce")
                break

    result = {
        "raw": df,
        "prices": pd.DataFrame(prices, index=df.index) if prices else pd.DataFrame()
    }
    return result


def parse_cash_3m_columns(df, metal_name):
    """Parse the Cash & 3M dataframe columns into standardized names."""
    result = pd.DataFrame(index=df.index)

    for col in df.columns:
        cl = col.lower()
        # Cash price
        if ("cash" in cl and "price" in cl) or ("spot" in cl and "price" in cl):
            result["cash_price"] = pd.to_numeric(df[col], errors="coerce")
        # 3M price
        elif ("3m" in cl and "price" in cl) or ("forward" in cl and "price" in cl):
            result["3m_price"] = pd.to_numeric(df[col], errors="coerce")
        # 3M volume
        elif "3m" in cl and "volume" in cl:
            result["3m_volume"] = pd.to_numeric(df[col], errors="coerce")
        # 3M open interest
        elif "3m" in cl and ("open" in cl or "oi" in cl or "int" in cl):
            result["3m_oi"] = pd.to_numeric(df[col], errors="coerce")
        # Spread price
        elif "spread" in cl and "price" in cl:
            result["spread_price"] = pd.to_numeric(df[col], errors="coerce")
        # Spread volume
        elif "spread" in cl and "volume" in cl:
            result["spread_volume"] = pd.to_numeric(df[col], errors="coerce")

    # Compute spread if not present but cash & 3m are
    if "spread_price" not in result.columns and "cash_price" in result.columns and "3m_price" in result.columns:
        result["spread_price"] = result["cash_price"] - result["3m_price"]

    # Compute returns
    if "cash_price" in result.columns:
        result["cash_return"] = np.log(result["cash_price"] / result["cash_price"].shift(1))
    if "3m_price" in result.columns:
        result["3m_return"] = np.log(result["3m_price"] / result["3m_price"].shift(1))

    return result


# ═══════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════

with st.sidebar:
    st.markdown('<p class="main-title">⚙️ Metals Dashboard</p>', unsafe_allow_html=True)
    st.markdown('<p class="main-subtitle">Risk Premia & Market Structure</p>', unsafe_allow_html=True)
    st.divider()

    # File uploads
    st.markdown("##### 📂 Data Files")
    cash_file = st.file_uploader("Metals Cash and 3M", type=["xlsx", "xls"], key="cash")
    curve_file = st.file_uploader("Metals Futures Curve", type=["xlsx", "xls", "csv", "xlsm"], key="curve")

    st.divider()

    # Global controls
    st.markdown("##### ⚙️ Controls")

    LME_METALS = ["Copper", "Aluminium", "Zinc", "Nickel", "Lead", "Tin"]

    if cash_file:
        cash_data = load_cash_3m_data(cash_file)
        available_metals = [m for m in LME_METALS if m in cash_data]
    else:
        available_metals = LME_METALS
        cash_data = {}

    selected_metal = st.selectbox("Metal", available_metals, index=0)

    # Date range
    if cash_data and selected_metal in cash_data:
        df_dates = cash_data[selected_metal]
        min_date = df_dates.index.min().date()
        max_date = df_dates.index.max().date()
    else:
        min_date = pd.Timestamp("2006-01-01").date()
        max_date = pd.Timestamp("2025-12-31").date()

    date_range = st.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

    if len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date, end_date = min_date, max_date

    st.divider()
    st.caption("NYU Financial Engineering")
    st.caption("Metals Risk Premia Project")


# ═══════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════

def filter_date(df, start, end):
    """Filter dataframe by date range. Handles both DatetimeIndex and non-datetime indexes."""
    if df.empty:
        return df
    try:
        # Ensure index is datetime
        if not isinstance(df.index, pd.DatetimeIndex):
            df = df.copy()
            df.index = pd.to_datetime(df.index, errors="coerce")
            df = df[df.index.notna()]
        if df.empty:
            return df
        mask = (df.index >= pd.Timestamp(start)) & (df.index <= pd.Timestamp(end))
        return df[mask]
    except Exception:
        return df


def metric_card(label, value, delta=None, unit=""):
    """Render a styled metric card."""
    delta_html = ""
    if delta is not None:
        cls = "delta-pos" if delta >= 0 else "delta-neg"
        sign = "+" if delta >= 0 else ""
        delta_html = f'<span class="{cls}">{sign}{delta:.2f}%</span>'

    st.markdown(f"""
    <div class="metric-card">
        <h4>{label}</h4>
        <p class="value">{value}{unit}</p>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)


def section_header(text):
    st.markdown(f'<div class="section-header">{text}</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# MAIN CONTENT
# ═══════════════════════════════════════════════

if not cash_file:
    st.markdown("## 📂 Upload Data to Begin")
    st.info("Upload **Metals Cash and 3M.xlsx** and optionally **Metals Futures Curve** file using the sidebar to explore the dashboard.")
    st.markdown("""
    **Expected file structure:**

    **File 1 — Metals Cash and 3M.xlsx:**
    One sheet per LME metal (LME Copper, LME Aluminium, ...) with columns for
    Cash Price, 3M Forward Price/Volume/OI, Cash-3M Spread Price/Volume.
    Plus a CME Cash Prices sheet for Gold, Silver, Platinum, Palladium, Copper ($/lb).

    **File 2 — Metals Futures Curve (.xlsx or .csv):**
    One sheet per metal with F1 through F27, each having Price, Volume, Open Interest columns.
    """)
    st.stop()


# Load futures curve data if available
curve_data = {}
if curve_file:
    curve_data = load_futures_curve_data(curve_file)


# Parse current metal data
if selected_metal in cash_data:
    metal_df = parse_cash_3m_columns(cash_data[selected_metal], selected_metal)
    metal_df = filter_date(metal_df, start_date, end_date)
else:
    st.warning(f"No data found for {selected_metal}")
    st.stop()


# ═══════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Market Overview",
    "📈 Term Structure",
    "💰 Cash vs 3M (Carry)",
    "📉 Volume & Open Interest",
    "🔗 Cross-Metal",
    "📋 Statistics"
])


# ══════════════════════════════════════════════════════
# TAB 1: MARKET OVERVIEW
# ══════════════════════════════════════════════════════

with tab1:
    st.markdown("### Market Overview")
    st.caption("Latest snapshot across all metals")

    # Build summary for all LME metals
    summary_rows = []
    for metal in LME_METALS:
        if metal not in cash_data:
            continue
        mdf = parse_cash_3m_columns(cash_data[metal], metal)
        mdf = filter_date(mdf, start_date, end_date)
        if mdf.empty:
            continue

        last = mdf.iloc[-1]
        prev = mdf.iloc[-2] if len(mdf) > 1 else mdf.iloc[-1]

        cash_p = last.get("cash_price", np.nan)
        tm_p = last.get("3m_price", np.nan)
        spread = last.get("spread_price", np.nan)
        cash_chg = ((cash_p / prev.get("cash_price", np.nan)) - 1) * 100 if pd.notna(prev.get("cash_price")) else 0

        summary_rows.append({
            "Metal": metal,
            "Cash": cash_p,
            "3M Forward": tm_p,
            "Cash-3M Spread": spread,
            "Daily Chg (%)": cash_chg,
            "Structure": "Backwardation" if (pd.notna(spread) and spread > 0) else "Contango",
        })

    if summary_rows:
        # Metric cards row
        cols = st.columns(min(len(summary_rows), 6))
        for i, row in enumerate(summary_rows):
            with cols[i % len(cols)]:
                badge = "backwardation" if row["Structure"] == "Backwardation" else "contango"
                metric_card(
                    row["Metal"],
                    f"${row['Cash']:,.0f}" if pd.notna(row["Cash"]) else "N/A",
                    row["Daily Chg (%)"] if pd.notna(row["Daily Chg (%)"]) else None,
                    ""
                )
                st.markdown(
                    f'<span class="badge-{badge}">{row["Structure"]}</span>',
                    unsafe_allow_html=True
                )

        st.markdown("")

        # CME Cash Prices
        if "CME_Cash" in cash_data:
            section_header("Precious Metals & COMEX Copper (Latest)")
            cme = cash_data["CME_Cash"]
            cme = filter_date(cme, start_date, end_date)
            if not cme.empty:
                last_cme = cme.iloc[-1]
                cme_cols = st.columns(min(len(cme.columns), 5))
                for j, col_name in enumerate(cme.columns):
                    with cme_cols[j % len(cme_cols)]:
                        val = last_cme[col_name]
                        # Determine unit
                        if "lb" in col_name.lower():
                            label = col_name.replace("($/lb)", "").strip()
                            unit_str = " $/lb"
                        else:
                            label = col_name.replace("($/oz)", "").strip()
                            unit_str = " $/oz"
                        if pd.notna(val):
                            metric_card(label, f"${val:,.2f}", unit=unit_str)

        st.divider()

        # Spread sparklines
        section_header("Cash-3M Spread — Last 90 Days")
        spark_cols = st.columns(3)
        for i, metal in enumerate([m for m in LME_METALS if m in cash_data]):
            mdf = parse_cash_3m_columns(cash_data[metal], metal)
            mdf = filter_date(mdf, start_date, end_date)
            if "spread_price" not in mdf.columns or mdf.empty:
                continue
            last_90 = mdf["spread_price"].dropna().tail(90)
            if last_90.empty:
                continue

            with spark_cols[i % 3]:
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=last_90.index, y=last_90.values,
                    mode="lines",
                    fill="tozeroy",
                    fillcolor="rgba(52,211,153,0.1)" if last_90.iloc[-1] > 0 else "rgba(248,113,113,0.1)",
                    line=dict(
                        color=COLORS["green"] if last_90.iloc[-1] > 0 else COLORS["red"],
                        width=2
                    ),
                    name=metal,
                    hovertemplate="%{x|%b %d}: $%{y:,.1f}<extra></extra>"
                ))
                fig.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
                fig.update_layout(
                    **CHART_LAYOUT,
                    height=180,
                    title=dict(text=metal, font=dict(size=13, color="#E2E8F0")),
                    showlegend=False,
                    margin=dict(l=40, r=10, t=35, b=25),
                    yaxis_title=None, xaxis_title=None,
                )
                st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 2: TERM STRUCTURE
# ══════════════════════════════════════════════════════

with tab2:
    st.markdown("### Term Structure (Futures Curve)")

    if not curve_data:
        st.info("Upload the **Metals Futures Curve** file to view term structure analysis.")
    else:
        # Select which sheet/metal
        curve_metals = list(curve_data.keys())
        if curve_metals:
            curve_metal = st.selectbox("Select Metal (Curve)", curve_metals, key="curve_metal")

            if curve_metal in curve_data and "prices" in curve_data[curve_metal]:
                prices_df = curve_data[curve_metal]["prices"]
                
                # Ensure prices_df has a valid DatetimeIndex
                if not prices_df.empty:
                    if not isinstance(prices_df.index, pd.DatetimeIndex):
                        # Try getting index from the raw data
                        raw_df = curve_data[curve_metal].get("raw", pd.DataFrame())
                        if isinstance(raw_df.index, pd.DatetimeIndex):
                            prices_df.index = raw_df.index[:len(prices_df)]
                        else:
                            try:
                                prices_df.index = pd.to_datetime(prices_df.index, errors="coerce")
                                prices_df = prices_df[prices_df.index.notna()]
                            except Exception:
                                st.warning("Could not parse dates from futures curve data.")
                                prices_df = pd.DataFrame()
                
                prices_df = filter_date(prices_df, start_date, end_date)

                # Debug: show parsed columns
                with st.expander("🔍 Debug: Parsed Data Info", expanded=False):
                    raw_cols = list(curve_data[curve_metal].get("raw", pd.DataFrame()).columns[:20])
                    st.write(f"**Raw columns (first 20):** {raw_cols}")
                    st.write(f"**Parsed price columns:** {list(prices_df.columns)}")
                    st.write(f"**Index type:** {type(prices_df.index).__name__}")
                    st.write(f"**Shape:** {prices_df.shape}")
                    if not prices_df.empty:
                        st.write(f"**Date range:** {prices_df.index.min()} → {prices_df.index.max()}")

                if not prices_df.empty and not prices_df.columns.empty:
                    # Date slider
                    available_dates = prices_df.dropna(how="all").index
                    if len(available_dates) > 0:
                        col1, col2 = st.columns([3, 1])

                        with col2:
                            # Multi-date comparison
                            st.markdown("##### Compare Dates")
                            n_compare = st.slider("Number of curves", 1, 5, 1, key="n_curves")

                        with col1:
                            date_idx = st.slider(
                                "Select date",
                                0, len(available_dates) - 1,
                                len(available_dates) - 1,
                                format="",
                                key="curve_date_slider"
                            )
                            selected_curve_date = available_dates[date_idx]
                            st.caption(f"**{selected_curve_date.strftime('%Y-%m-%d')}**")

                        # Build comparison dates
                        compare_dates = [selected_curve_date]
                        if n_compare > 1:
                            step = max(len(available_dates) // n_compare, 1)
                            for j in range(1, n_compare):
                                idx = max(date_idx - j * step, 0)
                                compare_dates.append(available_dates[idx])

                        compare_dates = sorted(set(compare_dates))

                        # Plot
                        fig = go.Figure()
                        compare_colors = [COLORS["primary"], COLORS["amber"], COLORS["accent"],
                                          COLORS["pink"], COLORS["green"]]

                        for k, dt in enumerate(compare_dates):
                            row = prices_df.loc[dt].dropna()
                            if row.empty:
                                continue
                            contracts = [c for c in row.index]
                            contract_nums = list(range(1, len(contracts) + 1))

                            is_latest = (dt == selected_curve_date)
                            fig.add_trace(go.Scatter(
                                x=contract_nums,
                                y=row.values,
                                mode="lines+markers",
                                name=dt.strftime("%Y-%m-%d"),
                                line=dict(
                                    color=compare_colors[k % len(compare_colors)],
                                    width=3 if is_latest else 1.5,
                                ),
                                marker=dict(size=6 if is_latest else 4),
                                opacity=1 if is_latest else 0.6,
                                hovertemplate="F%{x}: $%{y:,.2f}<extra>" + dt.strftime("%b %d, %Y") + "</extra>"
                            ))

                        fig.update_layout(
                            **CHART_LAYOUT,
                            height=500,
                            title=dict(text=f"{curve_metal} — Forward Curve", font=dict(size=16)),
                            xaxis_title="Contract Month",
                            yaxis_title="Price",
                            xaxis=dict(
                                tickmode="linear", dtick=1,
                                gridcolor="rgba(45,55,72,0.5)"
                            ),
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        # Contango/Backwardation indicator
                        latest_row = prices_df.loc[selected_curve_date].dropna()
                        if len(latest_row) >= 2:
                            slope = latest_row.iloc[-1] - latest_row.iloc[0]
                            if slope > 0:
                                st.success(f"📈 **Contango** — Deferred contracts trading above prompt ({curve_metal}, {selected_curve_date.strftime('%Y-%m-%d')})")
                            else:
                                st.warning(f"📉 **Backwardation** — Prompt contracts trading above deferred ({curve_metal}, {selected_curve_date.strftime('%Y-%m-%d')})")
                else:
                    st.warning("Could not parse futures price columns. Check column naming (expecting F1, F2, ... pattern with Price).")
            else:
                st.warning(f"No price data found for {curve_metal}")


# ══════════════════════════════════════════════════════
# TAB 3: CASH VS 3M (CARRY)
# ══════════════════════════════════════════════════════

with tab3:
    st.markdown(f"### {selected_metal} — Cash vs 3M Forward (Carry Analysis)")

    if "cash_price" not in metal_df.columns or "3m_price" not in metal_df.columns:
        st.warning("Cash or 3M price data not found for this metal.")
    else:
        # ── Cash vs 3M Price Chart ──
        section_header("Price Comparison")
        fig = make_subplots(rows=2, cols=1, shared_xaxes=True, row_heights=[0.65, 0.35],
                           vertical_spacing=0.06)

        fig.add_trace(go.Scatter(
            x=metal_df.index, y=metal_df["cash_price"],
            name="Cash (Spot)", line=dict(color=COLORS["amber"], width=2),
            hovertemplate="%{x|%b %d, %Y}<br>Cash: $%{y:,.2f}<extra></extra>"
        ), row=1, col=1)

        fig.add_trace(go.Scatter(
            x=metal_df.index, y=metal_df["3m_price"],
            name="3M Forward", line=dict(color=COLORS["primary"], width=2),
            hovertemplate="%{x|%b %d, %Y}<br>3M: $%{y:,.2f}<extra></extra>"
        ), row=1, col=1)

        # Spread
        if "spread_price" in metal_df.columns:
            spread = metal_df["spread_price"].dropna()
            colors_spread = [COLORS["green"] if v > 0 else COLORS["red"] for v in spread.values]

            fig.add_trace(go.Bar(
                x=spread.index, y=spread.values,
                name="Cash-3M Spread",
                marker_color=colors_spread,
                opacity=0.7,
                hovertemplate="%{x|%b %d, %Y}<br>Spread: $%{y:,.2f}<extra></extra>"
            ), row=2, col=1)

            fig.add_hline(y=0, row=2, col=1, line_dash="dash", line_color="#475569")

        fig.update_layout(
            **CHART_LAYOUT,
            height=600,
            title=dict(text=f"{selected_metal}: Cash vs 3M Forward", font=dict(size=14)),
        )
        fig.update_yaxes(title_text="Price ($/MT)", row=1, col=1)
        fig.update_yaxes(title_text="Spread", row=2, col=1)
        st.plotly_chart(fig, use_container_width=True)

        # ── Annualized Carry ──
        section_header("Annualized Carry (%)")
        if "spread_price" in metal_df.columns and "3m_price" in metal_df.columns:
            carry_pct = (metal_df["spread_price"] / metal_df["3m_price"]) * (365 / 90) * 100
            carry_pct = carry_pct.dropna()

            fig_carry = go.Figure()
            fig_carry.add_trace(go.Scatter(
                x=carry_pct.index, y=carry_pct.values,
                fill="tozeroy",
                fillcolor="rgba(59,130,246,0.1)",
                line=dict(color=COLORS["primary"], width=1.5),
                hovertemplate="%{x|%b %d, %Y}<br>Carry: %{y:.2f}%<extra></extra>"
            ))
            fig_carry.add_hline(y=0, line_dash="dash", line_color="#475569")
            fig_carry.update_layout(
                **CHART_LAYOUT,
                height=350,
                title=dict(text="Annualized Carry (Spread / 3M × 365/90)", font=dict(size=13)),
                yaxis_title="Carry (%)",
            )
            st.plotly_chart(fig_carry, use_container_width=True)

        # ── Spread Distribution ──
        section_header("Spread Distribution")
        if "spread_price" in metal_df.columns:
            spread_data = metal_df["spread_price"].dropna()
            col1, col2, col3 = st.columns(3)
            backw_pct = (spread_data > 0).sum() / len(spread_data) * 100
            with col1:
                metric_card("Backwardation", f"{backw_pct:.1f}%")
            with col2:
                metric_card("Contango", f"{100 - backw_pct:.1f}%")
            with col3:
                metric_card("Avg Spread", f"${spread_data.mean():,.1f}")

            fig_hist = go.Figure()
            fig_hist.add_trace(go.Histogram(
                x=spread_data.values,
                nbinsx=80,
                marker_color=COLORS["primary"],
                opacity=0.7,
                hovertemplate="Spread: $%{x:,.1f}<br>Count: %{y}<extra></extra>"
            ))
            fig_hist.add_vline(x=0, line_dash="dash", line_color=COLORS["amber"], line_width=2)
            fig_hist.update_layout(
                **CHART_LAYOUT,
                height=300,
                title=dict(text="Distribution of Cash-3M Spread", font=dict(size=13)),
                xaxis_title="Spread ($/MT)",
                yaxis_title="Frequency",
            )
            st.plotly_chart(fig_hist, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 4: VOLUME & OPEN INTEREST
# ══════════════════════════════════════════════════════

with tab4:
    st.markdown(f"### {selected_metal} — Volume & Open Interest")

    has_vol = "3m_volume" in metal_df.columns
    has_oi = "3m_oi" in metal_df.columns

    if not has_vol and not has_oi:
        st.info("Volume and Open Interest data not available for this metal in Cash & 3M file.")
    else:
        # ── 3M Volume & OI ──
        section_header("3M Forward — Volume & Open Interest")
        fig_vol = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                row_heights=[0.5, 0.5], vertical_spacing=0.08)

        if has_vol:
            vol = metal_df["3m_volume"].dropna()
            fig_vol.add_trace(go.Bar(
                x=vol.index, y=vol.values,
                name="Volume", marker_color=COLORS["primary"], opacity=0.6,
                hovertemplate="%{x|%b %d, %Y}<br>Vol: %{y:,.0f}<extra></extra>"
            ), row=1, col=1)
            # Rolling avg
            vol_ma = vol.rolling(20).mean()
            fig_vol.add_trace(go.Scatter(
                x=vol_ma.index, y=vol_ma.values,
                name="20D Avg", line=dict(color=COLORS["amber"], width=2),
            ), row=1, col=1)

        if has_oi:
            oi = metal_df["3m_oi"].dropna()
            fig_vol.add_trace(go.Scatter(
                x=oi.index, y=oi.values,
                name="Open Interest", line=dict(color=COLORS["accent"], width=2),
                fill="tozeroy", fillcolor="rgba(6,182,212,0.1)",
                hovertemplate="%{x|%b %d, %Y}<br>OI: %{y:,.0f}<extra></extra>"
            ), row=2, col=1)

        fig_vol.update_layout(
            **CHART_LAYOUT,
            height=500,
            title=dict(text=f"{selected_metal} 3M Forward", font=dict(size=14)),
        )
        fig_vol.update_yaxes(title_text="Volume", row=1, col=1)
        fig_vol.update_yaxes(title_text="Open Interest", row=2, col=1)
        st.plotly_chart(fig_vol, use_container_width=True)

        # ── Spread Volume (if available) ──
        if "spread_volume" in metal_df.columns:
            section_header("Cash-3M Spread Volume")
            sv = metal_df["spread_volume"].dropna()
            if not sv.empty:
                fig_sv = go.Figure()
                fig_sv.add_trace(go.Bar(
                    x=sv.index, y=sv.values,
                    marker_color=COLORS["secondary"], opacity=0.6,
                    hovertemplate="%{x|%b %d, %Y}<br>Spread Vol: %{y:,.0f}<extra></extra>"
                ))
                sv_ma = sv.rolling(20).mean()
                fig_sv.add_trace(go.Scatter(
                    x=sv_ma.index, y=sv_ma.values,
                    name="20D Avg", line=dict(color=COLORS["amber"], width=2),
                ))
                fig_sv.update_layout(
                    **CHART_LAYOUT, height=300,
                    title=dict(text="Spread Volume", font=dict(size=13)),
                )
                st.plotly_chart(fig_sv, use_container_width=True)

        # ── Futures strip volume heatmap ──
        if curve_data:
            section_header("Futures Strip Volume Heatmap")

            # Find matching curve sheet
            curve_match = None
            for sheet_name in curve_data:
                if selected_metal.lower() in sheet_name.lower():
                    curve_match = sheet_name
                    break

            if curve_match and "raw" in curve_data[curve_match]:
                raw_curve = curve_data[curve_match]["raw"]
                raw_curve = filter_date(raw_curve, start_date, end_date)

                # Extract volume columns
                vol_cols = [c for c in raw_curve.columns if "volume" in c.lower()]
                if vol_cols:
                    vol_df = raw_curve[vol_cols].copy()
                    # Clean column names
                    vol_df.columns = [c.split("_")[0] if "_" in c else c for c in vol_df.columns]

                    # Resample to monthly for readability
                    vol_monthly = vol_df.resample("M").mean()
                    vol_monthly = vol_monthly.tail(36)  # Last 3 years

                    if not vol_monthly.empty:
                        fig_hm = go.Figure(data=go.Heatmap(
                            z=vol_monthly.values.T,
                            x=vol_monthly.index.strftime("%Y-%m"),
                            y=vol_monthly.columns,
                            colorscale="Viridis",
                            hovertemplate="Date: %{x}<br>Contract: %{y}<br>Avg Volume: %{z:,.0f}<extra></extra>"
                        ))
                        fig_hm.update_layout(
                            **CHART_LAYOUT,
                            height=400,
                            title=dict(text=f"{selected_metal} — Monthly Average Volume by Contract", font=dict(size=13)),
                            xaxis_title="Month",
                            yaxis_title="Contract",
                        )
                        st.plotly_chart(fig_hm, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 5: CROSS-METAL (LME Copper vs COMEX Copper)
# ══════════════════════════════════════════════════════

with tab5:
    st.markdown("### LME Copper vs COMEX Copper (HG)")
    st.caption("Location arbitrage: LME $/MT vs COMEX ¢/lb")

    has_lme_cu = "Copper" in cash_data
    has_cme_cu = "CME_Cash" in cash_data

    if not has_lme_cu or not has_cme_cu:
        st.info("Both **LME Copper** and **CME Cash Prices** sheets are needed for this analysis.")
    else:
        lme_cu = parse_cash_3m_columns(cash_data["Copper"], "Copper")
        lme_cu = filter_date(lme_cu, start_date, end_date)

        cme_cash = cash_data["CME_Cash"]
        cme_cash = filter_date(cme_cash, start_date, end_date)

        # Find copper column in CME cash
        cu_cme_col = [c for c in cme_cash.columns if "copper" in c.lower() or "cu" in c.lower()]

        if not cu_cme_col:
            st.warning("Copper column not found in CME Cash Prices sheet.")
        else:
            cu_cme_col = cu_cme_col[0]
            cme_cu_price = cme_cash[cu_cme_col].dropna()

            # Convert COMEX $/lb to $/MT for comparison
            # 1 MT = 2204.62 lbs
            LBS_PER_MT = 2204.62
            cme_cu_mt = cme_cu_price * LBS_PER_MT  # $/lb × lbs/MT = $/MT

            # Align dates
            combined = pd.DataFrame({
                "LME_Cash": lme_cu["cash_price"] if "cash_price" in lme_cu.columns else lme_cu.get("3m_price"),
                "COMEX_MT": cme_cu_mt,
            }).dropna()

            if combined.empty:
                st.warning("No overlapping dates between LME and COMEX copper data.")
            else:
                combined["Spread"] = combined["LME_Cash"] - combined["COMEX_MT"]
                combined["Ratio"] = combined["LME_Cash"] / combined["COMEX_MT"]

                # Summary metrics
                col1, col2, col3, col4 = st.columns(4)
                last = combined.iloc[-1]
                with col1:
                    metric_card("LME Cash", f"${last['LME_Cash']:,.0f}", unit=" /MT")
                with col2:
                    metric_card("COMEX (conv.)", f"${last['COMEX_MT']:,.0f}", unit=" /MT")
                with col3:
                    metric_card("Spread", f"${last['Spread']:,.0f}", unit=" /MT")
                with col4:
                    metric_card("Ratio", f"{last['Ratio']:.4f}")

                # Price comparison
                section_header("LME vs COMEX — Price in $/MT")
                fig_xm = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                       row_heights=[0.6, 0.4], vertical_spacing=0.06)

                fig_xm.add_trace(go.Scatter(
                    x=combined.index, y=combined["LME_Cash"],
                    name="LME Copper Cash", line=dict(color=COLORS["orange"], width=2),
                    hovertemplate="%{x|%b %d, %Y}<br>LME: $%{y:,.0f}/MT<extra></extra>"
                ), row=1, col=1)

                fig_xm.add_trace(go.Scatter(
                    x=combined.index, y=combined["COMEX_MT"],
                    name="COMEX Copper (conv. $/MT)", line=dict(color=COLORS["primary"], width=2),
                    hovertemplate="%{x|%b %d, %Y}<br>COMEX: $%{y:,.0f}/MT<extra></extra>"
                ), row=1, col=1)

                # Spread
                spread_colors = [COLORS["green"] if v > 0 else COLORS["red"] for v in combined["Spread"]]
                fig_xm.add_trace(go.Bar(
                    x=combined.index, y=combined["Spread"],
                    name="LME - COMEX Spread",
                    marker_color=spread_colors, opacity=0.6,
                    hovertemplate="%{x|%b %d, %Y}<br>Spread: $%{y:,.0f}/MT<extra></extra>"
                ), row=2, col=1)
                fig_xm.add_hline(y=0, row=2, col=1, line_dash="dash", line_color="#475569")

                fig_xm.update_layout(
                    **CHART_LAYOUT, height=550,
                    title=dict(text="LME vs COMEX Copper", font=dict(size=14)),
                )
                fig_xm.update_yaxes(title_text="Price ($/MT)", row=1, col=1)
                fig_xm.update_yaxes(title_text="Spread ($/MT)", row=2, col=1)
                st.plotly_chart(fig_xm, use_container_width=True)

                # Rolling correlation
                section_header("Rolling 60-Day Correlation")
                rolling_corr = combined["LME_Cash"].rolling(60).corr(combined["COMEX_MT"])

                fig_corr = go.Figure()
                fig_corr.add_trace(go.Scatter(
                    x=rolling_corr.index, y=rolling_corr.values,
                    fill="tozeroy", fillcolor="rgba(139,92,246,0.15)",
                    line=dict(color=COLORS["secondary"], width=2),
                    hovertemplate="%{x|%b %d, %Y}<br>Corr: %{y:.4f}<extra></extra>"
                ))
                fig_corr.update_layout(
                    **CHART_LAYOUT, height=300,
                    title=dict(text="LME-COMEX Rolling Correlation (60D)", font=dict(size=13)),
                    yaxis_title="Correlation",
                    yaxis_range=[0.5, 1.05],
                )
                st.plotly_chart(fig_corr, use_container_width=True)

                # Spread statistics
                section_header("Spread Statistics")
                col1, col2 = st.columns(2)
                with col1:
                    spread_stats = combined["Spread"].describe()
                    st.dataframe(spread_stats.to_frame("LME-COMEX Spread ($/MT)").style.format("{:,.2f}"))
                with col2:
                    fig_sp_hist = go.Figure()
                    fig_sp_hist.add_trace(go.Histogram(
                        x=combined["Spread"].values, nbinsx=60,
                        marker_color=COLORS["secondary"], opacity=0.7,
                    ))
                    fig_sp_hist.add_vline(x=0, line_dash="dash", line_color=COLORS["amber"])
                    fig_sp_hist.update_layout(
                        **CHART_LAYOUT, height=300,
                        title=dict(text="Spread Distribution", font=dict(size=13)),
                        xaxis_title="Spread ($/MT)",
                    )
                    st.plotly_chart(fig_sp_hist, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 6: STATISTICS
# ══════════════════════════════════════════════════════

with tab6:
    st.markdown("### Descriptive Statistics")

    # ── Summary table across all metals ──
    section_header("Summary Statistics — All LME Metals")

    stats_rows = []
    for metal in LME_METALS:
        if metal not in cash_data:
            continue
        mdf = parse_cash_3m_columns(cash_data[metal], metal)
        mdf = filter_date(mdf, start_date, end_date)
        if "cash_price" not in mdf.columns:
            continue

        cash = mdf["cash_price"].dropna()
        rets = mdf.get("cash_return", pd.Series(dtype=float)).dropna()
        spread = mdf.get("spread_price", pd.Series(dtype=float)).dropna()

        row = {
            "Metal": metal,
            "Obs": len(cash),
            "Start": cash.index.min().strftime("%Y-%m-%d") if len(cash) > 0 else "",
            "End": cash.index.max().strftime("%Y-%m-%d") if len(cash) > 0 else "",
            "Mean Price": cash.mean(),
            "Std Price": cash.std(),
            "Min Price": cash.min(),
            "Max Price": cash.max(),
        }

        if len(rets) > 10:
            row["Ann. Return"] = rets.mean() * 252 * 100
            row["Ann. Vol"] = rets.std() * np.sqrt(252) * 100
            row["Sharpe"] = (rets.mean() / rets.std()) * np.sqrt(252) if rets.std() > 0 else 0
            row["Skew"] = rets.skew()
            row["Kurtosis"] = rets.kurtosis()

        if len(spread) > 0:
            row["Avg Spread"] = spread.mean()
            row["Backw. %"] = (spread > 0).sum() / len(spread) * 100

        stats_rows.append(row)

    if stats_rows:
        stats_df = pd.DataFrame(stats_rows).set_index("Metal")
        fmt_dict = {c: "{:,.2f}" for c in stats_df.columns if stats_df[c].dtype in ["float64", "float32"]}
        fmt_dict.update({"Obs": "{:,.0f}", "Backw. %": "{:.1f}%"})
        st.dataframe(stats_df.style.format(fmt_dict, na_rep="—"), use_container_width=True)

    # ── Rolling Volatility ──
    section_header(f"{selected_metal} — Rolling Volatility")
    if "cash_return" in metal_df.columns:
        rets = metal_df["cash_return"].dropna()

        fig_vol = go.Figure()
        for window, color, name in [(30, COLORS["primary"], "30D"), (60, COLORS["amber"], "60D"), (90, COLORS["accent"], "90D")]:
            rv = rets.rolling(window).std() * np.sqrt(252) * 100
            fig_vol.add_trace(go.Scatter(
                x=rv.index, y=rv.values,
                name=name, line=dict(color=color, width=1.5),
                hovertemplate="%{x|%b %d, %Y}<br>" + name + ": %{y:.1f}%<extra></extra>"
            ))

        fig_vol.update_layout(
            **CHART_LAYOUT, height=400,
            title=dict(text=f"{selected_metal} — Annualized Realized Volatility", font=dict(size=14)),
            yaxis_title="Volatility (%)",
        )
        st.plotly_chart(fig_vol, use_container_width=True)

    # ── Return Distribution ──
    section_header(f"{selected_metal} — Return Distribution")
    if "cash_return" in metal_df.columns:
        rets = metal_df["cash_return"].dropna()

        col1, col2 = st.columns([2, 1])
        with col1:
            fig_rd = go.Figure()
            fig_rd.add_trace(go.Histogram(
                x=rets.values * 100, nbinsx=100,
                marker_color=COLORS["primary"], opacity=0.7,
                name="Daily Returns",
            ))
            fig_rd.add_vline(x=0, line_dash="dash", line_color=COLORS["amber"])
            fig_rd.update_layout(
                **CHART_LAYOUT, height=350,
                title=dict(text="Daily Log Return Distribution (%)", font=dict(size=13)),
                xaxis_title="Return (%)",
                yaxis_title="Frequency",
            )
            st.plotly_chart(fig_rd, use_container_width=True)

        with col2:
            st.markdown("##### Return Statistics")
            ret_stats = {
                "Mean (daily)": f"{rets.mean() * 100:.4f}%",
                "Std (daily)": f"{rets.std() * 100:.4f}%",
                "Skewness": f"{rets.skew():.4f}",
                "Kurtosis": f"{rets.kurtosis():.4f}",
                "Min": f"{rets.min() * 100:.2f}%",
                "Max": f"{rets.max() * 100:.2f}%",
                "Ann. Return": f"{rets.mean() * 252 * 100:.2f}%",
                "Ann. Volatility": f"{rets.std() * np.sqrt(252) * 100:.2f}%",
                "Sharpe": f"{(rets.mean() / rets.std()) * np.sqrt(252):.4f}" if rets.std() > 0 else "N/A",
            }
            for k, v in ret_stats.items():
                st.markdown(f"**{k}:** `{v}`")

    # ── Regime Analysis ──
    section_header(f"{selected_metal} — Regime Breakdown")
    if "cash_return" in metal_df.columns:
        regimes = {
            "Pre-2015": (pd.Timestamp("2000-01-01"), pd.Timestamp("2014-12-31")),
            "2015–2019": (pd.Timestamp("2015-01-01"), pd.Timestamp("2019-12-31")),
            "COVID (2020–2021)": (pd.Timestamp("2020-01-01"), pd.Timestamp("2021-12-31")),
            "Post-2022": (pd.Timestamp("2022-01-01"), pd.Timestamp("2030-12-31")),
        }

        regime_stats = []
        for name, (s, e) in regimes.items():
            mask = (metal_df.index >= s) & (metal_df.index <= e)
            sub = metal_df[mask]
            if sub.empty or "cash_return" not in sub.columns:
                continue
            r = sub["cash_return"].dropna()
            sp = sub.get("spread_price", pd.Series(dtype=float)).dropna()
            if len(r) < 10:
                continue
            regime_stats.append({
                "Regime": name,
                "Obs": len(r),
                "Ann. Return (%)": r.mean() * 252 * 100,
                "Ann. Vol (%)": r.std() * np.sqrt(252) * 100,
                "Sharpe": (r.mean() / r.std()) * np.sqrt(252) if r.std() > 0 else 0,
                "Avg Spread": sp.mean() if len(sp) > 0 else np.nan,
                "Backw. (%)": (sp > 0).sum() / len(sp) * 100 if len(sp) > 0 else np.nan,
            })

        if regime_stats:
            regime_df = pd.DataFrame(regime_stats).set_index("Regime")
            st.dataframe(
                regime_df.style.format({
                    "Obs": "{:,.0f}",
                    "Ann. Return (%)": "{:.2f}",
                    "Ann. Vol (%)": "{:.2f}",
                    "Sharpe": "{:.4f}",
                    "Avg Spread": "{:,.1f}",
                    "Backw. (%)": "{:.1f}",
                }, na_rep="—"),
                use_container_width=True
            )
