"""
========================================================
GREENHOUSE DASHBOARD - FOOP Greenhouse Project
========================================================
HOW THIS FILE IS ORGANISED:
  1. IMPORTS          — load the tools Python needs
  2. PAGE SETUP       — title, layout, styling
  3. DATA LOADING     — read your Excel file
  4. SIDEBAR          — filters (date range, crop, market)
  5. HEADER           — logo, photo, title
  6. TAB 1 — SALES OVERVIEW
  7. TAB 2 — CROP PERFORMANCE
  8. TAB 3 — MARKET COMPARISON (Scenarios)
  9. TAB 4 — PRODUCTION TRACKER (manual entry)

To run this app:
  python -m streamlit run dashboard.py
"""

# ─────────────────────────────────────────────────────────
# 1. IMPORTS
# ─────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
from pathlib import Path

# ─────────────────────────────────────────────────────────
# 2. PAGE SETUP
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FOOP Greenhouse Dashboard",
    page_icon="🌿",
    layout="wide"
)

st.markdown("""
<style>
    [data-testid="stMetric"] {
        border: 1px solid rgba(128,128,128,0.25);
        border-radius: 8px;
        padding: 12px 16px;
    }
    [data-testid="stMetricValue"] { font-size: 1.4rem; font-weight: 600; }
    [data-testid="stMetricLabel"] { font-size: 0.85rem; opacity: 0.75; }
    h1 { color: #2e7d32; }
    h2 { color: #388e3c; }
    h3 { color: #43a047; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
# 3. DATA LOADING
# ─────────────────────────────────────────────────────────
@st.cache_data
def load_sales_data(filepath):
    df = pd.read_excel(filepath, sheet_name="Daily Sales Record", header=2)
    df.columns = df.columns.str.strip()

    cols_needed = [
        "Date", "Market", "Vegetable",
        "Quantity Delivered (kg)", "Quantity Sold (kg)",
        "kg Spoiled/Loss", "Loss %",
        "Units Sold in ", "Units Sold",
        "Price per Unit (NLE)", "Total Revenue (NLE)",
        "Transport Cost (NLE)", "Net Revenue (NLE)",
        "NLE per Kg earned", "Notes"
    ]
    cols_present = [c for c in cols_needed if c in df.columns]
    df = df[cols_present].copy()

    df.rename(columns={
        "Quantity Delivered (kg)": "Qty Delivered",
        "Quantity Sold (kg)": "Qty Sold",
        "kg Spoiled/Loss": "Spoilage",
        "Loss %": "Loss Pct",
        "Units Sold in ": "Unit Type",
        "Units Sold": "Units Sold",
        "Price per Unit (NLE)": "Price per Unit",
        "Total Revenue (NLE)": "Total Revenue",
        "Transport Cost (NLE)": "Transport Cost",
        "Net Revenue (NLE)": "Net Revenue",
        "NLE per Kg earned": "NLE per kg"
    }, inplace=True)

    def fix_date(val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (int, float)):
            try:
                return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))
            except:
                return pd.NaT
        if isinstance(val, str):
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"]:
                try:
                    return pd.Timestamp(datetime.strptime(val.strip(), fmt))
                except:
                    continue
        try:
            return pd.Timestamp(val)
        except:
            return pd.NaT

    df["Date"] = df["Date"].apply(fix_date)
    df = df.dropna(subset=["Date", "Vegetable"])
    df = df[df["Vegetable"].astype(str).str.strip() != ""]

    for col in ["Qty Delivered", "Qty Sold", "Spoilage", "Total Revenue",
                "Net Revenue", "Transport Cost", "NLE per kg", "Price per Unit"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Market"] = df["Market"].astype(str).str.strip()
    df["Vegetable"] = df["Vegetable"].astype(str).str.strip()
    df = df.sort_values("Date").reset_index(drop=True)
    return df


@st.cache_data
def load_scenarios_data(filepath):
    """
    Reads the Scenarios Profitability Tracker sheet.
    Columns read by position to avoid duplicate-name confusion.
    """
    try:
        raw = pd.read_excel(
            filepath,
            sheet_name="Scenarios Profitability Tracker",
            header=None
        )
        data = raw.iloc[4:].copy()
        data.columns = range(len(data.columns))

        col_map = {
            0: "Date", 1: "Vegetable", 2: "Qty Delivered",
            3: "Tormabum Revenue", 4: "Tormabum Transport", 5: "Tormabum Net",
            6: "Bo Revenue",       7: "Bo Transport",       8: "Bo Net",
            9: "Waterloo Revenue", 10: "Waterloo Transport", 11: "Waterloo Net",
            12: "Tormabum Net/kg", 13: "Bo Net/kg", 14: "Waterloo Net/kg",
            15: "Best Market"
        }
        data = data.rename(columns=col_map)

        data = data[data["Vegetable"].notna()]
        data = data[~data["Vegetable"].astype(str).str.upper().isin(
            ["TOTAL", "NAN", "VEGETABLE", ""]
        )]
        data = data[data["Date"].notna()]
        data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
        data = data.dropna(subset=["Date"])

        for col in ["Qty Delivered", "Tormabum Revenue", "Tormabum Transport",
                    "Tormabum Net", "Bo Revenue", "Bo Transport", "Bo Net",
                    "Waterloo Revenue", "Waterloo Transport", "Waterloo Net",
                    "Tormabum Net/kg", "Bo Net/kg", "Waterloo Net/kg"]:
            data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

        data["Vegetable"] = data["Vegetable"].astype(str).str.strip()
        return data.reset_index(drop=True)

    except Exception as e:
        return pd.DataFrame({"_error": [str(e)]})


@st.cache_data
def load_market_prices(filepath):
    df = pd.read_excel(filepath, sheet_name="Market prices", header=2)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Market", "Vegetable"])
    return df


@st.cache_data
def load_production_data(filepath):
    """
    Reads the Production Tracker sheet from your Excel file.
    One row per crop batch. You manage this sheet directly in Excel.
    """
    try:
        df = pd.read_excel(filepath, sheet_name="Production Tracker", header=0)
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=["Crop"])
        df = df[df["Crop"].astype(str).str.strip() != ""]
        df["Date Planted"] = pd.to_datetime(df["Date Planted"], errors="coerce")
        for col in ["Rows", "Total Plants", "Expected Harvest (kg)"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in ["First Harvest Date", "Last Harvest Date"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
            else:
                df[col] = pd.NaT
        if "Status" not in df.columns:
            df["Status"] = "Unknown"
        df["Status"] = df["Status"].astype(str).str.strip()
        if "Location" not in df.columns:
            df["Location"] = "Unknown"
        df["Location"] = df["Location"].astype(str).str.strip()
        if "Notes" not in df.columns:
            df["Notes"] = ""
        return df.reset_index(drop=True)
    except Exception as e:
        return pd.DataFrame({"_error": [str(e)]})


@st.cache_data
def load_harvest_log(filepath):
    """
    Reads the Harvest Log sheet.
    Each row is one harvest event: Date, Crop, Kg Harvested, Notes.
    """
    try:
        df = pd.read_excel(filepath, sheet_name="Harvest Log", header=0)
        df.columns = df.columns.str.strip()
        if df.empty or "Crop" not in df.columns:
            return pd.DataFrame(columns=["Date", "Crop", "Kg Harvested", "Notes"])
        df = df.dropna(subset=["Crop"])
        df = df[df["Crop"].astype(str).str.strip() != ""]
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Kg Harvested"] = pd.to_numeric(df["Kg Harvested"], errors="coerce").fillna(0)
        df["Crop"] = df["Crop"].astype(str).str.strip()
        if "Notes" not in df.columns:
            df["Notes"] = ""
        return df.sort_values("Date").reset_index(drop=True)
    except Exception as e:
        return pd.DataFrame(columns=["Date", "Crop", "Kg Harvested", "Notes"])


# ─────────────────────────────────────────────────────────
# 4. DATA FILE
# The Excel file loads automatically from the same folder as this script.
# To update the data: just replace Greenhouse_Sales_Book.xlsx with the
# new version and click "Reload data" in the sidebar — no upload needed.
# ─────────────────────────────────────────────────────────
DATA_FILE = Path(__file__).parent / "Greenhouse_Sales_Book.xlsx"

if not DATA_FILE.exists():
    logo_path = Path(__file__).parent / "Foop_logo_spiga__1_.jpg"
    if logo_path.exists():
        st.image(str(logo_path), width=180)
    st.error("Data file not found! Make sure Greenhouse_Sales_Book.xlsx is in the same folder as dashboard.py.")
    st.code(str(DATA_FILE))
    st.stop()

# Load data — cached so it only reloads when you click Reload
df               = load_sales_data(str(DATA_FILE))
scenarios_df     = load_scenarios_data(str(DATA_FILE))
market_prices_df = load_market_prices(str(DATA_FILE))
production_df    = load_production_data(str(DATA_FILE))
harvest_df       = load_harvest_log(str(DATA_FILE))

# Sidebar reload button
st.sidebar.header("📂 Data")
st.sidebar.caption("Source: Greenhouse_Sales_Book.xlsx")
if st.sidebar.button("🔄 Reload data"):
    st.cache_data.clear()
    st.rerun()

# ── Filters ──
st.sidebar.header("🔍 Filters")

min_date = df["Date"].min().date()
max_date = df["Date"].max().date()

date_range = st.sidebar.date_input(
    "Date Range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date = end_date = date_range[0] if date_range else min_date

all_crops   = sorted(df["Vegetable"].unique().tolist())
all_markets = sorted(df["Market"].unique().tolist())

selected_crops = st.sidebar.multiselect(
    "Crops", options=all_crops, default=all_crops
)
selected_markets = st.sidebar.multiselect(
    "Markets", options=all_markets, default=all_markets
)

# Apply filters to sales data
mask = (
    (df["Date"].dt.date >= start_date) &
    (df["Date"].dt.date <= end_date) &
    (df["Vegetable"].isin(selected_crops)) &
    (df["Market"].isin(selected_markets))
)
filtered_df = df[mask].copy()

# Apply filters to scenarios data (date + crop only — no market filter
# because scenarios show all markets per sale)
if not scenarios_df.empty and "_error" not in scenarios_df.columns:
    scen_mask = (
        (scenarios_df["Date"].dt.date >= start_date) &
        (scenarios_df["Date"].dt.date <= end_date) &
        (scenarios_df["Vegetable"].isin(selected_crops))
    )
    filtered_scen = scenarios_df[scen_mask].copy()
else:
    filtered_scen = scenarios_df.copy()

if filtered_df.empty:
    st.warning("No data matches your filters. Try widening the date range or selecting more crops.")
    st.stop()

# ─────────────────────────────────────────────────────────
# 5. HEADER — Logo + Greenhouse photo + Title
# ─────────────────────────────────────────────────────────
base_dir     = Path(__file__).parent
logo_path    = base_dir / "Foop_logo_spiga__1_.jpg"
photo_path   = base_dir / "Greenhouse_Photo.jpg"

col_logo, col_title = st.columns([1, 5])

with col_logo:
    if logo_path.exists():
        st.image(str(logo_path), width=120)

with col_title:
    st.title("Greenhouse Project Dashboard")

# Greenhouse photo as a full-width banner
if photo_path.exists():
    st.image(str(photo_path), use_container_width=True,
             caption="FOOP Greenhouse Facility")

st.divider()

# ─────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Sales Overview",
    "🌿 Crop Performance",
    "🏪 Market Comparison",
    "🌱 Production Tracker"
])

# ══════════════════════════════════════════════════════════
# TAB 1 — SALES OVERVIEW
# ══════════════════════════════════════════════════════════
with tab1:
    st.header("Sales Overview")

    col1, col2, col3, col4, col5 = st.columns(5)
    total_revenue  = filtered_df["Total Revenue"].sum()
    total_net      = filtered_df["Net Revenue"].sum()
    total_kg_sold  = filtered_df["Qty Sold"].sum()
    total_spoilage = filtered_df["Spoilage"].sum()
    total_delivered = filtered_df["Qty Delivered"].sum()
    num_sales      = len(filtered_df)

    col1.metric("Total Revenue (NLE)", f"{total_revenue:,.0f}")
    col2.metric("Net Revenue (NLE)",   f"{total_net:,.0f}")
    col3.metric("Total kg Sold",       f"{total_kg_sold:,.1f} kg")
    col4.metric(
        "Total Spoilage",
        f"{total_spoilage:,.1f} kg",
        delta=f"{(total_spoilage / total_delivered * 100) if total_delivered else 0:.1f}% loss rate",
        delta_color="inverse"
    )
    col5.metric("Number of Sales", num_sales)

    st.divider()

    # Date axis range — constrained to actual data only
    date_min = filtered_df["Date"].min()
    date_max = filtered_df["Date"].max()

    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("Revenue Over Time")
        rev_over_time = (
            filtered_df.groupby("Date")[["Total Revenue", "Net Revenue"]]
            .sum().reset_index()
        )
        fig_time = px.line(
            rev_over_time, x="Date",
            y=["Total Revenue", "Net Revenue"],
            labels={"value": "NLE", "variable": ""},
            color_discrete_map={
                "Total Revenue": "#2e7d32",
                "Net Revenue": "#81c784"
            }
        )
        # Fix: lock x-axis to actual data range only
        fig_time.update_xaxes(range=[date_min, date_max])
        fig_time.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_time, use_container_width=True)

    with col_right:
        st.subheader("Revenue by Crop")
        rev_by_crop = (
            filtered_df.groupby("Vegetable")["Net Revenue"]
            .sum().sort_values(ascending=True).reset_index()
        )
        fig_crop = px.bar(
            rev_by_crop, x="Net Revenue", y="Vegetable",
            orientation="h", color="Net Revenue",
            color_continuous_scale="Greens",
            labels={"Net Revenue": "Net Revenue (NLE)"}
        )
        fig_crop.update_layout(
            coloraxis_showscale=False,
            margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_crop, use_container_width=True)

    col_left2, col_right2 = st.columns(2)

    with col_left2:
        st.subheader("Revenue by Market")
        rev_by_market = (
            filtered_df.groupby("Market")["Net Revenue"]
            .sum().reset_index()
        )
        fig_mkt = px.pie(
            rev_by_market, values="Net Revenue", names="Market",
            color_discrete_sequence=px.colors.sequential.Greens[2:]
        )
        fig_mkt.update_layout(margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig_mkt, use_container_width=True)

    with col_right2:
        st.subheader("kg Delivered vs Sold vs Spoiled")
        fig_bar = go.Figure(go.Bar(
            x=["Delivered", "Sold", "Spoiled"],
            y=[total_delivered, total_kg_sold, total_spoilage],
            marker_color=["#2e7d32", "#66bb6a", "#ef5350"]
        ))
        fig_bar.update_layout(
            yaxis_title="kg",
            margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    st.subheader("📋 Sales Records")
    st.dataframe(
        filtered_df[[
            "Date", "Market", "Vegetable",
            "Qty Delivered", "Qty Sold", "Spoilage",
            "Total Revenue", "Net Revenue", "NLE per kg", "Notes"
        ]].sort_values("Date", ascending=False),
        use_container_width=True, hide_index=True
    )

    csv = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download filtered data as CSV",
        data=csv, file_name="greenhouse_sales_filtered.csv", mime="text/csv"
    )


# ══════════════════════════════════════════════════════════
# TAB 2 — CROP PERFORMANCE
# ══════════════════════════════════════════════════════════
with tab2:
    st.header("Crop Performance Analysis")
    st.caption("Critically analyse how each crop is performing over time")

    date_min = filtered_df["Date"].min()
    date_max = filtered_df["Date"].max()

    st.subheader("NLE per kg Earned — by Crop Over Time")
    st.caption(
        "Higher = you earned more per kg sold. "
        "If a crop's line is falling, you're getting less per kg over time."
    )
    fig_nle = px.line(
        filtered_df.sort_values("Date"),
        x="Date", y="NLE per kg", color="Vegetable",
        markers=True,
        labels={"NLE per kg": "NLE earned per kg"}
    )
    # Fix: lock x-axis to actual data range only
    fig_nle.update_xaxes(range=[date_min, date_max])
    fig_nle.update_layout(
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(l=0, r=0, t=40, b=0)
    )
    st.plotly_chart(fig_nle, use_container_width=True)

    st.divider()

    st.subheader("Crop Summary — All Time")
    crop_summary = filtered_df.groupby("Vegetable").agg(
        Total_Deliveries=("Qty Delivered", "sum"),
        Total_Sold=("Qty Sold", "sum"),
        Total_Spoilage=("Spoilage", "sum"),
        Total_Revenue=("Total Revenue", "sum"),
        Total_Net_Revenue=("Net Revenue", "sum"),
        Avg_NLE_per_kg=("NLE per kg", "mean"),
        Num_Sales=("Date", "count")
    ).reset_index()

    crop_summary["Sell-through Rate"] = (
        crop_summary["Total_Sold"] / crop_summary["Total_Deliveries"] * 100
    ).round(1).astype(str) + "%"
    crop_summary["Avg_NLE_per_kg"]    = crop_summary["Avg_NLE_per_kg"].round(2)
    crop_summary["Total_Revenue"]     = crop_summary["Total_Revenue"].round(0)
    crop_summary["Total_Net_Revenue"] = crop_summary["Total_Net_Revenue"].round(0)

    st.dataframe(
        crop_summary.rename(columns={
            "Vegetable": "Crop",
            "Total_Deliveries": "kg Delivered",
            "Total_Sold": "kg Sold",
            "Total_Spoilage": "kg Spoiled",
            "Total_Revenue": "Total Revenue (NLE)",
            "Total_Net_Revenue": "Net Revenue (NLE)",
            "Avg_NLE_per_kg": "Avg NLE/kg",
            "Num_Sales": "# Sales"
        }),
        use_container_width=True, hide_index=True
    )

    st.divider()

    col_l, col_r = st.columns(2)

    with col_l:
        st.subheader("Spoilage Rate by Crop")
        st.caption("Lower % = less waste.")
        spoilage_data = filtered_df.groupby("Vegetable").agg(
            Delivered=("Qty Delivered", "sum"),
            Spoiled=("Spoilage", "sum")
        ).reset_index()
        spoilage_data["Spoilage %"] = (
            spoilage_data["Spoiled"] / spoilage_data["Delivered"] * 100
        ).round(1)
        fig_spoil = px.bar(
            spoilage_data, x="Vegetable", y="Spoilage %",
            color="Spoilage %", color_continuous_scale="RdYlGn_r"
        )
        fig_spoil.add_hline(
            y=10, line_dash="dash", line_color="red",
            annotation_text="10% warning threshold"
        )
        fig_spoil.update_layout(
            coloraxis_showscale=False, margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_spoil, use_container_width=True)

    with col_r:
        st.subheader("Average NLE/kg by Crop")
        st.caption("Which crop earns the most per kg sold?")
        avg_nle = (
            filtered_df.groupby("Vegetable")["NLE per kg"]
            .mean().reset_index()
            .sort_values("NLE per kg", ascending=False)
        )
        fig_avg = px.bar(
            avg_nle, x="Vegetable", y="NLE per kg",
            color="NLE per kg", color_continuous_scale="Greens"
        )
        fig_avg.update_layout(
            coloraxis_showscale=False, margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_avg, use_container_width=True)

    st.divider()
    st.subheader("Your Prices vs Market Survey Prices")
    st.caption(
        "Compares what you actually earned per kg vs what the market survey recorded. "
        "If you are below market, you might be under-pricing or selling in the wrong unit."
    )

    your_prices = filtered_df.groupby("Vegetable")["NLE per kg"].mean().reset_index()
    your_prices.columns = ["Vegetable", "Your NLE/kg"]

    mkt_prices = market_prices_df[["Vegetable", "NLE per kg"]].copy()
    mkt_prices["NLE per kg"] = pd.to_numeric(mkt_prices["NLE per kg"], errors="coerce")
    mkt_avg = mkt_prices.groupby("Vegetable")["NLE per kg"].mean().reset_index()
    mkt_avg.columns = ["Vegetable", "Market Survey NLE/kg"]

    combined = pd.merge(your_prices, mkt_avg, on="Vegetable", how="inner")

    if not combined.empty:
        fig_compare = go.Figure()
        fig_compare.add_trace(go.Bar(
            name="Your Sales",
            x=combined["Vegetable"],
            y=combined["Your NLE/kg"].round(2),
            marker_color="#2e7d32"
        ))
        fig_compare.add_trace(go.Bar(
            name="Market Survey",
            x=combined["Vegetable"],
            y=combined["Market Survey NLE/kg"].round(2),
            marker_color="#a5d6a7"
        ))
        fig_compare.update_layout(
            barmode="group", yaxis_title="NLE per kg",
            legend=dict(orientation="h"),
            margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_compare, use_container_width=True)
    else:
        st.info("No overlapping crops between your sales and market survey yet.")


# ══════════════════════════════════════════════════════════
# TAB 3 — MARKET COMPARISON
# ══════════════════════════════════════════════════════════
with tab3:
    st.header("Market Comparison — Scenario Analysis")
    st.caption(
        "For each sale, this shows what you actually earned vs what you would have "
        "earned if you sold in a different market."
    )

    if filtered_scen.empty or "_error" in filtered_scen.columns:
        err = filtered_scen["_error"].iloc[0] if "_error" in filtered_scen.columns else "No data"
        st.error(f"Could not load scenarios data. Details: {err}")
    else:
        plot_df = filtered_scen[[
            "Date", "Vegetable", "Tormabum Net", "Bo Net", "Waterloo Net"
        ]].copy()
        plot_df.columns = ["Date", "Vegetable", "TormaBum Net", "Bo Net", "Waterloo Net"]
        plot_df = plot_df.dropna(subset=["Date"])

        def best_market(row):
            vals = {
                "TormaBum": row["TormaBum Net"],
                "Bo":       row["Bo Net"],
                "Waterloo": row["Waterloo Net"]
            }
            return max(vals, key=vals.get)

        plot_df["Best Market"] = plot_df.apply(best_market, axis=1)

        st.subheader("Which market gives the best return most often?")
        best_counts = plot_df["Best Market"].value_counts().reset_index()
        best_counts.columns = ["Market", "Times Best"]

        col1, col2 = st.columns([1, 2])
        with col1:
            st.dataframe(best_counts, hide_index=True, use_container_width=True)
        with col2:
            fig_best = px.pie(
                best_counts, values="Times Best", names="Market",
                color_discrete_sequence=["#2e7d32", "#66bb6a", "#a5d6a7"]
            )
            st.plotly_chart(fig_best, use_container_width=True)

        st.divider()

        st.subheader("Average Net Revenue by Market & Crop")
        st.caption(
            "Compare what you earned at your market vs what you would have earned elsewhere. "
            "Bars below zero mean you would have lost money after transport costs."
        )

        melt_df = plot_df.melt(
            id_vars=["Date", "Vegetable"],
            value_vars=["TormaBum Net", "Bo Net", "Waterloo Net"],
            var_name="Market", value_name="Net Revenue"
        )
        melt_df["Market"] = melt_df["Market"].str.replace(" Net", "", regex=False)
        veg_avg = melt_df.groupby(["Vegetable", "Market"])["Net Revenue"].mean().reset_index()

        fig_scenario = px.bar(
            veg_avg, x="Vegetable", y="Net Revenue", color="Market",
            barmode="group",
            color_discrete_map={
                "TormaBum": "#2e7d32",
                "Bo":       "#66bb6a",
                "Waterloo": "#a5d6a7"
            },
            labels={"Net Revenue": "Avg Net Revenue (NLE)"}
        )
        fig_scenario.add_hline(
            y=0, line_color="red", line_dash="dash",
            annotation_text="Break-even"
        )
        fig_scenario.update_layout(margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig_scenario, use_container_width=True)

        st.divider()
        st.subheader("Full Scenario Data")
        st.dataframe(
            plot_df.sort_values("Date", ascending=False),
            use_container_width=True, hide_index=True
        )


# ══════════════════════════════════════════════════════════
# TAB 4 — PRODUCTION TRACKER
# Reads from Production Tracker and Harvest Log sheets.
# Update those sheets in Excel and push to GitHub to refresh.
# ══════════════════════════════════════════════════════════
with tab4:
    st.header("Production Tracker")
    st.caption(
        "Managed via the **Production Tracker** and **Harvest Log** sheets "
        "in your Excel file. Update and push to GitHub to refresh."
    )

    if production_df.empty or "_error" in production_df.columns:
        err = production_df["_error"].iloc[0] if "_error" in production_df.columns else ""
        if err:
            st.error(f"Could not load Production Tracker sheet. Error: {err}")
        else:
            st.info("No production records found. Add rows to the Production Tracker sheet.")
    else:
        # ── Build cumulative harvest from Harvest Log ──
        # For each crop, sum all Kg Harvested entries
        if not harvest_df.empty:
            cumulative = (
                harvest_df.groupby("Crop")["Kg Harvested"]
                .sum().reset_index()
                .rename(columns={"Kg Harvested": "Actual Harvest (kg)"})
            )
            prod_display = production_df.merge(cumulative, on="Crop", how="left")
        else:
            prod_display = production_df.copy()
            prod_display["Actual Harvest (kg)"] = 0

        prod_display["Actual Harvest (kg)"] = prod_display["Actual Harvest (kg)"].fillna(0)
        prod_display["Harvest Progress (%)"] = (
            prod_display["Actual Harvest (kg)"] /
            prod_display["Expected Harvest (kg)"].replace(0, 1) * 100
        ).round(1)
        prod_display["Remaining (kg)"] = (
            prod_display["Expected Harvest (kg)"] - prod_display["Actual Harvest (kg)"]
        ).clip(lower=0)

        # ── SUMMARY CARDS ──
        active_statuses = ["Seedling", "Growing", "Flowering", "Producing"]
        active = prod_display[prod_display["Status"].isin(active_statuses)]

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Active Crops",       len(active))
        col2.metric("Total Rows",         int(prod_display["Rows"].sum()))
        col3.metric("Total Plants",       int(prod_display["Total Plants"].sum()))
        col4.metric("Expected Harvest",   f"{prod_display['Expected Harvest (kg)'].sum():.1f} kg")
        col5.metric("Actual Harvest",     f"{prod_display['Actual Harvest (kg)'].sum():.1f} kg")

        st.divider()

        # ── HARVEST PROGRESS PER CROP ──
        st.subheader("Harvest Progress — Actual vs Expected (kg)")
        st.caption("How much of the expected harvest has been collected so far.")

        fig_progress = go.Figure()
        for _, row in prod_display.iterrows():
            label = f"{row['Crop']} ({row['Location']})"
            fig_progress.add_trace(go.Bar(
                name="Actual",
                x=[label],
                y=[row["Actual Harvest (kg)"]],
                marker_color="#2e7d32",
                showlegend=_ == 0
            ))
            fig_progress.add_trace(go.Bar(
                name="Remaining",
                x=[label],
                y=[row["Remaining (kg)"]],
                marker_color="#c8e6c9",
                showlegend=_ == 0
            ))

        fig_progress.update_layout(
            barmode="stack",
            yaxis_title="kg",
            legend=dict(orientation="h"),
            margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_progress, use_container_width=True)

        st.divider()

        # ── BY LOCATION ──
        col_l, col_r = st.columns(2)

        with col_l:
            st.subheader("Crops by Location")
            loc_counts = prod_display.groupby("Location")["Crop"].count().reset_index()
            loc_counts.columns = ["Location", "Number of Crops"]
            fig_loc = px.pie(
                loc_counts, values="Number of Crops", names="Location",
                color_discrete_sequence=px.colors.sequential.Greens
            )
            fig_loc.update_layout(margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig_loc, use_container_width=True)

        with col_r:
            st.subheader("Expected Harvest by Location (kg)")
            loc_harvest = prod_display.groupby("Location").agg(
                Expected=("Expected Harvest (kg)", "sum"),
                Actual=("Actual Harvest (kg)", "sum")
            ).reset_index()
            fig_loc_h = go.Figure()
            fig_loc_h.add_trace(go.Bar(
                name="Expected", x=loc_harvest["Location"],
                y=loc_harvest["Expected"], marker_color="#a5d6a7"
            ))
            fig_loc_h.add_trace(go.Bar(
                name="Actual", x=loc_harvest["Location"],
                y=loc_harvest["Actual"], marker_color="#2e7d32"
            ))
            fig_loc_h.update_layout(
                barmode="group", yaxis_title="kg",
                legend=dict(orientation="h"),
                margin=dict(l=0, r=0, t=30, b=0)
            )
            st.plotly_chart(fig_loc_h, use_container_width=True)

        st.divider()

        # ── FULL PRODUCTION TABLE ──
        st.subheader("All Production Records")
        display_cols = [
            "Crop", "Location", "Date Planted", "Rows", "Total Plants",
            "Expected Harvest (kg)", "Actual Harvest (kg)", "Harvest Progress (%)",
            "First Harvest Date", "Last Harvest Date", "Status", "Notes"
        ]
        display_cols = [c for c in display_cols if c in prod_display.columns]

        # Format dates nicely
        show_df = prod_display[display_cols].copy()
        for dcol in ["Date Planted", "First Harvest Date", "Last Harvest Date"]:
            if dcol in show_df.columns:
                show_df[dcol] = pd.to_datetime(show_df[dcol]).dt.strftime("%d %b %Y").replace("NaT", "—")

        st.dataframe(show_df, use_container_width=True, hide_index=True)

        st.divider()

        # ── HARVEST LOG ──
        st.subheader("Harvest Log")
        if harvest_df.empty:
            st.info(
                "No harvest records yet. Add rows to the **Harvest Log** sheet "
                "in your Excel file as you harvest."
            )
        else:
            # Summary per crop
            harvest_summary = harvest_df.groupby("Crop").agg(
                Total_Harvested=("Kg Harvested", "sum"),
                First_Date=("Date", "min"),
                Last_Date=("Date", "max"),
                Num_Harvests=("Date", "count")
            ).reset_index()
            harvest_summary.columns = [
                "Crop", "Total Harvested (kg)",
                "First Harvest", "Last Harvest", "# Harvests"
            ]

            col_l2, col_r2 = st.columns(2)
            with col_l2:
                st.markdown("**Cumulative Harvest per Crop**")
                st.dataframe(harvest_summary, use_container_width=True, hide_index=True)

            with col_r2:
                fig_harv = px.bar(
                    harvest_summary.sort_values("Total Harvested (kg)", ascending=False),
                    x="Crop", y="Total Harvested (kg)",
                    color="Total Harvested (kg)",
                    color_continuous_scale="Greens"
                )
                fig_harv.update_layout(
                    coloraxis_showscale=False,
                    margin=dict(l=0, r=0, t=30, b=0)
                )
                st.plotly_chart(fig_harv, use_container_width=True)

            # Harvest over time
            st.markdown("**Harvest Over Time**")
            fig_time = px.line(
                harvest_df.sort_values("Date"),
                x="Date", y="Kg Harvested", color="Crop",
                markers=True
            )
            fig_time.update_xaxes(
                range=[harvest_df["Date"].min(), harvest_df["Date"].max()]
            )
            fig_time.update_layout(margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig_time, use_container_width=True)

            # Full log
            st.markdown("**Full Harvest Log**")
            st.dataframe(
                harvest_df.sort_values("Date", ascending=False),
                use_container_width=True, hide_index=True
            )

        # Download buttons
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            prod_csv = prod_display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Download Production Data",
                data=prod_csv,
                file_name="production_tracker.csv",
                mime="text/csv"
            )
        with col_d2:
            if not harvest_df.empty:
                harv_csv = harvest_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Download Harvest Log",
                    data=harv_csv,
                    file_name="harvest_log.csv",
                    mime="text/csv"
                )


# ── FOOTER ──
st.divider()
st.caption("FOOP Greenhouse Dashboard · Built with Streamlit")
