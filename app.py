import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- Page --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

# -------------------- Config --------------------
FILE_NAME = "CALL RECORDS 2026.xlsx"  # must exist in repo root (same folder as app.py)

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "CALL STATUS"           # <-- your sheet uses Call Status (from screenshots)
TECH_COL = "TECH 1"                  # available: "TECH  1" or "TECH 1" (we normalize)
CALL_ID_COL = "TD REPORT NO."        # adjust if needed

# -------------------- Helpers --------------------
@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    # normalize column names
    df.columns = [" ".join(str(c).strip().upper().split()) for c in df.columns]
    return df

def ensure_col(df, want):
    """Try a few normalized variants to find the column."""
    want_n = " ".join(want.strip().upper().split())
    if want_n in df.columns:
        return want_n
    # common fallbacks
    aliases = {
        "CALL STATUS": ["STATUS", "CALLSTATUS", "CALL_STATUS"],
        "TECH 1": ["TECH1", "TECH  1", "TECHNICIAN", "TECH"],
        "TD REPORT NO.": ["TD REPORT NO", "TD_REPORT_NO", "REPORT NO", "REPORT NO."],
    }
    for a in aliases.get(want_n, []):
        a_n = " ".join(a.strip().upper().split())
        if a_n in df.columns:
            return a_n
    return None

def fmt_period_label(dt, mode):
    if mode == "Month":
        return dt.strftime("%b")  # Jan, Feb
    if mode == "Week":
        y, w, _ = dt.isocalendar()
        return f"{y}-W{int(w):02d}"
    return dt.strftime("%Y-%m-%d")

# -------------------- Load --------------------
df = load_data()

# resolve columns after normalization
DATE_COL = ensure_col(df, DATE_COL) or DATE_COL
CUSTOMER_COL = ensure_col(df, CUSTOMER_COL) or CUSTOMER_COL
STATUS_COL = ensure_col(df, STATUS_COL) or STATUS_COL
TECH_COL = ensure_col(df, TECH_COL) or TECH_COL
CALL_ID_COL = ensure_col(df, CALL_ID_COL) or CALL_ID_COL

missing = [c for c in [DATE_COL, CUSTOMER_COL, STATUS_COL] if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", list(df.columns))
    st.stop()

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

# Clean status
df[STATUS_COL] = df[STATUS_COL].astype(str).str.strip().str.upper().replace({"NAN": "BLANK", "": "BLANK"})
df.loc[df[STATUS_COL].isin(["NONE", "NULL"]), STATUS_COL] = "BLANK"

# -------------------- Layout: Sidebar + Main --------------------
# Make sidebar a bit wider
st.markdown(
    """
    <style>
      section[data-testid="stSidebar"] { width: 340px !important; }
      section[data-testid="stSidebar"] > div { width: 340px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("## Filters")

    customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().astype(str).unique().tolist())
    sel_customer = st.selectbox("Customer", customers, index=0, key="customer_filter")

    min_d = df[DATE_COL].min().date()
    max_d = df[DATE_COL].max().date()

    d1, d2 = st.date_input("Date range", (min_d, max_d), key="date_range")

# Filter dataframe
mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)

dff = df.loc[mask].copy()

# -------------------- Title + KPIs --------------------
st.title("Service Calls Dashboard")

k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# -------------------- TOP: Pie + Customer Trend --------------------
left, right = st.columns(2, gap="large")

# ---- Pie
with left:
    st.subheader("Call Status")
    status_counts = (
        dff[STATUS_COL]
        .fillna("BLANK")
        .replace({"NAN": "BLANK", "": "BLANK"})
        .value_counts()
        .reset_index()
    )
    status_counts.columns = ["STATUS", "COUNT"]

    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT", hole=0)
    fig_pie.update_layout(
        legend_title_text="",
        margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ---- Customer trend (stacked) with "Above the legend" dropdown
with right:
    customer_title = "Customer Trend (Stacked)"
    if sel_customer != "(All)":
        customer_title = f"{sel_customer} - Trend (Stacked)"
    st.subheader(customer_title)

    # Decide default view:
    # If full range selected => show Total (single bar), else Month
    full_range = (d1 == min_d and d2 == max_d)
    default_mode = "Total" if full_range else "Month"

    # Right side panel for dropdown (above legend), chart on left
    chart_col, legend_col = st.columns([4.2, 1.2], gap="large")

    with legend_col:
        st.markdown("**View**")
        trend_mode = st.selectbox(
            "",
            ["Total", "Day", "Week", "Month"],
            index=["Total", "Day", "Week", "Month"].index(default_mode),
            key="trend_mode",
            label_visibility="collapsed",
        )
        st.markdown(" ")  # little spacing so it sits nicely above the chart legend area

    # Build grouped data
    tmp = dff.copy()
    if trend_mode == "Total":
        tmp["PERIOD"] = "All"
        xcol = "PERIOD"
        xorder = ["All"]
    elif trend_mode == "Day":
        tmp["PERIOD"] = tmp[DATE_COL].dt.date.astype(str)
        xcol = "PERIOD"
        xorder = sorted(tmp["PERIOD"].unique())
    elif trend_mode == "Week":
        # Use dt.isocalendar safely (pandas returns DataFrame with year/week/day)
        iso = tmp[DATE_COL].dt.isocalendar()
        tmp["PERIOD"] = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
        xcol = "PERIOD"
        xorder = sorted(tmp["PERIOD"].unique())
    else:  # Month
        # Keep year-month as period, but label as Jan/Feb on axis
        tmp["PERIOD_DT"] = tmp[DATE_COL].dt.to_period("M").dt.to_timestamp()
        tmp["PERIOD"] = tmp["PERIOD_DT"].dt.strftime("%Y-%m")
        xcol = "PERIOD"
        xorder = sorted(tmp["PERIOD"].unique())

    grp = tmp.groupby([xcol, STATUS_COL]).size().reset_index(name="COUNT")

    fig_stack = px.bar(grp, x=xcol, y="COUNT", color=STATUS_COL, barmode="stack")
    fig_stack.update_layout(
        legend_title_text="",
        margin=dict(l=10, r=10, t=10, b=10),
    )

    # Pretty x labels for Month (Jan/Feb instead of Jan 1 / Feb 1)
    if trend_mode == "Month":
        # Map YYYY-MM -> "Jan", but keep order
        mapping = {}
        for p in xorder:
            # p = "YYYY-MM"
            dt = pd.to_datetime(p + "-01")
            mapping[p] = dt.strftime("%b")
        fig_stack.update_xaxes(
            categoryorder="array",
            categoryarray=xorder,
            tickmode="array",
            tickvals=xorder,
            ticktext=[mapping[p] for p in xorder],
            title_text="",
        )
    else:
        fig_stack.update_xaxes(
            categoryorder="array",
            categoryarray=xorder,
            title_text="",
        )

    with chart_col:
        st.plotly_chart(fig_stack, use_container_width=True)

# -------------------- Technician Performance (stacked row) --------------------
st.markdown("## Technician Performance (Sorted by Completion Rate)")

if TECH_COL in dff.columns:
    tech_df = dff.copy()
    tech_df[TECH_COL] = tech_df[TECH_COL].astype(str).str.strip()
    tech_df.loc[tech_df[TECH_COL].isin(["", "NAN", "NONE", "NULL"]), TECH_COL] = "BLANK"

    tech_counts = tech_df.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")

    # completion rate = completed / total per tech
    totals = tech_counts.groupby(TECH_COL)["COUNT"].sum().rename("TOTAL")
    completed = (
        tech_counts.loc[tech_counts[STATUS_COL] == "COMPLETED"]
        .groupby(TECH_COL)["COUNT"]
        .sum()
        .rename("COMPLETED")
    )
    rates = pd.concat([totals, completed], axis=1).fillna(0)
    rates["COMP_RATE"] = rates["COMPLETED"] / rates["TOTAL"].where(rates["TOTAL"] != 0, 1)

    order = rates.sort_values("COMP_RATE", ascending=False).index.tolist()

    fig_tech = px.bar(
        tech_counts,
        y=TECH_COL,
        x="COUNT",
        color=STATUS_COL,
        barmode="stack",
        orientation="h",
        category_orders={TECH_COL: order},
    )
    fig_tech.update_layout(
        legend_title_text="",
        margin=dict(l=10, r=10, t=10, b=10),
        yaxis_title="",
        xaxis_title="Count",
    )
    st.plotly_chart(fig_tech, use_container_width=True)
else:
    st.warning(f"Technician column not found ({TECH_COL}). Available columns: {list(dff.columns)}")

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
