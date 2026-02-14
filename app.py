import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- Page --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

# -------------------- Config --------------------
FILE_NAME = "CALL RECORDS 2026.xlsx"  # in repo root

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "CALL STATUS"           # will be auto-mapped if different
TECH_COL = "TECH 1"
CALL_ID_COL = "TD REPORT NO."

# -------------------- Helpers --------------------
@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [" ".join(str(c).strip().upper().split()) for c in df.columns]
    return df

def ensure_col(df, want):
    """Return actual column name in df that matches want or known aliases."""
    want_n = " ".join(want.strip().upper().split())
    if want_n in df.columns:
        return want_n

    aliases = {
        "CALL STATUS": ["STATUS", "CALLSTATUS", "CALL_STATUS", "CALL  STATUS", "CALL-STATUS"],
        "TECH 1": ["TECH1", "TECH  1", "TECHNICIAN", "TECH"],
        "TD REPORT NO.": ["TD REPORT NO", "TD_REPORT_NO", "REPORT NO", "REPORT NO."],
        "DATE": ["CALL DATE", "SERVICE DATE"],
        "CUSTOMER": ["CLIENT", "CUSTOMER NAME"],
    }

    for a in aliases.get(want_n, []):
        a_n = " ".join(a.strip().upper().split())
        if a_n in df.columns:
            return a_n

    return None

def status_palette():
    return {
        "COMPLETED": "#0068C9",
        "ATTENDED": "#83C9FF",
        "NOT ATTENDED": "#FF2B2B",
    }

def normalize_status(s):
    s = str(s).strip().upper()
    if s in ["NAN", "", "NONE", "NULL", "(BLANK)"]:
        return "BLANK"
    return s

# -------------------- Load --------------------
df = load_data()

# Resolve real column names (avoid KeyError)
DATE_COL_REAL = ensure_col(df, DATE_COL)
CUSTOMER_COL_REAL = ensure_col(df, CUSTOMER_COL)
STATUS_COL_REAL = ensure_col(df, STATUS_COL)
TECH_COL_REAL = ensure_col(df, TECH_COL)
CALL_ID_COL_REAL = ensure_col(df, CALL_ID_COL)  # can be None if not present

missing = [name for name, real in [
    ("DATE", DATE_COL_REAL),
    ("CUSTOMER", CUSTOMER_COL_REAL),
    ("CALL STATUS", STATUS_COL_REAL),
] if real is None]

if missing:
    st.error(f"Missing required columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", list(df.columns))
    st.stop()

# Use resolved names from here on
DATE_COL = DATE_COL_REAL
CUSTOMER_COL = CUSTOMER_COL_REAL
STATUS_COL = STATUS_COL_REAL
TECH_COL = TECH_COL_REAL
CALL_ID_COL = CALL_ID_COL_REAL  # may be None

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

# Normalize status (DO NOT remove BLANK rows from dataset)
df[STATUS_COL] = df[STATUS_COL].apply(normalize_status)

# -------------------- Sidebar (wider) --------------------
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
    sel_customer = st.selectbox("Customer", customers)

    min_d = df[DATE_COL].min().date()
    max_d = df[DATE_COL].max().date()
    d1, d2 = st.date_input("Date range", (min_d, max_d))

# Filter
mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)
dff = df.loc[mask].copy()

# -------------------- Title + KPIs --------------------
st.title("Service Calls Dashboard")

k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns and CALL_ID_COL else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# -------------------- Charts --------------------
left, right = st.columns(2, gap="large")

colors = status_palette()
status_order = ["COMPLETED", "ATTENDED", "NOT ATTENDED"]

# ---------- PIE ----------
with left:
    st.subheader("Call Status")

    # hide BLANK only for chart
    chart_df = dff[dff[STATUS_COL].isin(status_order)].copy()

    status_counts = (
        chart_df[STATUS_COL]
        .value_counts()
        .reindex(status_order)
        .dropna()
        .reset_index()
    )
    status_counts.columns = ["STATUS", "COUNT"]

    fig_pie = px.pie(
        status_counts,
        names="STATUS",
        values="COUNT",
        color="STATUS",
        color_discrete_map=colors,
        category_orders={"STATUS": status_order},
    )
    fig_pie.update_layout(
        legend_title_text="",
        legend=dict(font=dict(size=14)),
        margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ---------- CUSTOMER TREND ----------
with right:
    customer_title = "All Customers" if sel_customer == "(All)" else sel_customer
    st.subheader(customer_title)

    full_range = (d1 == min_d and d2 == max_d)
    default_mode = "Total" if full_range else "Month"

    chart_col, side_col = st.columns([3.3, 1.2], gap="large")

    # View should be BELOW legend area — so we use Plotly legend and then place View under it
    with side_col:
        st.markdown("**View**")
        trend_mode = st.selectbox(
            "",
            ["Total", "Day", "Week", "Month"],
            index=["Total", "Day", "Week", "Month"].index(default_mode),
            label_visibility="collapsed",
        )

    tmp = dff.copy()
    tmp = tmp[tmp[STATUS_COL].isin(status_order)].copy()  # hide BLANK only in chart

    if trend_mode == "Total":
        tmp["PERIOD"] = "All"
        xorder = ["All"]
        tickvals, ticktext = xorder, xorder

    elif trend_mode == "Day":
        tmp["PERIOD"] = tmp[DATE_COL].dt.date.astype(str)
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals, ticktext = xorder, xorder

    elif trend_mode == "Week":
        iso = tmp[DATE_COL].dt.isocalendar()
        tmp["PERIOD"] = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals, ticktext = xorder, xorder

    else:  # Month
        tmp["PERIOD"] = tmp[DATE_COL].dt.to_period("M").astype(str)
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals = xorder
        ticktext = [pd.to_datetime(p + "-01").strftime("%b") for p in xorder]  # Jan/Feb

    grp = tmp.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    fig_stack = px.bar(
        grp,
        x="PERIOD",
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
        color_discrete_map=colors,
        category_orders={"PERIOD": xorder, STATUS_COL: status_order},
    )

    fig_stack.update_layout(
        legend_title_text="",
        legend=dict(font=dict(size=14)),  # match pie
        margin=dict(l=10, r=10, t=10, b=10),
    )

    fig_stack.update_xaxes(
        categoryorder="array",
        categoryarray=xorder,
        tickmode="array",
        tickvals=tickvals,
        ticktext=ticktext,
        title_text="",
    )
    fig_stack.update_yaxes(title_text="Count")

    with chart_col:
        st.plotly_chart(fig_stack, use_container_width=True)

# -------------------- Technician Performance (stacked ROW) --------------------
st.markdown("## Technician Performance (Sorted by Completion Rate)")

if TECH_COL and TECH_COL in dff.columns:
    tech_df = dff.copy()
    tech_df = tech_df[tech_df[STATUS_COL].isin(status_order)].copy()  # hide BLANK only in chart

    tech_counts = tech_df.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")

    totals = tech_counts.groupby(TECH_COL)["COUNT"].sum().rename("TOTAL")
    completed = (
        tech_counts[tech_counts[STATUS_COL] == "COMPLETED"]
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
        color_discrete_map=colors,
        category_orders={TECH_COL: order, STATUS_COL: status_order},
    )
    fig_tech.update_layout(
        legend_title_text="",
        legend=dict(font=dict(size=14)),
        margin=dict(l=10, r=10, t=10, b=10),
        yaxis_title="",
        xaxis_title="Count",
    )
    st.plotly_chart(fig_tech, use_container_width=True)

# -------------------- Show Data --------------------
with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
