import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import timedelta

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"   # must match the file in your repo

@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = (
        pd.Index(df.columns)
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.upper()
    )
    return df

df = load_data()

# ---- expected columns (normalized) ----
DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "STATUS"
CALL_ID_COL = "TD REPORT NO."
# --------------------------------------

def pick_first_existing(cols, candidates):
    for c in candidates:
        if c in cols:
            return c
    return None

TECH_COL = pick_first_existing(
    df.columns,
    ["TECH 1", "TECH 2", "TECH", "TECHNICIAN", "TECHNICIAN NAME", "ENGINEER"]
)

missing = [c for c in [DATE_COL, CUSTOMER_COL, STATUS_COL] if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", list(df.columns))
    st.stop()

# Parse dates
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

# Normalize status values
df[STATUS_COL] = df[STATUS_COL].astype(str).str.strip().str.upper()
df[STATUS_COL] = df[STATUS_COL].replace({"NAN": "BLANK", "NONE": "BLANK"}).fillna("BLANK")

# -------------------- SIDEBAR FILTERS --------------------
st.sidebar.header("Filters")

# Customer
customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().astype(str).unique().tolist())
sel_customer = st.sidebar.selectbox("Customer", customers, key="customer")

# Date filtering logic
min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()

# Default = All data ON
all_data = st.sidebar.toggle("All data (default)", value=True, help="When ON, trend stays Monthly across full data.")

quick = st.sidebar.selectbox(
    "Quick range (optional)",
    ["None", "Past Week", "Past Month", "Past 3 Months", "Past 6 Months", "Past Year"],
    index=0,
    disabled=all_data,
    key="quick_range",
)

# Base range shown in picker:
default_range = (min_d, max_d)

def clamp(d):
    if d < min_d:
        return min_d
    if d > max_d:
        return max_d
    return d

# Compute quick range if chosen
if not all_data and quick != "None":
    end = max_d
    if quick == "Past Week":
        start = end - timedelta(days=7)
    elif quick == "Past Month":
        start = end - timedelta(days=30)
    elif quick == "Past 3 Months":
        start = end - timedelta(days=90)
    elif quick == "Past 6 Months":
        start = end - timedelta(days=180)
    else:  # Past Year
        start = end - timedelta(days=365)
    default_range = (clamp(start), end)

d1, d2 = st.sidebar.date_input(
    "Date range",
    value=default_range,
    min_value=min_d,
    max_value=max_d,
    disabled=all_data,
    key="date_range",
)

# Only show granularity when user is actually filtering dates
if all_data:
    group_by = "Month"  # always monthly when All data
else:
    group_by = st.sidebar.selectbox("Trend granularity", ["Day", "Week", "Month"], index=2, key="group_by")

# -------------------- APPLY FILTERS --------------------
if all_data:
    mask = pd.Series(True, index=df.index)
else:
    mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)

if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)

dff = df.loc[mask].copy()

# -------------------- KPIs --------------------
k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# -------------------- TOP CHARTS --------------------
c1, c2 = st.columns(2)

with c1:
    st.subheader("Call Status")
    status_counts = dff[STATUS_COL].value_counts().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]
    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT")
    st.plotly_chart(fig_pie, use_container_width=True)

with c2:
    st.subheader(f"{'Monthly' if group_by=='Month' else group_by} Trend (Stacked)")

    if group_by == "Day":
        dff["PERIOD"] = dff[DATE_COL].dt.date.astype(str)
    elif group_by == "Week":
        # ISO year-week label
        iso = dff[DATE_COL].dt.isocalendar()
        dff["PERIOD"] = (iso["YEAR"].astype(str) + "-W" + iso["WEEK"].astype(str).str.zfill(2))
    else:
        dff["PERIOD"] = dff[DATE_COL].dt.to_period("M").astype(str)

    period_status = dff.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    fig_stack = px.bar(
        period_status,
        x="PERIOD",
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
    )
    fig_stack.update_layout(xaxis_title="")
    st.plotly_chart(fig_stack, use_container_width=True)

# -------------------- TECHNICIAN (STACKED COLUMN) --------------------
st.subheader("Technician Performance (Sorted by Completion Rate)")

if TECH_COL is None:
    st.warning("No technician column found (TECH 1 / TECH 2). Technician chart is hidden.")
else:
    dff[TECH_COL] = (
        dff[TECH_COL]
        .astype(str)
        .str.strip()
        .replace({"NAN": "UNASSIGNED", "NONE": "UNASSIGNED", "": "UNASSIGNED"})
        .fillna("UNASSIGNED")
    )

    tech_status = dff.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")

    pivot = tech_status.pivot_table(index=TECH_COL, columns=STATUS_COL, values="COUNT", aggfunc="sum", fill_value=0)
    for col in ["COMPLETED", "ATTENDED", "NOT ATTENDED"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot["TOTAL"] = pivot.sum(axis=1)
    pivot["COMPLETION_RATE"] = (pivot["COMPLETED"] / pivot["TOTAL"]).where(pivot["TOTAL"] > 0, 0)
    pivot = pivot.sort_values(["COMPLETION_RATE", "TOTAL"], ascending=[False, False]).reset_index()

    tech_order = pivot[TECH_COL].tolist()
    tech_status[TECH_COL] = pd.Categorical(tech_status[TECH_COL], categories=tech_order, ordered=True)

    status_order = ["ATTENDED", "COMPLETED", "NOT ATTENDED", "BLANK"]
    present = [s for s in status_order if s in tech_status[STATUS_COL].unique().tolist()]
    extras = [s for s in tech_status[STATUS_COL].unique().tolist() if s not in present]
    final_status_order = present + extras
    tech_status[STATUS_COL] = pd.Categorical(tech_status[STATUS_COL], categories=final_status_order, ordered=True)

    fig_tech = px.bar(
        tech_status.sort_values(TECH_COL),
        x=TECH_COL,
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
        category_orders={TECH_COL: tech_order, STATUS_COL: final_status_order},
    )
    fig_tech.update_layout(xaxis_title="Technician", yaxis_title="Calls", legend_title="Status")
    st.plotly_chart(fig_tech, use_container_width=True)

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
