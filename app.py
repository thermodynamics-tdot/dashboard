import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"   # must match the file in your repo

@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    # normalize column names safely (strip + collapse spaces + uppercase)
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

# Auto-detect tech column (your file shows TECH 1 / TECH 2)
TECH_COL = pick_first_existing(
    df.columns,
    ["TECH 1", "TECH 2", "TECH", "TECHNICIAN", "TECHNICIAN NAME", "ENGINEER"]
)

# Validate required cols
missing = [c for c in [DATE_COL, CUSTOMER_COL, STATUS_COL] if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", list(df.columns))
    st.stop()

# Parse dates
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

# Filters
st.sidebar.header("Filters")
min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()
d1, d2 = st.sidebar.date_input("Date range", (min_d, max_d), key="date_range")

customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().astype(str).unique().tolist())
sel_customer = st.sidebar.selectbox("Customer", customers, key="customer")

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)

dff = df.loc[mask].copy()

# Normalize status values
dff[STATUS_COL] = dff[STATUS_COL].astype(str).str.strip().str.upper()
dff[STATUS_COL] = dff[STATUS_COL].replace({"NAN": "BLANK", "NONE": "BLANK"}).fillna("BLANK")

# KPIs
k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

c1, c2 = st.columns(2)

# Pie chart
with c1:
    st.subheader("Call Status")
    status_counts = dff[STATUS_COL].value_counts().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]
    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT")
    st.plotly_chart(fig_pie, use_container_width=True)

# Monthly stacked trend
with c2:
    st.subheader("Monthly Trend (Stacked)")
    dff["MONTH"] = dff[DATE_COL].dt.to_period("M").astype(str)
    month_status = dff.groupby(["MONTH", STATUS_COL]).size().reset_index(name="COUNT")

    fig_stack = px.bar(
        month_status,
        x="MONTH",
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
    )
    st.plotly_chart(fig_stack, use_container_width=True)

# Technician chart (stacked column, sorted by completion rate)
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

    fig_tech.update_layout(
        xaxis_title="Technician",
        yaxis_title="Calls",
        legend_title="Status",
    )

    st.plotly_chart(fig_tech, use_container_width=True)

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
