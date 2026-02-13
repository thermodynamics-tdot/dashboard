import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"  # must exist in the repo root

# IMPORTANT:
# Your file does NOT have a column named STATUS.
# From your screenshot, status-like values (COMPLETED / ATTENDED / NOT ATTENDED / NEED MATERIAL) are in:
# "TECHNICIAN FEEDBACK"
# If your status is in a different column, change STATUS_COL below.

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "TECHNICIAN FEEDBACK"   # change if needed
TECH_COL = "TECH  1"                 # NOTE: two spaces between TECH and 1
CALL_ID_COL = "TD REPORT NO."


@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [str(c).strip().upper() for c in df.columns]  # normalize
    return df


df = load_data()

# Validate required columns
required = [DATE_COL, CUSTOMER_COL, STATUS_COL, TECH_COL, CALL_ID_COL]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", df.columns.tolist())
    st.stop()

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

# Clean status text
df[STATUS_COL] = (
    df[STATUS_COL]
    .fillna("BLANK")
    .astype(str)
    .str.strip()
    .str.upper()
)

# Sidebar filters
st.sidebar.header("Filters")

min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()
d1, d2 = st.sidebar.date_input("Date range", (min_d, max_d))

customers = ["(ALL)"] + sorted(df[CUSTOMER_COL].dropna().astype(str).unique().tolist())
sel_customer = st.sidebar.selectbox("Customer", customers)

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(ALL)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)

dff = df.loc[mask].copy()

# KPIs
k1, k2, k3 = st.columns(3)

total_calls = dff[CALL_ID_COL].nunique()
completed = int((dff[STATUS_COL] == "COMPLETED").sum())
not_attended = int((dff[STATUS_COL] == "NOT ATTENDED").sum())

k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", completed)
k3.metric("Not Attended", not_attended)

# Charts
c1, c2 = st.columns(2)

# 1) Pie - Status
with c1:
    st.subheader("Call Status")
    status_counts = dff[STATUS_COL].value_counts().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]
    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT")
    st.plotly_chart(fig_pie, use_container_width=True)

# 2) Stacked monthly trend
with c2:
    st.subheader("Monthly Trend (Stacked)")
    dff["MONTH"] = dff[DATE_COL].dt.to_period("M").astype(str)

    month_status = (
        dff.groupby(["MONTH", STATUS_COL])
        .size()
        .reset_index(name="COUNT")
    )

    fig_stack = px.bar(
        month_status,
        x="MONTH",
        y="COUNT",
        color=STATUS_COL,
        barmode="stack"
    )
    st.plotly_chart(fig_stack, use_container_width=True)

# 3) Technician performance
st.subheader("Technician Performance")

dff2 = dff.dropna(subset=[TECH_COL, STATUS_COL]).copy()
dff2[TECH_COL] = dff2[TECH_COL].astype(str).str.strip()

tech_status = (
    dff2.groupby([TECH_COL, STATUS_COL])
        .size()
        .reset_index(name="COUNT")
)

fig_tech = px.bar(
    tech_status,
    y=TECH_COL,
    x="COUNT",
    color=STATUS_COL,
    orientation="h",
    barmode="group"
)
st.plotly_chart(fig_tech, use_container_width=True)

with st.expander("Show filtered data"):
    st.dataframe(dff, use_container_width=True)
