import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"   # must match the file name in your repo folder

@st.cache_data(ttl=300)  # refresh every 5 min
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [c.strip().upper() for c in df.columns]
    return df

df = load_data()

# ---- update these column names if your file differs ----
DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "STATUS"
TECH_COL = "TECH 1"          # in your file it looks like TECH 1
CALL_ID_COL = "TD REPORT NO."  # adjust if needed
# -------------------------------------------------------

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

# Filters
st.sidebar.header("Filters")

min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()

d1, d2 = st.sidebar.date_input("Date range", (min_d, max_d))

customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().unique().tolist())
sel_customer = st.sidebar.selectbox("Customer", customers)

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL] == sel_customer)

dff = df.loc[mask].copy()

# KPIs
k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

c1, c2 = st.columns(2)

# Pie chart - status
with c1:
    st.subheader("Call Status")
    status_counts = dff[STATUS_COL].fillna("BLANK").value_counts().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]
    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT")
    st.plotly_chart(fig_pie, use_container_width=True)

# Stacked monthly trend
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

# Technician performance
st.subheader("Technician Performance")
tech_status = (
    dff.groupby([TECH_COL, STATUS_COL])
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

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
