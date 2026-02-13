import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"  # must be in the same folder as app.py in the repo


@st.cache_data(ttl=300)  # refresh every 5 minutes
def load_data():
    df = pd.read_excel(FILE_NAME)
    # normalize headers
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def pick_col(cols, candidates):
    cols_set = set(cols)
    for c in candidates:
        if c in cols_set:
            return c
    return None


df = load_data()
cols = df.columns.tolist()

# Auto-detect columns (robust against header variations)
DATE_COL = pick_col(cols, ["DATE", "CALL DATE", "DATE."])
CUSTOMER_COL = pick_col(cols, ["CUSTOMER", "CLIENT", "ACCOUNT"])
STATUS_COL = pick_col(cols, ["STATUS", "CALL STATUS", "JOB STATUS", "CALLSTATUS"])
TECH_COL = pick_col(cols, ["TECH 1", "TECH1", "TECHNICIAN", "TECH", "TECH NAME", "TECHNICIANS"])
CALL_ID_COL = pick_col(cols, ["TD REPORT NO.", "TD REPORT NO", "REPORT NO.", "REPORT NO", "TD NO", "TICKET NO", "TICKET"])

missing = [name for name, val in [
    ("DATE", DATE_COL),
    ("CUSTOMER", CUSTOMER_COL),
    ("STATUS", STATUS_COL),
    ("TECH", TECH_COL),
    ("TD REPORT NO.", CALL_ID_COL),
] if val is None]

if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns in your file:")
    st.write(cols)
    st.stop()

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

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
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL].astype(str).str.upper() == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL].astype(str).str.upper() == "NOT ATTENDED").sum()))

# Charts row
c1, c2 = st.columns(2)

# 1) Pie - Status distribution
with c1:
    st.subheader("Call Status")
    status_counts = (
        dff[STATUS_COL]
        .fillna("BLANK")
        .astype(str)
        .str.strip()
        .str.upper()
        .value_counts()
        .reset_index()
    )
    status_counts.columns = ["STATUS", "COUNT"]
    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT")
    st.plotly_chart(fig_pie, use_container_width=True)

# 2) Stacked monthly trend
with c2:
    st.subheader("Monthly Trend (Stacked)")
    dff["MONTH"] = dff[DATE_COL].dt.to_period("M").astype(str)

    month_status = (
        dff.assign(_STATUS=dff[STATUS_COL].fillna("BLANK").astype(str).str.strip().str.upper())
          .groupby(["MONTH", "_STATUS"])
          .size()
          .reset_index(name="COUNT")
          .rename(columns={"_STATUS": STATUS_COL})
    )

    fig_stack = px.bar(
        month_status,
        x="MONTH",
        y="COUNT",
        color=STATUS_COL,
        barmode="stack"
    )
    st.plotly_chart(fig_stack, use_container_width=True)

# 3) Technician performance (safe)
st.subheader("Technician Performance")
dff2 = dff.dropna(subset=[TECH_COL, STATUS_COL]).copy()
dff2[TECH_COL] = dff2[TECH_COL].astype(str).str.strip()
dff2[STATUS_COL] = dff2[STATUS_COL].astype(str).str.strip().str.upper()

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

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
