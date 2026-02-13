import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"   # must match the file name in your repo

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
TECH_COL = "TECH 1"
CALL_ID_COL = "TD REPORT NO."
# -------------------------------------------------------

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

# Filters (left sidebar)
st.sidebar.header("Filters")

min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()

d1, d2 = st.sidebar.date_input("Date range", (min_d, max_d), key="date_range")

customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().unique().tolist())
sel_customer = st.sidebar.selectbox("Customer", customers, key="customer")

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL] == sel_customer)

dff = df.loc[mask].copy()

# Normalize Status + Tech
dff[STATUS_COL] = dff[STATUS_COL].astype(str).str.strip().str.upper()
dff[TECH_COL] = dff[TECH_COL].astype(str).str.strip()
dff.loc[dff[TECH_COL].isin(["", "NAN", "NONE"]), TECH_COL] = "UNASSIGNED"

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
    status_counts = dff[STATUS_COL].replace({"NAN": "BLANK", "NONE": "BLANK"}).fillna("BLANK").value_counts().reset_index()
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

# -------------------- FIXED TECHNICIAN PART --------------------
st.subheader("Technician Performance (Sorted by Completion Rate)")

# counts per tech + status
tech_status = (
    dff.groupby([TECH_COL, STATUS_COL])
    .size()
    .reset_index(name="COUNT")
)

# pivot to compute completion rate
pivot = tech_status.pivot_table(index=TECH_COL, columns=STATUS_COL, values="COUNT", aggfunc="sum", fill_value=0)

# ensure expected status columns exist (avoid KeyError)
for col in ["COMPLETED", "ATTENDED", "NOT ATTENDED"]:
    if col not in pivot.columns:
        pivot[col] = 0

pivot["TOTAL"] = pivot.sum(axis=1)
pivot["COMPLETION_RATE"] = (pivot["COMPLETED"] / pivot["TOTAL"]).where(pivot["TOTAL"] > 0, 0)

# sort by completion rate (high to low), then by total calls
pivot = pivot.sort_values(["COMPLETION_RATE", "TOTAL"], ascending=[False, False]).reset_index()

# apply sorted tech order back to chart data
tech_order = pivot[TECH_COL].tolist()
tech_status[TECH_COL] = pd.Categorical(tech_status[TECH_COL], categories=tech_order, ordered=True)

# keep chart legend in a nice consistent order
status_order = ["ATTENDED", "COMPLETED", "NOT ATTENDED"]
present_status = [s for s in status_order if s in tech_status[STATUS_COL].unique().tolist()]
# include any extra statuses at the end
extras = [s for s in tech_status[STATUS_COL].unique().tolist() if s not in present_status]
final_status_order = present_status + extras

tech_status[STATUS_COL] = pd.Categorical(tech_status[STATUS_COL], categories=final_status_order, ordered=True)

# stacked COLUMN chart (vertical), tech on x
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
# ---------------------------------------------------------------

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
