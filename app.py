import streamlit as st
import pandas as pd

st.set_page_config(page_title="Service Dashboard", layout="wide")

EXCEL_URL = st.secrets.get("EXCEL_URL", "")

@st.cache_data(ttl=300)  # refresh every 5 minutes
def load_data(url: str) -> pd.DataFrame:
    if not url:
        return pd.DataFrame()
    df = pd.read_excel(url)
    # normalize columns (adjust names to yours)
    df.columns = [c.strip().upper() for c in df.columns]
    return df

st.title("Service Calls Dashboard")

df = load_data(EXCEL_URL)

if df.empty:
    st.info("Set EXCEL_URL in Streamlit secrets to a direct-download .xlsx link.")
    st.stop()

# ---- Adjust these to match your file column names ----
DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "STATUS"
TECH_COL = "TECH"  # or TECHNICIAN
CALL_ID_COL = "TD REPORT NO"  # or Call Ref No / unique ID
# -----------------------------------------------------

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

# Filters
left, right = st.columns([2, 3])

with left:
    customers = ["(All)"] + sorted([c for c in df[CUSTOMER_COL].dropna().unique()])
    sel_customer = st.selectbox("Customer", customers)

with right:
    min_d, max_d = df[DATE_COL].min().date(), df[DATE_COL].max().date()
    d1, d2 = st.date_input("Date range", (min_d, max_d))

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL] == sel_customer)

dff = df.loc[mask].copy()

# KPIs
k1, k2, k3 = st.columns(3)
k1.metric("Total Calls", int(dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff else len(dff)))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

c1, c2 = st.columns(2)

# Pie: status breakdown
with c1:
    st.subheader("Status Breakdown")
    pie = dff[STATUS_COL].fillna("BLANK").value_counts().reset_index()
    pie.columns = ["Status", "Count"]
    st.plotly_chart(
        {
            "data": [{"type": "pie", "labels": pie["Status"], "values": pie["Count"]}],
            "layout": {"margin": {"l": 10, "r": 10, "t": 10, "b": 10}},
        },
        use_container_width=True,
    )

# Stacked: status by technician (counts)
with c2:
    st.subheader("Technician Status (Stacked)")
    piv = pd.crosstab(dff[TECH_COL].fillna("UNKNOWN"), dff[STATUS_COL].fillna("BLANK"))
    st.bar_chart(piv)

st.divider()

# Completion rate chart (ranked)
st.subheader("Completion Rate by Technician (Ranked)")
totals = dff.groupby(dff[TECH_COL].fillna("UNKNOWN")).size().rename("Total")
completed = dff[dff[STATUS_COL] == "COMPLETED"].groupby(dff[TECH_COL].fillna("UNKNOWN")).size().rename("Completed")
rate = pd.concat([totals, completed], axis=1).fillna(0)
rate["CompletionRate"] = (rate["Completed"] / rate["Total"]).where(rate["Total"] > 0, 0)
rate = rate.sort_values("CompletionRate", ascending=False)

st.dataframe(rate.reset_index().rename(columns={"index": "Technician"}), use_container_width=True)
