import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

# =========================
# CONFIG (match your file)
# =========================
FILE_NAME = "CALL RECORDS 2026.xlsx"   # must exist in repo root

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
CALL_ID_COL = "TD REPORT NO."
TECH1_COL = "TECH 1"
TECH2_COL = "TECH 2"
TIME_OUT_COL = "TIME OUT"

# =====================================
# Load + clean
# =====================================
@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    # normalize column names
    df.columns = [c.strip().upper() for c in df.columns]

    # Parse date
    if DATE_COL in df.columns:
        df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
        df = df.dropna(subset=[DATE_COL])

    # Ensure required cols exist
    for c in [CUSTOMER_COL, CALL_ID_COL, TECH1_COL, TECH2_COL, TIME_OUT_COL]:
        if c not in df.columns:
            df[c] = None

    # Clean strings
    for c in [CUSTOMER_COL, CALL_ID_COL, TECH1_COL, TECH2_COL, TIME_OUT_COL]:
        df[c] = df[c].astype(str).str.strip()
        df.loc[df[c].isin(["nan", "NaN", "None", "NaT"]), c] = ""

    # ---- Derive STATUS (because your file has no CALL STATUS column)
    def derive_status(row):
        if row[TIME_OUT_COL] != "":
            return "COMPLETED"
        elif row[TECH1_COL] != "":
            return "ATTENDED"
        else:
            return "NOT ATTENDED"

    df["STATUS"] = df.apply(derive_status, axis=1)

    return df

df = load_data()

# =====================================
# Filters (left sidebar)
# =====================================
st.sidebar.header("Filters")

min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()

d1, d2 = st.sidebar.date_input("Date range", (min_d, max_d))

customers = ["(All)"] + sorted([c for c in df[CUSTOMER_COL].dropna().unique().tolist() if c != ""])
sel_customer = st.sidebar.selectbox("Customer", customers)

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL] == sel_customer)

dff = df.loc[mask].copy()

# =====================================
# KPIs
# =====================================
k1, k2, k3 = st.columns(3)

total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
completed = (dff["STATUS"] == "COMPLETED").sum()
not_attended = (dff["STATUS"] == "NOT ATTENDED").sum()

k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int(completed))
k3.metric("Not Attended", int(not_attended))

# =====================================
# Top charts (2 columns)
# =====================================
left, right = st.columns([1, 1])

# ---- Pie: Call Status
with left:
    st.subheader("Call Status")

    status_counts = dff["STATUS"].fillna("BLANK").replace("", "BLANK").value_counts().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]

    fig_pie = px.pie(
        status_counts,
        names="STATUS",
        values="COUNT",
        hole=0
    )
    fig_pie.update_traces(textposition="inside", textinfo="percent+value")
    st.plotly_chart(fig_pie, use_container_width=True)

# ---- Customer Trend (Stacked) with "View" control placed near legend area
with right:
    heading = "All Customers" if sel_customer == "(All)" else sel_customer
    st.subheader(heading)

    # Place the view selector above legend area (right side)
    view_cols = st.columns([0.65, 0.35])
    with view_cols[1]:
        view_mode = st.selectbox("View", ["Total", "Day", "Week", "Month"], index=3, key="cust_view")

    # Build period column based on view
    dff["_DATE"] = pd.to_datetime(dff[DATE_COL], errors="coerce")

    if view_mode == "Total":
        dff["PERIOD"] = "Total"
        x_order = ["Total"]
    elif view_mode == "Day":
        dff["PERIOD"] = dff["_DATE"].dt.strftime("%Y-%m-%d")
        x_order = None
    elif view_mode == "Week":
        # ISO week safely
        iso = dff["_DATE"].dt.isocalendar()
        dff["PERIOD"] = iso["YEAR"].astype(str) + "-W" + iso["WEEK"].astype(int).astype(str).str.zfill(2)
        x_order = None
    else:  # Month
        dff["PERIOD"] = dff["_DATE"].dt.strftime("%b %Y")  # Jan 2026, Feb 2026
        # keep chronological order
        month_order = (
            dff[["_DATE"]].assign(M=dff["_DATE"].dt.to_period("M").astype(str))
            .drop_duplicates()["M"].tolist()
        )
        # Convert "YYYY-MM" to "Mon YYYY" in same order
        month_order_labels = []
        for m in month_order:
            y, mo = m.split("-")
            dt = pd.to_datetime(f"{y}-{mo}-01")
            month_order_labels.append(dt.strftime("%b %Y"))
        x_order = month_order_labels

    cust_trend = (
        dff.groupby(["PERIOD", "STATUS"])
        .size()
        .reset_index(name="COUNT")
    )

    fig_cust = px.bar(
        cust_trend,
        x="PERIOD",
        y="COUNT",
        color="STATUS",
        barmode="stack",
        category_orders={"PERIOD": x_order} if x_order else None,
    )

    # Make legend style similar to pie (simple list on right)
    fig_cust.update_layout(
        legend_title_text="STATUS",
        legend=dict(orientation="v", x=1.02, y=0.9),
        margin=dict(l=10, r=160, t=10, b=40)
    )

    st.plotly_chart(fig_cust, use_container_width=True)

# =====================================
# Technician Performance (STACKED ROW chart)
# =====================================
st.subheader("Technician Performance (Stacked)")

# Use TECH 1 primarily; fallback to TECH 2 if TECH 1 empty
tech_series = dff[TECH1_COL].replace("", pd.NA)
tech_series = tech_series.fillna(dff[TECH2_COL].replace("", pd.NA))
dff["TECH"] = tech_series.fillna("BLANK")

tech_status = (
    dff.groupby(["TECH", "STATUS"])
    .size()
    .reset_index(name="COUNT")
)

# Order techs by completed count (desc)
completed_by_tech = (
    tech_status[tech_status["STATUS"] == "COMPLETED"]
    .set_index("TECH")["COUNT"]
    .to_dict()
)
tech_order = sorted(tech_status["TECH"].unique(), key=lambda t: completed_by_tech.get(t, 0), reverse=True)

fig_tech = px.bar(
    tech_status,
    y="TECH",
    x="COUNT",
    color="STATUS",
    barmode="stack",
    category_orders={"TECH": tech_order},
    orientation="h"
)

fig_tech.update_layout(
    legend_title_text="STATUS",
    legend=dict(orientation="v", x=1.02, y=0.95),
    margin=dict(l=10, r=160, t=10, b=10),
    yaxis_title=""
)

st.plotly_chart(fig_tech, use_container_width=True)

# =====================================
# Show data (still full table for now)
# =====================================
with st.expander("Show data"):
    st.dataframe(dff.drop(columns=["_DATE"], errors="ignore"), use_container_width=True)
