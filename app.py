import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"  # must exist in repo

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

# ---- expected columns ----
DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "STATUS"
CALL_ID_COL = "TD REPORT NO."
# --------------------------

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

# Parse + clean
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL])

df[STATUS_COL] = df[STATUS_COL].astype(str).str.strip().str.upper()
df[STATUS_COL] = df[STATUS_COL].replace({"NAN": "BLANK", "NONE": "BLANK"}).fillna("BLANK")

# -------------------- SIDEBAR (ONLY 2 FILTERS) --------------------
st.sidebar.header("Filters")

customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().astype(str).unique().tolist())
sel_customer = st.sidebar.selectbox("Customer", customers, key="customer")

min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()

d1, d2 = st.sidebar.date_input(
    "Date range",
    value=(min_d, max_d),
    min_value=min_d,
    max_value=max_d,
    key="date_range",
)

# Apply filters
mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)

dff = df.loc[mask].copy()

# Helper
def months_in_range(start_date, end_date) -> int:
    sy, sm = start_date.year, start_date.month
    ey, em = end_date.year, end_date.month
    return (ey * 12 + em) - (sy * 12 + sm) + 1

months_span = months_in_range(d1, d2)

# -------------------- KPIs --------------------
k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# -------------------- TOP ROW CHARTS --------------------
c1, c2 = st.columns(2)

with c1:
    st.subheader("Call Status")
    status_counts = dff[STATUS_COL].value_counts().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]
    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT")
    st.plotly_chart(fig_pie, use_container_width=True)

with c2:
    # Title should reflect selected customer
    customer_title = "All Customers" if sel_customer == "(All)" else sel_customer
    st.subheader(customer_title)

    # chart + right controls (wider so View shows fully)
    left_chart, right_panel = st.columns([3.6, 1.6], gap="large")

    default_mode = "Total" if months_span <= 1 else "Month"
    trend_mode = right_panel.selectbox(
        "View",
        ["Total", "Day", "Week", "Month"],
        index=["Total", "Day", "Week", "Month"].index(default_mode),
        key="trend_mode",
    )

    # Build period key
    if trend_mode == "Total":
        dff["PERIOD_LABEL"] = "TOTAL"
        dff["PERIOD_SORT"] = pd.Timestamp("2000-01-01")
        x_col = "PERIOD_LABEL"
        tickformat = None

    elif trend_mode == "Day":
        dff["PERIOD_SORT"] = dff[DATE_COL].dt.floor("D")
        dff["PERIOD_LABEL"] = dff["PERIOD_SORT"].dt.strftime("%Y-%m-%d")
        x_col = "PERIOD_LABEL"
        tickformat = None

    elif trend_mode == "Week":
        # ISO week label + a sortable date (week start Monday)
        iso = dff[DATE_COL].dt.isocalendar()
        dff["PERIOD_LABEL"] = iso["YEAR"].astype(str) + "-W" + iso["WEEK"].astype(str).str.zfill(2)
        dff["PERIOD_SORT"] = dff[DATE_COL] - pd.to_timedelta(dff[DATE_COL].dt.weekday, unit="D")
        x_col = "PERIOD_LABEL"
        tickformat = None

    else:  # Month
        # Use first day of month as a real datetime so Plotly can format ticks nicely
        dff["PERIOD_SORT"] = dff[DATE_COL].dt.to_period("M").dt.to_timestamp()
        dff["PERIOD_LABEL"] = dff["PERIOD_SORT"]  # datetime on x-axis
        x_col = "PERIOD_LABEL"
        tickformat = "%b"  # Jan, Feb, ...

    period_status = (
        dff.groupby([x_col, "PERIOD_SORT", STATUS_COL])
        .size()
        .reset_index(name="COUNT")
        .sort_values("PERIOD_SORT")
    )

    # Consistent status order
    status_order = ["COMPLETED", "ATTENDED", "NOT ATTENDED", "BLANK"]
    present = [s for s in status_order if s in period_status[STATUS_COL].unique().tolist()]
    extras = [s for s in period_status[STATUS_COL].unique().tolist() if s not in present]
    final_status_order = present + extras
    period_status[STATUS_COL] = pd.Categorical(period_status[STATUS_COL], categories=final_status_order, ordered=True)

    fig_stack = px.bar(
        period_status,
        x=x_col,
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
        category_orders={STATUS_COL: final_status_order},
    )

    fig_stack.update_layout(
        xaxis_title="",
        yaxis_title="Count",
        showlegend=True,
        legend_title_text="STATUS",
        legend=dict(x=1.02, y=0.9, xanchor="left", yanchor="top"),
        margin=dict(t=20, r=160, l=10, b=10),
    )

    # Month tick formatting (Jan / Feb)
    if tickformat:
        fig_stack.update_xaxes(tickformat=tickformat)

    left_chart.plotly_chart(fig_stack, use_container_width=True)

# -------------------- TECHNICIAN (STACKED ROW / HORIZONTAL) --------------------
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

    pivot = tech_status.pivot_table(
        index=TECH_COL, columns=STATUS_COL, values="COUNT", aggfunc="sum", fill_value=0
    )
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
        y=TECH_COL,
        x="COUNT",
        color=STATUS_COL,
        orientation="h",
        barmode="stack",
        category_orders={TECH_COL: tech_order, STATUS_COL: final_status_order},
    )
    fig_tech.update_layout(
        xaxis_title="Calls",
        yaxis_title="Technician",
        legend_title="Status",
        margin=dict(t=30, r=10, l=10, b=10),
    )
    st.plotly_chart(fig_tech, use_container_width=True)

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
