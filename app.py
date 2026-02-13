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

# Helpers
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
    st.subheader("Customer Trend (Stacked)")

    # Two columns: chart left, legend+filter right
    left_chart, right_panel = st.columns([4, 1.2], gap="large")

    # Decide default view
    default_mode = "Total" if months_span <= 1 else "Month"
    trend_mode = right_panel.selectbox(
        "View",
        ["Total", "Day", "Week", "Month"],
        index=["Total", "Day", "Week", "Month"].index(default_mode),
        key="trend_mode",
    )

    # Create PERIOD based on trend_mode
    if trend_mode == "Total":
        dff["PERIOD"] = "TOTAL"
    elif trend_mode == "Day":
        dff["PERIOD"] = dff[DATE_COL].dt.date.astype(str)
    elif trend_mode == "Week":
        dff["PERIOD"] = dff[DATE_COL].dt.strftime("%G-W%V")  # ISO week label
    else:
        dff["PERIOD"] = dff[DATE_COL].dt.to_period("M").astype(str)

    period_status = dff.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    # Build chart (legend OFF, we’ll render legend in the right panel)
    fig_stack = px.bar(
        period_status,
        x="PERIOD",
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
    )
    fig_stack.update_layout(
        xaxis_title="",
        yaxis_title="Count",
        showlegend=False,
        margin=dict(t=20, r=10, l=10, b=10),
    )

    left_chart.plotly_chart(fig_stack, use_container_width=True)

    # Manual legend (right side)
    # Keep the same status order used elsewhere
    status_order = ["ATTENDED", "COMPLETED", "NOT ATTENDED", "BLANK"]
    present = [s for s in status_order if s in period_status[STATUS_COL].unique().tolist()]
    extras = [s for s in period_status[STATUS_COL].unique().tolist() if s not in present]
    legend_items = present + extras

    right_panel.markdown("**STATUS**")
    for s in legend_items:
        # little colored square using plotly default color mapping is hard to sync exactly,
        # but this looks clean and readable.
        right_panel.markdown(f"- {s}")

    # This creates the spacing so the dropdown sits where you marked (under legend)
    right_panel.markdown("<br><br><br>", unsafe_allow_html=True)

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
