import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- CONFIG --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.markdown("## Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"  # keep in same folder as app.py in your repo

# Your columns (from your Excel headers)
DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
TECH_COL = "TECH  1"          # NOTE: two spaces between TECH and 1
CALL_ID_COL = "TD REPORT NO."

# Try to find the correct status column automatically
STATUS_CANDIDATES = [
    "STATUS",
    "CALL STATUS",
    "JOB STATUS",
    "ACTION TAKEN",
    "TECHNICIAN FEEDBACK",
]

STATUS_ORDER = ["ATTENDED", "COMPLETED", "NOT ATTENDED"]


# -------------------- HELPERS --------------------
def pick_col(cols, candidates):
    cols_set = set(cols)
    for c in candidates:
        if c in cols_set:
            return c
    return None


def canon_status(s: str) -> str:
    s = (s or "").strip().upper()

    if s in ("", "BLANK", "NAN", "NONE"):
        return "OTHER"

    if "COMPLETE" in s:
        return "COMPLETED"
    if ("NOT" in s and "ATTEND" in s) or ("NO" in s and "ATTEND" in s):
        return "NOT ATTENDED"
    if "ATTEND" in s:
        return "ATTENDED"

    return "OTHER"


def make_period_col(df: pd.DataFrame, granularity: str) -> pd.Series:
    # granularity: "Day" | "Week" | "Month"
    if granularity == "Day":
        return df[DATE_COL].dt.date.astype(str)
    if granularity == "Week":
        # week start (Mon) label like 2026-02-10
        wk_start = df[DATE_COL].dt.to_period("W-MON").apply(lambda p: p.start_time.date())
        return wk_start.astype(str)
    # Month
    return df[DATE_COL].dt.to_period("M").astype(str)


@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


# -------------------- LOAD --------------------
df = load_data()
cols = df.columns.tolist()

STATUS_COL = pick_col(cols, STATUS_CANDIDATES)

required = [DATE_COL, CUSTOMER_COL, TECH_COL, CALL_ID_COL]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", cols)
    st.stop()

if not STATUS_COL:
    st.error("Could not find a Status column automatically.")
    st.write("Available columns:", cols)
    st.stop()

# Parse date + clean
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

df[CUSTOMER_COL] = df[CUSTOMER_COL].astype(str).str.strip()
df[TECH_COL] = df[TECH_COL].astype(str).str.strip()

df[STATUS_COL] = df[STATUS_COL].fillna("").astype(str).str.strip().str.upper()
df["STATUS_STD"] = df[STATUS_COL].map(canon_status)

# Only available dates in sheet
available_dates = sorted(df[DATE_COL].dt.date.unique())
min_d, max_d = available_dates[0], available_dates[-1]


# -------------------- TOP CONTROLS --------------------
c1, c2, c3 = st.columns([2.2, 2.2, 2.2])

with c1:
    customers = ["(ALL)"] + sorted(df[CUSTOMER_COL].dropna().unique().tolist())
    sel_customer = st.selectbox("Customer", customers, index=0)

with c2:
    date_mode = st.radio("Date filter", ["Single day", "Range"], horizontal=True)

with c3:
    granularity = st.selectbox("Trend granularity", ["Day", "Week", "Month"], index=2)

# Date selection (only valid dates)
if date_mode == "Single day":
    sel_day = st.selectbox("Select a date (available dates only)", available_dates, index=len(available_dates) - 1)
    start_d, end_d = sel_day, sel_day
else:
    start_d, end_d = st.date_input("Select range", (min_d, max_d), min_value=min_d, max_value=max_d)
    # keep it safe
    if start_d > end_d:
        start_d, end_d = end_d, start_d

# -------------------- APPLY FILTERS --------------------
mask = (df[DATE_COL].dt.date >= start_d) & (df[DATE_COL].dt.date <= end_d)
dff = df.loc[mask].copy()

if sel_customer != "(ALL)":
    dff = dff[dff[CUSTOMER_COL] == sel_customer].copy()

# -------------------- KPIs --------------------
k1, k2, k3 = st.columns(3)
total_calls = int(dff[CALL_ID_COL].nunique())
k1.metric("Total Calls", total_calls)
k2.metric("Completed", int((dff["STATUS_STD"] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff["STATUS_STD"] == "NOT ATTENDED").sum()))

st.divider()

# ==================== LAYOUT ====================
# TOP: 2 charts (Pie + Trend)
top_left, top_right = st.columns([1.05, 2.2], gap="large")

# ---------- PIE ----------
with top_left:
    st.subheader("Call Status")

    pie_counts = (
        dff["STATUS_STD"]
        .value_counts()
        .reindex(STATUS_ORDER + ["OTHER"])
        .fillna(0)
        .reset_index()
    )
    pie_counts.columns = ["STATUS", "COUNT"]
    pie_counts = pie_counts[pie_counts["COUNT"] > 0]

    fig_pie = px.pie(
        pie_counts,
        names="STATUS",
        values="COUNT",
        category_orders={"STATUS": STATUS_ORDER + ["OTHER"]},
        hole=0.0,
    )
    fig_pie.update_traces(textposition="inside", textinfo="value+percent")
    fig_pie.update_layout(margin=dict(l=0, r=0, t=0, b=0), legend_title_text="")
    st.plotly_chart(fig_pie, use_container_width=True)

# ---------- TREND (STACKED) ----------
with top_right:
    title_customer = sel_customer if sel_customer != "(ALL)" else "(ALL)"
    st.subheader(f"{title_customer} — Trend ({granularity})")

    dff["PERIOD"] = make_period_col(dff, granularity)

    trend = (
        dff.groupby(["PERIOD", "STATUS_STD"])
        .size()
        .reset_index(name="COUNT")
    )

    trend["STATUS_STD"] = pd.Categorical(
        trend["STATUS_STD"],
        categories=STATUS_ORDER + ["OTHER"],
        ordered=True,
    )
    trend = trend.sort_values(["PERIOD", "STATUS_STD"])

    fig_trend = px.bar(
        trend,
        x="PERIOD",
        y="COUNT",
        color="STATUS_STD",
        barmode="stack",
        category_orders={"STATUS_STD": STATUS_ORDER + ["OTHER"]},
    )
    fig_trend.update_layout(
        margin=dict(l=0, r=0, t=10, b=0),
        legend_title_text="",
        xaxis_title="",
        yaxis_title="",
    )
    st.plotly_chart(fig_trend, use_container_width=True)

st.divider()

# BOTTOM: Technician chart (STACKED) full width
st.subheader("Technician Performance (Stacked)")

dff_tech = dff.dropna(subset=[TECH_COL]).copy()
dff_tech = dff_tech[dff_tech[TECH_COL].astype(str).str.strip().ne("")]
dff_tech = dff_tech[dff_tech["STATUS_STD"].isin(STATUS_ORDER + ["OTHER"])]

tech_counts = (
    dff_tech.groupby([TECH_COL, "STATUS_STD"])
    .size()
    .reset_index(name="COUNT")
)

# Sort techs by completion rate (Completed / Total)
pivot = tech_counts.pivot_table(
    index=TECH_COL,
    columns="STATUS_STD",
    values="COUNT",
    aggfunc="sum",
    fill_value=0
)
for c in STATUS_ORDER + ["OTHER"]:
    if c not in pivot.columns:
        pivot[c] = 0

pivot["TOTAL"] = pivot[STATUS_ORDER + ["OTHER"]].sum(axis=1)
pivot["COMPLETION_RATE"] = (pivot.get("COMPLETED", 0) / pivot["TOTAL"]).fillna(0)

tech_order = (
    pivot.sort_values(["COMPLETION_RATE", "COMPLETED"], ascending=[False, False])
    .index
    .tolist()
)

# Horizontal stacked bars = easiest to read with many technicians
fig_tech = px.bar(
    tech_counts,
    y=TECH_COL,
    x="COUNT",
    color="STATUS_STD",
    orientation="h",
    barmode="stack",
    category_orders={
        TECH_COL: tech_order,
        "STATUS_STD": STATUS_ORDER + ["OTHER"]
    },
)
fig_tech.update_layout(
    margin=dict(l=0, r=0, t=10, b=0),
    legend_title_text="",
    xaxis_title="",
    yaxis_title="",
    height=750,  # gives space for many names; adjust if needed
)
st.plotly_chart(fig_tech, use_container_width=True)

with st.expander("Show filtered rows"):
    st.dataframe(dff, use_container_width=True)
