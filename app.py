import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- CONFIG --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.markdown("## Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"  # must be in repo root (same folder as app.py)

# Your columns (based on your screenshot list)
DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
TECH_COL = "TECH  1"          # NOTE: two spaces between TECH and 1
CALL_ID_COL = "TD REPORT NO."

# ✅ IMPORTANT:
# We will TRY to find the correct status column automatically.
# It will prefer these columns in this order:
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

    if s == "" or s == "BLANK":
        return "OTHER"

    # common patterns
    if "COMPLETE" in s:
        return "COMPLETED"
    if ("NOT" in s and "ATTEND" in s) or ("NO" in s and "ATTEND" in s):
        return "NOT ATTENDED"
    if "ATTEND" in s:
        return "ATTENDED"

    # sometimes written like "NEED MATERIAL", "RESCHEDULE", etc.
    return "OTHER"


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

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

# Clean fields
df[CUSTOMER_COL] = df[CUSTOMER_COL].astype(str).str.strip()
df[TECH_COL] = df[TECH_COL].astype(str).str.strip()

df[STATUS_COL] = df[STATUS_COL].fillna("").astype(str).str.strip().str.upper()
df["STATUS_STD"] = df[STATUS_COL].map(canon_status)

# Debug panel so you can SEE why it becomes OTHER (you can hide later)
with st.expander("Debug (to fix status mapping)", expanded=True):
    st.write(f"Using STATUS column: **{STATUS_COL}**")
    st.write("Top values in that column (raw):")
    st.dataframe(df[STATUS_COL].value_counts().head(30).reset_index().rename(columns={"index": STATUS_COL, STATUS_COL: "COUNT"}))
    st.write("Mapped status counts (ATTENDED/COMPLETED/NOT ATTENDED/OTHER):")
    st.dataframe(df["STATUS_STD"].value_counts().reset_index().rename(columns={"index": "STATUS_STD", "STATUS_STD": "COUNT"}))

# -------------------- FILTERS (TOP-LEFT LIKE EXCEL) --------------------
left, mid, right = st.columns([1.05, 1.9, 1.35], gap="large")

with left:
    st.markdown("### DATE")
    min_d = df[DATE_COL].min().date()
    max_d = df[DATE_COL].max().date()
    d1, d2 = st.date_input(" ", (min_d, max_d))

    mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
    dff_date = df.loc[mask].copy()

    # -------------------- PIE (LEFT) --------------------
    st.markdown("### Call Status")

    pie_counts = (
        dff_date["STATUS_STD"]
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
    )
    fig_pie.update_traces(textposition="inside", textinfo="value+percent")
    fig_pie.update_layout(margin=dict(l=0, r=0, t=0, b=0), legend_title_text="")
    st.plotly_chart(fig_pie, use_container_width=True)

# -------------------- CUSTOMER FILTER (MIDDLE TOP) --------------------
with mid:
    st.markdown("### CUSTOMER")
    customers = ["(ALL)"] + sorted(dff_date[CUSTOMER_COL].dropna().unique().tolist())
    sel_customer = st.selectbox(" ", customers, index=0)

    dff = dff_date.copy()
    if sel_customer != "(ALL)":
        dff = dff[dff[CUSTOMER_COL] == sel_customer].copy()

    title_customer = sel_customer if sel_customer != "(ALL)" else "(ALL)"
    st.markdown(f"## {title_customer}")

    # -------------------- MONTHLY STACKED (CENTER) --------------------
    dff["MONTH"] = dff[DATE_COL].dt.to_period("M").astype(str)

    monthly = (
        dff.groupby(["MONTH", "STATUS_STD"])
        .size()
        .reset_index(name="COUNT")
    )

    monthly["STATUS_STD"] = pd.Categorical(
        monthly["STATUS_STD"],
        categories=STATUS_ORDER + ["OTHER"],
        ordered=True,
    )
    monthly = monthly.sort_values(["MONTH", "STATUS_STD"])

    fig_stack = px.bar(
        monthly,
        x="MONTH",
        y="COUNT",
        color="STATUS_STD",
        barmode="stack",
        category_orders={"STATUS_STD": STATUS_ORDER + ["OTHER"]},
    )
    fig_stack.update_layout(
        margin=dict(l=0, r=0, t=10, b=0),
        legend_title_text="",
        xaxis_title="",
        yaxis_title="",
    )
    st.plotly_chart(fig_stack, use_container_width=True)

# -------------------- TECH PERFORMANCE (RIGHT) --------------------
with right:
    st.markdown("### Technician Performance")

    dff_right = dff.dropna(subset=[TECH_COL]).copy()
    dff_right = dff_right[dff_right[TECH_COL].astype(str).str.strip().ne("")]
    dff_right = dff_right[dff_right["STATUS_STD"].isin(STATUS_ORDER + ["OTHER"])]

    tech_counts = (
        dff_right.groupby([TECH_COL, "STATUS_STD"])
        .size()
        .reset_index(name="COUNT")
    )

    # Completion rate for sorting
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

    fig_tech = px.bar(
        tech_counts,
        y=TECH_COL,
        x="COUNT",
        color="STATUS_STD",
        orientation="h",
        barmode="group",
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
    )
    st.plotly_chart(fig_tech, use_container_width=True)

# -------------------- DATA TABLE (OPTIONAL) --------------------
with st.expander("Show filtered rows"):
    st.dataframe(dff, use_container_width=True)
