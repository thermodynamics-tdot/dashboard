import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- CONFIG --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

FILE_NAME = "CALL RECORDS 2026.xlsx"  # must be in repo root (same folder as app.py)

# Your actual columns (from your screenshot)
DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "TECHNICIAN FEEDBACK"   # this is your status-like field
TECH_COL = "TECH  1"                 # NOTE: two spaces between TECH and 1
CALL_ID_COL = "TD REPORT NO."

# The 3 statuses you want to visualize like Excel
STATUS_ORDER = ["ATTENDED", "COMPLETED", "NOT ATTENDED"]
STATUS_CANON = {
    "ATTENDED": "ATTENDED",
    "COMPLETED": "COMPLETED",
    "NOT ATTENDED": "NOT ATTENDED",
    "NOTATTENDED": "NOT ATTENDED",
    "NEED MATERIAL": "OTHER",
    "RESCHEDULE": "OTHER",
    "SCHEDULED": "OTHER",
    "BLANK": "OTHER",
    "": "OTHER",
}

# -------------------- LOAD --------------------
@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [str(c).strip().upper() for c in df.columns]  # normalize headers
    return df

df = load_data()

# Validate required columns
required = [DATE_COL, CUSTOMER_COL, STATUS_COL, TECH_COL, CALL_ID_COL]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", df.columns.tolist())
    st.stop()

# Clean + parse
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

df[CUSTOMER_COL] = df[CUSTOMER_COL].astype(str).str.strip()
df[TECH_COL] = df[TECH_COL].astype(str).str.strip()

df[STATUS_COL] = (
    df[STATUS_COL]
    .fillna("BLANK")
    .astype(str)
    .str.strip()
    .str.upper()
)

def canon_status(s: str) -> str:
    s = (s or "").strip().upper()
    if s in STATUS_CANON:
        return STATUS_CANON[s]
    # If it contains the keywords, normalize
    if "COMPLETE" in s:
        return "COMPLETED"
    if "NOT" in s and "ATTEND" in s:
        return "NOT ATTENDED"
    if "ATTEND" in s:
        return "ATTENDED"
    return "OTHER"

df["STATUS_STD"] = df[STATUS_COL].map(canon_status)

# -------------------- PAGE TITLE --------------------
st.markdown("## Service Calls Dashboard")

# -------------------- LAYOUT (MATCH EXCEL) --------------------
# Left: Filters + Pie
# Middle: Stacked monthly
# Right: Technician performance
left, mid, right = st.columns([1.05, 1.9, 1.35], gap="large")

# ---------- LEFT COLUMN ----------
with left:
    # Date filter (top-left like Excel timeline)
    st.markdown("### DATE")
    min_d = df[DATE_COL].min().date()
    max_d = df[DATE_COL].max().date()

    d1, d2 = st.date_input(" ", (min_d, max_d))
    if isinstance(d1, (list, tuple)):  # just in case
        d1, d2 = d1[0], d1[1]

    # Filter data by date
    mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
    dff = df.loc[mask].copy()

    # Pie chart (below the date filter)
    st.markdown("### Call Status")
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
        hole=0.0,
        category_orders={"STATUS": STATUS_ORDER + ["OTHER"]},
    )
    fig_pie.update_traces(textposition="inside", textinfo="value+percent")
    fig_pie.update_layout(
        margin=dict(l=0, r=0, t=0, b=0),
        legend_title_text="",
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ---------- MIDDLE COLUMN ----------
with mid:
    # Customer dropdown (top of middle like Excel customer filter)
    st.markdown("### CUSTOMER")
    customers = ["(ALL)"] + sorted(dff[CUSTOMER_COL].dropna().unique().tolist())
    sel_customer = st.selectbox(" ", customers, index=0)

    dff_mid = dff.copy()
    if sel_customer != "(ALL)":
        dff_mid = dff_mid[dff_mid[CUSTOMER_COL] == sel_customer].copy()

    title_customer = sel_customer if sel_customer != "(ALL)" else "(ALL)"
    st.markdown(f"## {title_customer}")

    # Monthly stacked chart (center)
    dff_mid["MONTH"] = dff_mid[DATE_COL].dt.to_period("M").astype(str)

    monthly = (
        dff_mid.groupby(["MONTH", "STATUS_STD"])
        .size()
        .reset_index(name="COUNT")
    )

    # Keep order consistent
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

# ---------- RIGHT COLUMN ----------
with right:
    st.markdown("### Technician Performance")

    dff_right = dff_mid.copy()  # respects date + customer
    dff_right = dff_right.dropna(subset=[TECH_COL]).copy()
    dff_right = dff_right[dff_right[TECH_COL].str.strip().ne("")]

    # Counts per tech per status
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

    # Sort by completion rate desc, then by completed count desc
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

# Optional: show data
with st.expander("Show filtered rows"):
    st.dataframe(dff_mid, use_container_width=True)
