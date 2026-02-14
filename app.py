import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- Page --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

# -------------------- Config --------------------
FILE_NAME = "CALL RECORDS 2026.xlsx"  # in repo root

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "CALL STATUS"
TECH_COL = "TECH 1"
CALL_ID_COL = "TD REPORT NO."

# -------------------- Helpers --------------------
@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [" ".join(str(c).strip().upper().split()) for c in df.columns]
    return df

def ensure_col(df, want):
    want_n = " ".join(want.strip().upper().split())
    if want_n in df.columns:
        return want_n
    aliases = {
        "CALL STATUS": ["STATUS", "CALLSTATUS", "CALL_STATUS"],
        "TECH 1": ["TECH1", "TECH  1", "TECHNICIAN", "TECH"],
        "TD REPORT NO.": ["TD REPORT NO", "TD_REPORT_NO", "REPORT NO", "REPORT NO."],
    }
    for a in aliases.get(want_n, []):
        a_n = " ".join(a.strip().upper().split())
        if a_n in df.columns:
            return a_n
    return None

def status_palette():
    return {
        "COMPLETED": "#0068C9",
        "ATTENDED": "#83C9FF",
        "NOT ATTENDED": "#FF2B2B",
    }

def normalize_status(s: str) -> str:
    s = str(s).strip().upper()
    if s in ["NAN", "", "NONE", "NULL", "(BLANK)", "BLANK"]:
        return ""  # will be removed
    return s

def month_short(period_str: str) -> str:
    # period_str like "2026-01"
    return pd.to_datetime(period_str + "-01").strftime("%b")

# -------------------- Load --------------------
df = load_data()

DATE_COL = ensure_col(df, DATE_COL) or DATE_COL
CUSTOMER_COL = ensure_col(df, CUSTOMER_COL) or CUSTOMER_COL
STATUS_COL = ensure_col(df, STATUS_COL) or STATUS_COL
TECH_COL = ensure_col(df, TECH_COL) or TECH_COL
CALL_ID_COL = ensure_col(df, CALL_ID_COL) or CALL_ID_COL

missing = [c for c in [DATE_COL, CUSTOMER_COL, STATUS_COL] if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", list(df.columns))
    st.stop()

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

# Clean status and REMOVE blanks everywhere
df[STATUS_COL] = df[STATUS_COL].apply(normalize_status)
df = df[df[STATUS_COL] != ""].copy()

# Keep only known statuses (optional; comment if you have more)
STATUS_ORDER = ["COMPLETED", "ATTENDED", "NOT ATTENDED"]
df = df[df[STATUS_COL].isin(STATUS_ORDER)].copy()

# -------------------- Sidebar (wider) --------------------
st.markdown(
    """
    <style>
      section[data-testid="stSidebar"] { width: 340px !important; }
      section[data-testid="stSidebar"] > div { width: 340px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("## Filters")

    customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().astype(str).unique().tolist())
    sel_customer = st.selectbox("Customer", customers, index=0, key="customer_filter")

    min_d = df[DATE_COL].min().date()
    max_d = df[DATE_COL].max().date()
    d1, d2 = st.date_input("Date range", (min_d, max_d), key="date_range")

# Filter dataframe
mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)
dff = df.loc[mask].copy()

# -------------------- Title + KPIs --------------------
st.title("Service Calls Dashboard")

k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# -------------------- TOP: Pie + Customer Trend --------------------
left, right = st.columns(2, gap="large")
colors = status_palette()

# ---- Pie
with left:
    st.subheader("Call Status")
    status_counts = dff[STATUS_COL].value_counts().reindex(STATUS_ORDER).dropna().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]

    fig_pie = px.pie(
        status_counts,
        names="STATUS",
        values="COUNT",
        color="STATUS",
        color_discrete_map=colors,
        category_orders={"STATUS": STATUS_ORDER},
    )
    fig_pie.update_layout(
        legend_title_text="",
        legend=dict(font=dict(size=14)),
        margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ---- Customer Trend
with right:
    customer_title = "All Customers" if sel_customer == "(All)" else sel_customer
    st.subheader(customer_title)

    # build 4 views
    tmp = dff.copy()

    # TOTAL
    total_df = tmp.copy()
    total_df["PERIOD"] = "All"
    total_grp = total_df.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    # DAY
    day_df = tmp.copy()
    day_df["PERIOD"] = day_df[DATE_COL].dt.date.astype(str)
    day_order = sorted(day_df["PERIOD"].unique())
    day_grp = day_df.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    # WEEK
    week_df = tmp.copy()
    iso = week_df[DATE_COL].dt.isocalendar()
    week_df["PERIOD"] = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
    week_order = sorted(week_df["PERIOD"].unique())
    week_grp = week_df.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    # MONTH (show Jan/Feb)
    month_df = tmp.copy()
    month_df["PERIOD"] = month_df[DATE_COL].dt.to_period("M").astype(str)  # "2026-01"
    month_order = sorted(month_df["PERIOD"].unique())
    month_grp = month_df.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    # unify all periods into one dataset with a "VIEW" column
    total_grp["VIEW"] = "Total"
    day_grp["VIEW"] = "Day"
    week_grp["VIEW"] = "Week"
    month_grp["VIEW"] = "Month"

    all_grp = pd.concat([total_grp, day_grp, week_grp, month_grp], ignore_index=True)
    all_grp = all_grp[all_grp[STATUS_COL].isin(STATUS_ORDER)].copy()

    # For prettier X labels, create DISPLAY column
    all_grp["DISPLAY"] = all_grp["PERIOD"]
    all_grp.loc[all_grp["VIEW"] == "Month", "DISPLAY"] = all_grp.loc[all_grp["VIEW"] == "Month", "PERIOD"].apply(month_short)

    # Make a single figure with animation_frame for view switching
    fig_stack = px.bar(
        all_grp,
        x="DISPLAY",
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
        color_discrete_map=colors,
        category_orders={STATUS_COL: STATUS_ORDER},
        animation_frame="VIEW",
    )

    # Default frame: Total if full range else Month
    min_d_all = df[DATE_COL].min().date()
    max_d_all = df[DATE_COL].max().date()
    full_range = (d1 == min_d_all and d2 == max_d_all)
    default_view = "Total" if full_range else "Month"

    # Plotly animation starts at first frame; reorder frames by putting default first
    if fig_stack.frames:
        frames = {fr.name: fr for fr in fig_stack.frames}
        ordered = []
        if default_view in frames:
            ordered.append(frames[default_view])
        for name in ["Total", "Day", "Week", "Month"]:
            if name in frames and name != default_view:
                ordered.append(frames[name])
        fig_stack.frames = ordered

        # Set initial data to default frame
        fig_stack.data = fig_stack.frames[0].data

    # Hide animation slider/buttons; we’ll add our own dropdown under legend
    fig_stack.update_layout(
        showlegend=True,
        legend_title_text="",
        legend=dict(font=dict(size=14)),
        margin=dict(l=10, r=10, t=10, b=10),
        updatemenus=[
            dict(
                type="dropdown",
                direction="down",
                x=1.02,            # put it to the right of plot area
                y=0.55,            # under the legend area (tweak if needed)
                xanchor="left",
                yanchor="top",
                showactive=True,
                buttons=[
                    dict(label="Total", method="animate", args=[["Total"], {"mode": "immediate", "frame": {"duration": 0}, "transition": {"duration": 0}}]),
                    dict(label="Day",   method="animate", args=[["Day"],   {"mode": "immediate", "frame": {"duration": 0}, "transition": {"duration": 0}}]),
                    dict(label="Week",  method="animate", args=[["Week"],  {"mode": "immediate", "frame": {"duration": 0}, "transition": {"duration": 0}}]),
                    dict(label="Month", method="animate", args=[["Month"], {"mode": "immediate", "frame": {"duration": 0}, "transition": {"duration": 0}}]),
                ],
                bgcolor="rgba(240,242,246,1)",
                bordercolor="rgba(240,242,246,1)",
                font=dict(size=14),
            )
        ],
        sliders=[],
    )

    # Remove play button
    fig_stack.layout["updatemenus"] = fig_stack.layout["updatemenus"]

    # Improve x-axis labels for month/day/week
    fig_stack.update_xaxes(title_text="")
    fig_stack.update_yaxes(title_text="Count")

    st.plotly_chart(fig_stack, use_container_width=True)

# -------------------- Technician Performance (stacked ROW) --------------------
st.markdown("## Technician Performance (Sorted by Completion Rate)")

if TECH_COL in dff.columns:
    tech_df = dff.copy()
    tech_df[TECH_COL] = tech_df[TECH_COL].astype(str).str.strip()
    tech_df = tech_df[~tech_df[TECH_COL].isin(["", "NAN", "NONE", "NULL"])].copy()

    tech_counts = tech_df.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")
    tech_counts = tech_counts[tech_counts[STATUS_COL].isin(STATUS_ORDER)].copy()

    totals = tech_counts.groupby(TECH_COL)["COUNT"].sum().rename("TOTAL")
    completed = (
        tech_counts.loc[tech_counts[STATUS_COL] == "COMPLETED"]
        .groupby(TECH_COL)["COUNT"]
        .sum()
        .rename("COMPLETED")
    )
    rates = pd.concat([totals, completed], axis=1).fillna(0)
    rates["COMP_RATE"] = rates["COMPLETED"] / rates["TOTAL"].where(rates["TOTAL"] != 0, 1)
    order = rates.sort_values("COMP_RATE", ascending=False).index.tolist()

    fig_tech = px.bar(
        tech_counts,
        y=TECH_COL,
        x="COUNT",
        color=STATUS_COL,
        barmode="stack",
        orientation="h",
        color_discrete_map=colors,
        category_orders={TECH_COL: order, STATUS_COL: STATUS_ORDER},
    )
    fig_tech.update_layout(
        legend_title_text="",
        legend=dict(font=dict(size=14)),
        margin=dict(l=10, r=10, t=10, b=10),
        yaxis_title="",
        xaxis_title="Count",
    )
    st.plotly_chart(fig_tech, use_container_width=True)
else:
    st.warning(f"Technician column not found ({TECH_COL}). Available columns: {list(dff.columns)}")

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
