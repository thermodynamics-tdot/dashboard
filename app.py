import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# -------------------- Page --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

# -------------------- Config --------------------
FILE_NAME = "CALL RECORDS 2026.xlsx"  # in repo root

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "CALL STATUS"           # will be auto-mapped if different
TECH_COL = "TECH 1"
CALL_ID_COL = "TD REPORT NO."

# -------------------- Helpers --------------------
@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [" ".join(str(c).strip().upper().split()) for c in df.columns]
    return df

def ensure_col(df, want):
    """Return actual column name in df that matches want or known aliases."""
    want_n = " ".join(want.strip().upper().split())
    if want_n in df.columns:
        return want_n

    aliases = {
        "CALL STATUS": ["STATUS", "CALLSTATUS", "CALL_STATUS", "CALL  STATUS", "CALL-STATUS"],
        "TECH 1": ["TECH1", "TECH  1", "TECHNICIAN", "TECH"],
        "TD REPORT NO.": ["TD REPORT NO", "TD_REPORT_NO", "REPORT NO", "REPORT NO."],
        "DATE": ["CALL DATE", "SERVICE DATE"],
        "CUSTOMER": ["CLIENT", "CUSTOMER NAME"],
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
        "BLANK": "#BDBDBD",
    }

def normalize_status(s):
    s = str(s).strip().upper()
    if s in ["NAN", "", "NONE", "NULL", "(BLANK)"]:
        return "BLANK"
    return s

def normalize_text(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    return x if x else None

def dropdown_checkbox_filter(title, options, key_prefix, default_all=True, height_px=220, expanded=False):
    """
    Dropdown-style filter (expander) with:
    - Search
    - All toggle
    - Scrollable checkbox list (label left, checkbox right)
    Returns selected list (or all if none selected).
    """
    options = [o for o in options if o is not None]
    options = sorted(options, key=lambda s: str(s).lower())

    # Keep "All" state
    all_key = f"{key_prefix}_all"
    if all_key not in st.session_state:
        st.session_state[all_key] = default_all

    # Nice compact title summary
    selected_count = sum(bool(st.session_state.get(f"{key_prefix}_opt_{hash(o)}", default_all)) for o in options)
    summary = "All" if selected_count == len(options) else f"{selected_count} selected"
    label = f"{title} — {summary}"

    with st.expander(label, expanded=expanded):
        q = st.text_input(
            "",
            placeholder=f"Search {title.lower()}...",
            key=f"{key_prefix}_search",
            label_visibility="collapsed",
        )

        visible = options
        if q:
            qn = q.strip().lower()
            visible = [o for o in options if qn in str(o).lower()]

        colL, colR = st.columns([6, 1])
        with colL:
            st.write("All")
        with colR:
            select_all = st.checkbox("", key=all_key, label_visibility="collapsed")

        st.markdown(
            f"""
            <div style="height:{height_px}px; overflow:auto; border:1px solid rgba(49,51,63,0.2);
                        border-radius:8px; padding:8px; background:rgba(255,255,255,0.02);">
            """,
            unsafe_allow_html=True,
        )

        selected = []
        for opt in visible:
            k = f"{key_prefix}_opt_{hash(opt)}"
            if k not in st.session_state:
                st.session_state[k] = default_all

            if select_all:
                st.session_state[k] = True

            a, b = st.columns([6, 1])
            with a:
                st.write(str(opt))
            with b:
                st.checkbox("", key=k, label_visibility="collapsed")

            if st.session_state[k]:
                selected.append(opt)

        st.markdown("</div>", unsafe_allow_html=True)

    # If user unchecks everything, treat as "All"
    return selected if selected else options

# -------------------- Load --------------------
df = load_data()

# Resolve real column names (avoid KeyError)
DATE_COL_REAL = ensure_col(df, DATE_COL)
CUSTOMER_COL_REAL = ensure_col(df, CUSTOMER_COL)
STATUS_COL_REAL = ensure_col(df, STATUS_COL)
TECH_COL_REAL = ensure_col(df, TECH_COL)
CALL_ID_COL_REAL = ensure_col(df, CALL_ID_COL)  # can be None if not present

missing = [name for name, real in [
    ("DATE", DATE_COL_REAL),
    ("CUSTOMER", CUSTOMER_COL_REAL),
    ("CALL STATUS", STATUS_COL_REAL),
] if real is None]

if missing:
    st.error(f"Missing required columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", list(df.columns))
    st.stop()

# Use resolved names from here on
DATE_COL = DATE_COL_REAL
CUSTOMER_COL = CUSTOMER_COL_REAL
STATUS_COL = STATUS_COL_REAL
TECH_COL = TECH_COL_REAL
CALL_ID_COL = CALL_ID_COL_REAL  # may be None

# Parse date
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

# Normalize fields
df[STATUS_COL] = df[STATUS_COL].apply(normalize_status)
df[CUSTOMER_COL] = df[CUSTOMER_COL].apply(normalize_text)
if TECH_COL and TECH_COL in df.columns:
    df[TECH_COL] = df[TECH_COL].apply(normalize_text)

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

colors = status_palette()
status_order = ["COMPLETED", "ATTENDED", "NOT ATTENDED"]

with st.sidebar:
    st.markdown("## Filters")

    # ✅ Dropdown filters (collapsed)
    customer_options = df[CUSTOMER_COL].dropna().unique().tolist()
    sel_customers = dropdown_checkbox_filter(
        "Customer",
        customer_options,
        key_prefix="cust",
        default_all=True,
        height_px=220,
        expanded=False,
    )

    if TECH_COL and TECH_COL in df.columns:
        tech_options = df[TECH_COL].dropna().unique().tolist()
        sel_techs = dropdown_checkbox_filter(
            "Technician",
            tech_options,
            key_prefix="tech",
            default_all=True,
            height_px=220,
            expanded=False,
        )
    else:
        sel_techs = None

    # ✅ Status filter (same UI as customer/technician)
    status_options = df[STATUS_COL].dropna().unique().tolist()
    sel_statuses = dropdown_checkbox_filter(
        "Status",
        status_options,
        key_prefix="status",
        default_all=True,
        height_px=220,
        expanded=False,
    )

    min_d = df[DATE_COL].min().date()
    max_d = df[DATE_COL].max().date()

    # ✅ Default end date = today (clamped to available data range)
    today = date.today()
    default_end = min(max_d, max(min_d, today))

    # Date range (opens below)
    with st.expander("Choose a date range", expanded=True):
        d1 = st.date_input("Start date", min_d, key="start_date")
        d2 = st.date_input("End date", default_end, key="end_date")

    if d1 > d2:
        d1, d2 = d2, d1

# -------------------- Filter data --------------------
mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)

if sel_customers:
    mask &= df[CUSTOMER_COL].isin(sel_customers)

if sel_techs is not None and sel_techs:
    mask &= df[TECH_COL].isin(sel_techs)

if sel_statuses:
    mask &= df[STATUS_COL].isin(sel_statuses)

dff = df.loc[mask].copy()

# -------------------- Title + KPIs --------------------
st.title("Service Calls Dashboard")

k1, k2, k3 = st.columns(3)

if CALL_ID_COL and CALL_ID_COL in dff.columns:
    total_calls = dff[CALL_ID_COL].nunique()
else:
    total_calls = len(dff)

k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# -------------------- Charts --------------------
left, right = st.columns(2, gap="large")

# Charts: keep original look by hiding BLANK unless user selected it
chart_base = dff.copy()
if "BLANK" in chart_base[STATUS_COL].unique() and "BLANK" not in sel_statuses:
    chart_base = chart_base[chart_base[STATUS_COL] != "BLANK"].copy()

# Build chart status order: main 3 first, then any extras selected
selected_set = set(sel_statuses) if sel_statuses else set()
chart_status_order = [s for s in status_order if s in selected_set]
extras = [s for s in sel_statuses if s not in chart_status_order]
chart_status_order = chart_status_order + extras

# ---------- PIE ----------
with left:
    st.subheader("Call Status")

    status_counts = (
        chart_base[STATUS_COL]
        .value_counts()
        .reindex(chart_status_order)
        .fillna(0)
        .reset_index()
    )
    status_counts.columns = ["STATUS", "COUNT"]

    fig_pie = px.pie(
        status_counts,
        names="STATUS",
        values="COUNT",
        color="STATUS",
        color_discrete_map=colors,
        category_orders={"STATUS": chart_status_order},
    )
    fig_pie.update_layout(
        legend_title_text="",
        legend=dict(font=dict(size=14)),
        margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ---------- CUSTOMER TREND ----------
with right:
    if len(sel_customers) == 1:
        customer_title = sel_customers[0]
    else:
        customer_title = "Selected Customers"
    st.subheader(customer_title)

    full_range = (d1 == df[DATE_COL].min().date() and d2 == df[DATE_COL].max().date())
    default_mode = "Total" if full_range else "Month"

    chart_col, side_col = st.columns([3.3, 1.2], gap="large")

    with side_col:
        st.markdown("**View**")
        trend_mode = st.selectbox(
            "",
            ["Total", "Day", "Week", "Month"],
            index=["Total", "Day", "Week", "Month"].index(default_mode),
            label_visibility="collapsed",
        )

    tmp = chart_base.copy()

    if len(tmp) == 0:
        with chart_col:
            st.info("No data for the selected filters.")
    else:
        if trend_mode == "Total":
            tmp["PERIOD"] = "All"
            xorder = ["All"]
            tickvals, ticktext = xorder, xorder
        elif trend_mode == "Day":
            tmp["PERIOD"] = tmp[DATE_COL].dt.date.astype(str)
            xorder = sorted(tmp["PERIOD"].unique())
            tickvals, ticktext = xorder, xorder
        elif trend_mode == "Week":
            iso = tmp[DATE_COL].dt.isocalendar()
            tmp["PERIOD"] = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
            xorder = sorted(tmp["PERIOD"].unique())
            tickvals, ticktext = xorder, xorder
        else:  # Month
            tmp["PERIOD"] = tmp[DATE_COL].dt.to_period("M").astype(str)
            xorder = sorted(tmp["PERIOD"].unique())
            tickvals = xorder
            ticktext = [pd.to_datetime(p + "-01").strftime("%b") for p in xorder]

        grp = tmp.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

        fig_stack = px.bar(
            grp,
            x="PERIOD",
            y="COUNT",
            color=STATUS_COL,
            barmode="stack",
            color_discrete_map=colors,
            category_orders={"PERIOD": xorder, STATUS_COL: chart_status_order},
        )

        fig_stack.update_layout(
            legend_title_text="",
            legend=dict(font=dict(size=14)),
            margin=dict(l=10, r=10, t=10, b=10),
        )
        fig_stack.update_xaxes(
            categoryorder="array",
            categoryarray=xorder,
            tickmode="array",
            tickvals=tickvals,
            ticktext=ticktext,
            title_text="",
        )
        fig_stack.update_yaxes(title_text="Count")

        with chart_col:
            st.plotly_chart(fig_stack, use_container_width=True)

# -------------------- Technician Performance --------------------
st.markdown("## Technician Performance (Sorted by Completion Rate)")

if TECH_COL and TECH_COL in dff.columns:
    tech_df = chart_base.copy()

    if len(tech_df) == 0:
        st.info("No technician data for the selected filters.")
    else:
        tech_counts = tech_df.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")

        totals = tech_counts.groupby(TECH_COL)["COUNT"].sum().rename("TOTAL")
        completed = (
            tech_counts[tech_counts[STATUS_COL] == "COMPLETED"]
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
            category_orders={TECH_COL: order, STATUS_COL: chart_status_order},
        )
        fig_tech.update_layout(
            legend_title_text="",
            legend=dict(font=dict(size=14)),
            margin=dict(l=10, r=10, t=10, b=10),
            yaxis_title="",
            xaxis_title="Count",
        )
        st.plotly_chart(fig_tech, use_container_width=True)

# -------------------- Show Data --------------------
with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
