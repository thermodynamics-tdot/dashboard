import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_plotly_events import plotly_events

# -------------------- Page --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

# -------------------- Config --------------------
FILE_NAME = "CALL RECORDS 2026.xlsx"

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
        "CALL STATUS": ["STATUS", "CALLSTATUS", "CALL_STATUS", "CALL  STATUS"],
        "TECH 1": ["TECH1", "TECH  1", "TECHNICIAN", "TECH"],
        "TD REPORT NO.": ["TD REPORT NO", "TD_REPORT_NO", "REPORT NO", "REPORT NO."],
    }
    for a in aliases.get(want_n, []):
        a_n = " ".join(a.strip().upper().split())
        if a_n in df.columns:
            return a_n
    return None

def status_palette(statuses):
    base = {
        "COMPLETED": "#0068C9",
        "ATTENDED": "#83C9FF",
        "NOT ATTENDED": "#FF2B2B",
        "BLANK": "#FFABAB",
        "(BLANK)": "#FFABAB",
        "OTHER": "#999999",
    }
    return {s: base.get(s, "#666666") for s in statuses}

def render_custom_legend(statuses, colors, title=""):
    if title:
        st.markdown(f"**{title}**")
    for s in statuses:
        c = colors.get(s, "#666")
        st.markdown(
            f"""
            <div style="display:flex;align-items:center;gap:10px;margin:6px 0;">
              <span style="width:12px;height:12px;background:{c};display:inline-block;border-radius:2px;"></span>
              <span style="font-size:14px;">{s}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

def init_selection_state():
    if "selection" not in st.session_state:
        st.session_state.selection = {
            "source": None,     # "pie" or "trend"
            "status": None,     # selected status (pie or trend)
            "period": None,     # selected period (trend only)
        }

init_selection_state()

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

# Parse / clean
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

df[STATUS_COL] = df[STATUS_COL].astype(str).str.strip().str.upper().replace({"NAN": "BLANK", "": "BLANK"})
df.loc[df[STATUS_COL].isin(["NONE", "NULL"]), STATUS_COL] = "BLANK"

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

# statuses stable order
status_order = ["COMPLETED", "ATTENDED", "NOT ATTENDED", "BLANK"]
present_statuses = [s for s in status_order if s in dff[STATUS_COL].unique()]
extras = [s for s in sorted(dff[STATUS_COL].unique()) if s not in present_statuses]
statuses = present_statuses + extras if len(dff) else status_order
colors = status_palette(statuses)

# -------------------- TOP: Pie + Customer Trend --------------------
left, right = st.columns(2, gap="large")

# ---- Pie with click
with left:
    st.subheader("Call Status")
    status_counts = (
        dff[STATUS_COL]
        .fillna("BLANK")
        .replace({"NAN": "BLANK", "": "BLANK"})
        .value_counts()
        .reset_index()
    )
    status_counts.columns = ["STATUS", "COUNT"]

    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT", hole=0)
    fig_pie.update_layout(legend_title_text="", margin=dict(l=10, r=10, t=10, b=10))

    # Use events to capture clicks
    pie_click = plotly_events(
        fig_pie,
        click_event=True,
        hover_event=False,
        select_event=False,
        override_height=520,
        key="pie_events",
    )

    # Update selection if clicked
    if pie_click:
        # for px.pie, clicked point has "label"
        picked_status = pie_click[0].get("label")
        if picked_status:
            st.session_state.selection = {"source": "pie", "status": picked_status, "period": None}

# ---- Customer Trend (stacked) with click + custom legend + dropdown under legend
with right:
    customer_title = "All Customers" if sel_customer == "(All)" else sel_customer
    st.subheader(customer_title)

    full_range = (d1 == df[DATE_COL].min().date() and d2 == df[DATE_COL].max().date())
    default_mode = "Total" if full_range else "Month"

    chart_col, legend_col = st.columns([3.3, 1.2], gap="large")

    with legend_col:
        st.markdown("**STATUS**")
        render_custom_legend(statuses, colors, title="")
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        st.markdown("**View**")
        trend_mode = st.selectbox(
            "",
            ["Total", "Day", "Week", "Month"],
            index=["Total", "Day", "Week", "Month"].index(default_mode),
            key="trend_mode",
            label_visibility="collapsed",
        )

    tmp = dff.copy()
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
        tmp["PERIOD"] = tmp[DATE_COL].dt.to_period("M").astype(str)  # "YYYY-MM"
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
        category_orders={"PERIOD": xorder, STATUS_COL: statuses},
    )
    fig_stack.update_layout(showlegend=False, margin=dict(l=10, r=10, t=10, b=10))
    fig_stack.update_xaxes(
        tickmode="array",
        tickvals=tickvals,
        ticktext=ticktext,
        title_text="",
    )
    fig_stack.update_yaxes(title_text="Count")

    with chart_col:
        trend_click = plotly_events(
            fig_stack,
            click_event=True,
            hover_event=False,
            select_event=False,
            override_height=520,
            key="trend_events",
        )

    if trend_click:
        # for bars: x is period, curveNumber traces status (depends), but Plotly gives "x" and "curveNumber"
        picked_period = trend_click[0].get("x")
        curve_idx = trend_click[0].get("curveNumber", None)
        picked_status = None
        # plotly_events doesn't always give status label directly, so map by curve index
        if curve_idx is not None and curve_idx < len(statuses):
            picked_status = statuses[curve_idx]
        st.session_state.selection = {"source": "trend", "status": picked_status, "period": picked_period}

# -------------------- Technician Performance (stacked row) --------------------
st.markdown("## Technician Performance (Sorted by Completion Rate)")

if TECH_COL in dff.columns:
    tech_df = dff.copy()
    tech_df[TECH_COL] = tech_df[TECH_COL].astype(str).str.strip()
    tech_df.loc[tech_df[TECH_COL].isin(["", "NAN", "NONE", "NULL"]), TECH_COL] = "BLANK"

    tech_counts = tech_df.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")

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
        category_orders={TECH_COL: order, STATUS_COL: statuses},
    )
    fig_tech.update_layout(legend_title_text="", margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_tech, use_container_width=True)

# -------------------- Click-to-filter Table --------------------
st.divider()

col_a, col_b = st.columns([1, 3])
with col_a:
    if st.button("Clear chart selection"):
        st.session_state.selection = {"source": None, "status": None, "period": None}

sel = st.session_state.selection

st.markdown("### Data (click a chart to filter)")

filtered = dff.copy()

if sel["source"] == "pie" and sel["status"]:
    filtered = filtered[filtered[STATUS_COL] == sel["status"]]
    st.caption(f"Filtered by Pie click → STATUS = **{sel['status']}**")

elif sel["source"] == "trend":
    if sel["status"]:
        filtered = filtered[filtered[STATUS_COL] == sel["status"]]
    if sel["period"]:
        # Apply period filter based on selected trend_mode
        if st.session_state.get("trend_mode") == "Total":
            pass
        elif st.session_state.get("trend_mode") == "Day":
            filtered = filtered[filtered[DATE_COL].dt.date.astype(str) == str(sel["period"])]
        elif st.session_state.get("trend_mode") == "Week":
            iso = filtered[DATE_COL].dt.isocalendar()
            wk = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
            filtered = filtered[wk == str(sel["period"])]
        else:  # Month
            mo = filtered[DATE_COL].dt.to_period("M").astype(str)
            filtered = filtered[mo == str(sel["period"])]

    cap = "Filtered by Trend click → "
    if sel["status"]:
        cap += f"STATUS = **{sel['status']}**  "
    if sel["period"]:
        cap += f"PERIOD = **{sel['period']}**"
    st.caption(cap)

st.dataframe(filtered, use_container_width=True)
