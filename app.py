import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- Page --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

# -------------------- Config --------------------
FILE_NAME = "CALL RECORDS 2026.xlsx"  # in repo root

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "CALL STATUS"          # your file uses "Call Status"
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

def status_palette(statuses):
    # simple stable palette (you can change later)
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

# Clean status
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

# -------------------- TOP: Pie + Customer Trend --------------------
left, right = st.columns(2, gap="large")

# ---- Pie (keep plotly legend style)
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
    fig_pie.update_layout(
        legend_title_text="",
        margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ---- Customer Trend: bar chart + legend column + dropdown UNDER legend (your marked box)
with right:
    customer_title = "All Customers" if sel_customer == "(All)" else sel_customer
    st.subheader(customer_title)

    full_range = (d1 == min_d and d2 == max_d)
    default_mode = "Total" if full_range else "Month"

    # 2 columns: left=chart, right=legend+dropdown (this makes the dropdown sit where you marked)
    chart_col, legend_col = st.columns([3.3, 1.2], gap="large")

    # Decide mode (dropdown will be rendered INSIDE legend_col, BELOW legend)
    # We'll build data AFTER reading trend_mode (but we also need statuses/colors for legend)
    trend_mode = None
    with legend_col:
        st.markdown("**STATUS**")

    # Prepare statuses list in a stable order
    status_order = ["COMPLETED", "ATTENDED", "NOT ATTENDED", "BLANK"]
    present_statuses = [s for s in status_order if s in dff[STATUS_COL].unique()]
    # add any extras at end
    extras = [s for s in sorted(dff[STATUS_COL].unique()) if s not in present_statuses]
    statuses = present_statuses + extras if len(dff) else status_order

    colors = status_palette(statuses)

    # Render custom legend first (like left chart legend)
    with legend_col:
        render_custom_legend(statuses, colors, title="")

        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)  # spacing

        st.markdown("**View**")
        trend_mode = st.selectbox(
            "",
            ["Total", "Day", "Week", "Month"],
            index=["Total", "Day", "Week", "Month"].index(default_mode),
            key="trend_mode",
            label_visibility="collapsed",
        )

    # Build grouped data
    tmp = dff.copy()
    if trend_mode == "Total":
        tmp["PERIOD"] = "All"
        xorder = ["All"]
        xcol = "PERIOD"
        tickvals = xorder
        ticktext = xorder

    elif trend_mode == "Day":
        tmp["PERIOD"] = tmp[DATE_COL].dt.date.astype(str)
        xcol = "PERIOD"
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals = xorder
        ticktext = xorder

    elif trend_mode == "Week":
        iso = tmp[DATE_COL].dt.isocalendar()
        tmp["PERIOD"] = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
        xcol = "PERIOD"
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals = xorder
        ticktext = xorder

    else:  # Month
        tmp["PERIOD"] = tmp[DATE_COL].dt.to_period("M").astype(str)  # "2026-01"
        xcol = "PERIOD"
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals = xorder
        # show Jan/Feb instead of Jan 1 / Feb 1
        ticktext = [pd.to_datetime(p + "-01").strftime("%b") for p in xorder]

    grp = tmp.groupby([xcol, STATUS_COL]).size().reset_index(name="COUNT")

    fig_stack = px.bar(
        grp,
        x=xcol,
        y="COUNT",
        color=STATUS_COL,
        barmode="stack",
        color_discrete_map=colors,
        category_orders={xcol: xorder, STATUS_COL: statuses},
    )

    # IMPORTANT: hide plotly legend (we use custom legend in legend_col)
    fig_stack.update_layout(
        showlegend=False,
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

# -------------------- Technician Performance (stacked ROW) --------------------
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

    colors2 = status_palette(statuses)

    fig_tech = px.bar(
        tech_counts,
        y=TECH_COL,
        x="COUNT",
        color=STATUS_COL,
        barmode="stack",
        orientation="h",
        color_discrete_map=colors2,
        category_orders={TECH_COL: order, STATUS_COL: statuses},
    )
    fig_tech.update_layout(
        legend_title_text="",
        margin=dict(l=10, r=10, t=10, b=10),
        yaxis_title="",
        xaxis_title="Count",
    )
    st.plotly_chart(fig_tech, use_container_width=True)
else:
    st.warning(f"Technician column not found ({TECH_COL}). Available columns: {list(dff.columns)}")

with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
