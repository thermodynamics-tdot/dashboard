import streamlit as st
import pandas as pd
import plotly.express as px

# Try to enable click events (optional)
try:
    from streamlit_plotly_events import plotly_events
    HAS_EVENTS = True
except Exception:
    HAS_EVENTS = False

st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.title("Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
STATUS_COL = "CALL STATUS"
TECH_COL = "TECH 1"
CALL_ID_COL = "TD REPORT NO."

@st.cache_data(ttl=300)
def load_data():
    df = pd.read_excel(FILE_NAME)
    df.columns = [" ".join(str(c).strip().upper().split()) for c in df.columns]
    return df

df = load_data()

# Validate required columns
missing = [c for c in [DATE_COL, CUSTOMER_COL, STATUS_COL] if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {', '.join(missing)}")
    st.write("Available columns:", list(df.columns))
    st.stop()

df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

df[STATUS_COL] = df[STATUS_COL].astype(str).str.strip().str.upper().replace({"NAN": "BLANK", "": "BLANK"})
df.loc[df[STATUS_COL].isin(["NONE", "NULL"]), STATUS_COL] = "BLANK"

# Sidebar (filters)
st.sidebar.header("Filters")

customers = ["(All)"] + sorted(df[CUSTOMER_COL].dropna().astype(str).unique().tolist())
sel_customer = st.sidebar.selectbox("Customer", customers, index=0)

min_d = df[DATE_COL].min().date()
max_d = df[DATE_COL].max().date()
d1, d2 = st.sidebar.date_input("Date range", (min_d, max_d))

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)
if sel_customer != "(All)":
    mask &= (df[CUSTOMER_COL].astype(str) == sel_customer)

dff = df.loc[mask].copy()

# KPIs
k1, k2, k3 = st.columns(3)
total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)
k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# Selection state for filtering table
if "selection" not in st.session_state:
    st.session_state.selection = {"source": None, "status": None, "period": None, "trend_mode": None}

def clear_selection():
    st.session_state.selection = {"source": None, "status": None, "period": None, "trend_mode": None}

# Top charts
left, right = st.columns(2, gap="large")

# ---- Pie (click to filter table)
with left:
    st.subheader("Call Status")
    status_counts = dff[STATUS_COL].value_counts().reset_index()
    status_counts.columns = ["STATUS", "COUNT"]

    fig_pie = px.pie(status_counts, names="STATUS", values="COUNT")
    fig_pie.update_layout(margin=dict(l=10, r=10, t=10, b=10))

    if HAS_EVENTS:
        pie_click = plotly_events(
            fig_pie,
            click_event=True,
            hover_event=False,
            select_event=False,
            override_height=520,
            key="pie_events",
        )
        if pie_click:
            picked_status = pie_click[0].get("label")
            if picked_status:
                st.session_state.selection = {"source": "pie", "status": picked_status, "period": None, "trend_mode": None}
    else:
        st.plotly_chart(fig_pie, use_container_width=True)
        st.info("Click-to-filter is disabled because 'streamlit-plotly-events' is not installed.")

# ---- Customer trend (click to filter table)
with right:
    title = "All Customers" if sel_customer == "(All)" else sel_customer
    st.subheader(title)

    # default view
    trend_mode = st.selectbox("View", ["Total", "Day", "Week", "Month"], index=0, key="trend_mode_ui")

    tmp = dff.copy()
    if trend_mode == "Total":
        tmp["PERIOD"] = "All"
        ticktext = ["All"]
        tickvals = ["All"]
        xorder = ["All"]
    elif trend_mode == "Day":
        tmp["PERIOD"] = tmp[DATE_COL].dt.date.astype(str)
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals = xorder
        ticktext = xorder
    elif trend_mode == "Week":
        iso = tmp[DATE_COL].dt.isocalendar()
        tmp["PERIOD"] = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals = xorder
        ticktext = xorder
    else:
        tmp["PERIOD"] = tmp[DATE_COL].dt.to_period("M").astype(str)  # YYYY-MM
        xorder = sorted(tmp["PERIOD"].unique())
        tickvals = xorder
        ticktext = [pd.to_datetime(p + "-01").strftime("%b") for p in xorder]  # Jan, Feb

    grp = tmp.groupby(["PERIOD", STATUS_COL]).size().reset_index(name="COUNT")

    fig_stack = px.bar(grp, x="PERIOD", y="COUNT", color=STATUS_COL, barmode="stack")
    fig_stack.update_layout(margin=dict(l=10, r=10, t=10, b=10))
    fig_stack.update_xaxes(tickmode="array", tickvals=tickvals, ticktext=ticktext, title_text="")
    fig_stack.update_yaxes(title_text="Count")

    if HAS_EVENTS:
        trend_click = plotly_events(
            fig_stack,
            click_event=True,
            hover_event=False,
            select_event=False,
            override_height=520,
            key="trend_events",
        )
        if trend_click:
            picked_period = trend_click[0].get("x")
            curve_idx = trend_click[0].get("curveNumber", None)

            # Map curve index to a status label (best-effort)
            status_list = list(grp[STATUS_COL].dropna().unique())
            picked_status = status_list[curve_idx] if (curve_idx is not None and curve_idx < len(status_list)) else None

            st.session_state.selection = {
                "source": "trend",
                "status": picked_status,
                "period": picked_period,
                "trend_mode": trend_mode,
            }
    else:
        st.plotly_chart(fig_stack, use_container_width=True)

# Technician performance (stacked row)
st.subheader("Technician Performance (Sorted by Completion Rate)")

if TECH_COL in dff.columns:
    tech_df = dff.copy()
    tech_df[TECH_COL] = tech_df[TECH_COL].astype(str).str.strip()
    tech_df.loc[tech_df[TECH_COL].isin(["", "NAN", "NONE", "NULL"]), TECH_COL] = "BLANK"

    tech_counts = tech_df.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")

    totals = tech_counts.groupby(TECH_COL)["COUNT"].sum().rename("TOTAL")
    completed = tech_counts.loc[tech_counts[STATUS_COL] == "COMPLETED"].groupby(TECH_COL)["COUNT"].sum().rename("COMPLETED")
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
        category_orders={TECH_COL: order},
    )
    st.plotly_chart(fig_tech, use_container_width=True)

st.divider()

c1, c2 = st.columns([1, 4])
with c1:
    st.button("Clear chart selection", on_click=clear_selection)

st.markdown("### Data (click a chart to filter)")

filtered = dff.copy()
sel = st.session_state.selection

if sel["source"] == "pie" and sel["status"]:
    filtered = filtered[filtered[STATUS_COL] == sel["status"]]
    st.caption(f"Filtered by Pie click → STATUS = **{sel['status']}**")

elif sel["source"] == "trend":
    if sel["status"]:
        filtered = filtered[filtered[STATUS_COL] == sel["status"]]
    if sel["period"]:
        mode = sel.get("trend_mode")
        if mode == "Day":
            filtered = filtered[filtered[DATE_COL].dt.date.astype(str) == str(sel["period"])]
        elif mode == "Week":
            iso = filtered[DATE_COL].dt.isocalendar()
            wk = iso["year"].astype(str) + "-W" + iso["week"].astype(int).astype(str).str.zfill(2)
            filtered = filtered[wk == str(sel["period"])]
        elif mode == "Month":
            mo = filtered[DATE_COL].dt.to_period("M").astype(str)
            filtered = filtered[mo == str(sel["period"])]
        # Total: no extra filter

    st.caption("Filtered by Trend click")

st.dataframe(filtered, use_container_width=True)
