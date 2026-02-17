import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# -------------------- Page --------------------
st.set_page_config(
    page_title="Service Calls Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    if pd.isna(s):
        return "BLANK"
    s = str(s).strip().upper()
    if s in ["", "NAN", "NONE", "NULL", "(BLANK)"]:
        return "BLANK"
    return s

def normalize_text(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    return x if x else None

def multiselect_with_all(label, options, default_all=True, key=None):
    options = [o for o in options if o is not None]
    options_sorted = sorted(options, key=lambda s: str(s).lower())

    all_label = "(All)"
    ui_options = [all_label] + options_sorted
    default = [all_label] if default_all else []

    chosen = st.multiselect(label, ui_options, default=default, key=key)

    if all_label in chosen or len(chosen) == 0:
        return None
    return chosen

# -------------------- Load --------------------
df = load_data()

DATE_COL = ensure_col(df, DATE_COL)
CUSTOMER_COL = ensure_col(df, CUSTOMER_COL)
STATUS_COL = ensure_col(df, STATUS_COL)
TECH_COL = ensure_col(df, TECH_COL)
CALL_ID_COL = ensure_col(df, CALL_ID_COL)

if None in [DATE_COL, CUSTOMER_COL, STATUS_COL]:
    st.error("Missing required columns in Excel.")
    st.stop()

df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

df[STATUS_COL] = df[STATUS_COL].apply(normalize_status)
df = df[df[STATUS_COL] != "BLANK"].copy()

df[CUSTOMER_COL] = df[CUSTOMER_COL].apply(normalize_text)
if TECH_COL and TECH_COL in df.columns:
    df[TECH_COL] = df[TECH_COL].apply(normalize_text)

# -------------------- Sidebar Width + Arrow Fix --------------------
st.markdown(
    """
    <style>
      section[data-testid="stSidebar"] { width: 340px !important; }
      section[data-testid="stSidebar"] > div { width: 340px !important; }

      /* Fix toggle arrow position */
      [data-testid="collapsedControl"] {
        position: fixed;
        top: 60px;
        left: 340px;
        z-index: 9999;
      }

      section[data-testid="stSidebar"][aria-expanded="false"] ~ div [data-testid="collapsedControl"] {
        left: 0px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

colors = status_palette()
status_order = ["COMPLETED", "ATTENDED", "NOT ATTENDED"]

# -------------------- Sidebar --------------------
with st.sidebar:
    st.markdown("## Filters")

    customer_options = df[CUSTOMER_COL].dropna().unique().tolist()
    sel_customers = multiselect_with_all("Customer", customer_options, key="cust_multi")

    if TECH_COL and TECH_COL in df.columns:
        tech_options = df[TECH_COL].dropna().unique().tolist()
        sel_techs = multiselect_with_all("Technician", tech_options, key="tech_multi")
    else:
        sel_techs = None

    status_dropdown = ["(All)"] + status_order
    sel_status = st.selectbox("Status", status_dropdown, index=0)

    min_d = df[DATE_COL].min().date()
    max_d = df[DATE_COL].max().date()

    today = date.today()
    if st.session_state.get("_end_date_last_set") != today:
        st.session_state["end_date"] = today
        st.session_state["_end_date_last_set"] = today

    with st.expander("Choose a date range", expanded=True):
        d1 = st.date_input("Start date", min_d, key="start_date")
        d2_ui = st.date_input("End date", key="end_date")

    if d1 > d2_ui:
        d1, d2_ui = d2_ui, d1

# -------------------- Filter --------------------
d2 = min(d2_ui, max_d)

mask = (df[DATE_COL].dt.date >= d1) & (df[DATE_COL].dt.date <= d2)

if sel_customers is not None:
    mask &= df[CUSTOMER_COL].isin(sel_customers)

if sel_techs is not None and TECH_COL and TECH_COL in df.columns:
    mask &= df[TECH_COL].isin(sel_techs)

if sel_status != "(All)":
    mask &= (df[STATUS_COL] == sel_status)

dff = df.loc[mask].copy()

# -------------------- KPIs --------------------
st.title("Service Calls Dashboard")

k1, k2, k3 = st.columns(3)

total_calls = dff[CALL_ID_COL].nunique() if CALL_ID_COL in dff.columns else len(dff)

k1.metric("Total Calls", int(total_calls))
k2.metric("Completed", int((dff[STATUS_COL] == "COMPLETED").sum()))
k3.metric("Not Attended", int((dff[STATUS_COL] == "NOT ATTENDED").sum()))

# -------------------- Charts --------------------
left, right = st.columns(2, gap="large")

chart_base = dff.copy()
if sel_status == "(All)":
    chart_base = chart_base[chart_base[STATUS_COL].isin(status_order)].copy()

with left:
    st.subheader("Call Status")

    status_counts = (
        chart_base[STATUS_COL]
        .value_counts()
        .reindex(status_order)
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
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# -------------------- Technician Performance --------------------
st.markdown("## Technician Performance (Sorted by Completion Rate)")

if TECH_COL and TECH_COL in dff.columns:
    tech_counts = dff.groupby([TECH_COL, STATUS_COL]).size().reset_index(name="COUNT")

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
        category_orders={TECH_COL: order},
    )
    st.plotly_chart(fig_tech, use_container_width=True)

# -------------------- Show Data --------------------
with st.expander("Show data"):
    st.dataframe(dff, use_container_width=True)
