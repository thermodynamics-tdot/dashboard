import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- CONFIG --------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")
st.markdown("## Service Calls Dashboard")

FILE_NAME = "CALL RECORDS 2026.xlsx"  # put this file in same folder as app.py

DATE_COL = "DATE"
CUSTOMER_COL = "CUSTOMER"
TECH_COL = "TECH  1"        # NOTE: two spaces between TECH and 1
CALL_ID_COL = "TD REPORT NO."

STATUS_CANDIDATES = ["STATUS", "CALL STATUS", "JOB STATUS", "ACTION TAKEN", "TECHNICIAN FEEDBACK"]
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


def label_day(ts: pd.Series) -> pd.Series:
    return ts.dt.date.astype(str)


def label_week(ts: pd.Series) -> pd.Series:
    wk_start = ts.dt.to_period("W-MON").apply(lambda p: p.start_time.date())
    return wk_start.astype(str)


def label_month(ts: pd.Series) -> pd.Series:
    return ts.dt.to_period("M").astype(str)


def apply_quick_range(mode: str, available_dates: list):
    min_d, max_d = available_dates[0], available_dates[-1]

    if mode == "Today (last available)":
        return max_d, max_d

    if mode == "Last 7 days":
        end_d = max_d
        start_d = end_d - pd.Timedelta(days=6)
        start_d = max(start_d.date(), min_d)
        return start_d, end_d

    if mode == "Last 30 days":
        end_d = max_d
        start_d = end_d - pd.Timedelta(days=29)
        start_d = max(start_d.date(), min_d)
        return start_d, end_d

    if mode == "This week":
        end_d = max_d
        p = pd.Period(end_d, freq="W-MON")
        start_d = p.start_time.date()
        if start_d < min_d:
            start_d = min_d
        return start_d, end_d

    if mode == "This month":
        end_d = max_d
        p = pd.Period(end_d, freq="M")
        start_d = p.start_time.date()
        if start_d < min_d:
            start_d = min_d
        return start_d, end_d

    return min_d, max_d  # All time


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

df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
df = df.dropna(subset=[DATE_COL]).copy()

df[CUSTOMER_COL] = df[CUSTOMER_COL].astype(str).str.strip()
df[TECH_COL] = df[TECH_COL].astype(str).str.strip()
df[STATUS_COL] = df[STATUS_COL].fillna("").astype(str).str.strip().str.upper()
df["STATUS_STD"] = df[STATUS_COL].map(canon_status)

available_dates = sorted(df[DATE_COL].dt.date.unique())
min_d, max_d = available_dates[0], available_dates[-1]


# -------------------- LAYOUT --------------------
content_col, filter_col = st.columns([4.7, 1.3], gap="large")

with filter_col:
    with st.expander("Filters", expanded=True):
        customers = ["(ALL)"] + sorted(df[CUSTOMER_COL].dropna().unique().tolist())
        sel_customer = st.selectbox("Customer", customers, index=0, key="customer_select")

        st.markdown("**Date**")
        tab_quick, tab_custom = st.tabs(["Quick", "Custom"])

        # defaults
        start_d, end_d = min_d, max_d
        group_by = "Month"

        with tab_quick:
            quick = st.selectbox(
                "Range",
                ["Today (last available)", "Last 7 days", "Last 30 days", "This week", "This month", "All time"],
                index=3,
                key="quick_range",
            )
            start_d, end_d = apply_quick_range(quick, available_dates)

            if start_d != end_d:
                group_by = st.selectbox("Group by", ["Day", "Week", "Month"], index=2, key="group_by_quick")

        with tab_custom:
            mode = st.radio(
                " ",
                ["Single day", "Range"],
                horizontal=True,
                label_visibility="collapsed",
                key="custom_mode",
            )

            if mode == "Single day":
                sel_day = st.selectbox("Day", available_dates, index=len(available_dates) - 1, key="single_day_pick")
                start_d, end_d = sel_day, sel_day
            else:
                start_d, end_d = st.date_input(
                    "Pick dates",
                    (min_d, max_d),
                    min_value=min_d,
                    max_value=max_d,
                    key="range_pick",
                )
                if start_d > end_d:
                    start_d, end_d = end_d, start_d

                group_by = st.selectbox("Group by", ["Day", "Week", "Month"], index=2, key="group_by_custom")

    with st.expander("Debug (optional)", expanded=False):
        st.write(f"Status column used: **{STATUS_COL}**")


# -------------------- APPLY FILTERS --------------------
mask = (df[DATE_COL].dt.date >= start_d) & (df[DATE_COL].dt.date <= end_d)
dff = df.loc[mask].copy()

if sel_customer != "(ALL)":
    dff = dff[dff[CUSTOMER_COL] == sel_customer].copy()

single_day_mode = (start_d == end_d)


# -------------------- CONTENT --------------------
with content_col:
    # KPIs
    k1, k2, k3 = st.columns(3)
    total_calls = int(dff[CALL_ID_COL].nunique())
    k1.metric("Total Calls", total_calls)
    k2.metric("Completed", int((dff["STATUS_STD"] == "COMPLETED").sum()))
    k3.metric("Not Attended", int((dff["STATUS_STD"] == "NOT ATTENDED").sum()))

    st.divider()

    # TOP: Pie + Trend
    left, right = st.columns([1.05, 2.2], gap="large")

    with left:
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
        )
        fig_pie.update_traces(textposition="inside", textinfo="value+percent")
        fig_pie.update_layout(margin=dict(l=0, r=0, t=0, b=0), legend_title_text="")
        st.plotly_chart(fig_pie, use_container_width=True)

    with right:
        title_customer = sel_customer if sel_customer != "(ALL)" else "(ALL)"

        if single_day_mode:
            st.subheader(f"{title_customer} — {start_d}")
            by_status = (
                dff["STATUS_STD"]
                .value_counts()
                .reindex(STATUS_ORDER + ["OTHER"])
                .fillna(0)
                .reset_index()
            )
            by_status.columns = ["STATUS", "COUNT"]
            by_status = by_status[by_status["COUNT"] > 0]

            fig_day = px.bar(
                by_status,
                x="STATUS",
                y="COUNT",
                category_orders={"STATUS": STATUS_ORDER + ["OTHER"]},
            )
            fig_day.update_layout(margin=dict(l=0, r=0, t=10, b=0), xaxis_title="", yaxis_title="")
            st.plotly_chart(fig_day, use_container_width=True)

        else:
            st.subheader(f"{title_customer} — Trend ({group_by})")

            dff_trend = dff.copy()
            if group_by == "Day":
                dff_trend["PERIOD"] = label_day(dff_trend[DATE_COL])
            elif group_by == "Week":
                dff_trend["PERIOD"] = label_week(dff_trend[DATE_COL])
            else:
                dff_trend["PERIOD"] = label_month(dff_trend[DATE_COL])

            trend = (
                dff_trend.groupby(["PERIOD", "STATUS_STD"])
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

    # BOTTOM: Technician stacked (sorted by completion rate)
    st.subheader("Technician Performance (Stacked)")

    dff_tech = dff.dropna(subset=[TECH_COL]).copy()
    dff_tech = dff_tech[dff_tech[TECH_COL].astype(str).str.strip().ne("")]
    dff_tech = dff_tech[dff_tech["STATUS_STD"].isin(STATUS_ORDER + ["OTHER"])]

    tech_counts = (
        dff_tech.groupby([TECH_COL, "STATUS_STD"])
        .size()
        .reset_index(name="COUNT")
    )

    pivot = tech_counts.pivot_table(
        index=TECH_COL,
        columns="STATUS_STD",
        values="COUNT",
        aggfunc="sum",
        fill_value=0,
    )
    for c in STATUS_ORDER + ["OTHER"]:
        if c not in pivot.columns:
            pivot[c] = 0

    pivot["TOTAL"] = pivot[STATUS_ORDER + ["OTHER"]].sum(axis=1)
    pivot["COMPLETION_RATE"] = (pivot.get("COMPLETED", 0) / pivot["TOTAL"]).fillna(0)

    tech_order = (
        pivot.sort_values(["COMPLETION_RATE", "COMPLETED"], ascending=[False, False])
        .index.tolist()
    )

    fig_tech = px.bar(
        tech_counts,
        y=TECH_COL,
        x="COUNT",
        color="STATUS_STD",
        orientation="h",
        barmode="stack",
        category_orders={TECH_COL: tech_order, "STATUS_STD": STATUS_ORDER + ["OTHER"]},
    )
    fig_tech.update_layout(
        margin=dict(l=0, r=0, t=10, b=0),
        legend_title_text="",
        xaxis_title="",
        yaxis_title="",
        height=750,
    )
    st.plotly_chart(fig_tech, use_container_width=True)

    with st.expander("Show filtered rows"):
        st.dataframe(dff, use_container_width=True)
