import streamlit as st
import pandas as pd
import plotly.express as px
import requests
from io import BytesIO

# -------------------------------------------------
# Page Config
# -------------------------------------------------
st.set_page_config(page_title="Service Calls Dashboard", layout="wide")

st.title("Service Calls Dashboard")

# -------------------------------------------------
# Load Excel from SharePoint / OneDrive
# -------------------------------------------------

EXCEL_URL = st.secrets["EXCEL_URL"]

@st.cache_data(ttl=300)  # refresh every 5 minutes
def load_data(url):
    try:
        response = requests.get(url, allow_redirects=True, timeout=60)
        response.raise_for_status()
        return pd.read_excel(BytesIO(response.content))
    except Exception as e:
        st.error("Failed to load Excel file.")
        st.stop()

df = load_data(EXCEL_URL)

# -------------------------------------------------
# Data Preparation
# -------------------------------------------------

# Ensure DATE column is datetime
if "DATE" in df.columns:
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")

# Drop rows without date
df = df.dropna(subset=["DATE"])

# Create Month column
df["Month"] = df["DATE"].dt.strftime("%b")

# -------------------------------------------------
# Sidebar Filters
# -------------------------------------------------

st.sidebar.header("Filters")

# Date filter
min_date = df["DATE"].min()
max_date = df["DATE"].max()

date_range = st.sidebar.date_input(
    "Select Date Range",
    [min_date, max_date]
)

# Customer filter
if "CUSTOMER" in df.columns:
    customers = ["All"] + sorted(df["CUSTOMER"].dropna().unique())
    selected_customer = st.sidebar.selectbox("Select Customer", customers)
else:
    selected_customer = "All"

# -------------------------------------------------
# Apply Filters
# -------------------------------------------------

filtered_df = df.copy()

if len(date_range) == 2:
    filtered_df = filtered_df[
        (filtered_df["DATE"] >= pd.to_datetime(date_range[0])) &
        (filtered_df["DATE"] <= pd.to_datetime(date_range[1]))
    ]

if selected_customer != "All":
    filtered_df = filtered_df[filtered_df["CUSTOMER"] == selected_customer]

# -------------------------------------------------
# Layout
# -------------------------------------------------

col1, col2 = st.columns(2)

# -------------------------------------------------
# 1️⃣ Pie Chart - Call Status
# -------------------------------------------------

if "STATUS" in filtered_df.columns:

    status_counts = filtered_df["STATUS"].value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]

    fig_pie = px.pie(
        status_counts,
        names="Status",
        values="Count",
        title="Call Status Distribution"
    )

    col1.plotly_chart(fig_pie, use_container_width=True)

# -------------------------------------------------
# 2️⃣ Stacked Column - Monthly Trend
# -------------------------------------------------

if "STATUS" in filtered_df.columns:

    month_status = (
        filtered_df
        .groupby(["Month", "STATUS"])
        .size()
        .reset_index(name="Count")
    )

    fig_bar = px.bar(
        month_status,
        x="Month",
        y="Count",
        color="STATUS",
        title="Monthly Call Trend",
        barmode="stack"
    )

    col2.plotly_chart(fig_bar, use_container_width=True)

# -------------------------------------------------
# 3️⃣ Technician Performance
# -------------------------------------------------

if "TECH 1" in filtered_df.columns and "STATUS" in filtered_df.columns:

    tech_perf = (
        filtered_df
        .groupby(["TECH 1", "STATUS"])
        .size()
        .reset_index(name="Count")
    )

    fig_tech = px.bar(
        tech_perf,
        y="TECH 1",
        x="Count",
        color="STATUS",
        orientation="h",
        title="Technician Performance",
        barmode="group"
    )

    st.plotly_chart(fig_tech, use_container_width=True)

# -------------------------------------------------
# Raw Data (Optional)
# -------------------------------------------------

with st.expander("View Raw Data"):
    st.dataframe(filtered_df)
