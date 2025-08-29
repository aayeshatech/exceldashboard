import pandas as pd
import streamlit as st
from datetime import datetime
import time

st.set_page_config(page_title="Live Option Chain Dashboard", page_icon="📊", layout="wide")

# === Auto Refresh ===
st_autorefresh = st.sidebar.checkbox("🔄 Auto Refresh (30 sec)", value=True)
if st_autorefresh:
    st.experimental_rerun()

# === Load File ===
file_path = "Live_Option_Chain_Terminal.xlsm"

@st.cache_data(ttl=30)
def load_excel_data():
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    data = {name: pd.read_excel(xls, name) for name in xls.sheet_names if name in ["OC_1","OC_2","OC_3","Dashboard","Screener","Sector Dashboard","PCR & OI Chart"]}
    return data

data_dict = load_excel_data()

st.title("⚡ Live Options Dashboard")
st.caption(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# === PCR & Sentiment ===
if "PCR & OI Chart" in data_dict:
    df_pcr = data_dict["PCR & OI Chart"]
    try:
        pcr_oi = df_pcr.iloc[0,2]
        pcr_vol = df_pcr.iloc[0,3]
        st.metric("⚖️ PCR (OI)", f"{pcr_oi:.2f}")
        st.metric("📊 PCR (Vol)", f"{pcr_vol:.2f}")

        if pcr_oi > 1.3:
            st.error("🐻 Bearish Sentiment (High PCR)")
        elif pcr_oi < 0.7:
            st.success("🐂 Bullish Sentiment (Low PCR)")
        else:
            st.warning("⚖️ Neutral Sentiment")
    except:
        st.warning("PCR data not found")

# === OC Summary (Support / Resistance) ===
if "OC_1" in data_dict:
    df = data_dict["OC_1"].dropna()
    if "Strike" in df.columns:
        ce_oi = df["CE_OI"].idxmax()
        pe_oi = df["PE_OI"].idxmax()
        support = df.loc[pe_oi, "Strike"]
        resistance = df.loc[ce_oi, "Strike"]

        st.subheader("📌 Nifty Support / Resistance")
        col1, col2 = st.columns(2)
        col1.success(f"Support: {support}")
        col2.error(f"Resistance: {resistance}")

# === Dashboard: Top Movers ===
if "Dashboard" in data_dict:
    st.subheader("🔥 Top 10 Movers in F&O")
    dash_df = data_dict["Dashboard"]

    tabs = st.tabs(["Top OI Gainers","Top OI Losers","Top Price Gainers","Top Price Losers"])
    try:
        with tabs[0]:
            st.dataframe(dash_df.head(10))
        with tabs[1]:
            st.dataframe(dash_df.tail(10))
    except:
        st.warning("Dashboard format may differ")

# === Screener Data ===
if "Screener" in data_dict:
    st.subheader("📈 Screener (Stocks)")
    st.dataframe(data_dict["Screener"].head(20))

# === Sector Dashboard ===
if "Sector Dashboard" in data_dict:
    st.subheader("🏦 Sector Dashboard")
    st.dataframe(data_dict["Sector Dashboard"])
