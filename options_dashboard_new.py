import pandas as pd
import streamlit as st
from datetime import datetime
from streamlit_autorefresh import st_autorefresh

st.set_page_config(page_title="Live Option Chain Dashboard", page_icon="📊", layout="wide")

# === Auto Refresh (every 30 sec) ===
if st.sidebar.checkbox("🔄 Auto Refresh (30 sec)", value=True):
    st_autorefresh(interval=30000, key="refresh")

# === File Path ===
file_path = "Live_Option_Chain_Terminal.xlsm"

@st.cache_data(ttl=30)
def load_excel_data():
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet_targets = [
        "OC_1","OC_2",
        "Dashboard","Screener","Sector Dashboard",
        "PCR & OI Chart","FII DII Data","Fiis&Diis Dashboard"
    ]
    data = {}
    for sheet in sheet_targets:
        if sheet in xls.sheet_names:
            try:
                data[sheet] = pd.read_excel(xls, sheet)
            except Exception as e:
                st.warning(f"⚠️ Could not load {sheet}: {e}")
    return data

data_dict = load_excel_data()

st.title("⚡ Live Options Summary Dashboard")
st.caption(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# === Market Sentiment (PCR & OI Chart) ===
if "PCR & OI Chart" in data_dict:
    df_pcr = data_dict["PCR & OI Chart"]
    st.header("📊 Market Sentiment")
    try:
        pcr_oi = float(df_pcr.iloc[0,2])
        pcr_vol = float(df_pcr.iloc[0,3])

        col1, col2 = st.columns(2)
        col1.metric("⚖️ PCR (OI)", f"{pcr_oi:.2f}")
        col2.metric("📊 PCR (Vol)", f"{pcr_vol:.2f}")

        if pcr_oi > 1.3:
            st.error("🐻 Bearish Sentiment (High PCR)")
        elif pcr_oi < 0.7:
            st.success("🐂 Bullish Sentiment (Low PCR)")
        else:
            st.warning("⚖️ Neutral Sentiment")
    except Exception as e:
        st.warning(f"⚠️ PCR format issue: {e}")

# === Support / Resistance & Max Pain ===
def support_resistance_maxpain(df):
    try:
        if {"Strike","CE_OI","PE_OI"}.issubset(df.columns):
            ce_oi_idx = df["CE_OI"].idxmax()
            pe_oi_idx = df["PE_OI"].idxmax()
            support = df.loc[pe_oi_idx, "Strike"]
            resistance = df.loc[ce_oi_idx, "Strike"]

            # Max Pain
            strikes = df["Strike"].dropna().unique()
            total_pain = {}
            for strike in strikes:
                call_pain = (df.loc[df["Strike"] < strike, "CE_OI"] * (strike - df.loc[df["Strike"] < strike, "Strike"])).sum()
                put_pain  = (df.loc[df["Strike"] > strike, "PE_OI"] * (df.loc[df["Strike"] > strike, "Strike"] - strike)).sum()
                total_pain[strike] = call_pain + put_pain
            max_pain = min(total_pain, key=total_pain.get) if total_pain else None

            return support, resistance, max_pain
    except:
        return None, None, None
    return None, None, None

# === Index Summary ===
st.header("📌 Index Summary")

col1, col2 = st.columns(2)

if "OC_1" in data_dict:
    support, resistance, maxpain = support_resistance_maxpain(data_dict["OC_1"])
    with col1:
        st.subheader("Nifty Options")
        if support and resistance:
            st.success(f"🟢 Support: {support}")
            st.error(f"🔴 Resistance: {resistance}")
            st.info(f"💰 Max Pain: {maxpain}")
        st.dataframe(data_dict["OC_1"].head(15))

if "OC_2" in data_dict:
    support, resistance, maxpain = support_resistance_maxpain(data_dict["OC_2"])
    with col2:
        st.subheader("BankNifty Options")
        if support and resistance:
            st.success(f"🟢 Support: {support}")
            st.error(f"🔴 Resistance: {resistance}")
            st.info(f"💰 Max Pain: {maxpain}")
        st.dataframe(data_dict["OC_2"].head(15))

# === Dashboard Movers ===
if "Dashboard" in data_dict:
    st.header("🔥 Top Movers in F&O")
    tabs = st.tabs(["Top OI Gainers","Top OI Losers","Top Price Gainers","Top Price Losers"])
    dash_df = data_dict["Dashboard"]
    try:
        with tabs[0]:
            st.dataframe(dash_df.head(10))
        with tabs[1]:
            st.dataframe(dash_df.tail(10))
    except:
        st.warning("⚠️ Dashboard format may differ")

# === Sector Dashboard ===
if "Sector Dashboard" in data_dict:
    st.header("🏦 Sector View")
    st.dataframe(data_dict["Sector Dashboard"].head(20))

# === Screener ===
if "Screener" in data_dict:
    st.header("📈 Screener (Stocks)")
    st.dataframe(data_dict["Screener"].head(20))

# === FII / DII Data ===
if "FII DII Data" in data_dict:
    st.header("🏦 Institutional Activity (FII / DII)")
    st.dataframe(data_dict["FII DII Data"].tail(10))

if "Fiis&Diis Dashboard" in data_dict:
    st.header("📊 FII / DII Dashboard")
    st.dataframe(data_dict["Fiis&Diis Dashboard"].head(20))
