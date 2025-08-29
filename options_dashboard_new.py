import pandas as pd
import streamlit as st
from datetime import datetime
import xlwings as xw

st.set_page_config(page_title="Live Option Chain Dashboard", page_icon="📊", layout="wide")

# === Connect to running Excel workbook ===
# Make sure the Excel file is OPEN with live data running
EXCEL_FILE = "Live_Option_Chain_Terminal.xlsm"

try:
    wb = xw.Book(EXCEL_FILE)
except Exception as e:
    st.error(f"⚠️ Could not connect to Excel file.\n\nPlease open {EXCEL_FILE} in Excel first.\n\nError: {e}")
    st.stop()

# === Utility to read sheet into DataFrame ===
def get_sheet_data(sheet_name):
    try:
        sht = wb.sheets[sheet_name]
        df = sht.range("A1").expand().options(pd.DataFrame, header=1, index=False).value
        return df
    except Exception as e:
        st.warning(f"⚠️ Could not read sheet {sheet_name}: {e}")
        return pd.DataFrame()

# === Page Title ===
st.title("⚡ Live Options Summary Dashboard (Direct from Excel)")
st.caption(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# === Market Sentiment (PCR & OI Chart) ===
pcr_df = get_sheet_data("PCR & OI Chart")
if not pcr_df.empty:
    st.header("📊 Market Sentiment")
    try:
        pcr_oi = float(pcr_df.iloc[0,2])
        pcr_vol = float(pcr_df.iloc[0,3])
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

# === Support / Resistance / Max Pain ===
def support_resistance_maxpain(df):
    try:
        if {"Strike","CE_OI","PE_OI"}.issubset(df.columns):
            ce_oi_idx = df["CE_OI"].idxmax()
            pe_oi_idx = df["PE_OI"].idxmax()
            support = df.loc[pe_oi_idx, "Strike"]
            resistance = df.loc[ce_oi_idx, "Strike"]

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

nifty_df = get_sheet_data("OC_1")
if not nifty_df.empty:
    support, resistance, maxpain = support_resistance_maxpain(nifty_df)
    with col1:
        st.subheader("Nifty Options")
        if support and resistance:
            st.success(f"🟢 Support: {support}")
            st.error(f"🔴 Resistance: {resistance}")
            st.info(f"💰 Max Pain: {maxpain}")
        st.dataframe(nifty_df.head(15))

bnf_df = get_sheet_data("OC_2")
if not bnf_df.empty:
    support, resistance, maxpain = support_resistance_maxpain(bnf_df)
    with col2:
        st.subheader("BankNifty Options")
        if support and resistance:
            st.success(f"🟢 Support: {support}")
            st.error(f"🔴 Resistance: {resistance}")
            st.info(f"💰 Max Pain: {maxpain}")
        st.dataframe(bnf_df.head(15))

# === Dashboard Movers ===
dash_df = get_sheet_data("Dashboard")
if not dash_df.empty:
    st.header("🔥 Top Movers in F&O")
    tabs = st.tabs(["Top OI Gainers","Top OI Losers","Top Price Gainers","Top Price Losers"])
    try:
        with tabs[0]:
            st.dataframe(dash_df.head(10))
        with tabs[1]:
            st.dataframe(dash_df.tail(10))
    except:
        st.warning("⚠️ Dashboard format may differ")

# === Sector Dashboard ===
sector_df = get_sheet_data("Sector Dashboard")
if not sector_df.empty:
    st.header("🏦 Sector View")
    st.dataframe(sector_df.head(20))

# === Screener ===
screener_df = get_sheet_data("Screener")
if not screener_df.empty:
    st.header("📈 Screener (Stocks)")
    st.dataframe(screener_df.head(20))

# === FII / DII Data ===
fiidii_df = get_sheet_data("FII DII Data")
if not fiidii_df.empty:
    st.header("🏦 Institutional Activity (FII / DII)")
    st.dataframe(fiidii_df.tail(10))

fiidii_dash = get_sheet_data("Fiis&Diis Dashboard")
if not fiidii_dash.empty:
    st.header("📊 FII / DII Dashboard")
    st.dataframe(fiidii_dash.head(20))

# === Global Market Snapshot ===
gm_df = get_sheet_data("Globlemarket")
if not gm_df.empty:
    st.header("🌍 Global Market Snapshot (Dow, Gold, Silver, Crude, BTC)")
    st.dataframe(gm_df.head(20))
