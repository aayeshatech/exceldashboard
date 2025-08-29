import pandas as pd
import streamlit as st
from datetime import datetime
import xlwings as xw

st.set_page_config(page_title="Live Option Chain Dashboard", page_icon="📊", layout="wide")

EXCEL_FILE = "Live_Option_Chain_Terminal.xlsm"  # must be OPEN in Excel

# Connect to running Excel
try:
    wb = xw.Book(EXCEL_FILE)
except Exception as e:
    st.error(f"⚠️ Please open {EXCEL_FILE} in Excel first.\n\nError: {e}")
    st.stop()

# Utility to read sheet into DataFrame
def get_sheet_data(sheet_name, rows=30):
    try:
        sht = wb.sheets[sheet_name]
        df = sht.range("A1").expand().options(pd.DataFrame, header=1, index=False).value
        return df.head(rows) if isinstance(df, pd.DataFrame) else pd.DataFrame()
    except Exception as e:
        st.warning(f"⚠️ Could not read sheet {sheet_name}: {e}")
        return pd.DataFrame()

st.title("⚡ Live Options Summary Dashboard (Excel Connected)")
st.caption(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# PCR
pcr_df = get_sheet_data("PCR & OI Chart")
if not pcr_df.empty:
    st.header("📊 Market Sentiment (PCR)")
    try:
        pcr_oi = float(pcr_df.iloc[0,2])
        pcr_vol = float(pcr_df.iloc[0,3])
        col1, col2 = st.columns(2)
        col1.metric("⚖️ PCR (OI)", f"{pcr_oi:.2f}")
        col2.metric("📊 PCR (Vol)", f"{pcr_vol:.2f}")
    except:
        st.warning("⚠️ PCR values not found")

# Support / Resistance / Max Pain
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

# Index Summary
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
        st.dataframe(nifty_df)

bnf_df = get_sheet_data("OC_2")
if not bnf_df.empty:
    support, resistance, maxpain = support_resistance_maxpain(bnf_df)
    with col2:
        st.subheader("BankNifty Options")
        if support and resistance:
            st.success(f"🟢 Support: {support}")
            st.error(f"🔴 Resistance: {resistance}")
            st.info(f"💰 Max Pain: {maxpain}")
        st.dataframe(bnf_df)

# Other Sheets
for sheet in ["Dashboard","Sector Dashboard","Screener","FII DII Data","Fiis&Diis Dashboard","Globlemarket"]:
    df = get_sheet_data(sheet, 20)
    if not df.empty:
        st.header(f"📊 {sheet}")
        st.dataframe(df)
