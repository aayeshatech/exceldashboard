import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import xlwings as xw
import time
import numpy as np
from typing import Dict, Tuple, Optional
import threading
import json

# Page Configuration
st.set_page_config(
    page_title="Live Option Chain Dashboard", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configuration
EXCEL_FILE = "Live_Option_Chain_Terminal.xlsm"
REFRESH_INTERVAL = 5  # seconds
MAX_RETRIES = 3

# Initialize session state for data storage
if 'last_update' not in st.session_state:
    st.session_state.last_update = datetime.now()
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = {}
if 'connection_status' not in st.session_state:
    st.session_state.connection_status = False

class ExcelDataManager:
    def __init__(self, excel_file: str):
        self.excel_file = excel_file
        self.wb = None
        self.connection_status = False
        self.last_successful_fetch = None
        
    def connect_to_excel(self) -> bool:
        """Establish connection to Excel workbook"""
        try:
            self.wb = xw.Book(self.excel_file)
            self.connection_status = True
            return True
        except Exception as e:
            st.error(f"⚠️ Cannot connect to {self.excel_file}. Please ensure it's open in Excel.\n\nError: {e}")
            self.connection_status = False
            return False
    
    def get_sheet_data(self, sheet_name: str, max_rows: int = 100) -> pd.DataFrame:
        """Fetch data from Excel sheet with error handling"""
        if not self.connection_status:
            return pd.DataFrame()
            
        for attempt in range(MAX_RETRIES):
            try:
                sht = self.wb.sheets[sheet_name]
                # Get the used range to avoid reading empty cells
                used_range = sht.used_range
                if used_range is None:
                    return pd.DataFrame()
                
                # Read data with proper options
                df = sht.range("A1").expand().options(
                    pd.DataFrame, 
                    header=1, 
                    index=False,
                    numbers=int
                ).value
                
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Clean the dataframe
                    df = df.dropna(how='all').head(max_rows)
                    # Convert numeric columns
                    for col in df.select_dtypes(include=['object']).columns:
                        df[col] = pd.to_numeric(df[col], errors='ignore')
                    return df
                    
            except Exception as e:
                if attempt == MAX_RETRIES - 1:
                    st.warning(f"⚠️ Failed to read sheet {sheet_name} after {MAX_RETRIES} attempts: {e}")
                time.sleep(0.5)
                
        return pd.DataFrame()
    
    def fetch_all_data(self) -> Dict[str, pd.DataFrame]:
        """Fetch data from all relevant sheets"""
        sheets_to_fetch = [
            "PCR & OI Chart",
            "OC_1",  # Nifty
            "OC_2",  # BankNifty
            "Dashboard",
            "Sector Dashboard", 
            "Screener",
            "FII DII Data",
            "Fiis&Diis Dashboard",
            "Globlemarket"
        ]
        
        data = {}
        if self.connect_to_excel():
            for sheet in sheets_to_fetch:
                try:
                    df = self.get_sheet_data(sheet)
                    if not df.empty:
                        data[sheet] = df
                        st.session_state.excel_data[sheet] = df
                except Exception as e:
                    st.warning(f"Could not fetch {sheet}: {e}")
                    
            self.last_successful_fetch = datetime.now()
            st.session_state.last_update = self.last_successful_fetch
            
        return data

# Initialize Excel Manager
@st.cache_resource
def get_excel_manager():
    return ExcelDataManager(EXCEL_FILE)

excel_manager = get_excel_manager()

def calculate_support_resistance_maxpain(df: pd.DataFrame) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    """Calculate support, resistance, and max pain levels"""
    try:
        if not {"Strike", "CE_OI", "PE_OI"}.issubset(df.columns):
            return None, None, None
            
        # Clean and convert data
        df_clean = df.dropna(subset=["Strike", "CE_OI", "PE_OI"])
        df_clean = df_clean[df_clean["Strike"] > 0]  # Remove invalid strikes
        
        if df_clean.empty:
            return None, None, None
            
        # Support (highest PE OI)
        pe_oi_idx = df_clean["PE_OI"].idxmax()
        support = df_clean.loc[pe_oi_idx, "Strike"]
        
        # Resistance (highest CE OI)
        ce_oi_idx = df_clean["CE_OI"].idxmax()
        resistance = df_clean.loc[ce_oi_idx, "Strike"]
        
        # Max Pain calculation
        strikes = sorted(df_clean["Strike"].unique())
        total_pain = {}
        
        for strike in strikes:
            call_pain = 0
            put_pain = 0
            
            # Calculate call pain (ITM calls)
            itm_calls = df_clean[df_clean["Strike"] < strike]
            if not itm_calls.empty:
                call_pain = (itm_calls["CE_OI"] * (strike - itm_calls["Strike"])).sum()
            
            # Calculate put pain (ITM puts)
            itm_puts = df_clean[df_clean["Strike"] > strike]
            if not itm_puts.empty:
                put_pain = (itm_puts["PE_OI"] * (itm_puts["Strike"] - strike)).sum()
            
            total_pain[strike] = call_pain + put_pain
        
        max_pain = min(total_pain, key=total_pain.get) if total_pain else None
        
        return float(support), float(resistance), float(max_pain) if max_pain else None
        
    except Exception as e:
        st.error(f"Error calculating levels: {e}")
        return None, None, None

def display_pcr_data(pcr_df: pd.DataFrame):
    """Display PCR data with enhanced formatting"""
    if pcr_df.empty:
        st.warning("No PCR data available")
        return
        
    st.header("📊 Market Sentiment (PCR)")
    
    try:
        # Try different column positions for PCR values
        pcr_oi = None
        pcr_vol = None
        
        if len(pcr_df.columns) >= 4:
            pcr_oi = pd.to_numeric(pcr_df.iloc[0, 2], errors='coerce')
            pcr_vol = pd.to_numeric(pcr_df.iloc[0, 3], errors='coerce')
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if pcr_oi is not None:
                color = "normal" if 0.8 <= pcr_oi <= 1.2 else "inverse"
                st.metric("⚖️ PCR (OI)", f"{pcr_oi:.3f}", delta=None)
                if pcr_oi > 1.2:
                    st.success("🐻 Bearish Sentiment")
                elif pcr_oi < 0.8:
                    st.error("🐂 Bullish Sentiment")
                else:
                    st.info("😐 Neutral Sentiment")
        
        with col2:
            if pcr_vol is not None:
                st.metric("📊 PCR (Vol)", f"{pcr_vol:.3f}", delta=None)
        
        with col3:
            st.metric("🕐 Last Update", 
                     st.session_state.last_update.strftime('%H:%M:%S'))
                     
        # Display raw PCR data
        with st.expander("📋 Raw PCR Data"):
            st.dataframe(pcr_df, use_container_width=True)
            
    except Exception as e:
        st.error(f"Error displaying PCR data: {e}")

def display_option_chain(df: pd.DataFrame, title: str):
    """Display option chain with key levels"""
    if df.empty:
        st.warning(f"No {title} data available")
        return
        
    st.subheader(title)
    
    # Calculate key levels
    support, resistance, max_pain = calculate_support_resistance_maxpain(df)
    
    # Display key levels
    level_cols = st.columns(3)
    with level_cols[0]:
        if support:
            st.success(f"🟢 Support: {support:.0f}")
    with level_cols[1]:
        if resistance:
            st.error(f"🔴 Resistance: {resistance:.0f}")
    with level_cols[2]:
        if max_pain:
            st.info(f"💰 Max Pain: {max_pain:.0f}")
    
    # Display option chain data
    try:
        # Format numeric columns
        display_df = df.copy()
        numeric_cols = display_df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            if 'OI' in col or 'Volume' in col:
                display_df[col] = display_df[col].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else "")
            elif 'IV' in col or 'Price' in col or 'LTP' in col:
                display_df[col] = display_df[col].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "")
        
        st.dataframe(display_df, use_container_width=True, height=400)
        
    except Exception as e:
        st.dataframe(df, use_container_width=True, height=400)

def display_general_data(df: pd.DataFrame, title: str):
    """Display general data sheets"""
    if df.empty:
        return
        
    st.header(f"📊 {title}")
    
    # Format dataframe for better display
    display_df = df.copy()
    
    # Format numeric columns
    try:
        for col in display_df.select_dtypes(include=[np.number]).columns:
            if display_df[col].max() > 1000000:
                display_df[col] = display_df[col].apply(lambda x: f"{x/1000000:.2f}M" if pd.notnull(x) and x != 0 else "")
            elif display_df[col].max() > 1000:
                display_df[col] = display_df[col].apply(lambda x: f"{x/1000:.1f}K" if pd.notnull(x) and x != 0 else "")
    except:
        pass
    
    st.dataframe(display_df, use_container_width=True, height=300)

# Sidebar Controls
st.sidebar.title("🎛️ Dashboard Controls")

# Auto-refresh toggle
auto_refresh = st.sidebar.checkbox("🔄 Auto Refresh", value=True)
refresh_interval = st.sidebar.slider("Refresh Interval (seconds)", 1, 60, REFRESH_INTERVAL)

# Manual refresh button
if st.sidebar.button("🔄 Refresh Now", type="primary"):
    with st.spinner("Fetching latest data..."):
        excel_manager.fetch_all_data()
    st.rerun()

# Connection status
connection_status = excel_manager.connection_status
st.sidebar.metric("📶 Connection Status", 
                 "🟢 Connected" if connection_status else "🔴 Disconnected")

if st.session_state.last_update:
    st.sidebar.metric("⏰ Last Update", 
                     st.session_state.last_update.strftime('%H:%M:%S'))

# Main Dashboard
st.title("⚡ Live Options Summary Dashboard")
st.caption(f"Connected to: {EXCEL_FILE} | Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# Auto-refresh mechanism
if auto_refresh:
    placeholder = st.empty()
    
    # Fetch initial data
    if not st.session_state.excel_data or (datetime.now() - st.session_state.last_update).seconds > refresh_interval:
        with st.spinner("Loading data..."):
            excel_manager.fetch_all_data()

# Display PCR Data
if "PCR & OI Chart" in st.session_state.excel_data:
    display_pcr_data(st.session_state.excel_data["PCR & OI Chart"])

st.divider()

# Display Option Chains
st.header("📌 Index Options Summary")

col1, col2 = st.columns(2)

with col1:
    if "OC_1" in st.session_state.excel_data:
        display_option_chain(st.session_state.excel_data["OC_1"], "Nifty Options")

with col2:
    if "OC_2" in st.session_state.excel_data:
        display_option_chain(st.session_state.excel_data["OC_2"], "BankNifty Options")

st.divider()

# Display Other Data Sheets
other_sheets = [
    "Dashboard",
    "Sector Dashboard", 
    "Screener",
    "FII DII Data",
    "Fiis&Diis Dashboard",
    "Globlemarket"
]

for sheet in other_sheets:
    if sheet in st.session_state.excel_data:
        display_general_data(st.session_state.excel_data[sheet], sheet)
        st.divider()

# Auto-refresh logic
if auto_refresh:
    time.sleep(refresh_interval)
    st.rerun()

# Footer
st.markdown("---")
st.markdown("📊 **Live Options Dashboard** | Data refreshed every few seconds from live Excel")
st.markdown("⚠️ **Note**: Ensure the Excel file is open and running for real-time data updates")
