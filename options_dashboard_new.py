import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import time
import numpy as np
from typing import Dict, Tuple, Optional
import openpyxl
from openpyxl import load_workbook
import os
import json

# Page Configuration
st.set_page_config(
    page_title="Live Option Chain Dashboard", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configuration
EXCEL_FILE = "Live_Option_Chain_Terminal.xlsx"  # Changed to .xlsx
REFRESH_INTERVAL = 5  # seconds
MAX_RETRIES = 3

# Initialize session state for data storage
if 'last_update' not in st.session_state:
    st.session_state.last_update = datetime.now()
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = {}
if 'connection_status' not in st.session_state:
    st.session_state.connection_status = False
if 'file_last_modified' not in st.session_state:
    st.session_state.file_last_modified = None

class ExcelDataManager:
    def __init__(self, excel_file: str):
        self.excel_file = excel_file
        self.connection_status = False
        self.last_successful_fetch = None
        
    def check_file_exists(self) -> bool:
        """Check if Excel file exists and is accessible"""
        try:
            return os.path.exists(self.excel_file) and os.path.isfile(self.excel_file)
        except Exception:
            return False
    
    def get_file_modified_time(self) -> Optional[float]:
        """Get file modification time"""
        try:
            return os.path.getmtime(self.excel_file)
        except Exception:
            return None
    
    def read_excel_sheet(self, sheet_name: str, max_rows: int = 100) -> pd.DataFrame:
        """Read Excel sheet using pandas and openpyxl"""
        if not self.check_file_exists():
            st.error(f"⚠️ Excel file {self.excel_file} not found!")
            return pd.DataFrame()
            
        for attempt in range(MAX_RETRIES):
            try:
                # Read Excel file with pandas
                df = pd.read_excel(
                    self.excel_file, 
                    sheet_name=sheet_name, 
                    engine='openpyxl',
                    nrows=max_rows
                )
                
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Clean the dataframe
                    df = df.dropna(how='all')
                    # Convert object columns to numeric where possible
                    for col in df.columns:
                        if df[col].dtype == 'object':
                            df[col] = pd.to_numeric(df[col], errors='ignore')
                    
                    self.connection_status = True
                    return df
                    
            except FileNotFoundError:
                st.error(f"⚠️ File {self.excel_file} not found. Please check the file path.")
                break
            except PermissionError:
                st.warning(f"⚠️ Permission denied. File may be open in Excel. Attempt {attempt + 1}/{MAX_RETRIES}")
                time.sleep(1)
            except Exception as e:
                if attempt == MAX_RETRIES - 1:
                    st.warning(f"⚠️ Failed to read sheet {sheet_name}: {str(e)}")
                time.sleep(0.5)
                
        self.connection_status = False
        return pd.DataFrame()
    
    def get_sheet_names(self) -> list:
        """Get all sheet names from Excel file"""
        try:
            xl_file = pd.ExcelFile(self.excel_file, engine='openpyxl')
            return xl_file.sheet_names
        except Exception:
            return []
    
    def fetch_all_data(self) -> Dict[str, pd.DataFrame]:
        """Fetch data from all relevant sheets"""
        if not self.check_file_exists():
            st.error(f"⚠️ Excel file {self.excel_file} not found!")
            return {}
        
        # Check if file has been modified
        current_modified_time = self.get_file_modified_time()
        if (current_modified_time and 
            st.session_state.file_last_modified and 
            current_modified_time == st.session_state.file_last_modified and
            st.session_state.excel_data):
            # File hasn't changed, use cached data
            return st.session_state.excel_data
        
        # Default sheets to fetch
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
        
        # Get available sheet names
        available_sheets = self.get_sheet_names()
        if available_sheets:
            st.sidebar.success(f"📋 Found {len(available_sheets)} sheets")
            # Filter to only existing sheets
            sheets_to_fetch = [sheet for sheet in sheets_to_fetch if sheet in available_sheets]
        
        data = {}
        successful_reads = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, sheet in enumerate(sheets_to_fetch):
            try:
                status_text.text(f"Loading {sheet}...")
                df = self.read_excel_sheet(sheet)
                if not df.empty:
                    data[sheet] = df
                    st.session_state.excel_data[sheet] = df
                    successful_reads += 1
                
                progress_bar.progress((i + 1) / len(sheets_to_fetch))
                
            except Exception as e:
                st.warning(f"Could not fetch {sheet}: {e}")
        
        progress_bar.empty()
        status_text.empty()
        
        if successful_reads > 0:
            self.last_successful_fetch = datetime.now()
            st.session_state.last_update = self.last_successful_fetch
            st.session_state.file_last_modified = current_modified_time
            st.success(f"✅ Successfully loaded {successful_reads} sheets")
        else:
            st.error("❌ Failed to load any data from Excel file")
            
        return data

# Initialize Excel Manager
@st.cache_resource
def get_excel_manager():
    return ExcelDataManager(EXCEL_FILE)

excel_manager = get_excel_manager()

def calculate_support_resistance_maxpain(df: pd.DataFrame) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    """Calculate support, resistance, and max pain levels"""
    try:
        # Check for required columns (case insensitive)
        df_columns = [col.upper() for col in df.columns]
        required_cols = ['STRIKE', 'CE_OI', 'PE_OI']
        
        # Find actual column names
        col_mapping = {}
        for req_col in required_cols:
            found = False
            for actual_col in df.columns:
                if req_col in actual_col.upper():
                    col_mapping[req_col] = actual_col
                    found = True
                    break
            if not found:
                return None, None, None
        
        # Clean and convert data
        df_clean = df.dropna(subset=list(col_mapping.values()))
        df_clean = df_clean[df_clean[col_mapping['STRIKE']] > 0]
        
        if df_clean.empty:
            return None, None, None
            
        # Support (highest PE OI)
        pe_oi_idx = df_clean[col_mapping['PE_OI']].idxmax()
        support = df_clean.loc[pe_oi_idx, col_mapping['STRIKE']]
        
        # Resistance (highest CE OI)
        ce_oi_idx = df_clean[col_mapping['CE_OI']].idxmax()
        resistance = df_clean.loc[ce_oi_idx, col_mapping['STRIKE']]
        
        # Max Pain calculation
        strikes = sorted(df_clean[col_mapping['STRIKE']].unique())
        total_pain = {}
        
        for strike in strikes:
            call_pain = 0
            put_pain = 0
            
            # Calculate call pain (ITM calls)
            itm_calls = df_clean[df_clean[col_mapping['STRIKE']] < strike]
            if not itm_calls.empty:
                call_pain = (itm_calls[col_mapping['CE_OI']] * (strike - itm_calls[col_mapping['STRIKE']])).sum()
            
            # Calculate put pain (ITM puts)
            itm_puts = df_clean[df_clean[col_mapping['STRIKE']] > strike]
            if not itm_puts.empty:
                put_pain = (itm_puts[col_mapping['PE_OI']] * (itm_puts[col_mapping['STRIKE']] - strike)).sum()
            
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
        # Try to find PCR values in different column positions
        pcr_oi = None
        pcr_vol = None
        
        # Look for PCR columns
        for i, col in enumerate(pcr_df.columns):
            if 'PCR' in str(col).upper() and 'OI' in str(col).upper():
                pcr_oi = pd.to_numeric(pcr_df.iloc[0, i], errors='coerce')
            elif 'PCR' in str(col).upper() and 'VOL' in str(col).upper():
                pcr_vol = pd.to_numeric(pcr_df.iloc[0, i], errors='coerce')
        
        # Fallback to column positions
        if pcr_oi is None and len(pcr_df.columns) >= 3:
            pcr_oi = pd.to_numeric(pcr_df.iloc[0, 2], errors='coerce')
        if pcr_vol is None and len(pcr_df.columns) >= 4:
            pcr_vol = pd.to_numeric(pcr_df.iloc[0, 3], errors='coerce')
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if pcr_oi is not None:
                st.metric("⚖️ PCR (OI)", f"{pcr_oi:.3f}")
                if pcr_oi > 1.2:
                    st.success("🐻 Bearish Sentiment")
                elif pcr_oi < 0.8:
                    st.error("🐂 Bullish Sentiment")
                else:
                    st.info("😐 Neutral Sentiment")
            else:
                st.warning("PCR OI data not found")
        
        with col2:
            if pcr_vol is not None:
                st.metric("📊 PCR (Vol)", f"{pcr_vol:.3f}")
            else:
                st.warning("PCR Volume data not found")
        
        with col3:
            st.metric("🕐 Last Update", 
                     st.session_state.last_update.strftime('%H:%M:%S'))
                     
        # Display raw PCR data
        with st.expander("📋 Raw PCR Data"):
            st.dataframe(pcr_df, use_container_width=True)
            
    except Exception as e:
        st.error(f"Error displaying PCR data: {e}")
        st.dataframe(pcr_df, use_container_width=True)

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
        else:
            st.info("🟢 Support: N/A")
    with level_cols[1]:
        if resistance:
            st.error(f"🔴 Resistance: {resistance:.0f}")
        else:
            st.info("🔴 Resistance: N/A")
    with level_cols[2]:
        if max_pain:
            st.info(f"💰 Max Pain: {max_pain:.0f}")
        else:
            st.info("💰 Max Pain: N/A")
    
    # Display option chain data with formatting
    try:
        display_df = df.copy()
        
        # Format numeric columns
        for col in display_df.columns:
            if display_df[col].dtype in ['float64', 'int64']:
                if 'OI' in col.upper() or 'VOLUME' in col.upper():
                    display_df[col] = display_df[col].apply(
                        lambda x: f"{x:,.0f}" if pd.notnull(x) and x != 0 else ""
                    )
                elif any(term in col.upper() for term in ['IV', 'PRICE', 'LTP', 'PREMIUM']):
                    display_df[col] = display_df[col].apply(
                        lambda x: f"{x:.2f}" if pd.notnull(x) else ""
                    )
        
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
        for col in display_df.columns:
            if display_df[col].dtype in ['float64', 'int64']:
                max_val = display_df[col].max()
                if pd.notnull(max_val):
                    if max_val > 10000000:  # 10M+
                        display_df[col] = display_df[col].apply(
                            lambda x: f"{x/10000000:.1f}Cr" if pd.notnull(x) and x != 0 else ""
                        )
                    elif max_val > 1000000:  # 1M+
                        display_df[col] = display_df[col].apply(
                            lambda x: f"{x/1000000:.2f}M" if pd.notnull(x) and x != 0 else ""
                        )
                    elif max_val > 1000:  # 1K+
                        display_df[col] = display_df[col].apply(
                            lambda x: f"{x/1000:.1f}K" if pd.notnull(x) and x != 0 else ""
                        )
    except:
        pass
    
    st.dataframe(display_df, use_container_width=True, height=300)

# File upload option for testing
st.sidebar.title("🎛️ Dashboard Controls")

# File upload option
uploaded_file = st.sidebar.file_uploader(
    "📂 Upload Excel File (Optional)", 
    type=['xlsx', 'xlsm'],
    help="Upload your Excel file if not in the same directory"
)

if uploaded_file:
    # Save uploaded file temporarily
    with open("temp_" + uploaded_file.name, "wb") as f:
        f.write(uploaded_file.getbuffer())
    EXCEL_FILE = "temp_" + uploaded_file.name
    excel_manager.excel_file = EXCEL_FILE

# Auto-refresh toggle
auto_refresh = st.sidebar.checkbox("🔄 Auto Refresh", value=True)
refresh_interval = st.sidebar.slider("Refresh Interval (seconds)", 1, 60, REFRESH_INTERVAL)

# Manual refresh button
if st.sidebar.button("🔄 Refresh Now", type="primary"):
    with st.spinner("Fetching latest data..."):
        # Clear cached data
        st.session_state.excel_data = {}
        st.session_state.file_last_modified = None
        excel_manager.fetch_all_data()
    st.rerun()

# File status
file_exists = excel_manager.check_file_exists()
st.sidebar.metric("📄 File Status", 
                 "🟢 Found" if file_exists else "🔴 Not Found")

# Connection status
st.sidebar.metric("📶 Data Status", 
                 "🟢 Loaded" if st.session_state.excel_data else "🔴 No Data")

if st.session_state.last_update:
    st.sidebar.metric("⏰ Last Update", 
                     st.session_state.last_update.strftime('%H:%M:%S'))

# Show available sheets
if st.session_state.excel_data:
    with st.sidebar.expander("📋 Available Sheets"):
        for sheet in st.session_state.excel_data.keys():
            st.write(f"✅ {sheet}")

# Main Dashboard
st.title("⚡ Live Options Summary Dashboard")
st.caption(f"Reading from: {EXCEL_FILE} | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# Initial data load
if not st.session_state.excel_data:
    with st.spinner("Loading initial data..."):
        excel_manager.fetch_all_data()

# Display PCR Data
if "PCR & OI Chart" in st.session_state.excel_data:
    display_pcr_data(st.session_state.excel_data["PCR & OI Chart"])
else:
    st.info("📊 PCR data not available")

st.divider()

# Display Option Chains
st.header("📌 Index Options Summary")

col1, col2 = st.columns(2)

with col1:
    if "OC_1" in st.session_state.excel_data:
        display_option_chain(st.session_state.excel_data["OC_1"], "Nifty Options")
    else:
        st.info("📈 Nifty options data not available")

with col2:
    if "OC_2" in st.session_state.excel_data:
        display_option_chain(st.session_state.excel_data["OC_2"], "BankNifty Options")
    else:
        st.info("🏦 BankNifty options data not available")

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
if auto_refresh and file_exists:
    time.sleep(refresh_interval)
    # Check if file was modified
    current_time = excel_manager.get_file_modified_time()
    if (current_time and 
        st.session_state.file_last_modified and 
        current_time > st.session_state.file_last_modified):
        st.rerun()

# Footer
st.markdown("---")
st.markdown("📊 **Live Options Dashboard** | No xlwings required - Direct Excel file reading")
st.markdown("💡 **Tip**: Save your Excel file frequently to see live updates")
