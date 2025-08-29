import pandas as pd
import numpy as np
import streamlit as st
import warnings
import time
import subprocess
import sys
import io
import os
from datetime import datetime
import requests
from io import BytesIO
warnings.filterwarnings('ignore')

# Set page config
st.set_page_config(
    page_title="NSE Options Dashboard",
    page_icon="⚡",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
.main-header {
    font-size: 2.5rem;
    font-weight: bold;
    text-align: center;
    color: #FF6B35;
    margin-bottom: 1rem;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
}
.success-box {
    background-color: #d4edda;
    color: #155724;
    padding: 1rem;
    border-radius: 5px;
    border-left: 5px solid #28a745;
    margin: 1rem 0;
}
.warning-box {
    background-color: #fff3cd;
    color: #856404;
    padding: 1rem;
    border-radius: 5px;
    border-left: 5px solid #ffc107;
    margin: 1rem 0;
}
.error-box {
    background-color: #f8d7da;
    color: #721c24;
    padding: 1rem;
    border-radius: 5px;
    border-left: 5px solid #dc3545;
    margin: 1rem 0;
}
.info-box {
    background-color: #d1ecf1;
    color: #0c5460;
    padding: 1rem;
    border-radius: 5px;
    border-left: 5px solid #17a2b8;
    margin: 1rem 0;
}
.code-box {
    background-color: #f8f9fa;
    color: #212529;
    padding: 1rem;
    border-radius: 5px;
    font-family: monospace;
    margin: 1rem 0;
    border: 1px solid #dee2e6;
}
.refresh-button {
    background-color: #4CAF50;
    border: none;
    color: white;
    padding: 10px 20px;
    text-align: center;
    text-decoration: none;
    display: inline-block;
    font-size: 16px;
    margin: 4px 2px;
    cursor: pointer;
    border-radius: 5px;
}
.path-fix {
    background-color: #e9ecef;
    padding: 10px;
    border-radius: 5px;
    font-family: monospace;
    margin: 10px 0;
}
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'last_refresh' not in st.session_state:
    st.session_state.last_refresh = None
if 'data_dict' not in st.session_state:
    st.session_state.data_dict = {}
if 'excel_file_path' not in st.session_state:
    st.session_state.excel_file_path = ""
if 'refresh_interval' not in st.session_state:
    st.session_state.refresh_interval = 30
if 'auto_refresh' not in st.session_state:
    st.session_state.auto_refresh = False
if 'file_method' not in st.session_state:
    st.session_state.file_method = "path"

def install_excel_libraries():
    """Attempt to install Excel support libraries"""
    try:
        st.info("Attempting to install Excel support libraries...")
        
        # Install openpyxl and xlrd
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "xlrd"])
        
        st.success("✅ Excel libraries installed successfully!")
        return True
    except Exception as e:
        st.error(f"❌ Failed to install Excel libraries: {str(e)}")
        return False

def check_excel_libraries():
    """Check if Excel libraries are available"""
    try:
        import openpyxl
        import xlrd
        return True
    except ImportError:
        return False

def fix_windows_path(path):
    """Fix Windows path issues"""
    # Replace backslashes with forward slashes
    path = path.replace("\\", "/")
    
    # Handle drive letter (e.g., D: becomes /D/)
    if ":" in path and path.index(":") == 1:
        drive_letter = path[0]
        path = f"/{drive_letter}/" + path[3:]
    
    return path

def read_excel_with_pandas(file_path):
    """Try to read Excel file with pandas using different engines"""
    try:
        # First try to detect available engines
        available_engines = []
        
        try:
            import openpyxl
            available_engines.append('openpyxl')
        except ImportError:
            pass
            
        try:
            import xlrd
            available_engines.append('xlrd')
        except ImportError:
            pass
        
        if not available_engines:
            return {'excel_libraries_missing': True}
        
        st.info(f"📁 Available Excel engines: {', '.join(available_engines)}")
        
        # Try each available engine
        for engine in available_engines:
            try:
                excel_file = pd.ExcelFile(file_path, engine=engine)
                sheet_count = len(excel_file.sheet_names)
                st.info(f"📊 Found {sheet_count} sheets in Excel file using {engine}")
                
                data_dict = {}
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
                        if not df.empty:
                            data_dict[sheet_name] = df
                            st.success(f"✅ Loaded sheet: {sheet_name} ({len(df)} rows, {len(df.columns)} columns)")
                        else:
                            st.warning(f"⚠️ Sheet '{sheet_name}' is empty")
                    except Exception as sheet_error:
                        st.warning(f"⚠️ Could not load sheet '{sheet_name}': {str(sheet_error)}")
                        continue
                
                if data_dict:
                    return data_dict
                    
            except Exception as engine_error:
                st.warning(f"⚠️ Engine {engine} failed: {str(engine_error)}")
                continue
        
        # If all engines failed, try without specifying engine
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_count = len(excel_file.sheet_names)
            st.info(f"📊 Found {sheet_count} sheets in Excel file (auto engine)")
            
            data_dict = {}
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    if not df.empty:
                        data_dict[sheet_name] = df
                        st.success(f"✅ Loaded sheet: {sheet_name} ({len(df)} rows, {len(df.columns)} columns)")
                    else:
                        st.warning(f"⚠️ Sheet '{sheet_name}' is empty")
                except Exception as sheet_error:
                    st.warning(f"⚠️ Could not load sheet '{sheet_name}': {str(sheet_error)}")
                    continue
            
            if data_dict:
                return data_dict
        except Exception as e:
            st.error(f"❌ Auto engine also failed: {str(e)}")
        
        return {'excel_read_error': True}
        
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")
        return {'excel_read_error': True}

def load_excel_data(file_path):
    """Load data from Excel file path"""
    # Try original path first
    if os.path.exists(file_path):
        st.success(f"✅ File found at: {file_path}")
    else:
        # Try fixing Windows path
        fixed_path = fix_windows_path(file_path)
        if fixed_path != file_path and os.path.exists(fixed_path):
            st.info(f"🔄 Using fixed path: {fixed_path}")
            file_path = fixed_path
        else:
            st.error(f"❌ File not found: {file_path}")
            st.info("💡 Try using the file upload option instead")
            return {}
    
    file_extension = file_path.split('.')[-1].lower()
    
    if file_extension not in ['xlsx', 'xlsm', 'xls']:
        st.error(f"❌ Not an Excel file: {file_path}")
        return {}
    
    # Check if Excel libraries are available
    if not check_excel_libraries():
        return {'excel_libraries_missing': True}
    
    # Try to read Excel file using pandas
    data_dict = read_excel_with_pandas(file_path)
    return data_dict

def load_uploaded_file(uploaded_file):
    """Load data from uploaded file"""
    try:
        # Save uploaded file to a temporary location
        with open("temp_excel_file.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Load the saved file
        data_dict = load_excel_data("temp_excel_file.xlsx")
        
        # Clean up
        if os.path.exists("temp_excel_file.xlsx"):
            os.remove("temp_excel_file.xlsx")
            
        return data_dict
    except Exception as e:
        st.error(f"❌ Error processing uploaded file: {str(e)}")
        return {}

def fetch_data_from_excel():
    """Fetch data from the Excel file"""
    try:
        if st.session_state.file_method == "path" and st.session_state.excel_file_path:
            data_dict = load_excel_data(st.session_state.excel_file_path)
            if data_dict and 'excel_libraries_missing' not in data_dict and 'excel_read_error' not in data_dict:
                st.session_state.data_dict = data_dict
                st.session_state.last_refresh = datetime.now()
                st.success("✅ Data refreshed successfully!")
            else:
                st.error("❌ Failed to load data from Excel file")
        else:
            st.error("❌ No Excel file path specified")
    except Exception as e:
        st.error(f"❌ Error fetching data: {str(e)}")

def calculate_pcr(df):
    """Calculate Put-Call Ratio"""
    try:
        call_oi_cols = [col for col in df.columns if 'CE_OI' in col and 'Change' not in col]
        put_oi_cols = [col for col in df.columns if 'PE_OI' in col and 'Change' not in col]
        
        if call_oi_cols and put_oi_cols:
            total_call_oi = df[call_oi_cols[0]].fillna(0).sum()
            total_put_oi = df[put_oi_cols[0]].fillna(0).sum()
            pcr_oi = total_put_oi / total_call_oi if total_call_oi > 0 else 0
            return pcr_oi, total_call_oi, total_put_oi
        
        return None, None, None
    except Exception:
        return None, None, None

def calculate_volume_pcr(df):
    """Calculate Volume PCR"""
    try:
        call_vol_cols = [col for col in df.columns if 'CE_' in col and 'Volume' in col]
        put_vol_cols = [col for col in df.columns if 'PE_' in col and 'Volume' in col]
        
        if call_vol_cols and put_vol_cols:
            total_call_vol = df[call_vol_cols[0]].fillna(0).sum()
            total_put_vol = df[put_vol_cols[0]].fillna(0).sum()
            pcr_vol = total_put_vol / total_call_vol if total_call_vol > 0 else 0
            return pcr_vol, total_call_vol, total_put_vol
        
        return None, None, None
    except Exception:
        return None, None, None

def calculate_max_pain(df):
    """Calculate Max Pain"""
    try:
        strike_col = None
        for col in df.columns:
            if 'strike' in col.lower():
                strike_col = col
                break
        
        call_oi_col = None
        put_oi_col = None
        
        for col in df.columns:
            if 'CE_OI' in col and 'Change' not in col:
                call_oi_col = col
            if 'PE_OI' in col and 'Change' not in col:
                put_oi_col = col
        
        if strike_col and call_oi_col and put_oi_col:
            clean_df = df[[strike_col, call_oi_col, put_oi_col]].dropna()
            
            if len(clean_df) == 0:
                return None
            
            strikes = clean_df[strike_col].sort_values()
            total_pain = []
            
            for strike in strikes:
                call_pain = 0
                put_pain = 0
                
                for _, row in clean_df.iterrows():
                    if row[strike_col] < strike:
                        call_pain += row[call_oi_col] * (strike - row[strike_col])
                    if row[strike_col] > strike:
                        put_pain += row[put_oi_col] * (row[strike_col] - strike)
                
                total_pain.append(call_pain + put_pain)
            
            if total_pain:
                max_pain_index = np.argmin(total_pain)
                return strikes.iloc[max_pain_index]
        
        return None
    except Exception:
        return None

def get_support_resistance(df):
    """Get support and resistance levels"""
    try:
        strike_col = None
        for col in df.columns:
            if 'strike' in col.lower():
                strike_col = col
                break
        
        call_oi_col = None
        put_oi_col = None
        
        for col in df.columns:
            if 'CE_OI' in col and 'Change' not in col:
                call_oi_col = col
            if 'PE_OI' in col and 'Change' not in col:
                put_oi_col = col
        
        if strike_col and call_oi_col and put_oi_col:
            clean_df = df[[strike_col, call_oi_col, put_oi_col]].dropna()
            
            if len(clean_df) == 0:
                return None, None
            
            max_call_oi_idx = clean_df[call_oi_col].idxmax()
            max_put_oi_idx = clean_df[put_oi_col].idxmax()
            
            resistance = clean_df.loc[max_call_oi_idx, strike_col]
            support = clean_df.loc[max_put_oi_idx, strike_col]
            
            return support, resistance
        
        return None, None
    except Exception:
        return None, None

def display_market_sentiment(pcr_oi):
    """Display market sentiment based on PCR"""
    if pcr_oi is None:
        return
    
    if pcr_oi > 1.3:
        st.markdown(f"""
        <div class="error-box">
        <strong>🐻 BEARISH SENTIMENT</strong><br>
        PCR is high ({pcr_oi:.3f}) - More puts than calls, indicating bearish sentiment
        </div>
        """, unsafe_allow_html=True)
    elif pcr_oi < 0.7:
        st.markdown(f"""
        <div class="success-box">
        <strong>🐂 BULLISH SENTIMENT</strong><br>
        PCR is low ({pcr_oi:.3f}) - More calls than puts, indicating bullish sentiment
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="warning-box">
        <strong>⚖️ NEUTRAL SENTIMENT</strong><br>
        PCR is balanced ({pcr_oi:.3f}) - No clear directional bias
        </div>
        """, unsafe_allow_html=True)

def create_charts(df):
    """Create charts using Streamlit native functionality"""
    try:
        # Find relevant columns
        strike_col = None
        for col in df.columns:
            if 'strike' in col.lower():
                strike_col = col
                break
        
        if not strike_col:
            st.warning("No Strike column found for charting")
            return
        
        # OI Chart
        call_oi_cols = [col for col in df.columns if 'CE_OI' in col and 'Change' not in col]
        put_oi_cols = [col for col in df.columns if 'PE_OI' in col and 'Change' not in col]
        
        if call_oi_cols and put_oi_cols:
            chart_data = df[[strike_col, call_oi_cols[0], put_oi_cols[0]]].copy()
            chart_data = chart_data.dropna()
            
            if not chart_data.empty:
                chart_data = chart_data.set_index(strike_col)
                chart_data.columns = ['Call OI', 'Put OI']
                
                st.subheader("📊 Open Interest Distribution")
                st.bar_chart(chart_data, height=400)
        
        # Volume Chart
        call_vol_cols = [col for col in df.columns if 'CE_' in col and 'Volume' in col]
        put_vol_cols = [col for col in df.columns if 'PE_' in col and 'Volume' in col]
        
        if call_vol_cols and put_vol_cols:
            vol_data = df[[strike_col, call_vol_cols[0], put_vol_cols[0]]].copy()
            vol_data = vol_data.dropna()
            
            if not vol_data.empty:
                vol_data = vol_data.set_index(strike_col)
                vol_data.columns = ['Call Volume', 'Put Volume']
                
                st.subheader("📈 Volume Distribution")
                st.bar_chart(vol_data, height=400)
        
        # IV Chart
        call_iv_cols = [col for col in df.columns if 'CE_IV' in col]
        put_iv_cols = [col for col in df.columns if 'PE_IV' in col]
        
        if call_iv_cols and put_iv_cols:
            iv_data = df[[strike_col, call_iv_cols[0], put_iv_cols[0]]].copy()
            iv_data = iv_data.dropna()
            
            if not iv_data.empty:
                iv_data = iv_data.set_index(strike_col)
                iv_data.columns = ['Call IV', 'Put IV']
                
                st.subheader("📉 Implied Volatility")
                st.line_chart(iv_data, height=400)
    
    except Exception as e:
        st.warning(f"Could not create charts: {str(e)}")

def display_top_strikes(df):
    """Display top strikes by OI and Volume"""
    try:
        strike_col = None
        for col in df.columns:
            if 'strike' in col.lower():
                strike_col = col
                break
        
        if not strike_col:
            st.warning("No Strike column found")
            return
        
        call_oi_cols = [col for col in df.columns if 'CE_OI' in col and 'Change' not in col]
        put_oi_cols = [col for col in df.columns if 'PE_OI' in col and 'Change' not in col]
        call_vol_cols = [col for col in df.columns if 'CE_' in col and 'Volume' in col]
        put_vol_cols = [col for col in df.columns if 'PE_' in col and 'Volume' in col]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🟢 Top Call Activity")
            if call_oi_cols:
                display_cols = [strike_col, call_oi_cols[0]]
                if call_vol_cols:
                    display_cols.append(call_vol_cols[0])
                
                top_call = df[display_cols].nlargest(5, call_oi_cols[0])
                st.dataframe(top_call, hide_index=True)
        
        with col2:
            st.subheader("🔴 Top Put Activity")
            if put_oi_cols:
                display_cols = [strike_col, put_oi_cols[0]]
                if put_vol_cols:
                    display_cols.append(put_vol_cols[0])
                
                top_put = df[display_cols].nlargest(5, put_oi_cols[0])
                st.dataframe(top_put, hide_index=True)
    
    except Exception as e:
        st.warning(f"Could not display top strikes: {str(e)}")

def main():
    st.markdown('<h1 class="main-header">⚡ NSE Options Dashboard</h1>', unsafe_allow_html=True)
    
    # Current time
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Sidebar
    st.sidebar.header("📁 Data Source Configuration")
    
    # File method selection
    file_method = st.sidebar.radio(
        "Select data source method:",
        ["File Path", "File Upload"],
        index=0 if st.session_state.file_method == "path" else 1,
        help="Choose how to provide your Excel file"
    )
    
    st.session_state.file_method = "path" if file_method == "File Path" else "upload"
    
    if st.session_state.file_method == "path":
        # Excel file path input
        excel_file_path = st.sidebar.text_input(
            "Excel File Path",
            value=st.session_state.excel_file_path,
            help="Enter the full path to your Excel file (e.g., C:/Data/Live_Option_Chain_Terminal.xlsm)"
        )
        
        if excel_file_path != st.session_state.excel_file_path:
            st.session_state.excel_file_path = excel_file_path
            st.session_state.data_dict = {}  # Clear previous data
            
        # Show path help
        with st.sidebar.expander("💡 Path Help"):
            st.markdown("""
            **Windows Path Examples:**
            - `C:/Users/Name/Documents/file.xlsx` (recommended)
            - `C:\\Users\\Name\\Documents\\file.xlsx` (escape backslashes)
            - `/c/Users/Name/Documents/file.xlsx` (Git Bash style)
            """)
            
    else:
        # File uploader
        uploaded_file = st.sidebar.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xlsm', 'xls'],
            help="Upload your Excel file directly"
        )
        
        if uploaded_file is not None:
            with st.spinner("Loading uploaded file..."):
                data_dict = load_uploaded_file(uploaded_file)
                if data_dict and 'excel_libraries_missing' not in data_dict and 'excel_read_error' not in data_dict:
                    st.session_state.data_dict = data_dict
                    st.session_state.last_refresh = datetime.now()
                    st.sidebar.success("✅ File uploaded successfully!")
                else:
                    st.sidebar.error("❌ Failed to load uploaded file")
    
    # Refresh settings
    st.sidebar.header("🔄 Refresh Settings")
    st.session_state.refresh_interval = st.sidebar.number_input(
        "Refresh Interval (seconds)", 
        min_value=5, 
        max_value=300, 
        value=30,
        help="How often to automatically refresh data from Excel"
    )
    
    st.session_state.auto_refresh = st.sidebar.checkbox(
        "Auto Refresh", 
        value=st.session_state.auto_refresh,
        help="Automatically refresh data at the specified interval"
    )
    
    # Manual refresh button
    if st.sidebar.button("🔄 Manual Refresh", type="primary"):
        if st.session_state.file_method == "path":
            fetch_data_from_excel()
        else:
            st.sidebar.info("Please upload a file first")
    
    # File format recommendation
    st.sidebar.markdown("""
    **📊 Excel File Requirements:**
    - ✅ File must be accessible from this app
    - ✅ File should not be locked by another process
    - ✅ File should contain options chain data
    - ✅ Supported formats: .xlsx, .xlsm, .xls
    """)
    
    # Display last refresh time
    if st.session_state.last_refresh:
        last_refresh_str = st.session_state.last_refresh.strftime("%Y-%m-%d %H:%M:%S")
        st.markdown(f"<div class='info-box'><strong>🕐 Last Refresh:</strong> {last_refresh_str}</div>", unsafe_allow_html=True)
    else:
        st.markdown(f"<div class='info-box'><strong>🕐 Current Time:</strong> {current_time}</div>", unsafe_allow_html=True)
    
    # Auto-refresh logic
    if st.session_state.auto_refresh and st.session_state.file_method == "path" and st.session_state.excel_file_path:
        refresh_interval = st.session_state.refresh_interval
        time_since_refresh = (datetime.now() - st.session_state.last_refresh).total_seconds() if st.session_state.last_refresh else refresh_interval + 1
        
        if time_since_refresh >= refresh_interval:
            fetch_data_from_excel()
            st.rerun()
        else:
            time_until_refresh = refresh_interval - time_since_refresh
            st.sidebar.info(f"⏳ Next refresh in {int(time_until_refresh)} seconds")
    
    # Check if we have data to display
    if st.session_state.data_dict:
        data_dict = st.session_state.data_dict
        
        # Sheet/Data selection
        st.sidebar.header("📊 Data Selection")
        sheet_names = list(data_dict.keys())
        
        if len(sheet_names) > 1:
            selected_sheet = st.sidebar.selectbox("Choose data to analyze", sheet_names)
        else:
            selected_sheet = sheet_names[0]
            st.sidebar.info(f"Analyzing: {selected_sheet}")
        
        if selected_sheet and selected_sheet in data_dict:
            df = data_dict[selected_sheet].copy()
            
            # Get symbol info
            symbol = "OPTIONS"
            symbol_cols = [col for col in df.columns if 'symbol' in col.lower()]
            if symbol_cols and len(df) > 0:
                try:
                    symbol = str(df[symbol_cols[0]].iloc[0])
                except:
                    pass
            
            st.markdown(f"""
            <div class="info-box">
            <strong>📊 Analyzing:</strong> {symbol} - {selected_sheet}<br>
            <strong>📋 Data:</strong> {len(df)} rows, {len(df.columns)} columns<br>
            <strong>📁 Source:</strong> {st.session_state.excel_file_path if st.session_state.file_method == "path" else "Uploaded File"}
            </div>
            """, unsafe_allow_html=True)
            
            # Calculate metrics
            pcr_oi, total_call_oi, total_put_oi = calculate_pcr(df)
            pcr_vol, total_call_vol, total_put_vol = calculate_volume_pcr(df)
            max_pain = calculate_max_pain(df)
            support, resistance = get_support_resistance(df)
            
            # Key metrics
            st.header("📊 Key Metrics")
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                if total_call_oi:
                    st.metric("📞 Call OI", f"{int(total_call_oi):,}")
                else:
                    st.metric("📞 Call OI", "N/A")
            
            with col2:
                if total_put_oi:
                    st.metric("📉 Put OI", f"{int(total_put_oi):,}")
                else:
                    st.metric("📉 Put OI", "N/A")
            
            with col3:
                if pcr_oi:
                    st.metric("⚖️ PCR (OI)", f"{pcr_oi:.3f}")
                else:
                    st.metric("⚖️ PCR (OI)", "N/A")
            
            with col4:
                if pcr_vol:
                    st.metric("📊 PCR (Vol)", f"{pcr_vol:.3f}")
                else:
                    st.metric("📊 PCR (Vol)", "N/A")
            
            with col5:
                if max_pain:
                    st.metric("💰 Max Pain", f"{int(max_pain):,}")
                else:
                    st.metric("💰 Max Pain", "N/A")
            
            # Market Analysis
            st.header("📈 Market Analysis")
            display_market_sentiment(pcr_oi)
            
            # Support/Resistance
            if support and resistance:
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"""
                    <div class="success-box">
                    <strong>🟢 Support Level</strong><br>
                    {int(support):,} (Max Put OI)
                    </div>
                    """, unsafe_allow_html=True)
                with col2:
                    st.markdown(f"""
                    <div class="error-box">
                    <strong>🔴 Resistance Level</strong><br>
                    {int(resistance):,} (Max Call OI)
                    </div>
                    """, unsafe_allow_html=True)
            
            # Tabs
            tab1, tab2, tab3, tab4 = st.tabs([
                "📊 Options Chain", 
                "📈 Charts", 
                "🔥 Top Strikes",
                "📋 Raw Data"
            ])
            
            with tab1:
                st.subheader(f"📊 {symbol} Options Summary")
                
                # Show important columns
                important_keywords = ['strike', 'oi', 'volume', 'ltp', 'change']
                display_cols = []
                
                for col in df.columns:
                    col_lower = col.lower()
                    if any(keyword in col_lower for keyword in important_keywords):
                        display_cols.append(col)
                
                if display_cols:
                    display_df = df[display_cols].copy()
                    numeric_cols = display_df.select_dtypes(include=[np.number]).columns
                    for col in numeric_cols:
                        display_df[col] = pd.to_numeric(display_df[col], errors='coerce').round(2)
                    
                    st.dataframe(display_df, use_container_width=True, height=500)
                else:
                    st.dataframe(df.head(20), use_container_width=True, height=500)
            
            with tab2:
                st.header("📈 Visual Analysis")
                create_charts(df)
            
            with tab3:
                st.header("🔥 Most Active Strikes")
                display_top_strikes(df)
            
            with tab4:
                st.subheader("📋 Complete Data")
                st.dataframe(df, use_container_width=True, height=600)
                
                # Download option
                csv = df.to_csv(index=False)
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv,
                    file_name=f"{symbol}_{selected_sheet}_data.csv",
                    mime="text/csv"
                )
    
    elif st.session_state.file_method == "path" and st.session_state.excel_file_path:
        # File path is set but no data loaded yet
        st.info("📁 Excel file path is set. Click 'Manual Refresh' to load data.")
        
        # Show path troubleshooting
        with st.expander("🔧 Path Troubleshooting"):
            st.markdown("""
            ### Common Windows Path Issues:
            
            1. **Backslash Issues**: Windows uses backslashes (`\\`) but Python often prefers forward slashes (`/`)
            
            2. **Try these formats for your path**:
            """)
            
            original_path = st.session_state.excel_file_path
            st.markdown(f"""
            <div class="path-fix">
            Original: {original_path}
            </div>
            """, unsafe_allow_html=True)
            
            # Show alternative formats
            alt1 = original_path.replace("\\", "/")
            alt2 = original_path.replace("\\", "\\\\")
            alt3 = f"/{original_path[0].lower()}/" + original_path[3:].replace("\\", "/")
            
            st.markdown(f"""
            <div class="path-fix">
            Option 1: {alt1}
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="path-fix">
            Option 2: {alt2}
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="path-fix">
            Option 3: {alt3}
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("""
            **Quick Fixes:**
            - Use the file upload option instead
            - Copy the file to a simpler path (like C:/data/file.xlsx)
            - Use relative paths if the file is in the same directory as this app
            """)
        
        if st.button("🔄 Load Data Now", type="primary"):
            fetch_data_from_excel()
            st.rerun()
    
    else:
        # Welcome screen
        st.markdown("""
        <div class="info-box">
        <h2>🚀 Welcome to NSE Options Dashboard!</h2>
        <p><strong>Configure your Excel file path to start analyzing live data!</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### ✨ **Key Features:**
            - 📊 **Live Options Chain Analysis**
            - 📈 **Interactive Charts**  
            - 💹 **Real-time PCR Monitoring**
            - ⚖️ **Max Pain Calculation**
            - 🎯 **Support/Resistance Levels**
            - 🔥 **Top Strikes Analysis**
            - 🔄 **Auto-Refresh**
            - 📥 **Data Export**
            """)
        
        with col2:
            st.markdown("""
            ### 📁 **Setup Instructions:**
            1. Choose between file path or file upload method
            2. Provide your Excel file with options data
            3. Set your preferred refresh interval
            4. Enable auto-refresh or manually refresh data
            5. Select the sheet you want to analyze
            
            ### 💡 **Tips:**
            - The Excel file should contain options chain data
            - File formats: .xlsx, .xlsm, .xls
            - Use the upload option if path issues persist
            """)
        
        # Sample data
        st.subheader("📊 Expected Data Format")
        sample_data = pd.DataFrame({
            'Strike': [22500, 22550, 22600],
            'CE_OI': [1500, 2300, 3400],
            'PE_OI': [1200, 1800, 2100],
            'CE_Total_Traded_Volume': [450, 680, 920],
            'PE_Total_Traded_Volume': [320, 490, 650],
            'CE_LTP': [245.5, 195.3, 148.7],
            'PE_LTP': [38.4, 58.9, 85.3]
        })
        st.dataframe(sample_data, use_container_width=True)

if __name__ == "__main__":
    main()
