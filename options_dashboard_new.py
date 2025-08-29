import pandas as pd
import numpy as np
import streamlit as st
import warnings
import time
import subprocess
import sys
import io
import os
import tempfile
from datetime import datetime
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
.upload-info {
    background-color: #f8f9fa;
    padding: 15px;
    border-radius: 5px;
    border-left: 4px solid #6c757d;
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
    st.session_state.file_method = "upload"
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None

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

def load_uploaded_file(uploaded_file):
    """Load data from uploaded file with better error handling"""
    try:
        # Check if Excel libraries are available
        if not check_excel_libraries():
            return {'excel_libraries_missing': True}
        
        # Create a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            # Write uploaded file to temporary file
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
        
        try:
            # Load the saved file
            data_dict = read_excel_with_pandas(tmp_path)
            return data_dict
        finally:
            # Clean up temporary file
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
                
    except Exception as e:
        st.error(f"❌ Error processing uploaded file: {str(e)}")
        return {'upload_error': str(e)}

def process_uploaded_file():
    """Process the uploaded file and update session state"""
    if st.session_state.uploaded_file is not None:
        with st.spinner("🔄 Processing uploaded file..."):
            data_dict = load_uploaded_file(st.session_state.uploaded_file)
            
            if data_dict and 'excel_libraries_missing' not in data_dict and 'excel_read_error' not in data_dict and 'upload_error' not in data_dict:
                st.session_state.data_dict = data_dict
                st.session_state.last_refresh = datetime.now()
                st.success("✅ File uploaded and processed successfully!")
                return True
            else:
                if 'excel_libraries_missing' in data_dict:
                    st.error("❌ Excel libraries are missing. Please install them.")
                    if st.button("🛠️ Install Excel Libraries"):
                        if install_excel_libraries():
                            st.rerun()
                elif 'excel_read_error' in data_dict:
                    st.error("❌ Could not read the Excel file. It may be corrupted or in an unsupported format.")
                elif 'upload_error' in data_dict:
                    st.error(f"❌ Error uploading file: {data_dict['upload_error']}")
                else:
                    st.error("❌ Failed to process the uploaded file.")
                return False
    return False

def calculate_pcr(df):
    """Calculate Put-Call Ratio"""
    try:
        # Try to find OI columns with different patterns
        call_oi_patterns = ['CE_OI', 'CE.OI', 'CALL_OI', 'CALL.OI']
        put_oi_patterns = ['PE_OI', 'PE.OI', 'PUT_OI', 'PUT.OI']
        
        call_oi_col = None
        put_oi_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_oi_col is None and any(pattern in col_upper for pattern in call_oi_patterns):
                call_oi_col = col
            if put_oi_col is None and any(pattern in col_upper for pattern in put_oi_patterns):
                put_oi_col = col
        
        if call_oi_col and put_oi_col:
            total_call_oi = df[call_oi_col].fillna(0).sum()
            total_put_oi = df[put_oi_col].fillna(0).sum()
            pcr_oi = total_put_oi / total_call_oi if total_call_oi > 0 else 0
            return pcr_oi, total_call_oi, total_put_oi
        
        return None, None, None
    except Exception:
        return None, None, None

def calculate_volume_pcr(df):
    """Calculate Volume PCR"""
    try:
        # Try to find volume columns with different patterns
        call_vol_patterns = ['CE_VOLUME', 'CE.VOLUME', 'CALL_VOLUME', 'CALL.VOLUME']
        put_vol_patterns = ['PE_VOLUME', 'PE.VOLUME', 'PUT_VOLUME', 'PUT.VOLUME']
        
        call_vol_col = None
        put_vol_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_vol_col is None and any(pattern in col_upper for pattern in call_vol_patterns):
                call_vol_col = col
            if put_vol_col is None and any(pattern in col_upper for pattern in put_vol_patterns):
                put_vol_col = col
        
        if call_vol_col and put_vol_col:
            total_call_vol = df[call_vol_col].fillna(0).sum()
            total_put_vol = df[put_vol_col].fillna(0).sum()
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
        
        if strike_col is None:
            return None
        
        call_oi_col = None
        put_oi_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_oi_col is None and ('CE_OI' in col_upper or 'CALL_OI' in col_upper):
                call_oi_col = col
            if put_oi_col is None and ('PE_OI' in col_upper or 'PUT_OI' in col_upper):
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
        
        if strike_col is None:
            return None, None
        
        call_oi_col = None
        put_oi_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_oi_col is None and ('CE_OI' in col_upper or 'CALL_OI' in col_upper):
                call_oi_col = col
            if put_oi_col is None and ('PE_OI' in col_upper or 'PUT_OI' in col_upper):
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
        call_oi_col = None
        put_oi_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_oi_col is None and ('CE_OI' in col_upper or 'CALL_OI' in col_upper):
                call_oi_col = col
            if put_oi_col is None and ('PE_OI' in col_upper or 'PUT_OI' in col_upper):
                put_oi_col = col
        
        if call_oi_col and put_oi_col:
            chart_data = df[[strike_col, call_oi_col, put_oi_col]].copy()
            chart_data = chart_data.dropna()
            
            if not chart_data.empty:
                chart_data = chart_data.set_index(strike_col)
                chart_data.columns = ['Call OI', 'Put OI']
                
                st.subheader("📊 Open Interest Distribution")
                st.bar_chart(chart_data, height=400)
        
        # Volume Chart
        call_vol_col = None
        put_vol_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_vol_col is None and ('CE_VOLUME' in col_upper or 'CALL_VOLUME' in col_upper):
                call_vol_col = col
            if put_vol_col is None and ('PE_VOLUME' in col_upper or 'PUT_VOLUME' in col_upper):
                put_vol_col = col
        
        if call_vol_col and put_vol_col:
            vol_data = df[[strike_col, call_vol_col, put_vol_col]].copy()
            vol_data = vol_data.dropna()
            
            if not vol_data.empty:
                vol_data = vol_data.set_index(strike_col)
                vol_data.columns = ['Call Volume', 'Put Volume']
                
                st.subheader("📈 Volume Distribution")
                st.bar_chart(vol_data, height=400)
        
        # IV Chart
        call_iv_col = None
        put_iv_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_iv_col is None and ('CE_IV' in col_upper or 'CALL_IV' in col_upper):
                call_iv_col = col
            if put_iv_col is None and ('PE_IV' in col_upper or 'PUT_IV' in col_upper):
                put_iv_col = col
        
        if call_iv_col and put_iv_col:
            iv_data = df[[strike_col, call_iv_col, put_iv_col]].copy()
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
        
        call_oi_col = None
        put_oi_col = None
        call_vol_col = None
        put_vol_col = None
        
        for col in df.columns:
            col_upper = col.upper()
            if call_oi_col is None and ('CE_OI' in col_upper or 'CALL_OI' in col_upper):
                call_oi_col = col
            if put_oi_col is None and ('PE_OI' in col_upper or 'PUT_OI' in col_upper):
                put_oi_col = col
            if call_vol_col is None and ('CE_VOLUME' in col_upper or 'CALL_VOLUME' in col_upper):
                call_vol_col = col
            if put_vol_col is None and ('PE_VOLUME' in col_upper or 'PUT_VOLUME' in col_upper):
                put_vol_col = col
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🟢 Top Call Activity")
            if call_oi_col:
                display_cols = [strike_col, call_oi_col]
                if call_vol_col:
                    display_cols.append(call_vol_col)
                
                top_call = df[display_cols].nlargest(5, call_oi_col)
                st.dataframe(top_call, hide_index=True)
        
        with col2:
            st.subheader("🔴 Top Put Activity")
            if put_oi_col:
                display_cols = [strike_col, put_oi_col]
                if put_vol_col:
                    display_cols.append(put_vol_col)
                
                top_put = df[display_cols].nlargest(5, put_oi_col)
                st.dataframe(top_put, hide_index=True)
    
    except Exception as e:
        st.warning(f"Could not display top strikes: {str(e)}")

def display_data_analysis(data_dict):
    """Display the data analysis section"""
    if not data_dict:
        return
    
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
        <strong>📁 Source:</strong> Uploaded File
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
            if total_call_oi is not None:
                st.metric("📞 Call OI", f"{int(total_call_oi):,}")
            else:
                st.metric("📞 Call OI", "N/A")
        
        with col2:
            if total_put_oi is not None:
                st.metric("📉 Put OI", f"{int(total_put_oi):,}")
            else:
                st.metric("📉 Put OI", "N/A")
        
        with col3:
            if pcr_oi is not None:
                st.metric("⚖️ PCR (OI)", f"{pcr_oi:.3f}")
            else:
                st.metric("⚖️ PCR (OI)", "N/A")
        
        with col4:
            if pcr_vol is not None:
                st.metric("📊 PCR (Vol)", f"{pcr_vol:.3f}")
            else:
                st.metric("📊 PCR (Vol)", "N/A")
        
        with col5:
            if max_pain is not None:
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
            important_keywords = ['strike', 'oi', 'volume', 'ltp', 'change', 'iv', 'symbol']
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

def main():
    st.markdown('<h1 class="main-header">⚡ NSE Options Dashboard</h1>', unsafe_allow_html=True)
    
    # Current time
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Sidebar
    st.sidebar.header("📁 Data Source Configuration")
    
    # File uploader
    st.sidebar.markdown("""
    <div class="upload-info">
    <strong>📤 Upload Your Excel File</strong><br>
    Supported formats: .xlsx, .xlsm, .xls
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.sidebar.file_uploader(
        "Choose Excel file",
        type=['xlsx', 'xlsm', 'xls'],
        help="Upload your Excel file with options data",
        label_visibility="collapsed"
    )
    
    # Store uploaded file in session state
    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file
        st.session_state.file_method = "upload"
    
    # Process button
    if st.sidebar.button("🔄 Process Uploaded File", type="primary", use_container_width=True):
        process_uploaded_file()
        st.rerun()
    
    # Refresh settings
    st.sidebar.header("🔄 Refresh Settings")
    st.session_state.refresh_interval = st.sidebar.number_input(
        "Refresh Interval (seconds)", 
        min_value=5, 
        max_value=300, 
        value=30,
        help="How often to automatically refresh data"
    )
    
    st.session_state.auto_refresh = st.sidebar.checkbox(
        "Auto Refresh", 
        value=st.session_state.auto_refresh,
        help="Automatically refresh data at the specified interval"
    )
    
    # File format recommendation
    st.sidebar.markdown("""
    **📊 Excel File Requirements:**
    - ✅ File should contain options chain data
    - ✅ Supported formats: .xlsx, .xlsm, .xls
    - ✅ Ensure the file is not password protected
    """)
    
    # Display last refresh time
    if st.session_state.last_refresh:
        last_refresh_str = st.session_state.last_refresh.strftime("%Y-%m-%d %H:%M:%S")
        st.markdown(f"<div class='info-box'><strong>🕐 Last Refresh:</strong> {last_refresh_str}</div>", unsafe_allow_html=True)
    else:
        st.markdown(f"<div class='info-box'><strong>🕐 Current Time:</strong> {current_time}</div>", unsafe_allow_html=True)
    
    # Check if we have data to display
    if st.session_state.data_dict:
        display_data_analysis(st.session_state.data_dict)
    
    elif st.session_state.uploaded_file is not None:
        # File is uploaded but not processed yet
        st.info("📁 Excel file is uploaded. Click 'Process Uploaded File' to analyze it.")
        
        # Show file info
        file_info = st.session_state.uploaded_file
        st.markdown(f"""
        <div class="upload-info">
        <strong>📄 Uploaded File:</strong> {file_info.name}<br>
        <strong>📊 Size:</strong> {file_info.size / (1024*1024):.2f} MB<br>
        <strong>🕒 Uploaded at:</strong> {current_time}
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🔄 Process File Now", type="primary"):
            if process_uploaded_file():
                st.rerun()
    
    else:
        # Welcome screen
        st.markdown("""
        <div class="info-box">
        <h2>🚀 Welcome to NSE Options Dashboard!</h2>
        <p><strong>Upload your Excel file to start analyzing options data!</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### ✨ **Key Features:**
            - 📊 **Options Chain Analysis**
            - 📈 **Interactive Charts**  
            - 💹 **PCR Monitoring**
            - ⚖️ **Max Pain Calculation**
            - 🎯 **Support/Resistance Levels**
            - 🔥 **Top Strikes Analysis**
            - 🔄 **Auto-Refresh**
            - 📥 **Data Export**
            """)
        
        with col2:
            st.markdown("""
            ### 📁 **Setup Instructions:**
            1. Upload your Excel file using the sidebar
            2. Click "Process Uploaded File"
            3. Select the sheet you want to analyze
            4. Set your preferred refresh interval
            5. Enable auto-refresh if desired
            
            ### 💡 **Tips:**
            - The Excel file should contain options chain data
            - File formats: .xlsx, .xlsm, .xls
            - Ensure the file is not password protected
            """)
        
        # Sample data
        st.subheader("📊 Expected Data Format")
        sample_data = pd.DataFrame({
            'Strike': [22500, 22550, 22600],
            'CE_OI': [1500, 2300, 3400],
            'PE_OI': [1200, 1800, 2100],
            'CE_Volume': [450, 680, 920],
            'PE_Volume': [320, 490, 650],
            'CE_LTP': [245.5, 195.3, 148.7],
            'PE_LTP': [38.4, 58.9, 85.3]
        })
        st.dataframe(sample_data, use_container_width=True)

if __name__ == "__main__":
    main()
