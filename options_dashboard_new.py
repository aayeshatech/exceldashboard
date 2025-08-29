import pandas as pd
import numpy as np
import streamlit as st
import warnings
import time
import subprocess
import sys
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
</style>
""", unsafe_allow_html=True)

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

@st.cache_data(ttl=30)
def load_data_file(file):
    """Load data file - supports Excel and CSV with fallback methods"""
    file_extension = file.name.split('.')[-1].lower()
    
    try:
        if file_extension == 'csv':
            # CSV files always work - no dependencies needed
            df = pd.read_csv(file)
            if not df.empty:
                st.success(f"✅ Loaded CSV file successfully ({len(df)} rows, {len(df.columns)} columns)")
                return {'CSV_Data': df}
        
        elif file_extension in ['xlsx', 'xlsm']:
            # Try multiple methods for Excel files
            data_dict = {}
            
            # Method 1: Check if Excel libraries are available
            try:
                # Try importing required libraries
                import openpyxl
                import xlrd
                
                # If we get here, libraries are available
                try:
                    excel_file = pd.ExcelFile(file)
                    st.info(f"📁 Found {len(excel_file.sheet_names)} sheets in Excel file")
                    
                    # Try to read each sheet individually
                    for sheet_name in excel_file.sheet_names:
                        try:
                            df = pd.read_excel(file, sheet_name=sheet_name)
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
                    else:
                        st.error("❌ No sheets could be loaded from the Excel file")
                        return {}
                
                except Exception as excel_error:
                    st.error(f"❌ Error reading Excel file: {str(excel_error)}")
                    return {'excel_libraries_missing': True}
            
            except ImportError:
                # If Excel libraries not available, return a special indicator
                return {'excel_libraries_missing': True}
            
            except Exception as e:
                # Try reading first sheet only
                try:
                    st.warning("⚠️ Trying to read first sheet only...")
                    df = pd.read_excel(file)
                    if not df.empty:
                        return {'Sheet1': df}
                except Exception:
                    st.error(f"""
                    ❌ **Cannot read Excel file**: {str(e)}
                    """)
                    return {}
        
        else:
            st.error(f"❌ Unsupported file format: {file_extension}")
            return {}
            
    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")
        return {}

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
    st.markdown(f"<div class='info-box'><strong>🕐 Last Updated:</strong> {current_time}</div>", unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.header("📁 Upload Your Data")
    
    # File uploader
    uploaded_file = st.sidebar.file_uploader(
        "Choose your file",
        type=['xlsx', 'xlsm', 'csv'],
        help="Upload Excel or CSV file with options data"
    )
    
    # File format recommendation
    st.sidebar.markdown("""
    **📄 Recommended: CSV Format**
    - ✅ Faster loading
    - ✅ Always works
    - ✅ No dependencies
    
    **📊 Excel Support**
    - ⚠️ May need additional packages
    - 💡 Convert to CSV if issues
    """)
    
    if uploaded_file is not None:
        # Load data
        with st.spinner("Loading file..."):
            data_dict = load_data_file(uploaded_file)
        
        # Check if Excel libraries are missing
        if 'excel_libraries_missing' in data_dict:
            st.error("""
            📊 **Excel file detected, but Excel support libraries not installed**
            """)
            
            st.markdown("""
            **Option 1: Install Excel Libraries (Recommended)**
            """)
            
            st.markdown("""
            <div class="code-box">
            pip install openpyxl xlrd
            </div>
            """, unsafe_allow_html=True)
            
            # Add install button
            if st.button("🛠️ Install Excel Libraries"):
                if install_excel_libraries():
                    st.rerun()
            
            st.markdown("""
            **Option 2: Convert to CSV (Quick Fix)**
            1. Open your Excel file
            2. Select the sheet you want (e.g., OC_1, OC_2, OC_3)
            3. File → Save As → CSV format
            4. Upload the CSV file instead
            
            **Why CSV is better:**
            - ✅ No extra packages needed
            - ✅ Faster loading
            - ✅ Works on all systems
            - ✅ Smaller file size
            
            **CSV files work without any extra packages!**
            """)
            
            # Show a sample of how to convert
            st.markdown("""
            <div class="info-box">
            <strong>📋 How to convert Excel to CSV:</strong><br>
            1. Open your Excel file<br>
            2. Go to File → Save As<br>
            3. Choose "CSV (Comma delimited)" as the file type<br>
            4. Save with a new name<br>
            5. Upload the CSV file here
            </div>
            """, unsafe_allow_html=True)
        
        elif data_dict:
            # Auto refresh option
            st.sidebar.header("🔄 Settings")
            auto_refresh = st.sidebar.checkbox("Auto Refresh (30 sec)", value=False)
            
            if auto_refresh:
                st.sidebar.success("✅ Auto-refresh enabled")
                time.sleep(30)
                st.rerun()
            
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
                <strong>📋 Data:</strong> {len(df)} rows, {len(df.columns)} columns
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
        
        else:
            st.error("❌ Could not load data from the file.")
    
    else:
        # Welcome screen
        st.markdown("""
        <div class="info-box">
        <h2>🚀 Welcome to NSE Options Dashboard!</h2>
        <p><strong>Upload your data file to start analyzing!</strong></p>
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
            - 🎯 **Support/Resistance**
            - 🔥 **Top Strikes Analysis**
            - 📱 **Auto-Refresh**
            - 📥 **Data Export**
            """)
        
        with col2:
            st.markdown("""
            ### 📁 **Supported Files:**
            - 📄 **CSV** (Recommended - always works!)
            - 📊 **Excel .xlsx**
            - 📊 **Excel .xlsm** (with macros)
            
            ### 💡 **Tips:**
            - **CSV files load faster**
            - **No dependencies needed for CSV**
            - **Convert Excel to CSV if issues**
            """)
        
        # How to convert
        with st.expander("📝 How to Convert Excel to CSV"):
            st.markdown("""
            **Quick Steps:**
            1. Open your Excel file (Live_Option_Chain_Terminal.xlsm)
            2. Select the sheet you want (e.g., OC_1, OC_2, OC_3)
            3. **File → Save As**
            4. Choose **"CSV (Comma delimited)"**
            5. Save with a new name
            6. Upload the CSV file here
            
            **Why CSV is better:**
            - ✅ Loads instantly
            - ✅ Works everywhere
            - ✅ Smaller file size
            - ✅ No compatibility issues
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
