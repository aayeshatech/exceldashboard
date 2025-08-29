import pandas as pd
import numpy as np
import streamlit as st
import warnings
import time
from datetime import datetime

warnings.filterwarnings('ignore')

# Check for required dependencies
def check_dependencies():
    """Check if required dependencies are installed"""
    missing_deps = []
    
    try:
        import openpyxl
    except ImportError:
        missing_deps.append('openpyxl')
    
    try:
        import xlrd
    except ImportError:
        missing_deps.append('xlrd')
    
    return missing_deps

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
.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 1rem;
    border-radius: 10px;
    color: white;
    text-align: center;
    margin: 0.5rem 0;
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
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=30)
def load_excel_data(file):
    """Load Excel data with error handling"""
    try:
        excel_file = pd.ExcelFile(file)
        data_dict = {}
        
        st.info(f"📁 Loading {len(excel_file.sheet_names)} sheets from Excel file...")
        
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file, sheet_name=sheet_name)
                if not df.empty:
                    data_dict[sheet_name] = df
                    st.success(f"✅ Loaded sheet: {sheet_name} ({len(df)} rows)")
            except Exception as e:
                st.warning(f"⚠️ Could not load sheet {sheet_name}: {str(e)}")
                continue
                
        return data_dict
    except Exception as e:
        st.error(f"❌ Error loading Excel file: {str(e)}")
        return {}

def safe_calculate_pcr(df):
    """Safely calculate Put-Call Ratio"""
    try:
        call_oi_cols = [col for col in df.columns if 'CE_OI' in col and 'Change' not in col]
        put_oi_cols = [col for col in df.columns if 'PE_OI' in col and 'Change' not in col]
        
        if call_oi_cols and put_oi_cols:
            total_call_oi = df[call_oi_cols[0]].fillna(0).sum()
            total_put_oi = df[put_oi_cols[0]].fillna(0).sum()
            pcr_oi = total_put_oi / total_call_oi if total_call_oi > 0 else 0
            return pcr_oi, total_call_oi, total_put_oi
        
        return None, None, None
    except Exception as e:
        st.warning(f"Could not calculate PCR: {str(e)}")
        return None, None, None

def safe_calculate_volume_pcr(df):
    """Safely calculate Volume PCR"""
    try:
        call_vol_cols = [col for col in df.columns if 'CE_' in col and 'Volume' in col]
        put_vol_cols = [col for col in df.columns if 'PE_' in col and 'Volume' in col]
        
        if call_vol_cols and put_vol_cols:
            total_call_vol = df[call_vol_cols[0]].fillna(0).sum()
            total_put_vol = df[put_vol_cols[0]].fillna(0).sum()
            pcr_vol = total_put_vol / total_call_vol if total_call_vol > 0 else 0
            return pcr_vol, total_call_vol, total_put_vol
        
        return None, None, None
    except Exception as e:
        st.warning(f"Could not calculate Volume PCR: {str(e)}")
        return None, None, None

def safe_calculate_max_pain(df):
    """Safely calculate Max Pain"""
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
    except Exception as e:
        st.warning(f"Could not calculate Max Pain: {str(e)}")
        return None

def get_support_resistance(df):
    """Get support and resistance levels safely"""
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
    except Exception as e:
        st.warning(f"Could not calculate support/resistance: {str(e)}")
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

def create_simple_charts(df):
    """Create simple charts using Streamlit native functionality"""
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
        # Find columns
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
    st.markdown('<h1 class="main-header">⚡ NSE Options Chain Dashboard</h1>', unsafe_allow_html=True)
    
    # Check dependencies first
    missing_deps = check_dependencies()
    if missing_deps:
        st.error("❌ Missing Required Dependencies!")
        st.markdown(f"""
        <div class="error-box">
        <strong>Missing packages:</strong> {', '.join(missing_deps)}<br><br>
        <strong>To fix this, run these commands:</strong><br>
        <code>pip install {' '.join(missing_deps)}</code><br><br>
        <strong>Or install all at once:</strong><br>
        <code>pip install openpyxl xlrd</code><br><br>
        <strong>If using conda:</strong><br>
        <code>conda install openpyxl xlrd</code>
        </div>
        """, unsafe_allow_html=True)
        
        st.info("🔄 After installing the packages, refresh this page to continue.")
        st.stop()
    
    # Current time
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.markdown(f"<div class='info-box'><strong>🕐 Last Updated:</strong> {current_time}</div>", unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.header("📁 Upload Options Data")
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xlsm'],
        help="Upload your Live_Option_Chain_Terminal.xlsm file"
    )
    
    if uploaded_file is not None:
        # Load data
        with st.spinner("Loading Excel file..."):
            data_dict = load_excel_data(uploaded_file)
        
        if data_dict:
            # Auto refresh
            st.sidebar.header("🔄 Settings")
            auto_refresh = st.sidebar.checkbox("Auto Refresh (30 sec)", value=False)
            
            if auto_refresh:
                st.sidebar.success("✅ Auto-refresh enabled")
                time.sleep(30)
                st.rerun()
            
            # Sheet selection
            st.sidebar.header("📊 Select Sheet")
            sheet_names = list(data_dict.keys())
            
            # Show sheet info
            st.sidebar.info(f"Found {len(sheet_names)} sheets")
            for sheet in sheet_names[:5]:  # Show first 5
                st.sidebar.text(f"• {sheet}")
            if len(sheet_names) > 5:
                st.sidebar.text(f"... and {len(sheet_names) - 5} more")
            
            # Filter for options sheets
            options_sheets = []
            for sheet in sheet_names:
                sheet_upper = sheet.upper()
                if any(x in sheet_upper for x in ['OC_', 'OPTION', 'CHAIN']):
                    options_sheets.append(sheet)
            
            if not options_sheets:
                options_sheets = sheet_names[:10]  # Take first 10 if no obvious options sheets
            
            selected_sheet = st.sidebar.selectbox("Choose Sheet", options_sheets)
            
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
                
                # Calculate all metrics safely
                pcr_oi, total_call_oi, total_put_oi = safe_calculate_pcr(df)
                pcr_vol, total_call_vol, total_put_vol = safe_calculate_volume_pcr(df)
                max_pain = safe_calculate_max_pain(df)
                support, resistance = get_support_resistance(df)
                
                # Display key metrics
                st.header("📊 Key Metrics")
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    if total_call_oi:
                        st.metric(
                            label="📞 Call OI",
                            value=f"{int(total_call_oi):,}"
                        )
                    else:
                        st.metric("📞 Call OI", "N/A")
                
                with col2:
                    if total_put_oi:
                        st.metric(
                            label="📉 Put OI", 
                            value=f"{int(total_put_oi):,}"
                        )
                    else:
                        st.metric("📉 Put OI", "N/A")
                
                with col3:
                    if pcr_oi:
                        st.metric(
                            label="⚖️ PCR (OI)",
                            value=f"{pcr_oi:.3f}"
                        )
                    else:
                        st.metric("⚖️ PCR (OI)", "N/A")
                
                with col4:
                    if pcr_vol:
                        st.metric(
                            label="📊 PCR (Vol)",
                            value=f"{pcr_vol:.3f}"
                        )
                    else:
                        st.metric("📊 PCR (Vol)", "N/A")
                
                with col5:
                    if max_pain:
                        st.metric(
                            label="💰 Max Pain",
                            value=f"₹{int(max_pain):,}"
                        )
                    else:
                        st.metric("💰 Max Pain", "N/A")
                
                # Market Sentiment
                st.header("📈 Market Analysis")
                display_market_sentiment(pcr_oi)
                
                # Support/Resistance
                if support and resistance:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(f"""
                        <div class="success-box">
                        <strong>🟢 Support Level</strong><br>
                        ₹{int(support):,} (Max Put OI)
                        </div>
                        """, unsafe_allow_html=True)
                    with col2:
                        st.markdown(f"""
                        <div class="error-box">
                        <strong>🔴 Resistance Level</strong><br>
                        ₹{int(resistance):,} (Max Call OI)
                        </div>
                        """, unsafe_allow_html=True)
                
                # Create tabs
                tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📊 Quick View", 
                    "📈 Charts", 
                    "🔥 Top Strikes",
                    "📋 Full Data", 
                    "ℹ️ All Sheets"
                ])
                
                with tab1:
                    st.subheader(f"📊 {symbol} Options Chain Summary")
                    
                    # Show important columns only
                    important_keywords = ['strike', 'oi', 'volume', 'ltp', 'change']
                    display_cols = []
                    
                    for col in df.columns:
                        col_lower = col.lower()
                        if any(keyword in col_lower for keyword in important_keywords):
                            display_cols.append(col)
                    
                    if display_cols:
                        display_df = df[display_cols].copy()
                        # Round numeric columns
                        numeric_cols = display_df.select_dtypes(include=[np.number]).columns
                        for col in numeric_cols:
                            display_df[col] = pd.to_numeric(display_df[col], errors='coerce').round(2)
                        
                        st.dataframe(display_df, use_container_width=True, height=500)
                    else:
                        st.dataframe(df.head(20), use_container_width=True, height=500)
                
                with tab2:
                    st.header("📈 Visual Analysis")
                    create_simple_charts(df)
                
                with tab3:
                    st.header("🔥 Most Active Strikes")
                    display_top_strikes(df)
                    
                    # Show OI changes if available
                    change_cols = [col for col in df.columns if 'change' in col.lower() and 'oi' in col.lower()]
                    if change_cols:
                        st.subheader("📊 Recent OI Changes")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            ce_change_cols = [col for col in change_cols if 'CE' in col]
                            if ce_change_cols:
                                st.write("**📈 Call OI Changes:**")
                                change_data = df[df[ce_change_cols[0]] != 0].nlargest(5, ce_change_cols[0])
                                if not change_data.empty:
                                    display_cols = ['Strike', ce_change_cols[0]] if 'Strike' in df.columns else [ce_change_cols[0]]
                                    st.dataframe(change_data[display_cols], hide_index=True)
                        
                        with col2:
                            pe_change_cols = [col for col in change_cols if 'PE' in col]
                            if pe_change_cols:
                                st.write("**📉 Put OI Changes:**")
                                change_data = df[df[pe_change_cols[0]] != 0].nlargest(5, pe_change_cols[0])
                                if not change_data.empty:
                                    display_cols = ['Strike', pe_change_cols[0]] if 'Strike' in df.columns else [pe_change_cols[0]]
                                    st.dataframe(change_data[display_cols], hide_index=True)
                
                with tab4:
                    st.subheader("📋 Complete Raw Data")
                    st.dataframe(df, use_container_width=True, height=600)
                    
                    # Download option
                    csv = df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download as CSV",
                        data=csv,
                        file_name=f"{symbol}_{selected_sheet}_data.csv",
                        mime="text/csv"
                    )
                
                with tab5:
                    st.subheader("📋 All Available Sheets")
                    
                    sheet_info = []
                    for sheet_name, sheet_df in data_dict.items():
                        options_cols = [col for col in sheet_df.columns if any(x in col for x in ['CE_', 'PE_', 'Call', 'Put'])]
                        
                        sheet_info.append({
                            'Sheet Name': sheet_name,
                            'Rows': len(sheet_df),
                            'Columns': len(sheet_df.columns),
                            'Options Columns': len(options_cols),
                            'Has Options Data': 'Yes' if options_cols else 'No'
                        })
                    
                    sheet_info_df = pd.DataFrame(sheet_info)
                    st.dataframe(sheet_info_df, hide_index=True, use_container_width=True)
                    
                    # Quick preview
                    st.subheader("👀 Quick Preview")
                    other_sheets = [s for s in sheet_names if s != selected_sheet]
                    if other_sheets:
                        preview_sheet = st.selectbox("Select sheet to preview", other_sheets)
                        
                        if preview_sheet:
                            st.write(f"**First 10 rows of {preview_sheet}:**")
                            preview_df = data_dict[preview_sheet].head(10)
                            st.dataframe(preview_df, use_container_width=True)
        
        else:
            st.error("❌ Could not load any data from the file. Please check the file format.")
    
    else:
        # Welcome screen
        st.markdown("""
        <div class="info-box">
        <h2>🚀 Welcome to NSE Options Chain Dashboard!</h2>
        <p><strong>Upload your Excel file to get started</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        ### ✨ Features:
        - 📊 **Options Chain Analysis** - Complete CE/PE data analysis
        - 📈 **Native Charts** - OI, Volume, and IV visualization  
        - 💹 **PCR Monitoring** - Put-Call Ratio for market sentiment
        - ⚖️ **Max Pain** - Automatic calculation
        - 🎯 **Support/Resistance** - Key levels from OI data
        - 🔥 **Active Strikes** - Most traded options identification
        - 📱 **Auto-Refresh** - Live data updates every 30 seconds
        - 📥 **Data Export** - Download analysis as CSV
        - 🛡️ **Error-Free** - Robust handling of any data format
        
        ### 📁 Supported Files:
        - `.xlsx` - Excel files
        - `.xlsm` - Excel files with macros (like your Live_Option_Chain_Terminal.xlsm)
        
        ### 🎯 Perfect for:
        - Live options trading analysis
        - Market sentiment tracking
        - Risk management
        - Strike selection
        - Options flow monitoring
        """)

if __name__ == "__main__":
    main()
