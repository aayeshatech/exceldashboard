import pandas as pd
import numpy as np
import streamlit as st
import warnings
import time
import os
from datetime import datetime
import traceback
warnings.filterwarnings('ignore')

# Set page config
st.set_page_config(
    page_title="NSE Options Dashboard",
    page_icon="⚡",
    layout="wide"
)

# Custom CSS (simplified)
st.markdown("""
<style>
.main-header {
    font-size: 2.5rem;
    font-weight: bold;
    text-align: center;
    color: #FF6B35;
    margin-bottom: 1rem;
}
.metric-box {
    background-color: #f0f2f6;
    padding: 1rem;
    border-radius: 5px;
    text-align: center;
    margin: 0.5rem 0;
}
.success { background-color: #d4edda; color: #155724; }
.warning { background-color: #fff3cd; color: #856404; }
.error { background-color: #f8d7da; color: #721c24; }
.info { background-color: #d1ecf1; color: #0c5460; }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'last_refresh' not in st.session_state:
    st.session_state.last_refresh = None
if 'data_dict' not in st.session_state:
    st.session_state.data_dict = {}
if 'excel_file_path' not in st.session_state:
    st.session_state.excel_file_path = ""
if 'debug_info' not in st.session_state:
    st.session_state.debug_info = []

def log_debug(message):
    """Add debug information"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.debug_info.append(f"[{timestamp}] {message}")
    if len(st.session_state.debug_info) > 20:  # Keep only last 20 entries
        st.session_state.debug_info = st.session_state.debug_info[-20:]

def read_excel_file(file_path):
    """Read Excel file with comprehensive error handling"""
    try:
        log_debug(f"Attempting to read file: {file_path}")
        
        # Check if file exists
        if not os.path.exists(file_path):
            log_debug(f"File not found: {file_path}")
            return {"error": f"File not found: {file_path}"}
        
        # Check file extension
        file_ext = file_path.lower().split('.')[-1]
        log_debug(f"File extension: {file_ext}")
        
        if file_ext not in ['xlsx', 'xlsm', 'xls']:
            return {"error": f"Unsupported file format: {file_ext}"}
        
        # Try to read the Excel file
        try:
            # First, get all sheet names
            excel_file = pd.ExcelFile(file_path, engine='openpyxl' if file_ext in ['xlsx', 'xlsm'] else 'xlrd')
            sheet_names = excel_file.sheet_names
            log_debug(f"Found sheets: {sheet_names}")
            
            data_dict = {}
            
            # Read each sheet
            for sheet_name in sheet_names:
                try:
                    log_debug(f"Reading sheet: {sheet_name}")
                    df = pd.read_excel(
                        file_path, 
                        sheet_name=sheet_name,
                        engine='openpyxl' if file_ext in ['xlsx', 'xlsm'] else 'xlrd'
                    )
                    
                    if not df.empty:
                        log_debug(f"Sheet {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
                        data_dict[sheet_name] = df
                    else:
                        log_debug(f"Sheet {sheet_name} is empty")
                        
                except Exception as sheet_error:
                    log_debug(f"Error reading sheet {sheet_name}: {str(sheet_error)}")
                    continue
            
            if data_dict:
                log_debug(f"Successfully loaded {len(data_dict)} sheets")
                return data_dict
            else:
                return {"error": "No data could be loaded from any sheets"}
                
        except Exception as read_error:
            log_debug(f"Error reading Excel file: {str(read_error)}")
            return {"error": f"Error reading Excel file: {str(read_error)}"}
            
    except Exception as e:
        log_debug(f"Unexpected error: {str(e)}")
        return {"error": f"Unexpected error: {str(e)}"}

def load_uploaded_file(uploaded_file):
    """Load data from uploaded file"""
    try:
        log_debug("Processing uploaded file")
        
        # Save uploaded file temporarily
        temp_path = "temp_uploaded_file.xlsx"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        log_debug(f"Saved uploaded file to: {temp_path}")
        
        # Read the file
        result = read_excel_file(temp_path)
        
        # Clean up temp file
        try:
            os.remove(temp_path)
            log_debug("Cleaned up temp file")
        except:
            pass
            
        return result
        
    except Exception as e:
        log_debug(f"Error processing uploaded file: {str(e)}")
        return {"error": f"Error processing uploaded file: {str(e)}"}

def calculate_pcr_metrics(df):
    """Calculate PCR and other metrics"""
    try:
        log_debug("Calculating PCR metrics")
        log_debug(f"DataFrame columns: {list(df.columns)}")
        
        # Find relevant columns (case insensitive)
        call_oi_cols = [col for col in df.columns if 'CE' in col.upper() and 'OI' in col.upper() and 'CHANGE' not in col.upper()]
        put_oi_cols = [col for col in df.columns if 'PE' in col.upper() and 'OI' in col.upper() and 'CHANGE' not in col.upper()]
        call_vol_cols = [col for col in df.columns if 'CE' in col.upper() and 'VOL' in col.upper()]
        put_vol_cols = [col for col in df.columns if 'PE' in col.upper() and 'VOL' in col.upper()]
        strike_cols = [col for col in df.columns if 'STRIKE' in col.upper()]
        
        log_debug(f"Call OI columns: {call_oi_cols}")
        log_debug(f"Put OI columns: {put_oi_cols}")
        log_debug(f"Strike columns: {strike_cols}")
        
        metrics = {}
        
        # Calculate PCR (OI)
        if call_oi_cols and put_oi_cols:
            call_oi_total = df[call_oi_cols[0]].fillna(0).sum()
            put_oi_total = df[put_oi_cols[0]].fillna(0).sum()
            
            if call_oi_total > 0:
                pcr_oi = put_oi_total / call_oi_total
                metrics['pcr_oi'] = pcr_oi
                metrics['call_oi_total'] = call_oi_total
                metrics['put_oi_total'] = put_oi_total
                log_debug(f"PCR (OI): {pcr_oi:.3f}")
        
        # Calculate PCR (Volume)
        if call_vol_cols and put_vol_cols:
            call_vol_total = df[call_vol_cols[0]].fillna(0).sum()
            put_vol_total = df[put_vol_cols[0]].fillna(0).sum()
            
            if call_vol_total > 0:
                pcr_vol = put_vol_total / call_vol_total
                metrics['pcr_vol'] = pcr_vol
                metrics['call_vol_total'] = call_vol_total
                metrics['put_vol_total'] = put_vol_total
                log_debug(f"PCR (Vol): {pcr_vol:.3f}")
        
        # Calculate Max Pain
        if strike_cols and call_oi_cols and put_oi_cols:
            try:
                strike_col = strike_cols[0]
                call_oi_col = call_oi_cols[0]
                put_oi_col = put_oi_cols[0]
                
                clean_df = df[[strike_col, call_oi_col, put_oi_col]].dropna()
                
                if len(clean_df) > 0:
                    strikes = sorted(clean_df[strike_col].unique())
                    total_pain = []
                    
                    for strike in strikes:
                        call_pain = 0
                        put_pain = 0
                        
                        for _, row in clean_df.iterrows():
                            if row[strike_col] < strike:
                                call_pain += row[call_oi_col] * (strike - row[strike_col])
                            elif row[strike_col] > strike:
                                put_pain += row[put_oi_col] * (row[strike_col] - strike)
                        
                        total_pain.append(call_pain + put_pain)
                    
                    if total_pain:
                        max_pain_idx = np.argmin(total_pain)
                        max_pain = strikes[max_pain_idx]
                        metrics['max_pain'] = max_pain
                        log_debug(f"Max Pain: {max_pain}")
                        
                        # Support and Resistance
                        max_put_oi_idx = clean_df[put_oi_col].idxmax()
                        max_call_oi_idx = clean_df[call_oi_col].idxmax()
                        
                        support = clean_df.loc[max_put_oi_idx, strike_col]
                        resistance = clean_df.loc[max_call_oi_idx, strike_col]
                        
                        metrics['support'] = support
                        metrics['resistance'] = resistance
                        log_debug(f"Support: {support}, Resistance: {resistance}")
            except Exception as max_pain_error:
                log_debug(f"Error calculating max pain: {str(max_pain_error)}")
        
        return metrics
        
    except Exception as e:
        log_debug(f"Error calculating metrics: {str(e)}")
        return {}

def display_data_info(df):
    """Display information about the loaded data"""
    st.subheader("📊 Data Information")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Rows", len(df))
    with col2:
        st.metric("Columns", len(df.columns))
    with col3:
        st.metric("Non-null values", df.count().sum())
    
    # Show column names
    with st.expander("📋 Column Names"):
        for i, col in enumerate(df.columns, 1):
            st.write(f"{i}. {col}")
    
    # Show sample data
    with st.expander("📄 Sample Data (First 5 rows)"):
        st.dataframe(df.head(), use_container_width=True)

def main():
    st.markdown('<h1 class="main-header">⚡ NSE Options Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.header("📁 Data Source")
    
    # Method selection
    method = st.sidebar.radio(
        "Choose data source:",
        ["File Path", "File Upload"]
    )
    
    data_loaded = False
    
    if method == "File Path":
        # File path input
        file_path = st.sidebar.text_input(
            "Excel File Path:",
            value=st.session_state.excel_file_path,
            placeholder="C:/path/to/your/file.xlsx"
        )
        
        if file_path != st.session_state.excel_file_path:
            st.session_state.excel_file_path = file_path
            st.session_state.data_dict = {}
            st.session_state.debug_info = []
        
        if st.sidebar.button("🔄 Load Data", type="primary"):
            if file_path:
                with st.spinner("Loading Excel file..."):
                    result = read_excel_file(file_path)
                    
                    if "error" in result:
                        st.error(f"Error: {result['error']}")
                    else:
                        st.session_state.data_dict = result
                        st.session_state.last_refresh = datetime.now()
                        data_loaded = True
                        st.success(f"✅ Loaded {len(result)} sheets successfully!")
            else:
                st.error("Please enter a file path")
    
    else:
        # File upload
        uploaded_file = st.sidebar.file_uploader(
            "Upload Excel File:",
            type=['xlsx', 'xlsm', 'xls']
        )
        
        if uploaded_file is not None:
            with st.spinner("Processing uploaded file..."):
                result = load_uploaded_file(uploaded_file)
                
                if "error" in result:
                    st.error(f"Error: {result['error']}")
                else:
                    st.session_state.data_dict = result
                    st.session_state.last_refresh = datetime.now()
                    data_loaded = True
                    st.success(f"✅ Loaded {len(result)} sheets successfully!")
    
    # Debug information
    with st.sidebar.expander("🔧 Debug Info"):
        if st.session_state.debug_info:
            for info in st.session_state.debug_info[-10:]:  # Show last 10 entries
                st.text(info)
        else:
            st.text("No debug information yet")
        
        if st.button("Clear Debug"):
            st.session_state.debug_info = []
            st.rerun()
    
    # Main content
    if st.session_state.data_dict:
        # Show last refresh time
        if st.session_state.last_refresh:
            st.info(f"📅 Last loaded: {st.session_state.last_refresh.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Sheet selection
        sheet_names = list(st.session_state.data_dict.keys())
        
        if len(sheet_names) > 1:
            selected_sheet = st.selectbox("📊 Select sheet to analyze:", sheet_names)
        else:
            selected_sheet = sheet_names[0]
            st.info(f"📊 Analyzing sheet: {selected_sheet}")
        
        if selected_sheet in st.session_state.data_dict:
            df = st.session_state.data_dict[selected_sheet].copy()
            
            # Display data info
            display_data_info(df)
            
            # Calculate metrics
            metrics = calculate_pcr_metrics(df)
            
            if metrics:
                st.header("📊 Key Metrics")
                
                cols = st.columns(5)
                
                with cols[0]:
                    if 'call_oi_total' in metrics:
                        st.metric("📞 Call OI", f"{int(metrics['call_oi_total']):,}")
                
                with cols[1]:
                    if 'put_oi_total' in metrics:
                        st.metric("📉 Put OI", f"{int(metrics['put_oi_total']):,}")
                
                with cols[2]:
                    if 'pcr_oi' in metrics:
                        st.metric("⚖️ PCR (OI)", f"{metrics['pcr_oi']:.3f}")
                
                with cols[3]:
                    if 'pcr_vol' in metrics:
                        st.metric("📊 PCR (Vol)", f"{metrics['pcr_vol']:.3f}")
                
                with cols[4]:
                    if 'max_pain' in metrics:
                        st.metric("💰 Max Pain", f"{int(metrics['max_pain']):,}")
                
                # Support/Resistance
                if 'support' in metrics and 'resistance' in metrics:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.success(f"🟢 Support: {int(metrics['support']):,}")
                    with col2:
                        st.error(f"🔴 Resistance: {int(metrics['resistance']):,}")
                
                # Market sentiment
                if 'pcr_oi' in metrics:
                    pcr = metrics['pcr_oi']
                    if pcr > 1.3:
                        st.error("🐻 BEARISH SENTIMENT - High PCR indicates more puts")
                    elif pcr < 0.7:
                        st.success("🐂 BULLISH SENTIMENT - Low PCR indicates more calls")
                    else:
                        st.warning("⚖️ NEUTRAL SENTIMENT - Balanced put/call ratio")
            
            # Data display tabs
            tab1, tab2, tab3 = st.tabs(["📊 Filtered Data", "📋 Full Data", "📈 Charts"])
            
            with tab1:
                st.subheader("📊 Key Columns")
                
                # Find important columns
                important_keywords = ['strike', 'oi', 'volume', 'ltp', 'change', 'iv']
                important_cols = []
                
                for col in df.columns:
                    col_lower = col.lower()
                    if any(keyword in col_lower for keyword in important_keywords):
                        important_cols.append(col)
                
                if important_cols:
                    filtered_df = df[important_cols].copy()
                    
                    # Format numeric columns
                    for col in filtered_df.columns:
                        if filtered_df[col].dtype in ['float64', 'int64']:
                            if 'oi' in col.lower() and 'change' not in col.lower():
                                filtered_df[col] = filtered_df[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) else "")
                            elif filtered_df[col].dtype == 'float64':
                                filtered_df[col] = filtered_df[col].round(2)
                    
                    st.dataframe(filtered_df, use_container_width=True, height=500)
                else:
                    st.dataframe(df.head(20), use_container_width=True, height=500)
            
            with tab2:
                st.subheader("📋 Complete Dataset")
                st.dataframe(df, use_container_width=True, height=600)
                
                # Download option
                csv = df.to_csv(index=False)
                st.download_button(
                    label="📥 Download CSV",
                    data=csv,
                    file_name=f"options_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with tab3:
                st.subheader("📈 Visual Analysis")
                
                # Find numeric columns for plotting
                numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
                
                if len(numeric_cols) >= 2:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        x_col = st.selectbox("X-axis:", numeric_cols, index=0)
                    with col2:
                        y_col = st.selectbox("Y-axis:", numeric_cols, index=1 if len(numeric_cols) > 1 else 0)
                    
                    if x_col != y_col:
                        # Create scatter plot
                        plot_df = df[[x_col, y_col]].dropna()
                        if not plot_df.empty:
                            st.scatter_chart(plot_df.set_index(x_col), use_container_width=True)
                    else:
                        st.info("Please select different columns for X and Y axis")
                else:
                    st.info("Not enough numeric columns for plotting")
    
    else:
        # Welcome screen
        st.info("👋 Welcome! Please load your Excel file using the sidebar options.")
        
        st.markdown("""
        ### 📋 Instructions:
        1. **Choose data source method** (File Path or File Upload)
        2. **Provide your Excel file** with options data
        3. **Click Load Data** to analyze the file
        
        ### 📊 Expected Data Format:
        Your Excel file should contain columns like:
        - Strike prices
        - Call/Put Open Interest (CE_OI, PE_OI)
        - Call/Put Volume 
        - Last Traded Price (LTP)
        - Implied Volatility (IV)
        """)
        
        # Sample data preview
        sample_data = pd.DataFrame({
            'Strike': [22500, 22550, 22600],
            'CE_OI': [1500, 2300, 3400],
            'PE_OI': [1200, 1800, 2100],
            'CE_LTP': [245.5, 195.3, 148.7],
            'PE_LTP': [38.4, 58.9, 85.3]
        })
        st.subheader("📊 Sample Data Format:")
        st.dataframe(sample_data, use_container_width=True)

if __name__ == "__main__":
    main()
