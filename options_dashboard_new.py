import pandas as pd
import streamlit as st
import subprocess
import sys
import os
from datetime import datetime

st.set_page_config(page_title="Options Dashboard", page_icon="📊", layout="wide")

def install_openpyxl():
    """Install openpyxl package"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        return True
    except Exception as e:
        st.error(f"Failed to install openpyxl: {e}")
        return False

def check_openpyxl():
    """Check if openpyxl is available"""
    try:
        import openpyxl
        return True
    except ImportError:
        return False

st.title("NSE Options Dashboard")

# Check for openpyxl
if not check_openpyxl():
    st.error("Missing openpyxl dependency")
    
    if st.button("Install openpyxl Now"):
        with st.spinner("Installing openpyxl..."):
            if install_openpyxl():
                st.success("openpyxl installed successfully! Please refresh the page.")
                st.balloons()
            else:
                st.error("Installation failed. Please run manually: pip install openpyxl")
    
    st.info("Alternative: Run this command in your terminal:")
    st.code("pip install openpyxl")
    st.stop()

# Now openpyxl should be available
import openpyxl

def read_excel_data(file_path):
    """Read Excel file with all sheets"""
    try:
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        data_dict = {}
        
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                if not df.empty:
                    data_dict[sheet_name] = df
                    st.success(f"Loaded {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
            except Exception as e:
                st.warning(f"Could not read {sheet_name}: {e}")
        
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        return {}

def calculate_metrics(df):
    """Calculate options metrics"""
    metrics = {}
    
    # Find columns
    call_oi_cols = [col for col in df.columns if 'CE' in col.upper() and 'OI' in col.upper() and 'CHANGE' not in col.upper()]
    put_oi_cols = [col for col in df.columns if 'PE' in col.upper() and 'OI' in col.upper() and 'CHANGE' not in col.upper()]
    
    if call_oi_cols and put_oi_cols:
        call_oi = df[call_oi_cols[0]].fillna(0).sum()
        put_oi = df[put_oi_cols[0]].fillna(0).sum()
        
        if call_oi > 0:
            pcr = put_oi / call_oi
            metrics = {
                'call_oi': call_oi,
                'put_oi': put_oi,
                'pcr': pcr
            }
    
    return metrics

# File upload
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xlsm', 'xls'])

if uploaded_file:
    # Save temporarily
    temp_path = f"temp_{uploaded_file.name}"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Read data
    with st.spinner("Reading Excel file..."):
        data_dict = read_excel_data(temp_path)
    
    # Clean up
    try:
        os.remove(temp_path)
    except:
        pass
    
    if data_dict:
        st.success(f"Successfully loaded {len(data_dict)} sheets")
        
        # Sheet selection
        sheet_names = list(data_dict.keys())
        if len(sheet_names) > 1:
            selected_sheet = st.selectbox("Select Sheet", sheet_names)
        else:
            selected_sheet = sheet_names[0]
            st.info(f"Analyzing: {selected_sheet}")
        
        if selected_sheet in data_dict:
            df = data_dict[selected_sheet]
            
            # Basic info
            st.subheader("Data Overview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Rows", len(df))
            col2.metric("Columns", len(df.columns))
            col3.metric("Non-null", df.count().sum())
            
            # Calculate metrics
            metrics = calculate_metrics(df)
            
            if metrics:
                st.subheader("Options Metrics")
                col1, col2, col3 = st.columns(3)
                col1.metric("Call OI", f"{int(metrics['call_oi']):,}")
                col2.metric("Put OI", f"{int(metrics['put_oi']):,}")
                col3.metric("PCR", f"{metrics['pcr']:.3f}")
                
                # Market sentiment
                pcr = metrics['pcr']
                if pcr > 1.3:
                    st.error("Bearish Sentiment (High PCR)")
                elif pcr < 0.7:
                    st.success("Bullish Sentiment (Low PCR)")
                else:
                    st.warning("Neutral Sentiment")
            
            # Data tabs
            tab1, tab2 = st.tabs(["Sample Data", "Full Data"])
            
            with tab1:
                st.dataframe(df.head(20))
            
            with tab2:
                st.dataframe(df)
                
                # Download
                csv = df.to_csv(index=False)
                st.download_button(
                    "Download CSV",
                    csv,
                    f"data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    "text/csv"
                )
    else:
        st.error("No data could be loaded from the Excel file")

else:
    st.info("Upload your Excel file to get started")
    
    # Show system status
    st.subheader("System Status")
    st.write("openpyxl:", "✅ Ready" if check_openpyxl() else "❌ Missing")
