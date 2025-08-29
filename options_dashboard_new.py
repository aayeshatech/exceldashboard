import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
import os

st.set_page_config(page_title="Options Trading Dashboard", page_icon="📊", layout="wide")

# Custom CSS for better styling
st.markdown("""
<style>
.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 1rem;
    border-radius: 10px;
    color: white;
    text-align: center;
    margin: 0.5rem 0;
}
.bullish { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); }
.bearish { background: linear-gradient(135deg, #fc466b 0%, #3f5efb 100%); }
.neutral { background: linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%); color: #333; }
.summary-box {
    background-color: #f8f9fa;
    border-left: 4px solid #007bff;
    padding: 1rem;
    margin: 1rem 0;
    border-radius: 5px;
}
</style>
""", unsafe_allow_html=True)

def read_excel_file(file_path):
    """Read all sheets from Excel file"""
    try:
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                if not df.empty:
                    data_dict[sheet_name] = df
            except Exception as e:
                st.warning(f"Could not read {sheet_name}: {str(e)}")
        
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return {}

def calculate_pcr_metrics(df):
    """Calculate PCR from options data"""
    try:
        call_oi_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
        put_oi_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
        
        if call_oi_cols and put_oi_cols:
            call_oi = pd.to_numeric(df[call_oi_cols[0]], errors='coerce').fillna(0).sum()
            put_oi = pd.to_numeric(df[put_oi_cols[0]], errors='coerce').fillna(0).sum()
            
            if call_oi > 0:
                pcr = put_oi / call_oi
                return {
                    'call_oi': call_oi,
                    'put_oi': put_oi,
                    'pcr': pcr,
                    'total_oi': call_oi + put_oi
                }
        return None
    except:
        return None

def get_max_pain(df):
    """Calculate Max Pain level"""
    try:
        strike_cols = [col for col in df.columns if 'STRIKE' in str(col).upper()]
        call_oi_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
        put_oi_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
        
        if strike_cols and call_oi_cols and put_oi_cols:
            strike_col = strike_cols[0]
            call_oi_col = call_oi_cols[0]
            put_oi_col = put_oi_cols[0]
            
            clean_df = df[[strike_col, call_oi_col, put_oi_col]].dropna()
            clean_df[strike_col] = pd.to_numeric(clean_df[strike_col], errors='coerce')
            clean_df[call_oi_col] = pd.to_numeric(clean_df[call_oi_col], errors='coerce')
            clean_df[put_oi_col] = pd.to_numeric(clean_df[put_oi_col], errors='coerce')
            clean_df = clean_df.dropna()
            
            if len(clean_df) > 0:
                strikes = sorted(clean_df[strike_col].unique())
                pain_values = []
                
                for strike in strikes:
                    total_pain = 0
                    for _, row in clean_df.iterrows():
                        if row[strike_col] < strike:
                            total_pain += row[call_oi_col] * (strike - row[strike_col])
                        elif row[strike_col] > strike:
                            total_pain += row[put_oi_col] * (row[strike_col] - strike)
                    pain_values.append(total_pain)
                
                if pain_values:
                    max_pain_idx = np.argmin(pain_values)
                    return strikes[max_pain_idx]
        return None
    except:
        return None

def get_support_resistance(df):
    """Get support and resistance levels"""
    try:
        strike_cols = [col for col in df.columns if 'STRIKE' in str(col).upper()]
        call_oi_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
        put_oi_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
        
        if strike_cols and call_oi_cols and put_oi_cols:
            strike_col = strike_cols[0]
            call_oi_col = call_oi_cols[0]
            put_oi_col = put_oi_cols[0]
            
            clean_df = df[[strike_col, call_oi_col, put_oi_col]].dropna()
            clean_df[call_oi_col] = pd.to_numeric(clean_df[call_oi_col], errors='coerce')
            clean_df[put_oi_col] = pd.to_numeric(clean_df[put_oi_col], errors='coerce')
            clean_df = clean_df.dropna()
            
            if len(clean_df) > 0:
                resistance_idx = clean_df[call_oi_col].idxmax()
                support_idx = clean_df[put_oi_col].idxmax()
                
                resistance = clean_df.loc[resistance_idx, strike_col]
                support = clean_df.loc[support_idx, strike_col]
                
                return support, resistance
        return None, None
    except:
        return None, None

def analyze_sheet_summary(data_dict):
    """Create comprehensive analysis of all sheets"""
    summary = {}
    
    # Key sheets for options analysis
    key_sheets = {
        'OC_1': 'Nifty Options',
        'OC_2': 'Bank Nifty Options', 
        'OC_3': 'Fin Nifty Options',
        'PCR & OI Chart': 'PCR Analysis',
        'Dashboard': 'Main Dashboard',
        'FII DII Data': 'FII/DII Data',
        'Globlemarket': 'Global Markets',
        'Screener': 'Stock Screener'
    }
    
    for sheet_key, sheet_desc in key_sheets.items():
        if sheet_key in data_dict:
            df = data_dict[sheet_key]
            analysis = {
                'description': sheet_desc,
                'rows': len(df),
                'columns': len(df.columns),
                'data': df
            }
            
            # Calculate metrics for option chains
            if 'OC_' in sheet_key:
                pcr_data = calculate_pcr_metrics(df)
                if pcr_data:
                    analysis['pcr'] = pcr_data['pcr']
                    analysis['call_oi'] = pcr_data['call_oi']
                    analysis['put_oi'] = pcr_data['put_oi']
                    analysis['total_oi'] = pcr_data['total_oi']
                
                max_pain = get_max_pain(df)
                if max_pain:
                    analysis['max_pain'] = max_pain
                
                support, resistance = get_support_resistance(df)
                if support and resistance:
                    analysis['support'] = support
                    analysis['resistance'] = resistance
            
            summary[sheet_key] = analysis
    
    return summary

def display_market_overview(summary):
    """Display market overview dashboard"""
    st.header("Market Overview Dashboard")
    
    # Main indices analysis
    col1, col2, col3 = st.columns(3)
    
    indices = [('OC_1', 'NIFTY', col1), ('OC_2', 'BANK NIFTY', col2), ('OC_3', 'FIN NIFTY', col3)]
    
    for sheet_key, index_name, col in indices:
        with col:
            if sheet_key in summary:
                data = summary[sheet_key]
                st.subheader(f"{index_name}")
                
                if 'pcr' in data:
                    pcr = data['pcr']
                    
                    # Determine sentiment
                    if pcr > 1.3:
                        sentiment = "BEARISH"
                        color_class = "bearish"
                    elif pcr < 0.7:
                        sentiment = "BULLISH" 
                        color_class = "bullish"
                    else:
                        sentiment = "NEUTRAL"
                        color_class = "neutral"
                    
                    st.markdown(f"""
                    <div class="metric-card {color_class}">
                        <h3>{sentiment}</h3>
                        <p>PCR: {pcr:.3f}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.metric("Call OI", f"{int(data['call_oi']):,}")
                    st.metric("Put OI", f"{int(data['put_oi']):,}")
                    
                    if 'max_pain' in data:
                        st.metric("Max Pain", f"{int(data['max_pain']):,}")
                    
                    if 'support' in data and 'resistance' in data:
                        st.success(f"Support: {int(data['support']):,}")
                        st.error(f"Resistance: {int(data['resistance']):,}")
                else:
                    st.info(f"Data loaded: {data['rows']} rows")

def display_detailed_analysis(data_dict, summary):
    """Display detailed analysis for each major sheet"""
    
    # PCR Analysis
    if 'PCR & OI Chart' in data_dict:
        st.header("PCR & Open Interest Analysis")
        df = data_dict['PCR & OI Chart']
        
        # Display PCR data
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("PCR Chart Data")
            pcr_display_df = df.head(10)
            st.dataframe(pcr_display_df, use_container_width=True)
        
        with col2:
            st.subheader("Key PCR Metrics")
            # Try to find PCR values in the data
            for col in df.columns:
                if 'PCR' in str(col).upper():
                    pcr_values = pd.to_numeric(df[col], errors='coerce').dropna()
                    if not pcr_values.empty:
                        current_pcr = pcr_values.iloc[-1] if len(pcr_values) > 0 else 0
                        st.metric(f"Current {col}", f"{current_pcr:.3f}")
    
    # FII/DII Analysis
    if 'FII DII Data' in data_dict:
        st.header("FII/DII Flow Analysis")
        df = data_dict['FII DII Data']
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Recent FII/DII Data")
            st.dataframe(df.head(10), use_container_width=True)
        
        with col2:
            st.subheader("Flow Summary")
            # Look for amount columns
            amount_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['AMOUNT', 'CRORE', 'VALUE'])]
            if amount_cols:
                for col in amount_cols[:3]:  # Show top 3 amount columns
                    try:
                        values = pd.to_numeric(df[col], errors='coerce').dropna()
                        if not values.empty:
                            total = values.sum()
                            st.metric(col, f"₹{total:,.0f} Cr")
                    except:
                        pass
    
    # Global Market Data
    if 'Globlemarket' in data_dict:
        st.header("Global Markets")
        df = data_dict['Globlemarket']
        st.dataframe(df, use_container_width=True)

def display_sheet_navigator(data_dict):
    """Display sheet navigator for detailed analysis"""
    st.header("Sheet Navigator")
    
    # Categorize sheets
    option_chains = [sheet for sheet in data_dict.keys() if 'OC_' in sheet or 'Option' in sheet]
    analysis_sheets = [sheet for sheet in data_dict.keys() if any(term in sheet for term in ['Dashboard', 'Screener', 'PCR', 'FII', 'Global'])]
    data_sheets = [sheet for sheet in data_dict.keys() if sheet not in option_chains and sheet not in analysis_sheets]
    
    tab1, tab2, tab3 = st.tabs(["Option Chains", "Analysis Sheets", "Data Sheets"])
    
    with tab1:
        for sheet in option_chains:
            with st.expander(f"{sheet} ({len(data_dict[sheet])} rows, {len(data_dict[sheet].columns)} cols)"):
                df = data_dict[sheet]
                
                # Show sample data
                st.dataframe(df.head(5), use_container_width=True)
                
                # Calculate metrics if possible
                pcr_data = calculate_pcr_metrics(df)
                if pcr_data:
                    col1, col2, col3 = st.columns(3)
                    col1.metric("PCR", f"{pcr_data['pcr']:.3f}")
                    col2.metric("Total Call OI", f"{int(pcr_data['call_oi']):,}")
                    col3.metric("Total Put OI", f"{int(pcr_data['put_oi']):,}")
    
    with tab2:
        for sheet in analysis_sheets:
            with st.expander(f"{sheet} ({len(data_dict[sheet])} rows, {len(data_dict[sheet].columns)} cols)"):
                st.dataframe(data_dict[sheet].head(10), use_container_width=True)
    
    with tab3:
        for sheet in data_sheets:
            with st.expander(f"{sheet} ({len(data_dict[sheet])} rows, {len(data_dict[sheet].columns)} cols)"):
                st.dataframe(data_dict[sheet].head(5), use_container_width=True)

def main():
    st.title("Comprehensive Options Trading Dashboard")
    st.caption(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Options Excel File", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        # Save and read file
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        with st.spinner("Loading comprehensive data..."):
            data_dict = read_excel_file(temp_path)
        
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data_dict:
            st.success(f"Successfully loaded {len(data_dict)} sheets with trading data")
            
            # Create summary analysis
            summary = analyze_sheet_summary(data_dict)
            
            # Display different views
            view_option = st.radio(
                "Select Dashboard View:",
                ["Market Overview", "Detailed Analysis", "Sheet Navigator", "Raw Data Explorer"]
            )
            
            if view_option == "Market Overview":
                display_market_overview(summary)
                
                # Quick summary box
                st.markdown("""
                <div class="summary-box">
                <h4>Trading Summary</h4>
                <p>Your Excel file contains comprehensive options data with multiple option chains, 
                FII/DII flows, global market data, and various analytical sheets. Use the different 
                view options above to explore specific aspects of the data.</p>
                </div>
                """, unsafe_allow_html=True)
            
            elif view_option == "Detailed Analysis":
                display_detailed_analysis(data_dict, summary)
            
            elif view_option == "Sheet Navigator":
                display_sheet_navigator(data_dict)
            
            elif view_option == "Raw Data Explorer":
                st.header("Raw Data Explorer")
                selected_sheet = st.selectbox("Select Sheet:", list(data_dict.keys()))
                
                if selected_sheet:
                    df = data_dict[selected_sheet]
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Rows", len(df))
                    col2.metric("Columns", len(df.columns))
                    col3.metric("Data Points", df.size)
                    
                    st.subheader(f"Data from {selected_sheet}")
                    st.dataframe(df, use_container_width=True)
                    
                    # Download option
                    csv_data = df.to_csv(index=False)
                    st.download_button(
                        f"Download {selected_sheet} as CSV",
                        csv_data,
                        f"{selected_sheet}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        "text/csv"
                    )
        else:
            st.error("Could not load data from Excel file")
    
    else:
        st.info("Upload your comprehensive options Excel file to begin analysis")
        
        st.markdown("""
        ### Expected Data Structure:
        - **Option Chains**: OC_1 (Nifty), OC_2 (Bank Nifty), etc.
        - **PCR Analysis**: PCR & OI Chart
        - **FII/DII Data**: Institutional flow data
        - **Dashboard**: Main analytical dashboard
        - **Global Markets**: International market data
        - **Screener**: Stock screening data
        """)

if __name__ == "__main__":
    main()
