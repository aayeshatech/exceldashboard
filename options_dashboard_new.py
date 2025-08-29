import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
import os

st.set_page_config(page_title="Trading Dashboard", page_icon="📊", layout="wide")

# Focused CSS for trading dashboard
st.markdown("""
<style>
.trading-card {
    background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
    text-align: center;
}
.bullish-card {
    background: linear-gradient(135deg, #56ab2f 0%, #a8e6cf 100%);
    color: #2d5016;
}
.bearish-card {
    background: linear-gradient(135deg, #cb2d3e 0%, #ef473a 100%);
    color: white;
}
.neutral-card {
    background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    color: white;
}
.summary-section {
    background-color: #f8f9fa;
    border: 1px solid #dee2e6;
    border-radius: 8px;
    padding: 1rem;
    margin: 1rem 0;
}
.metric-row {
    display: flex;
    justify-content: space-between;
    padding: 0.5rem 0;
    border-bottom: 1px solid #eee;
}
</style>
""", unsafe_allow_html=True)

def read_excel_comprehensive(file_path):
    """Read Excel file focusing on top 10 trading sheets"""
    priority_sheets = [
        'Dashboard', 'Screener', 'Multistrike', 'Buyer&Sellers Graph', 
        'Participant Chart', 'PCR & OI Chart', 'Stock Dashboard', 
        'Sector Dashboard', 'Greeks', 'FII DII Data'
    ]
    
    try:
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        # Read priority sheets first
        for sheet_name in priority_sheets:
            if sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    if not df.empty:
                        data_dict[sheet_name] = df
                        st.success(f"Loaded {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
                except Exception as e:
                    st.warning(f"Could not read {sheet_name}: {str(e)}")
        
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return {}

def analyze_dashboard_sheet(df):
    """Analyze main Dashboard sheet"""
    summary = {
        'total_stocks': 0,
        'bullish_stocks': 0,
        'bearish_stocks': 0,
        'neutral_stocks': 0,
        'top_gainers': [],
        'top_losers': [],
        'high_volume_stocks': []
    }
    
    try:
        # Look for percentage change columns
        change_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['%', 'CHANGE', 'CHG'])]
        volume_cols = [col for col in df.columns if 'VOLUME' in str(col).upper()]
        symbol_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['SYMBOL', 'STOCK', 'NAME'])]
        
        if change_cols and symbol_cols:
            change_col = change_cols[0]
            symbol_col = symbol_cols[0]
            
            # Clean and convert data
            clean_df = df[[symbol_col, change_col]].copy()
            clean_df[change_col] = pd.to_numeric(clean_df[change_col], errors='coerce')
            clean_df = clean_df.dropna()
            
            if len(clean_df) > 0:
                summary['total_stocks'] = len(clean_df)
                summary['bullish_stocks'] = len(clean_df[clean_df[change_col] > 1])
                summary['bearish_stocks'] = len(clean_df[clean_df[change_col] < -1])
                summary['neutral_stocks'] = len(clean_df[abs(clean_df[change_col]) <= 1])
                
                # Top gainers and losers
                top_gainers = clean_df.nlargest(5, change_col)
                top_losers = clean_df.nsmallest(5, change_col)
                
                summary['top_gainers'] = [(row[symbol_col], row[change_col]) for _, row in top_gainers.iterrows()]
                summary['top_losers'] = [(row[symbol_col], row[change_col]) for _, row in top_losers.iterrows()]
        
        return summary
    except Exception as e:
        st.warning(f"Error analyzing dashboard: {e}")
        return summary

def analyze_pcr_data(df):
    """Analyze PCR & OI Chart data"""
    pcr_summary = {}
    
    try:
        # Look for PCR columns
        pcr_cols = [col for col in df.columns if 'PCR' in str(col).upper()]
        oi_cols = [col for col in df.columns if 'OI' in str(col).upper() and 'PCR' not in str(col).upper()]
        
        if pcr_cols:
            for col in pcr_cols:
                values = pd.to_numeric(df[col], errors='coerce').dropna()
                if not values.empty:
                    current_pcr = values.iloc[-1] if len(values) > 0 else 0
                    pcr_summary[col] = current_pcr
        
        return pcr_summary
    except Exception as e:
        st.warning(f"Error analyzing PCR data: {e}")
        return {}

def analyze_fii_dii_data(df):
    """Analyze FII DII flow data"""
    fii_dii_summary = {}
    
    try:
        # Look for FII/DII amount columns
        amount_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['FII', 'DII', 'AMOUNT', 'CRORE'])]
        date_cols = [col for col in df.columns if 'DATE' in str(col).upper()]
        
        if amount_cols:
            for col in amount_cols[:3]:  # Top 3 amount columns
                values = pd.to_numeric(df[col], errors='coerce').dropna()
                if not values.empty:
                    total_flow = values.sum()
                    recent_flow = values.iloc[-1] if len(values) > 0 else 0
                    fii_dii_summary[col] = {
                        'total': total_flow,
                        'recent': recent_flow
                    }
        
        return fii_dii_summary
    except Exception as e:
        st.warning(f"Error analyzing FII/DII data: {e}")
        return {}

def analyze_sector_data(df):
    """Analyze Sector Dashboard data"""
    sector_summary = {}
    
    try:
        # Look for sector performance columns
        sector_cols = [col for col in df.columns if 'SECTOR' in str(col).upper()]
        change_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['%', 'CHANGE', 'CHG'])]
        
        if sector_cols and change_cols:
            sector_col = sector_cols[0]
            change_col = change_cols[0]
            
            clean_df = df[[sector_col, change_col]].copy()
            clean_df[change_col] = pd.to_numeric(clean_df[change_col], errors='coerce')
            clean_df = clean_df.dropna()
            
            if len(clean_df) > 0:
                top_sectors = clean_df.nlargest(5, change_col)
                bottom_sectors = clean_df.nsmallest(5, change_col)
                
                sector_summary['top_performing'] = [(row[sector_col], row[change_col]) for _, row in top_sectors.iterrows()]
                sector_summary['underperforming'] = [(row[sector_col], row[change_col]) for _, row in bottom_sectors.iterrows()]
        
        return sector_summary
    except Exception as e:
        st.warning(f"Error analyzing sector data: {e}")
        return {}

def display_trading_summary(dashboard_summary, pcr_summary, fii_dii_summary, sector_summary):
    """Display comprehensive trading summary"""
    
    st.header("Market Summary Dashboard")
    
    # Market Overview Cards
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_stocks = dashboard_summary.get('total_stocks', 0)
        st.markdown(f"""
        <div class="trading-card">
            <h3>{total_stocks}</h3>
            <p>Total Stocks</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        bullish_stocks = dashboard_summary.get('bullish_stocks', 0)
        st.markdown(f"""
        <div class="trading-card bullish-card">
            <h3>{bullish_stocks}</h3>
            <p>Bullish Stocks</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        bearish_stocks = dashboard_summary.get('bearish_stocks', 0)
        st.markdown(f"""
        <div class="trading-card bearish-card">
            <h3>{bearish_stocks}</h3>
            <p>Bearish Stocks</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        neutral_stocks = dashboard_summary.get('neutral_stocks', 0)
        st.markdown(f"""
        <div class="trading-card neutral-card">
            <h3>{neutral_stocks}</h3>
            <p>Neutral Stocks</p>
        </div>
        """, unsafe_allow_html=True)
    
    # PCR Analysis
    if pcr_summary:
        st.subheader("Put-Call Ratio Analysis")
        pcr_cols = st.columns(len(pcr_summary))
        
        for i, (pcr_name, pcr_value) in enumerate(pcr_summary.items()):
            with pcr_cols[i]:
                if pcr_value > 1.3:
                    sentiment = "BEARISH"
                    card_class = "bearish-card"
                elif pcr_value < 0.7:
                    sentiment = "BULLISH"
                    card_class = "bullish-card"
                else:
                    sentiment = "NEUTRAL"
                    card_class = "neutral-card"
                
                st.markdown(f"""
                <div class="trading-card {card_class}">
                    <h4>{pcr_name}</h4>
                    <h3>{pcr_value:.3f}</h3>
                    <p>{sentiment}</p>
                </div>
                """, unsafe_allow_html=True)
    
    # Market Movers
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Top Gainers")
        top_gainers = dashboard_summary.get('top_gainers', [])
        if top_gainers:
            for stock, change in top_gainers[:5]:
                st.success(f"{stock}: +{change:.2f}%")
        else:
            st.info("No gainer data available")
    
    with col2:
        st.subheader("Top Losers")
        top_losers = dashboard_summary.get('top_losers', [])
        if top_losers:
            for stock, change in top_losers[:5]:
                st.error(f"{stock}: {change:.2f}%")
        else:
            st.info("No loser data available")
    
    # FII/DII Flows
    if fii_dii_summary:
        st.subheader("FII/DII Flow Analysis")
        fii_cols = st.columns(len(fii_dii_summary))
        
        for i, (flow_name, flow_data) in enumerate(fii_dii_summary.items()):
            with fii_cols[i]:
                recent_flow = flow_data.get('recent', 0)
                total_flow = flow_data.get('total', 0)
                
                flow_color = "success" if recent_flow > 0 else "error" if recent_flow < 0 else "info"
                
                st.markdown(f"""
                <div class="summary-section">
                    <h5>{flow_name}</h5>
                    <div class="metric-row">
                        <span>Recent:</span>
                        <span>₹{recent_flow:,.0f} Cr</span>
                    </div>
                    <div class="metric-row">
                        <span>Total:</span>
                        <span>₹{total_flow:,.0f} Cr</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    
    # Sector Performance
    if sector_summary:
        st.subheader("Sector Performance")
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Top Performing Sectors**")
            top_sectors = sector_summary.get('top_performing', [])
            for sector, change in top_sectors[:5]:
                st.success(f"{sector}: +{change:.2f}%")
        
        with col2:
            st.write("**Underperforming Sectors**")
            bottom_sectors = sector_summary.get('underperforming', [])
            for sector, change in bottom_sectors[:5]:
                st.error(f"{sector}: {change:.2f}%")

def display_detailed_sheets(data_dict):
    """Display detailed view of individual sheets"""
    st.header("Detailed Sheet Analysis")
    
    sheet_tabs = st.tabs(list(data_dict.keys()))
    
    for i, (sheet_name, df) in enumerate(data_dict.items()):
        with sheet_tabs[i]:
            st.subheader(f"{sheet_name} Analysis")
            
            # Basic metrics
            col1, col2, col3 = st.columns(3)
            col1.metric("Rows", len(df))
            col2.metric("Columns", len(df.columns))
            col3.metric("Data Points", df.size)
            
            # Show sample data
            st.write("**Sample Data:**")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Column summary
            with st.expander("Column Details"):
                for col in df.columns:
                    col_info = f"{col} - {df[col].dtype}"
                    if df[col].dtype in ['float64', 'int64']:
                        col_info += f" (Min: {df[col].min():.2f}, Max: {df[col].max():.2f})"
                    st.write(col_info)

def main():
    st.title("Comprehensive Trading Dashboard")
    st.caption(f"Live Market Analysis - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    uploaded_file = st.file_uploader("Upload Trading Excel File", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        with st.spinner("Loading and analyzing top 10 trading sheets..."):
            data_dict = read_excel_comprehensive(temp_path)
        
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data_dict:
            st.success(f"Loaded {len(data_dict)} priority sheets for analysis")
            
            # Analyze key sheets
            dashboard_summary = {}
            pcr_summary = {}
            fii_dii_summary = {}
            sector_summary = {}
            
            if 'Dashboard' in data_dict:
                dashboard_summary = analyze_dashboard_sheet(data_dict['Dashboard'])
            
            if 'PCR & OI Chart' in data_dict:
                pcr_summary = analyze_pcr_data(data_dict['PCR & OI Chart'])
            
            if 'FII DII Data' in data_dict:
                fii_dii_summary = analyze_fii_dii_data(data_dict['FII DII Data'])
            
            if 'Sector Dashboard' in data_dict:
                sector_summary = analyze_sector_data(data_dict['Sector Dashboard'])
            
            # Display options
            view_mode = st.radio(
                "Select View:",
                ["Trading Summary", "Detailed Sheet Analysis", "Raw Data View"]
            )
            
            if view_mode == "Trading Summary":
                display_trading_summary(dashboard_summary, pcr_summary, fii_dii_summary, sector_summary)
            
            elif view_mode == "Detailed Sheet Analysis":
                display_detailed_sheets(data_dict)
            
            elif view_mode == "Raw Data View":
                selected_sheet = st.selectbox("Select Sheet:", list(data_dict.keys()))
                if selected_sheet:
                    st.subheader(f"Complete {selected_sheet} Data")
                    st.dataframe(data_dict[selected_sheet], use_container_width=True)
                    
                    # Download option
                    csv_data = data_dict[selected_sheet].to_csv(index=False)
                    st.download_button(
                        f"Download {selected_sheet}",
                        csv_data,
                        f"{selected_sheet}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        "text/csv"
                    )
        else:
            st.error("Could not load any trading sheets from the file")
    
    else:
        st.info("Upload your comprehensive trading Excel file to begin analysis")
        
        st.markdown("""
        ### Priority Sheets Analyzed:
        1. **Dashboard** - Main stock performance overview
        2. **Screener** - Stock screening results
        3. **PCR & OI Chart** - Put-Call ratio analysis
        4. **FII DII Data** - Institutional flow analysis
        5. **Sector Dashboard** - Sector performance metrics
        6. **Greeks** - Options Greeks analysis
        7. **Stock Dashboard** - Individual stock analysis
        8. **Multistrike** - Multi-strike options data
        9. **Participant Chart** - Market participant data
        10. **Buyer&Sellers Graph** - Supply/demand analysis
        """)

if __name__ == "__main__":
    main()
