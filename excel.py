import pandas as pd
import numpy as np
import streamlit as st
import warnings
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
.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 1rem;
    border-radius: 10px;
    color: white;
    text-align: center;
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
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=30)
def load_excel_data(file):
    """Load Excel data"""
    try:
        excel_file = pd.ExcelFile(file)
        data_dict = {}
        
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file, sheet_name=sheet_name)
                if not df.empty:
                    data_dict[sheet_name] = df
            except Exception as e:
                continue
                
        return data_dict
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return {}

def calculate_pcr(df):
    """Calculate Put-Call Ratio"""
    try:
        if 'CE_OI' in df.columns and 'PE_OI' in df.columns:
            total_call_oi = df['CE_OI'].sum()
            total_put_oi = df['PE_OI'].sum()
            pcr_oi = total_put_oi / total_call_oi if total_call_oi > 0 else 0
            return pcr_oi, total_call_oi, total_put_oi
        return None, None, None
    except:
        return None, None, None

def calculate_volume_pcr(df):
    """Calculate Volume PCR"""
    try:
        if 'CE_Total_Traded_Volume' in df.columns and 'PE_Total_Traded_Volume' in df.columns:
            total_call_vol = df['CE_Total_Traded_Volume'].sum()
            total_put_vol = df['PE_Total_Traded_Volume'].sum()
            pcr_vol = total_put_vol / total_call_vol if total_call_vol > 0 else 0
            return pcr_vol, total_call_vol, total_put_vol
        return None, None, None
    except:
        return None, None, None

def calculate_max_pain(df):
    """Calculate Max Pain"""
    try:
        if all(col in df.columns for col in ['Strike', 'CE_OI', 'PE_OI']):
            strikes = df['Strike'].dropna().sort_values()
            total_pain = []
            
            for strike in strikes:
                call_pain = 0
                put_pain = 0
                
                for _, row in df.iterrows():
                    if pd.notna(row['Strike']) and pd.notna(row['CE_OI']) and pd.notna(row['PE_OI']):
                        if row['Strike'] < strike:
                            call_pain += row['CE_OI'] * (strike - row['Strike'])
                        if row['Strike'] > strike:
                            put_pain += row['PE_OI'] * (row['Strike'] - strike)
                
                total_pain.append(call_pain + put_pain)
            
            if total_pain:
                max_pain_index = np.argmin(total_pain)
                return strikes.iloc[max_pain_index]
        return None
    except:
        return None

def get_support_resistance(df):
    """Get support and resistance levels"""
    try:
        if all(col in df.columns for col in ['Strike', 'CE_OI', 'PE_OI']):
            max_call_oi_idx = df['CE_OI'].idxmax()
            max_put_oi_idx = df['PE_OI'].idxmax()
            
            resistance = df.loc[max_call_oi_idx, 'Strike']
            support = df.loc[max_put_oi_idx, 'Strike']
            
            return support, resistance
        return None, None
    except:
        return None, None

def display_market_sentiment(pcr_oi):
    """Display market sentiment based on PCR"""
    if pcr_oi is None:
        return
    
    if pcr_oi > 1.3:
        st.markdown("""
        <div class="error-box">
        <strong>🐻 BEARISH SENTIMENT</strong><br>
        PCR is high ({:.3f}) - More puts than calls, indicating bearish sentiment
        </div>
        """.format(pcr_oi), unsafe_allow_html=True)
    elif pcr_oi < 0.7:
        st.markdown("""
        <div class="success-box">
        <strong>🐂 BULLISH SENTIMENT</strong><br>
        PCR is low ({:.3f}) - More calls than puts, indicating bullish sentiment
        </div>
        """.format(pcr_oi), unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="warning-box">
        <strong>⚖️ NEUTRAL SENTIMENT</strong><br>
        PCR is balanced ({:.3f}) - No clear directional bias
        </div>
        """.format(pcr_oi), unsafe_allow_html=True)

def create_simple_oi_chart(df):
    """Create simple OI chart using Streamlit native charts"""
    try:
        if all(col in df.columns for col in ['Strike', 'CE_OI', 'PE_OI']):
            # Prepare data for chart
            chart_data = df[['Strike', 'CE_OI', 'PE_OI']].copy()
            chart_data = chart_data.dropna()
            chart_data = chart_data.set_index('Strike')
            
            st.subheader("📊 Open Interest Distribution")
            st.bar_chart(chart_data)
            return True
    except Exception as e:
        st.error(f"Error creating OI chart: {e}")
        return False

def create_simple_volume_chart(df):
    """Create simple Volume chart"""
    try:
        if all(col in df.columns for col in ['Strike', 'CE_Total_Traded_Volume', 'PE_Total_Traded_Volume']):
            # Prepare data for chart
            chart_data = df[['Strike', 'CE_Total_Traded_Volume', 'PE_Total_Traded_Volume']].copy()
            chart_data = chart_data.dropna()
            chart_data = chart_data.set_index('Strike')
            chart_data.columns = ['Call Volume', 'Put Volume']
            
            st.subheader("📈 Volume Distribution")
            st.bar_chart(chart_data)
            return True
    except Exception as e:
        st.error(f"Error creating volume chart: {e}")
        return False

def create_iv_chart(df):
    """Create simple IV chart"""
    try:
        if all(col in df.columns for col in ['Strike', 'CE_IV(Spot)', 'PE_IV(Spot)']):
            chart_data = df[['Strike', 'CE_IV(Spot)', 'PE_IV(Spot)']].copy()
            chart_data = chart_data.dropna()
            chart_data = chart_data.set_index('Strike')
            chart_data.columns = ['Call IV', 'Put IV']
            
            st.subheader("📉 Implied Volatility")
            st.line_chart(chart_data)
            return True
    except Exception as e:
        st.error(f"Error creating IV chart: {e}")
        return False

def display_top_strikes(df):
    """Display top strikes by OI and Volume"""
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🔥 Top Call Activity")
        if 'CE_OI' in df.columns:
            top_call_oi = df.nlargest(5, 'CE_OI')[['Strike', 'CE_OI', 'CE_Total_Traded_Volume']].round(2)
            st.dataframe(top_call_oi, hide_index=True)
    
    with col2:
        st.subheader("🔥 Top Put Activity")
        if 'PE_OI' in df.columns:
            top_put_oi = df.nlargest(5, 'PE_OI')[['Strike', 'PE_OI', 'PE_Total_Traded_Volume']].round(2)
            st.dataframe(top_put_oi, hide_index=True)

def main():
    st.markdown('<h1 class="main-header">⚡ NSE Options Chain Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.header("📁 Upload Options Data")
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xlsm'],
        help="Upload your Live_Option_Chain_Terminal.xlsm file"
    )
    
    if uploaded_file is not None:
        # Load data
        data_dict = load_excel_data(uploaded_file)
        
        if data_dict:
            # Auto refresh
            st.sidebar.header("🔄 Settings")
            auto_refresh = st.sidebar.checkbox("Auto Refresh (30 sec)", value=False)
            
            if auto_refresh:
                st.sidebar.success("✅ Auto-refreshing enabled")
                st.rerun()
            
            # Sheet selection
            st.sidebar.header("📊 Select Sheet")
            sheet_names = list(data_dict.keys())
            
            # Filter for options sheets
            options_sheets = [sheet for sheet in sheet_names if any(x in sheet.upper() for x in ['OC_', 'OPTION', 'CHAIN'])]
            if not options_sheets:
                options_sheets = sheet_names[:5]  # Take first 5 sheets if no obvious options sheets
            
            selected_sheet = st.sidebar.selectbox("Choose Options Sheet", options_sheets)
            
            if selected_sheet and selected_sheet in data_dict:
                df = data_dict[selected_sheet].copy()
                
                # Get symbol info
                symbol = "OPTIONS"
                if 'FNO Symbol' in df.columns and len(df) > 0:
                    symbol = str(df['FNO Symbol'].iloc[0])
                
                st.subheader(f"📊 {symbol} - {selected_sheet} Analysis")
                
                # Calculate all metrics
                pcr_oi, total_call_oi, total_put_oi = calculate_pcr(df)
                pcr_vol, total_call_vol, total_put_vol = calculate_volume_pcr(df)
                max_pain = calculate_max_pain(df)
                support, resistance = get_support_resistance(df)
                
                # Display key metrics in columns
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    if total_call_oi:
                        st.metric(
                            label="📞 Total Call OI",
                            value=f"{total_call_oi:,.0f}"
                        )
                
                with col2:
                    if total_put_oi:
                        st.metric(
                            label="📉 Total Put OI", 
                            value=f"{total_put_oi:,.0f}"
                        )
                
                with col3:
                    if pcr_oi:
                        delta_color = "normal"
                        if pcr_oi > 1.2:
                            delta_color = "inverse"
                        st.metric(
                            label="⚖️ PCR (OI)",
                            value=f"{pcr_oi:.3f}"
                        )
                
                with col4:
                    if pcr_vol:
                        st.metric(
                            label="📊 PCR (Volume)",
                            value=f"{pcr_vol:.3f}"
                        )
                
                with col5:
                    if max_pain:
                        st.metric(
                            label="💰 Max Pain",
                            value=f"₹{max_pain:.0f}"
                        )
                
                # Market Sentiment
                st.header("📈 Market Sentiment Analysis")
                display_market_sentiment(pcr_oi)
                
                # Support/Resistance
                if support and resistance:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.success(f"🟢 **Support Level**: ₹{support:.0f} (Max Put OI)")
                    with col2:
                        st.error(f"🔴 **Resistance Level**: ₹{resistance:.0f} (Max Call OI)")
                
                # Create tabs for different views
                tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📊 Options Chain", 
                    "📈 Charts", 
                    "🔥 Top Strikes",
                    "📋 Raw Data", 
                    "ℹ️ Sheet Info"
                ])
                
                with tab1:
                    st.subheader(f"📊 {symbol} Options Chain")
                    
                    # Filter important columns for display
                    display_cols = []
                    important_cols = ['Strike', 'CE_OI', 'CE_OI_Change', 'CE_Total_Traded_Volume', 'CE_LTP', 
                                    'PE_OI', 'PE_OI_Change', 'PE_Total_Traded_Volume', 'PE_LTP']
                    
                    for col in important_cols:
                        if col in df.columns:
                            display_cols.append(col)
                    
                    if display_cols:
                        display_df = df[display_cols].copy()
                        # Round numeric columns
                        numeric_cols = display_df.select_dtypes(include=[np.number]).columns
                        display_df[numeric_cols] = display_df[numeric_cols].round(2)
                        
                        st.dataframe(display_df, use_container_width=True, height=500)
                    else:
                        st.dataframe(df, use_container_width=True, height=500)
                
                with tab2:
                    st.header("📈 Visual Analysis")
                    
                    # OI Chart
                    create_simple_oi_chart(df)
                    
                    st.markdown("---")
                    
                    # Volume Chart
                    create_simple_volume_chart(df)
                    
                    st.markdown("---")
                    
                    # IV Chart
                    create_iv_chart(df)
                
                with tab3:
                    st.header("🔥 Most Active Strikes")
                    display_top_strikes(df)
                    
                    # OI Changes
                    if 'CE_OI_Change' in df.columns and 'PE_OI_Change' in df.columns:
                        st.subheader("📊 Significant OI Changes")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write("**📈 Biggest Call OI Increases:**")
                            call_increases = df[df['CE_OI_Change'] > 0].nlargest(5, 'CE_OI_Change')[['Strike', 'CE_OI_Change', 'CE_OI']]
                            if not call_increases.empty:
                                st.dataframe(call_increases, hide_index=True)
                            else:
                                st.info("No significant call OI increases")
                        
                        with col2:
                            st.write("**📉 Biggest Put OI Increases:**")
                            put_increases = df[df['PE_OI_Change'] > 0].nlargest(5, 'PE_OI_Change')[['Strike', 'PE_OI_Change', 'PE_OI']]
                            if not put_increases.empty:
                                st.dataframe(put_increases, hide_index=True)
                            else:
                                st.info("No significant put OI increases")
                
                with tab4:
                    st.subheader("📋 Complete Data")
                    st.dataframe(df, use_container_width=True, height=600)
                    
                    # Download options
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
                        sheet_info.append({
                            'Sheet Name': sheet_name,
                            'Rows': len(sheet_df),
                            'Columns': len(sheet_df.columns),
                            'Has Options Data': 'Yes' if any(col.startswith(('CE_', 'PE_')) for col in sheet_df.columns) else 'No'
                        })
                    
                    sheet_info_df = pd.DataFrame(sheet_info)
                    st.dataframe(sheet_info_df, hide_index=True, use_container_width=True)
                    
                    # Show sample from other sheets
                    st.subheader("🔍 Quick Preview of Other Sheets")
                    other_sheet = st.selectbox("Select sheet to preview", [s for s in sheet_names if s != selected_sheet])
                    
                    if other_sheet:
                        st.write(f"**Preview of {other_sheet}:**")
                        preview_df = data_dict[other_sheet].head(10)
                        st.dataframe(preview_df, use_container_width=True)
        
        else:
            st.error("❌ Could not load data from the file. Please check the file format and try again.")
    
    else:
        # Welcome screen
        st.info("""
        ## 🚀 Welcome to NSE Options Chain Dashboard!
        
        **📤 Upload your Excel file to get started**
        
        ### ✨ Features:
        - 📊 **Options Chain Analysis** - Complete CE/PE data view
        - 📈 **Visual Charts** - OI, Volume, and IV analysis  
        - 💹 **PCR Monitoring** - Put-Call Ratio for market sentiment
        - ⚖️ **Max Pain Calculation** - Find equilibrium strike price
        - 🎯 **Support/Resistance** - Key levels based on OI
        - 🔥 **Active Strikes** - Most traded options
        - 📱 **Auto-Refresh** - Live data updates
        - 📥 **Data Export** - Download as CSV
        
        ### 📁 Supported Files:
        - `.xlsx` - Excel files
        - `.xlsm` - Excel files with macros
        
        ### 📋 Expected Data Format:
        Your Excel file should contain options chain data with columns like:
        - `Strike` - Strike prices
        - `CE_OI`, `PE_OI` - Call/Put Open Interest  
        - `CE_Total_Traded_Volume`, `PE_Total_Traded_Volume` - Volume
        - `CE_LTP`, `PE_LTP` - Last Traded Price
        - `CE_OI_Change`, `PE_OI_Change` - OI Changes
        """)
        
        # Sample data format
        st.subheader("📊 Sample Data Format")
        sample_data = pd.DataFrame({
            'Strike': [22500, 22550, 22600, 22650, 22700],
            'CE_OI': [1500, 2300, 3400, 2100, 1800],
            'PE_OI': [1200, 1800, 2100, 2800, 3200],
            'CE_Total_Traded_Volume': [450, 680, 920, 540, 380],
            'PE_Total_Traded_Volume': [320, 490, 650, 780, 890],
            'CE_LTP': [245.5, 195.3, 148.7, 105.2, 68.9],
            'PE_LTP': [38.4, 58.9, 85.3, 120.7, 165.2]
        })
        
        st.dataframe(sample_data, use_container_width=True)

if __name__ == "__main__":
    main()
