import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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
}
.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 1rem;
    border-radius: 10px;
    color: white;
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

def create_oi_chart(df):
    """Create OI chart"""
    try:
        if all(col in df.columns for col in ['Strike', 'CE_OI', 'PE_OI']):
            fig = go.Figure()
            
            fig.add_trace(go.Bar(
                x=df['Strike'],
                y=df['CE_OI'],
                name='Call OI',
                marker_color='green',
                opacity=0.7
            ))
            
            fig.add_trace(go.Bar(
                x=df['Strike'],
                y=-df['PE_OI'],
                name='Put OI',
                marker_color='red',
                opacity=0.7
            ))
            
            fig.update_layout(
                title='Open Interest Distribution',
                xaxis_title='Strike Price',
                yaxis_title='Open Interest',
                height=500,
                barmode='relative'
            )
            
            return fig
    except Exception as e:
        st.error(f"Error creating chart: {e}")
        return None

def create_volume_chart(df):
    """Create Volume chart"""
    try:
        if all(col in df.columns for col in ['Strike', 'CE_Total_Traded_Volume', 'PE_Total_Traded_Volume']):
            fig = go.Figure()
            
            fig.add_trace(go.Bar(
                x=df['Strike'],
                y=df['CE_Total_Traded_Volume'],
                name='Call Volume',
                marker_color='lightgreen',
                opacity=0.7
            ))
            
            fig.add_trace(go.Bar(
                x=df['Strike'],
                y=-df['PE_Total_Traded_Volume'],
                name='Put Volume',
                marker_color='lightcoral',
                opacity=0.7
            ))
            
            fig.update_layout(
                title='Volume Distribution',
                xaxis_title='Strike Price',
                yaxis_title='Volume',
                height=500,
                barmode='relative'
            )
            
            return fig
    except Exception as e:
        st.error(f"Error creating volume chart: {e}")
        return None

def main():
    st.markdown('<h1 class="main-header">⚡ NSE Options Chain Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.header("📁 Upload File")
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xlsm'],
        help="Upload your options chain Excel file"
    )
    
    if uploaded_file is not None:
        # Load data
        data_dict = load_excel_data(uploaded_file)
        
        if data_dict:
            # Auto refresh
            st.sidebar.header("🔄 Settings")
            auto_refresh = st.sidebar.checkbox("Auto Refresh", value=False)
            
            if auto_refresh:
                st.sidebar.info("Refreshing every 30 seconds")
                st.rerun()
            
            # Sheet selection
            st.sidebar.header("📊 Select Sheet")
            sheet_names = list(data_dict.keys())
            selected_sheet = st.sidebar.selectbox("Choose Sheet", sheet_names)
            
            if selected_sheet and selected_sheet in data_dict:
                df = data_dict[selected_sheet].copy()
                
                # Display basic info
                st.subheader(f"📊 {selected_sheet} Analysis")
                
                # Calculate metrics
                pcr_oi, total_call_oi, total_put_oi = calculate_pcr(df)
                max_pain = calculate_max_pain(df)
                
                # Display metrics
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    if total_call_oi:
                        st.metric("Call OI", f"{total_call_oi:,.0f}")
                
                with col2:
                    if total_put_oi:
                        st.metric("Put OI", f"{total_put_oi:,.0f}")
                
                with col3:
                    if pcr_oi:
                        st.metric("PCR (OI)", f"{pcr_oi:.3f}")
                
                with col4:
                    if max_pain:
                        st.metric("Max Pain", f"₹{max_pain:.0f}")
                
                # Create tabs
                tab1, tab2, tab3, tab4 = st.tabs(["📊 Data", "📈 OI Chart", "📊 Volume", "📋 Analysis"])
                
                with tab1:
                    st.subheader("Options Chain Data")
                    st.dataframe(df, use_container_width=True, height=600)
                    
                    # Download button
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "Download CSV",
                        csv,
                        file_name=f"{selected_sheet}_data.csv",
                        mime="text/csv"
                    )
                
                with tab2:
                    st.subheader("Open Interest Analysis")
                    oi_chart = create_oi_chart(df)
                    if oi_chart:
                        st.plotly_chart(oi_chart, use_container_width=True)
                    
                    # Top OI analysis
                    if 'CE_OI' in df.columns and 'PE_OI' in df.columns:
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write("**Top Call OI:**")
                            top_call = df.nlargest(5, 'CE_OI')[['Strike', 'CE_OI']]
                            st.dataframe(top_call, hide_index=True)
                        
                        with col2:
                            st.write("**Top Put OI:**")
                            top_put = df.nlargest(5, 'PE_OI')[['Strike', 'PE_OI']]
                            st.dataframe(top_put, hide_index=True)
                
                with tab3:
                    st.subheader("Volume Analysis")
                    vol_chart = create_volume_chart(df)
                    if vol_chart:
                        st.plotly_chart(vol_chart, use_container_width=True)
                
                with tab4:
                    st.subheader("Market Analysis")
                    
                    # Market sentiment
                    if pcr_oi:
                        if pcr_oi > 1.3:
                            st.error("🐻 Bearish Sentiment (High PCR)")
                        elif pcr_oi < 0.7:
                            st.success("🐂 Bullish Sentiment (Low PCR)")
                        else:
                            st.info("⚖️ Neutral Sentiment")
                    
                    # Support/Resistance
                    if 'CE_OI' in df.columns and 'PE_OI' in df.columns and 'Strike' in df.columns:
                        max_call_oi_strike = df.loc[df['CE_OI'].idxmax(), 'Strike']
                        max_put_oi_strike = df.loc[df['PE_OI'].idxmax(), 'Strike']
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("🔴 Resistance (Max Call OI)", f"₹{max_call_oi_strike:.0f}")
                        with col2:
                            st.metric("🟢 Support (Max Put OI)", f"₹{max_put_oi_strike:.0f}")
                
                # Additional sheets info
                st.subheader("📋 Available Sheets")
                sheet_info = pd.DataFrame({
                    'Sheet Name': list(data_dict.keys()),
                    'Rows': [len(data_dict[sheet]) for sheet in data_dict.keys()],
                    'Columns': [len(data_dict[sheet].columns) for sheet in data_dict.keys()]
                })
                st.dataframe(sheet_info, hide_index=True)
        
        else:
            st.error("Could not load data from the file. Please check the file format.")
    
    else:
        st.info("""
        **Welcome to NSE Options Dashboard!**
        
        📤 Upload your Excel file to get started
        
        **Features:**
        - 📊 Options chain analysis
        - 📈 Open Interest charts
        - 💹 PCR monitoring  
        - ⚖️ Max Pain calculation
        - 🎯 Support/Resistance levels
        
        **Supported files:** .xlsx, .xlsm
        """)

if __name__ == "__main__":
    main()
