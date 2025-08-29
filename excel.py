import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
import time
from datetime import datetime, timedelta
warnings.filterwarnings('ignore')

# Set page config
st.set_page_config(
    page_title="NSE Options Chain Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
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
    .sub-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .option-card {
        background-color: #f8f9fa;
        border: 2px solid #e9ecef;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .call-option {
        border-left: 5px solid #28a745;
    }
    .put-option {
        border-left: 5px solid #dc3545;
    }
    .stSelectbox > div > div > div {
        background-color: #f0f2f6;
    }
    .highlight-positive {
        color: #28a745;
        font-weight: bold;
    }
    .highlight-negative {
        color: #dc3545;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=30)  # Cache for 30 seconds for real-time updates
def load_options_data(file_path):
    """Load options chain data from Excel file"""
    try:
        # Load all sheets
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                data_dict[sheet_name] = df
            except:
                continue
                
        return data_dict
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return {}

def calculate_option_metrics(df):
    """Calculate key options metrics"""
    metrics = {}
    
    if 'CE_OI' in df.columns and 'PE_OI' in df.columns:
        # Total Call and Put OI
        metrics['total_call_oi'] = df['CE_OI'].sum()
        metrics['total_put_oi'] = df['PE_OI'].sum()
        metrics['pcr_oi'] = metrics['total_put_oi'] / metrics['total_call_oi'] if metrics['total_call_oi'] > 0 else 0
    
    if 'CE_Total_Traded_Volume' in df.columns and 'PE_Total_Traded_Volume' in df.columns:
        # Total Call and Put Volume
        metrics['total_call_volume'] = df['CE_Total_Traded_Volume'].sum()
        metrics['total_put_volume'] = df['PE_Total_Traded_Volume'].sum()
        metrics['pcr_volume'] = metrics['total_put_volume'] / metrics['total_call_volume'] if metrics['total_call_volume'] > 0 else 0
    
    # Max Pain calculation
    if all(col in df.columns for col in ['Strike', 'CE_OI', 'PE_OI']):
        max_pain = calculate_max_pain(df)
        metrics['max_pain'] = max_pain
    
    return metrics

def calculate_max_pain(df):
    """Calculate Max Pain strike price"""
    try:
        strikes = df['Strike'].dropna().sort_values()
        total_pain = []
        
        for strike in strikes:
            call_pain = df[df['Strike'] < strike]['CE_OI'].sum() * (strike - df[df['Strike'] < strike]['Strike']).sum()
            put_pain = df[df['Strike'] > strike]['PE_OI'].sum() * (df[df['Strike'] > strike]['Strike'] - strike).sum()
            total_pain.append(call_pain + put_pain)
        
        max_pain_strike = strikes.iloc[np.argmin(total_pain)]
        return max_pain_strike
    except:
        return None

def create_oi_chart(df, symbol):
    """Create Open Interest chart"""
    fig = go.Figure()
    
    if all(col in df.columns for col in ['Strike', 'CE_OI', 'PE_OI']):
        # Call OI
        fig.add_trace(go.Bar(
            x=df['Strike'],
            y=df['CE_OI'],
            name='Call OI',
            marker_color='green',
            opacity=0.7,
            yaxis='y'
        ))
        
        # Put OI (negative for visual effect)
        fig.add_trace(go.Bar(
            x=df['Strike'],
            y=-df['PE_OI'],
            name='Put OI',
            marker_color='red',
            opacity=0.7,
            yaxis='y'
        ))
    
    fig.update_layout(
        title=f'{symbol} - Open Interest Distribution',
        xaxis_title='Strike Price',
        yaxis_title='Open Interest',
        hovermode='x unified',
        height=500,
        barmode='relative'
    )
    
    return fig

def create_volume_chart(df, symbol):
    """Create Volume chart"""
    fig = go.Figure()
    
    if all(col in df.columns for col in ['Strike', 'CE_Total_Traded_Volume', 'PE_Total_Traded_Volume']):
        # Call Volume
        fig.add_trace(go.Bar(
            x=df['Strike'],
            y=df['CE_Total_Traded_Volume'],
            name='Call Volume',
            marker_color='lightgreen',
            opacity=0.7
        ))
        
        # Put Volume (negative for visual effect)
        fig.add_trace(go.Bar(
            x=df['Strike'],
            y=-df['PE_Total_Traded_Volume'],
            name='Put Volume',
            marker_color='lightcoral',
            opacity=0.7
        ))
    
    fig.update_layout(
        title=f'{symbol} - Volume Distribution',
        xaxis_title='Strike Price',
        yaxis_title='Volume',
        hovermode='x unified',
        height=500,
        barmode='relative'
    )
    
    return fig

def create_iv_chart(df, symbol):
    """Create Implied Volatility chart"""
    fig = go.Figure()
    
    if all(col in df.columns for col in ['Strike', 'CE_IV(Spot)', 'PE_IV(Spot)']):
        # Call IV
        fig.add_trace(go.Scatter(
            x=df['Strike'],
            y=df['CE_IV(Spot)'],
            mode='lines+markers',
            name='Call IV',
            line=dict(color='green', width=2),
            marker=dict(size=6)
        ))
        
        # Put IV
        fig.add_trace(go.Scatter(
            x=df['Strike'],
            y=df['PE_IV(Spot)'],
            mode='lines+markers',
            name='Put IV',
            line=dict(color='red', width=2),
            marker=dict(size=6)
        ))
    
    fig.update_layout(
        title=f'{symbol} - Implied Volatility Smile',
        xaxis_title='Strike Price',
        yaxis_title='Implied Volatility',
        hovermode='x unified',
        height=500
    )
    
    return fig

def create_greeks_chart(df, symbol, greek='Delta'):
    """Create Greeks chart"""
    fig = go.Figure()
    
    call_greek = f'CE_{greek}(Spot)'
    put_greek = f'PE_{greek}(Spot)'
    
    if all(col in df.columns for col in ['Strike', call_greek, put_greek]):
        # Call Greek
        fig.add_trace(go.Scatter(
            x=df['Strike'],
            y=df[call_greek],
            mode='lines+markers',
            name=f'Call {greek}',
            line=dict(color='green', width=2)
        ))
        
        # Put Greek
        fig.add_trace(go.Scatter(
            x=df['Strike'],
            y=df[put_greek],
            mode='lines+markers',
            name=f'Put {greek}',
            line=dict(color='red', width=2)
        ))
    
    fig.update_layout(
        title=f'{symbol} - {greek} Distribution',
        xaxis_title='Strike Price',
        yaxis_title=greek,
        hovermode='x unified',
        height=400
    )
    
    return fig

def display_option_chain_table(df, symbol):
    """Display formatted option chain table"""
    st.subheader(f"📊 {symbol} Options Chain")
    
    if df.empty:
        st.warning("No data available for this symbol")
        return
    
    # Select key columns for display
    display_columns = []
    
    # Call options columns
    call_cols = ['CE_OI', 'CE_OI_Change', 'CE_Total_Traded_Volume', 'CE_LTP', 'CE_LTP_Change', 'CE_IV(Spot)']
    # Put options columns  
    put_cols = ['PE_OI', 'PE_OI_Change', 'PE_Total_Traded_Volume', 'PE_LTP', 'PE_LTP_Change', 'PE_IV(Spot)']
    
    # Build display dataframe
    display_df = pd.DataFrame()
    
    if 'Strike' in df.columns:
        display_df['Strike'] = df['Strike']
    
    # Add available call columns
    for col in call_cols:
        if col in df.columns:
            display_df[col.replace('CE_', 'C_')] = df[col]
    
    # Add available put columns
    for col in put_cols:
        if col in df.columns:
            display_df[col.replace('PE_', 'P_')] = df[col]
    
    # Format the dataframe for better display
    numeric_cols = display_df.select_dtypes(include=[np.number]).columns
    display_df[numeric_cols] = display_df[numeric_cols].round(2)
    
    # Color coding for changes
    def highlight_changes(val):
        try:
            if 'Change' in str(val):
                if val > 0:
                    return 'background-color: lightgreen'
                elif val < 0:
                    return 'background-color: lightcoral'
        except:
            pass
        return ''
    
    styled_df = display_df.style.applymap(highlight_changes)
    st.dataframe(styled_df, use_container_width=True, height=600)

def main():
    st.markdown('<h1 class="main-header">⚡ Live NSE Options Chain Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar configuration
    st.sidebar.header("⚙️ Dashboard Configuration")
    
    # File uploader
    uploaded_file = st.sidebar.file_uploader(
        "Upload Options Chain Excel File",
        type=['xlsx', 'xlsm'],
        help="Upload your Live_Option_Chain_Terminal.xlsm file"
    )
    
    if uploaded_file is not None:
        # Load data
        data_dict = load_options_data(uploaded_file)
        
        if data_dict:
            # Auto refresh settings
            st.sidebar.header("🔄 Auto Refresh")
            auto_refresh = st.sidebar.checkbox("Enable Auto Refresh", value=True)
            refresh_interval = st.sidebar.slider("Refresh Interval (seconds)", 5, 60, 15)
            
            if auto_refresh:
                st.sidebar.info(f"Auto refreshing every {refresh_interval} seconds")
                time.sleep(refresh_interval)
                st.rerun()
            
            # Symbol selection
            st.sidebar.header("📈 Symbol Selection")
            
            # Get available option chains
            option_sheets = [sheet for sheet in data_dict.keys() if 'OC_' in sheet or 'Option' in sheet]
            
            if option_sheets:
                selected_sheet = st.sidebar.selectbox("Select Options Chain", option_sheets)
                
                if selected_sheet in data_dict:
                    df = data_dict[selected_sheet].copy()
                    
                    # Get symbol from the data
                    symbol = "OPTIONS"
                    if 'FNO Symbol' in df.columns:
                        symbol = df['FNO Symbol'].iloc[0] if len(df) > 0 else "OPTIONS"
                    
                    # Main dashboard layout
                    st.markdown(f'<h2 class="sub-header">📊 {symbol} Options Analysis</h2>', unsafe_allow_html=True)
                    
                    # Calculate metrics
                    metrics = calculate_option_metrics(df)
                    
                    # Display key metrics
                    col1, col2, col3, col4, col5 = st.columns(5)
                    
                    with col1:
                        if 'total_call_oi' in metrics:
                            st.metric(
                                "Total Call OI",
                                f"{metrics['total_call_oi']:,.0f}",
                                delta=None
                            )
                    
                    with col2:
                        if 'total_put_oi' in metrics:
                            st.metric(
                                "Total Put OI", 
                                f"{metrics['total_put_oi']:,.0f}",
                                delta=None
                            )
                    
                    with col3:
                        if 'pcr_oi' in metrics:
                            pcr_color = "normal"
                            if metrics['pcr_oi'] > 1.2:
                                pcr_color = "inverse"
                            elif metrics['pcr_oi'] < 0.8:
                                pcr_color = "inverse"
                            
                            st.metric(
                                "PCR (OI)",
                                f"{metrics['pcr_oi']:.3f}",
                                delta=None
                            )
                    
                    with col4:
                        if 'pcr_volume' in metrics:
                            st.metric(
                                "PCR (Volume)",
                                f"{metrics['pcr_volume']:.3f}",
                                delta=None
                            )
                    
                    with col5:
                        if 'max_pain' in metrics and metrics['max_pain']:
                            st.metric(
                                "Max Pain",
                                f"₹{metrics['max_pain']:.0f}",
                                delta=None
                            )
                    
                    # Create tabs for different views
                    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                        "📊 Option Chain", 
                        "📈 Open Interest", 
                        "📊 Volume", 
                        "📈 IV Smile", 
                        "🔢 Greeks",
                        "💹 Analysis"
                    ])
                    
                    with tab1:
                        display_option_chain_table(df, symbol)
                        
                        # Download option
                        csv = df.to_csv(index=False)
                        st.download_button(
                            label="📥 Download Options Data as CSV",
                            data=csv,
                            file_name=f"{symbol}_options_chain.csv",
                            mime="text/csv"
                        )
                    
                    with tab2:
                        fig_oi = create_oi_chart(df, symbol)
                        st.plotly_chart(fig_oi, use_container_width=True)
                        
                        # OI Analysis
                        if 'CE_OI' in df.columns and 'PE_OI' in df.columns:
                            st.subheader("📊 OI Analysis")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write("**Top 5 Call OI Strikes:**")
                                top_call_oi = df.nlargest(5, 'CE_OI')[['Strike', 'CE_OI']]
                                st.dataframe(top_call_oi, hide_index=True)
                            
                            with col2:
                                st.write("**Top 5 Put OI Strikes:**")
                                top_put_oi = df.nlargest(5, 'PE_OI')[['Strike', 'PE_OI']]
                                st.dataframe(top_put_oi, hide_index=True)
                    
                    with tab3:
                        fig_vol = create_volume_chart(df, symbol)
                        st.plotly_chart(fig_vol, use_container_width=True)
                        
                        # Volume Analysis
                        if all(col in df.columns for col in ['CE_Total_Traded_Volume', 'PE_Total_Traded_Volume']):
                            st.subheader("📊 Volume Analysis")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write("**Top 5 Call Volume Strikes:**")
                                top_call_vol = df.nlargest(5, 'CE_Total_Traded_Volume')[['Strike', 'CE_Total_Traded_Volume']]
                                st.dataframe(top_call_vol, hide_index=True)
                            
                            with col2:
                                st.write("**Top 5 Put Volume Strikes:**")
                                top_put_vol = df.nlargest(5, 'PE_Total_Traded_Volume')[['Strike', 'PE_Total_Traded_Volume']]
                                st.dataframe(top_put_vol, hide_index=True)
                    
                    with tab4:
                        fig_iv = create_iv_chart(df, symbol)
                        st.plotly_chart(fig_iv, use_container_width=True)
                        
                        # IV Statistics
                        if all(col in df.columns for col in ['CE_IV(Spot)', 'PE_IV(Spot)']):
                            st.subheader("📊 IV Statistics")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write("**Call IV Stats:**")
                                call_iv_stats = df['CE_IV(Spot)'].describe()
                                st.dataframe(call_iv_stats.to_frame('Call IV'), use_container_width=True)
                            
                            with col2:
                                st.write("**Put IV Stats:**")
                                put_iv_stats = df['PE_IV(Spot)'].describe()
                                st.dataframe(put_iv_stats.to_frame('Put IV'), use_container_width=True)
                    
                    with tab5:
                        # Greeks selection
                        greek_options = ['Delta', 'Theta', 'Vega', 'Gamma', 'Rho']
                        selected_greek = st.selectbox("Select Greek to Display", greek_options)
                        
                        fig_greek = create_greeks_chart(df, symbol, selected_greek)
                        st.plotly_chart(fig_greek, use_container_width=True)
                        
                        # Greeks summary
                        greek_cols = [f'CE_{selected_greek}(Spot)', f'PE_{selected_greek}(Spot)']
                        if all(col in df.columns for col in greek_cols):
                            st.subheader(f"📊 {selected_greek} Summary")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric(f"Avg Call {selected_greek}", f"{df[greek_cols[0]].mean():.4f}")
                            with col2:
                                st.metric(f"Avg Put {selected_greek}", f"{df[greek_cols[1]].mean():.4f}")
                    
                    with tab6:
                        st.subheader("💹 Market Analysis")
                        
                        # Support and Resistance levels based on OI
                        if all(col in df.columns for col in ['Strike', 'CE_OI', 'PE_OI']):
                            # Calculate support and resistance
                            max_put_oi_strike = df.loc[df['PE_OI'].idxmax(), 'Strike']
                            max_call_oi_strike = df.loc[df['CE_OI'].idxmax(), 'Strike']
                            
                            col1, col2, col3 = st.columns(3)
                            
                            with col1:
                                st.metric("🔴 Resistance (Max Call OI)", f"₹{max_call_oi_strike:.0f}")
                            
                            with col2:
                                st.metric("🟢 Support (Max Put OI)", f"₹{max_put_oi_strike:.0f}")
                            
                            with col3:
                                if 'max_pain' in metrics and metrics['max_pain']:
                                    st.metric("⚖️ Max Pain", f"₹{metrics['max_pain']:.0f}")
                        
                        # Market sentiment
                        st.subheader("📊 Market Sentiment")
                        
                        if 'pcr_oi' in metrics:
                            pcr = metrics['pcr_oi']
                            if pcr > 1.3:
                                sentiment = "🐻 Bearish (High PCR)"
                                sentiment_color = "red"
                            elif pcr < 0.7:
                                sentiment = "🐂 Bullish (Low PCR)"
                                sentiment_color = "green"
                            else:
                                sentiment = "⚖️ Neutral"
                                sentiment_color = "orange"
                            
                            st.markdown(f'<p style="color: {sentiment_color}; font-size: 1.2em; font-weight: bold;">{sentiment}</p>', unsafe_allow_html=True)
                            
                        # Additional sheets analysis
                        st.subheader("📋 Additional Data Sheets")
                        
                        other_sheets = [sheet for sheet in data_dict.keys() if sheet not in option_sheets]
                        selected_additional = st.selectbox("Select Additional Sheet", ["None"] + other_sheets)
                        
                        if selected_additional != "None" and selected_additional in data_dict:
                            additional_df = data_dict[selected_additional]
                            st.write(f"**{selected_additional} Data:**")
                            st.dataframe(additional_df.head(20), use_container_width=True)
            
            else:
                st.warning("No options chain sheets found in the uploaded file.")
                st.info("Available sheets: " + ", ".join(data_dict.keys()))
        
        else:
            st.error("Could not load data from the Excel file. Please check the file format.")
    
    else:
        st.info("""
        🚀 **Welcome to the Live NSE Options Chain Dashboard!**
        
        **Features:**
        - 📊 Real-time options chain analysis
        - 📈 Open Interest and Volume charts
        - 💹 PCR (Put-Call Ratio) monitoring
        - 🔢 Greeks analysis (Delta, Theta, Vega, Gamma)
        - 📉 Implied Volatility smile
        - ⚖️ Max Pain calculation
        - 🎯 Support/Resistance identification
        
        **How to use:**
        1. Upload your Live_Option_Chain_Terminal.xlsm file
        2. Select the options chain to analyze
        3. Enable auto-refresh for live updates
        4. Explore different tabs for comprehensive analysis
        
        **Auto-refresh keeps your data live!** ⚡
        ""
