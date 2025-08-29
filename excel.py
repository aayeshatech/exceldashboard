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
    page_title="NSE Trading Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        color: #1f77b4;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
    }
    .stSelectbox > div > div > div {
        background-color: #f0f2f6;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=60)  # Cache for 1 minute to allow for live updates
def load_excel_data(file_path, sheet_names=None):
    """Load data from Excel file with multiple sheets"""
    try:
        if sheet_names is None:
            # Get all sheet names
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
        
        data_dict = {}
        for sheet in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            data_dict[sheet] = df
            
        return data_dict, sheet_names
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return {}, []

def calculate_technical_indicators(df, price_col='Close'):
    """Calculate basic technical indicators"""
    if price_col not in df.columns:
        return df
    
    df = df.copy()
    
    # Moving Averages
    if len(df) >= 20:
        df['MA_20'] = df[price_col].rolling(window=20).mean()
    if len(df) >= 50:
        df['MA_50'] = df[price_col].rolling(window=50).mean()
    
    # RSI calculation
    if len(df) >= 14:
        delta = df[price_col].diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
        rs = gain / loss
        df['RSI'] = 100 - (100 / (1 + rs))
    
    # Volume moving average
    if 'Volume' in df.columns and len(df) >= 20:
        df['Volume_MA'] = df['Volume'].rolling(window=20).mean()
    
    return df

def create_candlestick_chart(df, symbol):
    """Create candlestick chart with technical indicators"""
    fig = make_subplots(
        rows=3, cols=1,
        shared_xaxis=True,
        vertical_spacing=0.1,
        row_heights=[0.6, 0.2, 0.2],
        subplot_titles=[f'{symbol} - Price Chart', 'Volume', 'RSI']
    )
    
    # Candlestick chart
    if all(col in df.columns for col in ['Open', 'High', 'Low', 'Close']):
        fig.add_trace(
            go.Candlestick(
                x=df.index,
                open=df['Open'],
                high=df['High'],
                low=df['Low'],
                close=df['Close'],
                name='OHLC'
            ),
            row=1, col=1
        )
        
        # Add moving averages if available
        if 'MA_20' in df.columns:
            fig.add_trace(
                go.Scatter(x=df.index, y=df['MA_20'], name='MA 20', line=dict(color='orange')),
                row=1, col=1
            )
        if 'MA_50' in df.columns:
            fig.add_trace(
                go.Scatter(x=df.index, y=df['MA_50'], name='MA 50', line=dict(color='red')),
                row=1, col=1
            )
    
    # Volume chart
    if 'Volume' in df.columns:
        colors = ['red' if close < open else 'green' 
                 for close, open in zip(df.get('Close', []), df.get('Open', []))]
        
        fig.add_trace(
            go.Bar(x=df.index, y=df['Volume'], name='Volume', marker_color=colors),
            row=2, col=1
        )
        
        if 'Volume_MA' in df.columns:
            fig.add_trace(
                go.Scatter(x=df.index, y=df['Volume_MA'], name='Volume MA', line=dict(color='blue')),
                row=2, col=1
            )
    
    # RSI chart
    if 'RSI' in df.columns:
        fig.add_trace(
            go.Scatter(x=df.index, y=df['RSI'], name='RSI', line=dict(color='purple')),
            row=3, col=1
        )
        fig.add_hline(y=70, line_dash="dash", line_color="red", row=3, col=1)
        fig.add_hline(y=30, line_dash="dash", line_color="green", row=3, col=1)
    
    fig.update_layout(
        title=f'{symbol} Trading Chart',
        xaxis_rangeslider_visible=False,
        height=800
    )
    
    return fig

def apply_filters(df, filters):
    """Apply various filters to the dataframe"""
    filtered_df = df.copy()
    
    # Price range filter
    if filters.get('price_min') is not None and 'Close' in df.columns:
        filtered_df = filtered_df[filtered_df['Close'] >= filters['price_min']]
    if filters.get('price_max') is not None and 'Close' in df.columns:
        filtered_df = filtered_df[filtered_df['Close'] <= filters['price_max']]
    
    # Volume filter
    if filters.get('volume_min') is not None and 'Volume' in df.columns:
        filtered_df = filtered_df[filtered_df['Volume'] >= filters['volume_min']]
    
    # RSI filter
    if filters.get('rsi_min') is not None and 'RSI' in df.columns:
        filtered_df = filtered_df[filtered_df['RSI'] >= filters['rsi_min']]
    if filters.get('rsi_max') is not None and 'RSI' in df.columns:
        filtered_df = filtered_df[filtered_df['RSI'] <= filters['rsi_max']]
    
    return filtered_df

def main():
    st.markdown('<h1 class="main-header">📈 NSE Live Trading Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar for file upload and controls
    st.sidebar.header("📁 Data Source")
    
    # File uploader
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel file with NSE data",
        type=['xlsx', 'xls'],
        help="Upload your Excel file containing live NSE data"
    )
    
    if uploaded_file is not None:
        # Load data
        data_dict, sheet_names = load_excel_data(uploaded_file)
        
        if data_dict:
            # Sheet selection
            st.sidebar.header("📊 Sheet Selection")
            selected_sheets = st.sidebar.multiselect(
                "Select sheets to analyze",
                sheet_names,
                default=sheet_names[:3] if len(sheet_names) >= 3 else sheet_names
            )
            
            # Filters
            st.sidebar.header("🔍 Filters")
            
            # Auto-refresh
            auto_refresh = st.sidebar.checkbox("Auto-refresh (60 seconds)", value=False)
            if auto_refresh:
                st.rerun()
            
            # Price filters
            price_filter = st.sidebar.expander("Price Filters")
            with price_filter:
                price_min = st.number_input("Minimum Price", min_value=0.0, value=0.0)
                price_max = st.number_input("Maximum Price", min_value=0.0, value=10000.0)
            
            # Volume filter
            volume_filter = st.sidebar.expander("Volume Filters")
            with volume_filter:
                volume_min = st.number_input("Minimum Volume", min_value=0, value=0)
            
            # RSI filter
            rsi_filter = st.sidebar.expander("RSI Filters")
            with rsi_filter:
                rsi_min = st.number_input("RSI Minimum", min_value=0.0, max_value=100.0, value=0.0)
                rsi_max = st.number_input("RSI Maximum", min_value=0.0, max_value=100.0, value=100.0)
            
            filters = {
                'price_min': price_min if price_min > 0 else None,
                'price_max': price_max if price_max < 10000 else None,
                'volume_min': volume_min if volume_min > 0 else None,
                'rsi_min': rsi_min if rsi_min > 0 else None,
                'rsi_max': rsi_max if rsi_max < 100 else None,
            }
            
            # Main dashboard
            if selected_sheets:
                # Create tabs for different views
                tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "📈 Charts", "📋 Data Tables", "🔍 Screener"])
                
                with tab1:
                    st.header("Market Overview")
                    
                    # Summary metrics
                    col1, col2, col3, col4 = st.columns(4)
                    
                    total_stocks = sum(len(data_dict[sheet]) for sheet in selected_sheets)
                    with col1:
                        st.metric("Total Stocks", total_stocks)
                    
                    # Calculate aggregate metrics
                    all_data = pd.concat([data_dict[sheet] for sheet in selected_sheets], ignore_index=True)
                    
                    if 'Close' in all_data.columns:
                        avg_price = all_data['Close'].mean()
                        with col2:
                            st.metric("Average Price", f"₹{avg_price:.2f}")
                    
                    if 'Volume' in all_data.columns:
                        total_volume = all_data['Volume'].sum()
                        with col3:
                            st.metric("Total Volume", f"{total_volume:,.0f}")
                    
                    with col4:
                        st.metric("Active Sheets", len(selected_sheets))
                    
                    # Top gainers/losers (if Change% column exists)
                    if 'Change%' in all_data.columns or any('change' in col.lower() for col in all_data.columns):
                        col1, col2 = st.columns(2)
                        
                        change_col = next((col for col in all_data.columns if 'change' in col.lower()), None)
                        if change_col:
                            with col1:
                                st.subheader("🟢 Top Gainers")
                                top_gainers = all_data.nlargest(5, change_col)
                                st.dataframe(top_gainers[['Symbol', change_col] if 'Symbol' in all_data.columns else top_gainers[[change_col]]])
                            
                            with col2:
                                st.subheader("🔴 Top Losers")
                                top_losers = all_data.nsmallest(5, change_col)
                                st.dataframe(top_losers[['Symbol', change_col] if 'Symbol' in all_data.columns else top_losers[[change_col]]])
                
                with tab2:
                    st.header("Trading Charts")
                    
                    # Chart selection
                    chart_sheet = st.selectbox("Select sheet for charting", selected_sheets)
                    
                    if chart_sheet in data_dict:
                        df = data_dict[chart_sheet].copy()
                        
                        # Add technical indicators
                        df = calculate_technical_indicators(df)
                        
                        # Symbol selection if Symbol column exists
                        if 'Symbol' in df.columns:
                            symbols = df['Symbol'].unique()
                            selected_symbol = st.selectbox("Select Symbol", symbols)
                            symbol_data = df[df['Symbol'] == selected_symbol].copy()
                            
                            if not symbol_data.empty:
                                # Create and display chart
                                fig = create_candlestick_chart(symbol_data, selected_symbol)
                                st.plotly_chart(fig, use_container_width=True)
                        else:
                            # If no Symbol column, use the whole dataset
                            fig = create_candlestick_chart(df, chart_sheet)
                            st.plotly_chart(fig, use_container_width=True)
                
                with tab3:
                    st.header("Data Tables")
                    
                    for sheet in selected_sheets:
                        with st.expander(f"📋 {sheet} Data", expanded=True):
                            df = data_dict[sheet].copy()
                            
                            # Add technical indicators
                            df = calculate_technical_indicators(df)
                            
                            # Apply filters
                            filtered_df = apply_filters(df, filters)
                            
                            st.write(f"Showing {len(filtered_df)} of {len(df)} records")
                            st.dataframe(filtered_df, use_container_width=True)
                            
                            # Download button
                            csv = filtered_df.to_csv(index=False)
                            st.download_button(
                                label=f"Download {sheet} data as CSV",
                                data=csv,
                                file_name=f"{sheet}_filtered_data.csv",
                                mime="text/csv"
                            )
                
                with tab4:
                    st.header("Stock Screener")
                    
                    # Combine all data for screening
                    all_data = pd.concat([data_dict[sheet] for sheet in selected_sheets], ignore_index=True)
                    
                    if not all_data.empty:
                        # Add technical indicators
                        all_data = calculate_technical_indicators(all_data)
                        
                        # Apply filters
                        screened_data = apply_filters(all_data, filters)
                        
                        st.write(f"Found {len(screened_data)} stocks matching criteria")
                        
                        if not screened_data.empty:
                            # Display screened results
                            st.dataframe(screened_data, use_container_width=True)
                            
                            # Visualization of screened data
                            if 'Close' in screened_data.columns and 'Volume' in screened_data.columns:
                                fig = px.scatter(
                                    screened_data, 
                                    x='Close', 
                                    y='Volume',
                                    title="Price vs Volume - Screened Stocks",
                                    hover_data=['Symbol'] if 'Symbol' in screened_data.columns else None
                                )
                                st.plotly_chart(fig, use_container_width=True)
            
            else:
                st.warning("Please select at least one sheet to analyze.")
        
        else:
            st.error("Could not load data from the uploaded file. Please check the file format.")
    
    else:
        st.info("""
        👋 Welcome to the NSE Trading Dashboard!
        
        **How to use:**
        1. Upload your Excel file containing live NSE data
        2. Select the sheets you want to analyze
        3. Apply filters to screen stocks
        4. View charts, tables, and analytics
        
        **Expected data format:**
        - Each sheet should contain stock data
        - Common columns: Symbol, Open, High, Low, Close, Volume, Change%
        - The dashboard will adapt to your data structure
        """)
        
        # Example data structure
        st.subheader("📋 Expected Data Format")
        example_data = pd.DataFrame({
            'Symbol': ['RELIANCE', 'TCS', 'INFY'],
            'Open': [2500.0, 3200.0, 1450.0],
            'High': [2520.0, 3250.0, 1470.0],
            'Low': [2480.0, 3180.0, 1440.0],
            'Close': [2510.0, 3230.0, 1460.0],
            'Volume': [1000000, 800000, 1200000],
            'Change%': [0.4, 0.9, 0.7]
        })
        st.dataframe(example_data)

if __name__ == "__main__":
    main()
