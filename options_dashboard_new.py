import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
import os

st.set_page_config(page_title="Complete F&O Trading Dashboard", page_icon="📊", layout="wide")

# Enhanced CSS for comprehensive display
st.markdown("""
<style>
.dashboard-header {
    background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
    color: white;
    padding: 1.5rem;
    border-radius: 10px;
    text-align: center;
    margin-bottom: 2rem;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}

.sector-performance {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem;
    text-align: center;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.bullish-sector {
    background: linear-gradient(135deg, #28a745, #20c997) !important;
}

.bearish-sector {
    background: linear-gradient(135deg, #dc3545, #fd7e14) !important;
}

.stock-card {
    background: white;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.5rem;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    border-left: 5px solid;
}

.long-buildup-card {
    border-left-color: #28a745;
    background: linear-gradient(135deg, rgba(40, 167, 69, 0.1), rgba(255, 255, 255, 1));
}

.short-covering-card {
    border-left-color: #17a2b8;
    background: linear-gradient(135deg, rgba(23, 162, 184, 0.1), rgba(255, 255, 255, 1));
}

.short-buildup-card {
    border-left-color: #dc3545;
    background: linear-gradient(135deg, rgba(220, 53, 69, 0.1), rgba(255, 255, 255, 1));
}

.long-unwinding-card {
    border-left-color: #ffc107;
    background: linear-gradient(135deg, rgba(255, 193, 7, 0.1), rgba(255, 255, 255, 1));
}

.metric-card {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 1.5rem;
    border-radius: 10px;
    text-align: center;
    margin: 0.5rem;
}

.live-indicator {
    animation: pulse 2s infinite;
    color: #28a745;
    font-weight: bold;
}

@keyframes pulse {
    0% { opacity: 1; }
    50% { opacity: 0.5; }
    100% { opacity: 1; }
}
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=30)
def read_excel_data(file_path):
    """Read Excel file with macro support"""
    try:
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        data_dict = {}
        
        progress_bar = st.sidebar.progress(0)
        status_text = st.sidebar.empty()
        
        for i, sheet_name in enumerate(excel_file.sheet_names):
            try:
                progress = (i + 1) / len(excel_file.sheet_names)
                progress_bar.progress(progress)
                status_text.text(f"Reading sheet: {sheet_name}")
                
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                if not df.empty:
                    data_dict[sheet_name] = df
                    
            except Exception as e:
                continue
        
        progress_bar.empty()
        status_text.empty()
        st.sidebar.success(f"✅ Loaded {len(data_dict)} sheets successfully")
        return data_dict
        
    except Exception as e:
        st.sidebar.error(f"Error reading file: {str(e)}")
        return {}

def extract_stock_data(data_dict):
    """Extract and categorize stock data"""
    categories = {
        'long_buildup': [],
        'short_covering': [],
        'short_buildup': [],
        'long_unwinding': [],
        'bullish_stocks': [],
        'bearish_stocks': []
    }
    
    for sheet_name, df in data_dict.items():
        # Look for stock data sheets
        if any(term in sheet_name.upper() for term in ['STOCK', 'BULLISH', 'BEARISH', 'NIFTY']):
            
            # Find columns by checking actual column names
            symbol_col = None
            change_col = None
            price_col = None
            oi_col = None
            volume_col = None
            buildup_col = None
            sentiment_col = None
            
            for col in df.columns:
                col_upper = str(col).upper()
                if 'STOCK' in col_upper and 'NAME' in col_upper:
                    symbol_col = col
                elif 'CHANGE' in col_upper and '%' in col_upper:
                    change_col = col
                elif col_upper == 'PRICE':
                    price_col = col
                elif col_upper == 'OI':
                    oi_col = col
                elif col_upper == 'VOLUME':
                    volume_col = col
                elif 'BUILDUP' in col_upper:
                    buildup_col = col
                elif 'SENTIMENT' in col_upper:
                    sentiment_col = col
            
            # Process rows if we found key columns
            if symbol_col and change_col:
                for _, row in df.iterrows():
                    try:
                        symbol = str(row[symbol_col]) if pd.notna(row[symbol_col]) else ''
                        if not symbol or symbol == 'nan':
                            continue
                            
                        # Clean symbol name
                        if 'NSE:' in symbol:
                            symbol = symbol.replace('NSE:', '')
                        
                        # Get values
                        change = float(row[change_col]) if pd.notna(row[change_col]) else 0
                        price = float(row[price_col]) if price_col and pd.notna(row[price_col]) else 0
                        oi = float(row[oi_col]) if oi_col and pd.notna(row[oi_col]) else 0
                        volume = float(row[volume_col]) if volume_col and pd.notna(row[volume_col]) else 0
                        buildup = str(row[buildup_col]) if buildup_col and pd.notna(row[buildup_col]) else ''
                        sentiment = str(row[sentiment_col]) if sentiment_col and pd.notna(row[sentiment_col]) else ''
                        
                        stock_info = {
                            'symbol': symbol,
                            'change': change,
                            'price': price,
                            'oi': oi,
                            'volume': volume,
                            'buildup': buildup,
                            'sentiment': sentiment
                        }
                        
                        # Categorize by buildup
                        buildup_lower = buildup.lower()
                        if 'longbuildup' in buildup_lower:
                            categories['long_buildup'].append(stock_info)
                        elif 'shortcover' in buildup_lower:
                            categories['short_covering'].append(stock_info)
                        elif 'shortbuildup' in buildup_lower:
                            categories['short_buildup'].append(stock_info)
                        elif 'longunwind' in buildup_lower:
                            categories['long_unwinding'].append(stock_info)
                        
                        # Categorize by performance
                        if change > 0.5:
                            categories['bullish_stocks'].append(stock_info)
                        elif change < -0.5:
                            categories['bearish_stocks'].append(stock_info)
                            
                    except:
                        continue
    
    # Sort categories
    for category in categories:
        if category == 'bearish_stocks':
            categories[category] = sorted(categories[category], key=lambda x: x['change'])[:30]
        else:
            categories[category] = sorted(categories[category], key=lambda x: x['change'], reverse=True)[:30]
    
    return categories

def extract_sector_data(data_dict):
    """Extract sector performance data"""
    sectors = {}
    
    for sheet_name, df in data_dict.items():
        if 'SECTOR' in sheet_name.upper() or 'Dashboard' in sheet_name:
            for _, row in df.iterrows():
                try:
                    first_val = str(row.iloc[0])
                    if any(keyword in first_val.upper() for keyword in ['NIFTY', 'BANK', 'IT', 'AUTO', 'PHARMA']):
                        # Look for percentage values in the row
                        for val in row:
                            try:
                                if isinstance(val, (int, float)) and 0 <= val <= 100:
                                    sectors[first_val] = {'bullish': val, 'bearish': 100-val}
                                    break
                            except:
                                continue
                except:
                    continue
    
    return sectors

def display_stock_cards(stocks, title, card_class):
    """Display stocks in card format"""
    if not stocks:
        st.info(f"No {title.lower()} found")
        return
    
    st.subheader(title)
    
    # Display in grid format
    cols_per_row = 3
    for i in range(0, len(stocks), cols_per_row):
        cols = st.columns(cols_per_row)
        for j, stock in enumerate(stocks[i:i+cols_per_row]):
            with cols[j]:
                st.markdown(f"""
                <div class="stock-card {card_class}">
                    <h4>{stock['symbol']}</h4>
                    <p><strong>Change:</strong> {stock['change']:+.2f}%</p>
                    <p><strong>Price:</strong> ₹{stock['price']:.2f}</p>
                    <p><strong>OI:</strong> {stock['oi']:,.0f}</p>
                    <p><strong>Volume:</strong> {stock['volume']:,.0f}</p>
                    <p><strong>Buildup:</strong> {stock['buildup']}</p>
                    <p><strong>Sentiment:</strong> {stock['sentiment']}</p>
                </div>
                """, unsafe_allow_html=True)

def display_dashboard(data_dict):
    """Display the main dashboard"""
    
    # Header
    st.markdown(f"""
    <div class="dashboard-header">
        <h1>📊 F&O Trading Dashboard</h1>
        <p class="live-indicator">● LIVE DATA</p>
        <p>Real-time Analysis - {datetime.now().strftime("%d %B %Y, %H:%M:%S")}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Extract data
    stock_categories = extract_stock_data(data_dict)
    sector_data = extract_sector_data(data_dict)
    
    # Display sector performance
    if sector_data:
        st.header("📊 Sector Performance")
        cols = st.columns(min(4, len(sector_data)))
        for i, (sector, data) in enumerate(sector_data.items()):
            if i < len(cols):
                with cols[i]:
                    sector_class = "bullish-sector" if data['bullish'] > 60 else "bearish-sector" if data['bullish'] < 40 else ""
                    st.markdown(f"""
                    <div class="sector-performance {sector_class}">
                        <h4>{sector}</h4>
                        <p>Bullish: {data['bullish']:.1f}%</p>
                        <p>Bearish: {data['bearish']:.1f}%</p>
                    </div>
                    """, unsafe_allow_html=True)
    
    # Display summary metrics
    st.header("📈 Market Summary")
    cols = st.columns(6)
    
    metrics = [
        ("Long Buildup", len(stock_categories['long_buildup'])),
        ("Short Covering", len(stock_categories['short_covering'])),
        ("Short Buildup", len(stock_categories['short_buildup'])),
        ("Long Unwinding", len(stock_categories['long_unwinding'])),
        ("Bullish Stocks", len(stock_categories['bullish_stocks'])),
        ("Bearish Stocks", len(stock_categories['bearish_stocks']))
    ]
    
    for i, (label, count) in enumerate(metrics):
        with cols[i]:
            st.markdown(f"""
            <div class="metric-card">
                <h3>{count}</h3>
                <p>{label}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Stock analysis tabs
    st.header("🎯 Stock Analysis")
    
    tabs = st.tabs(["Long Buildup", "Short Covering", "Short Buildup", "Long Unwinding", "All Bullish", "All Bearish"])
    
    with tabs[0]:
        display_stock_cards(stock_categories['long_buildup'], "Long Buildup Stocks", "long-buildup-card")
    
    with tabs[1]:
        display_stock_cards(stock_categories['short_covering'], "Short Covering Stocks", "short-covering-card")
    
    with tabs[2]:
        display_stock_cards(stock_categories['short_buildup'], "Short Buildup Stocks", "short-buildup-card")
    
    with tabs[3]:
        display_stock_cards(stock_categories['long_unwinding'], "Long Unwinding Stocks", "long-unwinding-card")
    
    with tabs[4]:
        display_stock_cards(stock_categories['bullish_stocks'], "All Bullish Stocks", "long-buildup-card")
    
    with tabs[5]:
        display_stock_cards(stock_categories['bearish_stocks'], "All Bearish Stocks", "short-buildup-card")
    
    # Data info
    st.markdown("---")
    total_stocks = sum(len(stocks) for stocks in stock_categories.values())
    st.markdown(f"""
    **Data Source:** {len(data_dict)} Excel sheets processed  
    **Last Updated:** {datetime.now().strftime("%H:%M:%S")}  
    **Total Stocks Analyzed:** {total_stocks}
    """)

def main():
    st.sidebar.title("📊 F&O Dashboard")
    
    # File upload
    st.sidebar.success("✅ Supports macro-enabled files (.xlsm)!")
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File", 
        type=["xlsx", "xls", "xlsm"]
    )
    
    # Auto-refresh option
    auto_refresh = st.sidebar.checkbox("Auto Refresh (30s)", value=False)
    
    if st.sidebar.button("🔄 Refresh"):
        st.experimental_rerun()
    
    # Display time
    st.sidebar.markdown(f"**Time:** {datetime.now().strftime('%H:%M:%S')}")
    
    if uploaded_file:
        # Process file
        file_extension = uploaded_file.name.split('.')[-1]
        temp_file = f"temp_file.{file_extension}"
        
        with open(temp_file, "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Load data
        with st.spinner("Processing Excel file..."):
            data_dict = read_excel_data(temp_file)
        
        # Clean up
        try:
            os.remove(temp_file)
        except:
            pass
        
        if data_dict:
            display_dashboard(data_dict)
            
            # Auto-refresh
            if auto_refresh:
                try:
                    from streamlit_autorefresh import st_autorefresh
                    st_autorefresh(interval=30000, key="dashboard_refresh")
                except ImportError:
                    st.sidebar.info("Install streamlit-autorefresh for auto-refresh")
        else:
            st.error("Could not process the Excel file")
    
    else:
        st.info("Please upload an Excel file to view the F&O dashboard")
        
        # Sample data option
        if st.sidebar.checkbox("🎯 Load Sample Data"):
            sample_data = {
                'Sample Sheet': pd.DataFrame({
                    'STOCK NAME': ['NSE:RELIANCE', 'NSE:TCS', 'NSE:INFY', 'NSE:HDFC'],
                    'CHANGE %': [2.45, 1.87, -1.25, 0.95],
                    'PRICE': [2850, 3650, 1750, 1650],
                    'OI': [145000, 125000, 180000, 95000],
                    'VOLUME': [85000, 65000, 95000, 75000],
                    'Buildup': ['longBuildup', 'longBuildup', 'shortBuildup', 'shortCover'],
                    'SENTIMENT': ['Bullish', 'Bullish', 'Wait For Signal', 'Bullish']
                })
            }
            display_dashboard(sample_data)

if __name__ == "__main__":
    main()
