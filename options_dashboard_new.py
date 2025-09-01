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

def extract_sector_data(data_dict):
    """Extract sector performance data specifically from columns X and Z"""
    sectors = {}
    
    for sheet_name, df in data_dict.items():
        # Look for sector dashboard sheet
        if 'SECTOR' in sheet_name.upper() or 'DASHBOARD' in sheet_name.upper():
            
            # Get column names (X and Z columns)
            col_names = list(df.columns)
            
            # Find X and Z columns (columns 23 and 25 if 0-indexed)
            x_col_idx = None
            z_col_idx = None
            
            for i, col_name in enumerate(col_names):
                if 'X' in str(col_name).upper() or i == 23:  # Column X (index 23)
                    x_col_idx = i
                elif 'Z' in str(col_name).upper() or i == 25:  # Column Z (index 25)
                    z_col_idx = i
            
            if x_col_idx is not None and z_col_idx is not None:
                st.sidebar.info(f"Found sector data in {sheet_name}: X-col={x_col_idx}, Z-col={z_col_idx}")
                
                # Extract sector data from these columns
                for index, row in df.iterrows():
                    try:
                        # Get sector name from X column
                        sector_name = str(row.iloc[x_col_idx]).strip()
                        
                        # Skip empty rows or non-sector rows
                        if (not sector_name or sector_name == 'nan' or 
                            not any(keyword in sector_name.upper() for keyword in 
                                   ['NIFTY', 'BANK', 'FIN', 'IT', 'AUTO', 'PHARMA', 'METAL', 'OIL', 'FMCG'])):
                            continue
                        
                        # Get bullish percentage from Z column
                        bullish_val = None
                        z_value = row.iloc[z_col_idx]
                        
                        if pd.notna(z_value):
                            try:
                                # Handle percentage values (e.g., "0.4%")
                                if isinstance(z_value, str) and '%' in z_value:
                                    bullish_val = float(z_value.replace('%', '').strip())
                                else:
                                    bullish_val = float(z_value)
                            except:
                                pass
                        
                        if bullish_val is not None and 0 <= bullish_val <= 100:
                            sectors[sector_name] = {
                                'bullish': bullish_val, 
                                'bearish': 100 - bullish_val
                            }
                            st.sidebar.success(f"Added sector: {sector_name} - Bullish: {bullish_val}%")
                            
                    except Exception as e:
                        continue
    
    return sectors

def extract_stock_data(data_dict):
    """Extract and categorize stock data - Fixed to match your exact Excel format"""
    categories = {
        'long_buildup': [],
        'short_covering': [],
        'short_buildup': [],
        'long_unwinding': [],
        'bullish_stocks': [],
        'bearish_stocks': []
    }
    
    for sheet_name, df in data_dict.items():
        # Look specifically for sheets like "Nifty 50 Bullish Stock"
        if any(term in sheet_name for term in ['Bullish Stock', 'Bearish Stock', 'Stock Dashboard', 'NIFTY']):
            
            # Map exact column names from your Excel
            column_mapping = {}
            
            for col in df.columns:
                col_str = str(col).strip()
                if 'STOCK NAME' in col_str:
                    column_mapping['symbol'] = col
                elif 'CHANGE' in col_str and '%' in col_str:
                    column_mapping['change'] = col
                elif 'PRICE' in col_str:
                    column_mapping['price'] = col
                elif 'OI' in col_str and 'Change' not in col_str:
                    column_mapping['oi'] = col
                elif 'Volume' in col_str and 'Fut' not in col_str:
                    column_mapping['volume'] = col
                elif 'Building' in col_str or 'Buildup' in col_str:
                    column_mapping['buildup'] = col
                elif 'SENTIMENT' in col_str:
                    column_mapping['sentiment'] = col
            
            # Process rows if we have the required columns
            if 'symbol' in column_mapping and 'change' in column_mapping:
                for index, row in df.iterrows():
                    try:
                        # Extract symbol
                        symbol = str(row[column_mapping['symbol']]) if pd.notna(row[column_mapping['symbol']]) else ''
                        if not symbol or symbol == 'nan' or symbol == '':
                            continue
                            
                        # Clean symbol name - remove NSE= prefix (your data uses = instead of :)
                        if symbol.startswith('NSE='):
                            symbol = symbol[4:]  # Remove 'NSE=' prefix
                        
                        # Get change percentage
                        try:
                            change = float(row[column_mapping['change']]) if pd.notna(row[column_mapping['change']]) else 0
                        except:
                            change = 0
                        
                        # Get other values
                        price = float(row[column_mapping['price']]) if 'price' in column_mapping and pd.notna(row[column_mapping['price']]) else 0
                        oi = float(row[column_mapping['oi']]) if 'oi' in column_mapping and pd.notna(row[column_mapping['oi']]) else 0
                        volume = float(row[column_mapping['volume']]) if 'volume' in column_mapping and pd.notna(row[column_mapping['volume']]) else 0
                        buildup = str(row[column_mapping['buildup']]).strip() if 'buildup' in column_mapping and pd.notna(row[column_mapping['buildup']]) else ''
                        sentiment = str(row[column_mapping['sentiment']]).strip() if 'sentiment' in column_mapping and pd.notna(row[column_mapping['sentiment']]) else ''
                        
                        stock_info = {
                            'symbol': symbol,
                            'change': change,
                            'price': price,
                            'oi': oi,
                            'volume': volume,
                            'buildup': buildup,
                            'sentiment': sentiment
                        }
                        
                        # Categorize by buildup type (exact match from your data)
                        if buildup == 'LongBuilding':
                            categories['long_buildup'].append(stock_info)
                        elif buildup == 'Shortcover':
                            categories['short_covering'].append(stock_info)
                        elif buildup == 'ShortBuildup':
                            categories['short_buildup'].append(stock_info)
                        elif buildup == 'LongUnwinding':
                            categories['long_unwinding'].append(stock_info)
                        
                        # Also categorize by performance
                        if change > 0.3:
                            categories['bullish_stocks'].append(stock_info)
                        elif change < -0.3:
                            categories['bearish_stocks'].append(stock_info)
                            
                    except Exception as e:
                        continue
    
    # Sort categories
    for category in categories:
        if category == 'bearish_stocks':
            categories[category] = sorted(categories[category], key=lambda x: x['change'])[:50]
        else:
            categories[category] = sorted(categories[category], key=lambda x: x['change'], reverse=True)[:50]
    
    return categories

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
    
    # Debug mode toggle
    debug_mode = st.checkbox("🔍 Debug Mode - Show Processing Details", value=False)
    
    if debug_mode:
        st.subheader("📋 Available Sheets")
        for sheet_name, df in data_dict.items():
            with st.expander(f"Sheet: {sheet_name} ({len(df)} rows)"):
                st.write("**Columns:**", list(df.columns))
                st.dataframe(df.head(3))
    
    # Extract data - Focus on sector data first
    sector_data = extract_sector_data(data_dict)
    stock_categories = extract_stock_data(data_dict)
    
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
    else:
        st.warning("No sector data found. Please check if your Excel has a 'Sector Dashboard' sheet with data in columns X and Z.")
    
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
    
    # Debug: Show extraction results
    if debug_mode:
        st.subheader("🔍 Extraction Results")
        st.write("**Sector Data:**", sector_data)
        for category, stocks in stock_categories.items():
            st.write(f"**{category}:** {len(stocks)} stocks")
            if stocks:
                sample_stock = stocks[0]
                st.write(f"Sample: {sample_stock['symbol']} - Change: {sample_stock['change']}% - Buildup: {sample_stock['buildup']}")
    
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
    
    # Sheet summary
    if debug_mode:
        with st.expander("📄 All Sheet Names"):
            for sheet_name in data_dict.keys():
                st.write(f"• {sheet_name}")

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
                'Sector Dashboard': pd.DataFrame({
                    'Unnamed: 23': ['NSE:NIFTYNXT50', 'NSE:HDFCBANK', 'NSE:RBLBANK', 'NSE:YESBANK'],
                    'Unnamed: 25': ['0.0%', '0.4%', '0.8%', '0.5%']
                })
            }
            display_dashboard(sample_data)

if __name__ == "__main__":
    main()
