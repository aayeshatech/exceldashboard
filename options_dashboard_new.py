import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
import os

st.set_page_config(page_title="F&O Trading Dashboard", page_icon="📊", layout="wide")

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
    """Extract sector performance data specifically from columns X and Z in Sector Dashboard sheet"""
    sectors = {}
    
    # Debug: Print all sheet names
    st.sidebar.write("Available sheets:", list(data_dict.keys()))
    
    # Look for a sheet that contains both 'SECTOR' and 'DASHBOARD' (case-insensitive)
    target_sheet = None
    for sheet_name in data_dict.keys():
        if 'SECTOR' in sheet_name.upper() and 'DASHBOARD' in sheet_name.upper():
            target_sheet = sheet_name
            break
    
    if target_sheet is None:
        st.sidebar.error("Sector Dashboard sheet not found")
        # Try to find any sheet that might contain sector data
        for sheet_name in data_dict.keys():
            if 'SECTOR' in sheet_name.upper():
                target_sheet = sheet_name
                st.sidebar.warning(f"Found possible sector sheet: {sheet_name}")
                break
    
    if target_sheet is None:
        st.sidebar.error("No sector-related sheet found")
        return sectors
    
    df = data_dict[target_sheet]
    st.sidebar.info(f"Processing sheet: {target_sheet} with {len(df)} rows")
    
    # Get column names and indices
    col_names = list(df.columns)
    st.sidebar.write(f"Column names: {col_names}")
    
    # Try to get columns by index (X is 23, Z is 25)
    if len(col_names) > 25:
        x_col = col_names[23]  # Column X
        z_col = col_names[25]  # Column Z
        st.sidebar.info(f"Using column X (index 23): {x_col} and column Z (index 25): {z_col}")
    else:
        st.sidebar.error(f"Sheet has only {len(col_names)} columns, need at least 26 columns")
        return sectors
    
    # Extract data from these columns
    for index, row in df.iterrows():
        try:
            # Get sector name from X column
            sector_name = str(row[x_col]).strip()
            
            # Skip empty rows
            if not sector_name or sector_name == 'nan':
                continue
            
            # Get bullish percentage from Z column
            bullish_val = None
            z_value = row[z_col]
            
            if pd.notna(z_value):
                try:
                    # Handle percentage values (e.g., "0.4%")
                    if isinstance(z_value, str) and '%' in z_value:
                        bullish_val = float(z_value.replace('%', '').strip())
                    else:
                        bullish_val = float(z_value)
                except:
                    pass
            
            if bullish_val is not None:
                sectors[sector_name] = {
                    'bullish': bullish_val, 
                    'bearish': 100 - bullish_val
                }
                st.sidebar.success(f"Added sector: {sector_name} - Bullish: {bullish_val}%")
                
        except Exception as e:
            st.sidebar.error(f"Error processing row {index}: {str(e)}")
            continue
    
    return sectors

def extract_stock_data(data_dict):
    """Extract and categorize stock data - Simplified version"""
    categories = {
        'long_buildup': [],
        'short_covering': [],
        'short_buildup': [],
        'long_unwinding': [],
        'bullish_stocks': [],
        'bearish_stocks': []
    }
    
    # Debug: Print all sheet names
    st.sidebar.write("Available sheets for stocks:", list(data_dict.keys()))
    
    # Look for a sheet that contains 'NIFTY' and 'BULLISH' and 'STOCK' (case-insensitive)
    target_sheet = None
    for sheet_name in data_dict.keys():
        if 'NIFTY' in sheet_name.upper() and 'BULLISH' in sheet_name.upper() and 'STOCK' in sheet_name.upper():
            target_sheet = sheet_name
            break
    
    if target_sheet is None:
        st.sidebar.warning("Nifty 50 Bullish Stock sheet not found")
        # Try to find any sheet that might contain stock data
        for sheet_name in data_dict.keys():
            if 'STOCK' in sheet_name.upper() or 'BULLISH' in sheet_name.upper():
                target_sheet = sheet_name
                st.sidebar.warning(f"Found possible stock sheet: {sheet_name}")
                break
    
    if target_sheet is None:
        st.sidebar.error("No stock-related sheet found")
        return categories
    
    df = data_dict[target_sheet]
    st.sidebar.info(f"Processing sheet: {target_sheet} with {len(df)} rows")
    
    # Display column names for debugging
    col_names = list(df.columns)
    st.sidebar.write(f"Column names: {col_names}")
    
    # Process rows
    for index, row in df.iterrows():
        try:
            # Extract symbol (first column)
            symbol = str(row.iloc[0]).strip()
            if not symbol or symbol == 'nan':
                continue
                
            # Clean symbol name - remove NSE= prefix
            if symbol.startswith('NSE='):
                symbol = symbol[4:]
            
            # Get change percentage (second column)
            try:
                change = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
            except:
                change = 0
            
            # Get price (third column)
            try:
                price = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
            except:
                price = 0
            
            # Get OI (fourth column)
            try:
                oi = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
            except:
                oi = 0
            
            # Get volume (fifth column)
            try:
                volume = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0
            except:
                volume = 0
            
            # Get buildup type (sixth column)
            buildup = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ''
            
            # Get sentiment (seventh column)
            sentiment = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else ''
            
            stock_info = {
                'symbol': symbol,
                'change': change,
                'price': price,
                'oi': oi,
                'volume': volume,
                'buildup': buildup,
                'sentiment': sentiment
            }
            
            # Categorize by buildup type
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
            st.sidebar.error(f"Error processing row {index}: {str(e)}")
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
    
    # Extract data
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
    
    # Use st.rerun() instead of st.experimental_rerun()
    if st.sidebar.button("🔄 Refresh"):
        st.rerun()
    
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
            # Create a sample Sector Dashboard with 4 sectors
            sample_data = {
                'Sector Dashboard': pd.DataFrame({
                    'X': ['NSE:NIFTYNXT50', 'NSE:HDFCBANK', 'NSE:RBLBANK', 'NSE:YESBANK'],
                    'Z': ['0.0%', '0.4%', '0.8%', '0.5%']
                }),
                'Nifty 50 Bullish Stock': pd.DataFrame({
                    'Stock Name': ['NSE:INFY', 'NSE:ASIANPAINT', 'NSE:HDFCBANK'],
                    'Change %': ['1.91%', '2.12%', '0.85%'],
                    'Price': [1497.7, 3456.8, 1567.3],
                    'OI': [12000, 8500, 9200],
                    'Volume': [150000, 98000, 110000],
                    'Buildup': ['LongBuilding', 'LongBuilding', 'Shortcover'],
                    'Sentiment': ['Bullish', 'Bullish', 'ShortCover']
                })
            }
            display_dashboard(sample_data)

if __name__ == "__main__":
    main()
