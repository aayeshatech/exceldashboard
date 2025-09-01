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
.dataframe-container {
    max-height: 600px;
    overflow-y: auto;
    border: 1px solid #ddd;
    border-radius: 8px;
    padding: 10px;
    background-color: #f9f9f9;
}
.filter-container {
    background-color: #f8f9fa;
    padding: 1rem;
    border-radius: 8px;
    margin-bottom: 1rem;
    border: 1px solid #dee2e6;
}
.sheet-info {
    background: linear-gradient(135deg, #17a2b8, #138496);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin-bottom: 1rem;
}
.column-selector {
    background-color: #e9ecef;
    padding: 1rem;
    border-radius: 8px;
    margin: 1rem 0;
    border: 1px solid #ced4da;
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

def get_sheet_column_config(sheet_name, df):
    """Get smart column configuration based on sheet name and content"""
    sheet_upper = sheet_name.upper()
    config = {
        'default_columns': [],
        'important_columns': [],
        'filter_columns': [],
        'display_name': sheet_name
    }
    
    # Define configurations for different sheet types
    if 'SECTOR' in sheet_upper and 'DASHBOARD' in sheet_upper:
        config.update({
            'display_name': '🏭 Sector Dashboard',
            'default_columns': [23, 25] if len(df.columns) > 25 else [0, 1],
            'important_columns': [23, 25],
            'filter_columns': [23, 25],
            'description': 'Sector performance analysis with bullish/bearish percentages'
        })
    
    elif 'NIFTY' in sheet_upper and 'BULLISH' in sheet_upper and 'STOCK' in sheet_upper:
        config.update({
            'display_name': '📈 Nifty 50 Bullish Stocks',
            'default_columns': [0, 1, 2, 3, 4, 5, 6] if len(df.columns) > 6 else list(range(min(7, len(df.columns)))),
            'important_columns': [0, 1, 2, 5, 6],
            'filter_columns': [5, 6],
            'description': 'Bullish stock analysis with price changes and build-up patterns'
        })
    
    elif 'OPTIONS' in sheet_upper or 'OPTION' in sheet_upper:
        config.update({
            'display_name': '⚡ Options Data',
            'default_columns': [0, 1, 2, 3, 4] if len(df.columns) > 4 else list(range(min(5, len(df.columns)))),
            'important_columns': [0, 1, 2, 3, 4],
            'filter_columns': [1, 2],
            'description': 'Options chain analysis and trading data'
        })
    
    elif 'FUTURES' in sheet_upper or 'FUTURE' in sheet_upper:
        config.update({
            'display_name': '🚀 Futures Data',
            'default_columns': [0, 1, 2, 3, 4] if len(df.columns) > 4 else list(range(min(5, len(df.columns)))),
            'important_columns': [0, 1, 2, 3, 4],
            'filter_columns': [1, 2],
            'description': 'Futures trading data and open interest analysis'
        })
    
    else:
        # Generic configuration
        num_cols = len(df.columns)
        config.update({
            'display_name': f'📋 {sheet_name}',
            'default_columns': list(range(min(5, num_cols))),
            'important_columns': list(range(min(8, num_cols))),
            'filter_columns': [],
            'description': f'Data sheet with {num_cols} columns and {len(df)} rows'
        })
    
    # Ensure column indices are valid
    config['default_columns'] = [i for i in config['default_columns'] if i < len(df.columns)]
    config['important_columns'] = [i for i in config['important_columns'] if i < len(df.columns)]
    config['filter_columns'] = [i for i in config['filter_columns'] if i < len(df.columns)]
    
    return config

def extract_sector_data(data_dict):
    """Extract sector performance data specifically from columns X and Z in Sector Dashboard sheet"""
    sectors = {}
    
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
                
        except Exception as e:
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
        return categories
    
    df = data_dict[target_sheet]
    
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
    cols_per_row = 4
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

def display_sheet_data(data_dict, selected_sheet):
    """Display the selected sheet data with smart filtering options"""
    if not selected_sheet or selected_sheet not in data_dict:
        return
    
    df = data_dict[selected_sheet]
    config = get_sheet_column_config(selected_sheet, df)
    
    # Sheet header with info
    st.markdown(f"""
    <div class="sheet-info">
        <h2>{config['display_name']}</h2>
        <p>{config['description']}</p>
        <p><strong>Dimensions:</strong> {len(df)} rows × {len(df.columns)} columns</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Create filtering section
    if config['filter_columns']:
        st.markdown('<div class="filter-container">', unsafe_allow_html=True)
        st.subheader("🔍 Filter Data")
        
        filter_cols = st.columns(min(len(config['filter_columns']), 3))
        filters_applied = {}
        
        for i, col_idx in enumerate(config['filter_columns']):
            if i < len(filter_cols):
                with filter_cols[i]:
                    col_name = df.columns[col_idx]
                    st.write(f"**{col_name}**")
                    
                    # Get unique values, handling different data types
                    try:
                        unique_values = df.iloc[:, col_idx].dropna().unique()
                        unique_values = [str(val) for val in unique_values if str(val) != 'nan']
                        unique_values = sorted(unique_values)
                        
                        if unique_values:
                            selected_value = st.selectbox(
                                f"Filter by {col_name}",
                                options=["All"] + unique_values,
                                key=f"filter_{col_idx}"
                            )
                            
                            if selected_value != "All":
                                filters_applied[col_idx] = selected_value
                    except Exception as e:
                        st.write(f"Cannot filter by this column: {str(e)}")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Apply filters
        filtered_df = df.copy()
        for col_idx, filter_value in filters_applied.items():
            filtered_df = filtered_df[filtered_df.iloc[:, col_idx].astype(str) == filter_value]
        
        if filters_applied:
            st.info(f"Filtered to {len(filtered_df)} rows (from {len(df)} total)")
    else:
        filtered_df = df.copy()
    
    # Column selection section
    st.markdown('<div class="column-selector">', unsafe_allow_html=True)
    st.subheader("📊 Column Selection")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Create column options with better formatting
        col_options = []
        for i, col_name in enumerate(df.columns):
            # Truncate long column names for display
            display_name = col_name if len(str(col_name)) < 30 else f"{str(col_name)[:27]}..."
            col_options.append(f"Col {i:2d}: {display_name}")
        
        # Pre-select columns based on sheet type
        default_selection = [col_options[i] for i in config['default_columns'] if i < len(col_options)]
        
        selected_cols = st.multiselect(
            "Select columns to display:",
            options=col_options,
            default=default_selection,
            help="Choose which columns to show in the table below"
        )
    
    with col2:
        st.write("**Quick Select:**")
        
        # Quick selection buttons
        if st.button("📌 Important Columns", help="Select commonly used columns for this sheet type"):
            important_cols = [col_options[i] for i in config['important_columns'] if i < len(col_options)]
            st.rerun()
        
        if st.button("📋 All Columns", help="Select all available columns"):
            st.rerun()
        
        if st.button("🔄 Reset to Default", help="Reset to recommended columns"):
            st.rerun()
        
        # Show column count
        st.metric("Selected", len(selected_cols))
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Extract column indices from selected options
    col_indices = []
    for col_opt in selected_cols:
        try:
            # Extract index from "Col XX: Column Name" format
            idx = int(col_opt.split(":")[0].replace("Col", "").strip())
            col_indices.append(idx)
        except:
            pass
    
    # Apply column selection
    if col_indices:
        display_df = filtered_df.iloc[:, col_indices]
    else:
        display_df = filtered_df
        st.warning("No columns selected. Showing all columns.")
    
    # Display the data table
    st.subheader("📈 Data Table")
    
    # Add download option
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        st.write(f"**Showing:** {len(display_df)} rows × {len(display_df.columns)} columns")
    with col2:
        # Download button for filtered data
        csv = display_df.to_csv(index=False)
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name=f"{selected_sheet}_filtered.csv",
            mime="text/csv"
        )
    
    # Display dataframe with improved styling
    st.markdown('<div class="dataframe-container">', unsafe_allow_html=True)
    
    # Use st.dataframe with better configuration
    st.dataframe(
        display_df,
        use_container_width=True,
        height=400
    )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Summary statistics for numeric columns
    numeric_cols = display_df.select_dtypes(include=[np.number]).columns
    if len(numeric_cols) > 0:
        with st.expander("📊 Summary Statistics"):
            st.dataframe(display_df[numeric_cols].describe())

def display_dashboard(data_dict, selected_sheet=None):
    """Display the main dashboard"""
    
    # Header
    st.markdown(f"""
    <div class="dashboard-header">
        <h1>📊 F&O Trading Dashboard</h1>
        <p class="live-indicator">● LIVE DATA</p>
        <p>Real-time Analysis - {datetime.now().strftime("%d %B %Y, %H:%M:%S")}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # If a specific sheet is selected, display it with filtering options
    if selected_sheet and selected_sheet in data_dict:
        display_sheet_data(data_dict, selected_sheet)
        
        # Add a separator
        st.markdown("---")
    
    # Extract and display sector data
    sector_data = extract_sector_data(data_dict)
    
    if sector_data:
        st.header("🏭 Sector Performance")
        
        # Display sectors in a responsive grid
        sector_items = list(sector_data.items())
        cols_per_row = min(4, len(sector_items))
        
        for i in range(0, len(sector_items), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, (sector, data) in enumerate(sector_items[i:i+cols_per_row]):
                with cols[j]:
                    sector_class = "bullish-sector" if data['bullish'] > 60 else "bearish-sector" if data['bullish'] < 40 else ""
                    st.markdown(f"""
                    <div class="sector-performance {sector_class}">
                        <h4>{sector}</h4>
                        <p>📈 Bullish: {data['bullish']:.1f}%</p>
                        <p>📉 Bearish: {data['bearish']:.1f}%</p>
                    </div>
                    """, unsafe_allow_html=True)
    
    # Extract and display stock data
    stock_categories = extract_stock_data(data_dict)
    
    # Display summary metrics
    st.header("📈 Market Summary")
    cols = st.columns(6)
    
    metrics = [
        ("Long Buildup", len(stock_categories['long_buildup']), "🟢"),
        ("Short Covering", len(stock_categories['short_covering']), "🔵"),
        ("Short Buildup", len(stock_categories['short_buildup']), "🔴"),
        ("Long Unwinding", len(stock_categories['long_unwinding']), "🟡"),
        ("Bullish Stocks", len(stock_categories['bullish_stocks']), "📈"),
        ("Bearish Stocks", len(stock_categories['bearish_stocks']), "📉")
    ]
    
    for i, (label, count, icon) in enumerate(metrics):
        with cols[i]:
            st.markdown(f"""
            <div class="metric-card">
                <h2>{icon}</h2>
                <h3>{count}</h3>
                <p>{label}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Stock analysis tabs
    st.header("🎯 Stock Analysis")
    
    tabs = st.tabs(["🟢 Long Buildup", "🔵 Short Covering", "🔴 Short Buildup", "🟡 Long Unwinding", "📈 All Bullish", "📉 All Bearish"])
    
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
    
    info_cols = st.columns(4)
    with info_cols[0]:
        st.metric("📊 Sheets Processed", len(data_dict))
    with info_cols[1]:
        st.metric("📈 Total Stocks", total_stocks)
    with info_cols[2]:
        st.metric("🏭 Sectors Found", len(sector_data))
    with info_cols[3]:
        st.metric("⏰ Last Updated", datetime.now().strftime("%H:%M:%S"))

def main():
    st.sidebar.title("📊 F&O Dashboard Control")
    
    # File upload
    st.sidebar.markdown("### 📁 File Upload")
    st.sidebar.success("✅ Supports macro-enabled files (.xlsm)!")
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File", 
        type=["xlsx", "xls", "xlsm"]
    )
    
    # Auto-refresh option
    st.sidebar.markdown("### ⚙️ Settings")
    auto_refresh = st.sidebar.checkbox("🔄 Auto Refresh (30s)", value=False)
    
    # Manual refresh button
    if st.sidebar.button("🔄 Refresh Data", help="Manually refresh the dashboard"):
        st.rerun()
    
    # Display current time
    st.sidebar.markdown(f"**🕒 Current Time:** {datetime.now().strftime('%H:%M:%S')}")
    
    if uploaded_file:
        # Process file
        file_extension = uploaded_file.name.split('.')[-1]
        temp_file = f"temp_file.{file_extension}"
        
        with open(temp_file, "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Load data with progress indicator
        with st.spinner("🔄 Processing Excel file..."):
            data_dict = read_excel_data(temp_file)
        
        # Clean up temporary file
        try:
            os.remove(temp_file)
        except:
            pass
        
        if data_dict:
            # Sheet selector with enhanced display
            st.sidebar.markdown("### 📋 Sheet Selection")
            sheet_names = list(data_dict.keys())
            
            # Create better sheet display options
            sheet_options = ["📊 Dashboard Overview"] + [f"📄 {name}" for name in sheet_names]
            selected_option = st.sidebar.selectbox(
                "Choose a view:",
                sheet_options,
                index=0,
                help="Select 'Dashboard Overview' for the main dashboard, or choose a specific sheet to view its data"
            )
            
            # Extract the actual sheet name
            if selected_option == "📊 Dashboard Overview":
                selected_sheet = None
            else:
                selected_sheet = selected_option.replace("📄 ", "")
            
            # Display sheet information in sidebar
            if selected_sheet:
                df = data_dict[selected_sheet]
                config = get_sheet_column_config(selected_sheet, df)
                st.sidebar.markdown("### 📈 Sheet Info")
                st.sidebar.info(f"**Rows:** {len(df)}\n**Columns:** {len(df.columns)}")
                st.sidebar.write(f"**Type:** {config['display_name']}")
            
            # Display dashboard
            display_dashboard(data_dict, selected_sheet)
            
            # Auto-refresh functionality
            if auto_refresh:
                try:
                    import time
                    time.sleep(30)
                    st.rerun()
                except ImportError:
                    st.sidebar.info("💡 Install streamlit-autorefresh for better auto-refresh")
        else:
            st.error("❌ Could not process the Excel file. Please check the file format and try again.")
    
    else:
        # Welcome screen when no file is uploaded
        st.markdown("""
        <div class="dashboard-header">
            <h1>📊 F&O Trading Dashboard</h1>
            <p>Welcome to your comprehensive F&O analysis tool</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        ## 🚀 Getting Started
        
        1. **Upload your Excel file** using the sidebar file uploader
        2. **Choose your view** - Dashboard Overview or specific sheet analysis
        3. **Filter and analyze** your data with interactive tools
        
        ## 📋 Supported Features
        
        - ✅ **Multi-sheet Excel files** (.xlsx, .xls, .xlsm)
        - ✅ **Smart column selection** based on sheet type
        - ✅ **Interactive filtering** for data analysis
        - ✅ **Sector performance** visualization
        - ✅ **Stock categorization** (Long Buildup, Short Covering, etc.)
        - ✅ **Real-time dashboard** with auto-refresh
        - ✅ **Data export** functionality
        
        ## 🎯 Sheet Types Recognized
        
        - **Sector Dashboard** - Automatically detects sector performance data
        - **Nifty 50 Bullish Stocks** - Stock analysis with build-up patterns
        - **Options Data** - Options chain analysis
        - **Futures Data** - Futures trading data
        - **Generic Sheets** - Smart column detection for any data
        """)
        
        # Sample data option
        st.sidebar.markdown("### 🎯 Demo")
        if st.sidebar.button("🎯 Load Sample Data", help="Try the dashboard with sample trading data"):
            # Create comprehensive sample data
            sample_data = {
                'Sector Dashboard': pd.DataFrame({
                    **{f'Col_{i}': [f'Data_{i}_{j}' for j in range(4)] for i in range(23)},
                    df.columns[23]: ['Banking', 'IT', 'Pharma', 'Auto'],  # Column X (23)
                    'Col_24': ['Sample_24_0', 'Sample_24_1', 'Sample_24_2', 'Sample_24_3'],
                    df.columns[25]: ['65.4%', '72.1%', '45.8%', '58.3%']  # Column Z (25)
                }),
                'Nifty 50 Bullish Stock': pd.DataFrame({
                    'Symbol': ['INFY', 'TCS', 'ASIANPAINT', 'HDFCBANK', 'RELIANCE', 'WIPRO'],
                    'Change %': [1.91, 2.12, 0.85, 1.45, -0.75, 3.21],
                    'Price': [1497.7, 3456.8, 1567.3, 1645.2, 2567.4, 445.6],
                    'OI': [12000, 8500, 9200, 15600, 22300, 7800],
                    'Volume': [150000, 98000, 110000, 185000, 245000, 92000],
                    'Buildup': ['LongBuilding', 'LongBuilding', 'Shortcover', 'LongBuilding', 'ShortBuildup', 'LongBuilding'],
                    'Sentiment': ['Bullish', 'Bullish', 'ShortCover', 'Bullish', 'Bearish', 'Bullish']
                }),
                'Options Data': pd.DataFrame({
                    'Strike': [15000, 15100, 15200, 15300, 15400],
                    'Call OI': [45000, 38000, 52000, 29000, 18000],
                    'Put OI': [18000, 25000, 41000, 56000, 72000],
                    'Call Volume': [12000, 8500, 15600, 7200, 4100],
                    'Put Volume': [5600, 9200, 18900, 25600, 31200]
                }),
                'Futures Data': pd.DataFrame({
                    'Symbol': ['NIFTY', 'BANKNIFTY', 'FINNIFTY'],
                    'Price': [19845.6, 43256.8, 18967.4],
                    'Change': [125.4, -245.2, 89.7],
                    'OI': [2500000, 1800000, 950000],
                    'Volume': [5600000, 3200000, 1800000]
                })
            }
            
            # Fix the column reference issue
            sample_sector_df = pd.DataFrame({
                **{f'Col_{i}': [f'Data_{i}_{j}' for j in range(4)] for i in range(26)}
            })
            sample_sector_df.iloc[:, 23] = ['Banking', 'IT', 'Pharma', 'Auto']  # Column X
            sample_sector_df.iloc[:, 25] = ['65.4%', '72.1%', '45.8%', '58.3%']  # Column Z
            
            sample_data['Sector Dashboard'] = sample_sector_df
            
            st.success("🎉 Sample data loaded! Explore the dashboard features.")
            display_dashboard(sample_data)

if __name__ == "__main__":
    main()
