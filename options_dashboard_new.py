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

.stock-table {
    background: white;
    border-radius: 10px;
    padding: 1rem;
    margin: 1rem 0;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
}

.long-buildup {
    background: linear-gradient(135deg, #28a745, #20c997);
    color: white;
    padding: 0.5rem;
    border-radius: 5px;
    margin: 0.2rem 0;
    font-size: 0.9rem;
}

.short-covering {
    background: linear-gradient(135deg, #17a2b8, #6f42c1);
    color: white;
    padding: 0.5rem;
    border-radius: 5px;
    margin: 0.2rem 0;
    font-size: 0.9rem;
}

.short-buildup {
    background: linear-gradient(135deg, #dc3545, #fd7e14);
    color: white;
    padding: 0.5rem;
    border-radius: 5px;
    margin: 0.2rem 0;
    font-size: 0.9rem;
}

.long-unwinding {
    background: linear-gradient(135deg, #ffc107, #fd7e14);
    color: white;
    padding: 0.5rem;
    border-radius: 5px;
    margin: 0.2rem 0;
    font-size: 0.9rem;
}

.bullish-stock {
    background: linear-gradient(135deg, #28a745, #20c997);
    color: white;
    padding: 0.8rem;
    border-radius: 8px;
    margin: 0.3rem 0;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.bearish-stock {
    background: linear-gradient(135deg, #dc3545, #fd7e14);
    color: white;
    padding: 0.8rem;
    border-radius: 8px;
    margin: 0.3rem 0;
    display: flex;
    justify-content: space-between;
    align-items: center;
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

.data-table {
    background: white;
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.metric-card {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 1.5rem;
    border-radius: 10px;
    text-align: center;
    margin: 0.5rem;
}

.wait-signal {
    background: linear-gradient(135deg, #6c757d, #495057);
    color: white;
    padding: 0.3rem 0.6rem;
    border-radius: 4px;
    font-size: 0.8rem;
}

.bullish-signal {
    background: linear-gradient(135deg, #28a745, #20c997);
    color: white;
    padding: 0.3rem 0.6rem;
    border-radius: 4px;
    font-size: 0.8rem;
}
</style>
""", unsafe_allow_html=True)

def find_column_by_keywords(columns, keywords):
    """Find a column that contains any of the given keywords"""
    for col in columns:
        col_upper = str(col).upper()
        for keyword in keywords:
            if keyword in col_upper:
                return col
    return None

def create_sample_data():
    """Create sample F&O data for demonstration"""
    sample_stocks = [
        # Long Buildup stocks
        {'symbol': 'RELIANCE', 'change': 2.45, 'price': 2850, 'oi': 145000, 'oi_change': 15.2, 'volume': 85000, 'buildup': 'longBuildup', 'sentiment': 'Bullish'},
        {'symbol': 'TCS', 'change': 1.87, 'price': 3650, 'oi': 125000, 'oi_change': 12.8, 'volume': 65000, 'buildup': 'longBuildup', 'sentiment': 'Bullish'},
        {'symbol': 'INFY', 'change': 1.65, 'price': 1750, 'oi': 180000, 'oi_change': 18.5, 'volume': 95000, 'buildup': 'longBuildup', 'sentiment': 'Bullish'},
        
        # Short Covering stocks
        {'symbol': 'HDFC', 'change': 1.25, 'price': 1650, 'oi': 95000, 'oi_change': -8.5, 'volume': 75000, 'buildup': 'shortCover', 'sentiment': 'Bullish'},
        {'symbol': 'ICICIBANK', 'change': 0.95, 'price': 1150, 'oi': 125000, 'oi_change': -12.3, 'volume': 85000, 'buildup': 'shortCover', 'sentiment': 'Wait For Signal'},
        
        # Short Buildup stocks
        {'symbol': 'BAJFINANCE', 'change': -2.15, 'price': 6800, 'oi': 85000, 'oi_change': 22.5, 'volume': 45000, 'buildup': 'shortBuildup', 'sentiment': 'Wait For Signal'},
        {'symbol': 'HDFCBANK', 'change': -1.85, 'price': 1580, 'oi': 195000, 'oi_change': 15.8, 'volume': 125000, 'buildup': 'shortBuildup', 'sentiment': 'Wait For Signal'},
        
        # Long Unwinding stocks
        {'symbol': 'WIPRO', 'change': -1.45, 'price': 450, 'oi': 65000, 'oi_change': -18.5, 'volume': 55000, 'buildup': 'longUnwind', 'sentiment': 'Wait For Signal'},
        {'symbol': 'TECHM', 'change': -0.85, 'price': 1250, 'oi': 75000, 'oi_change': -15.2, 'volume': 65000, 'buildup': 'longUnwind', 'sentiment': 'Wait For Signal'},
        
        # Additional stocks
        {'symbol': 'ASIANPAINT', 'change': 1.97, 'price': 3200, 'oi': 55000, 'oi_change': 8.5, 'volume': 35000, 'buildup': 'longBuildup', 'sentiment': 'Bullish'},
        {'symbol': 'MARUTI', 'change': 1.21, 'price': 11500, 'oi': 45000, 'oi_change': 5.2, 'volume': 25000, 'buildup': 'shortCover', 'sentiment': 'Bullish'},
        {'symbol': 'LT', 'change': -1.15, 'price': 3450, 'oi': 85000, 'oi_change': 12.5, 'volume': 65000, 'buildup': 'shortBuildup', 'sentiment': 'Wait For Signal'},
    ]
    
    # Convert to DataFrame format that matches expected structure
    df = pd.DataFrame(sample_stocks)
    
    # Create a data_dict structure
    data_dict = {
        'Sample_Stock_Data': df,
        'Dashboard': pd.DataFrame({
            'Sector': ['NIFTY 50', 'NIFTY BANK', 'IT', 'PHARMA', 'AUTO', 'METAL'],
            'BULLISH': [65.5, 45.2, 78.8, 55.5, 60.0, 35.5],
            'BEARISH': [34.5, 54.8, 21.2, 44.5, 40.0, 64.5]
        })
    }
    
    return data_dict

def convert_csv_to_data_dict(df):
    """Convert CSV DataFrame to expected data_dict format"""
    data_dict = {
        'CSV_Stock_Data': df
    }
    return data_dict
    """Find a column that contains any of the given keywords"""
    for col in columns:
        col_upper = str(col).upper()
        for keyword in keywords:
            if keyword in col_upper:
                return col
    return None

@st.cache_data(ttl=30)  # Cache for 30 seconds for auto-refresh
def read_comprehensive_data(file_path):
    """Read all relevant sheets for F&O analysis - supports .xlsm files"""
    try:
        # Try reading with openpyxl engine which supports .xlsm files
        try:
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        except:
            # Fallback to default engine
            excel_file = pd.ExcelFile(file_path)
            
        data_dict = {}
        
        st.sidebar.info(f"📊 Processing {len(excel_file.sheet_names)} sheets...")
        
        # Read all available sheets
        for i, sheet_name in enumerate(excel_file.sheet_names):
            try:
                # Show progress
                progress = (i + 1) / len(excel_file.sheet_names)
                st.sidebar.progress(progress, text=f"Reading sheet: {sheet_name}")
                
                # Read sheet with error handling for macro-enabled files
                df = pd.read_excel(
                    file_path, 
                    sheet_name=sheet_name,
                    engine='openpyxl'  # Explicitly use openpyxl for .xlsm support
                )
                
                if not df.empty:
                    data_dict[sheet_name] = df
                    st.sidebar.success(f"✅ {sheet_name}: {len(df)} rows")
                else:
                    st.sidebar.warning(f"⚠️ {sheet_name}: Empty sheet")
                    
            except Exception as e:
                st.sidebar.error(f"❌ Error reading {sheet_name}: {str(e)}")
                continue
        
        st.sidebar.success(f"🎉 Successfully loaded {len(data_dict)} sheets")
        return data_dict
        
    except Exception as e:
        st.sidebar.error(f"❌ Error reading Excel file: {str(e)}")
        
        # Try alternative approach for problematic macro files
        try:
            st.sidebar.info("🔄 Trying alternative method for macro file...")
            
            # Use xlrd for older .xls files or specific .xlsm handling
            if file_path.endswith('.xlsm'):
                # For .xlsm files, try reading without macros
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                data_dict = {}
                
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                        if not df.empty:
                            data_dict[sheet_name] = df
                    except:
                        continue
                        
                return data_dict
            
        except Exception as e2:
            st.sidebar.error(f"❌ Alternative method failed: {str(e2)}")
            
        return {}

def extract_sector_performance(data_dict):
    """Extract sector-wise performance data"""
    sector_data = {}
    
    for sheet_name, df in data_dict.items():
        if 'SECTOR' in sheet_name.upper() or 'Dashboard' in sheet_name:
            # Look for sector performance data
            for _, row in df.iterrows():
                try:
                    sector_name = str(row.iloc[0]).strip()
                    if any(sector in sector_name.upper() for sector in ['NIFTY', 'BANK', 'IT', 'PHARMA', 'METAL', 'AUTO', 'ENERGY', 'FMCG']):
                        # Extract bullish/bearish percentages
                        for col in df.columns:
                            if 'BULLISH' in str(col).upper():
                                bullish_val = pd.to_numeric(row[col], errors='coerce')
                                if not pd.isna(bullish_val):
                                    sector_data[sector_name] = {
                                        'bullish': bullish_val,
                                        'bearish': 100 - bullish_val if bullish_val <= 100 else 0
                                    }
                                break
                except:
                    continue
    
    return sector_data

def extract_complete_stock_data(data_dict):
    """Extract complete stock data with all buildup types"""
    stock_data = {
        'long_buildup': [],
        'short_covering': [],
        'short_buildup': [],
        'long_unwinding': [],
        'bullish_stocks': [],
        'bearish_stocks': [],
        'all_stocks': []
    }
    
    for sheet_name, df in data_dict.items():
        if any(term in sheet_name.upper() for term in ['STOCK', 'NIFTY 50', 'DASHBOARD']):
            try:
                # Find relevant columns
                symbol_col = find_column_by_keywords(df.columns, ['STOCK NAME', 'SYMBOL', 'NAME'])
                change_col = find_column_by_keywords(df.columns, ['CHANGE %', '%', 'CHG'])
                price_col = find_column_by_keywords(df.columns, ['PRICE', 'LTP', 'CLOSE'])
                oi_col = find_column_by_keywords(df.columns, ['OI', 'OPEN INT'])
                oi_change_col = find_column_by_keywords(df.columns, ['Change in OI', 'OI CHG'])
                volume_col = find_column_by_keywords(df.columns, ['Volume', 'VOL'])
                buildup_col = find_column_by_keywords(df.columns, ['Buildup', 'BUILDUP', 'TREND'])
                sentiment_col = find_column_by_keywords(df.columns, ['SENTIMENT', 'SIGNAL'])
                
                if not symbol_col:
                    continue
                
                # Process each stock
                for _, row in df.iterrows():
                    try:
                        symbol = str(row[symbol_col]).strip()
                        if symbol == 'nan' or symbol == '' or 'NSE:' not in symbol:
                            continue
                        
                        # Clean symbol name
                        symbol = symbol.replace('NSE:', '')
                        
                        # Extract all data
                        change = pd.to_numeric(row[change_col], errors='coerce') if change_col else 0
                        price = pd.to_numeric(row[price_col], errors='coerce') if price_col else 0
                        oi = pd.to_numeric(row[oi_col], errors='coerce') if oi_col else 0
                        oi_change = pd.to_numeric(row[oi_change_col], errors='coerce') if oi_change_col else 0
                        volume = pd.to_numeric(row[volume_col], errors='coerce') if volume_col else 0
                        buildup = str(row[buildup_col]).strip() if buildup_col else 'Unknown'
                        sentiment = str(row[sentiment_col]).strip() if sentiment_col else 'Wait For Signal'
                        
                        if pd.isna(change):
                            continue
                        
                        stock_info = {
                            'symbol': symbol,
                            'change': change,
                            'price': price,
                            'oi': oi,
                            'oi_change': oi_change,
                            'volume': volume,
                            'buildup': buildup,
                            'sentiment': sentiment
                        }
                        
                        # Add to all stocks
                        stock_data['all_stocks'].append(stock_info)
                        
                        # Categorize by buildup type
                        if 'longBuildup' in buildup or 'Long Buildup' in buildup:
                            stock_data['long_buildup'].append(stock_info)
                        elif 'shortCover' in buildup or 'Short Cover' in buildup:
                            stock_data['short_covering'].append(stock_info)
                        elif 'shortBuildup' in buildup or 'Short Buildup' in buildup:
                            stock_data['short_buildup'].append(stock_info)
                        elif 'longUnwind' in buildup or 'Long Unwind' in buildup:
                            stock_data['long_unwinding'].append(stock_info)
                        
                        # Categorize by performance
                        if change > 0.5:
                            stock_data['bullish_stocks'].append(stock_info)
                        elif change < -0.5:
                            stock_data['bearish_stocks'].append(stock_info)
                    
                    except Exception as e:
                        continue
            
            except Exception as e:
                st.write(f"Error processing sheet {sheet_name}: {str(e)}")
                continue
    
    # Sort all categories
    for category in stock_data:
        if category == 'bearish_stocks':
            stock_data[category] = sorted(stock_data[category], key=lambda x: x['change'])[:50]
        else:
            stock_data[category] = sorted(stock_data[category], key=lambda x: x['change'], reverse=True)[:50]
    
    return stock_data

def display_sector_performance(sector_data):
    """Display sector performance grid"""
    if not sector_data:
        return
    
    st.header("📊 All Sector Performance")
    
    # Create grid layout for sectors
    cols = st.columns(4)
    col_idx = 0
    
    for sector_name, data in sector_data.items():
        with cols[col_idx % 4]:
            sector_class = "bullish-sector" if data['bullish'] > 60 else "bearish-sector" if data['bullish'] < 40 else ""
            
            st.markdown(f"""
            <div class="sector-performance {sector_class}">
                <h4>{sector_name}</h4>
                <div style="display: flex; justify-content: space-between;">
                    <div>
                        <strong>BULLISH</strong><br>
                        <span style="font-size: 1.2em;">{data['bullish']:.1f}%</span>
                    </div>
                    <div>
                        <strong>BEARISH</strong><br>
                        <span style="font-size: 1.2em;">{data['bearish']:.1f}%</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        col_idx += 1

def display_stock_analysis_table(stock_data, category, title, css_class):
    """Display stock data in a formatted table"""
    if not stock_data[category]:
        return
    
    st.subheader(title)
    
    # Create DataFrame for better display
    df_display = pd.DataFrame(stock_data[category])
    
    # Format the data for display
    df_display['Change %'] = df_display['change'].apply(lambda x: f"{x:+.2f}%")
    df_display['Price'] = df_display['price'].apply(lambda x: f"₹{x:.2f}" if x > 0 else "N/A")
    df_display['OI'] = df_display['oi'].apply(lambda x: f"{x:,.0f}" if x > 0 else "N/A")
    df_display['Volume'] = df_display['volume'].apply(lambda x: f"{x:,.0f}" if x > 0 else "N/A")
    df_display['OI Change %'] = df_display['oi_change'].apply(lambda x: f"{x:+.2f}%" if not pd.isna(x) and x != 0 else "N/A")
    
    # Select columns for display
    display_cols = ['symbol', 'Change %', 'Price', 'OI', 'OI Change %', 'Volume', 'buildup', 'sentiment']
    df_show = df_display[display_cols]
    df_show.columns = ['Stock', 'Change %', 'Price', 'OI', 'OI Change %', 'Volume', 'Buildup', 'Sentiment']
    
    # Display in chunks of 10
    for i in range(0, min(30, len(df_show)), 10):
        chunk = df_show.iloc[i:i+10]
        
        # Create HTML table for better formatting
        table_html = f"""
        <div class="data-table">
            <table style="width: 100%; border-collapse: collapse;">
                <thead style="background: #f8f9fa;">
                    <tr>
                        <th style="padding: 10px; text-align: left; border-bottom: 2px solid #dee2e6;">Stock</th>
                        <th style="padding: 10px; text-align: center; border-bottom: 2px solid #dee2e6;">Change %</th>
                        <th style="padding: 10px; text-align: center; border-bottom: 2px solid #dee2e6;">Price</th>
                        <th style="padding: 10px; text-align: center; border-bottom: 2px solid #dee2e6;">OI</th>
                        <th style="padding: 10px; text-align: center; border-bottom: 2px solid #dee2e6;">OI Change %</th>
                        <th style="padding: 10px; text-align: center; border-bottom: 2px solid #dee2e6;">Volume</th>
                        <th style="padding: 10px; text-align: center; border-bottom: 2px solid #dee2e6;">Buildup</th>
                        <th style="padding: 10px; text-align: center; border-bottom: 2px solid #dee2e6;">Sentiment</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for _, row in chunk.iterrows():
            # Determine row color based on change
            row_style = ""
            if category in ['long_buildup', 'bullish_stocks']:
                row_style = "background-color: rgba(40, 167, 69, 0.1);"
            elif category in ['short_covering']:
                row_style = "background-color: rgba(23, 162, 184, 0.1);"
            elif category in ['short_buildup', 'bearish_stocks']:
                row_style = "background-color: rgba(220, 53, 69, 0.1);"
            elif category in ['long_unwinding']:
                row_style = "background-color: rgba(255, 193, 7, 0.1);"
            
            sentiment_class = "bullish-signal" if "Bullish" in str(row['Sentiment']) else "wait-signal"
            
            table_html += f"""
                <tr style="{row_style}">
                    <td style="padding: 8px; border-bottom: 1px solid #dee2e6;"><strong>{row['Stock']}</strong></td>
                    <td style="padding: 8px; text-align: center; border-bottom: 1px solid #dee2e6;"><strong>{row['Change %']}</strong></td>
                    <td style="padding: 8px; text-align: center; border-bottom: 1px solid #dee2e6;">{row['Price']}</td>
                    <td style="padding: 8px; text-align: center; border-bottom: 1px solid #dee2e6;">{row['OI']}</td>
                    <td style="padding: 8px; text-align: center; border-bottom: 1px solid #dee2e6;">{row['OI Change %']}</td>
                    <td style="padding: 8px; text-align: center; border-bottom: 1px solid #dee2e6;">{row['Volume']}</td>
                    <td style="padding: 8px; text-align: center; border-bottom: 1px solid #dee2e6;"><span class="{css_class}">{row['Buildup']}</span></td>
                    <td style="padding: 8px; text-align: center; border-bottom: 1px solid #dee2e6;"><span class="{sentiment_class}">{row['Sentiment']}</span></td>
                </tr>
            """
        
        table_html += """
                </tbody>
            </table>
        </div>
        """
        
        st.markdown(table_html, unsafe_allow_html=True)

def display_summary_metrics(stock_data):
    """Display summary metrics"""
    st.header("📈 Market Summary")
    
    cols = st.columns(6)
    
    metrics = [
        ("Long Buildup", len(stock_data['long_buildup']), "#28a745"),
        ("Short Covering", len(stock_data['short_covering']), "#17a2b8"),
        ("Short Buildup", len(stock_data['short_buildup']), "#dc3545"),
        ("Long Unwinding", len(stock_data['long_unwinding']), "#ffc107"),
        ("Bullish Stocks", len(stock_data['bullish_stocks']), "#28a745"),
        ("Bearish Stocks", len(stock_data['bearish_stocks']), "#dc3545")
    ]
    
    for i, (label, count, color) in enumerate(metrics):
        with cols[i]:
            st.markdown(f"""
            <div class="metric-card" style="background: linear-gradient(135deg, {color}, #495057);">
                <h3>{count}</h3>
                <p>{label}</p>
            </div>
            """, unsafe_allow_html=True)

def display_live_dashboard(data_dict):
    """Display complete live F&O dashboard"""
    
    # Live header
    st.markdown(f"""
    <div class="dashboard-header">
        <h1>📊 Complete F&O Trading Dashboard</h1>
        <p class="live-indicator">● LIVE DATA</p>
        <p>Real-time Analysis - {datetime.now().strftime("%d %B %Y, %H:%M:%S")}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Extract data
    sector_data = extract_sector_performance(data_dict)
    stock_data = extract_complete_stock_data(data_dict)
    
    # Display sector performance
    display_sector_performance(sector_data)
    
    # Display summary metrics
    display_summary_metrics(stock_data)
    
    # Display stock analysis by categories
    st.header("🎯 Detailed Stock Analysis")
    
    # Create tabs for different categories
    tabs = st.tabs(["Long Buildup", "Short Covering", "Short Buildup", "Long Unwinding", "All Bullish", "All Bearish"])
    
    with tabs[0]:
        display_stock_analysis_table(stock_data, 'long_buildup', "🟢 Long Buildup Stocks", "long-buildup")
    
    with tabs[1]:
        display_stock_analysis_table(stock_data, 'short_covering', "🔵 Short Covering Stocks", "short-covering")
    
    with tabs[2]:
        display_stock_analysis_table(stock_data, 'short_buildup', "🔴 Short Buildup Stocks", "short-buildup")
    
    with tabs[3]:
        display_stock_analysis_table(stock_data, 'long_unwinding', "🟡 Long Unwinding Stocks", "long-unwinding")
    
    with tabs[4]:
        display_stock_analysis_table(stock_data, 'bullish_stocks', "📈 All Bullish Stocks", "bullish-signal")
    
    with tabs[5]:
        display_stock_analysis_table(stock_data, 'bearish_stocks', "📉 All Bearish Stocks", "wait-signal")
    
    # Data source info
    st.markdown("---")
    st.markdown(f"""
    **Data Source:** {len(data_dict)} Excel sheets processed  
    **Last Updated:** {datetime.now().strftime("%H:%M:%S")}  
    **Total Stocks Analyzed:** {len(stock_data['all_stocks'])}  
    **Auto-refresh:** Every 30 seconds
    """)

# Main execution
def main():
    # Sidebar controls
    st.sidebar.title("📊 Dashboard Controls")
    
    # File upload with macro file handling
    st.sidebar.markdown("### 📁 Upload Options")
    
    # Display warning about macro files
    st.sidebar.warning("⚠️ Macro-enabled files (.xlsm) not supported. Please convert to .xlsx format.")
    
    uploaded_file = st.sidebar.file_uploader(
        "Upload F&O Excel File", 
        type=["xlsx", "xls", "xlsm"],
        help="Supports .xlsx, .xls, and .xlsm (macro-enabled) files"
    )
    
    # Alternative data input methods
    st.sidebar.markdown("### 🔄 Alternative Input Methods")
    
    data_dict = None
    
    # Option to paste CSV data
    if st.sidebar.checkbox("📋 Paste CSV Data"):
        csv_data = st.sidebar.text_area(
            "Paste your stock data (CSV format):",
            height=150,
            placeholder="Symbol,Change%,Price,OI,Volume,Buildup,Sentiment\nRELIANCE,1.45,2500,150000,50000,longBuildup,Bullish"
        )
        
        if csv_data:
            try:
                import io
                df = pd.read_csv(io.StringIO(csv_data))
                st.sidebar.success(f"✅ {len(df)} rows loaded from CSV")
                data_dict = convert_csv_to_data_dict(df)
            except Exception as e:
                st.sidebar.error(f"❌ Error parsing CSV: {str(e)}")
    
    # Sample data option for demo
    if st.sidebar.checkbox("🎯 Load Sample Data"):
        data_dict = create_sample_data()
        st.sidebar.success("✅ Sample data loaded for demonstration")
    
    # Refresh controls
    auto_refresh = st.sidebar.checkbox("Auto Refresh (30s)", value=True)
    
    if st.sidebar.button("🔄 Manual Refresh"):
        st.experimental_rerun()
    
    # Display current time
    st.sidebar.markdown(f"**Current Time:** {datetime.now().strftime('%H:%M:%S')}")
    
    # Instructions updated for macro files
    if not data_dict:
        st.markdown("""
        ## 📊 F&O Dashboard - Now Supports Macro Files!
        
        ### ✅ **Supported File Formats:**
        - **Excel Workbook (.xlsx)**
        - **Excel 97-2003 (.xls)**
        - **Excel Macro-Enabled (.xlsm)** - NEW!
        
        ### 🔧 **How Macro Files Are Handled:**
        
        **What Works:**
        - ✅ All data extraction from sheets
        - ✅ Multiple sheet processing  
        - ✅ Complex formulas (calculated values)
        - ✅ Charts and tables data
        
        **What's Ignored:**
        - ⚠️ VBA Macros (security reason)
        - ⚠️ ActiveX controls
        - ⚠️ Custom functions
        
        ### 💡 **Quick Start Options:**
        
        **Option 1: Upload Your .xlsm File**
        - Simply upload your macro-enabled file
        - The system will extract all data automatically
        - Macros will be safely ignored
        
        **Option 2: Try Sample Data**  
        - Click "🎯 Load Sample Data" to see the dashboard
        
        **Option 3: Manual CSV Input**
        - Export your data to CSV
        - Use "📋 Paste CSV Data" option
        
        ### 📋 **Expected Data Structure:**
        ```
        Stock Name    | Change% | Price | OI      | Volume  | Buildup      | Sentiment
        RELIANCE      | 1.45    | 2500  | 150000  | 85000   | longBuildup  | Bullish
        TCS           | 1.87    | 3650  | 125000  | 65000   | longBuildup  | Bullish
        HDFC          | 1.25    | 1650  | 95000   | 75000   | shortCover   | Bullish
        ```
        """)
    
    if uploaded_file is not None:
        # Determine file type and show processing info
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension == 'xlsm':
            st.info("🔄 Processing macro-enabled file... Extracting data only.")
        
        # Save uploaded file with original extension
        temp_filename = f"temp_fo_file.{file_extension}"
        with open(temp_filename, "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Read and process data
        with st.spinner(f"Loading {file_extension.upper()} file..."):
            data_dict = read_comprehensive_data(temp_filename)
            
        # Clean up temp file
        try:
            os.remove(temp_filename)
        except:
            pass
    
    if data_dict:
        # Display dashboard
        display_live_dashboard(data_dict)
        
        # Auto-refresh functionality
        if auto_refresh:
            try:
                from streamlit_autorefresh import st_autorefresh
                st_autorefresh(interval=30 * 1000, key="fo_data_refresh")
            except ImportError:
                st.sidebar.warning("Install streamlit-autorefresh for auto-refresh: `pip install streamlit-autorefresh`")
    
    elif not data_dict and not uploaded_file:
        # Show file format help
        st.info("📁 Please choose one of the options in the sidebar to load your F&O data.")
        
        # Show expected data columns
        st.markdown("""
        ### 📊 Expected Data Columns:
        | Column | Description | Example |
        |--------|-------------|---------|
        | **Symbol/Stock Name** | Stock symbol | RELIANCE, TCS, INFY |
        | **Change %** | Price change percentage | 1.45, -2.15 |
        | **Price/LTP** | Last traded price | 2500, 1650 |
        | **OI** | Open Interest | 150000, 95000 |
        | **Volume** | Trading volume | 85000, 75000 |
        | **Buildup** | Position type | longBuildup, shortCover |
        | **Sentiment** | Market sentiment | Bullish, Wait For Signal |
        """)

if __name__ == "__main__":
    main()
