import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
import os

st.set_page_config(page_title="Stock Trading Dashboard", page_icon="📈", layout="wide")

# Trading-focused CSS
st.markdown("""
<style>
.buy-signal {
    background: linear-gradient(135deg, #00b09b, #96c93d);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
    text-align: center;
}
.sell-signal {
    background: linear-gradient(135deg, #ee0979, #ff6a00);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
    text-align: center;
}
.hold-signal {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
    text-align: center;
}
.sector-card {
    border: 1px solid #ddd;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.5rem 0;
    background-color: #f8f9fa;
}
.metric-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 1rem;
    margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

def read_excel_trading_sheets(file_path):
    """Read specific trading-focused sheets"""
    key_sheets = [
        'Dashboard', 'Stock Dashboard', 'Sector Dashboard', 
        'OC_1', 'OC_2', 'Screener', 'PCR & OI Chart'
    ]
    
    try:
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        for sheet_name in key_sheets:
            if sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    if not df.empty:
                        data_dict[sheet_name] = df
                        st.success(f"Loaded {sheet_name}: {len(df)} rows")
                except Exception as e:
                    st.warning(f"Could not read {sheet_name}: {str(e)}")
        
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return {}

def analyze_stock_signals(df):
    """Analyze stocks for buy/sell/hold signals"""
    signals = {
        'strong_buy': [],
        'buy': [],
        'hold': [],
        'sell': [],
        'strong_sell': []
    }
    
    try:
        # Find relevant columns
        symbol_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['SYMBOL', 'STOCK', 'NAME'])]
        change_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['CHANGE', '%', 'CHG'])]
        volume_cols = [col for col in df.columns if 'VOLUME' in str(col).upper()]
        oi_cols = [col for col in df.columns if 'OI' in str(col).upper()]
        
        if not symbol_cols or not change_cols:
            return signals
        
        symbol_col = symbol_cols[0]
        change_col = change_cols[0]
        
        # Clean data
        analysis_df = df[[symbol_col, change_col]].copy()
        if volume_cols:
            analysis_df['Volume'] = df[volume_cols[0]]
        if oi_cols:
            analysis_df['OI'] = df[oi_cols[0]]
        
        analysis_df[change_col] = pd.to_numeric(analysis_df[change_col], errors='coerce')
        analysis_df = analysis_df.dropna()
        
        for _, row in analysis_df.iterrows():
            symbol = row[symbol_col]
            change = row[change_col]
            
            # Trading signals based on change percentage
            if change > 5:
                signals['strong_buy'].append((symbol, change))
            elif change > 2:
                signals['buy'].append((symbol, change))
            elif change > -2:
                signals['hold'].append((symbol, change))
            elif change > -5:
                signals['sell'].append((symbol, change))
            else:
                signals['strong_sell'].append((symbol, change))
        
        # Sort by absolute change
        for signal_type in signals:
            signals[signal_type] = sorted(signals[signal_type], key=lambda x: abs(x[1]), reverse=True)[:10]
        
        return signals
    
    except Exception as e:
        st.warning(f"Error analyzing stock signals: {e}")
        return signals

def analyze_options_data(nifty_df, banknifty_df):
    """Analyze Nifty and Bank Nifty options data"""
    options_summary = {}
    
    for name, df in [('Nifty', nifty_df), ('Bank Nifty', banknifty_df)]:
        if df is not None and not df.empty:
            summary = {}
            
            # Find PCR data
            call_oi_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
            put_oi_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
            
            if call_oi_cols and put_oi_cols:
                call_oi = pd.to_numeric(df[call_oi_cols[0]], errors='coerce').fillna(0).sum()
                put_oi = pd.to_numeric(df[put_oi_cols[0]], errors='coerce').fillna(0).sum()
                
                if call_oi > 0:
                    pcr = put_oi / call_oi
                    summary['pcr'] = pcr
                    summary['call_oi'] = call_oi
                    summary['put_oi'] = put_oi
                    
                    # Trading signal based on PCR
                    if pcr > 1.5:
                        summary['signal'] = 'STRONG_BEARISH'
                    elif pcr > 1.2:
                        summary['signal'] = 'BEARISH'
                    elif pcr < 0.6:
                        summary['signal'] = 'STRONG_BULLISH'
                    elif pcr < 0.8:
                        summary['signal'] = 'BULLISH'
                    else:
                        summary['signal'] = 'NEUTRAL'
            
            # Find strike prices for support/resistance
            strike_cols = [col for col in df.columns if 'STRIKE' in str(col).upper()]
            if strike_cols and call_oi_cols and put_oi_cols:
                try:
                    strike_col = strike_cols[0]
                    call_col = call_oi_cols[0]
                    put_col = put_oi_cols[0]
                    
                    clean_df = df[[strike_col, call_col, put_col]].copy()
                    clean_df[call_col] = pd.to_numeric(clean_df[call_col], errors='coerce')
                    clean_df[put_col] = pd.to_numeric(clean_df[put_col], errors='coerce')
                    clean_df = clean_df.dropna()
                    
                    if not clean_df.empty:
                        resistance_idx = clean_df[call_col].idxmax()
                        support_idx = clean_df[put_col].idxmax()
                        
                        summary['resistance'] = clean_df.loc[resistance_idx, strike_col]
                        summary['support'] = clean_df.loc[support_idx, strike_col]
                except:
                    pass
            
            options_summary[name] = summary
    
    return options_summary

def analyze_sector_performance(df):
    """Analyze sector performance for trading opportunities"""
    sector_analysis = {}
    
    try:
        # Find sector and performance columns
        sector_cols = [col for col in df.columns if 'SECTOR' in str(col).upper()]
        change_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['CHANGE', '%', 'CHG'])]
        symbol_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['SYMBOL', 'STOCK', 'NAME'])]
        
        if not sector_cols or not change_cols:
            return sector_analysis
        
        sector_col = sector_cols[0]
        change_col = change_cols[0]
        
        # Group by sector
        sector_df = df.groupby(sector_col).agg({
            change_col: 'mean',
            symbol_cols[0] if symbol_cols else sector_col: 'count'
        }).reset_index()
        
        sector_df.columns = ['Sector', 'Avg_Change', 'Stock_Count']
        sector_df = sector_df.sort_values('Avg_Change', ascending=False)
        
        for _, row in sector_df.iterrows():
            sector = row['Sector']
            avg_change = row['Avg_Change']
            
            # Get individual stocks in this sector
            sector_stocks = df[df[sector_col] == sector]
            if symbol_cols and change_cols:
                stocks_list = []
                for _, stock_row in sector_stocks.iterrows():
                    stocks_list.append((stock_row[symbol_cols[0]], stock_row[change_col]))
                
                # Sort stocks by performance
                stocks_list = sorted(stocks_list, key=lambda x: x[1], reverse=True)
                
                sector_analysis[sector] = {
                    'avg_change': avg_change,
                    'stock_count': row['Stock_Count'],
                    'top_stocks': stocks_list[:5],
                    'signal': 'BUY' if avg_change > 1 else 'SELL' if avg_change < -1 else 'HOLD'
                }
        
        return sector_analysis
    
    except Exception as e:
        st.warning(f"Error analyzing sectors: {e}")
        return {}

def display_trading_dashboard(stock_signals, options_summary, sector_analysis):
    """Display comprehensive trading dashboard"""
    
    st.title("Stock Trading Command Center")
    st.caption(f"Live Analysis - {datetime.now().strftime('%H:%M:%S')}")
    
    # Options Market Overview
    st.header("Options Market Analysis")
    
    if options_summary:
        cols = st.columns(len(options_summary))
        
        for i, (index_name, data) in enumerate(options_summary.items()):
            with cols[i]:
                signal = data.get('signal', 'NEUTRAL')
                pcr = data.get('pcr', 0)
                
                if 'BULLISH' in signal:
                    card_class = 'buy-signal'
                elif 'BEARISH' in signal:
                    card_class = 'sell-signal'
                else:
                    card_class = 'hold-signal'
                
                st.markdown(f"""
                <div class="{card_class}">
                    <h3>{index_name}</h3>
                    <h4>{signal}</h4>
                    <p>PCR: {pcr:.3f}</p>
                </div>
                """, unsafe_allow_html=True)
                
                if 'support' in data and 'resistance' in data:
                    st.success(f"Support: {data['support']:,.0f}")
                    st.error(f"Resistance: {data['resistance']:,.0f}")
    
    # Stock Trading Signals
    st.header("Stock Trading Signals")
    
    signal_tabs = st.tabs(['Strong Buy', 'Buy', 'Hold', 'Sell', 'Strong Sell'])
    
    signal_types = ['strong_buy', 'buy', 'hold', 'sell', 'strong_sell']
    signal_colors = ['success', 'info', 'warning', 'error', 'error']
    
    for i, (tab, signal_type, color) in enumerate(zip(signal_tabs, signal_types, signal_colors)):
        with tab:
            stocks = stock_signals.get(signal_type, [])
            if stocks:
                st.subheader(f"{signal_type.replace('_', ' ').title()} Signals ({len(stocks)} stocks)")
                
                for stock, change in stocks[:10]:
                    if color == 'success':
                        st.success(f"{stock}: +{change:.2f}%")
                    elif color == 'error':
                        st.error(f"{stock}: {change:.2f}%")
                    elif color == 'warning':
                        st.warning(f"{stock}: {change:.2f}%")
                    else:
                        st.info(f"{stock}: {change:.2f}%")
            else:
                st.info(f"No {signal_type.replace('_', ' ')} signals detected")
    
    # Sector Analysis
    if sector_analysis:
        st.header("Sector Trading Opportunities")
        
        # Split sectors into columns
        sector_cols = st.columns(2)
        sectors_list = list(sector_analysis.items())
        
        for i, (sector, data) in enumerate(sectors_list):
            col = sector_cols[i % 2]
            
            with col:
                signal = data.get('signal', 'HOLD')
                avg_change = data.get('avg_change', 0)
                
                if signal == 'BUY':
                    signal_class = 'buy-signal'
                elif signal == 'SELL':
                    signal_class = 'sell-signal'
                else:
                    signal_class = 'hold-signal'
                
                st.markdown(f"""
                <div class="sector-card">
                    <h4>{sector}</h4>
                    <div class="{signal_class}" style="padding: 0.5rem; margin: 0.5rem 0;">
                        <strong>{signal}</strong> - {avg_change:.2f}%
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Show top stocks in sector
                top_stocks = data.get('top_stocks', [])
                if top_stocks:
                    st.write("**Top Stocks:**")
                    for stock, change in top_stocks[:3]:
                        if change > 0:
                            st.success(f"{stock}: +{change:.2f}%")
                        else:
                            st.error(f"{stock}: {change:.2f}%")
    
    # Quick Action Summary
    st.header("Quick Trading Summary")
    
    summary_cols = st.columns(4)
    
    with summary_cols[0]:
        strong_buys = len(stock_signals.get('strong_buy', []))
        st.markdown(f"""
        <div class="buy-signal">
            <h3>{strong_buys}</h3>
            <p>Strong Buy Signals</p>
        </div>
        """, unsafe_allow_html=True)
    
    with summary_cols[1]:
        buys = len(stock_signals.get('buy', []))
        st.markdown(f"""
        <div class="buy-signal" style="opacity: 0.8;">
            <h3>{buys}</h3>
            <p>Buy Signals</p>
        </div>
        """, unsafe_allow_html=True)
    
    with summary_cols[2]:
        sells = len(stock_signals.get('sell', []))
        st.markdown(f"""
        <div class="sell-signal" style="opacity: 0.8;">
            <h3>{sells}</h3>
            <p>Sell Signals</p>
        </div>
        """, unsafe_allow_html=True)
    
    with summary_cols[3]:
        strong_sells = len(stock_signals.get('strong_sell', []))
        st.markdown(f"""
        <div class="sell-signal">
            <h3>{strong_sells}</h3>
            <p>Strong Sell Signals</p>
        </div>
        """, unsafe_allow_html=True)

def main():
    uploaded_file = st.file_uploader("Upload Trading Excel File", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        with st.spinner("Analyzing trading data..."):
            data_dict = read_excel_trading_sheets(temp_path)
        
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data_dict:
            # Analyze stock signals
            stock_signals = {}
            if 'Stock Dashboard' in data_dict:
                stock_signals = analyze_stock_signals(data_dict['Stock Dashboard'])
            elif 'Dashboard' in data_dict:
                stock_signals = analyze_stock_signals(data_dict['Dashboard'])
            
            # Analyze options
            nifty_df = data_dict.get('OC_1')
            banknifty_df = data_dict.get('OC_2')
            options_summary = analyze_options_data(nifty_df, banknifty_df)
            
            # Analyze sectors
            sector_analysis = {}
            if 'Sector Dashboard' in data_dict:
                sector_analysis = analyze_sector_performance(data_dict['Sector Dashboard'])
            
            # Display dashboard
            display_trading_dashboard(stock_signals, options_summary, sector_analysis)
            
        else:
            st.error("Could not load trading data")
    else:
        st.info("Upload your Excel file to see trading signals and market analysis")

if __name__ == "__main__":
    main()
