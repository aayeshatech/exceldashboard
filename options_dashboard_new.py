import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime
import os

st.set_page_config(page_title="Professional Trading Dashboard", page_icon="⚡", layout="wide")

# Professional trading CSS
st.markdown("""
<style>
.trading-header {
    background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    text-align: center;
    margin-bottom: 2rem;
}
.strategy-card {
    border: 2px solid #007bff;
    border-radius: 10px;
    padding: 1.5rem;
    margin: 1rem 0;
    background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
}
.long-signal {
    background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
}
.short-signal {
    background: linear-gradient(135deg, #dc3545 0%, #fd7e14 100%);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
}
.neutral-signal {
    background: linear-gradient(135deg, #6c757d 0%, #adb5bd 100%);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
}
.pcr-bullish { background: #d4edda; color: #155724; padding: 0.5rem; border-radius: 5px; }
.pcr-bearish { background: #f8d7da; color: #721c24; padding: 0.5rem; border-radius: 5px; }
.pcr-neutral { background: #fff3cd; color: #856404; padding: 0.5rem; border-radius: 5px; }
.metrics-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
    gap: 1rem;
    margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

def read_all_trading_sheets(file_path):
    """Read all trading-relevant sheets from Excel"""
    try:
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        # Read all available sheets
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                if not df.empty:
                    data_dict[sheet_name] = df
                    st.info(f"Loaded {sheet_name}: {len(df)} rows, {len(df.columns)} cols")
            except Exception as e:
                st.warning(f"Skipped {sheet_name}: {str(e)}")
        
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return {}

def extract_pcr_strategy(pcr_df):
    """Extract PCR-based trading strategy"""
    pcr_strategy = {}
    
    try:
        # Look for PCR values
        for col in pcr_df.columns:
            if 'PCR' in str(col).upper():
                pcr_values = pd.to_numeric(pcr_df[col], errors='coerce').dropna()
                if not pcr_values.empty:
                    current_pcr = pcr_values.iloc[-1]
                    
                    if current_pcr > 1.5:
                        signal = "STRONG_BEARISH"
                        strategy = "Consider PUT buying or short positions"
                        confidence = "HIGH"
                    elif current_pcr > 1.2:
                        signal = "BEARISH"
                        strategy = "Bias towards PUT positions"
                        confidence = "MEDIUM"
                    elif current_pcr < 0.6:
                        signal = "STRONG_BULLISH"
                        strategy = "Consider CALL buying or long positions"
                        confidence = "HIGH"
                    elif current_pcr < 0.8:
                        signal = "BULLISH"
                        strategy = "Bias towards CALL positions"
                        confidence = "MEDIUM"
                    else:
                        signal = "NEUTRAL"
                        strategy = "Range-bound trading or straddle"
                        confidence = "LOW"
                    
                    pcr_strategy[col] = {
                        'value': current_pcr,
                        'signal': signal,
                        'strategy': strategy,
                        'confidence': confidence
                    }
        
        return pcr_strategy
    except Exception as e:
        st.warning(f"PCR analysis error: {e}")
        return {}

def extract_stock_strategies(dashboard_df):
    """Extract stock trading strategies from dashboard"""
    strategies = {
        'long_candidates': [],
        'short_candidates': [],
        'breakout_stocks': [],
        'high_volume_plays': []
    }
    
    try:
        # Find relevant columns
        symbol_cols = [col for col in dashboard_df.columns if any(term in str(col).upper() for term in ['SYMBOL', 'STOCK', 'NAME'])]
        change_cols = [col for col in dashboard_df.columns if any(term in str(col).upper() for term in ['CHANGE', '%'])]
        volume_cols = [col for col in dashboard_df.columns if 'VOLUME' in str(col).upper()]
        ltp_cols = [col for col in dashboard_df.columns if 'LTP' in str(col).upper()]
        
        if not symbol_cols or not change_cols:
            return strategies
        
        symbol_col = symbol_cols[0]
        change_col = change_cols[0]
        
        # Clean data
        analysis_df = dashboard_df[[symbol_col, change_col]].copy()
        if volume_cols:
            analysis_df['Volume'] = pd.to_numeric(dashboard_df[volume_cols[0]], errors='coerce')
        if ltp_cols:
            analysis_df['LTP'] = pd.to_numeric(dashboard_df[ltp_cols[0]], errors='coerce')
        
        analysis_df[change_col] = pd.to_numeric(analysis_df[change_col], errors='coerce')
        analysis_df = analysis_df.dropna()
        
        # Strategy classification
        for _, row in analysis_df.iterrows():
            symbol = row[symbol_col]
            change = row[change_col]
            volume = row.get('Volume', 0)
            ltp = row.get('LTP', 0)
            
            # Long candidates (strong bullish momentum)
            if change > 3:
                strategies['long_candidates'].append({
                    'symbol': symbol,
                    'change': change,
                    'volume': volume,
                    'ltp': ltp,
                    'strategy': f"LONG at {ltp:.2f}, Target: {ltp * 1.05:.2f}, SL: {ltp * 0.97:.2f}"
                })
            
            # Short candidates (strong bearish momentum)
            elif change < -3:
                strategies['short_candidates'].append({
                    'symbol': symbol,
                    'change': change,
                    'volume': volume,
                    'ltp': ltp,
                    'strategy': f"SHORT at {ltp:.2f}, Target: {ltp * 0.95:.2f}, SL: {ltp * 1.03:.2f}"
                })
            
            # Breakout stocks (moderate momentum with volume)
            elif 1.5 < abs(change) < 3 and volume > 100000:
                strategies['breakout_stocks'].append({
                    'symbol': symbol,
                    'change': change,
                    'volume': volume,
                    'ltp': ltp,
                    'strategy': f"BREAKOUT play - Watch for continuation"
                })
            
            # High volume plays
            if volume > 500000:
                strategies['high_volume_plays'].append({
                    'symbol': symbol,
                    'change': change,
                    'volume': volume,
                    'ltp': ltp
                })
        
        # Sort by performance
        for strategy_type in ['long_candidates', 'short_candidates']:
            strategies[strategy_type] = sorted(
                strategies[strategy_type], 
                key=lambda x: abs(x['change']), 
                reverse=True
            )[:10]
        
        return strategies
    
    except Exception as e:
        st.warning(f"Stock strategy analysis error: {e}")
        return strategies

def extract_sector_strategies(sector_df):
    """Extract sector-wise trading strategies"""
    sector_strategies = {}
    
    try:
        # Find sector columns
        sector_cols = [col for col in sector_df.columns if 'SECTOR' in str(col).upper()]
        change_cols = [col for col in sector_df.columns if any(term in str(col).upper() for term in ['CHANGE', '%'])]
        
        if not sector_cols or not change_cols:
            return sector_strategies
        
        sector_col = sector_cols[0]
        change_col = change_cols[0]
        
        # Group by sector
        sector_performance = sector_df.groupby(sector_col)[change_col].mean().sort_values(ascending=False)
        
        for sector, avg_change in sector_performance.items():
            if avg_change > 2:
                signal = "SECTOR_ROTATION_BUY"
                strategy = f"Rotate INTO {sector} - Strong outperformance"
                confidence = "HIGH"
            elif avg_change > 1:
                signal = "SECTOR_OVERWEIGHT"
                strategy = f"Overweight {sector} positions"
                confidence = "MEDIUM"
            elif avg_change < -2:
                signal = "SECTOR_ROTATION_SELL"
                strategy = f"Rotate OUT OF {sector} - Underperforming"
                confidence = "HIGH"
            elif avg_change < -1:
                signal = "SECTOR_UNDERWEIGHT"
                strategy = f"Underweight {sector} positions"
                confidence = "MEDIUM"
            else:
                signal = "SECTOR_NEUTRAL"
                strategy = f"Maintain neutral {sector} allocation"
                confidence = "LOW"
            
            sector_strategies[sector] = {
                'performance': avg_change,
                'signal': signal,
                'strategy': strategy,
                'confidence': confidence
            }
        
        return sector_strategies
    
    except Exception as e:
        st.warning(f"Sector strategy analysis error: {e}")
        return {}

def extract_screener_opportunities(screener_df):
    """Extract opportunities from screener data"""
    opportunities = []
    
    try:
        # Look for key screening criteria
        symbol_cols = [col for col in screener_df.columns if any(term in str(col).upper() for term in ['SYMBOL', 'STOCK'])]
        if not symbol_cols:
            return opportunities
        
        symbol_col = symbol_cols[0]
        
        # Add all screened stocks as opportunities
        for _, row in screener_df.iterrows():
            symbol = row[symbol_col]
            opportunity = {'symbol': symbol}
            
            # Add all available metrics
            for col in screener_df.columns:
                if col != symbol_col:
                    opportunity[col] = row[col]
            
            opportunities.append(opportunity)
        
        return opportunities[:20]  # Top 20 opportunities
    
    except Exception as e:
        st.warning(f"Screener analysis error: {e}")
        return []

def display_professional_dashboard(data_dict):
    """Display professional trading dashboard"""
    
    st.markdown("""
    <div class="trading-header">
        <h1>Professional Trading Strategy Dashboard</h1>
        <p>Data-Driven Trading Decisions | Live Market Analysis</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Extract strategies from different sheets
    pcr_strategy = {}
    stock_strategies = {}
    sector_strategies = {}
    screener_opportunities = []
    
    if 'PCR & OI Chart' in data_dict:
        pcr_strategy = extract_pcr_strategy(data_dict['PCR & OI Chart'])
    
    if 'Dashboard' in data_dict or 'Stock Dashboard' in data_dict:
        dashboard_data = data_dict.get('Stock Dashboard', data_dict.get('Dashboard'))
        stock_strategies = extract_stock_strategies(dashboard_data)
    
    if 'Sector Dashboard' in data_dict:
        sector_strategies = extract_sector_strategies(data_dict['Sector Dashboard'])
    
    if 'Screener' in data_dict:
        screener_opportunities = extract_screener_opportunities(data_dict['Screener'])
    
    # PCR Strategy Section
    if pcr_strategy:
        st.header("Options Market Strategy")
        
        for pcr_name, pcr_data in pcr_strategy.items():
            signal = pcr_data['signal']
            
            if 'BULLISH' in signal:
                card_class = 'pcr-bullish'
            elif 'BEARISH' in signal:
                card_class = 'pcr-bearish'
            else:
                card_class = 'pcr-neutral'
            
            st.markdown(f"""
            <div class="strategy-card">
                <h4>{pcr_name} Strategy</h4>
                <div class="{card_class}">
                    <strong>Signal:</strong> {signal}<br>
                    <strong>PCR Value:</strong> {pcr_data['value']:.3f}<br>
                    <strong>Strategy:</strong> {pcr_data['strategy']}<br>
                    <strong>Confidence:</strong> {pcr_data['confidence']}
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    # Stock Strategy Section
    if stock_strategies:
        st.header("Stock Trading Strategies")
        
        strategy_tabs = st.tabs(['Long Candidates', 'Short Candidates', 'Breakout Stocks', 'High Volume'])
        
        with strategy_tabs[0]:
            long_candidates = stock_strategies.get('long_candidates', [])
            if long_candidates:
                st.subheader(f"Top {len(long_candidates)} Long Opportunities")
                for stock in long_candidates:
                    st.markdown(f"""
                    <div class="long-signal">
                        <strong>{stock['symbol']}</strong> | Change: +{stock['change']:.2f}%<br>
                        {stock['strategy']}<br>
                        Volume: {stock['volume']:,.0f}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No strong long candidates identified")
        
        with strategy_tabs[1]:
            short_candidates = stock_strategies.get('short_candidates', [])
            if short_candidates:
                st.subheader(f"Top {len(short_candidates)} Short Opportunities")
                for stock in short_candidates:
                    st.markdown(f"""
                    <div class="short-signal">
                        <strong>{stock['symbol']}</strong> | Change: {stock['change']:.2f}%<br>
                        {stock['strategy']}<br>
                        Volume: {stock['volume']:,.0f}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No strong short candidates identified")
        
        with strategy_tabs[2]:
            breakout_stocks = stock_strategies.get('breakout_stocks', [])
            if breakout_stocks:
                st.subheader("Breakout Opportunities")
                for stock in breakout_stocks:
                    st.markdown(f"""
                    <div class="neutral-signal">
                        <strong>{stock['symbol']}</strong> | Change: {stock['change']:+.2f}%<br>
                        {stock['strategy']}<br>
                        Volume: {stock['volume']:,.0f}
                    </div>
                    """, unsafe_allow_html=True)
        
        with strategy_tabs[3]:
            high_volume = stock_strategies.get('high_volume_plays', [])[:10]
            if high_volume:
                st.subheader("High Volume Stocks to Watch")
                for stock in high_volume:
                    st.write(f"**{stock['symbol']}** | Change: {stock['change']:+.2f}% | Volume: {stock['volume']:,.0f}")
    
    # Sector Strategy Section
    if sector_strategies:
        st.header("Sector Rotation Strategy")
        
        sector_cols = st.columns(2)
        sectors_list = list(sector_strategies.items())
        
        for i, (sector, data) in enumerate(sectors_list):
            col = sector_cols[i % 2]
            
            with col:
                signal = data['signal']
                
                if 'BUY' in signal:
                    card_class = 'long-signal'
                elif 'SELL' in signal:
                    card_class = 'short-signal'
                else:
                    card_class = 'neutral-signal'
                
                st.markdown(f"""
                <div class="{card_class}">
                    <h5>{sector}</h5>
                    Performance: {data['performance']:+.2f}%<br>
                    {data['strategy']}<br>
                    Confidence: {data['confidence']}
                </div>
                """, unsafe_allow_html=True)
    
    # Screener Opportunities
    if screener_opportunities:
        st.header("Screener Opportunities")
        
        with st.expander(f"View All {len(screener_opportunities)} Screened Stocks"):
            screener_df = pd.DataFrame(screener_opportunities)
            st.dataframe(screener_df, use_container_width=True)
    
    # Trading Summary
    st.header("Executive Trading Summary")
    
    summary_cols = st.columns(4)
    
    with summary_cols[0]:
        long_count = len(stock_strategies.get('long_candidates', []))
        st.metric("Long Opportunities", long_count)
    
    with summary_cols[1]:
        short_count = len(stock_strategies.get('short_candidates', []))
        st.metric("Short Opportunities", short_count)
    
    with summary_cols[2]:
        breakout_count = len(stock_strategies.get('breakout_stocks', []))
        st.metric("Breakout Plays", breakout_count)
    
    with summary_cols[3]:
        screener_count = len(screener_opportunities)
        st.metric("Screener Stocks", screener_count)

def main():
    uploaded_file = st.file_uploader("Upload Professional Trading Excel", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        with st.spinner("Processing professional trading data..."):
            data_dict = read_all_trading_sheets(temp_path)
        
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data_dict:
            display_professional_dashboard(data_dict)
        else:
            st.error("Could not process trading data")
    else:
        st.info("Upload your Excel file with professional trading sheets")
        
        st.markdown("""
        ### Expected Sheets:
        - **Dashboard/Stock Dashboard**: Stock performance data
        - **PCR & OI Chart**: Options market sentiment
        - **Sector Dashboard**: Sector performance metrics  
        - **Screener**: Pre-filtered trading opportunities
        - **Greeks**: Options Greeks analysis
        - **FII DII Data**: Institutional flow data
        """)

if __name__ == "__main__":
    main()
