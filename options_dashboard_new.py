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
}
.fo-section {
    background: #f8f9fa;
    border: 2px solid #007bff;
    border-radius: 10px;
    padding: 1.5rem;
    margin: 1rem 0;
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
.pcr-analysis {
    background: linear-gradient(135deg, #6f42c1, #e83e8c);
    color: white;
    padding: 1.5rem;
    border-radius: 10px;
    margin: 1rem 0;
    text-align: center;
}
.oi-analysis {
    background: #fff3cd;
    border: 1px solid #ffeaa7;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.5rem 0;
}
.unwinding-alert {
    background: #f8d7da;
    border: 1px solid #f5c6cb;
    color: #721c24;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
}
.buildup-alert {
    background: #d4edda;
    border: 1px solid #c3e6cb;
    color: #155724;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
}
.stock-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 1rem;
    margin: 1rem 0;
}
.action-card {
    border-left: 5px solid;
    padding: 1rem;
    margin: 0.5rem 0;
    border-radius: 5px;
}
.long-action {
    background-color: rgba(40, 167, 69, 0.1);
    border-left-color: #28a745;
}
.short-action {
    background-color: rgba(220, 53, 69, 0.1);
    border-left-color: #dc3545;
}
.index-card {
    background: linear-gradient(135deg, #343a40, #495057);
    color: white;
    padding: 1.5rem;
    border-radius: 10px;
    margin: 1rem 0;
}
.debug-panel {
    background-color: #f8f9fa;
    border: 1px solid #dee2e6;
    border-radius: 5px;
    padding: 1rem;
    margin-bottom: 1rem;
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

def find_column_by_keywords(columns, keywords):
    """Find a column that contains any of the given keywords"""
    for col in columns:
        col_upper = str(col).upper()
        for keyword in keywords:
            if keyword in col_upper:
                return col
    return None

@st.cache_data(ttl=30)  # Cache for 30 seconds for stable display
def read_comprehensive_data(file_path):
    """Read all relevant sheets for F&O analysis - Silent loading"""
    try:
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        # Read all available sheets silently
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                if not df.empty:
                    data_dict[sheet_name] = df
            except:
                continue
        
        return data_dict
    except:
        return {}

def extract_fo_bullish_bearish_stocks(data_dict):
    """Extract detailed F&O bullish and bearish stocks with trading actions"""
    fo_stocks = {'bullish': [], 'bearish': [], 'trading_actions': []}
    
    # Look for F&O specific sheets or columns
    for sheet_name, df in data_dict.items():
        if any(term in sheet_name.upper() for term in ['STOCK', 'DASHBOARD', 'F&O', 'DATA']):
            try:
                # Find relevant columns with more flexible matching
                symbol_col = find_column_by_keywords(df.columns, ['SYMBOL', 'STOCK', 'NAME', 'SCRIP'])
                change_col = find_column_by_keywords(df.columns, ['CHANGE', '%', 'CHG'])
                ltp_col = find_column_by_keywords(df.columns, ['LTP', 'PRICE', 'CLOSE', 'LAST'])
                oi_col = find_column_by_keywords(df.columns, ['OI', 'OPEN INT'])
                volume_col = find_column_by_keywords(df.columns, ['VOLUME', 'VOL'])
                buildup_col = find_column_by_keywords(df.columns, ['BUILDUP', 'TREND'])
                
                if not symbol_col or not change_col:
                    continue
                
                # Process each stock
                for _, row in df.iterrows():
                    try:
                        symbol = str(row[symbol_col]).strip()
                        change = pd.to_numeric(row[change_col], errors='coerce')
                        
                        if pd.isna(change) or symbol == 'nan' or symbol == '':
                            continue
                        
                        # Get other values if available
                        ltp = pd.to_numeric(row[ltp_col], errors='coerce') if ltp_col else 0
                        oi = pd.to_numeric(row[oi_col], errors='coerce') if oi_col else 0
                        volume = pd.to_numeric(row[volume_col], errors='coerce') if volume_col else 0
                        buildup = str(row[buildup_col]).strip() if buildup_col else 'Unknown'
                        
                        stock_data = {
                            'symbol': symbol,
                            'change': change,
                            'ltp': ltp,
                            'oi': oi,
                            'volume': volume,
                            'buildup': buildup
                        }
                        
                        # Classify as bullish or bearish
                        if change > 1:  # 1% threshold for bullish
                            fo_stocks['bullish'].append(stock_data)
                            
                            # Generate trading action for bullish stocks
                            action = {
                                'symbol': symbol,
                                'action': 'LONG',
                                'ltp': ltp,
                                'target': ltp * 1.05 if ltp > 0 else 0,
                                'stop_loss': ltp * 0.97 if ltp > 0 else 0,
                                'confidence': 'HIGH' if change > 3 else 'MEDIUM',
                                'reason': f"Bullish momentum (+{change:.2f}%) with {buildup} buildup"
                            }
                            fo_stocks['trading_actions'].append(action)
                            
                        elif change < -1:  # -1% threshold for bearish
                            fo_stocks['bearish'].append(stock_data)
                            
                            # Generate trading action for bearish stocks
                            action = {
                                'symbol': symbol,
                                'action': 'SHORT',
                                'ltp': ltp,
                                'target': ltp * 0.95 if ltp > 0 else 0,
                                'stop_loss': ltp * 1.03 if ltp > 0 else 0,
                                'confidence': 'HIGH' if change < -3 else 'MEDIUM',
                                'reason': f"Bearish momentum ({change:.2f}%) with {buildup} buildup"
                            }
                            fo_stocks['trading_actions'].append(action)
                    
                    except:
                        continue
            
            except:
                continue
    
    # Sort by change percentage
    fo_stocks['bullish'] = sorted(fo_stocks['bullish'], key=lambda x: x['change'], reverse=True)[:50]
    fo_stocks['bearish'] = sorted(fo_stocks['bearish'], key=lambda x: x['change'])[:50]
    
    # Sort trading actions by confidence and change magnitude
    fo_stocks['trading_actions'] = sorted(
        fo_stocks['trading_actions'], 
        key=lambda x: (x['confidence'] == 'HIGH', abs(x['ltp'] - (x['target'] if x['action'] == 'LONG' else x['stop_loss']))), 
        reverse=True
    )[:20]
    
    return fo_stocks

def extract_comprehensive_pcr_data(data_dict):
    """Extract PCR data based on actual data structure"""
    pcr_data = {}
    
    # Look for sheets with options data
    for sheet_name, df in data_dict.items():
        try:
            # Check if this sheet has the expected structure from your images
            if any(col in df.columns for col in ['SP', 'NET COI', 'NET OI', 'COI PCR', 'OI PCR']):
                # Extract NIFTY PCR if available
                nifty_rows = df[df.iloc[:, 0].astype(str).str.contains('NIFTY', na=False)]
                if not nifty_rows.empty:
                    row = nifty_rows.iloc[0]
                    if 'COI PCR' in df.columns:
                        coi_pcr = pd.to_numeric(row['COI PCR'], errors='coerce')
                        if not pd.isna(coi_pcr):
                            pcr_data['NIFTY_COI_PCR'] = {
                                'value': coi_pcr,
                                'interpretation': get_pcr_interpretation(coi_pcr)
                            }
                    
                    if 'OI PCR' in df.columns:
                        oi_pcr = pd.to_numeric(row['OI PCR'], errors='coerce')
                        if not pd.isna(oi_pcr):
                            pcr_data['NIFTY_OI_PCR'] = {
                                'value': oi_pcr,
                                'interpretation': get_pcr_interpretation(oi_pcr)
                            }
                
                # Extract BANKNIFTY PCR if available
                bnf_rows = df[df.iloc[:, 0].astype(str).str.contains('BANKNIFTY', na=False)]
                if not bnf_rows.empty:
                    row = bnf_rows.iloc[0]
                    if 'COI PCR' in df.columns:
                        coi_pcr = pd.to_numeric(row['COI PCR'], errors='coerce')
                        if not pd.isna(coi_pcr):
                            pcr_data['BANKNIFTY_COI_PCR'] = {
                                'value': coi_pcr,
                                'interpretation': get_pcr_interpretation(coi_pcr)
                            }
                    
                    if 'OI PCR' in df.columns:
                        oi_pcr = pd.to_numeric(row['OI PCR'], errors='coerce')
                        if not pd.isna(oi_pcr):
                            pcr_data['BANKNIFTY_OI_PCR'] = {
                                'value': oi_pcr,
                                'interpretation': get_pcr_interpretation(oi_pcr)
                            }
            
            # Also look for individual PCR columns in other sheets
            for col in df.columns:
                col_str = str(col).upper()
                if 'PCR' in col_str:
                    pcr_values = pd.to_numeric(df[col], errors='coerce').dropna()
                    if not pcr_values.empty:
                        current_pcr = pcr_values.iloc[-1]
                        pcr_data[f'{sheet_name}_{col}'] = {
                            'value': current_pcr,
                            'interpretation': get_pcr_interpretation(current_pcr)
                        }
                        
        except Exception as e:
            continue
    
    return pcr_data

def get_pcr_interpretation(pcr_value):
    """Interpret PCR values"""
    if pcr_value > 1.5:
        return {'signal': 'STRONG_BEARISH', 'action': 'Consider PUT buying', 'confidence': 'HIGH'}
    elif pcr_value > 1.2:
        return {'signal': 'BEARISH', 'action': 'Bearish bias', 'confidence': 'MEDIUM'}
    elif pcr_value < 0.6:
        return {'signal': 'STRONG_BULLISH', 'action': 'Consider CALL buying', 'confidence': 'HIGH'}
    elif pcr_value < 0.8:
        return {'signal': 'BULLISH', 'action': 'Bullish bias', 'confidence': 'MEDIUM'}
    else:
        return {'signal': 'NEUTRAL', 'action': 'Range-bound', 'confidence': 'LOW'}

def extract_high_oi_analysis(data_dict):
    """Extract high OI strikes and generate option strategies based on actual data structure"""
    high_oi_data = {'nifty': {}, 'banknifty': {}, 'strategies': []}
    
    for sheet_name, df in data_dict.items():
        try:
            # Check if this sheet has options data structure like in your images
            if 'Strike Price' in df.columns or any(col in df.columns for col in ['SP', 'NET COI', 'NET OI']):
                
                # Determine index type from the sheet or data
                index_type = 'nifty'
                if 'BANKNIFTY' in sheet_name.upper() or any('BANKNIFTY' in str(cell) for cell in df.values.flatten() if pd.notna(cell)):
                    index_type = 'banknifty'
                
                # Look for strike-wise data
                strike_data = []
                
                for _, row in df.iterrows():
                    try:
                        # Try to find strike price
                        strike = None
                        if 'Strike Price' in df.columns:
                            strike = pd.to_numeric(row['Strike Price'], errors='coerce')
                        elif 'SP' in df.columns:
                            strike = pd.to_numeric(row['SP'], errors='coerce')
                        
                        if pd.isna(strike):
                            continue
                        
                        # Get OI data
                        call_oi = 0
                        put_oi = 0
                        
                        if 'NET COI' in df.columns:
                            call_oi = pd.to_numeric(row['NET COI'], errors='coerce') or 0
                        if 'NET OI' in df.columns:
                            put_oi = pd.to_numeric(row['NET OI'], errors='coerce') or 0
                        
                        # Get PCR data
                        coi_pcr = pd.to_numeric(row.get('COI PCR', 0), errors='coerce') or 0
                        oi_pcr = pd.to_numeric(row.get('OI PCR', 0), errors='coerce') or 0
                        
                        strike_data.append({
                            'strike': strike,
                            'call_oi': abs(call_oi),  # Use absolute value
                            'put_oi': abs(put_oi),    # Use absolute value
                            'coi_pcr': coi_pcr,
                            'oi_pcr': oi_pcr
                        })
                        
                    except:
                        continue
                
                if strike_data:
                    # Sort by OI to find highest
                    max_call_oi_strike = max(strike_data, key=lambda x: x['call_oi'])
                    max_put_oi_strike = max(strike_data, key=lambda x: x['put_oi'])
                    
                    # Calculate total PCR
                    total_call_oi = sum(s['call_oi'] for s in strike_data)
                    total_put_oi = sum(s['put_oi'] for s in strike_data)
                    pcr = total_put_oi / total_call_oi if total_call_oi > 0 else 0
                    
                    # Store analysis
                    high_oi_data[index_type] = {
                        'max_call_strike': max_call_oi_strike['strike'],
                        'max_call_oi': max_call_oi_strike['call_oi'],
                        'max_put_strike': max_put_oi_strike['strike'],
                        'max_put_oi': max_put_oi_strike['put_oi'],
                        'pcr': pcr,
                        'total_call_oi': total_call_oi,
                        'total_put_oi': total_put_oi
                    }
                    
                    # Generate strategies - SAFE VERSION
                    try:
                        call_strike_str = str(int(max_call_oi_strike['strike']))
                        call_oi_str = f"{max_call_oi_strike['call_oi']:,.0f}"
                        
                        high_oi_data['strategies'].append({
                            'index': index_type.upper(),
                            'strategy': 'SELL CALL',
                            'strike': call_strike_str,
                            'premium': 'Market Price',
                            'oi_text': f"Call OI: {call_oi_str}",
                            'reason': f'Maximum Call OI at {call_strike_str} - Strong Resistance',
                            'risk': 'UNLIMITED',
                            'reward': 'PREMIUM RECEIVED'
                        })
                    except:
                        pass
                    
                    try:
                        put_strike_str = str(int(max_put_oi_strike['strike']))
                        put_oi_str = f"{max_put_oi_strike['put_oi']:,.0f}"
                        
                        high_oi_data['strategies'].append({
                            'index': index_type.upper(),
                            'strategy': 'SELL PUT',
                            'strike': put_strike_str,
                            'premium': 'Market Price',
                            'oi_text': f"Put OI: {put_oi_str}",
                            'reason': f'Maximum Put OI at {put_strike_str} - Strong Support',
                            'risk': 'UNLIMITED',
                            'reward': 'PREMIUM RECEIVED'
                        })
                    except:
                        pass
                    
                    # PCR-based strategies
                    try:
                        pcr_str = f"{pcr:.3f}"
                        if pcr > 1.3:  # Bearish
                            high_oi_data['strategies'].append({
                                'index': index_type.upper(),
                                'strategy': 'BUY PUT',
                                'strike': 'ATM',
                                'premium': 'Market Price',
                                'oi_text': f'PCR: {pcr_str} (Bearish)',
                                'reason': 'High PCR indicates bearish sentiment',
                                'risk': 'PREMIUM PAID',
                                'reward': 'UNLIMITED DOWNSIDE'
                            })
                        elif pcr < 0.7:  # Bullish
                            high_oi_data['strategies'].append({
                                'index': index_type.upper(),
                                'strategy': 'BUY CALL',
                                'strike': 'ATM',
                                'premium': 'Market Price',
                                'oi_text': f'PCR: {pcr_str} (Bullish)',
                                'reason': 'Low PCR indicates bullish sentiment',
                                'risk': 'PREMIUM PAID',
                                'reward': 'UNLIMITED UPSIDE'
                            })
                    except:
                        pass
                        
        except:
            continue
    
    return high_oi_data

def detect_options_unwinding(data_dict):
    """Detect call and put unwinding patterns"""
    unwinding_data = {'call_unwinding': [], 'put_unwinding': [], 'fresh_positions': []}
    
    for sheet_name, df in data_dict.items():
        if 'OC_' in sheet_name or 'OPTIONS' in sheet_name.upper() or 'NIFTY' in sheet_name.upper() or 'BANKNIFTY' in sheet_name.upper():
            try:
                # Look for OI change columns
                strike_col = find_column_by_keywords(df.columns, ['STRIKE'])
                call_oi_change_col = find_column_by_keywords(df.columns, ['CE', 'CALL', 'CHANGE', 'OI'])
                put_oi_change_col = find_column_by_keywords(df.columns, ['PE', 'PUT', 'CHANGE', 'OI'])
                call_price_col = find_column_by_keywords(df.columns, ['CE', 'CALL', 'LTP', 'PRICE'])
                put_price_col = find_column_by_keywords(df.columns, ['PE', 'PUT', 'LTP', 'PRICE'])
                
                if not strike_col:
                    continue
                
                # Process each row
                for _, row in df.iterrows():
                    try:
                        strike = pd.to_numeric(row[strike_col], errors='coerce')
                        if pd.isna(strike):
                            continue
                        
                        # Check for call unwinding
                        if call_oi_change_col and call_price_col:
                            call_oi_change = pd.to_numeric(row[call_oi_change_col], errors='coerce')
                            call_price = pd.to_numeric(row[call_price_col], errors='coerce')
                            
                            if not pd.isna(call_oi_change) and not pd.isna(call_price):
                                if call_oi_change < -10000 and call_price > 0:
                                    unwinding_data['call_unwinding'].append({
                                        'strike': strike,
                                        'oi_change': call_oi_change,
                                        'price': call_price,
                                        'sheet': sheet_name
                                    })
                                
                                if call_oi_change > 50000:
                                    unwinding_data['fresh_positions'].append({
                                        'type': 'CALL',
                                        'strike': strike,
                                        'oi_change': call_oi_change,
                                        'sheet': sheet_name
                                    })
                        
                        # Check for put unwinding
                        if put_oi_change_col and put_price_col:
                            put_oi_change = pd.to_numeric(row[put_oi_change_col], errors='coerce')
                            put_price = pd.to_numeric(row[put_price_col], errors='coerce')
                            
                            if not pd.isna(put_oi_change) and not pd.isna(put_price):
                                if put_oi_change < -10000 and put_price > 0:
                                    unwinding_data['put_unwinding'].append({
                                        'strike': strike,
                                        'oi_change': put_oi_change,
                                        'price': put_price,
                                        'sheet': sheet_name
                                    })
                                
                                if put_oi_change > 50000:
                                    unwinding_data['fresh_positions'].append({
                                        'type': 'PUT',
                                        'strike': strike,
                                        'oi_change': put_oi_change,
                                        'sheet': sheet_name
                                    })
                    
                    except:
                        continue
            
            except:
                continue
    
    # Sort by magnitude of change
    for key in unwinding_data:
        if key != 'fresh_positions':
            unwinding_data[key] = sorted(unwinding_data[key], key=lambda x: abs(x['oi_change']), reverse=True)[:10]
        else:
            unwinding_data[key] = sorted(unwinding_data[key], key=lambda x: x['oi_change'], reverse=True)[:10]
    
    return unwinding_data

def get_index_interpretation(pcr, spot_price, resistance_strike, support_strike):
    """Interpret index options data"""
    interpretation = {
        'signal': 'NEUTRAL',
        'action': 'Range-bound expected',
        'confidence': 'LOW'
    }
    
    if pcr > 1.5:
        interpretation['signal'] = 'BEARISH'
        interpretation['action'] = 'Consider PUT positions'
        interpretation['confidence'] = 'HIGH'
    elif pcr < 0.6:
        interpretation['signal'] = 'BULLISH'
        interpretation['action'] = 'Consider CALL positions'
        interpretation['confidence'] = 'HIGH'
    elif pcr > 1.2:
        interpretation['signal'] = 'WEAKLY_BEARISH'
        interpretation['confidence'] = 'MEDIUM'
    elif pcr < 0.8:
        interpretation['signal'] = 'WEAKLY_BULLISH'
        interpretation['confidence'] = 'MEDIUM'
    
    return interpretation

def display_live_dashboard(data_dict):
    """Display live F&O dashboard with auto-refresh"""
    
    # Live header
    st.markdown(f"""
    <div class="dashboard-header">
        <h1>📊 LIVE F&O Trading Dashboard</h1>
        <p class="live-indicator">● LIVE</p>
        <p>Real-time Analysis - {datetime.now().strftime("%H:%M:%S")}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Extract all data
    fo_stocks = extract_fo_bullish_bearish_stocks(data_dict)
    pcr_data = extract_comprehensive_pcr_data(data_dict)
    unwinding_data = detect_options_unwinding(data_dict)
    high_oi_data = extract_high_oi_analysis(data_dict)
    
    # PCR Analysis Section - Prominently displayed
    st.header("📊 PCR & Market Sentiment")
    
    pcr_items = [k for k in pcr_data.keys() if k != 'oi_analysis']
    
    if pcr_items:
        pcr_cols = st.columns(min(4, len(pcr_items)))
        for i, pcr_name in enumerate(pcr_items[:4]):
            pcr_info = pcr_data[pcr_name]
            interpretation = pcr_info['interpretation']
            
            with pcr_cols[i]:
                signal_color = "#dc3545" if "BEARISH" in interpretation['signal'] else "#28a745" if "BULLISH" in interpretation['signal'] else "#6c757d"
                st.markdown(f"""
                <div class="pcr-analysis" style="background: linear-gradient(135deg, {signal_color}, #495057);">
                    <h4>{pcr_name.replace('_', ' ')}</h4>
                    <h2>{pcr_info['value']:.3f}</h2>
                    <p><strong>{interpretation['signal']}</strong></p>
                    <p>{interpretation['action']}</p>
                    <small>Confidence: {interpretation['confidence']}</small>
                </div>
                """, unsafe_allow_html=True)
    
    # High OI Analysis & Option Strategies
    st.header("Option Strategies Based on High OI")
    
    if high_oi_data['strategies']:
        # Show strategies in cards similar to your layout
        for strategy in high_oi_data['strategies'][:6]:
            strategy_color = "#28a745" if "BUY" in strategy['strategy'] else "#dc3545" if "SELL" in strategy['strategy'] else "#6f42c1"
            st.markdown(f"""
            <div class="action-card" style="border-left-color: {strategy_color}; background-color: rgba(111, 66, 193, 0.1);">
                <strong>{strategy['index']} - {strategy['strategy']}</strong><br>
                Strike: {strategy['strike']} | Premium: {strategy['premium']}<br>
                {strategy['oi_text']}<br>
                <small>{strategy['reason']}</small><br>
                <small>Risk: {strategy['risk']} | Reward: {strategy['reward']}</small>
            </div>
            """, unsafe_allow_html=True)
    
    # High OI Levels Display in tabular format like your images
    if high_oi_data['nifty'] or high_oi_data['banknifty']:
        st.subheader("High OI Analysis")
        
        oi_cols = st.columns(2)
        
        with oi_cols[0]:
            if high_oi_data['nifty']:
                nifty_oi = high_oi_data['nifty']
                st.markdown(f"""
                <div class="index-card">
                    <h3>NIFTY Options Analysis</h3>
                    <table style="width:100%; color:white;">
                    <tr><td>Max Call OI Strike:</td><td><strong>{nifty_oi['max_call_strike']:.0f}</strong></td></tr>
                    <tr><td>Max Call OI:</td><td>{nifty_oi['max_call_oi']:,.0f}</td></tr>
                    <tr><td>Max Put OI Strike:</td><td><strong>{nifty_oi['
