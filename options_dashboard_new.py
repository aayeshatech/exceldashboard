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
    """Extract comprehensive PCR and OI data from all relevant sheets"""
    pcr_data = {}
    
    # Look in multiple possible sheets for PCR data
    pcr_sheets = ['PCR & OI Chart', 'Dashboard', 'Data', 'Globlemarket']
    
    for sheet_name in pcr_sheets:
        if sheet_name in data_dict:
            df = data_dict[sheet_name]
            
            try:
                # Extract PCR values with flexible matching
                for col in df.columns:
                    col_str = str(col).upper()
                    if 'PCR' in col_str or 'PUT CALL RATIO' in col_str:
                        pcr_values = pd.to_numeric(df[col], errors='coerce').dropna()
                        if not pcr_values.empty:
                            current_pcr = pcr_values.iloc[-1]
                            pcr_data[f'{sheet_name}_{col}'] = {
                                'value': current_pcr,
                                'interpretation': get_pcr_interpretation(current_pcr)
                            }
                
                # Extract NIFTY PCR if available
                nifty_pcr_col = find_column_by_keywords(df.columns, ['NIFTY', 'PCR'])
                if nifty_pcr_col:
                    pcr_values = pd.to_numeric(df[nifty_pcr_col], errors='coerce').dropna()
                    if not pcr_values.empty:
                        current_pcr = pcr_values.iloc[-1]
                        pcr_data['NIFTY_PCR'] = {
                            'value': current_pcr,
                            'interpretation': get_pcr_interpretation(current_pcr)
                        }
                
                # Extract BANKNIFTY PCR if available
                banknifty_pcr_col = find_column_by_keywords(df.columns, ['BANKNIFTY', 'BNF', 'PCR'])
                if banknifty_pcr_col:
                    pcr_values = pd.to_numeric(df[banknifty_pcr_col], errors='coerce').dropna()
                    if not pcr_values.empty:
                        current_pcr = pcr_values.iloc[-1]
                        pcr_data['BANKNIFTY_PCR'] = {
                            'value': current_pcr,
                            'interpretation': get_pcr_interpretation(current_pcr)
                        }
                
                # Extract OI data with flexible matching
                oi_data = {}
                for col in df.columns:
                    col_str = str(col).upper()
                    if ('OI' in col_str and 'PCR' not in col_str) or 'OPEN INTEREST' in col_str:
                        oi_values = pd.to_numeric(df[col], errors='coerce').dropna()
                        if not oi_values.empty:
                            current_oi = oi_values.iloc[-1]
                            previous_oi = oi_values.iloc[-2] if len(oi_values) > 1 else current_oi
                            oi_change = current_oi - previous_oi
                            
                            oi_data[col] = {
                                'current': current_oi,
                                'change': oi_change,
                                'change_pct': (oi_change / previous_oi * 100) if previous_oi != 0 else 0
                            }
                
                if oi_data:
                    pcr_data['oi_analysis'] = oi_data
                    
            except:
                pass
    
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

def extract_high_oi_analysis(data_dict):
    """Extract high OI strikes and generate option strategies"""
    high_oi_data = {'nifty': {}, 'banknifty': {}, 'strategies': []}
    
    for sheet_name, df in data_dict.items():
        if 'OC_' in sheet_name or any(term in sheet_name.upper() for term in ['NIFTY', 'BANKNIFTY', 'OPTIONS']):
            try:
                # Determine index type
                index_type = 'nifty' if 'NIFTY' in sheet_name.upper() else 'banknifty' if 'BANKNIFTY' in sheet_name.upper() else 'nifty'
                
                # Find relevant columns
                strike_col = find_column_by_keywords(df.columns, ['STRIKE'])
                call_oi_col = find_column_by_keywords(df.columns, ['CE OI', 'CALL OI', 'CE_OI'])
                put_oi_col = find_column_by_keywords(df.columns, ['PE OI', 'PUT OI', 'PE_OI'])
                call_ltp_col = find_column_by_keywords(df.columns, ['CE LTP', 'CALL LTP', 'CE_LTP'])
                put_ltp_col = find_column_by_keywords(df.columns, ['PE LTP', 'PUT LTP', 'PE_LTP'])
                call_oi_change_col = find_column_by_keywords(df.columns, ['CE CHG OI', 'CALL CHANGE OI', 'CE_CHG_OI'])
                put_oi_change_col = find_column_by_keywords(df.columns, ['PE CHG OI', 'PUT CHANGE OI', 'PE_CHG_OI'])
                
                if not all([strike_col, call_oi_col, put_oi_col]):
                    continue
                
                # Get spot price from data if available
                spot_price = None
                for col in df.columns:
                    if any(term in str(col).upper() for term in ['SPOT', 'UNDERLYING', 'INDEX PRICE']):
                        spot_values = pd.to_numeric(df[col], errors='coerce').dropna()
                        if not spot_values.empty:
                            spot_price = spot_values.iloc[0]
                            break
                
                # Calculate total OI and find high OI strikes
                df['call_oi_num'] = pd.to_numeric(df[call_oi_col], errors='coerce')
                df['put_oi_num'] = pd.to_numeric(df[put_oi_col], errors='coerce')
                df['strike_num'] = pd.to_numeric(df[strike_col], errors='coerce')
                
                # Remove NaN values
                valid_data = df.dropna(subset=['call_oi_num', 'put_oi_num', 'strike_num'])
                
                if valid_data.empty:
                    continue
                
                # Find max OI strikes
                max_call_oi_row = valid_data.loc[valid_data['call_oi_num'].idxmax()]
                max_put_oi_row = valid_data.loc[valid_data['put_oi_num'].idxmax()]
                
                # Calculate PCR
                total_call_oi = valid_data['call_oi_num'].sum()
                total_put_oi = valid_data['put_oi_num'].sum()
                pcr = total_put_oi / total_call_oi if total_call_oi > 0 else 0
                
                # Store high OI data
                high_oi_data[index_type] = {
                    'max_call_strike': max_call_oi_row['strike_num'],
                    'max_call_oi': max_call_oi_row['call_oi_num'],
                    'max_put_strike': max_put_oi_row['strike_num'],
                    'max_put_oi': max_put_oi_row['put_oi_num'],
                    'pcr': pcr,
                    'spot_price': spot_price,
                    'total_call_oi': total_call_oi,
                    'total_put_oi': total_put_oi
                }
                
                # Generate option strategies based on OI analysis
                if call_ltp_col and put_ltp_col:
                    max_call_ltp = pd.to_numeric(max_call_oi_row[call_ltp_col], errors='coerce')
                    max_put_ltp = pd.to_numeric(max_put_oi_row[put_ltp_col], errors='coerce')
                    
                    # Strategy 1: High Call OI - Resistance level
                    if not pd.isna(max_call_ltp) and max_call_ltp > 0:
                        strategy = {
                            'index': index_type.upper(),
                            'strategy': 'SELL CALL',
                            'strike': max_call_oi_row['strike_num'],
                            'premium': max_call_ltp,
                            'oi': max_call_oi_row['call_oi_num'],
                            'reason': f'High Call OI at {max_call_oi_row["strike_num"]} - Strong Resistance',
                            'risk': 'LIMITED PREMIUM',
                            'reward': f'₹{max_call_ltp:.2f} per lot'
                        }
                        high_oi_data['strategies'].append(strategy)
                    
                    # Strategy 2: High Put OI - Support level
                    if not pd.isna(max_put_ltp) and max_put_ltp > 0:
                        strategy = {
                            'index': index_type.upper(),
                            'strategy': 'SELL PUT',
                            'strike': max_put_oi_row['strike_num'],
                            'premium': max_put_ltp,
                            'oi': max_put_oi_row['put_oi_num'],
                            'reason': f'High Put OI at {max_put_oi_row["strike_num"]} - Strong Support',
                            'risk': 'LIMITED PREMIUM',
                            'reward': f'₹{max_put_ltp:.2f} per lot'
                        }
                        high_oi_data['strategies'].append(strategy)
                
                # Generate PCR-based strategies
                if pcr > 1.3:  # Bearish
                    high_oi_data['strategies'].append({
                        'index': index_type.upper(),
                        'strategy': 'BUY PUT / SELL CALL',
                        'strike': f'ATM ({spot_price:.0f})' if spot_price else 'ATM',
                        'premium': 'Market Price',
                        'oi': f'PCR: {pcr:.3f}',
                        'reason': f'High PCR ({pcr:.3f}) indicates bearish sentiment',
                        'risk': 'PREMIUM PAID',
                        'reward': 'UNLIMITED DOWNSIDE'
                    })
                elif pcr < 0.7:  # Bullish
                    high_oi_data['strategies'].append({
                        'index': index_type.upper(),
                        'strategy': 'BUY CALL / SELL PUT',
                        'strike': f'ATM ({spot_price:.0f})' if spot_price else 'ATM',
                        'premium': 'Market Price',
                        'oi': f'PCR: {pcr:.3f}',
                        'reason': f'Low PCR ({pcr:.3f}) indicates bullish sentiment',
                        'risk': 'PREMIUM PAID',
                        'reward': 'UNLIMITED UPSIDE'
                    })
                
            except Exception as e:
                continue
    
    return high_oi_data
    """Extract Nifty and BankNifty options data for analysis"""
    index_data = {'nifty': {}, 'banknifty': {}}
    
    for sheet_name, df in data_dict.items():
        # Check if this sheet contains Nifty or BankNifty data
        index_type = None
        if 'NIFTY' in sheet_name.upper():
            index_type = 'nifty'
        elif 'BANKNIFTY' in sheet_name.upper():
            index_type = 'banknifty'
        
        if index_type is None:
            continue
        
        try:
            # Find relevant columns
            strike_col = find_column_by_keywords(df.columns, ['STRIKE'])
            call_oi_col = find_column_by_keywords(df.columns, ['CE', 'CALL', 'OI'])
            put_oi_col = find_column_by_keywords(df.columns, ['PE', 'PUT', 'OI'])
            call_oi_change_col = find_column_by_keywords(df.columns, ['CE', 'CALL', 'CHANGE', 'OI'])
            put_oi_change_col = find_column_by_keywords(df.columns, ['PE', 'PUT', 'CHANGE', 'OI'])
            call_price_col = find_column_by_keywords(df.columns, ['CE', 'CALL', 'LTP', 'PRICE'])
            put_price_col = find_column_by_keywords(df.columns, ['PE', 'PUT', 'LTP', 'PRICE'])
            
            if not strike_col or not call_oi_col or not put_oi_col:
                continue
            
            # Calculate PCR
            total_call_oi = pd.to_numeric(df[call_oi_col], errors='coerce').sum()
            total_put_oi = pd.to_numeric(df[put_oi_col], errors='coerce').sum()
            pcr = total_put_oi / total_call_oi if total_call_oi > 0 else 0
            
            # Find max OI strikes
            max_call_oi_idx = pd.to_numeric(df[call_oi_col], errors='coerce').idxmax()
            max_put_oi_idx = pd.to_numeric(df[put_oi_col], errors='coerce').idxmax()
            
            resistance_strike = df.iloc[max_call_oi_idx][strike_col] if not pd.isna(max_call_oi_idx) else None
            support_strike = df.iloc[max_put_oi_idx][strike_col] if not pd.isna(max_put_oi_idx) else None
            
            # Store index data
            index_data[index_type] = {
                'pcr': pcr,
                'resistance_strike': resistance_strike,
                'support_strike': support_strike,
                'interpretation': get_index_interpretation(pcr, None, resistance_strike, support_strike)
            }
            
        except:
            continue
    
    return index_data

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
    st.header("🎯 Option Strategies Based on High OI")
    
    if high_oi_data['strategies']:
        strategy_cols = st.columns(2)
        
        for i, strategy in enumerate(high_oi_data['strategies'][:6]):
            col = strategy_cols[i % 2]
            with col:
                strategy_color = "#28a745" if "BUY" in strategy['strategy'] else "#dc3545" if "SELL" in strategy['strategy'] else "#6f42c1"
                st.markdown(f"""
                <div class="action-card" style="border-left-color: {strategy_color}; background-color: rgba(111, 66, 193, 0.1);">
                    <strong>{strategy['index']} - {strategy['strategy']}</strong><br>
                    Strike: {strategy['strike']} | Premium: {strategy['premium']}<br>
                    OI: {strategy['oi']:,.0f} (if numeric)<br>
                    <small>{strategy['reason']}</small><br>
                    <small>Risk: {strategy['risk']} | Reward: {strategy['reward']}</small>
                </div>
                """, unsafe_allow_html=True)
    
    # High OI Levels Display
    if high_oi_data['nifty'] or high_oi_data['banknifty']:
        st.subheader("🔥 High OI Levels")
        
        oi_cols = st.columns(2)
        
        with oi_cols[0]:
            if high_oi_data['nifty']:
                nifty_oi = high_oi_data['nifty']
                st.markdown(f"""
                <div class="index-card">
                    <h3>NIFTY High OI</h3>
                    <p>Max Call OI: {nifty_oi['max_call_strike']:.0f} ({nifty_oi['max_call_oi']:,.0f})</p>
                    <p>Max Put OI: {nifty_oi['max_put_strike']:.0f} ({nifty_oi['max_put_oi']:,.0f})</p>
                    <p>PCR: {nifty_oi['pcr']:.3f}</p>
                    <p>Spot: {nifty_oi['spot_price']:.2f} (if available)</p>
                </div>
                """, unsafe_allow_html=True)
        
        with oi_cols[1]:
            if high_oi_data['banknifty']:
                bnf_oi = high_oi_data['banknifty']
                st.markdown(f"""
                <div class="index-card">
                    <h3>BANKNIFTY High OI</h3>
                    <p>Max Call OI: {bnf_oi['max_call_strike']:.0f} ({bnf_oi['max_call_oi']:,.0f})</p>
                    <p>Max Put OI: {bnf_oi['max_put_strike']:.0f} ({bnf_oi['max_put_oi']:,.0f})</p>
                    <p>PCR: {bnf_oi['pcr']:.3f}</p>
                    <p>Spot: {bnf_oi['spot_price']:.2f} (if available)</p>
                </div>
                """, unsafe_allow_html=True)
    
    # Trading Actions Section
    st.header("🔥 Active Trading Signals")
    
    long_actions = [action for action in fo_stocks['trading_actions'] if action['action'] == 'LONG']
    short_actions = [action for action in fo_stocks['trading_actions'] if action['action'] == 'SHORT']
    
    action_cols = st.columns(2)
    
    with action_cols[0]:
        st.subheader("🟢 LONG Positions")
        if long_actions:
            for action in long_actions[:8]:
                st.markdown(f"""
                <div class="action-card long-action">
                    <strong>{action['symbol']}</strong> - LONG<br>
                    Entry: ₹{action['ltp']:.2f} | Target: ₹{action['target']:.2f} | SL: ₹{action['stop_loss']:.2f}<br>
                    <small>Confidence: {action['confidence']}</small>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("No LONG signals active")
    
    with action_cols[1]:
        st.subheader("🔴 SHORT Positions")
        if short_actions:
            for action in short_actions[:8]:
                st.markdown(f"""
                <div class="action-card short-action">
                    <strong>{action['symbol']}</strong> - SHORT<br>
                    Entry: ₹{action['ltp']:.2f} | Target: ₹{action['target']:.2f} | SL: ₹{action['stop_loss']:.2f}<br>
                    <small>Confidence: {action['confidence']}</small>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("No SHORT signals active")
    
    # Market Overview
    st.header("📊 Market Overview")
    
    overview_cols = st.columns(4)
    
    with overview_cols[0]:
        st.metric("Bullish Stocks", len(fo_stocks['bullish']))
    
    with overview_cols[1]:
        st.metric("Bearish Stocks", len(fo_stocks['bearish']))
    
    with overview_cols[2]:
        total_unwinding = len(unwinding_data['call_unwinding']) + len(unwinding_data['put_unwinding'])
        st.metric("Unwinding Strikes", total_unwinding)
    
    with overview_cols[3]:
        st.metric("Fresh OI Buildup", len(unwinding_data['fresh_positions']))
    
    # Stock Performance
    if fo_stocks['bullish'] or fo_stocks['bearish']:
        st.header("📈 Top Performers")
        
        perf_cols = st.columns(2)
        
        with perf_cols[0]:
            if fo_stocks['bullish']:
                st.subheader("🟢 Top Gainers")
                for stock in fo_stocks['bullish'][:10]:
                    st.markdown(f"""
                    <div class="bullish-stock">
                        <div><strong>{stock['symbol']}</strong><br><small>₹{stock['ltp']:.2f}</small></div>
                        <div><h4>+{stock['change']:.2f}%</h4></div>
                    </div>
                    """, unsafe_allow_html=True)
        
        with perf_cols[1]:
            if fo_stocks['bearish']:
                st.subheader("🔴 Top Losers")
                for stock in fo_stocks['bearish'][:10]:
                    st.markdown(f"""
                    <div class="bearish-stock">
                        <div><strong>{stock['symbol']}</strong><br><small>₹{stock['ltp']:.2f}</small></div>
                        <div><h4>{stock['change']:.2f}%</h4></div>
                    </div>
                    """, unsafe_allow_html=True)
    
    # Options Activity
    if unwinding_data['call_unwinding'] or unwinding_data['put_unwinding']:
        st.header("🔄 Options Activity")
        
        unwind_cols = st.columns(2)
        
        with unwind_cols[0]:
            if unwinding_data['call_unwinding']:
                st.subheader("Call Unwinding")
                for unwind in unwinding_data['call_unwinding'][:5]:
                    st.markdown(f"""
                    <div class="unwinding-alert">
                        <strong>Strike {unwind['strike']:.0f}</strong><br>
                        OI Change: {unwind['oi_change']:,.0f}
                    </div>
                    """, unsafe_allow_html=True)
        
        with unwind_cols[1]:
            if unwinding_data['put_unwinding']:
                st.subheader("Put Unwinding")
                for unwind in unwinding_data['put_unwinding'][:5]:
                    st.markdown(f"""
                    <div class="unwinding-alert">
                        <strong>Strike {unwind['strike']:.0f}</strong><br>
                        OI Change: {unwind['oi_change']:,.0f}
                    </div>
                    """, unsafe_allow_html=True)
        
        with perf_cols[0]:
            if fo_stocks['bullish']:
                st.subheader("🟢 Top Gainers")
                for stock in fo_stocks['bullish'][:10]:
                    st.markdown(f"""
                    <div class="bullish-stock">
                        <div><strong>{stock['symbol']}</strong><br><small>₹{stock['ltp']:.2f}</small></div>
                        <div><h4>+{stock['change']:.2f}%</h4></div>
                    </div>
                    """, unsafe_allow_html=True)
        
        with perf_cols[1]:
            if fo_stocks['bearish']:
                st.subheader("🔴 Top Losers")
                for stock in fo_stocks['bearish'][:10]:
                    st.markdown(f"""
                    <div class="bearish-stock">
                        <div><strong>{stock['symbol']}</strong><br><small>₹{stock['ltp']:.2f}</small></div>
                        <div><h4>{stock['change']:.2f}%</h4></div>
                    </div>
                    """, unsafe_allow_html=True)
    
    # Options Activity
    if unwinding_data['call_unwinding'] or unwinding_data['put_unwinding']:
        st.header("🔄 Options Activity")
        
        unwind_cols = st.columns(2)
        
        with unwind_cols[0]:
            if unwinding_data['call_unwinding']:
                st.subheader("Call Unwinding")
                for unwind in unwinding_data['call_unwinding'][:5]:
                    st.markdown(f"""
                    <div class="unwinding-alert">
                        <strong>Strike {unwind['strike']:.0f}</strong><br>
                        OI Change: {unwind['oi_change']:,.0f}
                    </div>
                    """, unsafe_allow_html=True)
        
        with unwind_cols[1]:
            if unwinding_data['put_unwinding']:
                st.subheader("Put Unwinding")
                for unwind in unwinding_data['put_unwinding'][:5]:
                    st.markdown(f"""
                    <div class="unwinding-alert">
                        <strong>Strike {unwind['strike']:.0f}</strong><br>
                        OI Change: {unwind['oi_change']:,.0f}
                    </div>
                    """, unsafe_allow_html=True)

def main():
    # Auto-refresh every 5 seconds
    st.sidebar.button("🔄 Refresh", key="refresh")
    
    uploaded_file = st.file_uploader("Upload F&O Excel Data", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Silent data loading
        data_dict = read_comprehensive_data(temp_path)
        
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data_dict:
            display_live_dashboard(data_dict)
            
            # Auto refresh every 5 seconds
            st.rerun()
        else:
            st.error("Unable to process data")
    
    else:
        st.info("📁 Upload your F&O Excel file to start live analysis")

if __name__ == "__main__":
    main()
