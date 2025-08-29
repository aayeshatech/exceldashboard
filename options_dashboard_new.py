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
.no-data-message {
    color: #6c757d;
    font-style: italic;
    padding: 1rem;
    background-color: #f8f9fa;
    border-radius: 5px;
    margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

def read_comprehensive_data(file_path):
    """Read all relevant sheets for F&O analysis"""
    try:
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        # Target sheets for comprehensive F&O analysis
        target_sheets = [
            'Dashboard', 'Stock Dashboard', 'PCR & OI Chart', 'Sector Dashboard',
            'OC_1', 'OC_2', 'OC_3', 'Greeks', 'Screener', 'FII DII Data',
            'NIFTY', 'BANKNIFTY', 'INDEX', 'OPTIONS'
        ]
        
        # Read all available sheets
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                if not df.empty:
                    data_dict[sheet_name] = df
                    st.info(f"Loaded {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
            except Exception as e:
                st.warning(f"Could not read {sheet_name}: {str(e)}")
        
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return {}

def extract_fo_bullish_bearish_stocks(data_dict):
    """Extract detailed F&O bullish and bearish stocks with trading actions"""
    fo_stocks = {'bullish': [], 'bearish': [], 'trading_actions': []}
    
    # Look for F&O specific sheets or columns
    for sheet_name, df in data_dict.items():
        if any(term in sheet_name.upper() for term in ['STOCK', 'DASHBOARD', 'F&O']):
            try:
                # Find relevant columns
                symbol_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['SYMBOL', 'STOCK', 'NAME'])]
                change_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['CHANGE', '%'])]
                ltp_cols = [col for col in df.columns if 'LTP' in str(col).upper()]
                oi_cols = [col for col in df.columns if 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
                volume_cols = [col for col in df.columns if 'VOLUME' in str(col).upper()]
                buildup_cols = [col for col in df.columns if 'BUILDUP' in str(col).upper()]
                
                if not symbol_cols or not change_cols:
                    continue
                
                symbol_col = symbol_cols[0]
                change_col = change_cols[0]
                
                # Process each stock
                for _, row in df.iterrows():
                    try:
                        symbol = str(row[symbol_col]).strip()
                        change = pd.to_numeric(row[change_col], errors='coerce')
                        
                        if pd.isna(change) or symbol == 'nan':
                            continue
                        
                        stock_data = {
                            'symbol': symbol,
                            'change': change,
                            'ltp': pd.to_numeric(row[ltp_cols[0]] if ltp_cols else 0, errors='coerce'),
                            'oi': pd.to_numeric(row[oi_cols[0]] if oi_cols else 0, errors='coerce'),
                            'volume': pd.to_numeric(row[volume_cols[0]] if volume_cols else 0, errors='coerce'),
                            'buildup': str(row[buildup_cols[0]] if buildup_cols else 'Unknown').strip()
                        }
                        
                        # Classify as bullish or bearish
                        if change > 0.5:  # Bullish threshold
                            fo_stocks['bullish'].append(stock_data)
                            
                            # Generate trading action for bullish stocks
                            action = {
                                'symbol': symbol,
                                'action': 'LONG',
                                'ltp': stock_data['ltp'],
                                'target': stock_data['ltp'] * 1.05 if stock_data['ltp'] > 0 else 0,
                                'stop_loss': stock_data['ltp'] * 0.97 if stock_data['ltp'] > 0 else 0,
                                'confidence': 'HIGH' if change > 2 else 'MEDIUM',
                                'reason': f"Bullish momentum (+{change:.2f}%) with {stock_data['buildup']} buildup"
                            }
                            fo_stocks['trading_actions'].append(action)
                            
                        elif change < -0.5:  # Bearish threshold
                            fo_stocks['bearish'].append(stock_data)
                            
                            # Generate trading action for bearish stocks
                            action = {
                                'symbol': symbol,
                                'action': 'SHORT',
                                'ltp': stock_data['ltp'],
                                'target': stock_data['ltp'] * 0.95 if stock_data['ltp'] > 0 else 0,
                                'stop_loss': stock_data['ltp'] * 1.03 if stock_data['ltp'] > 0 else 0,
                                'confidence': 'HIGH' if change < -2 else 'MEDIUM',
                                'reason': f"Bearish momentum ({change:.2f}%) with {stock_data['buildup']} buildup"
                            }
                            fo_stocks['trading_actions'].append(action)
                    
                    except Exception as e:
                        continue
            
            except Exception as e:
                st.warning(f"Error processing {sheet_name}: {str(e)}")
    
    # Sort by change percentage
    fo_stocks['bullish'] = sorted(fo_stocks['bullish'], key=lambda x: x['change'], reverse=True)[:50]
    fo_stocks['bearish'] = sorted(fo_stocks['bearish'], key=lambda x: x['change'])[:50]
    
    # Sort trading actions by confidence and then by change magnitude
    fo_stocks['trading_actions'] = sorted(
        fo_stocks['trading_actions'], 
        key=lambda x: (x['confidence'] == 'HIGH', abs(x['ltp'] - (x['target'] if x['action'] == 'LONG' else x['stop_loss']))), 
        reverse=True
    )[:20]
    
    return fo_stocks

def extract_comprehensive_pcr_data(data_dict):
    """Extract comprehensive PCR and OI data"""
    pcr_data = {}
    
    if 'PCR & OI Chart' in data_dict:
        df = data_dict['PCR & OI Chart']
        
        try:
            # Extract PCR values
            for col in df.columns:
                if 'PCR' in str(col).upper():
                    pcr_values = pd.to_numeric(df[col], errors='coerce').dropna()
                    if not pcr_values.empty:
                        current_pcr = pcr_values.iloc[-1]
                        pcr_data[col] = {
                            'value': current_pcr,
                            'interpretation': get_pcr_interpretation(current_pcr)
                        }
            
            # Extract OI data
            oi_data = {}
            for col in df.columns:
                if 'OI' in str(col).upper() and 'PCR' not in str(col).upper():
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
            
            if oi_data:  # Only add oi_analysis if there's data
                pcr_data['oi_analysis'] = oi_data
            
        except Exception as e:
            st.warning(f"Error extracting PCR data: {str(e)}")
    
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
                strike_cols = [col for col in df.columns if 'STRIKE' in str(col).upper()]
                call_oi_change_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'CHANGE' in str(col).upper() and 'OI' in str(col).upper()]
                put_oi_change_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'CHANGE' in str(col).upper() and 'OI' in str(col).upper()]
                call_price_cols = [col for col in df.columns if 'CE' in str(col).upper() and ('LTP' in str(col).upper() or 'PRICE' in str(col).upper())]
                put_price_cols = [col for col in df.columns if 'PE' in str(col).upper() and ('LTP' in str(col).upper() or 'PRICE' in str(col).upper())]
                
                if not strike_cols:
                    continue
                
                strike_col = strike_cols[0]
                
                for _, row in df.iterrows():
                    strike = pd.to_numeric(row[strike_col], errors='coerce')
                    if pd.isna(strike):
                        continue
                    
                    # Check for call unwinding (OI decrease + price decrease)
                    if call_oi_change_cols and call_price_cols:
                        try:
                            call_oi_change = pd.to_numeric(row[call_oi_change_cols[0]], errors='coerce')
                            call_price = pd.to_numeric(row[call_price_cols[0]], errors='coerce')
                            
                            if call_oi_change < -10000 and call_price > 0:  # Significant OI decrease
                                unwinding_data['call_unwinding'].append({
                                    'strike': strike,
                                    'oi_change': call_oi_change,
                                    'price': call_price,
                                    'sheet': sheet_name
                                })
                        except:
                            pass
                    
                    # Check for put unwinding (OI decrease + price decrease)
                    if put_oi_change_cols and put_price_cols:
                        try:
                            put_oi_change = pd.to_numeric(row[put_oi_change_cols[0]], errors='coerce')
                            put_price = pd.to_numeric(row[put_price_cols[0]], errors='coerce')
                            
                            if put_oi_change < -10000 and put_price > 0:  # Significant OI decrease
                                unwinding_data['put_unwinding'].append({
                                    'strike': strike,
                                    'oi_change': put_oi_change,
                                    'price': put_price,
                                    'sheet': sheet_name
                                })
                        except:
                            pass
                    
                    # Check for fresh positions (OI increase)
                    if call_oi_change_cols:
                        try:
                            call_oi_change = pd.to_numeric(row[call_oi_change_cols[0]], errors='coerce')
                            if call_oi_change > 50000:  # Significant OI increase
                                unwinding_data['fresh_positions'].append({
                                    'type': 'CALL',
                                    'strike': strike,
                                    'oi_change': call_oi_change,
                                    'sheet': sheet_name
                                })
                        except:
                            pass
                    
                    if put_oi_change_cols:
                        try:
                            put_oi_change = pd.to_numeric(row[put_oi_change_cols[0]], errors='coerce')
                            if put_oi_change > 50000:  # Significant OI increase
                                unwinding_data['fresh_positions'].append({
                                    'type': 'PUT',
                                    'strike': strike,
                                    'oi_change': put_oi_change,
                                    'sheet': sheet_name
                                })
                        except:
                            pass
            
            except Exception as e:
                st.warning(f"Error analyzing unwinding in {sheet_name}: {str(e)}")
    
    # Sort by magnitude of change
    for key in unwinding_data:
        if key != 'fresh_positions':
            unwinding_data[key] = sorted(unwinding_data[key], key=lambda x: abs(x['oi_change']), reverse=True)[:10]
        else:
            unwinding_data[key] = sorted(unwinding_data[key], key=lambda x: x['oi_change'], reverse=True)[:10]
    
    return unwinding_data

def extract_index_options_data(data_dict):
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
            strike_cols = [col for col in df.columns if 'STRIKE' in str(col).upper()]
            call_oi_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
            put_oi_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
            call_oi_change_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'CHANGE' in str(col).upper() and 'OI' in str(col).upper()]
            put_oi_change_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'CHANGE' in str(col).upper() and 'OI' in str(col).upper()]
            call_price_cols = [col for col in df.columns if 'CE' in str(col).upper() and ('LTP' in str(col).upper() or 'PRICE' in str(col).upper())]
            put_price_cols = [col for col in df.columns if 'PE' in str(col).upper() and ('LTP' in str(col).upper() or 'PRICE' in str(col).upper())]
            
            if not strike_cols or not call_oi_cols or not put_oi_cols:
                continue
            
            strike_col = strike_cols[0]
            call_oi_col = call_oi_cols[0]
            put_oi_col = put_oi_cols[0]
            
            # Extract current spot price if available
            spot_price = None
            spot_cols = [col for col in df.columns if 'SPOT' in str(col).upper() or 'UNDERLYING' in str(col).upper() or 'INDEX' in str(col).upper()]
            if spot_cols:
                spot_values = pd.to_numeric(df[spot_cols[0]], errors='coerce').dropna()
                if not spot_values.empty:
                    spot_price = spot_values.iloc[-1]
            
            # Calculate PCR
            total_call_oi = pd.to_numeric(df[call_oi_col], errors='coerce').sum()
            total_put_oi = pd.to_numeric(df[put_oi_col], errors='coerce').sum()
            pcr = total_put_oi / total_call_oi if total_call_oi > 0 else 0
            
            # Find max OI strikes for support/resistance
            max_call_oi_idx = pd.to_numeric(df[call_oi_col], errors='coerce').idxmax()
            max_put_oi_idx = pd.to_numeric(df[put_oi_col], errors='coerce').idxmax()
            
            resistance_strike = df.iloc[max_call_oi_idx][strike_col] if not pd.isna(max_call_oi_idx) else None
            support_strike = df.iloc[max_put_oi_idx][strike_col] if not pd.isna(max_put_oi_idx) else None
            
            # Find significant OI changes
            significant_changes = []
            
            for _, row in df.iterrows():
                strike = row[strike_col]
                
                # Check call OI changes
                if call_oi_change_cols:
                    call_oi_change = pd.to_numeric(row[call_oi_change_cols[0]], errors='coerce')
                    if abs(call_oi_change) > 50000:  # Significant change threshold
                        significant_changes.append({
                            'strike': strike,
                            'type': 'CALL',
                            'oi_change': call_oi_change,
                            'oi': pd.to_numeric(row[call_oi_col], errors='coerce'),
                            'price': pd.to_numeric(row[call_price_cols[0]], errors='coerce') if call_price_cols else 0
                        })
                
                # Check put OI changes
                if put_oi_change_cols:
                    put_oi_change = pd.to_numeric(row[put_oi_change_cols[0]], errors='coerce')
                    if abs(put_oi_change) > 50000:  # Significant change threshold
                        significant_changes.append({
                            'strike': strike,
                            'type': 'PUT',
                            'oi_change': put_oi_change,
                            'oi': pd.to_numeric(row[put_oi_col], errors='coerce'),
                            'price': pd.to_numeric(row[put_price_cols[0]], errors='coerce') if put_price_cols else 0
                        })
            
            # Sort significant changes by absolute value
            significant_changes = sorted(significant_changes, key=lambda x: abs(x['oi_change']), reverse=True)[:10]
            
            # Store index data
            index_data[index_type] = {
                'spot_price': spot_price,
                'pcr': pcr,
                'resistance_strike': resistance_strike,
                'support_strike': support_strike,
                'significant_changes': significant_changes,
                'interpretation': get_index_interpretation(pcr, spot_price, resistance_strike, support_strike)
            }
            
        except Exception as e:
            st.warning(f"Error analyzing {sheet_name}: {str(e)}")
    
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
        interpretation['action'] = 'Consider PUT positions or shorting'
        interpretation['confidence'] = 'HIGH'
    elif pcr < 0.6:
        interpretation['signal'] = 'BULLISH'
        interpretation['action'] = 'Consider CALL positions or buying'
        interpretation['confidence'] = 'HIGH'
    elif pcr > 1.2:
        interpretation['signal'] = 'WEAKLY_BEARISH'
        interpretation['action'] = 'Bias towards PUT positions'
        interpretation['confidence'] = 'MEDIUM'
    elif pcr < 0.8:
        interpretation['signal'] = 'WEAKLY_BULLISH'
        interpretation['action'] = 'Bias towards CALL positions'
        interpretation['confidence'] = 'MEDIUM'
    
    # Add support/resistance analysis
    if spot_price and resistance_strike and support_strike:
        if spot_price > resistance_strike:
            interpretation['action'] += f". Breakout above {resistance_strike} resistance"
            interpretation['confidence'] = 'HIGH' if interpretation['confidence'] != 'LOW' else 'MEDIUM'
        elif spot_price < support_strike:
            interpretation['action'] += f". Breakdown below {support_strike} support"
            interpretation['confidence'] = 'HIGH' if interpretation['confidence'] != 'LOW' else 'MEDIUM'
        else:
            interpretation['action'] += f". Range between {support_strike} and {resistance_strike}"
    
    return interpretation

def display_comprehensive_dashboard(data_dict):
    """Display comprehensive F&O dashboard"""
    
    st.markdown("""
    <div class="dashboard-header">
        <h1>Comprehensive F&O Trading Dashboard</h1>
        <p>Live Analysis with PCR, OI Unwinding & Stock Alerts</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Extract all data
    fo_stocks = extract_fo_bullish_bearish_stocks(data_dict)
    pcr_data = extract_comprehensive_pcr_data(data_dict)
    unwinding_data = detect_options_unwinding(data_dict)
    index_data = extract_index_options_data(data_dict)
    
    # Trading Actions Section - Prominently displayed at the top
    st.header("🔥 Top Trading Actions")
    
    # Create two columns for LONG and SHORT actions
    long_actions = [action for action in fo_stocks['trading_actions'] if action['action'] == 'LONG']
    short_actions = [action for action in fo_stocks['trading_actions'] if action['action'] == 'SHORT']
    
    action_cols = st.columns(2)
    
    with action_cols[0]:
        st.subheader("🟢 LONG Positions")
        if long_actions:
            for action in long_actions[:5]:  # Show top 5 long positions
                st.markdown(f"""
                <div class="action-card long-action">
                    <strong>{action['symbol']}</strong> - LONG<br>
                    Entry: ₹{action['ltp']:.2f} | Target: ₹{action['target']:.2f} | SL: ₹{action['stop_loss']:.2f}<br>
                    <small>{action['reason']}</small><br>
                    <small>Confidence: {action['confidence']}</small>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No LONG positions identified at this time</div>', unsafe_allow_html=True)
    
    with action_cols[1]:
        st.subheader("🔴 SHORT Positions")
        if short_actions:
            for action in short_actions[:5]:  # Show top 5 short positions
                st.markdown(f"""
                <div class="action-card short-action">
                    <strong>{action['symbol']}</strong> - SHORT<br>
                    Entry: ₹{action['ltp']:.2f} | Target: ₹{action['target']:.2f} | SL: ₹{action['stop_loss']:.2f}<br>
                    <small>{action['reason']}</small><br>
                    <small>Confidence: {action['confidence']}</small>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No SHORT positions identified at this time</div>', unsafe_allow_html=True)
    
    # Nifty and BankNifty Analysis Section
    st.header("📈 Index Options Analysis")
    
    index_cols = st.columns(2)
    
    # Nifty Analysis
    with index_cols[0]:
        if index_data['nifty']:
            nifty = index_data['nifty']
            interpretation = nifty['interpretation']
            
            st.markdown(f"""
            <div class="index-card">
                <h3>NIFTY Analysis</h3>
                <p>Spot Price: {nifty['spot_price'] if nifty['spot_price'] else 'N/A'}</p>
                <p>PCR: {nifty['pcr']:.2f}</p>
                <p>Resistance: {nifty['resistance_strike'] if nifty['resistance_strike'] else 'N/A'}</p>
                <p>Support: {nifty['support_strike'] if nifty['support_strike'] else 'N/A'}</p>
                <p><strong>Signal: {interpretation['signal']}</strong></p>
                <p><strong>Action: {interpretation['action']}</strong></p>
                <p>Confidence: {interpretation['confidence']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Show significant OI changes
            if nifty['significant_changes']:
                st.subheader("Significant OI Changes")
                for change in nifty['significant_changes'][:5]:
                    alert_class = "unwinding-alert" if change['oi_change'] < 0 else "buildup-alert"
                    st.markdown(f"""
                    <div class="{alert_class}">
                        <strong>{change['type']} {change['strike']}</strong><br>
                        OI Change: {change['oi_change']:+,.0f}<br>
                        OI: {change['oi']:,.0f} | Price: ₹{change['price']:.2f}
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No Nifty data available in the uploaded file</div>', unsafe_allow_html=True)
    
    # BankNifty Analysis
    with index_cols[1]:
        if index_data['banknifty']:
            banknifty = index_data['banknifty']
            interpretation = banknifty['interpretation']
            
            st.markdown(f"""
            <div class="index-card">
                <h3>BANKNIFTY Analysis</h3>
                <p>Spot Price: {banknifty['spot_price'] if banknifty['spot_price'] else 'N/A'}</p>
                <p>PCR: {banknifty['pcr']:.2f}</p>
                <p>Resistance: {banknifty['resistance_strike'] if banknifty['resistance_strike'] else 'N/A'}</p>
                <p>Support: {banknifty['support_strike'] if banknifty['support_strike'] else 'N/A'}</p>
                <p><strong>Signal: {interpretation['signal']}</strong></p>
                <p><strong>Action: {interpretation['action']}</strong></p>
                <p>Confidence: {interpretation['confidence']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Show significant OI changes
            if banknifty['significant_changes']:
                st.subheader("Significant OI Changes")
                for change in banknifty['significant_changes'][:5]:
                    alert_class = "unwinding-alert" if change['oi_change'] < 0 else "buildup-alert"
                    st.markdown(f"""
                    <div class="{alert_class}">
                        <strong>{change['type']} {change['strike']}</strong><br>
                        OI Change: {change['oi_change']:+,.0f}<br>
                        OI: {change['oi']:,.0f} | Price: ₹{change['price']:.2f}
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No BankNifty data available in the uploaded file</div>', unsafe_allow_html=True)
    
    # PCR Analysis Section
    st.header("PCR & Options Analysis")
    
    # Get PCR items excluding 'oi_analysis'
    pcr_items = [k for k in pcr_data.keys() if k != 'oi_analysis']
    
    if pcr_items:
        pcr_cols = st.columns(len(pcr_items))
        col_idx = 0
        
        for pcr_name in pcr_items:
            pcr_info = pcr_data[pcr_name]
            interpretation = pcr_info['interpretation']
            signal = interpretation['signal']
            
            with pcr_cols[col_idx]:
                st.markdown(f"""
                <div class="pcr-analysis">
                    <h4>{pcr_name}</h4>
                    <h2>{pcr_info['value']:.3f}</h2>
                    <p><strong>{signal}</strong></p>
                    <p>{interpretation['action']}</p>
                    <small>Confidence: {interpretation['confidence']}</small>
                </div>
                """, unsafe_allow_html=True)
            
            col_idx += 1
    else:
        st.markdown('<div class="no-data-message">No PCR data available in the uploaded file</div>', unsafe_allow_html=True)
    
    # OI Analysis
    st.subheader("Open Interest Analysis")
    if 'oi_analysis' in pcr_data and pcr_data['oi_analysis']:
        oi_analysis = pcr_data['oi_analysis']
        
        # Check if there's any OI data before creating columns
        if oi_analysis:
            oi_cols = st.columns(min(4, len(oi_analysis)))
            for i, (oi_name, oi_info) in enumerate(list(oi_analysis.items())[:4]):
                with oi_cols[i]:
                    change_pct = oi_info['change_pct']
                    if abs(change_pct) > 5:
                        alert_class = "unwinding-alert" if change_pct < 0 else "buildup-alert"
                        st.markdown(f"""
                        <div class="{alert_class}">
                            <strong>{oi_name}</strong><br>
                            Current: {oi_info['current']:,.0f}<br>
                            Change: {oi_info['change']:+,.0f}<br>
                            Change %: {change_pct:+.2f}%
                        </div>
                        """, unsafe_allow_html=True)
    else:
        st.markdown('<div class="no-data-message">No Open Interest data available in the uploaded file</div>', unsafe_allow_html=True)
    
    # Options Unwinding Analysis
    st.header("Options Unwinding & Fresh Positions")
    
    unwind_tabs = st.tabs(['Call Unwinding', 'Put Unwinding', 'Fresh Positions'])
    
    with unwind_tabs[0]:
        call_unwinding = unwinding_data['call_unwinding']
        if call_unwinding:
            st.subheader(f"Call Unwinding Detected ({len(call_unwinding)} strikes)")
            for unwind in call_unwinding:
                st.markdown(f"""
                <div class="unwinding-alert">
                    <strong>Strike {unwind['strike']:.0f}</strong> - 
                    OI Change: {unwind['oi_change']:,.0f} - 
                    Price: ₹{unwind['price']:.2f}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No significant call unwinding detected</div>', unsafe_allow_html=True)
    
    with unwind_tabs[1]:
        put_unwinding = unwinding_data['put_unwinding']
        if put_unwinding:
            st.subheader(f"Put Unwinding Detected ({len(put_unwinding)} strikes)")
            for unwind in put_unwinding:
                st.markdown(f"""
                <div class="unwinding-alert">
                    <strong>Strike {unwind['strike']:.0f}</strong> - 
                    OI Change: {unwind['oi_change']:,.0f} - 
                    Price: ₹{unwind['price']:.2f}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No significant put unwinding detected</div>', unsafe_allow_html=True)
    
    with unwind_tabs[2]:
        fresh_positions = unwinding_data['fresh_positions']
        if fresh_positions:
            st.subheader(f"Fresh Positions Built ({len(fresh_positions)} strikes)")
            for position in fresh_positions:
                st.markdown(f"""
                <div class="buildup-alert">
                    <strong>{position['type']} Strike {position['strike']:.0f}</strong> - 
                    Fresh OI: +{position['oi_change']:,.0f}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No significant fresh positions detected</div>', unsafe_allow_html=True)
    
    # F&O Stocks Analysis
    st.header("F&O Stocks Analysis")
    
    stock_tabs = st.tabs(['Bullish Stocks', 'Bearish Stocks', 'Summary'])
    
    with stock_tabs[0]:
        bullish_stocks = fo_stocks['bullish']
        if bullish_stocks:
            st.subheader(f"Bullish F&O Stocks ({len(bullish_stocks)} stocks)")
            
            # Display in grid format
            cols = st.columns(2)
            for i, stock in enumerate(bullish_stocks[:20]):  # Show top 20
                col = cols[i % 2]
                with col:
                    st.markdown(f"""
                    <div class="bullish-stock">
                        <div>
                            <strong>{stock['symbol']}</strong><br>
                            <small>₹{stock['ltp']:.2f} | Vol: {stock['volume']:,.0f}</small><br>
                            <small>{stock['buildup']}</small>
                        </div>
                        <div>
                            <h4>+{stock['change']:.2f}%</h4>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No bullish F&O stocks identified with current thresholds</div>', unsafe_allow_html=True)
    
    with stock_tabs[1]:
        bearish_stocks = fo_stocks['bearish']
        if bearish_stocks:
            st.subheader(f"Bearish F&O Stocks ({len(bearish_stocks)} stocks)")
            
            # Display in grid format
            cols = st.columns(2)
            for i, stock in enumerate(bearish_stocks[:20]):  # Show top 20
                col = cols[i % 2]
                with col:
                    st.markdown(f"""
                    <div class="bearish-stock">
                        <div>
                            <strong>{stock['symbol']}</strong><br>
                            <small>₹{stock['ltp']:.2f} | Vol: {stock['volume']:,.0f}</small><br>
                            <small>{stock['buildup']}</small>
                        </div>
                        <div>
                            <h4>{stock['change']:.2f}%</h4>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-message">No bearish F&O stocks identified with current thresholds</div>', unsafe_allow_html=True)
    
    with stock_tabs[2]:
        # Summary metrics
        st.subheader("F&O Market Summary")
        
        summary_cols = st.columns(4)
        
        with summary_cols[0]:
            st.metric("Bullish Stocks", len(bullish_stocks))
        
        with summary_cols[1]:
            st.metric("Bearish Stocks", len(bearish_stocks))
        
        with summary_cols[2]:
            total_unwinding = len(unwinding_data['call_unwinding']) + len(unwinding_data['put_unwinding'])
            st.metric("Unwinding Strikes", total_unwinding)
        
        with summary_cols[3]:
            fresh_positions_count = len(unwinding_data['fresh_positions'])
            st.metric("Fresh Positions", fresh_positions_count)

def main():
    uploaded_file = st.file_uploader("Upload F&O Excel Data", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        with st.spinner("Processing comprehensive F&O data..."):
            data_dict = read_comprehensive_data(temp_path)
        
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data_dict:
            display_comprehensive_dashboard(data_dict)
        else:
            st.error("Could not process F&O data")
    
    else:
        st.info("Upload your comprehensive F&O Excel file")
        
        st.markdown("""
        ### This dashboard will extract:
        - **All F&O bullish/bearish stocks** with detailed metrics
        - **Specific trading actions** (LONG/SHORT) with entry, target, and stop loss
        - **Comprehensive PCR analysis** with interpretations
        - **Options unwinding detection** (Call/Put unwinding)
        - **Fresh position analysis** (New OI buildup)
        - **Nifty and BankNifty options analysis** with support/resistance levels
        - **Volume and OI change analysis**
        """)

if __name__ == "__main__":
    main()
