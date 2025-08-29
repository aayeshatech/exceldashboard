import pandas as pd
import streamlit as st
import numpy as np
from datetime import datetime, timedelta
import time
import os

st.set_page_config(page_title="Live Trading Alerts", page_icon="🚨", layout="wide")

# Live trading CSS with alerts
st.markdown("""
<style>
.alert-header {
    background: linear-gradient(90deg, #ff6b6b 0%, #feca57 100%);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    text-align: center;
    margin-bottom: 1rem;
    animation: pulse 2s infinite;
}
@keyframes pulse {
    0% { opacity: 1; }
    50% { opacity: 0.7; }
    100% { opacity: 1; }
}
.long-alert {
    background: linear-gradient(135deg, #2ecc71, #27ae60);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
    border-left: 5px solid #1abc9c;
    animation: slideIn 0.5s ease-in;
}
.short-alert {
    background: linear-gradient(135deg, #e74c3c, #c0392b);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
    border-left: 5px solid #e67e22;
    animation: slideIn 0.5s ease-in;
}
@keyframes slideIn {
    from { transform: translateX(-100%); }
    to { transform: translateX(0); }
}
.pcr-alert {
    background: linear-gradient(135deg, #9b59b6, #8e44ad);
    color: white;
    padding: 1rem;
    border-radius: 8px;
    margin: 0.5rem 0;
    text-align: center;
}
.support-resistance {
    background: #f8f9fa;
    border: 2px solid #007bff;
    padding: 1rem;
    border-radius: 8px;
    margin: 1rem 0;
}
.live-data {
    background: #1a1a1a;
    color: #00ff00;
    font-family: 'Courier New', monospace;
    padding: 1rem;
    border-radius: 8px;
    margin: 1rem 0;
}
.metric-alert {
    padding: 0.5rem;
    border-radius: 5px;
    margin: 0.2rem 0;
    font-weight: bold;
}
.bullish { background: #d4edda; color: #155724; }
.bearish { background: #f8d7da; color: #721c24; }
.neutral { background: #fff3cd; color: #856404; }
</style>
""", unsafe_allow_html=True)

# Session state for auto-refresh
if 'last_refresh' not in st.session_state:
    st.session_state.last_refresh = datetime.now()
if 'refresh_count' not in st.session_state:
    st.session_state.refresh_count = 0
if 'previous_data' not in st.session_state:
    st.session_state.previous_data = {}
if 'alerts_history' not in st.session_state:
    st.session_state.alerts_history = []

def read_trading_data(file_path):
    """Read trading data with error handling"""
    try:
        excel_file = pd.ExcelFile(file_path)
        data_dict = {}
        
        priority_sheets = [
            'Dashboard', 'Stock Dashboard', 'PCR & OI Chart', 
            'Sector Dashboard', 'Screener'
        ]
        
        for sheet_name in priority_sheets:
            if sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    if not df.empty:
                        data_dict[sheet_name] = df
                except Exception as e:
                    st.warning(f"Could not read {sheet_name}: {str(e)}")
        
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return {}

def analyze_pcr_alerts(pcr_df):
    """Generate PCR-based alerts"""
    alerts = []
    
    try:
        for col in pcr_df.columns:
            if 'PCR' in str(col).upper():
                pcr_values = pd.to_numeric(pcr_df[col], errors='coerce').dropna()
                if not pcr_values.empty:
                    current_pcr = pcr_values.iloc[-1]
                    
                    if current_pcr > 1.5:
                        alerts.append({
                            'type': 'PCR_BEARISH',
                            'message': f'STRONG BEARISH SIGNAL - {col}: {current_pcr:.3f}',
                            'action': 'Consider PUT buying or short positions',
                            'priority': 'HIGH'
                        })
                    elif current_pcr > 1.2:
                        alerts.append({
                            'type': 'PCR_BEARISH',
                            'message': f'Bearish bias - {col}: {current_pcr:.3f}',
                            'action': 'Bias towards PUT positions',
                            'priority': 'MEDIUM'
                        })
                    elif current_pcr < 0.6:
                        alerts.append({
                            'type': 'PCR_BULLISH',
                            'message': f'STRONG BULLISH SIGNAL - {col}: {current_pcr:.3f}',
                            'action': 'Consider CALL buying or long positions',
                            'priority': 'HIGH'
                        })
                    elif current_pcr < 0.8:
                        alerts.append({
                            'type': 'PCR_BULLISH',
                            'message': f'Bullish bias - {col}: {current_pcr:.3f}',
                            'action': 'Bias towards CALL positions',
                            'priority': 'MEDIUM'
                        })
        
        return alerts
    except Exception as e:
        st.warning(f"PCR analysis error: {e}")
        return []

def analyze_stock_alerts(stock_df):
    """Generate stock trading alerts"""
    alerts = []
    
    try:
        symbol_cols = [col for col in stock_df.columns if any(term in str(col).upper() for term in ['SYMBOL', 'STOCK', 'NAME'])]
        change_cols = [col for col in stock_df.columns if any(term in str(col).upper() for term in ['CHANGE', '%'])]
        volume_cols = [col for col in stock_df.columns if 'VOLUME' in str(col).upper()]
        
        if not symbol_cols or not change_cols:
            return alerts
        
        symbol_col = symbol_cols[0]
        change_col = change_cols[0]
        
        df_clean = stock_df[[symbol_col, change_col]].copy()
        df_clean[change_col] = pd.to_numeric(df_clean[change_col], errors='coerce')
        df_clean = df_clean.dropna()
        
        # Strong long alerts (>5% gain)
        strong_longs = df_clean[df_clean[change_col] > 5].nlargest(10, change_col)
        for _, row in strong_longs.iterrows():
            alerts.append({
                'type': 'LONG',
                'symbol': row[symbol_col],
                'change': row[change_col],
                'message': f'STRONG LONG SIGNAL - {row[symbol_col]} +{row[change_col]:.2f}%',
                'action': 'Consider long position with tight stop loss',
                'priority': 'HIGH'
            })
        
        # Strong short alerts (<-5% decline)
        strong_shorts = df_clean[df_clean[change_col] < -5].nsmallest(10, change_col)
        for _, row in strong_shorts.iterrows():
            alerts.append({
                'type': 'SHORT',
                'symbol': row[symbol_col],
                'change': row[change_col],
                'message': f'STRONG SHORT SIGNAL - {row[symbol_col]} {row[change_col]:.2f}%',
                'action': 'Consider short position or PUT options',
                'priority': 'HIGH'
            })
        
        # Medium alerts (2-5% moves)
        medium_longs = df_clean[(df_clean[change_col] > 2) & (df_clean[change_col] <= 5)].nlargest(5, change_col)
        for _, row in medium_longs.iterrows():
            alerts.append({
                'type': 'LONG',
                'symbol': row[symbol_col],
                'change': row[change_col],
                'message': f'Long opportunity - {row[symbol_col]} +{row[change_col]:.2f}%',
                'action': 'Monitor for continuation',
                'priority': 'MEDIUM'
            })
        
        medium_shorts = df_clean[(df_clean[change_col] < -2) & (df_clean[change_col] >= -5)].nsmallest(5, change_col)
        for _, row in medium_shorts.iterrows():
            alerts.append({
                'type': 'SHORT',
                'symbol': row[symbol_col],
                'change': row[change_col],
                'message': f'Short opportunity - {row[symbol_col]} {row[change_col]:.2f}%',
                'action': 'Monitor for further decline',
                'priority': 'MEDIUM'
            })
        
        return alerts
    
    except Exception as e:
        st.warning(f"Stock alerts error: {e}")
        return []

def detect_data_changes(current_data, previous_data):
    """Detect significant changes in data"""
    change_alerts = []
    
    try:
        if not previous_data:
            return change_alerts
        
        # Check for new strong movers
        for sheet_name in current_data:
            if sheet_name in previous_data:
                current_df = current_data[sheet_name]
                previous_df = previous_data[sheet_name]
                
                # Look for significant changes in key metrics
                if 'PCR' in sheet_name.upper():
                    # PCR changes
                    for col in current_df.columns:
                        if 'PCR' in str(col).upper():
                            try:
                                current_pcr = pd.to_numeric(current_df[col], errors='coerce').dropna().iloc[-1]
                                previous_pcr = pd.to_numeric(previous_df[col], errors='coerce').dropna().iloc[-1]
                                
                                pcr_change = abs(current_pcr - previous_pcr)
                                if pcr_change > 0.1:
                                    change_alerts.append({
                                        'type': 'PCR_CHANGE',
                                        'message': f'{col} changed significantly: {previous_pcr:.3f} → {current_pcr:.3f}',
                                        'priority': 'HIGH' if pcr_change > 0.2 else 'MEDIUM'
                                    })
                            except:
                                pass
        
        return change_alerts
    
    except Exception as e:
        st.warning(f"Change detection error: {e}")
        return []

def display_live_alerts(all_alerts):
    """Display live trading alerts"""
    
    if not all_alerts:
        st.info("No active alerts")
        return
    
    # Separate alerts by priority
    high_priority = [alert for alert in all_alerts if alert.get('priority') == 'HIGH']
    medium_priority = [alert for alert in all_alerts if alert.get('priority') == 'MEDIUM']
    
    if high_priority:
        st.markdown("""
        <div class="alert-header">
            🚨 HIGH PRIORITY TRADING ALERTS 🚨
        </div>
        """, unsafe_allow_html=True)
        
        for alert in high_priority[:10]:  # Top 10 high priority
            alert_type = alert.get('type', '')
            
            if 'LONG' in alert_type or 'BULLISH' in alert_type:
                card_class = 'long-alert'
            elif 'SHORT' in alert_type or 'BEARISH' in alert_type:
                card_class = 'short-alert'
            else:
                card_class = 'pcr-alert'
            
            st.markdown(f"""
            <div class="{card_class}">
                <strong>{alert['message']}</strong><br>
                Action: {alert.get('action', 'Monitor closely')}<br>
                Time: {datetime.now().strftime('%H:%M:%S')}
            </div>
            """, unsafe_allow_html=True)
    
    if medium_priority:
        st.subheader("Medium Priority Alerts")
        for alert in medium_priority[:10]:
            alert_type = alert.get('type', '')
            
            if 'LONG' in alert_type or 'BULLISH' in alert_type:
                st.success(f"🟢 {alert['message']} - {alert.get('action', '')}")
            elif 'SHORT' in alert_type or 'BEARISH' in alert_type:
                st.error(f"🔴 {alert['message']} - {alert.get('action', '')}")
            else:
                st.warning(f"🟡 {alert['message']} - {alert.get('action', '')}")

def display_live_summary(data_dict):
    """Display live market summary"""
    
    st.header("Live Market Summary")
    
    current_time = datetime.now().strftime('%H:%M:%S')
    
    summary_cols = st.columns(4)
    
    # Count alerts by type
    total_stocks = 0
    bullish_stocks = 0
    bearish_stocks = 0
    neutral_stocks = 0
    
    for sheet_name in data_dict:
        if 'Dashboard' in sheet_name:
            df = data_dict[sheet_name]
            change_cols = [col for col in df.columns if any(term in str(col).upper() for term in ['CHANGE', '%'])]
            if change_cols:
                changes = pd.to_numeric(df[change_cols[0]], errors='coerce').dropna()
                total_stocks += len(changes)
                bullish_stocks += len(changes[changes > 1])
                bearish_stocks += len(changes[changes < -1])
                neutral_stocks += len(changes[abs(changes) <= 1])
    
    with summary_cols[0]:
        st.markdown(f"""
        <div class="live-data">
            <h3>{total_stocks}</h3>
            <p>TOTAL STOCKS</p>
            <small>{current_time}</small>
        </div>
        """, unsafe_allow_html=True)
    
    with summary_cols[1]:
        st.markdown(f"""
        <div class="metric-alert bullish">
            <h3>{bullish_stocks}</h3>
            <p>BULLISH STOCKS</p>
        </div>
        """, unsafe_allow_html=True)
    
    with summary_cols[2]:
        st.markdown(f"""
        <div class="metric-alert bearish">
            <h3>{bearish_stocks}</h3>
            <p>BEARISH STOCKS</p>
        </div>
        """, unsafe_allow_html=True)
    
    with summary_cols[3]:
        st.markdown(f"""
        <div class="metric-alert neutral">
            <h3>{neutral_stocks}</h3>
            <p>NEUTRAL STOCKS</p>
        </div>
        """, unsafe_allow_html=True)

def display_support_resistance(data_dict):
    """Display support and resistance levels"""
    
    st.subheader("Key Levels")
    
    levels_data = []
    
    # Extract from PCR data if available
    if 'PCR & OI Chart' in data_dict:
        pcr_df = data_dict['PCR & OI Chart']
        
        # Look for support/resistance indicators
        for col in pcr_df.columns:
            if any(term in str(col).upper() for term in ['SUPPORT', 'RESISTANCE', 'LEVEL']):
                values = pd.to_numeric(pcr_df[col], errors='coerce').dropna()
                if not values.empty:
                    current_value = values.iloc[-1]
                    levels_data.append({
                        'Level': col,
                        'Value': current_value,
                        'Type': 'Support' if 'SUPPORT' in str(col).upper() else 'Resistance'
                    })
    
    if levels_data:
        levels_df = pd.DataFrame(levels_data)
        st.dataframe(levels_df, use_container_width=True)
    else:
        st.info("Support/Resistance levels will appear when data is available")

def main():
    st.markdown("""
    <div class="alert-header">
        🔴 LIVE TRADING ALERTS DASHBOARD 🔴
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar controls
    st.sidebar.header("Live Dashboard Controls")
    
    auto_refresh = st.sidebar.checkbox("Auto Refresh", value=True)
    refresh_interval = st.sidebar.slider("Refresh Interval (seconds)", 5, 60, 10)
    
    if st.sidebar.button("Manual Refresh Now"):
        st.session_state.refresh_count += 1
        st.rerun()
    
    st.sidebar.metric("Refresh Count", st.session_state.refresh_count)
    st.sidebar.metric("Last Refresh", st.session_state.last_refresh.strftime('%H:%M:%S'))
    
    # File upload
    uploaded_file = st.file_uploader("Upload Live Trading Data", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Read current data
        current_data = read_trading_data(temp_path)
        
        try:
            os.remove(temp_path)
        except:
            pass
        
        if current_data:
            # Generate alerts
            all_alerts = []
            
            # PCR alerts
            if 'PCR & OI Chart' in current_data:
                pcr_alerts = analyze_pcr_alerts(current_data['PCR & OI Chart'])
                all_alerts.extend(pcr_alerts)
            
            # Stock alerts
            for sheet_name in current_data:
                if 'Dashboard' in sheet_name:
                    stock_alerts = analyze_stock_alerts(current_data[sheet_name])
                    all_alerts.extend(stock_alerts)
            
            # Change detection alerts
            change_alerts = detect_data_changes(current_data, st.session_state.previous_data)
            all_alerts.extend(change_alerts)
            
            # Update previous data
            st.session_state.previous_data = current_data.copy()
            
            # Display alerts
            display_live_alerts(all_alerts)
            
            # Display summary
            display_live_summary(current_data)
            
            # Display support/resistance
            display_support_resistance(current_data)
            
            # Alert history
            if all_alerts:
                st.session_state.alerts_history.extend([
                    {
                        'time': datetime.now().strftime('%H:%M:%S'),
                        'alert': alert['message']
                    } for alert in all_alerts[:5]
                ])
                
                # Keep only last 20 alerts
                if len(st.session_state.alerts_history) > 20:
                    st.session_state.alerts_history = st.session_state.alerts_history[-20:]
            
            # Display alert history
            if st.session_state.alerts_history:
                with st.expander("Alert History"):
                    for alert_record in reversed(st.session_state.alerts_history[-10:]):
                        st.write(f"{alert_record['time']}: {alert_record['alert']}")
            
            # Auto-refresh logic
            if auto_refresh:
                time_since_refresh = (datetime.now() - st.session_state.last_refresh).total_seconds()
                
                if time_since_refresh >= refresh_interval:
                    st.session_state.last_refresh = datetime.now()
                    st.session_state.refresh_count += 1
                    time.sleep(1)
                    st.rerun()
                else:
                    time_remaining = refresh_interval - time_since_refresh
                    st.sidebar.info(f"Next refresh in {int(time_remaining)} seconds")
                    time.sleep(1)
                    st.rerun()
        
        else:
            st.error("Could not load trading data")
    
    else:
        st.info("Upload your Excel file to start live alerts")
        
        st.markdown("""
        ### Features:
        - Real-time PCR alerts for options trading
        - Long/Short signals based on stock performance  
        - Auto-refresh with customizable intervals
        - Change detection between data updates
        - Support/Resistance level monitoring
        - Alert history tracking
        """)

if __name__ == "__main__":
    main()
