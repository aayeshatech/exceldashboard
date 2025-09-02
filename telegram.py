#!/usr/bin/env python3
"""
Telegram Excel Monitor - Streamlit Version
Web application to monitor Excel files and send Telegram alerts
"""

import streamlit as st
import requests
import pandas as pd
import time
import json
from datetime import datetime
import io
import threading

# Page configuration
st.set_page_config(
    page_title="Telegram Excel Monitor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class TelegramMonitor:
    def __init__(self):
        self.initialize_session_state()
    
    def initialize_session_state(self):
        """Initialize session state variables"""
        if 'monitoring' not in st.session_state:
            st.session_state.monitoring = False
        if 'last_alert' not in st.session_state:
            st.session_state.last_alert = {"symbol": None, "type": None}
        if 'logs' not in st.session_state:
            st.session_state.logs = []
        if 'bot_token' not in st.session_state:
            st.session_state.bot_token = "7613703350:AAE-W4dJ37lngM4lO2Tnuns8-a-80jYRtxk"
        if 'chat_id' not in st.session_state:
            st.session_state.chat_id = "-1002840229810"
    
    def log_message(self, message):
        """Add message to logs"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        st.session_state.logs.append(log_entry)
        # Keep only last 50 logs
        if len(st.session_state.logs) > 50:
            st.session_state.logs = st.session_state.logs[-50:]
        print(log_entry)  # Also print to console
    
    def send_telegram_message(self, message):
        """Send message to Telegram"""
        try:
            bot_token = st.session_state.bot_token.strip()
            chat_id = st.session_state.chat_id.strip()
            
            if not bot_token or not chat_id:
                self.log_message("ERROR: Bot token or chat ID is missing")
                return False
            
            url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
            data = {
                "chat_id": chat_id,
                "text": message,
                "parse_mode": "HTML"
            }
            
            response = requests.post(url, json=data, timeout=10)
            result = response.json()
            
            if response.status_code == 200 and result.get("ok"):
                self.log_message("✅ Message sent successfully!")
                return True
            else:
                error_msg = result.get("description", "Unknown error")
                self.log_message(f"❌ Telegram API Error: {error_msg}")
                return False
                
        except requests.RequestException as e:
            self.log_message(f"❌ Network error: {str(e)}")
            return False
        except Exception as e:
            self.log_message(f"❌ Unexpected error: {str(e)}")
            return False
    
    def format_alert_message(self, signal):
        """Format alert message"""
        timestamp = datetime.now().strftime("%d/%m/%Y, %I:%M:%S %p")
        
        emoji_map = {
            'Long Buildup': '🚀',
            'Short Cover': '🔥', 
            'Strong Bullish': '💪',
            'Bullish': '📈'
        }
        
        emoji = emoji_map.get(signal['signalType'], '📈')
        
        message = f"""{emoji} <b>SECTOR DASHBOARD ALERT</b>

<b>Symbol:</b> {signal['symbol']}
<b>Signal:</b> {signal['signalType']}
<b>Time:</b> {timestamp}

📊 Monitor your positions accordingly!

<i>Auto-generated from Streamlit Excel Monitor</i>"""
        
        return message
    
    def analyze_dataframe(self, df):
        """Analyze dataframe for trading signals"""
        signals = []
        
        try:
            self.log_message(f"📊 Analyzing data: {len(df)} rows, {len(df.columns)} columns")
            
            # Convert all data to string for searching
            df_str = df.astype(str)
            
            # Search through all cells for NSE: symbols
            for col_idx, col in enumerate(df.columns):
                for row_idx, value in enumerate(df[col]):
                    if pd.isna(value):
                        continue
                    
                    value_str = str(value)
                    if 'NSE:' in value_str:
                        symbol = value_str.replace('NSE:', '').strip()
                        
                        # Analyze surrounding cells for signal type
                        signal_type = self.determine_signal_from_context(df_str, row_idx, col_idx)
                        
                        if signal_type:
                            signals.append({
                                'symbol': symbol,
                                'signalType': signal_type,
                                'row': row_idx,
                                'col': col_idx
                            })
                            self.log_message(f"📈 Found signal: {symbol} - {signal_type}")
            
            return signals
            
        except Exception as e:
            self.log_message(f"❌ Error analyzing data: {str(e)}")
            return []
    
    def determine_signal_from_context(self, df_str, row_idx, col_idx):
        """Determine signal type from surrounding context"""
        search_range = 5
        
        # Search nearby cells for signal keywords
        for r in range(max(0, row_idx - search_range), min(len(df_str), row_idx + search_range + 1)):
            for c in range(max(0, col_idx - search_range), min(len(df_str.columns), col_idx + search_range + 1)):
                try:
                    cell_value = str(df_str.iloc[r, c]).lower()
                    
                    if 'long' in cell_value and 'buildup' in cell_value:
                        return 'Long Buildup'
                    elif 'short' in cell_value and 'cover' in cell_value:
                        return 'Short Cover'
                    elif 'strong' in cell_value and 'bullish' in cell_value:
                        return 'Strong Bullish'
                    elif 'bullish' in cell_value:
                        return 'Bullish'
                except:
                    continue
        
        return None
    
    def check_for_signals(self, df):
        """Check dataframe for trading signals"""
        if df is None:
            self.log_message("❌ No data to analyze")
            return
        
        self.log_message("🔍 Checking for signals...")
        signals = self.analyze_dataframe(df)
        
        if not signals:
            self.log_message("ℹ️ No signals found in current scan")
            return
        
        # Find the highest priority signal
        priority_order = ['Long Buildup', 'Short Cover', 'Strong Bullish', 'Bullish']
        top_signal = None
        
        for priority in priority_order:
            signal = next((s for s in signals if s['signalType'] == priority), None)
            if signal:
                top_signal = signal
                break
        
        if top_signal:
            # Check if this is a new signal
            if (st.session_state.last_alert["symbol"] != top_signal['symbol'] or 
                st.session_state.last_alert["type"] != top_signal['signalType']):
                
                message = self.format_alert_message(top_signal)
                
                if self.send_telegram_message(message):
                    st.session_state.last_alert = {
                        "symbol": top_signal['symbol'],
                        "type": top_signal['signalType']
                    }
                    self.log_message(f"🚀 Alert sent: {top_signal['symbol']} - {top_signal['signalType']}")
                    st.success(f"Alert sent: {top_signal['symbol']} - {top_signal['signalType']}")
            else:
                self.log_message(f"ℹ️ No change in signal for {top_signal['symbol']}")

def main():
    monitor = TelegramMonitor()
    
    # Title and header
    st.title("📊 Telegram Excel Monitor")
    st.markdown("Monitor Excel files for trading signals and send alerts to Telegram")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Bot configuration
        bot_token = st.text_input(
            "Bot Token:",
            value=st.session_state.bot_token,
            type="password",
            help="Your Telegram bot token"
        )
        st.session_state.bot_token = bot_token
        
        chat_id = st.text_input(
            "Chat ID:",
            value=st.session_state.chat_id,
            help="Telegram chat ID (negative for groups)"
        )
        st.session_state.chat_id = chat_id
        
        st.divider()
        
        # Test alert
        st.header("🧪 Test Alert")
        if st.button("Send Test Alert", use_container_width=True):
            test_message = f"""🧪 <b>TEST ALERT</b>

This is a test message from the Streamlit Excel Monitor.
Timestamp: {datetime.now().strftime("%d/%m/%Y, %I:%M:%S %p")}"""
            
            if monitor.send_telegram_message(test_message):
                st.success("Test alert sent successfully!")
            else:
                st.error("Failed to send test alert. Check logs for details.")
        
        st.divider()
        
        # Manual alert
        st.header("📤 Manual Alert")
        manual_symbol = st.text_input("Stock Symbol:", placeholder="e.g., RELIANCE, TCS")
        manual_signal = st.selectbox(
            "Signal Type:",
            ["", "Long Buildup", "Short Cover", "Strong Bullish", "Bullish"]
        )
        
        if st.button("Send Manual Alert", use_container_width=True):
            if manual_symbol and manual_signal:
                signal = {
                    'symbol': manual_symbol.upper().replace('NSE:', ''),
                    'signalType': manual_signal
                }
                message = monitor.format_alert_message(signal)
                
                if monitor.send_telegram_message(message):
                    st.session_state.last_alert = {
                        "symbol": signal['symbol'],
                        "type": signal['signalType']
                    }
                    st.success(f"Alert sent: {signal['symbol']} - {signal['signalType']}")
                else:
                    st.error("Failed to send manual alert")
            else:
                st.error("Please enter both symbol and signal type")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 File Upload")
        
        # File uploader
        uploaded_file = st.file_uploader(
            "Upload your Excel file",
            type=['xlsx', 'xlsm', 'xls', 'csv'],
            help="Upload your Sector Dashboard Excel file or CSV export"
        )
        
        df = None
        if uploaded_file is not None:
            try:
                # Read the file
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                    monitor.log_message(f"📄 CSV file loaded: {uploaded_file.name}")
                else:
                    # Try to read Sector Dashboard sheet first
                    try:
                        df = pd.read_excel(uploaded_file, sheet_name='Sector Dashboard')
                        monitor.log_message(f"📊 Excel file loaded: {uploaded_file.name} (Sector Dashboard)")
                    except:
                        # If sheet doesn't exist, read first sheet
                        df = pd.read_excel(uploaded_file)
                        monitor.log_message(f"📊 Excel file loaded: {uploaded_file.name} (First sheet)")
                
                st.success(f"File loaded successfully: {len(df)} rows, {len(df.columns)} columns")
                
                # Show preview of the data
                with st.expander("📋 Data Preview"):
                    st.dataframe(df.head(10), use_container_width=True)
                
            except Exception as e:
                st.error(f"Error loading file: {str(e)}")
                monitor.log_message(f"❌ Error loading file: {str(e)}")
        
        # Analysis controls
        st.header("🔍 Analysis Controls")
        col_a, col_b = st.columns(2)
        
        with col_a:
            if st.button("🔍 Check Now", use_container_width=True):
                if df is not None:
                    monitor.check_for_signals(df)
                else:
                    st.error("Please upload a file first")
        
        with col_b:
            if st.button("🔄 Clear Logs", use_container_width=True):
                st.session_state.logs = []
                st.success("Logs cleared")
    
    with col2:
        st.header("📈 Status")
        
        # Status display
        status_col1, status_col2 = st.columns(2)
        with status_col1:
            st.metric("File Status", "Loaded" if df is not None else "Not Loaded")
        with status_col2:
            st.metric("Logs", len(st.session_state.logs))
        
        # Last alert info
        if st.session_state.last_alert["symbol"]:
            st.info(f"**Last Alert:**  \n{st.session_state.last_alert['symbol']} - {st.session_state.last_alert['type']}")
        else:
            st.info("**Last Alert:**  \nNone")
        
        # Auto-refresh toggle
        auto_refresh = st.checkbox("Auto-refresh every 30 seconds")
        if auto_refresh:
            time.sleep(30)
            st.rerun()
    
    # Logs section
    st.header("📋 Activity Logs")
    if st.session_state.logs:
        log_text = "\n".join(reversed(st.session_state.logs[-20:]))  # Show last 20 logs
        st.text_area("Recent Activity", value=log_text, height=200, disabled=True)
    else:
        st.info("No activity logs yet. Upload a file and run analysis to see logs.")
    
    # Instructions
    with st.expander("ℹ️ How to Use"):
        st.markdown("""
        ### Setup Instructions:
        1. **Configure Telegram**: Enter your bot token and chat ID in the sidebar
        2. **Test Connection**: Click "Send Test Alert" to verify your configuration
        3. **Upload File**: Upload your Excel file or CSV export of the Sector Dashboard
        4. **Check for Signals**: Click "Check Now" to analyze the data for trading signals
        5. **Manual Testing**: Use the manual alert feature to test specific stocks
        
        ### Supported Signal Types:
        - **Long Buildup** 🚀: Strong buying interest
        - **Short Cover** 🔥: Short covering activity  
        - **Strong Bullish** 💪: Very positive momentum
        - **Bullish** 📈: Positive momentum
        
        ### File Requirements:
        - Excel files (.xlsx, .xlsm, .xls) with a "Sector Dashboard" sheet
        - CSV exports from your Excel dashboard
        - Data should contain stock symbols in "NSE:SYMBOL" format
        """)

if __name__ == "__main__":
    main()
