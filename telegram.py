#!/usr/bin/env python3
"""
Telegram Excel Monitor - Python Version
A desktop application to monitor Excel files and send Telegram alerts
"""

import requests
import pandas as pd
import time
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from datetime import datetime
import os

class TelegramExcelMonitor:
    def __init__(self):
        self.bot_token = "7613703350:AAE-W4dJ37lngM4lO2Tnuns8-a-80jYRtxk"
        self.chat_id = "-1002840229810"
        self.excel_file = None
        self.monitoring = False
        self.last_alert = {"symbol": None, "type": None}
        
        # Create GUI
        self.create_gui()
    
    def create_gui(self):
        self.root = tk.Tk()
        self.root.title("Telegram Excel Monitor")
        self.root.geometry("800x600")
        
        # Configuration Frame
        config_frame = ttk.LabelFrame(self.root, text="Configuration", padding=10)
        config_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(config_frame, text="Bot Token:").grid(row=0, column=0, sticky="w")
        self.token_var = tk.StringVar(value=self.bot_token)
        token_entry = ttk.Entry(config_frame, textvariable=self.token_var, width=50, show="*")
        token_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(config_frame, text="Chat ID:").grid(row=1, column=0, sticky="w")
        self.chat_var = tk.StringVar(value=self.chat_id)
        chat_entry = ttk.Entry(config_frame, textvariable=self.chat_var, width=50)
        chat_entry.grid(row=1, column=1, padx=5)
        
        ttk.Label(config_frame, text="Interval (min):").grid(row=2, column=0, sticky="w")
        self.interval_var = tk.StringVar(value="1")
        interval_entry = ttk.Entry(config_frame, textvariable=self.interval_var, width=10)
        interval_entry.grid(row=2, column=1, sticky="w", padx=5)
        
        # File Selection Frame
        file_frame = ttk.LabelFrame(self.root, text="Excel File", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side="left")
        
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side="right")
        
        # Controls Frame
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(control_frame, text="Test Alert", command=self.test_alert).pack(side="left", padx=5)
        self.start_btn = ttk.Button(control_frame, text="Start Monitor", command=self.start_monitoring)
        self.start_btn.pack(side="left", padx=5)
        self.stop_btn = ttk.Button(control_frame, text="Stop Monitor", command=self.stop_monitoring, state="disabled")
        self.stop_btn.pack(side="left", padx=5)
        ttk.Button(control_frame, text="Check Now", command=self.check_now).pack(side="left", padx=5)
        
        # Manual Alert Frame
        manual_frame = ttk.LabelFrame(self.root, text="Manual Alert", padding=10)
        manual_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(manual_frame, text="Symbol:").grid(row=0, column=0, sticky="w")
        self.manual_symbol = tk.StringVar()
        ttk.Entry(manual_frame, textvariable=self.manual_symbol, width=20).grid(row=0, column=1, padx=5)
        
        ttk.Label(manual_frame, text="Signal:").grid(row=0, column=2, sticky="w", padx=(20,0))
        self.manual_signal = tk.StringVar()
        signal_combo = ttk.Combobox(manual_frame, textvariable=self.manual_signal, width=15)
        signal_combo['values'] = ('Long Buildup', 'Short Cover', 'Strong Bullish', 'Bullish')
        signal_combo.grid(row=0, column=3, padx=5)
        
        ttk.Button(manual_frame, text="Send Alert", command=self.send_manual_alert).grid(row=0, column=4, padx=10)
        
        # Status Frame
        status_frame = ttk.LabelFrame(self.root, text="Status", padding=10)
        status_frame.pack(fill="x", padx=10, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Status: Idle")
        self.status_label.pack(anchor="w")
        
        self.last_alert_label = ttk.Label(status_frame, text="Last Alert: None")
        self.last_alert_label.pack(anchor="w")
        
        # Log Frame
        log_frame = ttk.LabelFrame(self.root, text="Log", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        
        self.log("System initialized. Ready to monitor Excel files.")
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        print(log_entry.strip())
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.excel_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"Selected: {filename}")
            self.log(f"File selected: {filename}")
    
    def send_telegram_message(self, message):
        try:
            bot_token = self.token_var.get().strip()
            chat_id = self.chat_var.get().strip()
            
            if not bot_token or not chat_id:
                self.log("ERROR: Bot token or chat ID is missing")
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
                self.log("✅ Message sent successfully!")
                return True
            else:
                error_msg = result.get("description", "Unknown error")
                self.log(f"❌ Telegram API Error: {error_msg}")
                return False
                
        except requests.RequestException as e:
            self.log(f"❌ Network error: {str(e)}")
            return False
        except Exception as e:
            self.log(f"❌ Unexpected error: {str(e)}")
            return False
    
    def test_alert(self):
        test_message = """🧪 <b>TEST ALERT</b>

This is a test message from the Python Excel Monitor.
Timestamp: """ + datetime.now().strftime("%d/%m/%Y, %I:%M:%S %p")
        
        if self.send_telegram_message(test_message):
            messagebox.showinfo("Success", "Test alert sent successfully!")
        else:
            messagebox.showerror("Error", "Failed to send test alert. Check the log for details.")
    
    def format_alert_message(self, signal):
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

<i>Auto-generated from Python Excel Monitor</i>"""
        
        return message
    
    def send_manual_alert(self):
        symbol = self.manual_symbol.get().strip().upper()
        signal_type = self.manual_signal.get().strip()
        
        if not symbol or not signal_type:
            messagebox.showerror("Error", "Please enter both symbol and signal type")
            return
        
        signal = {
            'symbol': symbol.replace('NSE:', ''),
            'signalType': signal_type
        }
        
        message = self.format_alert_message(signal)
        
        if self.send_telegram_message(message):
            self.log(f"✅ Manual alert sent: {symbol} - {signal_type}")
            self.last_alert = {"symbol": symbol, "type": signal_type}
            self.update_status()
            # Clear the form
            self.manual_symbol.set("")
            self.manual_signal.set("")
            messagebox.showinfo("Success", f"Alert sent for {symbol} - {signal_type}")
        else:
            messagebox.showerror("Error", "Failed to send manual alert")
    
    def analyze_excel_file(self):
        if not self.excel_file:
            self.log("❌ No Excel file selected")
            return None
        
        try:
            # Read Excel file
            if self.excel_file.endswith('.csv'):
                df = pd.read_csv(self.excel_file)
            else:
                # Try to read the Sector Dashboard sheet
                try:
                    df = pd.read_excel(self.excel_file, sheet_name='Sector Dashboard')
                except:
                    # If sheet doesn't exist, read the first sheet
                    df = pd.read_excel(self.excel_file)
            
            self.log(f"📊 Loaded data: {len(df)} rows, {len(df.columns)} columns")
            
            # Look for stock symbols and signals
            signals = []
            
            # Search through all cells for NSE: symbols
            for col in df.columns:
                for idx, value in enumerate(df[col]):
                    if pd.isna(value):
                        continue
                    
                    value_str = str(value)
                    if 'NSE:' in value_str:
                        symbol = value_str.replace('NSE:', '').strip()
                        
                        # Analyze surrounding cells for signal type
                        signal_type = self.determine_signal_from_context(df, idx, col)
                        
                        if signal_type:
                            signals.append({
                                'symbol': symbol,
                                'signalType': signal_type,
                                'row': idx,
                                'col': col
                            })
            
            return signals
            
        except Exception as e:
            self.log(f"❌ Error analyzing Excel file: {str(e)}")
            return None
    
    def determine_signal_from_context(self, df, row_idx, col_name):
        # Look in surrounding cells for signal indicators
        search_range = 5
        
        # Get column index
        col_idx = df.columns.get_loc(col_name)
        
        # Search nearby cells for signal keywords
        for r in range(max(0, row_idx - search_range), min(len(df), row_idx + search_range + 1)):
            for c in range(max(0, col_idx - search_range), min(len(df.columns), col_idx + search_range + 1)):
                try:
                    cell_value = str(df.iloc[r, c]).lower()
                    
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
    
    def check_now(self):
        if not self.excel_file:
            messagebox.showerror("Error", "Please select an Excel file first")
            return
        
        self.log("🔍 Checking for signals...")
        signals = self.analyze_excel_file()
        
        if not signals:
            self.log("ℹ️ No signals found in current scan")
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
            if (self.last_alert["symbol"] != top_signal['symbol'] or 
                self.last_alert["type"] != top_signal['signalType']):
                
                message = self.format_alert_message(top_signal)
                
                if self.send_telegram_message(message):
                    self.last_alert = {
                        "symbol": top_signal['symbol'],
                        "type": top_signal['signalType']
                    }
                    self.log(f"🚀 Alert sent: {top_signal['symbol']} - {top_signal['signalType']}")
                    self.update_status()
            else:
                self.log(f"ℹ️ No change in signal for {top_signal['symbol']}")
    
    def update_status(self):
        status_text = "Status: Monitoring" if self.monitoring else "Status: Idle"
        self.status_label.config(text=status_text)
        
        if self.last_alert["symbol"]:
            alert_text = f"Last Alert: {self.last_alert['symbol']} - {self.last_alert['type']}"
        else:
            alert_text = "Last Alert: None"
        self.last_alert_label.config(text=alert_text)
    
    def start_monitoring(self):
        if not self.excel_file:
            messagebox.showerror("Error", "Please select an Excel file first")
            return
        
        if self.monitoring:
            return
        
        try:
            interval = int(self.interval_var.get())
            if interval < 1:
                raise ValueError("Interval must be at least 1 minute")
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid interval: {e}")
            return
        
        self.monitoring = True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        
        # Send startup message
        startup_msg = f"""🟢 <b>Excel Monitor Started</b>

Monitoring file: {os.path.basename(self.excel_file)}
Check interval: {interval} minute(s)

Looking for:
• Long Buildup signals
• Bullish patterns  
• Short Cover opportunities"""
        
        self.send_telegram_message(startup_msg)
        self.log("🚀 Monitoring started")
        self.update_status()
        
        # Start monitoring thread
        self.monitoring_thread = threading.Thread(target=self.monitoring_loop, args=(interval,), daemon=True)
        self.monitoring_thread.start()
    
    def monitoring_loop(self, interval_minutes):
        while self.monitoring:
            try:
                time.sleep(10)  # Initial delay
                if self.monitoring:  # Check again after sleep
                    self.check_now()
                
                # Sleep for the remaining time
                remaining_time = (interval_minutes * 60) - 10
                for _ in range(int(remaining_time)):
                    if not self.monitoring:
                        break
                    time.sleep(1)
                    
            except Exception as e:
                self.log(f"❌ Monitoring error: {str(e)}")
                time.sleep(60)  # Wait before retrying
    
    def stop_monitoring(self):
        if not self.monitoring:
            return
        
        self.monitoring = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        
        self.send_telegram_message("🔴 <b>Excel Monitor Stopped</b>")
        self.log("✅ Monitoring stopped")
        self.update_status()
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    # Check if required packages are installed
    required_packages = ['requests', 'pandas', 'openpyxl']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("Missing required packages. Please install them:")
        print(f"pip install {' '.join(missing_packages)}")
        exit(1)
    
    app = TelegramExcelMonitor()
    app.run()
