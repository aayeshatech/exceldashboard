import pandas as pd
import streamlit as st
import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import io

st.set_page_config(page_title="Options Dashboard", page_icon="📊", layout="wide")

def read_xlsx_without_openpyxl(file_path):
    """Read XLSX file without openpyxl by parsing the XML directly"""
    try:
        data_dict = {}
        
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            # Read workbook.xml to get sheet names
            try:
                workbook_xml = zip_file.read('xl/workbook.xml')
                root = ET.fromstring(workbook_xml)
                
                sheets = []
                for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                    sheet_id = sheet.get('sheetId')
                    sheet_name = sheet.get('name')
                    sheets.append((sheet_id, sheet_name))
                
                st.info(f"Found {len(sheets)} sheets")
                
                # Read shared strings
                shared_strings = []
                try:
                    shared_strings_xml = zip_file.read('xl/sharedStrings.xml')
                    ss_root = ET.fromstring(shared_strings_xml)
                    for si in ss_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                        t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                        shared_strings.append(t.text if t is not None else '')
                except:
                    pass
                
                # Read each sheet
                for sheet_id, sheet_name in sheets:
                    try:
                        sheet_xml = zip_file.read(f'xl/worksheets/sheet{sheet_id}.xml')
                        sheet_root = ET.fromstring(sheet_xml)
                        
                        rows_data = []
                        for row in sheet_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                            row_data = []
                            for cell in row.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                                cell_value = ""
                                v = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                                if v is not None:
                                    if cell.get('t') == 's':  # Shared string
                                        try:
                                            cell_value = shared_strings[int(v.text)]
                                        except:
                                            cell_value = v.text
                                    else:
                                        cell_value = v.text
                                row_data.append(cell_value)
                            if row_data:
                                rows_data.append(row_data)
                        
                        if rows_data:
                            # Convert to DataFrame
                            max_cols = max(len(row) for row in rows_data) if rows_data else 0
                            
                            # Pad shorter rows
                            for row in rows_data:
                                while len(row) < max_cols:
                                    row.append("")
                            
                            if len(rows_data) > 0:
                                headers = rows_data[0] if rows_data else []
                                data_rows = rows_data[1:] if len(rows_data) > 1 else []
                                
                                df = pd.DataFrame(data_rows, columns=headers)
                                
                                # Convert numeric columns
                                for col in df.columns:
                                    df[col] = pd.to_numeric(df[col], errors='ignore')
                                
                                data_dict[sheet_name] = df
                                st.success(f"Loaded {sheet_name}: {len(df)} rows")
                    
                    except Exception as e:
                        st.warning(f"Could not read sheet {sheet_name}: {str(e)}")
                        continue
                
                return data_dict
                
            except Exception as e:
                st.error(f"Error parsing workbook: {str(e)}")
                return {}
    
    except Exception as e:
        st.error(f"Error reading XLSX file: {str(e)}")
        return {}

def try_pandas_engines(file_path):
    """Try different pandas engines"""
    engines = [None, 'openpyxl', 'xlrd']
    
    for engine in engines:
        try:
            if engine:
                excel_file = pd.ExcelFile(file_path, engine=engine)
            else:
                excel_file = pd.ExcelFile(file_path)
            
            data_dict = {}
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
                    if not df.empty:
                        data_dict[sheet_name] = df
                        st.success(f"Loaded {sheet_name} with {engine or 'default'} engine")
                except Exception as e:
                    st.warning(f"Sheet {sheet_name} failed with {engine or 'default'}: {str(e)}")
            
            if data_dict:
                return data_dict
                
        except Exception as e:
            st.warning(f"Engine {engine or 'default'} failed: {str(e)}")
            continue
    
    return {}

def read_excel_file(file_path):
    """Read Excel file with fallback methods"""
    file_ext = file_path.lower().split('.')[-1]
    
    # First try pandas with available engines
    st.info("Trying pandas engines...")
    data_dict = try_pandas_engines(file_path)
    
    if data_dict:
        return data_dict
    
    # If pandas fails and it's an xlsx file, try manual parsing
    if file_ext == 'xlsx':
        st.info("Trying manual XLSX parsing...")
        return read_xlsx_without_openpyxl(file_path)
    
    # For xlsm files, convert file path suggestion
    elif file_ext == 'xlsm':
        st.error("XLSM files require openpyxl. Please save your file as .xlsx format and try again.")
        st.info("In Excel: File > Save As > Choose 'Excel Workbook (.xlsx)'")
        return {}
    
    return {}

def calculate_options_metrics(df):
    """Calculate basic options metrics"""
    metrics = {}
    
    # Find relevant columns (case insensitive)
    call_oi_cols = [col for col in df.columns if 'CE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
    put_oi_cols = [col for col in df.columns if 'PE' in str(col).upper() and 'OI' in str(col).upper() and 'CHANGE' not in str(col).upper()]
    
    if call_oi_cols and put_oi_cols:
        try:
            call_oi = pd.to_numeric(df[call_oi_cols[0]], errors='coerce').fillna(0).sum()
            put_oi = pd.to_numeric(df[put_oi_cols[0]], errors='coerce').fillna(0).sum()
            
            if call_oi > 0:
                pcr = put_oi / call_oi
                metrics = {
                    'call_oi': call_oi,
                    'put_oi': put_oi,
                    'pcr': pcr
                }
        except Exception as e:
            st.warning(f"Error calculating metrics: {e}")
    
    return metrics

def main():
    st.title("Options Dashboard")
    
    # Check what's available
    st.sidebar.header("System Status")
    
    # Check pandas engines
    engines_status = {}
    for engine in ['openpyxl', 'xlrd']:
        try:
            __import__(engine)
            engines_status[engine] = "✅ Available"
        except ImportError:
            engines_status[engine] = "❌ Missing"
    
    for engine, status in engines_status.items():
        st.sidebar.write(f"{engine}: {status}")
    
    if not any("✅" in status for status in engines_status.values()):
        st.sidebar.warning("No Excel engines available - using manual parsing")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        # Check file type
        file_ext = uploaded_file.name.lower().split('.')[-1]
        st.info(f"File type: {file_ext}")
        
        if file_ext == 'xlsm':
            st.warning("XLSM files may not work without openpyxl. Consider saving as .xlsx format.")
        
        # Save temporarily
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Read file
        with st.spinner("Reading Excel file..."):
            data_dict = read_excel_file(temp_path)
        
        # Clean up
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data_dict:
            st.success(f"Successfully loaded {len(data_dict)} sheets")
            
            # Sheet selector
            sheet_names = list(data_dict.keys())
            if len(sheet_names) > 1:
                selected_sheet = st.selectbox("Select Sheet to Analyze", sheet_names)
            else:
                selected_sheet = sheet_names[0]
                st.info(f"Analyzing sheet: {selected_sheet}")
            
            if selected_sheet in data_dict:
                df = data_dict[selected_sheet]
                
                # Basic info
                st.subheader("Data Overview")
                col1, col2, col3 = st.columns(3)
                col1.metric("Rows", len(df))
                col2.metric("Columns", len(df.columns))
                col3.metric("Non-empty cells", df.count().sum())
                
                # Show columns
                with st.expander("Column Names"):
                    for i, col in enumerate(df.columns, 1):
                        st.write(f"{i}. {col}")
                
                # Calculate metrics
                metrics = calculate_options_metrics(df)
                
                if metrics:
                    st.subheader("Options Metrics")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Call OI", f"{int(metrics['call_oi']):,}")
                    col2.metric("Put OI", f"{int(metrics['put_oi']):,}")
                    col3.metric("PCR Ratio", f"{metrics['pcr']:.3f}")
                    
                    # Market sentiment
                    pcr = metrics['pcr']
                    if pcr > 1.3:
                        st.error("📉 Bearish Sentiment (High PCR - More Puts)")
                    elif pcr < 0.7:
                        st.success("📈 Bullish Sentiment (Low PCR - More Calls)")
                    else:
                        st.warning("⚖️ Neutral Sentiment (Balanced)")
                else:
                    st.info("No options columns detected for metric calculation")
                
                # Data display
                tab1, tab2 = st.tabs(["Sample Data", "Full Dataset"])
                
                with tab1:
                    st.subheader("First 20 Rows")
                    st.dataframe(df.head(20), use_container_width=True)
                
                with tab2:
                    st.subheader("Complete Dataset")
                    st.dataframe(df, use_container_width=True)
                    
                    # Download option
                    csv_data = df.to_csv(index=False)
                    st.download_button(
                        "Download as CSV",
                        csv_data,
                        f"{selected_sheet}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        "text/csv"
                    )
        else:
            st.error("Could not read any data from the Excel file")
            
            if uploaded_file.name.lower().endswith('.xlsm'):
                st.info("💡 Try saving your file as .xlsx format in Excel and upload again")
    
    else:
        st.info("Upload your Excel file to begin analysis")
        
        # Show what file formats work
        st.subheader("Supported Formats")
        st.write("✅ .xlsx files (best compatibility)")
        st.write("⚠️ .xlsm files (may require openpyxl)")
        st.write("⚠️ .xls files (may require xlrd)")
        
        st.subheader("Troubleshooting")
        st.write("If your file doesn't load:")
        st.write("1. Save as .xlsx format in Excel")
        st.write("2. Ensure the file isn't corrupted")
        st.write("3. Check that sheets contain data")

if __name__ == "__main__":
    main()
