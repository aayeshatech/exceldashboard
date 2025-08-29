import pandas as pd
import streamlit as st
import subprocess
import sys
import os
from datetime import datetime

st.set_page_config(page_title="Options Dashboard", page_icon="📊", layout="wide")

def install_package(package):
    """Install a package using pip"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except Exception as e:
        st.error(f"Failed to install {package}: {e}")
        return False

def check_and_install_dependencies():
    """Check for required packages and install if missing"""
    required_packages = ['openpyxl', 'xlrd']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        st.warning(f"Missing packages: {missing_packages}")
        
        if st.button("Install Missing Packages"):
            with st.spinner("Installing packages..."):
                success = True
                for package in missing_packages:
                    if not install_package(package):
                        success = False
                
                if success:
                    st.success("Packages installed successfully! Please restart the app.")
                    st.stop()
                else:
                    st.error("Some packages failed to install. Please install manually:")
                    st.code(f"pip install {' '.join(missing_packages)}")
                    st.stop()
        else:
            st.info("Click the button above to install missing packages, or install manually:")
            st.code(f"pip install {' '.join(missing_packages)}")
            st.stop()
    
    return True

def read_excel_safely(file_path):
    """Try to read Excel with different engines"""
    engines = ['openpyxl', 'xlrd', None]
    
    for engine in engines:
        try:
            if engine:
                excel_file = pd.ExcelFile(file_path, engine=engine)
            else:
                excel_file = pd.ExcelFile(file_path)
            
            sheets = {}
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
                    if not df.empty:
                        sheets[sheet_name] = df
                        st.success(f"Loaded {sheet_name}: {len(df)} rows")
                except Exception as e:
                    st.warning(f"Skipped {sheet_name}: {e}")
            
            return sheets
            
        except Exception as e:
            if engine:
                st.warning(f"Engine {engine} failed: {e}")
                continue
            else:
                st.error(f"All engines failed: {e}")
                return {}
    
    return {}

def main():
    st.title("Options Dashboard")
    
    # Check dependencies first
    check_and_install_dependencies()
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xlsm', 'xls'])
    
    if uploaded_file:
        # Save temp file
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Read Excel
        with st.spinner("Reading Excel file..."):
            data = read_excel_safely(temp_path)
        
        # Clean up
        try:
            os.remove(temp_path)
        except:
            pass
        
        if data:
            st.success(f"Successfully loaded {len(data)} sheets")
            
            # Sheet selector
            sheet_names = list(data.keys())
            if len(sheet_names) > 1:
                selected_sheet = st.selectbox("Select Sheet", sheet_names)
            else:
                selected_sheet = sheet_names[0]
            
            # Display data
            if selected_sheet in data:
                df = data[selected_sheet]
                
                st.subheader(f"Sheet: {selected_sheet}")
                st.write(f"Rows: {len(df)}, Columns: {len(df.columns)}")
                
                # Show sample data
                st.dataframe(df.head(10))
                
                # Basic analysis if options columns found
                call_cols = [col for col in df.columns if 'CE' in col.upper() and 'OI' in col.upper()]
                put_cols = [col for col in df.columns if 'PE' in col.upper() and 'OI' in col.upper()]
                
                if call_cols and put_cols:
                    st.subheader("Options Analysis")
                    
                    call_oi = df[call_cols[0]].sum()
                    put_oi = df[put_cols[0]].sum()
                    pcr = put_oi / call_oi if call_oi > 0 else 0
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Call OI", f"{call_oi:,.0f}")
                    col2.metric("Put OI", f"{put_oi:,.0f}")
                    col3.metric("PCR", f"{pcr:.3f}")
                
                # Full data view
                with st.expander("Full Data"):
                    st.dataframe(df)
        else:
            st.error("No data could be loaded from the Excel file")
    
    else:
        st.info("Please upload an Excel file to get started")
        
        # Show installation status
        st.subheader("System Status")
        
        packages_status = {}
        for package in ['openpyxl', 'xlrd']:
            try:
                __import__(package)
                packages_status[package] = "✅ Installed"
            except ImportError:
                packages_status[package] = "❌ Missing"
        
        for package, status in packages_status.items():
            st.write(f"{package}: {status}")

if __name__ == "__main__":
    main()
