import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.set_page_config(layout="wide", page_title="Universal XL Aligner")

st.title("🚀 Universal XL Aligner: Vertical & Horizontal")

# --- 1. FILE UPLOAD ---
st.sidebar.header("Step 1: Upload Files")
uploaded_files = st.sidebar.file_uploader(
    "Upload your Excel files", 
    type=["xlsx"], 
    accept_multiple_files=True
)

# Map filenames to file objects for easy lookup
file_map = {f.name: f for f in uploaded_files} if uploaded_files else {}
all_filenames = list(file_map.keys())

if not uploaded_files:
    st.info("👈 Please upload Excel files in the sidebar to begin.")
    st.stop()

# --- 2. ALIGNMENT SETTINGS ---
st.sidebar.header("Step 2: Align Mode")
mode = st.sidebar.radio("Direction", ["Vertical (Across Files)", "Horizontal (Across Sheets)"])

master_filename = st.sidebar.selectbox("Select Master File", all_filenames)
master_file_obj = file_map[master_filename]

# Use pd.ExcelFile to read sheet names without loading the whole data yet
xl_master = pd.ExcelFile(master_file_obj)
all_sheets = xl_master.sheet_names

if mode == "Vertical (Across Files)":
    target_sheet = st.sidebar.selectbox("Select Sheet to align", all_sheets)
    df_master = pd.read_excel(master_file_obj, sheet_name=target_sheet)
    compare_targets = [f_name for f_name in all_filenames if f_name != master_filename]
else:
    master_sheet = st.sidebar.selectbox("Select Master Sheet", all_sheets)
    df_master = pd.read_excel(master_file_obj, sheet_name=master_sheet)
    compare_targets = [s for s in all_sheets if s != master_sheet]

# --- 3. MAIN DISPLAY ---
st.write(f"### Master Grid: {master_filename}")
event = st.dataframe(df_master, on_select="rerun", selection_mode="single-cell", width="stretch")

# --- 4. ACTION LOGIC ---
if event.selection.get("cells"):
    selected = event.selection["cells"][0]
    
    # Extract Row/Col (handles both dict and list return types from Streamlit)
    r = selected.get('row') if isinstance(selected, dict) else selected[0]
    c_raw = selected.get('column') if isinstance(selected, dict) else selected[1]
    
    # Handle column index vs column name
    if isinstance(c_raw, str):
        c = df_master.columns.get_loc(c_raw)
    else:
        c = c_raw
    
    master_val = df_master.iloc[r, c]
    
    st.divider()
    st.subheader(f"Results for Cell [{r}, {c}]")
    st.info(f"Master Value: **{master_val}**")

    results = []

    for target in compare_targets:
        try:
            if mode == "Vertical (Across Files)":
                # target is a filename; get file object from map
                target_file_obj = file_map[target]
                df_temp = pd.read_excel(target_file_obj, sheet_name=target_sheet)
                label = f"File: {target}"
            else:
                # target is a sheet name within the same master file
                df_temp = pd.read_excel(master_file_obj, sheet_name=target)
                label = f"Sheet: {target}"
            
            comp_val = df_temp.iloc[r, c]
            status = "✅ Match" if str(comp_val) == str(master_val) else "❌ Mismatch"
        except Exception as e:
            comp_val = "N/A"
            status = "⚠️ Missing/Error"
            
        results.append({mode.split()[0]: label, "Value": comp_val, "Status": status})

    st.table(pd.DataFrame(results))
else:
    st.info(f"Click a cell in the table above to run **{mode}** alignment.")
