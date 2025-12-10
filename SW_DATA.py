import streamlit as st
import win32com.client
import win32com.client.dynamic
import pythoncom
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="SolidWorks Batch Extractor",
    page_icon="🏭",
    layout="wide"
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    div.stButton > button:first-child {
        background-color: #d63031;
        color: white;
        border-radius: 8px;
        font-weight: bold;
    }
    .stDataFrame { background-color: white; border-radius: 10px; padding: 10px; }
</style>
""", unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---

def open_files_picker():
    """Opens a Windows file dialog to select MULTIPLE files."""
    try:
        root = tk.Tk()
        root.withdraw() 
        root.wm_attributes('-topmost', 1)
        # askopenfilenames returns a tuple of paths
        file_paths = filedialog.askopenfilenames(
            title="Select SolidWorks Files (Batch)",
            filetypes=[("SolidWorks Files", "*.sldprt;*.sldasm")]
        )
        root.destroy()
        return file_paths
    except Exception:
        return []

def get_custom_property(model, prop_name):
    """Safely extracts a custom property."""
    val = ""
    try:
        # 1. Try Active Configuration
        config = model.GetActiveConfiguration()
        cus_prop_mgr = config.CustomPropertyManager
        status, res_val, resolved_val = cus_prop_mgr.Get4(prop_name, False, "", "")
        
        if resolved_val.strip():
            return resolved_val

        # 2. Try General "Custom" Tab
        cus_prop_mgr = model.Extension.CustomPropertyManager("")
        status, res_val, resolved_val = cus_prop_mgr.Get4(prop_name, False, "", "")
        
        if resolved_val.strip():
            return resolved_val
    except:
        pass
    return ""

def get_material(model, doc_type):
    """Extracts material name. Only works reliably for Parts."""
    if doc_type != 1: 
        return "N/A (Assembly)"
    
    try:
        # 1. Try custom property "Material"
        mat_prop = get_custom_property(model, "Material")
        if mat_prop and "SW-Material" not in mat_prop:
            return mat_prop

        # 2. Use API to get internal material
        config_name = model.GetActiveConfiguration().Name
        db_xml, mat_name = model.GetMaterialPropertyName2(config_name, "", "")
        
        if mat_name:
            return mat_name
    except:
        pass
    return "Not Specified"

# --- MAIN LOGIC ---
def process_files(file_list, get_mass, get_dims, get_sm, get_props):
    pythoncom.CoInitialize()
    
    results_list = []
    
    try:
        try:
            swApp = win32com.client.GetActiveObject("SldWorks.Application")
        except:
            st.error("SolidWorks is not running. Please open it first.")
            return []

        # UI Progress
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, file_path in enumerate(file_list):
            file_name = os.path.basename(file_path)
            status_text.text(f"Processing ({i+1}/{len(file_list)}): {file_name}")
            progress_bar.progress((i + 1) / len(file_list))
            
            # --- OPEN FILE ---
            ext = os.path.splitext(file_path)[1].lower()
            doc_type = 1 if ext == ".sldprt" else 2
            
            swApp.DocumentVisible(False, doc_type)
            
            arg_err = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            arg_warn = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            
            raw_model = swApp.OpenDoc6(file_path, doc_type, 3, "", arg_err, arg_warn)

            if not raw_model:
                results_list.append({
                    "File Name": file_name,
                    "Status": "Failed to Open"
                })
                continue

            # Force Dynamic
            model = win32com.client.dynamic.Dispatch(raw_model)
            
            # --- EXTRACT DATA ---
            row = {
                "File Name": file_name,
                "Status": "Success",
                "Path": file_path
            }

            # 1. Properties
            if get_props:
                row["Part Number"] = get_custom_property(model, "Part Number")
                row["Description"] = get_custom_property(model, "Description")
                row["Revision"] = get_custom_property(model, "Revision")
            
            if get_props and doc_type == 1:
                row["Material"] = get_material(model, doc_type)
            else:
                row["Material"] = ""

            # 2. Mass Properties (Converted to Grams / mm^3 / mm^2)
            if get_mass:
                try:
                    mass_prop = model.Extension.CreateMassProperty()
                    mass_prop.UseSystemUnits = False # Returns SI (kg, m^3, m^2)
                    mass_prop.UpdateMassProperties()
                    
                    # Apply Conversions
                    # Mass: kg -> g (x1000)
                    row["Mass (g)"] = round(mass_prop.Mass * 1000.0, 2)
                    # Volume: m^3 -> mm^3 (x1,000,000,000)
                    row["Volume (mm^3)"] = round(mass_prop.Volume * 1000000000.0, 2)
                    # Area: m^2 -> mm^2 (x1,000,000)
                    row["Surface Area (mm^2)"] = round(mass_prop.SurfaceArea * 1000000.0, 2)
                    
                except:
                    # Legacy Fallback
                    try:
                        vals = model.GetMassProperties
                        if callable(vals): vals = vals()
                        
                        if vals:
                            # vals[3] = Volume (m^3), vals[4] = Area (m^2), vals[5] = Mass (kg)
                            row["Volume (mm^3)"] = round(vals[3] * 1000000000.0, 2)
                            row["Surface Area (mm^2)"] = round(vals[4] * 1000000.0, 2)
                            row["Mass (g)"] = round(vals[5] * 1000.0, 2)
                        else:
                            row["Mass (g)"] = 0
                            row["Volume (mm^3)"] = 0
                            row["Surface Area (mm^2)"] = 0
                    except:
                        row["Mass (g)"] = 0
                        row["Volume (mm^3)"] = 0
                        row["Surface Area (mm^2)"] = 0

            # 3. Sheet Metal
            if get_sm and doc_type == 1:
                is_sm = False
                thk = 0.0
                try:
                    fm = model.FeatureManager
                    fm_dyn = win32com.client.dynamic.Dispatch(fm)
                    all_feats = fm_dyn.GetFeatures(False) 
                    if all_feats:
                        for f_obj in all_feats:
                            f_dyn = win32com.client.dynamic.Dispatch(f_obj)
                            try:
                                tn = f_dyn.GetTypeName2
                                if callable(tn): tn = tn()
                            except: tn = ""
                            
                            if tn == "SheetMetal":
                                is_sm = True
                                try:
                                    fd = f_dyn.GetDefinition()
                                except: fd = f_dyn.GetDefinition
                                if fd:
                                    fd_dyn = win32com.client.dynamic.Dispatch(fd)
                                    thk = fd_dyn.Thickness
                                break
                except:
                    pass
                row["Is Sheet Metal"] = is_sm
                row["Thickness (mm)"] = round(thk * 1000.0, 3) if is_sm else 0

            # 4. Dimensions
            if get_dims:
                try:
                    box = None
                    if doc_type == 1:
                        try: box = model.GetPartBox(True)
                        except: 
                            conf = model.GetActiveConfiguration()
                            if conf:
                                c_d = win32com.client.dynamic.Dispatch(conf)
                                box = c_d.GetBox()
                    else:
                        box = model.GetBox(0)
                        
                    if box:
                        dims = sorted([abs(box[3]-box[0]), abs(box[4]-box[1]), abs(box[5]-box[2])])
                        row["Length (mm)"] = round(dims[2] * 1000.0, 2)
                        row["Width (mm)"] = round(dims[1] * 1000.0, 2)
                        row["Height (mm)"] = round(dims[0] * 1000.0, 2)
                except:
                    pass

            # Close
            try:
                title = model.GetTitle()
            except: title = model.GetTitle
            
            swApp.CloseDoc(title)
            results_list.append(row)
            swApp.DocumentVisible(True, doc_type)

        status_text.success("Batch Processing Complete!")
        progress_bar.empty()
        
    except Exception as e:
        st.error(f"Critical Error: {e}")
        try:
            swApp.DocumentVisible(True, 1)
            swApp.DocumentVisible(True, 2)
        except: pass
        
    return results_list

# --- FRONTEND UI ---

st.title("🏭 Batch SW Data Extractor")
st.markdown("Process multiple SolidWorks files and export data to Excel/CSV.")

# --- HELPER DROPDOWN ---
with st.expander("ℹ️ **User Guide: Inputs & Outputs Explained**"):
    st.markdown("""
    ### 📥 Inputs
    *   **File Selection:** Select one or multiple SolidWorks files. 
        *   Supports: `.SLDPRT` (Parts) and `.SLDASM` (Assemblies).
    *   **Configuration:** Use the Sidebar checkboxes to speed up processing by turning off data you don't need.

    ### 📤 Outputs Explained
    | Data Point | Units | Description |
    | :--- | :--- | :--- |
    | **Mass** | Grams ($g$) | Calculated based on material density. |
    | **Volume** | $mm^3$ | Total volume of the geometry. |
    | **Surface Area** | $mm^2$ | Total external surface area. |
    | **Dimensions** | $mm$ | Bounding Box: Length x Width x Height. *(Note: Aligned to XYZ planes)* |
    | **Sheet Metal** | $mm$ | Thickness (if valid Sheet Metal feature exists). |
    | **Properties** | Text | Extracts 'Part Number', 'Description', 'Revision', 'Material'. |

    ### ⚠️ Requirements
    1.  **SolidWorks** must be installed and running.
    2.  Files must be saved in your current SolidWorks version for best results.
    """)

if 'file_paths' not in st.session_state:
    st.session_state['file_paths'] = []
if 'data_df' not in st.session_state:
    st.session_state['data_df'] = None

# Sidebar Config
st.sidebar.header("Data to Extract")
opt_props = st.sidebar.checkbox("Properties (Desc, Rev, Material)", value=True)
opt_mass = st.sidebar.checkbox("Mass Properties (Mass, Vol, Area)", value=True)
opt_dims = st.sidebar.checkbox("Dimensions (L x W x H)", value=True)
opt_sm = st.sidebar.checkbox("Sheet Metal Thickness", value=True)

col1, col2 = st.columns([4, 1])
with col1:
    num_files = len(st.session_state['file_paths'])
    st.info(f"📂 Selected Files: {num_files}")
    if num_files > 0:
        with st.expander("View Selected Files"):
            st.write(st.session_state['file_paths'])

with col2:
    st.write("")
    if st.button("Browse Files"):
        paths = open_files_picker()
        if paths:
            st.session_state['file_paths'] = paths
            st.rerun()

if st.button("🚀 Run Batch Analysis", type="primary", disabled=(num_files == 0)):
    with st.spinner("Starting SolidWorks..."):
        data = process_files(
            st.session_state['file_paths'], 
            opt_mass, opt_dims, opt_sm, opt_props
        )
        
    if data:
        df = pd.DataFrame(data)
        
        # Order columns neatly
        desired_order = [
            "File Name", "Part Number", "Description", "Revision", "Material",
            "Mass (g)", "Volume (mm^3)", "Surface Area (mm^2)", 
            "Length (mm)", "Width (mm)", "Height (mm)", 
            "Is Sheet Metal", "Thickness (mm)", "Status"
        ]
        cols = [c for c in desired_order if c in df.columns] + [c for c in df.columns if c not in desired_order]
        df = df[cols]
        
        st.session_state['data_df'] = df

if st.session_state['data_df'] is not None:
    st.subheader("📊 Analysis Results")
    
    df = st.session_state['data_df']
    st.dataframe(df, use_container_width=True)
    
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Results as CSV",
        data=csv,
        file_name="SW_Batch_Results.csv",
        mime="text/csv"
    )