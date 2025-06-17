import streamlit as st
import pandas as pd
import io
import tempfile
import pdfplumber
import re

# --- Core Logic Functions ---

def compute_beds_and_trucks(panels, bed_width=2400, bed_weight_limit=2500, truck_weight_limit=15000, truck_max_length=13620):
    """Takes a list of panels and groups them into beds and trucks."""
    beds = []
    for panel in panels:
        placed = False
        for bed in beds:
            used_depth = sum(p['Depth'] for p in bed)
            total_weight = sum(p['Weight'] for p in bed)
            if used_depth + panel['Depth'] <= bed_width and total_weight + panel['Weight'] <= bed_weight_limit:
                bed.append(panel)
                placed = True
                break
        if not placed:
            beds.append([panel])
    
    bed_summaries = []
    for bed in beds:
        bed_length = max(p['Width'] for p in bed)
        bed_height = max(p['Height'] for p in bed)
        bed_weight = sum(p['Weight'] for p in bed)
        panel_types = ", ".join(str(p['Type']) for p in bed if pd.notna(p['Type']))
        bed_summaries.append({
            'Length': bed_length, 'Height': bed_height, 'Width': bed_width,
            'Weight': bed_weight, 'Num Panels': len(bed), 'Panel Types': panel_types
        })
        
    trucks = []
    for bed in bed_summaries:
        placed = False
        for truck in trucks:
            used_length = sum(b['Length'] for b in truck)
            total_weight = sum(b['Weight'] for b in truck)
            if used_length + bed['Length'] <= truck_max_length and total_weight + bed['Weight'] <= truck_weight_limit:
                truck.append(bed)
                placed = True
                break
        if not placed:
            trucks.append([bed])
            
    return bed_summaries, trucks

def export_to_excel(beds, trucks):
    """Exports the bed and truck data to an in-memory Excel file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame(beds).to_excel(writer, index=False, sheet_name="Beds")
        truck_summary = []
        for i, truck in enumerate(trucks):
            all_panel_types = ", ".join(b['Panel Types'] for b in truck)
            truck_summary.append({
                "Truck #": i + 1, "Num Beds": len(truck), "Total Weight (kg)": sum(b['Weight'] for b in truck),
                "Panel Types": all_panel_types
            })
        pd.DataFrame(truck_summary).to_excel(writer, index=False, sheet_name="Truck Summary")
        summary = pd.DataFrame({"Metric": ["Total Beds", "Total Trucks"], "Value": [len(beds), len(trucks)]})
        summary.to_excel(writer, index=False, sheet_name="Summary")
    output.seek(0)
    return output

def parse_pdf_panels(file_path, spacing=100, thickness=0.016, density=2100, buffer=0.10):
    """Parses panel data from a PDF file."""
    panels = []
    with pdfplumber.open(file_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
    pattern = re.compile(r"(Grc\.[\w\.]+)\s+(\d+)\s+(\d{3,4})\s+(\d{3,4})\s+(\d{3,4})")
    matches = pattern.findall(text)
    def estimate_weight(length_mm, height_mm):
        area_m2 = (length_mm / 1000) * (height_mm / 1000)
        volume_m3 = area_m2 * thickness
        return round(volume_m3 * density * (1 + buffer), 2)
    for match_data in matches:
        try:
            panel_type, qty, height, length, depth = match_data
            for _ in range(int(qty)):
                h = int(height) + 2 * spacing
                l = int(length) + 2 * spacing
                d = int(depth) + 2 * spacing
                weight = estimate_weight(int(length), int(height))
                panels.append({ "Type": panel_type, "Height": d, "Width": l, "Depth": h, "Weight": weight })
        except Exception as e:
            st.error(f"❌ Error parsing PDF match: {e}")
    return panels

def parse_excel_panels(df, spacing, column_map):
    """Parses panel data from a dataframe using a user-provided column map."""
    panels = []
    
    type_col = column_map["panel type"]
    len_col = column_map["length (mm)"]
    hgt_col = column_map["height (mm)"]
    dep_col = column_map["depth (mm)"]
    wgt_col = column_map.get("weight (kg)")

    if not all([type_col, len_col, hgt_col, dep_col]):
        st.error("Error: Please map all required dimension columns (Type, Length, Height, Depth).")
        return []

    for index, row in df.iterrows():
        try:
            panel_type_value = row[type_col]
            if pd.isna(panel_type_value) or str(panel_type_value).strip() == "":
                continue

            l_num = pd.to_numeric(row[len_col], errors='coerce')
            h_num = pd.to_numeric(row[hgt_col], errors='coerce')
            d_num = pd.to_numeric(row[dep_col], errors='coerce')

            if pd.isna(l_num) or pd.isna(h_num) or pd.isna(d_num):
                continue
            
            h, l, d = h_num + 2 * spacing, l_num + 2 * spacing, d_num + 2 * spacing

            weight = 0
            if wgt_col and pd.notna(row[wgt_col]): 
                weight_num = pd.to_numeric(row[wgt_col], errors='coerce')
                if pd.notna(weight_num) and weight_num > 0:
                    weight = weight_num
            
            if weight == 0:
                thickness, density, buffer = 0.016, 2100, 0.10
                area_m2 = (l_num / 1000) * (h_num / 1000)
                volume_m3 = area_m2 * (d_num / 1000 if d_num > 5 else thickness)
                weight = round(volume_m3 * density * (1 + buffer), 2)

            panels.append({ "Type": str(row[type_col]), "Height": d, "Width": l, "Depth": h, "Weight": weight })
        except Exception as e:
            st.error(f"❌ An error occurred on row index {index}: {e}")

    if not panels:
         st.warning("Warning: Could not parse any valid panels.")
    return panels

def display_ui_and_process(df, spacing):
    """Takes a cleaned dataframe and displays the UI for column mapping and analysis."""
    st.header("2. Data Preview")
    st.info("Here are the first 5 rows of your data. Use this to verify the correct header was selected.")
    st.dataframe(df.head())

    df.columns = df.columns.astype(str).str.strip()
    app_columns = [col for col in df.columns if col.strip() != '']

    st.header("3. Map Your Columns")
    st.info("The app will try to guess the correct columns. Please verify them.")
    
    app_columns_lower = [col.lower() for col in app_columns]
    def find_default_index(target_name):
        try: return app_columns_lower.index(target_name)
        except ValueError: return 0

    col1_map, col2_map = st.columns(2)
    with col1_map:
        type_col = st.selectbox("Panel Type/Name Column:", app_columns, index=find_default_index('cast unit'))
        len_col = st.selectbox("Length (mm) Column:", app_columns, index=find_default_index('length, mm'))
        wgt_col = st.selectbox("Weight (kg) Column (Optional):", [None] + app_columns)
    with col2_map:
        hgt_col = st.selectbox("Height (mm) Column:", app_columns, index=find_default_index('height, mm'))
        dep_col = st.selectbox("Depth/Width (mm) Column:", app_columns, index=find_default_index('width, mm'))

    st.header("4. Run Analysis")
    if st.button("Run Analysis with these settings"):
        column_map = {
            "panel type": type_col, "length (mm)": len_col,
            "height (mm)": hgt_col, "depth (mm)": dep_col,
            "weight (kg)": wgt_col,
        }
        panels = parse_excel_panels(df, spacing, column_map)
        
        if panels:
            beds, trucks = compute_beds_and_trucks(panels)
            st.success(f"Parsed {len(panels)} panels, which fit into {len(beds)} beds and {len(trucks)} trucks.")
            
            st.subheader("Bed Summary")
            st.dataframe(pd.DataFrame(beds))
            st.subheader("Truck Summary")
            truck_summary_list = []
            for i, truck in enumerate(trucks):
                all_panel_types = ", ".join(b['Panel Types'] for b in truck)
                truck_summary_list.append({
                    "Truck #": i + 1, "Num Beds": len(truck),
                    "Total Weight (kg)": sum(b['Weight'] for b in truck),
                    "Panel Types": all_panel_types
                })
            st.dataframe(pd.DataFrame(truck_summary_list))

            output = export_to_excel(beds, trucks)
            st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx")


# --- Main Streamlit App ---
st.set_page_config(page_title="GRC Transport Planner", layout="wide")
st.title("🚚 GRC Panel Transport & Storage Estimator")

uploaded_file = st.file_uploader("Upload a data file (CSV, XLSX) or a PDF", type=["csv", "pdf", "xlsx"])
spacing = st.number_input("Panel spacing (mm)", min_value=0, value=100)

if uploaded_file:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    df = None

    if file_extension == "pdf":
        analyze_pdf = st.button("Run PDF Analysis")
        if analyze_pdf:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                panels = parse_pdf_panels(tmp_file_path, spacing)
                if panels:
                    beds, trucks = compute_beds_and_trucks(panels)
                    st.success(f"Parsed {len(panels)} panels, which fit into {len(beds)} beds and {len(trucks)} trucks.")
                    output = export_to_excel(beds, trucks)
                    st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx")
            except Exception as e:
                 st.error(f"Failed to process PDF file: {e}")
    
    elif file_extension in ["csv", "xlsx"]:
        st.header("1. File Settings")
        
        delimiter = ";" # Default for Excel fallback
        if file_extension == "csv":
            delimiter_options = { "Semicolon (;)": ";", "Comma (,)": ",", "Tab": "\t" }
            delimiter_choice = st.selectbox("Column Delimiter:", options=list(delimiter_options.keys()))
            delimiter = delimiter_options[delimiter_choice]

        header_row = st.number_input("Which row number contains the headers? (First row is 0)", min_value=0, value=2)

        try:
            df_raw = None
            if file_extension == "xlsx":
                try:
                    df_raw = pd.read_excel(uploaded_file, header=None)
                except Exception as e:
                    if "Excel file format cannot be determined" in str(e):
                        st.warning("⚠️ This file is not a standard Excel file. Attempting to read as a semicolon-delimited CSV.")
                        uploaded_file.seek(0)
                        df_raw = pd.read_csv(uploaded_file, header=None, sep=';', encoding='utf-8-sig', engine='python')
                    else:
                        raise e
            else: # csv
                df_raw = pd.read_csv(uploaded_file, header=None, sep=delimiter, encoding='utf-8-sig', engine='python')

            # Promote the selected row to header
            new_header = df_raw.iloc[header_row]
            df = df_raw[header_row + 1:].copy()
            df.columns = new_header
            df = df.reset_index(drop=True)
            
            # Remove initial unnamed index column if it exists
            if 'unnamed' in str(df.columns[0]).lower():
                df = df.iloc[:, 1:].copy()

        except Exception as e:
            st.error(f"Error Reading File: {e}")
            st.info("Please ensure the file format, delimiter, and header row number are correct.")
            st.stop()
        
        # Call the common UI function to process the dataframe
        display_ui_and_process(df, spacing)

    else:
        st.error("Unsupported file format.")
