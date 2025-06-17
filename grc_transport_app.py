import streamlit as st
import pandas as pd
import io
import tempfile
import pdfplumber
import re

# --- Core Logic Functions (Unchanged) ---
def compute_beds_and_trucks(panels, bed_width=2400, bed_weight_limit=2500, truck_weight_limit=15000, truck_max_length=13620):
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
        panel_types = [p['Type'] for p in bed]
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
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame(beds).to_excel(writer, index=False, sheet_name="Beds")
        truck_summary = []
        for i, truck in enumerate(trucks):
            truck_summary.append({
                "Truck #": i + 1, "Num Beds": len(truck), "Total Weight (kg)": sum(b['Weight'] for b in truck),
                "Panel Types": ", ".join(str(pt) for b in truck for pt in b["Panel Types"] if pd.notna(pt))
            })
        pd.DataFrame(truck_summary).to_excel(writer, index=False, sheet_name="Truck Summary")
        summary = pd.DataFrame({"Metric": ["Total Beds", "Total Trucks"], "Value": [len(beds), len(trucks)]})
        summary.to_excel(writer, index=False, sheet_name="Summary")
    output.seek(0)
    return output

def parse_pdf_panels(file_path, spacing=100, thickness=0.016, density=2100, buffer=0.10):
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
            panel_type_value = row.get(type_col)
            if pd.isna(panel_type_value) or str(panel_type_value).strip() == "":
                continue

            l_num = pd.to_numeric(row[len_col], errors='coerce')
            h_num = pd.to_numeric(row[hgt_col], errors='coerce')
            d_num = pd.to_numeric(row[dep_col], errors='coerce')

            if pd.isna(l_num) or pd.isna(h_num) or pd.isna(d_num):
                continue
            
            h, l, d = h_num + 2 * spacing, l_num + 2 * spacing, d_num + 2 * spacing

            weight = 0
            if wgt_col: 
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
         st.warning("Warning: Could not parse any valid panels. Please check your column mappings and ensure rows have a panel type and valid numbers for dimensions.")
    return panels


# --- Streamlit App ---
st.set_page_config(page_title="GRC Transport Planner", layout="wide")
st.title("🚚 GRC Panel Transport & Storage Estimator")

uploaded_file = st.file_uploader("Upload a CSV or PDF File", type=["csv", "pdf"])
spacing = st.number_input("Panel spacing (mm)", min_value=0, value=100)

if uploaded_file:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    if file_extension == "pdf":
        analyze_pdf = st.button("Run PDF Analysis")
        if analyze_pdf:
            # PDF logic remains the same
            # ...
    
        elif file_extension == "csv":
        # FIX: Re-introduced the header row selection, as it's critical
        st.header("1. File Settings")
        header_row = st.number_input(
            "Select the row containing column names (the first row is 0):",
            min_value=0, max_value=20, value=2,
            help="This should be the row with names like 'cast unit', 'length, mm', etc."
        )

        df = None
        try:
            # Use encoding='utf-8-sig' to handle CSVs with a BOM
            df = pd.read_csv(uploaded_file, header=header_row, encoding='utf-8-sig')
        except Exception as e:
            st.error(f"Error Reading CSV File: {e}")
            st.info("Please ensure the 'header row' number is correct and the file is a standard comma-separated CSV.")
            st.stop()

        st.header("2. Data Preview")
        st.info("Here are the first 5 rows of your data. Use this to verify the correct header was selected.")
        st.dataframe(df.head())

        # Clean the column names to be safe
        df.columns = df.columns.str.strip()
        app_columns = [col for col in df.columns if col is not None and str(col).strip() != '']

        st.header("3. Map Your Columns")
        st.info("Select which column from your file corresponds to each required data field.")
        
        col1, col2 = st.columns(2)
        with col1:
            type_col = st.selectbox("Panel Type/Name Column:", app_columns)
            len_col = st.selectbox("Length (mm) Column:", app_columns)
            wgt_col = st.selectbox("Weight (kg) Column (Optional):", [None] + app_columns)
        with col2:
            hgt_col = st.selectbox("Height (mm) Column:", app_columns)
            dep_col = st.selectbox("Depth/Width (mm) Column:", app_columns)

        st.header("4. Run Analysis")
        analyze_data = st.button("Run Analysis with these settings")

        if analyze_data:
            column_map = {
                "panel type": type_col, "length (mm)": len_col,
                "height (mm)": hgt_col, "depth (mm)": dep_col,
                "weight (kg)": wgt_col,
            }
            panels = parse_excel_panels(df, spacing, column_map)
            
            if panels:
                beds, trucks = compute_beds_and_trucks(panels)
                st.success(f"Parsed {len(panels)} panels, which fit into {len(beds)} beds and {len(trucks)} trucks.")
                output = export_to_excel(beds, trucks)
                st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx")
    else:
        st.error("Unsupported file format. Please upload a .csv or .pdf file.")
