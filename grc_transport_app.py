import streamlit as st
import pandas as pd
import io
import tempfile
import pdfplumber
import re
import xlsxwriter
import difflib

# --- Parsing Logic ---
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
                panels.append({
                    "Type": panel_type,
                    "Height": d,
                    "Width": l,
                    "Depth": h,
                    "Weight": weight
                })
        except Exception as e:
            st.error(f"❌ Error parsing PDF match: {e}")
            st.write("🚨 Problematic match data:", match_data)
    return panels


def fuzzy_match_column(df_columns, target_keywords):
    for target in target_keywords:
        match = difflib.get_close_matches(target, df_columns, n=1, cutoff=0.6)
        if match:
            return match[0]
    return None

# FIX: The function now accepts a DataFrame directly, not a file path.
# This makes it independent of the file type (Excel or CSV).
def parse_excel_panels(df, spacing=100):
    # The first step is to clean the DataFrame that was passed in.
    df.replace(to_replace=r"\s+", value="", regex=True, inplace=True)
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df.columns = df.columns.str.strip().str.lower()
    colnames = df.columns.tolist()
    st.write("Detected columns:", colnames)

    column_map = {}
    targets = {
        "panel type": ["panel type", "type", "cast unit", "cast_unit", "name"],
        "height (mm)": ["height", "height (mm)", "augstums"],
        "length (mm)": ["length", "length (mm)", "garums"],
        "depth (mm)": ["depth", "depth (mm)", "platums", "width"],
        "weight (kg)": ["weight", "weight (kg)", "svars"],
        "dimensions": ["dimensions", "dimensions (l*h*w), mm", "izmēri"]
    }

    required_keys = ["panel type", "height (mm)", "length (mm)", "depth (mm)"]
    optional_keys = ["weight (kg)", "dimensions"]

    for key in required_keys:
        match = fuzzy_match_column(colnames, targets[key])
        if match: column_map[key] = match
    
    for key in optional_keys:
        match = fuzzy_match_column(colnames, targets[key])
        if match: column_map[key] = match

    panels = []
    # Using df.iterrows() on the passed-in DataFrame
    for index, row in df.iterrows():
        try:
            h_num, l_num, d_num = None, None, None

            # METHOD 1: Try individual dimensions
            if all(k in column_map for k in ["height (mm)", "length (mm)", "depth (mm)"]):
                h_cand = pd.to_numeric(row.get(column_map["height (mm)"]), errors='coerce')
                l_cand = pd.to_numeric(row.get(column_map["length (mm)"]), errors='coerce')
                d_cand = pd.to_numeric(row.get(column_map["depth (mm)"]), errors='coerce')
                if pd.notna(h_cand) and pd.notna(l_cand) and pd.notna(d_cand):
                    h_num, l_num, d_num = h_cand, l_cand, d_cand

            # METHOD 2: Fallback to combined 'dimensions' column
            if h_num is None and "dimensions" in column_map:
                dim_str = str(row.get(column_map["dimensions"]))
                parts = re.findall(r'(\d+\.?\d*)', dim_str)
                if len(parts) == 3:
                    l_num, h_num, d_num = float(parts[0]), float(parts[1]), float(parts[2])
            
            if h_num is None or l_num is None or d_num is None:
                continue

            h = h_num + 2 * spacing
            l = l_num + 2 * spacing
            d = d_num + 2 * spacing
            
            weight = 0
            weight_val = row.get(column_map.get("weight (kg)"))
            weight_num = pd.to_numeric(weight_val, errors='coerce')
            
            if pd.notna(weight_num) and weight_num > 0:
                weight = weight_num
            else:
                thickness, density, buffer = 0.016, 2100, 0.10
                area_m2 = (l_num / 1000) * (h_num / 1000)
                volume_m3 = area_m2 * (d_num / 1000 if d_num > 5 else thickness)
                weight = round(volume_m3 * density * (1 + buffer), 2)

            panels.append({
                "Type": str(row.get(column_map.get("panel type", "Unknown"))),
                "Height": d, "Width": l, "Depth": h, "Weight": weight
            })
        except Exception as e:
            st.error(f"❌ Error processing row at index {index}: {e}")
            st.write("🚨 Problematic row data:", row.to_dict())
            
    if not panels:
         st.warning("Could not parse any valid panels from the file. Please check the file format and header row selection.")
        
    return panels


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

# --- Streamlit App ---
st.set_page_config(page_title="GRC Transport Planner", layout="wide")
st.title("🚚 GRC Panel Transport & Storage Estimator")

# FIX: Add 'csv' to the list of accepted file types
uploaded_file = st.file_uploader("Upload a PDF, Excel, or CSV File", type=["pdf", "xlsx", "csv"])
spacing = st.number_input("Panel spacing (mm)", min_value=0, value=100)

if uploaded_file:
    # FIX: New logic to handle different file types (Excel, CSV, PDF)
    file_extension = uploaded_file.name.split('.')[-1].lower()
    df = None  # Initialize DataFrame

    try:
        # --- Handle Tabular Data (Excel or CSV) ---
        if file_extension in ["xlsx", "csv"]:
            st.subheader("File Preview & Settings")
            header_row = st.number_input("Select the header row (0-indexed, first row is 0)", min_value=0, max_value=20, value=2)
            
            if file_extension == "xlsx":
                df = pd.read_excel(uploaded_file, header=header_row)
            else: # Must be "csv"
                df = pd.read_csv(uploaded_file, header=header_row)
            
            st.dataframe(df.head()) # Show a preview of the loaded data
            
            analyze_data = st.button("Run Analysis")
            if analyze_data:
                # Pass the DataFrame directly to the parser
                panels = parse_excel_panels(df, spacing)
                if panels: # Only proceed if panels were successfully parsed
                    beds, trucks = compute_beds_and_trucks(panels)
                    st.success(f"Parsed {len(panels)} panels, which fit into {len(beds)} beds and {len(trucks)} trucks.")
                    output = export_to_excel(beds, trucks)
                    st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # --- Handle PDF Data ---
        elif file_extension == "pdf":
            analyze_pdf = st.button("Run PDF Analysis")
            if analyze_pdf:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                
                panels = parse_pdf_panels(tmp_file_path, spacing)
                beds, trucks = compute_beds_and_trucks(panels)
                st.success(f"Parsed {len(panels)} panels, which fit into {len(beds)} beds and {len(trucks)} trucks.")
                output = export_to_excel(beds, trucks)
                st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Unsupported file type. Please upload a PDF, XLSX, or CSV file.")

    except Exception as e:
        st.error(f"An error occurred during processing: {e}")
        st.info("Please check that the selected header row is correct and the file format is not corrupted.")
