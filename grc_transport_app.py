import streamlit as st
import pandas as pd
import io
import tempfile
import pdfplumber
import re
import xlsxwriter
import difflib

# --- Parsing Logic ---
# (Functions outside of parse_excel_panels are unchanged)
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
                    "Type": panel_type, "Height": d, "Width": l, "Depth": h, "Weight": weight
                })
        except Exception as e:
            st.error(f"❌ Error parsing PDF match: {e}")
    return panels

def fuzzy_match_column(df_columns, target_keywords):
    for target in target_keywords:
        match = difflib.get_close_matches(target, df_columns, n=1, cutoff=0.6)
        if match:
            return match[0]
    return None

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


# --- SIMPLIFIED PARSER per user request ---
def parse_excel_panels(df, spacing=100):
    # This simplified version only looks for single numeric values in the dimension columns.
    # It will skip rows with multi-part dimension strings.
    df.columns = df.columns.str.strip().str.lower()
    colnames = df.columns.tolist()

    column_map = {}
    # The targets no longer include fallbacks like 'name' or 'dimensions'
    targets = {
        "panel type": ["panel type", "type", "cast unit", "cast_unit"],
        "height (mm)": ["height", "height (mm)", "augstums"],
        "length (mm)": ["length", "length (mm)", "garums"],
        "depth (mm)": ["depth", "depth (mm)", "platums", "width"],
        "weight (kg)": ["weight", "weight (kg)", "svars"],
    }
    
    for key, potential_names in targets.items():
        match = fuzzy_match_column(colnames, potential_names)
        if match: column_map[key] = match

    panels = []
    for index, row in df.iterrows():
        try:
            # Directly try to convert the essential dimension columns to numbers
            l_num = pd.to_numeric(row.get(column_map.get("length (mm)")), errors='coerce')
            h_num = pd.to_numeric(row.get(column_map.get("height (mm)")), errors='coerce')
            d_num = pd.to_numeric(row.get(column_map.get("depth (mm)")), errors='coerce')

            # If any dimension is not a valid number, skip this row.
            if pd.isna(l_num) or pd.isna(h_num) or pd.isna(d_num):
                continue
            
            h = h_num + 2 * spacing
            l = l_num + 2 * spacing
            d = d_num + 2 * spacing

            # --- Weight Calculation ---
            weight = 0
            weight_num = pd.to_numeric(row.get(column_map.get("weight (kg)")), errors='coerce')
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
            st.error(f"❌ An error occurred on row index {index}: {e}")

    if not panels:
         st.warning("Warning: Could not parse any valid panel data. This may be because no rows contained individual numbers for height, length, and width within columns A-I.")
    return panels


# --- Streamlit App ---
st.set_page_config(page_title="GRC Transport Planner", layout="wide")
st.title("🚚 GRC Panel Transport & Storage Estimator")

uploaded_file = st.file_uploader("Upload a PDF, Excel, or CSV File", type=["pdf", "xlsx", "csv"])
spacing = st.number_input("Panel spacing (mm)", min_value=0, value=100)

if uploaded_file:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    df = None
    try:
        if file_extension in ["xlsx", "csv"]:
            st.subheader("File Settings")
            header_row = st.number_input("Select the header row (the first row is 0)", min_value=0, max_value=20, value=2)
            
            # FIX: Only read columns A through I
            if file_extension == "xlsx":
                df = pd.read_excel(uploaded_file, header=header_row, usecols='A:I')
            else: # "csv"
                df = pd.read_csv(uploaded_file, header=header_row, usecols=range(9)) # 9 columns = A through I
            
            st.subheader("Data Preview (First 5 Rows from Columns A-I)")
            st.dataframe(df.head())
            
            analyze_data = st.button("Run Analysis")
            if analyze_data:
                panels = parse_excel_panels(df, spacing)
                if panels:
                    beds, trucks = compute_beds_and_trucks(panels)
                    st.success(f"Parsed {len(panels)} panels, which fit into {len(beds)} beds and {len(trucks)} trucks.")
                    output = export_to_excel(beds, trucks)
                    st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx")

        elif file_extension == "pdf":
            # PDF logic remains unchanged
            analyze_pdf = st.button("Run PDF Analysis")
            if analyze_pdf:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                panels = parse_pdf_panels(tmp_file_path, spacing)
                beds, trucks = compute_beds_and_trucks(panels)
                st.success(f"Parsed {len(panels)} panels, which fit into {len(beds)} beds and {len(trucks)} trucks.")
                output = export_to_excel(beds, trucks)
                st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx")
        else:
            st.error("Unsupported file type. Please upload a PDF, XLSX, or CSV file.")

    except Exception as e:
        st.error(f"An error occurred during file processing: {e}")
        st.info("Please check that the selected header row is correct and the file format is not corrupted.")
