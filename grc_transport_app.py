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

    for panel_type, qty, height, length, depth in matches:
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

    return panels

def fuzzy_match_column(df_columns, target_keywords):
    for target in target_keywords:
        match = difflib.get_close_matches(target, df_columns, n=1, cutoff=0.6)
        if match:
            return match[0]
    return None

def parse_excel_panels(file_path, spacing=100):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip().str.lower()
    colnames = df.columns.tolist()

    column_map = {}
    targets = {
        "panel type": ["panel type", "type", "cast unit", "cast_unit"],
        "height (mm)": ["height", "height (mm)", "augstums"],
        "length (mm)": ["length", "length (mm)", "garums"],
        "depth (mm)": ["depth", "depth (mm)", "platums"],
        "weight (kg)": ["weight", "weight (kg)", "svars"]
    }

    missing = []
    for key, variants in targets.items():
        match = fuzzy_match_column(colnames, variants)
        if match:
            column_map[key] = match
        else:
            missing.append(key)

    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    panels = []
    for _, row in df.iterrows():
        h = row[column_map["height (mm)"]] + 2 * spacing
        l = row[column_map["length (mm)"]] + 2 * spacing
        d = row[column_map["depth (mm)"]] + 2 * spacing
        weight_col = column_map.get("weight (kg)")
        weight = row[weight_col] if weight_col and weight_col in row and pd.notna(row[weight_col]) else 0
        panels.append({
            "Type": row[column_map["panel type"]],
            "Height": d,
            "Width": l,
            "Depth": h,
            "Weight": weight
        })
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
        bed_summaries.append({
            'Length': bed_length,
            'Height': bed_height,
            'Width': bed_width,
            'Weight': bed_weight,
            'Num Panels': len(bed),
            'Panel Types': list(set(p['Type'] for p in bed))
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
                "Truck #": i + 1,
                "Num Beds": len(truck),
                "Total Weight (kg)": sum(b['Weight'] for b in truck),
                "Panel Types": ", ".join(set(pt for b in truck for pt in b['Panel Types']))
            })
        pd.DataFrame(truck_summary).to_excel(writer, index=False, sheet_name="Truck Summary")

        summary = pd.DataFrame({
            "Metric": ["Total Beds", "Total Trucks"],
            "Value": [len(beds), len(trucks)]
        })
        summary.to_excel(writer, index=False, sheet_name="Summary")

    output.seek(0)
    return output

# --- Streamlit App ---
st.set_page_config(page_title="GRC Transport Planner", layout="wide")
st.title("🚚 GRC Panel Transport & Storage Estimator")

uploaded_file = st.file_uploader("Upload a PDF or Excel File", type=["pdf", "xlsx"])
spacing = st.number_input("Panel spacing (mm)", min_value=0, value=100)

if uploaded_file:
    analyze = st.button("Analyze")
    if analyze:
        with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_file_path = tmp_file.name

        try:
            if uploaded_file.name.endswith(".pdf"):
                panels = parse_pdf_panels(tmp_file_path, spacing)
            else:
                panels = parse_excel_panels(tmp_file_path, spacing)

            beds, trucks = compute_beds_and_trucks(panels)
            st.success(f"Parsed {len(panels)} panels, {len(beds)} beds, {len(trucks)} trucks")

            output = export_to_excel(beds, trucks)
            st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Error processing file: {e}")
else:
    st.info("Upload a file to begin.")
