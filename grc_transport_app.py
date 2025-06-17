import streamlit as st
import pandas as pd
import io
import tempfile
import pdfplumber
import re
import xlsxwriter
import difflib

# --- Parsing Logic ---
def parse_excel_panels(file_path, spacing=100, header_row=0):
    df = pd.read_excel(file_path, header=header_row)
    df.replace(to_replace=r"\s+", value="", regex=True, inplace=True)
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df.columns = df.columns.str.strip().str.lower()
    colnames = df.columns.tolist()
    st.write("Detected columns:", colnames)

    column_map = {}
    targets = {
        "panel type": ["panel type", "type", "cast unit", "cast_unit"],
        "height (mm)": ["height", "height (mm)", "augstums"],
        "length (mm)": ["length", "length (mm)", "garums"],
        "depth (mm)": ["depth", "depth (mm)", "platums"],
        "weight (kg)": ["weight", "weight (kg)", "svars"]
    }

    required_keys = ["panel type", "height (mm)", "length (mm)", "depth (mm)"]
    optional_keys = ["weight (kg)"]

    missing = []
    for key in required_keys:
        match = fuzzy_match_column(colnames, targets[key])
        if match:
            column_map[key] = match
        else:
            missing.append(key)

    for key in optional_keys:
        match = fuzzy_match_column(colnames, targets[key])
        if match:
            column_map[key] = match

    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    panels = []
    # FIX: The logic inside the loop is now more robust to handle bad data.
    for index, row in df.iterrows():
        try:
            # --- FIX 1: Safely convert required dimensions to numbers ---
            # pd.to_numeric will turn empty strings and other non-numbers into NaN (Not a Number)
            h_num = pd.to_numeric(row[column_map["height (mm)"]], errors='coerce')
            l_num = pd.to_numeric(row[column_map["length (mm)"]], errors='coerce')
            d_num = pd.to_numeric(row[column_map["depth (mm)"]], errors='coerce')

            # --- FIX 2: Skip rows if essential data is missing or not a number ---
            if pd.isna(h_num) or pd.isna(l_num) or pd.isna(d_num):
                # Using the DataFrame index and header_row to give a helpful row number from the file
                st.warning(f"⚠️ Skipping row {index + header_row + 2} due to missing/invalid dimension values.")
                continue  # Go to the next row

            # If we get here, we have valid numbers. Now we can do the math.
            h = h_num + 2 * spacing
            l = l_num + 2 * spacing
            d = d_num + 2 * spacing
            weight = 0

            if "weight (kg)" in column_map:
                # Also use the safe conversion for the optional weight column
                weight_val = pd.to_numeric(row[column_map["weight (kg)"]], errors='coerce')
                # If conversion is successful (not NaN), use the value. Otherwise, weight remains 0.
                if pd.notna(weight_val):
                    weight = weight_val

            panel = {
                "Type": str(row[column_map["panel type"]]) if pd.notna(row[column_map["panel type"]]) else "Unknown",
                "Height": d,
                "Width": l,
                "Depth": h,
                "Weight": weight
            }
            panels.append(panel)

        except Exception as e:
            # FIX 3: Improved error message with the specific row index
            st.error(f"❌ Error parsing row at Excel index {index + header_row + 2}: {e}")
            st.write("🚨 Problematic row data:", row.to_dict())
    return panels


def fuzzy_match_column(df_columns, target_keywords):
    for target in target_keywords:
        match = difflib.get_close_matches(target, df_columns, n=1, cutoff=0.6)
        if match:
            return match[0]
    return None

def parse_excel_panels(file_path, spacing=100, header_row=0):
    df = pd.read_excel(file_path, header=header_row)
    df.replace(to_replace=r"\s+", value="", regex=True, inplace=True)
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df.columns = df.columns.str.strip().str.lower()
    colnames = df.columns.tolist()
    st.write("Detected columns:", colnames)

    column_map = {}
    targets = {
        "panel type": ["panel type", "type", "cast unit", "cast_unit"],
        "height (mm)": ["height", "height (mm)", "augstums"],
        "length (mm)": ["length", "length (mm)", "garums"],
        "depth (mm)": ["depth", "depth (mm)", "platums"],
        "weight (kg)": ["weight", "weight (kg)", "svars"]
    }

    required_keys = ["panel type", "height (mm)", "length (mm)", "depth (mm)"]
    optional_keys = ["weight (kg)"]

    missing = []
    for key in required_keys:
        match = fuzzy_match_column(colnames, targets[key])
        if match:
            column_map[key] = match
        else:
            missing.append(key)

    for key in optional_keys:
        match = fuzzy_match_column(colnames, targets[key])
        if match:
            column_map[key] = match

    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    panels = []
    # FIX: The whole 'for' loop body was rewritten to fix major indentation and logic errors.
    # Now, the entire processing for a single row is wrapped in one try/except block.
    for _, row in df.iterrows():
        try:
            h = row[column_map["height (mm)"]] + 2 * spacing
            l = row[column_map["length (mm)"]] + 2 * spacing
            d = row[column_map["depth (mm)"]] + 2 * spacing
            weight = 0

            if "weight (kg)" in column_map:
                val = row[column_map["weight (kg)"]]
                if pd.notna(val) and not isinstance(val, pd.Series):
                    # A nested try/except is fine for handling the specific float conversion
                    try:
                        weight = float(val)
                    except (ValueError, TypeError):
                        weight = 0 # If conversion fails, weight remains 0

            panel = {
                "Type": str(row[column_map["panel type"]]) if pd.notna(row[column_map["panel type"]]) else "Unknown",
                "Height": d,
                "Width": l,
                "Depth": h,
                "Weight": weight
            }
            panels.append(panel)
        except Exception as e:
            st.error(f"❌ Error parsing row: {e}")
            st.write("🚨 Problematic row data:", row.to_dict())
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
        panel_types = [
            str(p['Type']).strip()
            for p in bed
            if pd.notna(p['Type']) and isinstance(p['Type'], str) and str(p['Type']).strip().lower() not in ("", "nan", "none")
        ]
        bed_summaries.append({
            'Length': bed_length,
            'Height': bed_height,
            'Width': bed_width,
            'Weight': bed_weight,
            'Num Panels': len(bed),
            'Panel Types': panel_types
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
                "Panel Types": ", ".join(
                    str(pt)
                    for b in truck
                    if isinstance(b.get("Panel Types"), list)
                    for pt in b["Panel Types"]
                    if pd.notna(pt)
                )
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
    # We read the file into memory once to avoid issues with temp files
    uploaded_bytes = uploaded_file.getvalue()
    
    if uploaded_file.name.endswith(".xlsx"):
        try:
            # Use BytesIO to read the in-memory bytes
            preview_df = pd.read_excel(io.BytesIO(uploaded_bytes), header=None, nrows=5)
            preview_df.replace(to_replace=r"\s+", value="", regex=True, inplace=True)
            preview_df = preview_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            st.subheader("Preview First Rows (cleaned)")
            st.dataframe(preview_df)
            header_row = st.number_input("Select header row (0-indexed)", min_value=0, max_value=10, value=1)
            analyze = st.button("Run Analysis")

            if analyze:
                # Pass the in-memory bytes to the parsing function
                panels = parse_excel_panels(io.BytesIO(uploaded_bytes), spacing, header_row=header_row)
                beds, trucks = compute_beds_and_trucks(panels)
                st.success(f"Parsed {len(panels)} panels, {len(beds)} beds, {len(trucks)} trucks")

                output = export_to_excel(beds, trucks)
                st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing Excel: {e}")
    else: # Handles PDF
        analyze = st.button("Run Analysis")
        if analyze:
            # Using a temporary file is still a good approach for libraries like pdfplumber
            # that expect a file path.
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                tmp_file.write(uploaded_bytes)
                tmp_file_path = tmp_file.name
            
            try:
                panels = parse_pdf_panels(tmp_file_path, spacing)
                beds, trucks = compute_beds_and_trucks(panels)
                st.success(f"Parsed {len(panels)} panels, {len(beds)} beds, {len(trucks)} trucks")

                output = export_to_excel(beds, trucks)
                st.download_button("Download Transport Plan", data=output, file_name="transport_plan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error processing PDF: {e}")
