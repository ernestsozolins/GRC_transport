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

    # Match best-fit columns using fuzzy matching
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
        weight = row.get(column_map["weight (kg)"], 0)
        panels.append({
            "Type": row[column_map["panel type"]],
            "Height": d,
            "Width": l,
            "Depth": h,
            "Weight": weight
        })
    return panels

# Remaining unchanged compute_beds_and_trucks and export_to_excel functions below...
