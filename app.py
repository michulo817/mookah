import streamlit as st
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import column_index_from_string
from pypdf import PdfReader
import shutil
import os

# --- PAGE SETUP ---
st.set_page_config(page_title="SPUB Generator", page_icon="💊", layout="centered")
st.title("📄 SPUB to Excel Generator")
st.markdown("Upload the patient's SPUB PDF. The app will automatically map the data into your default templates.")

# --- TEMPLATE PATHS ---
# The app will automatically look for these exact files in the same folder
KEW_TEMPLATE = "KEW PS7 empty_2.xlsx"
LAMP_TEMPLATE = "Lampiran 2 empty_2.xlsx"

# --- BORDER STYLING ---
thin_border = Border(
    left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
)

# --- EXTRACTION LOGIC ---
def extract_data(pdf_path):
    patient_data = {"Nama": "Unknown_Patient", "No KP": "", "Tarikh": "13/10/2026", "Drugs": []}
    full_text = ""
    reader = PdfReader(pdf_path)
    for page in reader.pages:
        full_text += page.extract_text() + "\n"

    name_match = re.search(r"Nama:\s*(.*)", full_text)
    kp_match = re.search(r"No\.\s*KP:\s*(\d+)", full_text)
    if name_match: patient_data["Nama"] = name_match.group(1).strip()
    if kp_match: patient_data["No KP"] = kp_match.group(1).strip()

    blocks = re.split(r'WD_FM\nED\n(?:\d+\.\n)?', full_text)
    for block in blocks[1:]:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        qty_idx = -1
        for idx, line in enumerate(lines):
            if "hari" in line.lower():
                qty_idx = idx - 1
                break
        if qty_idx != -1:
            qty = lines[idx] if idx != -1 else lines[qty_idx] # Fallback protection
            last_paren_idx = -1
            for idx in range(qty_idx - 1, -1, -1):
                if ')' in lines[idx]:
                    last_paren_idx = idx
                    break
            if last_paren_idx != -1:
                drug_name = " ".join(lines[:last_paren_idx+1])
                dosage_raw = " ".join(lines[last_paren_idx+1:qty_idx])
                dosage = re.sub(r'\s+', ' ', dosage_raw).strip()
                dosage = re.sub(r'(\d)\s+Q\s+(\d)\s+W', r'\1Q\2W', dosage)
                dosage = re.sub(r'(\d)\s+Q\s+(\d)\s+D', r'\1Q\2D', dosage)
            else:
                drug_name = lines[0]
                dosage = " ".join(lines[1:qty_idx])
            patient_data["Drugs"].append({"Name": drug_name, "Dosage": dosage, "Quantity": lines[qty_idx]})
    return patient_data

def process_kew(patient_data, template_path, output_name):
    shutil.copy(template_path, output_name)
    wb = openpyxl.load_workbook(output_name)
    sheet = wb.active
    font_style = Font(name='Arial', size=10, bold=True)
    align_left = Alignment(vertical='top', horizontal='left', wrap_text=True)
    align_center = Alignment(vertical='top', horizontal='center', wrap_text=True)

    for merged_range in list(sheet.merged_cells.ranges):
        if 17 <= merged_range.min_row <= 33:
            sheet.unmerge_cells(str(merged_range))
            
    for r in range(17, 34):
        for c in ['A', 'B', 'C', 'D']: sheet[f'{c}{r}'] = None
        sheet.merge_cells(f'B{r}:C{r}')
        for c_idx in range(1, 16): sheet.cell(row=r, column=c_idx).border = thin_border

    for i, drug in enumerate(patient_data["Drugs"]):
        if i >= 17: break 
        r = 17 + i
        cell_a = sheet[f'A{r}']
        cell_a.value = i + 1
        cell_a.font, cell_a.alignment = font_style, align_center
        
        cell_b = sheet[f'B{r}']
        cell_b.value = drug["Name"]
        cell_b.font, cell_b.alignment = font_style, align_left
        
        cell_d = sheet[f'D{r}']
        cell_d.value = int(drug["Quantity"]) if drug["Quantity"].isdigit() else drug["Quantity"]
        cell_d.font, cell_d.alignment = font_style, align_center
    wb.save(output_name)

def process_lampiran(patient_data, template_path, output_name):
    shutil.copy(template_path, output_name)
    wb = openpyxl.load_workbook(output_name)
    sheet = wb.active
    font_style = Font(name='Arial', size=8, bold=False)
    align_center = Alignment(vertical='center', horizontal='center', wrap_text=True)
    align_left = Alignment(vertical='center', horizontal='left', wrap_text=True)

    for merged_range in list(sheet.merged_cells.ranges):
        if 10 <= merged_range.min_row <= 28:
            sheet.unmerge_cells
