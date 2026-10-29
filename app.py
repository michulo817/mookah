import streamlit as st
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import column_index_from_string
from pypdf import PdfReader
import shutil
import os
import io
import zipfile

st.set_page_config(page_title="SPUB Multi-Batch Generator", page_icon="💊", layout="wide")
st.title("📄 SPUB Multi-Batch Generator")
st.write("Upload one or multiple patient SPUB PDFs. If patient medications exceed table capacity, sequential pages are created while preserving the layout.")

KEW_TEMPLATE = "KEW PS7 empty.xlsx"
LAMP_TEMPLATE = "Lampiran 2 empty.xlsx"

LAMP_CAPACITY = 19   # Rows 10 to 28
KEW_CAPACITY = 17    # Rows 17 to 33

thin_border = Border(
    left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
)

if "batch_results" not in st.session_state:
    st.session_state["batch_results"] = None

def extract_data_from_pdf(pdf_stream):
    patient_data = {"Nama": "Unknown", "No KP": "", "Tarikh": "13/10/2026", "Drugs": []}
    
    if hasattr(pdf_stream, "seek"):
        pdf_stream.seek(0)
        
    full_text = ""
    reader = PdfReader(pdf_stream)
    for page in reader.pages:
        full_text += (page.extract_text() or "") + "\n"

    full_text = full_text.replace('\r\n', '\n').replace('\r', '\n')

    name_match = re.search(r"Nama:\s*(.*)", full_text)
    kp_match = re.search(r"No\.\s*KP:\s*(\d+)", full_text)
    if name_match: 
        patient_data["Nama"] = name_match.group(1).strip()
    if kp_match: 
        patient_data["No KP"] = kp_match.group(1).strip()

    blocks = re.split(r'WD_FM\s+ED\s+(?:\d+\.\s*)?', full_text)
    for block in blocks[1:]:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        qty_idx = -1
        for idx, line in enumerate(lines):
            if "hari" in line.lower():
                qty_idx = idx - 1
                break
        if qty_idx != -1 and qty_idx > 0:
            qty = lines[qty_idx]
            last_paren_idx = -1
            for idx in range(qty_idx - 1, -1, -1):
                if ')' in lines[idx]:
                    last_paren_idx = idx
                    break
            if last_paren_idx != -1:
                drug_name = " ".join(lines[:last_paren_idx+1])
                dosage_raw = " ".join(lines[last_paren_idx+1:qty_idx])
            else:
                drug_name = lines[0]
                dosage_raw = " ".join(lines[1:qty_idx])
                
            dosage = re.sub(r'\s+', ' ', dosage_raw).strip()
            dosage = re.sub(r'(\d)\s*Q\s*(\d)\s*W', r'\1Q\2W', dosage, flags=re.I)
            dosage = re.sub(r'(\d)\s*Q\s*(\d)\s*D', r'\1Q\2D', dosage, flags=re.I)
            patient_data["Drugs"].append({"Name": drug_name, "Dosage": dosage, "Quantity": qty})
            
    return patient_data

def build_kew_page(drugs_chunk, template_path, output_path, start_index=1):
    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    sheet = wb.active

    font_style = Font(name='Arial', size=10, bold=True)
    align_left = Alignment(vertical='top', horizontal='left', wrap_text=True)
    align_center = Alignment(vertical='top', horizontal='center', wrap_text=True)

    for merged_range in list(sheet.merged_cells.ranges):
        if 17 <= merged_range.min_row <= 33:
            sheet.unmerge_cells(str(merged_range))
            
    for r in range(17, 34):
        for c in ['A', 'B', 'C', 'D']: 
            sheet[f'{c}{r}'] = None
        sheet.merge_cells(f'B{r}:C{r}')
        for c_idx in range(1, 16): 
            sheet.cell(row=r, column=c_idx).border = thin_border

    for i, drug in enumerate(drugs_chunk):
        r = 17 + i
        cell_a = sheet[f'A{r}']
        cell_a.value = start_index + i
        cell_a.font, cell_a.alignment = font_style, align_center
        
        cell_b = sheet[f'B{r}']
        cell_b.value = drug["Name"]
        cell_b.font, cell_b.alignment = font_style, align_left
        
        cell_d = sheet[f'D{r}']
        cell_d.value = int(drug["Quantity"]) if str(drug["Quantity"]).isdigit() else drug["Quantity"]
        cell_d.font, cell_d.alignment = font_style, align_center

    wb.save(output_path)

def build_lampiran_page(patient_groups_in_page, template_path, output_path):
    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    sheet = wb.active

    font_style = Font(name='Arial', size=8, bold=False)
    align_center = Alignment(vertical='center', horizontal='center', wrap_text=True)
    align_left = Alignment(vertical='center', horizontal='left', wrap_text=True)

    for merged_range in list(sheet.merged_cells.ranges):
        if 10 <= merged_range.min_row <= 28:
            sheet.unmerge_cells(str(merged_range))
            
    for r in range(10, 29):
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'L', 'M']: 
            sheet[f'{col}{r}'] = None
        for c_idx in range(1, 14): 
            sheet.cell(row=r, column=c_idx).border = thin_border

    current_row = 10
    patient_number = 1

    for p in patient_groups_in_page:
        drugs = p["Drugs"]
        count = len(drugs)
        if count == 0:
            continue
        end_row = current_row + count - 1

        sheet[f'A{current_row}'] = patient_number
        sheet[f'B{current_row}'] = p["Tarikh"]
        sheet[f'C{current_row}'] = "HOSPITAL BINTULU"
        sheet[f'D{current_row}'] = p["Nama"]
        
        ic_cell = sheet[f'E{current_row}']
        ic_cell.value = str(p["No KP"])
        ic_cell.number_format = '@'

        if count > 1:
            for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                sheet.merge_cells(start_row=current_row, start_column=column_index_from_string(col),
