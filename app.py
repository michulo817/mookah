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
st.markdown("Upload your empty templates and the patient's SPUB PDF. The app will neatly extract and format the data for you to download.")

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
            qty = lines[qty_idx]
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
            patient_data["Drugs"].append({"Name": drug_name, "Dosage": dosage, "Quantity": qty})
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
            sheet.unmerge_cells(str(merged_range))
            
    for r in range(10, 29):
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'L', 'M']: sheet[f'{col}{r}'] = None
        for c_idx in range(1, 14): sheet.cell(row=r, column=c_idx).border = thin_border

    start_row = 10
    num_drugs = min(len(patient_data["Drugs"]), 19) 
    end_row = start_row + num_drugs - 1

    sheet[f'A{start_row}'] = 1
    sheet[f'B{start_row}'] = patient_data["Tarikh"]
    sheet[f'C{start_row}'] = "HOSPITAL BINTULU"
    sheet[f'D{start_row}'] = patient_data["Nama"]
    
    ic_cell = sheet[f'E{start_row}']
    ic_cell.value = str(patient_data["No KP"])
    ic_cell.number_format = '@' 
    
    if num_drugs > 1:
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            sheet.merge_cells(start_row=start_row, start_column=column_index_from_string(col),
                              end_row=end_row, end_column=column_index_from_string(col))
            cell = sheet[f'{col}{start_row}']
            cell.font, cell.alignment = font_style, align_center

    for i, drug in enumerate(patient_data["Drugs"]):
        if i >= 19: break
        r = start_row + i
        cell_g = sheet[f'G{r}']
        cell_g.value = drug["Name"]
        cell_g.font, cell_g.alignment = font_style, align_left
        
        cell_l = sheet[f'L{r}']
        cell_l.value = drug["Dosage"]
        cell_l.font, cell_l.alignment = font_style, align_center
        
        cell_m = sheet[f'M{r}']
        cell_m.value = int(drug["Quantity"]) if drug["Quantity"].isdigit() else drug["Quantity"]
        cell_m.font, cell_m.alignment = font_style, align_center
    wb.save(output_name)

# --- UPGRADED UI APP LAYOUT ---
st.sidebar.header("⚙️ Master Templates")
st.sidebar.info("Upload your empty Excel templates here. These serve as the master layout for this session.")
uploaded_kew = st.sidebar.file_uploader("KEW PS7 Template", type="xlsx")
uploaded_lamp = st.sidebar.file_uploader("Lampiran 2 Template", type="xlsx")

st.markdown("### 1. Upload Patient SPUB")
uploaded_pdf = st.file_uploader("Drop the patient's SPUB PDF here", type="pdf", label_visibility="collapsed")

if uploaded_pdf and uploaded_kew and uploaded_lamp:
    st.markdown("### 2. Preview & Generate")
    
    # Wide button
    if st.button("Extract & Format Files", type="primary", use_container_width=True):
        
        # Animated status box instead of a basic spinner
        with st.status("Processing documents...", expanded=True) as status:
            st.write("Reading PDF data...")
            with open("temp_pdf.pdf", "wb") as f: f.write(uploaded_pdf.getbuffer())
            with open("temp_kew.xlsx", "wb") as f: f.write(uploaded_kew.getbuffer())
            with open("temp_lamp.xlsx", "wb") as f: f.write(uploaded_lamp.getbuffer())

            patient_data = extract_data("temp_pdf.pdf")
            
            st.write("Formatting KEW PS7 & Lampiran 2...")
            safe_name = re.sub(r'[^a-zA-Z0-9]', '_', patient_data['Nama'])
            out_kew = f"KEW_PS7_{safe_name}.xlsx"
            out_lamp = f"Lampiran2_{safe_name}.xlsx"

            process_kew(patient_data, "temp_kew.xlsx", out_kew)
            process_lampiran(patient_data, "temp_lamp.xlsx", out_lamp)
            
            status.update(label="Extraction Complete!", state="complete", expanded=False)

        # Dashboard Metrics
        col1, col2, col3 = st.columns(3)
        col1.metric("Patient Name", patient_data["Nama"])
        col2.metric("IC Number", patient_data["No KP"])
        col3.metric("Drugs Extracted", len(patient_data["Drugs"]))

        # Live Data Preview Table
        with st.expander("👀 Preview Extracted Medications", expanded=True):
            st.dataframe(patient_data["Drugs"], use_container_width=True)

        # Full-width Download Buttons
        st.success("Your perfectly formatted Excel files are ready.")
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            with open(out_kew, "rb") as f:
                st.download_button("⬇️ Download KEW PS7", f, file_name=out_kew, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with dl_col2:
            with open(out_lamp, "rb") as f:
                st.download_button("⬇️ Download Lampiran 2", f, file_name=out_lamp, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
        # Clean up temp files behind the scenes
        os.remove("temp_pdf.pdf")
        os.remove("temp_kew.xlsx")
        os.remove("temp_lamp.xlsx")

elif not (uploaded_kew and uploaded_lamp):
    st.warning("👈 Please upload your empty Excel templates in the sidebar to begin.")
