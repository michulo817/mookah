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
                                  end_row=end_row, end_column=column_index_from_string(col))
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            sheet[f'{col}{current_row}'].font = font_style
            sheet[f'{col}{current_row}'].alignment = align_center

        for idx, drug in enumerate(drugs):
            r = current_row + idx
            cell_g = sheet[f'G{r}']
            cell_g.value = drug["Name"]
            cell_g.font, cell_g.alignment = font_style, align_left
            
            cell_l = sheet[f'L{r}']
            cell_l.value = drug["Dosage"]
            cell_l.font, cell_l.alignment = font_style, align_center
            
            cell_m = sheet[f'M{r}']
            cell_m.value = int(drug["Quantity"]) if str(drug["Quantity"]).isdigit() else drug["Quantity"]
            cell_m.font, cell_m.alignment = font_style, align_center

        current_row += count
        patient_number += 1

    wb.save(output_path)

st.write("---")

uploaded_pdfs = st.file_uploader(
    "Upload Patient SPUB PDFs (Single or Multiple)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_pdfs:
    st.info(f"Loaded {len(uploaded_pdfs)} PDF file(s).")

    if not os.path.exists(KEW_TEMPLATE) or not os.path.exists(LAMP_TEMPLATE):
        st.error(f"Cannot find: `{KEW_TEMPLATE}` or `{LAMP_TEMPLATE}` in repository root.")
    else:
        if st.button("Generate Excel Batches", type="primary", use_container_width=True):
            with st.spinner("Extracting and mapping patient data..."):
                all_patients = []
                for pdf_file in uploaded_pdfs:
                    p_data = extract_data_from_pdf(pdf_file)
                    all_patients.append(p_data)

                total_drugs_all = sum(len(p["Drugs"]) for p in all_patients)
                
                if total_drugs_all == 0:
                    st.error("⚠️ No medications could be detected in the uploaded PDF. Please make sure the uploaded PDF contains selectable text.")
                    st.session_state["batch_results"] = None
                else:
                    generated_files_dict = {}

                    lampiran_pages = []
                    current_page_patients = []
                    current_page_used_rows = 0

                    for patient in all_patients:
                        remaining_drugs = list(patient["Drugs"])
                        if not remaining_drugs:
                            continue
                        while remaining_drugs:
                            space_left = LAMP_CAPACITY - current_page_used_rows
                            if space_left <= 0:
                                lampiran_pages.append(current_page_patients)
                                current_page_patients = []
                                current_page_used_rows = 0
                                space_left = LAMP_CAPACITY

                            take_drugs = remaining_drugs[:space_left]
                            remaining_drugs = remaining_drugs[space_left:]

                            current_page_patients.append({
                                "Nama": patient["Nama"],
                                "No KP": patient["No KP"],
                                "Tarikh": patient["Tarikh"],
                                "Drugs": take_drugs
                            })
                            current_page_used_rows += len(take_drugs)

                    if current_page_patients:
                        lampiran_pages.append(current_page_patients)

                    for page_idx, page_patients in enumerate(lampiran_pages):
                        suffix = f"_Page_{page_idx+1}" if len(lampiran_pages) > 1 else ""
                        out_name = f"Lampiran2_Batch{suffix}.xlsx"
                        build_lampiran_page(page_patients, LAMP_TEMPLATE, out_name)
                        with open(out_name, "rb") as f:
                            generated_files_dict[out_name] = f.read()
                        if os.path.exists(out_name):
                            os.remove(out_name)

                    all_drugs_flat = []
                    for p in all_patients:
                        all_drugs_flat.extend(p["Drugs"])

                    kew_chunks = [all_drugs_flat[i:i + KEW_CAPACITY] for i in range(0, len(all_drugs_flat), KEW_CAPACITY)]
                    total_counter = 1

                    for page_idx, chunk in enumerate(kew_chunks):
                        suffix = f"_Page_{page_idx+1}" if len(kew_chunks) > 1 else ""
                        out_name = f"KEW_PS7_Batch{suffix}.xlsx"
                        build_kew_page(chunk, KEW_TEMPLATE, out_name, start_index=total_counter)
                        total_counter += len(chunk)
                        with open(out_name, "rb") as f:
                            generated_files_dict[out_name] = f.read()
                        if os.path.exists(out_name):
                            os.remove(out_name)

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for filename, filedata in generated_files_dict.items():
                            zf.writestr(filename, filedata)
                    zip_data = zip_buffer.getvalue()

                    st.session_state["batch_results"] = {
                        "patients": all_patients,
                        "files": generated_files_dict,
                        "zip": zip_data
                    }

if st.session_state["batch_results"]:
    results = st.session_state["batch_results"]
    patients = results["patients"]
    files = results["files"]
    zip_bytes = results["zip"]

    st.success(f"Successfully processed {len(patients)} patient(s) into {len(files)} page(s)!")

    st.subheader("📋 Extracted Patient Data Preview")
    for p in patients:
        with st.expander(f"👤 Patient: {p['Nama']} (IC: {p['No KP']}) — {len(p['Drugs'])} Drug(s)", expanded=True):
            if p["Drugs"]:
                st.dataframe(p["Drugs"], use_container_width=True)
            else:
                st.warning("No drug rows could be parsed for this patient.")

    st.write("---")
    st.subheader("📥 Download Generated Sheets")

    cols = st.columns(2)
    file_names = list(files.keys())
    for idx, fname in enumerate(file_names):
        with cols[idx % 2]:
            st.download_button(
                label=f"⬇️ {fname}",
                data=files[fname],
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{fname}",
                use_container_width=True
            )

    st.download_button(
        label="📦 Download All Pages as ZIP",
        data=zip_bytes,
        file_name="SPUB_Processed_Batch.zip",
        mime="application/zip",
        key="dl_zip_all",
        use_container_width=True
    )
