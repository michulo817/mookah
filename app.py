import streamlit as st
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import column_index_from_string
from pypdf import PdfReader
import shutil
import os
import io
import zipfile
from datetime import datetime

# --- 1. REFINED BRUTALIST CSS AESTHETIC ---
st.set_page_config(page_title="SPUB ENGINE", layout="wide")

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    .stApp {
        background-color: #000000;
        color: #ffffff;
    }
    
    h1, h2, h3, p, label {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .ye-title {
        font-size: 3rem;
        font-weight: 900;
        letter-spacing: 4px;
        margin-top: -40px;
        margin-bottom: 0px;
    }
    
    .ye-subtitle {
        color: #666666;
        font-size: 0.85rem;
        letter-spacing: 2px;
        margin-bottom: 30px;
    }

    .stButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border-radius: 0px !important;
        border: none !important;
        font-weight: 800 !important;
        letter-spacing: 2px !important;
        padding: 15px !important;
        transition: opacity 0.2s ease;
    }
    .stButton > button:hover {
        opacity: 0.8;
    }
    
    .stDownloadButton > button {
        background-color: transparent !important;
        color: #ffffff !important;
        border: 1px solid #ffffff !important;
        border-radius: 0px !important;
        font-weight: 700 !important;
    }
    .stDownloadButton > button:hover {
        background-color: #ffffff !important;
        color: #000000 !important;
    }

    [data-testid="stFileUploader"] {
        background-color: transparent !important;
    }
    [data-testid="stFileUploadDropzone"] {
        background-color: #080808 !important;
        border: 1px solid #333333 !important;
        border-radius: 0px !important;
    }
    [data-testid="stFileUploadDropzone"] button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border-radius: 0px !important;
        font-weight: 800 !important;
    }

    [data-testid="stAlert"] {
        background-color: transparent !important;
        border: 1px solid #444444 !important;
        border-radius: 0px !important;
        color: #ffffff !important;
    }
    
    [data-testid="stExpander"] {
        border: 1px solid #333333 !important;
        border-radius: 0px !important;
        background-color: transparent !important;
    }
    
    .streamlit-expanderHeader {
        font-weight: 600 !important;
        color: #ffffff !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. GLOBAL VARIABLES & STYLES ---
KEW_TEMPLATE = "KEW PS7 empty.xlsx"
LAMP_TEMPLATE = "Lampiran 2 empty.xlsx"
INDENT_TEMPLATE = "Copy of SPUB drug indent list google doc.xlsx"

LAMP_CAPACITY = 19
KEW_CAPACITY = 17

thin_border = Border(
    left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
)

COLORS = {
    "HBTU": "FFFF00FF",   # Magenta / Purple for Bintulu
    "HSIBU": "FFFFFF00",  # Yellow for Sibu
    "SGH": "FFFF9900",    # Orange for SGH
    "DEFAULT": "FFFFFFFF" 
}

SPECIALIZED_KEYWORDS = [
    "HYLO", "COMOD", "OPTHALMIC", "EYE DROP", "LATANOPROST", "DORZOLAMIDE", "TIMOLOL", 
    "INHALER", "RESPIMAT", "TURBUHALER", "ACCUHALER", "ELLIPTA", "EVOHALER", "BREEZHALER",
    "GLIPTIN", "GLIFLOZIN", "RYZODEG", "TRESIBA", "TOUJEO", "NOVOMIX", "PENFILL", "CARTRIDGE",
    "LEVETIRACETAM", "TOPIRAMATE", "LAMOTRIGINE", "CLOBAZAM", "PREGABALIN",
    "SILDENAFIL", "OXYBUTYNIN", "TAMSULOSIN", "DUTASTERIDE",
    "TAMOXIFEN", "MYCOPHENOLATE", "TACROLIMUS", "CYCLOSPORINE", 
    "CHOLECALCIFEROL", "CALCITRIOL"
]

def sanitize_excel_text(text):
    if not isinstance(text, str): return text
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)

# --- 3. EXTRACTION ENGINES ---
def extract_general_data(pdf_stream):
    patient_data = {"Nama": "UNKNOWN", "No KP": "", "Tarikh": datetime.now().strftime("%d/%m/%Y"), "Drugs": []}
    if hasattr(pdf_stream, "seek"): pdf_stream.seek(0)
        
    full_text = ""
    reader = PdfReader(pdf_stream)
    for page in reader.pages: full_text += (page.extract_text() or "") + "\n"
    full_text = full_text.replace('\r\n', '\n').replace('\r', '\n')

    name_match = re.search(r"Nama:\s*(.*)", full_text)
    kp_match = re.search(r"No\.\s*KP:\s*(\d+)", full_text)
    if name_match: patient_data["Nama"] = name_match.group(1).strip().upper()
    if kp_match: patient_data["No KP"] = kp_match.group(1).strip()

    blocks = re.split(r'\n(?:[A-Z_ ]+\n)?\d+\.\n', full_text)
    for block in blocks[1:]:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        qty_idx = next((i - 1 for i, line in enumerate(lines) if "hari" in line.lower()), -1)
        
        if qty_idx > 0:
            qty = lines[qty_idx]
            last_paren_idx = next((i for i in range(qty_idx - 1, -1, -1) if ')' in lines[i]), -1)
            
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

def extract_indent_data(pdf_stream):
    if hasattr(pdf_stream, "seek"): pdf_stream.seek(0)
    full_text = ""
    reader = PdfReader(pdf_stream)
    for page in reader.pages: full_text += (page.extract_text() or "") + "\n"
    # Scrub text for clean extraction
    full_text = full_text.replace('\r\n', '\n').replace('\r', '\n').replace('ƀ', '').replace('\x03', '')

    name = re.search(r"Nama:\s*(.*)", full_text)
    kp = re.search(r"No\.\s*KP:\s*(\d+)", full_text)
    mrn = re.search(r"(HBTU\d+|HSIBU\d+|SGH\d+|HUS\d+)", full_text)
    tca = re.search(r"TCA.*?:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", full_text, re.IGNORECASE)

    patient_data = {
        "Nama": name.group(1).strip().upper() if name else "UNKNOWN",
        "No KP": kp.group(1).strip() if kp else "UNKNOWN",
        "MRN": mrn.group(1).strip().upper() if mrn else "",
        "TCA": tca.group(1).strip() if tca else "NO TCA",
        "Hospital_Code": "DEFAULT",
        "Drugs": []
    }

    if patient_data["MRN"].startswith("HBTU"): patient_data["Hospital_Code"] = "HBTU"
    elif patient_data["MRN"].startswith("HSIB"): patient_data["Hospital_Code"] = "HSIBU"
    elif patient_data["MRN"].startswith("SGH") or patient_data["MRN"].startswith("HUS"): patient_data["Hospital_Code"] = "SGH"

    blocks = re.split(r'\n(?:[A-Z_ ]+\n)?\d+\.\n', full_text)
    for block in blocks[1:]:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        hari_idx = next((idx for idx, line in enumerate(lines) if "hari" in line.lower()), -1)
        
        if hari_idx > 0:
            next_date = lines[hari_idx + 1] if hari_idx + 1 < len(lines) else ""
            if not re.match(r'\d{1,2}/\d{1,2}/\d{2,4}', next_date):
                next_date = ""
                
            balance_qty = re.sub(r'[a-zA-Z\s]', '', lines[hari_idx]) 
            total_qty = lines[hari_idx - 1]
            
            # Combine raw drug info and explicitly remove pricing
            drug_info = " ".join(lines[:hari_idx - 1])
            drug_name_no_price = re.sub(r'\(RM[^)]+\)', '', drug_info).strip()
            
            # Smart dosage splitting
            match = re.search(r'(\d+\s*(?:DROP|TAB|CAP|MG|ML|PUFF|UNIT).*)$', drug_name_no_price, re.IGNORECASE)
            if match:
                dosage = match.group(1).strip()
                drug_name = drug_name_no_price[:match.start()].strip()
            else:
                drug_name = drug_name_no_price
                dosage = ""

            drug_name = re.sub(r'\s*\(\s*\)$', '', drug_name).strip()

            is_group_a = re.search(r'\((A/KK|A\*|A)\)', drug_name, re.IGNORECASE)
            is_specialized = any(kw in drug_name.upper() for kw in SPECIALIZED_KEYWORDS)
            
            if is_group_a or is_specialized:
                patient_data["Drugs"].append({
                    "Name": drug_name, 
                    "Dosage": dosage, 
                    "Balance_Qty": balance_qty,
                    "Next_Date": next_date
                })
                
    return patient_data

# --- 4. EXCEL BUILDERS ---
def append_to_indent_list(patients, template_path):
    output_path = "UPDATED_INDENT_LIST.xlsx"
    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    sheet = wb.active

    font_style = Font(name='Arial', size=10)
    align_left = Alignment(vertical='center', horizontal='left', wrap_text=True)
    align_center = Alignment(vertical='center', horizontal='center', wrap_text=True)
    
    green_fill = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid")
    transparent_fill = PatternFill(start_color="00000000", end_color="00000000", fill_type="solid")

    for p in patients:
        if not p["Drugs"]: continue 
        
        # 1. Determine Exact Hospital Color
        hosp_color_hex = COLORS.get(p["Hospital_Code"], COLORS["DEFAULT"])
        hosp_fill = PatternFill(start_color=hosp_color_hex, end_color=hosp_color_hex, fill_type="solid")
        
        # 2. Find exact insertion row right underneath the hospital's block
        insert_idx = sheet.max_row + 1
        last_populated = 4
        found_hospital = False
        
        for r in range(4, sheet.max_row + 1):
            c_e = sheet.cell(row=r, column=5)
            c_c = sheet.cell(row=r, column=3).value
            c_f = sheet.cell(row=r, column=6).value
            if c_e.fill and c_e.fill.start_color and c_e.fill.start_color.index == hosp_color_hex:
                found_hospital = True
                if c_c or c_f:
                    last_populated = r
                    
        if found_hospital:
            insert_idx = last_populated + 1
            
        # 3. Inject precisely required rows
        sheet.insert_rows(insert_idx, amount=len(p["Drugs"]))
        
        combined_details = f"{p['Nama']} ({p['No KP']})\n{p['MRN']} TCA: {p['TCA']}"

        for idx, drug in enumerate(p["Drugs"]):
            r = insert_idx + idx
            
            # Wipe formatting to prevent color bleed from adjacent blocks
            for col_idx in range(1, 15):
                sheet.cell(row=r, column=col_idx).fill = transparent_fill
            
            # Checkbox Box (B)
            cell_b = sheet[f'B{r}']
            cell_b.value = False 
            cell_b.alignment = align_center

            # Patient Details (C)
            cell_c = sheet[f'C{r}']
            cell_c.value = sanitize_excel_text(combined_details) if idx == 0 else ""
            cell_c.font, cell_c.alignment = font_style, align_left

            # Hospital Color Block (E)
            sheet[f'E{r}'].fill = hosp_fill

            # Clean Drug Name (F)
            cell_f = sheet[f'F{r}']
            cell_f.value = sanitize_excel_text(drug["Name"])
            cell_f.font, cell_f.alignment = font_style, align_left

            # Exact Dosage (G)
            cell_g = sheet[f'G{r}']
            cell_g.value = sanitize_excel_text(drug["Dosage"])
            cell_g.font, cell_g.alignment = font_style, align_left

            # Next Date - Green Block (H)
            sheet[f'H{r}'].value = sanitize_excel_text(drug["Next_Date"])
            sheet[f'H{r}'].font, sheet[f'H{r}'].alignment = font_style, align_center
            sheet[f'H{r}'].fill = green_fill

            # Balance Quantity (I)
            cell_i = sheet[f'I{r}']
            bal_qty = sanitize_excel_text(str(drug["Balance_Qty"]))
            cell_i.value = int(bal_qty) if bal_qty.isdigit() else bal_qty
            cell_i.font, cell_i.alignment = font_style, align_center
            
            # Future Date Blocks (K, N)
            sheet[f'K{r}'].fill = green_fill
            sheet[f'N{r}'].fill = green_fill

            # Master Border Enforcement
            for col_idx in range(2, 15):
                sheet.cell(row=r, column=col_idx).border = thin_border
                
    wb.save(output_path)
    return output_path

def build_kew_page(drugs_chunk, template_path, output_path, start_index=1):
    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    sheet = wb.active
    font_style = Font(name='Arial', size=10, bold=True)
    align_left = Alignment(vertical='top', horizontal='left', wrap_text=True)
    align_center = Alignment(vertical='top', horizontal='center', wrap_text=True)
    for merged_range in list(sheet.merged_cells.ranges):
        if 17 <= merged_range.min_row <= 33: sheet.unmerge_cells(str(merged_range))
    for r in range(17, 34):
        for c in ['A', 'B', 'C', 'D']: sheet[f'{c}{r}'] = None
        sheet.merge_cells(f'B{r}:C{r}')
        for c_idx in range(1, 16): sheet.cell(row=r, column=c_idx).border = thin_border
    for i, drug in enumerate(drugs_chunk):
        r = 17 + i
        sheet[f'A{r}'].value = start_index + i
        sheet[f'B{r}'].value = sanitize_excel_text(drug["Name"])
        raw_qty = sanitize_excel_text(str(drug["Quantity"]))
        sheet[f'D{r}'].value = int(raw_qty) if raw_qty.isdigit() else raw_qty
        for c in ['A', 'B', 'D']: 
            sheet[f'{c}{r}'].font = font_style
            sheet[f'{c}{r}'].alignment = align_center if c != 'B' else align_left
    wb.save(output_path)

def build_lampiran_page(patient_groups_in_page, template_path, output_path):
    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    sheet = wb.active
    font_style = Font(name='Arial', size=8, bold=False)
    align_center = Alignment(vertical='center', horizontal='center', wrap_text=True)
    align_left = Alignment(vertical='center', horizontal='left', wrap_text=True)
    for merged_range in list(sheet.merged_cells.ranges):
        if 10 <= merged_range.min_row <= 28: sheet.unmerge_cells(str(merged_range))
    for r in range(10, 29):
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'L', 'M']: sheet[f'{col}{r}'] = None
        for c_idx in range(1, 14): sheet.cell(row=r, column=c_idx).border = thin_border
    current_row = 10
    patient_number = 1
    for p in patient_groups_in_page:
        count = len(p["Drugs"])
        if count == 0: continue
        end_row = current_row + count - 1
        sheet[f'A{current_row}'] = patient_number
        sheet[f'B{current_row}'] = sanitize_excel_text(p["Tarikh"])
        sheet[f'C{current_row}'] = "HOSPITAL BINTULU"
        sheet[f'D{current_row}'] = sanitize_excel_text(p["Nama"])
        sheet[f'E{current_row}'].value = sanitize_excel_text(str(p["No KP"]))
        sheet[f'E{current_row}'].number_format = '@'
        if count > 1:
            for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                sheet.merge_cells(start_row=current_row, start_column=column_index_from_string(col), end_row=end_row, end_column=column_index_from_string(col))
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            sheet[f'{col}{current_row}'].font, sheet[f'{col}{current_row}'].alignment = font_style, align_center
        for idx, drug in enumerate(p["Drugs"]):
            r = current_row + idx
            sheet[f'G{r}'].value = sanitize_excel_text(drug["Name"])
            sheet[f'L{r}'].value = sanitize_excel_text(drug["Dosage"])
            raw_qty = sanitize_excel_text(str(drug["Quantity"]))
            sheet[f'M{r}'].value = int(raw_qty) if raw_qty.isdigit() else raw_qty
            for c in ['G', 'L', 'M']: 
                sheet[f'{c}{r}'].font = font_style
                sheet[f'{c}{r}'].alignment = align_center if c != 'G' else align_left
        current_row += count
        patient_number += 1
    wb.save(output_path)


# --- 5. UI ROUTING ---
st.markdown('<div class="ye-title">S P U B</div>', unsafe_allow_html=True)
st.markdown('<div class="ye-subtitle">DATA EXTRACTION ENGINE</div>', unsafe_allow_html=True)

app_mode = st.radio("SELECT MODULE", ["FORMS (KEW PS7 & LAMPIRAN 2)", "INDENT LIST (MASTER DATABASE)"], horizontal=True, label_visibility="collapsed")
st.write("---")

uploaded_pdfs = st.file_uploader("PROVIDE ARCHIVES (PDF)", type=["pdf"], accept_multiple_files=True)

if uploaded_pdfs:
    st.info(f"INPUT RECEIVED: {len(uploaded_pdfs)} DOCUMENT(S).")

    if app_mode == "FORMS (KEW PS7 & LAMPIRAN 2)":
        if not os.path.exists(KEW_TEMPLATE) or not os.path.exists(LAMP_TEMPLATE):
            st.error("SYSTEM ERROR: FORM TEMPLATES MISSING.")
        else:
            if st.button("INITIALIZE EXTRACTION", type="primary", use_container_width=True):
                with st.spinner("PROCESSING DATA..."):
                    all_patients = [extract_general_data(pdf) for pdf in uploaded_pdfs]
                    
                    if sum(len(p["Drugs"]) for p in all_patients) == 0:
                        st.error("EXTRACTION FAILED. NO USABLE TEXT.")
                    else:
                        generated_files_dict = {}
                        lampiran_pages, current_page_patients, current_page_used_rows = [], [], 0

                        for patient in all_patients:
                            remaining_drugs = list(patient["Drugs"])
                            while remaining_drugs:
                                space_left = LAMP_CAPACITY - current_page_used_rows
                                if space_left <= 0:
                                    lampiran_pages.append(current_page_patients)
                                    current_page_patients, current_page_used_rows, space_left = [], 0, LAMP_CAPACITY
                                take_drugs = remaining_drugs[:space_left]
                                remaining_drugs = remaining_drugs[space_left:]
                                current_page_patients.append({"Nama": patient["Nama"], "No KP": patient["No KP"], "Tarikh": patient["Tarikh"], "Drugs": take_drugs})
                                current_page_used_rows += len(take_drugs)
                        if current_page_patients: lampiran_pages.append(current_page_patients)

                        for page_idx, page_patients in enumerate(lampiran_pages):
                            out_name = f"LAMPIRAN_2_PT_{page_idx+1}.xlsx"
                            build_lampiran_page(page_patients, LAMP_TEMPLATE, out_name)
                            with open(out_name, "rb") as f: generated_files_dict[out_name] = f.read()
                            os.remove(out_name)

                        all_drugs_flat = [d for p in all_patients for d in p["Drugs"]]
                        kew_chunks = [all_drugs_flat[i:i + KEW_CAPACITY] for i in range(0, len(all_drugs_flat), KEW_CAPACITY)]
                        total_counter = 1

                        for page_idx, chunk in enumerate(kew_chunks):
                            out_name = f"KEW_PS7_PT_{page_idx+1}.xlsx"
                            build_kew_page(chunk, KEW_TEMPLATE, out_name, start_index=total_counter)
                            total_counter += len(chunk)
                            with open(out_name, "rb") as f: generated_files_dict[out_name] = f.read()
                            os.remove(out_name)

                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zf:
                            for filename, filedata in generated_files_dict.items(): zf.writestr(filename, filedata)
                        
                        st.session_state["form_results"] = {"patients": all_patients, "files": generated_files_dict, "zip": zip_buffer.getvalue()}

        if st.session_state.get("form_results"):
            st.success("PROCESS COMPLETE.")
            st.markdown('<br><div class="ye-subtitle" style="margin-bottom: 10px;">DATA LOG</div>', unsafe_allow_html=True)
            for p in st.session_state["form_results"]["patients"]:
                with st.expander(f"{p['Nama']} // {p['No KP']} // {len(p['Drugs'])} DRUGS", expanded=False):
                    st.dataframe(p["Drugs"], use_container_width=True) if p["Drugs"] else st.warning("EMPTY RECORD.")
            
            st.markdown('<br><div class="ye-subtitle" style="margin-bottom: 20px;">EXPORT ARCHIVES</div>', unsafe_allow_html=True)
            cols = st.columns(2)
            for idx, fname in enumerate(st.session_state["form_results"]["files"].keys()):
                with cols[idx % 2]: st.download_button(label=fname, data=st.session_state["form_results"]["files"][fname], file_name=fname, use_container_width=True)
            st.download_button("DOWNLOAD MASTER ARCHIVE (ZIP)", data=st.session_state["form_results"]["zip"], file_name="SPUB_FORMS.zip", mime="application/zip", use_container_width=True)

    elif app_mode == "INDENT LIST (MASTER DATABASE)":
        if not os.path.exists(INDENT_TEMPLATE):
            st.error(f"SYSTEM ERROR: MASTER TEMPLATE '{INDENT_TEMPLATE}' MISSING.")
        else:
            if st.button("INTEGRATE TO MASTER DATABASE", type="primary", use_container_width=True):
                with st.spinner("FILTERING RESTRICTED DRUGS..."):
                    indent_patients = [extract_indent_data(pdf) for pdf in uploaded_pdfs]
                    
                    if sum(len(p["Drugs"]) for p in indent_patients) == 0:
                        st.error("NO TARGET DRUGS DETECTED IN ARCHIVES.")
                    else:
                        out_path = append_to_indent_list(indent_patients, INDENT_TEMPLATE)
                        with open(out_path, "rb") as f:
                            st.session_state["indent_result"] = {"file": f.read(), "patients": indent_patients}
                        os.remove(out_path)

        if st.session_state.get("indent_result"):
            st.success("DATABASE INTEGRATION COMPLETE.")
            st.markdown('<br><div class="ye-subtitle" style="margin-bottom: 10px;">INDENT LOG (RESTRICTED DRUGS ONLY)</div>', unsafe_allow_html=True)
            for p in st.session_state["indent_result"]["patients"]:
                if p["Drugs"]:
                    with st.expander(f"{p['Nama']} // {p['MRN']} // {p['Hospital_Code']}", expanded=True):
                        st.dataframe(p["Drugs"], use_container_width=True)
            
            st.markdown('<br><div class="ye-subtitle" style="margin-bottom: 20px;">EXPORT DATABASE</div>', unsafe_allow_html=True)
            st.download_button("DOWNLOAD UPDATED INDENT LIST", data=st.session_state["indent_result"]["file"], file_name="UPDATED_INDENT_LIST.xlsx", use_container_width=True)
