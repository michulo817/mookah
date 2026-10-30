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
            dosage = re.sub(r'(\d)\s*Q\s*(\d)\s*D',)
