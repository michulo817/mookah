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
        full_text
