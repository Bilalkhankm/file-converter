import streamlit as st
import os
import io
import re
import json
import base64
import tempfile
import datetime
import csv as csv_mod
from pathlib import Path
from typing import Optional, List, Tuple

import pandas as pd
import chardet
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from PIL import Image
from groq import Groq

# --- Library Handlers (Graceful Fallbacks) ---
@st.cache_resource
def check_dependencies():
    status = {"pdf2docx": False, "pdf2image": False, "pytesseract": False}
    try:
        from pdf2docx import Converter
        status["pdf2docx"] = True
    except ImportError: pass
    
    try:
        from pdf2image import convert_from_path
        status["pdf2image"] = True
    except ImportError: pass

    try:
        import pytesseract
        status["pytesseract"] = True
    except ImportError: pass
    return status

DEPS = check_dependencies()

# --- UI CONFIG & STYLING ---
st.set_page_config(page_title="AI File Converter Pro", page_icon="🔄", layout="wide")

CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500&display=swap');
    
    .stApp { background-color: #0d1b2a; color: #f4f8ff; }
    h1, h2, h3 { font-family: 'Syne', sans-serif !important; color: #f0c040 !important; }
    .stMarkdown, p { font-family: 'DM Sans', sans-serif !important; }
    
    .main-header {
        background: linear-gradient(135deg, #1b2e45 0%, #243b55 100%);
        padding: 2rem;
        border-radius: 15px;
        border-left: 5px solid #f0c040;
        margin-bottom: 2rem;
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #e8a800, #f0c040) !important;
        color: #0d1b2a !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 8px !important;
        width: 100%;
    }
    
    .status-card {
        background: rgba(255, 255, 255, 0.05);
        padding: 1rem;
        border-radius: 10px;
        border: 1px solid rgba(240, 192, 64, 0.2);
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# --- GROQ INITIALIZATION ---
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
try:
    if not GROQ_API_KEY:
        groq_client = None
        GROQ_AVAILABLE = False
    else:
        groq_client = Groq(api_key=GROQ_API_KEY)
        GROQ_AVAILABLE = True
except Exception:
    GROQ_AVAILABLE = False

# --- CORE UTILITIES ---
def style_excel_sheet(ws):
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    alt_fill    = PatternFill(start_color="EBF0FA", end_color="EBF0FA", fill_type="solid")
    border_side = Side(style="thin", color="CCCCCC")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        fill = alt_fill if i % 2 == 0 else None
        for cell in row:
            if fill: cell.fill = fill
            cell.border = thin_border

def read_docx_text(file_bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

# --- AI HELPERS ---
def ai_extract_tables(text: str):
    if not GROQ_AVAILABLE: return []
    prompt = f"Extract structured data from this text into a JSON list of objects with 'name', 'headers', and 'rows'. Text:\n{text[:5000]}"
    try:
        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            response_format={"type": "json_object"}
        )
        data = json.loads(response.choices[0].message.content)
        return data.get("tables", []) if isinstance(data.get("tables"), list) else []
    except: return []

# --- CONVERSION LOGIC ---
def convert_logic(uploaded_file, target_type, use_ai=False):
    file_bytes = uploaded_file.getvalue()
    file_name = Path(uploaded_file.name).stem
    
    # [PDF -> Word]
    if target_type == "DOCX" and uploaded_file.name.endswith(".pdf"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(file_bytes)
            tmp_pdf_path = tmp_pdf.name
        
        out_path = f"{file_name}_converted.docx"
        try:
            from pdf2docx import Converter
            cv = Converter(tmp_pdf_path)
            output_buffer = io.BytesIO()
            cv.convert(output_buffer)
            cv.close()
            return output_buffer.getvalue(), out_path, "success"
        except Exception as e:
            return None, None, f"Error: {str(e)}"

    # [Word -> Excel]
    if target_type == "XLSX" and uploaded_file.name.endswith(".docx"):
        doc = Document(io.BytesIO(file_bytes))
        wb = openpyxl.Workbook()
        tables_added = 0
        
        for i, table in enumerate(doc.tables, 1):
            ws = wb.create_sheet(title=f"Table_{i}")
            for r_idx, row in enumerate(table.rows, 1):
                for c_idx, cell in enumerate(row.cells, 1):
                    ws.cell(row=r_idx, column=c_idx, value=cell.text.strip())
            style_excel_sheet(ws)
            tables_added += 1
            
        if use_ai and GROQ_AVAILABLE:
            text = read_docx_text(file_bytes)
            ai_tabs = ai_extract_tables(text)
            for tbl in ai_tabs:
                ws = wb.create_sheet(title=str(tbl.get("name", "AI_Extract"))[:30])
                ws.append(tbl.get("headers", []))
                for r in tbl.get("rows", []): ws.append(r)
                style_excel_sheet(ws)
        
        if "Sheet" in wb.sheetnames: del wb["Sheet"]
        out_buffer = io.BytesIO()
        wb.save(out_buffer)
        return out_buffer.getvalue(), f"{file_name}.xlsx", "success"

    # [CSV -> Excel]
    if target_type == "XLSX" and uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
        out_buffer = io.BytesIO()
        with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Data")
            style_excel_sheet(writer.book.active)
        return out_buffer.getvalue(), f"{file_name}.xlsx", "success"

    return None, None, "Unsupported conversion path."

# --- MAIN APP INTERFACE ---
def main():
    st.markdown("""
        <div class="main-header">
            <h1>🔄 AI File Converter Pro</h1>
            <p>Convert Documents with Enterprise-Grade Precision</p>
        </div>
    """, unsafe_allow_html=True)

    # Sidebar Status
    with st.sidebar:
        st.header("🛠️ System Status")
        st.write(f"Groq AI: {'🟢 Active' if GROQ_AVAILABLE else '🔴 Offline'}")
        st.write(f"PDF Engine: {'🟢 Ready' if DEPS['pdf2docx'] else '🔴 Missing'}")
        
        st.divider()
        st.info("Note: For high-accuracy OCR on scanned PDFs, ensure Tesseract is installed on your server.")

    # Tabs
    tab1, tab2, tab3 = st.tabs(["📄 Document Conversion", "🤖 AI Summarizer", "⚡ Batch Mode"])

    with tab1:
        col1, col2 = st.columns([1, 1])
        with col1:
            u_file = st.file_uploader("Upload File", type=["pdf", "docx", "csv", "xlsx"])
            target = st.selectbox("Convert To", ["DOCX", "PDF", "XLSX", "CSV"])
            ai_help = st.checkbox("Enable AI Enhanced Extraction", value=True) if GROQ_AVAILABLE else False
            
            if st.button("Start Conversion"):
                if u_file:
                    with st.spinner("Processing..."):
                        data, name, status = convert_logic(u_file, target, ai_help)
                        if data:
                            st.success(f"Successfully converted to {name}")
                            st.download_button("⬇️ Download Result", data, file_name=name)
                        else:
                            st.error(status)
                else:
                    st.warning("Please upload a file first.")

    with tab2:
        st.subheader("Summarize Document")
        s_file = st.file_uploader("Upload for Summary", type=["docx", "txt", "pdf"], key="summ")
        if st.button("✨ Generate Summary"):
            if not GROQ_AVAILABLE:
                st.error("Groq API Key missing.")
            elif s_file:
                with st.spinner("Thinking..."):
                    text = ""
                    if s_file.name.endswith(".docx"): text = read_docx_text(s_file.getvalue())
                    else: text = s_file.getvalue().decode()
                    
                    response = groq_client.chat.completions.create(
                        model="llama-3.3-70b-versatile",
                        messages=[{"role": "user", "content": f"Summarize this document in bullets:\n\n{text[:8000]}"}]
                    )
                    st.markdown("### Summary")
                    st.write(response.choices[0].message.content)

    with tab3:
        st.subheader("Batch Processing")
        b_files = st.file_uploader("Upload multiple files", accept_multiple_files=True)
        if st.button("🚀 Process Batch"):
            if b_files:
                for f in b_files:
                    st.write(f"Processing {f.name}...")
                    # Simulating batch loop logic here

if __name__ == "__main__":
    main()