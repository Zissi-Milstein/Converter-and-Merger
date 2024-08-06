import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document
from PyPDF2 import PdfMerger
import os
from excel2pdf import convert

def excel_to_pdf(excel_path, pdf_path, orientation='landscape'):
    # Convert Excel to PDF using excel2pdf
    convert(excel_path, pdf_path, orientation=orientation)

def docx_to_text(docx_path):
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def text_to_pdf(text, pdf_path, orientation='portrait'):
    if orientation == 'landscape':
        page_size = landscape(letter)
    elif orientation == 'portrait':
        page_size = portrait(letter)
    else:
        raise ValueError("Orientation must be 'landscape' or 'portrait'")

    pdf = SimpleDocTemplate(pdf_path, pagesize=page_size,
                            leftMargin=0.5 * inch, rightMargin=0.5 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.5 * inch)
    styles = getSampleStyleSheet()
    styleN = styles['Normal']

    paragraphs = [Paragraph(par, styleN) for par in text.split("\n")]
    pdf.build(paragraphs)

def merge_pdfs(pdf_paths, output_path):
    merger = PdfMerger()
    for pdf in pdf_paths:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

st.title("Document to PDF Converter and Merger")

uploaded_files = st.file_uploader("Choose files", type=["xlsx", "docx", "pdf"], accept_multiple_files=True)
orientation = st.radio("Choose orientation for Excel and Word files", ('landscape', 'portrait'))

if uploaded_files:
    pdf_paths = []
    for uploaded_file in uploaded_files:
        file_extension = os.path.splitext(uploaded_file.name)[1]
        if file_extension in [".xlsx", ".docx"]:
            with open(uploaded_file.name, "wb") as f:
                f.write(uploaded_file.getbuffer())
            input_path = uploaded_file.name
            output_path = f"{os.path.splitext(uploaded_file.name)[0]}.pdf"
            if file_extension == ".xlsx":
                excel_to_pdf(input_path, output_path, orientation)
            elif file_extension == ".docx":
                text = docx_to_text(input_path)
                text_to_pdf(text, output_path, orientation)
            pdf_paths.append(output_path)
        elif file_extension == ".pdf":
            with open(uploaded_file.name, "wb") as f:
                f.write(uploaded_file.getbuffer())
            pdf_paths.append(uploaded_file.name)

    if pdf_paths:
        merged_output_path = "merged_output.pdf"
        merge_pdfs(pdf_paths, merged_output_path)
        with open(merged_output_path, "rb") as pdf_file:
            pdf_data = pdf_file.read()

        st.success("PDFs merged successfully!")
        st.download_button(label="Download Merged PDF", data=pdf_data, file_name=merged_output_path, mime="application/pdf")
