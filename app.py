import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from PIL import Image
from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document
import os
from PyPDF2 import PdfMerger

def excel_to_image(excel_path, image_path):
    df = pd.read_excel(excel_path, engine='openpyxl')
    df.columns = [col if 'Unnamed' not in col else '' for col in df.columns]
    df = df.fillna('')
    
    fig, ax = plt.subplots(figsize=(12, len(df) * 0.3))  # Adjust size according to the data length
    ax.axis('off')
    
    # Create a table
    table = plt.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')
    
    # Style the table
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.2)
    
    plt.savefig(image_path, bbox_inches='tight', pad_inches=0.1)
    plt.close()

def image_to_pdf(image_path, pdf_path, orientation='portrait'):
    pdf_pages = PdfPages(pdf_path)
    images = [Image.open(image_path)]
    pdf_pages.savefig(images[0])
    pdf_pages.close()

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
                image_path = f"{os.path.splitext(uploaded_file.name)[0]}.png"
                excel_to_image(input_path, image_path)
                image_to_pdf(image_path, output_path, orientation)
                os.remove(image_path)
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
