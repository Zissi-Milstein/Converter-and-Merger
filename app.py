import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document
from PyPDF2 import PdfMerger
from PIL import Image
import openpyxl
import os

def image_to_pdf(image_path, pdf_path, orientation='portrait'):
    # Verify if the image file is valid
    try:
        img = Image.open(image_path)
        img.verify()
    except Exception as e:
        st.error(f"Error opening image file: {e}")
        return

    if orientation == 'landscape':
        figsize = (11.69, 8.27)  # A4 landscape size
    else:
        figsize = (8.27, 11.69)  # A4 portrait size

    fig, ax = plt.subplots(figsize=figsize)
    img = mpimg.imread(image_path)
    ax.imshow(img)
    ax.axis('off')
    plt.savefig(pdf_path, format='pdf', bbox_inches='tight', pad_inches=0)
    plt.close(fig)

def excel_to_pdf(excel_path, pdf_path, orientation='landscape'):
    # Load workbook and get the active sheet
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb.active
    
    # Get the data from the sheet
    data = list(sheet.values)
    
    # Extract header and rows
    header = data[0]
    rows = data[1:]

    # Determine page size
    if orientation == 'landscape':
        page_size = landscape(letter)
    elif orientation == 'portrait':
        page_size = portrait(letter)
    else:
        raise ValueError("Orientation must be 'landscape' or 'portrait'")

    # Create a PDF document
    pdf = SimpleDocTemplate(pdf_path, pagesize=page_size,
                            leftMargin=0.5 * inch, rightMargin=0.5 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.5 * inch)
    elements = []

    # Prepare data for PDF
    formatted_data = [header] + rows
    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    
    # Convert data to Paragraphs with formatting
    data = []
    for row in formatted_data:
        formatted_row = []
        for cell in row:
            if isinstance(cell, (int, float)):
                formatted_row.append(Paragraph(f"{cell:,.2f}", styleN))  # Format numbers with commas and two decimal places
            elif isinstance(cell, str):
                # Handle date and currency formatting here if needed
                formatted_row.append(Paragraph(cell, styleN))
            else:
                formatted_row.append(Paragraph(str(cell), styleN))
        data.append(formatted_row)

    # Create the table
    num_columns = len(header)
    column_width = (page_size[0] - 1.5 * inch) / num_columns
    table = Table(data, colWidths=[column_width] * num_columns)
    
    # Define table style
    table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10)
    ])
    table.setStyle(table_style)
    elements.append(table)
    
    # Build PDF
    pdf.build(elements)

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
