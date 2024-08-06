import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document
import os

def excel_to_pdf(excel_path, pdf_path, orientation='landscape'):
    df = pd.read_excel(excel_path)
    df.columns = [col if 'Unnamed' not in col else '' for col in df.columns]
    df = df.fillna('')
    
    if orientation == 'landscape':
        page_size = landscape(letter)
    elif orientation == 'portrait':
        page_size = portrait(letter)
    else:
        raise ValueError("Orientation must be 'landscape' or 'portrait'")

    left_margin = right_margin = top_margin = bottom_margin = 0.5 * inch
    effective_page_width = page_size[0] - left_margin - right_margin
    num_columns = len(df.columns)
    column_width = effective_page_width / num_columns

    pdf = SimpleDocTemplate(pdf_path, pagesize=page_size,
                            leftMargin=left_margin, rightMargin=right_margin,
                            topMargin=top_margin, bottomMargin=bottom_margin)
    elements = []

    data = [df.columns.to_list()] + df.values.tolist()
    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    data = [[Paragraph(str(cell), styleN) for cell in row] for row in data]

    table = Table(data, colWidths=[column_width] * num_columns)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10)
    ])
    table.setStyle(style)
    elements.append(table)
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

st.title("Document to PDF Converter")

uploaded_file = st.file_uploader("Choose a file", type=["xlsx", "docx"])
orientation = st.radio("Choose orientation", ('landscape', 'portrait'))

if uploaded_file is not None:
    file_extension = os.path.splitext(uploaded_file.name)[1]
    if file_extension == ".xlsx":
        with open(uploaded_file.name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        input_path = uploaded_file.name
        output_path = "output.pdf"
        excel_to_pdf(input_path, output_path, orientation)
    elif file_extension == ".docx":
        with open(uploaded_file.name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        input_path = uploaded_file.name
        output_path = "output.pdf"
        text = docx_to_text(input_path)
        text_to_pdf(text, output_path, orientation)
    else:
        st.error("Unsupported file type!")

    with open(output_path, "rb") as pdf_file:
        pdf_data = pdf_file.read()

    st.success("PDF generated successfully!")
    st.download_button(label="Download PDF", data=pdf_data, file_name=output_path, mime="application/pdf")
