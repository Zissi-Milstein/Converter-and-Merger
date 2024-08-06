import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
import os

def excel_to_pdf(excel_path, pdf_path, orientation='landscape'):
    df = pd.read_excel(excel_path)
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

st.title("Excel to PDF Converter")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
orientation = st.radio("Choose orientation", ('landscape', 'portrait'))

if uploaded_file is not None:
    with open(uploaded_file.name, "wb") as f:
        f.write(uploaded_file.getbuffer())
    input_path = uploaded_file.name
    output_path = "output.pdf"
    excel_to_pdf(input_path, output_path, orientation)
    st.success("PDF generated successfully!")
    st.download_button(label="Download PDF", file_name=output_path, mime="application/pdf")
