import streamlit as st
import os
import fitz
from openpyxl import Workbook

# Function to extract annotations and write to Excel
def extract_annotations_fitz_to_excel(directory):
    xlsx_file_path = os.path.join(directory, "annotations_output.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(['File Name', 'Page Number', 'Annotation'])

    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            doc = fitz.open(pdf_path)

            for page_number, page in enumerate(doc):
                for annot in page.annots():
                    annot_text = annot.info["content"]
                    if annot_text:
                        ws.append([filename, page_number + 1, annot_text])

            doc.close()

    wb.save(xlsx_file_path)
    return xlsx_file_path

# Streamlit UI
st.title("PDF Annotations Extractor")

# User input for directory
directory = st.text_input("Enter the directory of your PDF files:")

# Button to trigger annotation extraction
if st.button("Extract Annotations"):
    if directory and os.path.isdir(directory):
        output_file = extract_annotations_fitz_to_excel(directory)
        st.success(f"Annotations extracted successfully! File saved at: {output_file}")
    else:
        st.error("Please enter a valid directory.")
