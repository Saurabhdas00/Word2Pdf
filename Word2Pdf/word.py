import streamlit as st
from docx2pdf import convert
import tempfile
import os
import pythoncom
from pdf2docx import Converter  # For PDF to Word conversion
import pandas as pd  # For Excel to PDF conversion (requires additional libraries)

st.title("File Format Converter")

# Sidebar for conversion options
st.sidebar.title("Conversion Options")
conversion_type = st.sidebar.selectbox(
    "Select Conversion Type",
    ["Word to PDF", "PDF to Word", "Excel to PDF"]  # Add more options as needed
)

# File uploader
uploaded_file = st.file_uploader(f"Upload a file for {conversion_type}", type=["docx", "pdf", "xlsx"])

if uploaded_file is not None:
    try:
        # Save the uploaded file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as temp_file:
            temp_file.write(uploaded_file.getbuffer())
            temp_file_path = temp_file.name

        # Initialize the COM library (required for Windows)
        pythoncom.CoInitialize()

        # Perform the selected conversion
        if conversion_type == "Word to PDF":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf_file:
                temp_pdf_path = temp_pdf_file.name

            # Convert Word to PDF
            convert(temp_file_path, temp_pdf_path)

            # Read the converted PDF file
            with open(temp_pdf_path, "rb") as pdf_file:
                converted_bytes = pdf_file.read()

            # Provide a download button for the PDF
            st.download_button(
                label="Download PDF",
                data=converted_bytes,
                file_name="converted.pdf",
                mime="application/pdf"
            )

        elif conversion_type == "PDF to Word":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx_file:
                temp_docx_path = temp_docx_file.name

            # Convert PDF to Word
            cv = Converter(temp_file_path)
            cv.convert(temp_docx_path)
            cv.close()

            # Read the converted Word file
            with open(temp_docx_path, "rb") as docx_file:
                converted_bytes = docx_file.read()

            # Provide a download button for the Word file
            st.download_button(
                label="Download Word",
                data=converted_bytes,
                file_name="converted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        elif conversion_type == "Excel to PDF":
            # Convert Excel to PDF (requires additional libraries like pandas and fpdf)
            st.warning("Excel to PDF conversion is not implemented in this example. You can use libraries like `pandas` and `fpdf` to achieve this.")

        st.success(f"{conversion_type} conversion successful! Click the button above to download the converted file.")

    except Exception as e:
        st.error(f"An error occurred during conversion: {e}")

    finally:
        # Clean up temporary files
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
        if conversion_type == "Word to PDF" and os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if conversion_type == "PDF to Word" and os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)

        # Uninitialize the COM library
        pythoncom.CoUninitialize()