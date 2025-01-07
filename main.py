import streamlit as st
import subprocess
import os
import zipfile
import shutil
from io import BytesIO

# Function to convert Excel files to PDF using LibreOffice
def excel_to_pdf(input_file, output_folder):
    command = f'libreoffice --headless --convert-to pdf --outdir "{output_folder}" "{input_file}"'
    subprocess.run(command, shell=True)

# Function to convert DOCX files to PDF using LibreOffice
def docx_to_pdf(input_file, output_folder):
    command = f'libreoffice --headless --convert-to pdf --outdir "{output_folder}" "{input_file}"'
    subprocess.run(command, shell=True)

# Function to create a zip file from the converted PDFs
def zip_pdfs(pdf_folder, zip_filename):
    zip_path = f"/tmp/{zip_filename}"
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(pdf_folder):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), pdf_folder))
    return zip_path

# Streamlit app UI
st.title("Convert Files to PDF")

# Upload multiple files
uploaded_files = st.file_uploader("Upload Excel (.xlsx) and Word (.docx) files", type=["xlsx", "xls", "docx", "doc"], accept_multiple_files=True)

if uploaded_files:
    # Create a temporary folder to store converted PDFs
    pdf_folder = "/tmp/converted_pdfs"
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    # Display a spinner while converting files
    with st.spinner("Converting files to PDF..."):
        # Process each uploaded file
        for uploaded_file in uploaded_files:
            input_path = os.path.join("/tmp", uploaded_file.name)

            # Save the uploaded file temporarily
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Convert the file to PDF
            output_pdf_path = os.path.join(pdf_folder, f"{os.path.splitext(uploaded_file.name)[0]}.pdf")

            if uploaded_file.name.endswith(".xlsx"):
                excel_to_pdf(input_path, pdf_folder)
            if uploaded_file.name.endswith(".xls"):
                excel_to_pdf(input_path, pdf_folder)
            elif uploaded_file.name.endswith(".doc"):
                docx_to_pdf(input_path, pdf_folder)
            elif uploaded_file.name.endswith(".docx"):
                docx_to_pdf(input_path, pdf_folder)

        # Create a zip file containing all converted PDFs
        zip_filename = "converted_pdfs.zip"
        zip_path = zip_pdfs(pdf_folder, zip_filename)

    # Provide a download link for the zip file
    with open(zip_path, "rb") as f:
        st.download_button(
            label="Download Zipped PDFs",
            data=f.read(),
            file_name=zip_filename,
            mime="application/zip"
        )

    # Clean up temporary files
    shutil.rmtree(pdf_folder)
