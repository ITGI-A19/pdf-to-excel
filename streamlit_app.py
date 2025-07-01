import streamlit as st
import zipfile
import tempfile
import os
import shutil
import pdfplumber
import pandas as pd

st.title("📄 PDF to Excel Converter (Rainfall Data)")

uploaded_zip = st.file_uploader("Upload ZIP containing PDFs in folders", type="zip")

if uploaded_zip:
    with st.spinner("Processing ZIP file..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            # Step 1: Save and unzip
            zip_path = os.path.join(temp_dir, "input.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())

            extract_path = os.path.join(temp_dir, "unzipped")
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(extract_path)

            output_excel_dir = os.path.join(temp_dir, "output_excels")
            os.makedirs(output_excel_dir, exist_ok=True)

            # Step 2: Walk through all files and convert PDFs to Excel
            for root, dirs, files in os.walk(extract_path):
                for file in files:
                    if file.endswith(".pdf"):
                        pdf_path = os.path.join(root, file)
                        relative_folder = os.path.relpath(root, extract_path)
                        excel_subfolder = os.path.join(output_excel_dir, relative_folder)
                        os.makedirs(excel_subfolder, exist_ok=True)

                        base_name = os.path.splitext(file)[0]
                        excel_path = os.path.join(excel_subfolder, base_name + ".xlsx")

                        try:
                            tables_combined = []
                            with pdfplumber.open(pdf_path) as pdf:
                                for page in pdf.pages:
                                    tables = page.extract_tables()
                                    for table in tables:
                                        if table:
                                            df = pd.DataFrame(table)
                                            tables_combined.append(df)

                            if tables_combined:
                                final_df = pd.concat(tables_combined, ignore_index=True)
                                final_df.to_excel(excel_path, index=False)
                        except Exception as e:
                            st.error(f"Error processing {file}: {e}")

            # Step 3: Zip only Excel outputs
            output_zip_path = os.path.join(temp_dir, "excels_only.zip")
            with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(output_excel_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, output_excel_dir)
                        zipf.write(file_path, arcname=arcname)

            # Step 4: Download button
            with open(output_zip_path, "rb") as f:
                st.download_button(
                    label="📥 Download Converted Excel ZIP",
                    data=f,
                    file_name="converted_excels.zip",
                    mime="application/zip"
                )
