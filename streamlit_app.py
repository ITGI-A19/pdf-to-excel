import streamlit as st
import zipfile
import tempfile
import os
import shutil
import pdfplumber
import pandas as pd
import time

st.set_page_config(page_title="PDF to Excel with Progress", layout="wide")
st.title("📄 PDF to Excel Converter with Folder Structure + Progress + Logs")

uploaded_zip = st.file_uploader("📦 Upload a ZIP file containing folders with PDFs", type="zip")

if uploaded_zip:
    with st.spinner("⚙️ Unpacking and preparing..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, "input.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())

            extract_path = os.path.join(temp_dir, "unzipped")
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(extract_path)

            output_excel_dir = os.path.join(temp_dir, "output_excels")
            os.makedirs(output_excel_dir, exist_ok=True)

            # Gather all PDF files
            pdf_files = []
            for root, dirs, files in os.walk(extract_path):
                for file in files:
                    if file.lower().endswith(".pdf"):
                        pdf_files.append(os.path.join(root, file))

            total = len(pdf_files)
            progress_bar = st.progress(0)
            status_text = st.empty()
            log_messages = []

            start_time = time.time()
            success_count = 0
            fail_count = 0

            for idx, pdf_path in enumerate(pdf_files, 1):
                relative_folder = os.path.relpath(os.path.dirname(pdf_path), extract_path)
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                excel_subfolder = os.path.join(output_excel_dir, relative_folder)
                os.makedirs(excel_subfolder, exist_ok=True)
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
                        success_count += 1
                        log_messages.append(f"✅ {base_name}.pdf → Excel converted")
                    else:
                        log_messages.append(f"⚠️ {base_name}.pdf → No tables found")

                except Exception as e:
                    fail_count += 1
                    log_messages.append(f"❌ {base_name}.pdf → Error: {str(e)}")

                progress_bar.progress(idx / total)
                status_text.markdown(f"**Processed {idx} of {total} PDFs**")

            elapsed = round(time.time() - start_time, 2)

            # Create ZIP of Excel outputs
            output_zip_path = os.path.join(temp_dir, "excels_only.zip")
            with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(output_excel_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, output_excel_dir)
                        zipf.write(file_path, arcname)

            # Summary
            st.success("✅ All processing completed!")
            st.markdown(f"""
                - 🗂️ Total PDFs: **{total}**
                - ✅ Successes: **{success_count}**
                - ❌ Failures: **{fail_count}**
                - ⏱️ Time Taken: **{elapsed} seconds**
            """)

            # Logs
            with st.expander("🪵 Detailed Logs"):
                for msg in log_messages:
                    st.markdown(f"- {msg}")

            # Download ZIP
            with open(output_zip_path, "rb") as f:
                st.download_button(
                    label="📥 Download Excel ZIP",
                    data=f,
                    file_name="converted_excels.zip",
                    mime="application/zip"
                )
