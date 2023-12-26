import streamlit as st
import pandas as pd
import tabula
from io import BytesIO
import os
import shutil

def extract_tables_from_pdf(pdf_file):
    tables = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True)
    return tables

def main():
    st.title("PDF Table Alchemy")

    pdf_file = st.file_uploader("Upload a PDF file", type=["pdf"])

    if pdf_file is not None:
        st.sidebar.subheader("Uploaded PDF:")
        st.sidebar.write(pdf_file.name)

        tables = extract_tables_from_pdf(pdf_file)

        if tables:
            st.sidebar.subheader("Number of Tables Extracted:")
            st.sidebar.write(len(tables))

            output_folder = "output_tables"
            os.makedirs(output_folder, exist_ok=True)

            for i, table in enumerate(tables):
                excel_filename = os.path.join(output_folder, f"table_{i + 1}.xlsx")
                table.to_excel(excel_filename, index=False, engine='openpyxl')
                st.success(f"Table {i + 1} extracted. [Download Excel File]({excel_filename})")

            # Create a zip file containing all Excel files
            zip_filename = "output_tables.zip"
            shutil.make_archive(os.path.splitext(zip_filename)[0], 'zip', output_folder)

            # Provide a link to download the zip file
            st.sidebar.success(f"[Download All Tables as Zip]({zip_filename})")

if __name__ == "__main__":
    main()
