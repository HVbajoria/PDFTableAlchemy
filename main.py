import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import tabula
import os

# Function to extract tables from PDF and create a zip archive
def extract_and_zip(pdf_file):
    tables = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True)
    dataframes = {}
    for i, table in enumerate(tables):
        filename = f"Table_{i + 1}.xlsx"
        dataframes[filename] = table

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, df in dataframes.items():
            # Convert DataFrame to Excel file in memory
            excel_buffer = BytesIO()
            df.to_excel(excel_buffer, index=True, engine='openpyxl')
            # Add Excel file to the zip archive
            zip_file.writestr(filename, excel_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

def main():
    st.title("PDF Table Extractor and Zip Download")

    pdf_file = st.file_uploader("Upload a PDF file", type=["pdf"])

    if pdf_file is not None:
        st.sidebar.subheader("Uploaded PDF:")
        st.sidebar.write(pdf_file.name)

        # Display extracted tables
        tables = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True)
        for i, table in enumerate(tables):
            st.subheader(f"Table {i + 1}")
            st.dataframe(table)

        # Download button
        if st.button('Download Tables as Zip'):
            zip_buffer = extract_and_zip(pdf_file)
            st.download_button(label='Download Zip', data=zip_buffer, file_name='tables.zip', key='download_button')

if __name__ == "__main__":
    main()
