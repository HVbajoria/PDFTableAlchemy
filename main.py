import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import tabula
import os
import subprocess
import platform

hide_default_format = """ 
        <style> 
        footer {visibility: hidden;} 
        </style> 
        """ 
st.markdown(hide_default_format, unsafe_allow_html=True) 

def gradient_text(text, color1, color2):
    gradient_css = f"""
        background: -webkit-linear-gradient(left, {color1}, {color2});
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: bold;
        font-size: 42px;
    """
    return f'<span style="{gradient_css}">{text}</span>'

def gradient(text, color1, color2):
    gradient_css = f"""
        background: -webkit-linear-gradient(left, {color2}, {color1});
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: bold;
        font-size: 22px;
    """
    return f'<span style="{gradient_css}">{text}</span>'

color1 = "#0d3270"
color2 = "#0fab7b"
text = "PDFTableAlchemy"
  
# left_co, cent_co,last_co = st.columns(3)
# with cent_co:
#     st.image("images/logo.png", width=200)

styled_text = gradient_text(text, color1, color2)
st.write(f"<div style='text-align: center;'>{styled_text}</div>", unsafe_allow_html=True)

text="Transforming PDF chaos into Excel gold with Table Alchemy!"
styled_text = gradient(text, color1, color2)
st.write(f"<div style='text-align:center;'>{styled_text}</div>",unsafe_allow_html=True)

# Get path where Java is installed
def get_java_path():
    # Command to get the Java executable path
    command = "which java" if platform.system() != "Windows" else "where java"
    
    try:
        # Execute the command and capture the output
        java_path = subprocess.check_output(command, shell=True, text=True)
        return java_path.strip()
    except subprocess.CalledProcessError:
        return "Java not found or an error occurred."

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
    java_path = get_java_path()
    st.write(java_path)
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
        zip_buffer = extract_and_zip(pdf_file)
        st.sidebar.download_button(label='Download The Files', data=zip_buffer, file_name='tables.zip', key='download_button')

if __name__ == "__main__":
    main()

footer="""<style>

a:hover,  a:active {
color: red;
background-color: transparent;
text-decoration: underline;
}

.footer {
position: fixed;
left: 0;
bottom: 0;
width: 100%;
background-color: white;
color: black;
text-align: center;
}
</style>
<div class="footer">
<p>Developed with ❤️ by <a style='display: inline; text-align: center;' href="https://www.linkedin.com/in/harshavardhan-bajoria/" target="_blank">Harshavardhan Bajoria</a></p>
</div>
"""
st.markdown(footer,unsafe_allow_html=True)