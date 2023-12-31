import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import tabula
from form_recognizer_azure import analyze_document

st.set_page_config( 
    page_title="PDFTableAlchemy", 
    page_icon="🗃️", 
)

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

styled_text = gradient_text(text, color1, color2)
st.write(f"<div style='text-align: center;'>{styled_text}</div>", unsafe_allow_html=True)

text="Transforming PDF chaos into Excel gold with Table Alchemy!"
styled_text = gradient(text, color1, color2)
st.write(f"<div style='text-align:center;'>{styled_text}</div>", unsafe_allow_html=True)

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

def save_uploaded_file(uploaded_file):
    file_path = uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

def save_result_file(dataframes):
    result_path = "result.xlsx"
    with pd.ExcelWriter(result_path, engine='xlsxwriter') as writer:
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return result_path

def main():
    pdf_file = st.file_uploader("Upload a PDF file", type=["pdf"])

    if pdf_file is not None:
        st.sidebar.subheader("Uploaded PDF:")
        st.sidebar.write(pdf_file.name)

        # Save the uploaded file
        uploaded_file_path = save_uploaded_file(pdf_file)

        # Display extracted tables
        tables = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True)
        for i, table in enumerate(tables):
            st.subheader(f"Table {i + 1}")
            st.dataframe(table)

        # Download button
        zip_buffer = extract_and_zip(pdf_file)
        download_button = st.sidebar.download_button(label='Download The Files', data=zip_buffer, file_name='tables.zip', key='download_button')

        # Reprocess button
        if st.sidebar.button("Not Happy? Reprocess with Azure AI"):
            reprocessed_tables = analyze_document(False,uploaded_file_path)
            
if __name__ == "__main__":
    main()

footer = """<style>

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
st.markdown(footer, unsafe_allow_html=True)
