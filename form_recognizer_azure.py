from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
import os
from dotenv import load_dotenv
import streamlit as st
# Load environment variables from .env file
load_dotenv()

endpoint = os.environ.get('AZURE_FORM_RECOGNIZER_ENDPOINT')
key = "483b7c8454bd40ac81998900dab733ea"

from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
import os
from dotenv import load_dotenv
import logging
import json
import os
import logging
from json import JSONEncoder
from azure.core.exceptions import ResourceNotFoundError
import pandas as pd
import io
import urllib.parse

# Load environment variables from .env file
load_dotenv()

endpoint = os.environ.get('AZURE_FORM_RECOGNIZER_ENDPOINT')
key = os.environ.get('AZURE_FORM_RECOGNIZER_KEY')

def get_key_value_pairs(result):
    kvp = {}
    pagekvp = {}
    pagelen= len(result.pages)
    pagenum=None
    currpagenum=None
    for kv_pair in result.key_value_pairs:
        if pagenum is None:
            pagenum=kv_pair.key.bounding_regions[0].page_number
        elif (pagenum is not None) and (pagenum != kv_pair.key.bounding_regions[0].page_number):
            pagekvp[pagenum]=kvp
            kvp = {}
            pagenum=kv_pair.key.bounding_regions[0].page_number

        if kv_pair.key:
            if kv_pair.value:
                kvp[kv_pair.key.content] = kv_pair.value.content
    pagekvp[pagenum]=kvp
    return pagekvp

# Function to handle the download
def download_excel(excel_file_path):
    with open(excel_file_path, "rb") as f:
        data = f.read()
    st.sidebar.download_button(
        label="Download Excel File",
        key="download_button",
        on_click=lambda: download_excel(excel_file_path),
        file_name="output.xlsx",
        data=data,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def generate_excel(result, filename, add_keyvalue_pairs):
    if add_keyvalue_pairs:
        kvp = get_key_value_pairs(result)

    formtables = {}
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 2})

    current_page_num = None
    table_num = 1
    sheet_data = {}

    for table in result.tables:
        column_row_spans = []
        tableList = [[None for x in range(table.column_count)] for y in range(table.row_count)]

        for cell in table.cells:
            cellvalue = None
            if sum(c.isalnum() for c in cell.content) > 0:
                cellvalue = cell.content.replace(":unselected:", "").replace(":selected:", "")
                if cell.row_span > 1 and cell.column_span > 1:
                    column_row_spans.append([cell.row_index, cell.column_index, cell.row_index + cell.row_span - 1,
                                             cell.column_index + cell.column_span - 1, cellvalue])
                elif cell.row_span > 1 and cell.column_span == 1:
                    column_row_spans.append([cell.row_index, cell.column_index, cell.row_index + cell.row_span - 1,
                                             cell.column_index, cellvalue])
                elif cell.column_span > 1 and cell.row_span == 1:
                    column_row_spans.append([cell.row_index, cell.column_index, cell.row_index,
                                             cell.column_index + cell.column_span - 1, cellvalue])

            tableList[cell.row_index][cell.column_index] = cellvalue

        if current_page_num is None:
            current_page_num = table.bounding_regions[0].page_number
        elif (current_page_num is not None) and (current_page_num == table.bounding_regions[0].page_number):
            table_num += 1
        elif (current_page_num is not None) and (current_page_num != table.bounding_regions[0].page_number):
            table_num = 1
            current_page_num = table.bounding_regions[0].page_number

        excel_sheet_name = str(current_page_num) + '_' + str(table_num)
        df = pd.DataFrame.from_records(tableList)
        df.columns = df.iloc[0]
        df = df[1:]
        df = df[df.any(axis=1)]

        if not df.empty:
            if add_keyvalue_pairs and current_page_num in kvp:
                for key in kvp[current_page_num]:
                    if key not in df.columns:
                        df[key] = kvp[current_page_num][key]

            df.to_excel(writer, sheet_name=excel_sheet_name, index=False, engine='openpyxl')
            
            worksheet = writer.sheets[excel_sheet_name]
            for x in column_row_spans:
                worksheet.merge_range(x[0], x[1], x[2], x[3], x[4], merge_format)

            # Store sheet data for later display in the Streamlit app
            sheet_data[excel_sheet_name] = df

    excelname = filename + '.xlsx'
    writer.close()
    logging.info("writing excel for : " + excelname)

    # Save the excel file to the output folder
    output.seek(0)
    with open(excelname, "wb") as f:
        f.write(output.read())

    # Provide a new download button for the reprocessed tables stored in the result.xlsx file in the backend
    with open(excelname, "rb") as f:
        data = f.read()
    st.sidebar.download_button(
        label="Download Advanced File",
        key="azure_download_button",
        file_name="advanced_result.xlsx",
        data=data,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    

    return 'Individual table per sheet has been generated successfully in Excel: ' + excelname

def analyze_document( add_keyvalue_pairs, link):
    try:
        filename = "result"

        document_analysis_client = DocumentAnalysisClient(
            endpoint=endpoint, credential=AzureKeyCredential(key)   
        )
        model="prebuilt-document"
        with open(link, "rb") as f:
            poller = document_analysis_client.begin_analyze_document(
            model, document=f
        )
        result = poller.result()
        output_record = generate_excel(result,filename,add_keyvalue_pairs)
        
    except Exception as error:
        output_record =   "Error: " + str(error)
        
    print("Output record: " + output_record)
    return output_record
