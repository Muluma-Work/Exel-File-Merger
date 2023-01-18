# pip install openpyxl
import pandas as pd
import streamlit as st
import zipfile
import base64
import os
import numpy as np

# Web App Title
#st.title(":bar_chart: File Merger tool")
st.title("File Merger tool")
st.markdown("##")
st.sidebar.image(
    'https://muluma.co.za/wp-content/uploads/2021/08/Muluma-Logo-small.jpg', width=200)
#st.markdown('File Merger')
#st.set_page_config(page_title="Sales Dashboard")
# **Excel File Merger**)

# Excel file merge function


def excel_file_merge(zip_file_name):
    df = pd.DataFrame()
    archive = zipfile.ZipFile(zip_file_name, 'r')
    with zipfile.ZipFile(zip_file_name, "r") as f:
        for file in f.namelist():
            xlfile = archive.open(file)
            if file.endswith('.xlsx'):
                # Add a note indicating the file name that this dataframe originates from
                df_xl = pd.read_excel(xlfile, engine='openpyxl')
                df_xl['Note'] = file
                # Appends content of each Excel file iteratively
                df = df.append(df_xl, ignore_index=True)
    return df


# Upload CSV data
with st.sidebar.header('1. Upload your ZIP file here'):
    uploaded_file = st.sidebar.file_uploader(
        "This tool is going to merge all Excel files in the zip folder uploaded", type=["zip"])
    st.sidebar.markdown("""
[Download example ZIP input file](https://github.com/dataprofessor/excel-file-merge-app/raw/main/nba_data.zip)
""")

# File download


def filedownload(df):
    csv = df.to_csv(index=False)
    # strings <-> bytes conversions
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="new_merged_file.csv">Download Merged File as CSV</a>'
    return href


def xldownload(df):
    df.to_excel('data.xlsx', index=False)
    data = open('data.xlsx', 'rb').read()
    b64 = base64.b64encode(data).decode('UTF-8')
    # b64 = base64.b64encode(xl.encode()).decode()  # strings <-> bytes conversions
    href = f'<a href="data:file/xls;base64,{b64}" download="new_merged_file.xlsx">Download Merged File as XLSX</a>'
    return href


# Main panel
if st.sidebar.button('Submit zip file'):
    # @st.cache
    df = excel_file_merge(uploaded_file)
    st.header('**Merged data**')
    st.write(df)
    st.markdown(filedownload(df), unsafe_allow_html=True)
    st.markdown(xldownload(df), unsafe_allow_html=True)
else:
    st.info('Awaiting for ZIP file to be uploaded.')

st.sidebar.markdown('---')
st.sidebar.write('Developed by Tseke Maila')
st.sidebar.write('Contact at tseke@muluma.co.za')

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
