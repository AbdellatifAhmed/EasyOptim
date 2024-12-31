import pandas as pd
import ast
import streamlit as st
import tools as lataftaf
import os
import base64

output_dir = os.path.join(os.getcwd(), 'OutputFiles')
xls_PRFILEs = os.path.join(output_dir, 'PRFILE.xlsx')
st.set_page_config(
    page_title="EasyOptim - Audit PRFILE",
    layout="wide"
)

# Page Header
st.markdown(
    """
    <style>
    .header {
        background-color: #f8f9fa;
        padding: 20px;
        text-align: left;
        font-size: 30px;
        font-weight: bold;
        border-bottom: 2px solid #e0e0e0;
    }
    </style>
    <div class="header">
        EasyOptim - Audit PRFILE 
    </div>
    """,
    unsafe_allow_html=True,
)
with st.expander("Import RNCs PRFILEs", expanded=True):
    container1 = st.container()
    with container1:
        files_txt = st.file_uploader("Select RNCs PRFILEs:", type=[".txt"], accept_multiple_files=True)
        btn_checkPRFILE = st.button("Check PRFILEs", key="prfiles_button")
        if btn_checkPRFILE:
            output_status = lataftaf.audit_prfiles(files_txt)
            if output_status == "PRFILEs Preparation Done Successfully!":
                st.write(output_status)
                with open(xls_PRFILEs, "rb") as f:
                    file_data = f.read()
                    b64_file_data = base64.b64encode(file_data).decode()
                    download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(xls_PRFILEs)}">Click to download {os.path.basename(xls_PRFILEs)}</a>'
                st.markdown(download_link, unsafe_allow_html=True)
            else:
                st.error(output_status)



