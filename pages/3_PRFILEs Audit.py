import streamlit as st
import pandas as pd
import tools as lataftaf
import pydeck as pdk
from pydeck.types import String
import base64

st.set_page_config(
    page_title="EasyOptim - PRFILE Audit",
    layout="wide"  # Use the full width of the page
)
st.title("PRFILE Audit Tool")
st.write("tool take input PRFILEs and share organized output")

DB_work_Log = lataftaf.get_log_file()
with open(DB_work_Log, "rb") as f:
    file_data = f.read()
    b64_file_data = base64.b64encode(file_data).decode()
    download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(DB_work_Log)}">Click to download {os.path.basename(DB_work_Log)}</a>'
st.markdown(download_link, unsafe_allow_html=True)

st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        bottom: 0;
        width: 100%;
        background-color: #f8f9fa;
        padding: 10px 0;
        text-align: left;
        font-size: 16px;
        border-top: 2px solid #e0e0e0;
    }
    </style>
    <div class="footer">
        The Tool developed by Abdellatif Ahmed (abdellatif.ahmed@nokia.com)
        
    </div>
    
    """,
    unsafe_allow_html=True,
)
