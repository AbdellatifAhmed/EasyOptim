import streamlit as st
import tools as lataftaf
import base64
import os
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="EasyOptim - Test Page",
    layout="wide"  # Use the full width of the page
)
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
        Hello Nokia Optimization Team
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("Please select the desired tool on side bar, also updated DB can be uploaded")
recent_dB = lataftaf.get_log()
# try:
#     with open(recent_dB[1], "rb") as f:
#         file_data = f.read()
#     b64_file_data = base64.b64encode(file_data).decode()  # Encode log content to base64
#     href_KML = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(recent_dB[1])}">{recent_dB[0]}</a>'
    
#     with open(recent_dB[3], "rb") as f:
#             file_data = f.read()
#     b64_file_data = base64.b64encode(file_data).decode()  # Encode log content to base64
#     href_Dmp = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(recent_dB[3])}">{recent_dB[2]}</a>'
    
# except Exception as e:
#         st.error(f"Error reading intial Files location: {e}")

if "expanded" not in st.session_state:
    st.session_state.expanded = True

with st.expander("Upload/Update Dump & KML File",expanded=st.session_state.expanded):
    cont1 = st.container()
    with cont1:
        col1,col2 = st.columns(2)
        try:
            col1.markdown("Current KML File: " + recent_dB[0], unsafe_allow_html=True)
        except:
            pass
        file_Kml = col1.file_uploader("Sites DB File:", type=["csv"])
        col1_1,col1_2 = col1.columns(2)
        btn_Kml = col1_1.button("Update KML")
        Is_Update_Nbrs = col1_2.checkbox("Check to update estimated Nbrs")
        try:
            col2.markdown("Current Dump File: " + recent_dB[2], unsafe_allow_html=True)
        except:
            pass
        file_Dmp = col2.file_uploader("Param Dump File:", type=["xlsb"])  
        btn_Dmp = col2.button("Update Dump")


with st.expander("Upload KPIs & Performance Reports",expanded=st.session_state.expanded):
    cont2 = st.container()
    with cont2:
        col2G,col3G,col4G,col5G = st.columns(4)
        col2G.file_uploader("2G Performance Report:", type=["csv"])
        col2G.button("Update 2G Report")
        
        col3G.file_uploader("3G Performance Report:", type=["csv"])
        col3G.button("Update 3G Report")
        
        col4G.file_uploader("4G Performance Report:", type=["csv"])
        col4G.button("Update 4G Report")
        
        col5G.file_uploader("5G Performance Report:", type=["csv"])
        col5G.button("Update 5G Report")


if btn_Kml:
    print("Test1:")
    if file_Kml is not None:
        try:
            file_Kml_Name = file_Kml.name
            # Load and clean the data
            print("Test2:",file_Kml_Name)
            df_Kml = pd.read_csv(file_Kml, engine='python', encoding='Windows-1252')
            df_Kml = lataftaf.clean_Sites_db(df_Kml,Is_Update_Nbrs,file_Kml_Name)
            df_Kml["coordinates"] = df_Kml["coordinates"].apply(lambda x: [float(x[0]), float(x[1])] if isinstance(x, (list, tuple)) else None)
        except Exception as e:
            st.error(f"Error reading KML CSV File: {e}")
    else:
        st.error("No KML CSV file selected")
if btn_Dmp:
    if file_Dmp is not None:
        try:
            file_Dmp_Name = file_Dmp.name
            lataftaf.upload_Dmp(file_Dmp)
        except Exception as e:
            st.error(f"Error reading Parameters Dump File: {e}")
    else:
        st.error("No Parameters Dump file selected")

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
