import streamlit as st
import pandas as pd
import os
import pyxlsb 
import time
import tools as lataftaf
import base64

output_dir = os.path.join(os.getcwd(), 'OutputFiles')

st.set_page_config(
    page_title="EasyOptim - Audit LNREL",
    layout="wide"  # Use the full width of the page
)
st.title("Audit LNREL Tool")
st.write("Audit LNREL Tool uses the DB, LNREL Objects, Report (RSLTE031)")

with st.container():
    col1, col2,col3 = st.columns(3) 
    dB_file = col1.file_uploader("Sites DB File:", type=["csv", "xlsx", "txt"])
    dmp_file = col2.file_uploader("Parameters Dump File:", type=["xlsb", "xlsx"])
    rel_prfrmnce_file = col3.file_uploader("Performance Report File:", type=["xlsb", "xlsx"])


if st.button("Submit"):
    timer = st.empty()
    timer.text("Total consumed time ")
    if dB_file is not None:
        start_time = time.time()
        status = st.empty()
        status.text("Engineer Parameters File uploading ...")
        try:
            # Determine file type and read it
            if dB_file.name.endswith(".csv"):
                dB_data = pd.read_csv(dB_file, engine='python', encoding='Windows-1252')
            elif dB_file.name.endswith(".xlsx") or dB_file.name.endswith(".xls"):
                dB_data = pd.read_excel(dB_file)
            else:
                st.error("Unsupported file type!")
                dB_data = None
            status.text("Engineer Parameters File uploading ... Done ✅")
        except Exception as e:
            st.error(f"Error reading Engineer Parameters File: {e}")

    if dmp_file is not None:
        status2 = st.empty()
        status2.text("Parameters Dump File uploading ...")
        try:
            dataframes_LnRel = []
            dataframes_LnCel = []
            # Determine file type and read it
            if dmp_file.name.endswith(".xlsb"):
                xls = pd.ExcelFile(dmp_file, engine='pyxlsb')
            elif dmp_file.name.endswith(".xlsx"):
                xls = pd.ExcelFile(dmp_file)
            else:
                st.error("Unsupported file type!")
                dmp_data = None
            sheet_names_LnRel = [sheet for sheet in xls.sheet_names if sheet == 'LNREL' or sheet.startswith('LNREL_')]
            sheet_names_LnCel = [sheet for sheet in xls.sheet_names if sheet == 'LNCEL' ]
            for sheet in sheet_names_LnRel:
                df = pd.read_excel(xls, sheet_name=sheet, header=1)  # Read the sheet into a dataframe
                dataframes_LnRel.append(df)
                end_time =time.time()
                duration = str(round((end_time - start_time),0))+" Seconds"
                timer.text("Total consumed time " + duration)
                # st.write("loaded LnRel rows:",len(df))
            for sheet in sheet_names_LnCel:
                df = pd.read_excel(xls, sheet_name=sheet, header=1)  # Read the sheet into a dataframe
                dataframes_LnCel.append(df)
            if dataframes_LnRel:
                dmp_data = pd.concat(dataframes_LnRel, ignore_index=True)
                # st.write(dmp_data.head())
            if dataframes_LnCel:
                dmp_data_LnCel = pd.concat(dataframes_LnCel, ignore_index=True)
                # st.write(dmp_data_LnCel.head())    
            else:
                st.error("No sheets matching 'LNREL' or 'LNREL_' were found.")
            status2.text("Parameters Dump File uploading ... Done ✅")
        except Exception as e:
            st.error(f"Error reading Parameters Dump File: {e}")
    
    if rel_prfrmnce_file is not None:
        status3 = st.empty()
        status3.text("Performance Report File uploading ...")
        try:
            # Determine file type and read it
            if rel_prfrmnce_file.name.endswith(".xlsb"):
                rel_prfrmnce_data = pd.read_excel(rel_prfrmnce_file, header=0,engine='pyxlsb')
            elif rel_prfrmnce_file.name.endswith(".xlsx"):
                rel_prfrmnce_data = pd.read_excel(rel_prfrmnce_file, sheet_name='RSLTE031 - Neighbor HO analysis')
            else:
                st.error("Unsupported file type!")
                rel_prfrmnce_data = None
            rel_prfrmnce_data = rel_prfrmnce_data.iloc[1:].reset_index(drop=True)
            # st.write(rel_prfrmnce_data.head())
            status3.text("Performance Report File uploading ... Done ✅")
            
        except Exception as e:
            st.error(f"Error reading Performance Report File: {e}")
    if dmp_data is not None and dB_data is not None and rel_prfrmnce_data is not None:
        Lnrel_audit_form = {'fileParamatersDB':dmp_data,'LnCel':dmp_data_LnCel,
                                'fileSitesDB':dB_data,
                                'Perform_Data':rel_prfrmnce_data}
        LNREL_Audit_output = lataftaf.audit_Lnrel(Lnrel_audit_form)
    
    with open(LNREL_Audit_output, "rb") as f:
        file_data = f.read()
        b64_file_data = base64.b64encode(file_data).decode()
        download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(LNREL_Audit_output)}">Click to download {os.path.basename(LNREL_Audit_output)}</a>'
    st.markdown(download_link, unsafe_allow_html=True)
    
    end_time =time.time()
    duration = str(round((end_time - start_time),0))+" Seconds"
    timer.text("Total consumed time " + duration)

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
