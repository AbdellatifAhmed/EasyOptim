import streamlit as st
import pandas as pd
import tools as lataftaf
import os
import base64
output_dir = os.path.join(os.getcwd(), 'OutputFiles')

st.set_page_config(
    page_title="EasyOptim - LNADJGNB Audit",
    layout="wide"  # Use the full width of the page
)
st.title("Audit LNADJGNB Tool")
st.write("Audit LNADJGNB Tool uses DB, LNADJGNB object to find the instances with Missing IPs, also provide instance that should be deleted as they are very far")
with st.container():
    col1, col2 = st.columns(2) 
    dB_file = col1.file_uploader("Select Engineer Parameters File [Must contain NodeB ID,Lat, Long]", type=["csv", "xlsx", "txt"])
    dmp_file = col2.file_uploader("Select Parameters Dump file [Must contain LnAdjgNB tabe]", type=["xlsb", "xlsx"])
    col3, col4, col5, col6 =st.columns(4)
    number_input = col3.number_input("Delete Distance [Km]", min_value=0, max_value=30, value=6)

if st.button("Submit"):
    if dB_file is not None:
        st.write("Engineer Parameters File uploaded successfully!")
        try:
            # Determine file type and read it
            if dB_file.name.endswith(".csv"):
                dB_data = pd.read_csv(dB_file, engine='python', encoding='Windows-1252')
            elif dB_file.name.endswith(".xlsx") or dB_file.name.endswith(".xls"):
                dB_data = pd.read_excel(dB_file)
            else:
                st.error("Unsupported file type!")
                dB_data = None
        except Exception as e:
            st.error(f"Error reading Engineer Parameters File: {e}")
    if dmp_file is not None:
        st.write("Parameters Dump File uploaded successfully!")
        try:
            # Determine file type and read it
            if dmp_file.name.endswith(".xlsb"):
                import pyxlsb  # Needed for .xlsb files
                dmp_data = pd.read_excel(dmp_file, sheet_name='LNADJGNB',header=1,engine='pyxlsb')
            elif dmp_file.name.endswith(".xlsx"):
                dmp_data = pd.read_excel(dmp_file)
            else:
                st.error("Unsupported file type!")
                dmp_data = None
        except Exception as e:
            st.error(f"Error reading Parameters Dump File: {e}")
    if dmp_data is not None and dB_data is not None:
        Lnadjgnb_audit_form = {'fileParamatersDB':dmp_data,
                               'fileSitesDB':dB_data,
                               'no_nbrDelDistance':number_input}
        
        if 'output_files' not in st.session_state:
            st.write("Data processing in progress .... ")
            Lnadjgnb_audit_output = lataftaf.audit_Lnadjgnb(Lnadjgnb_audit_form)
            st.write("Elapsed time :",Lnadjgnb_audit_output['duration'])
            st.session_state.output_files = Lnadjgnb_audit_output['output_Files']
        
        for file_path in st.session_state.output_files:
            with open(file_path, "rb") as f:
                file_data = f.read()
                b64_file_data = base64.b64encode(file_data).decode()
                download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(file_path)}">Click to download {os.path.basename(file_path)}</a>'
            st.markdown(download_link, unsafe_allow_html=True)
