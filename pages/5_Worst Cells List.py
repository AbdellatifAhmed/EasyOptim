import pandas as pd
import ast
import streamlit as st
import tools as lataftaf
import os
import base64

output_dir = os.path.join(os.getcwd(), 'OutputFiles')
wcel_criteria = os.path.join(output_dir, 'WCL_Criteria.xlsx')
st.set_page_config(
    page_title="EasyOptim - Worst Cell List",
    layout="wide"
)
def update_Criteria(df_criteria):
    with pd.ExcelWriter(wcel_criteria, engine='openpyxl') as writer:
            df_criteria.to_excel(writer, sheet_name='WCL_Criteria', index=False)
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
        EasyOptim - Worst Cell List 
    </div>
    """,
    unsafe_allow_html=True,
)

intial_table =pd.DataFrame(
    {
        "Technology": ["2G"],
        "KPI_Name": [""],
        "Indicator1": [""],
        "Logical_Condition1": ["<"],
        "Threshold1": [0],
        "Indicator2": [""],
        "Logical_Condition2": ["<"],
        "Threshold2": [0],
        # "Indicator3": [""],
        # "Logical_Condition3": ["<"],
        # "Threshold3": [0],
    }
)
intial_table = intial_table[:-1]
tech_options = ['2G','3G','4G','5G']
default_value = '4G'
if "table_criteria" not in st.session_state:
    st.session_state.table_criteria = intial_table

# Table interface inside an expander
with st.expander("Specify or Add KPIs/Thresholds for Worst Cells Criteria", expanded=True):
    st.write("**Instructions:**")
    st.markdown("""
    - **Technology**: Limited to `2G`, `3G`, `4G`, `5G`.
    - **Logical_Condition1, Logical_Condition2**: Limited to `<`, `>`, `=`, `<=`, `>=`.
    - **Indicator1, Indicator1**: must be same name in KPIs report related to the technology.
    - **Input KPIs reports must be in CSV formate, [output report from Nokia Netact are with SemiColumn delimiter].
    """)
    tech_col,kpi_col,thrshld_col = st.columns(3)
    drpbx_tech = tech_col.selectbox("Technology:", options=tech_options, key="drpbx_tech",index =tech_options.index(default_value))
    kpi_alias_input = kpi_col.text_input("KPI Name/Alias:", "")
    threshold_days = thrshld_col.number_input("Number Of Days:", min_value=0, max_value=7, value=3)
    
    col1,col2,col3 = st.columns(3)
    kpi1_input = col1.text_input("KPI/Indicator 1", "")
    drpbx_log_cond_1 = col1.selectbox("Logical Condition:", options=['<','>','=','<=','>='], key="drpbx_log1")
    val1_input = col1.number_input("KPI1 Threshold", min_value=-1000.0, max_value=100000.0, value=0.0,step=0.0001)
    
    kpi2_input = col2.text_input("KPI/Indicator 2", "")
    drpbx_log_cond_2 = col2.selectbox("Logical Condition:", options=['<','>','=','<=','>='], key="drpbx_log2")
    val2_input = col2.number_input("KPI2 Threshold", min_value=-1000.0, max_value=100000.0, value=0.0,step=0.0001)
    
    
    file_criteria = col3.file_uploader("Select Existing Criteria Excel File:", type=["xlsx","csv"])
    col3_1,col3_2 =col3.columns(2)
    btn_uploadCriteria = col3_1.button("Upload Criteria", key="btn_uploadCriteria")
    chk_overrideCriteria = col3_2.checkbox('Override All Existing')

    btn_addKPI = st.button("Add to Criteria", key="btn_addKPI")
    criteria = st.empty()
    criteria.table(st.session_state.table_criteria)

    # btn_downloadCriteria = st.button("Download Criteria", key="btn_downloadCriteria")

    if btn_addKPI:
        if kpi_alias_input =="":
            st.error("You must Specify the KPI Name before adding it to craiteria")
        else:
            if kpi1_input =="" and kpi2_input == "" :
                st.error("You must Specify at least 1 Indicator to be used for chekcing the KPI "+ kpi_alias_input +" from Reports")
            else:
                new_row = {
                "Technology": drpbx_tech,
                "KPI_Name": kpi_alias_input,
                "Indicator1": kpi1_input,
                "Logical_Condition1": drpbx_log_cond_1,
                "Threshold1": val1_input,
                "Indicator2": kpi2_input,
                "Logical_Condition2": drpbx_log_cond_2,
                "Threshold2": val2_input,
            }
                st.session_state.table_criteria = pd.concat([st.session_state.table_criteria, pd.DataFrame([new_row])], ignore_index=True)
                criteria.table(st.session_state.table_criteria)
                update_Criteria(st.session_state.table_criteria)
                csv_data = st.session_state.table_criteria.to_csv(index=False)
                st.download_button(
                    label="Click to Download Criteria",
                    data=csv_data,
                    file_name="criteria.csv",
                    mime="text/csv",
                )

    if btn_uploadCriteria:
        if file_criteria is not None:
            try:
                if file_criteria.name.endswith(".csv"):
                    df_ext_Crit = pd.read_csv(file_criteria, engine='python', encoding='Windows-1252')
                elif file_criteria.name.endswith(".xlsx"):
                    df_ext_Crit = pd.read_excel(file_criteria)
                if chk_overrideCriteria:
                    st.session_state.table_criteria = df_ext_Crit
                else:
                    st.session_state.table_criteria = pd.concat([st.session_state.table_criteria, df_ext_Crit], ignore_index=True)
                
                criteria.table(st.session_state.table_criteria)
                update_Criteria(st.session_state.table_criteria)
                csv_data = st.session_state.table_criteria.to_csv(index=False)
                st.download_button(
                    label="Click to Download Criteria",
                    data=csv_data,
                    file_name="criteria.csv",
                    mime="text/csv",
                )
            except Exception as e:
                st.error("Problem in Selected Criteria File")
                st.error(e)
        else:
            st.error("No Criteria file selected to be updloaded")
          
with st.expander("Select the KPIs Reports for each Technology",expanded=True):
    col2G,col3G,col4G,col5G = st.columns(4)
    files_2G = col2G.file_uploader("Select 2G KPIs Reports:", type=["csv"],accept_multiple_files=True)
    files_3G = col3G.file_uploader("Select 3G KPIs Reports:", type=["csv"],accept_multiple_files=True) 
    files_4G = col4G.file_uploader("Select 4G KPIs Reports:", type=["csv"],accept_multiple_files=True) 
    files_5G = col5G.file_uploader("Select 5G KPIs Reports:", type=["csv"],accept_multiple_files=True)

    btn_initiate_Wcl = st.button("Start", key="btn_initiate_Wcl") 
    
    if btn_initiate_Wcl:
        if not files_2G and not files_3G and not files_4G and not files_5G:
            st.error("You did not provide any KPIs reports to check Worst Cells. Please upload reports and then press Start.")
        else:
            if files_2G:
                wcl_2G = lataftaf.get_wcl(st.session_state.table_criteria,files_2G,'2G',threshold_days)
                with open(wcl_2G, "rb") as f:
                    file_data = f.read()
                    b64_file_data = base64.b64encode(file_data).decode()
                    download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(wcl_2G)}">Click to download 2G WCL File {os.path.basename(wcl_2G)}</a>'
                st.markdown(download_link, unsafe_allow_html=True)
            
            if files_3G:
                wcl_3G = lataftaf.get_wcl(st.session_state.table_criteria,files_3G,'3G',threshold_days)
                with open(wcl_3G, "rb") as f:
                    file_data = f.read()
                    b64_file_data = base64.b64encode(file_data).decode()
                    download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(wcl_3G)}">Click to download 3G WCL File {os.path.basename(wcl_3G)}</a>'
                st.markdown(download_link, unsafe_allow_html=True)

            if files_4G:
                wcl_4G = lataftaf.get_wcl(st.session_state.table_criteria,files_4G,'4G',threshold_days)
                with open(wcl_4G, "rb") as f:
                    file_data = f.read()
                    b64_file_data = base64.b64encode(file_data).decode()
                    download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(wcl_4G)}">Click to download 4G WCL File {os.path.basename(wcl_4G)}</a>'
                st.markdown(download_link, unsafe_allow_html=True)

            if files_5G:
                wcl_5G = lataftaf.get_wcl(st.session_state.table_criteria,files_5G,'5G',threshold_days)
                with open(wcl_5G, "rb") as f:
                    file_data = f.read()
                    b64_file_data = base64.b64encode(file_data).decode()
                    download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(wcl_5G)}">Click to download 5G WCL File {os.path.basename(wcl_5G)}</a>'
                st.markdown(download_link, unsafe_allow_html=True)
        

