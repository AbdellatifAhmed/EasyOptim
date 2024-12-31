import os
import pandas as pd
import ast
import streamlit as st
import tools as lataftaf
import base64
# Setup
output_dir = os.path.join(os.getcwd(), 'OutputFiles')
xml_objects = os.path.join(output_dir, 'XML Objects.xlsx')
created_xml_link = os.path.join(output_dir, 'OutputXML.xml')
st.set_page_config(
    page_title="EasyOptim - Make XML",
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
        Hello Nokia Optimization Team - EasyOptim - Create XML Tool 
    </div>
    """,
    unsafe_allow_html=True,
)

# Load Data
df_xml_objects = pd.read_excel(xml_objects)
df_xml_objects["Mandatory_ID"] = df_xml_objects["Mandatory_ID"].apply(ast.literal_eval)
df_xml_objects["Parameters_List"] = df_xml_objects["Parameters_List"].apply(ast.literal_eval)

# Session State Initialization

if "selected_object" not in st.session_state:
    st.session_state.selected_object = 'ACBPR'
    filtered_row = df_xml_objects[df_xml_objects["Object"] == 'ACBPR'].iloc[0]
    st.session_state.mandatory_columns = filtered_row["Mandatory_ID"]

if "mandatory_columns" not in st.session_state:
    st.session_state.mandatory_columns = ''
    # st.session_state.parameters_list = filtered_row["Parameters_List"]

if "table_data" not in st.session_state:
    st.session_state.table_data = st.session_state.mandatory_columns 


def on_object_change():
    selected_object = st.session_state["object_dropdown"]
    st.session_state.selected_object = selected_object
    filtered_row = df_xml_objects[df_xml_objects["Object"] == selected_object].iloc[0]
    st.session_state.mandatory_columns = filtered_row["Mandatory_ID"]




# Main Tool 
with st.expander("Create XML Tool", expanded=True):
    container1 = st.container()
    with container1:
        col1, col2, col3 = st.columns(3)
        drp_object = col1.selectbox("Select an Object:", options=df_xml_objects["Object"], key="object_dropdown",on_change=on_object_change)
        col1.write(f"You are selecting selecting: {st.session_state.selected_object}")
        col1.write(f"Mandatory Feilds are: {st.session_state.mandatory_columns}")
        file_csv = col2.file_uploader("CSV changes File:", type=["csv"])
        rad_action = col3.radio("What is the Operation?",["update","delete","create"])
        btn_makeXML = st.button("Create XML", key="makeXML_button")
    if btn_makeXML:
        out = lataftaf.valide_make_XML(st.session_state.selected_object,file_csv,rad_action)
        st.write(out)
        with open(created_xml_link, "rb") as f:
            file_data = f.read()
            b64_file_data = base64.b64encode(file_data).decode()
            download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(created_xml_link)}">Click to download {os.path.basename(created_xml_link)}</a>'
        st.markdown(download_link, unsafe_allow_html=True)


    

if "expanded" not in st.session_state:
    st.session_state.expanded = False

with st.expander("Define XML Objects",expanded=st.session_state.expanded):
    file_Dmp = st.file_uploader("Select xlsb Dump file:", type=["xlsb"])
    btn_buildXML = st.button("Go Build XML Objects")

if btn_buildXML:
    st.write((lataftaf.build_XML_Object(file_Dmp)).head())
    