import pandas as pd
import ast
import streamlit as st
import tools as lataftaf
import os
import base64


if "selected_tech" not in st.session_state:
    st.session_state.selected_tech = '3G'


output_dir = os.path.join(os.getcwd(), 'OutputFiles')
overshooters = os.path.join(output_dir, 'Overshooting Sectors.xlsx')
st.set_page_config(
    page_title="EasyOptim - Find Overshooters",
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
        EasyOptim - Find Overshooters 
    </div>
    """,
    unsafe_allow_html=True,
)
tech_options = ['3G','4G']
default_value = None
with st.expander("Find Overshooters Tool",expanded=True):
    cont1 = st.container()
    with cont1:
        st.write("**Instructions:**")
        st.markdown("""
        - **Technology**: Limited to `3G`, `4G`.
        - **3G KPIs Report**: it must contain `WCEL name`, `WBTS ID`,`Period` or `Date` and PRACH counter must contain the text `opagation_delay` or `prach_delay_average`.
        - **3G Sites DB File**: Limited to `xlsx`, the columns `NodeB`,`Lat`,`Long`,`Bore` must exist.
        - **4G KPIs Report**: it must contain `LNCEL name`, `LNBTS name`,`Period` or `Date` and PRACH counter must contain the text `avg ue distance`.
        - **4G Sites DB File**: Limited to `xlsx`, the columns `eNodeB ID`,`Lat`,`Long` must exist.
         - **Function**: 
            1. Tool identify the closest sites within a `20Km` circle.
            2. Tool assume a coverage Arc of beamwidth `50degrees` and length equal to the Average propagation distance of the cell.
            3. Tool counts the sites that inside the coverage arc mentioned in point .
            4. The more sites inside the Coverage area, the more overshooting the cell introduce.
            5. If given Propagation report contains multiple values for the cell, only the Most recent value is considered.

        """)
        col1,col2,col3 = st.columns(3)
        drpbx_tech = col1.selectbox("Technology:", options=tech_options, key="selected_tech",index =None)
        btn_getOvershooter = col1.button("Find Overshooters")
        file_pd_report = col2.file_uploader("Propagation Delay Counters report:", type=["xlsx"])
        file_XML_Sites_DB = col3.file_uploader("Sites Engineering Parameters Database:", type=["xlsx"])
        if btn_getOvershooter:
            if file_pd_report:
                if file_XML_Sites_DB:
                    out_p = lataftaf.get_overshooters(drpbx_tech,file_pd_report,file_XML_Sites_DB)
                    st.write("Execution Done Successfully in: ", out_p)
                    with open(overshooters, "rb") as f:
                        file_data = f.read()
                        b64_file_data = base64.b64encode(file_data).decode()
                        download_link = f'<a href="data:application/octet-stream;base64,{b64_file_data}" download="{os.path.basename(overshooters)}">Click to download possible overshooter Cells {os.path.basename(overshooters)}</a>'
                    st.markdown(download_link, unsafe_allow_html=True)
                else:
                    st.error("Missing Sites DB")
            else:
                st.error("Missing Propagation delay KPIs report!")
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
