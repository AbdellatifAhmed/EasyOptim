import streamlit as st
import pandas as pd
import tools as lataftaf
import pydeck as pdk
import base64
import os

lataftaf.get_log()
st.set_page_config(
    page_title="EasyOptim",
    layout="wide"  # Use the full width of the page
)
st.markdown(
    """
    <style>
    .header {
        background-color: #f8f9fa;
        padding: 20px;
        text-align: left;
        font-size: 24px;
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
st.write("Please select any desired tool, also updated DB can be uploaded")
recent_dB = lataftaf.get_log() 

with st.container():
    status = st.empty()
    try:
        with open(recent_dB[1], "rb") as f:
            file_data = f.read()
        b64_file_data = base64.b64encode(file_data).decode()  # Encode log content to base64
        href = f'<a href="data:file/txt;base64,{b64_file_data}" download="{os.path.basename(recent_dB[1])}">{recent_dB[0]}</a>'
        sites_dB_comment = f"The Sites DB file [ {href} ] already exists. Use upload if a recent one needs to be used."
        status.markdown(sites_dB_comment, unsafe_allow_html=True)
        # st.markdown(sites_dB_comment, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error reading intial File location: {e}")
    dB_file = st.file_uploader("Sites DB File:", type=["csv"])
    Is_Update_Nbrs = st.checkbox('Check to update Nbrs')

if st.button("Submit"):
    if dB_file is not None:
        try:
            dB_file_name = dB_file.name
            dB_file_date = dB_file_name [:-4][3:]
            # Load and clean the data
            dB_data = pd.read_csv(dB_file, engine='python', encoding='Windows-1252')
            dB_data = lataftaf.clean_Sites_db(dB_data,Is_Update_Nbrs,dB_file_name)
            href = f'<a href="data:file/txt;base64,{b64_file_data}" download="{os.path.basename(recent_dB[1])}">{dB_file_name}</a>'
            sites_dB_comment = f"The Sites DB file [ {href} ] already exists. Use upload if a recent one needs to be used."
            status.markdown(sites_dB_comment, unsafe_allow_html=True)
            # Define the polygon layer
            Sectors_layer = pdk.Layer(
                "PolygonLayer",
                dB_data,
                get_polygon="polygon",  # Use the 'polygon' column
                get_fill_color="[200, 30, 90, 160]",  # Purple fill color with transparency
                get_line_color=[255, 255, 255],  # White border
                pickable=True,  # Enable click interactions
                extruded=False,  # Flat shapes
                
            )
            
            # Define the text layer for NodeB IDs
            text_layer = pdk.Layer(
                "TextLayer",
                dB_data,
                get_position="coordinates",  # Position for the text
                get_text="NodeB Id",  # Column to display as text
                get_color=[0, 0, 0, 0],  # Black text with slight transparency
                get_size=10,  # Font size
                get_angle=0,
                get_alignment_baseline="'top'",  # Position text above the sector
                pickable=False,  # Text should not be clickable
                
            )
            
            # Define the map view state
            view_state = pdk.ViewState(
                latitude=31.95938,  # Center based on average latitude
                longitude=35.862283,  # Center based on average longitude
                zoom=15,
                pitch=0,
            )
            
            # Render the map
            st.pydeck_chart(
                pdk.Deck(
                    layers=[Sectors_layer,text_layer], 
                    initial_view_state=view_state,
                    tooltip={"Vendor": "{Vendor}\n{BSCName}"},
                    map_style="mapbox://styles/mapbox/streets-v11"
                )
            )
        except Exception as e:
            st.error(f"Error reading Engineer Parameters File: {e}")
    else:
        st.error("No Sites DB file slected")

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
