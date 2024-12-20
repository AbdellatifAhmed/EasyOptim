import streamlit as st
import pandas as pd
import tools as lataftaf
import pydeck as pdk

st.set_page_config(
    page_title="EasyOptim",
    layout="wide"  # Use the full width of the page
)
st.write("Hello Nokia Optimization Team") 
st.write("Welcome to EasyOptim Developed by Abdellatif Ahmed")

with st.container():
    dB_file = st.file_uploader("Sites DB File:", type=["csv"])

if st.button("Submit"):
    if dB_file is not None:
        try:
            # Load and clean the data
            dB_data = pd.read_csv(dB_file, engine='python', encoding='Windows-1252')
            dB_data = lataftaf.clean_Sites_db(dB_data)
            
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
            st.write(dB_data.head())
        except Exception as e:
            st.error(f"Error reading Engineer Parameters File: {e}")

