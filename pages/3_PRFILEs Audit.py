import streamlit as st
import pandas as pd
import tools as lataftaf
import pydeck as pdk
from pydeck.types import String

st.set_page_config(
    page_title="EasyOptim - PRFILE Audit",
    layout="wide"  # Use the full width of the page
)
st.title("PRFILE Audit Tool")
st.write("tool take input PRFILEs and share organized output")


TEXT_LAYER_DATA = "https://raw.githubusercontent.com/visgl/deck.gl-data/master/website/bart-stations.json"  # noqa
df = pd.read_json(TEXT_LAYER_DATA)
st.write(df.head())

if st.button("Submit"):
    # Define a layer to display on a map
    layer = pdk.Layer(
        "TextLayer",
        df,
        pickable=True,
        get_position="coordinates",
        get_text="name",
        get_size=16,
        get_color=[0, 0, 0],
        get_angle=0,
        # Note that string constants in pydeck are explicitly passed as strings
        # This distinguishes them from columns in a data set
        get_text_anchor=String("middle"),
        get_alignment_baseline=String("center"),
    )

    # Set the viewport location
    view_state = pdk.ViewState(latitude=37.7749295, longitude=-122.4194155, zoom=10, bearing=0, pitch=45)

    # Render
    r = pdk.Deck(
        layers=[layer],
        initial_view_state=view_state,
        tooltip={"text": "{name}\n{address}"},
        map_style=pdk.map_styles.ROAD,
    )
    r.to_html("text_layer.html")
    st.pydeck_chart(
                        pdk.Deck(
                            layers=[layer], 
                            initial_view_state=view_state,
                            map_style="mapbox://styles/mapbox/streets-v11"
                        )
                    )