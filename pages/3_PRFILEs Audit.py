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


