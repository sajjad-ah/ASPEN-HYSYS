import streamlit as st

st.set_page_config(
    page_title="Emerson Thermowell Template Generator",
    layout="wide",
)

from app import main

main()