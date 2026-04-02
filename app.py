# app.py — Price File Builder (Streamlit)
import streamlit as st

from app_settings import render_settings_sidebar
from app_price import render_price_file
from app_specials import render_specials

st.set_page_config(layout="wide", page_title="Price File Builder")

st.warning(
    "⚠️ **This Streamlit app is still in testing.** "
    "Please use a test database — do not connect to a live production database.",
    icon=None,
)

render_settings_sidebar()

mode = st.radio(
    "Mode",
    ["Price File", "Specials File"],
    horizontal=True,
    key="app_mode",
)

st.divider()

if mode == "Price File":
    render_price_file()
else:
    render_specials()
