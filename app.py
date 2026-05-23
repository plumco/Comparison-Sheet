import streamlit as st
import streamlit.components.v1 as components

# THIS LINE FIXES THE WIDTH (It must be the first Streamlit command)
st.set_page_config(layout="wide")

# Read your HTML file
with open("index.html", "r", encoding="utf-8") as f:
    html_data = f.read()

# Render the HTML (I increased the height here so it fits better)
components.html(html_data, height=1300, scrolling=True)
