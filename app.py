import streamlit as st
import streamlit.components.v1 as components

# 1. Keep wide mode
st.set_page_config(layout="wide")

# 2. INJECT CSS TO REMOVE ALL BORDERS/PADDING
st.markdown("""
    <style>
        /* Removes padding from the main app container */
        .block-container {
            padding-top: 1rem;
            padding-bottom: 0rem;
            padding-left: 0rem;
            padding-right: 0rem;
            max-width: 100%;
        }
    </style>
""", unsafe_allow_html=True)

# 3. Read and render your HTML
with open("index.html", "r", encoding="utf-8") as f:
    html_data = f.read()

components.html(html_data, height=3000, scrolling=False)
