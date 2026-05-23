import streamlit as st
import streamlit.components.v1 as components

# 1. THIS MAKES IT WIDE (It MUST be the very first Streamlit command)
st.set_page_config(layout="wide")

# Read your HTML file
with open("index.html", "r", encoding="utf-8") as f:
    html_data = f.read()

# 2. MASSIVELY INCREASE HEIGHT & DISABLE INNER SCROLLBAR
# I set height to 3000 to ensure it fits all your rows.
components.html(html_data, height=3000, scrolling=False)
