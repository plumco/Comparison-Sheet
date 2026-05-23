import streamlit as st
import streamlit.components.v1 as components

# Open your HTML file and read it
with open("index.html", "r", encoding="utf-8") as f:
    html_data = f.read()

# Render the HTML inside your Streamlit app
components.html(html_data, height=800, scrolling=True)
