import streamlit as st
from pptx import Presentation

st.title("ACME Slide Checker")

uploaded = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

if uploaded:
    try:
        prs = Presentation(uploaded)
        st.success("File loaded successfully!")
        st.write(f"Slides detected: {len(prs.slides)}")
    except Exception as e:
        st.error(f"Could not read PPTX: {e}")
