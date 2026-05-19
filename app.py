import streamlit as st
from pptx import Presentation

st.title("ACME Slide Checker")

uploaded = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

if uploaded:
    try:
        prs = Presentation(uploaded)
        st.success("File loaded successfully!")
        st.write(f"Slides detected: {len(prs.slides)}")

        # --- RULES START HERE ---
        issues = []

        for i, slide in enumerate(prs.slides):
            title = slide.shapes.title.text if slide.shapes.title else ""
            if not title.strip():
                issues.append(f"Slide {i+1} has no title")

        if issues:
            st.error("Issues found:")
            for issue in issues:
                st.write(f"- {issue}")
        else:
            st.success("No issues found!")
        # --- RULES END HERE ---

    except Exception as e:
        st.error(f"Could not read PPTX: {e}")

