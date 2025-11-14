# Use view.officeapps.live.com (EASIEST + works in Streamlit)

import streamlit as st
import base64
import os

st.title("DOCX Viewer using Microsoft Office Online")

uploaded = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded:
    # Save file to a temporary dir that Streamlit can serve
    temp_path = os.path.join("temp", uploaded.name)
    os.makedirs("temp", exist_ok=True)
    
    with open(temp_path, "wb") as f:
        f.write(uploaded.getbuffer())

    # Generate a public URL for Streamlit static files
    file_url = st.experimental_get_query_params()
    file_url = f"http://localhost:8501/temp/{uploaded.name}"

    office_url = (
        "https://view.officeapps.live.com/op/embed.aspx?src=" + file_url
    )

    st.markdown(
        f"""
        <iframe src="{office_url}" width="100%" height="650px" frameborder="0"></iframe>
        """,
        unsafe_allow_html=True,
    )
