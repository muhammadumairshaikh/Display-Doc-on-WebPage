import streamlit as st
from pdf2image import convert_from_path
import win32com.client as win32
import pythoncom
import os

st.set_page_config(page_title="DOC VIEWER")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, "v2 Backend.docx")

POPLER_PATH = r"C:\Users\Imperium Dynamics\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"


def convert_docx_to_pdf(docx_path):
    pythoncom.CoInitialize()
    docx_path = os.path.abspath(docx_path)
    pdf_path = docx_path.replace(".docx", ".pdf")

    word = win32.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

    return pdf_path


def display_pdf(pdf_path):
    pdf_path = os.path.abspath(pdf_path)

    pages = convert_from_path(
        pdf_path,
        poppler_path=POPLER_PATH   
    )

    for i, page in enumerate(pages):
        st.markdown(f"### Page {i+1}")
        st.image(page, use_container_width=True)
        # st.markdown("---")


if FILE_PATH.lower().endswith(".pdf"):
    display_pdf(FILE_PATH)

elif FILE_PATH.lower().endswith(".docx"):
    converted_pdf = convert_docx_to_pdf(FILE_PATH)
    display_pdf(converted_pdf)

else:
    st.error("Unsupported file type.")
