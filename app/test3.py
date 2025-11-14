import streamlit as st
import base64

st.set_page_config(page_title="DOCX Viewer", layout="wide")
st.title("DOCX Viewer (docx-preview 0.3.7)")

uploaded = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded:
    file_bytes = uploaded.read()
    b64_doc = base64.b64encode(file_bytes).decode("utf-8")

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <script src="https://unpkg.com/docx-preview@0.3.7/dist/docx-preview.min.js"></script>
        <style>
            #docx-container {{
                width: 100%;
                height: 900px;
                border: 1px solid #ccc;
                overflow: auto;
                padding: 20px;
            }}
        </style>
    </head>

    <body>
        <div id="docx-container">Loading document…</div>

        <script>
            const base64Data = "{b64_doc}";

            // Convert base64 → Blob
            const binary = atob(base64Data);
            const len = binary.length;
            const buffer = new Uint8Array(len);
            for (let i = 0; i < len; i++) {{
                buffer[i] = binary.charCodeAt(i);
            }}

            const blob = new Blob([buffer], {{
                type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }});

            const container = document.getElementById("docx-container");

            // Render DOCX
            window.docx.renderAsync(blob, container)
                .then(() => console.log("DOCX rendered"))
                .catch(err => {{
                    container.innerHTML = "<p style='color:red;'>Failed to render document.</p>";
                    console.error(err);
                }});
        </script>
    </body>
    </html>
    """

    st.components.v1.html(html_code, height=950, scrolling=True)
