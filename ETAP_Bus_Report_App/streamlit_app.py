"""
streamlit_app.py
Web version of the app, for deployment on Streamlit Community Cloud (or any
Streamlit host). This is a SEPARATE entry point from ui.py (the PySide6
desktop GUI) and main.py (the argparse CLI) - Streamlit Cloud has no
display, so the Qt app cannot run there, and running main.py directly under
`streamlit run` triggers its argparse CLI instead of a web UI (that's the
"the following arguments are required: --pdf, --limits" error).

Deploy this file instead:
    streamlit run streamlit_app.py

All the actual parsing/writing logic still lives in pdf_parser.py /
word_writer.py / voltage_checker.py / utils.py - this file only adds a
browser-based front end calling the same generate_report() used by ui.py
and main.py.
"""
from __future__ import annotations

import io
import os
import tempfile

import streamlit as st

from main import generate_report, ReportGenerationError, DEFAULT_TEMPLATE

st.set_page_config(page_title="ETAP Bus Load Flow Report Generator", page_icon="⚡")

st.title("⚡ ETAP Bus Load Flow Report Generator")
st.caption(
    "Upload an ETAP Load Flow Analysis Report (PDF), enter acceptable "
    "voltage limits, and download a populated Bus Report Word document."
)

if not os.path.exists(DEFAULT_TEMPLATE):
    st.error(
        f"Bundled Word template not found at `{DEFAULT_TEMPLATE}`. "
        "Make sure assets/Bus_Template.docx is included in the deployment."
    )
    st.stop()

# --- Input 1: PDF upload --------------------------------------------------- #
uploaded_pdf = st.file_uploader(
    "Upload ETAP Load Flow Analysis Report (.pdf)", type=["pdf"]
)

# --- Input 2: voltage limits ------------------------------------------------ #
limits_text = st.text_input(
    "Acceptable Voltage Limits",
    placeholder="e.g. 95-106  or  95 - 105 %  or  90-110",
)

col1, col2 = st.columns([1, 1])
generate_clicked = col1.button("Generate Report", type="primary", use_container_width=True)
reset_clicked = col2.button("Reset", use_container_width=True)

if reset_clicked:
    st.rerun()

if generate_clicked:
    if uploaded_pdf is None:
        st.warning("Please upload an ETAP PDF report first.")
    elif not limits_text.strip():
        st.warning("Please enter voltage limits in the format Lower-Upper (Example: 95-106).")
    else:
        # generate_report() takes file paths, so the uploaded bytes need to
        # be written to a temp file first (this also gives pdf_parser a
        # real path to open with PyMuPDF).
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, uploaded_pdf.name or "report.pdf")
            with open(pdf_path, "wb") as f:
                f.write(uploaded_pdf.getbuffer())
            out_path = os.path.join(tmpdir, "Bus Report.docx")

            with st.spinner("Processing… extracting bus data and populating the report."):
                try:
                    generate_report(pdf_path, limits_text, out_path)
                except ReportGenerationError as exc:
                    st.error(str(exc))
                else:
                    st.success("Report Generated Successfully")
                    with open(out_path, "rb") as f:
                        report_bytes = f.read()
                    st.download_button(
                        label="Download Bus Report.docx",
                        data=report_bytes,
                        file_name="Bus Report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
