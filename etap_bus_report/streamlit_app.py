"""
streamlit_app.py
----------------
Web front end (Streamlit Community Cloud or any headless server).

It uses exactly the same pipeline as the desktop app - `pdf_parser`,
`voltage_checker`, `word_writer` - so both front ends always produce identical
documents.  Nothing here imports PySide6, which cannot run on a headless
server.

The whole page lives in :func:`render`, which must be called on *every* script
run: Streamlit re-executes the script on each interaction, and an already
imported module body would not run again.

    streamlit run streamlit_app.py      # directly
    streamlit run main.py               # via the shared entry point
"""

from __future__ import annotations

import os
import sys
import tempfile
import time

import streamlit as st

# Make the sibling modules importable when Streamlit runs this file directly.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pdf_parser          # noqa: E402
import voltage_checker     # noqa: E402
import word_writer         # noqa: E402
from utils import APP_NAME, APP_VERSION, ParserError  # noqa: E402

OUTPUT_NAME = "Bus Report.docx"

STYLE = """
<style>
  .block-container { max-width: 780px; padding-top: 2.5rem; }
</style>
"""


@st.cache_data(show_spinner=False)
def extract_buses(pdf_bytes: bytes, include_pseudo: bool):
    """Parse an uploaded report (cached, so changing the limits won't re-parse)."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as handle:
        handle.write(pdf_bytes)
        temp_path = handle.name
    try:
        return pdf_parser.load_report(temp_path, include_pseudo_buses=include_pseudo)
    finally:
        os.unlink(temp_path)


def render() -> None:
    """Draw the whole page. Called once per Streamlit script run."""
    st.set_page_config(page_title=APP_NAME, page_icon="⚡", layout="centered")
    st.markdown(STYLE, unsafe_allow_html=True)

    st.title("ETAP Bus Load Flow Report")
    st.caption(
        "Extracts the bus results from an ETAP Load Flow Analysis report and fills "
        f"the standard Bus table.  v{APP_VERSION}"
    )

    # -- Input 1: the ETAP report ------------------------------------------ #
    uploaded = st.file_uploader(
        "**1.  Upload ETAP Load Flow Analysis Report (.pdf)**",
        type=["pdf"],
        help="The Word template is built into the app - you never upload it.",
    )

    # -- Input 2: the acceptable voltage band ------------------------------ #
    limits_text = st.text_input(
        "**2.  Acceptable Voltage Limits**",
        value=st.session_state.get("limits_text", "95-106"),
        key="limits_text",
        placeholder="95-106",
        help="Accepted formats: 95-106  •  95 - 105 %  •  90-110  •  95 to 106",
    )

    with st.expander("Options"):
        include_pseudo = st.checkbox(
            "Include ETAP internal nodes (IDs containing '~')",
            value=False,
            key="include_pseudo",
            help="Cable and VFD terminal nodes such as Cable40~ or "
                 "S002-TR-CMP1_VFD~2. These are not real switchboards and are "
                 "excluded by default.",
        )

    left, right = st.columns(2)
    generate = left.button("Generate Report", type="primary", use_container_width=True)
    reset = right.button("Reset", use_container_width=True)

    if reset:
        st.session_state.pop("result", None)
        st.session_state.pop("limits_text", None)
        st.rerun()

    if generate:
        _generate(uploaded, limits_text, include_pseudo)

    _show_result()


def _generate(uploaded, limits_text: str, include_pseudo: bool) -> None:
    """Run the pipeline and store the outcome in the session state."""
    if uploaded is None:
        st.warning("Please upload an ETAP Load Flow Analysis report (.pdf) first.")
        return

    try:
        limits = voltage_checker.parse_limits(limits_text)
    except voltage_checker.LimitError as exc:
        st.error(str(exc))
        return

    started = time.perf_counter()
    try:
        with st.spinner("Reading the ETAP report ..."):
            buses = extract_buses(uploaded.getvalue(), include_pseudo)
        voltage_checker.apply_limits(buses, limits)
        document_bytes = word_writer.build_document_bytes(buses)
    except ParserError as exc:
        st.error(str(exc))
        return
    except Exception:  # pragma: no cover - unexpected failure
        st.error(pdf_parser.MSG_UNREADABLE)
        return

    st.session_state["result"] = {
        "bytes": document_bytes,
        "rows": [bus.as_row() for bus in buses],
        "not_ok": sum(1 for bus in buses if bus.remarks == "NOT ACCEPTABLE"),
        "limits": str(limits),
        "seconds": time.perf_counter() - started,
    }


def _show_result() -> None:
    result = st.session_state.get("result")
    if not result:
        return

    st.success("Report Generated Successfully")

    left, middle, right = st.columns(3)
    left.metric("Buses", len(result["rows"]))
    middle.metric("Outside limits", result["not_ok"])
    right.metric("Limits applied", result["limits"])

    st.download_button(
        "Download Bus Report.docx",
        data=result["bytes"],
        file_name=OUTPUT_NAME,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
    st.caption(f"Processed in {result['seconds']:.1f} s")

    with st.expander("Preview the extracted table"):
        st.dataframe(
            result["rows"],
            column_config={
                "0": "Bus ID",
                "1": "Nominal (kV, A)",
                "2": "Voltage %",
                "3": "kW Loading",
                "4": "Amp Loading",
                "5": "Remarks",
            },
            hide_index=True,
            use_container_width=True,
        )


if __name__ == "__main__":
    render()
