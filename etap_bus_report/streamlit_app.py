"""
streamlit_app.py
----------------
Web front end (Streamlit Community Cloud or any headless server).

Report types come from the :mod:`reports` registry.  The Load Flow module has
a configurable interface (select columns, set alert limits, choose units and
output format, preview, then export); the Short Circuit module keeps its
existing template-driven workflow untouched.

Nothing here imports PySide6, which cannot run on a headless server.

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

import loadflow_config as lf_config    # noqa: E402
import loadflow_fields as lf_fields    # noqa: E402
import reports                          # noqa: E402
import shortcircuit_config as sc_config  # noqa: E402
import shortcircuit_fields as sc_fields  # noqa: E402
import shortcircuit_processor as sc_processor  # noqa: E402
from utils import APP_NAME, APP_VERSION, ParserError  # noqa: E402

STYLE = """
<style>
  .block-container { max-width: 980px; padding-top: 2.5rem; }
  div[data-testid="stMetricValue"] { font-size: 1.3rem; }
</style>
"""

WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# --------------------------------------------------------------------------- #
# Cached work - the PDF is read once per uploaded file
# --------------------------------------------------------------------------- #


def _to_temp(pdf_bytes: bytes) -> str:
    handle = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    handle.write(pdf_bytes)
    handle.close()
    return handle.name


@st.cache_data(show_spinner=False, max_entries=3)
def parse_load_flow(pdf_bytes: bytes, include_pseudo_buses: bool):
    """Parse a Load Flow PDF once; cached so changing a checkbox is instant."""
    path = _to_temp(pdf_bytes)
    try:
        return reports.load_flow_records(path, include_pseudo_buses)
    finally:
        os.unlink(path)


@st.cache_data(show_spinner=False, max_entries=3)
def parse_short_circuit(pdf_bytes: bytes, with_detail: bool):
    """
    Parse a Short Circuit PDF once; cached so changing a checkbox is instant.

    ``with_detail`` adds Cfactor and X/R, which needs a whole-document scan and
    is therefore keyed separately rather than done every time.
    """
    path = _to_temp(pdf_bytes)
    try:
        info = reports.short_circuit_study_info(path)
        records = reports.short_circuit_records(path, include_nodes=True)
        if with_detail:
            reports.short_circuit_detail(path, records)
        return records, info
    finally:
        os.unlink(path)


# --------------------------------------------------------------------------- #
# Page
# --------------------------------------------------------------------------- #


def render() -> None:
    """Draw the whole page. Called once per Streamlit script run."""
    st.set_page_config(page_title=APP_NAME, page_icon="⚡", layout="centered")
    st.markdown(STYLE, unsafe_allow_html=True)

    st.title("ETAP Report Filler")
    st.caption(f"Builds the study tables from an ETAP report.  v{APP_VERSION}")

    label = st.radio("**Report Type**", reports.labels(), horizontal=True, key="report_label")
    report = reports.by_label(label)
    st.caption(report.description)

    if report.key == reports.LOAD_FLOW:
        _render_load_flow(report)
    else:
        _render_short_circuit(report)


# --------------------------------------------------------------------------- #
# Load Flow - configurable interface
# --------------------------------------------------------------------------- #


def _render_load_flow(report) -> None:
    uploaded = st.file_uploader(
        f"**1.  {report.upload_label}**", type=["pdf"], key="lf_upload",
    )

    include_pseudo = st.checkbox(
        "Include ETAP internal nodes (IDs containing '~')",
        value=False, key="lf_pseudo",
        help="Cable and VFD terminal nodes such as Cable40~. Not real switchboards.",
    )

    if uploaded is None:
        st.info("Upload an ETAP Load Flow Analysis report to configure and preview the table.")
        return

    try:
        with st.spinner("Reading the ETAP report ..."):
            records = parse_load_flow(uploaded.getvalue(), include_pseudo)
    except ParserError as exc:
        st.error(str(exc))
        return
    except Exception:
        st.error("Unable to read the ETAP report.")
        return

    st.success(f"{len(records)} buses extracted from {uploaded.name}")

    config = _load_flow_configuration()

    # -- preview + export ---------------------------------------------------- #
    try:
        table = reports.load_flow_table(records, config)
    except lf_config.ConfigError as exc:
        st.warning(str(exc))
        return

    st.markdown("**5.  Preview**")
    st.caption(
        f"{table.count} rows x {len(table.headers)} columns - "
        "exactly what the Word and Excel reports will contain."
    )
    st.dataframe(
        [dict(zip(table.headers, row)) for row in table.rows],
        hide_index=True, use_container_width=True, height=360,
    )

    for warning in table.warnings:
        st.warning(warning)

    left, right = st.columns([1, 2])
    if left.button("Generate Report", type="primary", use_container_width=True):
        started = time.perf_counter()
        result = reports.load_flow_result(table, config, source_file=uploaded.name)
        if result.wants_excel:
            result.excel_bytes = reports.build_excel(result, "Load Flow Analysis")
        st.session_state["lf_result"] = {
            "docx": result.document_bytes,
            "xlsx": result.excel_bytes,
            "docx_name": result.default_filename,
            "xlsx_name": result.excel_filename,
            "detail": result.detail,
            "count": result.count,
            "seconds": time.perf_counter() - started,
            "formats": config.output_format,
        }
    right.caption(_status_caption(table, config))

    _load_flow_downloads()


def _load_flow_configuration() -> lf_config.LoadFlowConfig:
    """The Bus Info / Load Flow Results / Alert / Display / Output panel."""
    st.markdown("**2.  Select the information to report**")
    st.caption("Bus ID is always the first column.")

    bus_column, results_column = st.columns(2)

    with bus_column:
        st.markdown("**Bus Info**")
        bus_info = _field_checkboxes(lf_fields.BUS_INFO_ORDER, "bus")

    with results_column:
        st.markdown("**Load Flow Results**")
        results = _field_checkboxes(lf_fields.RESULT_ORDER, "res")

    st.markdown("**3.  Alert limits**")
    st.caption("A bus at or beyond a limit is flagged. Clear a checkbox to switch that limit off.")
    loading = _alert_row("Loading", "loading",
                         lf_config.DEFAULT_LOADING_CRITICAL, lf_config.DEFAULT_LOADING_MARGINAL)
    overvoltage = _alert_row("Overvoltage", "over",
                             lf_config.DEFAULT_OVERVOLTAGE_CRITICAL,
                             lf_config.DEFAULT_OVERVOLTAGE_MARGINAL)
    undervoltage = _alert_row("Undervoltage", "under",
                              lf_config.DEFAULT_UNDERVOLTAGE_CRITICAL,
                              lf_config.DEFAULT_UNDERVOLTAGE_MARGINAL)

    with st.expander("Display options, remarks wording and output format", expanded=True):
        display_left, display_right = st.columns(2)
        power_unit = display_left.radio(
            "Power unit", list(lf_fields.POWER_UNITS), key="lf_power", horizontal=True,
            help="kVA shows kW / kvar / kVA as ETAP reports them; MVA divides by 1000.",
        )
        voltage_display = display_right.radio(
            "Voltage", lf_fields.VOLTAGE_DISPLAYS, key="lf_voltage_display", horizontal=True,
            help="Actual Value = Nominal kV x Voltage % / 100.",
        )

        remark_left, remark_right = st.columns(2)
        style = remark_left.radio(
            "Remarks wording",
            ["ACCEPTABLE / MARGINAL / CRITICAL", "ACCEPTABLE / NOT ACCEPTABLE"],
            key="lf_style",
            help="The second option is the wording earlier versions produced.",
        )
        include_cause = remark_right.checkbox(
            "Add an 'Alert' column naming the limit that was hit",
            value=False, key="lf_cause",
        )
        include_remarks = remark_right.checkbox("Include the Remarks column", value=True, key="lf_remarks")

        output_format = st.radio(
            "**4.  Output format**", lf_config.OUTPUT_FORMATS,
            index=lf_config.OUTPUT_FORMATS.index(lf_config.BOTH),
            key="lf_format", horizontal=True,
        )

    return lf_config.LoadFlowConfig(
        bus_info=tuple(bus_info),
        results=tuple(results),
        loading=loading,
        overvoltage=overvoltage,
        undervoltage=undervoltage,
        power_unit=power_unit,
        voltage_display=voltage_display,
        remark_style=(
            lf_config.STATUS_STYLE
            if style.startswith("ACCEPTABLE / MARGINAL")
            else lf_config.LEGACY_STYLE
        ),
        include_remarks=include_remarks,
        include_alert_cause=include_cause,
        output_format=output_format,
        include_pseudo_buses=st.session_state.get("lf_pseudo", False),
    )


def _field_checkboxes(order, prefix: str) -> list[str]:
    """A group of field checkboxes with Select all / Clear all."""
    defaults = set(lf_fields.DEFAULT_BUS_INFO if prefix == "bus" else lf_fields.DEFAULT_RESULTS)

    select_all, clear_all = st.columns(2)
    if select_all.button("Select all", key=f"{prefix}_all", use_container_width=True):
        for key in order:
            if lf_fields.get(key).available:
                st.session_state[f"{prefix}_{key}"] = True
    if clear_all.button("Clear all", key=f"{prefix}_none", use_container_width=True):
        for key in order:
            st.session_state[f"{prefix}_{key}"] = False

    chosen: list[str] = []
    for key in order:
        definition = lf_fields.get(key)
        label = definition.label_for(lf_fields.DisplayContext())
        if not definition.available:
            st.checkbox(f"{label} (not in this report)", value=False, disabled=True,
                        key=f"{prefix}_{key}", help=definition.note)
            continue
        if st.checkbox(label, value=key in defaults, key=f"{prefix}_{key}",
                       help=definition.note or None):
            chosen.append(key)
    return chosen


def _alert_row(label: str, prefix: str, critical: float, marginal: float) -> lf_config.AlertLimit:
    """One Alert row: enable + value for Critical and Marginal."""
    name, critical_on, critical_value, marginal_on, marginal_value = st.columns([2, 1, 2, 1, 2])
    name.markdown(f"<div style='padding-top:0.55rem'>{label}</div>", unsafe_allow_html=True)
    on_c = critical_on.checkbox("C", value=True, key=f"{prefix}_c_on",
                                label_visibility="collapsed", help=f"{label} critical limit on")
    value_c = critical_value.number_input(
        f"{label} critical %", value=float(critical), min_value=0.0, max_value=999.0,
        step=0.5, key=f"{prefix}_c", label_visibility="collapsed", disabled=not on_c,
    )
    on_m = marginal_on.checkbox("M", value=True, key=f"{prefix}_m_on",
                                label_visibility="collapsed", help=f"{label} marginal limit on")
    value_m = marginal_value.number_input(
        f"{label} marginal %", value=float(marginal), min_value=0.0, max_value=999.0,
        step=0.5, key=f"{prefix}_m", label_visibility="collapsed", disabled=not on_m,
    )
    return lf_config.AlertLimit(value_c, value_m, on_c, on_m)


def _status_caption(table, config) -> str:
    parts = []
    for status, count in table.status_counts.items():
        if not status:
            continue
        import loadflow_processor

        parts.append(f"{count} {loadflow_processor.remark_text(status, config)}")
    return "  |  ".join(parts)


def _load_flow_downloads() -> None:
    result = st.session_state.get("lf_result")
    if not result:
        return

    st.success("Report Generated Successfully")
    st.caption(f"{result['count']} buses  •  {result['detail']}  •  {result['seconds']:.1f} s")

    word_column, excel_column = st.columns(2)
    if result["docx"]:
        word_column.download_button(
            f"Download {result['docx_name']}", data=result["docx"],
            file_name=result["docx_name"], mime=WORD_MIME,
            type="primary", use_container_width=True,
        )
    if result["xlsx"]:
        excel_column.download_button(
            f"Download {result['xlsx_name']}", data=result["xlsx"],
            file_name=result["xlsx_name"], mime=EXCEL_MIME,
            use_container_width=True,
        )


# --------------------------------------------------------------------------- #
# Short Circuit - unchanged workflow
# --------------------------------------------------------------------------- #


def _render_short_circuit(report) -> None:
    uploaded = st.file_uploader(
        f"**1.  {report.upload_label}**", type=["pdf"], key="sc_upload",
    )

    if uploaded is None:
        st.info("Upload an ETAP Short Circuit Study report to configure and preview the table.")
        return

    # Parse once. Cfactor / X/R come from the per-bus detail pages, so the
    # expensive scan only runs when one of those columns is selected.
    want_detail = sc_fields.needs_detail(
        list(st.session_state.get("sc_info_sel", ())) + list(st.session_state.get("sc_res_sel", ()))
    )
    try:
        with st.spinner("Reading the ETAP report ..." if not want_detail
                        else "Reading the ETAP report (including per-bus detail pages) ..."):
            records, study_info = parse_short_circuit(uploaded.getvalue(), want_detail)
    except ParserError as exc:
        st.error(str(exc))
        return
    except Exception:
        st.error("Unable to read the ETAP Short Circuit Report.")
        return

    st.success(f"{len(records)} buses extracted from {uploaded.name}")

    config = _short_circuit_configuration(study_info)

    try:
        table = reports.short_circuit_table(records, config, study_info)
    except sc_config.ConfigError as exc:
        st.warning(str(exc))
        return
    except sc_processor.EmptyTableError as exc:
        st.warning(str(exc))
        return

    st.markdown("**5.  Preview**")
    st.caption(
        f"{table.count} rows x {len(table.headers)} columns - "
        "exactly what the Word and Excel reports will contain."
    )
    st.dataframe(
        [dict(zip(table.headers, row)) for row in table.rows],
        hide_index=True, use_container_width=True, height=360,
    )
    if table.filtered_out:
        st.caption("Filtered out: " + ", ".join(
            f"{count} {name}" for name, count in table.filtered_out.items()))
    for warning in table.warnings:
        st.warning(warning)

    left, right = st.columns([1, 2])
    if left.button("Generate Report", type="primary", use_container_width=True, key="sc_go"):
        started = time.perf_counter()
        result = reports.short_circuit_result(
            table, config, source_file=uploaded.name, study_info=study_info
        )
        if result.wants_excel:
            result.excel_bytes = reports.build_excel(result, "Short Circuit")
        st.session_state["sc_result"] = {
            "docx": result.document_bytes,
            "xlsx": result.excel_bytes,
            "docx_name": result.default_filename,
            "xlsx_name": result.excel_filename,
            "detail": result.detail,
            "count": result.count,
            "seconds": time.perf_counter() - started,
        }
    right.caption(table and _sc_status_caption(table, config))

    _short_circuit_downloads()


def _short_circuit_configuration(study_info: dict):
    """The Study / Info / Results / Alert / Units panel."""
    st.markdown("**2.  Study**")
    study_left, study_middle, study_right = st.columns(3)

    detected_standard = study_info.get("standard")
    standard = study_left.radio(
        "Standard", sc_config.STANDARDS,
        index=sc_config.STANDARDS.index(detected_standard)
        if detected_standard in sc_config.STANDARDS else 0,
        key="sc_standard", horizontal=True,
        help="Detected from the report header." if detected_standard else None,
    )
    detected_type = study_info.get("study_type")
    study_type = study_middle.radio(
        "Study type", sc_config.STUDY_TYPES,
        format_func=lambda t: sc_config.STUDY_TYPE_LABELS[t],
        index=sc_config.STUDY_TYPES.index(detected_type)
        if detected_type in sc_config.STUDY_TYPES else 0,
        key="sc_study_type",
    )
    reports_found = [study_info.get("study_case")] if study_info.get("study_case") else []
    selected_report = study_right.selectbox(
        "Report", reports_found or ["(unnamed)"], key="sc_report",
        help="The study cases this report contains, read from its header.",
    )

    if detected_standard and standard != detected_standard:
        st.warning(
            f"This report was run to {detected_standard}, not {standard}. "
            "The values shown are the ones ETAP calculated; changing the "
            "standard here only relabels the report, it does not recalculate."
        )
    if detected_type and study_type != detected_type:
        st.warning(
            f"This report contains {sc_config.STUDY_TYPE_LABELS[detected_type]} results only. "
            f"No {sc_config.STUDY_TYPE_LABELS[study_type]} data exists in the file, so the "
            "table below is still the fault currents ETAP actually calculated."
        )

    st.markdown("**3.  Information to report**")
    st.caption("ID is always the first column.")
    info_column, results_column = st.columns(2)
    with info_column:
        st.markdown("**Info**")
        info = _sc_checkboxes(sc_fields.INFO_ORDER, "sc_info", sc_fields.DEFAULT_INFO)
    with results_column:
        st.markdown("**Results**")
        results = _sc_checkboxes(sc_fields.RESULT_ORDER, "sc_res", sc_fields.DEFAULT_RESULTS)

    st.session_state["sc_info_sel"] = tuple(info)
    st.session_state["sc_res_sel"] = tuple(results)

    st.markdown("**4.  Alert, filters and units**")
    alert_column, filter_column = st.columns(2)

    with alert_column:
        critical_on, critical_value = st.columns([1, 2])
        on_c = critical_on.checkbox("Critical", value=True, key="sc_crit_on")
        crit = critical_value.number_input(
            "Critical %", value=float(sc_config.DEFAULT_CRITICAL), min_value=0.0,
            max_value=1000.0, step=1.0, key="sc_crit", label_visibility="collapsed",
            disabled=not on_c,
        )
        marginal_on, marginal_value = st.columns([1, 2])
        on_m = marginal_on.checkbox("Marginal", value=True, key="sc_marg_on")
        marg = marginal_value.number_input(
            "Marginal %", value=float(sc_config.DEFAULT_MARGINAL), min_value=0.0,
            max_value=1000.0, step=1.0, key="sc_marg", label_visibility="collapsed",
            disabled=not on_m,
        )
        st.caption("Percent of the equipment making (peak) capacity: ip / Rated ip.")
        skip_non_alerted = st.checkbox(
            "Skip non-alerted devices", value=False, key="sc_skip_alert",
            help="Keep only rows that reach a threshold.",
        )
        skip_nodes = st.checkbox(
            "Skip nodes", value=True, key="sc_skip_nodes",
            help="Exclude ETAP's auto-generated internal nodes (IDs containing '~').",
        )

    with filter_column:
        equipment = st.multiselect(
            "Device type", list(sc_fields.EQUIPMENT_TYPES), default=[],
            key="sc_equipment",
            help="How ETAP classifies each bus. Leave empty for every type.",
        )
        unit_left, unit_right = st.columns(2)
        current_unit = unit_left.selectbox(
            "Current", list(sc_fields.CURRENT_UNITS), key="sc_current_unit"
        )
        voltage_unit = unit_right.selectbox(
            "Voltage", list(sc_fields.VOLTAGE_UNITS), key="sc_voltage_unit"
        )
        include_duty = st.checkbox("Add a 'Duty (%)' column", value=False, key="sc_duty")
        style = st.radio(
            "Remarks wording",
            ["ACCEPTABLE / MARGINAL / CRITICAL", "ACCEPTABLE / NOT ACCEPTABLE"],
            key="sc_style",
        )
        output_format = st.radio(
            "Output format", sc_config.OUTPUT_FORMATS,
            index=sc_config.OUTPUT_FORMATS.index(sc_config.BOTH),
            key="sc_format", horizontal=True,
        )

    return sc_config.ShortCircuitConfig(
        standard=standard,
        study_type=study_type,
        selected_report=study_info.get("study_case", "") or "",
        info=tuple(info),
        results=tuple(results),
        include_duty=include_duty,
        remark_style=(
            sc_config.STATUS_STYLE if style.startswith("ACCEPTABLE / MARGINAL")
            else sc_config.LEGACY_STYLE
        ),
        critical=crit, critical_enabled=on_c,
        marginal=marg, marginal_enabled=on_m,
        skip_non_alerted=skip_non_alerted,
        skip_nodes=skip_nodes,
        equipment_types=tuple(equipment),
        current_unit=current_unit,
        voltage_unit=voltage_unit,
        output_format=output_format,
    )


def _sc_checkboxes(order, prefix: str, defaults) -> list[str]:
    """A group of Short Circuit field checkboxes with Select all / Clear all."""
    select_all, clear_all = st.columns(2)
    if select_all.button("Select all", key=f"{prefix}_all", use_container_width=True):
        for key in order:
            if sc_fields.get(key).available:
                st.session_state[f"{prefix}_{key}"] = True
    if clear_all.button("Clear all", key=f"{prefix}_none", use_container_width=True):
        for key in order:
            st.session_state[f"{prefix}_{key}"] = False

    chosen: list[str] = []
    for key in order:
        definition = sc_fields.get(key)
        label = definition.label_for(sc_fields.DisplayContext())
        if not definition.available:
            st.checkbox(f"{label} (not in this report)", value=False, disabled=True,
                        key=f"{prefix}_{key}", help=definition.note)
            continue
        suffix = "  (scans the whole report)" if definition.needs_detail else ""
        if st.checkbox(f"{label}{suffix}", value=key in defaults, key=f"{prefix}_{key}",
                       help=definition.note or None):
            chosen.append(key)
    return chosen


def _sc_status_caption(table, config) -> str:
    parts = []
    for status, count in table.status_counts.items():
        if not status:
            continue
        parts.append(f"{count} {sc_processor.remark_text(status, config)}")
    return "  |  ".join(parts)


def _short_circuit_downloads() -> None:
    result = st.session_state.get("sc_result")
    if not result:
        return

    st.success("Report Generated Successfully")
    st.caption(f"{result['count']} rows  •  {result['detail']}  •  {result['seconds']:.1f} s")

    word_column, excel_column = st.columns(2)
    if result["docx"]:
        word_column.download_button(
            f"Download {result['docx_name']}", data=result["docx"],
            file_name=result["docx_name"], mime=WORD_MIME,
            type="primary", use_container_width=True,
        )
    if result["xlsx"]:
        excel_column.download_button(
            f"Download {result['xlsx_name']}", data=result["xlsx"],
            file_name=result["xlsx_name"], mime=EXCEL_MIME,
            use_container_width=True,
        )


if __name__ == "__main__":
    render()
