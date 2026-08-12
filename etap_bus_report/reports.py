"""
reports.py
----------
Registry of the report types the application supports.

Every front end (desktop, web, CLI) drives the application through this module,
so adding a new ETAP report - Motor Starting, Arc Flash, Cable schedule - means
adding one :class:`ReportType` here plus its parser, writer and template.  No
front end changes.

Each report type declares:

* the parser and writer to use, and the template that ships with it,
* which extra inputs the UI must show (``inputs``),
* the columns of the preview table.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

import dynamic_word_writer
import excel_writer
import loadflow_config
import loadflow_processor
import pdf_parser
import sc_rules
import shortcircuit_config
import shortcircuit_fields
import shortcircuit_parser
import shortcircuit_processor
import shortcircuit_writer
import template_mapping
import voltage_checker
import word_writer
from utils import ParserError

LOAD_FLOW = "load_flow"
SHORT_CIRCUIT = "short_circuit"


# --------------------------------------------------------------------------- #
# Result
# --------------------------------------------------------------------------- #


@dataclass
class ReportResult:
    """Everything a front end needs after a successful run."""

    document_bytes: bytes
    headers: list[str]
    rows: list[list[str]]
    count: int
    flagged: int
    detail: str = ""
    default_filename: str = "Report.docx"
    #: The same table as an .xlsx workbook, filled in by :func:`generate`.
    excel_bytes: bytes = b""
    label: str = ""
    #: Notes about columns the ETAP report could not fill.
    warnings: list = field(default_factory=list)
    #: Which documents were requested (Word / Excel / Word + Excel).
    formats: str = "Word + Excel"

    @property
    def wants_word(self) -> bool:
        return self.formats in ("Word", "Word + Excel")

    @property
    def wants_excel(self) -> bool:
        return self.formats in ("Excel", "Word + Excel")

    #: Title printed at the top of the workbook (blank for template reports).
    title: str = ""

    @property
    def excel_filename(self) -> str:
        """Matching workbook name: same stem as the Word document."""
        return os.path.splitext(self.default_filename)[0] + ".xlsx"


# --------------------------------------------------------------------------- #
# Load Flow (unchanged behaviour)
# --------------------------------------------------------------------------- #


def load_flow_records(pdf_path: str, include_pseudo_buses: bool = False):
    """
    Parse a Load Flow report with the existing parser.

    Kept separate from :func:`_run_load_flow` so a front end can parse once,
    hold the records in session state, and rebuild the table as the user
    changes selections without touching the PDF again.
    """
    return pdf_parser.load_report(pdf_path, include_pseudo_buses=include_pseudo_buses)


def load_flow_table(records, config: loadflow_config.LoadFlowConfig):
    """Apply a configuration to already-parsed records (no PDF access)."""
    return loadflow_processor.build_table(records, config)


def load_flow_result(
    table,
    config: loadflow_config.LoadFlowConfig,
    source_file: Optional[str] = None,
) -> "ReportResult":
    """
    Build the Word and/or Excel documents from an existing final table.

    The preview, the document and the workbook are all rendered from *table*,
    so they cannot disagree.
    """
    metadata = dynamic_word_writer.default_metadata(
        source_file,
        **{"Limits applied": _limits_summary(config)},
    )

    document = b""
    if config.wants_word:
        document = dynamic_word_writer.build_document_bytes(
            table.headers, table.rows,
            title="LOAD FLOW ANALYSIS REPORT",
            metadata=metadata,
            notes=table.warnings,
        )

    result = ReportResult(
        document_bytes=document,
        headers=list(table.headers),
        rows=[list(row) for row in table.rows],
        count=table.count,
        flagged=table.flagged,
        detail=_status_summary(table, config),
        default_filename="LoadFlow_Report.docx",
        title="LOAD FLOW ANALYSIS REPORT",
        warnings=list(table.warnings),
        formats=config.output_format,
    )
    return result


def _limits_summary(config: loadflow_config.LoadFlowConfig) -> str:
    parts = []
    for name, limit in (
        ("Loading", config.loading),
        ("Overvoltage", config.overvoltage),
        ("Undervoltage", config.undervoltage),
    ):
        pieces = []
        if limit.critical_value is not None:
            pieces.append(f"critical {limit.critical_value:g} %")
        if limit.marginal_value is not None:
            pieces.append(f"marginal {limit.marginal_value:g} %")
        if pieces:
            parts.append(f"{name} {', '.join(pieces)}")
    return "; ".join(parts)


def _status_summary(table, config: loadflow_config.LoadFlowConfig) -> str:
    """Counts worded exactly as the Remarks column words them."""
    counts: dict[str, int] = {}
    for status, number in table.status_counts.items():
        if not status:
            continue
        text = loadflow_processor.remark_text(status, config)
        counts[text] = counts.get(text, 0) + number

    order = ["ACCEPTABLE", "MARGINAL", "CRITICAL", "NOT ACCEPTABLE"]
    parts = [f"{counts[text]} {text}" for text in order if counts.get(text)]
    if table.status_counts.get(""):
        parts.append(f"{table.status_counts['']} not assessed")
    return "  |  ".join(parts)


def _run_load_flow(pdf_path: str, options: dict) -> ReportResult:
    """
    One-shot Load Flow run (command line, tests, and any caller that has not
    adopted the configurable interface).

    ``config=`` takes a :class:`loadflow_config.LoadFlowConfig`; passing
    ``limits="95-106"`` instead reproduces the previous behaviour exactly.
    """
    config = options.get("config")
    if config is None:
        config = loadflow_config.from_legacy_limits(
            options.get("limits", "95-106"),
            include_pseudo_buses=options.get("include_pseudo_buses", False),
        )
    config.include_pseudo_buses = options.get(
        "include_pseudo_buses", config.include_pseudo_buses
    )
    config.validate()

    records = load_flow_records(pdf_path, config.include_pseudo_buses)
    table = load_flow_table(records, config)
    return load_flow_result(table, config, source_file=pdf_path)


# --------------------------------------------------------------------------- #
# Short Circuit
# --------------------------------------------------------------------------- #

#: ETAP equipment types that represent real switchgear assemblies.
EQUIPMENT_TYPES = {"switchgear", "switchboard", "mcc", "panelboard", "switchrack", "panel"}


def short_circuit_records(pdf_path: str, include_nodes: bool = True):
    """
    Parse a Short Circuit report with the existing parser, once.

    Nodes are kept here and filtered later, so toggling 'Skip Nodes' in the
    interface never needs another parse.
    """
    return shortcircuit_parser.load_report(pdf_path, include_pseudo_buses=include_nodes)


def short_circuit_study_info(pdf_path: str) -> dict:
    """Standard / study type / study case, as stated by the report itself."""
    return shortcircuit_parser.study_info(pdf_path)


def short_circuit_detail(pdf_path: str, records):
    """Add Cfactor and X/R from the per-bus detail pages (whole-document scan)."""
    return shortcircuit_parser.add_detail_fields(pdf_path, records)


def short_circuit_table(records, config, study_info: Optional[dict] = None):
    """Apply a configuration to already-parsed records (no PDF access)."""
    return shortcircuit_processor.build_table(records, config, study_info)


def short_circuit_result(
    table,
    config,
    source_file: Optional[str] = None,
    study_info: Optional[dict] = None,
) -> "ReportResult":
    """Build the Word and/or Excel documents from an existing final table."""
    info = study_info or {}
    metadata = dynamic_word_writer.default_metadata(
        source_file,
        **{
            "Standard": info.get("title") or config.standard,
            "Study type": shortcircuit_config.STUDY_TYPE_LABELS.get(
                config.study_type, config.study_type
            ),
            "Report": config.selected_report,
            "Units": f"current {config.current_unit}, voltage {config.voltage_unit}",
            "Filters": _filter_summary(config),
            "Alert thresholds": _threshold_summary(config),
        },
    )

    document = b""
    if config.wants_word:
        document = dynamic_word_writer.build_document_bytes(
            table.headers, table.rows,
            title="SHORT CIRCUIT ANALYSIS REPORT",
            metadata=metadata,
            notes=table.warnings,
        )

    # Name the file after the study case, whether it came from the interface
    # or straight from the report header.
    name = config.selected_report or info.get("study_case") or ""
    stem = f"ShortCircuit_{name}" if name else "ShortCircuit_Report"

    return ReportResult(
        document_bytes=document,
        headers=list(table.headers),
        rows=[list(row) for row in table.rows],
        count=table.count,
        flagged=table.flagged,
        detail=_sc_status_summary(table, config),
        default_filename=f"{stem}.docx",
        title="SHORT CIRCUIT ANALYSIS REPORT",
        warnings=list(table.warnings),
        formats=config.output_format,
    )


def _filter_summary(config) -> str:
    parts = []
    if config.equipment_types:
        parts.append("equipment type " + "/".join(config.equipment_types))
    else:
        parts.append("all equipment types")
    parts.append("nodes excluded" if config.skip_nodes else "nodes included")
    if config.skip_non_alerted:
        parts.append("alerted rows only")
    return "; ".join(parts)


def _threshold_summary(config) -> str:
    parts = []
    if config.critical_value is not None:
        parts.append(f"critical {config.critical_value:g} %")
    if config.marginal_value is not None:
        parts.append(f"marginal {config.marginal_value:g} %")
    return ", ".join(parts) + (" of the equipment rating" if parts else "none")


def _sc_status_summary(table, config) -> str:
    counts: dict[str, int] = {}
    for status, number in table.status_counts.items():
        if not status:
            continue
        text = shortcircuit_processor.remark_text(status, config)
        counts[text] = counts.get(text, 0) + number
    order = ["ACCEPTABLE", "MARGINAL", "CRITICAL", "NOT ACCEPTABLE"]
    parts = [f"{counts[t]} {t}" for t in order if counts.get(t)]
    if table.status_counts.get(""):
        parts.append(f"{table.status_counts['']} not rated")
    return "  |  ".join(parts)


def _run_short_circuit(pdf_path: str, options: dict) -> ReportResult:
    """
    One-shot Short Circuit run (command line, tests, and any caller that has
    not adopted the configurable interface).

    ``config=`` takes a :class:`shortcircuit_config.ShortCircuitConfig`.
    Without it, the previous template-driven behaviour is reproduced.
    """
    config = options.get("config")

    if config is None and not options.get("dynamic"):
        return _run_short_circuit_template(pdf_path, options)

    info = short_circuit_study_info(pdf_path)
    if config is None:
        config = shortcircuit_config.from_study_info(info)
        if options.get("equipment_only"):
            config.equipment_types = tuple(
                t for t in shortcircuit_fields.EQUIPMENT_TYPES if t != "Bus"
            )
        if "include_pseudo_buses" in options:
            config.skip_nodes = not options["include_pseudo_buses"]
    config.validate()

    records = short_circuit_records(pdf_path, include_nodes=True)
    if config.needs_detail_pages():
        short_circuit_detail(pdf_path, records)
    table = short_circuit_table(records, config, info)
    return short_circuit_result(table, config, source_file=pdf_path, study_info=info)


def _run_short_circuit_template(pdf_path: str, options: dict) -> ReportResult:
    """The original template-filling Short Circuit report, unchanged."""
    template = options.get("template") or shortcircuit_writer.template_path()
    field_keys = [key for _, key in shortcircuit_writer.describe_mapping(template)]

    buses = shortcircuit_parser.load_report(
        pdf_path,
        include_pseudo_buses=options.get("include_pseudo_buses", False),
        fault_type=options.get("fault_type", shortcircuit_parser.DEFAULT_FAULT_TYPE),
    )

    if options.get("equipment_only"):
        buses = [b for b in buses if (b.get("bus_type") or "").lower() in EQUIPMENT_TYPES]
        if not buses:
            raise ParserError(
                "No switchgear, switchboard, MCC or panel was found in the report. "
                "Turn off 'Switchgear and switchboards only' to list every bus."
            )

    if set(field_keys) & template_mapping.DETAIL_ONLY_FIELDS:
        shortcircuit_parser.add_detail_fields(pdf_path, buses)

    rule = options.get("rule") or sc_rules.DEFAULT_RULE
    document = shortcircuit_writer.build_document_bytes(buses, template=template, rule=rule)

    headers = [header for header, _ in shortcircuit_writer.describe_mapping(template)]
    rows = [
        [value if value is not None else "" for value in template_mapping.format_row(bus, field_keys)]
        for bus in buses
    ]
    flagged = sum(1 for bus in buses if bus.get("remarks") == "NOT ACCEPTABLE")
    rated = sum(1 for bus in buses if bus.get("rating_peak") is not None)

    return ReportResult(
        document_bytes=document,
        headers=headers,
        rows=rows,
        count=len(buses),
        flagged=flagged,
        detail=f"{rated} of {len(buses)} buses have an equipment rating in the report",
        default_filename="Short Circuit Report.docx",
    )


# --------------------------------------------------------------------------- #
# Registry
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class ReportType:
    key: str
    label: str
    upload_label: str
    template: str
    runner: Callable[[str, dict], ReportResult]
    default_filename: str
    #: Extra inputs the UI must render, in order.
    inputs: tuple = ()
    description: str = ""

    def run(self, pdf_path: str, **options: Any) -> ReportResult:
        return self.runner(pdf_path, options)

    def template_path(self) -> str:
        from utils import resource_path

        return resource_path(self.template)


REPORTS: dict[str, ReportType] = {
    LOAD_FLOW: ReportType(
        key=LOAD_FLOW,
        label="Load Flow Analysis",
        upload_label="Upload ETAP Load Flow Analysis Report (.pdf)",
        template=os.path.join("assets", "LoadFlow_Template.docx"),
        runner=_run_load_flow,
        default_filename="Bus Report.docx",
        inputs=("limits",),
        description="Bus voltages and loading, with selectable columns and configurable alert limits.",
    ),
    SHORT_CIRCUIT: ReportType(
        key=SHORT_CIRCUIT,
        label="Short Circuit Study",
        upload_label="Upload ETAP Short Circuit Study Report (.pdf)",
        template=os.path.join("assets", "ShortCircuit_Template.docx"),
        runner=_run_short_circuit,
        default_filename="ShortCircuit_Report.docx",
        inputs=(),
        description="Bus fault currents against equipment short-circuit ratings, "
                    "with selectable columns and configurable alert thresholds.",
    ),
}

ORDER = (LOAD_FLOW, SHORT_CIRCUIT)


def get(key: str) -> ReportType:
    try:
        return REPORTS[key]
    except KeyError:
        raise KeyError(f"Unknown report type {key!r}. Available: {', '.join(ORDER)}") from None


def by_label(label: str) -> ReportType:
    for report in REPORTS.values():
        if report.label == label:
            return report
    raise KeyError(f"Unknown report type {label!r}")


def labels() -> list[str]:
    return [REPORTS[key].label for key in ORDER]


def generate(report_key: str, pdf_path: str, **options: Any) -> ReportResult:
    """
    Run one report end to end and return both documents plus a summary.

    The Excel workbook is built from the very same headers and rows that were
    written into the Word document - the PDF is parsed once only.
    """
    report = get(report_key)
    result = report.run(pdf_path, **options)
    result.label = report.label
    if result.wants_excel:
        result.excel_bytes = build_excel(result, report.label)
    return result


def build_excel(result: ReportResult, label: str = "") -> bytes:
    """The result's table as a workbook - same headers, same rows."""
    return excel_writer.build_workbook_bytes(
        result.headers,
        result.rows,
        sheet_name=label or result.label or "Report",
        title=result.title,
    )
