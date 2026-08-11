"""
ui.py
-----
PySide6 desktop interface.

Two pages, chosen by the Report Type selector at the top:

* **Load Flow Analysis** - the configurable interface: pick the Bus Info and
  Load Flow Results columns, set the Critical/Marginal alert limits, choose
  display units and the output format, watch the preview update, then export.
  The PDF is parsed once when it is loaded; changing a selection only rebuilds
  the table.
* **Short Circuit Study** - the existing template-driven workflow, unchanged.

Parsing runs in a worker thread so the window never freezes.
"""

from __future__ import annotations

import os

from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import excel_writer
import loadflow_config as lf_config
import loadflow_fields as lf_fields
import reports
import sc_rules
from utils import APP_NAME, APP_VERSION, ParserError

STYLESHEET = """
QWidget       { background: #ffffff; color: #1f2430; font-size: 13px; }
QLabel#title  { font-size: 20px; font-weight: 600; color: #524689; }
QLabel#hint   { color: #6b7280; font-size: 11px; }
QLabel#group  { font-weight: 600; }
QLineEdit, QComboBox, QDoubleSpinBox { border: 1px solid #d0d5dd; border-radius: 5px; padding: 4px 6px; }
QPushButton   { border: 1px solid #d0d5dd; border-radius: 6px; padding: 7px 14px; background: #f7f7fb; }
QPushButton:hover { background: #eeeef6; }
QPushButton#primary { background: #524689; color: #ffffff; border: none; font-weight: 600; }
QPushButton#primary:hover   { background: #453a76; }
QPushButton#primary:disabled{ background: #b6b1cf; }
QFrame#card   { border: 1px solid #e5e7eb; border-radius: 8px; }
QTableWidget  { border: 1px solid #e5e7eb; gridline-color: #e5e7eb; }
QHeaderView::section { background: #524689; color: #ffffff; padding: 4px; border: none; font-weight: 600; }
"""


# --------------------------------------------------------------------------- #
# Workers
# --------------------------------------------------------------------------- #


class ParseWorker(QObject):
    """Parses a Load Flow PDF once, off the GUI thread."""

    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, pdf_path: str, include_pseudo: bool) -> None:
        super().__init__()
        self._pdf_path = pdf_path
        self._include_pseudo = include_pseudo

    def run(self) -> None:
        try:
            records = reports.load_flow_records(self._pdf_path, self._include_pseudo)
        except ParserError as exc:
            self.failed.emit(str(exc))
        except Exception:
            self.failed.emit("Unable to read the ETAP report.")
        else:
            self.finished.emit(records)


class ShortCircuitWorker(QObject):
    """Runs a Short Circuit report off the GUI thread (unchanged workflow)."""

    finished = Signal(object, str)
    failed = Signal(str)

    def __init__(self, pdf_path: str, output_path: str, options: dict) -> None:
        super().__init__()
        self._pdf_path = pdf_path
        self._output_path = output_path
        self._options = options

    def run(self) -> None:
        excel_path = excel_writer.excel_path_for(self._output_path)
        try:
            result = reports.generate(reports.SHORT_CIRCUIT, self._pdf_path, **self._options)
            folder = os.path.dirname(os.path.abspath(self._output_path))
            if folder:
                os.makedirs(folder, exist_ok=True)
            with open(self._output_path, "wb") as handle:
                handle.write(result.document_bytes)
            excel_writer.write_excel(result.excel_bytes, excel_path)
        except ParserError as exc:
            self.failed.emit(str(exc))
        except PermissionError:
            self.failed.emit(
                "Unable to save the report. Close the document if it is open in "
                "Word or Excel and try again."
            )
        except Exception:
            self.failed.emit("Unable to read the ETAP report.")
        else:
            self.finished.emit(result, excel_path)


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #


def card(title: str) -> QFrame:
    frame = QFrame()
    frame.setObjectName("card")
    box = QVBoxLayout(frame)
    box.setContentsMargins(14, 12, 14, 12)
    box.setSpacing(8)
    if title:
        label = QLabel(title)
        label.setObjectName("group")
        box.addWidget(label)
    return frame


def hint(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("hint")
    label.setWordWrap(True)
    return label


# --------------------------------------------------------------------------- #
# Load Flow page
# --------------------------------------------------------------------------- #


class LoadFlowPage(QWidget):
    """Configurable Load Flow interface."""

    status_message = Signal(str)
    busy = Signal(bool)

    def __init__(self) -> None:
        super().__init__()
        self._pdf_path: str | None = None
        self._records: list | None = None
        self._table = None
        self._thread: QThread | None = None
        self._worker: ParseWorker | None = None

        self._build()

    # -- construction ------------------------------------------------------- #

    def _build(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(12)

        outer.addWidget(self._upload_card())

        row = QHBoxLayout()
        row.addWidget(self._fields_card(), 2)
        row.addWidget(self._alert_card(), 3)
        outer.addLayout(row)

        outer.addWidget(self._display_card())
        outer.addWidget(self._preview_card(), 1)

        buttons = QHBoxLayout()
        self.reset_button = QPushButton("Reset")
        self.reset_button.clicked.connect(self.reset)
        self.generate_button = QPushButton("Generate Report")
        self.generate_button.setObjectName("primary")
        self.generate_button.clicked.connect(self.generate)
        buttons.addWidget(self.reset_button)
        buttons.addStretch(1)
        buttons.addWidget(self.generate_button)
        outer.addLayout(buttons)

    def _upload_card(self) -> QFrame:
        frame = card("1.  Upload ETAP Load Flow Analysis Report (.pdf)")
        row = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("hint")
        browse = QPushButton("Choose PDF")
        browse.clicked.connect(self.choose_pdf)
        row.addWidget(self.file_label, 1)
        row.addWidget(browse, 0)
        frame.layout().addLayout(row)

        self.pseudo_check = QCheckBox("Include ETAP internal nodes (IDs containing '~')")
        self.pseudo_check.toggled.connect(self._reparse)
        frame.layout().addWidget(self.pseudo_check)
        return frame

    def _fields_card(self) -> QFrame:
        frame = card("2.  Information to report")
        frame.layout().addWidget(hint("Bus ID is always the first column."))

        columns = QHBoxLayout()
        self.bus_checks = self._checkbox_group(
            "Bus Info", lf_fields.BUS_INFO_ORDER, lf_fields.DEFAULT_BUS_INFO
        )
        self.result_checks = self._checkbox_group(
            "Load Flow Results", lf_fields.RESULT_ORDER, lf_fields.DEFAULT_RESULTS
        )
        columns.addWidget(self.bus_checks["widget"])
        columns.addWidget(self.result_checks["widget"])
        frame.layout().addLayout(columns)
        return frame

    def _checkbox_group(self, title: str, order, defaults) -> dict:
        widget = QWidget()
        box = QVBoxLayout(widget)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(4)

        label = QLabel(title)
        label.setObjectName("group")
        box.addWidget(label)

        checks: dict[str, QCheckBox] = {}
        for key in order:
            definition = lf_fields.get(key)
            check = QCheckBox(definition.label_for(lf_fields.DisplayContext()))
            check.setChecked(key in defaults and definition.available)
            check.setEnabled(definition.available)
            if definition.note:
                check.setToolTip(definition.note)
            if not definition.available:
                check.setText(f"{check.text()} (not in this report)")
            check.toggled.connect(self.refresh_preview)
            checks[key] = check
            box.addWidget(check)

        buttons = QHBoxLayout()
        select_all = QPushButton("Select all")
        clear_all = QPushButton("Clear all")
        select_all.clicked.connect(lambda: self._set_all(checks, True))
        clear_all.clicked.connect(lambda: self._set_all(checks, False))
        buttons.addWidget(select_all)
        buttons.addWidget(clear_all)
        box.addLayout(buttons)
        box.addStretch(1)
        return {"widget": widget, "checks": checks}

    @staticmethod
    def _set_all(checks: dict, value: bool) -> None:
        for check in checks.values():
            if check.isEnabled():
                check.setChecked(value)

    def _alert_card(self) -> QFrame:
        frame = card("3.  Alert limits")
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)

        for column, text in ((1, "Critical (%)"), (3, "Marginal (%)")):
            label = QLabel(text)
            label.setObjectName("group")
            grid.addWidget(label, 0, column, 1, 2)

        self.alerts: dict[str, dict] = {}
        rows = (
            ("loading", "Loading", lf_config.DEFAULT_LOADING_CRITICAL,
             lf_config.DEFAULT_LOADING_MARGINAL),
            ("overvoltage", "Overvoltage", lf_config.DEFAULT_OVERVOLTAGE_CRITICAL,
             lf_config.DEFAULT_OVERVOLTAGE_MARGINAL),
            ("undervoltage", "Undervoltage", lf_config.DEFAULT_UNDERVOLTAGE_CRITICAL,
             lf_config.DEFAULT_UNDERVOLTAGE_MARGINAL),
        )
        for index, (key, label_text, critical, marginal) in enumerate(rows, start=1):
            grid.addWidget(QLabel(label_text), index, 0)
            entry = {}
            for offset, (kind, value) in enumerate((("critical", critical), ("marginal", marginal))):
                check = QCheckBox()
                check.setChecked(True)
                spin = QDoubleSpinBox()
                spin.setRange(0.0, 999.0)
                spin.setDecimals(1)
                spin.setSingleStep(0.5)
                spin.setValue(float(value))
                spin.setSuffix(" %")
                check.toggled.connect(spin.setEnabled)
                check.toggled.connect(self.refresh_preview)
                spin.valueChanged.connect(self.refresh_preview)
                grid.addWidget(check, index, 1 + offset * 2)
                grid.addWidget(spin, index, 2 + offset * 2)
                entry[kind] = (check, spin)
            self.alerts[key] = entry

        frame.layout().addLayout(grid)
        frame.layout().addWidget(
            hint("A bus at or beyond a limit is flagged. Clear a checkbox to switch that limit off.")
        )
        return frame

    def _display_card(self) -> QFrame:
        frame = card("Display options and output")
        grid = QGridLayout()
        grid.setHorizontalSpacing(14)

        grid.addWidget(QLabel("Power unit"), 0, 0)
        self.power_kva = QRadioButton("kVA")
        self.power_mva = QRadioButton("MVA")
        self.power_kva.setChecked(True)
        for index, button in enumerate((self.power_kva, self.power_mva)):
            button.toggled.connect(self.refresh_preview)
            grid.addWidget(button, 0, 1 + index)

        grid.addWidget(QLabel("Voltage"), 1, 0)
        self.voltage_percent = QRadioButton("%")
        self.voltage_actual = QRadioButton("Actual Value")
        self.voltage_percent.setChecked(True)
        for index, button in enumerate((self.voltage_percent, self.voltage_actual)):
            button.toggled.connect(self.refresh_preview)
            grid.addWidget(button, 1, 1 + index)

        grid.addWidget(QLabel("Remarks"), 2, 0)
        self.remark_combo = QComboBox()
        self.remark_combo.addItem("ACCEPTABLE / MARGINAL / CRITICAL", lf_config.STATUS_STYLE)
        self.remark_combo.addItem("ACCEPTABLE / NOT ACCEPTABLE", lf_config.LEGACY_STYLE)
        self.remark_combo.currentIndexChanged.connect(self.refresh_preview)
        grid.addWidget(self.remark_combo, 2, 1, 1, 2)

        self.cause_check = QCheckBox("Add an 'Alert' column naming the limit that was hit")
        self.cause_check.toggled.connect(self.refresh_preview)
        grid.addWidget(self.cause_check, 3, 0, 1, 3)

        grid.addWidget(QLabel("Output format"), 4, 0)
        self.format_combo = QComboBox()
        self.format_combo.addItems(list(lf_config.OUTPUT_FORMATS))
        self.format_combo.setCurrentText(lf_config.BOTH)
        grid.addWidget(self.format_combo, 4, 1, 1, 2)

        grid.setColumnStretch(3, 1)
        frame.layout().addLayout(grid)
        return frame

    def _preview_card(self) -> QFrame:
        frame = card("5.  Preview")
        self.preview_hint = hint("Upload a report to see the table.")
        frame.layout().addWidget(self.preview_hint)

        self.preview = QTableWidget(0, 0)
        self.preview.setAlternatingRowColors(True)
        self.preview.setEditTriggers(QTableWidget.NoEditTriggers)
        self.preview.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        self.preview.verticalHeader().setDefaultSectionSize(20)
        self.preview.setMinimumHeight(220)
        frame.layout().addWidget(self.preview)
        return frame

    # -- configuration ------------------------------------------------------ #

    def configuration(self) -> lf_config.LoadFlowConfig:
        def limit(key: str) -> lf_config.AlertLimit:
            critical_check, critical_spin = self.alerts[key]["critical"]
            marginal_check, marginal_spin = self.alerts[key]["marginal"]
            return lf_config.AlertLimit(
                critical_spin.value(), marginal_spin.value(),
                critical_check.isChecked(), marginal_check.isChecked(),
            )

        return lf_config.LoadFlowConfig(
            bus_info=tuple(k for k, c in self.bus_checks["checks"].items() if c.isChecked()),
            results=tuple(k for k, c in self.result_checks["checks"].items() if c.isChecked()),
            loading=limit("loading"),
            overvoltage=limit("overvoltage"),
            undervoltage=limit("undervoltage"),
            power_unit="kVA" if self.power_kva.isChecked() else "MVA",
            voltage_display="%" if self.voltage_percent.isChecked() else "Actual Value",
            remark_style=self.remark_combo.currentData(),
            include_alert_cause=self.cause_check.isChecked(),
            output_format=self.format_combo.currentText(),
            include_pseudo_buses=self.pseudo_check.isChecked(),
        )

    # -- actions ------------------------------------------------------------ #

    def choose_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select ETAP Load Flow Analysis report", "", "PDF files (*.pdf)"
        )
        if path:
            self._pdf_path = path
            self.file_label.setText(os.path.basename(path))
            self._reparse()

    def _reparse(self) -> None:
        if not self._pdf_path:
            return
        self.busy.emit(True)
        self.status_message.emit("Reading the ETAP report ...")

        self._thread = QThread(self)
        self._worker = ParseWorker(self._pdf_path, self.pseudo_check.isChecked())
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_parsed)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _on_parsed(self, records) -> None:
        self._records = records
        self.busy.emit(False)
        self.status_message.emit(f"{len(records)} buses extracted.")
        self.refresh_preview()

    def _on_failed(self, message: str) -> None:
        self._records = None
        self.busy.emit(False)
        self.status_message.emit(message)
        QMessageBox.warning(self, APP_NAME, message)

    def refresh_preview(self) -> None:
        """Rebuild the table from the parsed records - never re-reads the PDF."""
        if not self._records:
            return
        try:
            config = self.configuration()
            self._table = reports.load_flow_table(self._records, config)
        except lf_config.ConfigError as exc:
            self._table = None
            self.preview.clear()
            self.preview.setRowCount(0)
            self.preview.setColumnCount(0)
            self.preview_hint.setText(str(exc))
            return

        table = self._table
        self.preview.setColumnCount(len(table.headers))
        self.preview.setHorizontalHeaderLabels(table.headers)
        self.preview.setRowCount(len(table.rows))
        for row_index, row in enumerate(table.rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem(value)
                if column_index:
                    item.setTextAlignment(Qt.AlignCenter)
                self.preview.setItem(row_index, column_index, item)
        self.preview.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.preview.resizeColumnsToContents()

        notes = ("  " + "  ".join(table.warnings)) if table.warnings else ""
        self.preview_hint.setText(
            f"{table.count} rows x {len(table.headers)} columns - exactly what will be "
            f"exported.{notes}"
        )

    def reset(self) -> None:
        self._pdf_path = None
        self._records = None
        self._table = None
        self.file_label.setText("No file selected")
        self.preview.setRowCount(0)
        self.preview.setColumnCount(0)
        self.preview_hint.setText("Upload a report to see the table.")

    def generate(self) -> None:
        if not self._records:
            QMessageBox.warning(
                self, APP_NAME, "Please upload an ETAP Load Flow Analysis report (.pdf) first."
            )
            return
        try:
            config = self.configuration()
            config.validate()
        except lf_config.ConfigError as exc:
            QMessageBox.warning(self, APP_NAME, str(exc))
            return

        self.refresh_preview()
        if self._table is None:
            return

        suggested = os.path.join(os.path.dirname(self._pdf_path or ""), "LoadFlow_Report.docx")
        chosen, _ = QFileDialog.getSaveFileName(
            self, "Save report as", suggested, "Report (*.docx *.xlsx)"
        )
        if not chosen:
            return
        base = os.path.splitext(chosen)[0]

        result = reports.load_flow_result(self._table, config, source_file=self._pdf_path)
        written = []
        try:
            if config.wants_word:
                path = base + ".docx"
                with open(path, "wb") as handle:
                    handle.write(result.document_bytes)
                written.append(path)
            if config.wants_excel:
                path = base + ".xlsx"
                excel_writer.write_excel(
                    reports.build_excel(result, "Load Flow Analysis"), path
                )
                written.append(path)
        except PermissionError:
            QMessageBox.warning(
                self, APP_NAME,
                "Unable to save. Close the report if it is open in Word or Excel and try again.",
            )
            return

        self.status_message.emit("Report Generated Successfully  |  " + result.detail)
        QMessageBox.information(
            self, APP_NAME,
            "Report Generated Successfully\n\nSaved to:\n" + "\n".join(written),
        )


# --------------------------------------------------------------------------- #
# Short Circuit page (unchanged workflow)
# --------------------------------------------------------------------------- #


class ShortCircuitPage(QWidget):
    status_message = Signal(str)
    busy = Signal(bool)

    def __init__(self) -> None:
        super().__init__()
        self._pdf_path: str | None = None
        self._output_path: str | None = None
        self._thread: QThread | None = None
        self._worker: ShortCircuitWorker | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        upload = card("1.  Upload ETAP Short Circuit Study Report (.pdf)")
        row = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("hint")
        browse = QPushButton("Choose PDF")
        browse.clicked.connect(self.choose_pdf)
        row.addWidget(self.file_label, 1)
        row.addWidget(browse, 0)
        upload.layout().addLayout(row)
        layout.addWidget(upload)

        options = card("2.  Options")
        self.equipment_check = QCheckBox("Switchgear, switchboards, MCCs and panels only")
        self.pseudo_check = QCheckBox("Include ETAP internal nodes (IDs containing '~')")
        self.rule_combo = QComboBox()
        self.rule_combo.addItems(list(sc_rules.RULES))
        self.rule_combo.setCurrentText(sc_rules.DEFAULT_RULE)
        options.layout().addWidget(self.equipment_check)
        options.layout().addWidget(self.pseudo_check)
        rule_row = QHBoxLayout()
        rule_row.addWidget(QLabel("Remarks rule"))
        rule_row.addWidget(self.rule_combo, 1)
        options.layout().addLayout(rule_row)
        layout.addWidget(options)

        buttons = QHBoxLayout()
        self.reset_button = QPushButton("Reset")
        self.reset_button.clicked.connect(self.reset)
        self.generate_button = QPushButton("Generate Report")
        self.generate_button.setObjectName("primary")
        self.generate_button.clicked.connect(self.generate)
        buttons.addWidget(self.reset_button)
        buttons.addStretch(1)
        buttons.addWidget(self.generate_button)
        layout.addLayout(buttons)
        layout.addStretch(1)

    def choose_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select ETAP Short Circuit Study report", "", "PDF files (*.pdf)"
        )
        if path:
            self._pdf_path = path
            self.file_label.setText(os.path.basename(path))

    def reset(self) -> None:
        self._pdf_path = None
        self.file_label.setText("No file selected")
        self.equipment_check.setChecked(False)
        self.pseudo_check.setChecked(False)
        self.rule_combo.setCurrentText(sc_rules.DEFAULT_RULE)

    def generate(self) -> None:
        if not self._pdf_path:
            QMessageBox.warning(
                self, APP_NAME, "Please upload an ETAP Short Circuit Study report (.pdf) first."
            )
            return

        suggested = os.path.join(os.path.dirname(self._pdf_path), "Short Circuit Report.docx")
        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save report", suggested, "Word documents (*.docx)"
        )
        if not output_path:
            return
        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"
        self._output_path = output_path

        options = {
            "include_pseudo_buses": self.pseudo_check.isChecked(),
            "equipment_only": self.equipment_check.isChecked(),
            "rule": self.rule_combo.currentText(),
        }

        self.busy.emit(True)
        self.status_message.emit("Reading the ETAP report ...")

        self._thread = QThread(self)
        self._worker = ShortCircuitWorker(self._pdf_path, output_path, options)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _on_finished(self, result, excel_path: str) -> None:
        self.busy.emit(False)
        detail = f"  |  {result.detail}" if result.detail else ""
        self.status_message.emit(
            f"Report Generated Successfully  |  {result.count} buses  |  "
            f"{result.flagged} flagged{detail}"
        )
        QMessageBox.information(
            self, APP_NAME,
            f"Report Generated Successfully\n\nSaved to:\n{self._output_path}\n{excel_path}",
        )

    def _on_failed(self, message: str) -> None:
        self.busy.emit(False)
        self.status_message.emit(message)
        QMessageBox.warning(self, APP_NAME, message)


# --------------------------------------------------------------------------- #
# Main window
# --------------------------------------------------------------------------- #


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.resize(940, 860)
        self.setStyleSheet(STYLESHEET)

        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        title = QLabel("ETAP Report Filler")
        title.setObjectName("title")
        layout.addWidget(title)

        selector = card("Report Type")
        self.report_combo = QComboBox()
        for key in reports.ORDER:
            self.report_combo.addItem(reports.REPORTS[key].label, key)
        self.report_combo.currentIndexChanged.connect(self._on_report_changed)
        selector.layout().addWidget(self.report_combo)
        self.report_hint = hint("")
        selector.layout().addWidget(self.report_hint)
        layout.addWidget(selector)

        self.pages = QStackedWidget()
        self.load_flow_page = LoadFlowPage()
        self.short_circuit_page = ShortCircuitPage()
        for page in (self.load_flow_page, self.short_circuit_page):
            page.status_message.connect(self._set_status)
            page.busy.connect(self._set_busy)
            self.pages.addWidget(page)
        layout.addWidget(self.pages, 1)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(4)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.status = QLabel("")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        scroll = QScrollArea()
        scroll.setWidget(root)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(scroll)

        self._on_report_changed()

    def _on_report_changed(self) -> None:
        index = self.report_combo.currentIndex()
        self.pages.setCurrentIndex(index)
        self.report_hint.setText(reports.get(self.report_combo.currentData()).description)
        self.status.setText("")

    def _set_status(self, message: str) -> None:
        self.status.setText(message)

    def _set_busy(self, busy: bool) -> None:
        self.progress.setVisible(busy)
        self.report_combo.setEnabled(not busy)


def run() -> int:
    """Create the QApplication and show the main window."""
    app = QApplication.instance() or QApplication([])
    app.setApplicationName(APP_NAME)
    window = MainWindow()
    window.show()
    return app.exec()
