"""
ui.py
-----
Minimal PySide6 interface: one file picker, one limits field, two buttons.

All heavy work (PDF parsing, Word writing) runs in a worker thread so the
window never freezes, even on a 300-bus report.
"""

from __future__ import annotations

import os

from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

import pdf_parser
import voltage_checker
import word_writer
from utils import APP_NAME, APP_VERSION, ParserError

DEFAULT_OUTPUT_NAME = "Bus Report.docx"
DEFAULT_LIMITS = "95-106"

STYLESHEET = """
QWidget       { background: #ffffff; color: #1f2430; font-size: 14px; }
QLabel#title  { font-size: 20px; font-weight: 600; color: #524689; }
QLabel#hint   { color: #6b7280; font-size: 12px; }
QLabel#status { font-size: 13px; }
QLineEdit     { border: 1px solid #d0d5dd; border-radius: 6px; padding: 8px 10px; }
QLineEdit:focus { border: 1px solid #524689; }
QPushButton   { border: 1px solid #d0d5dd; border-radius: 6px; padding: 9px 16px; background: #f7f7fb; }
QPushButton:hover { background: #eeeef6; }
QPushButton#primary { background: #524689; color: #ffffff; border: none; font-weight: 600; }
QPushButton#primary:hover   { background: #453a76; }
QPushButton#primary:disabled{ background: #b6b1cf; }
QFrame#card   { border: 1px solid #e5e7eb; border-radius: 10px; }
"""


# --------------------------------------------------------------------------- #
# Worker
# --------------------------------------------------------------------------- #


class ReportWorker(QObject):
    """Runs the extraction pipeline off the GUI thread."""

    finished = Signal(str, int, int)   # output path, bus count, out-of-limit count
    failed = Signal(str)               # message to show the user

    def __init__(self, pdf_path: str, limits_text: str, output_path: str) -> None:
        super().__init__()
        self._pdf_path = pdf_path
        self._limits_text = limits_text
        self._output_path = output_path

    def run(self) -> None:
        try:
            limits = voltage_checker.parse_limits(self._limits_text)
            buses = pdf_parser.load_report(self._pdf_path)
            voltage_checker.apply_limits(buses, limits)
            word_writer.write_bus_report(buses, self._output_path)
        except voltage_checker.LimitError as exc:
            self.failed.emit(str(exc))
        except ParserError as exc:
            self.failed.emit(str(exc))
        except PermissionError:
            self.failed.emit(
                "Unable to save the report. Close 'Bus Report.docx' if it is open in Word and try again."
            )
        except Exception:  # pragma: no cover - unexpected failure
            self.failed.emit(pdf_parser.MSG_UNREADABLE)
        else:
            not_ok = sum(1 for bus in buses if bus.remarks == "NOT ACCEPTABLE")
            self.finished.emit(self._output_path, len(buses), not_ok)


# --------------------------------------------------------------------------- #
# Main window
# --------------------------------------------------------------------------- #


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumWidth(560)
        self.setStyleSheet(STYLESHEET)

        self._pdf_path: str | None = None
        self._thread: QThread | None = None
        self._worker: ReportWorker | None = None

        self.setCentralWidget(self._build_ui())

    # -- construction ------------------------------------------------------- #

    def _build_ui(self) -> QWidget:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(18)

        title = QLabel("ETAP Bus Load Flow Report")
        title.setObjectName("title")
        layout.addWidget(title)

        subtitle = QLabel("Extracts the bus results from an ETAP Load Flow report into the standard Bus table.")
        subtitle.setObjectName("hint")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        layout.addWidget(self._build_file_row())
        layout.addWidget(self._build_limits_row())
        layout.addLayout(self._build_buttons())

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)          # indeterminate
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(4)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.status = QLabel("")
        self.status.setObjectName("status")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        layout.addStretch(1)
        return root

    def _build_file_row(self) -> QWidget:
        card = QFrame()
        card.setObjectName("card")
        box = QVBoxLayout(card)
        box.setContentsMargins(16, 14, 16, 14)
        box.setSpacing(8)

        label = QLabel("1.  ETAP Load Flow Analysis Report (.pdf)")
        label.setFont(QFont(self.font().family(), self.font().pointSize(), QFont.DemiBold))
        box.addWidget(label)

        row = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("hint")
        self.file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        browse = QPushButton("Upload PDF")
        browse.clicked.connect(self.choose_pdf)
        row.addWidget(self.file_label, 1)
        row.addWidget(browse, 0)
        box.addLayout(row)
        return card

    def _build_limits_row(self) -> QWidget:
        card = QFrame()
        card.setObjectName("card")
        box = QVBoxLayout(card)
        box.setContentsMargins(16, 14, 16, 14)
        box.setSpacing(8)

        label = QLabel("2.  Acceptable Voltage Limits")
        label.setFont(QFont(self.font().family(), self.font().pointSize(), QFont.DemiBold))
        box.addWidget(label)

        self.limits_edit = QLineEdit(DEFAULT_LIMITS)
        self.limits_edit.setPlaceholderText("95-106")
        self.limits_edit.returnPressed.connect(self.generate)
        box.addWidget(self.limits_edit)

        hint = QLabel("Examples:  95-106    95 - 105 %    90-110")
        hint.setObjectName("hint")
        box.addWidget(hint)
        return card

    def _build_buttons(self) -> QHBoxLayout:
        row = QHBoxLayout()
        self.reset_button = QPushButton("Reset")
        self.reset_button.clicked.connect(self.reset)

        self.generate_button = QPushButton("Generate Report")
        self.generate_button.setObjectName("primary")
        self.generate_button.setDefault(True)
        self.generate_button.clicked.connect(self.generate)

        row.addWidget(self.reset_button, 0)
        row.addStretch(1)
        row.addWidget(self.generate_button, 0)
        return row

    # -- actions ------------------------------------------------------------ #

    def choose_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select ETAP Load Flow Analysis report", "", "PDF files (*.pdf)"
        )
        if path:
            self._pdf_path = path
            self.file_label.setText(os.path.basename(path))
            self.status.setText("")

    def reset(self) -> None:
        self._pdf_path = None
        self.file_label.setText("No file selected")
        self.limits_edit.setText(DEFAULT_LIMITS)
        self.status.setText("")

    def generate(self) -> None:
        if not self._pdf_path:
            self._warn("Please upload an ETAP Load Flow Analysis report (.pdf) first.")
            return

        try:
            voltage_checker.parse_limits(self.limits_edit.text())
        except voltage_checker.LimitError as exc:
            self._warn(str(exc))
            return

        suggested = os.path.join(os.path.dirname(self._pdf_path), DEFAULT_OUTPUT_NAME)
        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save Bus Report", suggested, "Word documents (*.docx)"
        )
        if not output_path:
            return
        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        self._set_busy(True)
        self.status.setText("Reading the ETAP report ...")

        self._thread = QThread(self)
        self._worker = ReportWorker(self._pdf_path, self.limits_edit.text(), output_path)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    # -- worker callbacks --------------------------------------------------- #

    def _on_finished(self, output_path: str, count: int, not_ok: int) -> None:
        self._set_busy(False)
        self.status.setText(
            f"Report Generated Successfully\n{count} buses written  •  "
            f"{not_ok} outside the acceptable limits\n{output_path}"
        )
        QMessageBox.information(
            self, APP_NAME, f"Report Generated Successfully\n\nSaved to:\n{output_path}"
        )

    def _on_failed(self, message: str) -> None:
        self._set_busy(False)
        self.status.setText(message)
        self._warn(message)

    # -- helpers ------------------------------------------------------------ #

    def _set_busy(self, busy: bool) -> None:
        self.generate_button.setEnabled(not busy)
        self.reset_button.setEnabled(not busy)
        self.progress.setVisible(busy)

    def _warn(self, message: str) -> None:
        QMessageBox.warning(self, APP_NAME, message)


def run() -> int:
    """Create the QApplication and show the main window."""
    app = QApplication.instance() or QApplication([])
    app.setApplicationName(APP_NAME)
    window = MainWindow()
    window.show()
    return app.exec()
