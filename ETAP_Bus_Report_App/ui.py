"""
ui.py
Minimal PySide6 desktop UI, per spec:
    Input 1: Upload ETAP Load Flow Analysis Report (.pdf)   [button]
    Input 2: Acceptable Voltage Limits                       [text field]
    Buttons: Generate Report, Reset (optional)
    Output:  "Report Generated Successfully" + save dialog

All processing logic lives in main.py / pdf_parser.py / word_writer.py /
voltage_checker.py - this module is presentation only, so the pipeline can
also be exercised headlessly (see main.py's CLI) without Qt installed.
"""
from __future__ import annotations

import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from main import generate_report, ReportGenerationError, DEFAULT_TEMPLATE

APP_TITLE = "ETAP Bus Load Flow Report Generator"
DEFAULT_OUTPUT_NAME = "Bus Report.docx"


class GenerateWorker(QThread):
    """Runs the (potentially several-second) parsing/writing pipeline off
    the UI thread so the window never freezes/greys out while processing."""

    finished_ok = Signal(str)
    finished_err = Signal(str)

    def __init__(self, pdf_path: str, limits_text: str, output_path: str):
        super().__init__()
        self.pdf_path = pdf_path
        self.limits_text = limits_text
        self.output_path = output_path

    def run(self):
        try:
            out = generate_report(self.pdf_path, self.limits_text, self.output_path)
        except ReportGenerationError as exc:
            self.finished_err.emit(str(exc))
        except Exception:  # noqa: BLE001 - surface unexpected errors too
            self.finished_err.emit(
                "An unexpected error occurred while generating the report:\n"
                + traceback.format_exc(limit=3)
            )
        else:
            self.finished_ok.emit(out)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setMinimumWidth(480)

        self.pdf_path: str | None = None
        self.worker: GenerateWorker | None = None

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(18)

        title = QLabel(APP_TITLE)
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # --- Input 1: PDF upload -------------------------------------- #
        pdf_row = QVBoxLayout()
        pdf_label = QLabel("Upload ETAP Load Flow Analysis Report (.pdf)")
        pdf_row.addWidget(pdf_label)

        pdf_btn_row = QHBoxLayout()
        self.pdf_button = QPushButton("Choose PDF…")
        self.pdf_button.clicked.connect(self._choose_pdf)
        self.pdf_file_label = QLabel("No file selected")
        self.pdf_file_label.setStyleSheet("color: #666;")
        pdf_btn_row.addWidget(self.pdf_button)
        pdf_btn_row.addWidget(self.pdf_file_label, stretch=1)
        pdf_row.addLayout(pdf_btn_row)
        layout.addLayout(pdf_row)

        # --- Input 2: Voltage limits ------------------------------------ #
        limits_row = QVBoxLayout()
        limits_label = QLabel("Acceptable Voltage Limits")
        limits_row.addWidget(limits_label)
        self.limits_input = QLineEdit()
        self.limits_input.setPlaceholderText("e.g. 95-106  or  95 - 105 %  or  90-110")
        limits_row.addWidget(self.limits_input)
        layout.addLayout(limits_row)

        # --- Buttons ------------------------------------------------------ #
        button_row = QHBoxLayout()
        self.generate_button = QPushButton("Generate Report")
        self.generate_button.setDefault(True)
        self.generate_button.clicked.connect(self._on_generate)
        self.reset_button = QPushButton("Reset")
        self.reset_button.clicked.connect(self._on_reset)
        button_row.addWidget(self.generate_button)
        button_row.addWidget(self.reset_button)
        layout.addLayout(button_row)

        # --- Status line ---------------------------------------------- #
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666;")
        layout.addWidget(self.status_label)

        layout.addStretch(1)

    # ------------------------------------------------------------------ #
    def _choose_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select ETAP Load Flow Analysis Report", "", "PDF Files (*.pdf)"
        )
        if path:
            self.pdf_path = path
            self.pdf_file_label.setText(os.path.basename(path))
            self.pdf_file_label.setStyleSheet("color: #222;")

    def _on_reset(self):
        self.pdf_path = None
        self.pdf_file_label.setText("No file selected")
        self.pdf_file_label.setStyleSheet("color: #666;")
        self.limits_input.clear()
        self.status_label.setText("")

    def _on_generate(self):
        if not self.pdf_path:
            QMessageBox.warning(self, APP_TITLE, "Please upload an ETAP PDF report first.")
            return

        limits_text = self.limits_input.text().strip()
        # Fast client-side check for empty field; format validation itself
        # happens in utils.parse_voltage_limits() inside the pipeline so
        # the exact spec error message is shown consistently either way.
        if not limits_text:
            QMessageBox.warning(
                self, APP_TITLE,
                "Please enter voltage limits in the format Lower-Upper (Example: 95-106).",
            )
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Bus Report As", DEFAULT_OUTPUT_NAME, "Word Document (*.docx)"
        )
        if not save_path:
            return  # user cancelled the save dialog
        if not save_path.lower().endswith(".docx"):
            save_path += ".docx"

        self.generate_button.setEnabled(False)
        self.status_label.setText("Processing… this may take a few seconds.")
        self.worker = GenerateWorker(self.pdf_path, limits_text, save_path)
        self.worker.finished_ok.connect(self._on_success)
        self.worker.finished_err.connect(self._on_error)
        self.worker.start()

    def _on_success(self, output_path: str):
        self.generate_button.setEnabled(True)
        self.status_label.setText("Report Generated Successfully")
        QMessageBox.information(
            self, APP_TITLE, f"Report Generated Successfully\n\nSaved to:\n{output_path}"
        )

    def _on_error(self, message: str):
        self.generate_button.setEnabled(True)
        self.status_label.setText("")
        QMessageBox.critical(self, APP_TITLE, message)


def main():
    if not os.path.exists(DEFAULT_TEMPLATE):
        print(f"Bundled template not found at {DEFAULT_TEMPLATE}", file=sys.stderr)
        sys.exit(1)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
