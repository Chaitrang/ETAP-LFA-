"""
main.py
Headless pipeline: PDF -> extracted bus data -> Remarks -> populated Word
report. This is called directly by ui.py (PySide6 GUI) and can also be run
standalone from the command line, which is useful for automation/testing
without a display:

    python main.py --pdf report.pdf --limits 95-106 --out "Bus Report.docx"

Exit codes: 0 success, 1 bad usage/limits, 2 PDF read error, 3 no bus data.
"""
from __future__ import annotations

import argparse
import os
import sys

from pdf_parser import extract_bus_data, NoBusOutputDataError, UnreadablePdfError
from voltage_checker import evaluate_remark
from word_writer import populate_template
from utils import parse_voltage_limits

def _default_template_path() -> str:
    """Resolve the bundled template both when running from source and when
    frozen into a PyInstaller executable (which extracts data files under
    sys._MEIPASS at runtime)."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "assets", "Bus_Template.docx")


DEFAULT_TEMPLATE = _default_template_path()

ERR_NO_BUS_DATA = "No Bus Output Data found in the uploaded ETAP report."
ERR_BAD_LIMITS = "Please enter voltage limits in the format Lower-Upper (Example: 95-106)."
ERR_UNREADABLE_PDF = "Unable to read the ETAP report."


class ReportGenerationError(Exception):
    """Wraps the standard user-facing error messages defined in the spec."""


def generate_report(pdf_path: str, limits_text: str, output_path: str,
                     template_path: str = DEFAULT_TEMPLATE) -> str:
    """
    Runs the full pipeline and writes the populated Word report to
    output_path. Raises ReportGenerationError with one of the three
    spec-defined messages on failure; callers (CLI or GUI) should display
    str(exc) to the user as-is.
    """
    limits = parse_voltage_limits(limits_text)
    if limits is None:
        raise ReportGenerationError(ERR_BAD_LIMITS)

    try:
        records = extract_bus_data(pdf_path)
    except UnreadablePdfError:
        raise ReportGenerationError(ERR_UNREADABLE_PDF)
    except NoBusOutputDataError:
        raise ReportGenerationError(ERR_NO_BUS_DATA)

    if not records:
        raise ReportGenerationError(ERR_NO_BUS_DATA)

    rows = []
    for rec in records:
        remark = evaluate_remark(rec.voltage_pct, limits)
        rows.append({
            "bus_id": rec.bus_id,
            "nominal": rec.nominal_display,
            "voltage_pct": rec.voltage_pct,
            "kw_loading": f"{rec.kw_loading:.2f}",
            "amp_loading": f"{rec.amp_loading:.1f}",
            "remarks": remark,
        })

    populate_template(template_path, output_path, rows)
    return output_path


def _cli():
    parser = argparse.ArgumentParser(
        description="Extract Bus Load Flow data from an ETAP PDF report and populate a Word report."
    )
    parser.add_argument("--pdf", required=True, help="Path to the ETAP Load Flow Analysis PDF report.")
    parser.add_argument("--limits", required=True, help='Acceptable voltage limits, e.g. "95-106".')
    parser.add_argument("--out", default="Bus Report.docx", help="Output .docx path.")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE, help="Path to the Word template (advanced use).")
    args = parser.parse_args()

    try:
        out = generate_report(args.pdf, args.limits, args.out, args.template)
    except ReportGenerationError as exc:
        print(str(exc), file=sys.stderr)
        if str(exc) == ERR_BAD_LIMITS:
            sys.exit(1)
        elif str(exc) == ERR_UNREADABLE_PDF:
            sys.exit(2)
        else:
            sys.exit(3)
    else:
        print(f"Report Generated Successfully: {out}")
        sys.exit(0)


if __name__ == "__main__":
    _cli()
