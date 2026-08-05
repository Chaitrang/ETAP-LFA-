"""
tests/test_pipeline.py
Lightweight smoke tests - runs with plain `python tests/test_pipeline.py`,
no pytest dependency required. Exercises:
  - utils: nominal voltage formatting, voltage-limit parsing
  - voltage_checker: boundary conditions
  - pdf_parser + word_writer + main: full pipeline against the sample PDF,
    plus the three spec-defined error paths.

Run from the project root:
    python tests/test_pipeline.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils import format_nominal_voltage, parse_voltage_limits
from voltage_checker import evaluate_remark, ACCEPTABLE, NOT_ACCEPTABLE

SAMPLE_PDF = "/mnt/user-data/uploads/Table_7_LFA_CASE_1.pdf"
TEMPLATE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "Bus_Template.docx")

passed = 0
failed = 0


def check(label, condition):
    global passed, failed
    if condition:
        passed += 1
    else:
        failed += 1
        print(f"FAIL: {label}")


def test_format_nominal_voltage():
    check("132kV bus", format_nominal_voltage("132.000") == "132 kV")
    check("11kV bus", format_nominal_voltage("11.000") == "11 kV")
    check("3.3kV bus", format_nominal_voltage("3.300") == "3.3 kV")
    check("415V bus", format_nominal_voltage("0.415") == "415 V")
    check("380V bus", format_nominal_voltage("0.380") == "380 V")
    check("11.5kV bus", format_nominal_voltage("11.500") == "11.5 kV")


def test_parse_voltage_limits():
    for text, expect in [
        ("95-106", (95.0, 106.0)),
        ("95 - 105 %", (95.0, 105.0)),
        ("90-110", (90.0, 110.0)),
        ("95%-106%", (95.0, 106.0)),
        ("95 to 106", (95.0, 106.0)),
        ("106-95", (95.0, 106.0)),  # reversed input still normalizes
    ]:
        r = parse_voltage_limits(text)
        check(f"limits '{text}'", r is not None and (r.lower, r.upper) == expect)

    for bad in ["95", "95%", "abc", "", None]:
        check(f"invalid limits {bad!r} rejected", parse_voltage_limits(bad) is None)


def test_evaluate_remark():
    limits = parse_voltage_limits("95-106")
    check("mid-range accepted", evaluate_remark(100.0, limits) == ACCEPTABLE)
    check("lower boundary accepted", evaluate_remark(95.0, limits) == ACCEPTABLE)
    check("upper boundary accepted", evaluate_remark(106.0, limits) == ACCEPTABLE)
    check("just below rejected", evaluate_remark(94.9, limits) == NOT_ACCEPTABLE)
    check("just above rejected", evaluate_remark(106.1, limits) == NOT_ACCEPTABLE)


def test_full_pipeline_success():
    if not os.path.exists(SAMPLE_PDF):
        print("SKIP: sample PDF not present in this environment")
        return
    from main import generate_report
    out_path = "/tmp/test_bus_report.docx"
    result = generate_report(SAMPLE_PDF, "95-106", out_path, TEMPLATE)
    check("output file created", os.path.exists(result))

    import docx
    d = docx.Document(result)
    t = d.tables[0]
    check("header row present", t.rows[0].cells[0].text == "Bus ID")
    check("283 buses + header row", len(t.rows) == 284)
    check("first data row is 132kV GRID", t.rows[1].cells[0].text == "132kV GRID")
    check("remarks column populated", t.rows[1].cells[5].text in (ACCEPTABLE, NOT_ACCEPTABLE))


def test_error_paths():
    from main import generate_report, ReportGenerationError

    # bad voltage limits
    try:
        generate_report(SAMPLE_PDF, "abc", "/tmp/x.docx", TEMPLATE)
        check("bad limits raises", False)
    except ReportGenerationError as e:
        check("bad limits message", "Lower-Upper" in str(e))

    # unreadable pdf
    try:
        generate_report("/nonexistent/file.pdf", "95-106", "/tmp/x.docx", TEMPLATE)
        check("unreadable pdf raises", False)
    except ReportGenerationError as e:
        check("unreadable pdf message", "Unable to read" in str(e))


if __name__ == "__main__":
    test_format_nominal_voltage()
    test_parse_voltage_limits()
    test_evaluate_remark()
    test_full_pipeline_success()
    test_error_paths()
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)
