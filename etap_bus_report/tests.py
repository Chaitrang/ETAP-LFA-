"""
tests.py
--------
Regression tests.  Run with ``pytest tests.py`` (or ``python tests.py`` for a
quick pass/fail summary without pytest).

The PDF-based tests are skipped automatically if no sample report is present in
``samples/``.
"""

from __future__ import annotations

import os
import tempfile

import pdf_parser
import voltage_checker
import word_writer
from utils import BusRecord, fmt_nominal_voltage

SAMPLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples", "sample_lfa_report.pdf")
HAS_SAMPLE = os.path.exists(SAMPLE)


# --------------------------------------------------------------------------- #
# Voltage limits
# --------------------------------------------------------------------------- #


def test_limits_parsing():
    for text, expected in [
        ("95-106", (95.0, 106.0)),
        ("95 - 105 %", (95.0, 105.0)),
        ("90-110", (90.0, 110.0)),
        ("95%-106%", (95.0, 106.0)),
        ("95 to 106", (95.0, 106.0)),
        ("94.5–105.5", (94.5, 105.5)),
    ]:
        limits = voltage_checker.parse_limits(text)
        assert (limits.lower, limits.upper) == expected, text


def test_limits_rejects_bad_input():
    for text in ["95", "95%", "abc", "", "106-95", "95-", "-106"]:
        try:
            voltage_checker.parse_limits(text)
        except voltage_checker.LimitError as exc:
            assert str(exc) == voltage_checker.LIMIT_FORMAT_MESSAGE
        else:  # pragma: no cover
            raise AssertionError(f"{text!r} should have been rejected")


def test_remarks():
    limits = voltage_checker.parse_limits("95-106")
    assert voltage_checker.evaluate(95.0, limits) == "ACCEPTABLE"      # inclusive
    assert voltage_checker.evaluate(106.0, limits) == "ACCEPTABLE"     # inclusive
    assert voltage_checker.evaluate(94.9, limits) == "NOT ACCEPTABLE"
    assert voltage_checker.evaluate(107.8, limits) == "NOT ACCEPTABLE"
    assert voltage_checker.evaluate(None, limits) == ""


def test_nominal_formatting():
    assert fmt_nominal_voltage(132.0) == "132 kV"
    assert fmt_nominal_voltage(11.0) == "11 kV"
    assert fmt_nominal_voltage(3.3) == "3.3 kV"
    assert fmt_nominal_voltage(3.45) == "3.45 kV"
    assert fmt_nominal_voltage(0.415) == "415 V"
    assert fmt_nominal_voltage(0.38) == "380 V"


# --------------------------------------------------------------------------- #
# Word writer (no PDF needed)
# --------------------------------------------------------------------------- #


def test_word_writer_row_count():
    from docx import Document

    buses = [
        BusRecord(f"BUS-{i}", 11.0, 100.0 + i / 10, 1234.5, 678.9, "ACCEPTABLE")
        for i in range(120)  # more than the blank rows in the template
    ]
    with tempfile.TemporaryDirectory() as folder:
        path = word_writer.write_bus_report(buses, os.path.join(folder, "out.docx"))
        table = Document(path).tables[0]
        assert len(table.rows) == len(buses) + 1                     # + header
        assert table.rows[1].cells[0].text == "BUS-0"
        assert table.rows[-1].cells[0].text == "BUS-119"
        assert table.rows[1].cells[1].text == "11 kV"

    # fewer buses than blank template rows -> extra rows removed
    with tempfile.TemporaryDirectory() as folder:
        path = word_writer.write_bus_report(buses[:3], os.path.join(folder, "small.docx"))
        assert len(Document(path).tables[0].rows) == 4


# --------------------------------------------------------------------------- #
# Parser (needs the sample report)
# --------------------------------------------------------------------------- #


def test_parser_sample():
    if not HAS_SAMPLE:
        print("skipped: samples/sample_lfa_report.pdf not present")
        return

    buses = pdf_parser.parse_bus_data(SAMPLE, include_pseudo_buses=True)

    # every bus of the report, in report order, no duplicates
    assert len(buses) == 283, len(buses)
    assert len({b.bus_id for b in buses}) == len(buses)
    assert buses[0].bus_id == "132kV GRID"

    by_id = {b.bus_id: b for b in buses}

    # a plain row
    grid = by_id["132kV GRID"]
    assert grid.nominal_text == "132 kV"
    assert grid.voltage_text == "100.0"
    assert grid.amp_loading == 272.8

    # a switchboard: kW must come from the Bus Loading Summary, not from the
    # directly connected load (which is zero for this bus)
    swbd = by_id["S002-SB-1AC001 [BUS A]"]
    assert swbd.nominal_text == "11 kV"
    assert swbd.voltage_text == "101.2"
    assert abs(swbd.kw_loading - 31.398 * 0.910 * 1000) < 1.0
    assert swbd.amp_loading == 1627.6

    # bus IDs wrapped over two lines in the PDF must be reassembled
    assert "S002-SB-PSS01N01 [BUS A]" in by_id
    assert "S003-DB-11N001-03 [BUS B]" in by_id

    # bus IDs that overflow their column must keep every word
    assert by_id["TYPICAL MOTOR FEEDER 1 A"].amp_loading == 1069.5
    assert by_id["POWER SOCKET FEEDER 1 B"].amp_loading == 130.6

    # ETAP internal nodes are dropped by default
    filtered = pdf_parser.parse_bus_data(SAMPLE)
    assert all("~" not in b.bus_id for b in filtered)
    assert len(filtered) < len(buses)


def test_end_to_end():
    if not HAS_SAMPLE:
        print("skipped: samples/sample_lfa_report.pdf not present")
        return

    import cli

    with tempfile.TemporaryDirectory() as folder:
        output = os.path.join(folder, "Bus Report.docx")
        path, buses = cli.build_report(SAMPLE, "95-106", output)
        assert os.path.exists(path)
        assert all(b.remarks in {"ACCEPTABLE", "NOT ACCEPTABLE", ""} for b in buses)
        # 107.8 % buses must be flagged
        flagged = [b for b in buses if b.bus_id == "Bus2"]
        assert flagged and flagged[0].remarks == "NOT ACCEPTABLE"


def test_missing_file():
    try:
        pdf_parser.parse_bus_data("does_not_exist.pdf")
    except Exception as exc:
        assert str(exc) == pdf_parser.MSG_UNREADABLE
    else:  # pragma: no cover
        raise AssertionError("missing file should raise")


if __name__ == "__main__":
    failures = 0
    for name, func in sorted(globals().items()):
        if name.startswith("test_") and callable(func):
            try:
                func()
                print(f"PASS  {name}")
            except AssertionError as exc:
                failures += 1
                print(f"FAIL  {name}: {exc}")
    print("\nAll tests passed." if not failures else f"\n{failures} test(s) failed.")
    raise SystemExit(1 if failures else 0)
