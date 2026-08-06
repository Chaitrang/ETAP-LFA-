"""
cli.py
------
Command line front end - same pipeline as the GUI, useful for batch runs,
regression testing and CI.

    python cli.py report.pdf --limits 95-106 --output "Bus Report.docx"
"""

from __future__ import annotations

import argparse
import sys
import time

import pdf_parser
import voltage_checker
import word_writer
from utils import APP_NAME, APP_VERSION, ParserError


def build_report(pdf_path: str, limits_text: str, output_path: str, include_pseudo: bool = False):
    """Run the full pipeline and return ``(output_path, buses)``."""
    limits = voltage_checker.parse_limits(limits_text)
    buses = pdf_parser.load_report(pdf_path, include_pseudo_buses=include_pseudo)
    voltage_checker.apply_limits(buses, limits)
    word_writer.write_bus_report(buses, output_path)
    return output_path, buses


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=f"{APP_NAME} {APP_VERSION}")
    parser.add_argument("pdf", help="ETAP Load Flow Analysis report (PDF)")
    parser.add_argument("-l", "--limits", default="95-106", help="Acceptable voltage limits, e.g. 95-106")
    parser.add_argument("-o", "--output", default="Bus Report.docx", help="Output .docx path")
    parser.add_argument(
        "--include-pseudo-buses",
        action="store_true",
        help="Also list ETAP auto-generated internal nodes (IDs containing '~')",
    )
    args = parser.parse_args(argv)

    started = time.perf_counter()
    try:
        output, buses = build_report(args.pdf, args.limits, args.output, args.include_pseudo_buses)
    except voltage_checker.LimitError as exc:
        print(exc, file=sys.stderr)
        return 2
    except ParserError as exc:
        print(exc, file=sys.stderr)
        return 3

    elapsed = time.perf_counter() - started
    not_ok = sum(1 for b in buses if b.remarks == "NOT ACCEPTABLE")
    print(f"Report Generated Successfully -> {output}")
    print(f"{len(buses)} buses, {not_ok} outside limits, {elapsed:.2f} s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
