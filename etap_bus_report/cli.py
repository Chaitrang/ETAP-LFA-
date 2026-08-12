"""
cli.py
------
Command line front end - same pipeline as the GUI and the web app.

    python cli.py report.pdf --type load_flow --limits 95-106 -o "Bus Report.docx"
    python cli.py sc.pdf --type short_circuit -o "Short Circuit Report.docx"
    python cli.py --show-mapping short_circuit
"""

from __future__ import annotations

import argparse
import os
import sys
import time

import excel_writer
import loadflow_config
import reports
import shortcircuit_config
import shortcircuit_fields
import shortcircuit_writer
import voltage_checker
from utils import APP_NAME, APP_VERSION, ParserError


def build_report(pdf_path: str, output_path: str, report_type: str = reports.LOAD_FLOW, **options):
    """
    Run one report end to end; returns ``(output_path, ReportResult)``.

    Writes the Word document and, beside it, the matching ``.xlsx`` workbook
    built from the same rows.
    """
    result = reports.generate(report_type, pdf_path, **options)
    folder = os.path.dirname(os.path.abspath(output_path))
    if folder:
        os.makedirs(folder, exist_ok=True)
    if result.wants_word and result.document_bytes:
        with open(output_path, "wb") as handle:
            handle.write(result.document_bytes)
    if result.wants_excel and result.excel_bytes:
        excel_writer.write_excel(result.excel_bytes, excel_writer.excel_path_for(output_path))
    return output_path, result


def _load_flow_config(args):
    """
    Build a configuration from the Load Flow flags.

    Returns ``None`` when no flag was given, so the legacy ``--limits``
    behaviour is preserved untouched.
    """
    flags = [args.bus_info, args.results, args.format]
    limits = {
        name: {kind: getattr(args, f"{name}_{kind}") for kind in ("critical", "marginal")}
        for name in ("loading", "overvoltage", "undervoltage")
    }
    if not any(flags) and not any(v is not None for row in limits.values() for v in row.values()):
        return None

    config = loadflow_config.from_legacy_limits(args.limits)
    if args.bus_info is not None:
        config.bus_info = tuple(f.strip() for f in args.bus_info.split(",") if f.strip())
    if args.results is not None:
        config.results = tuple(f.strip() for f in args.results.split(",") if f.strip())
    if args.format:
        config.output_format = args.format

    for name, values in limits.items():
        limit = getattr(config, name)
        for kind, value in values.items():
            if value is None:
                continue
            if value == 0:
                setattr(limit, f"{kind}_enabled", False)
            else:
                setattr(limit, kind, value)
                setattr(limit, f"{kind}_enabled", True)
        # once any limit is given explicitly, use the richer wording
        config.remark_style = loadflow_config.STATUS_STYLE
    return config


def _short_circuit_config(args):
    """
    Build a Short Circuit configuration from the flags.

    Returns ``None`` when no new flag was given, so the original
    template-filling behaviour is preserved untouched.
    """
    given = [args.sc_info, args.sc_results, args.device_type, args.critical,
             args.marginal, args.current_unit, args.voltage_unit]
    if not any(v is not None for v in given) and not (args.skip_non_alerted or args.keep_nodes
                                                     or args.format):
        return None

    config = shortcircuit_config.ShortCircuitConfig()
    if args.sc_info is not None:
        config.info = tuple(f.strip() for f in args.sc_info.split(",") if f.strip())
    if args.sc_results is not None:
        config.results = tuple(f.strip() for f in args.sc_results.split(",") if f.strip())
    if args.device_type:
        config.equipment_types = tuple(t.strip() for t in args.device_type.split(",") if t.strip())
    if args.critical is not None:
        config.critical, config.critical_enabled = args.critical, args.critical > 0
    if args.marginal is not None:
        config.marginal, config.marginal_enabled = args.marginal, args.marginal > 0
    if args.current_unit:
        config.current_unit = args.current_unit
    if args.voltage_unit:
        config.voltage_unit = args.voltage_unit
    if args.format:
        config.output_format = args.format
    config.skip_non_alerted = args.skip_non_alerted
    config.skip_nodes = not args.keep_nodes
    if args.equipment_only and not config.equipment_types:
        config.equipment_types = tuple(
            t for t in shortcircuit_fields.EQUIPMENT_TYPES if t != "Bus"
        )
    return config


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=f"{APP_NAME} {APP_VERSION}")
    parser.add_argument("pdf", nargs="?", help="ETAP report (PDF)")
    parser.add_argument(
        "-t", "--type", default=reports.LOAD_FLOW, choices=list(reports.ORDER),
        help="Report type to process",
    )
    parser.add_argument("-l", "--limits", default="95-106", help="Load Flow only: acceptable voltage limits")
    parser.add_argument("-o", "--output", help="Output .docx path")
    parser.add_argument(
        "--rule", default=None,
        help="Short Circuit only: Remarks rule (peak_vs_rating, peak_and_breaking, etap_flag, blank)",
    )
    parser.add_argument(
        "--equipment-only", action="store_true",
        help="Short Circuit only: list switchgear/switchboards/MCCs only, not every bus",
    )
    parser.add_argument(
        "--sc-info", default=None,
        help="Short Circuit only: comma separated Info columns "
             "(nominal_kv, bus_type, cfactor, rating_peak, rating_ib_sym, "
             "governing_device, xr_ratio)",
    )
    parser.add_argument(
        "--sc-results", default=None,
        help="Short Circuit only: comma separated Results columns "
             "(ik_initial, ip_peak, ik_steady, standard)",
    )
    parser.add_argument("--critical", type=float, default=None, metavar="PCT",
                        help="Short Circuit only: critical threshold %% of the rating")
    parser.add_argument("--marginal", type=float, default=None, metavar="PCT",
                        help="Short Circuit only: marginal threshold %% of the rating")
    parser.add_argument("--device-type", default=None,
                        help="Short Circuit only: comma separated equipment types "
                             "(Bus, Switchgear, Switchboard, Mcc, Switchrack, Panelboard)")
    parser.add_argument("--skip-non-alerted", action="store_true",
                        help="Short Circuit only: keep only rows reaching a threshold")
    parser.add_argument("--keep-nodes", action="store_true",
                        help="Short Circuit only: keep ETAP's internal '~' nodes")
    parser.add_argument("--current-unit", default=None, choices=list(shortcircuit_fields.CURRENT_UNITS),
                        help="Short Circuit only: current unit")
    parser.add_argument("--voltage-unit", default=None, choices=list(shortcircuit_fields.VOLTAGE_UNITS),
                        help="Short Circuit only: voltage unit")
    parser.add_argument(
        "--include-pseudo-buses", action="store_true",
        help="Also list ETAP auto-generated internal nodes (IDs containing '~')",
    )
    parser.add_argument(
        "--show-mapping", metavar="TYPE", nargs="?", const=reports.SHORT_CIRCUIT,
        help="Print how the bundled template's columns were mapped, then exit",
    )
    # -- Load Flow: configurable columns, limits and format ------------------ #
    parser.add_argument(
        "--bus-info", default=None,
        help="Load Flow only: comma separated Bus Info columns "
             "(nominal_kv, rated_amp). Default: nominal_kv",
    )
    parser.add_argument(
        "--results", default=None,
        help="Load Flow only: comma separated result columns (voltage, kw_loading, "
             "kvar_loading, apparent_loading, amp_loading, percent_loading)",
    )
    parser.add_argument(
        "--format", default=None, choices=list(loadflow_config.OUTPUT_FORMATS),
        help="Load Flow only: which documents to write (default: Word + Excel)",
    )
    for name in ("loading", "overvoltage", "undervoltage"):
        for kind in ("critical", "marginal"):
            parser.add_argument(
                f"--{name}-{kind}", type=float, default=None, metavar="PCT",
                help=f"Load Flow only: {name} {kind} limit in %% (omit to keep the default, "
                     f"pass 0 to switch it off)",
            )
    args = parser.parse_args(argv)

    if args.show_mapping:
        for header, key in shortcircuit_writer.describe_mapping():
            print(f"{header:45s} -> {key or '(left blank)'}")
        return 0

    if not args.pdf:
        parser.error("a PDF is required")

    report = reports.get(args.type)
    output = args.output or report.default_filename

    options = {"include_pseudo_buses": args.include_pseudo_buses}
    if args.type == reports.LOAD_FLOW:
        options["limits"] = args.limits
        config = _load_flow_config(args)
        if config is not None:
            options["config"] = config
    else:
        options["rule"] = args.rule
        options["equipment_only"] = args.equipment_only
        config = _short_circuit_config(args)
        if config is not None:
            options["config"] = config
            options["dynamic"] = True

    started = time.perf_counter()
    try:
        path, result = build_report(args.pdf, output, args.type, **options)
    except (voltage_checker.LimitError, loadflow_config.ConfigError,
            shortcircuit_config.ConfigError) as exc:
        print(exc, file=sys.stderr)
        return 2
    except ParserError as exc:
        print(exc, file=sys.stderr)
        return 3

    print("Report Generated Successfully")
    if result.wants_word and result.document_bytes:
        print(f"  -> {path}")
    if result.wants_excel and result.excel_bytes:
        print(f"  -> {excel_writer.excel_path_for(path)}")
    print(
        f"{result.count} buses, {result.flagged} flagged, "
        f"{time.perf_counter() - started:.2f} s"
    )
    if result.detail:
        print(result.detail)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
