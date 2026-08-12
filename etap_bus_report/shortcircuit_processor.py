"""
shortcircuit_processor.py
-------------------------
Turns the parsed Short Circuit dataset plus the user's configuration into
**one** final table, which the preview, the Word document and the Excel
workbook are all rendered from.

The PDF is never re-read here: the input is the list of records the existing
:mod:`shortcircuit_parser` already produced.

Classification
==============
The pre-existing rule (:func:`sc_rules.peak_vs_rating`) was: the equipment's
making (peak) capacity must cover the calculated peak current ``ip``.  That is
preserved exactly, with the comparison expressed as a percentage so the
thresholds can be configured:

    duty % = ip / Rated ip x 100

    duty % >= Critical  -> CRITICAL
    duty % >= Marginal  -> MARGINAL
    otherwise           -> ACCEPTABLE

With Critical 100 % this is identical to the old ``ip <= rating`` test, and it
matches ETAP's own ``*`` and ``#`` footnotes. A bus with no equipment rating
cannot be assessed, so it keeps a blank remark - never a guess.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Sequence

import shortcircuit_fields as fields
from shortcircuit_config import LEGACY_STYLE, ShortCircuitConfig
from utils import REMARK_ACCEPTABLE, REMARK_NOT_ACCEPTABLE, fmt_number

ACCEPTABLE = "ACCEPTABLE"
MARGINAL = "MARGINAL"
CRITICAL = "CRITICAL"


@dataclass
class ShortCircuitTable:
    """The final dataset: headers, rows, and what could not be filled."""

    headers: list[str]
    rows: list[list[str]]
    columns: list[str]
    status_counts: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    #: Rows removed by the filters, for an honest summary.
    filtered_out: dict = field(default_factory=dict)

    @property
    def count(self) -> int:
        return len(self.rows)

    @property
    def flagged(self) -> int:
        return sum(
            count for status, count in self.status_counts.items()
            if status not in (ACCEPTABLE, "")
        )


class EmptyTableError(ValueError):
    """Raised when the filters leave nothing to report."""


# --------------------------------------------------------------------------- #
# Classification
# --------------------------------------------------------------------------- #


def duty_percent(row: dict) -> Optional[float]:
    """Peak duty as a percentage of the equipment making capacity."""
    rating = row.get("rating_peak")
    peak = row.get("ip_peak")
    if not rating or peak is None:
        return None
    return peak / rating * 100.0


def classify(row: dict, config: ShortCircuitConfig) -> str:
    """ACCEPTABLE / MARGINAL / CRITICAL, or "" when the bus has no rating."""
    duty = duty_percent(row)
    if duty is None:
        return ""
    if config.critical_value is not None and duty >= config.critical_value:
        return CRITICAL
    if config.marginal_value is not None and duty >= config.marginal_value:
        return MARGINAL
    return ACCEPTABLE


def remark_text(status: str, config: ShortCircuitConfig) -> str:
    if not status:
        return ""
    if config.remark_style == LEGACY_STYLE:
        return REMARK_ACCEPTABLE if status == ACCEPTABLE else REMARK_NOT_ACCEPTABLE
    return status


# --------------------------------------------------------------------------- #
# Filtering
# --------------------------------------------------------------------------- #


def _is_node(row: dict) -> bool:
    """ETAP's auto-generated internal nodes carry a '~' in their ID."""
    return "~" in row.get("bus_id", "")


def apply_filters(records: Sequence[dict], config: ShortCircuitConfig):
    """
    Apply the Device Type / Skip Nodes / Skip Non-Alerted filters.

    Returns ``(kept_rows, removed_counts)``; the counts let the interface say
    exactly why rows disappeared instead of silently shrinking the report.
    """
    removed = {"nodes": 0, "equipment type": 0, "not alerted": 0}
    kept: list[dict] = []

    for record in records:
        row = dict(record)

        if config.skip_nodes and _is_node(row):
            removed["nodes"] += 1
            continue

        if config.equipment_types:
            if (row.get("bus_type") or "") not in config.equipment_types:
                removed["equipment type"] += 1
                continue

        row["status"] = classify(row, config)
        row["duty_percent"] = duty_percent(row)

        if config.skip_non_alerted and row["status"] in ("", ACCEPTABLE):
            removed["not alerted"] += 1
            continue

        kept.append(row)

    return kept, {k: v for k, v in removed.items() if v}


# --------------------------------------------------------------------------- #
# Table building
# --------------------------------------------------------------------------- #


def _format(value, decimals: Optional[int], trim: bool = False) -> str:
    if value is None:
        return ""
    if decimals is None or isinstance(value, str):
        return str(value)
    text = fmt_number(float(value), decimals)
    if trim and "." in text:
        text = text.rstrip("0").rstrip(".") or "0"
    return text


def build_table(
    records: Sequence[dict],
    config: ShortCircuitConfig,
    study_info: Optional[dict] = None,
) -> ShortCircuitTable:
    """
    Build the final table from parsed records and the user's configuration.

    Only the selected columns are produced, in the configured order; nothing is
    padded, reordered or invented.
    """
    config.validate()

    context = fields.DisplayContext(
        current_unit=config.current_unit, voltage_unit=config.voltage_unit
    )
    columns = config.columns()
    headers = [fields.get(key).label_for(context) for key in columns]

    kept, removed = apply_filters(records, config)
    if not kept:
        raise EmptyTableError(_empty_message(records, removed, config))

    standard_text = (study_info or {}).get("title") or config.standard

    rows: list[list[str]] = []
    status_counts: dict[str, int] = {}
    empty_counts = {key: 0 for key in columns}

    for record in kept:
        status = record.get("status", "")
        status_counts[status] = status_counts.get(status, 0) + 1
        record.setdefault("standard", standard_text)

        row: list[str] = []
        for key in columns:
            if key == "remarks":
                text = remark_text(status, config)
            else:
                definition = fields.get(key)
                text = _format(definition.value(record, context), definition.decimals,
                               definition.trim)
                if not text:
                    empty_counts[key] += 1
            row.append(text)
        rows.append(row)

    return ShortCircuitTable(
        headers=headers,
        rows=rows,
        columns=columns,
        status_counts=status_counts,
        warnings=_warnings(config, columns, empty_counts, len(rows), status_counts, context),
        filtered_out=removed,
    )


def _empty_message(records, removed: dict, config: ShortCircuitConfig) -> str:
    """Explain precisely why nothing is left, rather than exporting a blank table."""
    if not records:
        return (
            "The selected report contains no bus short-circuit results, so there "
            "is nothing to export."
        )
    reasons = ", ".join(f"{count} by '{name}'" for name, count in removed.items())
    # Report the last filter to run, since that is the one that emptied the table.
    if removed.get("not alerted"):
        threshold = config.marginal_value or config.critical_value
        return (
            "'Skip non-alerted devices' removed the last "
            f"{removed['not alerted']} row(s): none reaches the {threshold:g} % "
            "threshold, so nothing is alerted. Clear that option to list every "
            "row, or lower the threshold."
        )
    if removed.get("equipment type"):
        return (
            "No bus matches the selected equipment type(s): "
            f"{', '.join(config.equipment_types)}. Removed {reasons}."
        )
    return f"The filters removed every row ({reasons})."


def _warnings(config, columns, empty_counts, total, status_counts, context) -> list[str]:
    """Say plainly where the ETAP report could not fill a selected column."""
    messages: list[str] = []
    for key in columns:
        if key == "remarks":
            continue
        definition = fields.FIELDS[key]
        empty = empty_counts.get(key, 0)
        label = definition.label_for(context)
        if total and empty == total:
            messages.append(f"'{label}' is empty for every row in this report.")
        elif empty:
            messages.append(
                f"'{label}' is blank for {empty} of {total} rows - "
                "the report does not publish it for them."
            )

    unassessed = status_counts.get("", 0)
    if unassessed:
        messages.append(
            f"{unassessed} of {total} rows have no equipment rating in the report, "
            "so they cannot be assessed and their Remarks are blank."
        )
    return messages
