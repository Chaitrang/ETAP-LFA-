"""
loadflow_processor.py
---------------------
Turns the parsed Load Flow dataset plus the user's configuration into **one**
final table.

That table is the single source for the preview, the Word document and the
Excel workbook, which is what keeps the three identical.  The PDF is never
re-read here: the input is the list of :class:`utils.BusRecord` the existing
parser already produced.

Classification
==============
Each bus is assessed against the configured limits:

* **Overvoltage** - voltage % at or above the critical limit is CRITICAL,
  at or above the marginal limit is MARGINAL.
* **Undervoltage** - voltage % at or below the critical limit is CRITICAL,
  at or below the marginal limit is MARGINAL.
* **Loading** - % loading at or above the critical limit is CRITICAL, at or
  above the marginal limit is MARGINAL.

The worst outcome wins. A bus with no value for a check is simply not assessed
against it; nothing is assumed. With the marginal limits switched off and the
legacy wording selected, this reproduces the previous ACCEPTABLE /
NOT ACCEPTABLE behaviour exactly.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Sequence

import loadflow_fields as fields
from loadflow_config import LEGACY_STYLE, LoadFlowConfig
from utils import REMARK_ACCEPTABLE, REMARK_NOT_ACCEPTABLE, fmt_number

# Status ranking, worst last.
ACCEPTABLE = "ACCEPTABLE"
MARGINAL = "MARGINAL"
CRITICAL = "CRITICAL"
_RANK = {ACCEPTABLE: 0, MARGINAL: 1, CRITICAL: 2}

OVERVOLTAGE = "Overvoltage"
UNDERVOLTAGE = "Undervoltage"
LOADING = "Loading"


@dataclass
class LoadFlowTable:
    """The final dataset: headers, rows and what could not be filled."""

    headers: list[str]
    rows: list[list[str]]
    columns: list[str]
    #: Per-status counts, e.g. ``{"ACCEPTABLE": 231, "MARGINAL": 12}``.
    status_counts: dict = field(default_factory=dict)
    #: Human-readable notes about unavailable or partly empty columns.
    warnings: list[str] = field(default_factory=list)

    @property
    def count(self) -> int:
        return len(self.rows)

    @property
    def flagged(self) -> int:
        """Buses that are not ACCEPTABLE."""
        return sum(
            count for status, count in self.status_counts.items()
            if status not in (ACCEPTABLE, "")
        )


# --------------------------------------------------------------------------- #
# Classification
# --------------------------------------------------------------------------- #


def _worse(current: str, candidate: str) -> str:
    return candidate if _RANK[candidate] > _RANK[current] else current


def classify(record, config: LoadFlowConfig) -> tuple[str, str]:
    """
    Return ``(status, cause)`` for one bus.

    ``status`` is ACCEPTABLE / MARGINAL / CRITICAL; ``cause`` names the limit
    that produced it, or is empty when the bus is acceptable or could not be
    assessed at all.
    """
    status = ACCEPTABLE
    causes: list[str] = []
    assessed = False

    voltage = record.voltage_percent
    if voltage is not None:
        over, under = config.overvoltage, config.undervoltage

        if over.critical_value is not None or over.marginal_value is not None:
            assessed = True
            if over.critical_value is not None and voltage >= over.critical_value:
                status, _ = _worse(status, CRITICAL), causes.append(OVERVOLTAGE)
            elif over.marginal_value is not None and voltage >= over.marginal_value:
                status, _ = _worse(status, MARGINAL), causes.append(OVERVOLTAGE)

        if under.critical_value is not None or under.marginal_value is not None:
            assessed = True
            if under.critical_value is not None and voltage <= under.critical_value:
                status, _ = _worse(status, CRITICAL), causes.append(UNDERVOLTAGE)
            elif under.marginal_value is not None and voltage <= under.marginal_value:
                status, _ = _worse(status, MARGINAL), causes.append(UNDERVOLTAGE)

    percent_loading = record.meta.get("percent_loading")
    loading = config.loading
    if percent_loading is not None and (
        loading.critical_value is not None or loading.marginal_value is not None
    ):
        assessed = True
        if loading.critical_value is not None and percent_loading >= loading.critical_value:
            status, _ = _worse(status, CRITICAL), causes.append(LOADING)
        elif loading.marginal_value is not None and percent_loading >= loading.marginal_value:
            status, _ = _worse(status, MARGINAL), causes.append(LOADING)

    if not assessed:
        return "", ""

    # De-duplicate while keeping order, and only report causes when flagged.
    seen: list[str] = []
    for cause in causes:
        if cause not in seen:
            seen.append(cause)
    return status, ", ".join(seen) if status != ACCEPTABLE else ""


def remark_text(status: str, config: LoadFlowConfig) -> str:
    """Render a status with the wording the user selected."""
    if not status:
        return ""
    if config.remark_style == LEGACY_STYLE:
        return REMARK_ACCEPTABLE if status == ACCEPTABLE else REMARK_NOT_ACCEPTABLE
    return status


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


def build_table(records: Sequence, config: LoadFlowConfig) -> LoadFlowTable:
    """
    Build the final table from parsed records and the user's configuration.

    Only the selected columns are produced, in the configured order; nothing is
    padded, reordered or invented.
    """
    config.validate()

    context = fields.DisplayContext(
        power_unit=config.power_unit, voltage_display=config.voltage_display
    )
    columns = config.columns()
    headers = [fields.get(key).label_for(context) for key in columns]

    rows: list[list[str]] = []
    status_counts: dict[str, int] = {}
    empty_counts = {key: 0 for key in columns}

    for record in records:
        status, cause = classify(record, config)
        status_counts[status] = status_counts.get(status, 0) + 1

        row: list[str] = []
        for key in columns:
            if key == "remarks":
                text = remark_text(status, config)
            elif key == "alert_cause":
                text = cause
            else:
                definition = fields.get(key)
                text = _format(
                    definition.value(record, context), definition.decimals, definition.trim
                )
            if not text and key not in ("remarks", "alert_cause"):
                empty_counts[key] += 1
            row.append(text)
        rows.append(row)

    warnings = _warnings(config, columns, empty_counts, len(rows), status_counts)

    return LoadFlowTable(
        headers=headers,
        rows=rows,
        columns=columns,
        status_counts=status_counts,
        warnings=warnings,
    )


def _warnings(config, columns, empty_counts, total, status_counts) -> list[str]:
    """Say plainly where the ETAP report could not fill a selected column."""
    messages: list[str] = []

    for key in columns:
        definition = fields.FIELDS[key]
        if not definition.available:
            messages.append(
                f"'{definition.label_for(fields.DisplayContext())}' is not available: {definition.note}"
            )
            continue
        empty = empty_counts.get(key, 0)
        if total and empty == total:
            messages.append(
                f"'{definition.label_for(fields.DisplayContext(config.power_unit, config.voltage_display))}' "
                f"is empty for every bus in this report."
            )
        elif empty:
            messages.append(
                f"'{definition.label_for(fields.DisplayContext(config.power_unit, config.voltage_display))}' "
                f"is blank for {empty} of {total} buses - ETAP does not publish it for them."
            )

    unassessed = status_counts.get("", 0)
    if unassessed:
        messages.append(
            f"{unassessed} of {total} buses could not be assessed against the "
            f"enabled limits, so their Remarks are blank."
        )
    return messages
