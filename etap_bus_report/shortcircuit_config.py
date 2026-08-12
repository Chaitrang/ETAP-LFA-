"""
shortcircuit_config.py
----------------------
The user's configuration for a Short Circuit report: standard, study type,
selected report, chosen columns, alert thresholds, filters, units and output
format - plus validation.

Alert thresholds
================
The existing acceptance rule was "the equipment's making (peak) capacity must
cover the calculated peak current ip" (:mod:`sc_rules`). That meaning is kept;
the thresholds simply make the comparison configurable, and they match ETAP's
own footnotes in the report:

    * Indicates a device with calculated duty exceeding the device capability.
    # Indicates a device with calculated duty exceeding the device marginal
      limit ( 95 % times device capability).

So Critical 100 % and Marginal 95 % reproduce ETAP's own flags.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from shortcircuit_fields import (
    DEFAULT_CURRENT_UNIT,
    DEFAULT_INFO,
    DEFAULT_RESULTS,
    DEFAULT_VOLTAGE_UNIT,
    EQUIPMENT_TYPES,
    FIELDS,
    INFO_ORDER,
    RESULT_ORDER,
    CURRENT_UNITS,
    VOLTAGE_UNITS,
)

STANDARDS = ("IEC", "ANSI")
STUDY_TYPES = ("3-phase", "1-phase")
STUDY_TYPE_LABELS = {"3-phase": "3-Ph Device Duty", "1-phase": "1-Ph Device Duty"}

#: ETAP's own alert defaults, as printed in the report footnotes.
DEFAULT_CRITICAL = 100.0
DEFAULT_MARGINAL = 95.0

#: Remark wording. ``status`` adds MARGINAL; ``legacy`` is the pre-existing
#: ACCEPTABLE / NOT ACCEPTABLE pair produced by sc_rules.peak_vs_rating.
STATUS_STYLE = "status"
LEGACY_STYLE = "legacy"
REMARK_STYLES = (STATUS_STYLE, LEGACY_STYLE)

WORD = "Word"
EXCEL = "Excel"
BOTH = "Word + Excel"
OUTPUT_FORMATS = (WORD, EXCEL, BOTH)


class ConfigError(ValueError):
    """Raised when the configuration cannot produce a valid report."""


@dataclass
class ShortCircuitConfig:
    """Everything the user chose in the Short Circuit interface."""

    # -- study ------------------------------------------------------------- #
    standard: str = "IEC"
    study_type: str = "3-phase"
    #: Study case / data block the table is built from (e.g. "SC_Max").
    selected_report: str = ""

    # -- columns ----------------------------------------------------------- #
    info: tuple = DEFAULT_INFO
    results: tuple = DEFAULT_RESULTS
    include_remarks: bool = True
    include_duty: bool = False
    remark_style: str = STATUS_STYLE

    # -- alert -------------------------------------------------------------- #
    critical: float = DEFAULT_CRITICAL
    critical_enabled: bool = True
    marginal: float = DEFAULT_MARGINAL
    marginal_enabled: bool = True
    skip_non_alerted: bool = False
    skip_nodes: bool = True

    # -- filters ------------------------------------------------------------ #
    #: Empty tuple means "every equipment type".
    equipment_types: tuple = ()

    # -- units and output ---------------------------------------------------- #
    current_unit: str = DEFAULT_CURRENT_UNIT
    voltage_unit: str = DEFAULT_VOLTAGE_UNIT
    output_format: str = BOTH

    # -- derived ------------------------------------------------------------ #

    @property
    def wants_word(self) -> bool:
        return self.output_format in (WORD, BOTH)

    @property
    def wants_excel(self) -> bool:
        return self.output_format in (EXCEL, BOTH)

    @property
    def critical_value(self) -> Optional[float]:
        return self.critical if self.critical_enabled else None

    @property
    def marginal_value(self) -> Optional[float]:
        return self.marginal if self.marginal_enabled else None

    def columns(self) -> list[str]:
        """
        Deterministic column order:
        ID, selected Info, selected Results, [Duty %], [Remarks].
        """
        columns = ["bus_id"]
        columns += [key for key in INFO_ORDER if key in self.info]
        columns += [key for key in RESULT_ORDER if key in self.results]
        if self.include_duty:
            columns.append("duty_percent")
        if self.include_remarks:
            columns.append("remarks")
        return columns

    def needs_detail_pages(self) -> bool:
        from shortcircuit_fields import needs_detail

        return needs_detail(self.columns())

    # -- validation ---------------------------------------------------------- #

    def validate(self) -> None:
        unknown = [k for k in tuple(self.info) + tuple(self.results) if k not in FIELDS]
        if unknown:
            raise ConfigError(f"Unknown field(s) selected: {', '.join(unknown)}.")

        unavailable = [k for k in tuple(self.info) + tuple(self.results)
                       if not FIELDS[k].available]
        if unavailable:
            from shortcircuit_fields import DisplayContext

            names = ", ".join(FIELDS[k].label_for(DisplayContext()) for k in unavailable)
            raise ConfigError(
                f"This report does not publish: {names}. Clear those selections."
            )

        if not self.info and not self.results:
            raise ConfigError(
                "Select at least one Info or Results field - a report of IDs "
                "alone would be empty."
            )

        if self.standard not in STANDARDS:
            raise ConfigError(f"Standard must be one of: {', '.join(STANDARDS)}.")
        if self.study_type not in STUDY_TYPES:
            raise ConfigError(f"Study type must be one of: {', '.join(STUDY_TYPES)}.")
        if self.output_format not in OUTPUT_FORMATS:
            raise ConfigError(f"Choose an output format: {', '.join(OUTPUT_FORMATS)}.")
        if self.remark_style not in REMARK_STYLES:
            raise ConfigError(f"Unknown Remarks style: {self.remark_style!r}.")

        if self.current_unit not in CURRENT_UNITS:
            raise ConfigError(f"Current unit must be one of: {', '.join(CURRENT_UNITS)}.")
        if self.voltage_unit not in VOLTAGE_UNITS:
            raise ConfigError(f"Voltage unit must be one of: {', '.join(VOLTAGE_UNITS)}.")

        for value, name in ((self.critical_value, "Critical"), (self.marginal_value, "Marginal")):
            if value is None:
                continue
            if not isinstance(value, (int, float)) or value != value:
                raise ConfigError(f"The {name} threshold must be a number.")
            if not 0 < value <= 1000:
                raise ConfigError(f"The {name} threshold must be between 0 and 1000 %.")

        if (
            self.critical_value is not None
            and self.marginal_value is not None
            and self.marginal_value > self.critical_value
        ):
            raise ConfigError(
                f"The Marginal threshold ({self.marginal_value:g} %) must not be above "
                f"the Critical threshold ({self.critical_value:g} %)."
            )

        if self.include_remarks and self.critical_value is None and self.marginal_value is None:
            raise ConfigError(
                "Remarks are enabled but both thresholds are switched off. "
                "Enable a threshold, or turn the Remarks column off."
            )

        if self.skip_non_alerted and self.critical_value is None and self.marginal_value is None:
            raise ConfigError(
                "'Skip non-alerted devices' needs at least one threshold enabled, "
                "otherwise nothing can be classified as alerted."
            )

        unknown_types = [t for t in self.equipment_types if t not in EQUIPMENT_TYPES]
        if unknown_types:
            raise ConfigError(f"Unknown equipment type(s): {', '.join(unknown_types)}.")


def from_study_info(info: dict, **overrides) -> ShortCircuitConfig:
    """
    Build a configuration seeded from what the report itself states.

    The standard, study type and study case are detected rather than assumed,
    so opening the interface already reflects the uploaded report.
    """
    config = ShortCircuitConfig(**overrides)
    if info.get("standard") in STANDARDS:
        config.standard = info["standard"]
    if info.get("study_type") in STUDY_TYPES:
        config.study_type = info["study_type"]
    if info.get("study_case"):
        config.selected_report = info["study_case"]
    return config
