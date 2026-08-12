"""
loadflow_config.py
------------------
The user's configuration for a Load Flow report: which columns to include,
the alert limits, display units and output format - plus validation.

Alert limits
============
The application previously classified a bus by a single acceptable band
(``95-106`` by default): inside it "ACCEPTABLE", outside it "NOT ACCEPTABLE".
That meaning is preserved:

* the old lower limit is now the **undervoltage critical** limit,
* the old upper limit is now the **overvoltage critical** limit,

so with the marginal limits switched off the new engine flags exactly the buses
the old one did.  Marginal bands and the loading check are additions, off by
default only in the sense that they are configurable; their defaults follow the
reference interface.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Sequence

from loadflow_fields import (
    BUS_INFO_ORDER,
    DEFAULT_BUS_INFO,
    DEFAULT_POWER_UNIT,
    DEFAULT_RESULTS,
    DEFAULT_VOLTAGE_DISPLAY,
    FIELDS,
    RESULT_ORDER,
)

# --------------------------------------------------------------------------- #
# Defaults
# --------------------------------------------------------------------------- #

#: Existing application default acceptable band (voltage_checker's default).
LEGACY_LOWER = 95.0
LEGACY_UPPER = 106.0

#: Overvoltage: critical keeps the application's existing 106 % upper limit
#: rather than the reference screenshot's 105 %, as instructed.
DEFAULT_OVERVOLTAGE_CRITICAL = LEGACY_UPPER
DEFAULT_OVERVOLTAGE_MARGINAL = 102.0
#: Undervoltage: critical keeps the existing 95 % lower limit.
DEFAULT_UNDERVOLTAGE_CRITICAL = LEGACY_LOWER
DEFAULT_UNDERVOLTAGE_MARGINAL = 98.0
#: Loading limits follow the reference interface (ETAP's own alert defaults).
DEFAULT_LOADING_CRITICAL = 100.0
DEFAULT_LOADING_MARGINAL = 95.0

#: Remark wording.
STATUS_STYLE = "status"   # ACCEPTABLE / MARGINAL / CRITICAL
LEGACY_STYLE = "legacy"   # ACCEPTABLE / NOT ACCEPTABLE  (pre-existing wording)
REMARK_STYLES = (STATUS_STYLE, LEGACY_STYLE)

#: Output formats.
WORD = "Word"
EXCEL = "Excel"
BOTH = "Word + Excel"
OUTPUT_FORMATS = (WORD, EXCEL, BOTH)


class ConfigError(ValueError):
    """Raised when the configuration cannot produce a valid report."""


@dataclass
class AlertLimit:
    """One alert row: a critical and a marginal threshold, each switchable."""

    critical: Optional[float] = None
    marginal: Optional[float] = None
    critical_enabled: bool = True
    marginal_enabled: bool = True

    @property
    def critical_value(self) -> Optional[float]:
        return self.critical if self.critical_enabled else None

    @property
    def marginal_value(self) -> Optional[float]:
        return self.marginal if self.marginal_enabled else None


@dataclass
class LoadFlowConfig:
    """Everything the user chose in the Load Flow interface."""

    bus_info: tuple = DEFAULT_BUS_INFO
    results: tuple = DEFAULT_RESULTS

    loading: AlertLimit = field(
        default_factory=lambda: AlertLimit(DEFAULT_LOADING_CRITICAL, DEFAULT_LOADING_MARGINAL)
    )
    overvoltage: AlertLimit = field(
        default_factory=lambda: AlertLimit(
            DEFAULT_OVERVOLTAGE_CRITICAL, DEFAULT_OVERVOLTAGE_MARGINAL
        )
    )
    undervoltage: AlertLimit = field(
        default_factory=lambda: AlertLimit(
            DEFAULT_UNDERVOLTAGE_CRITICAL, DEFAULT_UNDERVOLTAGE_MARGINAL
        )
    )

    power_unit: str = DEFAULT_POWER_UNIT
    voltage_display: str = DEFAULT_VOLTAGE_DISPLAY

    remark_style: str = STATUS_STYLE
    include_remarks: bool = True
    include_alert_cause: bool = False

    output_format: str = BOTH
    include_pseudo_buses: bool = False

    # -- derived ------------------------------------------------------------ #

    @property
    def wants_word(self) -> bool:
        return self.output_format in (WORD, BOTH)

    @property
    def wants_excel(self) -> bool:
        return self.output_format in (EXCEL, BOTH)

    def columns(self) -> list[str]:
        """
        Deterministic column order:
        Bus ID, selected Bus Info, selected Load Flow Results, [Alert], Remarks.
        """
        columns = ["bus_id"]
        columns += [key for key in BUS_INFO_ORDER if key in self.bus_info]
        columns += [key for key in RESULT_ORDER if key in self.results]
        if self.include_alert_cause:
            columns.append("alert_cause")
        if self.include_remarks:
            columns.append("remarks")
        return columns

    # -- validation --------------------------------------------------------- #

    def validate(self) -> None:
        """Raise :class:`ConfigError` when the configuration is unusable."""
        unknown = [k for k in tuple(self.bus_info) + tuple(self.results) if k not in FIELDS]
        if unknown:
            raise ConfigError(f"Unknown field(s) selected: {', '.join(unknown)}.")

        if not self.bus_info and not self.results:
            raise ConfigError(
                "Select at least one Bus Info or Load Flow Results field - "
                "a report of Bus IDs alone would be empty."
            )

        if self.output_format not in OUTPUT_FORMATS:
            raise ConfigError(f"Choose an output format: {', '.join(OUTPUT_FORMATS)}.")

        if self.remark_style not in REMARK_STYLES:
            raise ConfigError(f"Unknown Remarks style: {self.remark_style!r}.")

        _check_limits("Loading", self.loading, higher_is_worse=True)
        _check_limits("Overvoltage", self.overvoltage, higher_is_worse=True)
        _check_limits("Undervoltage", self.undervoltage, higher_is_worse=False)

        if self.include_remarks and not self._any_limit_enabled():
            raise ConfigError(
                "Remarks are enabled but every alert limit is switched off. "
                "Enable at least one limit, or turn the Remarks column off."
            )

    def _any_limit_enabled(self) -> bool:
        for limit in (self.loading, self.overvoltage, self.undervoltage):
            if limit.critical_value is not None or limit.marginal_value is not None:
                return True
        return False

    def voltage_checks_enabled(self) -> bool:
        return (
            self.overvoltage.critical_value is not None
            or self.overvoltage.marginal_value is not None
            or self.undervoltage.critical_value is not None
            or self.undervoltage.marginal_value is not None
        )


def _check_limits(name: str, limit: AlertLimit, higher_is_worse: bool) -> None:
    for value, kind in ((limit.critical_value, "Critical"), (limit.marginal_value, "Marginal")):
        if value is None:
            continue
        if not isinstance(value, (int, float)) or value != value:
            raise ConfigError(f"{name} {kind} limit must be a number.")
        if not 0 < value < 1000:
            raise ConfigError(f"{name} {kind} limit must be between 0 and 1000 %.")

    critical, marginal = limit.critical_value, limit.marginal_value
    if critical is None or marginal is None:
        return

    if higher_is_worse and marginal > critical:
        raise ConfigError(
            f"{name}: the Marginal limit ({_fmt(marginal)} %) must not be above the "
            f"Critical limit ({_fmt(critical)} %)."
        )
    if not higher_is_worse and marginal < critical:
        raise ConfigError(
            f"{name}: the Marginal limit ({_fmt(marginal)} %) must not be below the "
            f"Critical limit ({_fmt(critical)} %)."
        )


def _fmt(value: float) -> str:
    return f"{value:g}"


def from_legacy_limits(limits_text: str, **overrides) -> LoadFlowConfig:
    """
    Build a configuration equivalent to the old ``Lower-Upper`` voltage band.

    Used by the command line and by any caller that still passes ``limits=``,
    so existing scripts keep producing the same classification.
    """
    import voltage_checker

    band = voltage_checker.parse_limits(limits_text)
    config = LoadFlowConfig(**overrides)
    config.undervoltage = AlertLimit(band.lower, None, True, False)
    config.overvoltage = AlertLimit(band.upper, None, True, False)
    config.loading = AlertLimit(DEFAULT_LOADING_CRITICAL, DEFAULT_LOADING_MARGINAL, False, False)
    config.remark_style = LEGACY_STYLE
    return config
