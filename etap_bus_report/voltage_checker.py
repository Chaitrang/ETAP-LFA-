"""
voltage_checker.py
------------------
Parsing of the user-entered acceptable voltage band and evaluation of the
Remarks column.

Accepted input formats (whitespace, '%' signs and various dashes are ignored)::

    95-106        95 - 105 %      90–110
    95%-106%      95 to 106       95,106
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

from utils import REMARK_ACCEPTABLE, REMARK_NOT_ACCEPTABLE

LIMIT_FORMAT_MESSAGE = (
    "Please enter voltage limits in the format Lower-Upper (Example: 95-106)."
)

#: hyphen, en dash, em dash, minus sign, "to", comma, slash
_SEPARATORS = r"(?:-|\u2010|\u2011|\u2012|\u2013|\u2014|\u2212|to|,|/)"
_LIMITS_RE = re.compile(
    r"^\s*(?P<low>\d+(?:\.\d+)?)\s*%?\s*" + _SEPARATORS + r"\s*(?P<high>\d+(?:\.\d+)?)\s*%?\s*$",
    re.IGNORECASE,
)


class LimitError(ValueError):
    """Raised when the acceptable voltage limits cannot be understood."""

    def __init__(self, message: str = LIMIT_FORMAT_MESSAGE) -> None:
        super().__init__(message)


@dataclass(frozen=True)
class VoltageLimits:
    """An inclusive acceptable band, in percent of nominal voltage."""

    lower: float
    upper: float

    def contains(self, voltage_percent: float) -> bool:
        return self.lower <= voltage_percent <= self.upper

    def __str__(self) -> str:  # pragma: no cover - display only
        return f"{_trim(self.lower)} - {_trim(self.upper)} %"


def _trim(value: float) -> str:
    return f"{value:.2f}".rstrip("0").rstrip(".")


def parse_limits(text: str) -> VoltageLimits:
    """
    Parse a user-entered limit string such as ``"95-106"`` or ``"95 - 105 %"``.

    Raises :class:`LimitError` for anything that is not a Lower-Upper pair,
    including a single value (``"95"``), text (``"abc"``) or an inverted band.
    """
    if text is None:
        raise LimitError()

    match = _LIMITS_RE.match(str(text).strip())
    if not match:
        raise LimitError()

    lower = float(match.group("low"))
    upper = float(match.group("high"))

    if lower >= upper:
        raise LimitError()
    if not (0 < lower < 1000 and 0 < upper < 1000):
        raise LimitError()

    return VoltageLimits(lower=lower, upper=upper)


def evaluate(voltage_percent: Optional[float], limits: VoltageLimits) -> str:
    """
    Return the Remarks text for a single bus.

    A bus without a voltage result (blank cell in the ETAP report) gets an
    empty remark rather than a misleading "NOT ACCEPTABLE".
    """
    if voltage_percent is None:
        return ""
    return REMARK_ACCEPTABLE if limits.contains(voltage_percent) else REMARK_NOT_ACCEPTABLE


def apply_limits(buses, limits: VoltageLimits):
    """Fill the ``remarks`` field of every :class:`utils.BusRecord` in *buses*."""
    for bus in buses:
        bus.remarks = evaluate(bus.voltage_percent, limits)
    return buses
