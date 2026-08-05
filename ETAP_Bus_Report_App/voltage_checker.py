"""
voltage_checker.py
Compares a bus's Voltage % against the user-entered acceptable limits and
returns the exact Remarks text required by the spec: "ACCEPTABLE" or
"NOT ACCEPTABLE". No colors, no symbols.
"""
from __future__ import annotations

from utils import VoltageLimits, safe_float

ACCEPTABLE = "ACCEPTABLE"
NOT_ACCEPTABLE = "NOT ACCEPTABLE"
UNKNOWN = "N/A"  # used only if a voltage value could not be read at all


def evaluate_remark(voltage_pct, limits: VoltageLimits) -> str:
    """
    voltage_pct : the bus's Voltage % value (str or float)
    limits      : a VoltageLimits(lower, upper)

    Returns "ACCEPTABLE" if lower <= voltage_pct <= upper, else
    "NOT ACCEPTABLE". Returns "N/A" only if the voltage value itself
    could not be parsed as a number (should not happen for a well-formed
    ETAP report, but guards against malformed/OCR'd input).
    """
    v = safe_float(voltage_pct)
    if v is None:
        return UNKNOWN
    if limits.lower <= v <= limits.upper:
        return ACCEPTABLE
    return NOT_ACCEPTABLE
