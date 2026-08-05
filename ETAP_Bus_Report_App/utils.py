"""
utils.py
Small, dependency-free helper functions shared across the application:
  - formatting the "Nominal (kV, A)" column exactly as ETAP would display it
  - parsing the free-form "Acceptable Voltage Limits" field the user types
  - light numeric helpers used by the parser / word writer
"""
from __future__ import annotations

import re
from dataclasses import dataclass


# --------------------------------------------------------------------------- #
# Nominal voltage formatting
# --------------------------------------------------------------------------- #
def format_nominal_voltage(kv_str: str) -> str:
    """
    Convert an ETAP kV value (e.g. "132.000", "0.415", "3.300") into the
    display format requested in the spec:
        132.000 -> "132 kV"
        11.000  -> "11 kV"
        3.300   -> "3.3 kV"
        0.415   -> "415 V"
        0.380   -> "380 V"

    Buses below 1 kV are displayed in Volts (V); buses at or above 1 kV are
    displayed in kV. Trailing zeros are trimmed but meaningful decimals are
    kept (3.300 -> 3.3, 11.500 -> 11.5).
    """
    try:
        kv = float(kv_str)
    except (TypeError, ValueError):
        return str(kv_str)

    if kv >= 1:
        value = kv
        unit = "kV"
    else:
        value = kv * 1000
        unit = "V"

    # Trim trailing zeros, but keep the value looking natural
    if value == int(value):
        text = str(int(value))
    else:
        text = f"{value:.3f}".rstrip("0").rstrip(".")
    return f"{text} {unit}"


# --------------------------------------------------------------------------- #
# Voltage limits parsing
# --------------------------------------------------------------------------- #
@dataclass
class VoltageLimits:
    lower: float
    upper: float


_LIMIT_RE = re.compile(
    r"""^\s*
        (?P<lower>\d+(?:\.\d+)?)
        \s*%?\s*
        (?:-|to|–|—)
        \s*
        (?P<upper>\d+(?:\.\d+)?)
        \s*%?\s*$
    """,
    re.VERBOSE | re.IGNORECASE,
)


def parse_voltage_limits(text: str) -> VoltageLimits | None:
    """
    Parse strings like:
        "95-106", "95 - 105 %", "90-110", "95%-106%", "95 to 106"
    into a VoltageLimits(lower, upper).

    Returns None if the text cannot be parsed (caller should show the
    standard error message defined in the spec).
    """
    if not text:
        return None
    match = _LIMIT_RE.match(text.strip())
    if not match:
        return None
    lower = float(match.group("lower"))
    upper = float(match.group("upper"))
    if lower > upper:
        lower, upper = upper, lower
    return VoltageLimits(lower=lower, upper=upper)


# --------------------------------------------------------------------------- #
# Small numeric helpers
# --------------------------------------------------------------------------- #
def safe_float(value, default=None):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def mva_pf_to_kw(mva, pf_percent):
    """Real power (kW) = MVA * (PF/100) * 1000."""
    mva_f = safe_float(mva)
    pf_f = safe_float(pf_percent)
    if mva_f is None or pf_f is None:
        return 0.0
    return round(mva_f * (pf_f / 100.0) * 1000.0, 2)
