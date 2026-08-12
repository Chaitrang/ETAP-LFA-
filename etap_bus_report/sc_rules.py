"""
sc_rules.py
-----------
Acceptance rules for the Remarks column of a Short Circuit report.

Rules are plain functions ``(bus: dict) -> str`` registered in :data:`RULES`.
Adding a project-specific rule is a function plus one dictionary entry - the
parser and the writer never change.
"""

from __future__ import annotations

from typing import Callable, Optional

from utils import REMARK_ACCEPTABLE, REMARK_NOT_ACCEPTABLE

#: Margin applied to the equipment rating, e.g. 0.95 to require a 5 % margin.
DEFAULT_MARGIN = 1.0


def peak_vs_rating(bus: dict, margin: float = DEFAULT_MARGIN) -> str:
    """
    Default rule: the equipment's making (peak) capacity must cover the
    calculated peak current ``ip``.

    Blank when the report gives no rating for the bus, because "NOT
    ACCEPTABLE" would then be an unverified claim rather than a result.
    """
    rating = bus.get("rating_peak")
    peak = bus.get("ip_peak")
    if rating is None or peak is None:
        return ""
    return REMARK_ACCEPTABLE if peak <= rating * margin else REMARK_NOT_ACCEPTABLE


def peak_and_breaking(bus: dict, margin: float = DEFAULT_MARGIN) -> str:
    """Both the peak capacity vs ip and the breaking capacity vs I"k must pass."""
    verdicts = []
    for rating_key, duty_key in (("rating_peak", "ip_peak"), ("rating_ib_sym", "ik_initial")):
        rating, duty = bus.get(rating_key), bus.get(duty_key)
        if rating is None or duty is None:
            continue
        verdicts.append(duty <= rating * margin)
    if not verdicts:
        return ""
    return REMARK_ACCEPTABLE if all(verdicts) else REMARK_NOT_ACCEPTABLE


def etap_flag(bus: dict) -> str:
    """
    Trust ETAP's own duty check: any device the report marked with ``*``
    (calculated duty exceeding the device capability) fails the bus.
    """
    if not bus.get("devices"):
        return ""
    return REMARK_NOT_ACCEPTABLE if bus.get("device_exceeded") else REMARK_ACCEPTABLE


def always_blank(bus: dict) -> str:
    """Leave Remarks empty - for templates where an engineer fills it in."""
    return ""


RULES: dict[str, Callable[[dict], str]] = {
    "peak_vs_rating": peak_vs_rating,
    "peak_and_breaking": peak_and_breaking,
    "etap_flag": etap_flag,
    "blank": always_blank,
}

DEFAULT_RULE = "peak_vs_rating"


def apply_rule(buses, rule: Optional[str] = None):
    """Fill ``bus["remarks"]`` for every record using the named rule."""
    function = RULES.get(rule or DEFAULT_RULE)
    if function is None:
        raise KeyError(f"Unknown Remarks rule: {rule!r}. Available: {', '.join(sorted(RULES))}")
    for bus in buses:
        bus["remarks"] = function(bus)
    return buses
