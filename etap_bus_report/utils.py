"""
utils.py
--------
Shared helpers: number parsing/formatting, geometry helpers used by the PDF
parser, application paths and small configuration constants.

Keeping these here (rather than inside pdf_parser.py) means future extractors
(short circuit, cable schedule, transformer report ...) can reuse them.
"""

from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass, field
from typing import Iterable, Optional

# --------------------------------------------------------------------------- #
# Configuration constants
# --------------------------------------------------------------------------- #

APP_NAME = "ETAP Bus Report Generator"
APP_VERSION = "1.0.0"

#: Number of decimals used when writing the Voltage % column.
VOLTAGE_DECIMALS = 1

#: Number of decimals used when writing the kW / Amp loading columns.
LOADING_DECIMALS = 1

#: Buses whose ID matches any of these patterns are ETAP auto-generated
#: "pseudo" nodes (cable/VFD/transformer internal terminals), not real
#: switchboards or busbars, so they are excluded from the Bus report.
#: Set to an empty list if every node in the report should be listed.
PSEUDO_BUS_PATTERNS = [
    re.compile(r"~"),  # Cable40~, S002-TR-CMP1_VFD~2, S003-Cable43~ ...
]

#: Text used in the Remarks column. The specification requires these exact words.
REMARK_ACCEPTABLE = "ACCEPTABLE"
REMARK_NOT_ACCEPTABLE = "NOT ACCEPTABLE"

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #


def resource_path(relative: str) -> str:
    """
    Return an absolute path to a bundled resource.

    Works both when running from source and when frozen with PyInstaller
    (which unpacks bundled data into ``sys._MEIPASS``).
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def template_path() -> str:
    """Absolute path of the bundled Word template."""
    return resource_path(os.path.join("assets", "Bus_Template.docx"))


# --------------------------------------------------------------------------- #
# Number helpers
# --------------------------------------------------------------------------- #

_NUMBER_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)$")


def is_number(token: str) -> bool:
    """True if *token* is a plain decimal number (no units, no thousands sep)."""
    return bool(_NUMBER_RE.match(token.strip().replace(",", "")))


def to_float(token: str) -> Optional[float]:
    """Parse *token* to float, returning ``None`` when it is not numeric."""
    token = (token or "").strip().replace(",", "")
    if not _NUMBER_RE.match(token):
        return None
    try:
        return float(token)
    except ValueError:
        return None


def fmt_number(value: Optional[float], decimals: int = LOADING_DECIMALS) -> str:
    """Format a float for a report cell; ``None`` becomes an empty string."""
    if value is None:
        return ""
    return f"{value:.{decimals}f}"


def fmt_nominal_voltage(kv: Optional[float]) -> str:
    """
    Format the bus nominal voltage the way an engineer writes it in a report.

    ETAP always prints kV (``0.380``, ``11.000``, ``3.300``).  Values below
    1 kV are converted to volts:

    >>> fmt_nominal_voltage(0.380), fmt_nominal_voltage(11.0)
    ('380 V', '11 kV')
    >>> fmt_nominal_voltage(3.45), fmt_nominal_voltage(0.415)
    ('3.45 kV', '415 V')
    """
    if kv is None:
        return ""
    if kv < 1.0:
        volts = kv * 1000.0
        return f"{_trim(volts)} V"
    return f"{_trim(kv)} kV"


def _trim(value: float) -> str:
    """Drop trailing zeros: 11.000 -> '11', 3.450 -> '3.45', 380.0 -> '380'."""
    text = f"{value:.3f}".rstrip("0").rstrip(".")
    return text or "0"


def is_pseudo_bus(bus_id: str) -> bool:
    """True for ETAP auto-generated internal nodes (see PSEUDO_BUS_PATTERNS)."""
    return any(p.search(bus_id) for p in PSEUDO_BUS_PATTERNS)


def clean_bus_id(raw: str) -> str:
    """
    Tidy a bus ID collected from PDF words.

    Removes the ETAP result flags that are printed immediately in front of the
    ID (``*`` = voltage regulated bus, ``#`` = load mismatch) and collapses the
    whitespace introduced by wrapped IDs such as ``S002-SB-PSS01N01 [BUS\\nA]``.
    """
    text = re.sub(r"\s+", " ", raw).strip()
    text = re.sub(r"^[*#&]+\s*", "", text)
    # "[BUS A]" sometimes arrives as "[BUS A ]" or "[ BUS A]" after wrapping.
    text = text.replace("[ ", "[").replace(" ]", "]")
    return text.strip()


# --------------------------------------------------------------------------- #
# Simple geometry helpers used when reading positioned PDF words
# --------------------------------------------------------------------------- #


def centre(word: dict) -> float:
    """Horizontal centre of a pdfplumber word box."""
    return (word["x0"] + word["x1"]) / 2.0


def cluster_lines(words: Iterable[dict], tolerance: float = 3.0) -> list[list[dict]]:
    """
    Group positioned words into visual lines.

    ETAP prints wrapped cells with a ~1 pt vertical offset relative to the
    numeric cells of the same record, so a small tolerance keeps a record
    together while still separating genuine rows (~9 pt apart).
    """
    ordered = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines: list[list[dict]] = []
    current: list[dict] = []
    current_top: Optional[float] = None

    for word in ordered:
        if current_top is None or abs(word["top"] - current_top) <= tolerance:
            current.append(word)
            current_top = word["top"] if current_top is None else current_top
        else:
            lines.append(sorted(current, key=lambda w: w["x0"]))
            current = [word]
            current_top = word["top"]
    if current:
        lines.append(sorted(current, key=lambda w: w["x0"]))
    return lines


def line_text(line: Iterable[dict]) -> str:
    """Join the words of a line into a single string."""
    return " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #


@dataclass
class BusRecord:
    """One row of the Bus report."""

    bus_id: str
    nominal_kv: Optional[float] = None
    voltage_percent: Optional[float] = None
    kw_loading: Optional[float] = None
    amp_loading: Optional[float] = None
    remarks: str = ""
    #: Free-form notes (e.g. which PDF table each value came from) - useful
    #: for debugging and for future Excel export.
    meta: dict = field(default_factory=dict)

    # -- formatted views used by the Word writer ---------------------------- #

    @property
    def nominal_text(self) -> str:
        return fmt_nominal_voltage(self.nominal_kv)

    @property
    def voltage_text(self) -> str:
        return fmt_number(self.voltage_percent, VOLTAGE_DECIMALS)

    @property
    def kw_text(self) -> str:
        return fmt_number(self.kw_loading, LOADING_DECIMALS)

    @property
    def amp_text(self) -> str:
        return fmt_number(self.amp_loading, LOADING_DECIMALS)

    def as_row(self) -> list[str]:
        """The six template columns, in order."""
        return [
            self.bus_id,
            self.nominal_text,
            self.voltage_text,
            self.kw_text,
            self.amp_text,
            self.remarks,
        ]


class ParserError(Exception):
    """Raised when the ETAP report cannot be read or does not contain buses."""
