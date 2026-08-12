"""
template_mapping.py
-------------------
Maps the **column headers of a Word template** onto the canonical data fields
produced by a parser, so that no report writer has to hard-code its columns.

Add a column to a template - or reorder the existing ones - and the writer
follows, as long as the header wording matches one of the patterns below.
Adding support for a new quantity is one entry in :data:`FIELD_PATTERNS` plus
the value in the parser.

The matching is deliberately ordered: ``Bus Rating (kV, A)`` and ``Switchgear
Rating (kA)`` both say "rating", so the voltage patterns are tested first.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Optional

from utils import fmt_nominal_voltage, fmt_number

#: Decimals used for short-circuit currents (kA).
CURRENT_DECIMALS = 2


# --------------------------------------------------------------------------- #
# Field catalogue
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class Field:
    """A quantity a template column can ask for."""

    key: str
    label: str
    formatter: Callable[[object], str]
    #: Alternative keys a parser may use for the same quantity.
    aliases: tuple = ()

    def read(self, record: dict):
        for key in (self.key,) + tuple(self.aliases):
            if record.get(key) is not None:
                return record[key]
        return None

    def format(self, value) -> str:
        if value is None or value == "":
            return ""
        return self.formatter(value)


def _current(value) -> str:
    return fmt_number(float(value), CURRENT_DECIMALS)


def _text(value) -> str:
    return str(value)


FIELDS: dict[str, Field] = {
    "bus_id": Field("bus_id", "Bus / Switchgear ID", _text, aliases=("id", "switchgear_id")),
    "nominal_voltage": Field(
        "nominal_voltage", "Nominal voltage",
        lambda v: fmt_nominal_voltage(float(v)),
        aliases=("nominal_kv", "kv"),
    ),
    "rating_peak": Field("rating_peak", "Equipment making (peak) capacity kA", _current),
    "rating_ib_sym": Field("rating_ib_sym", "Equipment breaking capacity kA sym", _current),
    "rating_ib_asym": Field("rating_ib_asym", "Equipment breaking capacity kA asym", _current),
    "rating_idc": Field("rating_idc", "Equipment DC capacity kA", _current),
    "ik_initial": Field("ik_initial", 'Initial symmetrical current I"k kA', _current),
    "ip_peak": Field("ip_peak", "Peak current ip kA", _current),
    "ib_sym": Field("ib_sym", "Breaking current Ib sym kA", _current),
    "ib_asym": Field("ib_asym", "Breaking current Ib asym kA", _current),
    "idc": Field("idc", "DC component Idc kA", _current),
    "ik_steady": Field("ik_steady", "Steady state current Ik kA", _current),
    "xr_ratio": Field("xr_ratio", "X/R ratio", lambda v: fmt_number(float(v), 1)),
    "governing_device": Field("governing_device", "Governing device", _text),
    "bus_type": Field("bus_type", "Equipment type", _text),
    "fault_type": Field("fault_type", "Fault type", _text),
    "voltage_percent": Field("voltage_percent", "Bus voltage %", lambda v: fmt_number(float(v), 1)),
    "kw_loading": Field("kw_loading", "kW loading", lambda v: fmt_number(float(v), 1)),
    "amp_loading": Field("amp_loading", "Amp loading", lambda v: fmt_number(float(v), 1)),
    "remarks": Field("remarks", "Remarks", _text),
}

#: Fields that only the per-bus detail pages can supply (expensive to read).
DETAIL_ONLY_FIELDS = {"xr_ratio"}

# --------------------------------------------------------------------------- #
# Header patterns - order matters, first match wins
# --------------------------------------------------------------------------- #

FIELD_PATTERNS: list[tuple[str, str]] = [
    (r"remark|comment|status|accept", "remarks"),
    # Load-flow quantities (kept here so the catalogue covers every template).
    (r"voltage\s*%|%\s*voltage|voltage\s*\(%\)|%\s*mag", "voltage_percent"),
    (r"\bkw\b", "kw_loading"),
    (r"\bamp\b.*load|load.*\bamp\b|\bamps?\b\s*$", "amp_loading"),
    # Voltage before rating: "Bus Rating (kV, A)" is a voltage column.
    (r"\bkv\b|nominal|voltage(?!.*factor)", "nominal_voltage"),
    (r"x\s*/\s*r", "xr_ratio"),
    (r"fault\s*type", "fault_type"),
    (r"(equipment|device|switchgear|bus)\s*type", "bus_type"),
    (r"(governing|protective)\s*device|device\s*id", "governing_device"),
    # Equipment ratings ("Switchgear Rating (kA) / Ip (peak)") - before the
    # generic ID rule, because they also mention the switchgear.
    (r"(rating|capacity|withstand).*asym", "rating_ib_asym"),
    (r"(rating|capacity|withstand).*(ib|breaking|sym)", "rating_ib_sym"),
    (r"(rating|capacity|withstand).*(idc|dc)", "rating_idc"),
    (r"rating|capacity|withstand", "rating_peak"),
    (r"\bid\b|\bname\b", "bus_id"),
    # ETAP results.  I"k first: it also contains the letters i and k.
    (r'i\s*["\'\u201c\u201d]\s*k|initial|symmetrical', "ik_initial"),
    (r"\bip\b|peak|making", "ip_peak"),
    (r"ib.*asym|asym", "ib_asym"),
    (r"ib.*sym|breaking|\bib\b", "ib_sym"),
    (r"\bidc\b|dc\s*component", "idc"),
    (r"\bik\b|steady", "ik_steady"),
]

_COMPILED = [(re.compile(pattern, re.I), key) for pattern, key in FIELD_PATTERNS]


def normalise_header(text: str) -> str:
    """Collapse whitespace and unify the various quote characters."""
    text = text.replace("\u201c", '"').replace("\u201d", '"').replace("\u2019", "'")
    return re.sub(r"\s+", " ", text).strip()


def match_field(header_text: str) -> Optional[str]:
    """Return the canonical field key for one template column header."""
    text = normalise_header(header_text)
    if not text:
        return None
    for pattern, key in _COMPILED:
        if pattern.search(text):
            return key
    return None


# --------------------------------------------------------------------------- #
# Reading a template's header
# --------------------------------------------------------------------------- #

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _cell_count(row) -> int:
    """Number of real ``<w:tc>`` elements in a row (merges not expanded)."""
    return len(row._tr.findall(f"{_W}tc"))


def header_row_count(table) -> int:
    """
    How many leading rows of *table* are header rows.

    A data row is the first row that has the table's full column count and is
    completely empty.  This copes with multi-row headers, with merged group
    headers and with the shaded spacer band some templates use.
    """
    grid_columns = len(table.columns)
    for index, row in enumerate(table.rows):
        if _cell_count(row) != grid_columns:
            continue  # merged banner / group header row
        if all(not cell.text.strip() for cell in row.cells):
            return index
    # No blank row at all - assume a single header row.
    return 1


def column_headers(table, header_rows: Optional[int] = None) -> list[str]:
    """Combined header text of each column (merged cells expanded and dedup'd)."""
    header_rows = header_rows if header_rows is not None else header_row_count(table)
    width = len(table.columns)
    headers: list[str] = []

    for index in range(width):
        parts: list[str] = []
        for row in table.rows[:header_rows]:
            try:
                text = normalise_header(row.cells[index].text)
            except IndexError:
                continue
            if text and text not in parts:
                parts.append(text)
        headers.append(" ".join(parts))
    return headers


def map_template(table) -> tuple[int, list[Optional[str]], list[str]]:
    """
    Inspect a template table.

    Returns ``(header_rows, field_keys, header_texts)`` where ``field_keys[i]``
    is the canonical field for column *i* (``None`` when nothing matched, in
    which case the column is left untouched).
    """
    rows = header_row_count(table)
    headers = column_headers(table, rows)
    return rows, [match_field(text) for text in headers], headers


def format_row(record: dict, field_keys: list[Optional[str]]) -> list[Optional[str]]:
    """
    Format one parsed record into the template's columns.

    ``None`` entries mean "leave this cell alone" so that unmatched columns
    keep whatever the template already contains.
    """
    values: list[Optional[str]] = []
    for key in field_keys:
        if key is None:
            values.append(None)
            continue
        field = FIELDS.get(key)
        values.append(field.format(field.read(record)) if field else "")
    return values
