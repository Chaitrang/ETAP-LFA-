"""
pdf_parser.py
-------------
Extraction of Bus load-flow results from an ETAP Load Flow Analysis report.

Design notes
============
The parser never relies on fixed page numbers.  It works purely from the
repeating *table headers* that ETAP prints at the top of every page of a
section, so it copes with reports of any length and with the header/footer
frame changing between ETAP versions.

Two result tables are used:

``LOAD FLOW REPORT``  (a.k.a. "Bus Output Data" / "Bus Load Flow Results")
    Source of the bus order, the nominal kV and the bus **Voltage %**.

``Bus Loading Summary Report``
    Source of the **total bus loading** (MVA and %PF -> kW) and of the
    **Amp Loading**.  This is the correct source for a bus loading table: in
    the Load Flow Report the "Load MW" column only shows load connected
    *directly* to the bus, so a switchboard that feeds everything through
    outgoing feeders would otherwise read 0 kW.

If the Loading Summary is missing from the report, the parser falls back to the
directly-connected load (Load MW) and the branch Amp of the Load Flow Report.

Words are read with their coordinates (pdfplumber) rather than as flat text,
because ETAP wraps long IDs onto a second line and pads columns with spaces -
both of which scramble plain text extraction.  Column boundaries are derived
from the header token positions on each page.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Optional

import pdfplumber

from utils import (
    BusRecord,
    ParserError,
    centre,
    clean_bus_id,
    cluster_lines,
    is_pseudo_bus,
    line_text,
    to_float,
)

# --------------------------------------------------------------------------- #
# Messages required by the specification
# --------------------------------------------------------------------------- #

MSG_NO_BUS_TABLE = "No Bus Output Data found in the uploaded ETAP report."
MSG_UNREADABLE = "Unable to read the ETAP report."


# --------------------------------------------------------------------------- #
# Column model
# --------------------------------------------------------------------------- #


@dataclass
class Column:
    """A table column, described by its horizontal span on the page."""

    name: str
    left: float
    right: float

    def contains(self, word: dict) -> bool:
        return self.left <= centre(word) < self.right


class ColumnMap:
    """Maps positioned words onto named columns for one page."""

    def __init__(self, columns: list[Column]) -> None:
        self.columns = columns
        self._by_name = {c.name: c for c in columns}

    def get(self, name: str) -> Optional[Column]:
        return self._by_name.get(name)

    def words_in(self, line: list[dict], name: str) -> list[dict]:
        column = self._by_name.get(name)
        if column is None:
            return []
        return [w for w in line if column.contains(w)]

    def text_in(self, line: list[dict], name: str) -> str:
        return " ".join(w["text"] for w in self.words_in(line, name))

    def value_in(self, line: list[dict], name: str) -> Optional[float]:
        words = self.words_in(line, name)
        for word in words:
            value = to_float(word["text"])
            if value is not None:
                return value
        return None

    def outside(self, line: list[dict], names: list[str]) -> list[dict]:
        """Words that fall in none of the named columns."""
        cols = [self._by_name[n] for n in names if n in self._by_name]
        return [w for w in line if not any(c.contains(w) for c in cols)]


def _columns_from_anchors(anchors: list[tuple[str, float]], page_width: float) -> ColumnMap:
    """
    Build a :class:`ColumnMap` from ``(name, header_centre)`` anchors.

    Boundaries are placed midway between consecutive header centres; the first
    column extends to the left page edge and the last to the right page edge.
    """
    anchors = sorted(anchors, key=lambda a: a[1])
    columns: list[Column] = []
    for index, (name, position) in enumerate(anchors):
        left = 0.0 if index == 0 else (anchors[index - 1][1] + position) / 2.0
        right = (
            page_width
            if index == len(anchors) - 1
            else (position + anchors[index + 1][1]) / 2.0
        )
        columns.append(Column(name, left, right))
    return ColumnMap(columns)


# --------------------------------------------------------------------------- #
# Header recognition
# --------------------------------------------------------------------------- #


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip().lower()


def _is_load_flow_header(text: str) -> bool:
    """
    Header of the bus Load Flow results table, e.g.::

        ID kV % Mag. Ang. MW Mvar MW Mvar ID MW Mvar Amp %PF %Tap

    The 'Amp' + '%PF' pair is what distinguishes it from the *Bus Input Data*
    header (which has the same leading columns but no flow results).
    """
    t = _norm(text)
    return (
        t.startswith("id kv")
        and "mag." in t
        and "ang" in t
        and " amp" in t
        and ("%pf" in t or "% pf" in t)
    )


def _is_loading_summary_header(text: str) -> bool:
    """Header of the Bus Loading Summary Report."""
    t = _norm(text)
    return (
        t.startswith("id kv")
        and "rated" in t
        and "mva" in t
        and ("%pf" in t or "% pf" in t)
        and "loading" in t
    )


def _group_labels(lines: list[list[dict]], header_index: int) -> list[dict]:
    """Words of the group-header line printed above the column sub-header."""
    if header_index == 0:
        return []
    return lines[header_index - 1]


def _load_flow_columns(header: list[dict], groups: list[dict], page_width: float) -> ColumnMap:
    """
    Derive the columns of interest of the Load Flow results table.

    ``MW``/``Mvar`` appear three times (Generation, Load, Load Flow); each
    occurrence is attributed to the nearest group label printed above it.
    """
    tokens = sorted(header, key=lambda w: w["x0"])
    anchors: list[tuple[str, float]] = []

    group_centres = {
        _norm(w["text"]): centre(w) for w in groups if _norm(w["text"]) in {"generation", "load", "flow"}
    }

    mw_seen = 0
    id_seen = 0
    for word in tokens:
        text = _norm(word["text"])
        pos = centre(word)
        if text == "id":
            id_seen += 1
            anchors.append(("bus_id" if id_seen == 1 else "branch_id", pos))
        elif text == "kv":
            anchors.append(("kv", pos))
        elif text in {"mag.", "mag"}:
            anchors.append(("voltage_percent", pos))
        elif text in {"ang.", "ang", "angle"}:
            anchors.append(("angle", pos))
        elif text == "mw":
            mw_seen += 1
            # 1st MW = Generation, 2nd = directly connected Load, 3rd = branch flow
            name = {1: "gen_mw", 2: "load_mw", 3: "flow_mw"}.get(mw_seen, f"mw_{mw_seen}")
            if group_centres:
                nearest = min(group_centres, key=lambda g: abs(group_centres[g] - pos))
                if nearest == "load" and mw_seen != 3:
                    name = "load_mw"
                elif nearest == "generation":
                    name = "gen_mw"
            anchors.append((name, pos))
        elif text == "mvar":
            anchors.append((f"mvar_{len([a for a in anchors if a[0].startswith('mvar')]) + 1}", pos))
        elif text == "amp":
            anchors.append(("flow_amp", pos))
        elif text in {"%pf", "pf"}:
            anchors.append(("pf", pos))
        elif text in {"%tap", "tap"}:
            anchors.append(("tap", pos))
        elif text == "%":
            continue  # part of "% Mag." - handled by the 'Mag.' token
        else:
            anchors.append((f"other_{len(anchors)}", pos))

    return _columns_from_anchors(anchors, page_width)


def _loading_summary_columns(header: list[dict], page_width: float) -> ColumnMap:
    """
    Derive the columns of the Bus Loading Summary Report.

    Sub-header order::

        ID kV Rated Amp MW Mvar (x4 load categories) MVA %PF Amp Loading
                                                     ^total ^   ^amp  ^percent
    """
    tokens = sorted(header, key=lambda w: w["x0"])
    anchors: list[tuple[str, float]] = []
    seen_mva = False
    seen_pf = False

    for word in tokens:
        text = _norm(word["text"])
        pos = centre(word)
        if text == "id" and not anchors:
            anchors.append(("bus_id", pos))
        elif text == "kv":
            anchors.append(("kv", pos))
        elif text == "rated":
            anchors.append(("rated_amp", pos))
        elif text == "mva":
            anchors.append(("total_mva", pos))
            seen_mva = True
        elif text in {"%pf", "pf"}:
            anchors.append(("pf", pos))
            seen_pf = True
        elif text == "amp":
            # 'Amp' occurs twice: once as part of 'Rated Amp', once as the
            # 'Amp Loading' value column that follows %PF.
            if seen_pf:
                anchors.append(("amp_loading", pos))
        elif text == "loading":
            if seen_mva:
                anchors.append(("percent_loading", pos))
        elif text in {"mw", "mvar"}:
            anchors.append((f"cat_{len(anchors)}", pos))
        else:
            anchors.append((f"other_{len(anchors)}", pos))

    return _columns_from_anchors(anchors, page_width)


# --------------------------------------------------------------------------- #
# Row recognition
# --------------------------------------------------------------------------- #

_FOOTNOTE_RE = re.compile(r"^\s*[*#]?\s*(indicates|total number|this transmission)", re.I)


def _is_footnote(text: str) -> bool:
    return bool(_FOOTNOTE_RE.match(text))


def _split_row(line: list[dict], columns: ColumnMap) -> tuple[list[dict], Optional[float]]:
    """
    Split a data line into (Bus ID words, nominal kV value).

    ETAP does not clip long Bus IDs: ``TYPICAL MOTOR FEEDER 1`` and
    ``POWER SOCKET FEEDER 1 B`` spill past the nominal ID column and into the
    kV column band.  The kV figure is always the *right-most* number of that
    band (ETAP right-aligns numbers), so everything printed to its left belongs
    to the Bus ID.
    """
    kv_column = columns.get("kv")
    if kv_column is None:
        return [], None

    candidates = [w for w in line if kv_column.contains(w) and to_float(w["text"]) is not None]
    if not candidates:
        return columns.words_in(line, "bus_id"), None

    kv_word = max(candidates, key=lambda w: w["x1"])
    bus_words = [w for w in line if w["x1"] <= kv_word["x0"] - 0.5]
    return bus_words, to_float(kv_word["text"])


def _continuation_words(line: list[dict], columns: ColumnMap, value_columns: list[str]) -> bool:
    """
    True when *line* only carries the tail of a wrapped Bus ID.

    ETAP wraps e.g. ``S002-SB-PSS01N01 [BUS`` / ``A]`` and
    ``TYPICAL MOTOR`` / ``FEEDER 1 A`` into the ID column while all numeric
    cells stay on the first line.
    """
    bus_words = columns.words_in(line, "bus_id")
    if not bus_words or len(bus_words) > 4:
        return False
    if len(bus_words) != len(line):
        return False  # something else on this line -> not a wrapped ID
    return not _is_footnote(line_text(line))


# --------------------------------------------------------------------------- #
# Page scanning
# --------------------------------------------------------------------------- #


@dataclass
class _PageTable:
    page_index: int
    kind: str  # 'load_flow' | 'loading_summary'
    columns: ColumnMap
    lines: list[list[dict]]
    header_index: int


def _scan_pages(pdf) -> list[_PageTable]:
    """Find every page that carries a bus results header, in document order."""
    tables: list[_PageTable] = []

    for index, page in enumerate(pdf.pages):
        try:
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        except Exception:  # pragma: no cover - corrupt page
            continue
        if not words:
            continue

        lines = cluster_lines(words)
        for line_index, line in enumerate(lines):
            text = line_text(line)
            if _is_load_flow_header(text):
                columns = _load_flow_columns(line, _group_labels(lines, line_index), page.width)
                tables.append(_PageTable(index, "load_flow", columns, lines, line_index))
                break
            if _is_loading_summary_header(text):
                columns = _loading_summary_columns(line, page.width)
                tables.append(_PageTable(index, "loading_summary", columns, lines, line_index))
                break

    return tables


def _last_block(tables: list[_PageTable], kind: str, after: int = -1) -> list[_PageTable]:
    """
    Return the last contiguous run of pages of the requested *kind*.

    ``after`` restricts the search to blocks starting at or after that page
    index, which is how the Loading Summary belonging to the final Load Flow
    section is located when several study cases are appended to one PDF.
    """
    blocks: list[list[_PageTable]] = []
    for table in tables:
        if table.kind != kind or table.page_index < after:
            continue
        if blocks and table.page_index == blocks[-1][-1].page_index + 1:
            blocks[-1].append(table)
        else:
            blocks.append([table])
    if not blocks:
        return []
    return blocks[0] if after >= 0 else blocks[-1]


# --------------------------------------------------------------------------- #
# Table readers
# --------------------------------------------------------------------------- #


def _read_load_flow(block: list[_PageTable]) -> list[BusRecord]:
    """Read Bus ID / nominal kV / Voltage % (and fallback loading) in PDF order."""
    records: list[BusRecord] = []

    for table in block:
        columns = table.columns
        for line in table.lines[table.header_index + 1 :]:
            text = line_text(line)
            if _is_footnote(text):
                continue

            bus_words, kv = _split_row(line, columns)
            vmag = columns.value_in(line, "voltage_percent")

            if bus_words and kv is not None and vmag is not None:
                record = BusRecord(
                    bus_id=clean_bus_id(" ".join(w["text"] for w in bus_words)),
                    nominal_kv=kv,
                    voltage_percent=vmag,
                )
                load_mw = columns.value_in(line, "load_mw")
                if load_mw is not None:
                    record.meta["load_mw"] = load_mw
                load_mvar = columns.value_in(line, "mvar_2")
                if load_mvar is not None:
                    record.meta["load_mvar"] = load_mvar
                record.meta["page"] = table.page_index + 1
                records.append(record)
            elif records and _continuation_words(line, columns, ["kv", "voltage_percent"]):
                # tail of a wrapped Bus ID -> glue it back onto the last record
                tail = " ".join(w["text"] for w in bus_words)
                records[-1].bus_id = clean_bus_id(records[-1].bus_id + " " + tail)

    return records


def _read_loading_summary(block: list[_PageTable]) -> dict[str, dict]:
    """Read total bus load (MVA, %PF) and Amp Loading, keyed by Bus ID."""
    loading: dict[str, dict] = {}
    last_id: Optional[str] = None

    for table in block:
        columns = table.columns
        for line in table.lines[table.header_index + 1 :]:
            text = line_text(line)
            if _is_footnote(text):
                continue

            bus_words, kv = _split_row(line, columns)

            if bus_words and kv is not None:
                bus_id = clean_bus_id(" ".join(w["text"] for w in bus_words))
                loading[bus_id] = {
                    "kv": kv,
                    "rated_amp": columns.value_in(line, "rated_amp"),
                    "mva": columns.value_in(line, "total_mva"),
                    "pf": columns.value_in(line, "pf"),
                    "amp": columns.value_in(line, "amp_loading"),
                    "percent": columns.value_in(line, "percent_loading"),
                }
                last_id = bus_id
            elif last_id and _continuation_words(line, columns, ["kv"]):
                tail = " ".join(w["text"] for w in bus_words)
                new_id = clean_bus_id(last_id + " " + tail)
                loading[new_id] = loading.pop(last_id)
                last_id = new_id

    return loading


def _merge(records: list[BusRecord], loading: dict[str, dict]) -> list[BusRecord]:
    """Attach loading data to the bus records read from the Load Flow Report."""
    for record in records:
        data = loading.get(record.bus_id)
        if data:
            mva, pf = data.get("mva"), data.get("pf")
            if mva is not None and pf is not None:
                record.kw_loading = mva * (pf / 100.0) * 1000.0
                record.meta["source_kw"] = "loading summary (MVA x %PF)"
            elif mva is not None:
                record.kw_loading = mva * 1000.0
                record.meta["source_kw"] = "loading summary (MVA, PF unavailable)"
            record.amp_loading = data.get("amp")
            record.meta["rated_amp"] = data.get("rated_amp")
            record.meta["percent_loading"] = data.get("percent")
            # Kept alongside the existing fields so that the configurable
            # report can offer apparent and reactive loading without a second
            # parse.  Nothing below changes any pre-existing value.
            record.meta["mva"] = mva
            record.meta["pf"] = pf
            if mva is not None and pf is not None:
                # Q = S * sin(acos(PF)); exactly consistent with P = S * PF above.
                ratio = max(0.0, 1.0 - (pf / 100.0) ** 2)
                record.meta["kvar_loading"] = mva * math.sqrt(ratio) * 1000.0
        if record.kw_loading is None and record.meta.get("load_mw") is not None:
            record.kw_loading = record.meta["load_mw"] * 1000.0
            record.meta.setdefault("source_kw", "load flow report (directly connected load)")
        if record.meta.get("kvar_loading") is None and record.meta.get("load_mvar") is not None:
            record.meta["kvar_loading"] = record.meta["load_mvar"] * 1000.0
    return records


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #


def parse_bus_data(pdf_path: str, include_pseudo_buses: bool = False) -> list[BusRecord]:
    """
    Extract every main bus of an ETAP Load Flow report, in report order.

    Parameters
    ----------
    pdf_path:
        Path to the ETAP Load Flow Analysis report (PDF).
    include_pseudo_buses:
        When ``False`` (default) ETAP's auto-generated internal nodes
        (``Cable40~``, ``S002-TR-CMP1_VFD~2`` ...) are omitted, since they are
        not real switchboards.

    Raises
    ------
    ParserError
        With :data:`MSG_UNREADABLE` if the file cannot be opened, or
        :data:`MSG_NO_BUS_TABLE` if no bus results table can be located.
    """
    try:
        pdf = pdfplumber.open(pdf_path)
    except Exception as exc:  # pragma: no cover - depends on the file
        raise ParserError(MSG_UNREADABLE) from exc

    try:
        try:
            tables = _scan_pages(pdf)
        except Exception as exc:  # pragma: no cover
            raise ParserError(MSG_UNREADABLE) from exc

        if not tables:
            raise ParserError(MSG_NO_BUS_TABLE)

        lf_block = _last_block(tables, "load_flow")
        if not lf_block:
            raise ParserError(MSG_NO_BUS_TABLE)

        summary_block = _last_block(tables, "loading_summary", after=lf_block[0].page_index)

        records = _read_load_flow(lf_block)
        if not records:
            raise ParserError(MSG_NO_BUS_TABLE)

        loading = _read_loading_summary(summary_block) if summary_block else {}
        records = _merge(records, loading)
    finally:
        pdf.close()

    if not include_pseudo_buses:
        records = [r for r in records if not is_pseudo_bus(r.bus_id)]

    if not records:
        raise ParserError(MSG_NO_BUS_TABLE)
    return records


def has_text_layer(pdf_path: str) -> bool:
    """
    True when the PDF already contains extractable text.

    Scanned reports return ``False``; :func:`ocr_pdf` can then be used to add a
    text layer before parsing.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:5]:
                if page.extract_words():
                    return True
    except Exception:
        return False
    return False


def ocr_pdf(pdf_path: str, output_path: Optional[str] = None) -> str:
    """
    Add a text layer to a scanned report using OCRmyPDF, returning the new path.

    OCRmyPDF (and Tesseract) are optional; if they are not installed a
    :class:`utils.ParserError` is raised so the caller can show a clear message.
    """
    import shutil
    import subprocess
    import tempfile

    if shutil.which("ocrmypdf") is None:
        raise ParserError(
            "This report contains no text layer and OCR is not available. "
            "Install OCRmyPDF/Tesseract, or export the report from ETAP as a text PDF."
        )

    output_path = output_path or tempfile.mktemp(suffix="_ocr.pdf")
    try:
        subprocess.run(
            ["ocrmypdf", "--skip-text", "--optimize", "0", pdf_path, output_path],
            check=True,
            capture_output=True,
        )
    except subprocess.CalledProcessError as exc:  # pragma: no cover
        raise ParserError(MSG_UNREADABLE) from exc
    return output_path


def load_report(pdf_path: str, include_pseudo_buses: bool = False) -> list[BusRecord]:
    """
    Convenience wrapper: OCR the report first when it has no text layer.
    """
    if not has_text_layer(pdf_path):
        pdf_path = ocr_pdf(pdf_path)
    return parse_bus_data(pdf_path, include_pseudo_buses=include_pseudo_buses)
