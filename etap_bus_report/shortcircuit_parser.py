"""
shortcircuit_parser.py
----------------------
Extraction of the bus short-circuit results from an ETAP Short Circuit Study
report.

Completely independent of the Load Flow parser: it shares only the generic
positioned-table toolkit in :mod:`table_reader`.

Which table is read
===================
ETAP's ``Short-Circuit Summary Report`` prints one table per fault type, each
page repeating the header::

    Device Capacity (kA)
    Bus            Device            Making   Short-Circuit Current (kA)
    ID   kV   ID   Type   Peak  Ib sym  Ib asym  Idc   I"k  ip  Ib sym  Ib asym  Idc  Ik

Rows come in groups: one **bus row** (``Type = Bus``) carrying the bus fault
currents, followed by one row per **protective device** on that bus carrying
that device's interrupting/making capacity and duty.

The parser is driven by that header alone - never by page numbers - and finds
the block by scanning backwards from the end of the document, so a 350-page
study is read in a couple of seconds.

Per-bus values (IEC 60909 naming)
    ``ik_initial`` I"k, ``ip_peak`` ip, ``ib_sym``/``ib_asym``, ``idc``,
    ``ik_steady`` Ik.
Per-bus device capacities (worst / lowest rated device on the bus)
    ``rating_peak``, ``rating_ib_sym``, ``rating_ib_asym``, ``rating_idc``.

``X/R`` and other quantities that appear only on the per-bus detail pages are
read on demand by :func:`add_detail_fields`, which is skipped unless the Word
template actually asks for them.
"""

from __future__ import annotations

import re
from typing import Optional

import pdfplumber

from table_reader import (
    ColumnMap,
    PageTable,
    columns_from_anchors,
    is_footnote,
    normalise,
    scan_backwards,
    split_leading_id,
)
from utils import ParserError, centre, clean_bus_id, is_pseudo_bus, line_text, to_float

# --------------------------------------------------------------------------- #
# Messages required by the specification
# --------------------------------------------------------------------------- #

MSG_NO_SC_TABLE = "No Bus Short Circuit Results table was found in the uploaded ETAP report."
MSG_UNREADABLE = "Unable to read the ETAP Short Circuit Report."

#: Fault type preferred when the report contains several summary tables.
DEFAULT_FAULT_TYPE = "3-phase"

#: How devices on one bus are combined into a single "switchgear rating".
#: ``min`` = the lowest rated device governs (the binding constraint, and what
#: ETAP flags with ``*``).  ``max`` and ``first`` are also accepted.
RATING_AGGREGATION = "min"


# --------------------------------------------------------------------------- #
# Header recognition
# --------------------------------------------------------------------------- #


def _is_summary_header(text: str) -> bool:
    """
    Header of the Short-Circuit Summary table.

    Signature: an ``ID kV ID Type`` prefix plus the IEC current symbols. Written
    tolerantly so that ETAP versions that drop or add duty columns still match.
    """
    t = normalise(text)
    if not t.startswith("id kv id type"):
        return False
    return ('i"k' in t or "i”k" in t or 'i" k' in t or "ik" in t) and (" ip " in f" {t} ")


def _fault_context(lines: list[list[dict]], header_index: int) -> dict:
    """Fault type printed above the table (``3-Phase Fault Currents`` ...)."""
    for line in reversed(lines[:header_index]):
        text = normalise(line_text(line))
        if "fault current" in text:
            if "3-phase" in text or "three" in text:
                return {"fault_type": "3-phase", "fault_title": text}
            if "line-to-ground" in text or "1-phase" in text or "lg" in text:
                return {"fault_type": "1-phase", "fault_title": text}
            return {"fault_type": text, "fault_title": text}
    return {"fault_type": DEFAULT_FAULT_TYPE, "fault_title": ""}


def _build_columns(header: list[dict], lines, header_index: int, page_width: float) -> ColumnMap:
    """
    Derive the columns from the header tokens.

    Column names are assigned by position within the two logical groups:
    the *Device Capacity* block (before ``I"k``) and the *Short-Circuit Current*
    block (from ``I"k`` onwards), which is how ETAP lays the table out.
    """
    tokens = sorted(header, key=lambda w: w["x0"])

    # Merge the two-word labels ("Ib sym", "Ib asym") into single anchors.
    merged: list[tuple[str, float]] = []
    index = 0
    while index < len(tokens):
        text = normalise(tokens[index]["text"])
        position = (tokens[index]["x0"] + tokens[index]["x1"]) / 2.0
        if text == "ib" and index + 1 < len(tokens):
            following = normalise(tokens[index + 1]["text"])
            if following in {"sym", "asym"}:
                position = (tokens[index]["x0"] + tokens[index + 1]["x1"]) / 2.0
                merged.append((f"ib {following}", position))
                index += 2
                continue
        merged.append((text, position))
        index += 1

    # Everything before I"k belongs to the device capacity block.
    capacity_names = {
        "peak": "rating_peak",
        "ib sym": "rating_ib_sym",
        "ib asym": "rating_ib_asym",
        "idc": "rating_idc",
    }
    current_names = {
        "ib sym": "ib_sym",
        "ib asym": "ib_asym",
        "idc": "idc",
        "ip": "ip_peak",
        "ik": "ik_steady",
    }

    anchors: list[tuple[str, float]] = []
    id_seen = 0
    in_currents = False

    for text, position in merged:
        if re.fullmatch(r'i\s*[”“"\']*\s*k', text) and not in_currents:
            in_currents = True
            anchors.append(("ik_initial", position))
            continue
        if text == "id":
            id_seen += 1
            anchors.append(("bus_id" if id_seen == 1 else "device_id", position))
        elif text == "kv":
            anchors.append(("kv", position))
        elif text == "type":
            anchors.append(("device_type", position))
        elif in_currents:
            anchors.append((current_names.get(text, f"current_{text}"), position))
        else:
            anchors.append((capacity_names.get(text, f"capacity_{text}"), position))

    return columns_from_anchors(anchors, page_width)


# --------------------------------------------------------------------------- #
# Row reading
# --------------------------------------------------------------------------- #

CURRENT_FIELDS = ("ik_initial", "ip_peak", "ib_sym", "ib_asym", "idc", "ik_steady")
RATING_FIELDS = ("rating_peak", "rating_ib_sym", "rating_ib_asym", "rating_idc")


def _aggregate(existing: Optional[float], value: Optional[float]) -> Optional[float]:
    """Combine device ratings according to :data:`RATING_AGGREGATION`."""
    if value is None:
        return existing
    if existing is None:
        return value
    if RATING_AGGREGATION == "max":
        return max(existing, value)
    if RATING_AGGREGATION == "first":
        return existing
    return min(existing, value)


def _device_id(line: list[dict], columns: ColumnMap, anchor: Optional[dict]) -> str:
    """
    Text of the Device ID cell.

    ETAP left-aligns device IDs directly after the kV column, well to the left
    of the ``ID`` header, so the cell is delimited by the kV value on one side
    and the Type column on the other rather than by the header centre.
    """
    type_column = columns.get("device_type")
    left = anchor["x1"] if anchor is not None else 0.0
    right = type_column.left if type_column else float("inf")
    words = [w for w in line if w["x0"] >= left - 0.5 and centre(w) < right]
    return " ".join(w["text"] for w in words).strip()


_PROSE_RE = re.compile(r"^[a-z]{2,}$")


def _is_id_tail(words: list[dict]) -> bool:
    """
    True when *words* look like the tail of a wrapped Bus ID rather than prose.

    ETAP closes the table with explanatory footnotes ("ip is calculated using
    method C", ...) whose first words land in the Bus ID column.  ID fragments
    are short and never plain lower-case words ("A]", "B]", "1 A", "BUS").
    """
    if not words or len(words) > 3:
        return False
    return not any(_PROSE_RE.match(w["text"]) for w in words)


def _read_block(block: list[PageTable]) -> list[dict]:
    """
    Read the summary block into one record per bus, in report order.

    Row grammar (ETAP prints the bus ID only once per group):

    ``<Bus ID> <kV> <Device ID> <Type> ...``
        the **bus row** - the fault currents of the bus itself.  ``Type`` is
        the bus's equipment type (``Bus``, ``Switchgear``, ``Panel``, ...);
        when the bus is modelled as switchgear the row also carries that
        assembly's rating.
    ``<kV> <Device ID> <Type> ...``
        a **device row** - one protective device on the same bus, with its
        capacity and its duty.
    ``<ID tail>``
        the continuation of a bus ID wrapped onto a second line.
    """
    buses: list[dict] = []
    current: Optional[dict] = None

    for table in block:
        columns = table.columns
        first_row_of_page = True
        for line in table.lines[table.header_index + 1 :]:
            text = line_text(line)
            if is_footnote(text):
                continue

            id_words, kv, anchor = split_leading_id(line, columns, "kv")

            if kv is None:
                # No kV on this line -> either the tail of a wrapped ID or one
                # of the explanatory footnotes printed under the table.
                tail_words = columns.words_in(line, "bus_id")
                if current is not None and current.get("_allow_tail") and _is_id_tail(tail_words):
                    tail = " ".join(w["text"] for w in tail_words)
                    current["bus_id"] = clean_bus_id(current["bus_id"] + " " + tail)
                continue

            device_type = normalise(columns.text_in(line, "device_type"))

            if id_words:
                bus_id = clean_bus_id(" ".join(w["text"] for w in id_words))
                # On a page break ETAP reprints the Bus ID on the first row of
                # the continued group - that is not a new bus.  The prefix test
                # catches the case where the repeated ID is itself wrapped, so
                # only its first line is printed here.
                repeated = (
                    current is not None
                    and first_row_of_page
                    and (bus_id == current["bus_id"] or current["bus_id"].startswith(bus_id + " "))
                )
            else:
                bus_id, repeated = "", False
            first_row_of_page = False

            if id_words and not repeated:
                # ---- bus row ------------------------------------------------
                current = {
                    "bus_id": bus_id,
                    "nominal_kv": kv,
                    "bus_type": device_type.title(),
                    "page": table.page_index + 1,
                    "fault_type": table.context.get("fault_type", ""),
                    "devices": [],
                    # A wrapped ID always continues on the very next line, i.e.
                    # before any device row of the same group.
                    "_allow_tail": True,
                }
                for field in CURRENT_FIELDS:
                    current[field] = columns.value_in(line, field)
                # A bus modelled as switchgear/panel carries its own rating.
                for field in RATING_FIELDS:
                    current[field] = columns.value_in(line, field)
                current["bus_rating_peak"] = current.get("rating_peak")
                buses.append(current)
                continue

            # ---- device row --------------------------------------------- #
            if current is None:
                continue
            current["_allow_tail"] = False
            device = {
                "device_id": _device_id(line, columns, anchor),
                "device_type": device_type.upper(),
                "exceeded": any(columns.flagged(line, f) for f in CURRENT_FIELDS),
            }
            for field in RATING_FIELDS:
                device[field] = columns.value_in(line, field)
            current["devices"].append(device)

    for bus in buses:
        _finalise_ratings(bus)
        bus.pop("_allow_tail", None)

    return buses


def _finalise_ratings(bus: dict) -> None:
    """
    Decide the single rating reported for the bus.

    The bus's own assembly rating wins when ETAP printed one (the bus is
    modelled as switchgear); otherwise the protective devices on the bus are
    combined per :data:`RATING_AGGREGATION` - by default the *lowest* rated
    device, because that is the binding constraint and the one ETAP flags
    with ``*``.
    """
    governing = _governing_device(bus)

    if bus.get("bus_rating_peak") is None:
        for field in RATING_FIELDS:
            value = None
            for device in bus["devices"]:
                value = _aggregate(value, device.get(field))
            bus[field] = value
        bus["governing_device"] = governing["device_id"] if governing else ""
    else:
        bus["governing_device"] = bus["bus_id"]

    bus["device_exceeded"] = any(d["exceeded"] for d in bus["devices"])


def _governing_device(bus: dict) -> Optional[dict]:
    """The device whose peak rating was kept for the bus (lowest by default)."""
    rated = [d for d in bus["devices"] if d.get("rating_peak") is not None]
    if not rated:
        return None
    if RATING_AGGREGATION == "max":
        return max(rated, key=lambda d: d["rating_peak"])
    if RATING_AGGREGATION == "first":
        return rated[0]
    return min(rated, key=lambda d: d["rating_peak"])


# --------------------------------------------------------------------------- #
# Optional detail pages (X/R and friends)
# --------------------------------------------------------------------------- #

_DETAIL_BUS_RE = re.compile(r"fault at bus\s*:?", re.I)
_CFACTOR_RE = re.compile(r"Voltage\s+c\s+Factor\s*=\s*([\d.]+)", re.I)


# --------------------------------------------------------------------------- #
# Study metadata (standard, study type, study case)
# --------------------------------------------------------------------------- #

_STANDARD_RE = re.compile(r"\b(IEC\s*60909|IEC|ANSI(?:/IEEE)?)\b", re.I)
_STUDY_TYPE_RE = re.compile(
    r"(3-?Phase|Three-?Phase|1-?Phase|Single-?Phase|Line-?to-?Ground|LG)\s+Fault", re.I
)
_STUDY_CASE_RE = re.compile(r"Study\s*Case\s*:\s*(\S+)", re.I)


def study_info(pdf_path: str) -> dict:
    """
    Read the study header: which standard, which fault type, which study case.

    All three are printed on the report's first page; nothing is guessed. Keys
    that the report does not state are returned empty.
    """
    info = {"standard": "", "study_type": "", "study_case": "", "title": ""}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return info
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return info

    match = _STANDARD_RE.search(text)
    if match:
        found = match.group(1).upper().replace(" ", "")
        info["standard"] = "IEC" if found.startswith("IEC") else "ANSI"
        info["title"] = match.group(0).strip()

    match = _STUDY_TYPE_RE.search(text)
    if match:
        found = match.group(1).lower().replace("-", "").replace(" ", "")
        info["study_type"] = (
            "3-phase" if found in {"3phase", "threephase"} else "1-phase"
        )

    match = _STUDY_CASE_RE.search(text)
    if match:
        info["study_case"] = match.group(1).strip()
    return info


def add_detail_fields(pdf_path: str, buses: list[dict]) -> list[dict]:
    """
    Add ``cfactor`` and ``xr_ratio`` from ETAP's per-bus detail pages.

    Each ``SHORT-CIRCUIT REPORT`` page prints the IEC voltage factor
    (``Voltage c Factor = 1.10``) and, on its ``Total`` contribution line, the
    X/R ratio of the whole fault contribution.

    The faulted bus is identified from that ``Total`` line rather than from the
    page heading: ETAP wraps long IDs across two lines in the heading, but the
    contribution table always prints the full ID.

    Only the top of each page is read, and the walk stops as soon as every bus
    has both values - this is still a whole-document scan, so it runs only when
    a selected column needs it.
    """
    wanted = {bus["bus_id"]: bus for bus in buses}
    outstanding = set(wanted)
    try:
        pdf = pdfplumber.open(pdf_path)
    except Exception as exc:
        raise ParserError(MSG_UNREADABLE) from exc

    try:
        for page in pdf.pages:
            if not outstanding:
                break
            try:
                # The heading, the c factor and the Total line all sit in the
                # top half; cropping roughly halves the extraction cost.
                text = page.crop((0, 0, page.width, page.height * 0.55)).extract_text() or ""
            except Exception:
                text = ""
            if "fault at bus" not in text.lower():
                page.flush_cache()
                continue

            cfactor = None
            match = _CFACTOR_RE.search(text)
            if match:
                cfactor = to_float(match.group(1))

            for line in text.split("\n"):
                if " Total " not in f" {line} ":
                    continue
                bus_id = clean_bus_id(line.split(" Total ")[0])
                bus = wanted.get(bus_id)
                if bus is None:
                    break
                numbers = [to_float(t) for t in line.split()]
                numbers = [n for n in numbers if n is not None]
                if bus.get("xr_ratio") is None and len(numbers) >= 2:
                    # "<ID> Total  %V  Real  Imaginary  X/R  Magnitude"
                    bus["xr_ratio"] = numbers[-2]
                if bus.get("cfactor") is None and cfactor is not None:
                    bus["cfactor"] = cfactor
                outstanding.discard(bus_id)
                break
            page.flush_cache()
    finally:
        pdf.close()
    return buses


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #


def parse_short_circuit(
    pdf_path: str,
    include_pseudo_buses: bool = False,
    fault_type: str = DEFAULT_FAULT_TYPE,
) -> list[dict]:
    """
    Extract every main bus of an ETAP Short Circuit report, in report order.

    Returns a list of plain dictionaries so that new fields can be added
    without touching the writer; the writer picks whatever the template asks
    for.

    Raises
    ------
    ParserError
        :data:`MSG_UNREADABLE` if the file cannot be opened,
        :data:`MSG_NO_SC_TABLE` if no summary table can be located.
    """
    try:
        pdf = pdfplumber.open(pdf_path)
    except Exception as exc:
        raise ParserError(MSG_UNREADABLE) from exc

    try:
        try:
            block = scan_backwards(
                pdf,
                is_header=_is_summary_header,
                build_columns=_build_columns,
                context=_fault_context,
            )
        except Exception as exc:  # pragma: no cover
            raise ParserError(MSG_UNREADABLE) from exc

        if not block:
            raise ParserError(MSG_NO_SC_TABLE)

        # Prefer the requested fault type when the report holds several tables.
        wanted = [t for t in block if t.context.get("fault_type") == fault_type]
        buses = _read_block(wanted or block)
    finally:
        pdf.close()

    if not include_pseudo_buses:
        buses = [b for b in buses if not is_pseudo_bus(b["bus_id"])]

    if not buses:
        raise ParserError(MSG_NO_SC_TABLE)
    return buses


def has_text_layer(pdf_path: str) -> bool:
    """True when the PDF already contains extractable text."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:5]:
                if page.extract_words():
                    return True
    except Exception:
        return False
    return False


def load_report(pdf_path: str, include_pseudo_buses: bool = False, **kwargs) -> list[dict]:
    """OCR the report first when it has no text layer, then parse it."""
    if not has_text_layer(pdf_path):
        import pdf_parser  # reuse the shared OCR helper

        pdf_path = pdf_parser.ocr_pdf(pdf_path)
    return parse_short_circuit(pdf_path, include_pseudo_buses=include_pseudo_buses, **kwargs)
