"""
pdf_parser.py
Extracts Main Bus data (Bus ID, Nominal Voltage, Voltage %, kW Loading,
Amp Loading) from an ETAP Load Flow Analysis Report PDF.

Design notes
------------
ETAP's PDF export does NOT put table cells in reading order in the raw text
stream - simple `pdftotext`/`get_text()` extraction interleaves columns and
is unusable for a table like this. Instead we read every word's bounding
box (`page.get_text("words")`) and reconstruct rows/columns from geometry:

  1. Locate the "LOAD FLOW REPORT" section (Bus ID, kV, Voltage % Mag., ...).
     This table gives us Bus ID, Nominal kV and Voltage %.
  2. Locate the "Bus Loading Summary Report" section (Bus ID, kV, Total Bus
     Load MVA, % PF, Amp Loading). This gives us kW Loading (MVA * PF/100 *
     1000) and Amp Loading for buses that carry directly-connected load.
  3. Merge the two tables on Bus ID. Buses with no directly connected load
     (pure junction/tie buses) get kW/Amp = 0.

Column positions are located dynamically from each section's header row
(searching for known header labels), so the parser tolerates the kind of
small positional drift you see between ETAP versions/report themes, rather
than depending on fixed page numbers or hard-coded pixel coordinates.

If a page contains no extractable text at all (a scanned/rasterized PDF),
we fall back to OCR via `pytesseract` on a rendered image of that page, if
available.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import fitz  # PyMuPDF

from utils import format_nominal_voltage, mva_pf_to_kw

NUM_RE = re.compile(r"^-?\d+(\.\d+)?$")
Y_TOL = 3.0  # points; tolerance for "same printed line"


class NoBusOutputDataError(Exception):
    """Raised when the Bus Output / Load Flow Results table cannot be found."""


class UnreadablePdfError(Exception):
    """Raised when the PDF cannot be opened/read at all."""


@dataclass
class BusRecord:
    bus_id: str
    kv_raw: str
    voltage_pct: str
    mva: Optional[str] = None
    pf: Optional[str] = None
    amp: Optional[str] = None
    order: int = 0

    @property
    def nominal_display(self) -> str:
        return format_nominal_voltage(self.kv_raw)

    @property
    def kw_loading(self) -> float:
        if self.mva is None or self.pf is None:
            return 0.0
        return mva_pf_to_kw(self.mva, self.pf)

    @property
    def amp_loading(self) -> float:
        try:
            return float(self.amp) if self.amp is not None else 0.0
        except ValueError:
            return 0.0


# --------------------------------------------------------------------------- #
# Section location
# --------------------------------------------------------------------------- #
def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text)


_FOOTNOTE_STOPWORDS = {"Indicates"}


def _collect_bus_id(words, busid_max_x, lo, hi):
    """
    Collect wrapped Bus ID text between y=[lo, hi), grouped into printed
    lines. Stops early if a line looks like the report's footnote text
    (e.g. "* Indicates a voltage regulated bus...") so trailing page
    footnotes never get glued onto the last bus of a page/section.
    """
    id_tokens = [
        w for w in words
        if w[0] < busid_max_x and lo <= w[1] < hi and w[4] not in ("*", "#")
    ]
    if not id_tokens:
        return ""
    # group into printed lines (tokens within ~4pt of each other vertically)
    id_tokens.sort(key=lambda w: (w[1], w[0]))
    lines = []
    current_line = [id_tokens[0]]
    current_y = id_tokens[0][1]
    for w in id_tokens[1:]:
        if abs(w[1] - current_y) <= 4:
            current_line.append(w)
        else:
            lines.append(current_line)
            current_line = [w]
            current_y = w[1]
    lines.append(current_line)

    kept_words = []
    for line in lines:
        line_sorted = sorted(line, key=lambda w: w[0])
        if any(t[4] in _FOOTNOTE_STOPWORDS for t in line_sorted):
            break  # footnote reached - stop collecting
        kept_words.extend(line_sorted)
    return _norm(" ".join(t[4] for t in kept_words)).strip()


def _find_section(doc, start_markers, end_markers):
    """Return (start_page_idx, end_page_idx) [start, end) for a section whose
    header text matches any of start_markers, ending just before a page that
    matches any of end_markers. Markers are matched against whitespace-
    normalized page text, so they tolerate ETAP's odd word wrapping."""
    start = None
    for i in range(len(doc)):
        norm = _norm(doc[i].get_text())
        if start is None:
            if any(m in norm for m in start_markers):
                start = i
                continue
        else:
            if any(m in norm for m in end_markers):
                return start, i
    if start is not None:
        return start, len(doc)
    return None, None


# --------------------------------------------------------------------------- #
# Generic column-anchor extraction
# --------------------------------------------------------------------------- #
def _words(page):
    words = page.get_text("words")
    if words:
        return words
    # Scanned page fallback: try OCR if available
    return _ocr_words(page)


def _ocr_words(page):
    try:
        import pytesseract
        from PIL import Image
        import io
    except ImportError:
        return []
    pix = page.get_pixmap(dpi=300)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    scale = 72.0 / 300.0  # convert pixel coords back to PDF points
    words = []
    for i in range(len(data["text"])):
        text = data["text"][i].strip()
        if not text:
            continue
        x0 = data["left"][i] * scale
        y0 = data["top"][i] * scale
        x1 = (data["left"][i] + data["width"][i]) * scale
        y1 = (data["top"][i] + data["height"][i]) * scale
        words.append((x0, y0, x1, y1, text, 0, 0, 0))
    return words


def _locate_header_x(words, label, y_max=270, x_min=0, x_max=600):
    """Find the x0 of a header word (case-sensitive exact match) within the
    top portion of the first section page. Returns None if not found."""
    for w in words:
        if w[4] != label:
            continue
        if y_max is not None and w[1] > y_max:
            continue
        if x_min <= w[0] <= x_max:
            return w[0]
    return None


# --------------------------------------------------------------------------- #
# LOAD FLOW REPORT (Bus ID, kV, Voltage %)
# --------------------------------------------------------------------------- #
def _extract_load_flow_table(doc, start_page, end_page):
    # Derive column x-positions from the header row on the first section page
    header_words = _words(doc[start_page])
    kv_x = _locate_header_x(header_words, "kV", y_max=260, x_min=90, x_max=200)
    mag_x = _locate_header_x(header_words, "Mag.", y_max=260, x_min=100, x_max=220)

    if kv_x is None:
        kv_x = 120.0
    if mag_x is None:
        mag_x = kv_x + 25.0

    kv_range = (kv_x - 15, kv_x + 20)
    volt_range = (mag_x - 12, mag_x + 25)
    busid_max_x = kv_range[0] - 3

    out = {}
    order = 0
    for pno in range(start_page, end_page):
        page = doc[pno]
        words = _words(page)
        if not words:
            continue

        kv_tokens = [w for w in words if kv_range[0] <= w[0] <= kv_range[1] and NUM_RE.match(w[4])]
        anchors = []
        for kvw in kv_tokens:
            y = kvw[1]
            volt_match = next(
                (w for w in words
                 if volt_range[0] <= w[0] <= volt_range[1] and abs(w[1] - y) <= Y_TOL
                 and NUM_RE.match(w[4])),
                None,
            )
            if volt_match is not None:
                anchors.append((y, kvw[4], volt_match[4]))
        anchors.sort(key=lambda a: a[0])

        for idx, (y, kv_text, volt_text) in enumerate(anchors):
            y_next = anchors[idx + 1][0] if idx + 1 < len(anchors) else 1e9
            lo, hi = y - 4, min(y_next - 3, y + 25)
            bus_id = _collect_bus_id(words, busid_max_x, lo, hi)
            if not bus_id:
                continue
            if bus_id not in out:
                order += 1
                out[bus_id] = BusRecord(
                    bus_id=bus_id, kv_raw=kv_text, voltage_pct=volt_text, order=order
                )
    return out


# --------------------------------------------------------------------------- #
# Bus Loading Summary Report (kW Loading via MVA*PF, Amp Loading)
# --------------------------------------------------------------------------- #
def _extract_bus_loading_table(doc, start_page, end_page):
    header_words = _words(doc[start_page])
    kv_x = _locate_header_x(header_words, "kV", y_max=265, x_min=90, x_max=200)
    amp_x = _locate_header_x(header_words, "Amp", y_max=265, x_min=480, x_max=560)
    pf_x = _locate_header_x(header_words, "PF", y_max=265, x_min=460, x_max=520)
    mva_x = _locate_header_x(header_words, "MVA", y_max=265, x_min=420, x_max=480)

    if kv_x is None:
        kv_x = 128.0
    if amp_x is None:
        amp_x = 512.0
    if pf_x is None:
        pf_x = 483.0
    if mva_x is None:
        mva_x = 454.0

    kv_range = (kv_x - 20, kv_x + 20)
    amp_range = (amp_x - 10, amp_x + 55)
    pf_range = (pf_x - 15, min(pf_x + 20, amp_range[0] - 1))
    mva_range = (mva_x - 15, min(mva_x + 20, pf_range[0] - 1))
    busid_max_x = kv_range[0] - 3

    out = {}
    for pno in range(start_page, end_page):
        page = doc[pno]
        words = _words(page)
        if not words:
            continue

        kv_tokens = [w for w in words if kv_range[0] <= w[0] <= kv_range[1] and NUM_RE.match(w[4])]
        anchors = [(w[1], w[4]) for w in kv_tokens]
        anchors.sort(key=lambda a: a[0])

        for idx, (y, kv_text) in enumerate(anchors):
            y_next = anchors[idx + 1][0] if idx + 1 < len(anchors) else 1e9
            lo, hi = y - 4, min(y_next - 3, y + 25)
            bus_id = _collect_bus_id(words, busid_max_x, lo, hi)
            if not bus_id:
                continue

            row_words = [w for w in words if abs(w[1] - y) <= Y_TOL]
            amp = pf = mva = None
            for w in row_words:
                if amp_range[0] <= w[0] <= amp_range[1] and NUM_RE.match(w[4]):
                    amp = w[4]
                elif pf_range[0] <= w[0] <= pf_range[1] and NUM_RE.match(w[4]):
                    pf = w[4]
                elif mva_range[0] <= w[0] <= mva_range[1] and NUM_RE.match(w[4]):
                    mva = w[4]
            out[bus_id] = (mva, pf, amp)
    return out


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #
def extract_bus_data(pdf_path: str) -> list[BusRecord]:
    """
    Main entry point. Returns a list of BusRecord in the same order the
    buses appear in the PDF's Load Flow Report table.

    Raises:
        UnreadablePdfError    - the file could not be opened as a PDF
        NoBusOutputDataError  - no Bus Output / Load Flow Results table found
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:  # noqa: BLE001
        raise UnreadablePdfError(str(exc)) from exc

    lf_start, lf_end = _find_section(
        doc,
        start_markers=["LOAD FLOW REPORT"],
        end_markers=["Bus Loading Summary Report", "Branch Loading Summary Report"],
    )
    if lf_start is None:
        raise NoBusOutputDataError(
            "No Bus Output Data found in the uploaded ETAP report."
        )

    bus_map = _extract_load_flow_table(doc, lf_start, lf_end)
    if not bus_map:
        raise NoBusOutputDataError(
            "No Bus Output Data found in the uploaded ETAP report."
        )

    bl_start, bl_end = _find_section(
        doc,
        start_markers=["Bus Loading Summary Report"],
        end_markers=["Branch Loading Summary Report", "Branch Losses Summary Report"],
    )
    loading_map = {}
    if bl_start is not None:
        loading_map = _extract_bus_loading_table(doc, bl_start, bl_end)

    records = sorted(bus_map.values(), key=lambda r: r.order)
    for rec in records:
        if rec.bus_id in loading_map:
            rec.mva, rec.pf, rec.amp = loading_map[rec.bus_id]

    return records
