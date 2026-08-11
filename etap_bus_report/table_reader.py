"""
table_reader.py
---------------
Generic toolkit for reading ETAP's fixed-pitch PDF tables from *positioned*
words rather than flat text.

It is the generalised version of the machinery proven in ``pdf_parser.py``
(Load Flow).  New report parsers - Short Circuit today, Motor Starting or Arc
Flash tomorrow - build on this module so that only the header signature and the
row semantics have to be written for each new ETAP table.

The Load Flow parser keeps its own inlined copy on purpose: it is in production
and must not be disturbed by changes made for other report types.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Iterable, Optional

from utils import centre, cluster_lines, line_text, to_float


# --------------------------------------------------------------------------- #
# Columns
# --------------------------------------------------------------------------- #


@dataclass
class Column:
    """A table column described by its horizontal span on the page."""

    name: str
    left: float
    right: float

    def contains(self, word: dict) -> bool:
        return self.left <= centre(word) < self.right


class ColumnMap:
    """Maps positioned words onto named columns for one page."""

    def __init__(self, columns: list[Column]) -> None:
        self.columns = columns
        self._by_name: dict[str, Column] = {}
        for column in columns:
            self._by_name.setdefault(column.name, column)

    # -- lookup ------------------------------------------------------------- #

    def get(self, name: str) -> Optional[Column]:
        return self._by_name.get(name)

    def has(self, name: str) -> bool:
        return name in self._by_name

    def words_in(self, line: list[dict], name: str) -> list[dict]:
        column = self._by_name.get(name)
        if column is None:
            return []
        return [w for w in line if column.contains(w)]

    def text_in(self, line: list[dict], name: str) -> str:
        return " ".join(w["text"] for w in self.words_in(line, name)).strip()

    def value_in(self, line: list[dict], name: str) -> Optional[float]:
        """First numeric value of a column (ETAP result flags ``*``/``#`` stripped)."""
        for word in self.words_in(line, name):
            value = to_float(word["text"].rstrip("*#"))
            if value is not None:
                return value
        return None

    def flagged(self, line: list[dict], name: str) -> bool:
        """True when ETAP marked the value of a column with ``*`` (limit exceeded)."""
        return any(w["text"].endswith(("*", "#")) for w in self.words_in(line, name))


def columns_from_anchors(anchors: list[tuple[str, float]], page_width: float) -> ColumnMap:
    """
    Build a :class:`ColumnMap` from ``(name, header_centre)`` anchors.

    Boundaries sit midway between consecutive header centres; the first column
    reaches the left page edge and the last one the right page edge.
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
# Row helpers
# --------------------------------------------------------------------------- #


def normalise(text: str) -> str:
    """Lower-cased, whitespace-collapsed text, for header matching."""
    return re.sub(r"\s+", " ", text).strip().lower()


def split_leading_id(
    line: list[dict],
    columns: ColumnMap,
    anchor_column: str,
) -> tuple[list[dict], Optional[float], Optional[dict]]:
    """
    Split a data line into (ID words, anchor value, anchor word).

    ETAP does not clip long IDs, so an ID can spill out of its nominal column
    and into the next one.  The anchor column holds a right-aligned number
    (typically the nominal kV), so its *right-most* numeric word marks the end
    of the ID: everything printed to the left of it belongs to the ID.
    """
    column = columns.get(anchor_column)
    if column is None:
        return [], None, None

    candidates = [
        w for w in line
        if column.contains(w) and to_float(w["text"].rstrip("*#")) is not None
    ]
    if not candidates:
        return [], None, None

    anchor = max(candidates, key=lambda w: w["x1"])
    id_words = [w for w in line if w["x1"] <= anchor["x0"] - 0.5]
    return id_words, to_float(anchor["text"].rstrip("*#")), anchor


def is_footnote(text: str, extra: str = "") -> bool:
    """True for ETAP's table footnotes and totals lines."""
    pattern = r"^\s*[*#]?\s*(indicates|total number|this |note[:s ])"
    if extra:
        pattern = pattern[:-1] + "|" + extra + ")"
    return bool(re.match(pattern, text, re.I))


# --------------------------------------------------------------------------- #
# Page scanning
# --------------------------------------------------------------------------- #


@dataclass
class PageTable:
    """One page of a multi-page ETAP table."""

    page_index: int
    columns: ColumnMap
    lines: list[list[dict]]
    header_index: int
    context: dict


def scan_backwards(
    pdf,
    is_header: Callable[[str], bool],
    build_columns: Callable[[list[dict], list[list[dict]], int, float], ColumnMap],
    context: Optional[Callable[[list[list[dict]], int], dict]] = None,
    max_pages: Optional[int] = None,
) -> list[PageTable]:
    """
    Find the **last** contiguous block of pages carrying a given table header.

    Scanning from the end of the document is what keeps large reports fast: an
    ETAP Short Circuit study can be 350+ pages, but its summary table sits near
    the end, so only a couple of dozen pages ever get read.

    Parameters
    ----------
    pdf:
        An open ``pdfplumber.PDF``.
    is_header:
        Predicate applied to each line of text; identifies the column header.
    build_columns:
        ``(header_words, lines, header_index, page_width) -> ColumnMap``.
    context:
        Optional ``(lines, header_index) -> dict`` returning per-page context
        (for example the fault type printed above the table).
    max_pages:
        Safety limit on how many pages to inspect.

    Returns
    -------
    list[PageTable]
        The pages of the block, in document order (may be empty).
    """
    found: list[PageTable] = []
    inspected = 0

    for index in range(len(pdf.pages) - 1, -1, -1):
        if max_pages is not None and inspected >= max_pages:
            break
        inspected += 1

        page = pdf.pages[index]
        try:
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        except Exception:  # pragma: no cover - corrupt page
            words = []

        matched = False
        if words:
            lines = cluster_lines(words)
            for line_index, line in enumerate(lines):
                if is_header(line_text(line)):
                    found.append(
                        PageTable(
                            page_index=index,
                            columns=build_columns(line, lines, line_index, page.width),
                            lines=lines,
                            header_index=line_index,
                            context=context(lines, line_index) if context else {},
                        )
                    )
                    matched = True
                    break

        # The block is contiguous: once it started, the first miss ends it.
        if found and not matched:
            break

        # Free the page cache - important on 350-page reports.
        page.flush_cache()

    found.reverse()
    return found


def iter_data_lines(block: Iterable[PageTable]) -> Iterable[tuple[PageTable, list[dict]]]:
    """Yield every line below the header, page by page, in document order."""
    for table in block:
        for line in table.lines[table.header_index + 1 :]:
            yield table, line
