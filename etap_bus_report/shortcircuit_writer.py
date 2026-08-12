"""
shortcircuit_writer.py
----------------------
Fills the bundled Short Circuit Word template.

Nothing about the template's columns is hard-coded: :mod:`template_mapping`
reads the header of the table in the template and decides which parsed field
belongs in which column.  Columns whose header matches nothing are left exactly
as the template has them.

As with the Load Flow writer, the template is opened and its existing table is
populated cell by cell - never re-created - so fonts, borders, shading, column
widths and spacing are preserved.
"""

from __future__ import annotations

import io
import os
from typing import Optional, Sequence

from docx import Document

import sc_rules
import template_mapping
from utils import resource_path
from word_writer import TemplateError, append_row, find_table, remove_row, set_cell_text

#: Bundled template for this report type.
TEMPLATE = os.path.join("assets", "ShortCircuit_Template.docx")


def template_path() -> str:
    """Absolute path of the bundled Short Circuit template."""
    return resource_path(TEMPLATE)


def _resize(table, data_rows: int, header_rows: int) -> None:
    """Make the table hold exactly *data_rows* data rows below its header."""
    current = len(table.rows) - header_rows
    if current <= 0:
        raise TemplateError("The Short Circuit template has no blank data row to copy.")

    prototype = table.rows[header_rows]
    while current < data_rows:
        append_row(table, prototype)
        current += 1
    while current > data_rows:
        remove_row(table, table.rows[-1])
        current -= 1


def build_document(
    buses: Sequence[dict],
    template: Optional[str] = None,
    rule: Optional[str] = None,
) -> Document:
    """
    Return the populated document.

    Parameters
    ----------
    buses:
        Records from :func:`shortcircuit_parser.parse_short_circuit`, in the
        order they must appear.
    template:
        Optional override of the bundled template.
    rule:
        Name of the Remarks rule (see :mod:`sc_rules`); ignored when the
        template has no Remarks column.
    """
    template = template or template_path()
    if not os.path.exists(template):
        raise TemplateError(f"Short Circuit template not found: {template}")

    document = Document(template)
    table = find_table(document)
    header_rows, field_keys, _ = template_mapping.map_template(table)

    if "remarks" in field_keys:
        sc_rules.apply_rule(buses, rule)

    _resize(table, len(buses), header_rows)

    for row, bus in zip(table.rows[header_rows:], buses):
        values = template_mapping.format_row(bus, field_keys)
        for cell, value in zip(row.cells, values):
            if value is not None:
                set_cell_text(cell, value)

    return document


def build_document_bytes(buses: Sequence[dict], template: Optional[str] = None, rule: Optional[str] = None) -> bytes:
    """Populated document as raw ``.docx`` bytes (downloads / in-memory use)."""
    buffer = io.BytesIO()
    build_document(buses, template, rule).save(buffer)
    return buffer.getvalue()


def write_short_circuit_report(
    buses: Sequence[dict],
    output_path: str,
    template: Optional[str] = None,
    rule: Optional[str] = None,
) -> str:
    """Fill the template and save the document to *output_path*."""
    document = build_document(buses, template, rule)

    folder = os.path.dirname(os.path.abspath(output_path))
    if folder:
        os.makedirs(folder, exist_ok=True)
    document.save(output_path)
    return output_path


def describe_mapping(template: Optional[str] = None) -> list[tuple[str, Optional[str]]]:
    """
    ``(header text, field key)`` for every column of the template.

    Handy when adapting a new template: it shows at a glance which columns the
    application recognised and which it will leave blank.
    """
    document = Document(template or template_path())
    table = find_table(document)
    header_rows, field_keys, headers = template_mapping.map_template(table)
    return list(zip(headers, field_keys))
