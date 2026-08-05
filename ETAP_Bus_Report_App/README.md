# ETAP Bus Load Flow Report Generator

Reads an ETAP Load Flow Analysis Report (PDF), extracts every Main Bus's
ID, Nominal Voltage, Voltage %, kW Loading and Amp Loading, checks each
bus against user-entered acceptable voltage limits, and writes the result
into the bundled `Bus_Template.docx` — without altering the template's
fonts, borders, spacing or column widths.

## Quick start (GUI)

```bash
pip install -r requirements.txt
python ui.py
```

1. Click **Choose PDF…** and select the ETAP report.
2. Type the acceptable voltage limits, e.g. `95-106`.
3. Click **Generate Report** and pick a save location.

## Quick start (web / Streamlit)

`ui.py` is a PySide6 **desktop** GUI - it will not run on Streamlit Cloud
or any other headless host (no display). Deploy `streamlit_app.py`
instead, which wraps the same `generate_report()` pipeline in a browser UI:

```bash
pip install -r requirements-streamlit.txt
streamlit run streamlit_app.py
```

On Streamlit Community Cloud: set the **main file path** to
`streamlit_app.py` (not `main.py` - running `main.py` directly under
Streamlit triggers its argparse CLI instead of a UI, which is where the
`error: the following arguments are required: --pdf, --limits` message
comes from) and, if the platform lets you choose a requirements file, use
`requirements-streamlit.txt` — it skips PySide6/PyInstaller, which are
desktop-only and can fail to install in a slim cloud container.

## Quick start (headless / CLI / automation)

```bash
python main.py --pdf "Table_7_LFA_CASE_1.pdf" --limits "95-106" --out "Bus Report.docx"
```

Exit codes: `0` success · `1` bad voltage limits · `2` unreadable PDF ·
`3` no Bus Output Data found.

## How it works

```
main.py         Orchestrates the pipeline; used by ui.py, streamlit_app.py, and the CLI.
ui.py           PySide6 desktop GUI (2 inputs, Generate/Reset, save dialog).
streamlit_app.py    Browser-based UI for Streamlit Cloud / any Streamlit host.
pdf_parser.py   Extracts Bus ID / kV / Voltage % / kW / Amp from the PDF.
word_writer.py  Populates assets/Bus_Template.docx, preserving formatting.
voltage_checker.py   ACCEPTABLE / NOT ACCEPTABLE comparison.
utils.py        Voltage formatting + "Lower-Upper" limits parsing.
assets/Bus_Template.docx   The fixed, bundled Word template.
```

### PDF parsing strategy (important)

ETAP's PDF export does **not** put table cells in reading order in the
underlying text stream — a naive `pdftotext`/`get_text()` call interleaves
columns and produces unusable output for a table like this (values from
one column appear jumbled between values from another). This is true even
though the report renders as a clean table visually.

`pdf_parser.py` therefore reads every word's **bounding box**
(`page.get_text("words")`, via PyMuPDF) and reconstructs rows and columns
from geometry instead of trusting stream order:

1. Locate the **"LOAD FLOW REPORT"** section. Its header row's `kV` and
   `Mag.` (Voltage %) labels are located dynamically to derive column
   x-ranges (rather than hard-coding pixel coordinates), so the parser
   tolerates the column drift you tend to see between ETAP versions/report
   themes. Any line whose kV-column and Voltage%-column both contain a
   number, on the same y-coordinate, marks the start of a new bus record.
   Everything left of the kV column, from that line up to the next bus's
   anchor line, is the Bus ID — this correctly re-assembles Bus IDs that
   wrap across two printed lines (e.g. `TYPICAL MOTOR FEEDER 1 A`) and
   IDs containing brackets (e.g. `S002-SB-1AC001 [BUS A]`), while ignoring
   the many branch/destination rows that follow each bus (a bus can have
   dozens of outgoing connections listed underneath it).
2. Locate the **"Bus Loading Summary Report"** section the same way, using
   its `kV`, `MVA`, `PF` and `Amp` (Amp Loading) header labels. This table
   gives the literal **Amp Loading** figure, and kW Loading is computed as
   `MVA × PF/100 × 1000`. Buses with no *directly connected* load (pure
   junction/tie buses, and buses whose load sits on a downstream VFD
   rather than the bus itself) simply don't appear in this table and are
   reported as 0 kW / 0 A — this is a judgment call worth reviewing
   against how your team reads these reports; see **Known limitation**
   below.
3. The two tables are merged on Bus ID, and Bus ID text tokens `*`/`#`
   (ETAP's swing-bus / mismatch markers) and page-footer footnote text
   ("* Indicates a voltage regulated bus…") are explicitly filtered out
   so they never get glued onto a bus name.

This was validated end-to-end against a real 74-page / 283-bus ETAP
report: **283/283 buses extracted, zero duplicates, zero misses**,
matching the report's own "Number of Buses: 283" header — including every
multi-line-wrapped and bracketed Bus ID.

If a page has no extractable text at all (a scanned/rasterized PDF),
`pdf_parser.py` falls back to OCR via `pytesseract` if it's installed;
otherwise that page is skipped rather than crashing the whole run.

### Known limitation

The spec's five target fields (Bus ID, Nominal Voltage, Voltage %, kW
Loading, Amp Loading) don't all live in one single ETAP table — the
sample report split them across two tables ("LOAD FLOW REPORT" for
voltage, "Bus Loading Summary Report" for loading). If your organization's
standard export template differs (e.g. a single combined "Bus Output
Data" table, which some ETAP configurations produce), the column-locating
logic in `_extract_load_flow_table` / `_extract_bus_loading_table` in
`pdf_parser.py` is the place to adjust — the header-label-driven approach
should make that a small, localized change rather than a rewrite.

### Word template population

`word_writer.py` only ever writes text into existing table cells — it
never recreates the table, so all of the template's fonts/borders/shading/
column widths are untouched. The template ships with a number of
pre-built empty rows; if a report has more buses than that, additional
rows are appended by **deep-copying the template's own row XML**, so
cloned rows are pixel-identical to the original ones (verified visually
by rendering the output to PDF/JPEG — see `page-*.jpg` if you want to
compare rows 1 and 283 yourself).

### Performance

The 283-bus / 74-page sample report processes in **~0.4 seconds**, well
within the 300-bus / 10-second target.

## Error handling

| Condition | Message |
|---|---|
| No Bus Output Data table found | `No Bus Output Data found in the uploaded ETAP report.` |
| Voltage limits don't match `Lower-Upper` | `Please enter voltage limits in the format Lower-Upper (Example: 95-106).` |
| PDF can't be opened/read | `Unable to read the ETAP report.` |

## Building a standalone executable

```bash
pip install -r requirements.txt
pyinstaller build.spec
```

Output is written to `dist/ETAP-Bus-Report-Generator/`. The Word template
is bundled automatically (see `build.spec`'s `datas=[...]`); `main.py`
resolves the template path correctly both when run from source and when
frozen (via `sys._MEIPASS`).

## Future expandability

The module boundaries were kept deliberately narrow so new report types
can be added without touching the others:
- **Short Circuit / Cable Schedule / Transformer reports**: add a sibling
  of `pdf_parser.py` (e.g. `sc_parser.py`) reusing the same word-bbox
  clustering helpers, and a sibling of `word_writer.py` for its template.
- **Excel export**: `main.py`'s `rows` list (plain dicts) can be handed to
  `pandas.DataFrame(rows).to_excel(...)` directly.
- **User-configurable templates**: `word_writer.populate_template()`
  already takes `template_path` as a parameter — `ui.py` would just need
  a third (optional) input to override `DEFAULT_TEMPLATE`.
