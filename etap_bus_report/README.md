# ETAP Bus Report Generator

Desktop application that reads an **ETAP Load Flow Analysis report (PDF)** and fills the
standard **Bus** table in a bundled Word template, flagging every bus against a
user-entered acceptable voltage band.

```
Upload ETAP LFA report (.pdf)  ─┐
Acceptable voltage limits       ─┴─►  Generate Report  ──►  Bus Report.docx
```

---

There are three front ends over one shared pipeline — pick whichever fits.

## 1. Install and run

**Desktop (PySide6)** — needs a machine with a display:

```bash
python -m venv .venv
.venv\Scripts\activate                    # Windows (source .venv/bin/activate elsewhere)
pip install -r requirements-desktop.txt
python main.py
```

**Web (Streamlit)** — works locally and on Streamlit Community Cloud:

```bash
pip install -r requirements.txt
streamlit run etap_bus_report/streamlit_app.py
```

**Command line** (batch runs, CI):

```bash
python cli.py "Umm Qasr-ST LFA.pdf" --limits 95-106 --output "Bus Report.docx"
```

### Deploying to Streamlit Community Cloud

In the app settings set:

| Setting | Value |
|---|---|
| Main file path | `etap_bus_report/streamlit_app.py` |
| Requirements | `etap_bus_report/requirements.txt` (auto-detected) |

**Do not put PySide6 in `requirements.txt`.** Streamlit Cloud is a headless
container with no display and no GUI system libraries, so importing PySide6
fails with `libglib-2.0.so.0: cannot open shared object file` — which is why
`main.py` (the *desktop* entry point) cannot be the cloud main module. PySide6
lives in `requirements-desktop.txt` only.

Tests:

```bash
pytest tests.py          # or: python tests.py
```

---

## 2. Project structure

```
main.py             desktop entry point (launches the GUI)
ui.py               PySide6 window: file picker, limits field, Generate / Reset
streamlit_app.py    web front end (Streamlit Cloud / any headless server)
cli.py              headless front end (batch runs, CI)
pdf_parser.py       ETAP PDF -> list[BusRecord]
voltage_checker.py  limit parsing ("95 - 105 %") and ACCEPTABLE / NOT ACCEPTABLE
word_writer.py      fills the bundled template, preserving all formatting
utils.py            data model, number/voltage formatting, geometry helpers, config
tests.py            regression tests (incl. the real 283-bus sample report)
assets/
    Bus_Template.docx   fixed template, bundled - the user never uploads it
samples/
    sample_lfa_report.pdf
requirements.txt            web / Streamlit Cloud (no PySide6)
requirements-desktop.txt    desktop build (adds PySide6)
build.spec          PyInstaller configuration
```

---

## 3. How the PDF is read

**No fixed page numbers are used anywhere.** The parser reads *positioned words*
(pdfplumber) and looks for the repeating ETAP column headers:

| Header signature found on a page | Table |
|---|---|
| `ID kV % Mag. Ang. MW Mvar MW Mvar ID MW Mvar Amp %PF %Tap` | Bus load-flow results (`LOAD FLOW REPORT` / `Bus Output Data`) |
| `ID kV Rated Amp … MVA %PF Amp Loading` | `Bus Loading Summary Report` |

Consecutive pages carrying the same header form one section; if a PDF contains
several study cases, the **last** load-flow section is used, together with the
loading summary that follows it. The `Bus Input Data` table is never used — its
header lacks the `Amp` / `%PF` result columns, which is exactly what tells the
two apart.

Column boundaries are computed per page from the header token positions, so
column shifts between ETAP versions are handled automatically. Two ETAP quirks
are handled explicitly:

* **Wrapped IDs** — `S002-SB-PSS01N01 [BUS` / `A]` is reassembled into
  `S002-SB-PSS01N01 [BUS A]`.
* **Overflowing IDs** — `TYPICAL MOTOR FEEDER 1 B` spills past its column, so
  the nominal kV is taken as the right-most number of the kV band and everything
  left of it is treated as the Bus ID.

Scanned PDFs: `pdf_parser.load_report()` detects a missing text layer and runs
OCRmyPDF automatically when it is installed (`pip install ocrmypdf`, plus
Tesseract); otherwise it reports that OCR is unavailable.

### Where each column comes from

| Word column | Source |
|---|---|
| Bus ID | Load Flow Report, in report order |
| Nominal (kV, A) | Load Flow Report `kV` — printed as `132 kV`, `3.3 kV`, `415 V`, `380 V` |
| Voltage % | Load Flow Report `% Mag.` |
| kW Loading | Bus Loading Summary: `Total Bus Load MVA × %PF × 1000` |
| Amp Loading | Bus Loading Summary: `Amp Loading` |
| Remarks | `ACCEPTABLE` / `NOT ACCEPTABLE` against the entered band |

**Why the loading summary and not the load-flow "Load MW"?** In the Load Flow
Report the `Load MW` column is only the load connected *directly* to the bus, so
a switchboard that distributes everything through outgoing feeders reads 0 MW.
`S002-SB-1AC001 [BUS A]`, for example, reads 0 MW there but 31.398 MVA @ 91.0 %
(= 28 572 kW, 1627.6 A) in the Bus Loading Summary — the figure that belongs in
a bus loading table. If a report has no loading summary, the parser falls back
to the directly-connected load and the branch amps and records that fact in
`BusRecord.meta["source_kw"]`.

### Two things worth confirming for your report

1. **"Main bus" definition.** Every bus row of the results table is written,
   except ETAP's auto-generated internal nodes whose ID contains `~`
   (`Cable40~`, `S002-TR-CMP1_VFD~2`, …) — these are cable/VFD terminals, not
   switchgear. For the sample report that is 255 rows out of 283. Change
   `PSEUDO_BUS_PATTERNS` in `utils.py` (set it to `[]`) if you want all 283, or
   add your own patterns to filter down to switchboards only.
2. **Rounding.** Voltage % is written to 1 decimal and the loadings to 1 decimal
   (`VOLTAGE_DECIMALS` / `LOADING_DECIMALS` in `utils.py`). The limit comparison
   always uses the full-precision value from ETAP.

---

## 4. Error handling

| Condition | Message |
|---|---|
| No bus results table in the PDF | `No Bus Output Data found in the uploaded ETAP report.` |
| Limits not a `Lower-Upper` pair (`95`, `95%`, `abc`, `106-95`) | `Please enter voltage limits in the format Lower-Upper (Example: 95-106).` |
| PDF cannot be opened / has no text and OCR unavailable | `Unable to read the ETAP report.` |
| Output file open in Word | Prompt to close it and retry |

Accepted limit formats: `95-106`, `95 - 105 %`, `90-110`, `95%-106%`,
`95 to 106`, en/em dashes. Both bounds are inclusive.

---

## 5. Performance

The 74-page / 283-bus sample report is parsed and written in ≈ 7 s on a laptop
(single pass over the pages, one `extract_words()` call per page). Parsing runs
in a worker thread, so the window stays responsive.

---

## 6. Building the executable

```bash
pip install -r requirements-desktop.txt pyinstaller
pyinstaller build.spec
# -> dist/ETAP Bus Report Generator.exe   (single file, windowed)
```

`build.spec` bundles `assets/Bus_Template.docx` as data; `utils.resource_path()`
resolves it both from source and from the frozen bundle, so the packaged
application carries its own template.

---

## 7. Extending it

The pipeline is deliberately split so that new report types plug in without
touching the UI:

* **Other ETAP tables** (short circuit, cable schedule, transformer report) —
  add a header signature plus a reader in `pdf_parser.py`; the column-mapping
  helpers (`ColumnMap`, `_columns_from_anchors`, `_split_row`) are generic.
* **Excel / CSV export** — `word_writer.export_rows()` already returns the
  formatted rows.
* **User-configurable templates** — `write_bus_report(..., template=...)`
  accepts any `.docx` whose first 6-column table is the bus table.
* **Extra columns** (rated amp, % loading, power factor) — already captured in
  `BusRecord.meta`, just add them to `BusRecord.as_row()`.
