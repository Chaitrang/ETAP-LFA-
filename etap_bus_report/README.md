# ETAP Report Filler

Reads an **ETAP study report (PDF)** and fills the matching standard Word table.

| Report type | Inputs | Output |
|---|---|---|
| **Load Flow Analysis** | Load Flow report + selected columns + alert limits | `LoadFlow_Report.docx` / `.xlsx` |
| **Short Circuit Study** | Short Circuit report + selected columns + thresholds | `ShortCircuit_<case>.docx` / `.xlsx` |

Both reports build their table from the columns you select and are generated
dynamically — **no template**. There are three front ends over one shared
pipeline. The original template-filling Short Circuit report is still available
(it is what `reports.generate("short_circuit", pdf)` returns without a config),
so nothing that depended on it broke.

---

## 1. Install and run

**Web (Streamlit)** — works locally and on Streamlit Community Cloud:

```bash
pip install -r requirements.txt
streamlit run etap_bus_report/main.py        # or streamlit_app.py - both work
```

**Desktop (PySide6)** — needs a machine with a display:

```bash
pip install -r requirements-desktop.txt
python main.py
```

**Command line** (batch runs, CI):

```bash
python cli.py "LFA.pdf" --type load_flow --limits 95-106 -o "Bus Report.docx"
python cli.py "SCA.pdf" --type short_circuit -o "Short Circuit Report.docx"
python cli.py --show-mapping                 # how the SC template was mapped
```

Tests: `pytest tests.py` (or `python tests.py`).

### Deploying to Streamlit Community Cloud

| Setting | Value |
|---|---|
| Main file path | `etap_bus_report/main.py` (or `etap_bus_report/streamlit_app.py`) |
| Requirements | `etap_bus_report/requirements.txt` |

`main.py` detects the Streamlit runtime and renders the web UI instead of the
Qt window, so the same file serves both. **Do not put PySide6 in
`requirements.txt`** — Streamlit Cloud is headless and importing it fails with
`libglib-2.0.so.0: cannot open shared object file`. It belongs in
`requirements-desktop.txt` only.

---

## 2. Project structure

```
main.py                 entry point - desktop GUI locally, web UI under Streamlit
ui.py                   PySide6 window (report selector + per-report inputs)
streamlit_app.py        web front end
cli.py                  batch front end
reports.py              registry: report type -> parser + writer + template + inputs

pdf_parser.py           Load Flow parser (shared by every Load Flow front end)
loadflow_fields.py      catalogue of every selectable column
loadflow_config.py      user selections, alert limits, validation
loadflow_processor.py   classification -> the one final table
dynamic_word_writer.py  Word report built from any headers/rows (no template)
voltage_checker.py      legacy Lower-Upper band parsing (still used by --limits)
word_writer.py          template-filling writer, kept for the Short Circuit path

table_reader.py         generic positioned-table toolkit for new parsers
shortcircuit_parser.py  Short Circuit parser (+ study metadata, detail pages)
shortcircuit_fields.py  catalogue of every selectable Short Circuit column
shortcircuit_config.py  study, selections, thresholds, filters, validation
shortcircuit_processor.py  classification and filtering -> the one final table
shortcircuit_writer.py  original template-driven writer, still available
template_mapping.py     maps Word column headers -> canonical data fields
sc_rules.py             pluggable Remarks rules for Short Circuit
excel_writer.py         .xlsx export, shared by every report type

utils.py                data model, formatting, config constants
tests.py                regression tests against both real sample reports
assets/
    LoadFlow_Template.docx
    ShortCircuit_Template.docx
samples/
requirements.txt / requirements-desktop.txt / build.spec
```

Each parser is independent: the Short Circuit module shares only the generic
geometry helpers in `table_reader.py`. The Load Flow modules were **not**
modified beyond exposing four table helpers publicly for reuse.

---

## 3. How the PDFs are read

**No fixed page numbers anywhere.** Both parsers read *positioned words*
(pdfplumber) and locate their table by its repeating column header.

### Load Flow

`ID kV % Mag. Ang. … Amp %PF %Tap` identifies the bus results,
`ID kV Rated Amp … MVA %PF Amp Loading` the Bus Loading Summary. Voltage %,
Amp Loading, Amp Rating and % Loading are read straight out of them.

**Where kW and kvar come from.** ETAP prints the same quantity in two places at
different precision, so the parser takes the more precise one:

| Source | Precision | Used |
|---|---|---|
| Load Flow Report: directly connected load + outgoing branch flows | 3 decimals of MW → ±0.5 kW at any size | **preferred** |
| Bus Loading Summary: total bus load MVA × %PF | %PF is printed to 1 decimal → error grows with bus size (≈9 kW on a 56 MW bus) | fallback |

The flow figure is used only when it agrees with the Loading Summary to within
1 %, which confirms the two describe the same quantity; otherwise the summary
value stands. Measured against ETAP's own results grid, the worst error over
the sample report is **0.4 kW** (it was 9.4 kW when derived from %PF alone).

The Load Flow Report's `Load MW` column is only the load connected *directly*
to a bus — zero for a switchboard that feeds everything onward — which is why
the outgoing flows have to be added to it.

### Short Circuit

The `Short-Circuit Summary Report` header is the anchor:

```
Bus            Device            Making   Short-Circuit Current (kA)
ID   kV   ID   Type   Peak  Ib sym  Ib asym  Idc   I"k  ip  Ib sym  Ib asym  Idc  Ik
```

The block is found by scanning **backwards** from the last page, so a 355-page
study is parsed in about 2.5 s instead of 30.

Row grammar (ETAP prints the Bus ID once per group):

* `<Bus ID> <kV> <Device ID> <Type> …` — the **bus row**, carrying the bus fault
  currents. `Type` is the bus's equipment type (`Bus`, `Switchgear`,
  `Switchboard`, `MCC`, `Panelboard`, `Switchrack`); when the bus is modelled as
  an assembly the row also carries that assembly's rating.
* `<kV> <Device ID> <Type> …` — a **device row**: one protective device on the
  same bus, with its capacity and its duty.

Handled explicitly, all of which occur in the sample report: bus IDs wrapped
over two lines, IDs that overflow their column, the Bus ID being reprinted on
the first row after a page break (not a second bus), and the explanatory
footnotes under the table whose first words land in the ID column.

Extracted per bus: `ik_initial` (I"k), `ip_peak` (ip), `ib_sym`, `ib_asym`,
`idc`, `ik_steady` (Ik), plus the equipment ratings `rating_peak`,
`rating_ib_sym`, `rating_ib_asym`, `rating_idc`, the governing device and
ETAP's own `*` duty flags. `X/R` is read from the per-bus detail pages **only
when a template column asks for it**, since that scan walks the whole document.

---

## 3b. The configurable Load Flow interface

Modelled on ETAP's own results window:

```
Bus Info                  Load Flow Results        Alert        Critical  Marginal
☑ Nominal kV              ☑ Voltage                Loading      [100] %   [95] %
☐ Amp Rating              ☑ kW Loading             Overvoltage  [106] %   [102] %
☐ Type (not published)    ☐ kvar Loading           Undervoltage [ 95] %   [ 98] %
                          ☐ kVA Loading
                          ☑ Amp Loading            Power: kVA / MVA
                          ☐ % Loading              Voltage: % / Actual Value
```

**Column order is deterministic:** Bus ID (always first, never removable), the
selected Bus Info fields, the selected Load Flow Results, an optional Alert
column, then Remarks. Unselected fields produce no column at all.

**One dataset, three outputs.** The PDF is parsed once when it is loaded;
`loadflow_processor.build_table()` turns records plus configuration into a
single table, and the preview, the Word document and the Excel workbook are all
rendered from it. Changing a checkbox rebuilds the table only — it never
re-reads the PDF (there is a test that asserts this).

### Classification

| Check | Flagged when |
|---|---|
| Overvoltage | voltage % ≥ critical → CRITICAL, ≥ marginal → MARGINAL |
| Undervoltage | voltage % ≤ critical → CRITICAL, ≤ marginal → MARGINAL |
| Loading | % loading ≥ critical → CRITICAL, ≥ marginal → MARGINAL |

Worst outcome wins. **Defaults keep the application's existing engineering
values:** overvoltage critical 106 % and undervoltage critical 95 % are the old
`95-106` band, not the screenshot's 105 %. Marginal bands (102 / 98) and the
loading limits (100 / 95) follow the reference interface. Selecting the
`ACCEPTABLE / NOT ACCEPTABLE` wording with the marginal limits off reproduces
the previous behaviour exactly — a test checks the same 24 buses are flagged.

### Nothing is invented

* **Type** is not published in an ETAP Load Flow report (it appears in the
  Short Circuit report), so the checkbox is disabled and says so.
* **Amp Rating** and **% Loading** exist only for buses ETAP gives a continuous
  rating; those cells stay blank and the interface reports how many.
* **Actual voltage** is `Nominal kV × Voltage % / 100` — offered only because
  both inputs are in the report.
* A bus that no enabled limit can assess gets a blank remark, never a guess.

---

## 3c. The configurable Short Circuit interface

Modelled on ETAP's Short Circuit Duty Analyzer:

```
Study                         Info                Results          Alert
Standard   IEC / ANSI         ☑ Nominal kV        ☑ I"k            ☑ Critical [100] %
Study type 3-Ph / 1-Ph        ☐ Type              ☑ ip             ☑ Marginal [ 95] %
Report     SC_Max             ☐ Cfactor           ☑ Ik             ☐ Skip non-alerted
                              ☑ Rated ip          ☐ Standard       ☑ Skip nodes
                              ☐ Rated device      Ib sym / Idc     Device type  [multi]
                              ☐ X/R                (not published)  Units  kA/A  kV/V
```

**Standard, study type and study case are detected from the report header**, not
assumed — the sample reads `IEC 60909`, `3-Phase`, `SC_Max`. Choosing a
different standard or study type only relabels the report and says so; it never
recalculates, because the numbers are ETAP's.

**Column order is deterministic:** ID (always first), selected Info, selected
Results, optional Duty %, then Remarks. Unselected fields produce no column.

**Thresholds preserve the existing rule.** The old test was "peak current must
not exceed the equipment making capacity". That is now expressed as
`duty % = ip / Rated ip × 100`, flagged CRITICAL at ≥ 100 % and MARGINAL at
≥ 95 % — which is exactly what ETAP's own `*` and `#` footnotes use. A test
asserts bus-by-bus agreement with the original `sc_rules.peak_vs_rating`.

### Filters

| Control | Effect |
|---|---|
| Device type | Keeps only the selected ETAP bus classes (Bus, Switchgear, Switchboard, Mcc, Switchrack, Panelboard) |
| Skip nodes | Drops ETAP's auto-generated `~` nodes |
| Skip non-alerted devices | Keeps only rows reaching a threshold |

If a combination removes every row the export is **refused with an explanation**
naming the filter responsible, rather than producing a plausible-looking empty
report.

### Nothing is invented

* **Ib sym / Idc** are published per protective device and per breaking time,
  never as a single per-bus figure, so those checkboxes are disabled and say so.
* **Cfactor** and **X/R** come from each bus's detail page. Selecting either
  triggers a whole-document scan (~20 s on the 355-page sample), which is why
  they are off by default and labelled "scans the whole report".
* **Rated ip** exists only where the report gives the bus a rated device (95 of
  256 in the sample); those rows stay blank and get a blank remark.
* Units convert exactly (kA↔A, kV↔V) and never change a classification.

---

## 4. Template-driven columns (the original Short Circuit report)

`template_mapping.py` reads the header of the table in the Word template and
matches each column to a canonical field. The supplied Short Circuit template
maps like this (`python cli.py --show-mapping`):

| Template column | Field |
|---|---|
| Switchgear ID | `bus_id` |
| Bus Rating (kV, A) | `nominal_voltage` |
| Switchgear Rating (kA) / Ip (peak) | `rating_peak` |
| 3 Phase ETAP Results (kA) / I" k | `ik_initial` |
| 3 Phase ETAP Results (kA) / Ip | `ip_peak` |
| 3 Phase ETAP Results (kA) / Ik | `ik_steady` |
| Remarks | `remarks` |

Multi-row headers, merged group headers and shaded spacer bands are all
detected. Reorder the columns, rename them, or add `X/R Ratio`, `Ib sym`,
`Idc`, `Device ID` — the writer follows. A column whose header matches nothing
is left exactly as the template has it, and nothing else about the document is
touched: fonts, borders, shading, widths and spacing all come from the
template, and the table is populated, never re-created.

Adding a quantity = one entry in `FIELD_PATTERNS` + the value in the parser.

---

## 4b. Excel export

Every run writes a workbook beside the Word document, with the same base name:

```
Bus Report.docx              Short Circuit Report.docx
Bus Report.xlsx              Short Circuit Report.xlsx
```

`excel_writer.py` is fed the **same headers and rows that were just written
into the Word table** - `reports.generate()` parses the PDF once and hands the
one table to both writers, so the two files cannot drift apart. A test opens
both outputs and compares them cell by cell.

Because it works from that generic table, any future report type gets its Excel
export with no extra code.

The sheet is styled as a schedule, not a data dump: bold header band in the
template's purple, thin borders, frozen header row, auto-fitted column widths
(capped, with wrapping for long IDs), centred values, IDs left aligned, an
autofilter, and landscape fit-to-width printing with the header repeated on
every page. Numeric cells are stored as **numbers** carrying the same decimal
count the Word document shows (`19.09` displays identically but sorts and
calculates), while values with units (`132 kV`, `380 V`) and text such as
`ACCEPTABLE` stay text - exactly as in the document.

---

## 5. Remarks rules (Short Circuit)

`sc_rules.py` holds the acceptance logic as interchangeable functions:

| Rule | Verdict |
|---|---|
| `peak_vs_rating` *(default)* | `ip` must not exceed the equipment making (peak) capacity |
| `peak_and_breaking` | as above **and** I"k vs the breaking capacity |
| `etap_flag` | trust ETAP's own `*` duty flags |
| `blank` | leave Remarks for the engineer |

A bus with **no rating in the report gets a blank remark**, not "NOT
ACCEPTABLE" — an unverifiable claim is worse than an obvious gap. Add a project
rule by writing a function and registering it in `RULES`; parser and writer are
untouched.

### Two things worth confirming for your report

1. **Which rating is reported.** When the bus is modelled as switchgear the
   assembly's own rating is used. Otherwise the devices on the bus are combined
   with `RATING_AGGREGATION` in `shortcircuit_parser.py`, default `min` — the
   lowest rated device governs, which is the binding constraint and the one
   ETAP flags. Set it to `max` or `first` if your practice differs.
2. **Which buses to list.** By default every bus in the results table is
   written. In the sample only 95 of 256 have an equipment rating, because most
   entries are plain bus nodes with no protective device. The *"Switchgear,
   switchboards, MCCs and panels only"* option narrows the table to the 67 real
   assemblies — probably what a schedule headed "Switchgear ID" wants.

---

## 6. Error handling

| Condition | Message |
|---|---|
| No bus results table (Load Flow) | `No Bus Output Data found in the uploaded ETAP report.` |
| No column selected | `Select at least one Bus Info or Load Flow Results field ...` |
| Marginal limit beyond its critical limit | names the row and both values |
| Remarks on with every limit off | asks for a limit or for Remarks to be turned off |
| No summary table (Short Circuit) | `No Bus Short Circuit Results table was found in the uploaded ETAP report.` |
| Limits not a `Lower-Upper` pair | `Please enter voltage limits in the format Lower-Upper (Example: 95-106).` |
| Load Flow PDF unreadable | `Unable to read the ETAP report.` |
| Short Circuit PDF unreadable | `Unable to read the ETAP Short Circuit Report.` |
| Output open in Word | Prompt to close it and retry |

Scanned PDFs: both parsers detect a missing text layer and run OCRmyPDF when it
is installed; otherwise they say so.

---

## 7. Performance

Measured on the supplied reports:

Both documents are produced in one pass; the workbook adds well under a second.

| Report | Pages | Buses | Time |
|---|---|---|---|
| Short Circuit study | 355 | 276 | ~2.5 s |
| Load Flow study | 74 | 283 | ~7 s |

The Short Circuit parser is the faster of the two despite the longer document,
because of the backward scan; the Load Flow parser reads every page. Parsing
runs in a worker thread (desktop) and is cached on the uploaded bytes (web).

---

## 8. Building the executable

```bash
pip install -r requirements-desktop.txt pyinstaller
pyinstaller build.spec
# -> dist/ETAP Report Filler.exe   (single file, windowed, both templates inside)
```

---

## 9. Adding the next report type

Motor Starting, Arc Flash, PDC, cable or transformer reports all follow the
same three steps, and **no front end changes**:

1. Write `<name>_parser.py` on top of `table_reader.py` — a header predicate, a
   column builder and a row reader.
2. Drop `assets/<Name>_Template.docx` in and, if its columns are new, add
   patterns to `template_mapping.FIELD_PATTERNS`. The generic writer handles the
   rest.
3. Register a `ReportType` in `reports.py` declaring the parser, writer,
   template, default filename and any extra UI inputs.

The desktop selector, the web radio buttons and the CLI `--type` choices are all
generated from that registry.
