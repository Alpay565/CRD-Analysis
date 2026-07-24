# CRD-Analysis

Quick analysis tool for **vehicle-application / carpark cross-reference** lists —
the kind of workbook you build when comparing competitor products across their
covered vehicles (ktypes) and the carpark behind each one.

Feed it a spreadsheet like `Book1.xlsx` and it produces, in seconds, the views
that are painful to build by hand:

- **Carpark per product** (and per competitor) — how much market each product covers
- **Ktype master list** — every unique ktype/car application with its carpark
- **Total carpark per ktype**, deduplicated, with **cross-competitor overlap / gap** analysis
- **Shared ktypes** — every ktype covered by more than one product, so overlapping applications are easy to spot
- **Simplified application list** in the reporting-template column layout, with **carpark added** and a **Covered in Products** column

## Requirements

**None.** The tool uses only the Python 3 standard library — no `pip install`,
no pandas/openpyxl. It reads `.xlsx` directly and writes Excel, HTML and CSV.
Any machine with Python 3.8+ can run it.

## Usage

**Windows — easiest:** double-click **`run_analysis.bat`** to analyze `Book1.xlsx`,
or drag any `.xlsx` file onto it. Results appear in the `analysis_output` folder.
(Run the command below in **Command Prompt** — the `C:\...>` window — not inside
the Python `>>>` prompt.)

```bash
python3 analyze_crd.py [INPUT.xlsx] [--outdir DIR]
```

- `INPUT.xlsx` — the workbook to analyze (default: `Book1.xlsx`)
- `--outdir` — where to write results (default: `analysis_output/`)

Examples:

```bash
python3 analyze_crd.py                          # analyze Book1.xlsx -> analysis_output/
python3 analyze_crd.py competitors_2026.xlsx    # analyze a different file
python3 analyze_crd.py data.xlsx --outdir out   # choose the output folder
```

A summary is printed to the console; full results go to `--outdir`.

### Analyzing a different file

The tool works on **any** workbook with the same column layout (see below) — not
just `Book1.xlsx`. Three ways to point it at a new file:

- **Drag-and-drop (Windows, no typing):** drag the `.xlsx` onto `run_analysis.bat`.
- **Put it next to the tool:** drop your file in this folder and run
  `python analyze_crd.py YOURFILE.xlsx`.
- **Full path (file kept elsewhere):** wrap the path in quotes —
  `python analyze_crd.py "C:\Users\amutlu\Documents\competitors_2026.xlsx"`.

Each run overwrites `analysis_output/`; add `--outdir some_folder` to keep several
analyses side by side.

## Outputs

| File | What it contains |
|------|------------------|
| `CRD_Analysis.xlsx` | Formatted workbook with all analyses on separate sheets (Summary, Carpark per Product, Carpark per Competitor, Competitor Overlap, Ktype Master List, Shared Ktypes, Simplified List). Header row is frozen and filterable; carpark is a real number. |
| `CRD_Analysis.html` | Self-contained dashboard — opens in any browser, no internet needed. KPI cards + every table with inline carpark bars. |
| `carpark_per_product.csv` | Applications and total carpark per product, with % of total. |
| `carpark_per_competitor.csv` | Same, grouped by competitor. |
| `competitor_overlap.csv` | Exclusive vs. shared ktypes/carpark per competitor (gap analysis). |
| `shared_ktypes.csv` | Every ktype covered by more than one product, with the covering products, competitors and carpark. |
| `ktype_master.csv` | One row per unique ktype: vehicle details, carpark, competitors, plus **one column per product** marked `x` where that product covers the ktype (a ktype shared by several products shows several `x`s). |
| `simplified_list.csv` | Reporting-template columns (Covered Vehicle, Vehicle ID, Car Marker, Model, Version, …), one row per ktype, with **Carpark** appended and **one `x`-marked column per product on the left** (mirroring the source template's part-number columns). `Covered Vehicle` is left blank (a manual flag for newly discovered applications). |

## Expected input

The first worksheet must have a header row that includes at least a **ktype**
column (`KtypNr`) and a **carpark** column (`Carpark per KtypNr`). Columns are
matched **by header name** (with common aliases), so the tool keeps working if
columns are reordered or lightly renamed. Recognized columns:

`Generic Product Designation`, `Reference`, `Competitor`, `Product Name`,
`Range`, `KtypNr`, `Vehicle type`, `Manufacturer`, `Car Model`, `Version`,
`Model year from`, `Model year to`, `Body type`, `Drive type`,
`Displacement (l)`, `Displacement (ccm techn.)`, `Fuel Type`, `kW`,
`Horse Power`, `Cylinders`, `Valves`, `Engine Type`, `Engine codes`,
`Limitations`, `Carpark per KtypNr`.

## How totals are calculated

- **Grand-total carpark** is **deduplicated by ktype**: if the same ktype appears
  under several products/competitors, its carpark is counted **once**.
- **Per-product / per-competitor totals** sum the carpark of the ktypes that
  product/competitor covers. When products share ktypes, these can add up to more
  than the grand total (that is expected — it reflects overlapping coverage), and
  `% of Total Carpark` is coverage relative to the deduped grand total.
- Rows with a missing/invalid carpark are counted as `0` and flagged in the
  Summary; ktypes with conflicting carpark values keep the first value and are
  flagged.

## Notes

`analysis_output/` is checked in as a ready-made example built from `Book1.xlsx`.
The values in `Book1.xlsx` are placeholders (real ktype/carpark data is
confidential); the tool works on structure, so it runs identically on the real files.
