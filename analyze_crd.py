#!/usr/bin/env python3
"""
CRD Vehicle Application Data Analyzer
=====================================

Analyze a competitor vehicle-application / carpark cross-reference workbook
(such as ``Book1.xlsx``) and produce quick, reusable analysis outputs:

  * Carpark per product (and per competitor)
  * A master list of every unique ktype with its carpark
  * Total carpark per ktype / car application, with cross-competitor overlap
  * A simplified application list matching the reporting template layout,
    with carpark added

The tool is intentionally **dependency-free** - it uses only the Python
standard library, so it runs anywhere Python 3 is installed (no ``pip install``
required). It reads ``.xlsx`` directly and writes an Excel workbook, a
self-contained HTML dashboard and CSV exports.

Usage
-----
    python3 analyze_crd.py [INPUT.xlsx] [--outdir DIR]

Defaults: INPUT = ``Book1.xlsx``, DIR = ``analysis_output``.

The input sheet is expected to contain (in any column order) a header row with
at least a ktype column (``KtypNr``) and a carpark column
(``Carpark per KtypNr``). Columns are matched by header name, so the tool keeps
working if columns are reordered or lightly renamed.
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import html as _html
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS = "{" + MAIN_NS + "}"


# --------------------------------------------------------------------------- #
# XLSX reading (standard library only)
# --------------------------------------------------------------------------- #
def _col_to_index(cell_ref: str) -> int:
    """'B7' -> 1 (zero-based column index)."""
    letters = re.match(r"([A-Za-z]+)", cell_ref).group(1).upper()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_xlsx(path: str):
    """Return (headers, rows) from the first worksheet of an .xlsx file.

    ``headers`` is a list of strings; ``rows`` is a list of lists (all strings),
    each padded to len(headers).
    """
    with zipfile.ZipFile(path) as z:
        # Shared strings
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot:
                shared.append("".join(t.text or "" for t in si.iter(_NS + "t")))

        # Locate the first worksheet part via the workbook relationships
        sheet_part = _first_sheet_part(z)
        sroot = ET.fromstring(z.read(sheet_part))

        raw_rows = []
        max_col = 0
        for row in sroot.iter(_NS + "row"):
            cells = {}
            for c in row.findall(_NS + "c"):
                ref = c.get("r") or ""
                ci = _col_to_index(ref) if ref else len(cells)
                cells[ci] = _cell_value(c, shared)
                if ci > max_col:
                    max_col = ci
            raw_rows.append(cells)

    width = max_col + 1
    grid = [[(r.get(i, "") or "").strip() for i in range(width)] for r in raw_rows]

    header_idx = _find_header_row(grid)
    headers = grid[header_idx]
    rows = [r for r in grid[header_idx + 1:] if any(v != "" for v in r)]
    # Pad/trim rows to header width
    rows = [(r + [""] * width)[:width] for r in rows]
    return headers, rows


def _first_sheet_part(z: zipfile.ZipFile) -> str:
    """Resolve the path of the first worksheet part."""
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        first = wb.find(_NS + "sheets").find(_NS + "sheet")
        rid = first.get(rns + "id")
        for rel in rels.iter(rel_ns + "Relationship"):
            if rel.get("Id") == rid:
                target = rel.get("Target")
                return target if target.startswith("xl/") else "xl/" + target.lstrip("/")
    except Exception:
        pass
    # Fallback
    return "xl/worksheets/sheet1.xml"


def _cell_value(c, shared) -> str:
    t = c.get("t")
    if t == "s":
        v = c.find(_NS + "v")
        return shared[int(v.text)] if v is not None else ""
    if t == "inlineStr":
        is_ = c.find(_NS + "is")
        return "".join(x.text or "" for x in is_.iter(_NS + "t")) if is_ is not None else ""
    v = c.find(_NS + "v")
    return v.text if v is not None else ""


def _find_header_row(grid) -> int:
    """Find the header row by looking for the ktype/carpark signature."""
    for i, row in enumerate(grid[:15]):
        norm = [_norm(x) for x in row]
        if any("ktyp" in x for x in norm) and any("carpark" in x for x in norm):
            return i
    # Fallback: first non-empty row
    for i, row in enumerate(grid):
        if any(v != "" for v in row):
            return i
    return 0


# --------------------------------------------------------------------------- #
# Column resolution
# --------------------------------------------------------------------------- #
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


# Logical field -> list of accepted header aliases (normalized, substring match).
FIELD_ALIASES = OrderedDict([
    ("generic_designation", ["generic product designation", "generic designation"]),
    ("reference", ["reference", "ref"]),
    ("competitor", ["competitor"]),
    ("product_name", ["product name"]),
    ("range", ["range"]),
    ("ktyp", ["ktypnr", "ktyp", "ktype", "vehicle id"]),
    ("vehicle_type", ["vehicle type", "covered vehicle"]),
    ("manufacturer", ["manufacturer", "car marker", "make", "brand"]),
    ("model", ["car model", "model"]),
    ("version", ["version"]),
    ("year_from", ["model year from", "from"]),
    ("year_to", ["model year to", "to"]),
    ("body", ["body type"]),
    ("drive", ["drive type"]),
    ("displ_l", ["displacement (l)", "litters", "liters", "displacement l"]),
    ("displ_cc", ["displacement (ccm", "ccm", "cc"]),
    ("fuel", ["fuel type", "fuel"]),
    ("kw", ["kw"]),
    ("hp", ["horse power", "hp"]),
    ("cylinders", ["cylinders"]),
    ("valves", ["valves"]),
    ("engine_type", ["engine type"]),
    ("engine_codes", ["engine codes", "engine code"]),
    ("limitations", ["limitations"]),
    ("carpark", ["carpark per ktypnr", "carpark"]),
])


def resolve_columns(headers):
    """Map logical field name -> column index using header aliases."""
    norm_headers = [_norm(h) for h in headers]
    used = set()
    mapping = {}
    for field, aliases in FIELD_ALIASES.items():
        found = None
        # Exact match first
        for alias in aliases:
            for i, h in enumerate(norm_headers):
                if i in used:
                    continue
                if h == alias:
                    found = i
                    break
            if found is not None:
                break
        # Substring match next
        if found is None:
            for alias in aliases:
                for i, h in enumerate(norm_headers):
                    if i in used:
                        continue
                    if alias in h or (h and h in alias):
                        found = i
                        break
                if found is not None:
                    break
        if found is not None:
            mapping[field] = found
            used.add(found)
    return mapping


def parse_carpark(value):
    """Tolerant integer parse: '1,461' -> 1461, '' -> None, '12.0' -> 12."""
    if value is None:
        return None
    s = str(value).strip().replace(",", "").replace(" ", "")
    if s == "":
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


# --------------------------------------------------------------------------- #
# Analysis
# --------------------------------------------------------------------------- #
class Sheet:
    """A simple tabular result rendered to xlsx/html/csv."""

    def __init__(self, name, headers, rows, num_cols=None, pct_cols=None,
                 bar_col=None, note=None):
        self.name = name
        self.headers = headers
        self.rows = rows
        self.num_cols = set(num_cols or [])
        self.pct_cols = set(pct_cols or [])
        self.bar_col = bar_col
        self.note = note


def _g(row, mapping, field):
    idx = mapping.get(field)
    return row[idx] if idx is not None and idx < len(row) else ""


def analyze(headers, rows):
    """Run all analyses; return (summary_dict, list_of_Sheet, warnings)."""
    m = resolve_columns(headers)
    warnings = []

    if "ktyp" not in m:
        raise SystemExit("ERROR: could not find a KtypNr column in the input.")
    if "carpark" not in m:
        warnings.append("No 'Carpark per KtypNr' column found - carpark totals will be 0.")

    # --- Per-ktype carpark (deduped, with consistency check) ---------------- #
    ktype_carpark = OrderedDict()       # ktype -> carpark (int)
    ktype_info = {}                     # ktype -> representative attribute row
    ktype_products = {}                 # ktype -> set of product labels
    ktype_competitors = {}             # ktype -> set of competitors
    ktype_limits = {}                  # ktype -> ordered list of distinct limitations
    product_labels_order = []          # distinct product labels, in first-seen order
    product_seen = set()
    missing_ktyp = 0
    missing_cp = 0
    inconsistent = []

    def product_label(row):
        rng = _g(row, m, "range")
        pn = _g(row, m, "product_name")
        ref = _g(row, m, "reference")
        comp = _g(row, m, "competitor")
        base = rng or pn or (("Ref " + ref) if ref else "") or "Unknown"
        if base in ("Undefined", "") and pn:
            base = pn
        return f"{comp}: {base}".strip(": ") if comp else base

    for row in rows:
        kt = _g(row, m, "ktyp").strip()
        cp = parse_carpark(_g(row, m, "carpark"))
        if not kt:
            missing_ktyp += 1
            continue
        if cp is None:
            missing_cp += 1
            cp = 0
        if kt in ktype_carpark:
            if ktype_carpark[kt] != cp and cp != 0:
                if ktype_carpark[kt] == 0:
                    ktype_carpark[kt] = cp
                else:
                    inconsistent.append(kt)
        else:
            ktype_carpark[kt] = cp
            ktype_info[kt] = row
        lbl = product_label(row)
        ktype_products.setdefault(kt, set()).add(lbl)
        if lbl not in product_seen:
            product_seen.add(lbl)
            product_labels_order.append(lbl)
        ktype_competitors.setdefault(kt, set()).add(_g(row, m, "competitor") or "?")
        lim_v = _g(row, m, "limitations").strip()
        kl = ktype_limits.setdefault(kt, [])
        if lim_v and lim_v not in kl:
            kl.append(lim_v)

    grand_total = sum(ktype_carpark.values())
    unique_ktypes = len(ktype_carpark)

    if missing_ktyp:
        warnings.append(f"{missing_ktyp} row(s) had no KtypNr and were skipped.")
    if missing_cp:
        warnings.append(f"{missing_cp} row(s) had no/invalid carpark value (counted as 0).")
    if inconsistent:
        uniq = sorted(set(inconsistent))
        warnings.append(
            f"{len(uniq)} ktype(s) had inconsistent carpark values across rows "
            f"(first value kept): {', '.join(uniq[:10])}"
            + (" ..." if len(uniq) > 10 else "")
        )

    # --- Carpark per product ------------------------------------------------ #
    prod_order = []
    prod_data = OrderedDict()   # key -> dict(comp, ref, name, range, ktypes set)
    for row in rows:
        kt = _g(row, m, "ktyp").strip()
        if not kt:
            continue
        key = (_g(row, m, "competitor"), _g(row, m, "reference"),
               _g(row, m, "product_name"), _g(row, m, "range"))
        if key not in prod_data:
            prod_data[key] = {"ktypes": set()}
            prod_order.append(key)
        prod_data[key]["ktypes"].add(kt)

    product_rows = []
    for key in prod_order:
        comp, ref, name, rng = key
        kts = prod_data[key]["ktypes"]
        total = sum(ktype_carpark.get(k, 0) for k in kts)
        product_rows.append([
            comp, ref, name, rng, len(kts), total,
            (total / grand_total) if grand_total else 0.0,
        ])
    product_rows.sort(key=lambda r: r[5], reverse=True)

    product_sheet = Sheet(
        "Carpark per Product",
        ["Competitor", "Reference", "Product Name", "Range",
         "Applications (ktypes)", "Total Carpark", "% of Total Carpark"],
        product_rows, num_cols=[4, 5], pct_cols=[6], bar_col=5,
        note="'% of Total Carpark' is coverage vs. the deduped grand total; "
             "values can sum to more than 100% when products share ktypes.",
    )

    # --- Carpark per competitor -------------------------------------------- #
    comp_ktypes = OrderedDict()
    for kt in ktype_carpark:
        for c in ktype_competitors.get(kt, ()):
            comp_ktypes.setdefault(c, set()).add(kt)
    competitor_rows = []
    for c, kts in comp_ktypes.items():
        total = sum(ktype_carpark.get(k, 0) for k in kts)
        competitor_rows.append([
            c, len(kts), total, (total / grand_total) if grand_total else 0.0,
        ])
    competitor_rows.sort(key=lambda r: r[2], reverse=True)
    competitor_sheet = Sheet(
        "Carpark per Competitor",
        ["Competitor", "Applications (ktypes)", "Total Carpark", "% of Total Carpark"],
        competitor_rows, num_cols=[1, 2], pct_cols=[3], bar_col=2,
    )

    # --- Competitor overlap / gap ------------------------------------------ #
    overlap_rows = []
    for c, kts in comp_ktypes.items():
        unique_only = [k for k in kts if len(ktype_competitors.get(k, ())) == 1]
        shared = [k for k in kts if len(ktype_competitors.get(k, ())) > 1]
        cp_unique = sum(ktype_carpark.get(k, 0) for k in unique_only)
        cp_shared = sum(ktype_carpark.get(k, 0) for k in shared)
        overlap_rows.append([
            c, len(kts), len(unique_only), cp_unique, len(shared), cp_shared,
        ])
    overlap_rows.sort(key=lambda r: r[1], reverse=True)
    overlap_sheet = Sheet(
        "Competitor Overlap",
        ["Competitor", "Total ktypes", "Exclusive ktypes", "Exclusive Carpark",
         "Shared ktypes", "Shared Carpark"],
        overlap_rows, num_cols=[1, 2, 3, 4, 5],
        note="'Exclusive' = ktypes only this competitor covers; 'Shared' = also "
             "covered by at least one other competitor. Highlights market gaps.",
    )

    # --- Ktype master list -------------------------------------------------- #
    attr_fields = ["vehicle_type", "manufacturer", "model", "version",
                   "year_from", "year_to", "body", "drive", "displ_l", "displ_cc",
                   "fuel", "kw", "hp", "cylinders", "valves", "engine_type",
                   "engine_codes"]
    # "Product Limitations" is appended separately: it is collected across all of
    # the ktype's rows (ktype_limits), not read from the representative row.
    attr_headers = ["Vehicle Type", "Manufacturer", "Model", "Version",
                    "Year From", "Year To", "Body Type", "Drive Type",
                    "Displacement (l)", "Displacement (cc)", "Fuel Type", "kW",
                    "HP", "Cylinders", "Valves", "Engine Type", "Engine Codes",
                    "Product Limitations"]
    master_rows = []
    for kt, cp in ktype_carpark.items():
        info = ktype_info.get(kt, [""] * len(headers))
        lims = ktype_limits.get(kt, [])
        attrs = [_g(info, m, f) for f in attr_fields] + [
            " / ".join(lims) if lims else "None"]
        pset = ktype_products.get(kt, set())
        comps = sorted(ktype_competitors.get(kt, []))
        coverage = ["x" if lbl in pset else "" for lbl in product_labels_order]
        master_rows.append(
            [kt] + attrs + [cp, len(pset), ", ".join(comps)] + coverage
        )
    cp_idx = len(attr_headers) + 1
    master_rows.sort(key=lambda r: r[cp_idx], reverse=True)
    master_sheet = Sheet(
        "Ktype Master List",
        ["KtypNr"] + attr_headers + ["Carpark", "# Products", "Competitors"]
        + product_labels_order,
        master_rows, num_cols=[cp_idx, cp_idx + 1], bar_col=cp_idx,
        note="One row per unique ktype. The product columns on the right are marked "
             "'x' for every product that covers the ktype, so a ktype shared across "
             "products shows more than one 'x'.",
    )

    # --- Shared ktypes (covered by more than one product) ------------------- #
    shared_rows = []
    for kt, cp in ktype_carpark.items():
        prods = sorted(ktype_products.get(kt, []))
        if len(prods) > 1:
            info = ktype_info.get(kt, [""] * len(headers))
            comps = sorted(ktype_competitors.get(kt, []))
            lims = ktype_limits.get(kt, [])
            shared_rows.append([
                kt, _g(info, m, "manufacturer"), _g(info, m, "model"),
                _g(info, m, "version"), cp, len(prods), len(comps),
                " | ".join(prods), ", ".join(comps),
                " / ".join(lims) if lims else "None",
            ])
    shared_rows.sort(key=lambda r: (r[5], r[4]), reverse=True)
    shared_count = len(shared_rows)
    shared_sheet = Sheet(
        "Shared Ktypes",
        ["KtypNr", "Manufacturer", "Model", "Version", "Carpark",
         "# Products", "# Competitors", "Products", "Competitors",
         "Product Limitations"],
        shared_rows, num_cols=[4, 5, 6], bar_col=4,
        note=("Ktypes covered by more than one product - the same car application "
              "sold under several references. "
              + ("No shared ktypes were found in this file."
                 if shared_count == 0 else f"{shared_count} shared ktype(s).")),
    )

    # --- Simplified list (reporting-template layout, one row per ktype) ----- #
    simplified_vehicle_headers = [
        "Covered Vehicle", "Vehicle ID (Ktype)", "Car Marker", "Model", "Version",
        "From", "To", "Body Type", "Drive Type", "Litters", "cc", "Fuel type",
        "kW", "HP", "Cylinders", "Valves", "Engine Type", "Engine Code",
        "Product Limitations", "Carpark", "# Products",
    ]
    simplified_headers = product_labels_order + simplified_vehicle_headers
    nprod = len(product_labels_order)
    simplified_rows = []
    for kt, cp in ktype_carpark.items():
        info = ktype_info.get(kt, [""] * len(headers))
        pset = ktype_products.get(kt, set())
        lims = ktype_limits.get(kt, [])
        lim = " / ".join(lims) if lims else "None"
        coverage = ["x" if lbl in pset else "" for lbl in product_labels_order]
        simplified_rows.append(coverage + [
            "",  # Covered Vehicle: manual "newly discovered application" flag - left blank
            kt, _g(info, m, "manufacturer"), _g(info, m, "model"),
            _g(info, m, "version"), _g(info, m, "year_from"), _g(info, m, "year_to"),
            _g(info, m, "body"), _g(info, m, "drive"), _g(info, m, "displ_l"),
            _g(info, m, "displ_cc"), _g(info, m, "fuel"), _g(info, m, "kw"),
            _g(info, m, "hp"), _g(info, m, "cylinders"), _g(info, m, "valves"),
            _g(info, m, "engine_type"), _g(info, m, "engine_codes"), lim, cp,
            len(pset),
        ])
    cp_col = nprod + 19       # Carpark column position
    nprod_col = nprod + 20    # "# Products" column position
    simplified_rows.sort(key=lambda r: (r[nprod_col], r[cp_col]), reverse=True)
    simplified_sheet = Sheet(
        "Simplified List", simplified_headers, simplified_rows,
        num_cols=[cp_col, nprod_col], bar_col=cp_col,
        note="Reporting-template layout, one row per unique ktype. The product "
             "columns on the left are marked 'x' for each product that covers the "
             "ktype (mirroring the source template's part-number columns), so a "
             "ktype shared across products shows more than one 'x'. 'Covered "
             "Vehicle' is left blank (a manual flag for newly discovered applications).",
    )

    summary = OrderedDict([
        ("Total Carpark (unique ktypes)", grand_total),
        ("Unique ktypes / car applications", unique_ktypes),
        ("Application rows in file", len([r for r in rows if _g(r, m, "ktyp").strip()])),
        ("Competitors", len(comp_ktypes)),
        ("Products", len(prod_order)),
        ("Ktypes shared across products", shared_count),
    ])

    sheets = [product_sheet, competitor_sheet, overlap_sheet,
              master_sheet, shared_sheet, simplified_sheet]
    return summary, sheets, warnings


# --------------------------------------------------------------------------- #
# Output: console
# --------------------------------------------------------------------------- #
def _fmt_int(v):
    return f"{v:,}" if isinstance(v, (int, float)) else str(v)


def print_console(summary, sheets, warnings):
    print("=" * 60)
    print("  CRD VEHICLE APPLICATION DATA - ANALYSIS SUMMARY")
    print("=" * 60)
    for k, v in summary.items():
        print(f"  {k:<38} {_fmt_int(v):>15}")
    prod = next(s for s in sheets if s.name == "Carpark per Product")
    print("\n  Top products by carpark:")
    for r in prod.rows[:8]:
        label = r[3] or r[2] or r[1]
        print(f"    {(str(r[0]) + ' | ' + str(label))[:44]:<46} "
              f"{_fmt_int(r[5]):>10}  ({r[6] * 100:4.1f}%)")
    if warnings:
        print("\n  Warnings:")
        for w in warnings:
            print(f"    - {w}")
    print("=" * 60)


# --------------------------------------------------------------------------- #
# Output: CSV
# --------------------------------------------------------------------------- #
def read_csv_rows(path):
    """Read a CSV written by write_csv(), skipping the leading 'sep=' hint."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        first = f.readline()
        if not first.lower().startswith("sep="):
            f.seek(0)
        return list(csv.reader(f))


def _csv_value(sheet, ci, v):
    if ci in sheet.pct_cols and isinstance(v, (int, float)):
        return round(v * 100, 1)
    return v


def write_csv(sheet, path):
    """Write the sheet as CSV.

    The leading ``sep=,`` line is an Excel hint: without it, Excel splits
    columns using the machine's regional list separator (a semicolon in many
    locales) and a comma-separated file lands entirely in column A. Excel
    consumes this line instead of showing it. Readers that don't understand it
    (including this tool's own viewer) skip it explicitly.
    """
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        f.write("sep=,\r\n")
        w = csv.writer(f)
        headers = [h + " (%)" if ci in sheet.pct_cols else h
                   for ci, h in enumerate(sheet.headers)]
        w.writerow(headers)
        for row in sheet.rows:
            w.writerow([_csv_value(sheet, ci, v) for ci, v in enumerate(row)])


# --------------------------------------------------------------------------- #
# Output: HTML dashboard (self-contained)
# --------------------------------------------------------------------------- #
_HTML_CSS = """
:root{--bg:#f4f6fb;--card:#fff;--ink:#1f2733;--muted:#5b6675;--line:#e2e7f0;
--accent:#305496;--accent2:#4472c4;--bar:#cfe0ff;}
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
background:var(--bg);color:var(--ink);font-size:14px;line-height:1.45}
header{background:var(--accent);color:#fff;padding:22px 28px}
header h1{margin:0;font-size:20px;font-weight:650}
header .meta{opacity:.85;font-size:12.5px;margin-top:4px}
main{max-width:1200px;margin:0 auto;padding:22px 20px 60px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:26px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 18px;
box-shadow:0 1px 2px rgba(20,40,80,.04)}
.kpi .v{font-size:26px;font-weight:700;color:var(--accent)}
.kpi .l{font-size:12px;color:var(--muted);margin-top:3px}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;
padding:16px 18px 6px;margin-bottom:22px;box-shadow:0 1px 2px rgba(20,40,80,.04)}
section h2{font-size:15.5px;margin:2px 0 4px}
section .note{font-size:12px;color:var(--muted);margin:0 0 10px}
.tablewrap{overflow:auto;max-height:520px;border:1px solid var(--line);border-radius:8px}
table{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}
thead th{position:sticky;top:0;background:var(--accent);color:#fff;text-align:left;
padding:8px 10px;font-weight:600;z-index:1}
tbody td{padding:6px 10px;border-bottom:1px solid var(--line)}
tbody tr:nth-child(even){background:#fafbfe}
td.num{text-align:right;font-variant-numeric:tabular-nums}
.barcell{position:relative}
.barcell .bar{position:absolute;left:0;top:0;bottom:0;background:var(--bar);z-index:0}
.barcell span{position:relative;z-index:1}
footer{max-width:1200px;margin:0 auto;padding:0 20px 40px;color:var(--muted);font-size:12px}
@media (prefers-color-scheme:dark){
:root{--bg:#12161d;--card:#1a2029;--ink:#e6ebf2;--muted:#9aa6b6;--line:#2a323d;
--accent:#3a5da8;--bar:#25406e}
tbody tr:nth-child(even){background:#1e2530}}
"""


def _cell_html(sheet, ci, v, bar_max):
    is_num = ci in sheet.num_cols or ci in sheet.pct_cols
    if ci in sheet.pct_cols and isinstance(v, (int, float)):
        text = f"{v * 100:.1f}%"
    elif ci in sheet.num_cols and isinstance(v, (int, float)):
        text = f"{v:,}"
    else:
        text = "" if v is None else str(v)
    esc = _html.escape(text)
    if sheet.bar_col == ci and isinstance(v, (int, float)) and bar_max:
        pct = max(0.0, min(100.0, v / bar_max * 100))
        return (f'<td class="num barcell"><span class="bar" style="width:{pct:.1f}%"></span>'
                f'<span>{esc}</span></td>')
    cls = ' class="num"' if is_num else ""
    return f"<td{cls}>{esc}</td>"


def _section_html(sheet):
    bar_max = 0
    if sheet.bar_col is not None:
        vals = [r[sheet.bar_col] for r in sheet.rows
                if isinstance(r[sheet.bar_col], (int, float))]
        bar_max = max(vals) if vals else 0
    head = "".join(f"<th>{_html.escape(h)}</th>" for h in sheet.headers)
    body = []
    for row in sheet.rows:
        cells = "".join(_cell_html(sheet, ci, v, bar_max) for ci, v in enumerate(row))
        body.append(f"<tr>{cells}</tr>")
    note = f'<p class="note">{_html.escape(sheet.note)}</p>' if sheet.note else ""
    count = f'<p class="note">{len(sheet.rows):,} rows</p>'
    return (f'<section><h2>{_html.escape(sheet.name)}</h2>{note}{count}'
            f'<div class="tablewrap"><table><thead><tr>{head}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div></section>')


def write_html(summary, sheets, warnings, path, source_name):
    now = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    kpis = "".join(
        f'<div class="kpi"><div class="v">{_fmt_int(v)}</div>'
        f'<div class="l">{_html.escape(k)}</div></div>'
        for k, v in summary.items()
    )
    warn_html = ""
    if warnings:
        items = "".join(f"<li>{_html.escape(w)}</li>" for w in warnings)
        warn_html = (f'<section><h2>Data notes</h2><ul style="margin:6px 0 12px;'
                     f'padding-left:20px;color:var(--muted);font-size:13px">{items}</ul></section>')
    sections = "".join(_section_html(s) for s in sheets)
    doc = (
        f"<!doctype html><html lang='en'><head><meta charset='utf-8'>"
        f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
        f"<title>CRD Analysis - {_html.escape(source_name)}</title>"
        f"<style>{_HTML_CSS}</style></head><body>"
        f"<header><h1>CRD Vehicle Application Analysis</h1>"
        f"<div class='meta'>Source: {_html.escape(source_name)} &middot; Generated {now}</div></header>"
        f"<main><div class='kpis'>{kpis}</div>{warn_html}{sections}</main>"
        f"<footer>Generated by analyze_crd.py &middot; carpark totals are deduplicated "
        f"by unique ktype.</footer></body></html>"
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


# --------------------------------------------------------------------------- #
# Output: XLSX writer (standard library only, inline strings)
# --------------------------------------------------------------------------- #
def _xml_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _index_to_col(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


# Style indices (see _styles_xml): 0 text, 1 header, 2 int #,##0,
# 3 bold int, 4 percent 0.00%
_S_TEXT, _S_HEADER, _S_INT, _S_BOLDINT, _S_PCT = 0, 1, 2, 3, 4


def _styles_xml():
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<styleSheet xmlns='" + MAIN_NS + "'>"
        "<fonts count='2'>"
        "<font><sz val='11'/><name val='Calibri'/></font>"
        "<font><b/><sz val='11'/><color rgb='FFFFFFFF'/><name val='Calibri'/></font>"
        "</fonts>"
        "<fills count='3'>"
        "<fill><patternFill patternType='none'/></fill>"
        "<fill><patternFill patternType='gray125'/></fill>"
        "<fill><patternFill patternType='solid'><fgColor rgb='FF305496'/>"
        "<bgColor indexed='64'/></patternFill></fill>"
        "</fills>"
        "<borders count='1'><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
        "<cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs>"
        "<cellXfs count='5'>"
        "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>"
        "<xf numFmtId='0' fontId='1' fillId='2' borderId='0' xfId='0' applyFont='1' "
        "applyFill='1' applyAlignment='1'><alignment vertical='center' wrapText='1'/></xf>"
        "<xf numFmtId='3' fontId='0' fillId='0' borderId='0' xfId='0' applyNumberFormat='1'/>"
        "<xf numFmtId='3' fontId='1' fillId='0' borderId='0' xfId='0' applyNumberFormat='1' applyFont='1'/>"
        "<xf numFmtId='10' fontId='0' fillId='0' borderId='0' xfId='0' applyNumberFormat='1'/>"
        "</cellXfs>"
        "<cellStyles count='1'><cellStyle name='Normal' xfId='0' builtinId='0'/></cellStyles>"
        "</styleSheet>"
    )


def _sheet_xml(sheet):
    ncols = len(sheet.headers)
    nrows = len(sheet.rows) + 1
    dim = f"A1:{_index_to_col(ncols - 1)}{nrows}"

    # Column widths from header/content length
    widths = []
    for ci, h in enumerate(sheet.headers):
        maxlen = len(str(h))
        for r in sheet.rows[:400]:
            maxlen = max(maxlen, len(str(r[ci]) if r[ci] is not None else ""))
        widths.append(min(52, max(9, maxlen + 2)))
    cols = "<cols>" + "".join(
        f"<col min='{i + 1}' max='{i + 1}' width='{w}' customWidth='1'/>"
        for i, w in enumerate(widths)
    ) + "</cols>"

    parts = []
    # Header row
    hc = []
    for ci, h in enumerate(sheet.headers):
        ref = f"{_index_to_col(ci)}1"
        hc.append(f"<c r='{ref}' s='{_S_HEADER}' t='inlineStr'><is><t xml:space='preserve'>"
                  f"{_xml_escape(h)}</t></is></c>")
    parts.append(f"<row r='1'>{''.join(hc)}</row>")

    # Data rows
    for ri, row in enumerate(sheet.rows, start=2):
        cells = []
        for ci, v in enumerate(row):
            ref = f"{_index_to_col(ci)}{ri}"
            if ci in sheet.pct_cols and isinstance(v, (int, float)):
                cells.append(f"<c r='{ref}' s='{_S_PCT}'><v>{v}</v></c>")
            elif ci in sheet.num_cols and isinstance(v, (int, float)):
                cells.append(f"<c r='{ref}' s='{_S_INT}'><v>{v}</v></c>")
            else:
                text = "" if v is None else str(v)
                cells.append(f"<c r='{ref}' t='inlineStr'><is><t xml:space='preserve'>"
                             f"{_xml_escape(text)}</t></is></c>")
        parts.append(f"<row r='{ri}'>{''.join(cells)}</row>")

    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<worksheet xmlns='" + MAIN_NS + "'>"
        f"<dimension ref='{dim}'/>"
        "<sheetViews><sheetView workbookViewId='0'>"
        "<pane ySplit='1' topLeftCell='A2' activePane='bottomLeft' state='frozen'/>"
        "<selection pane='bottomLeft' activeCell='A2' sqref='A2'/>"
        "</sheetView></sheetViews>"
        "<sheetFormatPr defaultRowHeight='15'/>"
        f"{cols}"
        f"<sheetData>{''.join(parts)}</sheetData>"
        f"<autoFilter ref='{dim}'/>"
        "</worksheet>"
    )


def _safe_sheet_name(name, used):
    name = re.sub(r"[\[\]\*\?:/\\]", " ", name)[:31].strip() or "Sheet"
    base, n = name, 1
    while name.lower() in used:
        suffix = f" {n}"
        name = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(name.lower())
    return name


def write_xlsx(summary, sheets, warnings, path):
    # Build a Summary sheet (KPIs + warnings) as the first tab.
    sum_rows = [[k, v] for k, v in summary.items()]
    if warnings:
        sum_rows.append(["", ""])
        sum_rows.append(["Data notes", ""])
        for w in warnings:
            sum_rows.append([w, ""])
    summary_sheet = Sheet("Summary", ["Metric", "Value"], sum_rows, num_cols=[1])
    all_sheets = [summary_sheet] + sheets

    used = set()
    named = [(_safe_sheet_name(s.name, used), s) for s in all_sheets]

    content_types = (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>"
        "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>"
        "<Default Extension='xml' ContentType='application/xml'/>"
        "<Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>"
        "<Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/>"
        + "".join(
            f"<Override PartName='/xl/worksheets/sheet{i + 1}.xml' "
            f"ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>"
            for i in range(len(named))
        )
        + "</Types>"
    )
    root_rels = (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>"
        "</Relationships>"
    )
    sheets_xml = "".join(
        f"<sheet name='{_xml_escape(nm)}' sheetId='{i + 1}' r:id='rId{i + 1}'/>"
        for i, (nm, _) in enumerate(named)
    )
    workbook = (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<workbook xmlns='" + MAIN_NS + "' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>"
        f"<sheets>{sheets_xml}</sheets></workbook>"
    )
    wb_rels = (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        + "".join(
            f"<Relationship Id='rId{i + 1}' "
            f"Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' "
            f"Target='worksheets/sheet{i + 1}.xml'/>"
            for i in range(len(named))
        )
        + f"<Relationship Id='rId{len(named) + 1}' "
        f"Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' "
        f"Target='styles.xml'/>"
        "</Relationships>"
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", _styles_xml())
        for i, (_, s) in enumerate(named):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", _sheet_xml(s))


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Analyze a vehicle-application / carpark cross-reference .xlsx"
    )
    ap.add_argument("input", nargs="?", default="Book1.xlsx",
                    help="Input .xlsx file (default: Book1.xlsx)")
    ap.add_argument("--outdir", default="analysis_output",
                    help="Output directory (default: analysis_output)")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.input):
        ap.error(f"input file not found: {args.input}")
    os.makedirs(args.outdir, exist_ok=True)

    headers, rows = read_xlsx(args.input)
    summary, sheets, warnings = analyze(headers, rows)

    source_name = os.path.basename(args.input)
    xlsx_path = os.path.join(args.outdir, "CRD_Analysis.xlsx")
    html_path = os.path.join(args.outdir, "CRD_Analysis.html")
    write_xlsx(summary, sheets, warnings, xlsx_path)
    write_html(summary, sheets, warnings, html_path, source_name)

    csv_map = {
        "Carpark per Product": "carpark_per_product.csv",
        "Carpark per Competitor": "carpark_per_competitor.csv",
        "Competitor Overlap": "competitor_overlap.csv",
        "Shared Ktypes": "shared_ktypes.csv",
        "Ktype Master List": "ktype_master.csv",
        "Simplified List": "simplified_list.csv",
    }
    for s in sheets:
        if s.name in csv_map:
            write_csv(s, os.path.join(args.outdir, csv_map[s.name]))

    print_console(summary, sheets, warnings)
    print(f"\nOutputs written to: {os.path.abspath(args.outdir)}")
    print(f"  - {os.path.basename(xlsx_path)}   (Excel workbook, {len(sheets) + 1} sheets)")
    print(f"  - {os.path.basename(html_path)}   (HTML dashboard)")
    for name in csv_map.values():
        print(f"  - {name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
