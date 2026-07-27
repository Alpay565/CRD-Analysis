#!/usr/bin/env python3
"""
CRD Analyzer - desktop application
==================================

A single, self-contained application around ``analyze_crd.py``. Launch it
(double-click the built ``CRD_Analyzer.exe`` or run ``python crd_app.py``) and it
opens a small page in your default browser where you can:

  * drag-and-drop a vehicle-application / carpark ``.xlsx`` onto it,
  * see the list of documents it produced,
  * view them (the dashboard and CSVs open right in the browser), and
  * download whichever ones you want to your PC (individually or as a .zip).

It runs a private local web server bound to 127.0.0.1 - nothing is uploaded
anywhere, everything stays on your machine. Runtime dependencies: none beyond the
Python standard library (the analysis engine is reused from analyze_crd.py).
"""

from __future__ import annotations

import argparse
import html as _html
import io
import json
import os
import secrets
import sys
import tempfile
import threading
import urllib.parse
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# Make the sibling analysis module importable both as a script and when frozen
# into an .exe by PyInstaller (which unpacks to sys._MEIPASS).
_BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
if _BASE not in sys.path:
    sys.path.insert(0, _BASE)
import analyze_crd  # noqa: E402


# --------------------------------------------------------------------------- #
# Analysis (reuses analyze_crd.py end to end)
# --------------------------------------------------------------------------- #
# Output file name -> short description, in the order to present them.
OUTPUT_FILES = [
    ("CRD_Analysis.xlsx", "BEST FOR EXCEL - every table as a formatted sheet"),
    ("CRD_Analysis.html", "Interactive dashboard - open this first"),
    ("simplified_list.csv", "Simplified vehicle list + carpark + product columns"),
    ("ktype_master.csv", "Every unique ktype with carpark & product coverage"),
    ("shared_ktypes.csv", "Ktypes covered by more than one product"),
    ("carpark_per_product.csv", "Carpark totals per product"),
    ("carpark_per_competitor.csv", "Carpark totals per competitor"),
    ("competitor_overlap.csv", "Exclusive vs. shared coverage per competitor"),
]

CSV_MAP = {
    "Carpark per Product": "carpark_per_product.csv",
    "Carpark per Competitor": "carpark_per_competitor.csv",
    "Competitor Overlap": "competitor_overlap.csv",
    "Shared Ktypes": "shared_ktypes.csv",
    "Ktype Master List": "ktype_master.csv",
    "Simplified List": "simplified_list.csv",
}

SESSIONS = {}          # session id -> output directory
SESSIONS_LOCK = threading.Lock()


def run_analysis(input_path, outdir, source_name):
    """Run the full analysis into outdir; return (summary, warnings)."""
    headers, rows = analyze_crd.read_xlsx(input_path)
    summary, sheets, warnings = analyze_crd.analyze(headers, rows)
    analyze_crd.write_xlsx(summary, sheets, warnings,
                           os.path.join(outdir, "CRD_Analysis.xlsx"))
    analyze_crd.write_html(summary, sheets, warnings,
                           os.path.join(outdir, "CRD_Analysis.html"), source_name)
    for s in sheets:
        if s.name in CSV_MAP:
            analyze_crd.write_csv(s, os.path.join(outdir, CSV_MAP[s.name]))
    return summary, warnings


def analyze_upload(data, filename):
    """Analyze uploaded bytes; return a JSON-serializable result dict."""
    name = os.path.basename(filename or "upload.xlsx")
    if not name.lower().endswith((".xlsx", ".xlsm")):
        return {"ok": False, "error": f"'{name}' is not an .xlsx file. "
                "Please drop an Excel .xlsx workbook."}

    session = secrets.token_hex(8)
    outdir = tempfile.mkdtemp(prefix="crd_")
    input_path = os.path.join(outdir, "_input_" + name)
    with open(input_path, "wb") as f:
        f.write(data)

    try:
        summary, warnings = run_analysis(input_path, outdir, name)
    except Exception as exc:  # bad/corrupt/wrong-shape workbook
        return {"ok": False, "error": f"Could not analyze '{name}': {exc}"}
    finally:
        try:
            os.remove(input_path)
        except OSError:
            pass

    with SESSIONS_LOCK:
        SESSIONS[session] = outdir

    files = []
    for fname, desc in OUTPUT_FILES:
        p = os.path.join(outdir, fname)
        if os.path.isfile(p):
            files.append({"name": fname, "desc": desc,
                          "size": _human_size(os.path.getsize(p))})
    return {
        "ok": True,
        "session": session,
        "source": name,
        "files": files,
        "kpis": [{"label": k, "value": _fmt(v)} for k, v in summary.items()],
        "warnings": list(warnings),
    }


def _fmt(v):
    return f"{v:,}" if isinstance(v, (int, float)) else str(v)


def _human_size(n):
    size = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def _safe_output_path(session, name):
    """Return an absolute path inside the session dir, or None if unsafe."""
    with SESSIONS_LOCK:
        outdir = SESSIONS.get(session)
    if not outdir:
        return None
    name = os.path.basename(urllib.parse.unquote(name))
    path = os.path.realpath(os.path.join(outdir, name))
    if os.path.commonpath([path, os.path.realpath(outdir)]) != os.path.realpath(outdir):
        return None
    return path if os.path.isfile(path) else None


CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".csv": "text/csv; charset=utf-8",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".zip": "application/zip",
}


# --------------------------------------------------------------------------- #
# CSV -> readable HTML table (for the in-app "View" button)
# --------------------------------------------------------------------------- #
_VIEWER_CSS = """
:root{--bg:#f4f6fb;--card:#fff;--ink:#1f2733;--muted:#5b6675;--line:#e2e7f0;
--accent:#305496;}
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
background:var(--bg);color:var(--ink);font-size:14px;line-height:1.45}
header{background:var(--accent);color:#fff;padding:16px 22px;display:flex;
align-items:center;gap:14px;flex-wrap:wrap}
header h1{margin:0;font-size:17px;font-weight:650;flex:1;min-width:220px}
header .count{opacity:.85;font-size:12.5px;font-weight:400}
.btn{display:inline-block;border:1px solid #fff;background:#fff;color:var(--accent);
padding:6px 12px;border-radius:8px;font-size:13px;text-decoration:none;font-weight:600}
.btn.ghost{background:transparent;color:#fff}
main{padding:16px 22px 50px}
.tools{margin-bottom:12px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
#q{padding:8px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px;
min-width:260px;background:var(--card);color:var(--ink)}
#shown{color:var(--muted);font-size:12.5px}
.tablewrap{overflow:auto;max-height:calc(100vh - 190px);border:1px solid var(--line);
border-radius:10px;background:var(--card)}
table{border-collapse:collapse;width:100%;font-size:13px}
thead th{position:sticky;top:0;background:var(--accent);color:#fff;text-align:left;
padding:9px 11px;font-weight:600;white-space:nowrap;z-index:1}
tbody td{padding:6px 11px;border-bottom:1px solid var(--line);vertical-align:top;
max-width:340px;overflow-wrap:anywhere}
tbody tr:nth-child(even){background:#fafbfe}
td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
td.mark{text-align:center;font-weight:700;color:var(--accent)}
tbody tr.hide{display:none}
@media (prefers-color-scheme:dark){
:root{--bg:#12161d;--card:#1a2029;--ink:#e6ebf2;--muted:#9aa6b6;--line:#2a323d;
--accent:#3a5da8}
tbody tr:nth-child(even){background:#1e2530}
.btn{background:#e9eef8;border-color:#e9eef8}}
"""


def _looks_numeric(s):
    """True for plain numbers like '1234', '1,234', '81', '1.5' (not '1.5 dCi')."""
    t = s.strip().replace(",", "")
    if not t or t in "-+":
        return False
    try:
        float(t)
        return True
    except ValueError:
        return False


def _csv_to_html(path, name, session):
    """Render a CSV file as a readable, searchable HTML table page."""
    # Uses the shared reader so the Excel 'sep=' hint line is skipped.
    table = analyze_crd.read_csv_rows(path)

    head = table[0] if table else []
    body = table[1:] if len(table) > 1 else []

    thead = "".join(f"<th>{_html.escape(h)}</th>" for h in head)

    rows_html = []
    for row in body:
        # Pad short rows so columns stay aligned with the header.
        if len(row) < len(head):
            row = row + [""] * (len(head) - len(row))
        cells = []
        for v in row:
            v = v or ""
            if v.strip().lower() == "x":
                cells.append(f'<td class="mark">{_html.escape(v)}</td>')
            elif _looks_numeric(v):
                cells.append(f'<td class="num">{_html.escape(v)}</td>')
            else:
                esc = _html.escape(v)
                # Keep long text scannable: full value available on hover.
                title = f' title="{esc}"' if len(v) > 60 else ""
                cells.append(f"<td{title}>{esc}</td>")
        rows_html.append("<tr>" + "".join(cells) + "</tr>")

    n = len(body)
    dl = f"/download/{session}/{urllib.parse.quote(name)}"
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_html.escape(name)} - CRD Analyzer</title>
<style>{_VIEWER_CSS}</style></head><body>
<header>
  <h1>{_html.escape(name)} <span class="count">&middot; {n:,} rows</span></h1>
  <a class="btn ghost" href="/">&#8592; Back to results</a>
  <a class="btn" href="{dl}">&#11015; Download this file</a>
</header>
<main>
  <div class="tools">
    <input id="q" type="search" placeholder="Filter rows... (e.g. RENAULT, 1.5 dCi)">
    <span id="shown">{n:,} of {n:,} rows</span>
  </div>
  <div class="tablewrap">
    <table><thead><tr>{thead}</tr></thead>
    <tbody>{''.join(rows_html)}</tbody></table>
  </div>
</main>
<script>
const q = document.getElementById('q');
const shown = document.getElementById('shown');
const rows = Array.from(document.querySelectorAll('tbody tr'));
const total = rows.length;
const fmt = n => n.toLocaleString();
q.addEventListener('input', () => {{
  const term = q.value.trim().toLowerCase();
  let count = 0;
  rows.forEach(tr => {{
    const hit = !term || tr.textContent.toLowerCase().includes(term);
    tr.classList.toggle('hide', !hit);
    if (hit) count++;
  }});
  shown.textContent = fmt(count) + ' of ' + fmt(total) + ' rows';
}});
</script>
</body></html>"""


# --------------------------------------------------------------------------- #
# HTTP handler
# --------------------------------------------------------------------------- #
class Handler(BaseHTTPRequestHandler):
    server_version = "CRDAnalyzer"

    def log_message(self, *args):
        pass  # keep the console quiet

    def _send(self, code, ctype, body, extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except BrokenPipeError:
            pass

    def _send_json(self, obj, code=200):
        self._send(code, "application/json; charset=utf-8", json.dumps(obj))

    # -- GET ---------------------------------------------------------------- #
    def do_GET(self):
        parts = urllib.parse.urlparse(self.path)
        path = parts.path
        if path in ("/", "/index.html"):
            self._send(200, "text/html; charset=utf-8", PAGE_HTML)
            return

        seg = [s for s in path.split("/") if s != ""]
        if len(seg) == 3 and seg[0] in ("view", "download"):
            p = _safe_output_path(seg[1], seg[2])
            if not p:
                self._send(404, "text/plain", "Not found")
                return
            ext = os.path.splitext(p)[1].lower()
            # For "view", show HTML/CSV inline; xlsx has no inline view -> download.
            inline = seg[0] == "view" and ext in (".html", ".csv")
            ctype = CONTENT_TYPES.get(ext, "application/octet-stream")
            if seg[0] == "view" and ext == ".csv":
                # Render the CSV as a readable table instead of raw text.
                page = _csv_to_html(p, os.path.basename(p), seg[1])
                self._send(200, "text/html; charset=utf-8", page)
                return
            with open(p, "rb") as f:
                data = f.read()
            extra = {}
            if not inline:
                extra["Content-Disposition"] = f'attachment; filename="{os.path.basename(p)}"'
            self._send(200, ctype, data, extra)
            return

        if len(seg) == 2 and seg[0] == "download_all":
            with SESSIONS_LOCK:
                outdir = SESSIONS.get(seg[1])
            if not outdir:
                self._send(404, "text/plain", "Not found")
                return
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for fname, _ in OUTPUT_FILES:
                    p = os.path.join(outdir, fname)
                    if os.path.isfile(p):
                        z.write(p, fname)
            self._send(200, "application/zip", buf.getvalue(),
                       {"Content-Disposition": 'attachment; filename="CRD_Analysis.zip"'})
            return

        self._send(404, "text/plain", "Not found")

    # -- POST --------------------------------------------------------------- #
    def do_POST(self):
        parts = urllib.parse.urlparse(self.path)
        if parts.path == "/analyze":
            length = int(self.headers.get("Content-Length", 0) or 0)
            if length <= 0:
                self._send_json({"ok": False, "error": "No file received."}, 400)
                return
            data = self.rfile.read(length)
            filename = self.headers.get("X-Filename", "upload.xlsx")
            try:
                result = analyze_upload(data, filename)
            except Exception as exc:
                result = {"ok": False, "error": f"Unexpected error: {exc}"}
            self._send_json(result, 200 if result.get("ok") else 400)
            return

        if parts.path == "/quit":
            self._send_json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return

        self._send(404, "text/plain", "Not found")


# --------------------------------------------------------------------------- #
# Front-end (single self-contained page)
# --------------------------------------------------------------------------- #
PAGE_HTML = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CRD Analyzer</title>
<style>
:root{--bg:#f4f6fb;--card:#fff;--ink:#1f2733;--muted:#5b6675;--line:#e2e7f0;
--accent:#305496;--accent2:#4472c4;--good:#e9f0ff;--danger:#c0392b;}
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
background:var(--bg);color:var(--ink);font-size:14px;line-height:1.45}
header{background:var(--accent);color:#fff;padding:20px 26px}
header h1{margin:0;font-size:20px;font-weight:650}
header .sub{opacity:.85;font-size:12.5px;margin-top:3px}
main{max-width:900px;margin:0 auto;padding:24px 20px 60px}
#drop{background:var(--card);border:2.5px dashed #b9c6de;border-radius:14px;
padding:46px 24px;text-align:center;cursor:pointer;transition:.15s;color:var(--muted)}
#drop.drag{border-color:var(--accent2);background:var(--good);color:var(--accent)}
#drop .big{font-size:17px;color:var(--ink);font-weight:600;margin-bottom:6px}
#drop .em{font-size:34px;margin-bottom:8px}
.btn{display:inline-block;border:1px solid var(--accent);background:var(--accent);
color:#fff;padding:7px 13px;border-radius:8px;font-size:13px;text-decoration:none;
cursor:pointer;font-weight:550}
.btn.ghost{background:#fff;color:var(--accent)}
.btn:hover{filter:brightness(1.05)}
.hidden{display:none}
#status{margin:20px 0;color:var(--muted)}
.spinner{display:inline-block;width:15px;height:15px;border:2.5px solid #ccd6ea;
border-top-color:var(--accent);border-radius:50%;animation:spin .8s linear infinite;
vertical-align:-3px;margin-right:8px}
@keyframes spin{to{transform:rotate(360deg)}}
.err{background:#fdecea;border:1px solid #f5c6c2;color:var(--danger);
padding:12px 14px;border-radius:10px;margin:16px 0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:6px 0 20px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px}
.kpi .v{font-size:22px;font-weight:700;color:var(--accent)}
.kpi .l{font-size:11.5px;color:var(--muted);margin-top:2px}
.section-title{font-size:15px;font-weight:650;margin:22px 0 10px}
.file{background:var(--card);border:1px solid var(--line);border-radius:10px;
padding:12px 14px;margin-bottom:10px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.file .meta{flex:1;min-width:180px}
.file .n{font-weight:600}
.file .d{color:var(--muted);font-size:12.5px}
.file .s{color:var(--muted);font-size:12px;white-space:nowrap}
.toolbar{display:flex;gap:10px;flex-wrap:wrap;margin:6px 0 18px;align-items:center}
.warn{background:#fff8e6;border:1px solid #ffe2a8;color:#7a5b00;padding:10px 13px;
border-radius:10px;margin:6px 0 16px;font-size:12.5px}
footer{max-width:900px;margin:0 auto;padding:0 20px 40px;color:var(--muted);font-size:12px}
@media (prefers-color-scheme:dark){
:root{--bg:#12161d;--card:#1a2029;--ink:#e6ebf2;--muted:#9aa6b6;--line:#2a323d;
--accent:#3a5da8;--good:#1c2740}
#drop{border-color:#33405a}}
</style></head>
<body>
<header>
  <h1>CRD Analyzer</h1>
  <div class="sub">Drag in a vehicle-application / carpark Excel file and get the full analysis.</div>
</header>
<main>
  <div id="drop">
    <div class="em">&#128228;</div>
    <div class="big">Drag your .xlsx file here</div>
    <div>or click to choose a file</div>
    <input id="file" type="file" accept=".xlsx,.xlsm" class="hidden">
  </div>

  <div id="status" class="hidden"></div>
  <div id="error" class="err hidden"></div>
  <div id="results" class="hidden"></div>
</main>
<footer>
  Runs locally on your PC - nothing is uploaded anywhere.
  <a href="#" id="quit" style="color:inherit">Quit the app</a>
</footer>

<script>
const drop = document.getElementById('drop');
const fileInput = document.getElementById('file');
const statusEl = document.getElementById('status');
const errorEl = document.getElementById('error');
const resultsEl = document.getElementById('results');

drop.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', e => { if (e.target.files[0]) analyze(e.target.files[0]); });
['dragenter','dragover'].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.add('drag');
}));
['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.remove('drag');
}));
drop.addEventListener('drop', e => {
  const f = e.dataTransfer.files[0];
  if (f) analyze(f);
});

function show(el){ el.classList.remove('hidden'); }
function hide(el){ el.classList.add('hidden'); }

async function analyze(file){
  hide(errorEl); hide(resultsEl);
  statusEl.innerHTML = '<span class="spinner"></span>Analyzing <b>' + esc(file.name) + '</b> ...';
  show(statusEl);
  try{
    const buf = await file.arrayBuffer();
    const res = await fetch('/analyze', {
      method:'POST',
      headers:{'X-Filename': encodeURIComponent(file.name)},
      body: buf
    });
    const data = await res.json();
    hide(statusEl);
    if(!data.ok){ errorEl.textContent = data.error || 'Analysis failed.'; show(errorEl); return; }
    render(data);
  }catch(err){
    hide(statusEl);
    errorEl.textContent = 'Something went wrong: ' + err;
    show(errorEl);
  }
}

function render(d){
  let h = '';
  h += '<div class="section-title">Results for ' + esc(d.source) + '</div>';
  h += '<div class="kpis">';
  d.kpis.forEach(k => {
    h += '<div class="kpi"><div class="v">' + esc(k.value) + '</div><div class="l">' + esc(k.label) + '</div></div>';
  });
  h += '</div>';
  if(d.warnings && d.warnings.length){
    h += '<div class="warn"><b>Notes:</b><ul style="margin:6px 0 0;padding-left:18px">';
    d.warnings.forEach(w => h += '<li>' + esc(w) + '</li>');
    h += '</ul></div>';
  }
  h += '<div class="section-title">Documents</div>';
  h += '<div class="warn">Working in Excel? Download <b>CRD_Analysis.xlsx</b> &ndash; ' +
       'it has every table as a properly formatted sheet. The .csv files are plain ' +
       'text exports for other tools.</div>';
  h += '<div class="toolbar"><a class="btn" href="/download_all/' + d.session + '">&#11015; Download all (.zip)</a>';
  h += '<a class="btn ghost" href="#" onclick="reset();return false;">Analyze another file</a></div>';
  d.files.forEach(f => {
    const view = '/view/' + d.session + '/' + encodeURIComponent(f.name);
    const dl = '/download/' + d.session + '/' + encodeURIComponent(f.name);
    // Only .html and .csv can be shown in the browser; .xlsx opens in Excel.
    const canView = /\.(html|csv)$/i.test(f.name);
    h += '<div class="file"><div class="meta"><div class="n">' + esc(f.name) + '</div>' +
         '<div class="d">' + esc(f.desc) + '</div></div>' +
         '<div class="s">' + esc(f.size) + '</div>' +
         (canView ? '<a class="btn ghost" target="_blank" href="' + view + '">View</a>' : '') +
         '<a class="btn" href="' + dl + '">Download</a></div>';
  });
  resultsEl.innerHTML = h;
  show(resultsEl);
  window.scrollTo({top: 0, behavior:'smooth'});
}

function reset(){
  hide(resultsEl); hide(errorEl); fileInput.value='';
}

document.getElementById('quit').addEventListener('click', async e => {
  e.preventDefault();
  try{ await fetch('/quit', {method:'POST'}); }catch(_){}
  document.body.innerHTML = '<main><p style="padding:40px 24px;color:#5b6675">' +
    'The CRD Analyzer has stopped. You can close this browser tab.</p></main>';
});

function esc(s){
  return String(s).replace(/[&<>"']/g, c => (
    {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
</script>
</body></html>
"""


# --------------------------------------------------------------------------- #
# Launcher
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="CRD Analyzer desktop app")
    ap.add_argument("--port", type=int, default=0,
                    help="Port to bind (default: an automatically chosen free port)")
    ap.add_argument("--host", default="127.0.0.1", help="Host to bind (default 127.0.0.1)")
    ap.add_argument("--no-browser", action="store_true",
                    help="Do not open a browser (for testing / headless use)")
    args = ap.parse_args(argv)

    server = ThreadingHTTPServer((args.host, args.port), Handler)
    port = server.server_address[1]
    url = f"http://{args.host}:{port}/"
    print("=" * 56)
    print("  CRD Analyzer is running.")
    print(f"  Open this in your browser:  {url}")
    print("  (Close this window, or click 'Quit the app', to stop.)")
    print("=" * 56)

    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
