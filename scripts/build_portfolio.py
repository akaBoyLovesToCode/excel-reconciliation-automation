"""Generate genuine portfolio visual assets and the 3-page PDF case study.

Usage:
    # Generate static workflow architecture diagram (no third-party Python deps required):
    python scripts/build_portfolio.py --workflow-only

    # Full end-to-end execution + real Excel rendering + 3-page PDF case study
    # (requires polars and openpyxl in active Python interpreter):
    PYTHONPATH=src python scripts/build_portfolio.py
"""

from __future__ import annotations

import argparse
import base64
import html
import importlib.util
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
PORTFOLIO_DIR = REPO_ROOT / "portfolio"
CHROME_APP_BIN = Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")

_ARROW_SVG = (
    '<svg class="arrow-svg" viewBox="0 0 24 24" fill="none" stroke-width="2.4" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M5 12h14M13 5l7 7-7 7"/></svg>'
)


def _find_chrome_binary() -> Path:
    """Locate Google Chrome or Chromium binary for headless rendering."""
    if CHROME_APP_BIN.exists():
        return CHROME_APP_BIN
    for candidate in ("google-chrome", "chromium", "chromium-browser"):
        found = shutil.which(candidate)
        if found:
            return Path(found)
    raise FileNotFoundError("Google Chrome binary not found for headless diagram/PDF rendering.")


def _workflow_html() -> str:
    """Return self-contained high-DPI HTML/SVG for 04_workflow.png based on actual modules."""
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    width: 1600px;
    height: 920px;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, sans-serif;
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
    color: #f8fafc;
    padding: 44px 52px;
    overflow: hidden;
  }}
  .header {{
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    border-bottom: 1px solid rgba(148, 163, 184, 0.25);
    padding-bottom: 20px;
    margin-bottom: 34px;
  }}
  .title-block h1 {{
    font-size: 30px;
    font-weight: 700;
    letter-spacing: -0.4px;
    color: #ffffff;
  }}
  .title-block p {{
    font-size: 15px;
    color: #94a3b8;
    margin-top: 6px;
  }}
  .badge {{
    background: rgba(56, 189, 248, 0.14);
    border: 1px solid rgba(56, 189, 248, 0.4);
    color: #38bdf8;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 999px;
  }}
  .pipeline {{
    display: grid;
    grid-template-columns: 1.05fr 36px 1.05fr 36px 1.15fr 36px 1.15fr 36px 0.95fr 36px 1.1fr;
    align-items: stretch;
    height: 540px;
  }}
  .stage {{
    background: rgba(30, 41, 59, 0.85);
    border: 1px solid rgba(148, 163, 184, 0.22);
    border-radius: 14px;
    padding: 20px 18px;
    display: flex;
    flex-direction: column;
    box-shadow: 0 12px 28px rgba(0, 0, 0, 0.28);
  }}
  .stage-num {{
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #38bdf8;
    margin-bottom: 6px;
  }}
  .stage h2 {{
    font-size: 18px;
    font-weight: 700;
    color: #ffffff;
    margin-bottom: 4px;
  }}
  .module-tag {{
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-size: 11.5px;
    background: rgba(15, 23, 42, 0.7);
    padding: 3px 8px;
    border-radius: 5px;
    align-self: flex-start;
    margin-bottom: 14px;
    color: #cbd5e1;
  }}
  .items {{
    display: flex;
    flex-direction: column;
    gap: 9px;
    flex: 1;
  }}
  .card {{
    background: rgba(15, 23, 42, 0.65);
    border-left: 3px solid #38bdf8;
    padding: 9px 11px;
    border-radius: 6px;
  }}
  .card.green {{ border-left-color: #22c55e; }}
  .card.amber {{ border-left-color: #f59e0b; }}
  .card.purple {{ border-left-color: #a855f7; }}
  .card.rose {{ border-left-color: #f43f5e; }}
  .card-title {{
    font-size: 12.5px;
    font-weight: 600;
    color: #f1f5f9;
  }}
  .card-sub {{
    font-size: 11px;
    color: #94a3b8;
    margin-top: 3px;
    line-height: 1.35;
  }}
  .arrow-col {{
    display: flex;
    align-items: center;
    justify-content: center;
  }}
  .arrow-svg {{
    width: 28px;
    height: 28px;
    stroke: #38bdf8;
  }}
  .footer-bar {{
    margin-top: 28px;
    background: rgba(15, 23, 42, 0.78);
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-radius: 12px;
    padding: 16px 24px;
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 20px;
  }}
  .guarantee-label {{
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: #38bdf8;
  }}
  .guarantee-text {{
    font-size: 12.5px;
    color: #e2e8f0;
    margin-top: 4px;
  }}
</style>
</head>
<body>
  <div class="header">
    <div class="title-block">
      <h1>Excel Reconciliation Automation — Pipeline Architecture</h1>
      <p>
        Deterministic Data Cleaning, Multi-Tier Record Matching
        &amp; Three-Way Financial Reconciliation
      </p>
    </div>
    <div class="badge">Python 3.12+ &bull; Polars &bull; openpyxl &bull; Decimal</div>
  </div>

  <div class="pipeline">
    <div class="stage">
      <div class="stage-num">Stage 01</div>
      <h2>Excel Inputs</h2>
      <div class="module-tag">excel_io.py</div>
      <div class="items">
        <div class="card">
          <div class="card-title">product_master.xlsx</div>
          <div class="card-sub">Canonical SKUs, supplier IDs, UOMs &amp; legacy aliases</div>
        </div>
        <div class="card">
          <div class="card-title">purchase_orders.xlsx</div>
          <div class="card-sub">Ordered quantities, contracted unit prices &amp; formulas</div>
        </div>
        <div class="card">
          <div class="card-title">shipments.xlsx</div>
          <div class="card-sub">Warehouse/ASN receipts, split &amp; partial deliveries</div>
        </div>
        <div class="card">
          <div class="card-title">supplier_invoices.xlsx</div>
          <div class="card-sub">AP billing lines, unit prices &amp; extended totals</div>
        </div>
        <div class="card purple">
          <div class="card-title">SHA-256 + Formula Reader</div>
          <div class="card-sub">Dual-pass openpyxl load preserving raw formulas</div>
        </div>
      </div>
    </div>

    <div class="arrow-col">{_ARROW_SVG}</div>

    <div class="stage">
      <div class="stage-num">Stage 02</div>
      <h2>Data Cleaning</h2>
      <div class="module-tag">cleaning.py</div>
      <div class="items">
        <div class="card">
          <div class="card-title">Schema Validation</div>
          <div class="card-sub">Header verification without silent column drops</div>
        </div>
        <div class="card">
          <div class="card-title">SKU &amp; String Hygiene</div>
          <div class="card-sub">Separator/case canonicalization; Unicode preserved</div>
        </div>
        <div class="card amber">
          <div class="card-title">Semantic Deduplication</div>
          <div class="card-sub">
            Structured identity groups; all conflicting versions quarantined.
          </div>
        </div>
        <div class="card rose">
          <div class="card-title">Type &amp; Formula Checks</div>
          <div class="card-sub">Decimal/ISO-date parsing; #REF! / #DIV/0! capture</div>
        </div>
        <div class="card green">
          <div class="card-title">Raw Value Audit Copies</div>
          <div class="card-sub">100% row retention with raw_* columns intact</div>
        </div>
      </div>
    </div>

    <div class="arrow-col">{_ARROW_SVG}</div>

    <div class="stage">
      <div class="stage-num">Stage 03</div>
      <h2>Record Matching</h2>
      <div class="module-tag">matching.py</div>
      <div class="items">
        <div class="card green">
          <div class="card-title">Tier 1: Exact ID Match</div>
          <div class="card-sub">Normalized canonical SKU + (po_number, po_line_id)</div>
        </div>
        <div class="card">
          <div class="card-title">Tier 2: Explicit Alias Map</div>
          <div class="card-sub">Verified 1-to-1 supplier/legacy SKU resolution</div>
        </div>
        <div class="card purple">
          <div class="card-title">Tier 3: Composite Key</div>
          <div class="card-sub">Unique (po_number, product_id) when line ID absent</div>
        </div>
        <div class="card rose">
          <div class="card-title">Tier 4: Ambiguity Quarantine</div>
          <div class="card-sub">Multi-candidate aliases &amp; multi-line POs never guessed</div>
        </div>
      </div>
    </div>

    <div class="arrow-col">{_ARROW_SVG}</div>

    <div class="stage">
      <div class="stage-num">Stage 04</div>
      <h2>Reconciliation</h2>
      <div class="module-tag">reconciliation.py</div>
      <div class="items">
        <div class="card green">
          <div class="card-title">Polars Pre-Aggregation</div>
          <div class="card-sub">Independent group_by eliminates M&times;N join inflation</div>
        </div>
        <div class="card">
          <div class="card-title">Decimal Financial Math</div>
          <div class="card-sub">ROUND_HALF_UP (0.01) with explicit qty/price tolerances</div>
        </div>
        <div class="card amber">
          <div class="card-title">Zero vs. Missing Semantics</div>
          <div class="card-sub">Valid 0-qty notices distinguished from missing docs</div>
        </div>
        <div class="card rose">
          <div class="card-title">9 Exception Categories</div>
          <div class="card-sub">Price/qty variances, unbilled receipts, rogue invoices</div>
        </div>
      </div>
    </div>

    <div class="arrow-col">{_ARROW_SVG}</div>

    <div class="stage">
      <div class="stage-num">Stage 05</div>
      <h2>Excel Report</h2>
      <div class="module-tag">reporting.py</div>
      <div class="items">
        <div class="card green">
          <div class="card-title">1. Summary</div>
          <div class="card-sub">Executive KPIs, live COUNTIF/SUM checks &amp; breakdown</div>
        </div>
        <div class="card">
          <div class="card-title">2. Reconciled Records</div>
          <div class="card-sub">Raw + normalized IDs, 3-way quantities &amp; variances</div>
        </div>
        <div class="card rose">
          <div class="card-title">3. Exceptions</div>
          <div class="card-sub">Severity-ranked queue, source row &amp; next action</div>
        </div>
        <div class="card purple">
          <div class="card-title">4. Audit Log</div>
          <div class="card-sub">UTC timestamp, SHA-256 hashes, rules &amp; warnings</div>
        </div>
      </div>
    </div>

    <div class="arrow-col">{_ARROW_SVG}</div>

    <div class="stage">
      <div class="stage-num">Stage 06</div>
      <h2>Validation</h2>
      <div class="module-tag">validation.py</div>
      <div class="items">
        <div class="card green">
          <div class="card-title">Structure &amp; Panes</div>
          <div class="card-sub">Verifies 4 required tabs, frozen headers &amp; filters</div>
        </div>
        <div class="card">
          <div class="card-title">Formula Integrity Scan</div>
          <div class="card-sub">Zero #REF!, #NAME?, #VALUE! or #DIV/0! tokens</div>
        </div>
        <div class="card purple">
          <div class="card-title">Cross-Tab KPI Audit</div>
          <div class="card-sub">Reconciles Summary totals against detail row counts</div>
        </div>
      </div>
    </div>
  </div>

  <div class="footer-bar">
    <div>
      <div class="guarantee-label">Zero Silent Drops</div>
      <div class="guarantee-text">
        Every invalid, duplicate, or orphan row is preserved and traced to its exact source row.
      </div>
    </div>
    <div>
      <div class="guarantee-label">Deterministic Matching</div>
      <div class="guarantee-text">
        Multi-candidate SKUs and order lines are quarantined as AMBIGUOUS—never arbitrarily chosen.
      </div>
    </div>
    <div>
      <div class="guarantee-label">Cardinality-Safe Totals</div>
      <div class="guarantee-text">
        Split shipments and split invoices pre-aggregate independently before 3-way comparison.
      </div>
    </div>
    <div>
      <div class="guarantee-label">Exact Monetary Arithmetic</div>
      <div class="guarantee-text">
        All currency calculations use Python decimal.Decimal with commercial ROUND_HALF_UP rounding.
      </div>
    </div>
  </div>
</body>
</html>
"""


def render_workflow_diagram(output_png: Path) -> Path:
    """Render portfolio/04_workflow.png at 2x Retina resolution using headless Chrome."""
    chrome = _find_chrome_binary()
    output_png.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        html_path = Path(tmp) / "workflow.html"
        html_path.write_text(_workflow_html(), encoding="utf-8")
        subprocess.run(
            [
                str(chrome),
                "--headless=new",
                "--disable-gpu",
                "--hide-scrollbars",
                "--force-device-scale-factor=2",
                "--window-size=1600,920",
                f"--screenshot={output_png}",
                html_path.as_uri(),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
    return output_png


def _render_xlsx_via_quicklook(xlsx_path: Path, output_png: Path) -> bool:
    """Attempt native macOS QuickLook rendering of an actual .xlsx document."""
    qlmanage = shutil.which("qlmanage")
    if not qlmanage:
        return False
    with tempfile.TemporaryDirectory() as tmp:
        res = subprocess.run(
            [qlmanage, "-t", "-s", "2200", "-o", tmp, str(xlsx_path)],
            check=False,
            capture_output=True,
            text=True,
        )
        candidate = Path(tmp) / f"{xlsx_path.name}.png"
        if res.returncode == 0 and candidate.exists():
            output_png.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(candidate, output_png)
            return True
    return False


def _png_to_data_uri(png_path: Path) -> str:
    """Encode a PNG file as an inline base64 data URI for standalone PDF rendering."""
    encoded = base64.b64encode(png_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def build_case_study_pdf(
    output_pdf: Path,
    *,
    summary_metrics: dict[str, str] | None,
    screenshot_paths: dict[str, Path],
) -> Path:
    """Compile the 3-page portfolio/case_study.pdf using headless Chrome print-to-pdf."""
    chrome = _find_chrome_binary()
    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    wf_uri = (
        _png_to_data_uri(screenshot_paths["workflow"])
        if "workflow" in screenshot_paths and screenshot_paths["workflow"].exists()
        else ""
    )
    s1_uri = (
        _png_to_data_uri(screenshot_paths["before_after"])
        if "before_after" in screenshot_paths and screenshot_paths["before_after"].exists()
        else ""
    )
    s2_uri = (
        _png_to_data_uri(screenshot_paths["dashboard"])
        if "dashboard" in screenshot_paths and screenshot_paths["dashboard"].exists()
        else ""
    )
    s3_uri = (
        _png_to_data_uri(screenshot_paths["exceptions"])
        if "exceptions" in screenshot_paths and screenshot_paths["exceptions"].exists()
        else ""
    )

    assert summary_metrics is not None
    m = {k: html.escape(str(v)) for k, v in summary_metrics.items()}
    s1_tag = f'<img src="{s1_uri}" alt="Before and After">' if s1_uri else ""
    s2_tag = f'<img src="{s2_uri}" alt="Summary Dashboard">' if s2_uri else ""
    s3_tag = f'<img src="{s3_uri}" alt="Exception Report">' if s3_uri else ""
    wf_tag = f'<img src="{wf_uri}" alt="Workflow Architecture">' if wf_uri else ""

    pdf_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<style>
  @page {{ size: A4 portrait; margin: 14mm 15mm; }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, sans-serif;
    color: #0f172a;
    font-size: 10pt;
    line-height: 1.45;
  }}
  .page {{
    page-break-after: always;
    height: 267mm;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }}
  .page:last-child {{ page-break-after: auto; }}
  .banner {{
    background: #1b365d;
    color: #ffffff;
    padding: 18px 22px;
    border-radius: 8px;
    margin-bottom: 16px;
  }}
  .banner .eyebrow {{
    font-size: 8.5pt;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    color: #93c5fd;
    font-weight: 700;
  }}
  .banner h1 {{ font-size: 19pt; margin-top: 4px; }}
  .banner p {{ font-size: 10pt; color: #e2e8f0; margin-top: 5px; }}
  h2 {{
    font-size: 12.5pt;
    color: #1b365d;
    border-bottom: 2px solid #cbd5e1;
    padding-bottom: 4px;
    margin: 14px 0 8px 0;
  }}
  p {{ margin-bottom: 8px; color: #334155; }}
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }}
  .box {{
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 11px 13px;
  }}
  .box h3 {{ font-size: 10.5pt; color: #0f172a; margin-bottom: 5px; }}
  ul {{ padding-left: 16px; color: #334155; }}
  li {{ margin-bottom: 4px; font-size: 9.5pt; }}
  table {{
    width: 100%;
    border-collapse: collapse;
    margin: 8px 0;
    font-size: 9pt;
  }}
  th {{
    background: #1b365d;
    color: #ffffff;
    text-align: left;
    padding: 6px 8px;
  }}
  td {{
    border: 1px solid #cbd5e1;
    padding: 5px 8px;
    color: #1e293b;
  }}
  tr:nth-child(even) td {{ background: #f8fafc; }}
  .img-frame {{
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    overflow: hidden;
    margin: 5px 0;
    background: #ffffff;
    text-align: center;
  }}
  .img-frame img {{ width: 100%; height: auto; display: block; margin: 0 auto; }}
  .img-pair .img-frame img {{ max-height: 68mm; width: auto; max-width: 100%; }}
  .img-wide img {{ max-height: 82mm; width: auto; max-width: 100%; }}
  .img-arch img {{ max-height: 96mm; width: auto; max-width: 100%; }}
  .footer {{
    border-top: 1px solid #e2e8f0;
    padding-top: 6px;
    font-size: 8pt;
    color: #64748b;
    display: flex;
    justify-content: space-between;
  }}
  code {{
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-size: 8.8pt;
    background: #f1f5f9;
    padding: 1px 4px;
    border-radius: 3px;
  }}
</style>
</head>
<body>
  <div class="page">
    <div>
      <div class="banner">
        <div class="eyebrow">Technical Portfolio Case Study &bull; Python Data Engineering</div>
        <h1>Excel Reconciliation Automation</h1>
        <p>
          Automated Data Cleaning, Deterministic Record Matching
          &amp; Three-Way Financial Reconciliation
        </p>
      </div>

      <h2>1. Business Challenge</h2>
      <p>
        Procurement, logistics, and accounts payable teams in trading and e-commerce
        organizations routinely exchange separate Excel exports from ERP, warehouse
        management, and supplier billing systems. Before payment approval, analysts
        must reconcile <strong>Purchase Orders</strong>, <strong>Shipment Receipts</strong>,
        and <strong>Supplier Invoices</strong> against a central
        <strong>Product Master</strong>.
      </p>
      <p>
        In manual spreadsheet workflows, VLOOKUP/XLOOKUP chains break down when
        confronted with inconsistent SKU punctuation (<code>sku_elec_0042</code> vs.
        <code>SKU-ELEC-0042</code>), legacy supplier part numbers, split deliveries
        across multiple shipments, duplicate data exports, missing line identifiers,
        and malformed numeric or formula cells. Worse, naive table joins across split
        shipments and split invoices inflate quantities and dollar totals.
      </p>

      <h2>2. Automated Engineering Solution</h2>
      <p>
        <strong>Excel Reconciliation Automation</strong> replaces fragile manual
        spreadsheet work with an auditable, deterministic Python CLI pipeline built on
        <strong>Polars</strong>, <strong>openpyxl</strong>, and exact
        <strong>decimal.Decimal</strong> financial arithmetic. Every input workbook is
        hashed with SHA-256, cleaned without dropping invalid rows, matched across four
        deterministic priority tiers, reconciled at line level, exported to a formatted
        four-tab Excel workbook (<code>reconciliation_report.xlsx</code>), and then
        validated for report consistency and formula integrity.
      </p>

      <h2>3. Core Capabilities</h2>
      <div class="grid-2">
        <div class="box">
          <h3>Non-Destructive Data Hygiene</h3>
          <ul>
            <li>
              Preserves every original cell in <code>raw_*</code> audit columns
              alongside normalized values.
            </li>
            <li>Canonicalizes SKU separators, casing, whitespace, and Unicode text.</li>
            <li>
              Groups exports by structured identity and quarantines all conflicting
              versions without grouping null identities as duplicates.
            </li>
          </ul>
        </div>
        <div class="box">
          <h3>Deterministic Multi-Tier Matching</h3>
          <ul>
            <li><strong>Tier 1:</strong> Exact normalized SKU &amp; PO line ID match.</li>
            <li><strong>Tier 2:</strong> Explicit 1-to-1 product alias resolution.</li>
            <li>
              <strong>Tier 3:</strong> Composite <code>(po_number, product_id)</code>
              key when line ID is omitted.
            </li>
            <li>
              <strong>Tier 4:</strong> Quarantines ambiguous multi-candidate
              matches—never guesses.
            </li>
          </ul>
        </div>
        <div class="box">
          <h3>Cardinality-Safe 3-Way Reconciliation</h3>
          <ul>
            <li>
              Pre-aggregates shipments and invoices independently before joining to
              prevent M&times;N inflation.
            </li>
            <li>
              Uses <code>decimal.Decimal</code> (<code>ROUND_HALF_UP</code>) with
              explicit tolerances.
            </li>
            <li>
              Distinguishes valid zero-quantity notices from missing shipment or
              invoice records.
            </li>
          </ul>
        </div>
        <div class="box">
          <h3>Executive &amp; Operational Reporting</h3>
          <ul>
            <li>
              Generates <code>Summary</code>, <code>Reconciled Records</code>,
              <code>Exceptions</code>, and <code>Audit Log</code> tabs.
            </li>
            <li>
              Includes frozen headers, auto-filters, conditional formatting, and live
              Excel check formulas.
            </li>
            <li>
              Built-in <code>validate</code> command verifies zero broken formulas
              (<code>#REF!</code>, <code>#DIV/0!</code>).
            </li>
          </ul>
        </div>
      </div>
    </div>
    <div class="footer">
      <span>github.com/akaBoyLovesToCode/excel-reconciliation-automation</span>
      <span>Page 1 of 3</span>
    </div>
  </div>

  <div class="page">
    <div>
      <h2>4. Synthetic Benchmark Demonstration (Deterministic Seed: 20260401)</h2>
      <p>
        Evaluated on four deterministic synthetic workbooks (1,500 products,
        2,000 PO rows, 2,200 shipments, and 2,100 supplier invoices) containing
        injected formatting noise, split shipments, duplicate bills, ambiguous
        aliases, price overcharges, and broken formulas:
      </p>
      <table>
        <thead>
          <tr>
            <th>Metric</th>
            <th>Result</th>
            <th>Metric</th>
            <th>Result</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Total PO Rows Ingested</td>
            <td>{m["total_po_rows"]}</td>
            <td>Unique Valid PO Lines</td>
            <td>{m["unique_valid_po_lines"]}</td>
          </tr>
          <tr>
            <td>Product Matching Coverage</td>
            <td>{m["matching_coverage"]}</td>
            <td>Fully Reconciled PO Lines</td>
            <td>{m["reconciled_lines"]} ({m["reconciliation_rate"]})</td>
          </tr>
          <tr>
            <td>Partial Fulfillment Lines</td>
            <td>{m["partial_lines"]}</td>
            <td>Total Exception Records (All Sources)</td>
            <td>{m["total_exceptions"]}</td>
          </tr>
          <tr>
            <td>Expected PO Commitment</td>
            <td>{m["expected_amount"]}</td>
            <td>Actual Invoiced Total</td>
            <td>{m["invoiced_amount"]}</td>
          </tr>
          <tr>
            <td>Absolute Discrepancy Exposure</td>
            <td>{m["abs_discrepancy"]}</td>
            <td>Unbilled Shipped Exposure</td>
            <td>{m["unbilled_amount"]}</td>
          </tr>
          <tr>
            <td>Net Billed Monetary Variance</td>
            <td>{m["net_variance"]}</td>
            <td>Quarantined Invoice Amount</td>
            <td>{m["quarantined_invoice_amount"]}</td>
          </tr>
        </tbody>
      </table>

      <div class="grid-2 img-pair">
        <div>
          <div class="img-frame">{s1_tag}</div>
        </div>
        <div>
          <div class="img-frame">{s2_tag}</div>
        </div>
      </div>
      <div class="img-frame img-wide">{s3_tag}</div>
    </div>
    <div class="footer">
      <span>Synthetic Business Dataset Demonstration &bull; Computed from Execution</span>
      <span>Page 2 of 3</span>
    </div>
  </div>

  <div class="page">
    <div>
      <h2>5. Pipeline Architecture</h2>
      <div class="img-frame img-arch">{wf_tag}</div>

      <h2>6. Engineering Quality, Validation &amp; Reproducibility</h2>
      <div class="grid-2">
        <div class="box">
          <h3>Technology Stack</h3>
          <ul>
            <li>
              <strong>Python 3.12+:</strong> Strictly typed domain models
              (<code>dataclasses</code>, <code>StrEnum</code>, <code>Decimal</code>).
            </li>
            <li>
              <strong>Polars:</strong> Vectorized cleaning expressions and
              independent pre-aggregation.
            </li>
            <li>
              <strong>openpyxl:</strong> Dual-pass formula preservation, workbook
              styling, and conditional formatting.
            </li>
            <li>
              <strong>Ruff &amp; pytest:</strong> Automated linting, formatting, unit
              tests, and CI verification.
            </li>
          </ul>
        </div>
        <div class="box">
          <h3>Automated Test &amp; Audit Controls</h3>
          <ul>
            <li>
              Regression tests guarding against M&times;N join-cardinality inflation
              and duplicate billing inflation.
            </li>
            <li>
              End-to-end ground-truth verification across all 2,000 synthetic
              purchase order rows.
            </li>
            <li>
              Post-generation workbook validator scanning every cell for Excel formula
              errors and KPI parity.
            </li>
            <li>
              Portable <code>pyproject.toml</code> + GitHub Actions CI running clean
              builds on every push.
            </li>
          </ul>
        </div>
      </div>

      <h2>7. Command-Line Reproducibility</h2>
      <div class="box">
        <p style="margin-bottom:4px;">
          <code>python -m excel_reconciliation generate-demo</code>
        </p>
        <p style="margin-bottom:4px;">
          <code>
            python -m excel_reconciliation reconcile
            --input examples/input --output examples/output
          </code>
        </p>
        <p style="margin-bottom:0;">
          <code>
            python -m excel_reconciliation validate
            --report examples/output/reconciliation_report.xlsx
          </code>
        </p>
      </div>
    </div>
    <div class="footer">
      <span>coreyzhou.com &bull; MIT License</span>
      <span>Page 3 of 3</span>
    </div>
  </div>
</body>
</html>
"""
    with tempfile.TemporaryDirectory() as tmp:
        html_file = Path(tmp) / "case_study.html"
        html_file.write_text(pdf_html, encoding="utf-8")
        subprocess.run(
            [
                str(chrome),
                "--headless=new",
                "--disable-gpu",
                "--no-pdf-header-footer",
                f"--print-to-pdf={output_pdf}",
                html_file.as_uri(),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
    return output_pdf


def main() -> int:
    """Entry point for generating portfolio visual assets."""
    parser = argparse.ArgumentParser(description="Build portfolio diagrams, screenshots, and PDF.")
    parser.add_argument(
        "--workflow-only",
        action="store_true",
        help="Render only 04_workflow.png (does not require polars or openpyxl).",
    )
    args = parser.parse_args()

    wf_path = render_workflow_diagram(PORTFOLIO_DIR / "04_workflow.png")
    print(f"Rendered architecture diagram: {wf_path}")
    if args.workflow_only:
        return 0

    missing = [mod for mod in ("polars", "openpyxl") if importlib.util.find_spec(mod) is None]
    if missing:
        print(
            f"ERROR: Cannot generate Excel-rendered screenshots (01..03) and case_study.pdf "
            f"until missing dependencies are available: {', '.join(missing)}.",
            file=sys.stderr,
        )
        return 2

    sys.path.insert(0, str(REPO_ROOT / "src"))
    from excel_reconciliation.demo_data import generate_demo_datasets
    from excel_reconciliation.reconciliation import run_reconciliation_from_directory
    from excel_reconciliation.reporting import write_reconciliation_report, write_safe_cell
    from excel_reconciliation.validation import validate_report_workbook

    in_dir = REPO_ROOT / "examples" / "input"
    out_dir = REPO_ROOT / "examples" / "output"
    report_path = out_dir / "reconciliation_report.xlsx"

    generate_demo_datasets(in_dir)
    result = run_reconciliation_from_directory(in_dir)
    write_reconciliation_report(result, report_path)
    val = validate_report_workbook(report_path)
    if not val.is_valid:
        raise RuntimeError(f"Generated report failed validation: {val.errors}")

    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    p_before = PORTFOLIO_DIR / "01_before_after.png"
    p_dash = PORTFOLIO_DIR / "02_summary_dashboard.png"
    p_exc = PORTFOLIO_DIR / "03_exception_report.png"

    # 1. Dashboard view rendered directly from reconciliation_report.xlsx (Summary is sheet 1)
    if not _render_xlsx_via_quicklook(report_path, p_dash):
        raise RuntimeError("macOS QuickLook failed to render 02_summary_dashboard.png")

    with tempfile.TemporaryDirectory() as tmp:
        tdir = Path(tmp)
        wb_rep = openpyxl.load_workbook(report_path, data_only=False)

        navy_fill = PatternFill("solid", fgColor="1B365D")
        amber_fill = PatternFill("solid", fgColor="92400E")
        green_fill = PatternFill("solid", fgColor="166534")
        zebra_fill = PatternFill("solid", fgColor="F8FAFC")
        hdr_font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        banner_font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
        body_font = Font(name="Calibri", size=8.5, color="0F172A")
        mono_font = Font(name="Consolas", size=8, color="1E293B")
        thin_side = Side(style="thin", color="CBD5E1")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        status_fills: dict[str, tuple[PatternFill, Font]] = {
            "RECONCILED": (
                PatternFill("solid", fgColor="DCFCE7"),
                Font(name="Calibri", size=8.5, bold=True, color="166534"),
            ),
            "PARTIAL_FULFILLMENT": (
                PatternFill("solid", fgColor="DBEAFE"),
                Font(name="Calibri", size=8.5, bold=True, color="1E40AF"),
            ),
            "QUANTITY_MISMATCH": (
                PatternFill("solid", fgColor="FEE2E2"),
                Font(name="Calibri", size=8.5, bold=True, color="991B1B"),
            ),
            "PRICE_MISMATCH": (
                PatternFill("solid", fgColor="FEE2E2"),
                Font(name="Calibri", size=8.5, bold=True, color="991B1B"),
            ),
            "MISSING_SHIPMENT": (
                PatternFill("solid", fgColor="FEF3C7"),
                Font(name="Calibri", size=8.5, bold=True, color="92400E"),
            ),
            "MISSING_INVOICE": (
                PatternFill("solid", fgColor="FEF3C7"),
                Font(name="Calibri", size=8.5, bold=True, color="92400E"),
            ),
            "AMBIGUOUS_REFERENCE": (
                PatternFill("solid", fgColor="F3E8FF"),
                Font(name="Calibri", size=8.5, bold=True, color="6B21A8"),
            ),
            "DUPLICATE_TRANSACTION": (
                PatternFill("solid", fgColor="FEF3C7"),
                Font(name="Calibri", size=8.5, bold=True, color="92400E"),
            ),
            "DATA_VALIDATION_ERROR": (
                PatternFill("solid", fgColor="FEE2E2"),
                Font(name="Calibri", size=8.5, bold=True, color="991B1B"),
            ),
            "UNEXPECTED_EXTRA_RECORD": (
                PatternFill("solid", fgColor="FEE2E2"),
                Font(name="Calibri", size=8.5, bold=True, color="991B1B"),
            ),
        }
        sev_fills: dict[str, tuple[PatternFill, Font]] = {
            "HIGH": (
                PatternFill("solid", fgColor="FEE2E2"),
                Font(name="Calibri", size=8.5, bold=True, color="991B1B"),
            ),
            "MEDIUM": (
                PatternFill("solid", fgColor="FEF3C7"),
                Font(name="Calibri", size=8.5, bold=True, color="92400E"),
            ),
            "LOW": (
                PatternFill("solid", fgColor="DBEAFE"),
                Font(name="Calibri", size=8.5, bold=True, color="1E40AF"),
            ),
        }

        # 2. Exception report view built from genuine Exceptions worksheet rows across categories
        ws_exc_src = wb_rep["Exceptions"]
        wb_exc_view = openpyxl.Workbook()
        ws_ev = wb_exc_view.active
        assert ws_ev is not None
        ws_ev.title = "Exceptions"
        ws_ev.page_setup.orientation = ws_ev.ORIENTATION_LANDSCAPE
        ws_ev.page_setup.paperSize = ws_ev.PAPERSIZE_TABLOID
        ws_ev.sheet_properties.pageSetUpPr.fitToPage = True
        ws_ev.page_setup.fitToWidth = 1
        ws_ev.page_setup.fitToHeight = 0

        ws_ev.merge_cells("A1:H1")
        ws_ev["A1"] = (
            f"Exception Review Queue — Representative Records from {ws_exc_src.max_row - 1:,} "
            "Flagged Exceptions (reconciliation_report.xlsx)"
        )
        ws_ev["A1"].fill = navy_fill
        ws_ev["A1"].font = banner_font
        ws_ev["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws_ev.row_dimensions[1].height = 26

        ev_headers = [
            "Exc ID",
            "Exception Category",
            "Sev",
            "Dataset : Row",
            "PO / Source Ref",
            "Normalized SKU",
            "Reason Codes",
            "Impact ($)",
        ]
        ws_ev.row_dimensions[2].height = 22
        for col_idx, hdr_val in enumerate(ev_headers, start=1):
            hc = write_safe_cell(ws_ev, 2, col_idx, hdr_val)
            hc.fill = navy_fill
            hc.font = hdr_font
            hc.alignment = Alignment(horizontal="center", vertical="center")
            hc.border = border

        exc_quotas = {
            "QUANTITY_MISMATCH": 3,
            "PRICE_MISMATCH": 3,
            "MISSING_SHIPMENT": 3,
            "MISSING_INVOICE": 3,
            "PARTIAL_FULFILLMENT": 3,
            "DUPLICATE_TRANSACTION": 3,
            "AMBIGUOUS_REFERENCE": 3,
            "UNEXPECTED_EXTRA_RECORD": 3,
            "DATA_VALIDATION_ERROR": 2,
        }
        picked_exc_rows: list[int] = []
        for r_idx in range(2, ws_exc_src.max_row + 1):
            cat_val = str(ws_exc_src.cell(row=r_idx, column=2).value or "")
            if exc_quotas.get(cat_val, 0) > 0:
                picked_exc_rows.append(r_idx)
                exc_quotas[cat_val] -= 1
            if sum(exc_quotas.values()) == 0:
                break

        for out_r, src_r in enumerate(picked_exc_rows, start=3):
            reasons = str(ws_exc_src.cell(row=src_r, column=11).value or "")
            ws_ev.row_dimensions[out_r].height = 26 if ";" in reasons else 23.5
            cat_str = str(ws_exc_src.cell(row=src_r, column=2).value or "")
            sev_str = str(ws_exc_src.cell(row=src_r, column=3).value or "")
            ds_str = str(ws_exc_src.cell(row=src_r, column=4).value or "")
            srow_val = ws_exc_src.cell(row=src_r, column=5).value
            po_ref = str(ws_exc_src.cell(row=src_r, column=6).value or "")
            norm_ref = str(ws_exc_src.cell(row=src_r, column=8).value or "")
            disp_ref = norm_ref if po_ref in ("", "MISSING") else po_ref
            sku_str = str(
                ws_exc_src.cell(row=src_r, column=10).value
                or ws_exc_src.cell(row=src_r, column=9).value
                or ""
            )
            mon_imp = ws_exc_src.cell(row=src_r, column=13).value

            ev_vals = [
                ws_exc_src.cell(row=src_r, column=1).value,
                cat_str,
                sev_str,
                f"{ds_str}:{srow_val}",
                disp_ref,
                sku_str,
                reasons,
                mon_imp,
            ]
            for c_i, cell_val in enumerate(ev_vals, start=1):
                c = write_safe_cell(ws_ev, out_r, c_i, cell_val)
                c.font = mono_font if c_i in (1, 5, 6, 7) else body_font
                c.border = border
                if out_r % 2 == 1:
                    c.fill = zebra_fill
                if c_i == 2 and cat_str in status_fills:
                    cf, cfont = status_fills[cat_str]
                    c.fill = cf
                    c.font = cfont
                elif c_i == 3 and sev_str in sev_fills:
                    sf, sfont = sev_fills[sev_str]
                    c.fill = sf
                    c.font = sfont
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i == 7:
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif c_i == 8:
                    c.number_format = '$#,##0.00;($#,##0.00);"$0.00"'
                    c.alignment = Alignment(horizontal="right", vertical="center")

        ev_widths = [8.5, 22.0, 6.5, 18.0, 15.0, 13.0, 28.0, 11.5]
        for idx_w, w_val in enumerate(ev_widths, start=1):
            ws_ev.column_dimensions[get_column_letter(idx_w)].width = w_val

        exc_xlsx = tdir / "exceptions_view.xlsx"
        wb_exc_view.save(exc_xlsx)
        wb_exc_view.close()
        if not _render_xlsx_via_quicklook(exc_xlsx, p_exc):
            raise RuntimeError("macOS QuickLook failed to render 03_exception_report.png")

        # 3. Before & After comparison workbook built directly from real input + output workbooks
        wb_po_raw = openpyxl.load_workbook(in_dir / "purchase_orders.xlsx", data_only=False)
        ws_po_raw = wb_po_raw.active
        ws_rec = wb_rep["Reconciled Records"]

        wb_cmp = openpyxl.Workbook()
        ws_cmp = wb_cmp.active
        assert ws_cmp is not None
        ws_cmp.title = "Before vs After"
        ws_cmp.page_setup.orientation = ws_cmp.ORIENTATION_LANDSCAPE
        ws_cmp.page_setup.paperSize = ws_cmp.PAPERSIZE_TABLOID
        ws_cmp.sheet_properties.pageSetUpPr.fitToPage = True
        ws_cmp.page_setup.fitToWidth = 1
        ws_cmp.page_setup.fitToHeight = 0

        ws_cmp.merge_cells("A1:D1")
        ws_cmp["A1"] = "BEFORE — Raw Input (purchase_orders.xlsx)"
        ws_cmp["A1"].fill = amber_fill
        ws_cmp["A1"].font = banner_font
        ws_cmp["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws_cmp.merge_cells("F1:J1")
        ws_cmp["F1"] = "AFTER — Cleaned, Matched & Reconciled (reconciliation_report.xlsx)"
        ws_cmp["F1"].fill = green_fill
        ws_cmp["F1"].font = banner_font
        ws_cmp["F1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_cmp.row_dimensions[1].height = 26

        cmp_headers = [
            "Row",
            "Raw PO Input",
            "Raw SKU / Alias",
            "Raw Qty @ Price",
            "",
            "Canonical PO Line",
            "Canonical SKU",
            "Match Tier",
            "Ord/Shp/Inv",
            "Reconciliation Status",
        ]
        ws_cmp.row_dimensions[2].height = 22
        for col_idx, text in enumerate(cmp_headers, start=1):
            if text:
                cell = write_safe_cell(ws_cmp, 2, col_idx, text)
                cell.fill = navy_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

        selected_rows: list[int] = []
        bucket_quotas = {
            "DIRTY_RECONCILED": 5,
            "ALIAS_RECONCILED": 5,
            "PARTIAL_FULFILLMENT": 3,
            "PRICE_MISMATCH": 3,
            "QUANTITY_MISMATCH": 3,
            "MISSING_SHIPMENT": 2,
            "MISSING_INVOICE": 2,
            "AMBIGUOUS_REFERENCE": 2,
            "DUPLICATE_TRANSACTION": 1,
            "DATA_VALIDATION_ERROR": 2,
        }
        assert ws_po_raw is not None
        for r_idx in range(2, ws_rec.max_row + 1):
            raw_po = str(ws_rec.cell(row=r_idx, column=2).value or "")
            raw_sku = str(ws_rec.cell(row=r_idx, column=4).value or "")
            canon_sku = str(ws_rec.cell(row=r_idx, column=8).value or "")
            match_tier = str(ws_rec.cell(row=r_idx, column=10).value or "")
            status = str(ws_rec.cell(row=r_idx, column=21).value or "")

            bucket: str | None = None
            if status == "RECONCILED" and match_tier == "ALIAS_MAP":
                bucket = "ALIAS_RECONCILED"
            elif status == "RECONCILED" and (raw_sku != canon_sku or raw_po != raw_po.strip()):
                bucket = "DIRTY_RECONCILED"
            elif status in bucket_quotas:
                bucket = status

            if bucket and bucket_quotas.get(bucket, 0) > 0:
                selected_rows.append(r_idx)
                bucket_quotas[bucket] -= 1
            if sum(bucket_quotas.values()) == 0:
                break

        for out_row_offset, rec_row_idx in enumerate(selected_rows, start=3):
            ws_cmp.row_dimensions[out_row_offset].height = 23.0
            src_row = int(ws_rec.cell(row=rec_row_idx, column=1).value or rec_row_idx)
            raw_po_val = str(ws_po_raw.cell(row=src_row, column=1).value or "")
            raw_sku_val = str(ws_po_raw.cell(row=src_row, column=5).value or "")
            raw_qty_val = ws_po_raw.cell(row=src_row, column=7).value
            raw_prc_val = ws_po_raw.cell(row=src_row, column=8).value
            status_val = str(ws_rec.cell(row=rec_row_idx, column=21).value or "")

            ord_v = ws_rec.cell(row=rec_row_idx, column=13).value
            shp_v = ws_rec.cell(row=rec_row_idx, column=14).value
            inv_v = ws_rec.cell(row=rec_row_idx, column=15).value
            q_triplet = (
                f"{ord_v if ord_v is not None else '-'}/"
                f"{shp_v if shp_v is not None else '-'}/"
                f"{inv_v if inv_v is not None else '-'}"
            )

            row_vals = [
                src_row,
                f'"{raw_po_val}"' if raw_po_val != raw_po_val.strip() else raw_po_val,
                f'"{raw_sku_val}"' if raw_sku_val != raw_sku_val.strip() else raw_sku_val,
                f"{raw_qty_val} @ ${raw_prc_val}",
                "",
                ws_rec.cell(row=rec_row_idx, column=6).value,
                ws_rec.cell(row=rec_row_idx, column=8).value or "(unmapped)",
                ws_rec.cell(row=rec_row_idx, column=10).value,
                q_triplet,
                status_val,
            ]
            for col_idx, cell_val in enumerate(row_vals, start=1):
                if col_idx == 5:
                    continue
                c = write_safe_cell(ws_cmp, out_row_offset, col_idx, cell_val)
                c.font = mono_font if col_idx in (2, 3, 6, 7, 9) else body_font
                c.border = border
                if out_row_offset % 2 == 1:
                    c.fill = zebra_fill
                if col_idx == 1:
                    c.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx in (4, 8, 9):
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 10 and status_val in status_fills:
                    s_fill, s_font = status_fills[status_val]
                    c.fill = s_fill
                    c.font = s_font
                    c.alignment = Alignment(horizontal="center", vertical="center")

        cmp_widths = [5.0, 15.5, 15.0, 13.0, 1.8, 15.5, 13.0, 10.5, 11.5, 20.5]
        for idx_cw, cw_val in enumerate(cmp_widths, start=1):
            ws_cmp.column_dimensions[get_column_letter(idx_cw)].width = cw_val

        before_after_xlsx = tdir / "before_after_view.xlsx"
        wb_cmp.save(before_after_xlsx)
        if not _render_xlsx_via_quicklook(before_after_xlsx, p_before):
            raise RuntimeError("macOS QuickLook failed to render 01_before_after.png")

        wb_po_raw.close()
        wb_rep.close()

    s = result.summary
    metrics_map = {
        "total_po_rows": f"{s.total_po_rows:,}",
        "unique_valid_po_lines": f"{s.unique_valid_po_lines:,}",
        "matching_coverage": f"{float(s.matching_coverage_rate) * 100:.2f}%",
        "reconciled_lines": f"{s.reconciled_lines:,}",
        "reconciliation_rate": f"{float(s.reconciliation_rate) * 100:.2f}%",
        "partial_lines": f"{s.partial_fulfillment_lines:,}",
        "total_exceptions": f"{s.total_exception_records:,}",
        "expected_amount": f"${s.total_expected_amount:,.2f}",
        "invoiced_amount": f"${s.total_invoiced_amount:,.2f}",
        "net_variance": f"${s.net_billed_variance:,.2f}",
        "abs_discrepancy": f"${s.absolute_monetary_discrepancy:,.2f}",
        "unbilled_amount": f"${s.unbilled_shipped_amount:,.2f}",
        "quarantined_invoice_amount": f"${s.quarantined_invoice_amount:,.2f}",
    }
    pdf_out = build_case_study_pdf(
        PORTFOLIO_DIR / "case_study.pdf",
        summary_metrics=metrics_map,
        screenshot_paths={
            "workflow": wf_path,
            "before_after": p_before,
            "dashboard": p_dash,
            "exceptions": p_exc,
        },
    )
    print(f"Generated 3-page PDF case study: {pdf_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
