"""Professional Excel report generation with openpyxl styling, formulas, and audit trails."""

from __future__ import annotations

import os
import re
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from excel_reconciliation.models import (
    ExceptionCategory,
    ReconciliationResult,
    ReconciliationStatus,
    Severity,
)
from excel_reconciliation.validation import WorkbookFormulaCache, evaluate_workbook_formula

# Color palette (clean corporate finance theme)
_NAVY_HEADER_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
_SECTION_BANNER_FILL = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
_SUBHEADER_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
_ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
_TOTAL_ROW_FILL = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")

_GREEN_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
_GREEN_FONT = Font(name="Calibri", size=10, bold=True, color="166534")

_RED_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
_RED_FONT = Font(name="Calibri", size=10, bold=True, color="991B1B")

_AMBER_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
_AMBER_FONT = Font(name="Calibri", size=10, bold=True, color="92400E")

_BLUE_FILL = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
_BLUE_FONT = Font(name="Calibri", size=10, bold=True, color="1E40AF")

_PURPLE_FILL = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
_PURPLE_FONT = Font(name="Calibri", size=10, bold=True, color="6B21A8")

_TITLE_FONT = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
_SECTION_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
_BOLD_BODY_FONT = Font(name="Calibri", size=10, bold=True, color="0F172A")
_BODY_FONT = Font(name="Calibri", size=10, bold=False, color="1E293B")
_MONO_FONT = Font(name="Consolas", size=9, bold=False, color="334155")

_THIN_GRAY = Side(border_style="thin", color="CBD5E1")
_MED_NAVY = Side(border_style="medium", color="1B365D")
_CELL_BORDER = Border(left=_THIN_GRAY, right=_THIN_GRAY, top=_THIN_GRAY, bottom=_THIN_GRAY)
_TOTAL_BORDER = Border(top=_THIN_GRAY, bottom=_MED_NAVY, left=_THIN_GRAY, right=_THIN_GRAY)

_FMT_INT = "#,##0"
_FMT_QTY = "#,##0.###"
_FMT_CURRENCY = '$#,##0.00;($#,##0.00);"$0.00"'
_FMT_PERCENT = "0.00%"

_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_DANGEROUS_TEXT_PREFIXES: tuple[str, ...] = ("=", "+", "-", "@", "\t", "\r", "\n")

_STATUS_STYLE_MAP: dict[str, tuple[PatternFill, Font]] = {
    ReconciliationStatus.RECONCILED.value: (_GREEN_FILL, _GREEN_FONT),
    ReconciliationStatus.PARTIAL_FULFILLMENT.value: (_BLUE_FILL, _BLUE_FONT),
    ReconciliationStatus.QUANTITY_MISMATCH.value: (_RED_FILL, _RED_FONT),
    ReconciliationStatus.PRICE_MISMATCH.value: (_RED_FILL, _RED_FONT),
    ReconciliationStatus.MISSING_SHIPMENT.value: (_AMBER_FILL, _AMBER_FONT),
    ReconciliationStatus.MISSING_INVOICE.value: (_AMBER_FILL, _AMBER_FONT),
    ReconciliationStatus.AMBIGUOUS_REFERENCE.value: (_PURPLE_FILL, _PURPLE_FONT),
    ReconciliationStatus.DUPLICATE_TRANSACTION.value: (_AMBER_FILL, _AMBER_FONT),
    ReconciliationStatus.DATA_VALIDATION_ERROR.value: (_RED_FILL, _RED_FONT),
}

_SEVERITY_STYLE_MAP: dict[str, tuple[PatternFill, Font]] = {
    Severity.HIGH.value: (_RED_FILL, _RED_FONT),
    Severity.MEDIUM.value: (_AMBER_FILL, _AMBER_FONT),
    Severity.LOW.value: (_BLUE_FILL, _BLUE_FONT),
}


def sanitize_excel_text(value: str) -> tuple[str, bool]:
    """Strip invalid XML control chars and flag strings needing literal quotePrefix protection."""
    cleaned = _ILLEGAL_XML_RE.sub("", value)
    needs_quote_prefix = cleaned.startswith(_DANGEROUS_TEXT_PREFIXES) or cleaned.lstrip(
        " "
    ).startswith(_DANGEROUS_TEXT_PREFIXES)
    return cleaned, needs_quote_prefix


def write_safe_cell(
    ws: Worksheet,
    row: int,
    col: int,
    value: object,
    *,
    allow_formula: bool = False,
) -> openpyxl.cell.Cell:
    """Write a cell value safely without formula injection or unintended type coercion."""
    cell = ws.cell(row=row, column=col)
    if value is None:
        cell.value = None
        return cell
    if isinstance(value, bool):
        cell.value = value
        return cell
    if isinstance(value, str):
        if allow_formula and value.startswith("="):
            cell.value = value
            return cell
        cleaned, needs_quote = sanitize_excel_text(value)
        cell.value = cleaned
        cell.data_type = "s"
        if needs_quote:
            cell.quotePrefix = True
        return cell
    cell.value = value
    return cell


def _format_cache_decimal(dec: Decimal) -> str:
    """Render Decimal as a plain non-scientific OpenXML <v> string."""
    if dec == Decimal("0"):
        return "0"
    text = format(dec.normalize(), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def _display_width(text: str) -> int:
    """Estimate visual character width accounting for East Asian wide glyphs."""
    width = 0
    for ch in text:
        width += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return width


def _autofit_columns(
    ws: Worksheet,
    *,
    min_width: int = 12,
    max_width: int = 54,
    sample_limit: int = 300,
) -> None:
    """Set readable Excel column widths using header and sampled body content."""
    max_col = ws.max_column or 1
    max_row = ws.max_row or 1
    inspect_rows = min(max_row, sample_limit)

    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        best = min_width
        for row_idx in range(1, inspect_rows + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = cell.value
            if val is None:
                continue
            if cell.data_type == "f" and isinstance(val, str) and val.startswith("="):
                best = max(best, 16)
                continue
            text = str(val)
            best = max(best, min(_display_width(text) + 4, max_width))
        ws.column_dimensions[col_letter].width = best


def _style_table_header(ws: Worksheet, row_idx: int, col_count: int) -> None:
    """Apply corporate navy styling to a worksheet table header row."""
    ws.row_dimensions[row_idx].height = 24
    for col_idx in range(1, col_count + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.fill = _NAVY_HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER


def _build_summary_sheet(
    ws: Worksheet,
    result: ReconciliationResult,
    rec_last_row: int,
    exc_last_row: int,
) -> None:
    """Populate the Summary dashboard worksheet with pure cross-sheet Excel formulas."""
    s = result.summary
    cfg = result.audit_log.config
    ws.views.sheetView[0].showGridLines = True

    # Banner
    ws.merge_cells("A1:D1")
    title_cell = write_safe_cell(ws, 1, 1, "Excel Reconciliation Automation — Executive Summary")
    title_cell.fill = _NAVY_HEADER_FILL
    title_cell.font = _TITLE_FONT
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:D2")
    sub_cell = write_safe_cell(
        ws,
        2,
        1,
        (
            f"Generated UTC: {result.audit_log.generated_at_utc}  |  "
            f"Currency: {cfg.reporting_currency}  |  "
            f"Rounding: {cfg.rounding_mode} ({cfg.currency_quantum})  |  "
            f"Qty Tol: {cfg.quantity_tolerance}  |  "
            f"Price Tol: ${cfg.unit_price_tolerance}"
        ),
    )
    sub_cell.fill = _SUBHEADER_FILL
    sub_cell.font = _BODY_FONT
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 20

    # Section 1: Portfolio & Financial KPIs
    write_safe_cell(ws, 4, 1, "Key Reconciliation Metric")
    write_safe_cell(ws, 4, 2, "Computed Value")
    write_safe_cell(ws, 4, 3, "Live Excel Formula Check")
    write_safe_cell(ws, 4, 4, "Operational Definition")
    _style_table_header(ws, 4, 4)
    ws.freeze_panes = "A5"

    rec_has_rows = len(result.reconciled_records) > 0
    exc_has_rows = len(result.exceptions) > 0

    rec_id_range = f"'Reconciled Records'!$A$2:$A${rec_last_row}"
    rec_prod_tier_range = f"'Reconciled Records'!$J$2:$J${rec_last_row}"
    rec_ord_range = f"'Reconciled Records'!$M$2:$M${rec_last_row}"
    rec_shp_range = f"'Reconciled Records'!$N$2:$N${rec_last_row}"
    rec_inv_qty_range = f"'Reconciled Records'!$O$2:$O${rec_last_row}"
    rec_out_range = f"'Reconciled Records'!$P$2:$P${rec_last_row}"
    rec_exp_amt_range = f"'Reconciled Records'!$R$2:$R${rec_last_row}"
    rec_inv_amt_range = f"'Reconciled Records'!$S$2:$S${rec_last_row}"
    rec_var_range = f"'Reconciled Records'!$T$2:$T${rec_last_row}"
    rec_status_range = f"'Reconciled Records'!$U$2:$U${rec_last_row}"
    rec_valid_range = f"'Reconciled Records'!$AB$2:$AB${rec_last_row}"
    rec_disc_range = f"'Reconciled Records'!$AC$2:$AC${rec_last_row}"
    rec_accrual_range = f"'Reconciled Records'!$AD$2:$AD${rec_last_row}"

    cov_formula = (
        f"=IF(B6=0,0,ROUND((COUNTIFS({rec_valid_range},TRUE,"
        f'{rec_prod_tier_range},"EXACT_ID")+'
        f"COUNTIFS({rec_valid_range},TRUE,"
        f'{rec_prod_tier_range},"ALIAS_MAP"))/B6,4))'
        if rec_has_rows
        else "=0"
    )
    quarantined_inv_formula = (
        f"=ROUND(SUMIFS(Exceptions!$M$2:$M${exc_last_row},"
        f'Exceptions!$D$2:$D${exc_last_row},"supplier_invoices",'
        f'Exceptions!$B$2:$B${exc_last_row},"<>PRICE_MISMATCH"),2)'
        if exc_has_rows
        else "=0"
    )

    kpi_rows: list[tuple[str, object, str, str, str]] = [
        (
            "Total Purchase Order Rows",
            s.total_po_rows,
            f"=COUNTA({rec_id_range})" if rec_has_rows else "=0",
            _FMT_INT,
            "All rows ingested from purchase_orders.xlsx",
        ),
        (
            "Unique Valid Purchase Order Lines",
            s.unique_valid_po_lines,
            f"=COUNTIF({rec_valid_range},TRUE)" if rec_has_rows else "=0",
            _FMT_INT,
            "Deduplicated, schema-valid PO lines evaluated in 3-way match",
        ),
        (
            "Successfully Reconciled PO Lines",
            s.reconciled_lines,
            f'=COUNTIF({rec_status_range},"RECONCILED")' if rec_has_rows else "=0",
            _FMT_INT,
            "Exact 3-way match across PO, Shipment(s), and Invoice(s) within tolerance",
        ),
        (
            "Partial Fulfillment PO Lines",
            s.partial_fulfillment_lines,
            f'=COUNTIF({rec_status_range},"PARTIAL_FULFILLMENT")' if rec_has_rows else "=0",
            _FMT_INT,
            "Partially shipped and accurately billed for shipped quantity",
        ),
        (
            "PO Lines Requiring Review",
            s.po_exception_lines,
            f'=COUNTIF({rec_status_range},"<>RECONCILED")' if rec_has_rows else "=0",
            _FMT_INT,
            "PO rows with quantity/price variances, missing docs, duplicates, or errors",
        ),
        (
            "Total Exception Records (All Sources)",
            s.total_exception_records,
            f"=COUNTA(Exceptions!$A$2:$A${exc_last_row})" if exc_has_rows else "=0",
            _FMT_INT,
            "All actionable exceptions across Products, POs, Shipments, and Invoices",
        ),
        (
            "Product Matching Coverage",
            s.matching_coverage_rate,
            cov_formula,
            _FMT_PERCENT,
            "Share of unique valid PO lines resolved to canonical Product Master SKUs",
        ),
        (
            "Three-Way Auto-Reconciliation Rate",
            s.reconciliation_rate,
            "=IF(B6=0,0,ROUND(B7/B6,4))",
            _FMT_PERCENT,
            "Share of unique valid PO lines fully reconciled without human intervention",
        ),
        (
            "Total Ordered Quantity (Valid Unique POs)",
            s.total_ordered_qty,
            f"=SUMIF({rec_valid_range},TRUE,{rec_ord_range})" if rec_has_rows else "=0",
            _FMT_QTY,
            "Total units ordered across valid, non-duplicate purchase order lines",
        ),
        (
            "Total Shipped Quantity",
            s.total_shipped_qty,
            f"=SUM({rec_shp_range})" if rec_has_rows else "=0",
            _FMT_QTY,
            "Pre-aggregated units received across matched non-duplicate shipments",
        ),
        (
            "Total Invoiced Quantity",
            s.total_invoiced_qty,
            f"=SUM({rec_inv_qty_range})" if rec_has_rows else "=0",
            _FMT_QTY,
            "Pre-aggregated units billed across matched non-duplicate supplier invoices",
        ),
        (
            "Outstanding Unshipped Quantity",
            s.total_outstanding_qty,
            f"=SUM({rec_out_range})" if rec_has_rows else "=0",
            _FMT_QTY,
            "Open backordered quantity remaining on valid purchase order lines",
        ),
        (
            "Total Expected PO Amount",
            s.total_expected_amount,
            f"=ROUND(SUMIF({rec_valid_range},TRUE,{rec_exp_amt_range}),2)"
            if rec_has_rows
            else "=0",
            _FMT_CURRENCY,
            "Contracted purchase commitment across valid unique PO lines",
        ),
        (
            "Total Actual Invoiced Amount",
            s.total_invoiced_amount,
            f"=ROUND(SUM({rec_inv_amt_range}),2)" if rec_has_rows else "=0",
            _FMT_CURRENCY,
            "Total billed amount across matched non-duplicate supplier invoices",
        ),
        (
            "Net Billed Monetary Variance",
            s.net_billed_variance,
            f"=ROUND(SUM({rec_var_range}),2)" if rec_has_rows else "=0",
            _FMT_CURRENCY,
            "Signed sum of (Invoiced Amount - Authorized Received Value) on invoiced PO lines",
        ),
        (
            "Absolute Monetary Discrepancy Exposure",
            s.absolute_monetary_discrepancy,
            f"=ROUND(SUM({rec_disc_range}),2)" if rec_has_rows else "=0",
            _FMT_CURRENCY,
            "Gross exposure separating unit-price over/undercharges from quantity variances",
        ),
        (
            "Unbilled Shipped Value (Accrual Exposure)",
            s.unbilled_shipped_amount,
            f"=ROUND(SUM({rec_accrual_range}),2)" if rec_has_rows else "=0",
            _FMT_CURRENCY,
            "Contracted value of received goods awaiting supplier invoice",
        ),
        (
            "Quarantined / Unmatched Invoice Amount",
            s.quarantined_invoice_amount,
            quarantined_inv_formula,
            _FMT_CURRENCY,
            f"{result.audit_log.config.reporting_currency} invoice value held in AP quarantine "
            "(unmatched, duplicates, errors; foreign currencies excluded)",
        ),
    ]

    for idx, (label, val, formula_str, num_fmt, note) in enumerate(kpi_rows, start=5):
        ws.row_dimensions[idx].height = 20
        c_label = write_safe_cell(ws, idx, 1, label)
        c_val = write_safe_cell(ws, idx, 2, val)
        c_form = write_safe_cell(ws, idx, 3, formula_str, allow_formula=True)
        c_note = write_safe_cell(ws, idx, 4, note)

        c_label.font = _BOLD_BODY_FONT
        c_val.font = _BOLD_BODY_FONT
        c_form.font = _BODY_FONT
        c_note.font = _BODY_FONT

        c_val.number_format = num_fmt
        c_form.number_format = num_fmt
        c_val.alignment = Alignment(horizontal="right", vertical="center")
        c_form.alignment = Alignment(horizontal="right", vertical="center")
        c_label.alignment = Alignment(horizontal="left", vertical="center")
        c_note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        row_fill = _ZEBRA_FILL if idx % 2 == 1 else None
        for cell in (c_label, c_val, c_form, c_note):
            cell.border = _CELL_BORDER
            if row_fill is not None:
                cell.fill = row_fill

    # Section 2: Exception Distribution Table
    dist_header_row = 5 + len(kpi_rows) + 2
    ws.merge_cells(
        start_row=dist_header_row - 1, start_column=1, end_row=dist_header_row - 1, end_column=4
    )
    sec_cell = write_safe_cell(ws, dist_header_row - 1, 1, "Exception Distribution by Category")
    sec_cell.fill = _SECTION_BANNER_FILL
    sec_cell.font = _SECTION_FONT
    sec_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[dist_header_row - 1].height = 24

    write_safe_cell(ws, dist_header_row, 1, "Exception Category")
    write_safe_cell(ws, dist_header_row, 2, "Exception Count")
    write_safe_cell(ws, dist_header_row, 3, "Excel Formula Check")
    write_safe_cell(ws, dist_header_row, 4, "Share of Total Exceptions")
    _style_table_header(ws, dist_header_row, 4)

    categories = list(ExceptionCategory)
    first_cat_row = dist_header_row + 1
    total_cat_row = first_cat_row + len(categories)

    for offset, cat in enumerate(categories):
        r_idx = first_cat_row + offset
        ws.row_dimensions[r_idx].height = 19
        count_val = s.exception_counts_by_category.get(cat.value, 0)
        count_formula = (
            f"=COUNTIF(Exceptions!$B$2:$B${exc_last_row},A{r_idx})" if exc_has_rows else "=0"
        )
        share_formula = f"=IF($B${total_cat_row}=0,0,B{r_idx}/$B${total_cat_row})"

        c1 = write_safe_cell(ws, r_idx, 1, cat.value)
        c2 = write_safe_cell(ws, r_idx, 2, count_val)
        c3 = write_safe_cell(ws, r_idx, 3, count_formula, allow_formula=True)
        c4 = write_safe_cell(ws, r_idx, 4, share_formula, allow_formula=True)

        c1.font = _MONO_FONT
        c2.font = _BODY_FONT
        c3.font = _BODY_FONT
        c4.font = _BODY_FONT

        c2.number_format = _FMT_INT
        c3.number_format = _FMT_INT
        c4.number_format = _FMT_PERCENT
        c2.alignment = Alignment(horizontal="right", vertical="center")
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c4.alignment = Alignment(horizontal="right", vertical="center")

        for cell in (c1, c2, c3, c4):
            cell.border = _CELL_BORDER
            if offset % 2 == 1:
                cell.fill = _ZEBRA_FILL

    # Total row for Exception Distribution
    ws.row_dimensions[total_cat_row].height = 21
    t1 = write_safe_cell(ws, total_cat_row, 1, "TOTAL EXCEPTIONS")
    t2 = write_safe_cell(ws, total_cat_row, 2, s.total_exception_records)
    t3 = write_safe_cell(
        ws, total_cat_row, 3, f"=SUM(C{first_cat_row}:C{total_cat_row - 1})", allow_formula=True
    )
    t4 = write_safe_cell(
        ws, total_cat_row, 4, f"=SUM(D{first_cat_row}:D{total_cat_row - 1})", allow_formula=True
    )
    t2.number_format = _FMT_INT
    t3.number_format = _FMT_INT
    t4.number_format = _FMT_PERCENT
    t2.alignment = Alignment(horizontal="right", vertical="center")
    t3.alignment = Alignment(horizontal="right", vertical="center")
    t4.alignment = Alignment(horizontal="right", vertical="center")
    for cell in (t1, t2, t3, t4):
        cell.font = _BOLD_BODY_FONT
        cell.fill = _TOTAL_ROW_FILL
        cell.border = _TOTAL_BORDER

    _autofit_columns(ws, min_width=18, max_width=52)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 64


def _build_reconciled_sheet(ws: Worksheet, result: ReconciliationResult) -> int:
    """Populate the Reconciled Records worksheet and return the last data row index."""
    headers = (
        "Source Row",
        "Raw PO Number",
        "Raw PO Line ID",
        "Raw SKU",
        "Normalized PO Number",
        "Normalized PO Line ID",
        "Normalized SKU",
        "Canonical SKU",
        "Matched Business Key",
        "Product Match Tier",
        "Product Description",
        "Supplier ID",
        "Ordered Qty",
        "Shipped Qty",
        "Invoiced Qty",
        "Outstanding Qty",
        "PO Unit Price",
        "Expected Amount",
        "Actual Invoiced Amount",
        "Monetary Variance",
        "Reconciliation Status",
        "Reason Codes",
        "Shipment Count",
        "Invoice Count",
        "Shipment References",
        "Invoice References",
        "Preserved PO Formula",
        "Valid PO Line",
        "Discrepancy Exposure",
        "Unbilled Receipt Accrual",
    )
    for c_idx, header in enumerate(headers, start=1):
        write_safe_cell(ws, 1, c_idx, header)
    _style_table_header(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    for idx, rec in enumerate(result.reconciled_records, start=2):
        reason_str = "; ".join(rc.value for rc in rec.reason_codes)
        row_values: list[object] = [
            rec.source_row,
            rec.raw_po_number,
            rec.raw_po_line_id,
            rec.raw_sku,
            rec.normalized_po_number,
            rec.normalized_po_line_id,
            rec.normalized_sku,
            rec.canonical_sku or "",
            rec.matched_business_key,
            rec.product_match_status.value,
            rec.product_description,
            rec.supplier_id,
            rec.ordered_qty,
            rec.shipped_qty,
            rec.invoiced_qty,
            rec.outstanding_qty,
            rec.po_unit_price,
            rec.expected_amount,
            rec.actual_invoiced_amount,
            rec.monetary_variance,
            rec.reconciliation_status.value,
            reason_str,
            rec.shipment_count,
            rec.invoice_count,
            rec.shipment_refs,
            rec.invoice_refs,
            rec.preserved_po_formula,
            rec.is_valid_po_line,
            rec.discrepancy_exposure if rec.is_valid_po_line else None,
            rec.unbilled_receiving_amount if rec.is_valid_po_line else None,
        ]
        is_zebra = idx % 2 == 1
        for c_idx, val in enumerate(row_values, start=1):
            cell = write_safe_cell(ws, idx, c_idx, val)
            cell.font = _BODY_FONT
            cell.border = _CELL_BORDER
            if is_zebra:
                cell.fill = _ZEBRA_FILL

        ws.cell(row=idx, column=1).number_format = _FMT_INT
        for q_col in (13, 14, 15, 16):
            ws.cell(row=idx, column=q_col).number_format = _FMT_QTY
        for m_col in (17, 18, 19, 20, 29, 30):
            ws.cell(row=idx, column=m_col).number_format = _FMT_CURRENCY
        ws.cell(row=idx, column=23).number_format = _FMT_INT
        ws.cell(row=idx, column=24).number_format = _FMT_INT

        status_cell = ws.cell(row=idx, column=21)
        status_cell.alignment = Alignment(horizontal="center", vertical="center")
        if rec.reconciliation_status.value in _STATUS_STYLE_MAP:
            st_fill, st_font = _STATUS_STYLE_MAP[rec.reconciliation_status.value]
            status_cell.fill = st_fill
            status_cell.font = st_font

        var_cell = ws.cell(row=idx, column=20)
        if (
            rec.monetary_variance is not None
            and abs(rec.monetary_variance) > result.audit_log.config.amount_tolerance
        ):
            var_cell.fill = _RED_FILL
            var_cell.font = _RED_FONT

    last_row = max(len(result.reconciled_records) + 1, 2)
    last_col_letter = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    status_range = f"U2:U{last_row}"
    for status_val, (fill_obj, font_obj) in _STATUS_STYLE_MAP.items():
        ws.conditional_formatting.add(
            status_range,
            CellIsRule(
                operator="equal",
                formula=[f'"{status_val}"'],
                stopIfTrue=True,
                fill=fill_obj,
                font=font_obj,
            ),
        )

    tol_str = str(result.audit_log.config.amount_tolerance)
    var_range = f"T2:T{last_row}"
    ws.conditional_formatting.add(
        var_range,
        CellIsRule(operator="greaterThan", formula=[tol_str], fill=_RED_FILL, font=_RED_FONT),
    )
    ws.conditional_formatting.add(
        var_range,
        CellIsRule(operator="lessThan", formula=[f"-{tol_str}"], fill=_RED_FILL, font=_RED_FONT),
    )

    _autofit_columns(ws, min_width=12, max_width=46)
    return last_row


def _build_exceptions_sheet(ws: Worksheet, result: ReconciliationResult) -> int:
    """Populate the Exceptions review queue worksheet and return last row index."""
    headers = (
        "Exception ID",
        "Exception Category",
        "Severity",
        "Source Dataset",
        "Source Row",
        "PO Reference",
        "Raw Reference",
        "Normalized Reference",
        "Raw SKU",
        "Normalized SKU",
        "Reason Codes",
        "Quantity Impact",
        "Monetary Impact",
        "Explanation",
        "Suggested Next Action",
    )
    for c_idx, header in enumerate(headers, start=1):
        write_safe_cell(ws, 1, c_idx, header)
    _style_table_header(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    for idx, exc in enumerate(result.exceptions, start=2):
        reason_str = "; ".join(rc.value for rc in exc.reason_codes)
        row_values: list[object] = [
            exc.exception_id,
            exc.category.value,
            exc.severity.value,
            exc.source_dataset.value,
            exc.source_row,
            exc.po_reference,
            exc.raw_reference,
            exc.normalized_reference,
            exc.raw_sku,
            exc.normalized_sku,
            reason_str,
            exc.quantity_impact,
            exc.monetary_impact,
            exc.explanation,
            exc.suggested_action,
        ]
        is_zebra = idx % 2 == 1
        for c_idx, val in enumerate(row_values, start=1):
            cell = write_safe_cell(ws, idx, c_idx, val)
            cell.font = _BODY_FONT
            cell.border = _CELL_BORDER
            if is_zebra:
                cell.fill = _ZEBRA_FILL

        ws.cell(row=idx, column=5).number_format = _FMT_INT
        ws.cell(row=idx, column=12).number_format = _FMT_QTY
        ws.cell(row=idx, column=13).number_format = _FMT_CURRENCY

        sev_cell = ws.cell(row=idx, column=3)
        sev_cell.alignment = Alignment(horizontal="center", vertical="center")
        if exc.severity.value in _SEVERITY_STYLE_MAP:
            s_fill, s_font = _SEVERITY_STYLE_MAP[exc.severity.value]
            sev_cell.fill = s_fill
            sev_cell.font = s_font

        cat_cell = ws.cell(row=idx, column=2)
        cat_cell.font = _BOLD_BODY_FONT

    last_row = max(len(result.exceptions) + 1, 2)
    last_col_letter = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    sev_range = f"C2:C{last_row}"
    for sev_val, (fill_obj, font_obj) in _SEVERITY_STYLE_MAP.items():
        ws.conditional_formatting.add(
            sev_range,
            CellIsRule(
                operator="equal",
                formula=[f'"{sev_val}"'],
                stopIfTrue=True,
                fill=fill_obj,
                font=font_obj,
            ),
        )

    _autofit_columns(ws, min_width=13, max_width=58)
    return last_row


def _build_audit_log_sheet(ws: Worksheet, result: ReconciliationResult) -> None:
    """Populate the Audit Log provenance, file hash, rule, and hygiene worksheet."""
    audit = result.audit_log
    headers = ("Section", "Key / Dataset", "Value / Status", "Details")
    for c_idx, header in enumerate(headers, start=1):
        write_safe_cell(ws, 1, c_idx, header)
    _style_table_header(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    rows: list[tuple[str, str, str, str]] = [
        (
            "Execution",
            "Generated Timestamp (UTC)",
            audit.generated_at_utc,
            "ISO-8601 UTC execution time",
        ),
        ("Execution", "Python Interpreter", audit.python_version, "Runtime Python version"),
        (
            "Execution",
            "Polars Engine",
            audit.polars_version,
            "Tabular cleaning & pre-aggregation engine",
        ),
        (
            "Execution",
            "openpyxl Engine",
            audit.openpyxl_version,
            "Workbook I/O & Excel formatting engine",
        ),
        (
            "Tolerance Rules",
            "Reporting Currency",
            audit.config.reporting_currency,
            "Single-currency boundary enforced across PO and Invoice inputs",
        ),
        (
            "Tolerance Rules",
            "Monetary Rounding",
            f"{audit.config.rounding_mode} ({audit.config.currency_quantum})",
            "Exact decimal.Decimal commercial ROUND_HALF_UP arithmetic",
        ),
        (
            "Tolerance Rules",
            "Quantity Tolerance",
            str(audit.config.quantity_tolerance),
            "Permitted absolute unit difference on 3-way match",
        ),
        (
            "Tolerance Rules",
            "Unit Price Tolerance",
            f"${audit.config.unit_price_tolerance}",
            "Permitted unit price delta before PRICE_MISMATCH",
        ),
        (
            "Tolerance Rules",
            "Extended Amount Tolerance",
            f"${audit.config.amount_tolerance}",
            "Permitted line-amount delta across split invoices",
        ),
    ]

    for meta in audit.input_files:
        rows.append(
            (
                "Input File Hash",
                meta.filename,
                meta.sha256,
                (
                    f"sheet={meta.sheet_name}, rows={meta.row_count}, "
                    f"formulas_preserved={meta.formula_cell_count}, "
                    f"broken_formulas={meta.broken_formula_count}"
                ),
            )
        )

    for idx_rule, rule_desc in enumerate(audit.matching_rules_applied, start=1):
        rows.append(("Matching Priority", f"Priority Rule {idx_rule}", "APPLIED", rule_desc))

    for sv in audit.schema_validations:
        status_txt = "VALID" if sv.is_valid else "MISSING_COLUMNS"
        detail_txt = (
            f"expected={len(sv.expected_columns)} cols; missing={list(sv.missing_columns)}; "
            f"extra={list(sv.extra_columns)}"
        )
        rows.append(("Schema Validation", sv.dataset_kind.value, status_txt, detail_txt))

    for cs in audit.cleaning_summaries:
        rows.append(
            (
                "Cleaning Summary",
                cs.dataset_kind.value,
                f"{cs.valid_unique_rows}/{cs.total_rows} valid unique",
                (
                    f"duplicates={cs.duplicate_rows}, "
                    f"conflicting_duplicates={cs.conflicting_duplicate_rows}, "
                    f"whitespace_fixed={cs.whitespace_normalized_count}, "
                    f"sku_normalized={cs.sku_normalized_count}, "
                    f"missing_vals={cs.missing_value_rows}, "
                    f"invalid_nums={cs.invalid_numeric_rows}, "
                    f"invalid_dates={cs.invalid_date_rows}, "
                    f"broken_formulas={cs.broken_formula_rows}"
                ),
            )
        )

    if audit.warnings:
        for idx_w, warn_msg in enumerate(audit.warnings, start=1):
            rows.append(("Processing Warning", f"WARN-{idx_w:02d}", "ATTENTION", warn_msg))
    else:
        rows.append(("Processing Warning", "Warnings", "NONE", "Zero processing warnings emitted"))

    for r_idx, row_tuple in enumerate(rows, start=2):
        is_zebra = r_idx % 2 == 1
        for c_idx, val in enumerate(row_tuple, start=1):
            cell = write_safe_cell(ws, r_idx, c_idx, val)
            cell.font = (
                _MONO_FONT if c_idx == 3 and row_tuple[0] == "Input File Hash" else _BODY_FONT
            )
            cell.border = _CELL_BORDER
            if is_zebra:
                cell.fill = _ZEBRA_FILL
        ws.cell(row=r_idx, column=1).font = _BOLD_BODY_FONT

    last_row = max(len(rows) + 1, 2)
    ws.auto_filter.ref = f"A1:D{last_row}"
    _autofit_columns(ws, min_width=18, max_width=70)


def _configure_landscape_fit(ws: Worksheet) -> None:
    """Configure landscape page setup so spreadsheet viewers render full table width cleanly."""
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def _compute_summary_formula_cache(wb: openpyxl.Workbook) -> dict[str, str]:
    """Evaluate every Summary formula cell from worksheet contents to populate <v> caches."""
    ws_sum = wb["Summary"]
    cache = WorkbookFormulaCache(wb)
    caches: dict[str, str] = {}
    for row in ws_sum.iter_rows():
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                eval_dec = evaluate_workbook_formula(wb, "Summary", cell.value, cache=cache)
                caches[cell.coordinate] = _format_cache_decimal(eval_dec)
    return caches


def _patch_sheet1_formula_cache_atomic(
    raw_xlsx_path: Path,
    dest_xlsx_path: Path,
    formula_cache: dict[str, str],
) -> None:
    """Write ECMA-376 <v> caches computed from <f> onto Summary (sheet1) atomically."""
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", ns_main)
    ET.register_namespace(
        "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    f_tag = f"{{{ns_main}}}f"
    v_tag = f"{{{ns_main}}}v"
    c_tag = f"{{{ns_main}}}c"

    fd, patched_tmp_name = tempfile.mkstemp(
        dir=dest_xlsx_path.parent,
        prefix=f".{dest_xlsx_path.stem}.",
        suffix=".patched.xlsx",
    )
    os.close(fd)
    patched_tmp = Path(patched_tmp_name)
    try:
        with (
            zipfile.ZipFile(raw_xlsx_path, "r") as zin,
            zipfile.ZipFile(patched_tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout,
        ):
            for item in zin.infolist():
                raw_bytes = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml" and formula_cache:
                    root = ET.fromstring(raw_bytes)
                    for cell_elem in root.iter(c_tag):
                        coord = cell_elem.attrib.get("r", "")
                        if coord in formula_cache and cell_elem.find(f_tag) is not None:
                            cell_elem.attrib.pop("t", None)
                            v_elem = cell_elem.find(v_tag)
                            if v_elem is None:
                                v_elem = ET.SubElement(cell_elem, v_tag)
                            v_elem.text = formula_cache[coord]
                    raw_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                zout.writestr(item, raw_bytes)
        os.replace(patched_tmp, dest_xlsx_path)
    finally:
        if patched_tmp.exists():
            patched_tmp.unlink()


def write_reconciliation_report(result: ReconciliationResult, output_path: Path) -> Path:
    """Write the complete four-tab reconciliation_report.xlsx workbook atomically."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()

    ws_summary = wb.active
    assert ws_summary is not None
    ws_summary.title = "Summary"
    ws_reconciled = wb.create_sheet(title="Reconciled Records")
    ws_exceptions = wb.create_sheet(title="Exceptions")
    ws_audit = wb.create_sheet(title="Audit Log")

    rec_last_row = _build_reconciled_sheet(ws_reconciled, result)
    exc_last_row = _build_exceptions_sheet(ws_exceptions, result)
    _build_summary_sheet(ws_summary, result, rec_last_row=rec_last_row, exc_last_row=exc_last_row)
    _build_audit_log_sheet(ws_audit, result)

    for ws in (ws_summary, ws_reconciled, ws_exceptions, ws_audit):
        _configure_landscape_fit(ws)

    formula_cache = _compute_summary_formula_cache(wb)

    fd, raw_tmp_name = tempfile.mkstemp(
        dir=output_path.parent,
        prefix=f".{output_path.stem}.",
        suffix=".raw.xlsx",
    )
    os.close(fd)
    raw_tmp = Path(raw_tmp_name)
    try:
        wb.save(raw_tmp)
        wb.close()
        _patch_sheet1_formula_cache_atomic(raw_tmp, output_path, formula_cache)
    finally:
        if raw_tmp.exists():
            raw_tmp.unlink()

    return output_path
