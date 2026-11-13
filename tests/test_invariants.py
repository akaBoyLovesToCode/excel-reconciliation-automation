"""Independent business-invariant tests across multi-supplier reconciliation scenarios."""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from decimal import Decimal
from pathlib import Path

import openpyxl
import polars as pl

from excel_reconciliation.cleaning import (
    clean_invoices_df,
    clean_products_df,
    clean_purchase_orders_df,
    clean_shipments_df,
)
from excel_reconciliation.models import REQUIRED_COLUMNS, DatasetKind
from excel_reconciliation.reconciliation import reconcile_datasets
from excel_reconciliation.reporting import write_reconciliation_report
from excel_reconciliation.validation import evaluate_workbook_formula, validate_report_workbook


def test_nine_business_invariants_on_multi_supplier_fixture(tmp_path: Path) -> None:
    """Assert 9 structural, cardinality, supplier, key, exposure, accrual, and parity invariants."""
    products_df = pl.DataFrame(
        [
            {
                "product_id": "PRD-A",
                "canonical_sku": "SKU-A",
                "product_name": "Alpha Sensor",
                "category": "Industrial",
                "uom": "EA",
                "standard_cost": "10.00",
                "supplier_id": "SUP-1",
                "aliases": "ALS-A",
                "active_flag": "Y",
            },
            {
                "product_id": "PRD-B",
                "canonical_sku": "SKU-B",
                "product_name": "Beta Valve",
                "category": "Industrial",
                "uom": "EA",
                "standard_cost": "20.00",
                "supplier_id": "SUP-2",
                "aliases": "ALS-B",
                "active_flag": "Y",
            },
        ],
        schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PRODUCTS], pl.String),
    )

    po_df = pl.DataFrame(
        [
            # PO with delimiter collision attempt A::B | C vs A | B::C
            {
                "po_number": "A::B",
                "po_line_id": "C",
                "order_date": "2026-03-01",
                "supplier_id": "SUP-1",
                "sku": "SKU-A",
                "product_description": "Alpha",
                "ordered_qty": "10",
                "unit_price": "10.00",
                "currency": "USD",
                "line_total": "100.00",
            },
            {
                "po_number": "A",
                "po_line_id": "B::C",
                "order_date": "2026-03-01",
                "supplier_id": "SUP-2",
                "sku": "SKU-B",
                "product_description": "Beta",
                "ordered_qty": "8",
                "unit_price": "20.00",
                "currency": "USD",
                "line_total": "160.00",
            },
            # Exact duplicate of PO A|B::C (quarantined)
            {
                "po_number": " a ",
                "po_line_id": " b::c ",
                "order_date": "2026-03-01",
                "supplier_id": "SUP-2",
                "sku": "SKU-B",
                "product_description": "Beta",
                "ordered_qty": "8",
                "unit_price": "20.00",
                "currency": "USD",
                "line_total": "160.00",
            },
            # Partially invoiced PO line (10 ordered, 8 shipped, 5 invoiced -> accrual $30.00)
            {
                "po_number": "PO-ACC",
                "po_line_id": "1",
                "order_date": "2026-03-01",
                "supplier_id": "SUP-1",
                "sku": "SKU-A",
                "product_description": "Alpha",
                "ordered_qty": "10",
                "unit_price": "10.00",
                "currency": "USD",
                "line_total": "100.00",
            },
            # Offset price+qty variance (10 ordered/shipped @ $24, 12 invoiced @ $20 -> net 0, $80)
            {
                "po_number": "PO-OFF",
                "po_line_id": "1",
                "order_date": "2026-03-01",
                "supplier_id": "SUP-2",
                "sku": "SKU-B",
                "product_description": "Beta",
                "ordered_qty": "10",
                "unit_price": "24.00",
                "currency": "USD",
                "line_total": "240.00",
            },
        ],
        schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS], pl.String),
    )

    ship_df = pl.DataFrame(
        [
            # Split shipments on A::B / C (4 + 6 = 10) plus 1 exact repeat duplicate
            {
                "shipment_id": "SHP-1",
                "po_number": "A::B",
                "po_line_id": "C",
                "ship_date": "2026-03-02",
                "carrier": "UPS",
                "tracking_ref": "T1",
                "sku": "SKU-A",
                "shipped_qty": "4",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-1",
                "po_number": "A::B",
                "po_line_id": "C",
                "ship_date": "2026-03-02",
                "carrier": "UPS",
                "tracking_ref": "T1",
                "sku": "als_a",
                "shipped_qty": "4",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-2",
                "po_number": "A::B",
                "po_line_id": "C",
                "ship_date": "2026-03-03",
                "carrier": "UPS",
                "tracking_ref": "T2",
                "sku": "SKU-A",
                "shipped_qty": "6",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-3",
                "po_number": "A",
                "po_line_id": "B::C",
                "ship_date": "2026-03-02",
                "carrier": "DHL",
                "tracking_ref": "T3",
                "sku": "SKU-B",
                "shipped_qty": "8",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-4",
                "po_number": "PO-ACC",
                "po_line_id": "1",
                "ship_date": "2026-03-02",
                "carrier": "DHL",
                "tracking_ref": "T4",
                "sku": "SKU-A",
                "shipped_qty": "8",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-5",
                "po_number": "PO-OFF",
                "po_line_id": "1",
                "ship_date": "2026-03-02",
                "carrier": "DHL",
                "tracking_ref": "T5",
                "sku": "SKU-B",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
        ],
        schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.SHIPMENTS], pl.String),
    )

    inv_df = pl.DataFrame(
        [
            # Two suppliers sharing INV-SHARED-01 line 1 (supplier-scoped identity)
            {
                "invoice_id": "INV-SHARED-01",
                "invoice_line_id": "1",
                "po_number": "A::B",
                "po_line_id": "C",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-1",
                "sku": "SKU-A",
                "invoiced_qty": "5",
                "unit_price": "10.00",
                "line_amount": "50.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-SHARED-01",
                "invoice_line_id": "1",
                "po_number": "A::B",
                "po_line_id": "C",
                "invoice_date": "2026-03-05",
                "supplier_id": "SUP-1",
                "sku": "SKU-A",
                "invoiced_qty": "5",
                "unit_price": "10.00",
                "line_amount": "50.00",
                "currency": "USD",
            },
        ],
        schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.INVOICES], pl.String),
    )
    # Use unique invoice_line_id for second split invoice on SUP-1 and reuse line 1 on SUP-2
    inv_rows = inv_df.to_dicts()
    inv_rows[1]["invoice_line_id"] = "2"
    inv_rows.extend(
        [
            {
                "invoice_id": "INV-SHARED-01",
                "invoice_line_id": "1",
                "po_number": "A",
                "po_line_id": "B::C",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-2",
                "sku": "SKU-B",
                "invoiced_qty": "8",
                "unit_price": "20.00",
                "line_amount": "160.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-ACC",
                "invoice_line_id": "1",
                "po_number": "PO-ACC",
                "po_line_id": "1",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-1",
                "sku": "SKU-A",
                "invoiced_qty": "5",
                "unit_price": "10.00",
                "line_amount": "50.00",
                "currency": "USD",
            },
            # Offset price+qty on PO-OFF: 12 * $20.00 = $240.00 vs PO 10 * $24.00 = $240.00
            # price delta = 10 * (-4) = -$40, qty delta = +2 * $20 = +$40 -> net $0, exposure $80!
            {
                "invoice_id": "INV-OFF",
                "invoice_line_id": "1",
                "po_number": "PO-OFF",
                "po_line_id": "1",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-2",
                "sku": "SKU-B",
                "invoiced_qty": "12",
                "unit_price": "20.00",
                "line_amount": "240.00",
                "currency": "USD",
            },
            # Cross-supplier hijacked invoice claiming SUP-1 on SUP-2's PO-OFF -> quarantined
            {
                "invoice_id": "INV-HIJACK",
                "invoice_line_id": "1",
                "po_number": "PO-OFF",
                "po_line_id": "1",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-1",
                "sku": "SKU-B",
                "invoiced_qty": "2",
                "unit_price": "20.00",
                "line_amount": "40.00",
                "currency": "USD",
            },
        ]
    )
    inv_df = pl.DataFrame(
        inv_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.INVOICES], pl.String)
    )

    result = reconcile_datasets(
        products_clean=clean_products_df(products_df),
        po_clean=clean_purchase_orders_df(po_df),
        shipments_clean=clean_shipments_df(ship_df),
        invoices_clean=clean_invoices_df(inv_df),
    )

    s = result.summary
    # Invariant 1 & 2: PO row conservation & status partition
    assert len(result.reconciled_records) == s.total_po_rows == 5
    assert s.unique_valid_po_lines + s.duplicate_po_rows + s.invalid_po_rows == s.total_po_rows
    assert s.reconciled_lines + s.po_exception_lines == s.total_po_rows

    # Invariant 3 & 4 & 6: Dedup, M x N split, and collision-free keys never inflate totals
    assert s.total_ordered_qty == Decimal("38")  # 10 + 8 + 10 + 10
    assert s.total_shipped_qty == Decimal("36")  # 10 + 8 + 8 + 10
    assert s.total_invoiced_qty == Decimal("35")  # 10 + 8 + 5 + 12

    # Invariant 5: Supplier-scoped INV-SHARED-01 ok on A::B|C & A|B::C; INV-HIJACK quarantined
    assert s.quarantined_invoice_amount == Decimal("40.00")

    # Invariant 7: Discrepancy exposure >= |monetary_variance| on valid lines (PO-OFF=$80)
    for rec in result.reconciled_records:
        if rec.is_valid_po_line and rec.monetary_variance is not None:
            assert rec.discrepancy_exposure >= abs(rec.monetary_variance)
    assert s.absolute_monetary_discrepancy == Decimal("80.00")

    # Invariant 8: Partial invoicing receipt accrual = (8 - 5) * $10 = $30.00
    assert s.unbilled_shipped_amount == Decimal("30.00")

    # Invariant 9: 3-way parity (Col B == Evaluated Formula == OpenXML <v>) in report
    report_path = tmp_path / "inv_report.xlsx"
    write_reconciliation_report(result, report_path)
    val = validate_report_workbook(report_path)
    assert val.is_valid, val.errors

    wb = openpyxl.load_workbook(report_path, data_only=False)
    try:
        ws_sum = wb["Summary"]
        with zipfile.ZipFile(report_path, "r") as zf:
            root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            xml_v = {
                el.attrib["r"]: child.text
                for el in root.iter()
                if (el.tag.endswith("}c") or el.tag == "c") and "r" in el.attrib
                for child in el
                if child.tag.endswith("}v") or child.tag == "v"
            }
        for r_idx in range(5, 23):
            col_b = Decimal(str(ws_sum.cell(row=r_idx, column=2).value))
            eval_c = evaluate_workbook_formula(
                wb, "Summary", str(ws_sum.cell(row=r_idx, column=3).value)
            )
            cached_v = Decimal(str(xml_v[f"C{r_idx}"]))
            assert abs(col_b - eval_c) <= Decimal("0.0001")
            assert abs(cached_v - eval_c) <= Decimal("0.0001")
    finally:
        wb.close()
