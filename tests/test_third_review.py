"""Independent third-review regressions; run with PYTHONPATH=src against either tree."""

from __future__ import annotations

import datetime as dt
import xml.etree.ElementTree as ET
import zipfile
from decimal import Decimal

import openpyxl
import polars as pl
import pytest
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, to_excel

from excel_reconciliation.cleaning import (
    clean_invoices_df,
    clean_products_df,
    clean_purchase_orders_df,
    clean_shipments_df,
)
from excel_reconciliation.excel_io import (
    evaluate_worksheet_formula,
    is_permitted_source_formula,
    read_excel_dataset,
)
from excel_reconciliation.matching import ProductMatcher
from excel_reconciliation.models import (
    REQUIRED_COLUMNS,
    DatasetKind,
    MatchStatus,
    ReconciliationStatus,
)
from excel_reconciliation.reconciliation import reconcile_datasets
from excel_reconciliation.reporting import (
    _compute_summary_formula_cache,
    _patch_sheet1_formula_cache_atomic,
    write_reconciliation_report,
)
from excel_reconciliation.validation import validate_report_workbook

PRODUCT = {
    "product_id": "P1",
    "canonical_sku": "SKU",
    "product_name": "Widget",
    "category": "Part",
    "uom": "EA",
    "standard_cost": "1",
    "supplier_id": "S1",
    "aliases": "LEG",
    "active_flag": "Y",
}
ORDER = {
    "po_number": "PO",
    "po_line_id": "L1",
    "order_date": "2026-01-01",
    "supplier_id": "S1",
    "sku": "SKU",
    "product_description": "Widget",
    "ordered_qty": "10",
    "unit_price": "10",
    "currency": "USD",
    "line_total": "100",
}
SHIPMENT = {
    "shipment_id": "SH",
    "po_number": "PO",
    "po_line_id": "L1",
    "ship_date": "2026-01-02",
    "carrier": "UPS",
    "tracking_ref": "T1",
    "sku": "SKU",
    "shipped_qty": "10",
    "warehouse_code": "W",
}
INVOICE = {
    "invoice_id": "IV",
    "invoice_line_id": "1",
    "po_number": "PO",
    "po_line_id": "L1",
    "invoice_date": "2026-01-03",
    "supplier_id": "S1",
    "sku": "SKU",
    "invoiced_qty": "10",
    "unit_price": "10",
    "line_amount": "100",
    "currency": "USD",
}


def frame(kind, rows):
    return pl.DataFrame(rows, schema=dict.fromkeys(REQUIRED_COLUMNS[kind], pl.String))


def result(*, orders=None, products=None, shipments=None, invoices=None):
    return reconcile_datasets(
        clean_products_df(frame(DatasetKind.PRODUCTS, [PRODUCT] if products is None else products)),
        clean_purchase_orders_df(
            frame(DatasetKind.PURCHASE_ORDERS, [ORDER] if orders is None else orders)
        ),
        clean_shipments_df(
            frame(DatasetKind.SHIPMENTS, [SHIPMENT] if shipments is None else shipments)
        ),
        clean_invoices_df(frame(DatasetKind.INVOICES, [INVOICE] if invoices is None else invoices)),
    )


def recache(wb, path):
    """Save caches afresh so corruption tests cannot fail only because openpyxl cleared them."""
    raw = path.with_suffix(".raw.xlsx")
    cache = _compute_summary_formula_cache(wb)
    wb.save(raw)
    wb.close()
    _patch_sheet1_formula_cache_atomic(raw, path, cache)
    raw.unlink()


@pytest.mark.parametrize(
    "totals",
    [("100", "999"), ("999", "100"), ("100", "TBD"), ("TBD", "100"), ("100", None), (None, "100")],
)
def test_f1_conflicting_po_totals_quarantine_every_version(totals):
    r = result(orders=[dict(ORDER, line_total=t) for t in totals])
    assert [x.reconciliation_status for x in r.reconciled_records] == [
        ReconciliationStatus.DUPLICATE_TRANSACTION
    ] * 2
    assert all(not x.is_valid_po_line for x in r.reconciled_records)
    assert r.summary.total_expected_amount == Decimal("0.00")
    assert r.summary.total_shipped_qty == 0
    assert r.summary.total_invoiced_amount == 0
    assert r.summary.quarantined_invoice_amount == Decimal("100.00")


@pytest.mark.parametrize("aliases", [(None, "LEG"), ("LEG", None), ("LEG;ALT", "ALT;LEG")])
def test_f2_alias_metadata_survives_duplicate_order(aliases):
    cleaned = clean_products_df(
        frame(DatasetKind.PRODUCTS, [dict(PRODUCT, aliases=a) for a in aliases])
    )
    assert cleaned.summary.duplicate_rows == 1
    assert cleaned.summary.conflicting_duplicate_rows == 0
    assert ProductMatcher(cleaned.dataframe).resolve_sku("LEG").status == MatchStatus.ALIAS_MAP
    assert cleaned.dataframe["raw_aliases"].to_list() == list(aliases)


@pytest.mark.parametrize("n", [1000, 2000, 10000])
def test_f4_complex_formula_quarantined_without_uncaught_recursion(tmp_path, n):
    formula = "=" + "+".join(["1"] * n)
    assert is_permitted_source_formula(formula) is False
    assert evaluate_worksheet_formula(formula, {}) in ("#VALUE!", "#NAME?")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["amount"])
    ws.append([formula])
    path = tmp_path / "complex.xlsx"
    wb.save(path)
    wb.close()
    loaded = read_excel_dataset(path, DatasetKind.INVOICES)
    assert (
        loaded.dataframe["_has_unpermitted_formula"][0]
        or loaded.dataframe["_has_broken_formula"][0]
    )


@pytest.mark.parametrize("epoch", [CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904])
def test_f5_date_references_remain_numeric_in_formulas(tmp_path, epoch):
    wb = openpyxl.Workbook()
    wb.epoch = epoch
    ws = wb.active
    ws.append(["date", "sum", "elapsed", "date_plus_one"])
    ws.append([dt.datetime(2026, 1, 1), "=SUM(A2:A2)", "=A3-A2", "=A2+1"])
    ws.append([dt.datetime(2026, 1, 4)])
    path = tmp_path / "dates.xlsx"
    wb.save(path)
    wb.close()
    row = read_excel_dataset(path, DatasetKind.PURCHASE_ORDERS).dataframe.to_dicts()[0]
    serial = Decimal(str(to_excel(dt.datetime(2026, 1, 1), epoch=epoch)))
    assert row["date"] == "2026-01-01"
    assert Decimal(row["sum"]) == serial
    assert Decimal(row["elapsed"]) == 3
    assert Decimal(row["date_plus_one"]) == serial + 1
    assert not row["_has_broken_formula"]


@pytest.mark.parametrize(
    "corruption",
    [
        "constant",
        "self_reference",
        "extra_formula",
        "numeric_text",
        "cache_type",
        "status",
        "known_status",
        "exception_source",
        "source",
        "headers",
        "coordinated_amount",
    ],
)
def test_f3_validator_rejects_contract_corruption(tmp_path, corruption):
    path = tmp_path / "report.xlsx"
    write_reconciliation_report(result(invoices=[]), path)
    assert validate_report_workbook(path).is_valid
    wb = openpyxl.load_workbook(path)
    if corruption == "constant":
        wb["Summary"]["C13"] = "=10"
    elif corruption == "self_reference":
        wb["Summary"]["C13"] = "=B13"
        wb["Summary"]["B13"] = 999
    elif corruption == "extra_formula":
        wb["Summary"]["D100"] = "=1+1"
    elif corruption == "numeric_text":
        wb["Reconciled Records"]["AC2"] = "0.00"
    elif corruption == "status":
        wb["Reconciled Records"]["U2"] = "NOT_A_STATUS"
    elif corruption == "known_status":
        wb["Reconciled Records"]["U2"] = "QUANTITY_MISMATCH"
    elif corruption == "exception_source":
        wb["Exceptions"]["E2"] = 3
    elif corruption == "source":
        wb["Reconciled Records"]["A2"] = "FORGED"
    elif corruption == "headers":
        wb["Reconciled Records"]["R1"] = "Tax Total"
    elif corruption == "coordinated_amount":
        wb["Reconciled Records"]["R2"] = 999
        wb["Summary"]["B17"] = 999
    recache(wb, path)
    if corruption == "cache_type":
        with zipfile.ZipFile(path) as z:
            entries = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
        name = "xl/worksheets/sheet1.xml"
        root = ET.fromstring(entries[name][1])
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        root.find(".//s:c[@r='C13']", ns).set("t", "str")
        entries[name] = (entries[name][0], ET.tostring(root, encoding="utf-8"))
        with zipfile.ZipFile(path, "w") as z:
            for info, data in entries.values():
                z.writestr(info, data)
    validation = validate_report_workbook(path)
    assert not validation.is_valid, (corruption, validation)
