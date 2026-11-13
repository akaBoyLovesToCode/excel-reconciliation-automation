"""Unit and regression tests for three-way reconciliation, cardinality safety, and Decimal math."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import openpyxl
import polars as pl
import pytest

from excel_reconciliation.cleaning import (
    clean_invoices_df,
    clean_products_df,
    clean_purchase_orders_df,
    clean_shipments_df,
)
from excel_reconciliation.models import (
    REQUIRED_COLUMNS,
    DatasetKind,
    ExceptionCategory,
    ReasonCode,
    ReconciliationConfig,
    ReconciliationStatus,
    SchemaValidationError,
    excel_round,
)
from excel_reconciliation.reconciliation import reconcile_datasets


def _run_mini_reconciliation(
    products_df: pl.DataFrame,
    po_rows: list[dict[str, str]],
    ship_rows: list[dict[str, str]],
    inv_rows: list[dict[str, str]],
    *,
    config: ReconciliationConfig | None = None,
):
    p_clean = clean_products_df(products_df)
    po_clean = clean_purchase_orders_df(
        pl.DataFrame(
            po_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS], pl.String)
        )
    )
    sh_clean = clean_shipments_df(
        pl.DataFrame(
            ship_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.SHIPMENTS], pl.String)
        )
    )
    iv_clean = clean_invoices_df(
        pl.DataFrame(
            inv_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.INVOICES], pl.String)
        )
    )
    return reconcile_datasets(p_clean, po_clean, sh_clean, iv_clean, config=config)


class TestReconciliationRegressionAndFinancialPrecision:
    """Verify M x N join-cardinality safety, duplicate isolation, zero handling, and tolerances."""

    def test_mxn_split_shipments_and_split_invoices_never_inflate_quantities_or_amounts(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        """Regression test: 3 shipments x 2 invoices on one PO line must NOT fan out to 6 rows."""
        po_rows = [
            {
                "po_number": "PO-500",
                "po_line_id": "PO-500-L01",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "USB-C Dock",
                "ordered_qty": "100",
                "unit_price": "19.99",
                "currency": "USD",
                "line_total": "1999.00",
            }
        ]
        ship_rows = [
            {
                "shipment_id": "SHP-A",
                "po_number": "PO-500",
                "po_line_id": "PO-500-L01",
                "ship_date": "2026-02-03",
                "carrier": "FedEx",
                "tracking_ref": "T-A",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "30",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-B",
                "po_number": "PO-500",
                "po_line_id": "PO-500-L01",
                "ship_date": "2026-02-05",
                "carrier": "FedEx",
                "tracking_ref": "T-B",
                "sku": "LEG-001",
                "shipped_qty": "45",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-C",
                "po_number": "PO-500",
                "po_line_id": "PO-500-L01",
                "ship_date": "2026-02-07",
                "carrier": "FedEx",
                "tracking_ref": "T-C",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "25",
                "warehouse_code": "WH-1",
            },
        ]
        inv_rows = [
            {
                "invoice_id": "INV-1",
                "invoice_line_id": "INV-1-01",
                "po_number": "PO-500",
                "po_line_id": "PO-500-L01",
                "invoice_date": "2026-02-08",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "60",
                "unit_price": "19.99",
                "line_amount": "1199.40",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-2",
                "invoice_line_id": "INV-2-01",
                "po_number": "PO-500",
                "po_line_id": "PO-500-L01",
                "invoice_date": "2026-02-10",
                "supplier_id": "SUP-01",
                "sku": "SUP01-DOCK",
                "invoiced_qty": "40",
                "unit_price": "19.99",
                "line_amount": "799.60",
                "currency": "USD",
            },
        ]

        res = _run_mini_reconciliation(sample_products_df, po_rows, ship_rows, inv_rows)
        assert len(res.reconciled_records) == 1
        rec = res.reconciled_records[0]

        assert rec.reconciliation_status == ReconciliationStatus.RECONCILED
        assert rec.shipment_count == 3
        assert rec.invoice_count == 2
        assert rec.ordered_qty == Decimal("100")
        assert rec.shipped_qty == Decimal("100")
        assert rec.invoiced_qty == Decimal("100")
        assert rec.expected_amount == Decimal("1999.00")
        assert rec.actual_invoiced_amount == Decimal("1999.00")
        assert rec.monetary_variance == Decimal("0.00")
        assert res.summary.total_shipped_qty == Decimal("100")
        assert res.summary.total_invoiced_amount == Decimal("1999.00")
        assert len(res.exceptions) == 0

    def test_duplicate_invoices_and_shipments_do_not_inflate_reconciled_amounts(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        """Regression test: duplicate export rows must be quarantined before aggregation."""
        po_rows = [
            {
                "po_number": "PO-600",
                "po_line_id": "PO-600-L01",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "USB-C Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            }
        ]
        ship_rows = [
            {
                "shipment_id": "SHP-600",
                "po_number": "PO-600",
                "po_line_id": "PO-600-L01",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "T600",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
            # Exact duplicate with whitespace/casing noise
            {
                "shipment_id": "  shp-600 ",
                "po_number": "po-600",
                "po_line_id": "po-600-l01",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "T600",
                "sku": "sku_elec_001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
        ]
        inv_rows = [
            {
                "invoice_id": "INV-600",
                "invoice_line_id": "INV-600-01",
                "po_number": "PO-600",
                "po_line_id": "PO-600-L01",
                "invoice_date": "2026-02-05",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
            # Duplicate billing line
            {
                "invoice_id": " inv-600 ",
                "invoice_line_id": " inv-600-01 ",
                "po_number": "PO-600",
                "po_line_id": "PO-600-L01",
                "invoice_date": "2026-02-05",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
        ]

        res = _run_mini_reconciliation(sample_products_df, po_rows, ship_rows, inv_rows)
        rec = res.reconciled_records[0]
        assert rec.reconciliation_status == ReconciliationStatus.RECONCILED
        assert rec.shipped_qty == Decimal("10")
        assert rec.invoiced_qty == Decimal("10")
        assert rec.actual_invoiced_amount == Decimal("500.00")

        dup_excs = [
            e for e in res.exceptions if e.category == ExceptionCategory.DUPLICATE_TRANSACTION
        ]
        assert len(dup_excs) == 2

    def test_valid_zero_quantity_distinguished_from_missing_records(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        po_rows = [
            # Line 1: Confirmed 0-qty cancelled line with 0-qty shipment notice and $0 invoice
            {
                "po_number": "PO-700",
                "po_line_id": "PO-700-L01",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock Cancelled Line",
                "ordered_qty": "0",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "0.00",
            },
            # Line 2: Open line with NO shipment and NO invoice (missing = None, not 0)
            {
                "po_number": "PO-701",
                "po_line_id": "PO-701-L01",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock Unshipped",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
        ]
        ship_rows = [
            {
                "shipment_id": "SHP-700",
                "po_number": "PO-700",
                "po_line_id": "PO-700-L01",
                "ship_date": "2026-02-02",
                "carrier": "UPS",
                "tracking_ref": "T700",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "0",
                "warehouse_code": "WH-1",
            }
        ]
        inv_rows = [
            {
                "invoice_id": "INV-700",
                "invoice_line_id": "INV-700-01",
                "po_number": "PO-700",
                "po_line_id": "PO-700-L01",
                "invoice_date": "2026-02-03",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "0",
                "unit_price": "50.00",
                "line_amount": "0.00",
                "currency": "USD",
            }
        ]

        res = _run_mini_reconciliation(sample_products_df, po_rows, ship_rows, inv_rows)
        r_zero, r_missing = res.reconciled_records

        assert r_zero.reconciliation_status == ReconciliationStatus.RECONCILED
        assert ReasonCode.ZERO_VALUE_CONFIRMED in r_zero.reason_codes
        assert r_zero.shipped_qty == Decimal("0")
        assert r_zero.actual_invoiced_amount == Decimal("0.00")

        assert r_missing.reconciliation_status == ReconciliationStatus.MISSING_SHIPMENT
        assert r_missing.shipped_qty is None
        assert r_missing.actual_invoiced_amount is None

    def test_price_mismatch_partial_fulfillment_and_orphan_records(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        po_rows = [
            # Line 1: Price mismatch ($55.00 billed vs $50.00 PO)
            {
                "po_number": "PO-801",
                "po_line_id": "PO-801-L01",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            # Line 2: Clean partial fulfillment (6 of 10 shipped & billed at $50.00)
            {
                "po_number": "PO-802",
                "po_line_id": "PO-802-L01",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
        ]
        ship_rows = [
            {
                "shipment_id": "SHP-801",
                "po_number": "PO-801",
                "po_line_id": "PO-801-L01",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "T801",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-802",
                "po_number": "PO-802",
                "po_line_id": "PO-802-L01",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "T802",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "6",
                "warehouse_code": "WH-1",
            },
            # Orphan shipment with unknown PO
            {
                "shipment_id": "SHP-ORPHAN",
                "po_number": "PO-99999",
                "po_line_id": "",
                "ship_date": "2026-02-04",
                "carrier": "UPS",
                "tracking_ref": "T999",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "5",
                "warehouse_code": "WH-1",
            },
        ]
        inv_rows = [
            {
                "invoice_id": "INV-801",
                "invoice_line_id": "INV-801-01",
                "po_number": "PO-801",
                "po_line_id": "PO-801-L01",
                "invoice_date": "2026-02-05",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "55.00",
                "line_amount": "550.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-802",
                "invoice_line_id": "INV-802-01",
                "po_number": "PO-802",
                "po_line_id": "PO-802-L01",
                "invoice_date": "2026-02-05",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "6",
                "unit_price": "50.00",
                "line_amount": "300.00",
                "currency": "USD",
            },
        ]

        res = _run_mini_reconciliation(sample_products_df, po_rows, ship_rows, inv_rows)
        r_price, r_partial = res.reconciled_records

        assert r_price.reconciliation_status == ReconciliationStatus.PRICE_MISMATCH
        assert ReasonCode.UNIT_PRICE_DISCREPANCY in r_price.reason_codes
        assert r_price.monetary_variance == Decimal("50.00")

        assert r_partial.reconciliation_status == ReconciliationStatus.PARTIAL_FULFILLMENT
        assert r_partial.outstanding_qty == Decimal("4")
        assert r_partial.monetary_variance == Decimal("0.00")

        # Net billed variance on Summary equals exact sum of line-level monetary variances
        line_var_sum = sum(
            (
                r.monetary_variance
                for r in res.reconciled_records
                if r.monetary_variance is not None
            ),
            Decimal("0.00"),
        )
        assert res.summary.net_billed_variance == line_var_sum == Decimal("50.00")

        categories = {e.category for e in res.exceptions}
        assert ExceptionCategory.UNEXPECTED_EXTRA_RECORD in categories
        assert ExceptionCategory.PRICE_MISMATCH in categories
        assert ExceptionCategory.PARTIAL_FULFILLMENT in categories

    def test_shared_shipment_id_alias_duplicates_and_conflicting_quarantine(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        po_rows = [
            {
                "po_number": "PO-901",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            {
                "po_number": "PO-901",
                "po_line_id": "2",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-02",
                "sku": "SKU-HOME-002",
                "product_description": "Grinder",
                "ordered_qty": "5",
                "unit_price": "120.00",
                "currency": "USD",
                "line_total": "600.00",
            },
            # PO-902 exercises canonical vs alias repeat dedup without inflation
            {
                "po_number": "PO-902",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            # PO-903 exercises conflicting duplicate shipment quarantine (all rows held)
            {
                "po_number": "PO-903",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
        ]
        ship_rows = [
            # Shared shipment_id across distinct PO lines on PO-901 (preserved!)
            {
                "shipment_id": "SHP-SHARED",
                "po_number": "PO-901",
                "po_line_id": "1",
                "ship_date": "2026-02-03",
                "carrier": "DHL",
                "tracking_ref": "TRK-S1",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-SHARED",
                "po_number": "PO-901",
                "po_line_id": "2",
                "ship_date": "2026-02-03",
                "carrier": "DHL",
                "tracking_ref": "TRK-S1",
                "sku": "SKU-HOME-002",
                "shipped_qty": "5",
                "warehouse_code": "WH-1",
            },
            # Exact semantic repeat using product alias LEG-001 for PO-902
            {
                "shipment_id": "SHP-DUP-ALIAS",
                "po_number": "PO-902",
                "po_line_id": "1",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "TRK-A1",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-DUP-ALIAS",
                "po_number": "PO-902",
                "po_line_id": "1",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "TRK-A1",
                "sku": "leg_001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
            # Conflicting duplicate shipments for PO-903 (qty 10 vs 12)
            {
                "shipment_id": "SHP-CONF",
                "po_number": "PO-903",
                "po_line_id": "1",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "TRK-C1",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "SHP-CONF",
                "po_number": "PO-903",
                "po_line_id": "1",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": "TRK-C1",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "12",
                "warehouse_code": "WH-1",
            },
        ]
        inv_rows = [
            {
                "invoice_id": "INV-901-1",
                "invoice_line_id": "1",
                "po_number": "PO-901",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-901-2",
                "invoice_line_id": "1",
                "po_number": "PO-901",
                "po_line_id": "2",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-02",
                "sku": "SKU-HOME-002",
                "invoiced_qty": "5",
                "unit_price": "120.00",
                "line_amount": "600.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-902",
                "invoice_line_id": "1",
                "po_number": "PO-902",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-903",
                "invoice_line_id": "1",
                "po_number": "PO-903",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
        ]

        res = _run_mini_reconciliation(sample_products_df, po_rows, ship_rows, inv_rows)
        rec_901_1, rec_901_2, rec_902, rec_903 = res.reconciled_records

        assert rec_901_1.reconciliation_status == ReconciliationStatus.RECONCILED
        assert rec_901_2.reconciliation_status == ReconciliationStatus.RECONCILED
        assert rec_902.reconciliation_status == ReconciliationStatus.RECONCILED
        assert rec_902.shipped_qty == Decimal("10")
        assert rec_903.reconciliation_status == ReconciliationStatus.AMBIGUOUS_REFERENCE
        assert rec_903.shipped_qty is None
        assert rec_903.shipment_count == 0

        conflict_excs = [
            e for e in res.exceptions if ReasonCode.CONFLICTING_DUPLICATE_SHIPMENT in e.reason_codes
        ]
        assert len(conflict_excs) == 2

    def test_financial_exposure_scenarios_accruals_rounding_and_provenance(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        # 3.3 Commercial ROUND_HALF_UP tie-breaking verification
        assert excel_round(Decimal("1.005"), 2) == Decimal("1.01")
        assert excel_round(Decimal("2.675"), 2) == Decimal("2.68")

        po_rows = [
            # 1: Quantity discrepancy only (inv 12 vs ship/po 10 @ $50) -> variance +100, exp 100
            {
                "po_number": "PO-F1",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            # 2: Price discrepancy only (inv 10 @ $55 vs po $50) -> variance +50, exp 50
            {
                "po_number": "PO-F2",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            # 3: Simultaneous same-sign qty + price (12 @ $55 vs 10 @ $50) -> no double count (160)
            {
                "po_number": "PO-F3",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            # 4: Opposite-sign offsetting qty & price (PO 10 @ $60 = $600; INV 12 @ $50 = $600)
            # Net variance is $0.00 while gross discrepancy exposure is $200.00!
            {
                "po_number": "PO-F4",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "60.00",
                "currency": "USD",
                "line_total": "600.00",
            },
            # 5: Clean partial fulfillment (ordered 10, shipped 6, invoiced 6 @ $50)
            {
                "po_number": "PO-F5",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            # 6: Partial invoicing accrual (ordered 10, shipped 8, invoiced 5 @ $50 -> accrual $150)
            {
                "po_number": "PO-F6",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Dock",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
        ]
        ship_rows = [
            {
                "shipment_id": f"SHP-F{i}",
                "po_number": f"PO-F{i}",
                "po_line_id": "1",
                "ship_date": "2026-02-03",
                "carrier": "UPS",
                "tracking_ref": f"TF{i}",
                "sku": "SKU-ELEC-001",
                "shipped_qty": str(sq),
                "warehouse_code": "WH-1",
            }
            for i, sq in [(1, 10), (2, 10), (3, 10), (4, 10), (5, 6), (6, 8)]
        ]
        inv_rows = [
            {
                "invoice_id": "INV-F1",
                "invoice_line_id": "1",
                "po_number": "PO-F1",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "12",
                "unit_price": "50.00",
                "line_amount": "600.00",
                "currency": "USD",
            },
            # Exact duplicate of INV-F1 -> quarantined, must not inflate totals
            {
                "invoice_id": "INV-F1",
                "invoice_line_id": "1",
                "po_number": "PO-F1",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "12",
                "unit_price": "50.00",
                "line_amount": "600.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-F2",
                "invoice_line_id": "INV-F2-01",
                "po_number": "PO-F2",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "55.00",
                "line_amount": "550.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-F3",
                "invoice_line_id": "INV-F3-01",
                "po_number": "PO-F3",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "12",
                "unit_price": "55.00",
                "line_amount": "660.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-F4",
                "invoice_line_id": "INV-F4-01",
                "po_number": "PO-F4",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "12",
                "unit_price": "50.00",
                "line_amount": "600.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-F5",
                "invoice_line_id": "1",
                "po_number": "PO-F5",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "6",
                "unit_price": "50.00",
                "line_amount": "300.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-F6",
                "invoice_line_id": "1",
                "po_number": "PO-F6",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "5",
                "unit_price": "50.00",
                "line_amount": "250.00",
                "currency": "USD",
            },
            # Unmatched orphan invoice ($500) -> quarantined
            {
                "invoice_id": "INV-ORPHAN",
                "invoice_line_id": "1",
                "po_number": "PO-NOPE",
                "po_line_id": "1",
                "invoice_date": "2026-02-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
        ]

        res = _run_mini_reconciliation(sample_products_df, po_rows, ship_rows, inv_rows)
        f1, f2, f3, f4, f5, f6 = res.reconciled_records

        assert f1.monetary_variance == Decimal("100.00")
        assert f1.discrepancy_exposure == Decimal("100.00")

        assert f2.monetary_variance == Decimal("50.00")
        assert f2.discrepancy_exposure == Decimal("50.00")

        # Simultaneous same-sign qty + price does not double-count
        assert f3.monetary_variance == Decimal("160.00")
        assert f3.discrepancy_exposure == Decimal("160.00")

        # Offset qty (-$100) + price (+$100) yields net 0.00 but gross exposure 200.00
        assert f4.monetary_variance == Decimal("0.00")
        assert f4.discrepancy_exposure == Decimal("200.00")

        assert f5.reconciliation_status == ReconciliationStatus.PARTIAL_FULFILLMENT
        assert f5.discrepancy_exposure == Decimal("0.00")
        assert f5.unbilled_receiving_amount == Decimal("0.00")

        # Partial invoicing (5 of 8 received units billed) accrues unbilled $150.00
        assert f6.unbilled_receiving_amount == Decimal("150.00")
        assert res.summary.unbilled_shipped_amount == Decimal("150.00")

        # Quarantined duplicate ($600) + orphan ($500) = $1,100.00
        assert res.summary.quarantined_invoice_amount == Decimal("1100.00")

        # 3.6 PRICE_MISMATCH provenance points directly to DatasetKind.INVOICES and invoice row
        price_excs = [e for e in res.exceptions if e.category == ExceptionCategory.PRICE_MISMATCH]
        assert len(price_excs) == 3  # F2, F3, F4
        assert all(e.source_dataset == DatasetKind.INVOICES for e in price_excs)
        assert {e.raw_reference for e in price_excs} == {"INV-F2-01", "INV-F3-01", "INV-F4-01"}
        # Source row for INV-F2 is row index 4 in sheet (after F1 + F1 duplicate -> row 4)
        inv_f2_exc = next(e for e in price_excs if e.raw_reference == "INV-F2-01")
        assert inv_f2_exc.source_row == 4

    def test_missing_required_column_raises_schema_validation_error(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        bad_po_df = pl.DataFrame(
            [{"po_number": "PO-1", "po_line_id": "1"}],
            schema={"po_number": pl.String, "po_line_id": pl.String},
        )
        ship_df = pl.DataFrame(
            [], schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.SHIPMENTS], pl.String)
        )
        inv_df = pl.DataFrame(
            [], schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.INVOICES], pl.String)
        )
        with pytest.raises(SchemaValidationError):
            reconcile_datasets(
                products_clean=clean_products_df(sample_products_df),
                po_clean=clean_purchase_orders_df(bad_po_df),
                shipments_clean=clean_shipments_df(ship_df),
                invoices_clean=clean_invoices_df(inv_df),
            )

    def test_mixed_currency_quarantine_excludes_foreign_amounts(
        self, sample_products_df: pl.DataFrame, tmp_path: Path
    ) -> None:
        from excel_reconciliation.reporting import write_reconciliation_report
        from excel_reconciliation.validation import validate_report_workbook

        po_rows = [
            {
                "po_number": "PO-USD-1",
                "po_line_id": "1",
                "order_date": "2026-03-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Precision Widget",
                "ordered_qty": "10",
                "unit_price": "50.00",
                "currency": "USD",
                "line_total": "500.00",
            },
            # Non-USD Purchase Order (400.00 EUR) must not appear with USD totals
            {
                "po_number": "PO-EUR-1",
                "po_line_id": "1",
                "order_date": "2026-03-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "Precision Widget",
                "ordered_qty": "8",
                "unit_price": "50.00",
                "currency": "EUR",
                "line_total": "400.00",
            },
        ]
        ship_rows = [
            {
                "shipment_id": "SHP-USD-1",
                "po_number": "PO-USD-1",
                "po_line_id": "1",
                "ship_date": "2026-03-02",
                "carrier": "DHL",
                "tracking_ref": "TRK-USD-1",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            }
        ]
        inv_rows = [
            # Matched USD invoice ($500.00)
            {
                "invoice_id": "INV-USD-1",
                "invoice_line_id": "1",
                "po_number": "PO-USD-1",
                "po_line_id": "1",
                "invoice_date": "2026-03-03",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
            # Duplicate USD invoice ($125.50) + orphan USD invoice ($300.00) -> $425.50 quarantined
            {
                "invoice_id": "INV-USD-ORPHAN",
                "invoice_line_id": "1",
                "po_number": "PO-NOPE",
                "po_line_id": "1",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "6",
                "unit_price": "50.00",
                "line_amount": "300.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-USD-DUP",
                "invoice_line_id": "1",
                "po_number": "PO-USD-DUP",
                "po_line_id": "1",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "2",
                "unit_price": "62.75",
                "line_amount": "125.50",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-USD-DUP",
                "invoice_line_id": "1",
                "po_number": "PO-USD-DUP",
                "po_line_id": "1",
                "invoice_date": "2026-03-04",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "2",
                "unit_price": "62.75",
                "line_amount": "125.50",
                "currency": "USD",
            },
            # Foreign-currency invoices: EUR (200.00 x2), GBP conflict (150 vs 175), CAD (900)
            {
                "invoice_id": "INV-EUR-1",
                "invoice_line_id": "1",
                "po_number": "PO-USD-1",
                "po_line_id": "1",
                "invoice_date": "2026-03-05",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "4",
                "unit_price": "50.00",
                "line_amount": "200.00",
                "currency": "EUR",
            },
            {
                "invoice_id": "INV-EUR-1",
                "invoice_line_id": "1",
                "po_number": "PO-USD-1",
                "po_line_id": "1",
                "invoice_date": "2026-03-05",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "4",
                "unit_price": "50.00",
                "line_amount": "200.00",
                "currency": "EUR",
            },
            {
                "invoice_id": "INV-GBP-C",
                "invoice_line_id": "1",
                "po_number": "PO-USD-1",
                "po_line_id": "1",
                "invoice_date": "2026-03-06",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "3",
                "unit_price": "50.00",
                "line_amount": "150.00",
                "currency": "GBP",
            },
            {
                "invoice_id": "INV-GBP-C",
                "invoice_line_id": "1",
                "po_number": "PO-USD-1",
                "po_line_id": "1",
                "invoice_date": "2026-03-06",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "3",
                "unit_price": "58.33",
                "line_amount": "175.00",
                "currency": "GBP",
            },
            {
                "invoice_id": "INV-CAD-1",
                "invoice_line_id": "1",
                "po_number": "PO-NOPE",
                "po_line_id": "1",
                "invoice_date": "2026-03-06",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "90.00",
                "line_amount": "900.00",
                "currency": "CAD",
            },
        ]

        res = _run_mini_reconciliation(sample_products_df, po_rows, ship_rows, inv_rows)

        # Only USD quarantined invoices (300.00 + 125.50 + 125.50 duplicate/orphan USD = 551.00)
        # Wait: INV-USD-DUP row 1 is also an orphan PO-USD-DUP ($125.50), row 2 is dup ($125.50)
        assert res.summary.total_expected_amount == Decimal("500.00")
        assert res.summary.total_invoiced_amount == Decimal("500.00")
        assert res.summary.quarantined_invoice_amount == Decimal("551.00")

        fx_excs = [
            exc
            for exc in res.exceptions
            if exc.original_currency is not None and exc.original_currency != "USD"
        ]
        assert len(fx_excs) == 6  # 1 EUR PO + 2 EUR inv + 2 GBP inv + 1 CAD inv
        for fx_exc in fx_excs:
            assert fx_exc.monetary_impact is None
            assert ReasonCode.CURRENCY_MISMATCH in fx_exc.reason_codes
            assert fx_exc.original_monetary_amount is not None
            assert fx_exc.original_currency in fx_exc.explanation
            assert f"{fx_exc.original_monetary_amount:.2f}" in fx_exc.explanation

        report_path = tmp_path / "fx_quarantine_report.xlsx"
        write_reconciliation_report(res, report_path)
        val = validate_report_workbook(report_path)
        assert val.is_valid, val.errors

        wb = openpyxl.load_workbook(report_path, data_only=False)
        try:
            ws_sum = wb["Summary"]
            assert Decimal(str(ws_sum["B22"].value)) == Decimal("551.00")
        finally:
            wb.close()
