"""Unit tests for deterministic 4-tier product and order-line matching."""

from __future__ import annotations

import polars as pl

from excel_reconciliation.cleaning import (
    clean_invoices_df,
    clean_products_df,
    clean_purchase_orders_df,
    clean_shipments_df,
)
from excel_reconciliation.matching import (
    OrderLineMatcher,
    ProductMatcher,
    match_purchase_orders_df,
)
from excel_reconciliation.models import (
    REQUIRED_COLUMNS,
    DatasetKind,
    MatchStatus,
    PoLineKey,
    ReasonCode,
    format_business_key,
)
from excel_reconciliation.reconciliation import reconcile_datasets


def _make_po_df(rows: list[dict[str, str]]) -> pl.DataFrame:
    return pl.DataFrame(
        rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS], pl.String)
    )


def _make_ship_df(rows: list[dict[str, str]]) -> pl.DataFrame:
    return pl.DataFrame(
        rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.SHIPMENTS], pl.String)
    )


class TestProductAndOrderMatching:
    """Verify exact ID, explicit alias, composite key, and ambiguity quarantine."""

    def test_product_matcher_priority_and_ambiguous_alias_quarantine(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        products_clean = clean_products_df(sample_products_df)
        matcher = ProductMatcher(products_clean.dataframe)

        # Tier 1: Exact canonical SKU
        r_exact = matcher.resolve_sku("SKU-ELEC-001")
        assert r_exact.status == MatchStatus.EXACT_ID
        assert r_exact.reason == ReasonCode.MATCHED_EXACT_ID
        assert r_exact.product is not None and r_exact.product.product_id == "PRD-001"

        # Tier 2: Explicit unique alias
        r_alias = matcher.resolve_sku("SUP01-DOCK")
        assert r_alias.status == MatchStatus.ALIAS_MAP
        assert r_alias.reason == ReasonCode.MATCHED_PRODUCT_ALIAS
        assert r_alias.product is not None and r_alias.product.product_id == "PRD-001"

        # Tier 4a: Colliding alias mapped to PRD-002 and PRD-003 -> never guessed!
        r_amb = matcher.resolve_sku("SHARED-AMB-99")
        assert r_amb.status == MatchStatus.AMBIGUOUS
        assert r_amb.reason == ReasonCode.AMBIGUOUS_ALIAS_MAPPING
        assert r_amb.product is None
        assert set(r_amb.candidate_product_ids) == {"PRD-002", "PRD-003"}

        # Tier 4b: Unrecognized SKU
        r_unk = matcher.resolve_sku("SKU-NONEXISTENT-999")
        assert r_unk.status == MatchStatus.UNMATCHED
        assert r_unk.reason == ReasonCode.UNKNOWN_SKU

    def test_transaction_matching_tiers_composite_and_multiline_ambiguity(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        products_clean = clean_products_df(sample_products_df)
        prod_matcher = ProductMatcher(products_clean.dataframe)

        po_df = _make_po_df(
            [
                # PO-100: single line for PRD-001 (line_id = "1")
                {
                    "po_number": "PO-100",
                    "po_line_id": "1",
                    "order_date": "2026-01-10",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "product_description": "Dock",
                    "ordered_qty": "20",
                    "unit_price": "50.00",
                    "currency": "USD",
                    "line_total": "1000.00",
                },
                # PO-200: two separate lines for the SAME SKU PRD-001 ("1" and "2")
                {
                    "po_number": "PO-200",
                    "po_line_id": "1",
                    "order_date": "2026-01-12",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "product_description": "Dock Batch A",
                    "ordered_qty": "10",
                    "unit_price": "50.00",
                    "currency": "USD",
                    "line_total": "500.00",
                },
                {
                    "po_number": "PO-200",
                    "po_line_id": "2",
                    "order_date": "2026-01-12",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "product_description": "Dock Batch B",
                    "ordered_qty": "30",
                    "unit_price": "50.00",
                    "currency": "USD",
                    "line_total": "1500.00",
                },
            ]
        )
        po_clean = clean_purchase_orders_df(po_df)
        matched_po = match_purchase_orders_df(po_clean.dataframe, prod_matcher)
        line_matcher = OrderLineMatcher(matched_po, prod_matcher)

        ship_df = _make_ship_df(
            [
                # 0: Exact PO + line_id + canonical SKU -> EXACT_ID
                {
                    "shipment_id": "SHP-1",
                    "po_number": "PO-100",
                    "po_line_id": "1",
                    "ship_date": "2026-01-15",
                    "carrier": "UPS",
                    "tracking_ref": "T1",
                    "sku": "sku_elec_001",
                    "shipped_qty": "10",
                    "warehouse_code": "WH1",
                },
                # 1: Exact PO + line_id + product alias -> ALIAS_MAP
                {
                    "shipment_id": "SHP-2",
                    "po_number": "PO-100",
                    "po_line_id": "1",
                    "ship_date": "2026-01-16",
                    "carrier": "UPS",
                    "tracking_ref": "T2",
                    "sku": "sup01_dock",
                    "shipped_qty": "10",
                    "warehouse_code": "WH1",
                },
                # 2: Missing po_line_id on single-line PO-100 -> COMPOSITE_KEY
                {
                    "shipment_id": "SHP-3",
                    "po_number": "PO-100",
                    "po_line_id": "",
                    "ship_date": "2026-01-17",
                    "carrier": "UPS",
                    "tracking_ref": "T3",
                    "sku": "SKU-ELEC-001",
                    "shipped_qty": "5",
                    "warehouse_code": "WH1",
                },
                # 3: Missing po_line_id on multi-line same-SKU PO-200 -> AMBIGUOUS!
                {
                    "shipment_id": "SHP-4",
                    "po_number": "PO-200",
                    "po_line_id": "",
                    "ship_date": "2026-01-18",
                    "carrier": "UPS",
                    "tracking_ref": "T4",
                    "sku": "SKU-ELEC-001",
                    "shipped_qty": "10",
                    "warehouse_code": "WH1",
                },
                # 4: Line ID "1" on PO-100, but SKU is PRD-002 (product conflict!) -> AMBIGUOUS!
                {
                    "shipment_id": "SHP-5",
                    "po_number": "PO-100",
                    "po_line_id": "1",
                    "ship_date": "2026-01-19",
                    "carrier": "UPS",
                    "tracking_ref": "T5",
                    "sku": "SKU-HOME-002",
                    "shipped_qty": "10",
                    "warehouse_code": "WH1",
                },
            ]
        )
        ship_clean = clean_shipments_df(ship_df)
        matched_ship = line_matcher.match_transactions_df(ship_clean.dataframe).to_dicts()

        assert matched_ship[0]["match_status"] == MatchStatus.EXACT_ID.value
        assert matched_ship[0]["matched_po_number"] == "PO-100"
        assert matched_ship[0]["matched_po_line_id"] == "1"

        assert matched_ship[1]["match_status"] == MatchStatus.ALIAS_MAP.value
        assert matched_ship[1]["matched_po_number"] == "PO-100"
        assert matched_ship[1]["matched_po_line_id"] == "1"

        assert matched_ship[2]["match_status"] == MatchStatus.COMPOSITE_KEY.value
        assert matched_ship[2]["matched_po_number"] == "PO-100"
        assert matched_ship[2]["matched_po_line_id"] == "1"

        assert matched_ship[3]["match_status"] == MatchStatus.AMBIGUOUS.value
        assert matched_ship[3]["match_reason"] == ReasonCode.MULTIPLE_CANDIDATE_PO_LINES.value
        assert matched_ship[3]["matched_po_number"] is None
        assert matched_ship[3]["matched_po_line_id"] is None
        cand_pairs = {
            (c["po_number"], c["po_line_id"]) for c in matched_ship[3]["candidate_po_lines"]
        }
        assert cand_pairs == {("PO-200", "1"), ("PO-200", "2")}

        assert matched_ship[4]["match_status"] == MatchStatus.AMBIGUOUS.value
        assert matched_ship[4]["match_reason"] == ReasonCode.LINE_PRODUCT_CONFLICT.value
        assert matched_ship[4]["matched_po_number"] is None
        assert matched_ship[4]["matched_po_line_id"] is None

    def test_supplier_scoped_invoice_ids_and_supplier_mismatch_quarantine(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        prod_clean = clean_products_df(sample_products_df)
        prod_matcher = ProductMatcher(prod_clean.dataframe)

        po_df = _make_po_df(
            [
                {
                    "po_number": "PO-SUP1",
                    "po_line_id": "1",
                    "order_date": "2026-01-10",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "product_description": "Dock",
                    "ordered_qty": "10",
                    "unit_price": "50.00",
                    "currency": "USD",
                    "line_total": "500.00",
                },
                {
                    "po_number": "PO-SUP2",
                    "po_line_id": "1",
                    "order_date": "2026-01-10",
                    "supplier_id": "SUP-02",
                    "sku": "SKU-HOME-002",
                    "product_description": "Grinder",
                    "ordered_qty": "5",
                    "unit_price": "120.00",
                    "currency": "USD",
                    "line_total": "600.00",
                },
            ]
        )
        po_clean = clean_purchase_orders_df(po_df)
        matched_po = match_purchase_orders_df(po_clean.dataframe, prod_matcher)
        line_matcher = OrderLineMatcher(matched_po, prod_matcher)

        # Two suppliers using identical (invoice_id, invoice_line_id) plus one supplier mismatch
        inv_rows = [
            {
                "invoice_id": "INV-1000",
                "invoice_line_id": "1",
                "po_number": "PO-SUP1",
                "po_line_id": "1",
                "invoice_date": "2026-01-12",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-1000",
                "invoice_line_id": "1",
                "po_number": "PO-SUP2",
                "po_line_id": "1",
                "invoice_date": "2026-01-12",
                "supplier_id": "SUP-02",
                "sku": "SKU-HOME-002",
                "invoiced_qty": "5",
                "unit_price": "120.00",
                "line_amount": "600.00",
                "currency": "USD",
            },
            {
                "invoice_id": "INV-WRONG-SUP",
                "invoice_line_id": "1",
                "po_number": "PO-SUP1",
                "po_line_id": "1",
                "invoice_date": "2026-01-12",
                "supplier_id": "SUP-99",
                "sku": "SKU-ELEC-001",
                "invoiced_qty": "10",
                "unit_price": "50.00",
                "line_amount": "500.00",
                "currency": "USD",
            },
        ]
        inv_df = pl.DataFrame(
            inv_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.INVOICES], pl.String)
        )
        inv_clean = clean_invoices_df(inv_df)
        assert inv_clean.summary.duplicate_rows == 0

        matched_inv = line_matcher.match_transactions_df(
            inv_clean.dataframe, validate_supplier=True
        ).to_dicts()
        assert matched_inv[0]["match_status"] == MatchStatus.EXACT_ID.value
        assert matched_inv[1]["match_status"] == MatchStatus.EXACT_ID.value
        assert matched_inv[2]["match_status"] == MatchStatus.UNMATCHED.value
        assert matched_inv[2]["match_reason"] == ReasonCode.SUPPLIER_MISMATCH.value

    def test_collision_free_business_keys_across_delimiter_boundaries(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        # Injective formatting never collides when identifiers contain '|' or '::'
        assert PoLineKey("A::B", "C") != PoLineKey("A", "B::C")
        assert format_business_key("A|B", "C", "PRD-001") != format_business_key(
            "A", "B|C", "PRD-001"
        )

        prod_clean = clean_products_df(sample_products_df)
        po_df = _make_po_df(
            [
                {
                    "po_number": "A::B",
                    "po_line_id": "C",
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
                    "po_number": "A",
                    "po_line_id": "B::C",
                    "order_date": "2026-02-01",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "product_description": "Dock",
                    "ordered_qty": "20",
                    "unit_price": "50.00",
                    "currency": "USD",
                    "line_total": "1000.00",
                },
                {
                    "po_number": "A|B",
                    "po_line_id": "C",
                    "order_date": "2026-02-01",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "product_description": "Dock",
                    "ordered_qty": "30",
                    "unit_price": "50.00",
                    "currency": "USD",
                    "line_total": "1500.00",
                },
                {
                    "po_number": "A",
                    "po_line_id": "B|C",
                    "order_date": "2026-02-01",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "product_description": "Dock",
                    "ordered_qty": "40",
                    "unit_price": "50.00",
                    "currency": "USD",
                    "line_total": "2000.00",
                },
            ]
        )
        ship_df = _make_ship_df(
            [
                {
                    "shipment_id": f"SHP-{idx}",
                    "po_number": po_num,
                    "po_line_id": lid,
                    "ship_date": "2026-02-03",
                    "carrier": "DHL",
                    "tracking_ref": f"T-{idx}",
                    "sku": "SKU-ELEC-001",
                    "shipped_qty": str(qty),
                    "warehouse_code": "WH-1",
                }
                for idx, (po_num, lid, qty) in enumerate(
                    [("A::B", "C", 10), ("A", "B::C", 20), ("A|B", "C", 30), ("A", "B|C", 40)],
                    start=1,
                )
            ]
        )
        inv_df = pl.DataFrame(
            [
                {
                    "invoice_id": f"INV-{idx}",
                    "invoice_line_id": "1",
                    "po_number": po_num,
                    "po_line_id": lid,
                    "invoice_date": "2026-02-04",
                    "supplier_id": "SUP-01",
                    "sku": "SKU-ELEC-001",
                    "invoiced_qty": str(qty),
                    "unit_price": "50.00",
                    "line_amount": f"{qty * 50}.00",
                    "currency": "USD",
                }
                for idx, (po_num, lid, qty) in enumerate(
                    [("A::B", "C", 10), ("A", "B::C", 20), ("A|B", "C", 30), ("A", "B|C", 40)],
                    start=1,
                )
            ],
            schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.INVOICES], pl.String),
        )

        result = reconcile_datasets(
            products_clean=prod_clean,
            po_clean=clean_purchase_orders_df(po_df),
            shipments_clean=clean_shipments_df(ship_df),
            invoices_clean=clean_invoices_df(inv_df),
        )
        assert result.summary.reconciled_lines == 4
        assert result.summary.total_ordered_qty == 100
        assert result.summary.total_shipped_qty == 100
        assert result.summary.total_invoiced_qty == 100
        assert [r.ordered_qty for r in result.reconciled_records] == [10, 20, 30, 40]
