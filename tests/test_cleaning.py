"""Unit tests for string/SKU normalization, numeric/date validation, schema, and deduplication."""

from __future__ import annotations

import os
import shutil
import subprocess
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
    normalize_identifier,
    normalize_sku,
    normalize_string,
    parse_date_value,
    parse_decimal_value,
    split_and_normalize_aliases,
    validate_schema,
)
from excel_reconciliation.excel_io import read_excel_dataset
from excel_reconciliation.models import REQUIRED_COLUMNS, DatasetKind, ReasonCode


class TestNormalizationPrimitives:
    """Verify deterministic string, identifier, SKU, and alias normalization."""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("  sku_elec_0042  ", "SKU-ELEC-0042"),
            ("SKU.ELEC.0042", "SKU-ELEC-0042"),
            ("Sku Elec 0042\t", "SKU-ELEC-0042"),
            ("sku//elec\\0042", "SKU-ELEC-0042"),
            ("---sku__ind--09---", "SKU-IND-09"),
            ("\u200bsku-001\ufeff", "SKU-001"),
            ("   ", None),
            ("-", None),
            ("NULL", None),
            (None, None),
        ],
    )
    def test_normalize_sku_variants(self, raw: str | None, expected: str | None) -> None:
        assert normalize_sku(raw) == expected

    def test_normalize_string_preserves_unicode_and_collapses_spaces(self) -> None:
        raw = "  Précision   Café Grinder — 日本の抹茶 500g \t "
        assert normalize_string(raw) == "Précision Café Grinder — 日本の抹茶 500g"
        assert normalize_identifier("  po-2026-10001 ") == "PO-2026-10001"

    def test_split_and_normalize_aliases_deduplicates_tokens(self) -> None:
        raw = " leg_001 ; SUP01.DOCK | leg-001 ,  "
        assert split_and_normalize_aliases(raw) == ("LEG-001", "SUP01-DOCK")


class TestScalarValidation:
    """Verify exact Decimal and calendar date validation rules."""

    def test_parse_decimal_distinguishes_zero_from_missing_and_invalid(self) -> None:
        assert parse_decimal_value("0") == (Decimal("0"), None)
        assert parse_decimal_value("0.00") == (Decimal("0.00"), None)
        assert parse_decimal_value("$1,250.75") == (Decimal("1250.75"), None)
        assert parse_decimal_value(None, required=True) == (None, ReasonCode.INVALID_NUMERIC_VALUE)
        assert parse_decimal_value(None, required=False) == (None, None)
        assert parse_decimal_value("TBD") == (None, ReasonCode.INVALID_NUMERIC_VALUE)
        assert parse_decimal_value("12O") == (None, ReasonCode.INVALID_NUMERIC_VALUE)
        assert parse_decimal_value("NaN") == (None, ReasonCode.INVALID_NUMERIC_VALUE)
        assert parse_decimal_value("=#REF!*H2") == (None, ReasonCode.BROKEN_EXCEL_FORMULA)

    def test_parse_decimal_rejects_negative_quantities_while_preserving_value(self) -> None:
        val, err = parse_decimal_value("-15.50", allow_negative=False)
        assert val == Decimal("-15.50")
        assert err == ReasonCode.NEGATIVE_QUANTITY_OR_PRICE

        val_ok, err_ok = parse_decimal_value("($15.50)", allow_negative=True)
        assert val_ok == Decimal("-15.50")
        assert err_ok is None

    @pytest.mark.parametrize(
        ("raw_date", "expected_iso", "expected_err"),
        [
            ("2026-03-15", "2026-03-15", None),
            ("03/15/2026", "2026-03-15", None),
            ("2026/03/15", "2026-03-15", None),
            ("15-Mar-2026", "2026-03-15", None),
            ("2026-02-29", None, ReasonCode.INVALID_DATE_VALUE),
            ("2026-02-30", None, ReasonCode.INVALID_DATE_VALUE),
            ("PENDING", None, ReasonCode.INVALID_DATE_VALUE),
            ("#VALUE!", None, ReasonCode.BROKEN_EXCEL_FORMULA),
        ],
    )
    def test_parse_date_formats_and_invalid_dates(
        self,
        raw_date: str,
        expected_iso: str | None,
        expected_err: ReasonCode | None,
    ) -> None:
        assert parse_date_value(raw_date) == (expected_iso, expected_err)


class TestDataFrameCleaningAndSchema:
    """Verify Polars DataFrame cleaning preserves raw data, flags duplicates, and checks schemas."""

    def test_schema_validation_detects_missing_and_extra_columns(self) -> None:
        observed = ["po_number", "po_line_id", "sku", "unexpected_notes"]
        res = validate_schema(DatasetKind.PURCHASE_ORDERS, observed)
        assert not res.is_valid
        assert "ordered_qty" in res.missing_columns
        assert "unit_price" in res.missing_columns
        assert res.extra_columns == ("unexpected_notes",)

    def test_cleaning_preserves_raw_values_and_never_drops_invalid_or_duplicate_rows(
        self, sample_products_df: pl.DataFrame
    ) -> None:
        dirty_rows = pl.DataFrame(
            [
                {
                    "product_id": " prd-001 ",
                    "canonical_sku": "sku_elec_001",
                    "product_name": "Duplicate Copy",
                    "category": "Electronics",
                    "uom": "EA",
                    "standard_cost": "50.00",
                    "supplier_id": "SUP-01",
                    "aliases": "",
                    "active_flag": "Y",
                },
                {
                    "product_id": "PRD-999",
                    "canonical_sku": "SKU-BAD-999",
                    "product_name": "Invalid Cost Row",
                    "category": "Industrial",
                    "uom": "EA",
                    "standard_cost": "TBD",
                    "supplier_id": "SUP-01",
                    "aliases": "",
                    "active_flag": "Y",
                },
            ],
            schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PRODUCTS], pl.String),
        )
        combined = pl.concat([sample_products_df, dirty_rows], how="vertical")
        cleaned = clean_products_df(combined)

        # All 5 rows retained
        assert cleaned.dataframe.height == 5
        assert cleaned.summary.total_rows == 5
        assert cleaned.summary.duplicate_rows == 1
        assert cleaned.summary.invalid_numeric_rows == 1
        assert cleaned.summary.valid_unique_rows == 3

        rows = cleaned.dataframe.to_dicts()
        # Row index 3 (4th row) is duplicate of PRD-001; raw string preserved
        assert rows[3]["raw_product_id"] == " prd-001 "
        assert rows[3]["norm_product_id"] == "PRD-001"
        assert rows[3]["_is_duplicate"] is True
        assert rows[3]["_duplicate_of_row"] == 2

        # Row index 4 (5th row) is invalid numeric, preserved with reason code
        assert rows[4]["raw_standard_cost"] == "TBD"
        assert rows[4]["_is_valid_record"] is False
        assert ReasonCode.INVALID_NUMERIC_VALUE.value in rows[4]["_validation_codes"]

    def test_null_keys_are_never_grouped_as_duplicates_of_each_other(self) -> None:
        ship_rows = [
            {
                "shipment_id": "",
                "po_number": "PO-100",
                "po_line_id": "PO-100-L01",
                "ship_date": "2026-02-01",
                "carrier": "DHL",
                "tracking_ref": "TRK-1",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "10",
                "warehouse_code": "WH-1",
            },
            {
                "shipment_id": "",
                "po_number": "PO-100",
                "po_line_id": "PO-100-L01",
                "ship_date": "2026-02-02",
                "carrier": "DHL",
                "tracking_ref": "TRK-2",
                "sku": "SKU-ELEC-001",
                "shipped_qty": "15",
                "warehouse_code": "WH-1",
            },
        ]
        df = pl.DataFrame(
            ship_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.SHIPMENTS], pl.String)
        )
        cleaned = clean_shipments_df(df)
        assert cleaned.summary.duplicate_rows == 0

    def test_po_cleaning_normalizes_skus_and_flags_bad_numbers(self) -> None:
        po_rows = [
            {
                "po_number": "  po-2026-001 ",
                "po_line_id": " po-2026-001-l01 ",
                "order_date": "02/14/2026",
                "supplier_id": "sup-01",
                "sku": " sku.elec.001 ",
                "product_description": "  USB-C Dock ",
                "ordered_qty": "20",
                "unit_price": "55.00",
                "currency": "usd",
                "line_total": "1100.00",
            },
            {
                "po_number": "PO-2026-002",
                "po_line_id": "PO-2026-002-L01",
                "order_date": "2026-02-15",
                "supplier_id": "SUP-01",
                "sku": "SKU-ELEC-001",
                "product_description": "USB-C Dock",
                "ordered_qty": "-5",
                "unit_price": "55.00",
                "currency": "USD",
                "line_total": "-275.00",
            },
        ]
        df = pl.DataFrame(
            po_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS], pl.String)
        )
        cleaned = clean_purchase_orders_df(df)
        out = cleaned.dataframe.to_dicts()
        assert out[0]["norm_po_number"] == "PO-2026-001"
        assert out[0]["norm_sku"] == "SKU-ELEC-001"
        assert out[0]["norm_order_date"] == "2026-02-14"
        assert out[0]["_is_valid_record"] is True
        assert out[1]["_is_valid_record"] is False
        assert ReasonCode.NEGATIVE_QUANTITY_OR_PRICE.value in out[1]["_validation_codes"]

    @pytest.mark.parametrize(
        "bad_val",
        [
            "1_000",
            "1,2,3",
            "($-10)",
            "(+10)",
            "1e50",
            "1e5",
            True,
            False,
            12.5,
        ],
    )
    def test_parse_decimal_rejects_malformed_strings_bools_and_floats(
        self, bad_val: object
    ) -> None:
        parsed, reason = parse_decimal_value(bad_val, allow_negative=True)
        assert parsed is None
        assert reason == ReasonCode.INVALID_NUMERIC_VALUE

    def test_product_master_conflicts_and_shared_canonical_skus_quarantined(self) -> None:
        rows = [
            # Conflicting pair sharing PRD-001 with different cost & canonical SKU
            {
                "product_id": "PRD-001",
                "canonical_sku": "SKU-001",
                "product_name": "Sensor A",
                "category": "Industrial",
                "uom": "EA",
                "standard_cost": "10.00",
                "supplier_id": "SUP-01",
                "aliases": "ALS-A",
                "active_flag": "Y",
            },
            {
                "product_id": "PRD-001",
                "canonical_sku": "SKU-001-REV",
                "product_name": "Sensor A",
                "category": "Industrial",
                "uom": "EA",
                "standard_cost": "12.50",
                "supplier_id": "SUP-01",
                "aliases": "ALS-B",
                "active_flag": "Y",
            },
            # Distinct product_ids sharing the same canonical SKU
            {
                "product_id": "PRD-002",
                "canonical_sku": "SKU-SHARED",
                "product_name": "Valve X",
                "category": "Industrial",
                "uom": "EA",
                "standard_cost": "25.00",
                "supplier_id": "SUP-02",
                "aliases": "",
                "active_flag": "Y",
            },
            {
                "product_id": "PRD-003",
                "canonical_sku": "SKU-SHARED",
                "product_name": "Valve Y",
                "category": "Industrial",
                "uom": "EA",
                "standard_cost": "25.00",
                "supplier_id": "SUP-03",
                "aliases": "",
                "active_flag": "Y",
            },
        ]
        df = pl.DataFrame(
            rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PRODUCTS], pl.String)
        )
        cleaned = clean_products_df(df)
        out = cleaned.dataframe.to_dicts()

        assert cleaned.summary.conflicting_duplicate_rows == 4
        assert ReasonCode.CONFLICTING_MASTER_RECORD.value in out[0]["_validation_codes"]
        assert ReasonCode.CONFLICTING_MASTER_RECORD.value in out[1]["_validation_codes"]
        assert out[0]["_is_valid_record"] is False
        assert out[1]["_is_valid_record"] is False

        assert ReasonCode.AMBIGUOUS_MASTER_SKU.value in out[2]["_validation_codes"]
        assert ReasonCode.AMBIGUOUS_MASTER_SKU.value in out[3]["_validation_codes"]
        assert out[2]["_is_valid_record"] is False
        assert out[3]["_is_valid_record"] is False

    def test_currency_mismatch_and_conflicting_duplicate_po_quarantine(self) -> None:
        po_rows = [
            {
                "po_number": "PO-1",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-001",
                "product_description": "Widget",
                "ordered_qty": "10",
                "unit_price": "20.00",
                "currency": "USD",
                "line_total": "200.00",
            },
            # Conflicting duplicate PO line with different qty
            {
                "po_number": "PO-1",
                "po_line_id": "1",
                "order_date": "2026-02-01",
                "supplier_id": "SUP-01",
                "sku": "SKU-001",
                "product_description": "Widget",
                "ordered_qty": "15",
                "unit_price": "20.00",
                "currency": "USD",
                "line_total": "300.00",
            },
            # Non-USD currency
            {
                "po_number": "PO-2",
                "po_line_id": "1",
                "order_date": "2026-02-02",
                "supplier_id": "SUP-01",
                "sku": "SKU-001",
                "product_description": "Widget",
                "ordered_qty": "10",
                "unit_price": "20.00",
                "currency": "EUR",
                "line_total": "200.00",
            },
        ]
        po_df = pl.DataFrame(
            po_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS], pl.String)
        )
        po_clean = clean_purchase_orders_df(po_df, reporting_currency="USD")
        out_po = po_clean.dataframe.to_dicts()
        assert po_clean.summary.conflicting_duplicate_rows == 2
        assert ReasonCode.CONFLICTING_DUPLICATE_PO_LINE.value in out_po[0]["_validation_codes"]
        assert ReasonCode.CONFLICTING_DUPLICATE_PO_LINE.value in out_po[1]["_validation_codes"]
        assert ReasonCode.CURRENCY_MISMATCH.value in out_po[2]["_validation_codes"]

        inv_rows = [
            {
                "invoice_id": "INV-EUR",
                "invoice_line_id": "1",
                "po_number": "PO-1",
                "po_line_id": "1",
                "invoice_date": "2026-02-03",
                "supplier_id": "SUP-01",
                "sku": "SKU-001",
                "invoiced_qty": "10",
                "unit_price": "20.00",
                "line_amount": "200.00",
                "currency": "EUR",
            }
        ]
        inv_df = pl.DataFrame(
            inv_rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.INVOICES], pl.String)
        )
        inv_clean = clean_invoices_df(inv_df, reporting_currency="USD")
        out_inv = inv_clean.dataframe.to_dicts()
        assert ReasonCode.CURRENCY_MISMATCH.value in out_inv[0]["_validation_codes"]


def _resolve_synchronous_soffice() -> str | None:
    """Locate a synchronous soffice executable, unwrapping macOS Nix `open -na` launchers."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        return None
    candidate = Path(soffice).resolve()
    app_binary = candidate.parent.parent / "Applications/LibreOffice.app/Contents/MacOS/soffice"
    if app_binary.is_file():
        return str(app_binary)
    return str(candidate)


def _recalculate_xlsx_via_libreoffice(xlsx_path: Path, work_dir: Path) -> Path | None:
    """Recalculate an xlsx workbook via headless LibreOffice Calc when available."""
    soffice = _resolve_synchronous_soffice()
    if soffice is None:
        if os.environ.get("REQUIRE_LIBREOFFICE") == "1":
            raise AssertionError("REQUIRE_LIBREOFFICE=1 set, but soffice binary was not found.")
        return None

    out_dir = work_dir / "lo_out"
    profile_dir = work_dir / "lo_profile"
    out_dir.mkdir(parents=True, exist_ok=True)
    profile_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        f"-env:UserInstallation={profile_dir.as_uri()}",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(out_dir),
        str(xlsx_path),
    ]
    proc = subprocess.run(cmd, check=True, capture_output=True, text=True)
    recalculated = out_dir / xlsx_path.name
    assert recalculated.exists(), f"LibreOffice output missing: {proc.stdout} {proc.stderr}"
    return recalculated


class TestExcelSumSemantics:
    """Verify Excel SUM reference semantics and LibreOffice Calc parity."""

    def test_excel_sum_ignores_text_formatted_numbers_and_matches_libreoffice(
        self, tmp_path: Path
    ) -> None:
        wb_path = tmp_path / "sum_semantics.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Data"

        headers = ["val", "qty", "unit_price", "formula_col", "err_col", "probe_out"]
        ws.append(headers)

        # Rows 2..11 exercise numeric cells, numeric-looking text, general text, blank, formulas
        # Numerics in A2:A11: 10 (A2), 25.5 (A3), 14.50 (A4), -5 (A5), formula 10*3=30 (A11) -> 75
        rows_a: list[tuple[object, str | None, bool]] = [
            (10, None, False),  # Row 2: int (n) -> 10
            (25.5, None, False),  # Row 3: float (n) -> 25.5
            (Decimal("14.50"), None, False),  # Row 4: Decimal (n) -> 14.50
            (-5, None, False),  # Row 5: negative int (n) -> -5  [A2:A5 sum = 45]
            ("100", "s", False),  # Row 6: numeric-looking text -> ignored by SUM(range)
            ("25.50", "s", True),  # Row 7: quote-prefixed numeric text -> ignored by SUM(range)
            ("$40.00", "s", False),  # Row 8: currency-formatted text -> ignored by SUM(range)
            ("TBD", "s", False),  # Row 9: non-numeric text -> ignored by SUM(range)
            (None, None, False),  # Row 10: blank -> ignored by SUM(range)
            ("=10*3", "f", False),  # Row 11: numeric formula (=30) -> summed
            ("=10/2", "f", False),  # Row 12: numeric formula (=5) -> outside A2:A11
            (True, "b", False),  # Row 13: boolean -> ignored by Excel SUM(ref) (LO differs)
        ]
        for idx, (val, dtype, q_prefix) in enumerate(rows_a, start=2):
            c = ws.cell(row=idx, column=1, value=val)
            if dtype is not None:
                c.data_type = dtype
            if q_prefix:
                c.quotePrefix = True
            ws.cell(row=idx, column=2, value=10)
            p_cell = ws.cell(row=idx, column=3, value="25.50")
            p_cell.data_type = "s"

        # Error propagation helper column E
        ws.cell(row=2, column=5, value=10)
        ws.cell(row=3, column=5, value="=1/0")

        # Probes 2..10 match both Excel and LibreOffice Calc semantics
        lo_parity_probes: list[tuple[int, str, str]] = [
            (2, "=SUM(A2:A11)", "75"),  # numeric + formula cells only
            (3, "=SUM(A1:A11)", "75"),  # header text in A1 ignored
            (4, "=SUM(A2,A6,A9)", "10"),  # discrete refs ignore text
            (5, "=SUM(A2:A5,5,A6*2)", "250"),  # range(45) + literal(5) + arith(200)
            (6, "=ROUND(SUM(A2:A11)/3,2)", "25.00"),  # rounded aggregate
            (7, "=B2*C2", "255.00"),  # normal arithmetic still coerces numeric text C2="25.50"
            (8, "=A2+A6", "110"),  # normal arithmetic coerces direct text ref A6="100"
            (9, "=SUM(E2:E3)", "#DIV/0!"),  # error cell inside range propagates
            (10, "=A2+A9", "#VALUE!"),  # non-numeric text in normal arithmetic gives #VALUE!
        ]
        # Probe 11 checks Excel's boolean-in-reference rule (Calc historically treats TRUE as 1)
        excel_only_probes: list[tuple[int, str, str]] = [
            (11, "=SUM(A2,A13)", "10"),
        ]
        for row_idx, formula_str, _ in lo_parity_probes + excel_only_probes:
            ws.cell(row=row_idx, column=6, value=formula_str)

        wb.save(wb_path)
        wb.close()

        loaded = read_excel_dataset(wb_path, DatasetKind.PURCHASE_ORDERS, sheet_name="Data")
        out_by_row = {int(r["_source_row"]): r["probe_out"] for r in loaded.dataframe.to_dicts()}
        for row_idx, formula_str, expected_str in lo_parity_probes + excel_only_probes:
            actual_str = out_by_row[row_idx]
            if expected_str.startswith("#"):
                assert actual_str == expected_str, f"{formula_str}: {actual_str} != {expected_str}"
            else:
                assert Decimal(actual_str) == Decimal(expected_str), (
                    f"{formula_str}: {actual_str} != {expected_str}"
                )

        # When LibreOffice is available (or required via REQUIRE_LIBREOFFICE=1), verify parity
        recalculated_path = _recalculate_xlsx_via_libreoffice(wb_path, tmp_path)
        if recalculated_path is not None:
            wb_lo = openpyxl.load_workbook(recalculated_path, data_only=True)
            try:
                ws_lo = wb_lo["Data"]
                for row_idx, formula_str, _ in lo_parity_probes:
                    lo_val = ws_lo.cell(row=row_idx, column=6).value
                    py_val = out_by_row[row_idx]
                    if str(py_val).startswith("#"):
                        assert str(lo_val) == py_val, f"LibreOffice mismatch on {formula_str}"
                    else:
                        assert Decimal(str(lo_val)) == Decimal(str(py_val)), (
                            f"LibreOffice mismatch on {formula_str}: LO={lo_val} vs PY={py_val}"
                        )
            finally:
                wb_lo.close()
