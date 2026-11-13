"""End-to-end tests for deterministic synthetic data generation and ground-truth accuracy."""

from __future__ import annotations

import json
from pathlib import Path

from excel_reconciliation.demo_data import build_synthetic_records, generate_demo_datasets
from excel_reconciliation.reconciliation import run_reconciliation_from_directory
from excel_reconciliation.reporting import write_reconciliation_report
from excel_reconciliation.validation import validate_report_workbook


class TestSyntheticDataAndEndToEndGroundTruth:
    """Verify deterministic synthetic dataset generation and 100% ground-truth agreement."""

    def test_synthetic_record_builder_is_deterministic_and_hits_target_sizes(self) -> None:
        b1 = build_synthetic_records(seed=20260401)
        b2 = build_synthetic_records(seed=20260401)
        b_other = build_synthetic_records(seed=999999)

        assert b1.products == b2.products
        assert b1.purchase_orders == b2.purchase_orders
        assert b1.shipments == b2.shipments
        assert b1.invoices == b2.invoices
        assert b1.purchase_orders != b_other.purchase_orders

        assert len(b1.products) == 1500
        assert len(b1.purchase_orders) == 2000
        assert len(b1.shipments) == 2200
        assert len(b1.invoices) == 2100

    def test_full_demo_pipeline_matches_ground_truth_and_validates(self, tmp_path: Path) -> None:
        input_dir = tmp_path / "input"
        output_dir = tmp_path / "output"

        written = generate_demo_datasets(input_dir, seed=20260401)
        assert (input_dir / "product_master.xlsx").exists()
        assert (input_dir / "purchase_orders.xlsx").exists()
        assert (input_dir / "shipments.xlsx").exists()
        assert (input_dir / "supplier_invoices.xlsx").exists()

        gt = json.loads(written["ground_truth"].read_text(encoding="utf-8"))
        result = run_reconciliation_from_directory(input_dir)

        assert result.summary.total_po_rows == 2000
        assert result.summary.unique_valid_po_lines == gt["expected_unique_valid_po_lines"]
        assert result.summary.duplicate_po_rows == 25

        # Verify actual status distribution across all 2,000 PO rows matches ground truth
        observed_dist: dict[str, int] = {}
        for rec in result.reconciled_records:
            st = rec.reconciliation_status.value
            observed_dist[st] = observed_dist.get(st, 0) + 1

        assert observed_dist == gt["expected_po_status_counts"]

        # Verify all 2,000 PO source rows match expected_po_status_by_source_row
        expected_by_row: dict[str, str] = gt["expected_po_status_by_source_row"]
        assert len(expected_by_row) == 2000
        for rec in result.reconciled_records:
            row_key = str(rec.source_row)
            assert row_key in expected_by_row
            assert rec.reconciliation_status.value == expected_by_row[row_key]

        # Verify independently computed financial totals and all-dataset exception counts
        fin_gt: dict[str, str] = gt["expected_financial_totals"]
        assert str(result.summary.total_ordered_qty) == fin_gt["total_ordered_qty"]
        assert str(result.summary.total_shipped_qty) == fin_gt["total_shipped_qty"]
        assert str(result.summary.total_invoiced_qty) == fin_gt["total_invoiced_qty"]
        assert str(result.summary.total_outstanding_qty) == fin_gt["total_outstanding_qty"]
        assert f"{result.summary.total_expected_amount:.2f}" == fin_gt["total_expected_amount"]
        assert f"{result.summary.total_invoiced_amount:.2f}" == fin_gt["total_invoiced_amount"]
        assert f"{result.summary.net_billed_variance:.2f}" == fin_gt["net_billed_variance"]
        assert (
            f"{result.summary.absolute_monetary_discrepancy:.2f}"
            == fin_gt["absolute_monetary_discrepancy"]
        )
        assert f"{result.summary.unbilled_shipped_amount:.2f}" == fin_gt["unbilled_shipped_amount"]
        assert (
            f"{result.summary.quarantined_invoice_amount:.2f}"
            == fin_gt["quarantined_invoice_amount"]
        )

        exc_by_ds: dict[str, int] = {}
        for exc in result.exceptions:
            ds_key = exc.source_dataset.value
            exc_by_ds[ds_key] = exc_by_ds.get(ds_key, 0) + 1
        assert exc_by_ds == gt["expected_exception_counts_by_dataset"]
        assert result.summary.total_exception_records == gt["expected_total_exceptions"]

        # Verify every individual valid unique PO line matches its exact ground-truth status
        expected_line_map: dict[str, str] = gt["expected_po_line_statuses"]
        for rec in result.reconciled_records:
            if rec.reconciliation_status.value == "DUPLICATE_TRANSACTION":
                continue
            if rec.normalized_po_line_id in expected_line_map:
                assert (
                    rec.reconciliation_status.value == expected_line_map[rec.normalized_po_line_id]
                ), (
                    f"Mismatch on {rec.normalized_po_line_id}: "
                    f"got {rec.reconciliation_status.value}, "
                    f"expected {expected_line_map[rec.normalized_po_line_id]}"
                )

        report_path = output_dir / "reconciliation_report.xlsx"
        write_reconciliation_report(result, report_path)
        val_res = validate_report_workbook(report_path)
        assert val_res.is_valid, f"Report validation failed: {val_res.errors}"
