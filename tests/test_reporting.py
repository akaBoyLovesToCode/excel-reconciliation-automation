"""Tests for Excel formula preservation, report generation, and workbook validation."""

from __future__ import annotations

import time
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from excel_reconciliation.excel_io import evaluate_worksheet_formula, read_excel_dataset
from excel_reconciliation.models import (
    REQUIRED_COLUMNS,
    DatasetKind,
    ReasonCode,
    ReconciliationStatus,
)
from excel_reconciliation.reconciliation import run_reconciliation_from_directory
from excel_reconciliation.reporting import write_reconciliation_report
from excel_reconciliation.validation import (
    REQUIRED_WORKSHEETS,
    FormulaEvaluationLimits,
    FormulaResourceLimitError,
    WorkbookFormulaCache,
    evaluate_workbook_formula,
    validate_report_workbook,
)


class TestExcelFormulaPreservationAndReporting:
    """Verify input formula evaluation/preservation and output Excel workbook integrity."""

    def test_input_formulas_evaluated_preserved_and_report_validated(
        self, tmp_path: Path, write_xlsx_helper
    ) -> None:
        in_dir = tmp_path / "in"
        in_dir.mkdir()

        def _save_sheet(fname: str, cols: tuple[str, ...], rows: list[list[object]]) -> None:
            wb = openpyxl.Workbook()
            ws = wb.active
            assert ws is not None
            ws.append(list(cols))
            for r in rows:
                ws.append(r)
            wb.save(in_dir / fname)
            wb.close()

        _save_sheet(
            "product_master.xlsx",
            REQUIRED_COLUMNS[DatasetKind.PRODUCTS],
            [
                [
                    "PRD-01",
                    "SKU-A-01",
                    "Precision Sensor",
                    "Industrial",
                    "EA",
                    Decimal("20.00"),
                    "SUP-01",
                    "LEG-A01",
                    "Y",
                ]
            ],
        )
        _save_sheet(
            "purchase_orders.xlsx",
            REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS],
            [
                # Row 2: Live formula =G2*H2 (no cached value saved by Excel)
                [
                    "PO-01",
                    "PO-01-L01",
                    "2026-03-01",
                    "SUP-01",
                    "SKU-A-01",
                    "Precision Sensor",
                    10,
                    Decimal("25.00"),
                    "USD",
                    "=G2*H2",
                ],
                # Row 3: Broken Excel formula =#REF!*H3
                [
                    "PO-02",
                    "PO-02-L01",
                    "2026-03-02",
                    "SUP-01",
                    "SKU-A-01",
                    "Precision Sensor",
                    5,
                    Decimal("25.00"),
                    "USD",
                    "=#REF!*H3",
                ],
            ],
        )
        _save_sheet(
            "shipments.xlsx",
            REQUIRED_COLUMNS[DatasetKind.SHIPMENTS],
            [["SHP-01", "PO-01", "PO-01-L01", "2026-03-04", "DHL", "TRK1", "SKU-A-01", 10, "WH-1"]],
        )
        _save_sheet(
            "supplier_invoices.xlsx",
            REQUIRED_COLUMNS[DatasetKind.INVOICES],
            [
                [
                    "INV-01",
                    "INV-01-01",
                    "PO-01",
                    "PO-01-L01",
                    "2026-03-06",
                    "SUP-01",
                    "LEG-A01",
                    10,
                    Decimal("25.00"),
                    "=H2*I2",
                    "USD",
                ]
            ],
        )

        result = run_reconciliation_from_directory(in_dir)
        rec_clean, rec_broken = result.reconciled_records

        assert rec_clean.reconciliation_status == ReconciliationStatus.RECONCILED
        assert rec_clean.preserved_po_formula == "=G2*H2"
        assert rec_clean.expected_amount == Decimal("250.00")
        assert rec_clean.actual_invoiced_amount == Decimal("250.00")

        assert rec_broken.reconciliation_status == ReconciliationStatus.DATA_VALIDATION_ERROR
        assert ReasonCode.BROKEN_EXCEL_FORMULA in rec_broken.reason_codes

        report_path = tmp_path / "out" / "reconciliation_report.xlsx"
        write_reconciliation_report(result, report_path)
        assert report_path.exists()

        val_report = validate_report_workbook(report_path)
        assert val_report.is_valid, f"Unexpected validation errors: {val_report.errors}"
        assert val_report.worksheets_found == REQUIRED_WORKSHEETS
        assert val_report.formula_count > 0
        assert len(val_report.broken_formula_cells) == 0

        # Confirm preserved formula cell AA2 is stored as string ('s'), not evaluated formula ('f')
        wb_check = openpyxl.load_workbook(report_path, data_only=False)
        try:
            ws_rec = wb_check["Reconciled Records"]
            assert ws_rec.freeze_panes == "A2"
            assert ws_rec.auto_filter.ref is not None
            aa2 = ws_rec["AA2"]
            assert aa2.value == "=G2*H2"
            assert aa2.data_type == "s"
        finally:
            wb_check.close()

    def test_validator_detects_tampered_broken_formula_in_report(self, tmp_path: Path) -> None:
        bad_report = tmp_path / "broken_report.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Summary"
        ws["B5"] = "=#REF!+10"
        wb.create_sheet("Reconciled Records")
        wb.create_sheet("Exceptions")
        wb.create_sheet("Audit Log")
        wb.save(bad_report)
        wb.close()

        val = validate_report_workbook(bad_report)
        assert not val.is_valid
        assert any("#REF!" in cell_desc for cell_desc in val.broken_formula_cells)

    def test_validator_detects_summary_monetary_variance_mismatch(self, tmp_path: Path) -> None:
        in_dir = tmp_path / "in"
        in_dir.mkdir()

        def _write_xlsx(path: Path, cols: list[str], rows: list[list[object]]) -> None:
            wb_in = openpyxl.Workbook()
            ws_in = wb_in.active
            assert ws_in is not None
            ws_in.append(cols)
            for r in rows:
                ws_in.append(r)
            wb_in.save(path)
            wb_in.close()

        _write_xlsx(
            in_dir / "product_master.xlsx",
            list(REQUIRED_COLUMNS[DatasetKind.PRODUCTS]),
            [["P-01", "SKU-001", "Sensor", "Electronics", "EA", "10.00", "SUP-01", "", "Y"]],
        )
        _write_xlsx(
            in_dir / "purchase_orders.xlsx",
            list(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS]),
            [
                [
                    "PO-01",
                    "PO-01-L01",
                    "2026-03-01",
                    "SUP-01",
                    "SKU-001",
                    "Sensor",
                    10,
                    "20.00",
                    "USD",
                    200,
                ]
            ],
        )
        _write_xlsx(
            in_dir / "shipments.xlsx",
            list(REQUIRED_COLUMNS[DatasetKind.SHIPMENTS]),
            [["SHP-01", "PO-01", "PO-01-L01", "2026-03-03", "DHL", "TRK-1", "SKU-001", 4, "WH-1"]],
        )
        _write_xlsx(
            in_dir / "supplier_invoices.xlsx",
            list(REQUIRED_COLUMNS[DatasetKind.INVOICES]),
            [
                [
                    "INV-01",
                    "INV-01-1",
                    "PO-01",
                    "PO-01-L01",
                    "2026-03-04",
                    "SUP-01",
                    "SKU-001",
                    4,
                    "20.00",
                    80,
                    "USD",
                ]
            ],
        )

        result = run_reconciliation_from_directory(in_dir)
        report_path = tmp_path / "out" / "reconciliation_report.xlsx"
        write_reconciliation_report(result, report_path)
        assert validate_report_workbook(report_path).is_valid

        # Tamper with Summary Net Billed Monetary Variance (B19) and verify validator rejects it
        wb = openpyxl.load_workbook(report_path)
        wb["Summary"]["B19"] = 999.99
        wb.save(report_path)
        wb.close()

        tampered_val = validate_report_workbook(report_path)
        assert not tampered_val.is_valid
        assert any("Net Billed Monetary Variance" in err for err in tampered_val.errors)

    def test_formula_injection_defense_and_strict_validator_checks(self, tmp_path: Path) -> None:
        import xml.etree.ElementTree as ET
        import zipfile

        in_dir = tmp_path / "sec_in"
        in_dir.mkdir()

        def _save_sheet(fname: str, cols: tuple[str, ...], rows: list[list[object]]) -> None:
            wb_in = openpyxl.Workbook()
            ws_in = wb_in.active
            assert ws_in is not None
            ws_in.append(list(cols))
            for r in rows:
                ws_in.append(r)
            wb_in.save(in_dir / fname)
            wb_in.close()

        _save_sheet(
            "product_master.xlsx",
            REQUIRED_COLUMNS[DatasetKind.PRODUCTS],
            [
                [
                    "PRD-01",
                    "SKU-SAFE-01",
                    "Safe Widget",
                    "Industrial",
                    "EA",
                    "10.00",
                    "SUP-01",
                    "",
                    "Y",
                ]
            ],
        )

        # Store literal text payloads as explicit strings and one real unpermitted formula
        wb_po = openpyxl.Workbook()
        ws_po = wb_po.active
        assert ws_po is not None
        ws_po.append(list(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS]))

        payloads = [
            "=1+1",
            "+cmd|'/C calc'!A0",
            "-2+3",
            "@SUM(1,2)",
            "\t=1",
            ' =HYPERLINK("http://evil.example","bad")',
            "=1+",
        ]
        for idx, desc_payload in enumerate(payloads, start=2):
            po_num = f"PO-SEC-{idx}"
            lid = f"{po_num}-1"
            ws_po.append(
                [
                    po_num,
                    lid,
                    "2026-03-01",
                    "SUP-01",
                    "SKU-SAFE-01",
                    None,
                    10,
                    Decimal("20.00"),
                    "USD",
                    f"=G{idx}*H{idx}",
                ]
            )
            c = ws_po.cell(row=idx, column=6)
            c.value = desc_payload
            c.data_type = "s"

        # Add one genuine unpermitted Excel formula row (=HYPERLINK(...))
        bad_row_idx = len(payloads) + 2
        ws_po.append(
            [
                "PO-BAD-FORM",
                "PO-BAD-FORM-1",
                "2026-03-01",
                "SUP-01",
                "SKU-SAFE-01",
                "Widget",
                10,
                Decimal("20.00"),
                "USD",
                '=HYPERLINK("https://evil.example","Click")',
            ]
        )
        wb_po.save(in_dir / "purchase_orders.xlsx")
        wb_po.close()

        _save_sheet(
            "shipments.xlsx",
            REQUIRED_COLUMNS[DatasetKind.SHIPMENTS],
            [
                [
                    f"SHP-{idx}",
                    f"PO-SEC-{idx}",
                    f"PO-SEC-{idx}-1",
                    "2026-03-02",
                    "DHL",
                    f"T{idx}",
                    "SKU-SAFE-01",
                    10,
                    "WH1",
                ]
                for idx in range(2, len(payloads) + 2)
            ],
        )
        _save_sheet(
            "supplier_invoices.xlsx",
            REQUIRED_COLUMNS[DatasetKind.INVOICES],
            [
                [
                    f"INV-{idx}",
                    f"INV-{idx}-1",
                    f"PO-SEC-{idx}",
                    f"PO-SEC-{idx}-1",
                    "2026-03-03",
                    "SUP-01",
                    "SKU-SAFE-01",
                    10,
                    Decimal("20.00"),
                    200,
                    "USD",
                ]
                for idx in range(2, len(payloads) + 2)
            ],
        )

        # Preserve dangerous text prefixes cleanly (leading whitespace stripped by normalize_string)
        result = run_reconciliation_from_directory(in_dir)
        expected_descs = [p.strip() for p in payloads]
        assert [
            r.product_description for r in result.reconciled_records[: len(payloads)]
        ] == expected_descs
        bad_rec = result.reconciled_records[bad_row_idx - 2]
        assert bad_rec.reconciliation_status == ReconciliationStatus.DATA_VALIDATION_ERROR
        assert ReasonCode.UNPERMITTED_EXCEL_FORMULA in bad_rec.reason_codes

        report_path = tmp_path / "sec_out" / "reconciliation_report.xlsx"
        write_reconciliation_report(result, report_path)
        assert validate_report_workbook(report_path).is_valid

        wb_out = openpyxl.load_workbook(report_path, data_only=False)
        try:
            ws_rec = wb_out["Reconciled Records"]
            for offset, expected_txt in enumerate(expected_descs, start=2):
                cell = ws_rec.cell(row=offset, column=11)
                assert cell.value == expected_txt
                assert cell.data_type == "s"
                assert cell.quotePrefix is True

            # Verify Summary uses pure cross-sheet references (zero hardcoded quantities/dollars)
            ws_sum = wb_out["Summary"]
            for r_idx in range(5, 23):
                c_val = ws_sum.cell(row=r_idx, column=3)
                assert c_val.data_type == "f"
                assert isinstance(c_val.value, str) and c_val.value.startswith("=")
        finally:
            wb_out.close()

        # Independent raw OpenXML check via zipfile + ET: zero <f> in data sheets, quotePrefix set
        with zipfile.ZipFile(report_path, "r") as zf:
            for data_xml in (
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
                "xl/worksheets/sheet4.xml",
            ):
                root = ET.fromstring(zf.read(data_xml))
                f_nodes = [el for el in root.iter() if el.tag.endswith("}f") or el.tag == "f"]
                assert f_nodes == []
            styles_xml = zf.read("xl/styles.xml").decode("utf-8")
            assert 'quotePrefix="1"' in styles_xml

        # Validator rejects missing Col B KPI, injected formula on data tab, and stale XML <v> cache
        blank_kpi_path = tmp_path / "blank_kpi.xlsx"
        blank_kpi_path.write_bytes(report_path.read_bytes())
        wb_b = openpyxl.load_workbook(blank_kpi_path)
        wb_b["Summary"]["B5"] = None
        wb_b.save(blank_kpi_path)
        wb_b.close()
        assert not validate_report_workbook(blank_kpi_path).is_valid

        injected_path = tmp_path / "injected.xlsx"
        injected_path.write_bytes(report_path.read_bytes())
        wb_inj = openpyxl.load_workbook(injected_path)
        wb_inj["Reconciled Records"]["AA2"] = "=SUM(M2:M3)"
        wb_inj.save(injected_path)
        wb_inj.close()
        assert not validate_report_workbook(injected_path).is_valid

        tampered_cache_path = tmp_path / "stale_cache.xlsx"
        with (
            zipfile.ZipFile(report_path, "r") as zin,
            zipfile.ZipFile(tampered_cache_path, "w", compression=zipfile.ZIP_DEFLATED) as zout,
        ):
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    root = ET.fromstring(data)
                    for el in root.iter():
                        if (el.tag.endswith("}c") or el.tag == "c") and el.attrib.get("r") == "C5":
                            for child in el:
                                if child.tag.endswith("}v") or child.tag == "v":
                                    child.text = "99999"
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                zout.writestr(item, data)
        val_cache = validate_report_workbook(tampered_cache_path)
        assert not val_cache.is_valid
        assert any("<v> cache" in err for err in val_cache.errors)

    def test_oversized_formula_ranges_rejected_fast_and_limits_configurable(
        self, tmp_path: Path
    ) -> None:
        # 1. Input ingestion rejects huge worksheet ranges quickly without hanging
        oversized_input_path = tmp_path / "oversized_po.xlsx"
        wb_in = openpyxl.Workbook()
        ws_in = wb_in.active
        assert ws_in is not None
        ws_in.append(list(REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS]))
        ws_in.append(
            [
                "PO-BIG-1",
                "PO-BIG-1-L01",
                "2026-03-01",
                "SUP-01",
                "SKU-001",
                "Widget",
                10,
                Decimal("20.00"),
                "USD",
                "=SUM(A1:XFD1048576)",
            ]
        )
        wb_in.save(oversized_input_path)
        wb_in.close()

        t0 = time.perf_counter()
        loaded = read_excel_dataset(oversized_input_path, DatasetKind.PURCHASE_ORDERS)
        elapsed_in = time.perf_counter() - t0
        assert elapsed_in < 1.0
        row_out = loaded.dataframe.to_dicts()[0]
        assert row_out["_has_broken_formula"] is True
        assert row_out["line_total"] == "#VALUE!"

        # Direct evaluate_worksheet_formula normal vs custom limit
        small_grid = {(1, 1): 10, (2, 1): 15, (3, 1): 20}
        assert evaluate_worksheet_formula("=SUM(A1:A3)", small_grid) == "45"
        assert evaluate_worksheet_formula("=SUM(A1:XFD1048576)", small_grid) == "#VALUE!"
        assert (
            evaluate_worksheet_formula("=SUM(A1:A10)", small_grid, max_range_cells=5) == "#VALUE!"
        )

        # 2. Report evaluator and validate_report_workbook reject =SUM(A1:XFD1048576) quickly
        wb = openpyxl.Workbook()
        ws_sum = wb.active
        assert ws_sum is not None
        ws_sum.title = "Summary"
        ws_data = wb.create_sheet("Reconciled Records")
        for idx in range(1, 21):
            ws_data.cell(row=idx, column=1, value=idx)
            ws_data.cell(row=idx, column=2, value="RECONCILED")

        # Normal bounded formulas work
        cache = WorkbookFormulaCache(wb)
        sumifs_formula = (
            "=SUMIFS('Reconciled Records'!$A$1:$A$20,"
            "'Reconciled Records'!$B$1:$B$20,\"RECONCILED\")"
        )
        assert evaluate_workbook_formula(
            wb, "Summary", "=SUM('Reconciled Records'!$A$1:$A$20)", cache=cache
        ) == Decimal("210")
        assert evaluate_workbook_formula(
            wb,
            "Summary",
            sumifs_formula,
            cache=cache,
        ) == Decimal("210")

        # Oversized range fails in milliseconds with FormulaResourceLimitError
        t1 = time.perf_counter()
        with pytest.raises(FormulaResourceLimitError, match="max_range_cells"):
            evaluate_workbook_formula(
                wb, "Summary", "=SUM('Reconciled Records'!A1:XFD1048576)", cache=cache
            )
        assert time.perf_counter() - t1 < 0.2

        # Multi-range SUMIFS exceeding max_evaluated_cells fails cleanly
        tight_limits = FormulaEvaluationLimits(max_range_cells=25, max_evaluated_cells=30)
        with pytest.raises(FormulaResourceLimitError, match="max_evaluated_cells"):
            evaluate_workbook_formula(
                wb,
                "Summary",
                sumifs_formula,
                limits=tight_limits,
                cache=cache,
            )
        wb.close()

        # Tamper a valid minimal report with =SUM(A1:XFD1048576) and check validate_report_workbook
        in_dir = tmp_path / "min_in"
        in_dir.mkdir()
        for fname, cols, data_rows in (
            (
                "product_master.xlsx",
                REQUIRED_COLUMNS[DatasetKind.PRODUCTS],
                [["P-1", "SKU-1", "Item", "Cat", "EA", "10.00", "SUP-1", "", "Y"]],
            ),
            (
                "purchase_orders.xlsx",
                REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS],
                [
                    [
                        "PO-1",
                        "PO-1-L1",
                        "2026-03-01",
                        "SUP-1",
                        "SKU-1",
                        "Item",
                        5,
                        "10.00",
                        "USD",
                        50,
                    ]
                ],
            ),
            (
                "shipments.xlsx",
                REQUIRED_COLUMNS[DatasetKind.SHIPMENTS],
                [["SHP-1", "PO-1", "PO-1-L1", "2026-03-02", "DHL", "T1", "SKU-1", 5, "WH-1"]],
            ),
            (
                "supplier_invoices.xlsx",
                REQUIRED_COLUMNS[DatasetKind.INVOICES],
                [
                    [
                        "INV-1",
                        "INV-1-1",
                        "PO-1",
                        "PO-1-L1",
                        "2026-03-03",
                        "SUP-1",
                        "SKU-1",
                        5,
                        "10.00",
                        50,
                        "USD",
                    ]
                ],
            ),
        ):
            w = openpyxl.Workbook()
            s = w.active
            assert s is not None
            s.append(list(cols))
            for dr in data_rows:
                s.append(dr)
            w.save(in_dir / fname)
            w.close()

        rep_path = tmp_path / "rep.xlsx"
        write_reconciliation_report(run_reconciliation_from_directory(in_dir), rep_path)

        wb_bad = openpyxl.load_workbook(rep_path)
        wb_bad["Summary"]["C14"] = "=SUM('Reconciled Records'!A1:XFD1048576)"
        wb_bad.save(rep_path)
        wb_bad.close()

        t2 = time.perf_counter()
        val_bad = validate_report_workbook(rep_path)
        assert time.perf_counter() - t2 < 1.0
        assert not val_bad.is_valid
        assert any("max_range_cells" in err for err in val_bad.errors)
