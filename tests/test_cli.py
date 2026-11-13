"""CLI integration tests for generate-demo, reconcile, and validate commands."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_reconciliation.cli import main


class TestCommandLineInterface:
    """Verify end-to-end CLI subcommands and error exit codes."""

    def test_cli_generate_reconcile_and_validate_end_to_end(self, tmp_path: Path) -> None:
        in_dir = tmp_path / "demo_in"
        out_dir = tmp_path / "demo_out"
        summary_json = out_dir / "summary.json"
        report_file = out_dir / "reconciliation_report.xlsx"

        rc_gen = main(["--quiet", "generate-demo", "--output", str(in_dir), "--seed", "42"])
        assert rc_gen == 0
        assert (in_dir / "purchase_orders.xlsx").exists()

        rc_rec = main(
            [
                "--quiet",
                "reconcile",
                "--input",
                str(in_dir),
                "--output",
                str(out_dir),
                "--summary-json",
                str(summary_json),
            ]
        )
        assert rc_rec == 0
        assert report_file.exists()
        assert summary_json.exists()

        summary_payload = json.loads(summary_json.read_text(encoding="utf-8"))
        assert summary_payload["total_po_rows"] == 2000

        rc_val = main(["--quiet", "validate", "--report", str(report_file)])
        assert rc_val == 0

    def test_cli_missing_input_directory_returns_nonzero(self, tmp_path: Path) -> None:
        rc = main(
            [
                "--quiet",
                "reconcile",
                "--input",
                str(tmp_path / "does_not_exist"),
                "--output",
                str(tmp_path / "out"),
            ]
        )
        assert rc == 1

    def test_cli_rejects_negative_tolerance(self) -> None:
        with pytest.raises(SystemExit):
            main(["reconcile", "--qty-tolerance", "-1"])

    def test_cli_missing_required_column_exits_nonzero(
        self, tmp_path: Path, write_xlsx_helper, capsys: pytest.CaptureFixture[str]
    ) -> None:
        import openpyxl

        in_dir = tmp_path / "bad_schema_in"
        in_dir.mkdir()
        main(["--quiet", "generate-demo", "--output", str(in_dir), "--seed", "7"])

        # Remove a required column ('currency') from purchase_orders.xlsx
        po_path = in_dir / "purchase_orders.xlsx"
        wb = openpyxl.load_workbook(po_path)
        ws = wb.active
        assert ws is not None
        ws.delete_cols(9, 1)
        wb.save(po_path)
        wb.close()

        rc = main(["reconcile", "--input", str(in_dir), "--output", str(tmp_path / "out")])
        captured = capsys.readouterr()
        assert rc == 1
        assert "ERROR:" in captured.err
