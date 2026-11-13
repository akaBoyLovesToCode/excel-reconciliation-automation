"""Command-line interface for Excel Reconciliation Automation."""

from __future__ import annotations

import argparse
import importlib.util
import json
import logging
import os
import sys
import tempfile
from collections.abc import Sequence
from dataclasses import asdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

from excel_reconciliation import __version__
from excel_reconciliation.models import ReconciliationConfig, SchemaValidationError

LOGGER = logging.getLogger("excel_reconciliation")


def _decimal_arg(value: str) -> Decimal:
    """Parse a non-negative Decimal CLI tolerance argument."""
    try:
        parsed = Decimal(value)
    except InvalidOperation as exc:
        raise argparse.ArgumentTypeError(f"Invalid decimal value: {value!r}") from exc
    if not parsed.is_finite() or parsed < Decimal("0"):
        raise argparse.ArgumentTypeError(f"Tolerance must be finite and >= 0, got {value!r}")
    return parsed


def _check_dependencies(required_modules: tuple[str, ...]) -> list[str]:
    """Return any required third-party modules missing from the active Python interpreter."""
    missing: list[str] = []
    for mod in required_modules:
        if sys.modules.get(mod) is not None:
            continue
        try:
            if importlib.util.find_spec(mod) is None:
                missing.append(mod)
        except (ModuleNotFoundError, ValueError):
            missing.append(mod)
    return missing


def build_parser() -> argparse.ArgumentParser:
    """Construct the top-level CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="python -m excel_reconciliation",
        description=(
            "Excel Reconciliation Automation: clean messy workbooks, match SKUs and "
            "order lines deterministically, and generate auditable 3-way Excel reports."
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"excel-reconciliation-automation {__version__}",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress informational progress logs and console summary tables.",
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Subcommand: generate-demo
    gen_parser = subparsers.add_parser(
        "generate-demo",
        help="Generate deterministic synthetic Excel workbooks and ground-truth manifest.",
    )
    gen_parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("examples/input"),
        help="Target directory for synthetic Excel datasets (default: examples/input).",
    )
    gen_parser.add_argument(
        "--seed",
        type=int,
        default=20260401,
        help="Deterministic PRNG seed (default: 20260401).",
    )

    # Subcommand: reconcile
    rec_parser = subparsers.add_parser(
        "reconcile",
        help="Clean, match, and reconcile Excel workbooks into reconciliation_report.xlsx.",
    )
    rec_parser.add_argument(
        "--input",
        "-i",
        type=Path,
        default=Path("examples/input"),
        help="Directory containing input Excel workbooks (default: examples/input).",
    )
    rec_parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("examples/output"),
        help=(
            "Directory where reconciliation_report.xlsx will be written (default: examples/output)."
        ),
    )
    rec_parser.add_argument(
        "--report-name",
        type=str,
        default="reconciliation_report.xlsx",
        help="Output Excel filename (default: reconciliation_report.xlsx).",
    )
    rec_parser.add_argument(
        "--qty-tolerance",
        type=_decimal_arg,
        default=Decimal("0"),
        help="Permitted absolute quantity difference (default: 0).",
    )
    rec_parser.add_argument(
        "--price-tolerance",
        type=_decimal_arg,
        default=Decimal("0.01"),
        help="Permitted unit price difference in USD (default: 0.01).",
    )
    rec_parser.add_argument(
        "--amount-tolerance",
        type=_decimal_arg,
        default=Decimal("0.02"),
        help="Permitted extended line-amount difference in USD (default: 0.02).",
    )
    rec_parser.add_argument(
        "--summary-json",
        type=Path,
        default=None,
        help="Optional path to export machine-readable summary KPIs as JSON.",
    )

    # Subcommand: validate
    val_parser = subparsers.add_parser(
        "validate",
        help="Validate worksheet structure, formula integrity, and KPI consistency of a report.",
    )
    val_parser.add_argument(
        "--report",
        "-r",
        type=Path,
        default=Path("examples/output/reconciliation_report.xlsx"),
        help="Path to reconciliation_report.xlsx to validate.",
    )

    return parser


def _print_reconciliation_summary(report_path: Path, summary_obj: object) -> None:
    """Format a concise console summary of reconciliation outcomes."""
    from excel_reconciliation.models import SummaryMetrics

    assert isinstance(summary_obj, SummaryMetrics)
    s = summary_obj
    cov_pct = float(s.matching_coverage_rate) * 100.0
    rec_pct = float(s.reconciliation_rate) * 100.0

    lines = [
        "",
        "=== Excel Reconciliation Summary ===",
        f"Report written to           : {report_path}",
        f"Total PO rows ingested      : {s.total_po_rows:,}",
        f"Unique valid PO lines       : {s.unique_valid_po_lines:,}",
        f"Product matching coverage   : {cov_pct:.2f}%",
        f"Fully reconciled PO lines   : {s.reconciled_lines:,} ({rec_pct:.2f}%)",
        f"Partial fulfillment lines   : {s.partial_fulfillment_lines:,}",
        f"PO lines requiring review   : {s.po_exception_lines:,}",
        f"Total exceptions (all tabs) : {s.total_exception_records:,}",
        f"Outstanding unshipped qty   : {s.total_outstanding_qty:,}",
        f"Expected PO commitment      : ${s.total_expected_amount:,.2f}",
        f"Actual invoiced total       : ${s.total_invoiced_amount:,.2f}",
        f"Absolute discrepancy        : ${s.absolute_monetary_discrepancy:,.2f}",
        f"Unbilled shipped exposure   : ${s.unbilled_shipped_amount:,.2f}",
        f"Quarantined invoice amount  : ${s.quarantined_invoice_amount:,.2f}",
        "====================================",
    ]
    print("\n".join(lines))


def main(argv: Sequence[str] | None = None) -> int:
    """CLI entry point returning POSIX exit code."""
    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if args.command == "generate-demo":
        missing = _check_dependencies(("openpyxl",))
        if missing:
            print(
                f"ERROR: Missing required Python package(s): {', '.join(missing)}. "
                "Ensure dependencies declared in pyproject.toml are available to Python.",
                file=sys.stderr,
            )
            return 2

        from excel_reconciliation.demo_data import generate_demo_datasets

        written = generate_demo_datasets(args.output, seed=args.seed)
        if not args.quiet:
            print(f"Generated synthetic Excel datasets (seed={args.seed}) in {args.output}:")
            for name, fpath in written.items():
                print(f"  - {name}: {fpath}")
        return 0

    if args.command == "reconcile":
        missing = _check_dependencies(("polars", "openpyxl"))
        if missing:
            print(
                f"ERROR: Missing required Python package(s): {', '.join(missing)}. "
                "Ensure dependencies declared in pyproject.toml are available to Python.",
                file=sys.stderr,
            )
            return 2

        from excel_reconciliation.reconciliation import run_reconciliation_from_directory
        from excel_reconciliation.reporting import write_reconciliation_report

        config = ReconciliationConfig(
            quantity_tolerance=args.qty_tolerance,
            unit_price_tolerance=args.price_tolerance,
            amount_tolerance=args.amount_tolerance,
        )
        try:
            result = run_reconciliation_from_directory(args.input, config=config)
        except (FileNotFoundError, SchemaValidationError) as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1

        report_path = args.output / args.report_name
        write_reconciliation_report(result, report_path)

        if args.summary_json is not None:
            args.summary_json.parent.mkdir(parents=True, exist_ok=True)
            serializable = {
                k: str(v) if isinstance(v, Decimal) else v
                for k, v in asdict(result.summary).items()
            }
            payload = json.dumps(serializable, indent=2, sort_keys=True) + "\n"
            fd, tmp_name = tempfile.mkstemp(
                dir=args.summary_json.parent,
                prefix=f".{args.summary_json.stem}.",
                suffix=".tmp.json",
            )
            os.close(fd)
            tmp_path = Path(tmp_name)
            try:
                tmp_path.write_text(payload, encoding="utf-8")
                os.replace(tmp_path, args.summary_json)
            finally:
                if tmp_path.exists():
                    tmp_path.unlink()

        if not args.quiet:
            _print_reconciliation_summary(report_path, result.summary)
        return 0

    if args.command == "validate":
        missing = _check_dependencies(("openpyxl",))
        if missing:
            print(
                f"ERROR: Missing required Python package(s): {', '.join(missing)}. "
                "Ensure dependencies declared in pyproject.toml are available to Python.",
                file=sys.stderr,
            )
            return 2

        from excel_reconciliation.validation import validate_report_workbook

        val_res = validate_report_workbook(args.report)
        if not args.quiet:
            status_word = "PASSED" if val_res.is_valid else "FAILED"
            print(f"Report validation {status_word}: {val_res.report_path}")
            print(f"  Worksheets : {', '.join(val_res.worksheets_found)}")
            print(f"  Formulas   : {val_res.formula_count} verified")
            for warn in val_res.warnings:
                print(f"  WARNING    : {warn}")
        if not val_res.is_valid:
            for err in val_res.errors:
                print(f"  ERROR      : {err}", file=sys.stderr)
            return 1
        return 0

    return 1
