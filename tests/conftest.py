"""Shared pytest fixtures for Excel Reconciliation Automation tests."""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any

import openpyxl
import polars as pl
import pytest

from excel_reconciliation.models import REQUIRED_COLUMNS, DatasetKind


@pytest.fixture
def sample_products_df() -> pl.DataFrame:
    """Small deterministic Product Master DataFrame including clean, alias, and ambiguous items."""
    rows: list[dict[str, Any]] = [
        {
            "product_id": "PRD-001",
            "canonical_sku": "SKU-ELEC-001",
            "product_name": "USB-C Dock Dual 4K",
            "category": "Electronics",
            "uom": "EA",
            "standard_cost": "50.00",
            "supplier_id": "SUP-01",
            "aliases": "LEG-001; SUP01-DOCK",
            "active_flag": "Y",
        },
        {
            "product_id": "PRD-002",
            "canonical_sku": "SKU-HOME-002",
            "product_name": "Précision Café Grinder — Edelstahl",
            "category": "Home & Kitchen",
            "uom": "EA",
            "standard_cost": "120.00",
            "supplier_id": "SUP-02",
            "aliases": "LEG-002; SHARED-AMB-99",
            "active_flag": "Y",
        },
        {
            "product_id": "PRD-003",
            "canonical_sku": "SKU-IND-003",
            "product_name": "Pneumatic Valve 24V",
            "category": "Industrial",
            "uom": "EA",
            "standard_cost": "35.50",
            "supplier_id": "SUP-03",
            "aliases": "LEG-003; SHARED-AMB-99",
            "active_flag": "Y",
        },
    ]
    return pl.DataFrame(
        rows, schema=dict.fromkeys(REQUIRED_COLUMNS[DatasetKind.PRODUCTS], pl.String)
    )


@pytest.fixture
def write_xlsx_helper(tmp_path: Path):
    """Helper factory to write an arbitrary worksheet to a temporary .xlsx file."""

    def _writer(
        filename: str,
        headers: Sequence[str] | tuple[str, ...],
        rows: list[list[Any]],
        *,
        sheet_name: str = "Sheet1",
    ) -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = sheet_name
        ws.append(list(headers))
        for row in rows:
            ws.append(row)
        out_path = tmp_path / filename
        wb.save(out_path)
        wb.close()
        return out_path

    return _writer
