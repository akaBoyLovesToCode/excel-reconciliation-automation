"""Three-way financial reconciliation engine across Purchase Orders, Shipments, and Invoices."""

from __future__ import annotations

import datetime as dt
import json
import platform
from collections import Counter
from dataclasses import dataclass, replace
from decimal import Decimal, localcontext
from pathlib import Path
from typing import Any

import openpyxl
import polars as pl

from excel_reconciliation.cleaning import (
    CleanedDataset,
    clean_invoices_df,
    clean_products_df,
    clean_purchase_orders_df,
    clean_shipments_df,
)
from excel_reconciliation.excel_io import LoadedWorkbookData, read_excel_dataset
from excel_reconciliation.matching import (
    MATCHING_RULES_DESCRIPTIONS,
    OrderLineMatcher,
    ProductMatcher,
    deduplicate_matched_invoices,
    deduplicate_matched_shipments,
    match_purchase_orders_df,
)
from excel_reconciliation.models import (
    ARITHMETIC_CONTEXT,
    AuditLogBundle,
    DatasetKind,
    ExceptionCategory,
    ExceptionRecord,
    MatchStatus,
    PoLineKey,
    ReasonCode,
    ReconciledRecord,
    ReconciliationConfig,
    ReconciliationResult,
    ReconciliationStatus,
    SchemaValidationError,
    Severity,
    SummaryMetrics,
    quantize_rate,
)


@dataclass(frozen=True, slots=True)
class MatchedShipmentAggregate:
    """Pre-aggregated valid shipment summary for a single PO line."""

    shipment_count: int
    shipped_qty: Decimal
    shipment_refs: str
    used_alias: bool


@dataclass(frozen=True, slots=True)
class MatchedInvoiceLine:
    """Single validated, deduplicated invoice line with original source provenance."""

    source_row: int
    raw_invoice_id: str
    raw_invoice_line_id: str
    norm_invoice_id: str
    norm_invoice_line_id: str
    raw_sku: str
    norm_sku: str
    invoiced_qty: Decimal
    unit_price: Decimal
    line_amount: Decimal
    has_internal_math_err: bool
    used_alias: bool


@dataclass(frozen=True, slots=True)
class MatchedInvoiceAggregate:
    """Pre-aggregated valid supplier invoice lines for a single PO line."""

    invoice_count: int
    invoiced_qty: Decimal
    invoiced_amount: Decimal
    unit_prices: tuple[Decimal, ...]
    invoice_refs: str
    has_internal_math_err: bool
    used_alias: bool
    lines: tuple[MatchedInvoiceLine, ...]


@dataclass(frozen=True, slots=True)
class LineFinancials:
    """Independent financial metrics computed identically for every eligible PO line."""

    ordered_qty: Decimal
    shipped_qty: Decimal | None
    invoiced_qty: Decimal | None
    outstanding_qty: Decimal
    expected_amount: Decimal
    actual_invoiced_amount: Decimal | None
    authorized_supported_value: Decimal
    monetary_variance: Decimal | None
    price_discrepancy_exposure: Decimal
    quantity_overbilling_exposure: Decimal
    discrepancy_exposure: Decimal
    unbilled_receiving_amount: Decimal
    has_unit_price_discrepancy: bool
    has_invoice_math_discrepancy: bool


def _extract_formula(formulas_json: str | None, col_name: str) -> str | None:
    """Extract preserved Excel formula string for a specific column if present."""
    if not formulas_json:
        return None
    try:
        mapping = json.loads(formulas_json)
    except json.JSONDecodeError:
        return None
    val = mapping.get(col_name)
    return str(val) if val is not None else None


def _row_candidate_po_keys(row: dict[str, Any]) -> list[PoLineKey]:
    """Extract structured PoLineKey objects touched by an ambiguous or conflicting transaction."""
    keys: list[PoLineKey] = []
    seen: set[PoLineKey] = set()

    def _add(po_n: str | None, po_l: str | None) -> None:
        if po_n and po_l:
            k = PoLineKey(po_number=str(po_n), po_line_id=str(po_l))
            if k not in seen:
                seen.add(k)
                keys.append(k)

    _add(row.get("matched_po_number"), row.get("matched_po_line_id"))
    for cand in row.get("candidate_po_lines") or ():
        if isinstance(cand, dict):
            _add(cand.get("po_number"), cand.get("po_line_id"))
    if not keys:
        _add(row.get("norm_po_number"), row.get("norm_po_line_id"))
    return keys


def _aggregate_matched_shipments(
    matched_shipments_df: pl.DataFrame,
) -> dict[PoLineKey, MatchedShipmentAggregate]:
    """Pre-aggregate valid, non-duplicate shipments per structured PoLineKey.

    Independent group_by prior to joining prevents M x N join-cardinality inflation.
    """
    if matched_shipments_df.is_empty():
        return {}

    clean_matched = matched_shipments_df.filter(
        pl.col("_is_valid_record")
        & ~pl.col("_is_duplicate")
        & ~pl.col("_is_conflicting_duplicate")
        & pl.col("matched_po_number").is_not_null()
        & pl.col("matched_po_line_id").is_not_null()
    )
    if clean_matched.is_empty():
        return {}

    grouped = clean_matched.group_by(
        ["matched_po_number", "matched_po_line_id"], maintain_order=True
    ).agg(
        pl.len().alias("shipment_count"),
        pl.col("norm_shipment_id").alias("shipment_ids"),
        pl.col("norm_shipped_qty").alias("shipped_qtys"),
        pl.col("match_status").alias("match_statuses"),
    )

    agg_map: dict[PoLineKey, MatchedShipmentAggregate] = {}
    for row in grouped.to_dicts():
        key = PoLineKey(
            po_number=str(row["matched_po_number"]),
            po_line_id=str(row["matched_po_line_id"]),
        )
        qty_sum = sum((Decimal(q) for q in row["shipped_qtys"] if q is not None), Decimal("0"))
        refs = ", ".join(sid for sid in row["shipment_ids"] if sid is not None)
        agg_map[key] = MatchedShipmentAggregate(
            shipment_count=int(row["shipment_count"]),
            shipped_qty=qty_sum,
            shipment_refs=refs,
            used_alias=MatchStatus.ALIAS_MAP.value in (row["match_statuses"] or []),
        )
    return agg_map


def _aggregate_matched_invoices(
    matched_invoices_df: pl.DataFrame,
    config: ReconciliationConfig,
) -> dict[PoLineKey, MatchedInvoiceAggregate]:
    """Pre-aggregate valid, non-duplicate invoices per structured PoLineKey."""
    if matched_invoices_df.is_empty():
        return {}

    clean_matched = matched_invoices_df.filter(
        pl.col("_is_valid_record")
        & ~pl.col("_is_duplicate")
        & ~pl.col("_is_conflicting_duplicate")
        & pl.col("matched_po_number").is_not_null()
        & pl.col("matched_po_line_id").is_not_null()
    )
    if clean_matched.is_empty():
        return {}

    by_key_rows: dict[PoLineKey, list[dict[str, Any]]] = {}
    for row in clean_matched.to_dicts():
        key = PoLineKey(
            po_number=str(row["matched_po_number"]),
            po_line_id=str(row["matched_po_line_id"]),
        )
        by_key_rows.setdefault(key, []).append(row)

    agg_map: dict[PoLineKey, MatchedInvoiceAggregate] = {}
    for key, grp_rows in by_key_rows.items():
        grp_rows_sorted = sorted(grp_rows, key=lambda r: int(r["_source_row"]))
        inv_lines: list[MatchedInvoiceLine] = []
        any_math_err = False
        any_alias = False
        qty_sum = Decimal("0")
        amt_sum = Decimal("0.00")

        for r in grp_rows_sorted:
            q_val = Decimal(str(r["norm_invoiced_qty"]))
            p_val = Decimal(str(r["norm_unit_price"]))
            a_val = config.quantize_money(Decimal(str(r["norm_line_amount"])))
            expected_line_amt = config.quantize_money(q_val * p_val)
            line_math_err = abs(a_val - expected_line_amt) > config.amount_tolerance
            if line_math_err:
                any_math_err = True
            used_alias = r["match_status"] == MatchStatus.ALIAS_MAP.value
            if used_alias:
                any_alias = True

            qty_sum += q_val
            amt_sum += a_val
            inv_lines.append(
                MatchedInvoiceLine(
                    source_row=int(r["_source_row"]),
                    raw_invoice_id=str(r["raw_invoice_id"] or ""),
                    raw_invoice_line_id=str(r["raw_invoice_line_id"] or ""),
                    norm_invoice_id=str(r["norm_invoice_id"] or ""),
                    norm_invoice_line_id=str(r["norm_invoice_line_id"] or ""),
                    raw_sku=str(r["raw_sku"] or ""),
                    norm_sku=str(r["norm_sku"] or ""),
                    invoiced_qty=q_val,
                    unit_price=p_val,
                    line_amount=a_val,
                    has_internal_math_err=line_math_err,
                    used_alias=used_alias,
                )
            )

        refs = ", ".join(
            line.norm_invoice_line_id for line in inv_lines if line.norm_invoice_line_id
        )
        agg_map[key] = MatchedInvoiceAggregate(
            invoice_count=len(inv_lines),
            invoiced_qty=qty_sum,
            invoiced_amount=config.quantize_money(amt_sum),
            unit_prices=tuple(line.unit_price for line in inv_lines),
            invoice_refs=refs,
            has_internal_math_err=any_math_err,
            used_alias=any_alias,
            lines=tuple(inv_lines),
        )
    return agg_map


def compute_po_line_financials(
    *,
    ordered_qty: Decimal,
    po_unit_price: Decimal,
    po_line_total_val: Decimal | None,
    ship_agg: MatchedShipmentAggregate | None,
    inv_agg: MatchedInvoiceAggregate | None,
    config: ReconciliationConfig,
) -> LineFinancials:
    """Compute line-level financial metrics independently of primary exception classification.

    Formulas:
      - Q_auth = min(Q_ship, Q_po)
      - Walking matched invoices in source-row order with remaining authorized capacity R = Q_auth:
          q_supp_j = min(q_j, R);  R <- R - q_supp_j;  q_exc_j = q_j - q_supp_j
          a_supp_j = ROUND_HALF_UP(a_j * q_supp_j / q_j) if q_j > 0 else a_j
          a_exc_j  = ROUND_HALF_UP(a_j - a_supp_j)
          v_supp_j = ROUND_HALF_UP(q_supp_j * P_po)
          delta_price_j = a_supp_j - v_supp_j
          delta_qty_j   = a_exc_j
      - Net Billed Monetary Variance = A_inv - sum(v_supp_j)
      - Discrepancy Exposure = sum(|delta_price_j| + |delta_qty_j|) when price, math,
          or overbilling/premature-billing discrepancy exceeds configured tolerances ($0 otherwise)
      - Unbilled Receipt Accrual = ROUND_HALF_UP(R * P_po)
          = ROUND_HALF_UP(max(Q_auth - Q_inv, 0) * P_po)
    """
    ship_count = ship_agg.shipment_count if ship_agg else 0
    shipped_qty = ship_agg.shipped_qty if ship_agg else None
    effective_shipped = shipped_qty if shipped_qty is not None else Decimal("0")
    outstanding_qty = max(ordered_qty - effective_shipped, Decimal("0"))
    expected_amount = config.quantize_money(ordered_qty * po_unit_price)

    inv_count = inv_agg.invoice_count if inv_agg else 0
    invoiced_qty = inv_agg.invoiced_qty if inv_agg else None
    actual_invoiced_amount = inv_agg.invoiced_amount if inv_agg else None

    authorized_received_qty = min(effective_shipped, ordered_qty)
    remaining_authorized = authorized_received_qty

    sum_v_supp = Decimal("0.00")
    gross_price_exp = Decimal("0.00")
    gross_qty_exp = Decimal("0.00")

    if inv_agg is not None:
        with localcontext(ARITHMETIC_CONTEXT):
            for inv in inv_agg.lines:
                q_j = inv.invoiced_qty
                a_j = inv.line_amount
                q_supp = min(q_j, remaining_authorized)
                remaining_authorized -= q_supp
                q_exc = q_j - q_supp

                if q_j > Decimal("0") and q_exc > Decimal("0"):
                    a_supp = config.quantize_money((a_j * q_supp) / q_j)
                    a_exc = config.quantize_money(a_j - a_supp)
                elif q_exc > Decimal("0"):
                    a_supp = Decimal("0.00")
                    a_exc = a_j
                else:
                    a_supp = a_j
                    a_exc = Decimal("0.00")

                v_supp = config.quantize_money(q_supp * po_unit_price)
                delta_price = config.quantize_money(a_supp - v_supp)
                delta_qty = a_exc

                sum_v_supp += v_supp
                gross_price_exp += abs(delta_price)
                gross_qty_exp += abs(delta_qty)

    sum_v_supp = config.quantize_money(sum_v_supp)
    gross_price_exp = config.quantize_money(gross_price_exp)
    gross_qty_exp = config.quantize_money(gross_qty_exp)
    unbilled_receiving = config.quantize_money(
        max(remaining_authorized, Decimal("0")) * po_unit_price
    )

    monetary_variance: Decimal | None
    if inv_count > 0 and actual_invoiced_amount is not None:
        monetary_variance = config.quantize_money(actual_invoiced_amount - sum_v_supp)
    else:
        monetary_variance = None

    has_inv_math_err = bool(inv_agg and inv_agg.has_internal_math_err)
    has_unit_price_disc = False
    if inv_agg is not None and actual_invoiced_amount is not None and invoiced_qty is not None:
        expected_for_inv_qty = config.quantize_money(invoiced_qty * po_unit_price)
        has_unit_price_disc = any(
            abs(p - po_unit_price) > config.unit_price_tolerance for p in inv_agg.unit_prices
        ) or (abs(actual_invoiced_amount - expected_for_inv_qty) > config.amount_tolerance)

    po_math_off = (
        po_line_total_val is not None
        and abs(po_line_total_val - expected_amount) > config.amount_tolerance
    )

    premature_or_overbilled = False
    if (
        inv_count > 0
        and invoiced_qty is not None
        and (ship_count == 0 or invoiced_qty > authorized_received_qty + config.quantity_tolerance)
    ):
        premature_or_overbilled = True

    if inv_count == 0:
        discrepancy_exposure = Decimal("0.00")
    elif has_inv_math_err or has_unit_price_disc or premature_or_overbilled:
        discrepancy_exposure = config.quantize_money(gross_price_exp + gross_qty_exp)
    elif po_math_off:
        assert po_line_total_val is not None
        discrepancy_exposure = config.quantize_money(abs(po_line_total_val - expected_amount))
    else:
        discrepancy_exposure = Decimal("0.00")

    return LineFinancials(
        ordered_qty=ordered_qty,
        shipped_qty=shipped_qty,
        invoiced_qty=invoiced_qty,
        outstanding_qty=outstanding_qty,
        expected_amount=expected_amount,
        actual_invoiced_amount=actual_invoiced_amount,
        authorized_supported_value=sum_v_supp,
        monetary_variance=monetary_variance,
        price_discrepancy_exposure=gross_price_exp,
        quantity_overbilling_exposure=gross_qty_exp,
        discrepancy_exposure=discrepancy_exposure,
        unbilled_receiving_amount=unbilled_receiving,
        has_unit_price_discrepancy=has_unit_price_disc,
        has_invoice_math_discrepancy=has_inv_math_err,
    )


def _make_reconciled_record(
    *,
    source_row: int,
    raw_po: str,
    raw_line: str,
    raw_sku: str,
    norm_po: str,
    norm_line: str,
    norm_sku: str,
    matched_product_id: str | None,
    canonical_sku: str | None,
    matched_business_key: str,
    product_description: str,
    supplier_id: str,
    order_date: str | None,
    product_match_status: MatchStatus,
    ordered_qty: Decimal | None,
    shipped_qty: Decimal | None,
    invoiced_qty: Decimal | None,
    outstanding_qty: Decimal | None,
    po_unit_price: Decimal | None,
    expected_amount: Decimal | None,
    actual_invoiced_amount: Decimal | None,
    monetary_variance: Decimal | None,
    shipment_count: int,
    invoice_count: int,
    shipment_refs: str,
    invoice_refs: str,
    preserved_po_formula: str | None,
    reconciliation_status: ReconciliationStatus,
    reason_codes: tuple[ReasonCode, ...],
    is_valid_po_line: bool,
    discrepancy_exposure: Decimal = Decimal("0.00"),
    unbilled_receiving_amount: Decimal = Decimal("0.00"),
) -> ReconciledRecord:
    """Single constructor helper ensuring all ReconciledRecord invariants are populated."""
    return ReconciledRecord(
        source_row=source_row,
        raw_po_number=raw_po,
        raw_po_line_id=raw_line,
        raw_sku=raw_sku,
        normalized_po_number=norm_po,
        normalized_po_line_id=norm_line,
        normalized_sku=norm_sku,
        matched_product_id=matched_product_id,
        canonical_sku=canonical_sku,
        matched_business_key=matched_business_key,
        product_description=product_description,
        supplier_id=supplier_id,
        order_date=order_date,
        product_match_status=product_match_status,
        ordered_qty=ordered_qty,
        shipped_qty=shipped_qty,
        invoiced_qty=invoiced_qty,
        outstanding_qty=outstanding_qty,
        po_unit_price=po_unit_price,
        expected_amount=expected_amount,
        actual_invoiced_amount=actual_invoiced_amount,
        monetary_variance=monetary_variance,
        shipment_count=shipment_count,
        invoice_count=invoice_count,
        shipment_refs=shipment_refs,
        invoice_refs=invoice_refs,
        preserved_po_formula=preserved_po_formula,
        reconciliation_status=reconciliation_status,
        reason_codes=reason_codes,
        is_valid_po_line=is_valid_po_line,
        discrepancy_exposure=discrepancy_exposure,
        unbilled_receiving_amount=unbilled_receiving_amount,
    )


def _collect_product_exceptions(
    products_df: pl.DataFrame,
    exceptions: list[ExceptionRecord],
    next_exc_id: Any,
) -> None:
    """Collect Product Master duplicate, conflict, and validation exceptions."""
    has_conflict_col = "_is_conflicting_duplicate" in products_df.columns
    for row in products_df.to_dicts():
        is_conflict = bool(has_conflict_col and row["_is_conflicting_duplicate"])
        if is_conflict:
            v_codes = tuple(ReasonCode(c) for c in (row["_validation_codes"] or []))
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.AMBIGUOUS_REFERENCE,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.PRODUCTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_product_id"] or ""),
                    normalized_reference=str(row["norm_product_id"] or ""),
                    raw_sku=str(row["raw_canonical_sku"] or ""),
                    normalized_sku=str(row["norm_canonical_sku"] or ""),
                    po_reference="N/A",
                    reason_codes=v_codes or (ReasonCode.CONFLICTING_MASTER_RECORD,),
                    explanation=(
                        f"Conflicting Product Master definition for "
                        f"product_id={row['norm_product_id']}, "
                        f"canonical_sku={row['norm_canonical_sku']}; quarantined."
                    ),
                    monetary_impact=None,
                    quantity_impact=None,
                    suggested_action=(
                        "Resolve conflicting SKU, cost, or supplier attributes in Product Master."
                    ),
                )
            )
        elif row["_is_duplicate"]:
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DUPLICATE_TRANSACTION,
                    severity=Severity.LOW,
                    source_dataset=DatasetKind.PRODUCTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_product_id"] or ""),
                    normalized_reference=str(row["norm_product_id"] or ""),
                    raw_sku=str(row["raw_canonical_sku"] or ""),
                    normalized_sku=str(row["norm_canonical_sku"] or ""),
                    po_reference="N/A",
                    reason_codes=(ReasonCode.DUPLICATE_PRODUCT_ROW,),
                    explanation=(
                        f"Duplicate product master row for {row['norm_product_id']} "
                        f"(first seen at row {row['_duplicate_of_row']}); "
                        "excluded from catalog index."
                    ),
                    monetary_impact=None,
                    quantity_impact=None,
                    suggested_action="Remove duplicate row from Product Master export.",
                )
            )
        elif not row["_is_valid_record"]:
            v_codes = tuple(ReasonCode(c) for c in (row["_validation_codes"] or []))
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DATA_VALIDATION_ERROR,
                    severity=Severity.MEDIUM,
                    source_dataset=DatasetKind.PRODUCTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_product_id"] or ""),
                    normalized_reference=str(row["norm_product_id"] or ""),
                    raw_sku=str(row["raw_canonical_sku"] or ""),
                    normalized_sku=str(row["norm_canonical_sku"] or ""),
                    po_reference="N/A",
                    reason_codes=v_codes or (ReasonCode.INVALID_NUMERIC_VALUE,),
                    explanation=(
                        "Invalid Product Master record "
                        f"(standard_cost={row['raw_standard_cost']!r})."
                    ),
                    monetary_impact=None,
                    quantity_impact=None,
                    suggested_action="Correct standard_cost or missing SKU in Product Master.",
                )
            )


def _collect_shipment_exceptions(
    matched_ship_df: pl.DataFrame,
    exceptions: list[ExceptionRecord],
    ambiguous_po_keys: set[PoLineKey],
    match_status_counter: Counter[str],
    next_exc_id: Any,
) -> None:
    """Collect Shipment duplicate, conflict, validation, ambiguous, and orphan exceptions."""
    for row in matched_ship_df.to_dicts():
        m_status = str(row["match_status"])
        match_status_counter[f"shipment:{m_status}"] += 1
        qty_dec = Decimal(row["norm_shipped_qty"]) if row["norm_shipped_qty"] is not None else None

        if row["_is_conflicting_duplicate"]:
            ambiguous_po_keys.update(_row_candidate_po_keys(row))
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DUPLICATE_TRANSACTION,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.SHIPMENTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_shipment_id"] or ""),
                    normalized_reference=str(row["norm_shipment_id"] or ""),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=(ReasonCode.CONFLICTING_DUPLICATE_SHIPMENT,),
                    explanation=(
                        f"Conflicting duplicate shipment row {row['norm_shipment_id']} "
                        f"(qty={row['raw_shipped_qty']!r}, date={row['raw_ship_date']!r}); "
                        "all conflicting rows quarantined."
                    ),
                    monetary_impact=None,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Confirm delivered quantity with warehouse and remove conflicting ASNs."
                    ),
                )
            )
            continue

        if row["_is_duplicate"]:
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DUPLICATE_TRANSACTION,
                    severity=Severity.MEDIUM,
                    source_dataset=DatasetKind.SHIPMENTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_shipment_id"] or ""),
                    normalized_reference=str(row["norm_shipment_id"] or ""),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=(ReasonCode.DUPLICATE_SHIPMENT,),
                    explanation=(
                        f"Duplicate shipment row {row['norm_shipment_id']} "
                        f"(first seen at row {row['_duplicate_of_row']}); "
                        "quarantined to prevent quantity inflation."
                    ),
                    monetary_impact=None,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Deduplicate WMS/carrier shipment feed before posting receipt."
                    ),
                )
            )
            continue

        if not row["_is_valid_record"]:
            v_codes = tuple(ReasonCode(c) for c in (row["_validation_codes"] or []))
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DATA_VALIDATION_ERROR,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.SHIPMENTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_shipment_id"] or ""),
                    normalized_reference=str(row["norm_shipment_id"] or ""),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=v_codes or (ReasonCode.INVALID_NUMERIC_VALUE,),
                    explanation=(
                        f"Shipment record failed validation (qty={row['raw_shipped_qty']!r}, "
                        f"date={row['raw_ship_date']!r})."
                    ),
                    monetary_impact=None,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Request corrected packing list / ASN from warehouse or carrier."
                    ),
                )
            )
            continue

        if m_status == MatchStatus.AMBIGUOUS.value:
            cand_keys = _row_candidate_po_keys(row)
            ambiguous_po_keys.update(cand_keys)
            m_reason = ReasonCode(row["match_reason"])
            cand_label = (
                ", ".join(k.format_display() for k in cand_keys) or "multiple master products"
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.AMBIGUOUS_REFERENCE,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.SHIPMENTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_shipment_id"] or ""),
                    normalized_reference=str(row["norm_shipment_id"] or ""),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or ""),
                    reason_codes=(m_reason,),
                    explanation=(
                        f"Shipment {row['norm_shipment_id']} could not be uniquely assigned "
                        f"({m_reason.value}; candidates: {cand_label})."
                    ),
                    monetary_impact=None,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Assign explicit PO line ID or disambiguate product alias in master data."
                    ),
                )
            )
            continue

        if row["matched_po_number"] is None:
            m_reason = ReasonCode(row["match_reason"])
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.UNEXPECTED_EXTRA_RECORD,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.SHIPMENTS,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_shipment_id"] or ""),
                    normalized_reference=str(row["norm_shipment_id"] or ""),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=(m_reason,),
                    explanation=(
                        f"Unmatched shipment {row['norm_shipment_id']} ({m_reason.value}) "
                        f"referencing PO={row['raw_po_number']!r}, SKU={row['raw_sku']!r}."
                    ),
                    monetary_impact=None,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Verify whether shipment is unauthorized or missing PO reference."
                    ),
                )
            )


def _with_reason_code(codes: tuple[ReasonCode, ...], extra: ReasonCode) -> tuple[ReasonCode, ...]:
    """Append a reason code to a tuple if not already present."""
    return codes if extra in codes else (*codes, extra)


def _format_foreign_currency_note(
    parsed_amt: Decimal | None,
    raw_amt: Any,
    norm_curr: str | None,
    raw_curr: Any,
    reporting_currency: str,
) -> str:
    """Format an auditable note preserving original foreign-currency amount and code."""
    curr_label = norm_curr or (str(raw_curr).strip() if raw_curr is not None else "") or "MISSING"
    amt_txt = f"{parsed_amt:.2f}" if parsed_amt is not None else repr(raw_amt)
    return (
        f" Original amount {amt_txt} {curr_label} (raw_amount={raw_amt!r}, currency={raw_curr!r}) "
        f"excluded from {reporting_currency} totals without FX conversion."
    )


def _collect_invoice_exceptions(
    matched_inv_df: pl.DataFrame,
    config: ReconciliationConfig,
    exceptions: list[ExceptionRecord],
    ambiguous_po_keys: set[PoLineKey],
    match_status_counter: Counter[str],
    next_exc_id: Any,
) -> None:
    """Collect Invoice duplicate, conflict, validation, ambiguous & orphan exceptions."""
    req_currency = config.reporting_currency.strip().upper()
    for row in matched_inv_df.to_dicts():
        m_status = str(row["match_status"])
        match_status_counter[f"invoice:{m_status}"] += 1
        qty_dec = (
            Decimal(row["norm_invoiced_qty"]) if row["norm_invoiced_qty"] is not None else None
        )
        parsed_amt_dec = (
            config.quantize_money(Decimal(row["norm_line_amount"]))
            if row["norm_line_amount"] is not None
            else None
        )
        norm_curr = row["norm_currency"]
        raw_curr = row["raw_currency"]
        raw_curr_clean = str(raw_curr or "").strip()
        currency_ok = norm_curr == req_currency
        # Never aggregate foreign-currency or missing-currency amounts into reporting-currency KPIs
        amt_dec = parsed_amt_dec if currency_ok else None
        orig_curr = norm_curr or (raw_curr_clean if raw_curr_clean else None)
        fx_note = (
            ""
            if currency_ok
            else _format_foreign_currency_note(
                parsed_amt_dec,
                row["raw_line_amount"],
                norm_curr,
                raw_curr,
                req_currency,
            )
        )

        if row["_is_conflicting_duplicate"]:
            ambiguous_po_keys.update(_row_candidate_po_keys(row))
            r_codes = (ReasonCode.CONFLICTING_DUPLICATE_INVOICE,)
            if not currency_ok:
                r_codes = _with_reason_code(r_codes, ReasonCode.CURRENCY_MISMATCH)
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DUPLICATE_TRANSACTION,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.INVOICES,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_invoice_line_id"] or row["raw_invoice_id"] or ""),
                    normalized_reference=str(
                        row["norm_invoice_line_id"] or row["norm_invoice_id"] or ""
                    ),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=r_codes,
                    explanation=(
                        f"Conflicting duplicate invoice line {row['norm_invoice_line_id']} "
                        f"for supplier {row['norm_supplier_id']}; "
                        f"all conflicting rows quarantined.{fx_note}"
                    ),
                    monetary_impact=amt_dec,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Place all conflicting versions on AP hold and verify with supplier."
                    ),
                    original_monetary_amount=parsed_amt_dec,
                    original_currency=orig_curr,
                )
            )
            continue

        if row["_is_duplicate"]:
            r_codes = (ReasonCode.DUPLICATE_INVOICE,)
            if not currency_ok:
                r_codes = _with_reason_code(r_codes, ReasonCode.CURRENCY_MISMATCH)
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DUPLICATE_TRANSACTION,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.INVOICES,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_invoice_line_id"] or row["raw_invoice_id"] or ""),
                    normalized_reference=str(
                        row["norm_invoice_line_id"] or row["norm_invoice_id"] or ""
                    ),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=r_codes,
                    explanation=(
                        f"Duplicate supplier invoice line {row['norm_invoice_line_id']} "
                        f"(first seen at row {row['_duplicate_of_row']}); "
                        f"blocked to prevent double payment.{fx_note}"
                    ),
                    monetary_impact=amt_dec,
                    quantity_impact=qty_dec,
                    suggested_action="Reject duplicate supplier billing line in Accounts Payable.",
                    original_monetary_amount=parsed_amt_dec,
                    original_currency=orig_curr,
                )
            )
            continue

        if not row["_is_valid_record"]:
            v_codes = tuple(ReasonCode(c) for c in (row["_validation_codes"] or []))
            if not currency_ok:
                v_codes = _with_reason_code(v_codes, ReasonCode.CURRENCY_MISMATCH)
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.DATA_VALIDATION_ERROR,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.INVOICES,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_invoice_line_id"] or row["raw_invoice_id"] or ""),
                    normalized_reference=str(
                        row["norm_invoice_line_id"] or row["norm_invoice_id"] or ""
                    ),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=v_codes or (ReasonCode.INVALID_NUMERIC_VALUE,),
                    explanation=(
                        "Supplier invoice line failed validation "
                        f"(qty={row['raw_invoiced_qty']!r}, price={row['raw_unit_price']!r}, "
                        f"amount={row['raw_line_amount']!r}, "
                        f"currency={row['raw_currency']!r}).{fx_note}"
                    ),
                    monetary_impact=amt_dec,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Place invoice on AP hold and request corrected invoice from supplier."
                    ),
                    original_monetary_amount=parsed_amt_dec,
                    original_currency=orig_curr,
                )
            )
            continue

        if m_status == MatchStatus.AMBIGUOUS.value:
            cand_keys = _row_candidate_po_keys(row)
            ambiguous_po_keys.update(cand_keys)
            m_reason = ReasonCode(row["match_reason"])
            r_codes = (m_reason,)
            if not currency_ok:
                r_codes = _with_reason_code(r_codes, ReasonCode.CURRENCY_MISMATCH)
            cand_label = (
                ", ".join(k.format_display() for k in cand_keys) or "multiple master products"
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.AMBIGUOUS_REFERENCE,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.INVOICES,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_invoice_line_id"] or ""),
                    normalized_reference=str(row["norm_invoice_line_id"] or ""),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or ""),
                    reason_codes=r_codes,
                    explanation=(
                        f"Invoice line {row['norm_invoice_line_id']} matched multiple candidates "
                        f"({m_reason.value}; candidates: {cand_label}).{fx_note}"
                    ),
                    monetary_impact=amt_dec,
                    quantity_impact=qty_dec,
                    suggested_action=(
                        "Confirm target PO line with buyer before releasing invoice payment."
                    ),
                    original_monetary_amount=parsed_amt_dec,
                    original_currency=orig_curr,
                )
            )
            continue

        if row["matched_po_number"] is None:
            m_reason = ReasonCode(row["match_reason"])
            r_codes = (m_reason,)
            if not currency_ok:
                r_codes = _with_reason_code(r_codes, ReasonCode.CURRENCY_MISMATCH)
            action = (
                "Hold payment; invoice supplier_id does not match target Purchase Order supplier."
                if m_reason == ReasonCode.SUPPLIER_MISMATCH
                else "Hold payment; verify PO authorization with procurement."
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=next_exc_id(),
                    category=ExceptionCategory.UNEXPECTED_EXTRA_RECORD,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.INVOICES,
                    source_row=int(row["_source_row"]),
                    raw_reference=str(row["raw_invoice_line_id"] or ""),
                    normalized_reference=str(row["norm_invoice_line_id"] or ""),
                    raw_sku=str(row["raw_sku"] or ""),
                    normalized_sku=str(row["norm_sku"] or ""),
                    po_reference=str(row["norm_po_number"] or "MISSING"),
                    reason_codes=r_codes,
                    explanation=(
                        f"Unexpected supplier invoice line {row['norm_invoice_line_id']} "
                        f"({m_reason.value}) with supplier={row['raw_supplier_id']!r}, "
                        f"PO={row['raw_po_number']!r}, SKU={row['raw_sku']!r}.{fx_note}"
                    ),
                    monetary_impact=amt_dec,
                    quantity_impact=qty_dec,
                    suggested_action=action,
                    original_monetary_amount=parsed_amt_dec,
                    original_currency=orig_curr,
                )
            )


def reconcile_datasets(
    products_clean: CleanedDataset,
    po_clean: CleanedDataset,
    shipments_clean: CleanedDataset,
    invoices_clean: CleanedDataset,
    *,
    config: ReconciliationConfig | None = None,
    input_metadata: tuple[LoadedWorkbookData, ...] = (),
) -> ReconciliationResult:
    """Execute deterministic matching and three-way financial reconciliation on cleaned datasets."""
    cfg = config or ReconciliationConfig()
    invalid_schemas = tuple(
        ds.schema_validation
        for ds in (products_clean, po_clean, shipments_clean, invoices_clean)
        if not ds.schema_validation.is_valid
    )
    if invalid_schemas:
        raise SchemaValidationError(invalid_schemas)

    warnings: list[str] = []

    product_matcher = ProductMatcher(products_clean.dataframe)
    matched_po_df = match_purchase_orders_df(po_clean.dataframe, product_matcher)
    line_matcher = OrderLineMatcher(matched_po_df, product_matcher)

    raw_matched_ship_df = line_matcher.match_transactions_df(
        shipments_clean.dataframe, validate_supplier=False
    )
    matched_ship_df, effective_ship_summary = deduplicate_matched_shipments(
        raw_matched_ship_df, shipments_clean.summary
    )

    raw_matched_inv_df = line_matcher.match_transactions_df(
        invoices_clean.dataframe, validate_supplier=True
    )
    matched_inv_df, effective_inv_summary = deduplicate_matched_invoices(
        raw_matched_inv_df, invoices_clean.summary
    )

    exceptions: list[ExceptionRecord] = []
    exc_seq = 0

    def _next_exc_id() -> str:
        nonlocal exc_seq
        exc_seq += 1
        return f"EXC-{exc_seq:04d}"

    ambiguous_po_keys: set[PoLineKey] = set()
    match_status_counter: Counter[str] = Counter()

    _collect_product_exceptions(products_clean.dataframe, exceptions, _next_exc_id)
    _collect_shipment_exceptions(
        matched_ship_df, exceptions, ambiguous_po_keys, match_status_counter, _next_exc_id
    )
    _collect_invoice_exceptions(
        matched_inv_df, cfg, exceptions, ambiguous_po_keys, match_status_counter, _next_exc_id
    )

    ship_agg_by_key = _aggregate_matched_shipments(matched_ship_df)
    inv_agg_by_key = _aggregate_matched_invoices(matched_inv_df, cfg)

    reconciled_records: list[ReconciledRecord] = []
    total_ordered_qty = Decimal("0")
    total_shipped_qty = Decimal("0")
    total_invoiced_qty = Decimal("0")
    total_outstanding_qty = Decimal("0")
    total_expected_amount = Decimal("0.00")
    total_invoiced_amount = Decimal("0.00")
    net_billed_variance = Decimal("0.00")
    absolute_monetary_discrepancy = Decimal("0.00")
    unbilled_shipped_amount = Decimal("0.00")

    unique_valid_po_lines = 0
    duplicate_po_rows = 0
    invalid_po_rows = 0
    reconciled_lines = 0
    partial_fulfillment_lines = 0
    po_exception_lines = 0
    matched_po_product_count = 0

    req_currency = cfg.reporting_currency.strip().upper()
    for row in matched_po_df.to_dicts():
        source_row = int(row["_source_row"])
        raw_po = str(row["raw_po_number"] or "")
        raw_line = str(row["raw_po_line_id"] or "")
        raw_sku = str(row["raw_sku"] or "")
        norm_po = str(row["norm_po_number"] or "")
        norm_line = str(row["norm_po_line_id"] or "")
        norm_sku = str(row["norm_sku"] or "")
        prod_match_status = MatchStatus(row["product_match_status"])
        match_status_counter[f"po:{prod_match_status.value}"] += 1

        preserved_formula = _extract_formula(row["_formulas_json"], "line_total")
        ordered_qty = (
            Decimal(row["norm_ordered_qty"]) if row["norm_ordered_qty"] is not None else None
        )
        po_unit_price = (
            Decimal(row["norm_unit_price"]) if row["norm_unit_price"] is not None else None
        )
        po_line_total_val = (
            Decimal(row["norm_line_total"]) if row["norm_line_total"] is not None else None
        )
        parsed_expected_amt = (
            cfg.quantize_money(ordered_qty * po_unit_price)
            if (ordered_qty is not None and po_unit_price is not None)
            else None
        )
        po_norm_curr = row["norm_currency"]
        po_raw_curr = row["raw_currency"]
        po_curr_clean = str(po_raw_curr or "").strip()
        po_currency_ok = po_norm_curr == req_currency
        expected_amt = parsed_expected_amt if po_currency_ok else None
        safe_po_unit_price = po_unit_price if po_currency_ok else None
        po_orig_curr = po_norm_curr or (po_curr_clean if po_curr_clean else None)
        po_fx_note = (
            ""
            if po_currency_ok
            else _format_foreign_currency_note(
                parsed_expected_amt,
                row["raw_line_total"],
                po_norm_curr,
                po_raw_curr,
                req_currency,
            )
        )
        desc = str(row["norm_product_description"] or row["canonical_product_name"] or raw_sku)
        supplier_id = str(row["norm_supplier_id"] or "")

        # 5a. Conflicting or identical duplicate PO line -> quarantined
        if row["_is_conflicting_duplicate"]:
            duplicate_po_rows += 1
            po_exception_lines += 1
            r_codes = (ReasonCode.CONFLICTING_DUPLICATE_PO_LINE,)
            if not po_currency_ok:
                r_codes = _with_reason_code(r_codes, ReasonCode.CURRENCY_MISMATCH)
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=row["matched_product_id"],
                    canonical_sku=row["canonical_sku"],
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=None,
                    invoiced_qty=None,
                    outstanding_qty=None,
                    po_unit_price=safe_po_unit_price,
                    expected_amount=expected_amt,
                    actual_invoiced_amount=None,
                    monetary_variance=None,
                    shipment_count=0,
                    invoice_count=0,
                    shipment_refs="",
                    invoice_refs="",
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.DUPLICATE_TRANSACTION,
                    reason_codes=r_codes,
                    is_valid_po_line=False,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.DUPLICATE_TRANSACTION,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line or raw_po,
                    normalized_reference=norm_line or norm_po,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po or "MISSING",
                    reason_codes=r_codes,
                    explanation=(
                        f"Conflicting duplicate PO line {norm_line} on {norm_po}; "
                        f"all conflicting rows quarantined.{po_fx_note}"
                    ),
                    monetary_impact=expected_amt,
                    quantity_impact=ordered_qty,
                    suggested_action="Resolve conflicting PO line terms in ERP export.",
                    original_monetary_amount=parsed_expected_amt,
                    original_currency=po_orig_curr,
                )
            )
            continue

        if row["_is_duplicate"]:
            duplicate_po_rows += 1
            po_exception_lines += 1
            r_codes = (ReasonCode.DUPLICATE_PO_LINE,)
            if not po_currency_ok:
                r_codes = _with_reason_code(r_codes, ReasonCode.CURRENCY_MISMATCH)
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=row["matched_product_id"],
                    canonical_sku=row["canonical_sku"],
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=None,
                    invoiced_qty=None,
                    outstanding_qty=None,
                    po_unit_price=safe_po_unit_price,
                    expected_amount=expected_amt,
                    actual_invoiced_amount=None,
                    monetary_variance=None,
                    shipment_count=0,
                    invoice_count=0,
                    shipment_refs="",
                    invoice_refs="",
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.DUPLICATE_TRANSACTION,
                    reason_codes=r_codes,
                    is_valid_po_line=False,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.DUPLICATE_TRANSACTION,
                    severity=Severity.MEDIUM,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line or raw_po,
                    normalized_reference=norm_line or norm_po,
                    raw_sku=raw_sku,
                    normalized_sku=str(norm_sku),
                    po_reference=norm_po or "MISSING",
                    reason_codes=r_codes,
                    explanation=(
                        f"Duplicate PO line {norm_line} (first seen at row "
                        f"{row['_duplicate_of_row']}); excluded from open order "
                        f"commitments.{po_fx_note}"
                    ),
                    monetary_impact=expected_amt,
                    quantity_impact=ordered_qty,
                    suggested_action="Remove duplicate PO line from ERP purchase order export.",
                    original_monetary_amount=parsed_expected_amt,
                    original_currency=po_orig_curr,
                )
            )
            continue

        # 5b. Invalid PO row
        if not row["_is_valid_record"]:
            invalid_po_rows += 1
            po_exception_lines += 1
            v_codes = tuple(ReasonCode(c) for c in (row["_validation_codes"] or []))
            if not po_currency_ok:
                v_codes = _with_reason_code(v_codes, ReasonCode.CURRENCY_MISMATCH)
            if not v_codes:
                v_codes = (ReasonCode.INVALID_NUMERIC_VALUE,)
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=row["matched_product_id"],
                    canonical_sku=row["canonical_sku"],
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=None,
                    invoiced_qty=None,
                    outstanding_qty=None,
                    po_unit_price=safe_po_unit_price,
                    expected_amount=expected_amt,
                    actual_invoiced_amount=None,
                    monetary_variance=None,
                    shipment_count=0,
                    invoice_count=0,
                    shipment_refs="",
                    invoice_refs="",
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.DATA_VALIDATION_ERROR,
                    reason_codes=v_codes,
                    is_valid_po_line=False,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.DATA_VALIDATION_ERROR,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line or raw_po,
                    normalized_reference=norm_line or norm_po,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po or "MISSING",
                    reason_codes=v_codes,
                    explanation=(
                        f"PO line failed validation (qty={row['raw_ordered_qty']!r}, "
                        f"unit_price={row['raw_unit_price']!r}, date={row['raw_order_date']!r}, "
                        f"currency={row['raw_currency']!r}).{po_fx_note}"
                    ),
                    monetary_impact=expected_amt,
                    quantity_impact=ordered_qty,
                    suggested_action="Amend invalid numeric/date/currency/formula values on PO.",
                    original_monetary_amount=parsed_expected_amt,
                    original_currency=po_orig_curr,
                )
            )
            continue

        # Valid unique PO line
        unique_valid_po_lines += 1
        assert ordered_qty is not None and po_unit_price is not None and expected_amt is not None
        po_line_key = PoLineKey(po_number=norm_po, po_line_id=norm_line)

        if prod_match_status in (MatchStatus.EXACT_ID, MatchStatus.ALIAS_MAP):
            matched_po_product_count += 1

        ship_agg = ship_agg_by_key.get(po_line_key)
        inv_agg = inv_agg_by_key.get(po_line_key)

        fin = compute_po_line_financials(
            ordered_qty=ordered_qty,
            po_unit_price=po_unit_price,
            po_line_total_val=po_line_total_val,
            ship_agg=ship_agg,
            inv_agg=inv_agg,
            config=cfg,
        )

        total_ordered_qty += fin.ordered_qty
        if fin.shipped_qty is not None:
            total_shipped_qty += fin.shipped_qty
        if fin.invoiced_qty is not None:
            total_invoiced_qty += fin.invoiced_qty
        total_outstanding_qty += fin.outstanding_qty
        total_expected_amount += fin.expected_amount
        if fin.actual_invoiced_amount is not None:
            total_invoiced_amount += fin.actual_invoiced_amount

        ship_count = ship_agg.shipment_count if ship_agg else 0
        ship_refs = ship_agg.shipment_refs if ship_agg else ""
        inv_count = inv_agg.invoice_count if inv_agg else 0
        inv_refs = inv_agg.invoice_refs if inv_agg else ""

        # 5c. Ambiguous or unmapped SKU on PO line
        if prod_match_status == MatchStatus.AMBIGUOUS:
            po_exception_lines += 1
            p_reason = ReasonCode(row["product_match_reason"])
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=None,
                    canonical_sku=None,
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=fin.shipped_qty,
                    invoiced_qty=fin.invoiced_qty,
                    outstanding_qty=fin.outstanding_qty,
                    po_unit_price=po_unit_price,
                    expected_amount=fin.expected_amount,
                    actual_invoiced_amount=fin.actual_invoiced_amount,
                    monetary_variance=None,
                    shipment_count=ship_count,
                    invoice_count=inv_count,
                    shipment_refs=ship_refs,
                    invoice_refs=inv_refs,
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.AMBIGUOUS_REFERENCE,
                    reason_codes=(p_reason,),
                    is_valid_po_line=True,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.AMBIGUOUS_REFERENCE,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line,
                    normalized_reference=norm_line,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po,
                    reason_codes=(p_reason,),
                    explanation=(
                        f"PO line {norm_line} SKU {raw_sku!r} maps to multiple or conflicted "
                        f"Product Master items ({p_reason.value})."
                    ),
                    monetary_impact=expected_amt,
                    quantity_impact=ordered_qty,
                    suggested_action="Resolve ambiguous alias/SKU mapping in Product Master.",
                )
            )
            continue

        if prod_match_status in (MatchStatus.UNMATCHED, MatchStatus.INVALID_KEY):
            po_exception_lines += 1
            p_reason = ReasonCode(row["product_match_reason"])
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=None,
                    canonical_sku=None,
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=fin.shipped_qty,
                    invoiced_qty=fin.invoiced_qty,
                    outstanding_qty=fin.outstanding_qty,
                    po_unit_price=po_unit_price,
                    expected_amount=fin.expected_amount,
                    actual_invoiced_amount=fin.actual_invoiced_amount,
                    monetary_variance=None,
                    shipment_count=ship_count,
                    invoice_count=inv_count,
                    shipment_refs=ship_refs,
                    invoice_refs=inv_refs,
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.DATA_VALIDATION_ERROR,
                    reason_codes=(p_reason,),
                    is_valid_po_line=True,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.DATA_VALIDATION_ERROR,
                    severity=Severity.MEDIUM,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line,
                    normalized_reference=norm_line,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po,
                    reason_codes=(p_reason,),
                    explanation=f"PO line {norm_line} references unrecognized SKU {raw_sku!r}.",
                    monetary_impact=expected_amt,
                    quantity_impact=ordered_qty,
                    suggested_action="Add canonical SKU or supplier alias to Product Master.",
                )
            )
            continue

        # 5d. PO line touched by quarantined ambiguous or conflicting transaction(s)
        if po_line_key in ambiguous_po_keys:
            po_exception_lines += 1
            r_codes = (ReasonCode.AMBIGUOUS_TRANSACTION_LINK,)
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=row["matched_product_id"],
                    canonical_sku=row["canonical_sku"],
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=fin.shipped_qty,
                    invoiced_qty=fin.invoiced_qty,
                    outstanding_qty=fin.outstanding_qty,
                    po_unit_price=po_unit_price,
                    expected_amount=fin.expected_amount,
                    actual_invoiced_amount=fin.actual_invoiced_amount,
                    monetary_variance=None,
                    shipment_count=ship_count,
                    invoice_count=inv_count,
                    shipment_refs=ship_refs,
                    invoice_refs=inv_refs,
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.AMBIGUOUS_REFERENCE,
                    reason_codes=r_codes,
                    is_valid_po_line=True,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.AMBIGUOUS_REFERENCE,
                    severity=Severity.MEDIUM,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line,
                    normalized_reference=norm_line,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po,
                    reason_codes=r_codes,
                    explanation=(
                        f"PO line {norm_line} has quarantined ambiguous or conflicting "
                        f"shipment/invoice records on {norm_po}."
                    ),
                    monetary_impact=expected_amt,
                    quantity_impact=ordered_qty,
                    suggested_action=(
                        "Resolve quarantined shipments/invoices (line IDs or conflicts) and re-run."
                    ),
                )
            )
            continue

        # Accumulate line-level financial metrics independently of primary exception classification
        if fin.monetary_variance is not None:
            net_billed_variance += fin.monetary_variance
        absolute_monetary_discrepancy += fin.discrepancy_exposure
        unbilled_shipped_amount += fin.unbilled_receiving_amount

        # 5e. Missing Shipment
        if ship_count == 0 or fin.shipped_qty is None:
            po_exception_lines += 1
            if inv_count > 0 and inv_agg is not None and fin.actual_invoiced_amount is not None:
                codes_list = [ReasonCode.INVOICED_WITHOUT_SHIPMENT]
                if fin.has_unit_price_discrepancy:
                    codes_list.append(ReasonCode.UNIT_PRICE_DISCREPANCY)
                sev = Severity.HIGH
                expl = (
                    f"Supplier billed {fin.invoiced_qty} units (${fin.actual_invoiced_amount}) "
                    f"on {inv_refs}, but zero shipment records exist for {norm_line}."
                )
                action = "Block payment immediately; verify proof of delivery with receiving."
                mon_impact = fin.discrepancy_exposure
            else:
                codes_list = [ReasonCode.NO_SHIPMENT_RECORDED]
                sev = Severity.MEDIUM
                expl = f"Open PO line {norm_line} ({ordered_qty} units) has no recorded shipment."
                action = "Follow up with supplier on expected ship date or cancel backorder."
                mon_impact = expected_amt

            r_codes = tuple(codes_list)
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=row["matched_product_id"],
                    canonical_sku=row["canonical_sku"],
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=None,
                    invoiced_qty=fin.invoiced_qty,
                    outstanding_qty=fin.outstanding_qty,
                    po_unit_price=po_unit_price,
                    expected_amount=fin.expected_amount,
                    actual_invoiced_amount=fin.actual_invoiced_amount,
                    monetary_variance=fin.monetary_variance,
                    shipment_count=0,
                    invoice_count=inv_count,
                    shipment_refs="",
                    invoice_refs=inv_refs,
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.MISSING_SHIPMENT,
                    reason_codes=r_codes,
                    is_valid_po_line=True,
                    discrepancy_exposure=fin.discrepancy_exposure,
                    unbilled_receiving_amount=fin.unbilled_receiving_amount,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.MISSING_SHIPMENT,
                    severity=sev,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line,
                    normalized_reference=norm_line,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po,
                    reason_codes=r_codes,
                    explanation=expl,
                    monetary_impact=mon_impact,
                    quantity_impact=ordered_qty,
                    suggested_action=action,
                )
            )
            continue

        # 5f. Missing Invoice (goods shipped, zero invoices)
        if inv_count == 0 or inv_agg is None or fin.invoiced_qty is None:
            po_exception_lines += 1
            if fin.shipped_qty < ordered_qty - cfg.quantity_tolerance:
                codes_list = [ReasonCode.PARTIAL_SHIPMENT_UNBILLED]
            else:
                codes_list = [ReasonCode.NO_INVOICE_RECORDED]
            if fin.shipped_qty > ordered_qty + cfg.quantity_tolerance:
                codes_list.append(ReasonCode.OVER_SHIPMENT)

            r_codes = tuple(codes_list)
            reconciled_records.append(
                _make_reconciled_record(
                    source_row=source_row,
                    raw_po=raw_po,
                    raw_line=raw_line,
                    raw_sku=raw_sku,
                    norm_po=norm_po,
                    norm_line=norm_line,
                    norm_sku=norm_sku,
                    matched_product_id=row["matched_product_id"],
                    canonical_sku=row["canonical_sku"],
                    matched_business_key=row["matched_business_key"],
                    product_description=desc,
                    supplier_id=supplier_id,
                    order_date=row["norm_order_date"],
                    product_match_status=prod_match_status,
                    ordered_qty=ordered_qty,
                    shipped_qty=fin.shipped_qty,
                    invoiced_qty=None,
                    outstanding_qty=fin.outstanding_qty,
                    po_unit_price=po_unit_price,
                    expected_amount=fin.expected_amount,
                    actual_invoiced_amount=None,
                    monetary_variance=None,
                    shipment_count=ship_count,
                    invoice_count=0,
                    shipment_refs=ship_refs,
                    invoice_refs="",
                    preserved_po_formula=preserved_formula,
                    reconciliation_status=ReconciliationStatus.MISSING_INVOICE,
                    reason_codes=r_codes,
                    is_valid_po_line=True,
                    discrepancy_exposure=fin.discrepancy_exposure,
                    unbilled_receiving_amount=fin.unbilled_receiving_amount,
                )
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.MISSING_INVOICE,
                    severity=Severity.MEDIUM,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line,
                    normalized_reference=norm_line,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po,
                    reason_codes=r_codes,
                    explanation=(
                        f"Shipped {fin.shipped_qty} of {ordered_qty} units ({ship_refs}) "
                        f"with unbilled receipt accrual ${fin.unbilled_receiving_amount}, "
                        "but no supplier invoice has been received."
                    ),
                    monetary_impact=fin.unbilled_receiving_amount,
                    quantity_impact=fin.shipped_qty,
                    suggested_action=(
                        "Accrue unbilled receipt in GL and request invoice from supplier."
                    ),
                )
            )
            continue

        # 5g. Full Three-Way Evaluation (PO + Shipment(s) + Invoice(s) present)
        assert fin.monetary_variance is not None and fin.actual_invoiced_amount is not None
        po_math_off = (
            po_line_total_val is not None
            and abs(po_line_total_val - fin.expected_amount) > cfg.amount_tolerance
        )
        shipped_matches_po = abs(fin.shipped_qty - ordered_qty) <= cfg.quantity_tolerance
        invoiced_matches_shipped = abs(fin.invoiced_qty - fin.shipped_qty) <= cfg.quantity_tolerance
        is_clean_partial = (
            Decimal("0") < fin.shipped_qty < ordered_qty - cfg.quantity_tolerance
            and invoiced_matches_shipped
        )
        partially_invoiced_receipt = (
            Decimal("0")
            < fin.invoiced_qty
            < min(fin.shipped_qty, ordered_qty) - cfg.quantity_tolerance
        )

        reason_list: list[ReasonCode] = []
        if fin.has_invoice_math_discrepancy:
            reason_list.append(ReasonCode.INVOICE_MATH_DISCREPANCY)
        if fin.has_unit_price_discrepancy:
            reason_list.append(ReasonCode.UNIT_PRICE_DISCREPANCY)
        if po_math_off:
            reason_list.append(ReasonCode.PO_MATH_DISCREPANCY)

        if fin.shipped_qty > ordered_qty + cfg.quantity_tolerance:
            reason_list.append(ReasonCode.OVER_SHIPMENT)
        elif fin.shipped_qty < ordered_qty - cfg.quantity_tolerance and not is_clean_partial:
            reason_list.append(ReasonCode.UNDER_SHIPMENT)

        if fin.invoiced_qty > fin.shipped_qty + cfg.quantity_tolerance:
            reason_list.append(ReasonCode.INVOICE_QTY_EXCEEDS_SHIPPED)
        elif not invoiced_matches_shipped:
            reason_list.append(ReasonCode.INVOICE_QTY_MISMATCH)

        if partially_invoiced_receipt:
            reason_list.append(ReasonCode.PARTIAL_SHIPMENT_UNBILLED)

        if fin.has_invoice_math_discrepancy or fin.has_unit_price_discrepancy or po_math_off:
            po_exception_lines += 1
            status = ReconciliationStatus.PRICE_MISMATCH
            r_codes = tuple(reason_list)
            obs_prices_str = ", ".join(f"${p}" for p in inv_agg.unit_prices)

            # Provenance: cite offending supplier_invoices row for invoice price/math discrepancies
            if fin.has_invoice_math_discrepancy or fin.has_unit_price_discrepancy:
                offending = next(
                    (
                        line
                        for line in inv_agg.lines
                        if line.has_internal_math_err
                        or abs(line.unit_price - po_unit_price) > cfg.unit_price_tolerance
                    ),
                    inv_agg.lines[0],
                )
                exc_dataset = DatasetKind.INVOICES
                exc_row = offending.source_row
                exc_raw_ref = offending.raw_invoice_line_id or offending.raw_invoice_id
                exc_norm_ref = offending.norm_invoice_line_id or offending.norm_invoice_id
                exc_raw_sku = offending.raw_sku
                exc_norm_sku = offending.norm_sku
            else:
                exc_dataset = DatasetKind.PURCHASE_ORDERS
                exc_row = source_row
                exc_raw_ref = raw_line
                exc_norm_ref = norm_line
                exc_raw_sku = raw_sku
                exc_norm_sku = norm_sku

            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.PRICE_MISMATCH,
                    severity=Severity.HIGH,
                    source_dataset=exc_dataset,
                    source_row=exc_row,
                    raw_reference=exc_raw_ref,
                    normalized_reference=exc_norm_ref,
                    raw_sku=exc_raw_sku,
                    normalized_sku=exc_norm_sku,
                    po_reference=norm_po,
                    reason_codes=r_codes,
                    explanation=(
                        f"Monetary discrepancy on {norm_line}: PO unit price ${po_unit_price} vs "
                        f"invoiced [{obs_prices_str}]; billed ${fin.actual_invoiced_amount} for "
                        f"{fin.invoiced_qty} units (exposure=${fin.discrepancy_exposure}, "
                        f"net_variance=${fin.monetary_variance})."
                    ),
                    monetary_impact=fin.discrepancy_exposure
                    if fin.discrepancy_exposure != Decimal("0.00")
                    else fin.monetary_variance,
                    quantity_impact=fin.invoiced_qty - fin.shipped_qty,
                    suggested_action=(
                        "Request credit memo from supplier or verify approved PO price amendment."
                    ),
                )
            )
        elif is_clean_partial:
            partial_fulfillment_lines += 1
            po_exception_lines += 1
            status = ReconciliationStatus.PARTIAL_FULFILLMENT
            r_codes = (ReasonCode.PARTIAL_SHIPMENT_BILLED,)
            remaining_val = cfg.quantize_money(fin.outstanding_qty * po_unit_price)
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.PARTIAL_FULFILLMENT,
                    severity=Severity.LOW,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line,
                    normalized_reference=norm_line,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po,
                    reason_codes=r_codes,
                    explanation=(
                        f"Partial fulfillment on {norm_line}: {fin.shipped_qty} of {ordered_qty} "
                        f"units shipped and accurately billed (${fin.actual_invoiced_amount}); "
                        f"{fin.outstanding_qty} units open."
                    ),
                    monetary_impact=remaining_val,
                    quantity_impact=fin.outstanding_qty,
                    suggested_action="Keep PO line open for remaining backordered quantity.",
                )
            )
        elif not (shipped_matches_po and invoiced_matches_shipped):
            po_exception_lines += 1
            status = ReconciliationStatus.QUANTITY_MISMATCH
            r_codes = tuple(reason_list) or (ReasonCode.INVOICE_QTY_MISMATCH,)
            qty_mon_impact = (
                fin.discrepancy_exposure
                if fin.discrepancy_exposure > Decimal("0.00")
                else fin.unbilled_receiving_amount
            )
            exceptions.append(
                ExceptionRecord(
                    exception_id=_next_exc_id(),
                    category=ExceptionCategory.QUANTITY_MISMATCH,
                    severity=Severity.HIGH,
                    source_dataset=DatasetKind.PURCHASE_ORDERS,
                    source_row=source_row,
                    raw_reference=raw_line,
                    normalized_reference=norm_line,
                    raw_sku=raw_sku,
                    normalized_sku=norm_sku,
                    po_reference=norm_po,
                    reason_codes=r_codes,
                    explanation=(
                        f"Quantity mismatch on {norm_line}: ordered={ordered_qty}, "
                        f"shipped={fin.shipped_qty} ({ship_refs}), "
                        f"invoiced={fin.invoiced_qty} ({inv_refs}); "
                        f"exposure=${fin.discrepancy_exposure}, "
                        f"unbilled=${fin.unbilled_receiving_amount}."
                    ),
                    monetary_impact=qty_mon_impact,
                    quantity_impact=fin.invoiced_qty - fin.shipped_qty,
                    suggested_action=(
                        "Reconcile receiving dock count against supplier invoice before paying."
                    ),
                )
            )
        else:
            reconciled_lines += 1
            status = ReconciliationStatus.RECONCILED
            base_code = (
                ReasonCode.ZERO_VALUE_CONFIRMED
                if ordered_qty == Decimal("0")
                else ReasonCode.THREE_WAY_MATCH
            )
            used_alias = (
                prod_match_status == MatchStatus.ALIAS_MAP
                or (ship_agg is not None and ship_agg.used_alias)
                or inv_agg.used_alias
            )
            r_codes = (base_code, ReasonCode.MATCHED_PRODUCT_ALIAS) if used_alias else (base_code,)

        reconciled_records.append(
            _make_reconciled_record(
                source_row=source_row,
                raw_po=raw_po,
                raw_line=raw_line,
                raw_sku=raw_sku,
                norm_po=norm_po,
                norm_line=norm_line,
                norm_sku=norm_sku,
                matched_product_id=row["matched_product_id"],
                canonical_sku=row["canonical_sku"],
                matched_business_key=row["matched_business_key"],
                product_description=desc,
                supplier_id=supplier_id,
                order_date=row["norm_order_date"],
                product_match_status=prod_match_status,
                ordered_qty=ordered_qty,
                shipped_qty=fin.shipped_qty,
                invoiced_qty=fin.invoiced_qty,
                outstanding_qty=fin.outstanding_qty,
                po_unit_price=po_unit_price,
                expected_amount=fin.expected_amount,
                actual_invoiced_amount=fin.actual_invoiced_amount,
                monetary_variance=fin.monetary_variance,
                shipment_count=ship_count,
                invoice_count=inv_count,
                shipment_refs=ship_refs,
                invoice_refs=inv_refs,
                preserved_po_formula=preserved_formula,
                reconciliation_status=status,
                reason_codes=r_codes,
                is_valid_po_line=True,
                discrepancy_exposure=fin.discrepancy_exposure,
                unbilled_receiving_amount=fin.unbilled_receiving_amount,
            )
        )

    coverage_rate = quantize_rate(matched_po_product_count, unique_valid_po_lines)
    recon_rate = quantize_rate(reconciled_lines, unique_valid_po_lines)

    quarantined_invoice_amount = cfg.quantize_money(
        sum(
            (
                exc.monetary_impact
                for exc in exceptions
                if exc.source_dataset == DatasetKind.INVOICES
                and exc.category != ExceptionCategory.PRICE_MISMATCH
                and exc.monetary_impact is not None
                and (exc.original_currency is None or exc.original_currency.upper() == req_currency)
            ),
            Decimal("0.00"),
        )
    )

    severity_rank = {Severity.HIGH: 0, Severity.MEDIUM: 1, Severity.LOW: 2}
    ordered_exceptions = sorted(
        enumerate(exceptions),
        key=lambda item: (severity_rank[item[1].severity], item[1].source_row, item[0]),
    )
    exceptions = [
        replace(exc_item, exception_id=f"EXC-{new_idx:04d}")
        for new_idx, (_, exc_item) in enumerate(ordered_exceptions, start=1)
    ]

    exc_cat_counts: dict[str, int] = {cat.value: 0 for cat in ExceptionCategory}
    for exc in exceptions:
        exc_cat_counts[exc.category.value] = exc_cat_counts.get(exc.category.value, 0) + 1

    summary = SummaryMetrics(
        total_po_rows=matched_po_df.height,
        unique_valid_po_lines=unique_valid_po_lines,
        duplicate_po_rows=duplicate_po_rows,
        invalid_po_rows=invalid_po_rows,
        reconciled_lines=reconciled_lines,
        partial_fulfillment_lines=partial_fulfillment_lines,
        po_exception_lines=po_exception_lines,
        total_exception_records=len(exceptions),
        matching_coverage_rate=coverage_rate,
        reconciliation_rate=recon_rate,
        total_ordered_qty=total_ordered_qty,
        total_shipped_qty=total_shipped_qty,
        total_invoiced_qty=total_invoiced_qty,
        total_outstanding_qty=total_outstanding_qty,
        total_expected_amount=cfg.quantize_money(total_expected_amount),
        total_invoiced_amount=cfg.quantize_money(total_invoiced_amount),
        net_billed_variance=cfg.quantize_money(net_billed_variance),
        absolute_monetary_discrepancy=cfg.quantize_money(absolute_monetary_discrepancy),
        unbilled_shipped_amount=cfg.quantize_money(unbilled_shipped_amount),
        quarantined_invoice_amount=quarantined_invoice_amount,
        exception_counts_by_category=exc_cat_counts,
        match_counts_by_status=dict(match_status_counter),
    )

    effective_summaries = (
        products_clean.summary,
        po_clean.summary,
        effective_ship_summary,
        effective_inv_summary,
    )
    for ds_sum in effective_summaries:
        if ds_sum.broken_formula_rows > 0:
            warnings.append(
                f"{ds_sum.dataset_kind.value}: detected {ds_sum.broken_formula_rows} row(s) "
                "containing broken or unpermitted Excel formulas."
            )
        if ds_sum.duplicate_rows > 0:
            warnings.append(
                f"{ds_sum.dataset_kind.value}: isolated {ds_sum.duplicate_rows} duplicate row(s)."
            )
        if ds_sum.conflicting_duplicate_rows > 0:
            warnings.append(
                f"{ds_sum.dataset_kind.value}: quarantined {ds_sum.conflicting_duplicate_rows} "
                "conflicting duplicate row(s)."
            )

    audit_log = AuditLogBundle(
        generated_at_utc=dt.datetime.now(dt.UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        python_version=platform.python_version(),
        polars_version=getattr(pl, "__version__", "unknown"),
        openpyxl_version=getattr(openpyxl, "__version__", "unknown"),
        config=cfg,
        input_files=tuple(item.metadata for item in input_metadata),
        schema_validations=(
            products_clean.schema_validation,
            po_clean.schema_validation,
            shipments_clean.schema_validation,
            invoices_clean.schema_validation,
        ),
        cleaning_summaries=effective_summaries,
        matching_rules_applied=MATCHING_RULES_DESCRIPTIONS,
        warnings=tuple(warnings),
    )

    return ReconciliationResult(
        summary=summary,
        reconciled_records=tuple(reconciled_records),
        exceptions=tuple(exceptions),
        audit_log=audit_log,
    )


def run_reconciliation_from_directory(
    input_dir: Path,
    *,
    config: ReconciliationConfig | None = None,
) -> ReconciliationResult:
    """Load all four Excel workbooks from input_dir, clean, match, and reconcile."""
    cfg = config or ReconciliationConfig()
    file_map = {
        DatasetKind.PRODUCTS: input_dir / "product_master.xlsx",
        DatasetKind.PURCHASE_ORDERS: input_dir / "purchase_orders.xlsx",
        DatasetKind.SHIPMENTS: input_dir / "shipments.xlsx",
        DatasetKind.INVOICES: input_dir / "supplier_invoices.xlsx",
    }

    loaded_products = read_excel_dataset(file_map[DatasetKind.PRODUCTS], DatasetKind.PRODUCTS)
    loaded_pos = read_excel_dataset(
        file_map[DatasetKind.PURCHASE_ORDERS], DatasetKind.PURCHASE_ORDERS
    )
    loaded_shipments = read_excel_dataset(file_map[DatasetKind.SHIPMENTS], DatasetKind.SHIPMENTS)
    loaded_invoices = read_excel_dataset(file_map[DatasetKind.INVOICES], DatasetKind.INVOICES)

    products_clean = clean_products_df(loaded_products.dataframe, loaded_products.observed_headers)
    po_clean = clean_purchase_orders_df(
        loaded_pos.dataframe,
        loaded_pos.observed_headers,
        reporting_currency=cfg.reporting_currency,
    )
    shipments_clean = clean_shipments_df(
        loaded_shipments.dataframe, loaded_shipments.observed_headers
    )
    invoices_clean = clean_invoices_df(
        loaded_invoices.dataframe,
        loaded_invoices.observed_headers,
        reporting_currency=cfg.reporting_currency,
    )

    return reconcile_datasets(
        products_clean=products_clean,
        po_clean=po_clean,
        shipments_clean=shipments_clean,
        invoices_clean=invoices_clean,
        config=cfg,
        input_metadata=(loaded_products, loaded_pos, loaded_shipments, loaded_invoices),
    )
