"""Deterministic multi-tier record matching and post-match semantic deduplication."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, replace
from decimal import Decimal
from typing import Any

import polars as pl

from excel_reconciliation.models import (
    CleaningSummary,
    InvoiceLineKey,
    MatchStatus,
    PoLineKey,
    ReasonCode,
    ShipmentLineKey,
    format_business_key,
)

MATCHING_RULES_DESCRIPTIONS: tuple[str, ...] = (
    (
        "Tier 1 (EXACT_ID): Normalized canonical SKU + exact normalized "
        "(po_number, po_line_id) with supplier verification."
    ),
    "Tier 2 (ALIAS_MAP): Explicit single-target product alias mapping to canonical product_id.",
    (
        "Tier 3 (COMPOSITE_KEY): Unique composite (po_number, canonical_product_id) "
        "when po_line_id is absent or non-canonical."
    ),
    (
        "Tier 4 (AMBIGUOUS / UNMATCHED): Quarantine multi-candidate aliases, "
        "conflicting catalog items, supplier mismatches, multi-line POs without po_line_id, "
        "product conflicts, and unknown refs."
    ),
)


@dataclass(frozen=True, slots=True)
class ProductEntry:
    """Canonical product record indexed from cleaned Product Master."""

    product_id: str
    canonical_sku: str
    product_name: str
    supplier_id: str | None


@dataclass(frozen=True, slots=True)
class SkuResolution:
    """Result of resolving a normalized SKU token against Product Master."""

    status: MatchStatus
    reason: ReasonCode
    product: ProductEntry | None
    candidate_product_ids: tuple[str, ...] = ()


class ProductMatcher:
    """Auditable SKU and explicit alias resolver backed by cleaned Product Master."""

    def __init__(self, cleaned_products_df: pl.DataFrame) -> None:
        canonical_buckets: dict[str, dict[str, ProductEntry]] = defaultdict(dict)
        alias_buckets: dict[str, dict[str, ProductEntry]] = defaultdict(dict)
        conflicted_canonical: dict[str, set[str]] = defaultdict(set)
        conflicted_aliases: dict[str, set[str]] = defaultdict(set)

        if not cleaned_products_df.is_empty():
            has_conflict_col = "_is_conflicting_duplicate" in cleaned_products_df.columns
            valid_cond = pl.col("_is_valid_record") & ~pl.col("_is_duplicate")
            if has_conflict_col:
                valid_cond = valid_cond & ~pl.col("_is_conflicting_duplicate")

            valid_df = cleaned_products_df.filter(valid_cond)
            for row in valid_df.to_dicts():
                prod_id: str = row["norm_product_id"]
                canon_sku: str = row["norm_canonical_sku"]
                entry = ProductEntry(
                    product_id=prod_id,
                    canonical_sku=canon_sku,
                    product_name=row["norm_product_name"] or canon_sku,
                    supplier_id=row["norm_supplier_id"],
                )
                canonical_buckets[canon_sku][prod_id] = entry
                for alias_token in row["norm_aliases"] or []:
                    if alias_token != canon_sku:
                        alias_buckets[alias_token][prod_id] = entry

            if has_conflict_col:
                conflict_df = cleaned_products_df.filter(pl.col("_is_conflicting_duplicate"))
                for row in conflict_df.to_dicts():
                    c_pid = row["norm_product_id"] or f"ROW-{row['_source_row']}"
                    c_sku = row["norm_canonical_sku"]
                    if c_sku is not None:
                        conflicted_canonical[c_sku].add(c_pid)
                    for alias_token in row["norm_aliases"] or []:
                        if alias_token != c_sku:
                            conflicted_aliases[alias_token].add(c_pid)

        self._canonical_index: dict[str, tuple[ProductEntry, ...]] = {
            sku: tuple(entries.values()) for sku, entries in canonical_buckets.items()
        }
        self._alias_index: dict[str, tuple[ProductEntry, ...]] = {
            alias: tuple(entries.values()) for alias, entries in alias_buckets.items()
        }
        self._conflicted_canonical: dict[str, tuple[str, ...]] = {
            sku: tuple(sorted(pids)) for sku, pids in conflicted_canonical.items()
        }
        self._conflicted_aliases: dict[str, tuple[str, ...]] = {
            alias: tuple(sorted(pids)) for alias, pids in conflicted_aliases.items()
        }

    def resolve_sku(self, norm_sku: str | None) -> SkuResolution:
        """Resolve a normalized SKU string using strict Tier 1 -> Tier 2 -> Tier 4 priority."""
        if norm_sku is None:
            return SkuResolution(
                status=MatchStatus.INVALID_KEY,
                reason=ReasonCode.MISSING_SKU,
                product=None,
            )

        if norm_sku in self._conflicted_canonical:
            valid_cands = tuple(e.product_id for e in self._canonical_index.get(norm_sku, ()))
            all_cands = tuple(dict.fromkeys((*self._conflicted_canonical[norm_sku], *valid_cands)))
            return SkuResolution(
                status=MatchStatus.AMBIGUOUS,
                reason=ReasonCode.AMBIGUOUS_MASTER_SKU,
                product=None,
                candidate_product_ids=all_cands,
            )

        if norm_sku in self._canonical_index:
            candidates = self._canonical_index[norm_sku]
            if len(candidates) == 1:
                return SkuResolution(
                    status=MatchStatus.EXACT_ID,
                    reason=ReasonCode.MATCHED_EXACT_ID,
                    product=candidates[0],
                    candidate_product_ids=(candidates[0].product_id,),
                )
            return SkuResolution(
                status=MatchStatus.AMBIGUOUS,
                reason=ReasonCode.AMBIGUOUS_MASTER_SKU,
                product=None,
                candidate_product_ids=tuple(c.product_id for c in candidates),
            )

        if norm_sku in self._conflicted_aliases:
            valid_cands = tuple(e.product_id for e in self._alias_index.get(norm_sku, ()))
            all_cands = tuple(dict.fromkeys((*self._conflicted_aliases[norm_sku], *valid_cands)))
            return SkuResolution(
                status=MatchStatus.AMBIGUOUS,
                reason=ReasonCode.AMBIGUOUS_ALIAS_MAPPING,
                product=None,
                candidate_product_ids=all_cands,
            )

        if norm_sku in self._alias_index:
            candidates = self._alias_index[norm_sku]
            if len(candidates) == 1:
                return SkuResolution(
                    status=MatchStatus.ALIAS_MAP,
                    reason=ReasonCode.MATCHED_PRODUCT_ALIAS,
                    product=candidates[0],
                    candidate_product_ids=(candidates[0].product_id,),
                )
            return SkuResolution(
                status=MatchStatus.AMBIGUOUS,
                reason=ReasonCode.AMBIGUOUS_ALIAS_MAPPING,
                product=None,
                candidate_product_ids=tuple(c.product_id for c in candidates),
            )

        return SkuResolution(
            status=MatchStatus.UNMATCHED,
            reason=ReasonCode.UNKNOWN_SKU,
            product=None,
        )


def match_purchase_orders_df(
    cleaned_po_df: pl.DataFrame,
    product_matcher: ProductMatcher,
) -> pl.DataFrame:
    """Resolve product SKUs on Purchase Order lines and annotate matching metadata."""
    matched_prod_ids: list[str | None] = []
    canonical_skus: list[str | None] = []
    canonical_names: list[str | None] = []
    match_statuses: list[str] = []
    match_reasons: list[str] = []
    business_keys: list[str] = []

    for row in cleaned_po_df.to_dicts():
        res = product_matcher.resolve_sku(row["norm_sku"])
        po_num = row["norm_po_number"] or "MISSING-PO"
        po_line = row["norm_po_line_id"] or f"ROW-{row['_source_row']}"

        if res.product is not None:
            matched_prod_ids.append(res.product.product_id)
            canonical_skus.append(res.product.canonical_sku)
            canonical_names.append(res.product.product_name)
            bkey = format_business_key(po_num, po_line, res.product.product_id)
        else:
            matched_prod_ids.append(None)
            canonical_skus.append(None)
            canonical_names.append(None)
            bkey = format_business_key(po_num, po_line, "UNRESOLVED")

        match_statuses.append(res.status.value)
        match_reasons.append(res.reason.value)
        business_keys.append(bkey)

    return cleaned_po_df.with_columns(
        pl.Series("matched_product_id", matched_prod_ids, dtype=pl.String),
        pl.Series("canonical_sku", canonical_skus, dtype=pl.String),
        pl.Series("canonical_product_name", canonical_names, dtype=pl.String),
        pl.Series("product_match_status", match_statuses, dtype=pl.String),
        pl.Series("product_match_reason", match_reasons, dtype=pl.String),
        pl.Series("matched_business_key", business_keys, dtype=pl.String),
    )


@dataclass(frozen=True, slots=True)
class IndexedPoLine:
    """Lookup projection of a valid, unique Purchase Order line."""

    key: PoLineKey
    product_id: str | None
    canonical_sku: str | None
    norm_sku: str
    supplier_id: str
    business_key: str


def _cand_structs(entries: list[IndexedPoLine] | tuple[IndexedPoLine, ...]) -> list[dict[str, str]]:
    """Serialize candidate PO line keys into typed Polars structs without delimiter flattening."""
    return [
        {"po_number": entry.key.po_number, "po_line_id": entry.key.po_line_id} for entry in entries
    ]


class OrderLineMatcher:
    """Matches Shipment and Supplier Invoice rows to valid Purchase Order lines."""

    def __init__(
        self,
        matched_po_df: pl.DataFrame,
        product_matcher: ProductMatcher,
    ) -> None:
        self._product_matcher = product_matcher
        self._known_pos: set[str] = set()
        self._suppliers_for_po: dict[str, set[str]] = defaultdict(set)
        self._by_exact_line: dict[PoLineKey, IndexedPoLine] = {}
        composite_buckets: dict[tuple[str, str], list[IndexedPoLine]] = defaultdict(list)

        if not matched_po_df.is_empty():
            valid_cond = pl.col("_is_valid_record") & ~pl.col("_is_duplicate")
            if "_is_conflicting_duplicate" in matched_po_df.columns:
                valid_cond = valid_cond & ~pl.col("_is_conflicting_duplicate")

            valid_po_df = matched_po_df.filter(valid_cond)
            for row in valid_po_df.to_dicts():
                po_num: str = row["norm_po_number"]
                po_line_id: str = row["norm_po_line_id"]
                supplier_id: str = row["norm_supplier_id"] or ""
                line_key = PoLineKey(po_number=po_num, po_line_id=po_line_id)
                prod_id: str | None = row["matched_product_id"]
                entry = IndexedPoLine(
                    key=line_key,
                    product_id=prod_id,
                    canonical_sku=row["canonical_sku"],
                    norm_sku=row["norm_sku"],
                    supplier_id=supplier_id,
                    business_key=row["matched_business_key"],
                )
                self._known_pos.add(po_num)
                if supplier_id:
                    self._suppliers_for_po[po_num].add(supplier_id)
                self._by_exact_line[line_key] = entry
                if prod_id is not None:
                    composite_buckets[(po_num, prod_id)].append(entry)

        self._by_composite: dict[tuple[str, str], tuple[IndexedPoLine, ...]] = {
            k: tuple(v) for k, v in composite_buckets.items()
        }

    def match_transactions_df(
        self,
        cleaned_tx_df: pl.DataFrame,
        *,
        validate_supplier: bool = False,
    ) -> pl.DataFrame:
        """Match Shipment or Invoice records to PO lines without arbitrary tie-breaking.

        When validate_supplier=True (Supplier Invoices), supplier_id on the transaction must
        match the supplier_id of the target Purchase Order line; mismatches are rejected with
        MatchStatus.UNMATCHED and ReasonCode.SUPPLIER_MISMATCH.
        """
        matched_po_numbers: list[str | None] = []
        matched_po_line_ids: list[str | None] = []
        matched_bkeys: list[str | None] = []
        matched_prod_ids: list[str | None] = []
        canonical_skus: list[str | None] = []
        match_statuses: list[str] = []
        match_reasons: list[str] = []
        candidate_po_lines_col: list[list[dict[str, str]]] = []

        for row in cleaned_tx_df.to_dicts():
            po_num: str | None = row["norm_po_number"]
            po_line_id: str | None = row["norm_po_line_id"]
            norm_sku: str | None = row["norm_sku"]
            tx_supplier: str | None = row.get("norm_supplier_id")

            sku_res = self._product_matcher.resolve_sku(norm_sku)
            prod_id = sku_res.product.product_id if sku_res.product else None
            canon_sku = sku_res.product.canonical_sku if sku_res.product else None

            matched_prod_ids.append(prod_id)
            canonical_skus.append(canon_sku)

            def _record_no_match(
                status: MatchStatus,
                reason: ReasonCode,
                cands: list[dict[str, str]] | None = None,
            ) -> None:
                matched_po_numbers.append(None)
                matched_po_line_ids.append(None)
                matched_bkeys.append(None)
                match_statuses.append(status.value)
                match_reasons.append(reason.value)
                candidate_po_lines_col.append(cands or [])

            # 1. Missing or unknown PO reference -> UNMATCHED
            if po_num is None:
                _record_no_match(MatchStatus.UNMATCHED, ReasonCode.MISSING_ORDER_REFERENCE)
                continue

            if po_num not in self._known_pos:
                _record_no_match(MatchStatus.UNMATCHED, ReasonCode.UNKNOWN_PO_NUMBER)
                continue

            # 1b. Supplier-level validation before matching candidate PO lines
            if validate_supplier and (
                tx_supplier is None or tx_supplier not in self._suppliers_for_po.get(po_num, set())
            ):
                _record_no_match(MatchStatus.UNMATCHED, ReasonCode.SUPPLIER_MISMATCH)
                continue

            # 2. Ambiguous SKU / Alias -> quarantine immediately; mark affected candidate PO lines
            if sku_res.status == MatchStatus.AMBIGUOUS:
                seen_keys: set[PoLineKey] = set()
                affected_entries: list[IndexedPoLine] = []
                for cand_pid in sku_res.candidate_product_ids:
                    for line_entry in self._by_composite.get((po_num, cand_pid), ()):
                        if validate_supplier and line_entry.supplier_id != tx_supplier:
                            continue
                        if line_entry.key not in seen_keys:
                            seen_keys.add(line_entry.key)
                            affected_entries.append(line_entry)
                if not affected_entries and po_line_id is not None:
                    exact_key = PoLineKey(po_number=po_num, po_line_id=po_line_id)
                    if exact_key in self._by_exact_line:
                        cand_entry = self._by_exact_line[exact_key]
                        if not validate_supplier or cand_entry.supplier_id == tx_supplier:
                            affected_entries.append(cand_entry)

                _record_no_match(
                    MatchStatus.AMBIGUOUS,
                    sku_res.reason,
                    _cand_structs(affected_entries),
                )
                continue

            # 3. Tier 1 / Tier 2: Direct normalized (po_number, po_line_id) lookup
            if po_line_id is not None:
                exact_key = PoLineKey(po_number=po_num, po_line_id=po_line_id)
                if exact_key in self._by_exact_line:
                    target_line = self._by_exact_line[exact_key]
                    if validate_supplier and (
                        tx_supplier is None or tx_supplier != target_line.supplier_id
                    ):
                        _record_no_match(MatchStatus.UNMATCHED, ReasonCode.SUPPLIER_MISMATCH)
                        continue

                    if target_line.product_id is None:
                        _record_no_match(MatchStatus.UNMATCHED, sku_res.reason)
                        continue

                    if prod_id is not None and prod_id == target_line.product_id:
                        matched_po_numbers.append(target_line.key.po_number)
                        matched_po_line_ids.append(target_line.key.po_line_id)
                        matched_bkeys.append(target_line.business_key)
                        if sku_res.status == MatchStatus.ALIAS_MAP:
                            match_statuses.append(MatchStatus.ALIAS_MAP.value)
                            match_reasons.append(ReasonCode.MATCHED_PRODUCT_ALIAS.value)
                        else:
                            match_statuses.append(MatchStatus.EXACT_ID.value)
                            match_reasons.append(ReasonCode.MATCHED_EXACT_ID.value)
                        candidate_po_lines_col.append(_cand_structs([target_line]))
                    else:
                        _record_no_match(
                            MatchStatus.AMBIGUOUS,
                            ReasonCode.LINE_PRODUCT_CONFLICT,
                            _cand_structs([target_line]),
                        )
                    continue

            # 4. Tier 3: Composite Business-Key Match (po_number, canonical_product_id)
            if prod_id is None:
                _record_no_match(sku_res.status, sku_res.reason)
                continue

            all_candidates = self._by_composite.get((po_num, prod_id), ())
            if validate_supplier:
                supplier_candidates = tuple(
                    c
                    for c in all_candidates
                    if tx_supplier is not None and c.supplier_id == tx_supplier
                )
                if all_candidates and not supplier_candidates:
                    _record_no_match(MatchStatus.UNMATCHED, ReasonCode.SUPPLIER_MISMATCH)
                    continue
                candidates = supplier_candidates
            else:
                candidates = all_candidates

            if len(candidates) == 1:
                only = candidates[0]
                matched_po_numbers.append(only.key.po_number)
                matched_po_line_ids.append(only.key.po_line_id)
                matched_bkeys.append(only.business_key)
                match_statuses.append(MatchStatus.COMPOSITE_KEY.value)
                match_reasons.append(ReasonCode.MATCHED_COMPOSITE_KEY.value)
                candidate_po_lines_col.append(_cand_structs([only]))
            elif len(candidates) > 1:
                _record_no_match(
                    MatchStatus.AMBIGUOUS,
                    ReasonCode.MULTIPLE_CANDIDATE_PO_LINES,
                    _cand_structs(candidates),
                )
            else:
                _record_no_match(MatchStatus.UNMATCHED, ReasonCode.PRODUCT_NOT_ON_PO)

        cand_struct_dtype = pl.List(
            pl.Struct([pl.Field("po_number", pl.String), pl.Field("po_line_id", pl.String)])
        )
        return cleaned_tx_df.with_columns(
            pl.Series("matched_po_number", matched_po_numbers, dtype=pl.String),
            pl.Series("matched_po_line_id", matched_po_line_ids, dtype=pl.String),
            pl.Series("matched_business_key", matched_bkeys, dtype=pl.String),
            pl.Series("matched_product_id", matched_prod_ids, dtype=pl.String),
            pl.Series("canonical_sku", canonical_skus, dtype=pl.String),
            pl.Series("match_status", match_statuses, dtype=pl.String),
            pl.Series("match_reason", match_reasons, dtype=pl.String),
            pl.Series("candidate_po_lines", candidate_po_lines_col, dtype=cand_struct_dtype),
        )


def _norm_dec_token(norm_val: str | None, raw_val: Any) -> tuple[str, str]:
    """Compare Decimal values numerically ('10.00' == '10') with raw fallback."""
    if norm_val is not None:
        dec = Decimal(norm_val).normalize()
        return ("DEC", "0") if dec == Decimal("0") else ("DEC", format(dec, "f"))
    return ("RAW", str(raw_val or ""))


def _shipment_intrinsic_valid(row: dict[str, Any]) -> bool:
    """Return True if non-duplicate shipment row passed schema/numeric/date validation."""
    codes = set(row.get("_validation_codes") or [])
    non_dup_blockers = {
        ReasonCode.BROKEN_EXCEL_FORMULA.value,
        ReasonCode.UNPERMITTED_EXCEL_FORMULA.value,
        ReasonCode.INVALID_NUMERIC_VALUE.value,
        ReasonCode.NEGATIVE_QUANTITY_OR_PRICE.value,
        ReasonCode.INVALID_DATE_VALUE.value,
        ReasonCode.MISSING_SKU.value,
    }
    return (
        row["norm_shipment_id"] is not None
        and row["norm_sku"] is not None
        and row["norm_ship_date"] is not None
        and row["norm_shipped_qty"] is not None
        and not row["_has_broken_formula"]
        and not row["_has_unpermitted_formula"]
        and not bool(codes & non_dup_blockers)
    )


def _invoice_intrinsic_valid(row: dict[str, Any]) -> bool:
    """Return True if non-duplicate invoice row passed schema/currency/numeric/date checks."""
    codes = set(row.get("_validation_codes") or [])
    non_dup_blockers = {
        ReasonCode.BROKEN_EXCEL_FORMULA.value,
        ReasonCode.UNPERMITTED_EXCEL_FORMULA.value,
        ReasonCode.CURRENCY_MISMATCH.value,
        ReasonCode.INVALID_NUMERIC_VALUE.value,
        ReasonCode.NEGATIVE_QUANTITY_OR_PRICE.value,
        ReasonCode.INVALID_DATE_VALUE.value,
        ReasonCode.MISSING_SKU.value,
    }
    return (
        row["norm_invoice_id"] is not None
        and row["norm_invoice_line_id"] is not None
        and row["norm_supplier_id"] is not None
        and row["norm_sku"] is not None
        and row["norm_invoice_date"] is not None
        and row["norm_invoiced_qty"] is not None
        and row["norm_unit_price"] is not None
        and row["norm_line_amount"] is not None
        and not row["_has_broken_formula"]
        and not row["_has_unpermitted_formula"]
        and not bool(codes & non_dup_blockers)
    )


def deduplicate_matched_shipments(
    matched_ship_df: pl.DataFrame,
    base_summary: CleaningSummary,
) -> tuple[pl.DataFrame, CleaningSummary]:
    """Re-evaluate shipment duplicates after SKU alias and PO-line resolution.

    - Preserves same shipment_id across different PO lines or different split shipment_ids.
    - Detects canonical SKU vs alias duplicates pointing to the same matched product & order line.
    - Quarantines ALL rows in conflicting duplicate groups rather than keeping an arbitrary row.
    """
    if matched_ship_df.is_empty():
        return matched_ship_df, base_summary

    rows = matched_ship_df.to_dicts()
    n = len(rows)
    is_dup = [False] * n
    dup_of: list[int | None] = [None] * n
    is_conflict = [False] * n

    groups: dict[ShipmentLineKey, list[int]] = defaultdict(list)
    for idx, row in enumerate(rows):
        sid = row["norm_shipment_id"]
        prod_key = row["matched_product_id"] or row["norm_sku"]
        if sid is None or prod_key is None:
            continue

        if row["matched_po_number"] is not None and row["matched_po_line_id"] is not None:
            scope_kind = "MATCHED_LINE"
            scope_po = str(row["matched_po_number"])
            scope_line = str(row["matched_po_line_id"])
        elif row["norm_po_number"] is not None or row["norm_po_line_id"] is not None:
            scope_kind = "RAW_REF"
            scope_po = str(row["norm_po_number"] or "")
            scope_line = str(row["norm_po_line_id"] or "")
        else:
            continue

        skey = ShipmentLineKey(
            shipment_id=sid,
            order_scope_kind=scope_kind,
            po_number=scope_po,
            po_line_id=scope_line,
            product_key=prod_key,
        )
        groups[skey].append(idx)

    for indices in groups.values():
        if len(indices) <= 1:
            continue
        sorted_indices = sorted(indices, key=lambda i: int(rows[i]["_source_row"]))
        payloads = {
            (
                _norm_dec_token(rows[i]["norm_shipped_qty"], rows[i]["raw_shipped_qty"]),
                rows[i]["norm_ship_date"] or str(rows[i]["raw_ship_date"] or ""),
                rows[i]["norm_carrier"],
                rows[i]["norm_tracking_ref"],
                rows[i]["norm_warehouse_code"],
            )
            for i in sorted_indices
        }
        if len(payloads) > 1:
            for i in sorted_indices:
                is_conflict[i] = True
        else:
            first_row_num = int(rows[sorted_indices[0]]["_source_row"])
            for i in sorted_indices[1:]:
                is_dup[i] = True
                dup_of[i] = first_row_num

    updated_valid: list[bool] = []
    updated_codes: list[list[str]] = []
    for idx, row in enumerate(rows):
        codes = [
            c
            for c in (row["_validation_codes"] or [])
            if c
            not in (
                ReasonCode.DUPLICATE_SHIPMENT.value,
                ReasonCode.CONFLICTING_DUPLICATE_SHIPMENT.value,
            )
        ]
        if is_conflict[idx]:
            codes.append(ReasonCode.CONFLICTING_DUPLICATE_SHIPMENT.value)
            updated_valid.append(False)
        else:
            if is_dup[idx]:
                codes.append(ReasonCode.DUPLICATE_SHIPMENT.value)
            updated_valid.append(_shipment_intrinsic_valid(row))
        updated_codes.append(codes)

    updated_df = matched_ship_df.with_columns(
        pl.Series("_is_duplicate", is_dup, dtype=pl.Boolean),
        pl.Series("_duplicate_of_row", dup_of, dtype=pl.Int64),
        pl.Series("_is_conflicting_duplicate", is_conflict, dtype=pl.Boolean),
        pl.Series("_is_valid_record", updated_valid, dtype=pl.Boolean),
        pl.Series("_validation_codes", updated_codes, dtype=pl.List(pl.String)),
    )

    dup_count = sum(1 for flag in is_dup if flag)
    conflict_count = sum(1 for flag in is_conflict if flag)
    valid_unique = int(
        updated_df.select(
            (
                pl.col("_is_valid_record")
                & ~pl.col("_is_duplicate")
                & ~pl.col("_is_conflicting_duplicate")
                & pl.col("norm_po_number").is_not_null()
            ).sum()
        ).item()
        or 0
    )
    updated_summary = replace(
        base_summary,
        valid_unique_rows=valid_unique,
        duplicate_rows=dup_count,
        conflicting_duplicate_rows=conflict_count,
    )
    return updated_df, updated_summary


def deduplicate_matched_invoices(
    matched_inv_df: pl.DataFrame,
    base_summary: CleaningSummary,
) -> tuple[pl.DataFrame, CleaningSummary]:
    """Re-evaluate supplier-scoped invoice duplicates after SKU alias and PO-line resolution."""
    if matched_inv_df.is_empty():
        return matched_inv_df, base_summary

    rows = matched_inv_df.to_dicts()
    n = len(rows)
    is_dup = [False] * n
    dup_of: list[int | None] = [None] * n
    is_conflict = [False] * n

    groups: dict[InvoiceLineKey, list[int]] = defaultdict(list)
    for idx, row in enumerate(rows):
        sup_id = row["norm_supplier_id"]
        inv_id = row["norm_invoice_id"]
        inv_line = row["norm_invoice_line_id"]
        if sup_id is not None and inv_id is not None and inv_line is not None:
            groups[InvoiceLineKey(sup_id, inv_id, inv_line)].append(idx)

    for indices in groups.values():
        if len(indices) <= 1:
            continue
        sorted_indices = sorted(indices, key=lambda i: int(rows[i]["_source_row"]))
        payloads = {
            (
                rows[i]["matched_po_number"] or rows[i]["norm_po_number"],
                rows[i]["matched_po_line_id"] or rows[i]["norm_po_line_id"],
                rows[i]["matched_product_id"] or rows[i]["norm_sku"],
                _norm_dec_token(rows[i]["norm_invoiced_qty"], rows[i]["raw_invoiced_qty"]),
                _norm_dec_token(rows[i]["norm_unit_price"], rows[i]["raw_unit_price"]),
                _norm_dec_token(rows[i]["norm_line_amount"], rows[i]["raw_line_amount"]),
                rows[i]["norm_currency"],
                rows[i]["norm_invoice_date"] or str(rows[i]["raw_invoice_date"] or ""),
            )
            for i in sorted_indices
        }
        if len(payloads) > 1:
            for i in sorted_indices:
                is_conflict[i] = True
        else:
            first_row_num = int(rows[sorted_indices[0]]["_source_row"])
            for i in sorted_indices[1:]:
                is_dup[i] = True
                dup_of[i] = first_row_num

    updated_valid: list[bool] = []
    updated_codes: list[list[str]] = []
    for idx, row in enumerate(rows):
        codes = [
            c
            for c in (row["_validation_codes"] or [])
            if c
            not in (
                ReasonCode.DUPLICATE_INVOICE.value,
                ReasonCode.CONFLICTING_DUPLICATE_INVOICE.value,
            )
        ]
        if is_conflict[idx]:
            codes.append(ReasonCode.CONFLICTING_DUPLICATE_INVOICE.value)
            updated_valid.append(False)
        else:
            if is_dup[idx]:
                codes.append(ReasonCode.DUPLICATE_INVOICE.value)
            updated_valid.append(_invoice_intrinsic_valid(row))
        updated_codes.append(codes)

    updated_df = matched_inv_df.with_columns(
        pl.Series("_is_duplicate", is_dup, dtype=pl.Boolean),
        pl.Series("_duplicate_of_row", dup_of, dtype=pl.Int64),
        pl.Series("_is_conflicting_duplicate", is_conflict, dtype=pl.Boolean),
        pl.Series("_is_valid_record", updated_valid, dtype=pl.Boolean),
        pl.Series("_validation_codes", updated_codes, dtype=pl.List(pl.String)),
    )

    dup_count = sum(1 for flag in is_dup if flag)
    conflict_count = sum(1 for flag in is_conflict if flag)
    valid_unique = int(
        updated_df.select(
            (
                pl.col("_is_valid_record")
                & ~pl.col("_is_duplicate")
                & ~pl.col("_is_conflicting_duplicate")
                & pl.col("norm_po_number").is_not_null()
            ).sum()
        ).item()
        or 0
    )
    updated_summary = replace(
        base_summary,
        valid_unique_rows=valid_unique,
        duplicate_rows=dup_count,
        conflicting_duplicate_rows=conflict_count,
    )
    return updated_df, updated_summary
