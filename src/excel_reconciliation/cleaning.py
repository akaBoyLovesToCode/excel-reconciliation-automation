"""Reusable data cleaning, normalization, schema validation, and deduplication."""

from __future__ import annotations

import datetime as dt
import re
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any

import polars as pl

from excel_reconciliation.models import (
    EXCEL_ERROR_TOKENS,
    MAX_DECIMAL_MAGNITUDE,
    REQUIRED_COLUMNS,
    CleaningSummary,
    DatasetKind,
    ReasonCode,
    SchemaValidationResult,
)

_ZERO_WIDTH_RE = re.compile(r"[\u200b\u200c\u200d\ufeff]")
_WHITESPACE_RE = re.compile(r"\s+")
_SKU_SEPARATOR_RE = re.compile(r"[\s._/\\-]+")
_ALIAS_SPLIT_RE = re.compile(r"[;|,]+")
_PAREN_DECIMAL_RE = re.compile(r"^\(\s*\$?\s*(?:(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?|\.\d+)\s*\)$")
_SIGNED_DECIMAL_RE = re.compile(r"^[+-]?\s*\$?\s*(?:(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?|\.\d+)$")
_DATE_FORMATS: tuple[str, ...] = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%m/%d/%Y",
    "%d.%m.%Y",
    "%Y-%m-%d %H:%M:%S",
)
_EN_MONTH_RE = re.compile(r"^(\d{1,2})-([A-Za-z]{3})-(\d{4})$")
_EN_MONTHS: dict[str, int] = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}
_MISSING_SENTINELS: frozenset[str] = frozenset({"", "NULL", "NONE", "NIL", "-", "--", "—"})


@dataclass(frozen=True, slots=True)
class CleanedDataset:
    """Cleaned Polars DataFrame paired with schema and hygiene diagnostics."""

    dataset_kind: DatasetKind
    dataframe: pl.DataFrame
    schema_validation: SchemaValidationResult
    summary: CleaningSummary


def normalize_string(value: str | None) -> str | None:
    """Normalize arbitrary text by removing zero-width chars and collapsing whitespace.

    Preserves Unicode characters (accents, CJK, symbols) while trimming outer spaces.
    Returns None for empty or sentinel-null inputs.
    """
    if value is None:
        return None
    cleaned = _ZERO_WIDTH_RE.sub("", value).strip()
    cleaned = _WHITESPACE_RE.sub(" ", cleaned)
    if cleaned.upper() in _MISSING_SENTINELS:
        return None
    return cleaned


def normalize_identifier(value: str | None) -> str | None:
    """Normalize business identifiers (PO numbers, line IDs, shipment/invoice IDs)."""
    base = normalize_string(value)
    if base is None:
        return None
    return base.upper()


def normalize_sku(value: str | None) -> str | None:
    """Canonicalize SKU codes across inconsistent casing, whitespace, and separators.

    Examples:
        "  sku_elec_0042 " -> "SKU-ELEC-0042"
        "SKU.ELEC.0042"    -> "SKU-ELEC-0042"
        "Sku Elec 0042"    -> "SKU-ELEC-0042"
    """
    if value is None:
        return None
    cleaned = _ZERO_WIDTH_RE.sub("", value).strip()
    if not cleaned or cleaned.upper() in _MISSING_SENTINELS:
        return None
    canonical = _SKU_SEPARATOR_RE.sub("-", cleaned.upper()).strip("-")
    return canonical if canonical else None


def split_and_normalize_aliases(value: str | None) -> tuple[str, ...]:
    """Parse a delimited alias string into deduplicated, normalized SKU alias tokens."""
    if value is None:
        return ()
    seen: dict[str, None] = {}
    for part in _ALIAS_SPLIT_RE.split(value):
        norm = normalize_sku(part)
        if norm is not None and norm not in seen:
            seen[norm] = None
    return tuple(seen.keys())


def parse_decimal_value(
    value: str | int | Decimal | None,
    *,
    required: bool = True,
    allow_negative: bool = False,
) -> tuple[Decimal | None, ReasonCode | None]:
    """Parse a raw scalar cell into a bounded exact Decimal without silent coercion.

    Rejects malformed strings ('TBD', '12O', '1_000', '1,2,3', '($-10)', '(+10)', '1e50'),
    binary floats/bools, non-finite values, and Excel error tokens (#REF!, #DIV/0!).
    Distinguishes missing values (None) from valid zero values (Decimal('0')).
    """
    if value is None:
        return (None, ReasonCode.INVALID_NUMERIC_VALUE if required else None)

    if isinstance(value, (bool, float)):
        return (None, ReasonCode.INVALID_NUMERIC_VALUE)

    if isinstance(value, Decimal):
        if not value.is_finite() or abs(value) > MAX_DECIMAL_MAGNITUDE:
            return (None, ReasonCode.INVALID_NUMERIC_VALUE)
        if not allow_negative and value < Decimal("0"):
            return (value, ReasonCode.NEGATIVE_QUANTITY_OR_PRICE)
        return (value, None)

    if isinstance(value, int):
        trimmed = str(value)
    elif isinstance(value, str):
        trimmed = _ZERO_WIDTH_RE.sub("", value).strip()
    else:
        return (None, ReasonCode.INVALID_NUMERIC_VALUE)

    if not trimmed or trimmed.upper() in _MISSING_SENTINELS:
        return (None, ReasonCode.INVALID_NUMERIC_VALUE if required else None)

    upper = trimmed.upper()
    if any(tok in upper for tok in EXCEL_ERROR_TOKENS):
        return (None, ReasonCode.BROKEN_EXCEL_FORMULA)

    is_paren_negative = False
    if _PAREN_DECIMAL_RE.match(trimmed):
        is_paren_negative = True
        core = trimmed[1:-1].strip()
        sign = "-"
    elif _SIGNED_DECIMAL_RE.match(trimmed):
        core = trimmed
        if core.startswith("-"):
            sign = "-"
            core = core[1:].strip()
        elif core.startswith("+"):
            sign = ""
            core = core[1:].strip()
        else:
            sign = ""
    else:
        return (None, ReasonCode.INVALID_NUMERIC_VALUE)

    core = core.removeprefix("$").strip().replace(",", "")
    if not core:
        return (None, ReasonCode.INVALID_NUMERIC_VALUE)

    try:
        parsed = Decimal(f"{sign}{core}")
    except InvalidOperation:
        return (None, ReasonCode.INVALID_NUMERIC_VALUE)

    if not parsed.is_finite() or abs(parsed) > MAX_DECIMAL_MAGNITUDE:
        return (None, ReasonCode.INVALID_NUMERIC_VALUE)

    if is_paren_negative and parsed > Decimal("0"):
        parsed = -parsed

    if not allow_negative and parsed < Decimal("0"):
        return (parsed, ReasonCode.NEGATIVE_QUANTITY_OR_PRICE)

    return (parsed, None)


def parse_date_value(
    value: str | None,
    *,
    required: bool = True,
) -> tuple[str | None, ReasonCode | None]:
    """Validate and normalize date strings into ISO-8601 YYYY-MM-DD format."""
    norm = normalize_string(value)
    if norm is None:
        return (None, ReasonCode.INVALID_DATE_VALUE if required else None)

    if any(tok in norm.upper() for tok in EXCEL_ERROR_TOKENS):
        return (None, ReasonCode.BROKEN_EXCEL_FORMULA)

    en_match = _EN_MONTH_RE.match(norm)
    if en_match:
        day_s, mon_s, year_s = en_match.groups()
        mon_num = _EN_MONTHS.get(mon_s.upper())
        if mon_num is not None:
            try:
                return (dt.date(int(year_s), mon_num, int(day_s)).isoformat(), None)
            except ValueError:
                return (None, ReasonCode.INVALID_DATE_VALUE)

    for fmt in _DATE_FORMATS:
        try:
            parsed_dt = dt.datetime.strptime(norm, fmt)
            return (parsed_dt.date().isoformat(), None)
        except ValueError:
            continue

    return (None, ReasonCode.INVALID_DATE_VALUE)


def validate_schema(
    dataset_kind: DatasetKind,
    observed_columns: tuple[str, ...] | list[str],
) -> SchemaValidationResult:
    """Validate workbook column headers against expected dataset schema."""
    expected = REQUIRED_COLUMNS[dataset_kind]
    observed_clean = tuple(
        col.strip().lower() for col in observed_columns if not col.startswith("_")
    )
    observed_set = set(observed_clean)
    expected_set = set(expected)

    missing = tuple(col for col in expected if col not in observed_set)
    extra = tuple(col for col in observed_clean if col not in expected_set)

    return SchemaValidationResult(
        dataset_kind=dataset_kind,
        is_valid=(len(missing) == 0),
        expected_columns=expected,
        observed_columns=observed_clean,
        missing_columns=missing,
        extra_columns=extra,
    )


def _polars_norm_str_expr(col_name: str) -> pl.Expr:
    """Vectorized Polars expression for whitespace and sentinel normalization."""
    cleaned = (
        pl.col(col_name)
        .cast(pl.String)
        .str.replace_all(r"[\u200b\u200c\u200d\ufeff]", "")
        .str.strip_chars()
        .str.replace_all(r"\s+", " ")
    )
    return (
        pl.when(cleaned.is_null() | cleaned.str.to_uppercase().is_in(list(_MISSING_SENTINELS)))
        .then(pl.lit(None, dtype=pl.String))
        .otherwise(cleaned)
    )


def _polars_norm_id_expr(col_name: str) -> pl.Expr:
    """Vectorized Polars expression for uppercase business key normalization."""
    return _polars_norm_str_expr(col_name).str.to_uppercase()


def _polars_norm_sku_expr(col_name: str) -> pl.Expr:
    """Vectorized Polars expression for SKU separator and casing canonicalization."""
    cleaned = (
        pl.col(col_name)
        .cast(pl.String)
        .str.replace_all(r"[\u200b\u200c\u200d\ufeff]", "")
        .str.strip_chars()
    )
    canonical = cleaned.str.to_uppercase().str.replace_all(r"[\s._/\\-]+", "-").str.strip_chars("-")
    return (
        pl.when(
            cleaned.is_null()
            | cleaned.str.to_uppercase().is_in(list(_MISSING_SENTINELS))
            | (canonical == "")
        )
        .then(pl.lit(None, dtype=pl.String))
        .otherwise(canonical)
    )


def _preserve_raw_columns(df: pl.DataFrame, columns: tuple[str, ...]) -> pl.DataFrame:
    """Ensure every expected column exists and preserve raw_<col> audit copies."""
    missing_exprs = [
        pl.lit(None, dtype=pl.String).alias(col) for col in columns if col not in df.columns
    ]
    if missing_exprs:
        df = df.with_columns(missing_exprs)

    if "_source_row" not in df.columns:
        df = df.with_row_index(name="_source_row", offset=2).with_columns(
            pl.col("_source_row").cast(pl.Int64)
        )
    if "_formulas_json" not in df.columns:
        df = df.with_columns(pl.lit(None, dtype=pl.String).alias("_formulas_json"))
    if "_has_broken_formula" not in df.columns:
        df = df.with_columns(pl.lit(False, dtype=pl.Boolean).alias("_has_broken_formula"))
    if "_has_unpermitted_formula" not in df.columns:
        df = df.with_columns(pl.lit(False, dtype=pl.Boolean).alias("_has_unpermitted_formula"))

    raw_copies = [pl.col(col).cast(pl.String).alias(f"raw_{col}") for col in columns]
    return df.with_columns(raw_copies)


def _dec_cmp_key(dec_val: Decimal | None, raw_val: Any) -> tuple[str, str]:
    """Canonical Decimal comparison token so '10.00' and '10' compare equal."""
    if dec_val is not None:
        normalized = dec_val.normalize()
        # Handle -0 / 0.00 uniformly
        if normalized == Decimal("0"):
            return ("DEC", "0")
        return ("DEC", format(normalized, "f"))
    return ("RAW", str(raw_val or ""))


def _classify_key_groups(
    rows: list[dict[str, Any]],
    key_fn: Any,
    payload_fn: Any,
) -> tuple[list[bool], list[int | None], list[bool]]:
    """Classify non-null key groups into unique firsts, identical duplicates, and conflicts.

    Conflicting groups (same identity key with differing business attributes) quarantine
    every row in the group rather than silently retaining one row.
    """
    n = len(rows)
    is_dup = [False] * n
    dup_of: list[int | None] = [None] * n
    is_conflict = [False] * n

    groups: dict[tuple[Any, ...], list[int]] = defaultdict(list)
    for idx, row in enumerate(rows):
        key = key_fn(row)
        if key is not None:
            groups[key].append(idx)

    for indices in groups.values():
        if len(indices) <= 1:
            continue
        sorted_indices = sorted(indices, key=lambda i: int(rows[i]["_source_row"]))
        payloads = {payload_fn(rows[i]) for i in sorted_indices}
        if len(payloads) > 1:
            for i in sorted_indices:
                is_conflict[i] = True
        else:
            first_row_num = int(rows[sorted_indices[0]]["_source_row"])
            for i in sorted_indices[1:]:
                is_dup[i] = True
                dup_of[i] = first_row_num

    return is_dup, dup_of, is_conflict


def _count_whitespace_fixes(df: pl.DataFrame, cols: list[str]) -> int:
    """Count cells across given columns that had leading/trailing or repeated whitespace."""
    if df.is_empty():
        return 0
    total = 0
    for col in cols:
        raw_c = f"raw_{col}"
        changed = df.select(
            (
                pl.col(raw_c).is_not_null()
                & (
                    (pl.col(raw_c) != pl.col(raw_c).str.strip_chars())
                    | pl.col(raw_c).str.contains(r"\s{2,}|\t|\n")
                )
            ).sum()
        ).item()
        total += int(changed or 0)
    return total


def _count_sku_fixes(df: pl.DataFrame, raw_col: str, norm_col: str) -> int:
    """Count rows where SKU normalization altered separator, case, or whitespace formatting."""
    if df.is_empty():
        return 0
    changed = df.select(
        (
            pl.col(raw_col).is_not_null()
            & pl.col(norm_col).is_not_null()
            & (pl.col(raw_col) != pl.col(norm_col))
        ).sum()
    ).item()
    return int(changed or 0)


def _append_unique_code(codes: list[str], code: ReasonCode | str) -> None:
    code_str = code.value if isinstance(code, ReasonCode) else str(code)
    if code_str not in codes:
        codes.append(code_str)


def clean_products_df(
    df: pl.DataFrame,
    observed_headers: tuple[str, ...] | list[str] | None = None,
) -> CleanedDataset:
    """Clean and validate Product Master without silently resolving conflicting records."""
    headers = observed_headers if observed_headers is not None else df.columns
    schema_res = validate_schema(DatasetKind.PRODUCTS, headers)
    df = _preserve_raw_columns(df, REQUIRED_COLUMNS[DatasetKind.PRODUCTS])

    df = df.with_columns(
        _polars_norm_id_expr("product_id").alias("norm_product_id"),
        _polars_norm_sku_expr("canonical_sku").alias("norm_canonical_sku"),
        _polars_norm_str_expr("product_name").alias("norm_product_name"),
        _polars_norm_str_expr("category").alias("norm_category"),
        _polars_norm_id_expr("uom").alias("norm_uom"),
        _polars_norm_id_expr("supplier_id").alias("norm_supplier_id"),
        _polars_norm_id_expr("active_flag").alias("norm_active_flag"),
    )

    rows = df.to_dicts()
    n = len(rows)

    parsed_costs: list[tuple[Decimal | None, ReasonCode | None]] = []
    aliases_per_row: list[list[str]] = []
    skus_by_pid: dict[str, set[str]] = defaultdict(set)
    pids_by_sku: dict[str, set[str]] = defaultdict(set)
    pair_indices: dict[tuple[str, str], list[int]] = defaultdict(list)

    for idx, row in enumerate(rows):
        cost_dec, cost_err = parse_decimal_value(row["raw_standard_cost"], required=True)
        parsed_costs.append((cost_dec, cost_err))
        aliases = list(split_and_normalize_aliases(row["raw_aliases"]))
        aliases_per_row.append(aliases)

        pid = row["norm_product_id"]
        csku = row["norm_canonical_sku"]
        if pid is not None and csku is not None:
            skus_by_pid[pid].add(csku)
            pids_by_sku[csku].add(pid)
            pair_indices[(pid, csku)].append(idx)

    # Identify conflicting pairs and identical duplicates deterministically
    is_dup = [False] * n
    dup_of: list[int | None] = [None] * n
    is_conflict = [False] * n
    sku_collision_conflict = [False] * n

    for idx, row in enumerate(rows):
        pid = row["norm_product_id"]
        csku = row["norm_canonical_sku"]
        if pid is not None and len(skus_by_pid[pid]) > 1:
            is_conflict[idx] = True
        if csku is not None and len(pids_by_sku[csku]) > 1:
            is_conflict[idx] = True
            sku_collision_conflict[idx] = True

    for indices in pair_indices.values():
        if len(indices) <= 1:
            continue
        sorted_indices = sorted(indices, key=lambda i: int(rows[i]["_source_row"]))
        if any(is_conflict[i] for i in sorted_indices):
            for i in sorted_indices:
                is_conflict[i] = True
            continue

        # Check semantic conflicts among rows sharing (product_id, canonical_sku)
        distinct_costs = {
            _dec_cmp_key(parsed_costs[i][0], rows[i]["raw_standard_cost"])
            for i in sorted_indices
            if parsed_costs[i][1] is None
        }
        distinct_sups = {
            rows[i]["norm_supplier_id"]
            for i in sorted_indices
            if rows[i]["norm_supplier_id"] is not None
        }
        distinct_nonempty_aliases = {
            frozenset(aliases_per_row[i]) for i in sorted_indices if aliases_per_row[i]
        }
        has_invalid_cost_mix = any(parsed_costs[i][1] is not None for i in sorted_indices) and any(
            parsed_costs[i][1] is None for i in sorted_indices
        )

        if (
            len(distinct_costs) > 1
            or len(distinct_sups) > 1
            or len(distinct_nonempty_aliases) > 1
            or has_invalid_cost_mix
        ):
            for i in sorted_indices:
                is_conflict[i] = True
        else:
            # Empty aliases are incomplete metadata, not evidence to discard known aliases.
            # All non-empty sets agree here, so enrich normalized rows without changing raw_*.
            merged_aliases = sorted({alias for i in sorted_indices for alias in aliases_per_row[i]})
            for i in sorted_indices:
                aliases_per_row[i] = merged_aliases.copy()
            first_row_num = int(rows[sorted_indices[0]]["_source_row"])
            for i in sorted_indices[1:]:
                is_dup[i] = True
                dup_of[i] = first_row_num

    norm_costs: list[str | None] = []
    valid_flags: list[bool] = []
    issue_codes_col: list[list[str]] = []

    missing_rows = 0
    invalid_num_rows = 0
    broken_formula_rows = 0

    for idx, row in enumerate(rows):
        codes: list[str] = []
        if row["_has_broken_formula"]:
            _append_unique_code(codes, ReasonCode.BROKEN_EXCEL_FORMULA)
            broken_formula_rows += 1
        if row["_has_unpermitted_formula"]:
            _append_unique_code(codes, ReasonCode.UNPERMITTED_EXCEL_FORMULA)
            broken_formula_rows += 1

        if row["norm_product_id"] is None or row["norm_canonical_sku"] is None:
            missing_rows += 1
            if row["norm_canonical_sku"] is None:
                _append_unique_code(codes, ReasonCode.MISSING_SKU)

        cost_dec, cost_err = parsed_costs[idx]
        norm_costs.append(str(cost_dec) if cost_dec is not None else None)
        if cost_err is not None:
            _append_unique_code(codes, cost_err)
            invalid_num_rows += 1

        if is_conflict[idx]:
            _append_unique_code(codes, ReasonCode.CONFLICTING_MASTER_RECORD)
            if sku_collision_conflict[idx]:
                _append_unique_code(codes, ReasonCode.AMBIGUOUS_MASTER_SKU)
        elif is_dup[idx]:
            _append_unique_code(codes, ReasonCode.DUPLICATE_PRODUCT_ROW)

        is_valid = (
            row["norm_product_id"] is not None
            and row["norm_canonical_sku"] is not None
            and cost_err is None
            and not row["_has_broken_formula"]
            and not row["_has_unpermitted_formula"]
            and not is_conflict[idx]
        )
        valid_flags.append(is_valid)
        issue_codes_col.append(codes)

    df = df.with_columns(
        pl.Series("_is_duplicate", is_dup, dtype=pl.Boolean),
        pl.Series("_duplicate_of_row", dup_of, dtype=pl.Int64),
        pl.Series("_is_conflicting_duplicate", is_conflict, dtype=pl.Boolean),
        pl.Series("norm_standard_cost", norm_costs, dtype=pl.String),
        pl.Series("norm_aliases", aliases_per_row, dtype=pl.List(pl.String)),
        pl.Series("_is_valid_record", valid_flags, dtype=pl.Boolean),
        pl.Series("_validation_codes", issue_codes_col, dtype=pl.List(pl.String)),
    )

    dup_count = (
        int(df.select(pl.col("_is_duplicate").sum()).item() or 0) if not df.is_empty() else 0
    )
    conflict_count = (
        int(df.select(pl.col("_is_conflicting_duplicate").sum()).item() or 0)
        if not df.is_empty()
        else 0
    )
    valid_unique = (
        int(
            df.select(
                (
                    pl.col("_is_valid_record")
                    & ~pl.col("_is_duplicate")
                    & ~pl.col("_is_conflicting_duplicate")
                ).sum()
            ).item()
            or 0
        )
        if not df.is_empty()
        else 0
    )

    summary = CleaningSummary(
        dataset_kind=DatasetKind.PRODUCTS,
        total_rows=df.height,
        valid_unique_rows=valid_unique,
        duplicate_rows=dup_count,
        conflicting_duplicate_rows=conflict_count,
        whitespace_normalized_count=_count_whitespace_fixes(
            df, ["product_id", "canonical_sku", "product_name", "supplier_id", "aliases"]
        ),
        sku_normalized_count=_count_sku_fixes(df, "raw_canonical_sku", "norm_canonical_sku"),
        missing_value_rows=missing_rows,
        invalid_numeric_rows=invalid_num_rows,
        invalid_date_rows=0,
        broken_formula_rows=broken_formula_rows,
    )
    return CleanedDataset(DatasetKind.PRODUCTS, df, schema_res, summary)


def clean_purchase_orders_df(
    df: pl.DataFrame,
    observed_headers: tuple[str, ...] | list[str] | None = None,
    *,
    reporting_currency: str = "USD",
) -> CleanedDataset:
    """Clean and validate Purchase Orders without discarding or conflating rows."""
    headers = observed_headers if observed_headers is not None else df.columns
    schema_res = validate_schema(DatasetKind.PURCHASE_ORDERS, headers)
    df = _preserve_raw_columns(df, REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS])
    req_currency = reporting_currency.strip().upper()

    df = df.with_columns(
        _polars_norm_id_expr("po_number").alias("norm_po_number"),
        _polars_norm_id_expr("po_line_id").alias("norm_po_line_id"),
        _polars_norm_id_expr("supplier_id").alias("norm_supplier_id"),
        _polars_norm_sku_expr("sku").alias("norm_sku"),
        _polars_norm_str_expr("product_description").alias("norm_product_description"),
        _polars_norm_id_expr("currency").alias("norm_currency"),
    )

    rows = df.to_dicts()
    n = len(rows)

    parsed_dates: list[tuple[str | None, ReasonCode | None]] = []
    parsed_qtys: list[tuple[Decimal | None, ReasonCode | None]] = []
    parsed_prices: list[tuple[Decimal | None, ReasonCode | None]] = []
    parsed_totals: list[tuple[Decimal | None, ReasonCode | None]] = []

    for row in rows:
        parsed_dates.append(parse_date_value(row["raw_order_date"], required=True))
        parsed_qtys.append(parse_decimal_value(row["raw_ordered_qty"], required=True))
        parsed_prices.append(parse_decimal_value(row["raw_unit_price"], required=True))
        parsed_totals.append(parse_decimal_value(row["raw_line_total"], required=False))

    enriched_rows = [
        {
            **row,
            "_parsed_qty": parsed_qtys[i][0],
            "_parsed_price": parsed_prices[i][0],
            "_parsed_total": parsed_totals[i][0],
            "_total_error": parsed_totals[i][1],
            "_parsed_date": parsed_dates[i][0],
        }
        for i, row in enumerate(rows)
    ]

    is_dup, dup_of, is_conflict = _classify_key_groups(
        enriched_rows,
        key_fn=lambda r: (
            (r["norm_po_number"], r["norm_po_line_id"])
            if r["norm_po_number"] is not None and r["norm_po_line_id"] is not None
            else None
        ),
        payload_fn=lambda r: (
            r["norm_supplier_id"],
            r["norm_sku"],
            _dec_cmp_key(r["_parsed_qty"], r["raw_ordered_qty"]),
            _dec_cmp_key(r["_parsed_price"], r["raw_unit_price"]),
            _dec_cmp_key(r["_parsed_total"], r["raw_line_total"]),
            r["_total_error"],
            bool(r["_has_broken_formula"]),
            bool(r["_has_unpermitted_formula"]),
            r["norm_currency"],
            r["_parsed_date"] or str(r["raw_order_date"] or ""),
        ),
    )

    norm_dates: list[str | None] = []
    norm_qtys: list[str | None] = []
    norm_prices: list[str | None] = []
    norm_totals: list[str | None] = []
    valid_flags: list[bool] = []
    issue_codes_col: list[list[str]] = []

    missing_rows = 0
    invalid_num_rows = 0
    invalid_date_rows = 0
    broken_formula_rows = 0

    for idx, row in enumerate(rows):
        codes: list[str] = []
        row_num_err = False

        if row["_has_broken_formula"]:
            _append_unique_code(codes, ReasonCode.BROKEN_EXCEL_FORMULA)
            broken_formula_rows += 1
        if row["_has_unpermitted_formula"]:
            _append_unique_code(codes, ReasonCode.UNPERMITTED_EXCEL_FORMULA)
            broken_formula_rows += 1

        if (
            row["norm_po_number"] is None
            or row["norm_po_line_id"] is None
            or row["norm_supplier_id"] is None
            or row["norm_sku"] is None
        ):
            missing_rows += 1
            if (
                row["norm_po_number"] is None
                or row["norm_po_line_id"] is None
                or row["norm_supplier_id"] is None
            ):
                _append_unique_code(codes, ReasonCode.MISSING_ORDER_REFERENCE)
            if row["norm_sku"] is None:
                _append_unique_code(codes, ReasonCode.MISSING_SKU)

        currency_ok = row["norm_currency"] == req_currency
        if not currency_ok:
            _append_unique_code(codes, ReasonCode.CURRENCY_MISMATCH)

        date_iso, date_err = parsed_dates[idx]
        norm_dates.append(date_iso)
        if date_err is not None:
            _append_unique_code(codes, date_err)
            invalid_date_rows += 1

        qty_dec, qty_err = parsed_qtys[idx]
        norm_qtys.append(str(qty_dec) if qty_dec is not None else None)
        if qty_err is not None:
            _append_unique_code(codes, qty_err)
            row_num_err = True

        price_dec, price_err = parsed_prices[idx]
        norm_prices.append(str(price_dec) if price_dec is not None else None)
        if price_err is not None:
            _append_unique_code(codes, price_err)
            row_num_err = True

        total_dec, total_err = parsed_totals[idx]
        norm_totals.append(str(total_dec) if total_dec is not None else None)
        if total_err is not None:
            _append_unique_code(codes, total_err)
            row_num_err = True

        if row_num_err:
            invalid_num_rows += 1

        if is_conflict[idx]:
            _append_unique_code(codes, ReasonCode.CONFLICTING_DUPLICATE_PO_LINE)
        elif is_dup[idx]:
            _append_unique_code(codes, ReasonCode.DUPLICATE_PO_LINE)

        is_valid = (
            row["norm_po_number"] is not None
            and row["norm_po_line_id"] is not None
            and row["norm_supplier_id"] is not None
            and row["norm_sku"] is not None
            and currency_ok
            and date_err is None
            and qty_err is None
            and price_err is None
            and total_err is None
            and not row["_has_broken_formula"]
            and not row["_has_unpermitted_formula"]
            and not is_conflict[idx]
        )
        valid_flags.append(is_valid)
        issue_codes_col.append(codes)

    df = df.with_columns(
        pl.Series("_is_duplicate", is_dup, dtype=pl.Boolean),
        pl.Series("_duplicate_of_row", dup_of, dtype=pl.Int64),
        pl.Series("_is_conflicting_duplicate", is_conflict, dtype=pl.Boolean),
        pl.Series("norm_order_date", norm_dates, dtype=pl.String),
        pl.Series("norm_ordered_qty", norm_qtys, dtype=pl.String),
        pl.Series("norm_unit_price", norm_prices, dtype=pl.String),
        pl.Series("norm_line_total", norm_totals, dtype=pl.String),
        pl.Series("_is_valid_record", valid_flags, dtype=pl.Boolean),
        pl.Series("_validation_codes", issue_codes_col, dtype=pl.List(pl.String)),
    )

    dup_count = sum(1 for flag in is_dup if flag) if n else 0
    conflict_count = sum(1 for flag in is_conflict if flag) if n else 0
    valid_unique = (
        int(
            df.select(
                (
                    pl.col("_is_valid_record")
                    & ~pl.col("_is_duplicate")
                    & ~pl.col("_is_conflicting_duplicate")
                ).sum()
            ).item()
            or 0
        )
        if not df.is_empty()
        else 0
    )

    summary = CleaningSummary(
        dataset_kind=DatasetKind.PURCHASE_ORDERS,
        total_rows=df.height,
        valid_unique_rows=valid_unique,
        duplicate_rows=dup_count,
        conflicting_duplicate_rows=conflict_count,
        whitespace_normalized_count=_count_whitespace_fixes(
            df, ["po_number", "po_line_id", "supplier_id", "sku", "product_description"]
        ),
        sku_normalized_count=_count_sku_fixes(df, "raw_sku", "norm_sku"),
        missing_value_rows=missing_rows,
        invalid_numeric_rows=invalid_num_rows,
        invalid_date_rows=invalid_date_rows,
        broken_formula_rows=broken_formula_rows,
    )
    return CleanedDataset(DatasetKind.PURCHASE_ORDERS, df, schema_res, summary)


def clean_shipments_df(
    df: pl.DataFrame,
    observed_headers: tuple[str, ...] | list[str] | None = None,
) -> CleanedDataset:
    """Clean and validate Shipments with order-line-scoped preliminary deduplication."""
    headers = observed_headers if observed_headers is not None else df.columns
    schema_res = validate_schema(DatasetKind.SHIPMENTS, headers)
    df = _preserve_raw_columns(df, REQUIRED_COLUMNS[DatasetKind.SHIPMENTS])

    df = df.with_columns(
        _polars_norm_id_expr("shipment_id").alias("norm_shipment_id"),
        _polars_norm_id_expr("po_number").alias("norm_po_number"),
        _polars_norm_id_expr("po_line_id").alias("norm_po_line_id"),
        _polars_norm_str_expr("carrier").alias("norm_carrier"),
        _polars_norm_id_expr("tracking_ref").alias("norm_tracking_ref"),
        _polars_norm_sku_expr("sku").alias("norm_sku"),
        _polars_norm_id_expr("warehouse_code").alias("norm_warehouse_code"),
    )

    rows = df.to_dicts()
    n = len(rows)

    parsed_dates: list[tuple[str | None, ReasonCode | None]] = []
    parsed_qtys: list[tuple[Decimal | None, ReasonCode | None]] = []
    for row in rows:
        parsed_dates.append(parse_date_value(row["raw_ship_date"], required=True))
        parsed_qtys.append(parse_decimal_value(row["raw_shipped_qty"], required=True))

    enriched_rows = [
        {
            **row,
            "_parsed_qty": parsed_qtys[i][0],
            "_parsed_date": parsed_dates[i][0],
        }
        for i, row in enumerate(rows)
    ]

    # Preliminary order-line-scoped shipment identity; alias-aware dedup runs post-matching
    is_dup, dup_of, is_conflict = _classify_key_groups(
        enriched_rows,
        key_fn=lambda r: (
            (
                r["norm_shipment_id"],
                r["norm_po_number"] or "",
                r["norm_po_line_id"] or "",
                r["norm_sku"],
            )
            if (
                r["norm_shipment_id"] is not None
                and r["norm_sku"] is not None
                and (r["norm_po_number"] is not None or r["norm_po_line_id"] is not None)
            )
            else None
        ),
        payload_fn=lambda r: (
            _dec_cmp_key(r["_parsed_qty"], r["raw_shipped_qty"]),
            r["_parsed_date"] or str(r["raw_ship_date"] or ""),
            r["norm_carrier"],
            r["norm_tracking_ref"],
            r["norm_warehouse_code"],
        ),
    )

    norm_dates: list[str | None] = []
    norm_qtys: list[str | None] = []
    valid_flags: list[bool] = []
    issue_codes_col: list[list[str]] = []

    missing_rows = 0
    invalid_num_rows = 0
    invalid_date_rows = 0
    broken_formula_rows = 0

    for idx, row in enumerate(rows):
        codes: list[str] = []
        if row["_has_broken_formula"]:
            _append_unique_code(codes, ReasonCode.BROKEN_EXCEL_FORMULA)
            broken_formula_rows += 1
        if row["_has_unpermitted_formula"]:
            _append_unique_code(codes, ReasonCode.UNPERMITTED_EXCEL_FORMULA)
            broken_formula_rows += 1

        if (
            row["norm_shipment_id"] is None
            or row["norm_po_number"] is None
            or row["norm_sku"] is None
        ):
            missing_rows += 1
            if row["norm_po_number"] is None:
                _append_unique_code(codes, ReasonCode.MISSING_ORDER_REFERENCE)
            if row["norm_sku"] is None:
                _append_unique_code(codes, ReasonCode.MISSING_SKU)

        date_iso, date_err = parsed_dates[idx]
        norm_dates.append(date_iso)
        if date_err is not None:
            _append_unique_code(codes, date_err)
            invalid_date_rows += 1

        qty_dec, qty_err = parsed_qtys[idx]
        norm_qtys.append(str(qty_dec) if qty_dec is not None else None)
        if qty_err is not None:
            _append_unique_code(codes, qty_err)
            invalid_num_rows += 1

        if is_conflict[idx]:
            _append_unique_code(codes, ReasonCode.CONFLICTING_DUPLICATE_SHIPMENT)
        elif is_dup[idx]:
            _append_unique_code(codes, ReasonCode.DUPLICATE_SHIPMENT)

        is_valid = (
            row["norm_shipment_id"] is not None
            and row["norm_sku"] is not None
            and date_err is None
            and qty_err is None
            and not row["_has_broken_formula"]
            and not row["_has_unpermitted_formula"]
            and not is_conflict[idx]
        )
        valid_flags.append(is_valid)
        issue_codes_col.append(codes)

    df = df.with_columns(
        pl.Series("_is_duplicate", is_dup, dtype=pl.Boolean),
        pl.Series("_duplicate_of_row", dup_of, dtype=pl.Int64),
        pl.Series("_is_conflicting_duplicate", is_conflict, dtype=pl.Boolean),
        pl.Series("norm_ship_date", norm_dates, dtype=pl.String),
        pl.Series("norm_shipped_qty", norm_qtys, dtype=pl.String),
        pl.Series("_is_valid_record", valid_flags, dtype=pl.Boolean),
        pl.Series("_validation_codes", issue_codes_col, dtype=pl.List(pl.String)),
    )

    dup_count = sum(1 for flag in is_dup if flag) if n else 0
    conflict_count = sum(1 for flag in is_conflict if flag) if n else 0
    valid_unique = (
        int(
            df.select(
                (
                    pl.col("_is_valid_record")
                    & ~pl.col("_is_duplicate")
                    & ~pl.col("_is_conflicting_duplicate")
                    & pl.col("norm_po_number").is_not_null()
                ).sum()
            ).item()
            or 0
        )
        if not df.is_empty()
        else 0
    )

    summary = CleaningSummary(
        dataset_kind=DatasetKind.SHIPMENTS,
        total_rows=df.height,
        valid_unique_rows=valid_unique,
        duplicate_rows=dup_count,
        conflicting_duplicate_rows=conflict_count,
        whitespace_normalized_count=_count_whitespace_fixes(
            df, ["shipment_id", "po_number", "po_line_id", "sku", "tracking_ref"]
        ),
        sku_normalized_count=_count_sku_fixes(df, "raw_sku", "norm_sku"),
        missing_value_rows=missing_rows,
        invalid_numeric_rows=invalid_num_rows,
        invalid_date_rows=invalid_date_rows,
        broken_formula_rows=broken_formula_rows,
    )
    return CleanedDataset(DatasetKind.SHIPMENTS, df, schema_res, summary)


def clean_invoices_df(
    df: pl.DataFrame,
    observed_headers: tuple[str, ...] | list[str] | None = None,
    *,
    reporting_currency: str = "USD",
) -> CleanedDataset:
    """Clean and validate Supplier Invoices with supplier-scoped deduplication."""
    headers = observed_headers if observed_headers is not None else df.columns
    schema_res = validate_schema(DatasetKind.INVOICES, headers)
    df = _preserve_raw_columns(df, REQUIRED_COLUMNS[DatasetKind.INVOICES])
    req_currency = reporting_currency.strip().upper()

    df = df.with_columns(
        _polars_norm_id_expr("invoice_id").alias("norm_invoice_id"),
        _polars_norm_id_expr("invoice_line_id").alias("norm_invoice_line_id"),
        _polars_norm_id_expr("po_number").alias("norm_po_number"),
        _polars_norm_id_expr("po_line_id").alias("norm_po_line_id"),
        _polars_norm_id_expr("supplier_id").alias("norm_supplier_id"),
        _polars_norm_sku_expr("sku").alias("norm_sku"),
        _polars_norm_id_expr("currency").alias("norm_currency"),
    )

    rows = df.to_dicts()
    n = len(rows)

    parsed_dates: list[tuple[str | None, ReasonCode | None]] = []
    parsed_qtys: list[tuple[Decimal | None, ReasonCode | None]] = []
    parsed_prices: list[tuple[Decimal | None, ReasonCode | None]] = []
    parsed_amts: list[tuple[Decimal | None, ReasonCode | None]] = []

    for row in rows:
        parsed_dates.append(parse_date_value(row["raw_invoice_date"], required=True))
        parsed_qtys.append(parse_decimal_value(row["raw_invoiced_qty"], required=True))
        parsed_prices.append(parse_decimal_value(row["raw_unit_price"], required=True))
        parsed_amts.append(parse_decimal_value(row["raw_line_amount"], required=True))

    enriched_rows = [
        {
            **row,
            "_parsed_qty": parsed_qtys[i][0],
            "_parsed_price": parsed_prices[i][0],
            "_parsed_amt": parsed_amts[i][0],
            "_parsed_date": parsed_dates[i][0],
        }
        for i, row in enumerate(rows)
    ]

    # Supplier-scoped invoice identity prevents invoice number collisions across suppliers
    is_dup, dup_of, is_conflict = _classify_key_groups(
        enriched_rows,
        key_fn=lambda r: (
            (r["norm_supplier_id"], r["norm_invoice_id"], r["norm_invoice_line_id"])
            if (
                r["norm_supplier_id"] is not None
                and r["norm_invoice_id"] is not None
                and r["norm_invoice_line_id"] is not None
            )
            else None
        ),
        payload_fn=lambda r: (
            r["norm_po_number"],
            r["norm_po_line_id"],
            r["norm_sku"],
            _dec_cmp_key(r["_parsed_qty"], r["raw_invoiced_qty"]),
            _dec_cmp_key(r["_parsed_price"], r["raw_unit_price"]),
            _dec_cmp_key(r["_parsed_amt"], r["raw_line_amount"]),
            r["norm_currency"],
            r["_parsed_date"] or str(r["raw_invoice_date"] or ""),
        ),
    )

    norm_dates: list[str | None] = []
    norm_qtys: list[str | None] = []
    norm_prices: list[str | None] = []
    norm_amounts: list[str | None] = []
    valid_flags: list[bool] = []
    issue_codes_col: list[list[str]] = []

    missing_rows = 0
    invalid_num_rows = 0
    invalid_date_rows = 0
    broken_formula_rows = 0

    for idx, row in enumerate(rows):
        codes: list[str] = []
        row_num_err = False

        if row["_has_broken_formula"]:
            _append_unique_code(codes, ReasonCode.BROKEN_EXCEL_FORMULA)
            broken_formula_rows += 1
        if row["_has_unpermitted_formula"]:
            _append_unique_code(codes, ReasonCode.UNPERMITTED_EXCEL_FORMULA)
            broken_formula_rows += 1

        if (
            row["norm_invoice_id"] is None
            or row["norm_invoice_line_id"] is None
            or row["norm_supplier_id"] is None
            or row["norm_po_number"] is None
            or row["norm_sku"] is None
        ):
            missing_rows += 1
            if row["norm_supplier_id"] is None:
                _append_unique_code(codes, ReasonCode.SUPPLIER_MISMATCH)
            if row["norm_po_number"] is None:
                _append_unique_code(codes, ReasonCode.MISSING_ORDER_REFERENCE)
            if row["norm_sku"] is None:
                _append_unique_code(codes, ReasonCode.MISSING_SKU)

        currency_ok = row["norm_currency"] == req_currency
        if not currency_ok:
            _append_unique_code(codes, ReasonCode.CURRENCY_MISMATCH)

        date_iso, date_err = parsed_dates[idx]
        norm_dates.append(date_iso)
        if date_err is not None:
            _append_unique_code(codes, date_err)
            invalid_date_rows += 1

        qty_dec, qty_err = parsed_qtys[idx]
        norm_qtys.append(str(qty_dec) if qty_dec is not None else None)
        if qty_err is not None:
            _append_unique_code(codes, qty_err)
            row_num_err = True

        price_dec, price_err = parsed_prices[idx]
        norm_prices.append(str(price_dec) if price_dec is not None else None)
        if price_err is not None:
            _append_unique_code(codes, price_err)
            row_num_err = True

        amt_dec, amt_err = parsed_amts[idx]
        norm_amounts.append(str(amt_dec) if amt_dec is not None else None)
        if amt_err is not None:
            _append_unique_code(codes, amt_err)
            row_num_err = True

        if row_num_err:
            invalid_num_rows += 1

        if is_conflict[idx]:
            _append_unique_code(codes, ReasonCode.CONFLICTING_DUPLICATE_INVOICE)
        elif is_dup[idx]:
            _append_unique_code(codes, ReasonCode.DUPLICATE_INVOICE)

        is_valid = (
            row["norm_invoice_id"] is not None
            and row["norm_invoice_line_id"] is not None
            and row["norm_supplier_id"] is not None
            and row["norm_sku"] is not None
            and currency_ok
            and date_err is None
            and qty_err is None
            and price_err is None
            and amt_err is None
            and not row["_has_broken_formula"]
            and not row["_has_unpermitted_formula"]
            and not is_conflict[idx]
        )
        valid_flags.append(is_valid)
        issue_codes_col.append(codes)

    df = df.with_columns(
        pl.Series("_is_duplicate", is_dup, dtype=pl.Boolean),
        pl.Series("_duplicate_of_row", dup_of, dtype=pl.Int64),
        pl.Series("_is_conflicting_duplicate", is_conflict, dtype=pl.Boolean),
        pl.Series("norm_invoice_date", norm_dates, dtype=pl.String),
        pl.Series("norm_invoiced_qty", norm_qtys, dtype=pl.String),
        pl.Series("norm_unit_price", norm_prices, dtype=pl.String),
        pl.Series("norm_line_amount", norm_amounts, dtype=pl.String),
        pl.Series("_is_valid_record", valid_flags, dtype=pl.Boolean),
        pl.Series("_validation_codes", issue_codes_col, dtype=pl.List(pl.String)),
    )

    dup_count = sum(1 for flag in is_dup if flag) if n else 0
    conflict_count = sum(1 for flag in is_conflict if flag) if n else 0
    valid_unique = (
        int(
            df.select(
                (
                    pl.col("_is_valid_record")
                    & ~pl.col("_is_duplicate")
                    & ~pl.col("_is_conflicting_duplicate")
                    & pl.col("norm_po_number").is_not_null()
                ).sum()
            ).item()
            or 0
        )
        if not df.is_empty()
        else 0
    )

    summary = CleaningSummary(
        dataset_kind=DatasetKind.INVOICES,
        total_rows=df.height,
        valid_unique_rows=valid_unique,
        duplicate_rows=dup_count,
        conflicting_duplicate_rows=conflict_count,
        whitespace_normalized_count=_count_whitespace_fixes(
            df, ["invoice_id", "invoice_line_id", "po_number", "po_line_id", "supplier_id", "sku"]
        ),
        sku_normalized_count=_count_sku_fixes(df, "raw_sku", "norm_sku"),
        missing_value_rows=missing_rows,
        invalid_numeric_rows=invalid_num_rows,
        invalid_date_rows=invalid_date_rows,
        broken_formula_rows=broken_formula_rows,
    )
    return CleanedDataset(DatasetKind.INVOICES, df, schema_res, summary)
