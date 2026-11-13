"""Excel workbook ingestion, SHA-256 hashing, and safe formula-aware cell extraction."""

from __future__ import annotations

import ast
import datetime as dt
import hashlib
import json
import math
import operator
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
from typing import Any

import openpyxl
import polars as pl
from openpyxl.utils import column_index_from_string
from openpyxl.utils.datetime import to_excel

from excel_reconciliation.models import (
    ARITHMETIC_CONTEXT,
    EXCEL_ERROR_TOKENS,
    DatasetKind,
    InputFileMetadata,
    excel_round,
)

_FORBIDDEN_FORMULA_CHARS = frozenset({"[", "]", "!", "|", '"', "'", ";"})
_RANGE_TOKEN_RE = re.compile(
    r"\$?([A-Za-z]{1,3})\$?([1-9]\d*)\s*:\s*\$?([A-Za-z]{1,3})\$?([1-9]\d*)"
)
_CELL_TOKEN_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?([1-9]\d*)(?!\s*\()")
_PERMITTED_FUNCTIONS = frozenset({"SUM", "ROUND"})

DEFAULT_MAX_INPUT_FORMULA_LENGTH = 2_048
DEFAULT_MAX_INPUT_RANGE_CELLS = 100_000
DEFAULT_MAX_INPUT_EVALUATED_CELLS = 250_000


@dataclass(frozen=True, slots=True)
class _GridCell:
    """Typed worksheet cell entry distinguishing true formula cells and original cell types."""

    raw_value: Any
    is_formula: bool
    cached_value: Any = None
    data_type: str = "n"
    excel_serial: Decimal | None = None


@dataclass(frozen=True, slots=True)
class LoadedWorkbookData:
    """Container holding raw Polars DataFrame and provenance metadata."""

    dataset_kind: DatasetKind
    dataframe: pl.DataFrame
    observed_headers: tuple[str, ...]
    metadata: InputFileMetadata


def compute_sha256(filepath: Path) -> str:
    """Compute hex SHA-256 digest of a file."""
    digest = hashlib.sha256()
    with filepath.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _format_plain_decimal(dec: Decimal) -> str:
    """Format a finite Decimal without scientific notation while preserving exact scale."""
    raw = str(dec)
    if "e" not in raw and "E" not in raw:
        return raw
    plain = format(dec, "f")
    if "." in plain:
        plain = plain.rstrip("0").rstrip(".")
    return plain or "0"


def _cell_scalar_to_str(value: Any) -> str | None:
    """Convert an openpyxl cell scalar into an auditable string without losing precision."""
    if value is None:
        return None
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, dt.datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return _format_plain_decimal(value) if value.is_finite() else str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return "#NUM!"
        return _format_plain_decimal(Decimal(str(value)))
    text = str(value)
    return text if text != "" else None


def _contains_excel_error(text: str) -> bool:
    """Return True if string contains any standard Excel error token."""
    upper = text.strip().upper()
    return any(token in upper for token in EXCEL_ERROR_TOKENS)


def _preprocess_formula_expr(formula: str) -> str | None:
    """Translate intra-sheet ranges and cell coordinates into safe AST helper calls."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    body = formula[1:].strip()
    if not body:
        return None
    if any(ch in body for ch in _FORBIDDEN_FORMULA_CHARS):
        return None

    with_ranges = _RANGE_TOKEN_RE.sub(
        lambda m: (
            f'__RANGE__("{m.group(1).upper()}", {m.group(2)}, "{m.group(3).upper()}", {m.group(4)})'
        ),
        body,
    )
    with_cells = _CELL_TOKEN_RE.sub(
        lambda m: f'__CELL__("{m.group(1).upper()}", {m.group(2)})',
        with_ranges,
    )
    return with_cells


def _validate_formula_ast(node: ast.AST) -> bool:
    """Verify AST uses only permitted arithmetic operators, cell refs, SUM, and ROUND."""
    if isinstance(node, ast.Expression):
        return _validate_formula_ast(node.body)
    if isinstance(node, ast.Constant):
        return isinstance(node.value, (int, float)) and not isinstance(node.value, bool)
    if isinstance(node, ast.UnaryOp):
        return isinstance(node.op, (ast.UAdd, ast.USub)) and _validate_formula_ast(node.operand)
    if isinstance(node, ast.BinOp):
        return (
            isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div))
            and _validate_formula_ast(node.left)
            and _validate_formula_ast(node.right)
        )
    if isinstance(node, ast.Call):
        if node.keywords or not isinstance(node.func, ast.Name):
            return False
        fn_name = node.func.id.upper()
        if fn_name == "__CELL__":
            return (
                len(node.args) == 2
                and isinstance(node.args[0], ast.Constant)
                and isinstance(node.args[0].value, str)
                and isinstance(node.args[1], ast.Constant)
                and isinstance(node.args[1].value, int)
            )
        if fn_name == "__RANGE__":
            return (
                len(node.args) == 4
                and isinstance(node.args[0], ast.Constant)
                and isinstance(node.args[0].value, str)
                and isinstance(node.args[1], ast.Constant)
                and isinstance(node.args[1].value, int)
                and isinstance(node.args[2], ast.Constant)
                and isinstance(node.args[2].value, str)
                and isinstance(node.args[3], ast.Constant)
                and isinstance(node.args[3].value, int)
            )
        if fn_name == "SUM":
            return len(node.args) >= 1 and all(_validate_formula_ast(arg) for arg in node.args)
        if fn_name == "ROUND":
            return len(node.args) == 2 and all(_validate_formula_ast(arg) for arg in node.args)
        return False
    return False


def is_permitted_source_formula(
    formula: str,
    *,
    max_formula_length: int = DEFAULT_MAX_INPUT_FORMULA_LENGTH,
) -> bool:
    """Return True iff formula is a safe intra-sheet arithmetic/SUM/ROUND Excel formula."""
    if (
        not isinstance(formula, str)
        or not formula.startswith("=")
        or len(formula) > max_formula_length
    ):
        return False
    if _contains_excel_error(formula):
        return False
    preprocessed = _preprocess_formula_expr(formula)
    if preprocessed is None:
        return False
    try:
        parsed = ast.parse(preprocessed, mode="eval")
        return _validate_formula_ast(parsed)
    except (SyntaxError, RecursionError):
        return False


class _FormulaEvalError(Exception):
    """Internal exception carrying an Excel error token during AST evaluation."""

    def __init__(self, token: str) -> None:
        super().__init__(token)
        self.token = token


def _extract_excel_error(text: str) -> str:
    """Extract the matching standard Excel error token from text."""
    upper = text.strip().upper()
    for token in EXCEL_ERROR_TOKENS:
        if token in upper:
            return token
    return "#ERROR!"


def evaluate_worksheet_formula(
    formula: str,
    grid_values: dict[tuple[int, int], Any],
    *,
    depth: int = 0,
    max_formula_length: int = DEFAULT_MAX_INPUT_FORMULA_LENGTH,
    max_range_cells: int = DEFAULT_MAX_INPUT_RANGE_CELLS,
    max_evaluated_cells: int = DEFAULT_MAX_INPUT_EVALUATED_CELLS,
    _evaluated_counter: list[int] | None = None,
) -> str | None:
    """Safely evaluate permitted intra-sheet Excel arithmetic/SUM/ROUND formulas.

    Uses exact Decimal arithmetic in a 50-digit context, follows Excel SUM cell-reference
    rules (ignoring text-formatted numbers and booleans in referenced cells while summing
    numeric cells and numeric formula results), and enforces bounded range/cell limits.
    Returns an Excel error token if the formula is unpermitted, exceeds evaluation limits,
    references broken/non-numeric cells, or divides by zero.
    """
    if depth > 8:
        return "#VALUE!"
    if not isinstance(formula, str) or len(formula) > max_formula_length:
        return "#VALUE!"
    if _contains_excel_error(formula):
        return _extract_excel_error(formula)

    preprocessed = _preprocess_formula_expr(formula)
    if preprocessed is None:
        return "#NAME?"

    try:
        parsed = ast.parse(preprocessed, mode="eval")
        permitted = _validate_formula_ast(parsed)
    except (SyntaxError, RecursionError):
        return "#VALUE!"

    if not permitted:
        return "#NAME?"

    eval_counter = _evaluated_counter if _evaluated_counter is not None else [0]

    def _charge_eval(cell_count: int) -> None:
        eval_counter[0] += cell_count
        if eval_counter[0] > max_evaluated_cells:
            raise _FormulaEvalError("#VALUE!")

    def _eval_node(node: ast.AST) -> Decimal:
        if isinstance(node, ast.Expression):
            return _eval_node(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return Decimal(str(node.value))
        if isinstance(node, ast.UnaryOp):
            operand = _eval_node(node.operand)
            return +operand if isinstance(node.op, ast.UAdd) else -operand
        if isinstance(node, ast.BinOp):
            left = _eval_node(node.left)
            right = _eval_node(node.right)
            if isinstance(node.op, ast.Div) and right == Decimal("0"):
                raise _FormulaEvalError("#DIV/0!")
            ops = {
                ast.Add: operator.add,
                ast.Sub: operator.sub,
                ast.Mult: operator.mul,
                ast.Div: operator.truediv,
            }
            with localcontext(ARITHMETIC_CONTEXT):
                return ops[type(node.op)](left, right)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            fn_name = node.func.id.upper()
            if fn_name == "__CELL__":
                _charge_eval(1)
                col_s = str(node.args[0].value)  # type: ignore[attr-defined]
                row_i = int(node.args[1].value)  # type: ignore[attr-defined]
                c_idx = column_index_from_string(col_s)
                resolved = _resolve_grid_decimal(
                    grid_values.get((row_i, c_idx)),
                    grid_values,
                    depth=depth + 1,
                    max_formula_length=max_formula_length,
                    max_range_cells=max_range_cells,
                    max_evaluated_cells=max_evaluated_cells,
                    eval_counter=eval_counter,
                )
                if isinstance(resolved, str):
                    raise _FormulaEvalError(resolved)
                return resolved if resolved is not None else Decimal("0")
            if fn_name == "__RANGE__":
                raise _FormulaEvalError("#VALUE!")
            if fn_name == "SUM":
                total = Decimal("0")
                with localcontext(ARITHMETIC_CONTEXT):
                    for arg in node.args:
                        if (
                            isinstance(arg, ast.Call)
                            and isinstance(arg.func, ast.Name)
                            and arg.func.id == "__RANGE__"
                        ):
                            col1_s = str(arg.args[0].value)  # type: ignore[attr-defined]
                            row1_i = int(arg.args[1].value)  # type: ignore[attr-defined]
                            col2_s = str(arg.args[2].value)  # type: ignore[attr-defined]
                            row2_i = int(arg.args[3].value)  # type: ignore[attr-defined]
                            c1, c2 = (
                                column_index_from_string(col1_s),
                                column_index_from_string(col2_s),
                            )
                            r_lo, r_hi = min(row1_i, row2_i), max(row1_i, row2_i)
                            c_lo, c_hi = min(c1, c2), max(c1, c2)
                            range_cells = (r_hi - r_lo + 1) * (c_hi - c_lo + 1)
                            if range_cells > max_range_cells:
                                raise _FormulaEvalError("#VALUE!")
                            _charge_eval(range_cells)
                            for r_i in range(r_lo, r_hi + 1):
                                for c_i in range(c_lo, c_hi + 1):
                                    resolved = _resolve_grid_sum_ref(
                                        grid_values.get((r_i, c_i)),
                                        grid_values,
                                        depth=depth + 1,
                                        max_formula_length=max_formula_length,
                                        max_range_cells=max_range_cells,
                                        max_evaluated_cells=max_evaluated_cells,
                                        eval_counter=eval_counter,
                                    )
                                    if isinstance(resolved, str):
                                        raise _FormulaEvalError(resolved)
                                    if resolved is not None:
                                        total += resolved
                        elif (
                            isinstance(arg, ast.Call)
                            and isinstance(arg.func, ast.Name)
                            and arg.func.id == "__CELL__"
                        ):
                            _charge_eval(1)
                            col_s = str(arg.args[0].value)  # type: ignore[attr-defined]
                            row_i = int(arg.args[1].value)  # type: ignore[attr-defined]
                            c_idx = column_index_from_string(col_s)
                            resolved = _resolve_grid_sum_ref(
                                grid_values.get((row_i, c_idx)),
                                grid_values,
                                depth=depth + 1,
                                max_formula_length=max_formula_length,
                                max_range_cells=max_range_cells,
                                max_evaluated_cells=max_evaluated_cells,
                                eval_counter=eval_counter,
                            )
                            if isinstance(resolved, str):
                                raise _FormulaEvalError(resolved)
                            if resolved is not None:
                                total += resolved
                        else:
                            total += _eval_node(arg)
                return total
            if fn_name == "ROUND":
                val_dec = _eval_node(node.args[0])
                digits_dec = _eval_node(node.args[1])
                digits_int = int(digits_dec)
                if Decimal(digits_int) != digits_dec or abs(digits_int) > 15:
                    raise _FormulaEvalError("#VALUE!")
                return excel_round(val_dec, digits_int)
        raise _FormulaEvalError("#VALUE!")

    try:
        result = _eval_node(parsed)
    except _FormulaEvalError as exc:
        return exc.token
    except ZeroDivisionError:
        return "#DIV/0!"
    except Exception:
        return "#VALUE!"

    return _format_plain_decimal(result)


def _eval_nested_formula_decimal(
    formula_text: str,
    grid_values: dict[tuple[int, int], Any],
    *,
    depth: int,
    max_formula_length: int,
    max_range_cells: int,
    max_evaluated_cells: int,
    eval_counter: list[int],
) -> Decimal | str | None:
    """Evaluate a nested cell formula to Decimal or Excel error string."""
    norm_formula = formula_text if formula_text.startswith("=") else f"={formula_text}"
    if _contains_excel_error(norm_formula):
        return _extract_excel_error(norm_formula)
    if not is_permitted_source_formula(norm_formula, max_formula_length=max_formula_length):
        return "#NAME?"
    evaluated = evaluate_worksheet_formula(
        norm_formula,
        grid_values,
        depth=depth,
        max_formula_length=max_formula_length,
        max_range_cells=max_range_cells,
        max_evaluated_cells=max_evaluated_cells,
        _evaluated_counter=eval_counter,
    )
    if evaluated is None:
        return None
    if _contains_excel_error(evaluated):
        return _extract_excel_error(evaluated)
    try:
        dec = Decimal(evaluated)
        return dec if dec.is_finite() else "#NUM!"
    except InvalidOperation:
        return "#VALUE!"


def _resolve_grid_sum_ref(
    raw_entry: Any,
    grid_values: dict[tuple[int, int], Any],
    *,
    depth: int,
    max_formula_length: int,
    max_range_cells: int,
    max_evaluated_cells: int,
    eval_counter: list[int],
) -> Decimal | str | None:
    """Resolve a cell reference inside SUM(...) following Excel type semantics.

    Numeric cells and numeric formula results are summed; blanks, text-formatted
    cells (even numeric-looking strings), and booleans are ignored (None);
    error cells and errored formulas propagate their Excel error token.
    """
    if raw_entry is None:
        return None

    if isinstance(raw_entry, _GridCell):
        if raw_entry.is_formula:
            if raw_entry.cached_value is not None:
                cached_str = _cell_scalar_to_str(raw_entry.cached_value)
                if cached_str is not None and _contains_excel_error(cached_str):
                    return _extract_excel_error(cached_str)
            return _eval_nested_formula_decimal(
                str(raw_entry.raw_value or ""),
                grid_values,
                depth=depth,
                max_formula_length=max_formula_length,
                max_range_cells=max_range_cells,
                max_evaluated_cells=max_evaluated_cells,
                eval_counter=eval_counter,
            )
        if raw_entry.data_type == "e":
            err_str = _cell_scalar_to_str(raw_entry.raw_value)
            return _extract_excel_error(err_str) if err_str else "#ERROR!"
        if raw_entry.excel_serial is not None:
            return raw_entry.excel_serial
        val = raw_entry.raw_value
        if val is None:
            return None
        if raw_entry.data_type in ("s", "inlineStr", "str", "b") or isinstance(val, (str, bool)):
            return None
        if (
            raw_entry.data_type == "n"
            and isinstance(val, (int, float, Decimal))
            and not isinstance(val, bool)
        ):
            if isinstance(val, Decimal):
                return val if val.is_finite() else "#NUM!"
            if isinstance(val, int):
                return Decimal(val)
            if not math.isfinite(val):
                return "#NUM!"
            return Decimal(str(val))
        return None

    if isinstance(raw_entry, str):
        if raw_entry.startswith("="):
            return _eval_nested_formula_decimal(
                raw_entry,
                grid_values,
                depth=depth,
                max_formula_length=max_formula_length,
                max_range_cells=max_range_cells,
                max_evaluated_cells=max_evaluated_cells,
                eval_counter=eval_counter,
            )
        if _contains_excel_error(raw_entry):
            return _extract_excel_error(raw_entry)
        return None
    if isinstance(raw_entry, bool):
        return None
    if isinstance(raw_entry, Decimal):
        return raw_entry if raw_entry.is_finite() else "#NUM!"
    if isinstance(raw_entry, int):
        return Decimal(raw_entry)
    if isinstance(raw_entry, float):
        return Decimal(str(raw_entry)) if math.isfinite(raw_entry) else "#NUM!"
    return None


def _resolve_grid_decimal(
    raw_entry: Any,
    grid_values: dict[tuple[int, int], Any],
    *,
    depth: int,
    max_formula_length: int = DEFAULT_MAX_INPUT_FORMULA_LENGTH,
    max_range_cells: int = DEFAULT_MAX_INPUT_RANGE_CELLS,
    max_evaluated_cells: int = DEFAULT_MAX_INPUT_EVALUATED_CELLS,
    eval_counter: list[int] | None = None,
) -> Decimal | str | None:
    """Resolve a cell entry or nested formula for arithmetic operators."""
    if raw_entry is None:
        return None
    counter = eval_counter if eval_counter is not None else [0]

    if isinstance(raw_entry, _GridCell):
        if raw_entry.is_formula:
            if raw_entry.cached_value is not None:
                cached_str = _cell_scalar_to_str(raw_entry.cached_value)
                if cached_str is not None and _contains_excel_error(cached_str):
                    return _extract_excel_error(cached_str)
            return _eval_nested_formula_decimal(
                str(raw_entry.raw_value or ""),
                grid_values,
                depth=depth,
                max_formula_length=max_formula_length,
                max_range_cells=max_range_cells,
                max_evaluated_cells=max_evaluated_cells,
                eval_counter=counter,
            )
        if raw_entry.data_type == "e":
            err_str = _cell_scalar_to_str(raw_entry.raw_value)
            return _extract_excel_error(err_str) if err_str else "#ERROR!"
        if raw_entry.excel_serial is not None:
            return raw_entry.excel_serial
        scalar_val = raw_entry.raw_value
    else:
        if isinstance(raw_entry, str) and raw_entry.startswith("="):
            return _eval_nested_formula_decimal(
                raw_entry,
                grid_values,
                depth=depth,
                max_formula_length=max_formula_length,
                max_range_cells=max_range_cells,
                max_evaluated_cells=max_evaluated_cells,
                eval_counter=counter,
            )
        scalar_val = raw_entry

    if isinstance(scalar_val, bool):
        return Decimal("1") if scalar_val else Decimal("0")

    as_str = _cell_scalar_to_str(scalar_val)
    if as_str is None:
        return None
    if _contains_excel_error(as_str):
        return _extract_excel_error(as_str)
    # Literal text strings starting with '=' are text, never formulas
    if isinstance(scalar_val, str) and as_str.lstrip().startswith("="):
        return "#VALUE!"
    cleaned_num = as_str.strip().replace(",", "").removeprefix("$")
    try:
        dec = Decimal(cleaned_num)
        return dec if dec.is_finite() else "#NUM!"
    except InvalidOperation:
        return "#VALUE!"


def read_excel_dataset(
    filepath: Path,
    dataset_kind: DatasetKind,
    *,
    sheet_name: str | None = None,
    max_formula_length: int = DEFAULT_MAX_INPUT_FORMULA_LENGTH,
    max_range_cells: int = DEFAULT_MAX_INPUT_RANGE_CELLS,
    max_evaluated_cells: int = DEFAULT_MAX_INPUT_EVALUATED_CELLS,
) -> LoadedWorkbookData:
    """Read an Excel worksheet into a string-typed Polars DataFrame with formula provenance.

    Strictly separates genuine Excel formula cells (cell.data_type == 'f') from
    untrusted text strings starting with '='. Only explicitly permitted intra-sheet
    arithmetic/SUM/ROUND formulas are evaluated and preserved; unpermitted or broken
    formulas are quarantined.
    """
    if not filepath.exists():
        raise FileNotFoundError(f"Input workbook not found: {filepath}")

    sha256_hex = compute_sha256(filepath)
    wb_formulas = openpyxl.load_workbook(filepath, data_only=False)
    wb_values = openpyxl.load_workbook(filepath, data_only=True)

    try:
        target_sheet = sheet_name or wb_formulas.sheetnames[0]
        if target_sheet not in wb_formulas.sheetnames:
            raise ValueError(f"Worksheet '{target_sheet}' not found in {filepath.name}")

        ws_f = wb_formulas[target_sheet]
        ws_v = wb_values[target_sheet]

        max_row = ws_f.max_row or 0
        max_col = ws_f.max_column or 0
        if max_row < 1 or max_col < 1:
            empty_meta = InputFileMetadata(
                dataset_kind=dataset_kind,
                filename=filepath.name,
                filepath=str(filepath),
                sha256=sha256_hex,
                sheet_name=target_sheet,
                row_count=0,
                formula_cell_count=0,
                broken_formula_count=0,
            )
            return LoadedWorkbookData(
                dataset_kind=dataset_kind,
                dataframe=pl.DataFrame(),
                observed_headers=(),
                metadata=empty_meta,
            )

        # Extract header row (row 1); ignore completely blank header cells if sheet is empty
        observed_headers: list[str] = []
        col_names: list[str] = []
        seen_names: dict[str, int] = {}
        any_header_value = False
        for col_idx in range(1, max_col + 1):
            raw_hdr = ws_f.cell(row=1, column=col_idx).value
            if raw_hdr is not None and str(raw_hdr).strip() != "":
                any_header_value = True
                hdr_text = str(raw_hdr).strip().lower()
            else:
                hdr_text = f"unnamed_{col_idx}"
            observed_headers.append(hdr_text)
            if hdr_text in seen_names:
                seen_names[hdr_text] += 1
                col_names.append(f"{hdr_text}__dup{seen_names[hdr_text]}")
            else:
                seen_names[hdr_text] = 0
                col_names.append(hdr_text)

        if not any_header_value:
            observed_headers = []
            col_names = []

        # Build typed grid cache separating true formula cells and preserving data types
        grid_values: dict[tuple[int, int], _GridCell] = {}
        for r_idx in range(1, max_row + 1):
            for c_idx in range(1, max_col + 1):
                f_cell = ws_f.cell(row=r_idx, column=c_idx)
                v_cell = ws_v.cell(row=r_idx, column=c_idx)
                is_formula = f_cell.data_type == "f" and isinstance(f_cell.value, str)
                grid_values[(r_idx, c_idx)] = _GridCell(
                    raw_value=f_cell.value,
                    is_formula=is_formula,
                    cached_value=v_cell.value,
                    data_type=str(f_cell.data_type or "n"),
                    excel_serial=(
                        Decimal(str(to_excel(f_cell.value, epoch=wb_formulas.epoch)))
                        if isinstance(f_cell.value, (dt.datetime, dt.date, dt.time, dt.timedelta))
                        else None
                    ),
                )

        rows: list[dict[str, Any]] = []
        formula_cell_count = 0
        broken_formula_count = 0

        for r_idx in range(2, max_row + 1):
            row_dict: dict[str, Any] = {"_source_row": r_idx}
            row_formulas: dict[str, str] = {}
            row_has_broken_formula = False
            row_has_unpermitted_formula = False
            any_non_empty = False

            for c_idx, col_name in enumerate(col_names, start=1):
                f_cell = ws_f.cell(row=r_idx, column=c_idx)
                v_cell = ws_v.cell(row=r_idx, column=c_idx)
                f_val = f_cell.value
                v_val = v_cell.value

                is_true_formula = f_cell.data_type == "f" and isinstance(f_val, str)
                if is_true_formula:
                    any_non_empty = True
                    formula_str = str(f_val)
                    if not formula_str.startswith("="):
                        formula_str = f"={formula_str}"

                    if f_cell.data_type == "e" or _contains_excel_error(formula_str):
                        broken_formula_count += 1
                        row_has_broken_formula = True
                        row_dict[col_name] = formula_str
                    elif not is_permitted_source_formula(
                        formula_str, max_formula_length=max_formula_length
                    ):
                        broken_formula_count += 1
                        row_has_unpermitted_formula = True
                        row_dict[col_name] = None
                    else:
                        formula_cell_count += 1
                        row_formulas[col_name] = formula_str
                        cached_str = _cell_scalar_to_str(v_val) if v_val is not None else None
                        if cached_str is not None and _contains_excel_error(cached_str):
                            broken_formula_count += 1
                            row_has_broken_formula = True
                            row_dict[col_name] = cached_str
                        else:
                            evaluated = evaluate_worksheet_formula(
                                formula_str,
                                grid_values,
                                max_formula_length=max_formula_length,
                                max_range_cells=max_range_cells,
                                max_evaluated_cells=max_evaluated_cells,
                            )
                            if evaluated is not None and _contains_excel_error(evaluated):
                                broken_formula_count += 1
                                row_has_broken_formula = True
                            row_dict[col_name] = evaluated
                elif f_cell.data_type == "e":
                    any_non_empty = True
                    broken_formula_count += 1
                    row_has_broken_formula = True
                    row_dict[col_name] = _cell_scalar_to_str(f_val) or "#ERROR!"
                else:
                    # Literal scalar or text cell (even if starting with '=', '+', '-', '@')
                    cell_str = _cell_scalar_to_str(f_val)
                    if cell_str is not None:
                        any_non_empty = True
                    row_dict[col_name] = cell_str

            if not any_non_empty:
                continue

            row_dict["_formulas_json"] = (
                json.dumps(row_formulas, sort_keys=True) if row_formulas else None
            )
            row_dict["_has_broken_formula"] = row_has_broken_formula
            row_dict["_has_unpermitted_formula"] = row_has_unpermitted_formula
            rows.append(row_dict)

        schema_spec: dict[str, pl.DataType] = {"_source_row": pl.Int64}
        for col_name in col_names:
            schema_spec[col_name] = pl.String
        schema_spec["_formulas_json"] = pl.String
        schema_spec["_has_broken_formula"] = pl.Boolean
        schema_spec["_has_unpermitted_formula"] = pl.Boolean

        df = pl.DataFrame(rows, schema=schema_spec) if rows else pl.DataFrame(schema=schema_spec)
        metadata = InputFileMetadata(
            dataset_kind=dataset_kind,
            filename=filepath.name,
            filepath=str(filepath),
            sha256=sha256_hex,
            sheet_name=target_sheet,
            row_count=df.height,
            formula_cell_count=formula_cell_count,
            broken_formula_count=broken_formula_count,
        )
        return LoadedWorkbookData(
            dataset_kind=dataset_kind,
            dataframe=df,
            observed_headers=tuple(observed_headers),
            metadata=metadata,
        )
    finally:
        wb_formulas.close()
        wb_values.close()
