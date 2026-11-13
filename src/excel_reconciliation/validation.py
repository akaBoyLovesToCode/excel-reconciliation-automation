"""Structural, formula-safety, OpenXML cache, and accounting validation for Excel reports."""

from __future__ import annotations

import math
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

from excel_reconciliation.models import (
    ARITHMETIC_CONTEXT,
    EXCEL_ERROR_TOKENS,
    DatasetKind,
    ExceptionCategory,
    MatchStatus,
    ReconciliationStatus,
    Severity,
    ValidationReport,
    excel_round,
)

REQUIRED_WORKSHEETS: tuple[str, ...] = (
    "Summary",
    "Reconciled Records",
    "Exceptions",
    "Audit Log",
)

REQUIRED_SUMMARY_KPI_LABELS: tuple[str, ...] = (
    "Total Purchase Order Rows",
    "Unique Valid Purchase Order Lines",
    "Successfully Reconciled PO Lines",
    "Partial Fulfillment PO Lines",
    "PO Lines Requiring Review",
    "Total Exception Records (All Sources)",
    "Product Matching Coverage",
    "Three-Way Auto-Reconciliation Rate",
    "Total Ordered Quantity (Valid Unique POs)",
    "Total Shipped Quantity",
    "Total Invoiced Quantity",
    "Outstanding Unshipped Quantity",
    "Total Expected PO Amount",
    "Total Actual Invoiced Amount",
    "Net Billed Monetary Variance",
    "Absolute Monetary Discrepancy Exposure",
    "Unbilled Shipped Value (Accrual Exposure)",
    "Quarantined / Unmatched Invoice Amount",
)

_DANGEROUS_TEXT_PREFIXES: tuple[str, ...] = ("=", "+", "-", "@", "\t", "\r", "\n")

DEFAULT_MAX_FORMULA_LENGTH = 2_048
DEFAULT_MAX_FORMULA_RANGE_CELLS = 100_000
DEFAULT_MAX_FORMULA_EVALUATED_CELLS = 250_000
DEFAULT_MAX_FORMULA_RECURSION_DEPTH = 25

_CELL_REF_RE = re.compile(r"^(?:(?:'([^']+)'|([A-Za-z0-9_.-]+))!)?\$?([A-Za-z]{1,3})\$?([0-9]+)$")
_RANGE_REF_RE = re.compile(
    r"^(?:(?:'([^']+)'|([A-Za-z0-9_.-]+))!)?\$?([A-Za-z]{1,3})\$?([0-9]+):\$?([A-Za-z]{1,3})\$?([0-9]+)$"
)


class FormulaResourceLimitError(ValueError):
    """Raised when a formula exceeds configured length, range-size, or cell evaluation limits."""


@dataclass(frozen=True, slots=True)
class FormulaEvaluationLimits:
    """Configurable resource limits preventing unbounded formula evaluation."""

    max_formula_length: int = DEFAULT_MAX_FORMULA_LENGTH
    max_range_cells: int = DEFAULT_MAX_FORMULA_RANGE_CELLS
    max_evaluated_cells: int = DEFAULT_MAX_FORMULA_EVALUATED_CELLS
    max_recursion_depth: int = DEFAULT_MAX_FORMULA_RECURSION_DEPTH

    def __post_init__(self) -> None:
        for field_name in (
            "max_formula_length",
            "max_range_cells",
            "max_evaluated_cells",
            "max_recursion_depth",
        ):
            val = getattr(self, field_name)
            if isinstance(val, bool) or not isinstance(val, int) or val < 1:
                raise ValueError(f"{field_name} must be a positive integer, got {val!r}")


@dataclass(slots=True)
class _CachedSheetData:
    """Bounded, read-only snapshot of a worksheet's dimensions and non-empty cells."""

    max_row: int
    max_col: int
    cells: dict[tuple[int, int], tuple[Any, str]]


class WorkbookFormulaCache:
    """Caches worksheet dimensions, non-empty cells, and evaluated formula cells per workbook."""

    def __init__(self, wb: openpyxl.Workbook) -> None:
        self.wb = wb
        self.sheetnames: tuple[str, ...] = tuple(wb.sheetnames)
        self._sheet_set: frozenset[str] = frozenset(self.sheetnames)
        self._sheets: dict[str, _CachedSheetData] = {}
        self.formula_results: dict[tuple[str, int, int], Any] = {}
        self.active_cells: set[tuple[str, int, int]] = set()

    def has_sheet(self, sheet_name: str) -> bool:
        """Return True iff worksheet exists in the cached workbook."""
        return sheet_name in self._sheet_set

    def get_sheet(self, sheet_name: str) -> _CachedSheetData:
        """Build once and return cached dimensions and populated cell map for a sheet."""
        cached = self._sheets.get(sheet_name)
        if cached is not None:
            return cached
        if sheet_name not in self._sheet_set:
            raise ValueError(f"Formula references missing worksheet: {sheet_name!r}")

        ws = self.wb[sheet_name]
        cells: dict[tuple[int, int], tuple[Any, str]] = {}
        max_row = 0
        max_col = 0
        raw_cells = getattr(ws, "_cells", None)
        if isinstance(raw_cells, dict):
            for (r_idx, c_idx), cell in raw_cells.items():
                val = cell.value
                dtype = str(cell.data_type or "n")
                if val is not None or dtype == "e":
                    cells[(r_idx, c_idx)] = (val, dtype)
                    if r_idx > max_row:
                        max_row = r_idx
                    if c_idx > max_col:
                        max_col = c_idx
        else:
            for r_idx, row in enumerate(ws.iter_rows(), start=1):
                for c_idx, cell in enumerate(row, start=1):
                    val = cell.value
                    dtype = str(cell.data_type or "n")
                    if val is not None or dtype == "e":
                        cells[(r_idx, c_idx)] = (val, dtype)
                        if r_idx > max_row:
                            max_row = r_idx
                        if c_idx > max_col:
                            max_col = c_idx

        cached = _CachedSheetData(max_row=max_row, max_col=max_col, cells=cells)
        self._sheets[sheet_name] = cached
        return cached


@dataclass(frozen=True, slots=True)
class _Token:
    kind: str
    value: str


@dataclass(frozen=True, slots=True)
class _CellRangeRef:
    sheet_name: str
    min_col: int
    min_row: int
    max_col: int
    max_row: int


@dataclass(frozen=True, slots=True)
class _CellRef:
    sheet_name: str
    col: int
    row: int


def _tokenize_formula(expr: str) -> list[_Token]:
    """Tokenize an Excel formula body (without leading '=') into structured tokens."""
    tokens: list[_Token] = []
    i = 0
    n = len(expr)
    while i < n:
        ch = expr[i]
        if ch.isspace():
            i += 1
            continue
        if ch == '"':
            j = i + 1
            chars: list[str] = []
            closed = False
            while j < n:
                if expr[j] == '"':
                    if j + 1 < n and expr[j + 1] == '"':
                        chars.append('"')
                        j += 2
                        continue
                    closed = True
                    j += 1
                    break
                chars.append(expr[j])
                j += 1
            if not closed:
                raise ValueError(f"Unterminated string literal in formula: {expr!r}")
            tokens.append(_Token("STRING", "".join(chars)))
            i = j
            continue

        # Check multi-character comparison operators first
        if expr.startswith(("<>", "<=", ">="), i):
            tokens.append(_Token("OP", expr[i : i + 2]))
            i += 2
            continue
        if ch in ("+", "-", "*", "/", "=", "<", ">"):
            tokens.append(_Token("OP", ch))
            i += 1
            continue
        if ch == "(":
            tokens.append(_Token("LPAREN", ch))
            i += 1
            continue
        if ch == ")":
            tokens.append(_Token("RPAREN", ch))
            i += 1
            continue
        if ch == ",":
            tokens.append(_Token("COMMA", ch))
            i += 1
            continue

        # Quoted sheet reference starting at single quote: 'Sheet Name'!$A$1[:$B$2]
        if ch == "'":
            j = i + 1
            sheet_chars: list[str] = []
            closed = False
            while j < n:
                if expr[j] == "'":
                    if j + 1 < n and expr[j + 1] == "'":
                        sheet_chars.append("'")
                        j += 2
                        continue
                    closed = True
                    j += 1
                    break
                sheet_chars.append(expr[j])
                j += 1
            if not closed or j >= n or expr[j] != "!":
                raise ValueError(f"Invalid quoted sheet reference in formula: {expr!r}")
            # Read remaining cell/range ref after '!'
            k = j + 1
            while k < n and (expr[k].isalnum() or expr[k] in ("$", ":")):
                k += 1
            raw_ref = f"'{''.join(sheet_chars)}'!{expr[j + 1 : k]}"
            tokens.append(_Token("IDENT", raw_ref))
            i = k
            continue

        # Numeric literal (strictly digits / decimal / exponent, not starting with letters)
        if ch.isdigit() or (ch == "." and i + 1 < n and expr[i + 1].isdigit()):
            j = i
            while j < n and (expr[j].isdigit() or expr[j] == "."):
                j += 1
            if j < n and expr[j] in ("e", "E"):
                exp_idx = j + 1
                if exp_idx < n and expr[exp_idx] in ("+", "-"):
                    exp_idx += 1
                if exp_idx < n and expr[exp_idx].isdigit():
                    j = exp_idx
                    while j < n and expr[j].isdigit():
                        j += 1
            tokens.append(_Token("NUMBER", expr[i:j]))
            i = j
            continue

        # Identifier, function name, boolean, sheet!ref, or range ref
        if ch.isalpha() or ch in ("$", "_"):
            j = i
            while j < n and (expr[j].isalnum() or expr[j] in ("$", "_", ".", "!", ":")):
                j += 1
            tokens.append(_Token("IDENT", expr[i:j]))
            i = j
            continue

        raise ValueError(f"Unexpected character {ch!r} in formula: {expr!r}")

    return tokens


def _to_decimal_strict(val: Any) -> Decimal | None:
    """Convert a cell value or scalar to finite Decimal when numeric (excluding bool)."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, Decimal):
        return val if val.is_finite() else None
    if isinstance(val, int):
        return Decimal(val)
    if isinstance(val, float):
        if not math.isfinite(val):
            return None
        return Decimal(str(val))
    return None


def _matches_criterion(cell_val: Any, criterion: Any) -> bool:
    """Evaluate Excel COUNTIF/SUMIF criterion semantics against a cell value."""
    if isinstance(criterion, bool):
        return isinstance(cell_val, bool) and cell_val is criterion

    crit_dec = _to_decimal_strict(criterion)
    if crit_dec is not None:
        cell_dec = _to_decimal_strict(cell_val)
        return cell_dec is not None and cell_dec == crit_dec

    if isinstance(criterion, str):
        if criterion.startswith("<>"):
            target = criterion[2:]
            if target == "":
                return cell_val is not None and cell_val != ""
            if cell_val is None:
                return True
            target_dec = _to_decimal_strict(cell_val)
            try:
                t_num = Decimal(target)
            except InvalidOperation:
                t_num = None
            if t_num is not None and target_dec is not None:
                return target_dec != t_num
            return str(cell_val).casefold() != target.casefold()

        for op in (">=", "<=", ">", "<", "="):
            if criterion.startswith(op):
                rhs_str = criterion[len(op) :]
                if op == "=" and rhs_str == "":
                    return cell_val is None or cell_val == ""
                cell_dec = _to_decimal_strict(cell_val)
                try:
                    rhs_dec = Decimal(rhs_str)
                except InvalidOperation:
                    rhs_dec = None
                if rhs_dec is not None and cell_dec is not None:
                    if op == ">=":
                        return cell_dec >= rhs_dec
                    if op == "<=":
                        return cell_dec <= rhs_dec
                    if op == ">":
                        return cell_dec > rhs_dec
                    if op == "<":
                        return cell_dec < rhs_dec
                    return cell_dec == rhs_dec
                if op == "=":
                    return str(cell_val or "").casefold() == rhs_str.casefold()
                return False

        # Plain string equality (case-insensitive per Excel)
        if cell_val is None:
            return criterion == ""
        return str(cell_val).casefold() == criterion.casefold()

    return cell_val == criterion


class _WorkbookFormulaEvaluator:
    """Recursive-descent Excel formula evaluator with cached cell lookups and resource limits."""

    def __init__(
        self,
        cache: WorkbookFormulaCache,
        current_sheet: str,
        *,
        limits: FormulaEvaluationLimits | None = None,
        depth: int = 0,
        evaluated_cells: list[int] | None = None,
    ) -> None:
        self.cache = cache
        self.wb = cache.wb
        self.current_sheet = current_sheet
        self.limits = limits or FormulaEvaluationLimits()
        self.depth = depth
        self._evaluated_cells = evaluated_cells if evaluated_cells is not None else [0]

    def _charge_cells(self, count: int, *, context: str) -> None:
        self._evaluated_cells[0] += count
        if self._evaluated_cells[0] > self.limits.max_evaluated_cells:
            raise FormulaResourceLimitError(
                f"Formula resource limit exceeded ({context}): evaluated "
                f"{self._evaluated_cells[0]:,} cells > "
                f"max_evaluated_cells={self.limits.max_evaluated_cells:,}."
            )

    def _check_range_budget(
        self, rng: _CellRangeRef, *, multiplier: int = 1, context: str
    ) -> _CachedSheetData:
        sheet_data = self.cache.get_sheet(rng.sheet_name)
        range_cells = (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1)
        if range_cells > self.limits.max_range_cells:
            raise FormulaResourceLimitError(
                f"Formula resource limit exceeded ({context}): range on '{rng.sheet_name}' "
                f"spans {range_cells:,} cells > max_range_cells={self.limits.max_range_cells:,}."
            )
        self._charge_cells(range_cells * multiplier, context=context)
        return sheet_data

    def evaluate(self, formula: str) -> Any:
        stripped = formula.strip()
        if len(stripped) > self.limits.max_formula_length:
            raise FormulaResourceLimitError(
                f"Formula resource limit exceeded: formula length {len(stripped):,} > "
                f"max_formula_length={self.limits.max_formula_length:,}."
            )
        if not stripped.startswith("="):
            raise ValueError(f"Expected '=' formula prefix, got: {formula!r}")
        tokens = _tokenize_formula(stripped[1:])
        self._tokens = tokens
        self._pos = 0
        if not tokens:
            raise ValueError("Empty Excel formula.")
        val = self._parse_comparison()
        if self._pos < len(self._tokens):
            tok = self._tokens[self._pos].value
            raise ValueError(f"Unexpected trailing token {tok!r} in formula {formula!r}")
        return val

    def _peek(self) -> _Token | None:
        if self._pos < len(self._tokens):
            return self._tokens[self._pos]
        return None

    def _consume(self, expected_kind: str | None = None) -> _Token:
        tok = self._peek()
        if tok is None:
            raise ValueError("Unexpected end of Excel formula.")
        if expected_kind is not None and tok.kind != expected_kind:
            raise ValueError(f"Expected {expected_kind}, got {tok.kind} ({tok.value!r}).")
        self._pos += 1
        return tok

    def _parse_comparison(self) -> Any:
        left = self._parse_additive()
        while (
            (tok := self._peek()) is not None
            and tok.kind == "OP"
            and tok.value
            in (
                "=",
                "<>",
                "<=",
                ">=",
                "<",
                ">",
            )
        ):
            op = self._consume("OP").value
            right = self._parse_additive()
            left = self._compare_values(left, right, op)
        return left

    def _parse_additive(self) -> Any:
        left = self._parse_multiplicative()
        while (tok := self._peek()) is not None and tok.kind == "OP" and tok.value in ("+", "-"):
            op = self._consume("OP").value
            right = self._parse_multiplicative()
            l_dec = self._require_decimal(left)
            r_dec = self._require_decimal(right)
            if op == "+":
                left = ARITHMETIC_CONTEXT.add(l_dec, r_dec)
            else:
                left = ARITHMETIC_CONTEXT.subtract(l_dec, r_dec)
        return left

    def _parse_multiplicative(self) -> Any:
        left = self._parse_unary()
        while (tok := self._peek()) is not None and tok.kind == "OP" and tok.value in ("*", "/"):
            op = self._consume("OP").value
            right = self._parse_unary()
            l_dec = self._require_decimal(left)
            r_dec = self._require_decimal(right)
            if op == "*":
                left = ARITHMETIC_CONTEXT.multiply(l_dec, r_dec)
            else:
                if r_dec == Decimal("0"):
                    raise ValueError("Division by zero in Excel formula.")
                left = ARITHMETIC_CONTEXT.divide(l_dec, r_dec)
        return left

    def _parse_unary(self) -> Any:
        tok = self._peek()
        if tok is not None and tok.kind == "OP" and tok.value in ("+", "-"):
            op = self._consume("OP").value
            val = self._parse_unary()
            dec = self._require_decimal(val)
            return dec if op == "+" else -dec
        return self._parse_primary()

    def _parse_primary(self) -> Any:
        tok = self._peek()
        if tok is None:
            raise ValueError("Unexpected end of expression in Excel formula.")

        if tok.kind == "NUMBER":
            self._consume("NUMBER")
            return Decimal(tok.value)
        if tok.kind == "STRING":
            self._consume("STRING")
            return tok.value
        if tok.kind == "LPAREN":
            self._consume("LPAREN")
            val = self._parse_comparison()
            self._consume("RPAREN")
            return val
        if tok.kind == "IDENT":
            ident = self._consume("IDENT").value
            nxt = self._peek()
            if nxt is not None and nxt.kind == "LPAREN":
                return self._parse_function_call(ident.upper())
            upper = ident.upper()
            if upper == "TRUE":
                return True
            if upper == "FALSE":
                return False
            ref = self._parse_ref(ident)
            if isinstance(ref, _CellRef):
                self._charge_cells(1, context="cell reference")
                return self._eval_cell(ref.sheet_name, ref.row, ref.col)
            return ref

        raise ValueError(f"Unexpected token {tok.value!r} in Excel formula.")

    def _skip_expression_tokens(self) -> None:
        """Consume and discard an unchosen branch of IF() so 1/0 is not evaluated lazily."""
        depth = 0
        while (tok := self._peek()) is not None:
            if tok.kind == "LPAREN":
                depth += 1
                self._consume()
            elif tok.kind == "RPAREN":
                if depth == 0:
                    return
                depth -= 1
                self._consume()
            elif tok.kind == "COMMA" and depth == 0:
                return
            else:
                self._consume()

    def _parse_function_call(self, fn_name: str) -> Any:
        self._consume("LPAREN")
        if fn_name == "IF":
            cond = self._parse_comparison()
            self._consume("COMMA")
            truthy = self._is_truthy(cond)
            if truthy:
                result = self._parse_comparison()
                self._consume("COMMA")
                self._skip_expression_tokens()
            else:
                self._skip_expression_tokens()
                self._consume("COMMA")
                result = self._parse_comparison()
            self._consume("RPAREN")
            return result

        args: list[Any] = []
        if (tok := self._peek()) is not None and tok.kind != "RPAREN":
            args.append(self._parse_comparison())
            while (nxt := self._peek()) is not None and nxt.kind == "COMMA":
                self._consume("COMMA")
                args.append(self._parse_comparison())
        self._consume("RPAREN")

        if fn_name == "ROUND":
            if len(args) != 2:
                raise ValueError(f"ROUND expects 2 arguments, got {len(args)}.")
            val = self._require_decimal(args[0])
            digits = int(self._require_decimal(args[1]))
            return excel_round(val, digits)

        if fn_name == "SUM":
            total = Decimal("0")
            for arg in args:
                if isinstance(arg, _CellRangeRef):
                    sheet_data = self._check_range_budget(arg, context="SUM")
                    eff_max_row = min(arg.max_row, sheet_data.max_row)
                    eff_max_col = min(arg.max_col, sheet_data.max_col)
                    if eff_max_row >= arg.min_row and eff_max_col >= arg.min_col:
                        for r in range(arg.min_row, eff_max_row + 1):
                            for c in range(arg.min_col, eff_max_col + 1):
                                c_val = self._eval_cell(arg.sheet_name, r, c)
                                dec = _to_decimal_strict(c_val)
                                if dec is not None:
                                    total = ARITHMETIC_CONTEXT.add(total, dec)
                else:
                    dec = _to_decimal_strict(arg)
                    if dec is not None:
                        total = ARITHMETIC_CONTEXT.add(total, dec)
            return total

        if fn_name == "COUNTA":
            count = 0
            for arg in args:
                if isinstance(arg, _CellRangeRef):
                    sheet_data = self._check_range_budget(arg, context="COUNTA")
                    eff_max_row = min(arg.max_row, sheet_data.max_row)
                    eff_max_col = min(arg.max_col, sheet_data.max_col)
                    if eff_max_row >= arg.min_row and eff_max_col >= arg.min_col:
                        for r in range(arg.min_row, eff_max_row + 1):
                            for c in range(arg.min_col, eff_max_col + 1):
                                c_val = self._eval_cell(arg.sheet_name, r, c)
                                if c_val is not None and c_val != "":
                                    count += 1
                elif arg is not None and arg != "":
                    count += 1
            return Decimal(count)

        if fn_name == "COUNTIF":
            if len(args) != 2 or not isinstance(args[0], _CellRangeRef):
                raise ValueError("COUNTIF expects (range, criterion).")
            rng, crit = args[0], args[1]
            self._check_range_budget(rng, context="COUNTIF")
            count = 0
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if _matches_criterion(self._eval_cell(rng.sheet_name, r, c), crit):
                        count += 1
            return Decimal(count)

        if fn_name == "COUNTIFS":
            if len(args) < 2 or len(args) % 2 != 0:
                raise ValueError("COUNTIFS expects an even number of (range, criterion) args.")
            pairs: list[tuple[_CellRangeRef, Any]] = []
            for idx in range(0, len(args), 2):
                if not isinstance(args[idx], _CellRangeRef):
                    raise ValueError("COUNTIFS criteria argument must be a range reference.")
                pairs.append((args[idx], args[idx + 1]))
            base = pairs[0][0]
            height = base.max_row - base.min_row + 1
            width = base.max_col - base.min_col + 1
            for rng, _ in pairs[1:]:
                if (rng.max_row - rng.min_row + 1) != height or (
                    rng.max_col - rng.min_col + 1
                ) != width:
                    raise ValueError("COUNTIFS ranges must have identical dimensions.")
            for rng, _ in pairs:
                self._check_range_budget(rng, context="COUNTIFS")
            count = 0
            for dr in range(height):
                for dc in range(width):
                    if all(
                        _matches_criterion(
                            self._eval_cell(rng.sheet_name, rng.min_row + dr, rng.min_col + dc),
                            crit,
                        )
                        for rng, crit in pairs
                    ):
                        count += 1
            return Decimal(count)

        if fn_name == "SUMIF":
            if len(args) not in (2, 3) or not isinstance(args[0], _CellRangeRef):
                raise ValueError("SUMIF expects (criteria_range, criterion, [sum_range]).")
            c_rng = args[0]
            crit = args[1]
            s_rng = args[2] if len(args) == 3 else c_rng
            if not isinstance(s_rng, _CellRangeRef):
                raise ValueError("SUMIF sum_range must be a range reference.")
            self._check_range_budget(c_rng, context="SUMIF criteria_range")
            if s_rng is not c_rng:
                self._check_range_budget(s_rng, context="SUMIF sum_range")
            height = c_rng.max_row - c_rng.min_row + 1
            width = c_rng.max_col - c_rng.min_col + 1
            total = Decimal("0")
            for dr in range(height):
                for dc in range(width):
                    c_val = self._eval_cell(
                        c_rng.sheet_name, c_rng.min_row + dr, c_rng.min_col + dc
                    )
                    if _matches_criterion(c_val, crit):
                        s_val = self._eval_cell(
                            s_rng.sheet_name, s_rng.min_row + dr, s_rng.min_col + dc
                        )
                        dec = _to_decimal_strict(s_val)
                        if dec is not None:
                            total = ARITHMETIC_CONTEXT.add(total, dec)
            return total

        if fn_name == "SUMIFS":
            if len(args) < 3 or (len(args) - 1) % 2 != 0 or not isinstance(args[0], _CellRangeRef):
                raise ValueError("SUMIFS expects (sum_range, criteria_range1, crit1, ...).")
            s_rng = args[0]
            pairs = []
            for idx in range(1, len(args), 2):
                if not isinstance(args[idx], _CellRangeRef):
                    raise ValueError("SUMIFS criteria_range must be a range reference.")
                pairs.append((args[idx], args[idx + 1]))
            height = s_rng.max_row - s_rng.min_row + 1
            width = s_rng.max_col - s_rng.min_col + 1
            for rng, _ in pairs:
                if (rng.max_row - rng.min_row + 1) != height or (
                    rng.max_col - rng.min_col + 1
                ) != width:
                    raise ValueError("SUMIFS ranges must have identical dimensions.")
            self._check_range_budget(s_rng, context="SUMIFS sum_range")
            for rng, _ in pairs:
                self._check_range_budget(rng, context="SUMIFS criteria_range")
            total = Decimal("0")
            for dr in range(height):
                for dc in range(width):
                    if all(
                        _matches_criterion(
                            self._eval_cell(rng.sheet_name, rng.min_row + dr, rng.min_col + dc),
                            crit,
                        )
                        for rng, crit in pairs
                    ):
                        s_val = self._eval_cell(
                            s_rng.sheet_name, s_rng.min_row + dr, s_rng.min_col + dc
                        )
                        dec = _to_decimal_strict(s_val)
                        if dec is not None:
                            total = ARITHMETIC_CONTEXT.add(total, dec)
            return total

        raise ValueError(f"Unsupported Excel function {fn_name!r} in report formula.")

    def _parse_ref(self, token_text: str) -> _CellRef | _CellRangeRef:
        m_range = _RANGE_REF_RE.match(token_text)
        if m_range:
            quoted_sheet, bare_sheet, c1, r1, c2, r2 = m_range.groups()
            sheet = quoted_sheet or bare_sheet or self.current_sheet
            col1 = column_index_from_string(c1.upper())
            col2 = column_index_from_string(c2.upper())
            row1 = int(r1)
            row2 = int(r2)
            return _CellRangeRef(
                sheet_name=sheet,
                min_col=min(col1, col2),
                min_row=min(row1, row2),
                max_col=max(col1, col2),
                max_row=max(row1, row2),
            )
        m_cell = _CELL_REF_RE.match(token_text)
        if m_cell:
            quoted_sheet, bare_sheet, col_str, row_str = m_cell.groups()
            sheet = quoted_sheet or bare_sheet or self.current_sheet
            return _CellRef(
                sheet_name=sheet,
                col=column_index_from_string(col_str.upper()),
                row=int(row_str),
            )
        raise ValueError(f"Unsupported Excel identifier/reference: {token_text!r}")

    def _eval_cell(self, sheet_name: str, row: int, col: int) -> Any:
        sheet_data = self.cache.get_sheet(sheet_name)
        entry = sheet_data.cells.get((row, col))
        if entry is None:
            return None
        val, dtype = entry
        if dtype == "e":
            raise ValueError(f"Referenced cell {sheet_name}!R{row}C{col} has Excel error {val!r}.")
        if dtype == "f" and isinstance(val, str):
            cell_key = (sheet_name, row, col)
            if cell_key in self.cache.formula_results:
                return self.cache.formula_results[cell_key]
            if cell_key in self.cache.active_cells:
                raise ValueError(f"Circular formula reference at {sheet_name}!R{row}C{col}.")
            self.cache.active_cells.add(cell_key)
            try:
                evaluated = evaluate_workbook_formula(
                    self.wb,
                    sheet_name,
                    val,
                    limits=self.limits,
                    cache=self.cache,
                    _depth=self.depth + 1,
                    _evaluated_cells=self._evaluated_cells,
                )
                self.cache.formula_results[cell_key] = evaluated
                return evaluated
            finally:
                self.cache.active_cells.discard(cell_key)
        return val

    def _require_decimal(self, value: Any) -> Decimal:
        if value is None:
            return Decimal("0")
        dec = _to_decimal_strict(value)
        if dec is not None:
            return dec
        raise ValueError(f"Expected numeric value in formula arithmetic, got {value!r}")

    @staticmethod
    def _is_truthy(value: Any) -> bool:
        if isinstance(value, bool):
            return value
        dec = _to_decimal_strict(value)
        if dec is not None:
            return dec != Decimal("0")
        return bool(value)

    @staticmethod
    def _compare_values(left: Any, right: Any, op: str) -> bool:
        l_dec = _to_decimal_strict(left)
        r_dec = _to_decimal_strict(right)
        if l_dec is not None and r_dec is not None:
            if op == "=":
                return l_dec == r_dec
            if op == "<>":
                return l_dec != r_dec
            if op == "<=":
                return l_dec <= r_dec
            if op == ">=":
                return l_dec >= r_dec
            if op == "<":
                return l_dec < r_dec
            return l_dec > r_dec
        l_norm = "" if left is None else str(left).casefold()
        r_norm = "" if right is None else str(right).casefold()
        if op == "=":
            return l_norm == r_norm
        if op == "<>":
            return l_norm != r_norm
        if op == "<=":
            return l_norm <= r_norm
        if op == ">=":
            return l_norm >= r_norm
        if op == "<":
            return l_norm < r_norm
        return l_norm > r_norm


def evaluate_workbook_formula(
    wb: openpyxl.Workbook,
    current_sheet: str,
    formula: str,
    *,
    limits: FormulaEvaluationLimits | None = None,
    cache: WorkbookFormulaCache | None = None,
    _depth: int = 0,
    _evaluated_cells: list[int] | None = None,
) -> Decimal:
    """Evaluate an Excel formula against an in-memory openpyxl Workbook and return Decimal."""
    active_limits = limits or FormulaEvaluationLimits()
    if _depth > active_limits.max_recursion_depth:
        raise FormulaResourceLimitError(
            f"Exceeded formula evaluation recursion limit "
            f"(max_recursion_depth={active_limits.max_recursion_depth}) on {formula!r}"
        )
    active_cache = cache if cache is not None else WorkbookFormulaCache(wb)
    evaluator = _WorkbookFormulaEvaluator(
        active_cache,
        current_sheet,
        limits=active_limits,
        depth=_depth,
        evaluated_cells=_evaluated_cells,
    )
    raw = evaluator.evaluate(formula)
    if isinstance(raw, bool):
        return Decimal("1") if raw else Decimal("0")
    dec = _to_decimal_strict(raw)
    if dec is None:
        raise ValueError(
            f"Formula {formula!r} did not evaluate to a finite numeric result: {raw!r}"
        )
    return dec


def _resolve_summary_xml_path(zf: zipfile.ZipFile) -> str:
    """Locate the OpenXML worksheet path corresponding to the 'Summary' sheet."""
    default_path = "xl/worksheets/sheet1.xml"
    try:
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except Exception:
        return default_path

    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rid_by_name: dict[str, str] = {}
    for elem in wb_root.iter():
        if elem.tag.endswith("}sheet") or elem.tag == "sheet":
            name = elem.attrib.get("name", "")
            rid = elem.attrib.get(f"{{{r_ns}}}id") or elem.attrib.get("r:id", "")
            if name and rid:
                rid_by_name[name] = rid

    target_rid = rid_by_name.get("Summary")
    if not target_rid:
        return default_path

    for rel in rels_root.iter():
        if (rel.tag.endswith("}Relationship") or rel.tag == "Relationship") and rel.attrib.get(
            "Id"
        ) == target_rid:
            target = rel.attrib.get("Target", "").lstrip("/")
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            return target
    return default_path


def _read_xml_formula_caches(
    report_path: Path,
) -> tuple[dict[str, tuple[str, str | None]], list[str]]:
    """Read raw OpenXML <f> and cached <v> values from the Summary worksheet XML."""
    caches: dict[str, tuple[str, str | None]] = {}
    xml_errors: list[str] = []
    with zipfile.ZipFile(report_path, "r") as zf:
        summary_xml_path = _resolve_summary_xml_path(zf)
        if summary_xml_path not in zf.namelist():
            return caches, [f"Missing worksheet XML '{summary_xml_path}' inside workbook archive."]
        root = ET.fromstring(zf.read(summary_xml_path))
        for c_elem in root.iter():
            if not (c_elem.tag.endswith("}c") or c_elem.tag == "c"):
                continue
            coord = c_elem.attrib.get("r", "")
            cell_type = c_elem.attrib.get("t", "")
            if cell_type == "e":
                xml_errors.append(f"Summary!{coord} has OpenXML error type t='e'.")
            f_elem = None
            v_elem = None
            for child in c_elem:
                if child.tag.endswith("}f") or child.tag == "f":
                    f_elem = child
                elif child.tag.endswith("}v") or child.tag == "v":
                    v_elem = child
            if f_elem is not None:
                if cell_type not in ("", "n"):
                    xml_errors.append(
                        f"Summary!{coord} formula cache must have numeric OpenXML type."
                    )
                f_text = f_elem.text or ""
                v_text = v_elem.text if v_elem is not None else None
                caches[coord] = (f_text, v_text)
    return caches, xml_errors


def _column_sum_decimal(sheet_data: _CachedSheetData, col_idx: int) -> Decimal:
    """Sum all finite numeric values in rows 2..max_row for a cached worksheet column."""
    total = Decimal("0")
    for r in range(2, sheet_data.max_row + 1):
        entry = sheet_data.cells.get((r, col_idx))
        if entry is not None:
            dec = _to_decimal_strict(entry[0])
            if dec is not None:
                total = ARITHMETIC_CONTEXT.add(total, dec)
    return total


_RECONCILED_HEADERS: tuple[str, ...] = (
    "Source Row",
    "Raw PO Number",
    "Raw PO Line ID",
    "Raw SKU",
    "Normalized PO Number",
    "Normalized PO Line ID",
    "Normalized SKU",
    "Canonical SKU",
    "Matched Business Key",
    "Product Match Tier",
    "Product Description",
    "Supplier ID",
    "Ordered Qty",
    "Shipped Qty",
    "Invoiced Qty",
    "Outstanding Qty",
    "PO Unit Price",
    "Expected Amount",
    "Actual Invoiced Amount",
    "Monetary Variance",
    "Reconciliation Status",
    "Reason Codes",
    "Shipment Count",
    "Invoice Count",
    "Shipment References",
    "Invoice References",
    "Preserved PO Formula",
    "Valid PO Line",
    "Discrepancy Exposure",
    "Unbilled Receipt Accrual",
)

_EXCEPTION_HEADERS: tuple[str, ...] = (
    "Exception ID",
    "Exception Category",
    "Severity",
    "Source Dataset",
    "Source Row",
    "PO Reference",
    "Raw Reference",
    "Normalized Reference",
    "Raw SKU",
    "Normalized SKU",
    "Reason Codes",
    "Quantity Impact",
    "Monetary Impact",
    "Explanation",
    "Suggested Next Action",
)


def _validate_report_contract(wb: openpyxl.Workbook, errors: list[str]) -> None:
    """Check emitted schema and formula dependencies independently of saved KPI values."""
    ws_rec = wb["Reconciled Records"]
    ws_exc = wb["Exceptions"]
    ws_sum = wb["Summary"]
    for ws, headers in ((ws_rec, _RECONCILED_HEADERS), (ws_exc, _EXCEPTION_HEADERS)):
        actual = tuple(ws.cell(1, col).value for col in range(1, len(headers) + 1))
        if actual != headers:
            errors.append(f"Worksheet '{ws.title}' has incorrect report column headers.")

    def numeric(
        cell: Any, *, required: bool = False, integer: bool = False, minimum: int | None = None
    ) -> Decimal | None:
        value = cell.value
        if value is None and not required:
            return None
        dec = _to_decimal_strict(value)
        if (
            cell.data_type != "n"
            or dec is None
            or (integer and dec != dec.to_integral_value())
            or (minimum is not None and dec < minimum)
        ):
            errors.append(f"Invalid numeric cell {cell.parent.title}!{cell.coordinate}: {value!r}.")
            return None
        return dec

    def enum_cell(cell: Any, choices: set[str]) -> None:
        if cell.data_type != "s" or cell.value not in choices:
            errors.append(
                f"Invalid category at {cell.parent.title}!{cell.coordinate}: {cell.value!r}."
            )

    source_rows: set[Decimal] = set()
    rec_has_rows = False
    for row in range(2, ws_rec.max_row + 1):
        if all(ws_rec.cell(row, col).value is None for col in range(1, 31)):
            continue
        rec_has_rows = True
        source = numeric(ws_rec.cell(row, 1), required=True, integer=True, minimum=2)
        if source is not None:
            if source in source_rows:
                errors.append(f"Repeated PO source row {source} in Reconciled Records.")
            source_rows.add(source)
        enum_cell(ws_rec.cell(row, 10), {s.value for s in MatchStatus})
        enum_cell(ws_rec.cell(row, 21), {s.value for s in ReconciliationStatus})
        valid_cell = ws_rec.cell(row, 28)
        if valid_cell.data_type != "b" or not isinstance(valid_cell.value, bool):
            errors.append(f"Reconciled Records!AB{row} must be a Boolean cell.")
        values = {col: numeric(ws_rec.cell(row, col)) for col in (*range(13, 21), 29, 30)}
        for col in (23, 24):
            numeric(ws_rec.cell(row, col), required=True, integer=True, minimum=0)
        if valid_cell.value is True:
            for col in (13, 16, 17, 18, 29, 30):
                if ws_rec.cell(row, col).value is None:
                    errors.append(
                        f"Missing required value at Reconciled Records!"
                        f"{get_column_letter(col)}{row}."
                    )
            qty, shipped, price, expected, outstanding = (values[c] for c in (13, 14, 17, 18, 16))
            if qty is not None and price is not None and expected is not None:
                try:
                    contracted = excel_round(ARITHMETIC_CONTEXT.multiply(qty, price), 2)
                except InvalidOperation:
                    errors.append(f"Reconciled Records!R{row} exceeds monetary precision limits.")
                else:
                    if abs(expected - contracted) > Decimal("0.0001"):
                        errors.append(
                            f"Reconciled Records!R{row} does not equal rounded quantity * price."
                        )
            if qty is not None and outstanding is not None:
                expected_outstanding = max(
                    ARITHMETIC_CONTEXT.subtract(qty, shipped or Decimal(0)), Decimal(0)
                )
                if outstanding != expected_outstanding:
                    errors.append(f"Reconciled Records!P{row} has incorrect outstanding quantity.")
        elif valid_cell.value is False:
            for col in (14, 15, 16, 19, 20, 29, 30):
                if ws_rec.cell(row, col).value is not None:
                    errors.append(
                        f"Excluded PO row contributes at Reconciled Records!"
                        f"{get_column_letter(col)}{row}."
                    )

    exc_has_rows = False
    exception_ids: set[str] = set()
    for row in range(2, ws_exc.max_row + 1):
        if all(ws_exc.cell(row, col).value is None for col in range(1, 16)):
            continue
        exc_has_rows = True
        exception_id = ws_exc.cell(row, 1).value
        if not isinstance(exception_id, str) or not exception_id or exception_id in exception_ids:
            errors.append(f"Invalid or repeated exception ID at Exceptions!A{row}.")
        else:
            exception_ids.add(exception_id)
        enum_cell(ws_exc.cell(row, 2), {s.value for s in ExceptionCategory})
        enum_cell(ws_exc.cell(row, 3), {s.value for s in Severity})
        enum_cell(ws_exc.cell(row, 4), {s.value for s in DatasetKind})
        numeric(ws_exc.cell(row, 5), required=True, integer=True, minimum=2)
        numeric(ws_exc.cell(row, 12))
        numeric(ws_exc.cell(row, 13))

    # Each non-price PO status has one primary PO exception with the same source row.
    po_exception_keys = {
        (ws_exc.cell(row, 5).value, ws_exc.cell(row, 2).value)
        for row in range(2, ws_exc.max_row + 1)
        if ws_exc.cell(row, 4).value == DatasetKind.PURCHASE_ORDERS.value
    }
    price_statuses: dict[str, int] = {}
    price_exceptions: dict[str, int] = {}
    for row in range(2, ws_exc.max_row + 1):
        if ws_exc.cell(row, 2).value == ExceptionCategory.PRICE_MISMATCH.value:
            po = str(ws_exc.cell(row, 6).value or "")
            price_exceptions[po] = price_exceptions.get(po, 0) + 1
    for row in range(2, ws_rec.max_row + 1):
        status_value = ws_rec.cell(row, 21).value
        if status_value is None:
            continue
        if status_value == ReconciliationStatus.PRICE_MISMATCH.value:
            po = str(ws_rec.cell(row, 5).value or "")
            price_statuses[po] = price_statuses.get(po, 0) + 1
        elif status_value != ReconciliationStatus.RECONCILED.value:
            if (ws_rec.cell(row, 1).value, status_value) not in po_exception_keys:
                errors.append(f"Reconciled Records!U{row} lacks its primary source-row exception.")
    if price_statuses != price_exceptions:
        errors.append("PRICE_MISMATCH statuses do not match price exceptions by PO number.")

    last_rec = max(ws_rec.max_row, 2)
    last_exc = max(ws_exc.max_row, 2)

    def rec_range(col: str) -> str:
        return f"'Reconciled Records'!${col}$2:${col}${last_rec}"

    def exc_range(col: str) -> str:
        return f"Exceptions!${col}$2:${col}${last_exc}"

    valid, tiers, status = (rec_range(col) for col in ("AB", "J", "U"))
    formulas = [
        f"=COUNTA({rec_range('A')})",
        f"=COUNTIF({valid},TRUE)",
        f'=COUNTIF({status},"RECONCILED")',
        f'=COUNTIF({status},"PARTIAL_FULFILLMENT")',
        f'=COUNTIF({status},"<>RECONCILED")',
        f"=COUNTA({exc_range('A')})" if exc_has_rows else "=0",
        f'=IF(B6=0,0,ROUND((COUNTIFS({valid},TRUE,{tiers},"EXACT_ID")+'
        f'COUNTIFS({valid},TRUE,{tiers},"ALIAS_MAP"))/B6,4))',
        "=IF(B6=0,0,ROUND(B7/B6,4))",
        f"=SUMIF({valid},TRUE,{rec_range('M')})",
        f"=SUM({rec_range('N')})",
        f"=SUM({rec_range('O')})",
        f"=SUM({rec_range('P')})",
        f"=ROUND(SUMIF({valid},TRUE,{rec_range('R')}),2)",
        f"=ROUND(SUM({rec_range('S')}),2)",
        f"=ROUND(SUM({rec_range('T')}),2)",
        f"=ROUND(SUM({rec_range('AC')}),2)",
        f"=ROUND(SUM({rec_range('AD')}),2)",
        f'=ROUND(SUMIFS({exc_range("M")},{exc_range("D")},"supplier_invoices",'
        f'{exc_range("B")},"<>PRICE_MISMATCH"),2)'
        if exc_has_rows
        else "=0",
    ]
    if not rec_has_rows:
        formulas = [
            formula if row in (10, 12, 22) else "=0"
            for row, formula in enumerate(formulas, start=5)
        ]
    expected_formulas = {f"C{row}": formula for row, formula in enumerate(formulas, start=5)}
    for row in range(26, 35):
        expected_formulas[f"C{row}"] = (
            f"=COUNTIF({exc_range('B')},A{row})" if exc_has_rows else "=0"
        )
        expected_formulas[f"D{row}"] = f"=IF($B$35=0,0,B{row}/$B$35)"
    expected_formulas.update(C35="=SUM(C26:C34)", D35="=SUM(D26:D34)")

    def tokens(formula: str) -> list[tuple[str, str]]:
        return [
            (t.kind, t.value.casefold() if t.kind == "IDENT" else t.value)
            for t in _tokenize_formula(formula[1:])
        ]

    expected_labels = dict(enumerate(REQUIRED_SUMMARY_KPI_LABELS, start=5))
    expected_labels.update({row: cat.value for row, cat in enumerate(ExceptionCategory, start=26)})
    expected_labels[35] = "TOTAL EXCEPTIONS"
    for row, label in expected_labels.items():
        if ws_sum.cell(row, 1).value != label:
            errors.append(f"Unexpected Summary label at A{row}.")
    for coord, expected_formula in expected_formulas.items():
        cell = ws_sum[coord]
        try:
            matches = (
                cell.data_type == "f"
                and isinstance(cell.value, str)
                and tokens(cell.value) == tokens(expected_formula)
            )
        except ValueError:
            matches = False
        if not matches:
            errors.append(f"Unexpected Summary formula at {coord}; expected {expected_formula!r}.")
    for row in ws_sum:
        for cell in row:
            if cell.data_type == "f" and cell.coordinate not in expected_formulas:
                errors.append(f"Unauthorized Summary formula at {cell.coordinate}.")


def validate_report_workbook(
    report_path: Path,
    *,
    limits: FormulaEvaluationLimits | None = None,
) -> ValidationReport:
    """Inspect reconciliation_report.xlsx for structure, formula safety, cache, and KPIs."""
    if not report_path.exists():
        return ValidationReport(
            report_path=str(report_path),
            is_valid=False,
            worksheets_found=(),
            formula_count=0,
            broken_formula_cells=(),
            errors=(f"Report file does not exist: {report_path}",),
            warnings=(),
        )

    errors: list[str] = []
    warnings: list[str] = []
    broken_cells: list[str] = []
    formula_count = 0

    try:
        wb = openpyxl.load_workbook(report_path, data_only=False)
    except Exception as exc:
        return ValidationReport(
            report_path=str(report_path),
            is_valid=False,
            worksheets_found=(),
            formula_count=0,
            broken_formula_cells=(),
            errors=(f"Failed to open Excel workbook: {exc}",),
            warnings=(),
        )

    try:
        sheets_found = tuple(wb.sheetnames)
        for req_sheet in REQUIRED_WORKSHEETS:
            if req_sheet not in wb.sheetnames:
                errors.append(f"Missing required worksheet: '{req_sheet}'.")

        # 1. Inspect every cell across all sheets for error tokens, unquoted text, or rogue formulas
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if ws.freeze_panes is None:
                warnings.append(f"Worksheet '{sheet_name}' does not have frozen header panes.")

            for row in ws.iter_rows():
                for cell in row:
                    val = cell.value
                    coord = f"{sheet_name}!{cell.coordinate}"
                    if cell.data_type == "e":
                        broken_cells.append(f"{coord} ({val})")
                    elif cell.data_type == "f" and isinstance(val, str):
                        formula_count += 1
                        upper = val.upper()
                        if any(tok in upper for tok in EXCEL_ERROR_TOKENS):
                            broken_cells.append(f"{coord} ({val})")
                        elif sheet_name != "Summary" or cell.column not in (3, 4):
                            broken_cells.append(f"{coord} ({val})")
                            errors.append(
                                f"Unexpected live formula outside Summary check columns at {coord}."
                            )
                    elif (
                        isinstance(val, str)
                        and sheet_name in ("Reconciled Records", "Exceptions", "Audit Log")
                        and (
                            val.startswith(_DANGEROUS_TEXT_PREFIXES)
                            or val.lstrip(" ").startswith(_DANGEROUS_TEXT_PREFIXES)
                        )
                        and (cell.data_type != "s" or not cell.quotePrefix)
                    ):
                        errors.append(
                            f"Unsafe unquoted formula-like text cell at {coord}: {val!r}."
                        )

        if broken_cells:
            errors.append(
                f"Detected {len(broken_cells)} broken or unauthorized Excel formula/error cell(s): "
                + ", ".join(broken_cells[:10])
            )

        if formula_count == 0:
            errors.append("Workbook contains zero live Excel formulas on Summary worksheet.")

        # 2. Cross-sheet KPI, distribution, formula evaluation, and XML <v> cache parity checks
        if all(s in wb.sheetnames for s in REQUIRED_WORKSHEETS):
            _validate_report_contract(wb, errors)
            active_limits = limits or FormulaEvaluationLimits()
            wb_cache = WorkbookFormulaCache(wb)
            ws_sum = wb["Summary"]
            ws_rec = wb["Reconciled Records"]
            ws_exc = wb["Exceptions"]
            ws_aud = wb["Audit Log"]
            sheet_sum = wb_cache.get_sheet("Summary")
            sheet_rec = wb_cache.get_sheet("Reconciled Records")
            sheet_exc = wb_cache.get_sheet("Exceptions")

            for tab_name, tab_ws in (
                ("Reconciled Records", ws_rec),
                ("Exceptions", ws_exc),
                ("Audit Log", ws_aud),
            ):
                if not tab_ws.auto_filter.ref:
                    errors.append(f"Worksheet '{tab_name}' is missing auto-filter configuration.")

            kpi_row_by_label: dict[str, int] = {}
            kpi_by_label: dict[str, object] = {}
            for r_idx in range(5, sheet_sum.max_row + 1):
                lbl_entry = sheet_sum.cells.get((r_idx, 1))
                val_entry = sheet_sum.cells.get((r_idx, 2))
                lbl = lbl_entry[0] if lbl_entry is not None else None
                val = val_entry[0] if val_entry is not None else None
                if isinstance(lbl, str) and lbl.strip():
                    clean_lbl = lbl.strip()
                    kpi_row_by_label[clean_lbl] = r_idx
                    kpi_by_label[clean_lbl] = val

            # Require all documented KPI and exception distribution labels with finite numeric Col B
            for req_label in REQUIRED_SUMMARY_KPI_LABELS:
                if req_label not in kpi_by_label:
                    errors.append(f"Summary worksheet missing required KPI row: '{req_label}'.")
                elif _to_decimal_strict(kpi_by_label[req_label]) is None:
                    errors.append(
                        f"Summary KPI '{req_label}' must contain finite numeric Col B "
                        f"(got {kpi_by_label[req_label]!r})."
                    )

            required_dist_labels = [cat.value for cat in ExceptionCategory] + ["TOTAL EXCEPTIONS"]
            for dist_label in required_dist_labels:
                if dist_label not in kpi_by_label:
                    errors.append(
                        f"Summary worksheet missing Exception Distribution row: '{dist_label}'."
                    )
                elif _to_decimal_strict(kpi_by_label[dist_label]) is None:
                    errors.append(
                        f"Summary distribution '{dist_label}' must contain a finite numeric value "
                        f"in Column B (got {kpi_by_label[dist_label]!r})."
                    )

            # Verify Col B matches actual data worksheet counts & column sums
            rec_rows = sum(1 for r in range(2, sheet_rec.max_row + 1) if (r, 1) in sheet_rec.cells)
            exc_rows = sum(1 for r in range(2, sheet_exc.max_row + 1) if (r, 1) in sheet_exc.cells)

            total_po_dec = _to_decimal_strict(kpi_by_label.get("Total Purchase Order Rows"))
            reconciled_dec = _to_decimal_strict(
                kpi_by_label.get("Successfully Reconciled PO Lines")
            )
            review_dec = _to_decimal_strict(kpi_by_label.get("PO Lines Requiring Review"))
            total_exc_dec = _to_decimal_strict(
                kpi_by_label.get("Total Exception Records (All Sources)")
            )
            dist_total_dec = _to_decimal_strict(kpi_by_label.get("TOTAL EXCEPTIONS"))

            if total_po_dec is not None and total_po_dec != Decimal(rec_rows):
                errors.append(
                    f"Summary 'Total Purchase Order Rows' ({total_po_dec}) != "
                    f"Reconciled Records row count ({rec_rows})."
                )
            if (
                total_po_dec is not None
                and reconciled_dec is not None
                and review_dec is not None
                and (reconciled_dec + review_dec != total_po_dec)
            ):
                errors.append(
                    f"Summary reconciled ({reconciled_dec}) + review ({review_dec}) != "
                    f"total PO rows ({total_po_dec})."
                )
            if total_exc_dec is not None and total_exc_dec != Decimal(exc_rows):
                errors.append(
                    f"Summary 'Total Exception Records' ({total_exc_dec}) != "
                    f"Exceptions worksheet row count ({exc_rows})."
                )
            if (
                total_exc_dec is not None
                and dist_total_dec is not None
                and total_exc_dec != dist_total_dec
            ):
                errors.append(
                    f"Exception distribution total ({dist_total_dec}) != "
                    f"Summary exception count ({total_exc_dec})."
                )

            column_kpi_checks = (
                ("Total Shipped Quantity", 14, Decimal("0.001")),
                ("Total Invoiced Quantity", 15, Decimal("0.001")),
                ("Outstanding Unshipped Quantity", 16, Decimal("0.001")),
                ("Total Actual Invoiced Amount", 19, Decimal("0.01")),
                ("Net Billed Monetary Variance", 20, Decimal("0.01")),
                ("Absolute Monetary Discrepancy Exposure", 29, Decimal("0.01")),
                ("Unbilled Shipped Value (Accrual Exposure)", 30, Decimal("0.01")),
            )
            for label, col_idx, tol in column_kpi_checks:
                expected_dec = _to_decimal_strict(kpi_by_label.get(label))
                if expected_dec is not None:
                    actual_dec = _column_sum_decimal(sheet_rec, col_idx)
                    if abs(expected_dec - actual_dec) > tol:
                        errors.append(
                            f"Summary '{label}' ({expected_dec}) != "
                            f"Reconciled Records column {col_idx} ({get_column_letter(col_idx)}) "
                            f"sum ({actual_dec})."
                        )

            # 3-way parity: Column B == Independent Formula Evaluation == OpenXML <v> cache
            xml_caches, xml_errors = _read_xml_formula_caches(report_path)
            errors.extend(xml_errors)

            parity_tol = Decimal("0.0001")
            checked_labels = list(REQUIRED_SUMMARY_KPI_LABELS) + required_dist_labels
            checked_coords: set[str] = set()
            for lbl in checked_labels:
                r_idx = kpi_row_by_label.get(lbl)
                if r_idx is None:
                    continue
                b_entry = sheet_sum.cells.get((r_idx, 2))
                col_b_dec = _to_decimal_strict(b_entry[0] if b_entry is not None else None)
                c_cell = ws_sum.cell(row=r_idx, column=3)
                checked_coords.add(f"C{r_idx}")
                if c_cell.data_type != "f" or not isinstance(c_cell.value, str):
                    errors.append(
                        f"Summary formula check cell C{r_idx} ('{lbl}') is not a live formula."
                    )
                    continue

                try:
                    eval_c = evaluate_workbook_formula(
                        wb,
                        "Summary",
                        c_cell.value,
                        limits=active_limits,
                        cache=wb_cache,
                    )
                except Exception as exc:
                    errors.append(
                        f"Failed evaluating Summary formula C{r_idx} ({c_cell.value!r}): {exc}"
                    )
                    continue

                if col_b_dec is not None and abs(col_b_dec - eval_c) > parity_tol:
                    errors.append(
                        f"Summary '{lbl}' Col B ({col_b_dec}) != evaluated C{r_idx} ({eval_c})."
                    )

                coord_c = f"C{r_idx}"
                cache_c = xml_caches.get(coord_c)
                if cache_c is None or cache_c[1] is None or not str(cache_c[1]).strip():
                    errors.append(f"Summary formula cell {coord_c} is missing OpenXML <v> cache.")
                else:
                    try:
                        cached_c_dec = Decimal(str(cache_c[1]))
                    except InvalidOperation:
                        errors.append(
                            f"Summary cell {coord_c} has non-numeric <v> cache: {cache_c[1]!r}."
                        )
                    else:
                        if not cached_c_dec.is_finite() or abs(cached_c_dec - eval_c) > parity_tol:
                            errors.append(
                                f"Summary {coord_c} XML <v> cache ({cached_c_dec}) != "
                                f"evaluated formula ({eval_c})."
                            )

                # Distribution rows also validate Column D share formulas and their XML caches
                if lbl in required_dist_labels:
                    d_cell = ws_sum.cell(row=r_idx, column=4)
                    checked_coords.add(f"D{r_idx}")
                    if d_cell.data_type != "f" or not isinstance(d_cell.value, str):
                        errors.append(
                            f"Summary distribution share cell D{r_idx} ('{lbl}') is not a formula."
                        )
                        continue
                    try:
                        eval_d = evaluate_workbook_formula(
                            wb,
                            "Summary",
                            d_cell.value,
                            limits=active_limits,
                            cache=wb_cache,
                        )
                    except Exception as exc:
                        errors.append(
                            f"Failed evaluating Summary D{r_idx} ({d_cell.value!r}): {exc}"
                        )
                        continue
                    coord_d = f"D{r_idx}"
                    cache_d = xml_caches.get(coord_d)
                    if cache_d is None or cache_d[1] is None or not str(cache_d[1]).strip():
                        errors.append(
                            f"Summary formula cell {coord_d} is missing OpenXML <v> cache."
                        )
                    else:
                        try:
                            cached_d_dec = Decimal(str(cache_d[1]))
                        except InvalidOperation:
                            errors.append(
                                f"Summary cell {coord_d} has non-numeric <v> cache: {cache_d[1]!r}."
                            )
                        else:
                            if (
                                not cached_d_dec.is_finite()
                                or abs(cached_d_dec - eval_d) > parity_tol
                            ):
                                errors.append(
                                    f"Summary {coord_d} XML <v> cache ({cached_d_dec}) != "
                                    f"evaluated formula ({eval_d})."
                                )

            # Ensure any other Summary formula cells also obey resource & syntax limits
            for (r_idx, c_idx), (val, dtype) in sorted(sheet_sum.cells.items()):
                if dtype != "f" or not isinstance(val, str):
                    continue
                coord = f"{get_column_letter(c_idx)}{r_idx}"
                if coord in checked_coords:
                    continue
                try:
                    evaluate_workbook_formula(
                        wb,
                        "Summary",
                        val,
                        limits=active_limits,
                        cache=wb_cache,
                    )
                except Exception as exc:
                    errors.append(f"Failed evaluating Summary formula {coord} ({val!r}): {exc}")

        return ValidationReport(
            report_path=str(report_path),
            is_valid=(len(errors) == 0),
            worksheets_found=sheets_found,
            formula_count=formula_count,
            broken_formula_cells=tuple(broken_cells),
            errors=tuple(errors),
            warnings=tuple(warnings),
        )
    finally:
        wb.close()
