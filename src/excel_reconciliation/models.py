"""Domain models, structured business keys, enumerations, and financial rounding rules."""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Context, Decimal, localcontext
from enum import StrEnum

ARITHMETIC_CONTEXT = Context(prec=50, rounding=ROUND_HALF_UP)
MAX_DECIMAL_MAGNITUDE = Decimal("1000000000000")


def excel_round(value: Decimal, digits: int = 2) -> Decimal:
    """Round a Decimal using Excel-compatible commercial half-up (away-from-zero) semantics."""
    with localcontext(ARITHMETIC_CONTEXT):
        quantum = Decimal("1").scaleb(-digits)
        return value.quantize(quantum, rounding=ROUND_HALF_UP)


def quantize_rate(
    numerator: int | Decimal,
    denominator: int | Decimal,
    quantum: Decimal = Decimal("0.0001"),
) -> Decimal:
    """Compute a ratio quantized with commercial ROUND_HALF_UP semantics."""
    if Decimal(denominator) == Decimal("0"):
        return Decimal("0").quantize(quantum, rounding=ROUND_HALF_UP)
    with localcontext(ARITHMETIC_CONTEXT):
        ratio = Decimal(numerator) / Decimal(denominator)
        return ratio.quantize(quantum, rounding=ROUND_HALF_UP)


def _escape_key_component(part: str) -> str:
    """Escape backslashes and pipe separators so formatted display keys are injective."""
    return part.replace("\\", "\\\\").replace("|", "\\|")


def format_business_key(po_number: str, po_line_id: str, product_scope: str) -> str:
    """Format an injective, human-auditable display key from structured key components."""
    return (
        f"{_escape_key_component(po_number)}|"
        f"{_escape_key_component(po_line_id)}|"
        f"{_escape_key_component(product_scope)}"
    )


class DatasetKind(StrEnum):
    """Supported input workbook types."""

    PRODUCTS = "product_master"
    PURCHASE_ORDERS = "purchase_orders"
    SHIPMENTS = "shipments"
    INVOICES = "supplier_invoices"


class MatchStatus(StrEnum):
    """Deterministic record matching status ordered by resolution priority."""

    EXACT_ID = "EXACT_ID"
    ALIAS_MAP = "ALIAS_MAP"
    COMPOSITE_KEY = "COMPOSITE_KEY"
    AMBIGUOUS = "AMBIGUOUS"
    UNMATCHED = "UNMATCHED"
    INVALID_KEY = "INVALID_KEY"


class ReconciliationStatus(StrEnum):
    """Primary reconciliation status assigned to each purchase order line."""

    RECONCILED = "RECONCILED"
    PARTIAL_FULFILLMENT = "PARTIAL_FULFILLMENT"
    QUANTITY_MISMATCH = "QUANTITY_MISMATCH"
    PRICE_MISMATCH = "PRICE_MISMATCH"
    MISSING_SHIPMENT = "MISSING_SHIPMENT"
    MISSING_INVOICE = "MISSING_INVOICE"
    AMBIGUOUS_REFERENCE = "AMBIGUOUS_REFERENCE"
    DUPLICATE_TRANSACTION = "DUPLICATE_TRANSACTION"
    DATA_VALIDATION_ERROR = "DATA_VALIDATION_ERROR"


class ExceptionCategory(StrEnum):
    """Actionable exception classification for operational and accounting review."""

    QUANTITY_MISMATCH = "QUANTITY_MISMATCH"
    PRICE_MISMATCH = "PRICE_MISMATCH"
    MISSING_SHIPMENT = "MISSING_SHIPMENT"
    MISSING_INVOICE = "MISSING_INVOICE"
    PARTIAL_FULFILLMENT = "PARTIAL_FULFILLMENT"
    DUPLICATE_TRANSACTION = "DUPLICATE_TRANSACTION"
    UNEXPECTED_EXTRA_RECORD = "UNEXPECTED_EXTRA_RECORD"
    AMBIGUOUS_REFERENCE = "AMBIGUOUS_REFERENCE"
    DATA_VALIDATION_ERROR = "DATA_VALIDATION_ERROR"


class Severity(StrEnum):
    """Operational urgency level for exceptions."""

    HIGH = "HIGH"
    MEDIUM = "MEDIUM"
    LOW = "LOW"


class ReasonCode(StrEnum):
    """Auditable machine-readable reason codes explaining match and reconciliation outcomes."""

    # Matching outcomes
    MATCHED_EXACT_ID = "MATCHED_EXACT_ID"
    MATCHED_PRODUCT_ALIAS = "MATCHED_PRODUCT_ALIAS"
    MATCHED_COMPOSITE_KEY = "MATCHED_COMPOSITE_KEY"
    AMBIGUOUS_MASTER_SKU = "AMBIGUOUS_MASTER_SKU"
    AMBIGUOUS_ALIAS_MAPPING = "AMBIGUOUS_ALIAS_MAPPING"
    CONFLICTING_MASTER_RECORD = "CONFLICTING_MASTER_RECORD"
    MULTIPLE_CANDIDATE_PO_LINES = "MULTIPLE_CANDIDATE_PO_LINES"
    LINE_PRODUCT_CONFLICT = "LINE_PRODUCT_CONFLICT"
    SUPPLIER_MISMATCH = "SUPPLIER_MISMATCH"
    MISSING_SKU = "MISSING_SKU"
    UNKNOWN_SKU = "UNKNOWN_SKU"
    MISSING_ORDER_REFERENCE = "MISSING_ORDER_REFERENCE"
    UNKNOWN_PO_NUMBER = "UNKNOWN_PO_NUMBER"
    PRODUCT_NOT_ON_PO = "PRODUCT_NOT_ON_PO"

    # Reconciliation & validation outcomes
    THREE_WAY_MATCH = "THREE_WAY_MATCH"
    ZERO_VALUE_CONFIRMED = "ZERO_VALUE_CONFIRMED"
    PARTIAL_SHIPMENT_BILLED = "PARTIAL_SHIPMENT_BILLED"
    PARTIAL_SHIPMENT_UNBILLED = "PARTIAL_SHIPMENT_UNBILLED"
    OVER_SHIPMENT = "OVER_SHIPMENT"
    UNDER_SHIPMENT = "UNDER_SHIPMENT"
    INVOICE_QTY_EXCEEDS_SHIPPED = "INVOICE_QTY_EXCEEDS_SHIPPED"
    INVOICE_QTY_MISMATCH = "INVOICE_QTY_MISMATCH"
    UNIT_PRICE_DISCREPANCY = "UNIT_PRICE_DISCREPANCY"
    INVOICE_MATH_DISCREPANCY = "INVOICE_MATH_DISCREPANCY"
    PO_MATH_DISCREPANCY = "PO_MATH_DISCREPANCY"
    NO_SHIPMENT_RECORDED = "NO_SHIPMENT_RECORDED"
    INVOICED_WITHOUT_SHIPMENT = "INVOICED_WITHOUT_SHIPMENT"
    NO_INVOICE_RECORDED = "NO_INVOICE_RECORDED"
    DUPLICATE_PRODUCT_ROW = "DUPLICATE_PRODUCT_ROW"
    DUPLICATE_PO_LINE = "DUPLICATE_PO_LINE"
    DUPLICATE_SHIPMENT = "DUPLICATE_SHIPMENT"
    DUPLICATE_INVOICE = "DUPLICATE_INVOICE"
    CONFLICTING_DUPLICATE_PO_LINE = "CONFLICTING_DUPLICATE_PO_LINE"
    CONFLICTING_DUPLICATE_SHIPMENT = "CONFLICTING_DUPLICATE_SHIPMENT"
    CONFLICTING_DUPLICATE_INVOICE = "CONFLICTING_DUPLICATE_INVOICE"
    AMBIGUOUS_TRANSACTION_LINK = "AMBIGUOUS_TRANSACTION_LINK"
    CURRENCY_MISMATCH = "CURRENCY_MISMATCH"
    INVALID_NUMERIC_VALUE = "INVALID_NUMERIC_VALUE"
    NEGATIVE_QUANTITY_OR_PRICE = "NEGATIVE_QUANTITY_OR_PRICE"
    INVALID_DATE_VALUE = "INVALID_DATE_VALUE"
    BROKEN_EXCEL_FORMULA = "BROKEN_EXCEL_FORMULA"
    UNPERMITTED_EXCEL_FORMULA = "UNPERMITTED_EXCEL_FORMULA"


@dataclass(frozen=True, slots=True, order=True)
class PoLineKey:
    """Collision-free structured composite key identifying a unique Purchase Order line."""

    po_number: str
    po_line_id: str

    def format_display(self) -> str:
        """Return an injective display representation without delimiter collisions."""
        return f"{_escape_key_component(self.po_number)}|{_escape_key_component(self.po_line_id)}"


@dataclass(frozen=True, slots=True, order=True)
class ShipmentLineKey:
    """Explicit semantic identity for a shipment line after alias and order-line resolution."""

    shipment_id: str
    order_scope_kind: str
    po_number: str
    po_line_id: str
    product_key: str


@dataclass(frozen=True, slots=True, order=True)
class InvoiceLineKey:
    """Supplier-scoped identity for a supplier invoice line."""

    supplier_id: str
    invoice_id: str
    invoice_line_id: str


REQUIRED_COLUMNS: dict[DatasetKind, tuple[str, ...]] = {
    DatasetKind.PRODUCTS: (
        "product_id",
        "canonical_sku",
        "product_name",
        "category",
        "uom",
        "standard_cost",
        "supplier_id",
        "aliases",
        "active_flag",
    ),
    DatasetKind.PURCHASE_ORDERS: (
        "po_number",
        "po_line_id",
        "order_date",
        "supplier_id",
        "sku",
        "product_description",
        "ordered_qty",
        "unit_price",
        "currency",
        "line_total",
    ),
    DatasetKind.SHIPMENTS: (
        "shipment_id",
        "po_number",
        "po_line_id",
        "ship_date",
        "carrier",
        "tracking_ref",
        "sku",
        "shipped_qty",
        "warehouse_code",
    ),
    DatasetKind.INVOICES: (
        "invoice_id",
        "invoice_line_id",
        "po_number",
        "po_line_id",
        "invoice_date",
        "supplier_id",
        "sku",
        "invoiced_qty",
        "unit_price",
        "line_amount",
        "currency",
    ),
}

EXCEL_ERROR_TOKENS: frozenset[str] = frozenset(
    {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#ERROR!"}
)


@dataclass(frozen=True, slots=True)
class SchemaValidationResult:
    """Outcome of validating workbook headers against expected dataset columns."""

    dataset_kind: DatasetKind
    is_valid: bool
    expected_columns: tuple[str, ...]
    observed_columns: tuple[str, ...]
    missing_columns: tuple[str, ...]
    extra_columns: tuple[str, ...]


class SchemaValidationError(ValueError):
    """Raised when one or more input workbooks are missing required schema columns."""

    def __init__(self, invalid_results: tuple[SchemaValidationResult, ...]) -> None:
        self.validation_results = invalid_results
        details = "; ".join(
            f"{res.dataset_kind.value} missing [{', '.join(res.missing_columns)}]"
            for res in invalid_results
        )
        super().__init__(f"Required column validation failed: {details}")


@dataclass(frozen=True, slots=True)
class ReconciliationConfig:
    """Explicit rounding, currency, and tolerance rules for three-way financial reconciliation.

    Rounding rule:
        All monetary amounts are quantized to 0.01 using standard commercial
        half-up rounding (decimal.ROUND_HALF_UP) in a 50-digit Decimal context.

    Tolerances:
        quantity_tolerance: Maximum permitted absolute quantity difference (default 0).
        unit_price_tolerance: Maximum permitted absolute unit-price difference (default $0.01).
        amount_tolerance: Maximum permitted extended line-amount difference across
            aggregated invoices (default $0.02 to accommodate split-invoice penny rounding).
        reporting_currency: Single reporting currency enforced across POs and
            Invoices (default USD).
    """

    quantity_tolerance: Decimal = Decimal("0")
    unit_price_tolerance: Decimal = Decimal("0.01")
    amount_tolerance: Decimal = Decimal("0.02")
    currency_quantum: Decimal = Decimal("0.01")
    rounding_mode: str = ROUND_HALF_UP
    reporting_currency: str = "USD"

    def quantize_money(self, value: Decimal) -> Decimal:
        """Round a Decimal monetary value using configured commercial rounding."""
        with localcontext(ARITHMETIC_CONTEXT):
            return value.quantize(self.currency_quantum, rounding=self.rounding_mode)


@dataclass(frozen=True, slots=True)
class InputFileMetadata:
    """Audit metadata captured when reading an input Excel workbook."""

    dataset_kind: DatasetKind
    filename: str
    filepath: str
    sha256: str
    sheet_name: str
    row_count: int
    formula_cell_count: int
    broken_formula_count: int


@dataclass(frozen=True, slots=True)
class CleaningSummary:
    """Row-level data hygiene metrics emitted during dataset cleaning."""

    dataset_kind: DatasetKind
    total_rows: int
    valid_unique_rows: int
    duplicate_rows: int
    whitespace_normalized_count: int
    sku_normalized_count: int
    missing_value_rows: int
    invalid_numeric_rows: int
    invalid_date_rows: int
    broken_formula_rows: int
    conflicting_duplicate_rows: int = 0


@dataclass(frozen=True, slots=True)
class ReconciledRecord:
    """Auditable three-way reconciliation record for a purchase order line."""

    source_row: int
    raw_po_number: str
    raw_po_line_id: str
    raw_sku: str
    normalized_po_number: str
    normalized_po_line_id: str
    normalized_sku: str
    matched_product_id: str | None
    canonical_sku: str | None
    matched_business_key: str
    product_description: str
    supplier_id: str
    order_date: str | None
    product_match_status: MatchStatus
    ordered_qty: Decimal | None
    shipped_qty: Decimal | None
    invoiced_qty: Decimal | None
    outstanding_qty: Decimal | None
    po_unit_price: Decimal | None
    expected_amount: Decimal | None
    actual_invoiced_amount: Decimal | None
    monetary_variance: Decimal | None
    shipment_count: int
    invoice_count: int
    shipment_refs: str
    invoice_refs: str
    preserved_po_formula: str | None
    reconciliation_status: ReconciliationStatus
    reason_codes: tuple[ReasonCode, ...]
    is_valid_po_line: bool = False
    discrepancy_exposure: Decimal = Decimal("0.00")
    unbilled_receiving_amount: Decimal = Decimal("0.00")


@dataclass(frozen=True, slots=True)
class ExceptionRecord:
    """Structured exception requiring operational, procurement, or AP review."""

    exception_id: str
    category: ExceptionCategory
    severity: Severity
    source_dataset: DatasetKind
    source_row: int
    raw_reference: str
    normalized_reference: str
    raw_sku: str
    normalized_sku: str
    po_reference: str
    reason_codes: tuple[ReasonCode, ...]
    explanation: str
    monetary_impact: Decimal | None
    quantity_impact: Decimal | None
    suggested_action: str
    original_monetary_amount: Decimal | None = None
    original_currency: str | None = None


@dataclass(frozen=True, slots=True)
class SummaryMetrics:
    """Aggregate KPIs for executive reconciliation reporting."""

    total_po_rows: int
    unique_valid_po_lines: int
    duplicate_po_rows: int
    invalid_po_rows: int
    reconciled_lines: int
    partial_fulfillment_lines: int
    po_exception_lines: int
    total_exception_records: int
    matching_coverage_rate: Decimal
    reconciliation_rate: Decimal
    total_ordered_qty: Decimal
    total_shipped_qty: Decimal
    total_invoiced_qty: Decimal
    total_outstanding_qty: Decimal
    total_expected_amount: Decimal
    total_invoiced_amount: Decimal
    net_billed_variance: Decimal
    absolute_monetary_discrepancy: Decimal
    unbilled_shipped_amount: Decimal
    quarantined_invoice_amount: Decimal = Decimal("0.00")
    exception_counts_by_category: dict[str, int] = field(default_factory=dict)
    match_counts_by_status: dict[str, int] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class AuditLogBundle:
    """Full provenance log embedded into the Audit Log worksheet."""

    generated_at_utc: str
    python_version: str
    polars_version: str
    openpyxl_version: str
    config: ReconciliationConfig
    input_files: tuple[InputFileMetadata, ...]
    schema_validations: tuple[SchemaValidationResult, ...]
    cleaning_summaries: tuple[CleaningSummary, ...]
    matching_rules_applied: tuple[str, ...]
    warnings: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class ValidationReport:
    """Result of validating a generated reconciliation_report.xlsx file."""

    report_path: str
    is_valid: bool
    worksheets_found: tuple[str, ...]
    formula_count: int
    broken_formula_cells: tuple[str, ...]
    errors: tuple[str, ...]
    warnings: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class ReconciliationResult:
    """Complete output of an end-to-end reconciliation run."""

    summary: SummaryMetrics
    reconciled_records: tuple[ReconciledRecord, ...]
    exceptions: tuple[ExceptionRecord, ...]
    audit_log: AuditLogBundle
