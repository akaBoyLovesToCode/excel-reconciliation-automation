"""Deterministic synthetic Excel dataset and ground-truth generator."""

from __future__ import annotations

import datetime as dt
import json
import os
import random
import tempfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from excel_reconciliation.models import (
    REQUIRED_COLUMNS,
    DatasetKind,
    ReconciliationConfig,
    ReconciliationStatus,
    excel_round,
)

DEFAULT_DEMO_SEED = 20260401

_CATEGORIES: tuple[tuple[str, str], ...] = (
    ("ELEC", "Electronics"),
    ("HOME", "Home & Kitchen"),
    ("IND", "Industrial"),
    ("OFFC", "Office Supplies"),
    ("FOOD", "Specialty Food"),
)

_UOMS: tuple[str, ...] = ("EA", "BOX", "CASE", "PACK", "SET")
_CARRIERS: tuple[str, ...] = (
    "FedEx Freight",
    "UPS Supply Chain",
    "DHL Global Forwarding",
    "Maersk Overland",
    "XPO Logistics",
)
_WAREHOUSES: tuple[str, ...] = ("WH-US-EAST", "WH-US-WEST", "WH-EU-DE", "WH-APAC-SG")

_STANDARD_DESCRIPTIONS: dict[str, tuple[str, ...]] = {
    "ELEC": (
        "USB-C Docking Station Dual 4K",
        "Industrial Barcode Scanner 2D",
        "Managed Gigabit PoE+ Switch 24-Port",
        "Wireless Thermal Label Printer 300dpi",
        "Uninterruptible Power Supply 1500VA",
    ),
    "HOME": (
        "Insulated Stainless Carafe 1.5L",
        "Commercial Induction Cooktop 3500W",
        "HEPA Air Filtration Cartridge Pro",
        "Digital Precision Kitchen Scale 10kg",
        "Heavy-Duty Silicone Baking Mat Set",
    ),
    "IND": (
        "Pneumatic Solenoid Valve 24V DC",
        "Digital Torque Wrench Calibrated 200Nm",
        "Braided Hydraulic Hose Assembly 2m",
        "Laser Distance Meter IP65 100m",
        "Modular DIN-Rail Power Supply 480W",
    ),
    "OFFC": (
        "Ergonomic Mesh Task Chair Adjustable",
        "Acoustic Desk Privacy Panel 120cm",
        "Heavy-Duty Cross-Cut Paper Shredder",
        "Dual-Monitor Articulating Gas Arm",
        "Archival Acid-Free Storage Ledger Box",
    ),
    "FOOD": (
        "Single-Origin Whole Bean Espresso 1kg",
        "Organic Cold-Pressed Extra Virgin Olive Oil",
        "Artisanal Sea Salt Flakes Bulk 2.5kg",
        "Raw Wildflower Honey Glass Jar 500g",
        "Sparkling Mineral Water Case 24x330ml",
    ),
}

_UNICODE_DESCRIPTIONS: tuple[str, ...] = (
    "Précision Café Grinder — Edelstahl 220V",
    "日本の宇治抹茶パウダー Ceremonial Grade 500g",
    "Válvula Hidráulica Proporcional Ø25mm",
    "Bürostuhl Ergonomisch Größe L — Schwarz",
    "Caméra Thermique Infrarouge — Série Élite",
    "Cafetera Automática Presión 19 Bar — Edición Pro",
    "Messschieber Präzision ±0.01mm Rostfrei",
    "Thé Vert Sencha Supérieur — Récolte Printemps",
)


@dataclass(frozen=True, slots=True)
class SyntheticBundle:
    """In-memory representation of generated rows and ground-truth expectations."""

    seed: int
    products: list[dict[str, Any]]
    purchase_orders: list[dict[str, Any]]
    shipments: list[dict[str, Any]]
    invoices: list[dict[str, Any]]
    ground_truth: dict[str, Any]


def _dirty_sku_variant(sku: str, rng: random.Random) -> str:
    """Apply realistic formatting noise (separators, casing, outer spaces) to a canonical SKU."""
    mode = rng.randrange(5)
    if mode == 0:
        return f"  {sku.lower()}  "
    if mode == 1:
        return sku.replace("-", "_")
    if mode == 2:
        return sku.replace("-", ".").lower()
    if mode == 3:
        return f" {sku.replace('-', ' ')}\t"
    return f"  {sku} "


def _dirty_id_variant(identifier: str, rng: random.Random) -> str:
    """Apply realistic whitespace or casing noise to a business reference ID."""
    mode = rng.randrange(4)
    if mode == 0:
        return f"  {identifier} "
    if mode == 1:
        return identifier.lower()
    if mode == 2:
        return f"\t{identifier.lower()}  "
    return identifier


_EN_MONTH_ABBRS: tuple[str, ...] = (
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
)


def _format_date_variant(date_obj: dt.date, rng: random.Random) -> str:
    """Format a calendar date in mixed enterprise export styles (locale-independent)."""
    mode = rng.randrange(10)
    if mode < 7:
        return date_obj.isoformat()
    if mode == 7:
        return date_obj.strftime("%m/%d/%Y")
    if mode == 8:
        return date_obj.strftime("%Y/%m/%d")
    mon_abbr = _EN_MONTH_ABBRS[date_obj.month - 1]
    return f"{date_obj.day:02d}-{mon_abbr}-{date_obj.year:04d}"


def build_synthetic_records(seed: int = DEFAULT_DEMO_SEED) -> SyntheticBundle:
    """Deterministically construct synthetic datasets and exact ground-truth expectations."""
    rng = random.Random(seed)
    cfg = ReconciliationConfig()
    base_date = dt.date(2026, 1, 5)

    # -------------------------------------------------------------------------
    # 1. Product Master: 1,500 rows (1,485 valid unique + 10 duplicates + 5 invalid)
    # -------------------------------------------------------------------------
    products: list[dict[str, Any]] = []
    catalog: list[dict[str, Any]] = []

    for idx in range(1, 1486):
        cat_code, cat_name = _CATEGORIES[(idx - 1) % len(_CATEGORIES)]
        prod_id = f"PRD-{idx:04d}"
        canon_sku = f"SKU-{cat_code}-{idx:04d}"
        sup_num = ((idx - 1) % 25) + 1
        supplier_id = f"SUP-{sup_num:02d}"
        uom = _UOMS[(idx - 1) % len(_UOMS)]
        cost = Decimal(str(rng.randint(850, 48000))) / Decimal("100")

        if idx % 12 == 0:
            prod_name = _UNICODE_DESCRIPTIONS[(idx // 12) % len(_UNICODE_DESCRIPTIONS)]
        else:
            templates = _STANDARD_DESCRIPTIONS[cat_code]
            prod_name = f"{templates[(idx - 1) % len(templates)]} #{idx:04d}"

        leg_alias = f"LEG-{cat_code}-{idx:04d}"
        sup_alias = f"{supplier_id.replace('-', '')}-P{idx:04d}"

        # Products 1474..1485 (6 pairs) share colliding aliases AMB-ALIAS-01..06
        if idx >= 1474:
            pair_num = ((idx - 1474) // 2) + 1
            amb_alias = f"AMB-ALIAS-{pair_num:02d}"
            alias_str = f"{leg_alias}; {amb_alias}"
            primary_alias = leg_alias
        else:
            alias_str = f"{leg_alias}; {sup_alias}"
            primary_alias = leg_alias if idx % 2 == 0 else sup_alias

        raw_sku_cell = _dirty_sku_variant(canon_sku, rng) if idx % 7 == 0 else canon_sku
        raw_name_cell = f"  {prod_name}  " if idx % 11 == 0 else prod_name

        item_meta = {
            "product_id": prod_id,
            "canonical_sku": canon_sku,
            "product_name": prod_name,
            "supplier_id": supplier_id,
            "cost": cost,
            "primary_alias": primary_alias,
        }
        if idx <= 1473:
            catalog.append(item_meta)

        products.append(
            {
                "product_id": f" {prod_id} " if idx % 13 == 0 else prod_id,
                "canonical_sku": raw_sku_cell,
                "product_name": raw_name_cell,
                "category": cat_name,
                "uom": uom,
                "standard_cost": f"{cost:.2f}",
                "supplier_id": supplier_id,
                "aliases": alias_str,
                "active_flag": "Y",
            }
        )

    # 10 duplicate product rows (copies of PRD-0010 .. PRD-0019 with whitespace/case noise)
    for dup_idx in range(9, 19):
        orig = products[dup_idx]
        products.append(
            {
                "product_id": f"  {orig['product_id'].strip().lower()} ",
                "canonical_sku": _dirty_sku_variant(orig["canonical_sku"].strip(), rng),
                "product_name": orig["product_name"],
                "category": orig["category"],
                "uom": orig["uom"],
                "standard_cost": orig["standard_cost"],
                "supplier_id": orig["supplier_id"],
                "aliases": orig["aliases"],
                "active_flag": "Y",
            }
        )

    # 5 invalid-cost product rows (PRD-1486 .. PRD-1490)
    bad_costs = ("TBD", "N/A", "-15.00", "ERR", "pending")
    for offset, bad_cost in enumerate(bad_costs, start=1486):
        cat_code, cat_name = _CATEGORIES[offset % len(_CATEGORIES)]
        products.append(
            {
                "product_id": f"PRD-{offset:04d}",
                "canonical_sku": f"SKU-{cat_code}-{offset:04d}",
                "product_name": f"Unpriced Prototype Item #{offset}",
                "category": cat_name,
                "uom": "EA",
                "standard_cost": bad_cost,
                "supplier_id": "SUP-01",
                "aliases": "",
                "active_flag": "Y",
            }
        )

    assert len(products) == 1500

    # -------------------------------------------------------------------------
    # 2. Purchase Orders (2,000 rows), Shipments (2,200 rows), Invoices (2,100 rows)
    # -------------------------------------------------------------------------
    purchase_orders: list[dict[str, Any]] = []
    shipments: list[dict[str, Any]] = []
    invoices: list[dict[str, Any]] = []
    expected_po_status: dict[str, str] = {}

    po_counter = 10000
    ship_counter = 50000
    inv_counter = 80000

    def _next_po_num() -> str:
        nonlocal po_counter
        po_counter += 1
        return f"PO-2026-{po_counter:05d}"

    def _next_ship_id() -> str:
        nonlocal ship_counter
        ship_counter += 1
        return f"SHP-2026-{ship_counter:05d}"

    def _next_inv_ids() -> tuple[str, str]:
        nonlocal inv_counter
        inv_counter += 1
        inv_id = f"INV-2026-{inv_counter:05d}"
        return inv_id, f"{inv_id}-01"

    def _pick_po_sku(item: dict[str, Any], idx_hint: int) -> str:
        if idx_hint % 6 == 0:
            return item["primary_alias"]
        if idx_hint % 4 == 0:
            return _dirty_sku_variant(item["canonical_sku"], rng)
        return item["canonical_sku"]

    def _pick_tx_sku(item: dict[str, Any], idx_hint: int) -> str:
        if idx_hint % 5 == 0:
            return _dirty_sku_variant(item["primary_alias"], rng)
        if idx_hint % 3 == 0:
            return _dirty_sku_variant(item["canonical_sku"], rng)
        return item["canonical_sku"]

    def _add_po_row(
        po_num: str,
        line_id: str,
        order_dt: str,
        supplier_id: str,
        sku_str: str,
        desc: str,
        qty_cell: int | str,
        price_cell: str,
        *,
        line_total_override: str | None = None,
        noisy: bool = False,
    ) -> None:
        purchase_orders.append(
            {
                "po_number": _dirty_id_variant(po_num, rng) if noisy else po_num,
                "po_line_id": _dirty_id_variant(line_id, rng) if noisy else line_id,
                "order_date": order_dt,
                "supplier_id": supplier_id,
                "sku": sku_str,
                "product_description": f"  {desc} " if noisy else desc,
                "ordered_qty": qty_cell,
                "unit_price": price_cell,
                "currency": "USD",
                "_line_total_override": line_total_override,
            }
        )

    def _add_shipment(
        po_num: str | None,
        line_id: str | None,
        ship_dt: str,
        sku_str: str,
        qty_cell: int | str,
        *,
        ship_id_override: str | None = None,
        carrier_override: str | None = None,
        tracking_override: str | None = None,
        warehouse_override: str | None = None,
        noisy: bool = False,
    ) -> str:
        sid = ship_id_override or _next_ship_id()
        default_tracking = f"TRK-{rng.randint(10000000, 99999999)}"
        shipments.append(
            {
                "shipment_id": _dirty_id_variant(sid, rng) if noisy else sid,
                "po_number": (
                    (_dirty_id_variant(po_num, rng) if noisy else po_num)
                    if po_num is not None
                    else ""
                ),
                "po_line_id": (
                    (_dirty_id_variant(line_id, rng) if noisy else line_id)
                    if line_id is not None
                    else ""
                ),
                "ship_date": ship_dt,
                "carrier": carrier_override or _CARRIERS[len(shipments) % len(_CARRIERS)],
                "tracking_ref": tracking_override or default_tracking,
                "sku": sku_str,
                "shipped_qty": qty_cell,
                "warehouse_code": warehouse_override
                or _WAREHOUSES[len(shipments) % len(_WAREHOUSES)],
            }
        )
        return sid

    def _add_invoice(
        po_num: str | None,
        line_id: str | None,
        inv_dt: str,
        supplier_id: str,
        sku_str: str,
        qty_cell: int | str,
        price_cell: str,
        *,
        inv_id_override: tuple[str, str] | None = None,
        amount_override: str | None = None,
        noisy: bool = False,
    ) -> tuple[str, str]:
        inv_id, inv_line_id = inv_id_override or _next_inv_ids()
        invoices.append(
            {
                "invoice_id": _dirty_id_variant(inv_id, rng) if noisy else inv_id,
                "invoice_line_id": _dirty_id_variant(inv_line_id, rng) if noisy else inv_line_id,
                "po_number": (
                    (_dirty_id_variant(po_num, rng) if noisy else po_num)
                    if po_num is not None
                    else ""
                ),
                "po_line_id": (
                    (_dirty_id_variant(line_id, rng) if noisy else line_id)
                    if line_id is not None
                    else ""
                ),
                "invoice_date": inv_dt,
                "supplier_id": supplier_id,
                "sku": sku_str,
                "invoiced_qty": qty_cell,
                "unit_price": price_cell,
                "_amount_override": amount_override,
                "currency": "USD",
            }
        )
        return inv_id, inv_line_id

    # Single-line POs guarantee unique (po_number, product_id) when po_line_id is omitted
    # Cohort 1: 1,170 clean reconciled lines
    #   - First 145 lines get 2 split shipments (+145 shipments)
    #   - Next 285 lines get 2 split invoices (+285 invoices)
    #   - Remaining 740 lines have 1 shipment + 1 invoice
    for i in range(1170):
        item = catalog[i % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=i % 75)
        sdt = odt + dt.timedelta(days=4)
        idt = sdt + dt.timedelta(days=3)

        qty = rng.randint(4, 60) * 2  # even number for clean 50/50 splits
        unit_price = cfg.quantize_money(item["cost"] * Decimal("1.15"))
        noisy = i % 3 == 0
        omit_line_id_on_tx = i % 4 == 0  # exercises Tier 3 COMPOSITE_KEY matching

        _add_po_row(
            po_num,
            line_id,
            _format_date_variant(odt, rng),
            item["supplier_id"],
            _pick_po_sku(item, i),
            item["product_name"],
            qty,
            f"{unit_price:.2f}",
            noisy=noisy,
        )
        expected_po_status[line_id] = ReconciliationStatus.RECONCILED.value

        tx_line_ref = None if omit_line_id_on_tx else line_id
        if i < 145:
            q1 = qty // 2
            q2 = qty - q1
            _add_shipment(
                po_num, tx_line_ref, sdt.isoformat(), _pick_tx_sku(item, i), q1, noisy=noisy
            )
            _add_shipment(
                po_num,
                tx_line_ref,
                (sdt + dt.timedelta(days=2)).isoformat(),
                _pick_tx_sku(item, i + 1),
                q2,
            )
        else:
            _add_shipment(
                po_num, tx_line_ref, sdt.isoformat(), _pick_tx_sku(item, i), qty, noisy=noisy
            )

        if 145 <= i < 430:
            iq1 = qty // 2
            iq2 = qty - iq1
            _add_invoice(
                po_num,
                tx_line_ref,
                idt.isoformat(),
                item["supplier_id"],
                _pick_tx_sku(item, i),
                iq1,
                f"{unit_price:.2f}",
                noisy=noisy,
            )
            _add_invoice(
                po_num,
                tx_line_ref,
                (idt + dt.timedelta(days=2)).isoformat(),
                item["supplier_id"],
                _pick_tx_sku(item, i + 2),
                iq2,
                f"{unit_price:.2f}",
            )
        else:
            _add_invoice(
                po_num,
                tx_line_ref,
                idt.isoformat(),
                item["supplier_id"],
                _pick_tx_sku(item, i),
                qty,
                f"{unit_price:.2f}",
                noisy=noisy,
            )

    # Cohort 2: 150 multi-shipment reconciled lines (135 with 2 shipments, 15 with 3 shipments)
    for i in range(150):
        item = catalog[(i + 200) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=(i + 10) % 70)
        qty = rng.randint(6, 30) * 3
        unit_price = cfg.quantize_money(item["cost"] * Decimal("1.10"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            _pick_po_sku(item, i),
            item["product_name"],
            qty,
            f"{unit_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.RECONCILED.value

        if i < 15:
            part = qty // 3
            for leg in range(3):
                leg_q = part if leg < 2 else (qty - 2 * part)
                _add_shipment(
                    po_num,
                    line_id,
                    (odt + dt.timedelta(days=3 + leg * 2)).isoformat(),
                    _pick_tx_sku(item, i + leg),
                    leg_q,
                )
        else:
            q1 = qty // 2
            _add_shipment(
                po_num, line_id, (odt + dt.timedelta(days=3)).isoformat(), item["canonical_sku"], q1
            )
            _add_shipment(
                po_num,
                line_id,
                (odt + dt.timedelta(days=6)).isoformat(),
                item["primary_alias"],
                qty - q1,
            )

        _add_invoice(
            po_num,
            line_id,
            (odt + dt.timedelta(days=8)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            qty,
            f"{unit_price:.2f}",
        )

    # Cohort 3: 60 M x N lines (2 shipments AND 2 invoices -> join-inflation safety)
    for i in range(60):
        item = catalog[(i + 500) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=i % 60)
        qty = rng.randint(10, 40) * 2
        unit_price = cfg.quantize_money(item["cost"] * Decimal("1.20"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            qty,
            f"{unit_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.RECONCILED.value

        half = qty // 2
        _add_shipment(
            po_num, line_id, (odt + dt.timedelta(days=2)).isoformat(), item["canonical_sku"], half
        )
        _add_shipment(
            po_num, line_id, (odt + dt.timedelta(days=5)).isoformat(), item["canonical_sku"], half
        )
        _add_invoice(
            po_num,
            line_id,
            (odt + dt.timedelta(days=4)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            half,
            f"{unit_price:.2f}",
        )
        _add_invoice(
            po_num,
            line_id,
            (odt + dt.timedelta(days=7)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            half,
            f"{unit_price:.2f}",
        )

    # Cohort 4: 5 valid zero-quantity confirmed lines (distinguishes valid 0 from missing None)
    for i in range(5):
        item = catalog[(i + 700) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=15 + i)
        unit_price = cfg.quantize_money(item["cost"])

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            0,
            f"{unit_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.RECONCILED.value
        _add_shipment(
            po_num, line_id, (odt + dt.timedelta(days=1)).isoformat(), item["canonical_sku"], 0
        )
        _add_invoice(
            po_num,
            line_id,
            (odt + dt.timedelta(days=2)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            0,
            f"{unit_price:.2f}",
        )

    # Cohort 5: 95 clean partial fulfillments (shipped < ordered, invoiced == shipped at PO price)
    for i in range(95):
        item = catalog[(i + 750) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=i % 50)
        ordered_qty = rng.randint(20, 100)
        shipped_qty = rng.randint(5, ordered_qty - 5)
        unit_price = cfg.quantize_money(item["cost"] * Decimal("1.12"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            _pick_po_sku(item, i),
            item["product_name"],
            ordered_qty,
            f"{unit_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.PARTIAL_FULFILLMENT.value
        _add_shipment(
            po_num,
            line_id,
            (odt + dt.timedelta(days=4)).isoformat(),
            item["canonical_sku"],
            shipped_qty,
        )
        _add_invoice(
            po_num,
            line_id,
            (odt + dt.timedelta(days=6)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            shipped_qty,
            f"{unit_price:.2f}",
        )

    # Cohort 6: 90 quantity mismatches (over-shipments or over-billed partial shipments)
    for i in range(90):
        item = catalog[(i + 900) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=i % 60)
        ordered_qty = rng.randint(15, 80)
        unit_price = cfg.quantize_money(item["cost"] * Decimal("1.08"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            ordered_qty,
            f"{unit_price:.2f}",
            noisy=(i % 2 == 0),
        )
        expected_po_status[line_id] = ReconciliationStatus.QUANTITY_MISMATCH.value

        if i % 3 == 0:
            # Over-shipment billed at over-shipped qty
            ship_q = ordered_qty + rng.randint(2, 12)
            inv_q = ship_q
        elif i % 3 == 1:
            # Partial shipment, but supplier billed full ordered_qty!
            ship_q = ordered_qty - rng.randint(3, 10)
            inv_q = ordered_qty
        else:
            # Full shipment received, supplier over-billed extra units
            ship_q = ordered_qty
            inv_q = ordered_qty + rng.randint(1, 8)

        _add_shipment(
            po_num, line_id, (odt + dt.timedelta(days=4)).isoformat(), item["canonical_sku"], ship_q
        )
        _add_invoice(
            po_num,
            line_id,
            (odt + dt.timedelta(days=6)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            inv_q,
            f"{unit_price:.2f}",
        )

    # Cohort 7: 85 price / extended-amount mismatches
    for i in range(85):
        item = catalog[(i + 1050) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=i % 55)
        qty = rng.randint(5, 50)
        po_price = cfg.quantize_money(item["cost"] * Decimal("1.10"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            _pick_po_sku(item, i),
            item["product_name"],
            qty,
            f"{po_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.PRICE_MISMATCH.value
        _add_shipment(
            po_num, line_id, (odt + dt.timedelta(days=3)).isoformat(), item["canonical_sku"], qty
        )

        if i < 12:
            # Unit price matches PO, but hardcoded invoice line_amount has a surcharge/math error
            bogus_amt = cfg.quantize_money(Decimal(qty) * po_price + Decimal("47.50"))
            _add_invoice(
                po_num,
                line_id,
                (odt + dt.timedelta(days=5)).isoformat(),
                item["supplier_id"],
                item["canonical_sku"],
                qty,
                f"{po_price:.2f}",
                amount_override=f"{bogus_amt:.2f}",
            )
        else:
            delta_pct = Decimal("1.07") if i % 2 == 0 else Decimal("0.94")
            billed_price = cfg.quantize_money(po_price * delta_pct)
            if abs(billed_price - po_price) <= Decimal("0.05"):
                billed_price = po_price + Decimal("1.25")
            _add_invoice(
                po_num,
                line_id,
                (odt + dt.timedelta(days=5)).isoformat(),
                item["supplier_id"],
                _pick_tx_sku(item, i),
                qty,
                f"{billed_price:.2f}",
            )

    # Cohort 8: 95 missing invoices (goods shipped, supplier invoice missing)
    for i in range(95):
        item = catalog[(i + 1150) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=i % 60)
        qty = rng.randint(8, 45)
        po_price = cfg.quantize_money(item["cost"] * Decimal("1.15"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            qty,
            f"{po_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.MISSING_INVOICE.value
        ship_q = qty if i % 4 != 0 else max(qty - 3, 1)
        _add_shipment(
            po_num, line_id, (odt + dt.timedelta(days=4)).isoformat(), item["canonical_sku"], ship_q
        )

    # Cohort 9: 135 open PO lines with missing shipment (and no invoice)
    for i in range(135):
        item = catalog[(i + 1250) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=60 + (i % 25))
        qty = rng.randint(10, 75)
        po_price = cfg.quantize_money(item["cost"] * Decimal("1.10"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            _pick_po_sku(item, i),
            item["product_name"],
            qty,
            f"{po_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.MISSING_SHIPMENT.value

    # Cohort 10: 25 prematurely invoiced PO lines (invoice present, shipment missing!)
    for i in range(25):
        item = catalog[(i + 300) % len(catalog)]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=40 + i)
        qty = rng.randint(5, 30)
        po_price = cfg.quantize_money(item["cost"] * Decimal("1.12"))

        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            qty,
            f"{po_price:.2f}",
        )
        expected_po_status[line_id] = ReconciliationStatus.MISSING_SHIPMENT.value
        _add_invoice(
            po_num,
            line_id,
            (odt + dt.timedelta(days=2)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            qty,
            f"{po_price:.2f}",
        )

    # Cohort 11: 15 multi-line POs with identical SKU on L01 & L02 (30 PO lines)
    # Shipments and invoices omit po_line_id -> must quarantine as AMBIGUOUS_REFERENCE
    for i in range(15):
        item = catalog[(i + 400) % len(catalog)]
        po_num = _next_po_num()
        line1 = f"{po_num}-L01"
        line2 = f"{po_num}-L02"
        odt = base_date + dt.timedelta(days=20 + i)
        po_price = cfg.quantize_money(item["cost"] * Decimal("1.10"))

        _add_po_row(
            po_num,
            line1,
            odt.isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            20,
            f"{po_price:.2f}",
        )
        _add_po_row(
            po_num,
            line2,
            odt.isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            35,
            f"{po_price:.2f}",
        )
        expected_po_status[line1] = ReconciliationStatus.AMBIGUOUS_REFERENCE.value
        expected_po_status[line2] = ReconciliationStatus.AMBIGUOUS_REFERENCE.value

        # Shipment & invoice omit po_line_id -> 2 candidate PO lines -> AMBIGUOUS!
        _add_shipment(
            po_num, None, (odt + dt.timedelta(days=4)).isoformat(), item["canonical_sku"], 20
        )
        _add_invoice(
            po_num,
            None,
            (odt + dt.timedelta(days=6)).isoformat(),
            item["supplier_id"],
            item["canonical_sku"],
            20,
            f"{po_price:.2f}",
        )

    # Cohort 12: 10 PO lines referencing colliding ambiguous aliases AMB-ALIAS-01..06
    for i in range(10):
        amb_sku = f"AMB-ALIAS-{(i % 6) + 1:02d}"
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=30 + i)
        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            "SUP-02",
            amb_sku,
            f"Legacy Catalog Component ({amb_sku})",
            12,
            "45.00",
        )
        expected_po_status[line_id] = ReconciliationStatus.AMBIGUOUS_REFERENCE.value

    # Cohort 13: 10 PO lines referencing unknown uncataloged SKUs
    for i in range(10):
        unk_sku = f"SKU-UNLISTED-9{i:02d}"
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        odt = base_date + dt.timedelta(days=35 + i)
        _add_po_row(
            po_num,
            line_id,
            odt.isoformat(),
            "SUP-05",
            unk_sku,
            f"Special Non-Catalog Item {unk_sku}",
            10,
            "99.00",
        )
        expected_po_status[line_id] = ReconciliationStatus.DATA_VALIDATION_ERROR.value

    assert len(purchase_orders) == 1960

    # Add 25 duplicate PO rows (copies of first 25 PO lines with whitespace/casing noise)
    for dup_i in range(25):
        orig_po = purchase_orders[dup_i]
        purchase_orders.append(
            {
                "po_number": f"  {orig_po['po_number'].strip().lower()}  ",
                "po_line_id": f" {orig_po['po_line_id'].strip().lower()} ",
                "order_date": orig_po["order_date"],
                "supplier_id": orig_po["supplier_id"],
                "sku": orig_po["sku"],
                "product_description": orig_po["product_description"],
                "ordered_qty": orig_po["ordered_qty"],
                "unit_price": orig_po["unit_price"],
                "currency": "USD",
                "_line_total_override": None,
            }
        )

    # Add 15 invalid PO rows (bad numbers, negative qty, impossible dates, broken formulas)
    invalid_po_specs: tuple[tuple[Any, str, str, str | None], ...] = (
        ("TBD", "25.00", "2026-02-10", None),
        ("N/A", "40.00", "2026-02-11", None),
        ("12O", "19.50", "2026-02-12", None),
        ("-8", "32.00", "2026-02-13", None),
        ("-15", "14.00", "2026-02-14", None),
        (10, "TBD", "2026-02-15", None),
        (20, "FREE?", "2026-02-16", None),
        (15, "-9.99", "2026-02-17", None),
        (12, "22.00", "2026-02-30", None),
        (18, "31.00", "PENDING", None),
        (25, "18.00", "2026-13-01", None),
        (10, "55.00", "not-a-date", None),
        (14, "27.50", "2026-02-18", "=#REF!*H1988"),
        (16, "44.00", "2026-02-19", "=#DIV/0!"),
        ("ERR", "12.00", "2026-02-20", None),
    )
    for idx_inv, (bad_q, bad_p, bad_dt, bad_formula) in enumerate(invalid_po_specs):
        item = catalog[idx_inv]
        po_num = _next_po_num()
        line_id = f"{po_num}-L01"
        _add_po_row(
            po_num,
            line_id,
            bad_dt,
            item["supplier_id"],
            item["canonical_sku"],
            item["product_name"],
            bad_q,
            bad_p,
            line_total_override=bad_formula,
        )

    assert len(purchase_orders) == 2000

    # Complete Shipments to exactly 2,200 rows:
    # Currently 2,135 rows; add 30 duplicates + 25 unexpected + 10 invalid = 2,200
    assert len(shipments) == 2135

    for dup_s in range(30):
        orig_sh = shipments[dup_s]
        _add_shipment(
            orig_sh["po_number"].strip(),
            orig_sh["po_line_id"].strip() or None,
            orig_sh["ship_date"],
            orig_sh["sku"],
            orig_sh["shipped_qty"],
            ship_id_override=orig_sh["shipment_id"].strip(),
            carrier_override=orig_sh["carrier"],
            tracking_override=orig_sh["tracking_ref"],
            warehouse_override=orig_sh["warehouse_code"],
            noisy=True,
        )

    for orp_s in range(25):
        item = catalog[(orp_s + 50) % len(catalog)]
        if orp_s < 12:
            # Missing PO reference
            _add_shipment(None, None, "2026-03-01", item["canonical_sku"], rng.randint(5, 25))
        elif orp_s < 20:
            # Unknown PO number
            _add_shipment(
                f"PO-2026-999{orp_s:02d}",
                None,
                "2026-03-02",
                item["canonical_sku"],
                rng.randint(5, 25),
            )
        else:
            # Existing PO, but SKU was never ordered on that PO!
            existing_po = purchase_orders[orp_s]["po_number"].strip().upper()
            wrong_item = catalog[1400 + orp_s]
            _add_shipment(existing_po, None, "2026-03-03", wrong_item["canonical_sku"], 10)

    bad_ship_specs: tuple[tuple[Any, str], ...] = (
        ("N/A", "2026-03-04"),
        ("TBD", "2026-03-04"),
        ("-6", "2026-03-05"),
        ("-12", "2026-03-05"),
        ("2O", "2026-03-06"),
        ("ERR", "2026-03-06"),
        (10, "PENDING"),
        (15, "2026-02-29"),
        (8, "2026-99-99"),
        ("boxes:10", "2026-03-07"),
    )
    for idx_bs, (bad_sq, bad_sdt) in enumerate(bad_ship_specs):
        target_po = purchase_orders[50 + idx_bs]["po_number"].strip()
        item = catalog[50 + idx_bs]
        _add_shipment(target_po, None, bad_sdt, item["canonical_sku"], bad_sq)

    assert len(shipments) == 2200

    # Complete Supplier Invoices to exactly 2,100 rows:
    # Currently 2,040 rows; add 30 duplicates + 20 unexpected + 10 invalid = 2,100
    assert len(invoices) == 2040

    for dup_iv in range(30):
        orig_iv = invoices[dup_iv]
        _add_invoice(
            orig_iv["po_number"].strip(),
            orig_iv["po_line_id"].strip() or None,
            orig_iv["invoice_date"],
            orig_iv["supplier_id"],
            orig_iv["sku"],
            orig_iv["invoiced_qty"],
            orig_iv["unit_price"],
            inv_id_override=(
                orig_iv["invoice_id"].strip(),
                orig_iv["invoice_line_id"].strip(),
            ),
            noisy=True,
        )

    for orp_iv in range(20):
        item = catalog[(orp_iv + 80) % len(catalog)]
        if orp_iv < 10:
            # Missing PO reference
            _add_invoice(
                None,
                None,
                "2026-03-10",
                item["supplier_id"],
                item["canonical_sku"],
                10,
                f"{item['cost']:.2f}",
            )
        elif orp_iv < 16:
            # Unknown PO number
            _add_invoice(
                f"PO-2026-988{orp_iv:02d}",
                None,
                "2026-03-11",
                item["supplier_id"],
                item["canonical_sku"],
                12,
                f"{item['cost']:.2f}",
            )
        else:
            # Existing PO, unordered SKU
            existing_po = purchase_orders[100 + orp_iv]["po_number"].strip().upper()
            wrong_item = catalog[1420 + orp_iv]
            _add_invoice(
                existing_po,
                None,
                "2026-03-12",
                wrong_item["supplier_id"],
                wrong_item["canonical_sku"],
                5,
                f"{wrong_item['cost']:.2f}",
            )

    bad_inv_specs: tuple[tuple[Any, str, str, str | None], ...] = (
        ("TBD", "25.00", "2026-03-12", None),
        ("N/A", "30.00", "2026-03-12", None),
        ("-5", "20.00", "2026-03-13", None),
        (10, "TBD", "2026-03-13", None),
        (10, "-18.50", "2026-03-14", None),
        (10, "22.00", "PENDING", None),
        (10, "22.00", "2026-02-30", None),
        (10, "24.00", "2026-03-15", "=#REF!*I2098"),
        (12, "19.00", "2026-03-15", "=#VALUE!"),
        (8, "35.00", "2026-03-15", "=#DIV/0!"),
    )
    for idx_bi, (bad_iq, bad_ip, bad_idt, bad_amt_f) in enumerate(bad_inv_specs):
        target_po = purchase_orders[150 + idx_bi]["po_number"].strip()
        item = catalog[150 + idx_bi]
        _add_invoice(
            target_po,
            None,
            bad_idt,
            item["supplier_id"],
            item["canonical_sku"],
            bad_iq,
            bad_ip,
            amount_override=bad_amt_f,
        )

    assert len(invoices) == 2100

    status_distribution: dict[str, int] = {}
    for st in expected_po_status.values():
        status_distribution[st] = status_distribution.get(st, 0) + 1
    status_distribution[ReconciliationStatus.DUPLICATE_TRANSACTION.value] = 25
    status_distribution[ReconciliationStatus.DATA_VALIDATION_ERROR.value] = (
        status_distribution.get(ReconciliationStatus.DATA_VALIDATION_ERROR.value, 0) + 15
    )

    # Expand ground truth across all 2,000 PO source rows and compute independent financials
    expected_po_status_by_source_row: dict[str, str] = {}
    for src_row, po_row in enumerate(purchase_orders, start=2):
        if src_row <= 1961:
            norm_lid = str(po_row["po_line_id"]).strip().upper()
            expected_po_status_by_source_row[str(src_row)] = expected_po_status[norm_lid]
        elif src_row <= 1986:
            expected_po_status_by_source_row[str(src_row)] = (
                ReconciliationStatus.DUPLICATE_TRANSACTION.value
            )
        else:
            expected_po_status_by_source_row[str(src_row)] = (
                ReconciliationStatus.DATA_VALIDATION_ERROR.value
            )

    ambiguous_pos = {
        str(po_row["po_number"]).strip().upper()
        for po_row in purchase_orders[:1960]
        if expected_po_status.get(str(po_row["po_line_id"]).strip().upper())
        == ReconciliationStatus.AMBIGUOUS_REFERENCE.value
        and str(po_row["po_line_id"]).strip().upper().endswith("-L02")
    }
    lines_by_po: dict[str, list[str]] = {}
    for po_row in purchase_orders[:1960]:
        pnum = str(po_row["po_number"]).strip().upper()
        lid = str(po_row["po_line_id"]).strip().upper()
        lines_by_po.setdefault(pnum, []).append(lid)

    ship_qty_by_line: dict[tuple[str, str], Decimal] = {}
    for sh_row in shipments[:2135]:
        pnum = str(sh_row["po_number"]).strip().upper()
        if pnum in ambiguous_pos:
            continue
        lid = str(sh_row["po_line_id"]).strip().upper()
        if not lid:
            lid = lines_by_po[pnum][0]
        ship_qty_by_line[(pnum, lid)] = ship_qty_by_line.get((pnum, lid), Decimal("0")) + Decimal(
            str(sh_row["shipped_qty"]).strip()
        )

    inv_agg_by_line: dict[tuple[str, str], tuple[Decimal, Decimal]] = {}
    quarantined_inv_amt = Decimal("0.00")
    for iv_row in invoices[:2040]:
        pnum = str(iv_row["po_number"]).strip().upper()
        iq = Decimal(str(iv_row["invoiced_qty"]).strip())
        ip = Decimal(str(iv_row["unit_price"]).strip())
        amt = (
            excel_round(Decimal(str(iv_row["_amount_override"])), 2)
            if iv_row["_amount_override"] is not None
            else excel_round(iq * ip, 2)
        )
        if pnum in ambiguous_pos:
            quarantined_inv_amt += amt
            continue
        lid = str(iv_row["po_line_id"]).strip().upper()
        if not lid:
            lid = lines_by_po[pnum][0]
        prev_q, prev_a = inv_agg_by_line.get((pnum, lid), (Decimal("0"), Decimal("0.00")))
        inv_agg_by_line[(pnum, lid)] = (prev_q + iq, prev_a + amt)

    for iv_row in invoices[2040:]:
        amt_ovr = iv_row.get("_amount_override")
        if amt_ovr is not None and str(amt_ovr).startswith("=#"):
            continue
        try:
            iq = Decimal(str(iv_row["invoiced_qty"]).strip())
            ip = Decimal(str(iv_row["unit_price"]).strip())
        except InvalidOperation:
            continue
        amt = (
            excel_round(Decimal(str(amt_ovr)), 2)
            if amt_ovr is not None
            else excel_round(iq * ip, 2)
        )
        quarantined_inv_amt += amt

    gt_ordered_qty = Decimal("0")
    gt_shipped_qty = Decimal("0")
    gt_invoiced_qty = Decimal("0")
    gt_outstanding_qty = Decimal("0")
    gt_expected_amt = Decimal("0.00")
    gt_invoiced_amt = Decimal("0.00")
    gt_net_variance = Decimal("0.00")
    gt_abs_exposure = Decimal("0.00")
    gt_unbilled = Decimal("0.00")

    for po_row in purchase_orders[:1960]:
        pnum = str(po_row["po_number"]).strip().upper()
        lid = str(po_row["po_line_id"]).strip().upper()
        oq = Decimal(str(po_row["ordered_qty"]).strip())
        op = Decimal(str(po_row["unit_price"]).strip())
        exp_amt = excel_round(oq * op, 2)

        sq = ship_qty_by_line.get((pnum, lid), Decimal("0"))
        has_inv = (pnum, lid) in inv_agg_by_line
        iq, iamt = inv_agg_by_line.get((pnum, lid), (Decimal("0"), Decimal("0.00")))
        out_q = max(oq - sq, Decimal("0"))
        auth_q = min(sq, oq)
        auth_val = excel_round(auth_q * op, 2)

        gt_ordered_qty += oq
        gt_shipped_qty += sq
        gt_invoiced_qty += iq
        gt_outstanding_qty += out_q
        gt_expected_amt += exp_amt
        gt_invoiced_amt += iamt

        if has_inv:
            gt_net_variance += excel_round(iamt - auth_val, 2)
            price_comp = iamt - excel_round(iq * op, 2)
            qty_comp = excel_round((iq - auth_q) * op, 2)
            gt_abs_exposure += excel_round(abs(price_comp) + abs(qty_comp), 2)
            if Decimal("0") < iq < auth_q:
                gt_unbilled += excel_round((auth_q - iq) * op, 2)
        elif auth_q > Decimal("0"):
            gt_unbilled += auth_val

    po_non_reconciled = sum(
        1
        for st_val in expected_po_status_by_source_row.values()
        if st_val != ReconciliationStatus.RECONCILED.value
    )
    price_mismatch_count = status_distribution.get(ReconciliationStatus.PRICE_MISMATCH.value, 0)
    expected_exception_counts_by_dataset = {
        DatasetKind.PRODUCTS.value: 15,
        DatasetKind.PURCHASE_ORDERS.value: po_non_reconciled - price_mismatch_count,
        DatasetKind.SHIPMENTS.value: 80,
        DatasetKind.INVOICES.value: price_mismatch_count + 75,
    }
    expected_total_exceptions = sum(expected_exception_counts_by_dataset.values())

    ground_truth = {
        "seed": seed,
        "row_counts": {
            DatasetKind.PRODUCTS.value: len(products),
            DatasetKind.PURCHASE_ORDERS.value: len(purchase_orders),
            DatasetKind.SHIPMENTS.value: len(shipments),
            DatasetKind.INVOICES.value: len(invoices),
        },
        "expected_unique_valid_po_lines": 1960,
        "expected_duplicate_counts": {
            DatasetKind.PRODUCTS.value: 10,
            DatasetKind.PURCHASE_ORDERS.value: 25,
            DatasetKind.SHIPMENTS.value: 30,
            DatasetKind.INVOICES.value: 30,
        },
        "expected_po_status_counts": status_distribution,
        "expected_po_line_statuses": expected_po_status,
        "expected_po_status_by_source_row": expected_po_status_by_source_row,
        "expected_exception_counts_by_dataset": expected_exception_counts_by_dataset,
        "expected_total_exceptions": expected_total_exceptions,
        "expected_financial_totals": {
            "total_ordered_qty": str(gt_ordered_qty),
            "total_shipped_qty": str(gt_shipped_qty),
            "total_invoiced_qty": str(gt_invoiced_qty),
            "total_outstanding_qty": str(gt_outstanding_qty),
            "total_expected_amount": f"{excel_round(gt_expected_amt, 2):.2f}",
            "total_invoiced_amount": f"{excel_round(gt_invoiced_amt, 2):.2f}",
            "net_billed_variance": f"{excel_round(gt_net_variance, 2):.2f}",
            "absolute_monetary_discrepancy": f"{excel_round(gt_abs_exposure, 2):.2f}",
            "unbilled_shipped_amount": f"{excel_round(gt_unbilled, 2):.2f}",
            "quarantined_invoice_amount": f"{excel_round(quarantined_inv_amt, 2):.2f}",
        },
    }

    return SyntheticBundle(
        seed=seed,
        products=products,
        purchase_orders=purchase_orders,
        shipments=shipments,
        invoices=invoices,
        ground_truth=ground_truth,
    )


def _to_excel_number_if_clean(val: Any) -> int | Decimal | str:
    """Write clean numeric strings as native Excel numbers while keeping dirty tokens raw."""
    if isinstance(val, int):
        return val
    text = str(val)
    if text != text.strip():
        return text
    try:
        dec = Decimal(text)
        if dec.is_finite():
            return dec
    except Exception:
        pass
    return text


def _save_workbook_atomic(wb: Any, dest: Path) -> None:
    """Save an openpyxl workbook via a sibling temp file and atomic os.replace."""
    fd, tmp_name = tempfile.mkstemp(
        dir=dest.parent,
        prefix=f".{dest.stem}.",
        suffix=".tmp.xlsx",
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, dest)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def _write_text_atomic(dest: Path, content: str) -> None:
    """Write UTF-8 text atomically via a sibling temp file and os.replace."""
    fd, tmp_name = tempfile.mkstemp(
        dir=dest.parent,
        prefix=f".{dest.stem}.",
        suffix=".tmp",
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        tmp_path.write_text(content, encoding="utf-8")
        os.replace(tmp_path, dest)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def generate_demo_datasets(
    output_dir: Path,
    *,
    seed: int = DEFAULT_DEMO_SEED,
) -> dict[str, Path]:
    """Generate the four synthetic Excel workbooks and ground_truth.json in output_dir."""
    import openpyxl

    output_dir.mkdir(parents=True, exist_ok=True)
    bundle = build_synthetic_records(seed=seed)
    written: dict[str, Path] = {}

    # 1. product_master.xlsx
    wb_prod = openpyxl.Workbook()
    ws_prod = wb_prod.active
    assert ws_prod is not None
    ws_prod.title = "Products"
    prod_cols = REQUIRED_COLUMNS[DatasetKind.PRODUCTS]
    ws_prod.append(list(prod_cols))
    for row in bundle.products:
        ws_prod.append(
            [
                _to_excel_number_if_clean(row[c]) if c == "standard_cost" else row[c]
                for c in prod_cols
            ]
        )
    prod_path = output_dir / "product_master.xlsx"
    _save_workbook_atomic(wb_prod, prod_path)
    wb_prod.close()
    written["product_master"] = prod_path

    # 2. purchase_orders.xlsx (with live Excel =G{row}*H{row} formulas)
    wb_po = openpyxl.Workbook()
    ws_po = wb_po.active
    assert ws_po is not None
    ws_po.title = "PurchaseOrders"
    po_cols = REQUIRED_COLUMNS[DatasetKind.PURCHASE_ORDERS]
    ws_po.append(list(po_cols))
    for r_idx, row in enumerate(bundle.purchase_orders, start=2):
        formula_or_override = row["_line_total_override"] or f"=G{r_idx}*H{r_idx}"
        ws_po.append(
            [
                row["po_number"],
                row["po_line_id"],
                row["order_date"],
                row["supplier_id"],
                row["sku"],
                row["product_description"],
                _to_excel_number_if_clean(row["ordered_qty"]),
                _to_excel_number_if_clean(row["unit_price"]),
                row["currency"],
                formula_or_override,
            ]
        )
    po_path = output_dir / "purchase_orders.xlsx"
    _save_workbook_atomic(wb_po, po_path)
    wb_po.close()
    written["purchase_orders"] = po_path

    # 3. shipments.xlsx
    wb_sh = openpyxl.Workbook()
    ws_sh = wb_sh.active
    assert ws_sh is not None
    ws_sh.title = "Shipments"
    sh_cols = REQUIRED_COLUMNS[DatasetKind.SHIPMENTS]
    ws_sh.append(list(sh_cols))
    for row in bundle.shipments:
        ws_sh.append(
            [_to_excel_number_if_clean(row[c]) if c == "shipped_qty" else row[c] for c in sh_cols]
        )
    sh_path = output_dir / "shipments.xlsx"
    _save_workbook_atomic(wb_sh, sh_path)
    wb_sh.close()
    written["shipments"] = sh_path

    # 4. supplier_invoices.xlsx (with live Excel =H{row}*I{row} formulas & intentional overrides)
    wb_inv = openpyxl.Workbook()
    ws_inv = wb_inv.active
    assert ws_inv is not None
    ws_inv.title = "Invoices"
    inv_cols = REQUIRED_COLUMNS[DatasetKind.INVOICES]
    ws_inv.append(list(inv_cols))
    for r_idx, row in enumerate(bundle.invoices, start=2):
        amt_raw = row["_amount_override"]
        amt_cell = (
            (_to_excel_number_if_clean(amt_raw) if not str(amt_raw).startswith("=") else amt_raw)
            if amt_raw is not None
            else f"=H{r_idx}*I{r_idx}"
        )
        ws_inv.append(
            [
                row["invoice_id"],
                row["invoice_line_id"],
                row["po_number"],
                row["po_line_id"],
                row["invoice_date"],
                row["supplier_id"],
                row["sku"],
                _to_excel_number_if_clean(row["invoiced_qty"]),
                _to_excel_number_if_clean(row["unit_price"]),
                amt_cell,
                row["currency"],
            ]
        )
    inv_path = output_dir / "supplier_invoices.xlsx"
    _save_workbook_atomic(wb_inv, inv_path)
    wb_inv.close()
    written["supplier_invoices"] = inv_path

    # 5. ground_truth.json
    gt_path = output_dir / "ground_truth.json"
    _write_text_atomic(gt_path, json.dumps(bundle.ground_truth, indent=2, sort_keys=True) + "\n")
    written["ground_truth"] = gt_path

    return written
