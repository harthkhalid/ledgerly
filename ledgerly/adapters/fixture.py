"""Fixture adapter — loads sample_data.json so the repo runs with zero credentials."""

from __future__ import annotations

import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from ledgerly.models import (
    Customer,
    InventoryLevel,
    LineItem,
    Order,
    Product,
    Variant,
)

_FIXTURE_PATH = Path(__file__).resolve().parent.parent / "fixtures" / "sample_data.json"


def _parse_dt(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)


def _dec(value: Any) -> Decimal:
    return Decimal(str(value))


class FixtureSource:
    """Loads realistic fake DTC fragrance data from the bundled JSON fixture."""

    def __init__(self, path: Path | None = None) -> None:
        self._path = path or _FIXTURE_PATH
        with open(self._path, encoding="utf-8") as fh:
            raw = json.load(fh)
        self._products = [Product.model_validate(p) for p in raw["products"]]
        self._variants = [
            Variant(
                id=v["id"],
                product_id=v["product_id"],
                sku=v["sku"],
                title=v["title"],
                price=_dec(v["price"]),
                cost=_dec(v["cost"]),
                inventory_item_id=v["inventory_item_id"],
            )
            for v in raw["variants"]
        ]
        self._inventory = [
            InventoryLevel(
                inventory_item_id=i["inventory_item_id"],
                available=i["available"],
                updated_at=_parse_dt(i["updated_at"]),
            )
            for i in raw["inventory"]
        ]
        self._customers = [
            Customer(
                id=c["id"],
                created_at=_parse_dt(c["created_at"]),
                email_hash=c["email_hash"],
            )
            for c in raw["customers"]
        ]
        self._orders = [
            Order(
                id=o["id"],
                created_at=_parse_dt(o["created_at"]),
                customer_id=o["customer_id"],
                line_items=[
                    LineItem(
                        variant_id=li["variant_id"],
                        sku=li["sku"],
                        quantity=li["quantity"],
                        unit_price=_dec(li["unit_price"]),
                        discount_allocated=_dec(li.get("discount_allocated", 0)),
                    )
                    for li in o["line_items"]
                ],
                total_discounts=_dec(o.get("total_discounts", 0)),
                total_shipping_charged=_dec(o.get("total_shipping_charged", 0)),
                financial_status=o.get("financial_status", "paid"),
                refunded_amount=_dec(o.get("refunded_amount", 0)),
            )
            for o in raw["orders"]
        ]

    def fetch_orders(self, since: datetime, until: datetime) -> list[Order]:
        return [o for o in self._orders if since <= o.created_at < until]

    def fetch_products(self) -> tuple[list[Product], list[Variant]]:
        return list(self._products), list(self._variants)

    def fetch_inventory(self) -> list[InventoryLevel]:
        return list(self._inventory)

    def fetch_customers(self) -> list[Customer]:
        return list(self._customers)

    def fetch_all_orders(self) -> list[Order]:
        """Return every order in the fixture (used for cohort / trailing windows)."""
        return list(self._orders)
