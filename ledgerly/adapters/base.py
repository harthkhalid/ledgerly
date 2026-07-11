"""DataSource protocol — the only interface the engine consumes."""

from __future__ import annotations

from datetime import datetime
from typing import Protocol, runtime_checkable

from ledgerly.models import InventoryLevel, Order, Product, Variant


@runtime_checkable
class DataSource(Protocol):
    """Adapter contract. Engine imports this protocol, never Shopify specifics."""

    def fetch_orders(self, since: datetime, until: datetime) -> list[Order]:
        """Return orders whose created_at falls in [since, until)."""
        ...

    def fetch_products(self) -> tuple[list[Product], list[Variant]]:
        """Return products and their variants."""
        ...

    def fetch_inventory(self) -> list[InventoryLevel]:
        """Return current inventory levels keyed by inventory_item_id."""
        ...
