"""Tests for pydantic data models."""

from datetime import datetime
from decimal import Decimal

import pytest
from pydantic import ValidationError

from ledgerly.adapters.fixture import FixtureSource
from ledgerly.models import (
    Customer,
    InventoryLevel,
    LineItem,
    Order,
    Product,
    Variant,
)


def test_product_roundtrip():
    p = Product(id="p1", title="Midnight Amber", product_type="Eau de Parfum", vendor="LF")
    assert p.title == "Midnight Amber"
    assert p.model_dump()["id"] == "p1"


def test_variant_money_is_decimal():
    v = Variant(
        id="v1",
        product_id="p1",
        sku="LF-01-50",
        title="50ml",
        price=Decimal("78.00"),
        cost=Decimal("22.00"),
        inventory_item_id="i1",
    )
    assert isinstance(v.price, Decimal)
    assert isinstance(v.cost, Decimal)
    assert v.price - v.cost == Decimal("56.00")


def test_order_with_line_items():
    li = LineItem(
        variant_id="v1",
        sku="LF-01-50",
        quantity=2,
        unit_price=Decimal("78.00"),
        discount_allocated=Decimal("5.00"),
    )
    o = Order(
        id="o1",
        created_at=datetime(2026, 6, 1, 12, 0),
        customer_id="c1",
        line_items=[li],
        total_discounts=Decimal("5.00"),
        total_shipping_charged=Decimal("12.50"),
        financial_status="paid",
        refunded_amount=Decimal("0"),
    )
    assert len(o.line_items) == 1
    assert o.line_items[0].quantity == 2


def test_customer_stores_hash_not_email():
    c = Customer(
        id="c1",
        created_at=datetime(2026, 1, 15),
        email_hash="abc123def456",
    )
    dumped = c.model_dump()
    assert "email" not in dumped
    assert dumped["email_hash"] == "abc123def456"


def test_inventory_level():
    inv = InventoryLevel(
        inventory_item_id="i1",
        available=42,
        updated_at=datetime(2026, 6, 30),
    )
    assert inv.available == 42


def test_variant_rejects_missing_sku():
    with pytest.raises(ValidationError):
        Variant(
            id="v1",
            product_id="p1",
            title="x",
            price=Decimal("10"),
            cost=Decimal("5"),
            inventory_item_id="i1",
        )


def test_fixture_source_loads():
    src = FixtureSource()
    products, variants = src.fetch_products()
    inventory = src.fetch_inventory()
    customers = src.fetch_customers()
    orders = src.fetch_all_orders()

    assert len(products) >= 14
    assert len(variants) >= 35
    assert len(customers) >= 450
    assert len(orders) >= 800
    assert len(inventory) == len(variants)

    # Money fields are Decimal
    assert isinstance(variants[0].price, Decimal)
    assert isinstance(orders[0].line_items[0].unit_price, Decimal)

    # Date filter works
    since = datetime(2026, 6, 1)
    until = datetime(2026, 7, 1)
    june = src.fetch_orders(since, until)
    assert all(since <= o.created_at < until for o in june)
    assert len(june) > 0
