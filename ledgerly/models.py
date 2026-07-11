"""Pydantic data models for Ledgerly. All money fields use Decimal."""

from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from typing import Optional

from pydantic import BaseModel, Field


class Product(BaseModel):
    id: str
    title: str
    product_type: str
    vendor: str


class Variant(BaseModel):
    id: str
    product_id: str
    sku: str
    title: str
    price: Decimal
    cost: Decimal  # COGS unit cost
    inventory_item_id: str


class InventoryLevel(BaseModel):
    inventory_item_id: str
    available: int
    updated_at: datetime


class Customer(BaseModel):
    id: str
    created_at: datetime  # acquisition date
    email_hash: str  # never store raw PII


class LineItem(BaseModel):
    variant_id: str
    sku: str
    quantity: int
    unit_price: Decimal
    discount_allocated: Decimal = Field(default=Decimal("0"))


class Order(BaseModel):
    id: str
    created_at: datetime
    customer_id: str
    line_items: list[LineItem]
    total_discounts: Decimal = Field(default=Decimal("0"))
    total_shipping_charged: Decimal = Field(default=Decimal("0"))
    financial_status: str = "paid"
    refunded_amount: Decimal = Field(default=Decimal("0"))


class CatalogBundle(BaseModel):
    """Convenience container returned by adapters."""

    products: list[Product]
    variants: list[Variant]
    inventory: list[InventoryLevel]
    customers: list[Customer]
    orders: list[Order]
