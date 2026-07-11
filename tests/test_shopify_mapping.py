"""Mocked-response unit tests for Shopify GraphQL → pydantic mapping."""

from __future__ import annotations

from decimal import Decimal

from ledgerly.adapters.shopify import (
    map_inventory_page,
    map_orders_page,
    map_products_page,
)

PRODUCTS_PAYLOAD = {
    "data": {
        "products": {
            "pageInfo": {"hasNextPage": False, "endCursor": None},
            "edges": [
                {
                    "node": {
                        "id": "gid://shopify/Product/1001",
                        "title": "Midnight Amber",
                        "productType": "Eau de Parfum",
                        "vendor": "Ledgerly Fragrances",
                        "variants": {
                            "edges": [
                                {
                                    "node": {
                                        "id": "gid://shopify/ProductVariant/2001",
                                        "sku": "LF-01-78",
                                        "title": "50ml",
                                        "price": "78.00",
                                        "inventoryItem": {
                                            "id": "gid://shopify/InventoryItem/3001",
                                            "unitCost": {"amount": "22.50"},
                                        },
                                    }
                                }
                            ]
                        },
                    }
                }
            ],
        }
    }
}

ORDERS_PAYLOAD = {
    "data": {
        "orders": {
            "pageInfo": {"hasNextPage": False, "endCursor": None},
            "edges": [
                {
                    "node": {
                        "id": "gid://shopify/Order/9001",
                        "createdAt": "2026-06-15T14:30:00Z",
                        "displayFinancialStatus": "PAID",
                        "totalDiscountsSet": {"shopMoney": {"amount": "5.00"}},
                        "totalShippingPriceSet": {"shopMoney": {"amount": "12.50"}},
                        "totalRefundedSet": {"shopMoney": {"amount": "0.00"}},
                        "customer": {"id": "gid://shopify/Customer/5001"},
                        "lineItems": {
                            "edges": [
                                {
                                    "node": {
                                        "variant": {
                                            "id": "gid://shopify/ProductVariant/2001",
                                            "sku": "LF-01-78",
                                        },
                                        "quantity": 2,
                                        "originalUnitPriceSet": {
                                            "shopMoney": {"amount": "78.00"}
                                        },
                                        "discountAllocations": [
                                            {
                                                "allocatedAmountSet": {
                                                    "shopMoney": {"amount": "5.00"}
                                                }
                                            }
                                        ],
                                    }
                                }
                            ]
                        },
                    }
                }
            ],
        }
    }
}

INVENTORY_PAYLOAD = {
    "data": {
        "inventoryItems": {
            "pageInfo": {"hasNextPage": False, "endCursor": None},
            "edges": [
                {
                    "node": {
                        "id": "gid://shopify/InventoryItem/3001",
                        "updatedAt": "2026-06-30T12:00:00Z",
                        "inventoryLevels": {
                            "edges": [
                                {
                                    "node": {
                                        "quantities": [
                                            {"name": "available", "quantity": 42}
                                        ]
                                    }
                                }
                            ]
                        },
                    }
                }
            ],
        }
    }
}


def test_map_products_page():
    products, variants = map_products_page(PRODUCTS_PAYLOAD)
    assert len(products) == 1
    assert products[0].id == "1001"
    assert products[0].title == "Midnight Amber"
    assert len(variants) == 1
    assert variants[0].sku == "LF-01-78"
    assert variants[0].price == Decimal("78.00")
    assert variants[0].cost == Decimal("22.50")
    assert variants[0].inventory_item_id == "3001"


def test_map_orders_page():
    orders = map_orders_page(ORDERS_PAYLOAD)
    assert len(orders) == 1
    o = orders[0]
    assert o.id == "9001"
    assert o.customer_id == "5001"
    assert o.financial_status == "paid"
    assert o.total_shipping_charged == Decimal("12.50")
    assert len(o.line_items) == 1
    li = o.line_items[0]
    assert li.sku == "LF-01-78"
    assert li.quantity == 2
    assert li.unit_price == Decimal("78.00")
    assert li.discount_allocated == Decimal("5.00")


def test_map_inventory_page():
    levels = map_inventory_page(INVENTORY_PAYLOAD)
    assert len(levels) == 1
    assert levels[0].inventory_item_id == "3001"
    assert levels[0].available == 42
