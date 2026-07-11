"""Shopify Admin GraphQL adapter (API version 2026-01).

Credentials from env / constructor only — never hardcoded, never logged.
HTTP transport is separate from GraphQL→pydantic mapping so mapping can be
unit-tested with mocked payloads.
"""

from __future__ import annotations

import logging
import time
from datetime import datetime
from decimal import Decimal
from typing import Any, Callable

import requests

from ledgerly.models import (
    InventoryLevel,
    LineItem,
    Order,
    Product,
    Variant,
)

API_VERSION = "2026-01"
logger = logging.getLogger(__name__)

# --- GraphQL documents -------------------------------------------------------

ORDERS_QUERY = """
query Orders($cursor: String, $query: String!) {
  orders(first: 50, after: $cursor, query: $query, sortKey: CREATED_AT) {
    pageInfo { hasNextPage endCursor }
    edges {
      node {
        id
        createdAt
        displayFinancialStatus
        totalDiscountsSet { shopMoney { amount } }
        totalShippingPriceSet { shopMoney { amount } }
        totalRefundedSet { shopMoney { amount } }
        customer { id }
        lineItems(first: 50) {
          edges {
            node {
              variant { id sku }
              quantity
              originalUnitPriceSet { shopMoney { amount } }
              discountAllocations {
                allocatedAmountSet { shopMoney { amount } }
              }
            }
          }
        }
      }
    }
  }
}
"""

PRODUCTS_QUERY = """
query Products($cursor: String) {
  products(first: 50, after: $cursor) {
    pageInfo { hasNextPage endCursor }
    edges {
      node {
        id
        title
        productType
        vendor
        variants(first: 50) {
          edges {
            node {
              id
              sku
              title
              price
              inventoryItem {
                id
                unitCost { amount }
              }
            }
          }
        }
      }
    }
  }
}
"""

INVENTORY_QUERY = """
query InventoryItems($cursor: String) {
  inventoryItems(first: 50, after: $cursor) {
    pageInfo { hasNextPage endCursor }
    edges {
      node {
        id
        updatedAt
        inventoryLevels(first: 10) {
          edges {
            node {
              quantities(names: ["available"]) {
                name
                quantity
              }
            }
          }
        }
      }
    }
  }
}
"""


# --- Mapping (pure, no HTTP) -------------------------------------------------

def _gid_tail(gid: str) -> str:
    """Shrink Shopify GIDs to a stable short id: gid://shopify/Order/123 → 123."""
    if not gid:
        return ""
    return gid.rsplit("/", 1)[-1]


def _dec(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value))


def _parse_dt(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)


def map_products_page(payload: dict) -> tuple[list[Product], list[Variant]]:
    """Map one products GraphQL page into Product + Variant models."""
    products: list[Product] = []
    variants: list[Variant] = []
    edges = payload.get("data", {}).get("products", {}).get("edges", [])
    for edge in edges:
        node = edge["node"]
        pid = _gid_tail(node["id"])
        products.append(
            Product(
                id=pid,
                title=node.get("title") or "",
                product_type=node.get("productType") or "",
                vendor=node.get("vendor") or "",
            )
        )
        for vedge in node.get("variants", {}).get("edges", []):
            vnode = vedge["node"]
            inv = vnode.get("inventoryItem") or {}
            cost_amount = (inv.get("unitCost") or {}).get("amount")
            variants.append(
                Variant(
                    id=_gid_tail(vnode["id"]),
                    product_id=pid,
                    sku=vnode.get("sku") or f"UNSKU-{_gid_tail(vnode['id'])}",
                    title=vnode.get("title") or "",
                    price=_dec(vnode.get("price")),
                    cost=_dec(cost_amount),
                    inventory_item_id=_gid_tail(inv.get("id") or ""),
                )
            )
    return products, variants


def map_orders_page(payload: dict) -> list[Order]:
    """Map one orders GraphQL page into Order models."""
    orders: list[Order] = []
    edges = payload.get("data", {}).get("orders", {}).get("edges", [])
    for edge in edges:
        node = edge["node"]
        customer = node.get("customer") or {}
        line_items: list[LineItem] = []
        for ledge in node.get("lineItems", {}).get("edges", []):
            lnode = ledge["node"]
            variant = lnode.get("variant") or {}
            disc = Decimal("0")
            for alloc in lnode.get("discountAllocations") or []:
                amt = (
                    (alloc.get("allocatedAmountSet") or {})
                    .get("shopMoney", {})
                    .get("amount")
                )
                disc += _dec(amt)
            price = (lnode.get("originalUnitPriceSet") or {}).get("shopMoney", {}).get(
                "amount"
            )
            line_items.append(
                LineItem(
                    variant_id=_gid_tail(variant.get("id") or ""),
                    sku=variant.get("sku") or "",
                    quantity=int(lnode.get("quantity") or 0),
                    unit_price=_dec(price),
                    discount_allocated=disc,
                )
            )
        orders.append(
            Order(
                id=_gid_tail(node["id"]),
                created_at=_parse_dt(node["createdAt"]),
                customer_id=_gid_tail(customer.get("id") or "0"),
                line_items=line_items,
                total_discounts=_dec(
                    (node.get("totalDiscountsSet") or {})
                    .get("shopMoney", {})
                    .get("amount")
                ),
                total_shipping_charged=_dec(
                    (node.get("totalShippingPriceSet") or {})
                    .get("shopMoney", {})
                    .get("amount")
                ),
                financial_status=(node.get("displayFinancialStatus") or "paid").lower(),
                refunded_amount=_dec(
                    (node.get("totalRefundedSet") or {})
                    .get("shopMoney", {})
                    .get("amount")
                ),
            )
        )
    return orders


def map_inventory_page(payload: dict) -> list[InventoryLevel]:
    """Map one inventoryItems GraphQL page into InventoryLevel models."""
    levels: list[InventoryLevel] = []
    edges = payload.get("data", {}).get("inventoryItems", {}).get("edges", [])
    for edge in edges:
        node = edge["node"]
        iid = _gid_tail(node["id"])
        available = 0
        for ledges in node.get("inventoryLevels", {}).get("edges", []):
            for q in ledges["node"].get("quantities") or []:
                if q.get("name") == "available":
                    available += int(q.get("quantity") or 0)
        levels.append(
            InventoryLevel(
                inventory_item_id=iid,
                available=available,
                updated_at=_parse_dt(node.get("updatedAt") or "2026-01-01T00:00:00Z"),
            )
        )
    return levels


# --- HTTP client -------------------------------------------------------------

class ShopifySource:
    """DataSource backed by Shopify Admin GraphQL API 2026-01."""

    def __init__(
        self,
        store: str,
        access_token: str,
        session: requests.Session | None = None,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        store = store.strip().removesuffix(".myshopify.com")
        self._url = (
            f"https://{store}.myshopify.com/admin/api/{API_VERSION}/graphql.json"
        )
        self._token = access_token
        self._session = session or requests.Session()
        self._sleep = sleep

    def _headers(self) -> dict[str, str]:
        return {
            "Content-Type": "application/json",
            "X-Shopify-Access-Token": self._token,
        }

    def _post(self, query: str, variables: dict | None = None) -> dict:
        """Execute a GraphQL request with THROTTLED backoff."""
        body = {"query": query, "variables": variables or {}}
        for attempt in range(8):
            resp = self._session.post(
                self._url, json=body, headers=self._headers(), timeout=60
            )
            resp.raise_for_status()
            payload = resp.json()
            errors = payload.get("errors") or []
            throttled = any(
                (e.get("extensions") or {}).get("code") == "THROTTLED"
                or "THROTTLED" in str(e.get("message", "")).upper()
                for e in errors
            )
            if throttled:
                cost = (payload.get("extensions") or {}).get("cost") or {}
                throttle = cost.get("throttleStatus") or {}
                restore = float(throttle.get("restoreRate") or 100.0)
                currently = float(throttle.get("currentlyAvailable") or 0)
                maximum = float(throttle.get("maximumAvailable") or 1000.0)
                # Wait until we regain a meaningful chunk of budget
                needed = max(maximum * 0.2 - currently, 1.0)
                wait = needed / restore if restore else 2.0
                wait = min(max(wait, 0.5), 10.0)
                logger.warning("Shopify THROTTLED — backing off %.1fs", wait)
                self._sleep(wait)
                continue
            if errors:
                # Non-throttle GraphQL errors
                messages = "; ".join(e.get("message", str(e)) for e in errors)
                raise RuntimeError(f"Shopify GraphQL error: {messages}")
            return payload
        raise RuntimeError("Shopify API still throttled after retries")

    def _paginate(self, query: str, root: str, variables: dict | None = None):
        """Yield successive page payloads for a connection under data[root]."""
        variables = dict(variables or {})
        cursor = None
        while True:
            variables["cursor"] = cursor
            payload = self._post(query, variables)
            yield payload
            page_info = payload.get("data", {}).get(root, {}).get("pageInfo", {})
            if not page_info.get("hasNextPage"):
                break
            cursor = page_info.get("endCursor")

    def fetch_orders(self, since: datetime, until: datetime) -> list[Order]:
        # Shopify search query syntax for created_at window
        q = (
            f"created_at:>={since.strftime('%Y-%m-%d')} "
            f"created_at:<{until.strftime('%Y-%m-%d')}"
        )
        orders: list[Order] = []
        for page in self._paginate(ORDERS_QUERY, "orders", {"query": q}):
            orders.extend(map_orders_page(page))
        return orders

    def fetch_products(self) -> tuple[list[Product], list[Variant]]:
        products: list[Product] = []
        variants: list[Variant] = []
        for page in self._paginate(PRODUCTS_QUERY, "products"):
            p, v = map_products_page(page)
            products.extend(p)
            variants.extend(v)
        return products, variants

    def fetch_inventory(self) -> list[InventoryLevel]:
        levels: list[InventoryLevel] = []
        for page in self._paginate(INVENTORY_QUERY, "inventoryItems"):
            levels.extend(map_inventory_page(page))
        return levels
