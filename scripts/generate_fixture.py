"""Generate fixtures/sample_data.json — ~14 products, ~40 variants, ~900 orders, ~500 customers."""

from __future__ import annotations

import hashlib
import json
import random
from datetime import datetime, timedelta
from decimal import Decimal
from pathlib import Path

SEED = 42
START = datetime(2026, 1, 1)
END = datetime(2026, 7, 1)  # exclusive — 6 months of data

PRODUCT_DEFS = [
    ("Midnight Amber", "Eau de Parfum", "Ledgerly Fragrances", [48, 78, 118]),
    ("Cedar & Smoke", "Eau de Parfum", "Ledgerly Fragrances", [52, 85, 125]),
    ("Rose Absolute", "Eau de Parfum", "Ledgerly Fragrances", [55, 90, 135]),
    ("Ocean Drift", "Eau de Toilette", "Ledgerly Fragrances", [38, 62, 95]),
    ("Vanilla Noir", "Eau de Parfum", "Ledgerly Fragrances", [50, 82, 120]),
    ("Citrus Grove", "Eau de Toilette", "Ledgerly Fragrances", [35, 58, 88]),
    ("Sandalwood Mist", "Eau de Parfum", "Ledgerly Fragrances", [54, 88, 130]),
    ("Jasmine Veil", "Eau de Parfum", "Ledgerly Fragrances", [56, 92, 138]),
    ("Leather Bound", "Eau de Parfum", "Ledgerly Fragrances", [60, 98, 145]),
    ("Fig & Moss", "Eau de Toilette", "Ledgerly Fragrances", [40, 65, 98]),
    ("Discovery Set", "Gift Set", "Ledgerly Fragrances", [42]),
    ("Travel Duo", "Gift Set", "Ledgerly Fragrances", [55]),
    ("Candle — Amber", "Home", "Ledgerly Home", [36, 48]),
    ("Candle — Cedar", "Home", "Ledgerly Home", [36, 48]),
]

# Cost as fraction of price; a few SKUs intentionally negative-margin after fees
COST_RATIOS = {
    "default": 0.28,
    "slow": 0.32,
    "negative": 0.72,  # high COGS → negative contribution after allocations
}

SIZE_BY_INDEX = {
    0: "30ml",
    1: "50ml",
    2: "100ml",
}
SPECIAL_SIZES = {
    ("Discovery Set", 0): "5x2ml",
    ("Travel Duo", 0): "2x15ml",
    ("Candle — Amber", 0): "6oz",
    ("Candle — Amber", 1): "10oz",
    ("Candle — Cedar", 0): "6oz",
    ("Candle — Cedar", 1): "10oz",
}


def _size_label(title: str, vi: int) -> str:
    return SPECIAL_SIZES.get((title, vi), SIZE_BY_INDEX.get(vi, f"size{vi}"))


def _sku(product_idx: int, price: int) -> str:
    return f"LF-{product_idx + 1:02d}-{price}"


def _hash_email(n: int) -> str:
    return hashlib.sha256(f"customer{n}@example.com".encode()).hexdigest()[:16]


def generate() -> dict:
    rng = random.Random(SEED)

    products = []
    variants = []
    inventory = []
    # Mark a few SKUs as slow-moving / negative-margin
    negative_skus: set[str] = set()
    slow_skus: set[str] = set()

    for pi, (title, ptype, vendor, prices) in enumerate(PRODUCT_DEFS):
        pid = f"prod_{pi + 1:03d}"
        products.append({"id": pid, "title": title, "product_type": ptype, "vendor": vendor})
        for vi, price in enumerate(prices):
            sku = _sku(pi, price)
            vid = f"var_{pi + 1:03d}_{vi + 1}"
            iid = f"inv_{pi + 1:03d}_{vi + 1}"
            size = _size_label(title, vi)
            # Assign cost profile
            if pi in (8, 9) and vi == 0:  # Leather Bound 30ml, Fig & Moss 30ml
                ratio = COST_RATIOS["negative"]
                negative_skus.add(sku)
            elif pi in (12, 13):  # candles — slow movers
                ratio = COST_RATIOS["slow"]
                slow_skus.add(sku)
            else:
                ratio = COST_RATIOS["default"] + rng.uniform(-0.04, 0.04)
            cost = round(price * ratio, 2)
            variants.append({
                "id": vid,
                "product_id": pid,
                "sku": sku,
                "title": f"{title} / {size}",
                "price": price,
                "cost": cost,
                "inventory_item_id": iid,
            })
            # Inventory: slow movers have high stock, popular have moderate
            if sku in slow_skus:
                avail = rng.randint(180, 320)
            elif sku in negative_skus:
                avail = rng.randint(40, 90)
            else:
                avail = rng.randint(15, 200)
            inventory.append({
                "inventory_item_id": iid,
                "available": avail,
                "updated_at": "2026-06-30T12:00:00",
            })

    # ~500 customers with acquisition spread across 6 months
    customers = []
    n_customers = 500
    for ci in range(n_customers):
        # More acquisitions in spring (seasonal)
        month_weights = [0.12, 0.14, 0.18, 0.20, 0.18, 0.18]  # Jan–Jun
        month = rng.choices(range(6), weights=month_weights, k=1)[0]
        day = rng.randint(1, 28)
        created = datetime(2026, month + 1, day, rng.randint(8, 20), rng.randint(0, 59))
        customers.append({
            "id": f"cust_{ci + 1:04d}",
            "created_at": created.isoformat(),
            "email_hash": _hash_email(ci + 1),
        })

    # Build variant lookup for order generation
    var_by_sku = {v["sku"]: v for v in variants}
    all_skus = list(var_by_sku.keys())
    # Popularity weights — slow SKUs get low weight
    weights = []
    for sku in all_skus:
        if sku in slow_skus:
            weights.append(0.3)
        elif sku in negative_skus:
            weights.append(1.5)
        else:
            weights.append(2.0 + rng.uniform(0, 3))

    # Seasonal revenue multipliers by month
    season = {1: 0.75, 2: 0.85, 3: 1.0, 4: 1.15, 5: 1.25, 6: 1.35}

    orders = []
    order_id = 1
    # Track which customers have ordered (for repeat behavior)
    customer_order_months: dict[str, list[int]] = {c["id"]: [] for c in customers}

    # First orders: most customers order in their acquisition month
    for cust in customers:
        if rng.random() < 0.92:  # 8% never convert
            created = datetime.fromisoformat(cust["created_at"])
            _place_order(
                orders, order_id, cust["id"], created, all_skus, weights, var_by_sku, rng
            )
            customer_order_months[cust["id"]].append(created.month)
            order_id += 1

    # Repeat purchases across subsequent months (cohort LTV non-trivial)
    for cust in customers:
        acq = datetime.fromisoformat(cust["created_at"])
        for month_offset in range(1, 6):
            target_month = acq.month + month_offset
            if target_month > 6:
                break
            # Retention decays
            p_repeat = 0.35 * (0.7 ** (month_offset - 1)) * season.get(target_month, 1.0)
            if rng.random() < p_repeat:
                day = rng.randint(1, 28)
                dt = datetime(2026, target_month, day, rng.randint(9, 21), rng.randint(0, 59))
                _place_order(
                    orders, order_id, cust["id"], dt, all_skus, weights, var_by_sku, rng
                )
                customer_order_months[cust["id"]].append(target_month)
                order_id += 1

    # Pad with guest-like extra orders to reach ~900
    while len(orders) < 900:
        cust = rng.choice(customers)
        month = rng.choices(range(1, 7), weights=[season[m] for m in range(1, 7)], k=1)[0]
        day = rng.randint(1, 28)
        dt = datetime(2026, month, day, rng.randint(8, 22), rng.randint(0, 59))
        # Don't order before acquisition
        acq = datetime.fromisoformat(cust["created_at"])
        if dt < acq:
            continue
        _place_order(orders, order_id, cust["id"], dt, all_skus, weights, var_by_sku, rng)
        order_id += 1

    return {
        "products": products,
        "variants": variants,
        "inventory": inventory,
        "customers": customers,
        "orders": orders,
        "meta": {
            "negative_margin_skus": sorted(negative_skus),
            "slow_moving_skus": sorted(slow_skus),
            "order_count": len(orders),
            "customer_count": len(customers),
            "variant_count": len(variants),
            "product_count": len(products),
        },
    }


def _place_order(orders, order_id, customer_id, created, all_skus, weights, var_by_sku, rng):
    n_lines = rng.choices([1, 2, 3], weights=[0.65, 0.28, 0.07], k=1)[0]
    chosen = rng.choices(all_skus, weights=weights, k=n_lines)
    # Deduplicate SKUs in one order
    seen = set()
    line_items = []
    total_discount = Decimal("0")
    for sku in chosen:
        if sku in seen:
            continue
        seen.add(sku)
        v = var_by_sku[sku]
        qty = rng.choices([1, 2], weights=[0.85, 0.15], k=1)[0]
        unit_price = Decimal(str(v["price"]))
        # Occasional line discount
        disc = Decimal("0")
        if rng.random() < 0.12:
            disc = (unit_price * qty * Decimal("0.10")).quantize(Decimal("0.01"))
            total_discount += disc
        line_items.append({
            "variant_id": v["id"],
            "sku": sku,
            "quantity": qty,
            "unit_price": float(unit_price),
            "discount_allocated": float(disc),
        })
    if not line_items:
        return
    shipping = float(rng.choice([0, 5.99, 8.99, 12.50]))
    orders.append({
        "id": f"ord_{order_id:05d}",
        "created_at": created.isoformat(),
        "customer_id": customer_id,
        "line_items": line_items,
        "total_discounts": float(total_discount),
        "total_shipping_charged": shipping,
        "financial_status": "paid",
        "refunded_amount": 0.0 if rng.random() > 0.03 else float(rng.choice([10, 25, 48])),
    })


def main() -> None:
    data = generate()
    out = Path(__file__).resolve().parent.parent / "ledgerly" / "fixtures" / "sample_data.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
    meta = data["meta"]
    print(
        f"Wrote {out}\n"
        f"  products={meta['product_count']} variants={meta['variant_count']} "
        f"customers={meta['customer_count']} orders={meta['order_count']}\n"
        f"  negative_margin_skus={meta['negative_margin_skus']}\n"
        f"  slow_moving_skus={meta['slow_moving_skus']}"
    )


if __name__ == "__main__":
    main()
