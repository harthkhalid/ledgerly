"""Shared helpers for engine tests."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from ledgerly.adapters.fixture import FixtureSource
from ledgerly.engine.workbook import build_workbook, save_workbook


def make_report(tmp_path: Path, month: str = "2026-06") -> tuple[Path, Workbook]:
    src = FixtureSource()
    products, variants = src.fetch_products()
    inventory = src.fetch_inventory()
    customers = src.fetch_customers()
    all_orders = src.fetch_all_orders()
    year, mon = map(int, month.split("-"))
    since = datetime(year, mon, 1)
    until = datetime(year + (1 if mon == 12 else 0), 1 if mon == 12 else mon + 1, 1)
    # Engine uses all_orders for cohorts/velocity; month filter is informational
    wb = build_workbook(
        orders=all_orders,
        products=products,
        variants=variants,
        inventory=inventory,
        customers=customers,
        report_month=month,
        all_orders=all_orders,
    )
    path = tmp_path / f"ledgerly_{month}.xlsx"
    save_workbook(wb, path)
    reloaded = load_workbook(path)
    return path, reloaded
