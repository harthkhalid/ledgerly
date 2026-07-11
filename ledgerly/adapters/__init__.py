"""Adapter package exports."""

from ledgerly.adapters.base import DataSource
from ledgerly.adapters.fixture import FixtureSource

__all__ = ["DataSource", "FixtureSource"]

# ShopifySource is imported lazily so fixture-only usage needs no requests session setup.

