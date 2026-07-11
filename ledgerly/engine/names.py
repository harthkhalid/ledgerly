"""Defined-name constants and helpers for workbook-level named ranges."""

from __future__ import annotations

from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedName

# Exact names required by the spec — every downstream formula references these.
SHIPPING_COST_PER_ORDER = "ShippingCostPerOrder"
PACKAGING_COST_PER_ORDER = "PackagingCostPerOrder"
PAYMENT_PROCESSING_PCT = "PaymentProcessingPct"
MONTHLY_AD_SPEND = "MonthlyAdSpend"
MARGIN_FLOOR_PCT = "MarginFloorPct"
LEAD_TIME_DAYS = "LeadTimeDays"
SAFETY_STOCK_DAYS = "SafetyStockDays"
REPORT_MONTH = "ReportMonth"

ALL_CONTROL_NAMES = (
    SHIPPING_COST_PER_ORDER,
    PACKAGING_COST_PER_ORDER,
    PAYMENT_PROCESSING_PCT,
    MONTHLY_AD_SPEND,
    MARGIN_FLOOR_PCT,
    LEAD_TIME_DAYS,
    SAFETY_STOCK_DAYS,
    REPORT_MONTH,
)

# Defaults (Controls tab only — never hardcode elsewhere)
DEFAULTS = {
    SHIPPING_COST_PER_ORDER: 12.50,
    PACKAGING_COST_PER_ORDER: 3.25,
    PAYMENT_PROCESSING_PCT: 0.029,
    MONTHLY_AD_SPEND: 4500,
    MARGIN_FLOOR_PCT: 0.35,
    LEAD_TIME_DAYS: 30,
    SAFETY_STOCK_DAYS: 14,
    REPORT_MONTH: "2026-06",
}

# Controls sheet cell map: name → absolute cell address (value column)
# Labels in column A, values in column B, starting row 3
CONTROL_CELLS = {
    SHIPPING_COST_PER_ORDER: "Controls!$B$3",
    PACKAGING_COST_PER_ORDER: "Controls!$B$4",
    PAYMENT_PROCESSING_PCT: "Controls!$B$5",
    MONTHLY_AD_SPEND: "Controls!$B$6",
    MARGIN_FLOOR_PCT: "Controls!$B$7",
    LEAD_TIME_DAYS: "Controls!$B$8",
    SAFETY_STOCK_DAYS: "Controls!$B$9",
    REPORT_MONTH: "Controls!$B$10",
}

# Row index on Controls for each name (for writing values)
CONTROL_ROWS = {
    SHIPPING_COST_PER_ORDER: 3,
    PACKAGING_COST_PER_ORDER: 4,
    PAYMENT_PROCESSING_PCT: 5,
    MONTHLY_AD_SPEND: 6,
    MARGIN_FLOOR_PCT: 7,
    LEAD_TIME_DAYS: 8,
    SAFETY_STOCK_DAYS: 9,
    REPORT_MONTH: 10,
}

CONTROL_LABELS = {
    SHIPPING_COST_PER_ORDER: "Shipping Cost Per Order",
    PACKAGING_COST_PER_ORDER: "Packaging Cost Per Order",
    PAYMENT_PROCESSING_PCT: "Payment Processing %",
    MONTHLY_AD_SPEND: "Monthly Ad Spend",
    MARGIN_FLOOR_PCT: "Margin Floor %",
    LEAD_TIME_DAYS: "Lead Time (Days)",
    SAFETY_STOCK_DAYS: "Safety Stock (Days)",
    REPORT_MONTH: "Report Month",
}


def register_control_names(wb: Workbook) -> None:
    """Register all eight Controls named ranges on the workbook."""
    for name, attr in CONTROL_CELLS.items():
        # Remove existing if re-building
        if name in wb.defined_names:
            del wb.defined_names[name]
        wb.defined_names.add(DefinedName(name=name, attr_text=attr))


def require_names(wb: Workbook, required: tuple[str, ...] | list[str]) -> None:
    """Raise ValueError listing every missing name — never silently fall back."""
    missing = [n for n in required if n not in wb.defined_names]
    if missing:
        raise ValueError(
            f"Template is missing required named ranges: {', '.join(sorted(missing))}"
        )
