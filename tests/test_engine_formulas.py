"""Assert analysis tabs contain live Excel formulas, not static values."""

from __future__ import annotations

from tests.conftest import make_report


def test_sku_margin_shipping_formula_uses_named_range(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["SKU Margin"]
    # First data row allocated shipping is column G
    cell = ws["G3"].value
    assert isinstance(cell, str)
    assert cell.startswith("=")
    assert "ShippingCostPerOrder" in cell


def test_sku_margin_packaging_and_processing_named_ranges(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["SKU Margin"]
    assert "PackagingCostPerOrder" in str(ws["H3"].value)
    assert "PaymentProcessingPct" in str(ws["I3"].value)
    assert "MonthlyAdSpend" in str(ws["J3"].value)


def test_rawdata_net_revenue_is_formula(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["RawData"]
    # net_revenue is column K
    val = ws["K3"].value
    assert isinstance(val, str)
    assert val.startswith("=")
    assert "H3" in val and "I3" in val


def test_rawdata_cogs_and_cohort_are_formulas(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["RawData"]
    assert str(ws["M3"].value).startswith("=")
    assert str(ws["E3"].value).startswith("=")
    assert "TEXT" in str(ws["E3"].value)


def test_dashboard_kpis_are_formulas(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["Dashboard"]
    for row in range(4, 8):
        val = ws.cell(row=row, column=2).value
        assert isinstance(val, str), f"B{row} should be formula, got {val!r}"
        assert val.startswith("="), f"B{row} should start with =, got {val!r}"


def test_inventory_formulas_reference_named_ranges(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["Inventory"]
    status = str(ws["G3"].value)
    assert status.startswith("=")
    assert "LeadTimeDays" in status
    assert "SafetyStockDays" in status
    reorder = str(ws["F3"].value)
    assert "LeadTimeDays" in reorder


def test_sku_margin_units_and_revenue_are_sumifs(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["SKU Margin"]
    assert str(ws["C3"].value).startswith("=SUMIFS")
    assert str(ws["D3"].value).startswith("=SUMIFS")
    assert str(ws["E3"].value).startswith("=SUMIFS")


def test_no_static_numbers_in_analysis_formula_cells(tmp_path):
    """Analysis cells that must be formulas must not be plain numbers."""
    _, wb = make_report(tmp_path)

    # SKU Margin data rows: cols C–L must be formulas
    sm = wb["SKU Margin"]
    row = 3
    while sm.cell(row=row, column=1).value and sm.cell(row=row, column=1).value != "TOTAL":
        for col in range(3, 13):
            val = sm.cell(row=row, column=col).value
            assert isinstance(val, str) and val.startswith("="), (
                f"SKU Margin R{row}C{col} expected formula, got {val!r}"
            )
        row += 1

    # Inventory data: cols C–G formulas
    inv = wb["Inventory"]
    row = 3
    while inv.cell(row=row, column=1).value:
        for col in range(3, 8):
            val = inv.cell(row=row, column=col).value
            assert isinstance(val, str) and val.startswith("="), (
                f"Inventory R{row}C{col} expected formula, got {val!r}"
            )
        row += 1
