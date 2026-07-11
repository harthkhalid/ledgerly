"""Assert conditional formatting rules are present and reference named ranges."""

from __future__ import annotations

from ledgerly.engine import names as N
from tests.conftest import make_report


def _iter_rules(ws):
    for cf_range in ws.conditional_formatting._cf_rules:
        for rule in ws.conditional_formatting[cf_range]:
            yield cf_range, rule


def test_sku_margin_formula_rule_references_margin_floor(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["SKU Margin"]
    # After reload, openpyxl yields Rule objects; FormulaRule is a factory.
    formula_rules = [rule for _, rule in _iter_rules(ws) if rule.type == "expression"]
    assert formula_rules, "SKU Margin should have a FormulaRule (type=expression)"
    found = False
    for rule in formula_rules:
        for f in rule.formula or []:
            if N.MARGIN_FLOOR_PCT in str(f):
                found = True
    assert found, "FormulaRule must reference MarginFloorPct by name"


def test_inventory_has_three_cell_is_rules(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["Inventory"]
    cell_is = [rule for _, rule in _iter_rules(ws) if rule.type == "cellIs"]
    assert len(cell_is) == 3, f"Expected 3 CellIsRules, got {len(cell_is)}"
    formulas = []
    for rule in cell_is:
        formulas.extend(str(f) for f in (rule.formula or []))
    joined = " ".join(formulas)
    assert "REORDER NOW" in joined
    assert "REORDER SOON" in joined
    assert "OK" in joined


def test_cohort_has_color_scale_rule(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["Cohort LTV"]
    scales = [rule for _, rule in _iter_rules(ws) if rule.type == "colorScale"]
    assert len(scales) >= 1, "Cohort LTV should have a ColorScaleRule"
