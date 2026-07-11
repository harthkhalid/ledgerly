"""Assert all eight defined names exist and are referenced downstream."""

from __future__ import annotations

from ledgerly.engine import names as N
from tests.conftest import make_report


def test_all_eight_defined_names_exist(tmp_path):
    _, wb = make_report(tmp_path)
    for name in N.ALL_CONTROL_NAMES:
        assert name in wb.defined_names, f"Missing defined name: {name}"


def test_defined_names_resolve_to_controls(tmp_path):
    _, wb = make_report(tmp_path)
    for name in N.ALL_CONTROL_NAMES:
        dn = wb.defined_names[name]
        # openpyxl stores attr_text like "Controls!$B$3"
        attr = dn.attr_text
        assert "Controls!" in attr, f"{name} should point at Controls, got {attr}"


def test_margin_floor_referenced_downstream(tmp_path):
    _, wb = make_report(tmp_path)
    hits = _count_formula_refs(wb, N.MARGIN_FLOOR_PCT)
    assert hits >= 3, f"MarginFloorPct referenced {hits} times, need >= 3"


def test_lead_time_referenced_downstream(tmp_path):
    _, wb = make_report(tmp_path)
    hits = _count_formula_refs(wb, N.LEAD_TIME_DAYS)
    assert hits >= 3, f"LeadTimeDays referenced {hits} times, need >= 3"


def test_payment_processing_referenced_downstream(tmp_path):
    _, wb = make_report(tmp_path)
    hits = _count_formula_refs(wb, N.PAYMENT_PROCESSING_PCT)
    assert hits >= 3, f"PaymentProcessingPct referenced {hits} times, need >= 3"


def _count_formula_refs(wb, name: str) -> int:
    count = 0
    for ws in wb.worksheets:
        if ws.title == "Controls":
            continue
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and name in cell.value:
                    count += 1
        # Also check conditional formatting formulas
        for cf_range in ws.conditional_formatting._cf_rules:
            for rule in ws.conditional_formatting[cf_range]:
                formulas = getattr(rule, "formula", None) or []
                for f in formulas:
                    if name in str(f):
                        count += 1
    return count
