"""Assert Dashboard embeds exactly one LineChart, BarChart, and PieChart."""

from __future__ import annotations

from openpyxl.chart import BarChart, LineChart, PieChart

from tests.conftest import make_report


def test_dashboard_has_three_chart_types(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["Dashboard"]
    charts = ws._charts
    lines = [c for c in charts if isinstance(c, LineChart)]
    bars = [c for c in charts if isinstance(c, BarChart)]
    pies = [c for c in charts if isinstance(c, PieChart)]
    assert len(lines) == 1, f"Expected 1 LineChart, got {len(lines)}"
    assert len(bars) == 1, f"Expected 1 BarChart, got {len(bars)}"
    assert len(pies) == 1, f"Expected 1 PieChart, got {len(pies)}"


def test_each_chart_has_title(tmp_path):
    _, wb = make_report(tmp_path)
    ws = wb["Dashboard"]
    for chart in ws._charts:
        title = chart.title
        assert title is not None, f"Chart {type(chart).__name__} missing title"
        # title may be a str or Title object
        text = title if isinstance(title, str) else str(title)
        assert len(text) > 0
