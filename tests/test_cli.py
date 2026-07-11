"""CLI tests — generate and validate against fixture source."""

from __future__ import annotations

from pathlib import Path

from click.testing import CliRunner

from ledgerly.cli import main


def test_generate_fixture_exits_zero_and_writes_file(tmp_path):
    runner = CliRunner()
    out = tmp_path / "report.xlsx"
    result = runner.invoke(
        main,
        ["generate", "--month", "2026-06", "--source", "fixture", "--out", str(out)],
    )
    assert result.exit_code == 0, result.output
    assert out.exists()
    assert out.stat().st_size > 0


def test_validate_passes_on_generated_report(tmp_path):
    runner = CliRunner()
    out = tmp_path / "report.xlsx"
    gen = runner.invoke(
        main,
        ["generate", "--month", "2026-06", "--source", "fixture", "--out", str(out)],
    )
    assert gen.exit_code == 0, gen.output
    val = runner.invoke(main, ["validate", str(out)])
    assert val.exit_code == 0, val.output
    assert "VALIDATION PASSED" in val.output


def test_generate_no_template_full_workbook(tmp_path):
    runner = CliRunner()
    out = tmp_path / "full.xlsx"
    result = runner.invoke(
        main,
        [
            "generate",
            "--month",
            "2026-06",
            "--source",
            "fixture",
            "--out",
            str(out),
            "--no-template",
        ],
    )
    assert result.exit_code == 0, result.output
    assert out.exists()
    val = runner.invoke(main, ["validate", str(out)])
    assert val.exit_code == 0, val.output
    assert "SKU Margin" in val.output or "VALIDATION PASSED" in val.output


def test_shopify_source_errors_without_env(tmp_path):
    runner = CliRunner(env={"SHOPIFY_STORE": "", "SHOPIFY_ACCESS_TOKEN": ""})
    # Ensure env vars are absent
    result = runner.invoke(
        main,
        ["generate", "--month", "2026-06", "--source", "shopify", "--out", str(tmp_path / "x.xlsx")],
        env={"SHOPIFY_STORE": None, "SHOPIFY_ACCESS_TOKEN": None},
    )
    # click may still see parent env; force clear
    import os

    env = os.environ.copy()
    env.pop("SHOPIFY_STORE", None)
    env.pop("SHOPIFY_ACCESS_TOKEN", None)
    result = runner.invoke(
        main,
        ["generate", "--month", "2026-06", "--source", "shopify", "--out", str(tmp_path / "x.xlsx")],
        env=env,
    )
    assert result.exit_code != 0
    assert "SHOPIFY_STORE" in result.output or "SHOPIFY_ACCESS_TOKEN" in result.output
