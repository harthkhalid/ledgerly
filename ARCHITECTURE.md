# Ledgerly Architecture

## Adapter pattern

The Excel engine (`ledgerly.engine`) imports **zero** Shopify code. It consumes only pydantic models (`Order`, `Product`, `Variant`, `InventoryLevel`, `Customer`) produced by anything that satisfies the `DataSource` protocol in `ledgerly.adapters.base`:

- `fetch_orders(since, until)`
- `fetch_products()`
- `fetch_inventory()`

`FixtureSource` and `ShopifySource` are interchangeable adapters. The engine never branches on source type. This keeps the workbook builder testable offline and embeddable in any host that can supply the data model.

```
┌─────────────┐     ┌──────────────┐     ┌─────────────────┐
│ Shopify API │──▶  │ ShopifySource│──▶  │                 │
└─────────────┘     └──────────────┘     │  engine.workbook│──▶ .xlsx
┌─────────────┐     ┌──────────────┐     │  (formulas, CF, │
│ sample.json │──▶  │ FixtureSource│──▶  │   charts, names)│
└─────────────┘     └──────────────┘     └─────────────────┘
```

## Why openpyxl over pandas `.to_excel`

`DataFrame.to_excel` writes **dead values**. Ledgerly's acceptance criteria require that changing `MarginFloorPct` on Controls re-flags SKU Margin rows when the file is reopened in Excel — that only works if cells contain formula strings and conditional formatting rules that reference workbook-level defined names.

Pandas is allowed for internal shaping only. Every analysis tab is assembled cell-by-cell (or table-by-table) through openpyxl so formulas, styles, charts, and defined names survive.

## Template-preservation strategy

Template mode (`template_mode.py`) is the flagship path for branded deliverables:

1. Load `templates/brand_report.xlsx` with `openpyxl.load_workbook(path)` — **never** `read_only=True`, **never** `data_only=True`.
2. Inject only into designated named ranges and the `tblTemplateData` body.
3. Leave every other cell, style, formula, column width, and sheet untouched.
4. When appending rows, copy the last template row's style and extend the Excel Table `ref`.

### Pitfalls

| Pitfall | Why it hurts | Mitigation |
|---|---|---|
| `data_only=True` | Replaces formulas with cached values (or `None`), destroying the live graph | Explicit comment + default `load_workbook` call |
| Forgetting style copy on append | New rows look unstyled / break brand | `_copy_row_style` from last seed row |
| Not extending Table `ref` | Structured references and table formatting stop at the old range | Update `table.ref` after write |
| Silent missing names | Injection writes nowhere; Summary formulas break quietly | `TemplateInjectionError` listing every missing name |

`scripts/make_template.py` rebuilds the template programmatically so provenance is reproducible and the repo stays self-contained.

## Embedding the engine

Because the engine is a pure **data model in → Workbook out** module, any host can call it:

```python
from ledgerly.engine.workbook import build_workbook, save_workbook

wb = build_workbook(
    orders=orders,
    products=products,
    variants=variants,
    inventory=inventory,
    customers=customers,
    report_month="2026-06",
    all_orders=orders,
)
save_workbook(wb, "report.xlsx")
```

Suitable hosts: a cron/scheduler job, a thin web service that returns the `.xlsx` as a download, an Airflow/Prefect task, or the Click CLI in `ledgerly.cli`. The CLI is a thin adapter around the same functions — not a second implementation.

## Workbook formula graph

```
Controls (named ranges)
    │
    ├── SKU Margin  (ShippingCostPerOrder, PackagingCostPerOrder,
    │                PaymentProcessingPct, MonthlyAdSpend, MarginFloorPct)
    ├── Inventory   (LeadTimeDays, SafetyStockDays)
    └── Dashboard   (MarginFloorPct via COUNTIF)
         ▲
RawData / RawInventory (hidden) ── SUMIFS / SUMPRODUCT sources
```

Zero assumption literals appear outside Controls. Downstream tabs reference defined names only.
