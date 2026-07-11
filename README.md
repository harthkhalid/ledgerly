# Ledgerly

Ledgerly is an internal operations tool for a DTC fragrance brand. It pulls order, product, and inventory data (from Shopify or a bundled fixture) and generates a fully editable multi-tab Excel workbook — live formulas, named ranges, conditional formatting, and native charts — so a business owner can change assumptions on the Controls tab and watch every downstream tab recalculate in Excel.

## Quickstart (no credentials)

```bash
pip install -e ".[dev]"
python scripts/make_template.py          # builds templates/brand_report.xlsx
ledgerly generate --month 2026-06 --source fixture --no-template --out report.xlsx
ledgerly validate report.xlsx
```

Open `report.xlsx` in Excel, change **Margin Floor %** on the Controls tab from 35% to 50%, and watch SKU Margin rows re-flag. Change **Lead Time (Days)** and Inventory statuses/reorder dates update live.

Template (branded) mode is the default when `--no-template` is omitted:

```bash
ledgerly generate --month 2026-06 --source fixture --out branded.xlsx
```

### Shopify source

```bash
set SHOPIFY_STORE=your-store
set SHOPIFY_ACCESS_TOKEN=shpat_...
ledgerly generate --month 2026-06 --source shopify --no-template --out report.xlsx
```

## Features

| Feature | Detail |
|---|---|
| Live Excel formulas | SUMIFS, SUMPRODUCT, IFERROR, COUNTIF — not static dumps |
| Named ranges | Eight Controls assumptions referenced by name everywhere |
| Conditional formatting | Margin-floor FormulaRule, inventory CellIsRules, cohort ColorScale |
| Native charts | LineChart, BarChart, PieChart on Dashboard (no images) |
| Template mode | Injects into pre-styled `brand_report.xlsx` without touching other cells |
| Fixture source | ~14 products, ~36 variants, ~900 orders, ~500 customers — zero network |
| Adapter pattern | Engine imports zero Shopify code; swap any DataSource |

## Tabs

1. **Controls** — editable assumptions (shipping, packaging, fees, ad spend, margin floor, lead time, safety stock, report month)
2. **RawData** / **RawInventory** — hidden flat tables + Excel Table `tblLines`
3. **SKU Margin** — contribution margin with allocated costs
4. **Inventory** — days of cover, reorder-by date, REORDER NOW / SOON / OK
5. **Cohort LTV** — acquisition-month triangle + 30/60/90-day averages
6. **Dashboard** — KPIs + three native charts

## Screenshots

<img width="895" height="642" alt="image" src="https://github.com/user-attachments/assets/8698152d-c83e-4744-ac15-6800f338970b" />


<img width="1346" height="786" alt="image" src="https://github.com/user-attachments/assets/b06a24c8-611b-4e93-aa0e-bb8c107dc475" />


<img width="1522" height="457" alt="image" src="https://github.com/user-attachments/assets/061ddf79-b02e-4409-9e52-45096eac1869" />


## Tests

```bash
pytest -v
```

## License

Internal use.
