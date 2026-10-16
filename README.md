# BrightMart Sales Pipeline, Dashboard & Power BI

Turns the raw `brightmart_sales_data.xlsx` spreadsheet into cleaned data,
aggregated metrics, an interactive web dashboard, a formatted Excel
report, and a Power BI project. Python does the data prep; view the
result in Flask (localhost), Excel, or Power BI.

## Stack

| Stage | Module | Responsibility |
|---|---|---|
| 1. Load | `src/loader.py` | Read the .xlsx, validate columns, coerce types |
| 2. Clean | `src/cleaning.py` | Dedup, impute blanks, normalize casing, flag catalog conflicts |
| 3. Metrics | `src/metrics.py` | Revenue math, daily aggregation, breakdowns, top products |
| 4. Report | `src/report.py` | Assemble everything into one `SalesReport` object |
| 5. Export | `src/export.py` | Formatted multi-sheet Excel workbook (also a CLI) |
| 6. Dashboard | `src/dashboard.py` | Flask web app with KPIs, SVG charts, and Excel download |
| 7. Power BI export | `src/powerbi.py` | Star-schema CSVs for the semantic model |
| 8. Power BI project | `powerbi/` | PBIP project: semantic model (TMDL) + report (PBIR) |
| — | `src/verify.py` | Independent number verification (audit CLI) |

### Cleaning policy (agreed with the data owner)

- Exact duplicate rows are removed (keep first).
- Rows whose `Category`/`Unit_Price` disagrees with the product catalog are
  **flagged for review, not fixed** (`Review_Flags` = `review`).
- Blank `Unit_Price` → catalog price; blank `Discount` → 0; blank
  `Customer_Type` → global majority; blank `Date`/`Quantity` → row dropped.
- Text casing is normalized to the modal spelling; `Invoice_ID` is never
  re-cased. Every imputation is tagged in `Review_Flags`.
- Revenue is `Quantity * Unit_Price * (1 - Discount)` in ETB; daily charts
  ignore time-of-day.

## Setup

```bash
pip install -r requirements.txt
```

## Run the pipeline (refresh the Power BI data)

```bash
python -m src.powerbi brightmart_sales_data.xlsx
```

Writes the star-schema CSVs into `powerbi/data/`:

| File | Role | Columns |
|---|---|---|
| `Sales.csv` | fact, 1 row per transaction | DateKey, Invoice_ID, Salesperson, Product, Category, Quantity, Unit_Price, Discount, Payment_Method, Customer_Type, Review_Flags, Revenue, Is_Flagged, Discount_Pct |
| `Date.csv` | dim: one row per calendar day | Date, Year, Month, MonthName, Day, WeekdayName |
| `Products.csv` | dim | Product, Category, Catalog_Price |
| `Salespersons.csv` | dim | Salesperson |
| `PaymentMethods.csv` | dim | Payment_Method |
| `CustomerTypes.csv` | dim | Customer_Type |
| `CleaningSummary.csv` | 1 row of cleaning counts | rows_in, duplicates_removed, rows_dropped, rows_out, flagged_for_review, prices_imputed, discounts_imputed, customer_types_imputed |
| `FlaggedRows.csv` | sales rows needing review | same as Sales minus Is_Flagged |

All CSVs are UTF-8 with a BOM; dates are ISO-8601. Aggregation happens in
DAX measures, so the fact table stays row-level and slices freely.

A different month's workbook:

```bash
python -m src.powerbi my_month_file.xlsx
```

## Run the dashboard

```bash
python -m src.dashboard
# then open http://127.0.0.1:5000
```

**Input-driven flow:** the landing page asks for a month's sales workbook.
Drop an `.xlsx` (or click to choose) and press **Generate report** — the file
is validated first (required columns, at least one data row, ≤ 25 MB) and the
dashboard is generated from it: KPI cards, daily revenue chart, top products,
breakdowns, cleaning summary, flagged rows, and a one-click **Excel download**
named after the source file. Uploading another month replaces the report;
invalid files are rejected with the reason shown. If
`brightmart_sales_data.xlsx` is present in the project root, a
**Use sample file instead** shortcut loads it without uploading.

Headless alternative (no browser), same numbers:

```bash
python -m src.export brightmart_sales_data.xlsx            # -> brightmart_sales_report.xlsx
python -m src.export brightmart_sales_data.xlsx out.xlsx   # custom output
```

The workbook contains: `Sales Dashboard`, `Summary`, `Daily Revenue`,
`Top Products`, `By Category`, `By Salesperson`, `By Payment Method`,
`By Customer Type`, `Cleaned Data`, `Flagged for Review`, and a hidden
`Chart Data` staging sheet, with native interactive Excel charts and
styled tables.

## Open the Power BI report

Double-click `powerbi/BrightMartSalesReport.pbip` (requires Power BI
Desktop). The project contains:

- **Semantic model** (`BrightMartSales.SemanticModel`) — the star schema:
  `Sales` fact joined (many-to-one, single direction) to `Date`, `Products`,
  `Salespersons`, `PaymentMethods`, and `CustomerTypes`; measures:

  | Measure | DAX |
  |---|---|
  | Total Revenue | `SUM(Sales[Revenue])` |
  | Orders | `COUNTROWS(Sales)` |
  | Units Sold | `SUM(Sales[Quantity])` |
  | Average Order Value | `DIVIDE([Total Revenue], [Orders], 0)` |
  | Flagged Rows | `CALCULATE(COUNTROWS(Sales), Sales[Is_Flagged] = TRUE())` |
  | Flagged Revenue | `CALCULATE([Total Revenue], Sales[Is_Flagged] = TRUE())` |

- **Report** (`BrightMartSales.Report`) — one *Sales Overview* page with
  KPI cards (revenue, orders, units, AOV), daily revenue area chart, top
  products, and breakdowns by category, salesperson, and payment method.

**Monthly workflow:** run `python -m src.powerbi <new month>.xlsx`, then hit
**Refresh** in Power BI Desktop — the visuals re-point at the fresh CSVs.

**Model tweaks:** edit tables/measures/relationships in
`powerbi/BrightMartSales.SemanticModel/definition/*.tmdl`, page layout in
`powerbi/BrightMartSales.Report/definition/report.json`, or regenerate both
with `python scripts/gen_powerbi.py` (re-run it if you move the project —
the CSV folder path is baked in at generation time).

## Verify the numbers

```bash
python -m src.powerbi brightmart_sales_data.xlsx   # refresh CSVs first
python -m src.verify brightmart_sales_data.xlsx
```

Recomputes the headline figures straight from a workbook — without
reusing the pipeline's own aggregation code — and reconciles them against
the pipeline, the Power BI CSV export, the exported Excel file, and the
rendered dashboard:

- manual per-invoice spot checks (raw cells vs pipeline output)
- total revenue vs the transaction-level sum
- every breakdown table cross-foots to the total
- duplicates excluded exactly once (raw duplicates == summary count)
- flagged rows consistent in every summary, revenue still included
- the Power BI CSV export matches the report (rows, revenue, flags)
- Excel Summary/daily/cleaned/flagged tables match the report
- dashboard page and `/export.xlsx` match the report

Exit code is 0 only if every check passes, so it can run in CI.

## Tests

```bash
python -m pytest
```

The real-file integration tests are skipped automatically if
`brightmart_sales_data.xlsx` is not present in the project root.
