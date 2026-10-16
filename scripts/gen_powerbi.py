"""Generate the BrightMart Power BI PBIP project (idempotent).

Written files:

    powerbi/BrightMartSalesReport/.pbip-settings.json
    powerbi/BrightMartSales.SemanticModel/.platform
    powerbi/BrightMartSales.SemanticModel/definition/model.tmdl
    powerbi/BrightMartSales.SemanticModel/definition/<Table>.tmdl  (per CSV)
    powerbi/BrightMartSales.Report/.platform
    powerbi/BrightMartSales.Report/definition/report.json          (PBIR)

The model binds to the CSVs produced by `python -m src.powerbi` in
powerbi/data via an absolute path recorded at generation time. Re-run
after moving the project:

    python scripts/gen_powerbi.py
"""

from __future__ import annotations

import json
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "powerbi" / "data"

PBI_ROOT = PROJECT_ROOT / "powerbi"
SEM_MODEL = PBI_ROOT / "BrightMartSales.SemanticModel"
REPORT = PBI_ROOT / "BrightMartSales.Report"

MODEL_NAME = "BrightMart Sales Model"
REPORT_NAME = "BrightMart Sales Report"

#: CSV filename -> (table name, [(column, M type)])
TABLES: dict[str, tuple[str, list[tuple[str, str]]]] = {
    "Sales.csv": (
        "Sales",
        [
            ("DateKey", "text"),
            ("Invoice_ID", "text"),
            ("Salesperson", "text"),
            ("Product", "text"),
            ("Category", "text"),
            ("Quantity", "number"),
            ("Unit_Price", "number"),
            ("Discount", "number"),
            ("Payment_Method", "text"),
            ("Customer_Type", "text"),
            ("Review_Flags", "text"),
            ("Revenue", "number"),
            ("Is_Flagged", "logical"),
            ("Discount_Pct", "number"),
        ],
    ),
    "Products.csv": (
        "Products",
        [
            ("Product", "text"),
            ("Category", "text"),
            ("Catalog_Price", "number"),
        ],
    ),
    "Salespersons.csv": ("Salespersons", [("Salesperson", "text")]),
    "PaymentMethods.csv": ("PaymentMethods", [("Payment_Method", "text")]),
    "CustomerTypes.csv": ("CustomerTypes", [("Customer_Type", "text")]),
    "Date.csv": (
        "Date",
        [
            ("Date", "date"),
            ("Year", "Int64.Type"),
            ("Month", "Int64.Type"),
            ("MonthName", "text"),
            ("Day", "Int64.Type"),
            ("WeekdayName", "text"),
        ],
    ),
    "CleaningSummary.csv": (
        "CleaningSummary",
        [
            ("rows_in", "Int64.Type"),
            ("duplicates_removed", "Int64.Type"),
            ("rows_dropped", "Int64.Type"),
            ("rows_out", "Int64.Type"),
            ("flagged_for_review", "Int64.Type"),
            ("prices_imputed", "Int64.Type"),
            ("discounts_imputed", "Int64.Type"),
            ("customer_types_imputed", "Int64.Type"),
        ],
    ),
    "FlaggedRows.csv": (
        "FlaggedRows",
        [
            ("DateKey", "text"),
            ("Invoice_ID", "text"),
            ("Salesperson", "text"),
            ("Product", "text"),
            ("Category", "text"),
            ("Quantity", "number"),
            ("Unit_Price", "number"),
            ("Discount", "number"),
            ("Payment_Method", "text"),
            ("Customer_Type", "text"),
            ("Review_Flags", "text"),
            ("Revenue", "number"),
            ("Discount_Pct", "number"),
        ],
    ),
}

#: (fact table, fact column, dim table, dim column)
RELATIONSHIPS = [
    ("Sales", "DateKey", "Date", "Date"),
    ("Sales", "Product", "Products", "Product"),
    ("Sales", "Salesperson", "Salespersons", "Salesperson"),
    ("Sales", "Payment_Method", "PaymentMethods", "Payment_Method"),
    ("Sales", "Customer_Type", "CustomerTypes", "Customer_Type"),
]

#: (measure name, DAX, format string) — homepage KPIs first
MEASURES: list[tuple[str, str, str]] = [
    ("Total Revenue", "SUM(Sales[Revenue])", "#,0.00"),
    ("Orders", "COUNTROWS(Sales)", "#,0"),
    ("Units Sold", "SUM(Sales[Quantity])", "#,0"),
    ("Average Order Value", "DIVIDE([Total Revenue], [Orders], 0)", "#,0.00"),
    ("Flagged Rows", "CALCULATE(COUNTROWS(Sales), Sales[Is_Flagged] = TRUE())", "#,0"),
    ("Flagged Revenue", "CALCULATE([Total Revenue], Sales[Is_Flagged] = TRUE())", "#,0.00"),
]

DATE_PARAM_NAME = "SharedDataFolder"


def m_query_for(csv_name: str) -> str:
    """M query binding one table to its CSV via the folder parameter."""
    _, columns = TABLES[csv_name]
    pairs = ",\n".join(
        f'{{"{col}", {mtype}}}' for col, mtype in columns
    )
    return (
        "let\n"
        f'    Source = Folder.Files({DATE_PARAM_NAME}),\n'
        f'    Row = Table.SelectRows(Source, each [Name] = "{csv_name}"){{0}},\n'
        "    Csv = Csv.Document(Row[Content], [Delimiter = \",\", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),\n"
        "    Headers = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),\n"
        "    Typed = Table.TransformColumnTypes(Headers, {\n"
        f"{pairs}\n"
        "    })\n"
        "in\n"
        "    Typed"
    )


def model_tmdl(data_folder: str) -> str:
    lines: list[str] = [
        "model Model",
        "	culture en-US",
        "	defaultDataAccessOption Import",
        "",
        "database",
        "	compatibilityLevel 1560",
        "",
        "model model",
    ]
    for table, _ in TABLES.values():
        lines.append(f"	table {table}")
    lines.append("")
    for i, (fact, fact_col, dim, dim_col) in enumerate(RELATIONSHIPS):
        lines += [
            f"	relationship r{i}",
            f"		fromTable {fact}",
            f"		fromColumn {fact_col}",
            f"		toTable {dim}",
            f"		toColumn {dim_col}",
        ]
    lines += [
        "",
        f"expression {DATE_PARAM_NAME} = \"{data_folder}\" meta [IsParameterQuery = true, ListValue = true]",
        "",
    ]
    for measure, dax, fmt in MEASURES:
        lines += [
            f"	measure '{measure}' = {dax}",
            f"		formatString: {fmt}",
            "		lineageTag: measure_" + measure.replace(" ", "_"),
        ]
    lines += ["", "annotation PBIDesktopVersion = \"2.140.0.0\""]
    return "\n".join(lines) + "\n"


def table_tmdl(csv_name: str) -> str:
    table, columns = TABLES[csv_name]
    lines: list[str] = [f"table {table}", ""]
    m = m_query_for(csv_name)
    lines += [
        "	partition CSV = m",
        "		mode: import",
        "		query:",
        *(f"			{line}" for line in m.splitlines()),
        "",
    ]
    for col, mtype in columns:
        lines += [
            f"	column {col}",
            f"		dataType: {TMDL_TYPES[mtype]}",
            "		summarizeBy: none",
            f"		sourceColumn: {col}",
            "",
        ]
    return "\n".join(lines) + "\n"


TMDL_TYPES = {
    "text": "string",
    "number": "double",
    "logical": "boolean",
    "Int64.Type": "int64",
    "date": "dateTime",
}


def platform_json(display_type: str, display_name: str, logical_id: str) -> str:
    return json.dumps(
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": display_type, "displayName": display_name},
            "config": {"version": "2.0", "logicalId": logical_id},
        },
        indent=2,
    )


def card_visual(name: str, x: int, y: int, measure: str, label: str) -> dict:
    return {
        "name": name,
        "type": "VisualContainer",
        "position": {"x": x, "y": y, "width": 220, "height": 130, "z": 0},
        "visual": {
            "visualType": "card",
            "objects": {},
            "query": {
                "Values": {
                    "Aggregations": [
                        {
                            "Aggregation": {
                                "Expression": {
                                    "Measure": {
                                        "Expression": {"SourceRef": {"Schema": "m"}},
                                        "Reference": measure,
                                    }
                                }
                            }
                        }
                    ]
                }
            },
        },
        "title": label,
    }


def report_json() -> dict:
    kpis = [
        ("Total Revenue", "Total Revenue (ETB)"),
        ("Orders", "Orders"),
        ("Units Sold", "Units Sold"),
        ("Average Order Value", "Avg Order Value (ETB)"),
    ]
    cards = [
        card_visual(f"kpi{i}", 24 + i * 236, 24, measure, label)
        for i, (measure, label) in enumerate(kpis)
    ]
    visuals = cards + [
        {
            "name": "dailyRevenueArea",
            "type": "VisualContainer",
            "position": {"x": 24, "y": 180, "width": 690, "height": 330, "z": 0},
            "visual": {"visualType": "areaChart", "objects": {}},
            "title": "Daily Revenue",
            "projections": {
                "Category": [{"field": {"Column": {"Property": "DateKey"}}, "queryRef": "Sales.DateKey"}],
                "Y": [{"field": {"Measure": {"Property": "Total Revenue"}}, "queryRef": "Sales.Total Revenue"}],
            },
        },
        {
            "name": "topProductsBars",
            "type": "VisualContainer",
            "position": {"x": 730, "y": 180, "width": 460, "height": 330, "z": 0},
            "visual": {"visualType": "barChart", "objects": {}},
            "title": "Top Products by Revenue",
            "projections": {
                "Category": [{"field": {"Column": {"Property": "Product"}}, "queryRef": "Sales.Product"}],
                "Y": [{"field": {"Measure": {"Property": "Total Revenue"}}, "queryRef": "Sales.Total Revenue"}],
            },
        },
        {
            "name": "revenueByCategory",
            "type": "VisualContainer",
            "position": {"x": 24, "y": 526, "width": 350, "height": 300, "z": 0},
            "visual": {"visualType": "donutChart", "objects": {}},
            "title": "Revenue by Category",
            "projections": {
                "Category": [{"field": {"Column": {"Property": "Category"}}, "queryRef": "Sales.Category"}],
                "Values": [{"field": {"Measure": {"Property": "Total Revenue"}}, "queryRef": "Sales.Total Revenue"}],
            },
        },
        {
            "name": "revenueBySalesperson",
            "type": "VisualContainer",
            "position": {"x": 390, "y": 526, "width": 350, "height": 300, "z": 0},
            "visual": {"visualType": "barChart", "objects": {}},
            "title": "Revenue by Salesperson",
            "projections": {
                "Category": [{"field": {"Column": {"Property": "Salesperson"}}, "queryRef": "Sales.Salesperson"}],
                "Y": [{"field": {"Measure": {"Property": "Total Revenue"}}, "queryRef": "Sales.Total Revenue"}],
            },
        },
        {
            "name": "revenueByPaymentMethod",
            "type": "VisualContainer",
            "position": {"x": 756, "y": 526, "width": 350, "height": 300, "z": 0},
            "visual": {"visualType": "columnChart", "objects": {}},
            "title": "Revenue by Payment Method",
            "projections": {
                "Category": [{"field": {"Column": {"Property": "Payment_Method"}}, "queryRef": "Sales.Payment_Method"}],
                "Y": [{"field": {"Measure": {"Property": "Total Revenue"}}, "queryRef": "Sales.Total Revenue"}],
            },
        },
        {
            "name": "cleaningSummaryCard",
            "type": "VisualContainer",
            "position": {"x": 1122, "y": 526, "width": 230, "height": 300, "z": 0},
            "visual": {"visualType": "image", "objects": {}},
            "title": "Cleaning summary",
        },
    ]
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbir/reportSchema/1.0.0/schema.json",
        "settings": {
            "useStylizedVisualContainerHeaders": True,
            "defaultVisualStyles": {},
        },
        "themeCollection": {"baseTheme": {"name": "CY24SU10"}},
        "layout": {"pageSize": {"type": "Default"}},
        "pages": [
            {
                "name": "sales-overview",
                "displayName": "Sales Overview",
                "visualContainers": visuals,
            }
        ],
    }


def main() -> None:
    data_folder = str(DATA_DIR).replace("\\", "\\\\")

    pbip = PBI_ROOT / "BrightMartSalesReport"
    pbip.mkdir(parents=True, exist_ok=True)
    (PBI_ROOT / "BrightMartSalesReport.pbip").write_text(
        json.dumps(
            {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
                "version": "1.0",
                "artifacts": [
                    {"report": {"name": REPORT_NAME}},
                    {"semanticModel": {"name": MODEL_NAME}},
                ],
                "settings": {},
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    (SEM_MODEL / "definition").mkdir(parents=True, exist_ok=True)
    (SEM_MODEL / ".platform").write_text(
        platform_json("SemanticModel", MODEL_NAME, "brightmart-semantic-model"),
        encoding="utf-8",
    )
    (SEM_MODEL / "definition" / "model.tmdl").write_text(
        model_tmdl(data_folder), encoding="utf-8"
    )
    for csv_name in TABLES:
        table = TABLES[csv_name][0]
        (SEM_MODEL / "definition" / f"{table}.tmdl").write_text(
            table_tmdl(csv_name), encoding="utf-8"
        )

    (REPORT / "definition").mkdir(parents=True, exist_ok=True)
    (REPORT / ".platform").write_text(
        platform_json("Report", REPORT_NAME, "brightmart-report"),
        encoding="utf-8",
    )
    (REPORT / "definition" / "report.json").write_text(
        json.dumps(report_json(), indent=2), encoding="utf-8"
    )

    print(f"PBIP project written under {PBI_ROOT}")
    print(f"  data folder bound at {DATA_DIR}")
    print("Open in Power BI Desktop: powerbi/BrightMartSalesReport.pbip")


if __name__ == "__main__":
    main()
