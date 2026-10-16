"""Power BI export stage: turn the cleaned pipeline output into a folder of
flat CSV files arranged as a star schema, ready to load into Power BI
Desktop via the PBIP project in ``powerbi/``.

Run from the project root:

    python -m src.powerbi [workbook.xlsx] [output_dir]

Writes ``output_dir`` (default ``powerbi/data``):

    Sales.csv           one row per transaction, with Revenue and flags
    Products.csv        dim: product -> category + catalog price
    Salespersons.csv    dim: salesperson
    PaymentMethods.csv  dim: payment method
    CustomerTypes.csv   dim: customer type
    Date.csv            dim: one row per calendar day in range
    CleaningSummary.csv one row of cleaning counts (+ source file, run stamp)
    FlaggedRows.csv     the subset of Sales rows flagged for review

All CSVs are UTF-8 with a BOM so Power BI reads them without an encoding
step, and dates are ISO-8601. Daily charts use DateKey (time-of-day is
ignored, per the cleaning policy). Aggregation happens inside Power BI via
DAX measures — the CSVs stay row-level so the model can slice freely.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from src.report import SalesReport, build_report

PROJECT_ROOT = Path(__file__).resolve().parents[1]

#: Numeric/text types Power BI will infer correctly without help.
DATE_FMT = "%Y-%m-%d"


def _date_dimension(min_date: pd.Timestamp, max_date: pd.Timestamp) -> pd.DataFrame:
    """One row per calendar day covering the sales period."""
    days = pd.date_range(min_date.normalize(), max_date.normalize(), freq="D")
    return pd.DataFrame(
        {
            "Date": days.strftime(DATE_FMT),
            "Year": days.year,
            "Month": days.month,
            "MonthName": days.strftime("%b"),
            "Day": days.day,
            "WeekdayName": days.strftime("%a"),
        }
    )


def product_dimension(clean: pd.DataFrame) -> pd.DataFrame:
    """Product -> Category and catalog price (modal price, lowest wins ties)."""
    catalog = (
        clean.groupby("Product", dropna=False)
        .agg(
            Category=("Category", lambda s: s.mode().iloc[0] if len(s.mode()) else ""),
            Catalog_Price=("Unit_Price", "median"),
        )
        .reset_index()
    )
    return catalog


def sales_fact(clean: pd.DataFrame) -> pd.DataFrame:
    """Row-level sales with derived Revenue, pre-formatted for CSV."""
    work = clean.copy()
    work["Revenue"] = (
        work["Quantity"] * work["Unit_Price"] * (1.0 - work["Discount"])
    ).round(2)
    work["Date"] = pd.to_datetime(work["Date"])
    work["Is_Flagged"] = (
        work["Review_Flags"].fillna("").str.contains("review", na=False)
    )
    # Discount as a percentage is friendlier for Power BI formatting rules.
    work["Discount_Pct"] = (work["Discount"].astype(float) * 100.0).round(2)
    return work


def write_csvs(report: SalesReport, output_dir: Path) -> list[Path]:
    """Write every star-schema CSV; returns the files written."""
    output_dir.mkdir(parents=True, exist_ok=True)

    clean = report.rows.copy()
    clean["Date"] = pd.to_datetime(clean["Date"])

    sales = sales_fact(clean)
    dims = {
        "Products.csv": product_dimension(clean),
        "Salespersons.csv": pd.DataFrame(
            {"Salesperson": sorted(clean["Salesperson"].dropna().unique())}
        ),
        "PaymentMethods.csv": pd.DataFrame(
            {"Payment_Method": sorted(clean["Payment_Method"].dropna().unique())}
        ),
        "CustomerTypes.csv": pd.DataFrame(
            {"Customer_Type": sorted(clean["Customer_Type"].dropna().unique())}
        ),
        "Date.csv": _date_dimension(
            clean["Date"].min(), clean["Date"].max()
        ),
    }

    sales = sales.rename(columns={"Date": "DateKey"})
    sales["DateKey"] = pd.to_datetime(sales["DateKey"]).dt.strftime(DATE_FMT)

    flags = clean["Review_Flags"].fillna("")
    summary = pd.DataFrame(
        [
            {
                key: report.cleaning_summary.get(key)
                for key in (
                    "rows_in",
                    "duplicates_removed",
                    "rows_dropped",
                    "rows_out",
                    "flagged_for_review",
                )
            }
            | {
                "prices_imputed": int(flags.str.contains("imputed_price", na=False).sum()),
                "discounts_imputed": int(flags.str.contains("imputed_discount", na=False).sum()),
                "customer_types_imputed": int(flags.str.contains("imputed_customer_type", na=False).sum()),
            }
        ]
    )

    flagged = sales[sales["Is_Flagged"]].drop(columns=["Is_Flagged"])

    written: list[Path] = []
    frames: dict[str, pd.DataFrame] = {
        "Sales.csv": sales,
        "CleaningSummary.csv": summary,
        "FlaggedRows.csv": flagged,
        **dims,
    }
    for name, frame in frames.items():
        target = output_dir / name
        frame.to_csv(target, index=False, encoding="utf-8-sig")
        written.append(target)
    return written


def main(argv: list[str]) -> int:
    """CLI: python -m src.powerbi [SOURCE.xlsx] [OUTPUT_DIR]"""
    source = Path(argv[1]) if len(argv) > 1 else PROJECT_ROOT / "brightmart_sales_data.xlsx"
    output_dir = Path(argv[2]) if len(argv) > 2 else PROJECT_ROOT / "powerbi" / "data"
    if not source.is_file():
        print(f"not a file: {source}", file=sys.stderr)
        return 2
    report = build_report(source)
    written = write_csvs(report, output_dir)
    totals = report.totals
    print(
        f"Power BI export -> {output_dir}\n"
        f"  {totals['orders']} sales rows, "
        f"revenue {totals['total_revenue']:,.2f} ETB, "
        f"{len(written)} files"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
