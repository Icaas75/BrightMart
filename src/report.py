"""Assemble the full sales report from the raw spreadsheet.

Stage 4 of the pipeline. Runs loader -> cleaner -> metrics and packages the
result as the single object the dashboard consumes.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from src.cleaning import clean_sales
from src.loader import load_sales
from src.metrics import (
    daily_revenue,
    revenue_by_column,
    summary_totals,
    top_products,
)


@dataclass
class SalesReport:
    """Everything the dashboard needs, computed once from the raw file."""

    totals: dict
    cleaning_summary: dict
    daily: pd.DataFrame          # Date, Revenue, Orders, Units
    by_category: pd.DataFrame    # Category, Total_Revenue
    by_salesperson: pd.DataFrame
    by_payment: pd.DataFrame
    by_customer: pd.DataFrame
    products: pd.DataFrame       # top products by revenue
    rows: pd.DataFrame           # cleaned row-level data incl. Review_Flags
    flagged: pd.DataFrame        # rows needing human review

    @property
    def revenue_series(self) -> pd.Series:
        """Daily revenue as a {label: value} dict for charting."""
        return {
            d.strftime("%b %d"): round(float(v), 2)
            for d, v in zip(self.daily["Date"], self.daily["Revenue"])
        }


def build_report(path, sheet: str | int | None = None) -> SalesReport:
    """Load, clean, and aggregate the spreadsheet into a SalesReport."""
    raw = load_sales(path, sheet=sheet)
    clean, cleaning_summary = clean_sales(raw)

    flagged = clean[
        clean["Review_Flags"].fillna("").str.contains("review", na=False)
    ].copy()

    return SalesReport(
        totals=summary_totals(clean),
        cleaning_summary=cleaning_summary,
        daily=daily_revenue(clean),
        by_category=revenue_by_column(clean, "Category"),
        by_salesperson=revenue_by_column(clean, "Salesperson"),
        by_payment=revenue_by_column(clean, "Payment_Method"),
        by_customer=revenue_by_column(clean, "Customer_Type"),
        products=top_products(clean, n=10),
        rows=clean,
        flagged=flagged,
    )
