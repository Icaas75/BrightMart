"""Calculate sales metrics from cleaned data.

Stage 3 of the pipeline. All monetary figures are in ETB and

    revenue = Quantity * Unit_Price * (1 - Discount)

with Discount expressed as a fraction (0.05 = 5%). Daily aggregations group
on the calendar day only, ignoring the time-of-day part of Date.
"""

from __future__ import annotations

import pandas as pd


def compute_revenue(df: pd.DataFrame) -> pd.Series:
    """Per-row revenue = Quantity * Unit_Price * (1 - Discount), in ETB."""
    return df["Quantity"] * df["Unit_Price"] * (1.0 - df["Discount"])


def add_revenue(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of ``df`` with a Revenue column (ETB)."""
    out = df.copy()
    out["Revenue"] = compute_revenue(out)
    return out


def daily_revenue(df: pd.DataFrame) -> pd.DataFrame:
    """Revenue by calendar day (time-of-day ignored), oldest first.

    Computes revenue itself, so ``df`` does not need a Revenue column.
    Returns a frame with columns: Date (midnight timestamps), Revenue,
    Orders (number of sales rows), Units (sum of Quantity).
    """
    work = add_revenue(df)
    work["Day"] = work["Date"].dt.normalize()
    grouped = work.groupby("Day").agg(
        Revenue=("Revenue", "sum"),
        Orders=("Revenue", "size"),
        Units=("Quantity", "sum"),
    )
    return grouped.sort_index().reset_index().rename(columns={"Day": "Date"})


def summary_totals(df: pd.DataFrame) -> dict:
    """Headline KPIs for the whole cleaned dataset."""
    revenue = compute_revenue(df)
    return {
        "total_revenue": float(revenue.sum()),
        "orders": int(len(df)),
        "units": int(df["Quantity"].sum()),
        "avg_order_value": float(revenue.mean()) if len(df) else 0.0,
        "start_date": df["Date"].min(),
        "end_date": df["Date"].max(),
    }


def revenue_by_column(df: pd.DataFrame, column: str) -> pd.DataFrame:
    """Total revenue grouped by a categorical column, highest first."""
    work = add_revenue(df)
    grouped = work.groupby(column, dropna=False)["Revenue"].sum()
    return (
        grouped.sort_values(ascending=False)
        .reset_index()
        .rename(columns={"Revenue": "Total_Revenue"})
    )


def top_products(df: pd.DataFrame, n: int = 5) -> pd.DataFrame:
    """Top ``n`` products by revenue, with units sold."""
    work = add_revenue(df)
    grouped = work.groupby("Product").agg(
        Total_Revenue=("Revenue", "sum"), Units=("Quantity", "sum")
    )
    return grouped.sort_values("Total_Revenue", ascending=False).head(n).reset_index()
