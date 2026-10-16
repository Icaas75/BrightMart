"""Tests for src.metrics."""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from src.cleaning import clean_sales
from src.loader import REQUIRED_COLUMNS, load_sales
from src.metrics import (
    add_revenue,
    compute_revenue,
    daily_revenue,
    revenue_by_column,
    summary_totals,
    top_products,
)


def _frame(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=REQUIRED_COLUMNS).astype({"Date": "datetime64[ns]"})


def _row(**over) -> dict:
    values = {
        "Date": dt.datetime(2026, 8, 1, 9, 0),
        "Invoice_ID": "INV-1",
        "Salesperson": "Sara",
        "Product": "USB-C Cable",
        "Category": "Accessories",
        "Quantity": 2,
        "Unit_Price": 900,
        "Discount": 0,
        "Payment_Method": "Cash",
        "Customer_Type": "Individual",
    }
    values.update(over)
    return values


# ------------------------------------------------------------- revenue math


def test_revenue_basic_formula():
    df = _frame([_row(Quantity=2, Unit_Price=900, Discount=0)])
    assert compute_revenue(df).iloc[0] == pytest.approx(1800.0)


def test_revenue_with_discount_fraction():
    # 3 * 68000 * (1 - 0.1) = 183600
    df = _frame([_row(Quantity=3, Unit_Price=68000, Discount=0.1)])
    assert compute_revenue(df).iloc[0] == pytest.approx(183_600.0)


def test_revenue_full_discount_is_zero():
    df = _frame([_row(Quantity=5, Unit_Price=900, Discount=1.0)])
    assert compute_revenue(df).iloc[0] == pytest.approx(0.0)


def test_add_revenue_returns_copy_with_column():
    df = _frame([_row()])
    out = add_revenue(df)
    assert "Revenue" not in df.columns  # original untouched
    assert out.loc[0, "Revenue"] == pytest.approx(1800.0)


# ------------------------------------------------------------ daily grouping


def test_daily_grouping_ignores_time_of_day():
    rows = [
        _row(Invoice_ID="INV-1", Date=dt.datetime(2026, 8, 1, 9, 30), Quantity=1,
             Unit_Price=900, Discount=0),
        _row(Invoice_ID="INV-2", Date=dt.datetime(2026, 8, 1, 23, 59), Quantity=2,
             Unit_Price=900, Discount=0),
        _row(Invoice_ID="INV-3", Date=dt.datetime(2026, 8, 2, 8, 0), Quantity=1,
             Unit_Price=900, Discount=0),
    ]
    daily = daily_revenue(add_revenue(_frame(rows)))
    assert len(daily) == 2
    assert daily.loc[0, "Date"] == pd.Timestamp("2026-08-01")
    assert daily.loc[0, "Revenue"] == pytest.approx(2700.0)
    assert daily.loc[0, "Orders"] == 2
    assert daily.loc[0, "Units"] == 3
    assert daily.loc[1, "Revenue"] == pytest.approx(900.0)
    # oldest day first
    assert daily["Date"].is_monotonic_increasing


# ----------------------------------------------------------------- totals


def test_summary_totals():
    rows = [
        _row(Invoice_ID="INV-1", Quantity=1, Unit_Price=1000, Discount=0.1),  # 900
        _row(Invoice_ID="INV-2", Quantity=2, Unit_Price=500, Discount=0),     # 1000
    ]
    totals = summary_totals(_frame(rows))
    assert totals["total_revenue"] == pytest.approx(1900.0)
    assert totals["orders"] == 2
    assert totals["units"] == 3
    assert totals["avg_order_value"] == pytest.approx(950.0)
    assert totals["start_date"] == pd.Timestamp("2026-08-01 09:00:00")


def test_summary_totals_empty_frame():
    totals = summary_totals(_frame([]))
    assert totals["total_revenue"] == 0.0
    assert totals["orders"] == 0
    assert totals["avg_order_value"] == 0.0


# -------------------------------------------------------------- breakdowns


def test_revenue_by_column_sorted_descending():
    rows = [
        _row(Invoice_ID="INV-1", Category="Phones", Quantity=1, Unit_Price=52000, Discount=0),
        _row(Invoice_ID="INV-2", Category="Accessories", Quantity=1, Unit_Price=900, Discount=0),
        _row(Invoice_ID="INV-3", Category="Accessories", Quantity=1, Unit_Price=900, Discount=0),
    ]
    by_cat = revenue_by_column(_frame(rows), "Category")
    assert list(by_cat["Category"]) == ["Phones", "Accessories"]
    assert by_cat.loc[0, "Total_Revenue"] == pytest.approx(52_000.0)
    assert by_cat.loc[1, "Total_Revenue"] == pytest.approx(1800.0)


def test_top_products_limits_and_sums():
    rows = [
        _row(Invoice_ID="INV-1", Product="USB-C Cable", Quantity=1, Unit_Price=900, Discount=0),
        _row(Invoice_ID="INV-2", Product="USB-C Cable", Quantity=2, Unit_Price=900, Discount=0),
        _row(Invoice_ID="INV-3", Product="iPhone 15", Quantity=1, Unit_Price=52000, Discount=0),
        _row(Invoice_ID="INV-4", Product="JBL Tune 510BT", Quantity=1, Unit_Price=4500, Discount=0),
    ]
    top = top_products(_frame(rows), n=2)
    assert list(top["Product"]) == ["iPhone 15", "JBL Tune 510BT"]
    assert top.loc[0, "Total_Revenue"] == pytest.approx(52_000.0)
    assert top.loc[1, "Total_Revenue"] == pytest.approx(4500.0)
    assert top.loc[1, "Units"] == 1


# --------------------------------------------------- real-file integration


def test_real_file_metrics_run(real_sales_file):
    df, _ = clean_sales(load_sales(real_sales_file))
    totals = summary_totals(df)
    daily = daily_revenue(add_revenue(df))

    # 498 clean rows over ~31 days in August 2026
    assert totals["orders"] == 498
    assert totals["start_date"].month == 8 and totals["start_date"].year == 2026
    assert len(daily) == 31
    assert daily["Revenue"].sum() == pytest.approx(totals["total_revenue"])
    assert (daily["Revenue"] > 0).all()
