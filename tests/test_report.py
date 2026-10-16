"""Tests for src.report."""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from src.cleaning import clean_sales
from src.loader import REQUIRED_COLUMNS, load_sales
from src.metrics import add_revenue, summary_totals
from src.report import build_report


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


def test_report_from_real_file(real_sales_file):
    report = build_report(real_sales_file)

    # cleaning summary matches the agreed policy outcomes
    assert report.cleaning_summary["rows_in"] == 505
    assert report.cleaning_summary["rows_out"] == 498
    assert report.cleaning_summary["duplicates_removed"] == 5
    assert report.cleaning_summary["flagged_for_review"] == 2

    # totals are internally consistent
    assert report.totals["orders"] == 498
    assert report.daily["Revenue"].sum() == pytest.approx(report.totals["total_revenue"])

    # breakdown tables are non-empty and sorted
    assert len(report.by_category) == 4
    assert report.by_category["Total_Revenue"].is_monotonic_decreasing
    assert len(report.by_salesperson) == 5
    assert len(report.by_payment) == 4  # casing fixed Mobile Money
    assert set(report.by_customer["Customer_Type"]) == {"Individual", "Business"}
    assert len(report.products) == 10

    # exactly the two known conflicting rows are flagged for review
    assert len(report.flagged) == 2
    assert set(report.flagged["Product"]) == {"iPhone 15", "Samsung Galaxy A55"}


def test_report_daily_has_31_days(real_sales_file):
    report = build_report(real_sales_file)
    assert len(report.daily) == 31
    assert report.daily["Date"].is_monotonic_increasing


def test_revenue_series_property(real_sales_file):
    report = build_report(real_sales_file)
    series = report.revenue_series
    assert len(series) == 31
    assert all(isinstance(v, float) for v in series.values())


def test_flagged_rows_survive_cleaning(real_sales_file):
    """Flagged rows must appear in row-level data with the review tag."""
    report = build_report(real_sales_file)
    review_ids = set(report.flagged["Invoice_ID"])
    tagged = set(report.rows.loc[
        report.rows["Review_Flags"].fillna("").str.contains("review"), "Invoice_ID"
    ])
    assert review_ids == tagged and len(review_ids) == 2


def test_clean_rows_have_no_missing_values(real_sales_file):
    report = build_report(real_sales_file)
    values = report.rows.drop(columns="Review_Flags")
    assert not values.isna().any().any()


def test_empty_flagged_table_when_no_conflicts(make_workbook):
    rows = [
        REQUIRED_COLUMNS,
        list(_row().values()),
    ]
    report = build_report(make_workbook(rows))
    assert len(report.flagged) == 0
    assert report.cleaning_summary["flagged_for_review"] == 0
