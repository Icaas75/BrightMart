"""Tests for src.cleaning: catalog building, dedup, blank policy, review flags."""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from src.cleaning import (
    DEFAULT_DISCOUNT,
    FLAG_IMPUTED_CUSTOMER,
    FLAG_IMPUTED_DISCOUNT,
    FLAG_IMPUTED_PRICE,
    FLAG_REVIEW,
    clean_sales,
)
from src.loader import REQUIRED_COLUMNS, load_sales

HEADER = REQUIRED_COLUMNS


def _frame(rows: list[dict]) -> pd.DataFrame:
    """Build a loader-shaped frame from dict rows."""
    return pd.DataFrame(rows, columns=HEADER).astype({"Date": "datetime64[ns]"})


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


# ------------------------------------------------------------------ pipeline


def test_clean_real_file_counts(real_sales_file):
    df, summary = clean_sales(load_sales(real_sales_file))
    # 505 in, 5 exact duplicates and 2 unrepairable blanks dropped -> 498
    assert summary["rows_in"] == 505
    assert summary["duplicates_removed"] == 5
    assert summary["rows_dropped"] == 2
    assert summary["rows_out"] == 498
    assert summary["flagged_for_review"] == 2


def test_real_file_imputations_are_flagged(real_sales_file):
    df, summary = clean_sales(load_sales(real_sales_file))
    flags = df["Review_Flags"].str.split(" ")
    assert FLAG_IMPUTED_PRICE in {f for fs in flags for f in fs}
    assert FLAG_IMPUTED_DISCOUNT in {f for fs in flags for f in fs}
    assert FLAG_IMPUTED_CUSTOMER in {f for fs in flags for f in fs}
    # no NaNs remain anywhere
    assert not df.drop(columns="Review_Flags").isna().any().any()


def test_case_variants_are_normalized(real_sales_file):
    df, _ = clean_sales(load_sales(real_sales_file))
    products = set(df["Product"])
    assert "iPhone 15" in products and "iphone 15" not in products
    assert "Samsung Galaxy A55" in products and "Samsung galaxy A55" not in products
    assert "sara" not in set(df["Salesperson"]) and "Sara" in set(df["Salesperson"])
    assert "Mobile money" not in set(df["Payment_Method"])


# ------------------------------------------------------- duplicates policy


def test_exact_duplicates_removed_keep_first():
    r1, r2 = _row(Invoice_ID="INV-1"), _row(Invoice_ID="INV-1")
    df, summary = clean_sales(_frame([r1, r2]))
    assert len(df) == 1
    assert summary["duplicates_removed"] == 1


def test_differing_rows_with_same_id_kept():
    # Same Invoice_ID but different content is NOT an exact duplicate.
    r1 = _row(Invoice_ID="INV-1", Quantity=1)
    r2 = _row(Invoice_ID="INV-1", Quantity=3)
    df, summary = clean_sales(_frame([r1, r2]))
    assert len(df) == 2
    assert summary["duplicates_removed"] == 0


# --------------------------------------------------- review flag policy


def test_conflicting_price_is_flagged_not_fixed():
    # 'usb-c cable' normally costs 900; this row says 500 -> flagged, kept.
    rows = [
        _row(Invoice_ID="INV-1"),
        _row(Invoice_ID="INV-2", Unit_Price=500),
    ]
    df, summary = clean_sales(_frame(rows))
    assert len(df) == 2
    assert summary["flagged_for_review"] == 1
    assert df.loc[1, "Unit_Price"] == 500  # untouched
    assert FLAG_REVIEW in df.loc[1, "Review_Flags"].split()
    assert FLAG_REVIEW not in df.loc[0, "Review_Flags"].split()


def test_conflicting_category_is_flagged_not_fixed():
    rows = [
        _row(Invoice_ID="INV-1"),
        _row(Invoice_ID="INV-2", Category="Phones"),
    ]
    df, summary = clean_sales(_frame(rows))
    assert df.loc[1, "Category"] == "Phones"  # untouched
    assert FLAG_REVIEW in df.loc[1, "Review_Flags"].split()


def test_conflicting_row_can_also_be_imputed():
    # Conflict AND blank discount -> row gets both flags.
    rows = [
        _row(Invoice_ID="INV-1"),
        _row(Invoice_ID="INV-2", Unit_Price=500, Discount=None),
    ]
    df, _ = clean_sales(_frame(rows))
    flags = df.loc[1, "Review_Flags"].split()
    assert FLAG_REVIEW in flags and FLAG_IMPUTED_DISCOUNT in flags


def test_casing_alone_is_not_a_conflict():
    rows = [
        _row(Invoice_ID="INV-1", Product="usb-c cable"),
        _row(Invoice_ID="INV-2", Product="USB-C Cable"),
    ]
    df, summary = clean_sales(_frame(rows))
    assert summary["flagged_for_review"] == 0
    assert set(df["Product"]) == {"USB-C Cable"}


# ---------------------------------------------------------- blanks policy


def test_blank_price_imputed_from_catalog():
    rows = [
        _row(Invoice_ID="INV-1"),  # establishes catalog price 900
        _row(Invoice_ID="INV-2", Unit_Price=None),
    ]
    df, _ = clean_sales(_frame(rows))
    assert df.loc[1, "Unit_Price"] == 900
    assert FLAG_IMPUTED_PRICE in df.loc[1, "Review_Flags"].split()


def test_blank_discount_becomes_zero():
    df, _ = clean_sales(_frame([_row(Discount=None)]))
    assert df.loc[0, "Discount"] == DEFAULT_DISCOUNT
    assert FLAG_IMPUTED_DISCOUNT in df.loc[0, "Review_Flags"].split()


def test_blank_customer_type_gets_majority():
    rows = [
        _row(Invoice_ID="INV-1", Customer_Type="Individual"),
        _row(Invoice_ID="INV-2", Customer_Type="Individual"),
        _row(Invoice_ID="INV-3", Customer_Type="Business"),
        _row(Invoice_ID="INV-4", Customer_Type=None),
    ]
    df, _ = clean_sales(_frame(rows))
    assert df.loc[3, "Customer_Type"] == "Individual"
    assert FLAG_IMPUTED_CUSTOMER in df.loc[3, "Review_Flags"].split()


def test_blank_date_row_is_dropped():
    df, summary = clean_sales(_frame([_row(Invoice_ID="INV-1", Date=None)]))
    assert len(df) == 0
    assert summary["rows_dropped"] == 1


def test_blank_quantity_row_is_dropped():
    df, summary = clean_sales(_frame([_row(Invoice_ID="INV-1", Quantity=None)]))
    assert len(df) == 0
    assert summary["rows_dropped"] == 1


def test_summary_counts_are_consistent():
    rows = [
        _row(Invoice_ID="INV-1"),
        _row(Invoice_ID="INV-1"),                      # duplicate
        _row(Invoice_ID="INV-2", Date=None),           # dropped
        _row(Invoice_ID="INV-3", Unit_Price=None),     # imputed
    ]
    _, summary = clean_sales(_frame(rows))
    assert summary["rows_in"] == 4
    assert summary["rows_out"] == 2
    assert summary["rows_in"] - summary["duplicates_removed"] - summary["rows_dropped"] \
        == summary["rows_out"]
