"""Tests for src.loader."""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest

from src.loader import (
    REQUIRED_COLUMNS,
    EmptySheetError,
    MissingColumnsError,
    load_sales,
)

HEADER = REQUIRED_COLUMNS


def _row(**overrides) -> list:
    """A valid data row, with optional per-column overrides."""
    values = {
        "Date": dt.datetime(2026, 8, 1, 9, 30),
        "Invoice_ID": "INV-00001",
        "Salesperson": "Sara",
        "Product": "USB-C Cable",
        "Category": "Accessories",
        "Quantity": 2,
        "Unit_Price": 900,
        "Discount": 0,
        "Payment_Method": "Cash",
        "Customer_Type": "Individual",
    }
    values.update(overrides)
    return [values[col] for col in REQUIRED_COLUMNS]


# ---------------------------------------------------------------- happy path


def test_loads_real_spreadsheet(real_sales_file):
    df = load_sales(real_sales_file)
    assert list(df.columns) == REQUIRED_COLUMNS
    assert len(df) == 505
    assert pd.api.types.is_datetime64_any_dtype(df["Date"])
    for col in ("Quantity", "Unit_Price", "Discount"):
        assert pd.api.types.is_numeric_dtype(df[col])
    # dates in the real file land in August 2026
    assert df["Date"].min().year == 2026 and df["Date"].min().month == 8
    assert df["Date"].max().year == 2026 and df["Date"].max().month == 8


def test_extra_columns_are_dropped(make_workbook):
    rows = [HEADER + ["Extra"], _row() + ["ignore me"]]
    df = load_sales(make_workbook(rows))
    assert list(df.columns) == REQUIRED_COLUMNS
    assert len(df) == 1


def test_loading_named_sheet(make_workbook):
    rows = [HEADER, _row()]
    df = load_sales(make_workbook(rows), sheet="Sales Data")
    assert len(df) == 1


# ---------------------------------------------------------------- error cases


def test_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError, match="nope.xlsx"):
        load_sales(tmp_path / "nope.xlsx")


def test_header_only_sheet_raises(make_workbook):
    with pytest.raises(EmptySheetError):
        load_sales(make_workbook([HEADER]))


def test_completely_blank_sheet_raises(make_workbook):
    with pytest.raises(EmptySheetError):
        load_sales(make_workbook([]))


def test_missing_required_column_raises(make_workbook):
    bad_header = [c for c in HEADER if c != "Unit_Price"]
    path = make_workbook([bad_header, _row()[:-1]])
    with pytest.raises(MissingColumnsError, match="Unit_Price"):
        load_sales(path)


# ------------------------------------------------------------ type coercion


def test_excel_serial_dates_are_converted(make_workbook):
    # 46235 days after 1899-12-30 is 2026-08-01
    path = make_workbook([HEADER, _row(Date=46235.0)])
    df = load_sales(path)
    assert df.loc[0, "Date"] == pd.Timestamp("2026-08-01")


def test_datetime_dates_pass_through(make_workbook):
    path = make_workbook([HEADER, _row(Date=dt.datetime(2026, 8, 2, 14, 45))])
    df = load_sales(path)
    assert df.loc[0, "Date"] == pd.Timestamp("2026-08-02 14:45:00")


def test_unparseable_date_becomes_nat(make_workbook):
    path = make_workbook([HEADER, _row(Date="not-a-date")])
    df = load_sales(path)
    assert pd.isna(df.loc[0, "Date"])


def test_non_numeric_quantity_becomes_nan(make_workbook):
    path = make_workbook([HEADER, _row(Quantity="abc")])
    df = load_sales(path)
    assert pd.isna(df.loc[0, "Quantity"])


def test_text_columns_are_stripped(make_workbook):
    path = make_workbook(
        [
            HEADER,
            _row(Invoice_ID="  INV-9 ", Salesperson=" Sara ", Product="USB-C Cable ",
                 Customer_Type=" Individual "),
        ]
    )
    df = load_sales(path)
    assert df.loc[0, "Invoice_ID"] == "INV-9"
    assert df.loc[0, "Salesperson"] == "Sara"
    assert df.loc[0, "Product"] == "USB-C Cable"
    assert df.loc[0, "Customer_Type"] == "Individual"
