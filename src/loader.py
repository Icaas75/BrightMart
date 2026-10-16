"""Load the BrightMart sales spreadsheet into a tidy DataFrame.

Stage 1 of the pipeline. Responsibilities (and nothing more):
- locate and open the .xlsx file,
- validate that all required columns are present,
- coerce types: dates (Excel serials or datetimes) and numerics,
- strip whitespace from text columns.

Cleaning (duplicates, blanks, casing) happens in src/cleaning.py.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

#: Columns the rest of the pipeline depends on, in display order.
REQUIRED_COLUMNS = [
    "Date",
    "Invoice_ID",
    "Salesperson",
    "Product",
    "Category",
    "Quantity",
    "Unit_Price",
    "Discount",
    "Payment_Method",
    "Customer_Type",
]

TEXT_COLUMNS = [
    "Invoice_ID",
    "Salesperson",
    "Product",
    "Category",
    "Payment_Method",
    "Customer_Type",
]

NUMERIC_COLUMNS = ["Quantity", "Unit_Price", "Discount"]

# Excel stores dates as days since 1899-12-30 (with a fictional 1900 leap day
# already accounted for by this origin).
EXCEL_EPOCH = "1899-12-30"


class EmptySheetError(ValueError):
    """The worksheet exists but contains no data rows."""


class MissingColumnsError(ValueError):
    """One or more REQUIRED_COLUMNS are absent from the worksheet."""


def load_sales(path: str | Path, sheet: str | int | None = None) -> pd.DataFrame:
    """Read the BrightMart sales spreadsheet.

    Parameters
    ----------
    path:
        Path to the .xlsx file.
    sheet:
        Sheet name or 0-based index. ``None`` means the first sheet.

    Returns
    -------
    pd.DataFrame
        Exactly ``REQUIRED_COLUMNS``; ``Date`` as datetime64 (serial numbers
        converted), the three measure columns as numeric, text columns as
        stripped strings. Invalid values become ``NaN``/``NaT`` rather than
        raising, so the cleaning stage can decide what to do with them.
    """
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"Spreadsheet not found: {path}")

    df = pd.read_excel(path, sheet_name=0 if sheet is None else sheet, engine="openpyxl")

    if df.empty:
        raise EmptySheetError(f"Sheet has no data rows: {path}")

    df.columns = [str(col).strip() for col in df.columns]
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise MissingColumnsError(
            f"Missing required columns {missing}; found {list(df.columns)}"
        )

    df = df[REQUIRED_COLUMNS].copy()

    df["Date"] = _to_datetime(df["Date"])
    for col in NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in TEXT_COLUMNS:
        df[col] = df[col].astype("string").str.strip()

    return df


def _to_datetime(series: pd.Series) -> pd.Series:
    """Parse dates that may be real datetimes or Excel serial numbers."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin=EXCEL_EPOCH, errors="coerce")
    return pd.to_datetime(series, errors="coerce")
