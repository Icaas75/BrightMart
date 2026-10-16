"""Clean the loaded sales data according to the agreed policy.

Stage 2 of the pipeline. The rules below were confirmed with the data owner:

1. Rows that are exact duplicates of an earlier row are removed (keep first).
2. Rows whose product disagrees with the product catalog on Category or
   Unit_Price are **not** fixed silently: they are flagged for human review
   (Review_Flags contains "review") and left as-is.
3. Blank cells:
   - Unit_Price blank -> fill from the catalog price for that product;
   - Discount blank   -> 0;
   - Customer_Type blank -> global majority value;
   - Date or Quantity blank -> the row cannot be reasonably repaired, so it is
     dropped and counted in the cleaning summary.
4. Text values are case-normalized to the modal spelling observed in the data
   (e.g. 'iphone 15' -> 'iPhone 15'). Invoice_ID is never re-cased.
5. The time-of-day part of Date is kept in the data but ignored by daily
   aggregations (see metrics.py).
"""

from __future__ import annotations

import pandas as pd

#: Value used when Discount is blank.
DEFAULT_DISCOUNT = 0.0

#: Flag written into Review_Flags for catalog-mismatch rows.
FLAG_REVIEW = "review"

#: Flags written into Review_Flags when a value had to be imputed.
FLAG_IMPUTED_PRICE = "imputed_price"
FLAG_IMPUTED_DISCOUNT = "imputed_discount"
FLAG_IMPUTED_CUSTOMER = "imputed_customer_type"

#: Text columns whose casing is normalized. Invoice_ID is deliberately
#: excluded: identifiers must not be re-cased.
CASE_COLUMNS = [
    "Salesperson",
    "Product",
    "Category",
    "Payment_Method",
    "Customer_Type",
]

MEASURE_COLUMNS = ["Quantity", "Unit_Price", "Discount"]


def clean_sales(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Apply the full cleaning policy.

    Returns
    -------
    (clean_df, summary)
        ``clean_df`` has all original columns plus ``Review_Flags`` (space-
        separated tags). ``summary`` is a dict of counts for the dashboard.
    """
    df = df.copy()
    summary: dict = {}

    # Measures become float64 so blank (NaN) values can be imputed regardless
    # of whether integers arrived from the loader.
    for col in MEASURE_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("float64")

    # Capture the modal original spelling per lowercase key BEFORE destroying
    # casing, so we can restore it at the end.
    case_maps = _build_case_maps(df)
    _lowercase_text(df)

    _remove_exact_duplicates(df, summary)

    # Catalog is built after dedup so trusted rows define the mapping.
    catalog = _build_catalog(df)

    df["Review_Flags"] = ""
    _flag_catalog_mismatches(df, catalog, summary)
    _restore_casing(df, case_maps)
    _impute_blanks(df, catalog, summary)
    _drop_unrepairable(df, summary)

    summary["rows_in"] = (
        summary.get("duplicates_removed", 0)
        + summary.get("rows_dropped", 0)
        + len(df)
    )
    summary["rows_out"] = len(df)
    return df, summary


# --------------------------------------------------------------------- steps


def _build_case_maps(df: pd.DataFrame) -> dict[str, dict[str, str]]:
    """Per column: lowercase key -> modal original spelling.

    Grouping by lowercase key and taking the modal *raw* spelling is what
    turns 'iphone 15' (1 row) into 'iPhone 15' (39 rows).
    """
    maps: dict[str, dict[str, str]] = {}
    for col in CASE_COLUMNS:
        raw = df[col]
        keys = raw.str.lower()
        mapping: dict[str, str] = {}
        for key in keys.dropna().unique():
            spelling = _modal_value(raw[keys == key])
            if spelling is not None:
                mapping[key] = spelling
        maps[col] = mapping
    return maps


def _lowercase_text(df: pd.DataFrame) -> None:
    """Lowercase display columns so casing differences collapse for matching."""
    for col in CASE_COLUMNS:
        df[col] = df[col].str.lower().str.strip()


def _build_catalog(df: pd.DataFrame) -> dict[str, dict]:
    """Lowercase product key -> Unit_Price and lowercase Category.

    Values are modal among the product's rows; count ties go to the earliest
    occurrence so a single conflicting row can never redefine the catalog
    against an equally-sized majority.
    """
    catalog: dict[str, dict] = {}
    for key, group in df.groupby("Product", dropna=False):
        catalog[str(key)] = {
            "category": _modal_value(group["Category"]),
            "price": _modal_float(group["Unit_Price"]),
        }
    return catalog


def _remove_exact_duplicates(df: pd.DataFrame, summary: dict) -> None:
    """Drop rows identical to an earlier row across all columns (keep first)."""
    before = len(df)
    df.drop_duplicates(keep="first", inplace=True)
    summary["duplicates_removed"] = before - len(df)


def _flag_catalog_mismatches(df, catalog, summary) -> None:
    """Flag rows whose Category/Unit_Price disagree with the catalog.

    Casing is already normalized, so a mismatch means a genuinely conflicting
    value: we leave it untouched and mark the row for human review.
    """
    def _mismatch(row) -> bool:
        entry = catalog.get(str(row["Product"]).lower())
        if entry is None:
            return False
        cat_ok = pd.isna(row["Category"]) or row["Category"] == entry["category"]
        price_ok = pd.isna(row["Unit_Price"]) or row["Unit_Price"] == entry["price"]
        return not (cat_ok and price_ok)

    mask = df.apply(_mismatch, axis=1)
    summary["flagged_for_review"] = int(mask.sum())
    df.loc[mask, "Review_Flags"] = FLAG_REVIEW


def _restore_casing(df, case_maps: dict[str, dict[str, str]]) -> None:
    """Rewrite lowercase text values with their modal original spellings."""
    for col in CASE_COLUMNS:
        df[col] = df[col].map(case_maps[col]).fillna(df[col])


def _impute_blanks(df, catalog, summary) -> None:
    """Fill repairable blanks per policy; tag imputed rows."""
    missing_price = df["Unit_Price"].isna()
    if missing_price.any():
        # Lookup keys are lowercase; row values may already be display-cased.
        df.loc[missing_price, "Unit_Price"] = (
            df.loc[missing_price, "Product"]
            .str.lower()
            .map(lambda k: catalog.get(k, {}).get("price", float("nan")))
            .astype("float64")
        )
        _add_flag(df, missing_price & df["Unit_Price"].notna(), FLAG_IMPUTED_PRICE)

    missing_discount = df["Discount"].isna()
    if missing_discount.any():
        df.loc[missing_discount, "Discount"] = DEFAULT_DISCOUNT
        _add_flag(df, missing_discount, FLAG_IMPUTED_DISCOUNT)

    missing_customer = df["Customer_Type"].isna()
    if missing_customer.any():
        majority = _majority_value(df.loc[~missing_customer, "Customer_Type"])
        df.loc[missing_customer, "Customer_Type"] = majority
        _add_flag(df, missing_customer, FLAG_IMPUTED_CUSTOMER)


def _majority_value(series: pd.Series) -> str:
    """Most frequent non-null value ('Individual' for the real dataset)."""
    mode = series.mode(dropna=True)
    return str(mode.iloc[0]) if not mode.empty else ""


def _add_flag(df: pd.DataFrame, mask: pd.Series, flag: str) -> None:
    """Append ``flag`` to Review_Flags for rows where ``mask`` is True."""
    if not mask.any():
        return
    existing = df.loc[mask, "Review_Flags"].fillna("")
    df.loc[mask, "Review_Flags"] = (existing + " " + flag).str.strip()


def _drop_unrepairable(df: pd.DataFrame, summary: dict) -> None:
    """Drop rows whose Date or Quantity is blank (cannot be imputed)."""
    bad = df["Date"].isna() | df["Quantity"].isna()
    summary["rows_dropped"] = int(bad.sum())
    df.drop(index=df.index[bad], inplace=True)


# ------------------------------------------------------------------ helpers


def _modal_value(series: pd.Series) -> str | None:
    """Most frequent value; ties -> most uppercase spelling, then earliest."""
    s = series.dropna()
    if s.empty:
        return None
    counts = s.value_counts()
    top = counts.max()
    candidates = [v for v in counts.index if counts[v] == top]
    order = {v: i for i, v in enumerate(s)}
    best = min(
        candidates,
        key=lambda v: (-counts[v], -sum(ch.isupper() for ch in str(v)), order[v]),
    )
    return str(best)


def _modal_float(series: pd.Series) -> float:
    """Most frequent numeric value; ties -> earliest occurrence."""
    s = series.dropna()
    if s.empty:
        return float("nan")
    counts = s.value_counts()
    top = counts.max()
    candidates = [v for v in counts.index if counts[v] == top]
    order = {v: i for i, v in enumerate(s)}
    return float(min(candidates, key=lambda v: (-counts[v], order[v])))
