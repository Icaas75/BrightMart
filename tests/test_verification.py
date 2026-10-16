"""Verification tests: traceability of every headline number.

Maps 1:1 to the five review items:
1. manual per-transaction revenue math
2. duplicates excluded exactly once
3. flagged records consistent in every summary
4. the Power BI CSV export matches the dashboard quantities
5. a different month's spreadsheet is handled transparently
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd
import pytest

from src.loader import REQUIRED_COLUMNS, EmptySheetError, MissingColumnsError
from src.powerbi import write_csvs
from src.report import build_report
from src.verify import TOL, main, run_verification

HEADER = REQUIRED_COLUMNS


def _row(**over) -> list:
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
    return [values[c] for c in REQUIRED_COLUMNS]


def _make_month_workbook(make_workbook, year: int, month: int) -> Path:
    """Two unambiguous rows in one month: revenue 2*900*1 + 1*1800*0.9 = 3420."""
    rows = [
        HEADER,
        _row(Invoice_ID="INV-A", Date=dt.datetime(year, month, 3, 10, 0),
             Quantity=2, Unit_Price=900, Discount=0),          # 1800.0
        _row(Invoice_ID="INV-B", Date=dt.datetime(year, month, 5, 11, 0),
             Quantity=1, Unit_Price=1800, Discount=0.1),       # 1620.0
    ]
    return make_workbook(rows)


# ------------------------------------------- 1. transaction-level spot check


def test_manual_revenue_matches_pipeline_real_file(real_sales_file):
    """Recompute revenue for several real invoices without pipeline helpers."""
    report = build_report(real_sales_file)
    raw = pd.read_excel(real_sales_file)
    raw.columns = [str(c).strip() for c in raw.columns]

    picked = 0
    for _, r in raw.dropna(
        subset=["Quantity", "Unit_Price", "Discount"]
    ).drop_duplicates().head(10).iterrows():
        match = report.rows[report.rows["Invoice_ID"] == str(r["Invoice_ID"]).strip()]
        if len(match) != 1:
            continue
        expected = (
            float(r["Quantity"]) * float(r["Unit_Price"]) * (1 - float(r["Discount"]))
        )
        row0 = match.iloc[0]
        actual = (
            float(row0["Quantity"])
            * float(row0["Unit_Price"])
            * (1 - float(row0["Discount"]))
        )
        assert actual == pytest.approx(expected, abs=TOL)
        picked += 1
    assert picked >= 3, "expected at least 3 unique spot-checkable invoices"


def test_total_equals_sum_of_transaction_revenue(real_sales_file):
    report = build_report(real_sales_file)
    rows = report.rows
    manual = float(
        (rows["Quantity"] * rows["Unit_Price"] * (1 - rows["Discount"])).sum()
    )
    assert manual == pytest.approx(report.totals["total_revenue"], abs=TOL)


def test_daily_and_breakdowns_crossfoot(real_sales_file):
    report = build_report(real_sales_file)
    total = report.totals["total_revenue"]
    for series in (
        report.daily["Revenue"],
        report.by_category["Total_Revenue"],
        report.by_salesperson["Total_Revenue"],
        report.by_payment["Total_Revenue"],
        report.by_customer["Total_Revenue"],
    ):
        assert float(series.sum()) == pytest.approx(total, abs=TOL)
    # top products is a subset (top 10 of more products overall)
    assert float(report.products["Total_Revenue"].sum()) <= total + TOL


# ------------------------------------------ 2. duplicates excluded exactly once


def test_duplicate_count_matches_raw_file(real_sales_file):
    """Raw file contains exactly 5 exact duplicates; none survive cleaning."""
    raw = pd.read_excel(real_sales_file)
    raw.columns = [str(c).strip() for c in raw.columns]
    assert int(raw.duplicated(keep="first").sum()) == 5

    report = build_report(real_sales_file)
    assert report.cleaning_summary["duplicates_removed"] == 5
    assert int(report.rows.duplicated().sum()) == 0


def test_single_extra_duplicate_removes_exactly_one_more(make_workbook):
    base = _row(Invoice_ID="INV-1")
    rows = [HEADER, base, base.copy()]
    report = build_report(make_workbook(rows))
    assert report.cleaning_summary["duplicates_removed"] == 1
    assert len(report.rows) == 1


# --------------------------------------- 3. flagged consistency in summaries


def test_flagged_consistency_real_file(real_sales_file):
    report = build_report(real_sales_file)
    rows = report.rows
    has_review = rows["Review_Flags"].fillna("").str.contains("review", na=False)

    assert int(has_review.sum()) == len(report.flagged) == 2
    assert report.cleaning_summary["flagged_for_review"] == 2
    # each flagged invoice appears exactly once in clean rows
    for invoice in report.flagged["Invoice_ID"]:
        assert int((rows["Invoice_ID"] == invoice).sum()) == 1
    # flagged revenue is part of, not on top of, the total
    rev = rows["Quantity"] * rows["Unit_Price"] * (1 - rows["Discount"])
    flagged_rev = float(rev[has_review].sum())
    rest_rev = float(rev[~has_review].sum())
    assert flagged_rev + rest_rev == pytest.approx(
        report.totals["total_revenue"], abs=TOL
    )


def test_flagged_row_revenue_uses_its_own_values(real_sales_file):
    """A row flagged for a conflicting price still books ITS OWN price."""
    report = build_report(real_sales_file)
    flagged = report.flagged
    assert len(flagged) == 2

    raw = pd.read_excel(real_sales_file)
    raw.columns = [str(c).strip() for c in raw.columns]
    for _, f in flagged.iterrows():
        raw_match = raw[raw["Invoice_ID"] == f["Invoice_ID"]].iloc[0]
        assert float(f["Unit_Price"]) == pytest.approx(
            float(raw_match["Unit_Price"]), abs=TOL
        )


def test_flagged_consistency_on_synthetic_conflict(make_workbook):
    """One normal row + one with a conflicting price and blank discount."""
    rows = [
        HEADER,
        _row(Invoice_ID="INV-1"),
        _row(Invoice_ID="INV-2", Unit_Price=500, Discount=None),
    ]
    report = build_report(make_workbook(rows))
    assert len(report.flagged) == 1
    assert report.cleaning_summary["flagged_for_review"] == 1
    flags = report.flagged.iloc[0]["Review_Flags"]
    assert "review" in flags.split() and "imputed_discount" in flags.split()
    # conflicting price preserved, blank discount imputed to 0
    assert float(report.flagged.iloc[0]["Unit_Price"]) == 500.0
    assert float(report.flagged.iloc[0]["Discount"]) == 0.0
    # totals include the row at its own (conflicting) price
    manual = 2 * 900 * 1.0 + 2 * 500 * (1 - 0.0)
    assert report.totals["total_revenue"] == pytest.approx(manual, abs=TOL)


# --------------------------------------------- 4. Power BI CSV vs report parity


def test_csv_export_matches_report(real_sales_file, tmp_path):
    report = build_report(real_sales_file)
    out_dir = tmp_path / "powerbi-data"
    written = write_csvs(report, out_dir)
    assert len(written) == 8

    sales = pd.read_csv(out_dir / "Sales.csv", encoding="utf-8-sig")
    total = report.totals["total_revenue"]

    assert len(sales) == report.totals["orders"] == 498
    assert float(sales["Revenue"].sum()) == pytest.approx(total, abs=TOL)
    assert sales["Invoice_ID"].is_unique

    flagged = pd.read_csv(out_dir / "FlaggedRows.csv", encoding="utf-8-sig")
    assert len(flagged) == 2
    assert float(flagged["Revenue"].sum()) == pytest.approx(
        float(
            (report.rows["Review_Flags"].fillna("").str.contains("review")).pipe(
                lambda m: (report.rows["Quantity"] * report.rows["Unit_Price"]
                           * (1 - report.rows["Discount"]))[m].sum()
            )
        ),
        abs=TOL,
    )

    # date dimension covers every sales day
    date_dim = pd.read_csv(out_dir / "Date.csv", encoding="utf-8-sig")
    assert len(date_dim) == len(report.daily) == 31

    # dimension cardinalities match the cleaned data
    clean = report.rows
    products = pd.read_csv(out_dir / "Products.csv", encoding="utf-8-sig")
    assert len(products) == clean["Product"].nunique()
    assert set(products["Category"]) == set(clean["Category"].dropna().unique())

    for frame_name, column in (
        ("Salespersons.csv", "Salesperson"),
        ("PaymentMethods.csv", "Payment_Method"),
        ("CustomerTypes.csv", "Customer_Type"),
    ):
        dim = pd.read_csv(out_dir / frame_name, encoding="utf-8-sig")
        col = dim.columns[0]
        assert set(dim[col]) == set(clean[column].dropna().unique())


def test_csv_cleaning_summary_round_trips(real_sales_file, tmp_path):
    report = build_report(real_sales_file)
    out_dir = tmp_path / "powerbi-data"
    write_csvs(report, out_dir)

    summary = pd.read_csv(out_dir / "CleaningSummary.csv", encoding="utf-8-sig").iloc[0]
    assert summary["rows_in"] == report.cleaning_summary["rows_in"] == 505
    assert summary["duplicates_removed"] == 5
    assert summary["rows_dropped"] == 2
    assert summary["rows_out"] == len(report.rows) == 498
    assert summary["flagged_for_review"] == 2


def test_csv_datekey_is_normalized_day(real_sales_file, tmp_path):
    """DateKey strips time-of-day: daily aggregation slices on it cleanly."""
    report = build_report(real_sales_file)
    out_dir = tmp_path / "powerbi-data"
    write_csvs(report, out_dir)

    sales = pd.read_csv(out_dir / "Sales.csv", encoding="utf-8-sig")
    date_dim = pd.read_csv(out_dir / "Date.csv", encoding="utf-8-sig")

    assert sales["DateKey"].str.len().eq(10).all()          # YYYY-MM-DD
    assert set(sales["DateKey"]).issubset(set(date_dim["Date"]))


def test_verify_cli_and_csvs_agree(real_sales_file, tmp_path):
    """The verify CLI also validates the freshly written CSV set."""
    report = build_report(real_sales_file)
    out_dir = tmp_path / "powerbi-data"
    write_csvs(report, out_dir)
    checks = run_verification(real_sales_file, csv_dir=out_dir)
    assert all(check.ok for check in checks), [c.name for c in checks if not c.ok]


# ------------------------------------- 5. different month / empty month files


def test_different_month_workbook_is_reported_as_its_own_month(make_workbook):
    path = _make_month_workbook(make_workbook, 2026, 9)
    report = build_report(path)
    assert report.totals["start_date"].month == 9
    assert report.totals["end_date"].month == 9
    assert report.totals["orders"] == 2
    assert report.totals["total_revenue"] == pytest.approx(3420.0, abs=TOL)
    assert len(report.daily) == 2


def test_month_summary_is_honest_about_the_input_month(real_sales_file, make_workbook):
    """August file says August; a September file must not be labeled August."""
    report = build_report(real_sales_file)
    start = report.totals["start_date"]
    assert (start.year, start.month) == (2026, 8)

    path = _make_month_workbook(make_workbook, 2026, 9)
    report_sep = build_report(path)
    start_sep = report_sep.totals["start_date"]
    assert (start_sep.year, start_sep.month) == (2026, 9)


def test_empty_month_raises_transparent_error(make_workbook):
    path = make_workbook([HEADER])  # header-only workbook: no transactions
    with pytest.raises(EmptySheetError, match="no data rows"):
        build_report(path)


def test_missing_column_raises_transparent_error(make_workbook):
    bad = [c for c in REQUIRED_COLUMNS if c != "Unit_Price"]
    with pytest.raises(MissingColumnsError, match="Unit_Price"):
        build_report(make_workbook([bad, _row()[:-1]]))


# ------------------------------------------------------------ verify CLI itself


def test_verify_cli_on_real_file(real_sales_file, tmp_path):
    out_dir = tmp_path / "powerbi-data"
    write_csvs(build_report(real_sales_file), out_dir)
    # patch the default CSV dir so the CLI checks the fresh export
    import src.verify as verify_mod

    original = verify_mod.DEFAULT_CSV_DIR
    verify_mod.DEFAULT_CSV_DIR = out_dir
    try:
        assert main(["verify", str(real_sales_file)]) == 0
    finally:
        verify_mod.DEFAULT_CSV_DIR = original


def test_verify_cli_bad_path_returns_1(tmp_path):
    assert main(["verify", str(tmp_path / "missing.xlsx")]) == 1
