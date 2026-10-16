"""Independent number verification for the BrightMart report.

Recomputes key figures straight from the raw workbook (without reusing the
pipeline's aggregation helpers) and reconciles them against the pipeline
result, the star-schema CSVs Power BI consumes, the exported Excel
workbook, and the rendered Flask dashboard:

    python -m src.verify [workbook.xlsx]

Exit code is 0 when every check passes, 1 otherwise. Each check prints a
PASS/FAIL line with the numbers behind it, so every headline figure is
traceable to transaction-level math.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from src.dashboard import app, clear_report_cache, fmt_int, fmt_money
from src.export import report_to_bytes
from src.report import build_report

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DATA = PROJECT_ROOT / "brightmart_sales_data.xlsx"
DEFAULT_CSV_DIR = PROJECT_ROOT / "powerbi" / "data"

#: Money comparisons tolerate half a cent.
TOL = 0.005


@dataclass
class Check:
    """One named reconciliation result."""

    name: str
    ok: bool
    detail: str


def _manual_revenue(qty, price, discount) -> float:
    """The agreed revenue formula, expressed independently of the pipeline."""
    return float(qty) * float(price) * (1.0 - float(discount))


def _read_raw(path) -> pd.DataFrame:
    """Read the workbook directly, bypassing the pipeline loader."""
    raw = pd.read_excel(path)
    raw.columns = [str(c).strip() for c in raw.columns]
    for col in ("Quantity", "Unit_Price", "Discount"):
        raw[col] = pd.to_numeric(raw[col], errors="coerce")
    raw["Invoice_ID"] = raw["Invoice_ID"].astype("string").str.strip()
    return raw


def _read_csvs(csv_dir: Path) -> dict[str, pd.DataFrame]:
    """Read the star-schema CSVs Power BI consumes."""
    frames: dict[str, pd.DataFrame] = {}
    for name in (
        "Sales",
        "Products",
        "Salespersons",
        "PaymentMethods",
        "CustomerTypes",
        "Date",
        "CleaningSummary",
        "FlaggedRows",
    ):
        target = csv_dir / f"{name}.csv"
        frames[name] = pd.read_csv(target, encoding="utf-8-sig")
    return frames


def fmt_money(value: float) -> str:
    return f"{float(value):,.2f}"


def fmt_int(value: float) -> str:
    return f"{int(value):,}"


def run_verification(path, csv_dir: Path | None = None) -> list[Check]:
    """Run every reconciliation check; returns them in printed order."""
    csv_dir = csv_dir or DEFAULT_CSV_DIR
    report = build_report(path)
    clean = report.rows
    summary = report.cleaning_summary
    totals = report.totals
    total = float(totals["total_revenue"])

    # Transaction-level revenue recomputed here, not via src.metrics.
    revenue = (
        clean["Quantity"] * clean["Unit_Price"] * (1.0 - clean["Discount"])
    )

    checks: list[Check] = []

    def add(name: str, ok: bool, detail: str) -> None:
        checks.append(Check(name, bool(ok), detail))

    # ------------------------------------------ 1. manual per-row spot check
    raw = _read_raw(path)
    candidates = raw.dropna(subset=["Quantity", "Unit_Price", "Discount"])
    candidates = candidates[~candidates.duplicated(keep=False)].head(3)
    problems, hits = [], []
    for _, r in candidates.iterrows():
        invoice = str(r["Invoice_ID"]).strip()
        match = clean[clean["Invoice_ID"] == invoice]
        if len(match) != 1:
            problems.append(f"{invoice}: {len(match)} matching clean rows")
            continue
        row = match.iloc[0]
        expected = _manual_revenue(r["Quantity"], r["Unit_Price"], r["Discount"])
        actual = _manual_revenue(
            row["Quantity"], row["Unit_Price"], row["Discount"]
        )
        if abs(expected - actual) > TOL:
            problems.append(f"{invoice}: pipeline {actual:,.2f} != manual {expected:,.2f}")
        else:
            hits.append(f"{invoice} = {expected:,.2f}")
    add(
        "Manual spot-check (up to 3 invoices)",
        not problems and len(hits) == len(candidates) and len(hits) > 0,
        "; ".join(problems or hits),
    )

    # --------------------------------- 2. total vs transaction-level maths
    manual_total = float(revenue.sum())
    add(
        "Total revenue matches transaction-level sum",
        abs(manual_total - total) <= TOL,
        f"{len(clean)} rows x qty*price*(1-disc) = {manual_total:,.2f} "
        f"vs headline {total:,.2f}",
    )

    # ------------------------------------------------- 3. cross-foot totals
    tables = {
        "daily": report.daily["Revenue"],
        "category": report.by_category["Total_Revenue"],
        "salesperson": report.by_salesperson["Total_Revenue"],
        "payment": report.by_payment["Total_Revenue"],
        "customer": report.by_customer["Total_Revenue"],
    }
    off = {
        name: float(series.sum()) - total
        for name, series in tables.items()
        if abs(float(series.sum()) - total) > TOL
    }
    add(
        "Every breakdown table cross-foots to the total",
        not off,
        "; ".join(f"{k} off by {v:,.2f}" for k, v in off.items())
        or f"all 5 tables sum to {total:,.2f}",
    )
    products_sum = float(report.products["Total_Revenue"].sum())
    add(
        "Top-products revenue is a subset of the total",
        products_sum <= total + TOL,
        f"top {len(report.products)} = {products_sum:,.2f} <= {total:,.2f}",
    )

    # ------------------------------------------------ 4. duplicate accounting
    raw_dupes = int(raw.duplicated(keep="first").sum())
    clean_dupes = int(clean.duplicated(keep="first").sum())
    arithmetic_ok = (
        summary["rows_in"] - summary["duplicates_removed"] - summary["rows_dropped"]
        == summary["rows_out"] == len(clean)
    )
    add(
        "Duplicates excluded exactly once",
        raw_dupes == summary["duplicates_removed"]
        and clean_dupes == 0
        and arithmetic_ok,
        f"raw file has {raw_dupes} exact-duplicate rows, summary removed "
        f"{summary['duplicates_removed']}, {clean_dupes} remain, "
        f"{summary['rows_in']} in - dup - dropped = {summary['rows_out']} out",
    )

    # ------------------------------------------------ 5. flagged consistency
    has_review = clean["Review_Flags"].fillna("").str.contains("review", na=False)
    n_review = int(has_review.sum())
    each_once = all(
        int((clean["Invoice_ID"] == inv).sum()) == 1
        for inv in report.flagged["Invoice_ID"]
    )
    flagged_rev = float(revenue[has_review].sum())
    rest_rev = float(revenue[~has_review].sum())
    add(
        "Flagged rows consistent in every summary",
        n_review == len(report.flagged) == summary["flagged_for_review"]
        and each_once
        and abs(flagged_rev + rest_rev - total) <= TOL,
        f"{n_review} flagged (summary={summary['flagged_for_review']}, "
        f"report={len(report.flagged)}), each appears once, their revenue "
        f"{flagged_rev:,.2f} is included in the total",
    )

    # ------------------------------------- 6. Power BI CSVs vs the report
    try:
        csvs = _read_csvs(csv_dir)
    except FileNotFoundError as exc:
        csvs = None
        add(
            "Power BI CSVs match the report",
            False,
            f"missing CSV in {csv_dir}: {exc}",
        )

    if csvs is not None:
        sales = csvs["Sales"]
        csv_total = float(sales["Revenue"].sum())
        flagged_csv = csvs["FlaggedRows"]
        date_rows = len(csvs["Date"])
        csv_summary = csvs["CleaningSummary"].iloc[0]
        dims_ok = (
            len(csvs["Products"]) == clean["Product"].nunique()
            and len(csvs["Salespersons"]) == clean["Salesperson"].nunique()
            and len(csvs["PaymentMethods"]) == clean["Payment_Method"].nunique()
            and len(csvs["CustomerTypes"]) == clean["Customer_Type"].nunique()
        )
        csv_ok = (
            len(sales) == totals["orders"]
            and abs(csv_total - total) <= TOL
            and len(flagged_csv) == n_review
            and abs(float(flagged_csv["Revenue"].sum()) - flagged_rev) <= TOL
            and date_rows == len(report.daily)
            and int(csv_summary["rows_out"]) == len(clean)
            and dims_ok
        )
        add(
            "Power BI CSVs match the report",
            csv_ok,
            f"{len(sales)} sales rows summing to {csv_total:,.2f}, "
            f"{len(flagged_csv)} flagged rows, {date_rows} date rows, "
            f"product/salesperson/payment/customer dims consistent"
            if csv_ok
            else f"sales={len(sales)} vs {totals['orders']} orders, "
            f"csv revenue {csv_total:,.2f} vs {total:,.2f}, "
            f"flagged {len(flagged_csv)} vs {n_review}, "
            f"dates {date_rows} vs {len(report.daily)}",
        )

    # ------------------------------------- 7. Excel workbook vs the report
    wb = load_workbook(BytesIO(report_to_bytes(report)))
    kpi = {
        row[0].value: row[1].value
        for row in wb["Summary"].iter_rows(min_col=1, max_col=2)
        if row[0].value is not None
    }
    daily_ws = wb["Daily Revenue"]
    daily_sum = sum(
        float(daily_ws.cell(row=r, column=2).value or 0)
        for r in range(2, daily_ws.max_row + 1)
    )
    excel_ok = (
        abs(float(kpi.get("Total revenue (ETB)", float("nan"))) - total) <= TOL
        and kpi.get("Orders") == totals["orders"]
        and kpi.get("Units sold") == totals["units"]
        and kpi.get("Rows flagged for review") == n_review
        and daily_ws.max_row - 1 == len(report.daily)
        and abs(daily_sum - total) <= TOL
        and wb["Cleaned Data"].max_row - 1 == len(clean)
        and wb["Flagged for Review"].max_row - 1 == n_review
    )
    add(
        "Excel workbook matches the report",
        excel_ok,
        f"Summary KPIs, {daily_ws.max_row - 1} daily rows summing to "
        f"{daily_sum:,.2f}, {wb['Cleaned Data'].max_row - 1} clean rows, "
        f"{wb['Flagged for Review'].max_row - 1} flagged rows",
    )

    # ------------------------------------------ 8. dashboard vs the report
    app.config["DATA_PATH"] = str(path)
    app.config["SHEET"] = None
    clear_report_cache()
    with app.test_client() as client:
        html = client.get("/").get_data(as_text=True)
        export_bytes = client.get("/export.xlsx").data
    clear_report_cache()
    dash_ok = (
        fmt_money(total) in html
        and fmt_int(totals["orders"]) in html
        and html.count(">review</span>") == n_review
        and export_bytes[:2] == b"PK"
    )
    add(
        "Dashboard page matches the report",
        dash_ok,
        f"revenue {fmt_money(total)} and {fmt_int(totals['orders'])} orders "
        f"rendered, {n_review} review badges, export downloadable",
    )

    return checks


def main(argv: list[str]) -> int:
    """CLI entry: print every check, exit nonzero if any fails."""
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT_DATA
    print(f"Verifying {path}")
    try:
        checks = run_verification(path)
    except Exception as exc:  # noqa: BLE001 - report the failure clearly
        print(f"FAIL  could not build a report: {exc}")
        return 1

    failed = 0
    for check in checks:
        mark = "PASS" if check.ok else "FAIL"
        failed += not check.ok
        print(f"{mark}  {check.name}\n      {check.detail}")
    print(f"\n{len(checks) - failed}/{len(checks)} checks passed")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
