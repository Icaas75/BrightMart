"""Formatted multi-sheet Excel export of the sales report.

Run from the project root:

    python -m src.export SOURCE.xlsx [OUTPUT.xlsx]
    # -> brightmart_sales_report.xlsx by default

Mirrors the web dashboard: one executive dashboard tab in a light BI style,
then Summary / Daily Revenue (with 7-day average and cumulative columns) /
Top Products / one sheet per breakdown / Cleaned Data / Flagged for Review,
plus a hidden `Chart Data` staging sheet. Native Excel charts (line, bars,
doughnut) are attached to styled tables so they stay interactive in Excel.

Layout rules that keep the dashboard tab readable:
- every chart gets enough width/height that axis labels never crowd,
- section banners are written AFTER their section's cells (merging last),
- KPI cards span three full-width columns each with a gap between cards.
"""

from __future__ import annotations

import sys
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.report import SalesReport, build_report

NAVY = "0F172A"
BLUE = "2563EB"
LIGHT = "F1F5F9"
BORDER_COLOR = "E2E8F0"
WHITE = "FFFFFF"
MUTED = "64748B"
FAINT = "94A3B8"

MONEY_FMT = "#,##0.00"
MONEY_FMT_S = "#,##0"
INT_FMT = "#,##0"
PCT_FMT = "0.0%"

HAIR = Side(style="thin", color=BORDER_COLOR)
BOX = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)

#: Order in write_report_workbook; "Chart Data" hosts the native charts.
EXPECTED_SHEETS = [
    "Summary",
    "Sales Dashboard",
    "Daily Revenue",
    "Top Products",
    "By Category",
    "By Salesperson",
    "By Payment Method",
    "By Customer Type",
    "Cleaned Data",
    "Flagged for Review",
    "Chart Data",
]


# -------------------------------------------------------------------- utils


def _clean_value(value):
    """Convert pandas/numpy scalars into types openpyxl accepts."""
    if value is None or value is pd.NaT or value is pd.NA:
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return float(value)
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item") and not isinstance(value, (int, float, str, bool)):
        return value.item()  # numpy scalars -> python scalars
    return value


def _banner(ws, cell: str, text: str, span: str, size: int = 11) -> None:
    """Navy banner heading across ``span`` starting at ``cell``."""
    ws[cell] = text
    ws[cell].font = Font(bold=True, color=WHITE, size=size)
    ws[cell].fill = PatternFill("solid", fgColor=NAVY)
    ws[cell].alignment = Alignment(vertical="center", horizontal="left", indent=1)
    ws.merge_cells(f"{cell}:{span}")
    ws.row_dimensions[ws[cell].row].height = 20


def _style_header_row(ws, row: int, ncols: int, first_col: int = 1) -> None:
    for c in range(first_col, first_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(vertical="center")
        cell.border = BOX
    ws.row_dimensions[row].height = 18
    ws.freeze_panes = ws.cell(row=row + 1, column=first_col)


def _autowidth(ws, skip_rows: int = 0) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows(min_row=skip_rows + 1):
        for cell in row:
            text = str(cell.value) if cell.value is not None else ""
            widths[cell.column] = max(widths.get(cell.column, 8), min(len(text) + 2, 34))
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def _write_frame(ws, frame: pd.DataFrame, money_cols=(), pct_cols=(), int_cols=()):
    """Write a frame with a styled header row and number formats."""
    for c, name in enumerate(frame.columns, start=1):
        ws.cell(row=1, column=c, value=name)
    _style_header_row(ws, 1, len(frame.columns))
    for r, (_, row) in enumerate(frame.iterrows(), start=2):
        for c, name in enumerate(frame.columns, start=1):
            cell = ws.cell(row=r, column=c, value=_clean_value(row[name]))
            cell.border = BOX
            if name in money_cols:
                cell.number_format = MONEY_FMT
            elif name in pct_cols:
                cell.number_format = PCT_FMT
            elif name in int_cols:
                cell.number_format = INT_FMT
    _autowidth(ws)


def _enriched_daily(daily: pd.DataFrame) -> pd.DataFrame:
    """Daily frame plus 7-day average and cumulative columns."""
    out = daily.copy()
    out["7-Day Avg"] = out["Revenue"].rolling(7, min_periods=1).mean().round(2)
    out["Cumulative"] = out["Revenue"].cumsum().round(2)
    return out


def _enriched_breakdown(breakdown: pd.DataFrame, label: str,
                        clean: pd.DataFrame, group_col: str) -> pd.DataFrame:
    """Label, revenue, share of total and order counts."""
    total = breakdown["Total_Revenue"].sum()
    orders = (
        clean.groupby(group_col, dropna=False).size().rename("Orders")
        .reindex(breakdown[label].to_list(), fill_value=0)
    )
    out = breakdown.copy()
    out["Share_of_Total"] = (out["Total_Revenue"] / total) if total else 0.0
    out["Orders"] = orders.to_list()
    return out[[label, "Total_Revenue", "Share_of_Total", "Orders"]]


def _cleaned_with_revenue(rows: pd.DataFrame) -> pd.DataFrame:
    """Row-level cleaned data with the computed Revenue column appended."""
    out = rows.copy()
    out["Revenue"] = (
        out["Quantity"] * out["Unit_Price"] * (1.0 - out["Discount"])
    ).round(2)
    return out


# ------------------------------------------------------- exec dashboard tab
#
# Column plan (widths in parentheses):
#   A gap(2) | B,C,D card/section(13 each) | E gap(3) | F,G,H (13) | I gap(3)
#   J,K,L (13) | M gap(3) | N,O,P (13) | Q gap(2)
# Row plan:
#   2 banner | 3 sub | 5-7 KPI cards | 9 section banners | 10-26 charts
#   28 banners | 29 headers | 30-40 tables | 42 banners | 43 headers | 44-49
#
KPI_SPANS = [(2, 4), (6, 8), (10, 12), (14, 16)]  # (first col, last col)


def _kpi_card(ws, row: int, col_start: int, col_end: int, title: str,
              value, note: str, accent: str, number_format: str) -> None:
    """One BI-style metric tile spanning col_start..col_end, 3 rows tall."""
    for r, (text, size, color, fill, fmt) in enumerate(
        (
            (title, 9, MUTED, LIGHT, None),
            (value, 18, NAVY, WHITE, number_format),
            (note, 8, FAINT, WHITE, None),
        ),
        start=row,
    ):
        ws.cell(row=r, column=col_start, value=text)
        ws.merge_cells(start_row=r, start_column=col_start,
                       end_row=r, end_column=col_end)
        top = ws.cell(row=r, column=col_start)
        top.font = Font(size=size, bold=(r == row + 1), color=color)
        top.fill = PatternFill("solid", fgColor=fill)
        top.alignment = Alignment(vertical="center", horizontal="left", indent=1)
        if fmt:
            top.number_format = fmt
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.border = Border(
                left=HAIR if c == col_start else None,
                right=HAIR if c == col_end else None,
                bottom=HAIR if r == row + 2 else (HAIR if r == row else None),
                top=Side(style="medium", color=accent) if r == row else None,
            )
    ws.row_dimensions[row].height = 16
    ws.row_dimensions[row + 1].height = 26
    ws.row_dimensions[row + 2].height = 14


def _mini_table(ws, start_row: int, col: int, headers: list[str],
                rows: list[tuple], money_cols=(), pct_cols=(), int_cols=()) -> int:
    """Small styled two/four-column table used on the dashboard tab.

    Returns the last row written.
    """
    for c, name in enumerate(headers, start=col):
        cell = ws.cell(row=start_row, column=c, value=name)
        cell.font = Font(bold=True, size=9, color=MUTED)
        cell.border = BOX
        cell.fill = PatternFill("solid", fgColor=LIGHT)
    for r, values in enumerate(rows, start=start_row + 1):
        for c, value in enumerate(values, start=col):
            cell = ws.cell(row=r, column=c, value=_clean_value(value))
            cell.border = BOX
            index = c - col
            if index in money_cols:
                cell.number_format = MONEY_FMT
            elif index in pct_cols:
                cell.number_format = PCT_FMT
            elif index in int_cols:
                cell.number_format = INT_FMT
    return start_row + len(rows)


def _line_chart(ws_anchor_ws, data_ws, col_data: int, col_cats: int, n_rows: int,
                width_cm: float, height_cm: float, title: str | None = None,
                tick_skip: int | None = None) -> LineChart:
    chart = LineChart()
    data = Reference(data_ws, min_col=col_data, min_row=1, max_row=1 + n_rows)
    cats = Reference(data_ws, min_col=col_cats, min_row=2, max_row=1 + n_rows)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = title
    chart.height, chart.width = height_cm, width_cm
    chart.legend = None
    chart.y_axis.number_format = MONEY_FMT_S
    chart.y_axis.majorGridlines = chart.y_axis.majorGridlines  # keep defaults
    if tick_skip:
        chart.x_axis.tickLblSkip = tick_skip
        chart.x_axis.tickMarkSkip = tick_skip
    return chart


def _write_dashboard_tab(wb: Workbook, report: SalesReport) -> None:
    """One-page executive dashboard and the hidden Chart Data staging sheet."""
    totals = report.totals
    chart_data = wb["Chart Data"]

    # ---- staging data for native charts (hidden) ----
    daily = report.daily
    chart_data["A1"] = "Date"
    chart_data["B1"] = "Revenue"
    for i, (d, v) in enumerate(
        zip(daily["Date"], daily["Revenue"]), start=2
    ):
        chart_data.cell(row=i, column=1, value=d.to_pydatetime().date())
        chart_data.cell(row=i, column=2, value=float(v))

    staging = (
        ("D", "E", "Category", report.by_category, "Category"),
        ("G", "H", "Product", report.products, "Product"),
        ("J", "K", "Customer", report.by_customer, "Customer_Type"),
        ("M", "N", "Payment", report.by_payment, "Payment_Method"),
        ("P", "Q", "Salesperson", report.by_salesperson, "Salesperson"),
    )
    col_letters = {"A": 1, "D": 4, "E": 5, "G": 7, "H": 8, "J": 10,
                   "K": 11, "M": 13, "N": 14, "P": 16, "Q": 17}
    for label_col, value_col, header, frame, field in staging:
        chart_data[f"{label_col}1"] = header
        chart_data[f"{value_col}1"] = "Revenue"
        for i, (label, v) in enumerate(
            zip(frame[field], frame["Total_Revenue"]), start=2
        ):
            chart_data.cell(row=i, column=col_letters[label_col], value=label)
            chart_data.cell(row=i, column=col_letters[value_col], value=float(v))

    # ---- the dashboard grid: explicit wide columns, nothing squeezed ----
    ws = wb["Sales Dashboard"]
    ws.sheet_view.showGridLines = False
    widths = {1: 2, 5: 3, 9: 3, 13: 3, 17: 2}
    for col in (2, 3, 4, 6, 7, 8, 10, 11, 12, 14, 15, 16):
        widths[col] = 13
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    # banner (merged first so nothing under it is a merged target later)
    n_days = len(daily)
    _banner(ws, "B2", "BrightMart — Executive Sales Dashboard", "P2", size=14)
    ws["B3"] = (
        f"Report period {totals['start_date']:%b %d, %Y} – "
        f"{totals['end_date']:%b %d, %Y}   ·   {totals['orders']:,} orders   ·   "
        f"{totals['units']:,} units"
    )
    ws["B3"].font = Font(size=9, color=FAINT)

    # ---- KPI cards: four tiles, each spanning 3 columns with gaps ----
    accents = (BLUE, "059669", "7C3AED", "F59E0B")
    units_per_order = (
        totals["units"] / totals["orders"] if totals["orders"] else 0.0
    )
    cards = (
        ("REVENUE (ETB)", totals["total_revenue"],
         f"{totals['orders']:,} orders · {totals['units']:,} units", MONEY_FMT),
        ("ORDERS", totals["orders"], "transactions booked", INT_FMT),
        ("UNITS SOLD", totals["units"], f"{units_per_order:.1f} units per order", INT_FMT),
        ("AVG ORDER VALUE", totals["avg_order_value"], "ETB per order", MONEY_FMT),
    )
    for (col_start, col_end), accent, (title, value, note, fmt) in zip(
        KPI_SPANS, accents, cards
    ):
        _kpi_card(ws, 5, col_start, col_end, title, value, note, accent, fmt)

    # ---- revenue trend: a full-width line chart with breathing room ----
    _banner(ws, "B9", "Revenue Trend — daily revenue in ETB", "H9")
    line = _line_chart(
        ws, chart_data, col_data=2, col_cats=1, n_rows=n_days,
        width_cm=16.5, height_cm=8.8, tick_skip=max(n_days // 8, 1),
    )
    ws.add_chart(line, "B10")

    # ---- revenue mix: doughnut with readable value labels ----
    _banner(ws, "J9", "Revenue Mix — by category", "P9")
    by_cat = report.by_category
    dough = DoughnutChart()
    ddata = Reference(chart_data, min_col=5, min_row=1, max_row=1 + 1 + len(by_cat))
    dcats = Reference(chart_data, min_col=4, min_row=2, max_row=1 + 1 + len(by_cat))
    dough.add_data(ddata, titles_from_data=True)
    dough.set_categories(dcats)
    dough.title = None
    dough.dataLabels = DataLabelList()
    dough.dataLabels.showVal = True
    dough.dataLabels.showPercent = False
    dough.dataLabels.showCatName = False
    dough.dataLabels.numFmt = MONEY_FMT_S
    dough.height, dough.width = 8.8, 9.2
    ws.add_chart(dough, "J10")

    # ---- tables, written BEFORE their banners (banners merge last) ----
    products = report.products
    prod_rows = products.head(10)
    product_table_rows = [
        (
            row["Product"],
            float(row["Total_Revenue"]),
            int(row["Units"]),
            float(row["Total_Revenue"]) / totals["total_revenue"],
        )
        for _, row in prod_rows.iterrows()
    ]
    _mini_table(
        ws, 29, 2,
        ["Product", "Revenue (ETB)", "Units", "Share"],
        product_table_rows,
        money_cols={1}, pct_cols={3}, int_cols={2},
    )
    _banner(ws, "B28", "Product Performance — top 10 by revenue", "H28")

    salesperson_table_rows = [
        (row["Salesperson"], float(row["Total_Revenue"]))
        for _, row in report.by_salesperson.iterrows()
    ]
    _mini_table(
        ws, 29, 10,
        ["Salesperson", "Revenue (ETB)"],
        salesperson_table_rows,
        money_cols={1},
    )
    _banner(ws, "J28", "By Salesperson", "P28")

    trend_end_row = 10 + int(8.8 / 0.53)  # chart is ~8.8cm tall
    tables_top = max(trend_end_row + 2, 29)

    payment_table_rows = [
        (row["Payment_Method"], float(row["Total_Revenue"]))
        for _, row in report.by_payment.iterrows()
    ]
    pay_last = _mini_table(
        ws, tables_top + 1, 2,
        ["Payment Method", "Revenue (ETB)"],
        payment_table_rows,
        money_cols={1},
    )
    _banner(ws, f"B{tables_top}", "By Payment Method", f"F{tables_top}")

    customer_table_rows = [
        (row["Customer_Type"], float(row["Total_Revenue"]))
        for _, row in report.by_customer.iterrows()
    ]
    _mini_table(
        ws, tables_top + 1, 10,
        ["Customer Type", "Revenue (ETB)"],
        customer_table_rows,
        money_cols={1},
    )
    _banner(ws, f"J{tables_top}", "By Customer Type", f"N{tables_top}")

    # give the two lower tables room in case product table is taller
    ws.row_dimensions[tables_top].height = 20


# ------------------------------------------------------------------- build


def write_report_workbook(report: SalesReport) -> Workbook:
    """Build the full styled workbook from a SalesReport (no file I/O)."""
    totals = report.totals
    clean = report.rows
    wb = Workbook()
    wb.remove(wb.active)
    for name in EXPECTED_SHEETS:
        wb.create_sheet(name)
    wb["Chart Data"].sheet_state = "hidden"

    # ------------------------------------------------------------- Summary
    ws = wb["Summary"]
    ws.sheet_view.showGridLines = False
    period = (
        f"{totals['start_date']:%b %d, %Y} – {totals['end_date']:%b %d, %Y}"
    )
    rows = {
        **{
            "Total revenue (ETB)": totals["total_revenue"],
            "Orders": totals["orders"],
            "Units sold": totals["units"],
            "Average order value (ETB)": totals["avg_order_value"],
            "Report period": period,
        },
        **{f"Rows {k.replace('_', ' ')}": v for k, v in report.cleaning_summary.items()},
    }
    for i, (label, value) in enumerate(rows.items(), start=3):
        ws.cell(row=i, column=1, value=label).border = BOX
        cell = ws.cell(row=i, column=2, value=_clean_value(value))
        cell.border = BOX
        if label.startswith(("Total revenue", "Average")):
            cell.number_format = MONEY_FMT
        elif isinstance(value, int):
            cell.number_format = INT_FMT

    # clickable table of contents, placed clear of the value column
    toc_pairs = [
        ("Sales Dashboard", "A1"),
        ("Daily Revenue", "A1"),
        ("Top Products", "A1"),
        ("Flagged for Review", "A1"),
    ]
    for i, (label, ref) in enumerate(toc_pairs):
        cell = ws.cell(row=3, column=4 + i * 2, value=label)
        cell.hyperlink = f"#'{label}'!{ref}"
        cell.font = Font(bold=True, color=BLUE)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.border = BOX
    _banner(ws, "A1", "Summary — BrightMart sales report", "B1")
    _autowidth(ws)
    ws.column_dimensions["A"].width = 26
    for i in range(4, 4 + 2 * len(toc_pairs)):
        ws.column_dimensions[get_column_letter(i)].width = 14

    # -------------------------------------------------------- Daily Revenue
    enriched = _enriched_daily(report.daily)
    ws = wb["Daily Revenue"]
    _write_frame(
        ws, enriched,
        money_cols=("Revenue", "7-Day Avg", "Cumulative"),
        int_cols=("Orders", "Units"),
    )
    # big line chart to the right of the 5-column table
    line = _line_chart(
        ws, ws, col_data=2, col_cats=1, n_rows=len(report.daily),
        width_cm=18, height_cm=10, title="Daily Revenue (ETB)",
        tick_skip=max(len(report.daily) // 8, 1),
    )
    ws.add_chart(line, "G2")

    # --------------------------------------------------------- Top Products
    ws = wb["Top Products"]
    _write_frame(
        ws, report.products, money_cols=("Total_Revenue",), int_cols=("Units",)
    )
    bar = BarChart()
    bar.type = "bar"
    n = len(report.products)
    data = Reference(ws, min_col=2, min_row=1, max_row=1 + n)
    cats = Reference(ws, min_col=1, min_row=2, max_row=1 + n)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.title = "Top Products by Revenue (ETB)"
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True
    bar.dataLabels.numFmt = MONEY_FMT_S
    bar.height, bar.width = 10, 16
    bar.legend = None
    ws.add_chart(bar, "E2")

    # ------------------------------------------------------------ breakdowns
    breakdown_pairs = (
        ("By Category", "Category", report.by_category, "bar", ()),
        ("By Salesperson", "Salesperson", report.by_salesperson, "bar", ()),
        ("By Payment Method", "Payment_Method", report.by_payment, "col", ()),
        ("By Customer Type", "Customer_Type", report.by_customer, "bar", ("doughnut",)),
    )
    for sheet, col, frame, chart_type, extras in breakdown_pairs:
        enriched_frame = _enriched_breakdown(frame, col, clean, col)
        ws = wb[sheet]
        _write_frame(
            ws, enriched_frame,
            money_cols=("Total_Revenue",), pct_cols=("Share_of_Total",),
            int_cols=("Orders",),
        )
        n = len(frame)
        data = Reference(ws, min_col=2, min_row=1, max_row=1 + n)
        cats = Reference(ws, min_col=1, min_row=2, max_row=1 + n)
        if "doughnut" in extras:
            chart = DoughnutChart()
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showVal = True
            chart.dataLabels.numFmt = MONEY_FMT_S
        else:
            chart = BarChart()
            chart.type = chart_type
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showVal = True
            chart.dataLabels.numFmt = MONEY_FMT_S
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.title = sheet.replace("By ", "Revenue by ")
        # height scales with category count so labels never cram
        chart.height = max(5.5, 1.6 * n)
        chart.width = 14
        chart.legend = None
        anchor_col = "F" if not extras else "H"
        ws.add_chart(chart, f"{anchor_col}2")
        if "doughnut" in extras:
            doughnut = DoughnutChart()
            ddata = Reference(ws, min_col=2, min_row=1, max_row=1 + n)
            dcats = Reference(ws, min_col=1, min_row=2, max_row=1 + n)
            doughnut.add_data(ddata, titles_from_data=True)
            doughnut.set_categories(dcats)
            doughnut.title = "Customer Type Split"
            doughnut.dataLabels = DataLabelList()
            doughnut.dataLabels.showVal = True
            doughnut.dataLabels.numFmt = MONEY_FMT_S
            doughnut.height, doughnut.width = 7, 9
            doughnut.legend = None
            ws.add_chart(doughnut, "V2")

    # ---------------------------------------------------- Cleaned / Flagged
    ws = wb["Cleaned Data"]
    _write_frame(ws, _cleaned_with_revenue(clean), money_cols=(
        "Unit_Price", "Revenue", "Discount",
    ))

    ws = wb["Flagged for Review"]
    _write_frame(ws, _cleaned_with_revenue(report.flagged), money_cols=(
        "Unit_Price", "Revenue", "Discount",
    ))

    _write_dashboard_tab(wb, report)

    # "Sales Dashboard" should sit first, right after "Summary"
    current = wb.sheetnames.index("Sales Dashboard")
    wb.move_sheet("Sales Dashboard", offset=-(current - 1))
    return wb


def export_report(report: SalesReport, target: Path | str) -> Path:
    """Write the workbook to ``target`` (xlsx) and return the path."""
    target = Path(target)
    wb = write_report_workbook(report)
    wb.save(target)
    return target


def report_to_bytes(report: SalesReport) -> bytes:
    """Serialize the workbook to .xlsx bytes (for the dashboard download)."""
    buffer = BytesIO()
    write_report_workbook(report).save(buffer)
    return buffer.getvalue()


def main(argv: list[str]) -> int:
    """CLI: python -m src.export SOURCE.xlsx [OUTPUT.xlsx]"""
    if not 2 <= len(argv) <= 3:
        print("usage: python -m src.export SOURCE.xlsx [OUTPUT.xlsx]", file=sys.stderr)
        return 2
    source = Path(argv[1])
    output = Path(argv[2]) if len(argv) == 3 else Path("brightmart_sales_report.xlsx")
    export_report(build_report(source), output)
    print(f"Wrote {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
