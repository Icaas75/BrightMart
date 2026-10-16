"""Tests for the Flask dashboard (src.dashboard)."""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from src.dashboard import (
    app,
    area_chart_svg,
    bar_chart_svg,
    clear_report_cache,
    fmt_compact,
    fmt_int,
    fmt_money,
)


@pytest.fixture()
def client(real_sales_file):
    """Test client pointed at the real spreadsheet."""
    app.config["DATA_PATH"] = str(real_sales_file)
    app.config["SHEET"] = None
    clear_report_cache()
    with app.test_client() as test_client:
        yield test_client
    clear_report_cache()


# ---------------------------------------------------------------- formatters


def test_formatters():
    assert fmt_money(1234.5) == "1,234.50"
    assert fmt_int(498) == "498"
    assert fmt_compact(26_351_224.0) == "26.4M"
    assert fmt_compact(5_200.0) == "5K"
    assert fmt_compact(450.0) == "450"


# --------------------------------------------------------------- chart SVGs


def test_bar_chart_svg_renders_bars_and_labels():
    svg = bar_chart_svg([("A", 10.0), ("B", 20.0)])
    assert svg.startswith("<svg")
    assert svg.count("<rect") == 2
    assert ">A</text>" in svg and ">B</text>" in svg
    assert "20" in svg  # value label


def test_bar_chart_svg_empty_returns_placeholder():
    assert "No data" in bar_chart_svg([])


def test_area_chart_svg_renders_line_and_grid():
    svg = area_chart_svg(["Aug 01", "Aug 02"], [100.0, 200.0])
    assert svg.startswith("<svg")
    assert "<polyline" in svg
    assert svg.count('class="grid"') == 4
    assert "Aug 01" in svg and "Aug 02" in svg


def test_area_chart_svg_empty_returns_placeholder():
    assert "No data" in area_chart_svg([], [])


# -------------------------------------------------------------------- pages


def test_index_returns_200_with_kpis(client):
    response = client.get("/")
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    # headline numbers from the real file
    assert "498" in html          # orders
    assert "505" in html          # rows in
    assert "iPhone 15" in html    # top products
    assert "BrightMart Sales Dashboard" in html


def test_index_renders_svg_charts(client):
    html = client.get("/").get_data(as_text=True)
    assert "<svg" in html
    assert html.count("<svg") >= 5  # daily + products + 4 breakdowns


def test_index_lists_flagged_rows(client):
    html = client.get("/").get_data(as_text=True)
    assert "review" in html          # flag badge
    assert "iPhone 15" in html       # known conflicting products
    assert "Samsung Galaxy A55" in html


def test_index_without_flagged_rows_shows_ok(make_workbook):
    from src.loader import REQUIRED_COLUMNS

    header = REQUIRED_COLUMNS
    row = [
        "2026-08-01 09:00", "INV-1", "Sara", "USB-C Cable", "Accessories",
        2, 900, 0, "Cash", "Individual",
    ]
    app.config["DATA_PATH"] = str(make_workbook([header, row]))
    app.config["SHEET"] = None
    clear_report_cache()
    with app.test_client() as test_client:
        html = test_client.get("/").get_data(as_text=True)
    assert "No conflicts" in html


# ------------------------------------------------------------------- routes


def test_export_route_returns_workbook(client):
    response = client.get("/export.xlsx")
    assert response.status_code == 200
    assert response.mimetype == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # export filename derives from the loaded source file
    assert "brightmart_sales_data_report.xlsx" in response.headers["Content-Disposition"]
    data = response.data
    assert data[:2] == b"PK"
    wb = openpyxl.load_workbook(BytesIO(data))
    assert "Summary" in wb.sheetnames


def test_refresh_redirects_to_index(client):
    response = client.get("/refresh")
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/")


def test_missing_data_file_shows_upload_page(real_sales_file):
    """A configured-but-missing file means 'nothing loaded' -> upload page."""
    app.config["DATA_PATH"] = "definitely/not/here.xlsx"
    app.config["SOURCE_NAME"] = ""
    clear_report_cache()
    with app.test_client() as test_client:
        response = test_client.get("/")
    assert response.status_code == 200
    assert "Drop an .xlsx here" in response.get_data(as_text=True)
