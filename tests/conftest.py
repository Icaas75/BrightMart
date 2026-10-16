"""Shared fixtures for the test suite."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
REAL_XLSX = PROJECT_ROOT / "brightmart_sales_data.xlsx"


@pytest.fixture(scope="session")
def real_sales_file() -> Path:
    """The real BrightMart spreadsheet, if it is present in the project root."""
    if not REAL_XLSX.is_file():
        pytest.skip("brightmart_sales_data.xlsx not present in project root")
    return REAL_XLSX


@pytest.fixture
def make_workbook(tmp_path):
    """Factory that writes a small .xlsx from a list of rows and returns its path."""

    def _make(rows: list[list], name: str = "tiny.xlsx", sheet: str = "Sales Data") -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet
        for row in rows:
            ws.append(row)
        path = tmp_path / name
        wb.save(path)
        return path

    return _make
