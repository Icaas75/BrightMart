"""Flask dashboard for the BrightMart sales report.

Run from the project root:

    python -m src.dashboard
    # then open http://127.0.0.1:5000

Input-driven: upload an .xlsx in the browser and the report is generated
from that file — KPI cards, daily revenue area chart, breakdown bar charts,
top products, cleaning summary, rows flagged for human review, and a
one-click Excel export (/export.xlsx). If brightmart_sales_data.xlsx exists
in the project root, the upload page offers it as a one-click sample.
Headless alternative: python -m src.export FILE.xlsx. All charts are inline
SVG — no CDN needed.
"""

from __future__ import annotations

import sys
import tempfile
import uuid
from functools import lru_cache
from pathlib import Path

import pandas as pd
from flask import (
    Flask,
    Response,
    abort,
    redirect,
    render_template,
    request,
    url_for,
)
from werkzeug.utils import secure_filename

# Allow running this file directly (python src/dashboard.py) as well as
# python -m src.dashboard: make the `src.*` imports work in both cases.
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.export import report_to_bytes  # noqa: E402
from src.report import SalesReport, build_report  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DATA = PROJECT_ROOT / "brightmart_sales_data.xlsx"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "brightmart_uploads"
MAX_UPLOAD_MB = 25

app = Flask(__name__)
app.config["DATA_PATH"] = None  # set after a successful upload
app.config["SOURCE_NAME"] = ""  # display name of the loaded file
app.config["SHEET"] = None
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.errorhandler(500)
def _transparent_error(error):
    """Show the real reason (e.g. empty sheet, missing column) instead of a
    generic 'internal server error' page."""
    original = getattr(error, "original_exception", None)
    reason = str(original) if original else getattr(
        error, "description", "Unexpected error while building the report."
    )
    return (
        f"<!DOCTYPE html><html><head><title>Report error</title></head>"
        f"<body style='font-family:sans-serif;max-width:640px;margin:60px auto;'>"
        f"<h1>Could not build the report</h1>"
        f"<p><strong>{reason}</strong></p>"
        f"<p style='color:#64748b'>Check that the spreadsheet is the right "
        f"month's file, has all required columns (Date, Invoice_ID, "
        f"Salesperson, Product, Category, Quantity, Unit_Price, Discount, "
        f"Payment_Method, Customer_Type), and contains at least one data row."
        f"</p><p><a href='/refresh'>Try again</a></p></body></html>",
        500,
    )


BLUE = "#2563eb"
GREEN = "#059669"


# ------------------------------------------------------------------ caching


@lru_cache(maxsize=2)
def _cached_report(path: str, sheet: str | None) -> SalesReport:
    return build_report(path, sheet=sheet)


def get_report() -> SalesReport:
    path = app.config["DATA_PATH"]
    if not path or not Path(path).is_file():
        abort(404, description="No spreadsheet loaded — upload one first.")
    return _cached_report(path, app.config["SHEET"])


def _source_name() -> str:
    configured = app.config.get("SOURCE_NAME")
    if configured:
        return configured
    path = app.config["DATA_PATH"]
    return Path(path).name if path else ""


def _clear_old_uploads(keep: Path) -> None:
    """Remove previously uploaded workbooks except the current one."""
    try:
        for old in UPLOAD_DIR.glob("upload-*.xlsx"):
            if old != keep:
                old.unlink(missing_ok=True)
    except OSError:
        pass


def clear_report_cache() -> None:
    _cached_report.cache_clear()


# ----------------------------------------------------------------- charting


def fmt_money(value: float) -> str:
    return f"{float(value):,.2f}"


def fmt_int(value: float) -> str:
    return f"{int(value):,}"


def fmt_compact(value: float) -> str:
    value = float(value)
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.1f}M"
    if abs(value) >= 1_000:
        return f"{value / 1_000:.0f}K"
    return f"{value:.0f}"


def bar_chart_svg(pairs, width: int = 460, height: int = 200,
                  color: str = BLUE) -> str:
    """Vertical bar chart from (label, value) pairs as inline SVG."""
    pairs = list(pairs)
    if not pairs:
        return '<p class="muted">No data</p>'
    max_v = max(v for _, v in pairs) or 1.0
    left, right, top, bottom = 10, 10, 22, 30
    chart_w, chart_h = width - left - right, height - top - bottom
    slot = chart_w / len(pairs)
    bar_w = min(slot * 0.62, 90)

    parts = [
        f'<svg viewBox="0 0 {width} {height}" role="img" '
        f'preserveAspectRatio="xMidYMid meet">'
    ]
    for i, (label, value) in enumerate(pairs):
        bar_h = max(chart_h * float(value) / max_v, 2)
        x = left + i * slot + (slot - bar_w) / 2
        y = top + chart_h - bar_h
        parts.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" '
            f'height="{bar_h:.1f}" rx="4" fill="{color}">'
            f"<title>{label}: {fmt_money(value)} ETB</title></rect>"
        )
        parts.append(
            f'<text x="{x + bar_w / 2:.1f}" y="{y - 6:.1f}" '
            f'class="val">{fmt_compact(value)}</text>'
        )
        short = str(label) if len(str(label)) <= 12 else str(label)[:11] + "…"
        parts.append(
            f'<text x="{x + bar_w / 2:.1f}" y="{height - 8:.1f}" '
            f'class="lbl">{short}</text>'
        )
    parts.append("</svg>")
    return "".join(parts)


def area_chart_svg(labels, values, width: int = 900, height: int = 260) -> str:
    """Daily revenue as an area/line chart from aligned label/value lists."""
    values = [float(v) for v in values]
    n = len(values)
    if n == 0:
        return '<p class="muted">No data</p>'
    max_v = max(values) or 1.0
    left, right, top, bottom = 52, 16, 14, 34
    chart_w, chart_h = width - left - right, height - top - bottom
    step = chart_w / (n - 1) if n > 1 else 0.0

    def px(i: int) -> float:
        return left + i * step

    def py(v: float) -> float:
        return top + chart_h - chart_h * v / max_v

    line = " ".join(f"{px(i):.1f},{py(v):.1f}" for i, v in enumerate(values))
    area = (
        f"M {px(0):.1f},{top + chart_h:.1f} "
        + " ".join(f"L {px(i):.1f},{py(v):.1f}" for i, v in enumerate(values))
        + f" L {px(n - 1):.1f},{top + chart_h:.1f} Z"
    )

    parts = [f'<svg viewBox="0 0 {width} {height}" role="img">']
    for frac in (0.0, 1 / 3, 2 / 3, 1.0):
        y = top + chart_h * (1 - frac)
        parts.append(
            f'<line x1="{left}" y1="{y:.1f}" x2="{width - right}" '
            f'y2="{y:.1f}" class="grid"/>'
        )
        parts.append(
            f'<text x="{left - 8}" y="{y + 4:.1f}" class="lbl" '
            f'text-anchor="end">{fmt_compact(max_v * frac)}</text>'
        )
    parts.append(f'<path d="{area}" fill="{BLUE}" opacity="0.12"/>')
    parts.append(
        f'<polyline points="{line}" fill="none" stroke="{BLUE}" '
        f'stroke-width="2.5" stroke-linejoin="round"/>'
    )
    label_every = max(1, (n + 7) // 8)
    for i, label in enumerate(labels):
        if i % label_every and i != n - 1:
            continue
        parts.append(
            f'<circle cx="{px(i):.1f}" cy="{py(values[i]):.1f}" r="3.5" '
            f'fill="{BLUE}"><title>{label}: {fmt_money(values[i])} ETB'
            f"</title></circle>"
        )
        parts.append(
            f'<text x="{px(i):.1f}" y="{height - 8:.1f}" class="lbl" '
            f'text-anchor="middle">{label}</text>'
        )
    parts.append("</svg>")
    return "".join(parts)


# -------------------------------------------------------------------- views


def _flagged_records(flagged: pd.DataFrame) -> list[dict]:
    records = []
    for row in flagged.itertuples(index=False):
        records.append(
            {
                "invoice": row.Invoice_ID,
                "date": row.Date.strftime("%b %d, %Y %H:%M"),
                "salesperson": row.Salesperson,
                "product": row.Product,
                "category": row.Category,
                "price": fmt_money(row.Unit_Price),
                "discount": f"{float(row.Discount) * 100:.0f}%",
                "flags": row.Review_Flags or "",
            }
        )
    return records


def _has_loaded_report() -> bool:
    path = app.config["DATA_PATH"]
    return bool(path and Path(path).is_file())


@app.route("/")
def index():
    if not _has_loaded_report():
        return render_template(
            "upload.html",
            error=None,
            has_default=DEFAULT_DATA.is_file(),
            has_report=False,
        )
    report = get_report()
    totals = report.totals
    summary = report.cleaning_summary

    daily_labels = [d.strftime("%b %d") for d in report.daily["Date"]]
    breakdowns = [
        ("By category", report.by_category, "Category"),
        ("By salesperson", report.by_salesperson, "Salesperson"),
        ("By payment method", report.by_payment, "Payment_Method"),
        ("By customer type", report.by_customer, "Customer_Type"),
    ]
    total_revenue = float(totals["total_revenue"]) or 1.0
    product_shares = [
        float(v) / total_revenue * 100.0
        for v in report.products["Total_Revenue"]
    ]

    context = {
        "orders": fmt_int(totals["orders"]),
        "units": fmt_int(totals["units"]),
        "units_per_order": (
            f"{totals['units'] / totals['orders']:.1f}" if totals["orders"] else "0.0"
        ),
        "revenue": fmt_money(totals["total_revenue"]),
        "aov": fmt_money(totals["avg_order_value"]),
        "range": (
            f"{totals['start_date']:%b %d, %Y} – {totals['end_date']:%b %d, %Y}"
        ),
        "daily_chart": area_chart_svg(daily_labels, report.daily["Revenue"]),
        "product_chart": bar_chart_svg(
            zip(report.products["Product"], report.products["Total_Revenue"])
        ),
        "products": [
            (p, fmt_money(v), fmt_int(u))
            for p, v, u in zip(
                report.products["Product"],
                report.products["Total_Revenue"],
                report.products["Units"],
            )
        ],
        "breakdown_charts": [
            (title, bar_chart_svg(zip(frame[label], frame["Total_Revenue"])))
            for title, frame, label in breakdowns
        ],
        "summary": summary,
        "product_share": [f"{s:.2f}" for s in product_shares],
        "product_share_pct": [
            f"{(float(v) / total_revenue * 100.0):.1f}%"
            for v in report.products["Total_Revenue"]
        ],
        "flagged": _flagged_records(report.flagged),
        "source_name": _source_name(),
    }
    return render_template("dashboard.html", **context)


@app.route("/upload", methods=["GET", "POST"])
def upload():
    """Receive an .xlsx file, generate the report, and show the dashboard."""
    if request.method == "GET":
        return render_template(
            "upload.html",
            error=None,
            has_default=DEFAULT_DATA.is_file(),
            has_report=_has_loaded_report(),
        )

    file = request.files.get("file")
    if file is None or file.filename == "":
        return _upload_error("No file selected — choose an .xlsx file and try again.")
    filename = secure_filename(file.filename)
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        return _upload_error(f"'{filename}' is not an Excel workbook (.xlsx).")

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    target = UPLOAD_DIR / f"upload-{uuid.uuid4().hex[:8]}-{filename}"
    file.save(target)

    try:
        # Validate (and pre-warm the cache) before switching the dashboard over.
        _cached_report(str(target), None)
    except Exception as exc:  # noqa: BLE001 - the reason is shown verbatim
        target.unlink(missing_ok=True)
        return _upload_error(str(exc))

    previous = app.config["DATA_PATH"]
    app.config["DATA_PATH"] = str(target)
    app.config["SOURCE_NAME"] = filename  # display name, not the uuid-prefixed path
    app.config["SHEET"] = None
    clear_report_cache()
    _cached_report(str(target), None)
    if previous and Path(previous).parent == UPLOAD_DIR:
        try:
            Path(previous).unlink(missing_ok=True)
        except OSError:
            pass
    _clear_old_uploads(target)
    return redirect(url_for("index"))


def _upload_error(message: str):
    """Re-render the upload page with the reason the file was rejected."""
    return render_template(
        "upload.html",
        error=message,
        has_default=DEFAULT_DATA.is_file(),
        has_report=_has_loaded_report(),
    ), 400


@app.route("/use-default", methods=["POST"])
def use_default():
    """Shortcut: use the bundled sample workbook from the project root."""
    if not DEFAULT_DATA.is_file():
        return _upload_error("No sample file found in the project root.")
    app.config["DATA_PATH"] = str(DEFAULT_DATA)
    app.config["SOURCE_NAME"] = DEFAULT_DATA.name
    app.config["SHEET"] = None
    clear_report_cache()
    return redirect(url_for("index"))


@app.errorhandler(413)
def _too_large(error):
    return _upload_error(f"File too large — the limit is {MAX_UPLOAD_MB} MB."), 413


def _export_filename() -> str:
    stem = secure_filename(Path(_source_name()).stem) or "report"
    return f"{stem}_report.xlsx"


@app.route("/export.xlsx")
def export():
    if not app.config["DATA_PATH"] or not Path(app.config["DATA_PATH"]).is_file():
        return redirect(url_for("upload"))
    data = report_to_bytes(get_report())
    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={_export_filename()}"},
    )


@app.route("/refresh")
def refresh():
    clear_report_cache()
    return redirect(url_for("index"))


if __name__ == "__main__":
    print("BrightMart dashboard -> http://127.0.0.1:5000 (upload an .xlsx to begin)")
    app.run(host="127.0.0.1", port=5000, debug=False)
