"""Inspect an .xlsx file using only the Python standard library.

An .xlsx file is a ZIP archive containing XML parts. This script parses the
workbook structure (sheet names, dimensions, shared strings) and the first
data sheet's raw cell values so we can review the data before installing
pandas or writing any application code.
"""

import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

MAIN = NS["m"]


def col_to_num(col_letters: str) -> int:
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27 ..."""
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def cell_ref_parts(ref: str) -> tuple[str, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall("m:si", NS):
        text = "".join(t.text or "" for t in si.iter(f"{{{MAIN}}}t"))
        strings.append(text)
    return strings


def load_sheet_names(zf: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    names = []
    for sheet in root.iter(f"{{{MAIN}}}sheet"):
        names.append(sheet.get("name"))
    return names


def load_rows(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]):
    """Return list of rows; each row is a dict {col_index: raw_value}."""
    root = ET.fromstring(zf.read(sheet_path))
    rows = []
    for row_el in root.iter(f"{{{MAIN}}}row"):
        row = {}
        for c in row_el.findall("m:c", NS):
            ref = c.get("r") or ""
            letters, _ = cell_ref_parts(ref) if ref else ("A", 1)
            col_idx = col_to_num(letters)
            ctype = c.get("t")
            if ctype == "s":
                v = c.find("m:v", NS)
                val = shared[int(v.text)] if v is not None and v.text else ""
            elif ctype == "inlineStr":
                is_el = c.find("m:is", NS)
                val = "".join(t.text or "" for t in is_el.iter(f"{{{MAIN}}}t")) if is_el is not None else ""
            else:
                v = c.find("m:v", NS)
                val = v.text if v is not None and v.text else ""
            row[col_idx] = val
        rows.append((int(row_el.get("r") or len(rows) + 1), row))
    return rows


def fmt(v):
    if v is None:
        return "·"
    s = str(v)
    return s if len(s) <= 22 else s[:19] + "..."


def main(path: str) -> None:
    zf = zipfile.ZipFile(path)
    print(f"=== {path} ===")
    print("parts:", ", ".join(n for n in zf.namelist() if not n.startswith("xl/printerSettings")))

    names = load_sheet_names(zf)
    print("\nsheets:", names)

    shared = load_shared_strings(zf)
    print(f"shared strings: {len(shared)}")

    # find the first worksheet part
    sheet_part = "xl/worksheets/sheet1.xml"
    if sheet_part not in zf.namelist():
        sheet_part = next(n for n in zf.namelist() if n.startswith("xl/worksheets/sheet"))

    rows = load_rows(zf, sheet_part, shared)
    print(f"row count (incl. header): {len(rows)}")

    max_col = max((max(r.keys()) for _, r in rows if r), default=0)

    print(f"\n--- first 12 rows (max {max_col} cols) ---")
    for rnum, row in rows[:12]:
        cells = [fmt(row.get(i)) for i in range(1, max_col + 1)]
        print(f"r{rnum:>4}: " + " | ".join(cells))

    # dirty-row hunt
    header_row_num, header = rows[0]
    headers = [header.get(i, f"col{i}") for i in range(1, max_col + 1)]
    print(f"\n--- column profile ---")
    for i in range(1, max_col + 1):
        vals = [row.get(i, "") for _, row in rows[1:]]
        nonempty = [v for v in vals if v.strip() != ""]
        is_numeric = all(re.fullmatch(r"-?\d+(\.\d+)?([eE][+-]?\d+)?", v) for v in nonempty) if nonempty else False
        distinct = sorted(set(nonempty))
        sample = distinct[:12]
        print(f"col {i:>2} '{headers[i-1]}': {len(nonempty)}/{len(vals)} filled, "
              f"{'NUMERIC' if is_numeric else 'text'}, {len(distinct)} distinct; "
              f"sample: {[fmt(v) for v in sample[:8]]}")

    print("\n--- rows with blank cells ---")
    for rnum, row in rows[1:]:
        blanks = [headers[i - 1] for i in range(1, max_col + 1) if str(row.get(i, "")).strip() == ""]
        if blanks:
            print(f"row {rnum}: blank {blanks} | INV={row.get(2)} SP={row.get(3)} Product={row.get(4)}")

    print("\n--- duplicate Invoice_IDs ---")
    seen = {}
    for rnum, row in rows[1:]:
        inv = row.get(2)
        seen.setdefault(inv, []).append(rnum)
    for inv, rn in seen.items():
        if len(rn) > 1:
            print(f"{inv}: rows {rn}")

    print("\n--- product -> category / unit_price mapping consistency ---")
    cat_of, price_of = {}, {}
    for rnum, row in rows[1:]:
        prod, cat, price = row.get(4), row.get(5), row.get(7)
        if prod and cat and cat_of.setdefault(prod, cat) != cat:
            print(f"row {rnum}: '{prod}' has category '{cat}' but also '{cat_of[prod]}'")
        if prod and price and price_of.setdefault(prod, price) != price:
            print(f"row {rnum}: '{prod}' has price {price} but also {price_of[prod]}")
    for prod in sorted(cat_of):
        print(f"  {prod:24s} -> {cat_of[prod]:12s} price={price_of.get(prod)}")

    print("\n--- date range (excel serials) ---")
    dates = sorted(float(v) for _, row in rows[1:] for v in [row.get(1)] if v)
    if dates:
        print(f"min={dates[0]:.2f} max={dates[-1]:.2f} (serial; 45292=2024-01-01, .0=midnight)"
              f" -> min+frac={dates[0] % 1:.3f}, has time-of-day frac")

    print("\n--- are duplicate-invoice rows exact copies? ---")
    by_inv = {}
    for rnum, row in rows[1:]:
        by_inv.setdefault(row.get(2), []).append((rnum, row))
    for inv, group in by_inv.items():
        if len(group) < 2:
            continue
        (r1, a), (r2, b) = group[0], group[1]
        diffs = [headers[i - 1] for i in range(1, max_col + 1)
                 if str(a.get(i, "")).strip() != str(b.get(i, "")).strip()]
        print(f"{inv}: rows {r1}&{r2} -> {'EXACT copy' if not diffs else 'differs in: ' + ', '.join(diffs)}")

    print("\n--- case-variant groups (product / salesperson / payment / category) ---")
    for col_idx, label in ((4, "Product"), (3, "Salesperson"), (9, "Payment_Method"), (5, "Category")):
        groups = {}
        for _, row in rows[1:]:
            v = str(row.get(col_idx, "")).strip()
            if v:
                groups.setdefault(v.lower(), []).append(v)
        for key, variants in sorted(groups.items()):
            if len(set(variants)) > 1:
                from collections import Counter
                counts = Counter(variants)
                print(f"{label:15s} '{key}': " + ", ".join(f"'{v}' x{c}" for v, c in counts.most_common()))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "brightmart_sales_data.xlsx")
