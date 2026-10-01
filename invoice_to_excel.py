"""Read PDF invoices from a folder and write one clean Excel workbook.

Usage:
    python invoice_to_excel.py samples output/invoices.xlsx

Sheets:
    Invoices  one row per invoice with supplier, number, date, currency, subtotal, tax, total and a status
    Lines     every line item with quantity, unit price and amount
    Check     invoices that need a human look, with the reason (lines not adding up, missing fields, unreadable)

Text PDFs are read directly, so numbers are copied exactly, never guessed. Nothing is sent over the network.
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

NUM = r"(-?\d[\d,]*\.\d{2})"
LINE = re.compile(rf"^(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<price>{NUM[1:-1]})\s+(?P<amount>{NUM[1:-1]})$")
LABELS = {
    "number": [r"Invoice\s*No\.?:?\s*([A-Z0-9][\w\-/]*)", r"Invoice\s*#\s*([A-Z0-9][\w\-/]*)"],
    "date": [r"(?:Invoice\s*)?Date:?\s*(\d{1,4}[./\-]\d{1,2}[./\-]\d{1,4})"],
    "currency": [r"Currency:?\s*([A-Z]{3})", r"\(([A-Z]{3})\)"],
    "subtotal": [rf"Subtotal:?\s*{NUM}"],
    "tax": [rf"(?:VAT|Tax)[^:\n]*:?\s*{NUM}"],
    "total": [rf"\bTotal(?:\s*due)?:?\s*{NUM}"],
}
DATE_FORMATS = ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"]


def to_num(s):
    return round(float(s.replace(",", "")), 2)


def norm_date(s):
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def parse_text(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    inv = {"supplier": lines[0] if lines else ""}
    for field, patterns in LABELS.items():
        for p in patterns:
            m = re.search(p, text, re.I)
            if m:
                inv[field] = m.group(1)
                break
    items = []
    for l in lines:
        m = LINE.match(l)
        if m and not re.match(r"(Subtotal|VAT|Tax|Total)", m.group("desc"), re.I):
            items.append({"description": m.group("desc"), "qty": float(m.group("qty")),
                          "unit_price": to_num(m.group("price")), "amount": to_num(m.group("amount"))})
    for k in ("subtotal", "tax", "total"):
        if k in inv:
            inv[k] = to_num(inv[k])
    if "date" in inv:
        inv["date"] = norm_date(inv["date"])
    inv["lines"] = items
    return inv


def check(inv):
    problems = []
    for f in ("number", "date", "total"):
        if not inv.get(f):
            problems.append(f"missing {f}")
    if not inv["lines"]:
        problems.append("no line items found")
    line_sum = round(sum(i["amount"] for i in inv["lines"]), 2)
    for i in inv["lines"]:
        if abs(round(i["qty"] * i["unit_price"], 2) - i["amount"]) > 0.01:
            problems.append(f"line '{i['description']}': qty x price != amount")
    if "subtotal" in inv and abs(line_sum - inv["subtotal"]) > 0.01:
        problems.append(f"lines add up to {line_sum:.2f}, subtotal says {inv['subtotal']:.2f}")
    if inv.get("total") is not None:
        expected = round(inv.get("subtotal", line_sum) + inv.get("tax", 0.0), 2)
        if abs(expected - inv["total"]) > 0.01:
            problems.append(f"subtotal + tax = {expected:.2f}, total says {inv['total']:.2f}")
    inv["line_sum"] = line_sum
    return problems


def read_pdf(path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    if not text.strip():
        raise ValueError("no text layer (scanned PDF, needs OCR)")
    return parse_text(text)


def process(folder):
    results = []
    for path in sorted(Path(folder).glob("*.pdf")):
        try:
            inv = read_pdf(path)
            problems = check(inv)
        except Exception as e:  # unreadable file goes to Check, never dropped silently
            inv, problems = {"lines": [], "line_sum": 0.0}, [f"could not read: {e}"]
        inv["file"] = path.name
        inv["problems"] = problems
        inv["status"] = "OK" if not problems else "CHECK"
        results.append(inv)
    return results


def write_excel(results, out):
    wb = Workbook()
    head = Font(bold=True)
    ws = wb.active
    ws.title = "Invoices"
    cols = ["file", "supplier", "number", "date", "currency", "subtotal", "tax", "total", "line_sum", "status"]
    ws.append(cols)
    for r in results:
        ws.append([r.get(c) for c in cols])
        if r["status"] == "CHECK":
            for cell in ws[ws.max_row]:
                cell.fill = PatternFill("solid", fgColor="FFE699")
    wl = wb.create_sheet("Lines")
    wl.append(["file", "description", "qty", "unit_price", "amount"])
    for r in results:
        for i in r["lines"]:
            wl.append([r["file"], i["description"], i["qty"], i["unit_price"], i["amount"]])
    wc = wb.create_sheet("Check")
    wc.append(["file", "reason"])
    for r in results:
        for p in r["problems"]:
            wc.append([r["file"], p])
    for sheet in (ws, wl, wc):
        for cell in sheet[1]:
            cell.font = head
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = max(12, min(45, max(len(str(c.value or "")) for c in col) + 2))
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def main(argv):
    if len(argv) != 3:
        print(__doc__)
        return 2
    results = process(argv[1])
    write_excel(results, argv[2])
    ok = sum(r["status"] == "OK" for r in results)
    print(f"{len(results)} invoices: {ok} OK, {len(results) - ok} need checking -> {argv[2]}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
