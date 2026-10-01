"""Generates three sample PDF invoices with different layouts. The third one has a deliberate total mismatch."""
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

OUT = Path(__file__).parent / "samples"

INVOICES = [
    dict(file="northwind_INV-1042.pdf", supplier="Northwind Office Supply Ltd", number="INV-1042", date="2026-09-12",
         currency="GBP", lines=[("A4 paper, 5 boxes", 5, 21.50), ("Toner cartridge HP 26A", 2, 64.00), ("Desk organiser", 3, 8.99)],
         tax_rate=0.20, total_override=None, style="classic"),
    dict(file="bluepeak_2026-0377.pdf", supplier="Blue Peak Logistics GmbH", number="2026-0377", date="18.09.2026",
         currency="EUR", lines=[("Pallet storage, September", 12, 9.50), ("Handling fee", 1, 45.00)],
         tax_rate=0.19, total_override=None, style="compact"),
    dict(file="greenline_GL-88.pdf", supplier="Greenline Cleaning Services", number="GL-88", date="25/09/2026",
         currency="USD", lines=[("Office cleaning, 4 visits", 4, 120.00), ("Window cleaning", 1, 80.00)],
         tax_rate=0.0, total_override=590.00, style="classic"),  # lines say 560.00, total says 590.00
]


def draw(inv):
    path = OUT / inv["file"]
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    sub = round(sum(q * p for _, q, p in inv["lines"]), 2)
    tax = round(sub * inv["tax_rate"], 2)
    total = inv["total_override"] if inv["total_override"] is not None else round(sub + tax, 2)
    y = h - 60
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, inv["supplier"])
    c.setFont("Helvetica", 11)
    if inv["style"] == "classic":
        c.drawString(50, y - 30, f"Invoice No: {inv['number']}")
        c.drawString(50, y - 46, f"Invoice Date: {inv['date']}")
        c.drawString(50, y - 62, f"Currency: {inv['currency']}")
    else:
        c.drawString(50, y - 30, f"Rechnung / Invoice #{inv['number']}   Date {inv['date']}   ({inv['currency']})")
    y -= 110
    c.setFont("Helvetica-Bold", 11)
    c.drawString(50, y, "Description"); c.drawString(330, y, "Qty"); c.drawString(390, y, "Unit price"); c.drawString(480, y, "Amount")
    c.setFont("Helvetica", 11)
    for desc, q, p in inv["lines"]:
        y -= 18
        c.drawString(50, y, desc); c.drawString(330, y, str(q)); c.drawString(390, y, f"{p:.2f}"); c.drawString(480, y, f"{q * p:.2f}")
    y -= 36
    c.drawString(390, y, "Subtotal:"); c.drawString(480, y, f"{sub:.2f}")
    y -= 18
    c.drawString(390, y, f"VAT {int(inv['tax_rate'] * 100)}%:"); c.drawString(480, y, f"{tax:.2f}")
    y -= 18
    c.setFont("Helvetica-Bold", 11)
    c.drawString(390, y, "Total due:"); c.drawString(480, y, f"{total:.2f}")
    c.save()


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    for inv in INVOICES:
        draw(inv)
    print("written", len(INVOICES), "invoices to", OUT)
