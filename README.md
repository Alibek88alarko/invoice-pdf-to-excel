# PDF invoices to one clean Excel sheet, with checks

Drop PDF invoices into a folder, run one command, get one Excel workbook:

| Sheet | What is in it |
|---|---|
| **Invoices** | one row per invoice: supplier, number, date, currency, subtotal, tax, total, sum of lines, status |
| **Lines** | every line item: description, quantity, unit price, amount |
| **Check** | every invoice that needs a human look, with the reason |

An invoice goes to **Check** when its line items do not add up to the subtotal, when subtotal plus tax does not match the
total, when a quantity times price does not match the line amount, when the number, date or total is missing, or when
the file cannot be read. Nothing is dropped silently.

Example from the three sample invoices in `samples/`: two come out OK, and one goes to Check with
`subtotal + tax = 560.00, total says 590.00`. Output: [`output/invoices.xlsx`](output/invoices.xlsx).

## Run it

```
pip install -r requirements.txt
python invoice_to_excel.py samples output/invoices.xlsx
```

Dates in different formats (2026-09-12, 18.09.2026, 25/09/2026) are written as one format. Numbers are copied from the
PDF text exactly, never guessed. Nothing is sent over the network.

## Tests

```
pip install pytest
python -m pytest -q
```

Five tests: two invoice layouts, date normalisation, the mismatch going to Check, an unreadable file being flagged,
and the three sheets in the workbook. `make_samples.py` regenerates the sample invoices.

## Next steps for real invoices

This demo reads text PDFs with clear labels. For real supplier invoices the usual additions are:
- OCR for scanned PDFs and photos;
- a language model only for layouts the rules cannot read, with its answer checked by the same rules before it is
  written to Excel, so a guessed number still lands in Check instead of the results;
- a per-supplier template when one supplier sends many invoices in the same layout.

MIT licence.
