"""Generates the two synthetic PDF fixtures committed alongside this script.

Not part of the shipped app, and not a project dependency -- run once, by
hand, after `pip install reportlab` (needed only to regenerate; the
committed .pdf files are what test code actually reads):

    python tablekit_tests/fixtures/generate_fixtures.py

Every figure in both fixtures is fabricated (a fictional "Acme Test
Holdings, Inc.") but internally self-consistent -- each subtotal reconciles
exactly, both years, the way a real statement would -- so the
reconciliation trust-layer has something real to check, not just placeholder
text. These are NOT real companies' data; the tool has separately been
tested during development against real-world financial statements (see
CHANGELOG.md), which are not committed here.

sample_digital_income_statement.pdf
    Real, selectable PDF text plus ruled table lines (drawn with reportlab),
    for the normal digital-text extraction path.

sample_scanned_balance_sheet.pdf
    The same kind of statement, rendered to a raster image (Pillow) and
    saved with NO text layer at all -- for the manual box-select "OCR this
    region" failsafe, the way an actual scanned page would exercise it.
"""
from pathlib import Path

HERE = Path(__file__).resolve().parent


def make_digital_income_statement(path):
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(str(path), pagesize=letter)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(72, 740, "ACME TEST HOLDINGS, INC.")
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(72, 726, "Synthetic company -- fabricated data, for software testing only")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(72, 702, "CONSOLIDATED INCOME STATEMENT")
    c.setFont("Helvetica", 9)
    c.drawString(72, 688, "(in thousands, except per share data) -- year ended December 31,")

    # Header row carries ONLY the years, no label in the leading column --
    # analyze()'s header-row heuristic (extract_all_tables.py's `year_only`
    # check) only recognizes a row as a year header when its label cell is
    # EMPTY; a label sharing the row with the years ("Year ended December
    # 31,  2024  2023") falls through and gets read back as a data row,
    # which fed spurious extra figures into the footing check (verified
    # empirically -- the reconciliation walk was off by exactly 2024/2023,
    # the two "figures" from that mislabeled header row).
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(470, 662, "2024")
    c.drawRightString(540, 662, "2023")
    c.line(72, 655, 540, 655)

    rows = [
        ("Revenue", "125,000", "110,000", False),
        ("Cost of revenue", "(72,000)", "(64,000)", False),
        ("Gross profit", "53,000", "46,000", True),
        ("Operating expenses", "(31,000)", "(28,000)", False),
        ("Operating income", "22,000", "18,000", True),
        ("Interest expense", "(2,000)", "(1,500)", False),
        ("Income before income taxes", "20,000", "16,500", True),
        ("Income tax expense", "(4,800)", "(3,960)", False),
        ("Net income", "15,200", "12,540", True),
    ]

    # A rule below EVERY row, not just subtotals: the detector reads the
    # ruled grid as its row structure, so two rows sharing no line between
    # them get read back as one merged row (verified empirically -- an
    # earlier version of this fixture with rules only above subtotals came
    # back from extract_all_tables.py with adjacent line items concatenated
    # into a single row).
    y = 638
    for label, v2024, v2023, subtotal in rows:
        c.setFont("Helvetica-Bold" if subtotal else "Helvetica", 9)
        c.drawString(72, y, label)
        c.drawRightString(470, y, v2024)
        c.drawRightString(540, y, v2023)
        c.setLineWidth(1.1 if subtotal else 0.4)
        c.line(72, y - 5, 540, y - 5)
        y -= 16

    c.setLineWidth(1.1)
    c.setFont("Helvetica", 7)
    c.drawString(72, 74, "Note: all figures are fabricated for testing and reconcile arithmetically "
                         "(gross profit = revenue less cost of revenue, and so on).")
    c.showPage()
    c.save()


def make_scanned_balance_sheet(path):
    from PIL import Image, ImageDraw, ImageFont

    dpi = 300
    s = dpi / 200          # every spatial constant below was tuned at 200dpi
    w, h = int(8.5 * dpi), int(11 * dpi)
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)

    def F(name, pt):
        return ImageFont.truetype(rf"C:\Windows\Fonts\{name}", round(pt * s))

    title = F("arialbd.ttf", 30)
    sub = F("ariali.ttf", 17)
    head = F("arialbd.ttf", 19)
    body = F("arial.ttf", 18)
    bold_body = F("arialbd.ttf", 18)
    note_font = F("ariali.ttf", 13)

    left = round(140 * s)
    right = w - round(140 * s)
    col_2024 = right - round(140 * s)
    col_2023 = right

    def rtext(x, y, txt, font):
        bbox = d.textbbox((0, 0), txt, font=font)
        d.text((x - (bbox[2] - bbox[0]), y), txt, font=font, fill="black")

    y = round(120 * s)
    d.text((left, y), "ACME TEST HOLDINGS, INC.", font=title, fill="black")
    y += round(46 * s)
    d.text((left, y), "Synthetic company -- fabricated data, for software testing only", font=sub, fill="black")
    y += round(50 * s)
    d.text((left, y), "CONSOLIDATED STATEMENT OF FINANCIAL POSITION", font=head, fill="black")
    y += round(30 * s)
    d.text((left, y), "(in thousands)", font=sub, fill="black")
    y += round(46 * s)

    d.text((left, y), "As at December 31,", font=bold_body, fill="black")
    rtext(col_2024, y, "2024", bold_body)
    rtext(col_2023, y, "2023", bold_body)
    y += round(26 * s)
    d.line([(left, y), (right, y)], fill="black", width=round(2 * s))
    y += round(18 * s)

    def line(label, v2024, v2023, *, bold=False, indent=0):
        nonlocal y
        f = bold_body if bold else body
        d.text((left + round(indent * s), y), label, font=f, fill="black")
        if v2024 is not None:
            rtext(col_2024, y, v2024, f)
            rtext(col_2023, y, v2023, f)
        y += round(27 * s)

    def rule(indent=0, width=1):
        d.line([(left + round(indent * s), y - round(4 * s)), (right, y - round(4 * s))],
               fill="black", width=round(width * s))

    line("ASSETS", None, None, bold=True)
    line("Current assets:", None, None, indent=20)
    line("Cash and cash equivalents", "18,500", "14,200", indent=40)
    line("Trade receivables", "22,300", "19,800", indent=40)
    line("Inventories", "9,700", "8,900", indent=40)
    rule(indent=40)
    line("Total current assets", "50,500", "42,900", bold=True, indent=20)
    y += round(6 * s)
    line("Non-current assets:", None, None, indent=20)
    line("Property and equipment", "64,000", "59,500", indent=40)
    line("Intangible assets", "12,000", "11,000", indent=40)
    rule(indent=40)
    line("Total non-current assets", "76,000", "70,500", bold=True, indent=20)
    rule()
    line("TOTAL ASSETS", "126,500", "113,400", bold=True)
    y += round(14 * s)

    line("LIABILITIES", None, None, bold=True)
    line("Current liabilities:", None, None, indent=20)
    line("Trade payables", "15,200", "13,100", indent=40)
    line("Short-term borrowings", "8,000", "7,500", indent=40)
    rule(indent=40)
    line("Total current liabilities", "23,200", "20,600", bold=True, indent=20)
    y += round(6 * s)
    line("Non-current liabilities:", None, None, indent=20)
    line("Long-term borrowings", "30,000", "28,000", indent=40)
    rule(indent=40)
    line("Total non-current liabilities", "30,000", "28,000", bold=True, indent=20)
    rule()
    line("Total liabilities", "53,200", "48,600", bold=True)
    y += round(14 * s)

    line("EQUITY", None, None, bold=True)
    line("Share capital", "20,000", "20,000", indent=20)
    line("Retained earnings", "53,300", "44,800", indent=20)
    rule(indent=20)
    line("Total equity", "73,300", "64,800", bold=True)
    y += round(14 * s)
    rule(width=2)
    line("TOTAL LIABILITIES AND EQUITY", "126,500", "113,400", bold=True)

    y += round(40 * s)
    d.text((left, y), "Note: figures are fabricated for testing and balance exactly (assets = liabilities + equity).",
           font=note_font, fill="black")

    img.save(str(path), "PDF", resolution=float(dpi))


if __name__ == "__main__":
    p1 = HERE / "sample_digital_income_statement.pdf"
    p2 = HERE / "sample_scanned_balance_sheet.pdf"
    make_digital_income_statement(p1)
    make_scanned_balance_sheet(p2)
    print(f"wrote {p1}")
    print(f"wrote {p2}")
