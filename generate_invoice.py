"""
Generate an invoice / bill PDF for ordered items.

Layout:
  1. Logo (left); brand name and contact details centred
  2. Date (right)
     "INVOICE" title centred between rows 2 and 3
  3. Customer name (left), Bill number (right, below the date)
  4. Item table: image | ID | name | qty | price per unit | amount
  5. Totals: total quantity and grand total

Order items can be passed on the command line or in an Excel file.

Command line (ID:QTY:PRICE per item):
  py generate_invoice.py --customer "Customer Name" --item BR0001:2:450 --item ML0002:1:1200

Excel (see orders/sample_order.xlsx):
  py generate_invoice.py --excel orders/sample_order.xlsx
  - A row with "Customer" in column A and the name in column B (optional if
    --customer is given; --customer wins when both are present).
  - A header row with ID | Qty | Price, followed by one row per item.
    An optional "Name" column overrides the product name for that row.

Product names come from products.xlsx (all sheets) by ID; product images are
taken from social_watermarked/<ID>.<ext>. Missing images show a placeholder.

Bill numbers look like AC-2026-0001 and continue from the last number used,
stored in output/invoices/bill_counter.json (restarts at 0001 each year).
The PDF is saved as output/invoices/<bill number>.pdf.
"""

import argparse
import json
import sys
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path

import openpyxl
from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from generate_crystal_catalogue import (
    CONTACT_EMAIL,
    CONTACT_WHATSAPP,
    EXCEL_FILE as PRODUCTS_FILE,
    load_categories,
    normalize_id,
    wrap_text,
)

# ------------------------- CONFIGURATION -------------------------
BASE_DIR = Path(__file__).resolve().parent
LOGO_FILE = BASE_DIR / "images" / "cover_logo.jpeg"
IMAGE_FOLDER = BASE_DIR / "social_watermarked"
INVOICE_FOLDER = BASE_DIR / "output" / "invoices"
COUNTER_FILE = INVOICE_FOLDER / "bill_counter.json"

BRAND_NAME = "Auravya Crystals"
BILL_PREFIX = "AC"
DATE_FORMAT = "%d-%m-%Y"
IMAGE_EXTENSIONS = (".jpeg", ".jpg", ".png", ".webp")

# Excel header names (matched case-insensitively).
ID_HEADERS = {"id", "item id", "product id"}
QTY_HEADERS = {"qty", "quantity"}
PRICE_HEADERS = {"price", "unit price", "price per unit", "rate"}
NAME_HEADERS = {"name", "item name", "product name"}
CUSTOMER_LABELS = {"customer", "customer name"}

# Page layout
PAGE_W, PAGE_H = A4
MARGIN = 36
CONTENT_W = PAGE_W - 2 * MARGIN
ROW_H = 46          # item row height
THUMB = 38          # thumbnail size (points)
HEADER_ROW_H = 22   # table header height
BOTTOM_LIMIT = 70   # keep this much space free at the bottom of each page

# Table columns: (title, width, alignment). Name takes the remaining width.
COLUMNS = [
    ("", 50, "center"),          # image
    ("Item ID", 62, "left"),
    ("Item Name", None, "left"),
    ("Qty", 40, "right"),
    ("Price / Unit", 80, "right"),
    ("Amount", 88, "right"),
]
_fixed = sum(w for _, w, _ in COLUMNS if w)
COLUMNS = [(t, w if w else CONTENT_W - _fixed, a) for t, w, a in COLUMNS]

# Colours (same palette as the catalogue)
GOLD = colors.HexColor("#B18A3A")
DARK_GOLD = colors.HexColor("#7A5420")
TEXT_COLOR = colors.HexColor("#4A4032")
MUTED = colors.HexColor("#9B9488")
LINE = colors.HexColor("#E6DCC8")
ROW_ALT = colors.HexColor("#FBF8F1")
HEADER_BG = colors.HexColor("#B18A3A")
IMAGE_BG = colors.HexColor("#F0EEE9")

# Fonts: TrueType fonts are needed for the rupee sign.
FONT, FONT_BOLD = "Helvetica", "Helvetica-Bold"
RUPEE = "Rs. "
for regular, bold in (
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("/mnt/c/Windows/Fonts/arial.ttf", "/mnt/c/Windows/Fonts/arialbd.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
):
    if Path(regular).is_file() and Path(bold).is_file():
        pdfmetrics.registerFont(TTFont("InvoiceFont", regular))
        pdfmetrics.registerFont(TTFont("InvoiceFont-Bold", bold))
        FONT, FONT_BOLD, RUPEE = "InvoiceFont", "InvoiceFont-Bold", "\u20B9"
        break


# ------------------------- INPUT -------------------------
class OrderError(ValueError):
    """Invalid order input (shown to the user without a traceback)."""


def parse_qty(value, where):
    try:
        qty = Decimal(str(value).strip())
    except (InvalidOperation, AttributeError):
        raise OrderError(f"{where}: quantity '{value}' is not a number")
    if qty <= 0 or qty != qty.to_integral_value():
        raise OrderError(f"{where}: quantity must be a whole number above 0 (got {value})")
    return int(qty)


def parse_price(value, where):
    text = str(value if value is not None else "").replace(",", "").replace("\u20B9", "").strip()
    if text.lower().startswith("rs"):
        text = text[2:].lstrip(". ")
    try:
        price = Decimal(text)
    except InvalidOperation:
        raise OrderError(f"{where}: price '{value}' is not a number")
    if price < 0:
        raise OrderError(f"{where}: price cannot be negative (got {value})")
    return price


def items_from_args(item_args):
    """Parse --item ID:QTY:PRICE values."""
    items = []
    for n, raw in enumerate(item_args, start=1):
        parts = raw.split(":")
        if len(parts) != 3:
            raise OrderError(f"--item {raw}: expected ID:QTY:PRICE, e.g. BR0001:2:450")
        item_id, qty, price = parts
        where = f"--item {raw}"
        if not normalize_id(item_id):
            raise OrderError(f"{where}: item ID is empty")
        items.append({"id": normalize_id(item_id), "qty": parse_qty(qty, where),
                      "price": parse_price(price, where), "name": None})
    return items


def order_from_excel(path):
    """Read (customer, items) from the first sheet of an order Excel file."""
    if not path.is_file():
        raise OrderError(f"Order Excel not found: {path}")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb.worksheets[0]
        customer, header, items = None, None, []
        for row_no, row in enumerate(ws.iter_rows(values_only=True), start=1):
            cells = ["" if v is None else str(v).strip() for v in row]
            if not any(cells):
                continue
            if header is None:
                lowered = [c.lower().rstrip(":") for c in cells]
                if lowered[0] in CUSTOMER_LABELS:
                    customer = next((c for c in cells[1:] if c), None)
                    continue

                def find(names):
                    return next((i for i, c in enumerate(lowered) if c in names), None)

                cols = {"id": find(ID_HEADERS), "qty": find(QTY_HEADERS),
                        "price": find(PRICE_HEADERS), "name": find(NAME_HEADERS)}
                if None not in (cols["id"], cols["qty"], cols["price"]):
                    header = cols
                continue

            def cell(key):
                i = header[key]
                return row[i] if i is not None and i < len(row) else None

            item_id = normalize_id(cell("id"))
            if not item_id:
                continue  # blank / notes row
            where = f"{path.name} row {row_no} ({item_id})"
            name = cell("name")
            items.append({
                "id": item_id,
                "qty": parse_qty(cell("qty"), where),
                "price": parse_price(cell("price"), where),
                "name": str(name).strip() if name not in (None, "") else None,
            })
        if header is None:
            raise OrderError(f"{path.name}: no header row with ID, Qty and Price columns found")
        return customer, items
    finally:
        wb.close()


def product_names():
    """ID -> product name from every sheet of products.xlsx."""
    if not PRODUCTS_FILE.is_file():
        return {}
    categories, _, _ = load_categories(PRODUCTS_FILE)
    return {p["id"]: p["name"] for _, products in categories for p in products}


def find_image(item_id):
    for ext in IMAGE_EXTENSIONS:
        for candidate in (IMAGE_FOLDER / f"{item_id}{ext}", IMAGE_FOLDER / f"{item_id}{ext.upper()}"):
            if candidate.is_file():
                return candidate
    return None


# ------------------------- BILL NUMBER -------------------------
def next_bill_number(today):
    """Peek at the next bill number (saved only after the PDF is created)."""
    data = {}
    if COUNTER_FILE.is_file():
        data = json.loads(COUNTER_FILE.read_text(encoding="utf-8"))
    year = str(today.year)
    number = int(data.get(year, 0)) + 1
    # Never reuse a number whose PDF already exists (e.g. counter file deleted).
    while (INVOICE_FOLDER / f"{BILL_PREFIX}-{year}-{number:04d}.pdf").exists():
        number += 1
    return f"{BILL_PREFIX}-{year}-{number:04d}", data, year, number


def save_bill_number(data, year, number):
    data[year] = number
    COUNTER_FILE.parent.mkdir(parents=True, exist_ok=True)
    COUNTER_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


# ------------------------- FORMATTING -------------------------
def money(value):
    """Indian digit grouping, e.g. 1234567.5 -> ₹12,34,567.50"""
    value = Decimal(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    sign = "-" if value < 0 else ""
    whole, frac = f"{abs(value):.2f}".split(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ",".join(groups + [tail])
    return f"{sign}{RUPEE}{whole}.{frac}"


def thumbnail(path):
    """Square, centre-cropped thumbnail at ~200 dpi to keep the PDF small."""
    px = round(THUMB / 72 * 200)
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")
        im = ImageOps.fit(im, (px, px), Image.LANCZOS)
        buffer = BytesIO()
        im.save(buffer, format="JPEG", quality=88)
    buffer.seek(0)
    return ImageReader(buffer)


# ------------------------- DRAWING -------------------------
def draw_text(pdf, text, x, y, width, align, font, size):
    pdf.setFont(font, size)
    if align == "right":
        pdf.drawRightString(x + width - 6, y, text)
    elif align == "center":
        pdf.drawCentredString(x + width / 2, y, text)
    else:
        pdf.drawString(x + 6, y, text)


def draw_brand_header(pdf):
    """Row 1: logo on the left; brand name and contact details centred on the
    page. Returns the y below the header."""
    top = PAGE_H - MARGIN
    logo_size = 78
    if LOGO_FILE.is_file():
        pdf.drawImage(str(LOGO_FILE), MARGIN, top - logo_size, logo_size, logo_size)
    centre_x = PAGE_W / 2
    pdf.setFillColor(GOLD)
    pdf.setFont("Times-Italic", 26)
    pdf.drawCentredString(centre_x, top - 36, BRAND_NAME)

    contacts = []
    if CONTACT_EMAIL:
        contacts.append(CONTACT_EMAIL)
    numbers = CONTACT_WHATSAPP.split(",") if isinstance(CONTACT_WHATSAPP, str) else CONTACT_WHATSAPP
    contacts += [str(n).strip() for n in numbers if str(n).strip()]
    pdf.setFillColor(MUTED)
    pdf.setFont(FONT, 8.5)
    pdf.drawCentredString(centre_x, top - 54, "  |  ".join(contacts))

    y = top - logo_size - 8
    pdf.setStrokeColor(GOLD)
    pdf.setLineWidth(1)
    pdf.line(MARGIN, y, PAGE_W - MARGIN, y)
    return y


def draw_bill_info(pdf, y, bill_no, bill_date, customer):
    """Rows 2-3: customer (left, row 3), "INVOICE" centred between the two
    rows, and Date (row 2) above Bill No (row 3) on the right."""
    size = 10.5
    row2_y, row3_y = y - 22, y - 42

    # Right block: labels share a left edge, values are right-aligned.
    right = PAGE_W - MARGIN
    rows = [("Date:", bill_date.strftime(DATE_FORMAT), row2_y), ("Bill No:", bill_no, row3_y)]
    label_w = max(stringWidth(label, FONT_BOLD, size) for label, _, _ in rows)
    value_w = max(stringWidth(value, FONT, size) for _, value, _ in rows)
    label_x = right - value_w - label_w - 6
    pdf.setFillColor(TEXT_COLOR)
    for label, value, row_y in rows:
        pdf.setFont(FONT_BOLD, size)
        pdf.drawString(label_x, row_y, label)
        pdf.setFont(FONT, size)
        pdf.drawRightString(right, row_y, value)

    # Customer on the left of row 3.
    pdf.setFont(FONT_BOLD, size)
    pdf.drawString(MARGIN, row3_y, "Customer:")
    pdf.setFont(FONT, size)
    pdf.drawString(MARGIN + stringWidth("Customer: ", FONT_BOLD, size), row3_y, customer)

    # "INVOICE" centred horizontally and vertically between rows 2 and 3
    # (midpoint of the two text lines, minus half the cap height).
    title_size = 18
    middle = (row2_y + size * 0.72 + row3_y) / 2
    pdf.setFillColor(DARK_GOLD)
    pdf.setFont(FONT_BOLD, title_size)
    pdf.drawCentredString(PAGE_W / 2, middle - title_size * 0.72 / 2, "INVOICE")
    return row3_y - 16


def draw_table_header(pdf, y):
    pdf.setFillColor(HEADER_BG)
    pdf.rect(MARGIN, y - HEADER_ROW_H, CONTENT_W, HEADER_ROW_H, stroke=0, fill=1)
    pdf.setFillColor(colors.white)
    x = MARGIN
    for title, width, align in COLUMNS:
        draw_text(pdf, title, x, y - HEADER_ROW_H + 7.5, width, align, FONT_BOLD, 9)
        x += width
    return y - HEADER_ROW_H


def draw_item_row(pdf, y, index, item):
    y -= ROW_H
    if index % 2:
        pdf.setFillColor(ROW_ALT)
        pdf.rect(MARGIN, y, CONTENT_W, ROW_H, stroke=0, fill=1)

    widths = [w for _, w, _ in COLUMNS]
    xs = [MARGIN + sum(widths[:i]) for i in range(len(widths))]
    text_y = y + ROW_H / 2 - 3.5

    # Image (or placeholder)
    img_x = xs[0] + (widths[0] - THUMB) / 2
    img_y = y + (ROW_H - THUMB) / 2
    if item["image"]:
        pdf.drawImage(thumbnail(item["image"]), img_x, img_y, THUMB, THUMB)
    else:
        pdf.setFillColor(IMAGE_BG)
        pdf.rect(img_x, img_y, THUMB, THUMB, stroke=0, fill=1)
        pdf.setFillColor(MUTED)
        pdf.setFont(FONT, 6)
        pdf.drawCentredString(img_x + THUMB / 2, img_y + THUMB / 2 - 2, "No image")

    pdf.setFillColor(DARK_GOLD)
    draw_text(pdf, item["id"], xs[1], text_y, widths[1], "left", FONT_BOLD, 9)

    # Name: up to two lines, vertically centred.
    pdf.setFillColor(TEXT_COLOR)
    lines = wrap_text(item["name"], FONT, 9, widths[2] - 12)[:2] or [""]
    first_y = text_y + (5.5 if len(lines) == 2 else 0)
    for n, line in enumerate(lines):
        draw_text(pdf, line, xs[2], first_y - n * 11, widths[2], "left", FONT, 9)

    draw_text(pdf, str(item["qty"]), xs[3], text_y, widths[3], "right", FONT, 9)
    draw_text(pdf, money(item["price"]), xs[4], text_y, widths[4], "right", FONT, 9)
    draw_text(pdf, money(item["amount"]), xs[5], text_y, widths[5], "right", FONT_BOLD, 9)

    pdf.setStrokeColor(LINE)
    pdf.setLineWidth(0.5)
    pdf.line(MARGIN, y, PAGE_W - MARGIN, y)
    return y


def draw_totals(pdf, y, total_qty, grand_total):
    widths = [w for _, w, _ in COLUMNS]
    xs = [MARGIN + sum(widths[:i]) for i in range(len(widths))]
    h = 28
    y -= h
    pdf.setFillColor(colors.HexColor("#F5EEDD"))
    pdf.rect(MARGIN, y, CONTENT_W, h, stroke=0, fill=1)
    pdf.setStrokeColor(GOLD)
    pdf.setLineWidth(1)
    pdf.line(MARGIN, y + h, PAGE_W - MARGIN, y + h)
    pdf.line(MARGIN, y, PAGE_W - MARGIN, y)

    text_y = y + h / 2 - 4
    pdf.setFillColor(DARK_GOLD)
    draw_text(pdf, "Total", xs[1], text_y, widths[1] + widths[2], "left", FONT_BOLD, 11)
    draw_text(pdf, str(total_qty), xs[3], text_y, widths[3], "right", FONT_BOLD, 11)
    draw_text(pdf, money(grand_total), xs[4], text_y, widths[4] + widths[5], "right", FONT_BOLD, 11)
    return y


def draw_footer(pdf, page, pages):
    pdf.setFillColor(MUTED)
    pdf.setFont(FONT, 8)
    pdf.drawCentredString(PAGE_W / 2, 30, f"Thank you for shopping with {BRAND_NAME}!")
    if pages > 1:
        pdf.drawRightString(PAGE_W - MARGIN, 30, f"Page {page} of {pages}")


def paginate(items, first_table_top, next_table_top):
    """Split items into pages; the last page must also fit the totals row."""
    totals_h = 28 + 10
    pages, current = [], []
    top = first_table_top - HEADER_ROW_H
    y = top
    for item in items:
        if y - ROW_H < BOTTOM_LIMIT:
            pages.append(current)
            current, y = [], next_table_top - HEADER_ROW_H
        current.append(item)
        y -= ROW_H
    if y - totals_h < BOTTOM_LIMIT and current:
        # Totals don't fit: move the last item to a new page with the totals.
        pages.append(current[:-1])
        current = current[-1:]
    pages.append(current)
    return [p for p in pages if p] or [[]]


def build_invoice(customer, items, output=None, bill_date=None):
    if not customer or not customer.strip():
        raise OrderError("Customer name is required (use --customer or a 'Customer' row in Excel)")
    if not items:
        raise OrderError("No order items given")

    names = product_names()
    unknown = []
    for item in items:
        if not item.get("name"):
            item["name"] = names.get(item["id"], "")
            if not item["name"]:
                unknown.append(item["id"])
        item["amount"] = item["price"] * item["qty"]
        item["image"] = find_image(item["id"])
    total_qty = sum(i["qty"] for i in items)
    grand_total = sum((i["amount"] for i in items), Decimal(0))

    bill_date = bill_date or date.today()
    bill_no, counter, year, number = next_bill_number(bill_date)
    output = Path(output) if output else INVOICE_FOLDER / f"{bill_no}.pdf"
    output.parent.mkdir(parents=True, exist_ok=True)

    pdf = canvas.Canvas(str(output), pagesize=A4)
    pdf.setTitle(f"Invoice {bill_no}")
    pdf.setAuthor(BRAND_NAME)

    # Measure header height once to know where the table starts.
    first_top = PAGE_H - MARGIN - 78 - 8 - 22 - 20 - 16
    next_top = PAGE_H - MARGIN - 78 - 8 - 22
    pages = paginate(items, first_top, next_top)

    row_index = 0
    for page_no, page_items in enumerate(pages, start=1):
        y = draw_brand_header(pdf)
        if page_no == 1:
            y = draw_bill_info(pdf, y, bill_no, bill_date, customer.strip())
        else:
            y -= 22
            pdf.setFillColor(MUTED)
            pdf.setFont(FONT, 9)
            pdf.drawString(MARGIN, y + 6, f"Bill No: {bill_no} (continued)")
        y = draw_table_header(pdf, y)
        for item in page_items:
            y = draw_item_row(pdf, y, row_index, item)
            row_index += 1
        if page_no == len(pages):
            draw_totals(pdf, y, total_qty, grand_total)
        draw_footer(pdf, page_no, len(pages))
        pdf.showPage()
    pdf.save()

    save_bill_number(counter, year, number)
    missing_images = [i["id"] for i in items if not i["image"]]
    print(f"Invoice created: {output}")
    print(f"Bill No: {bill_no} | Date: {bill_date.strftime(DATE_FORMAT)} | Customer: {customer.strip()}")
    print(f"Items: {len(items)} | Total qty: {total_qty} | Grand total: "
          f"{money(grand_total).replace(RUPEE, 'Rs. ')}")
    if unknown:
        print(f"Warning - IDs not found in {PRODUCTS_FILE.name} (no name): {unknown}")
    if missing_images:
        print(f"Warning - no image in {IMAGE_FOLDER.name}: {missing_images}")
    return output


def main():
    # Windows consoles may not support the rupee sign; avoid crashing on print.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description="Generate an invoice PDF.")
    parser.add_argument("--customer", help="customer name")
    parser.add_argument("--item", action="append", default=[], metavar="ID:QTY:PRICE",
                        help="ordered item, repeat for each item, e.g. --item BR0001:2:450")
    parser.add_argument("--excel", type=Path, help="order Excel file (Customer row + ID/Qty/Price table)")
    parser.add_argument("--output", type=Path, help="PDF path (default: output/invoices/<bill no>.pdf)")
    args = parser.parse_args()

    if not args.item and not args.excel:
        parser.error("give order items with --item or --excel")
    try:
        customer, items = args.customer, []
        if args.excel:
            excel_customer, items = order_from_excel(args.excel)
            customer = customer or excel_customer
        items += items_from_args(args.item)
        build_invoice(customer, items, args.output)
    except OrderError as exc:
        sys.exit(f"Error: {exc}")
    except PermissionError as exc:
        sys.exit(f"Error: cannot write {exc.filename} - close it if it is open in a PDF viewer.")


if __name__ == "__main__":
    main()
