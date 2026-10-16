"""
Generate a crystal product catalogue PDF from an Excel sheet and product images.

Folder setup:
  catalogue_project/
    generate_crystal_catalogue.py
    products.xlsx
    images/
      BR0001.jpg
      BR0002_product.png
      ML0001.jpeg
    output/                         (created automatically)

Excel columns (first worksheet, row 1):
  ID | Name | Cost

The script recreates the layout of crystal_catalogue_template.pdf:
A4 portrait, category heading, 3 columns x 6 rows (18 products/page),
ID badge on image top-left, wrapped product name and right-aligned cost.

Install:
  py -m pip install openpyxl pillow reportlab

Run:
  py generate_crystal_catalogue.py
"""

import re
from pathlib import Path
from io import BytesIO

import openpyxl
from PIL import Image, ImageOps
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

# ------------------------- CONFIGURATION -------------------------
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "products.xlsx"
IMAGE_FOLDER = BASE_DIR / "images"
OUTPUT_PDF = BASE_DIR / "output" / "product_catalog.pdf"

ID_COLUMN = "ID"
NAME_COLUMN = "Name"
COST_COLUMN = "Cost"

CATEGORY_HEADERS = {
    "BR": "Bracelets",
    "ML": "Mala",
    "DC": "Decoration",
    "TR": "Tree",
    "HL": "Healing",
    "PD": "Pendants",
    "PY": "Pyramids",
}

# Page layout aligned to the supplied 3-column x 6-row template.
PAGE_W, PAGE_H = A4
MARGIN_X = 30
GAP_X = 17
COLS = 3
ROWS = 6
GRID_TOP = PAGE_H - 69
GRID_BOTTOM = 24
GAP_Y = 8
CARD_W = (PAGE_W - 2 * MARGIN_X - (COLS - 1) * GAP_X) / COLS
CARD_H = (GRID_TOP - GRID_BOTTOM - (ROWS - 1) * GAP_Y) / ROWS
IMAGE_H = CARD_H * 0.76
LABEL_H = CARD_H - IMAGE_H
IMAGE_W = CARD_W

# Colours / fonts
BG = colors.HexColor("#FBFAF7")
GOLD = colors.HexColor("#B18A3A")
DARK_GOLD = colors.HexColor("#7A5420")
NAME_COLOR = colors.HexColor("#8D6B2F")
TEXT_COLOR = colors.HexColor("#4A4032")
IMAGE_BG = colors.HexColor("#F0EEE9")
BORDER = colors.HexColor("#D9D3C8")

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
ID_PATTERN = re.compile(
    r"(?<![A-Z0-9])(?:BR|ML|DC|TR|HL|PD|PY)\d+(?![A-Z0-9])",
    re.IGNORECASE,
)


def normalize_id(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip().upper()


def load_products(path):
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = workbook.active
        first_row = next(ws.iter_rows(min_row=1, max_row=1), [])
        headers = {
            str(cell.value).strip(): index
            for index, cell in enumerate(first_row)
            if cell.value is not None
        }
        required = {ID_COLUMN, NAME_COLUMN, COST_COLUMN}
        missing_columns = required - set(headers)
        if missing_columns:
            raise ValueError(
                f"Excel is missing columns: {', '.join(sorted(missing_columns))}. "
                f"Found: {', '.join(headers)}"
            )

        products = {}
        duplicate_ids = []
        for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            product_id = normalize_id(row[headers[ID_COLUMN]])
            if not product_id:
                continue
            if product_id in products:
                duplicate_ids.append((product_id, row_number))
                continue
            name_value = row[headers[NAME_COLUMN]]
            products[product_id] = {
                "id": product_id,
                "name": "" if name_value is None else str(name_value).strip(),
                "cost": row[headers[COST_COLUMN]],
            }
        return products, duplicate_ids
    finally:
        workbook.close()


def find_images(folder):
    """Find IDs in image names, e.g. BR0001.jpg or BR0001_front.png."""
    images_by_id = {}
    for path in folder.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        stem = path.stem.upper()
        codes = set(ID_PATTERN.findall(stem))
        # Also accept a filename exactly equal to an ID.
        if not codes and re.fullmatch(r"(?:BR|ML|DC|TR|HL|PD|PY)\d+", stem):
            codes = {stem}
        for code in codes:
            images_by_id.setdefault(code.upper(), []).append(path)
    return images_by_id


def format_cost(value):
    if value is None or str(value).strip() == "":
        return ""
    if isinstance(value, (int, float)):
        return f"{value:,.0f}" if float(value).is_integer() else f"{value:,.2f}"
    value = str(value).strip()
    # Don't add a second currency symbol if Excel already includes one.
    return value if value.startswith(("â‚¹", "$", "â‚¬", "Â£")) else f"{value}"


def wrap_text(text, font, size, max_width):
    words = str(text or "").split()
    lines, current = [], ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if stringWidth(candidate, font, size) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
            current = ""
        if stringWidth(word, font, size) <= max_width:
            current = word
        else:
            part = ""
            for ch in word:
                if stringWidth(part + ch, font, size) <= max_width:
                    part += ch
                else:
                    if part:
                        lines.append(part)
                    part = ch
            current = part
    if current:
        lines.append(current)
    return lines


def draw_image_cover(pdf, path, x, y, width, height, dpi=300):
    """Fill the whole placeholder: scale to cover and centre-crop the overflow
    (no stretching, no empty bands around the image)."""
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im)  # honour phone-camera rotation
        if im.mode in ("RGBA", "LA", "P"):
            # Flatten transparency onto the placeholder colour, not black.
            im = im.convert("RGBA")
            bg_rgb = tuple(int(c * 255) for c in IMAGE_BG.rgb())
            flat = Image.new("RGB", im.size, bg_rgb)
            flat.paste(im, mask=im.split()[-1])
            im = flat
        else:
            im = im.convert("RGB")

        # Centre-crop to the placeholder's aspect ratio.
        iw, ih = im.size
        box_ratio = width / height
        if iw / ih > box_ratio:  # too wide -> trim left/right
            new_w = round(ih * box_ratio)
            left = (iw - new_w) // 2
            im = im.crop((left, 0, left + new_w, ih))
        else:  # too tall -> trim top/bottom
            new_h = round(iw / box_ratio)
            top = (ih - new_h) // 2
            im = im.crop((0, top, iw, top + new_h))

        # Downscale to print resolution to keep the PDF size reasonable.
        target = (round(width / 72 * dpi), round(height / 72 * dpi))
        if im.width > target[0]:
            im = im.resize(target, Image.LANCZOS)

        buffer = BytesIO()
        im.save(buffer, format="JPEG", quality=92)
    buffer.seek(0)
    pdf.drawImage(ImageReader(buffer), x, y, width=width, height=height)


def draw_card(pdf, product, image_path, x, y):
    # Image panel
    image_y = y + LABEL_H
    pdf.setFillColor(IMAGE_BG)
    pdf.roundRect(x, image_y, IMAGE_W, IMAGE_H, 9, stroke=0, fill=1)
    pdf.saveState()
    clip = pdf.beginPath()
    clip.roundRect(x, image_y, IMAGE_W, IMAGE_H, 9)
    pdf.clipPath(clip, stroke=0, fill=0)
    draw_image_cover(pdf, image_path, x, image_y, IMAGE_W, IMAGE_H)
    pdf.restoreState()

    # Thin outline around image
    pdf.setStrokeColor(BORDER)
    pdf.setLineWidth(0.7)
    pdf.roundRect(x, image_y, IMAGE_W, IMAGE_H, 9, stroke=1, fill=0)

    # ID badge at top-left of image
    badge_text = product["id"]
    badge_w = max(42, stringWidth(badge_text, "Helvetica-Bold", 7.5) + 10)
    badge_h = 15
    badge_x = x + 5
    badge_y = image_y + IMAGE_H - badge_h - 5
    pdf.setFillColor(colors.HexColor("#FFFDF8"))
    pdf.roundRect(badge_x, badge_y, badge_w, badge_h, 3, stroke=0, fill=1)
    pdf.setFillColor(DARK_GOLD)
    pdf.setFont("Helvetica-Bold", 7.5)
    pdf.drawString(badge_x + 4, badge_y + 4.2, badge_text)

    # Name wraps below image; cost is right-aligned on the first label line.
    label_top = y + LABEL_H - 10
    name_font, name_size = "Times-Bold", 8.5
    cost_font, cost_size = "Helvetica-Bold", 7.5
    cost_text = format_cost(product["cost"])
    cost_width = stringWidth(cost_text, cost_font, cost_size) if cost_text else 0
    name_x = x + 2
    right_x = x + CARD_W - 2
    name_max_width = max(20, CARD_W - 4 - cost_width - (8 if cost_text else 0))
    lines = wrap_text(product["name"], name_font, name_size, name_max_width)

    pdf.setFillColor(NAME_COLOR)
    pdf.setFont(name_font, name_size)
    if lines:
        pdf.drawString(name_x, label_top, lines[0])
        line_y = label_top - 10
        for line in lines[1:2]:  # card label area allows up to two lines
            pdf.drawString(name_x, line_y, line)
            line_y -= 9
    if cost_text:
        pdf.setFillColor(TEXT_COLOR)
        pdf.setFont(cost_font, cost_size)
        pdf.drawRightString(right_x, label_top, cost_text)


def draw_page_header(pdf, category_code):
    heading = CATEGORY_HEADERS.get(category_code, "Other")
    pdf.setFillColor(BG)
    pdf.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)
    pdf.setFillColor(GOLD)
    pdf.setFont("Times-Italic", 25)
    pdf.drawString(MARGIN_X, PAGE_H - 42, f"{heading}")
    pdf.setFillColor(DARK_GOLD)
    pdf.setFont("Helvetica-Bold", 13)
    pdf.drawRightString(PAGE_W - MARGIN_X, PAGE_H - 36, "Auravya Crystals")
    pdf.setStrokeColor(colors.HexColor("#E6DCC8"))
    pdf.setLineWidth(0.8)
    pdf.line(MARGIN_X, PAGE_H - 54, PAGE_W - MARGIN_X, PAGE_H - 54)


def generate_pdf():
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")
    if not IMAGE_FOLDER.is_dir():
        raise FileNotFoundError(f"Image folder not found: {IMAGE_FOLDER}")

    products, duplicate_excel_ids = load_products(EXCEL_FILE)
    images_by_id = find_images(IMAGE_FOLDER)

    matched_by_category = {}
    missing_images = []
    ambiguous_images = []
    for product_id, product in products.items():
        candidates = images_by_id.get(product_id, [])
        if not candidates:
            missing_images.append(product_id)
            continue
        if len(candidates) > 1:
            ambiguous_images.append((product_id, [p.name for p in candidates]))
            continue
        category_code = product_id[:2]
        matched_by_category.setdefault(category_code, []).append((product, candidates[0]))

    # Keep category order as specified; within each category sort by ID.
    category_order = ["BR", "ML", "DC", "TR", "HL", "PD", "PY"]
    OUTPUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(OUTPUT_PDF), pagesize=A4)
    pdf.setTitle("Auravya Crystals Product Catalogue")

    pages_created = 0
    for category_code in category_order:
        items = sorted(matched_by_category.get(category_code, []), key=lambda item: item[0]["id"])
        if not items:
            continue
        for start in range(0, len(items), COLS * ROWS):
            page_items = items[start:start + COLS * ROWS]
            draw_page_header(pdf, category_code)
            for index, (product, image_path) in enumerate(page_items):
                row = index // COLS
                col = index % COLS
                x = MARGIN_X + col * (CARD_W + GAP_X)
                y_top = GRID_TOP - row * (CARD_H + GAP_Y)
                y = y_top - CARD_H
                draw_card(pdf, product, image_path, x, y)
            pdf.setFillColor(colors.HexColor("#9B9488"))
            pdf.setFont("Helvetica", 6.5)
            pdf.drawCentredString(
                PAGE_W / 2, 10,
                f"{heading_for(category_code)} â€¢ {len(page_items)} products on this page"
            )
            pdf.showPage()
            pages_created += 1
    pdf.save()

    print(f"Catalogue created: {OUTPUT_PDF}")
    print(f"Products in catalogue: {sum(len(v) for v in matched_by_category.values())}")
    print(f"Pages created: {pages_created}")
    print(f"Excel IDs without matching image: {missing_images}")
    print(f"IDs with multiple matching images (skipped): {ambiguous_images}")
    print(f"Duplicate IDs in Excel (later rows skipped): {duplicate_excel_ids}")


def heading_for(code):
    return f"{CATEGORY_HEADERS.get(code, 'Other')}"


if __name__ == "__main__":
    generate_pdf()