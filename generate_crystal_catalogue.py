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

Excel layout: one worksheet per category. The sheet name is used as the
category heading, and categories appear in the same order as the sheet tabs.
Each sheet has these columns in row 1:
  ID | Name | Cost
Sheets without these columns are skipped with a warning.

Images are matched by ID: BR0001.jpg, BR0001_front.png and BR0001-1.jpeg all
match BR0001 (but BR0001A.jpg only matches ID BR0001A).

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
from PIL import Image, ImageChops, ImageFilter, ImageOps
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ------------------------- CONFIGURATION -------------------------
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "products.xlsx"
IMAGE_FOLDER = BASE_DIR / "images"
OUTPUT_PDF = BASE_DIR / "output" / "product_catalog.pdf"

COVER_LOGO = IMAGE_FOLDER / "cover_logo.jpeg"

# Cover-page contact links. Replace the placeholders with real details.
# Leave a value empty ("") to hide that link.
CONTACT_EMAIL = "auravyacrystals@gmail.com"
CONTACT_WHATSAPP = "+917030155355"     # include country code
CONTACT_INSTAGRAM = "auravyacrystals"    # handle without @
CONTACT_FACEBOOK = "auravyacrystals"     # page name from facebook.com/<page>

ID_COLUMN = "ID"
NAME_COLUMN = "Name"
COST_COLUMN = "Cost"

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
# Foreground photo size relative to "fit fully inside" (1.0 = no crop).
# Values > 1.0 enlarge the sharp photo; any overflow (usually top/bottom of
# portrait photos) is trimmed evenly at the placeholder edge.
FOREGROUND_SCALE = 1.3

# Start the next category on the same page if at least this fraction of the
# grid area is still free; otherwise it starts on a new page.
CONTINUE_THRESHOLD = 0.5
SECTION_HEADING_H = 36  # vertical space used by a mid-page category heading

# Colours / fonts
BG = colors.HexColor("#FBFAF7")
GOLD = colors.HexColor("#B18A3A")
DARK_GOLD = colors.HexColor("#7A5420")
NAME_COLOR = colors.HexColor("#8D6B2F")
TEXT_COLOR = colors.HexColor("#4A4032")
IMAGE_BG = colors.HexColor("#F0EEE9")
BORDER = colors.HexColor("#D9D3C8")

# Built-in PDF fonts (Helvetica) have no rupee glyph, so use a TrueType font.
RUPEE = "\u20B9"
COST_FONT = "Helvetica-Bold"
for _font_path in (
    Path("C:/Windows/Fonts/arialbd.ttf"),
    Path("/mnt/c/Windows/Fonts/arialbd.ttf"),
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
):
    if _font_path.is_file():
        pdfmetrics.registerFont(TTFont("CostFont", str(_font_path)))
        COST_FONT = "CostFont"
        break
else:
    RUPEE = "Rs. "  # fallback when no rupee-capable font is available

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}


def normalize_id(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip().upper()


def load_categories(path):
    """Read every worksheet as one category.

    Returns (categories, duplicate_ids, skipped_sheets) where categories is a
    list of (sheet_name, [product, ...]) in workbook tab order.
    """
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        categories = []
        seen = {}  # product ID -> sheet name, to detect duplicates across sheets
        duplicate_ids = []
        skipped_sheets = []
        required = {ID_COLUMN, NAME_COLUMN, COST_COLUMN}

        for ws in workbook.worksheets:
            sheet_name = ws.title.strip()
            first_row = next(ws.iter_rows(min_row=1, max_row=1), [])
            headers = {
                str(cell.value).strip(): index
                for index, cell in enumerate(first_row)
                if cell.value is not None
            }
            missing_columns = required - set(headers)
            if missing_columns:
                skipped_sheets.append(
                    (sheet_name, f"missing columns: {', '.join(sorted(missing_columns))}")
                )
                continue

            def cell(row, column):
                index = headers[column]
                return row[index] if index < len(row) else None

            products = []
            for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                product_id = normalize_id(cell(row, ID_COLUMN))
                if not product_id:
                    continue  # blank row
                if product_id in seen:
                    duplicate_ids.append((product_id, f"{sheet_name}!row {row_number}"))
                    continue
                seen[product_id] = sheet_name
                name_value = cell(row, NAME_COLUMN)
                products.append({
                    "id": product_id,
                    "name": "" if name_value is None else str(name_value).strip(),
                    "cost": cell(row, COST_COLUMN),
                })

            if products:
                categories.append((sheet_name, products))
            else:
                skipped_sheets.append((sheet_name, "no products"))
        return categories, duplicate_ids, skipped_sheets
    finally:
        workbook.close()


def find_images(folder):
    """Map upper-case filename stem -> list of image paths."""
    images_by_stem = {}
    for path in folder.rglob("*"):
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS:
            images_by_stem.setdefault(path.stem.upper(), []).append(path)
    return images_by_stem


def images_for_id(product_id, images_by_stem):
    """Images whose name is the ID, or the ID followed by a separator
    (e.g. BR0001_front, BR0001-1). An exact name match wins."""
    if product_id in images_by_stem:
        return images_by_stem[product_id]
    pattern = re.compile(rf"{re.escape(product_id)}[\s_\-.(]")
    return [
        path
        for stem, paths in images_by_stem.items()
        if pattern.match(stem)
        for path in paths
    ]


def format_cost(value):
    if value is None or str(value).strip() == "":
        return ""
    if isinstance(value, (int, float)):
        amount = f"{value:,.0f}" if float(value).is_integer() else f"{value:,.2f}"
        return f"{RUPEE}{amount}"
    value = str(value).strip()
    # Don't add a second currency symbol if Excel already includes one.
    if value.startswith(("\u20B9", "Rs", "$", "\u20AC", "\u00A3")):
        return value
    return f"{RUPEE}{value}"


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


def autotrim(im, tolerance=20, pad=0.03):
    """Remove uniform-colour borders around the product so it is drawn larger.
    The corner pixel is taken as the background colour; a small padding is
    kept so the product does not touch the placeholder edge."""
    bg = Image.new(im.mode, im.size, im.getpixel((0, 0)))
    diff = ImageChops.difference(im, bg).convert("L")
    mask = diff.point(lambda p: 255 if p > tolerance else 0)
    box = mask.getbbox()
    if not box:
        return im
    left, top, right, bottom = box
    px, py = int((right - left) * pad), int((bottom - top) * pad)
    return im.crop((
        max(0, left - px),
        max(0, top - py),
        min(im.width, right + px),
        min(im.height, bottom + py),
    ))


def draw_image_fit(pdf, path, x, y, width, height, dpi=300):
    """Fill the placeholder without cropping the product:
    1. trim plain borders around the product,
    2. draw a blurred, cover-scaled copy as the background,
    3. paste the full (contain-scaled) image centred on top."""
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

        im = autotrim(im)
        iw, ih = im.size

        # Canvas at print resolution with the placeholder's aspect ratio.
        cw = max(1, round(width / 72 * dpi))
        ch = max(1, round(height / 72 * dpi))

        # Background: cover-scale + blur (cropping is fine, it is only backdrop).
        backdrop = ImageOps.fit(im, (cw, ch), Image.LANCZOS)
        backdrop = backdrop.filter(ImageFilter.GaussianBlur(radius=max(4, cw // 25)))

        # Foreground: contain-scale x FOREGROUND_SCALE, centred.
        # Never wider than the placeholder; overflow is clipped by paste().
        scale = min(min(cw / iw, ch / ih) * FOREGROUND_SCALE, cw / iw)
        fg = im.resize(
            (max(1, round(iw * scale)), max(1, round(ih * scale))), Image.LANCZOS
        )
        backdrop.paste(fg, ((cw - fg.width) // 2, (ch - fg.height) // 2))

        buffer = BytesIO()
        backdrop.save(buffer, format="JPEG", quality=92)
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
    draw_image_fit(pdf, image_path, x + 2, image_y + 2, IMAGE_W - 4, IMAGE_H - 4)
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
    cost_font, cost_size = COST_FONT, 7.5
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


def draw_page_header(pdf, heading):
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


def draw_section_heading(pdf, heading, y_top):
    """Draw a category heading inside the grid area; return the new top y."""
    pdf.setFillColor(GOLD)
    pdf.setFont("Times-Italic", 20)
    pdf.drawString(MARGIN_X, y_top - 22, heading)
    pdf.setStrokeColor(colors.HexColor("#E6DCC8"))
    pdf.setLineWidth(0.8)
    pdf.line(MARGIN_X, y_top - 29, PAGE_W - MARGIN_X, y_top - 29)
    return y_top - SECTION_HEADING_H


def contact_links():
    """Return [(icon, label, text, url), ...] for the configured contacts."""
    links = []
    if CONTACT_EMAIL:
        links.append(("email", "Email", CONTACT_EMAIL, f"mailto:{CONTACT_EMAIL}"))
    if CONTACT_WHATSAPP:
        digits = re.sub(r"\D", "", CONTACT_WHATSAPP)  # wa.me needs digits only
        links.append(("whatsapp", "WhatsApp", CONTACT_WHATSAPP, f"https://wa.me/{digits}"))
    if CONTACT_INSTAGRAM:
        handle = CONTACT_INSTAGRAM.lstrip("@")
        links.append(("instagram", "Instagram", f"@{handle}",
                      f"https://www.instagram.com/{handle}"))
    if CONTACT_FACEBOOK:
        page = CONTACT_FACEBOOK.strip("/")
        links.append(("facebook", "Facebook", f"facebook.com/{page}",
                      f"https://www.facebook.com/{page}"))
    return links


def draw_icon(pdf, kind, cx, cy, r):
    """Simple vector icon inside a gold circle of radius r centred at (cx, cy)."""
    pdf.setFillColor(GOLD)
    pdf.circle(cx, cy, r, stroke=0, fill=1)
    pdf.setStrokeColor(colors.white)
    pdf.setFillColor(colors.white)
    pdf.setLineWidth(1.1)
    s = r * 0.5
    if kind == "email":
        pdf.rect(cx - s, cy - s * 0.7, 2 * s, 1.4 * s, stroke=1, fill=0)
        path = pdf.beginPath()
        path.moveTo(cx - s, cy + s * 0.7)
        path.lineTo(cx, cy - s * 0.05)
        path.lineTo(cx + s, cy + s * 0.7)
        pdf.drawPath(path, stroke=1, fill=0)
    elif kind == "whatsapp":
        # Speech bubble with a tail and a small handset dot pattern.
        pdf.circle(cx, cy + s * 0.05, s, stroke=1, fill=0)
        path = pdf.beginPath()
        path.moveTo(cx - s * 0.75, cy - s * 0.6)
        path.lineTo(cx - s * 1.05, cy - s * 1.05)
        path.lineTo(cx - s * 0.3, cy - s * 0.9)
        pdf.drawPath(path, stroke=1, fill=0)
        pdf.setFont("Helvetica-Bold", r * 0.75)
        pdf.drawCentredString(cx, cy - r * 0.22, "W")
    elif kind == "instagram":
        pdf.roundRect(cx - s, cy - s, 2 * s, 2 * s, s * 0.55, stroke=1, fill=0)
        pdf.circle(cx, cy, s * 0.45, stroke=1, fill=0)
        pdf.circle(cx + s * 0.55, cy + s * 0.55, s * 0.1, stroke=0, fill=1)
    elif kind == "facebook":
        pdf.setFont("Helvetica-Bold", r * 1.25)
        pdf.drawCentredString(cx + r * 0.05, cy - r * 0.43, "f")


def draw_cover_page(pdf):
    """First page: company logo, catalogue title and clickable contact links."""
    # Use the logo's own background colour so the image blends into the page.
    with Image.open(COVER_LOGO) as im:
        r, g, b = im.convert("RGB").getpixel((5, 5))
    cover_bg = colors.Color(r / 255, g / 255, b / 255)
    pdf.setFillColor(cover_bg)
    pdf.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)

    # Logo (square image), centred in the upper part of the page.
    logo_size = 430
    logo_x = (PAGE_W - logo_size) / 2
    logo_y = PAGE_H - 50 - logo_size
    pdf.drawImage(str(COVER_LOGO), logo_x, logo_y, width=logo_size, height=logo_size)

    # Title
    y = logo_y - 20
    pdf.setFillColor(GOLD)
    pdf.setFont("Times-Italic", 30)
    pdf.drawCentredString(PAGE_W / 2, y, "Product Catalogue")
    y -= 22
    pdf.setStrokeColor(colors.HexColor("#E6DCC8"))
    pdf.setLineWidth(0.8)
    pdf.line(PAGE_W / 2 - 120, y, PAGE_W / 2 + 120, y)

    links = contact_links()
    if not links:
        return
    y -= 28
    pdf.setFillColor(DARK_GOLD)
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawCentredString(PAGE_W / 2, y, "CONNECT WITH US")

    # Contact rows: icon | label | value. Whole row is a clickable link.
    icon_r = 11
    row_gap = 34
    label_font, label_size = "Helvetica-Bold", 11
    text_font, text_size = "Helvetica", 11
    label_w = max(stringWidth(l, label_font, label_size) for _, l, _, _ in links)
    text_w = max(stringWidth(t, text_font, text_size) for _, _, t, _ in links)
    block_w = 2 * icon_r + 12 + label_w + 16 + text_w
    left = (PAGE_W - block_w) / 2
    y -= 34
    for icon, label, text, url in links:
        cx = left + icon_r
        draw_icon(pdf, icon, cx, y + 4, icon_r)
        label_x = left + 2 * icon_r + 12
        pdf.setFillColor(DARK_GOLD)
        pdf.setFont(label_font, label_size)
        pdf.drawString(label_x, y, label)
        text_x = label_x + label_w + 16
        pdf.setFillColor(TEXT_COLOR)
        pdf.setFont(text_font, text_size)
        pdf.drawString(text_x, y, text)
        # Underline the value to signal it is a link.
        pdf.setStrokeColor(GOLD)
        pdf.setLineWidth(0.5)
        pdf.line(text_x, y - 2, text_x + stringWidth(text, text_font, text_size), y - 2)
        pdf.linkURL(url, (left, y - icon_r + 4, left + block_w, y + icon_r + 4),
                    relative=0, thickness=0)
        y -= row_gap

    pdf.setFillColor(colors.HexColor("#9B9488"))
    pdf.setFont("Helvetica", 8)
    pdf.drawCentredString(PAGE_W / 2, 24, "Tap any contact above to reach us directly.")


def generate_pdf():
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")
    if not IMAGE_FOLDER.is_dir():
        raise FileNotFoundError(f"Image folder not found: {IMAGE_FOLDER}")

    categories, duplicate_excel_ids, skipped_sheets = load_categories(EXCEL_FILE)
    images_by_stem = find_images(IMAGE_FOLDER)

    # Keep sheet order; products keep their row order within each sheet.
    matched_categories = []
    missing_images = []
    ambiguous_images = []
    for heading, products in categories:
        items = []
        for product in products:
            candidates = images_for_id(product["id"], images_by_stem)
            if not candidates:
                missing_images.append(product["id"])
                continue
            if len(candidates) > 1:
                ambiguous_images.append((product["id"], [p.name for p in candidates]))
                continue
            items.append((product, candidates[0]))
        if items:
            matched_categories.append((heading, items))

    OUTPUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(OUTPUT_PDF), pagesize=A4)
    pdf.setTitle("Auravya Crystals Product Catalogue")

    if COVER_LOGO.is_file():
        draw_cover_page(pdf)
        pdf.showPage()
    else:
        print(f"Cover logo not found, cover page skipped: {COVER_LOGO}")

    usable_h = GRID_TOP - GRID_BOTTOM
    row_h = CARD_H + GAP_Y
    state = {"open": False, "y": GRID_TOP, "cats": [], "count": 0, "pages": 0}

    def start_page(heading):
        draw_page_header(pdf, heading)
        state.update(open=True, y=GRID_TOP, cats=[heading], count=0)

    def finish_page():
        pdf.setFillColor(colors.HexColor("#9B9488"))
        pdf.setFont("Helvetica", 6.5)
        headings = ", ".join(state["cats"])
        pdf.drawCentredString(
            PAGE_W / 2, 10,
            f"{headings} \u2022 {state['count']} products on this page"
        )
        pdf.showPage()
        state["open"] = False
        state["pages"] += 1

    for heading, items in matched_categories:
        # Free fraction in row units (trailing row gap counted as free), so
        # e.g. 3 of 6 rows used = exactly 50% free.
        free_ratio = (state["y"] - GRID_BOTTOM + GAP_Y) / (usable_h + GAP_Y)
        if state["open"] and free_ratio >= CONTINUE_THRESHOLD - 1e-9:
            # Enough room left: continue this category on the same page.
            state["y"] = draw_section_heading(pdf, heading, state["y"])
            state["cats"].append(heading)
        else:
            if state["open"]:
                finish_page()
            start_page(heading)

        for start in range(0, len(items), COLS):
            row_items = items[start:start + COLS]
            if state["y"] - CARD_H < GRID_BOTTOM - 0.01:
                # Row doesn't fit: category continues on a new page.
                finish_page()
                start_page(heading)
            for col, (product, image_path) in enumerate(row_items):
                x = MARGIN_X + col * (CARD_W + GAP_X)
                draw_card(pdf, product, image_path, x, state["y"] - CARD_H)
            state["count"] += len(row_items)
            state["y"] -= row_h

    if state["open"]:
        finish_page()
    pages_created = state["pages"]
    pdf.save()

    print(f"Catalogue created: {OUTPUT_PDF}")
    print(f"Categories (sheets): {[h for h, _ in matched_categories]}")
    print(f"Products in catalogue: {sum(len(v) for _, v in matched_categories)}")
    print(f"Pages created: {pages_created} product pages"
          f"{' + cover' if COVER_LOGO.is_file() else ''}")
    print(f"Excel IDs without matching image: {missing_images}")
    print(f"IDs with multiple matching images (skipped): {ambiguous_images}")
    print(f"Duplicate IDs in Excel (later rows skipped): {duplicate_excel_ids}")
    print(f"Sheets skipped: {skipped_sheets}")


if __name__ == "__main__":
    generate_pdf()