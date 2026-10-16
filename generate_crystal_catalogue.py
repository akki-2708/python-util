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
from urllib.parse import quote

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
# Contact icons: logos in assets/logos named email.png, whatsapp.png,
# instagram.png and facebook.png. If a file is missing, a built-in drawn
# icon is used instead.
LOGO_FOLDER = BASE_DIR / "assets" / "logos"

# Cover-page contact links. Replace the placeholders with real details.
# Leave a value empty ("") to hide that link.
CONTACT_EMAIL = "auravyacrystals@gmail.com"
# One or more WhatsApp numbers (include country code). Each number gets its
# own clickable row that opens a chat with that number.
CONTACT_WHATSAPP = [
    "+917030155355",
    "+919971828843"
]
# Message pre-filled in the WhatsApp chat (the user can edit it before sending).
# Set to "" to open the chat without a message.
CONTACT_WHATSAPP_MESSAGE = "I would like to enquire about this catalogue!"
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
        # Accept a list, or a single string with numbers separated by commas.
        numbers = CONTACT_WHATSAPP
        if isinstance(numbers, str):
            numbers = numbers.split(",")
        # URL-encode the message (spaces -> %20, ! etc.) for the ?text= parameter.
        text_param = (
            f"?text={quote(CONTACT_WHATSAPP_MESSAGE, safe='')}"
            if CONTACT_WHATSAPP_MESSAGE else ""
        )
        for number in numbers:
            number = str(number).strip()
            digits = re.sub(r"\D", "", number)  # wa.me needs digits only
            if digits:
                links.append(("whatsapp", "WhatsApp", number,
                              f"https://wa.me/{digits}{text_param}"))
    if CONTACT_INSTAGRAM:
        handle = CONTACT_INSTAGRAM.lstrip("@")
        links.append(("instagram", "Instagram", f"@{handle}",
                      f"https://www.instagram.com/{handle}"))
    if CONTACT_FACEBOOK:
        page = CONTACT_FACEBOOK.strip("/")
        links.append(("facebook", "Facebook", f"facebook.com/{page}",
                      f"https://www.facebook.com/{page}"))
    return links


def whatsapp_icon_image(px=256):
    """WhatsApp-style icon (white phone in a speech bubble on the WhatsApp
    green circle) drawn with Pillow. Returns a transparent RGBA image."""
    import math
    from PIL import ImageDraw

    S = 4 * px  # draw 4x larger, then downscale for smooth edges
    u = S / 100  # design on a 100 x 100 grid
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    white = (255, 255, 255, 255)
    d.ellipse((0, 0, S - 1, S - 1), fill=(37, 211, 102, 255))  # #25D366

    # Speech-bubble ring with a tail at the bottom-left.
    c, r, t = 50 * u, 30 * u, 5.5 * u
    d.ellipse((c - r, c - r, c + r, c + r), outline=white, width=round(t))
    d.polygon([(25 * u, 68 * u), (20 * u, 82 * u), (35 * u, 77 * u)], fill=white)

    # Handset: thick curve (overlapping discs) with rounded ear/mouth pieces.
    cx, cy, rc, w = 54 * u, 46 * u, 11.5 * u, 3.4 * u
    for i in range(61):
        a = math.radians(95 + i * 80 / 60)
        x, y = cx + rc * math.cos(a), cy + rc * math.sin(a)
        d.ellipse((x - w, y - w, x + w, y + w), fill=white)

    def blob(x, y, rx, ry, angle):
        mask = Image.new("L", (S, S), 0)
        ImageDraw.Draw(mask).ellipse((x - rx, y - ry, x + rx, y + ry), fill=255)
        mask = mask.rotate(angle, center=(x, y), resample=Image.BICUBIC)
        im.paste(white, (0, 0), mask)

    blob(cx - rc - 0.8 * u, cy - 4.5 * u, 4.6 * u, 6.2 * u, -15)  # earpiece
    blob(cx + 4.5 * u, cy + rc + 0.8 * u, 6.2 * u, 4.6 * u, -15)  # mouthpiece
    return im.resize((px, px), Image.LANCZOS)


_icon_cache = {}


def _icon_canvas(px):
    """Supersampled transparent canvas: returns (size, unit, image) where
    unit is 1/100 of the size, so shapes are designed on a 100 x 100 grid."""
    S = 4 * px
    return S, S / 100, Image.new("RGBA", (S, S), (0, 0, 0, 0))


def gmail_icon_image(px=256):
    """Gmail-style multicolour 'M' envelope on a white circle."""
    from PIL import ImageDraw

    S, u, im = _icon_canvas(px)
    d = ImageDraw.Draw(im)
    d.ellipse((0, 0, S - 1, S - 1), fill=(255, 255, 255, 255),
              outline=(218, 220, 224, 255), width=round(1.5 * u))

    def pts(points):
        return [(x * u, y * u) for x, y in points]

    blue, green, red = (66, 133, 244, 255), (52, 168, 83, 255), (234, 67, 53, 255)
    yellow, dark_red = (251, 188, 4, 255), (197, 34, 31, 255)
    # Side stems (rounded on the outer bottom corner).
    d.rounded_rectangle(pts([(19, 33), (31, 72)]), radius=3 * u, fill=blue)
    d.rectangle(pts([(25, 33), (31, 72)]), fill=blue)
    d.rounded_rectangle(pts([(69, 33), (81, 72)]), radius=3 * u, fill=green)
    d.rectangle(pts([(69, 33), (75, 72)]), fill=green)
    # Red 'V' of the envelope flap, then the coloured top corners.
    d.polygon(pts([(19, 33), (50, 56), (81, 33), (81, 46), (50, 69), (19, 46)]), fill=red)
    d.polygon(pts([(19, 33), (31, 41.9), (31, 54.9), (19, 46)]), fill=dark_red)
    d.polygon(pts([(81, 33), (69, 41.9), (69, 54.9), (81, 46)]), fill=yellow)
    d.ellipse(pts([(19, 27), (31, 39)]), fill=dark_red)
    d.ellipse(pts([(69, 27), (81, 39)]), fill=yellow)
    return im.resize((px, px), Image.LANCZOS)


def instagram_icon_image(px=256):
    """Instagram-style camera outline on the yellow-orange-pink-purple gradient."""
    from PIL import ImageDraw

    S, u, im = _icon_canvas(px)
    stops = [(0.0, (254, 218, 117)), (0.25, (250, 126, 30)), (0.5, (214, 41, 118)),
             (0.75, (150, 47, 191)), (1.0, (79, 91, 213))]

    def colour(t):
        for (t0, c0), (t1, c1) in zip(stops, stops[1:]):
            if t <= t1:
                f = (t - t0) / (t1 - t0)
                return tuple(round(a + (b - a) * f) for a, b in zip(c0, c1)) + (255,)
        return stops[-1][1] + (255,)

    # Radial gradient from the bottom-left corner; built small, then upscaled.
    n = 64
    small = Image.new("RGBA", (n, n))
    for y in range(n):
        for x in range(n):
            dx, dy = x / (n - 1) - 0.1, (n - 1 - y) / (n - 1) + 0.1
            small.putpixel((x, y), colour(min(1, (dx * dx + dy * dy) ** 0.5 / 1.3)))
    gradient = small.resize((S, S), Image.BICUBIC)
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, S - 1, S - 1), radius=26 * u, fill=255)
    im.paste(gradient, (0, 0), mask)

    d = ImageDraw.Draw(im)
    white = (255, 255, 255, 255)
    d.rounded_rectangle((21 * u, 21 * u, 79 * u, 79 * u), radius=17 * u,
                        outline=white, width=round(6.5 * u))
    d.ellipse((36 * u, 36 * u, 64 * u, 64 * u), outline=white, width=round(6.5 * u))
    d.ellipse((63 * u, 29 * u, 71 * u, 37 * u), fill=white)
    return im.resize((px, px), Image.LANCZOS)


def facebook_icon_image(px=256):
    """Facebook-style white 'f' on the Facebook blue circle."""
    from PIL import ImageDraw

    S, u, im = _icon_canvas(px)
    circle = Image.new("L", (S, S), 0)
    ImageDraw.Draw(circle).ellipse((0, 0, S - 1, S - 1), fill=255)
    im.paste((24, 119, 242, 255), (0, 0), circle)  # #1877F2

    f = Image.new("L", (S, S), 0)
    d = ImageDraw.Draw(f)
    d.rectangle((43 * u, 42 * u, 57 * u, 101 * u), fill=255)  # stem
    d.rectangle((33 * u, 50 * u, 67 * u, 62 * u), fill=255)   # crossbar
    # Curved top: a quarter ring (top-left quadrant of an ellipse ring).
    hook = Image.new("L", (S, S), 0)
    h = ImageDraw.Draw(hook)
    h.ellipse((43 * u, 18 * u, 89 * u, 66 * u), fill=255)
    h.ellipse((57 * u, 31 * u, 77 * u, 53 * u), fill=0)
    h.rectangle((66 * u, 18 * u, 100 * u, 100 * u), fill=0)
    h.rectangle((0, 42 * u, 100 * u, 100 * u), fill=0)
    f.paste(255, (0, 0), hook)
    d.rectangle((66 * u, 18 * u, 69 * u, 31 * u), fill=255)  # flat end of the hook
    # The stem runs off the bottom of the circle, as in the real logo.
    f = Image.composite(f, Image.new("L", (S, S), 0), circle)
    im.paste((255, 255, 255, 255), (0, 0), f)
    return im.resize((px, px), Image.LANCZOS)


ICON_BUILDERS = {
    "email": gmail_icon_image,
    "whatsapp": whatsapp_icon_image,
    "instagram": instagram_icon_image,
    "facebook": facebook_icon_image,
}


def icon_reader(kind):
    """Use assets/logos/<kind>.png if present, otherwise the generated
    icon. Cached so each icon is loaded only once."""
    if kind not in _icon_cache:
        logo = LOGO_FOLDER / f"{kind}.png"
        if logo.is_file():
            _icon_cache[kind] = ImageReader(str(logo))
        else:
            print(f"Logo not found, using built-in icon: {logo}")
            buffer = BytesIO()
            ICON_BUILDERS[kind]().save(buffer, format="PNG")
            buffer.seek(0)
            _icon_cache[kind] = ImageReader(buffer)
    return _icon_cache[kind]


def draw_icon(pdf, kind, cx, cy, r):
    """Brand-style icon of radius r centred at (cx, cy)."""
    pdf.drawImage(icon_reader(kind), cx - r, cy - r,
                  width=2 * r, height=2 * r, mask="auto")


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
    y -= 34
    # Shrink row spacing if many contacts (e.g. several WhatsApp numbers) so
    # the list never runs into the footer.
    available = y - 45
    row_gap = min(34, available / max(1, len(links) - 1)) if len(links) > 1 else 34
    row_gap = max(row_gap, 2 * icon_r + 2)
    label_font, label_size = "Helvetica-Bold", 11
    text_font, text_size = "Helvetica", 11
    label_w = max(stringWidth(l, label_font, label_size) for _, l, _, _ in links)
    text_w = max(stringWidth(t, text_font, text_size) for _, _, t, _ in links)
    block_w = 2 * icon_r + 12 + label_w + 16 + text_w
    left = (PAGE_W - block_w) / 2
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