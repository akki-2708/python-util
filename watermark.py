"""
Add the "auravyacrystals" watermark to images (e.g. for social media posts).

The original images are never modified; watermarked copies are written to a
separate output folder with the same file names.

Default folders:
  social/               input images (originals, left unchanged)
  social_watermarked/   output (created automatically)

Run:
  py watermark.py                       # watermark new images in social/
  py watermark.py --force               # redo all, overwrite existing copies
  py watermark.py -i new_pics -o out    # custom input/output folders
  py watermark.py --text "my text" --opacity 120

Images that already have a watermarked copy in the output folder are skipped,
so after adding new photos to social/ just run the script again.

The same apply_watermark() function is used by generate_crystal_catalogue.py,
so catalogue and social images carry an identical watermark.
"""

import argparse
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps

BASE_DIR = Path(__file__).resolve().parent
INPUT_FOLDER = BASE_DIR / "images"
OUTPUT_FOLDER = BASE_DIR / "social_watermarked"

WATERMARK_TEXT = "auravyacrystals"
WATERMARK_OPACITY = 90  # 0 (invisible) - 255 (solid)
WATERMARK_FONTS = (
    "C:/Windows/Fonts/georgiab.ttf",
    "/mnt/c/Windows/Fonts/georgiab.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
)

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
JPEG_QUALITY = 95


def _load_font(text, target_width):
    """Font sized so `text` is about `target_width` pixels wide."""
    for font_path in WATERMARK_FONTS:
        if Path(font_path).is_file():
            probe = ImageFont.truetype(font_path, 100)
            size = int(100 * target_width / probe.getlength(text))
            return ImageFont.truetype(font_path, max(10, size))
    return ImageFont.load_default(size=max(10, int(target_width / len(text))))


def apply_watermark(im, text=WATERMARK_TEXT, opacity=WATERMARK_OPACITY):
    """Return a copy of im with a semi-transparent diagonal text watermark.
    Transparent images keep their alpha channel."""
    has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
    base = im.convert("RGBA")
    w, h = base.size
    # Text spans ~80% of the image diagonal.
    font = _load_font(text, (w * w + h * h) ** 0.5 * 0.8)

    # Draw the text on its own layer, rotate it along the diagonal, then blend.
    left, top, right, bottom = font.getbbox(text)
    layer = Image.new("RGBA", (right - left + 20, bottom - top + 20), (0, 0, 0, 0))
    ImageDraw.Draw(layer).text(
        (10 - left, 10 - top), text, font=font,
        fill=(255, 255, 255, opacity),
        stroke_width=max(1, font.size // 40),
        stroke_fill=(0, 0, 0, opacity // 2),  # keeps it visible on light areas
    )
    layer = layer.rotate(math.degrees(math.atan2(h, w)), expand=True, resample=Image.BICUBIC)
    base.alpha_composite(layer, ((w - layer.width) // 2, (h - layer.height) // 2))
    return base if has_alpha else base.convert("RGB")


def watermark_file(src, dst, text=WATERMARK_TEXT, opacity=WATERMARK_OPACITY):
    """Watermark one image file and save it to dst (src is not modified)."""
    with Image.open(src) as im:
        icc_profile = im.info.get("icc_profile")
        im = ImageOps.exif_transpose(im)  # apply phone rotation before drawing
        result = apply_watermark(im, text, opacity)

    dst.parent.mkdir(parents=True, exist_ok=True)
    save_args = {"icc_profile": icc_profile} if icc_profile else {}
    suffix = dst.suffix.lower()
    if suffix in (".jpg", ".jpeg"):
        result.convert("RGB").save(dst, "JPEG", quality=JPEG_QUALITY, subsampling=0,
                                   optimize=True, **save_args)
    elif suffix == ".webp":
        result.save(dst, "WEBP", quality=JPEG_QUALITY, **save_args)
    else:
        result.save(dst, **save_args)


def watermark_folder(input_folder, output_folder, text=WATERMARK_TEXT,
                     opacity=WATERMARK_OPACITY, force=False):
    input_folder, output_folder = Path(input_folder), Path(output_folder)
    if not input_folder.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_folder}")
    if input_folder.resolve() == output_folder.resolve():
        raise ValueError("Output folder must differ from input folder "
                         "so the original images are not overwritten.")

    images = sorted(
        p for p in input_folder.rglob("*")
        if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS
        and output_folder.resolve() not in p.resolve().parents  # skip output if nested
    )
    done, skipped, failed = 0, 0, []
    for src in images:
        dst = output_folder / src.relative_to(input_folder)  # keeps sub-folders
        if dst.exists() and not force:
            skipped += 1
            continue
        try:
            watermark_file(src, dst, text, opacity)
            done += 1
        except Exception as exc:  # keep going if one file is broken
            failed.append((src.name, str(exc)))

    print(f"Input:  {input_folder}")
    print(f"Output: {output_folder}")
    print(f"Watermarked: {done}, skipped (already done): {skipped}, failed: {len(failed)}")
    for name, error in failed:
        print(f"  FAILED {name}: {error}")
    if skipped and not force:
        print("Use --force to re-create existing watermarked copies.")


def main():
    parser = argparse.ArgumentParser(description="Add a text watermark to images.")
    parser.add_argument("-i", "--input", type=Path, default=INPUT_FOLDER,
                        help=f"folder with original images (default: {INPUT_FOLDER.name})")
    parser.add_argument("-o", "--output", type=Path, default=OUTPUT_FOLDER,
                        help=f"folder for watermarked copies (default: {OUTPUT_FOLDER.name})")
    parser.add_argument("--text", default=WATERMARK_TEXT, help="watermark text")
    parser.add_argument("--opacity", type=int, default=WATERMARK_OPACITY,
                        help="0 (invisible) - 255 (solid), default %(default)s")
    parser.add_argument("--force", action="store_true",
                        help="overwrite watermarked copies that already exist")
    args = parser.parse_args()
    if not 0 <= args.opacity <= 255:
        parser.error("--opacity must be between 0 and 255")
    watermark_folder(args.input, args.output, args.text, args.opacity, args.force)


if __name__ == "__main__":
    main()
